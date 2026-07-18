# coding: utf-8
"""
答案格式清洗模块 - LLM 智能版

使用 DeepSeek-V3 大模型进行语义分析和格式清洗。

此模块依赖外部 LLM 服务（硅基流动 API）。
"""

import urllib.request
import json
import re
import os
import sys
import time
import concurrent.futures

# 处理导入路径，支持直接运行和模块运行两种方式
try:
    # 尝试相对导入（作为模块运行时）
    from .common import debug_log, get_active_wps
except ImportError:
    # 相对导入失败时，使用绝对导入（直接运行时）
    from common import debug_log, get_active_wps


def ask_llm(prompt):
    api_key = os.getenv("SILICONFLOW_API_KEY") or os.getenv("DEEPSEEK_API_KEY")
    if not api_key:
        debug_log("未检测到 API Key，请先设置环境变量 SILICONFLOW_API_KEY", "error")
        return None

    url = "https://api.siliconflow.cn/v1/chat/completions"
    data = {
        "model": "deepseek-ai/DeepSeek-V3",
        "messages": [
            {"role": "system", "content": "你是一个严谨的 JSON 数据处理程序。"},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.0,
        "max_tokens": 4000,
    }
    debug_log("发送请求到 DeepSeek...", "llm")
    req = urllib.request.Request(
        url,
        data=json.dumps(data).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
    )
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=120) as response:
                result = json.loads(response.read().decode("utf-8"))
                content = result["choices"][0]["message"]["content"]
                debug_log(f"LLM 返回内容长度: {len(content)}", "llm")
                return content
        except Exception as e:
            debug_log(f"LLM 请求失败 (尝试 {attempt + 1}/3): {e}", "error")
            time.sleep(2)
    return None


def extract_json(text):
    match = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        return match.group(0).strip()
    return text


def process_document_llm(doc=None):
    """
    使用 LLM 清洗文档
    
    Args:
        doc: WPS 文档对象，如果为 None 则自动获取当前活动文档
    """
    print("   ⚠️ format_answers_deepseek.py 已停用，避免继续走失效旧链路。")
    print("   👉 请改用: python 格式处理/main.py")
    print("   👉 或使用: python 格式处理/format_answers_ai.py")
    return False

    try:
        if doc is None:
            wps = get_active_wps()
            if not wps:
                print("   ❌ 无法连接到 WPS，请重试")
                return
            doc = wps.ActiveDocument
        doc_name = doc.Name
    except Exception as e:
        print(f"   ❌ 获取文档失败 (RPC错误): {e}")
        return

    print(f"   ▶️ 正在使用大模型清洗文档: {doc_name}")

    # 提前处理另存为逻辑，避免破坏原文档
    out_path = os.path.splitext(doc.FullName)[0] + "_标准化(大模型版).docx"
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except Exception as e:
            print(f"   ❌ 无法删除已有的标准化文档，请先关闭它。错误: {e}")
            return

    try:
        doc.SaveAs2(out_path)
    except Exception as e:
        print(f"   ❌ 另存为标准化文档失败，请检查文件是否被占用。错误: {e}")
        return

    try:
        paras = doc.Paragraphs
        count = paras.Count
    except Exception as e:
        print(f"   ❌ 读取段落失败 (RPC错误): {e}")
        return

    text_lines = []
    q_num_map = {}
    last_seen_q = None
    auto_q_num = 1
    q_num_pattern = re.compile(r"^\s*(\d+)[．.]")
    auto_q_num = 1

    for i in range(1, count + 1):
        # 用占位符标记包含图片和表格的行，防止 LLM 误删
        try:
            p = paras(i)
            has_rich = False
            if p.Range.Information(12):
                has_rich = True  # Table
            if p.Range.InlineShapes.Count > 0:
                has_rich = True

            try:
                list_str = p.Range.ListFormat.ListString
            except:
                list_str = ""

            raw_text = (
                p.Range.Text.replace("\r", "")
                .replace("\n", "")
                .replace("\x07", "")
                .strip()
            )

            if list_str:
                raw_text = list_str + " " + raw_text

            if not raw_text and has_rich:
                text_lines.append(f"[line {i}] [包含图片] 或 [表格内容]")
            else:
                # 截断过长文本，防止超过 LLM context
                match_q = q_num_pattern.match(raw_text)
                if match_q:
                    last_seen_q = match_q.group(1)

                # If there is no last_seen_q, but we hit an Answer, we use auto_q_num
                match_ans = re.search(
                    r"^\s*[【\[<]?(答案|参考答案)[】\]>]?\s*", raw_text
                )
                if match_ans:
                    if last_seen_q is None:
                        q_num_map[i] = str(auto_q_num)
                        auto_q_num += 1
                    else:
                        q_num_map[i] = last_seen_q
                else:
                    q_num_map[i] = last_seen_q

                text_lines.append(f"[line {i}] {raw_text[:100]}")
        except:
            text_lines.append(f"[line {i}] [读取错误]")

    # 为了防止上下文超限，分批处理
    batch_size = 150
    actions = {}

    print(f"   🤖 文档共 {count} 行，开始分批发送给大模型进行语义判定...")

    def process_batch(start_idx):
        batch = text_lines[start_idx : start_idx + batch_size]
        prompt = f"""分析以下试卷的文本行，判断每一行应该被删除还是保留，并提取相关题号。
        规则：
        1. 绝对删除 ("d"): 
           - 所有的题干（如“1. 下列关于...”或“(1)操作步骤”）、多项选择题的选项(A、B、C、D)、卷头废话、图片占位符、无用总结（如“故选C。”）。
           - 注意：不要误删真正的答案文本（如“15. 说法错误”属于答案首行，或者“①因为与异性...”属于答案的后续行，绝不能删）。
        2. 绝对保留 ("k"): 
           - 答案和解析内容必须保留！包括带有 `【答案】`、`[答案]`、`答案：` 开头的行，以及带有 `【解析】`、`【分析】`、`【详解】`、`解析：` 开头的行。
           - 如果一行文字是纯粹的答案（如“6. C”或“1. (1) cd”），哪怕它没有【答案】标签，也必须保留！
           - 即使没有这些标签，只要内容明显是解答、分析、评价步骤（如“第一步”、“正误判断”、“论据”等，或者选项的解释“A：...”、“B：...”），一律算作解析，必须保留！
           - 标记为 `[包含图片]` 或 `[表格内容]` 的行，如果明显属于题目/选项则删除，如果属于答案/解析则保留。
        3. 提取题号 ("k[题号]"):
           - 对于每一道题的“答案开始的第一行”（例如“【答案】C”），提取主题号并附加在k后面。
           - ⚠️ 核心要求：如果是大题的小问（如“(1)”、“①”），你**必须**将小问序号和主题号拼在一起提取！例如：主题号是10，当前是第(1)小问的答案，输出 "k10(1)"；是第(2)小问的答案，输出 "k10(2)"。如果你要删除的题干里包含小问序号，你必须把这个序号转移保存到紧跟其后的答案行的 k 标签里！绝不能只输出 "k10"！如果没有小问的普通题，只输出 "k15"。

        为了极致压缩输出体积，请返回一个极其紧凑的 JSON 对象（不是数组），键为行号字符串，值为 "d", "k", 或 "k题号"。不准有任何 Markdown 代码块符号前缀，严格要求：必须是标准合法的 JSON 格式，绝不能有单引号、注释或尾随逗号，直接输出纯 JSON：
        {{
          "1": "d",
          "5": "k1",
          "6": "k",
          "10": "k10(1)"
        }}
        文本：
        {chr(10).join(batch)}
        """
        debug_log(
            f"开始并发处理第 {start_idx} 到 {start_idx + batch_size} 行...", "info"
        )
        response = ask_llm(prompt)
        if not response:
            return {}

        json_str = extract_json(response)
        try:
            debug_log(f"第 {start_idx} 行批次返回 JSON (长度 {len(json_str)})", "json")
            return json.loads(json_str)
        except Exception as e:
            print(f"   ⚠️ 第 {start_idx}-{start_idx + batch_size} 行 JSON 解析失败: {e}")
            return {}

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = {
            executor.submit(process_batch, idx): idx
            for idx in range(0, count, batch_size)
        }
        for future in concurrent.futures.as_completed(futures):
            batch_result = future.result()
            for line_str, action_code in batch_result.items():
                try:
                    line_num = int(line_str)
                    if action_code == "d":
                        actions[line_num] = {"action": "delete"}
                    elif action_code.startswith("k"):
                        q_num_match = re.search(r"^k(.+)$", action_code)
                        actions[line_num] = {
                            "action": "keep",
                            "q_num": q_num_match.group(1).strip()
                            if q_num_match
                            else None,
                        }
                except ValueError:
                    pass
    print(f"   🔪 模型判定完毕，开始执行删除操作...")

    # 预处理：找出每个大题的第一次出现的答案行号
    first_line_for_main_q = {}
    last_deleted_sub_q = None
    for i in range(1, count + 1):
        action_item = actions.get(i, {})
        action = action_item.get("action")
        
        # 记录刚刚被删除的带有小问号的题干
        if action == "delete":
            try:
                try:
                    list_str = paras(i).Range.ListFormat.ListString
                except:
                    list_str = ""
                text = paras(i).Range.Text.replace('\r', '').replace('\n', '').replace('\x07', '').strip()
                if list_str:
                    text = list_str + " " + text
                
                # ✨ 核心修复：限制只捕获 (1) 这种明确的小题号，坚决不捕获 ①②③，避免把选项里的序号当作小题号！
                match = re.match(r'^([(（]\d+[)）])', text)
                if match:
                    last_deleted_sub_q = match.group(1)
            except:
                pass

        if action == "keep":
            q_num = action_item.get("q_num")
            if q_num is None:
                q_num = q_num_map.get(i)
            if q_num:
                clean_q_num = str(q_num).replace(" ", "")
                
                # 如果这个 q_num 只是一个纯数字的主题号，并且我们刚刚记下了一个小问号，就把它接上去！
                if re.match(r'^\d+$', clean_q_num) and last_deleted_sub_q:
                    clean_q_num = clean_q_num + last_deleted_sub_q
                    action_item["q_num"] = clean_q_num
                    
                sub_match = re.match(r"^(\d+)([(（].*[)）]|[①-⑳])$", clean_q_num)
                main_q = sub_match.group(1) if sub_match else clean_q_num
                if main_q not in first_line_for_main_q:
                    first_line_for_main_q[main_q] = i
                    
            # 遇到有效答案行后，清空记录，避免错误传递
            last_deleted_sub_q = None

    # 倒序删除
    deleted_count = 0
    for i in range(count, 0, -1):
        action_item = actions.get(i, {})
        action = action_item.get("action")

        if action == "delete":
            try:
                if paras(i).Range.Information(12):  # 如果在一个表格里
                    paras(i).Range.Tables(1).Delete()
                    deleted_count += 1
                    debug_log(f"行 {i} [删除]: 抹除整张题目表格", "action")
                else:
                    text_to_delete = (
                        paras(i)
                        .Range.Text.replace("\r", "")
                        .replace("\n", "")
                        .replace("\x07", "")[:20]
                    )
                    paras(i).Range.Delete()
                    deleted_count += 1
                    debug_log(f"行 {i} [删除]: {text_to_delete}", "action")
            except Exception as e:
                # 如果是多行表格，第一行删掉后整个表格都没了，后续行再操作会报错，直接忽略即可
                pass
        elif action == "keep":
            q_num = action_item.get("q_num")
            if q_num is None:
                q_num = q_num_map.get(i)
            try:
                rng = paras(i).Range
                line_text = (
                    rng.Text.replace("\r", "").replace("\n", "").replace("\x07", "")
                )

                if q_num is not None:
                    match_ans = re.search(
                        r"^\s*[【\[<]?(?:答案|参考答案)[】\]>]?\s*", line_text
                    )
                    if match_ans:
                        rng_replace = doc.Range(
                            rng.Start, rng.Start + len(match_ans.group(0))
                        )

                        clean_q_num = str(q_num).replace(" ", "")
                        sub_match = re.match(
                            r"^(\d+)([(（].*[)）]|[①-⑳])$", clean_q_num
                        )

                        if sub_match:
                            main_q = sub_match.group(1)
                            sub_q = sub_match.group(2)
                            if i <= first_line_for_main_q.get(main_q, i):
                                rng_replace.Text = f"{main_q}．{sub_q}"
                            else:
                                rng_replace.Text = f"{sub_q}"
                        else:
                            rng_replace.Text = f"{clean_q_num}．"

                        debug_log(
                            f"行 {i} [保留/替换答案头]: 注入题号 {rng_replace.Text}",
                            "action",
                        )

                match_ana = re.search(
                    r"^\s*[【\[<]?(解析|分析|解答)[】\]>]?\s*", line_text
                )
                if match_ana:
                    rng_replace = doc.Range(
                        rng.Start, rng.Start + len(match_ana.group(0))
                    )
                    rng_replace.Text = "解析："
                    debug_log(f"行 {i} [保留/替换解析头]: 变更为 解析：", "action")

                match_detail = re.search(r"^\s*[【\[]详解[】\]]\s*", line_text)
                if match_detail:
                    rng_replace = doc.Range(
                        rng.Start, rng.Start + len(match_detail.group(0))
                    )
                    rng_replace.Text = ""
                    debug_log(f"行 {i} [保留/剥除详解头]", "action")
            except:
                pass

    print(f"   ✅ 基于大模型的清洗完成，共删除 {deleted_count} 行无效内容。")
    print(f"   ⚠️ 注意：正在自动调用本地正则引擎进行段落合并...")
    try:
        format_answers.process_document(doc, keep_pictures=True, is_chained=True)
        print(f"   ✅ 本地正则合并完成。")
    except Exception as e:
        print(f"   ❌ 本地合并异常: {e}")

    # --- ✨ 新增：将全文颜色强制设置为纯黑色，并统一字体和字号 ---
    try:
        # wdColorAutomatic 为 0xFF000000（-16777216），为了保证纯黑，直接使用 wdColorBlack = 0
        font = doc.Content.Font
        font.Color = 0

        # 统一设置字体为：中文宋体，英文 Times New Roman
        font.Name = "Times New Roman"
        font.NameFarEast = "宋体"

        # 统一设置字号为小四（12磅）
        font.Size = 12

        # 取消可能存在的加粗和倾斜
        font.Bold = False
        font.Italic = False

        print(
            "   ✅ 全文已刷成黑色，并统一格式为：宋体/Times New Roman，小四，取消加粗/斜体"
        )
    except Exception as e:
        print(f"   ⚠️ 设置字体样式失败: {e}")

    try:
        doc.Save()
        print(f"   ✅ 成功生成大模型清洗版文档: {os.path.basename(out_path)}")
    except Exception as e:
        print(f"   ❌ 保存文档失败: {e}")


# LLM 版本的入口函数
def run_llm_engine(doc):
    """
    运行 LLM 智能引擎清洗文档
    
    Args:
        doc: WPS 文档对象
        
    Returns:
        bool: 是否成功
    """
    return bool(process_document_llm(doc))


# 向后兼容：保留 main 函数但提示用户使用新入口
if __name__ == "__main__":
    print("=" * 60)
    print("⚠️  提示：请使用新的统一入口运行")
    print("   运行方式: python -m modules.answer_format.main")
    print("=" * 60)
    print()
