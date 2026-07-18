# coding: utf-8
"""
格式模板C - 道法【答案】【解析】格式

适用文档特征：
- 选择题答案格式: `【答案】A`
- 选择题解析格式: `【解析】考查内容` + 逐项分析（①：...）+ `故本题选X。`
- 大题答案格式: `【答案】选择小白：...`（多行，直到 `【分析】`）
- 大题考点格式: `【分析】考点考查：...`（元数据，跳过不输出）
- 大题详解格式: `【解析】步骤...`（多行，直到下一题）
- 章节标题: `第一课 青春正当时`、`一、基础训练`
- 子标题: `1．1 青春的邀请`（题号形式但后接数字）
"""

import re
import os

# 格式识别特征
TEMPLATE_FEATURES = {
    "name": "安乡金海 - 初一 - 政治",
    "patterns": [
        r"^【答案】[A-D]",        # 选择题答案: 【答案】A
        r"^【解析】",              # 解析标记
        r"^【分析】",              # 分析标记（道法文档特有）
        r"故本题选[A-D]",          # 故本题选X（道法文档特有，区别于模板B的「故选X。」）
        r"考点考查",              # 道法文档解析特有词汇
        r"能力考查",              # 道法文档解析特有词汇
        r"核心素养",              # 道法文档解析特有词汇
    ],
    "match_threshold": 0.03,
}

# 需要删除的垃圾行模式
GARBAGE_PATTERNS = [
    r"^第[一二三四五六七八九十\d]+课",   # 第一课、第二课 等章节标题
    r"^\d+[．.]\d+[\s　]",              # 子标题：1．1 / 1.1 + 空格
    r"^\d[．.]\d[^\d]",                # 子标题：5.1人 / 5．1第（个位数.个位数+非数字，如1.1、5.1）
    r"^[一二三四五六七八九十]+[、．]",   # 一、 二、 等节标题
    r"^\s*$",                           # 空行
]

# 题号模式：同时支持全角 ． 和半角 . （有些文档用 7. 而非 7．）
QUESTION_PATTERN = re.compile(
    r"^(\d+)[．.](.+)"  # [．.] = 全角句号 U+FF0E 或 半角点 U+002E
)

# 选择题答案模式: 【答案】A / 【答案】AB 等
CHOICE_ANSWER_PATTERN = re.compile(r"^【答案】\s*([A-D]+)\s*$")

# 大题答案模式: 【答案】后跟非单纯字母内容
BIG_ANSWER_PATTERN = re.compile(r"^【答案】\s*(.+)$")

# 解析标记
ANALYSIS_PATTERN = re.compile(r"^【解析】\s*(.*)$")

# 详解标记（部分道法文档用【详解】代替【解析】）
DETAIL_PATTERN = re.compile(r"^【详解】\s*(.*)$")

# 大题小问编号模式: (1) (2) (3) 或 （1） （2） （3）等（支持中文/英文括号，后面可跟空格）
SUB_QUESTION_PATTERN = re.compile(r"^[(（]\d+[)）]\s*")

# 大题题号+小问编号连排模式（如 11（1）题文 或 11(1)题文，无句号分隔）
BIG_QUESTION_WITH_SUB_PATTERN = re.compile(r"^(\d+)[(（](\d+)[)）](.*)")

# 疑问词模式 - 用于判断小问内容是问题还是答案（含简述/分析等作答动词）
QUESTION_WORDS_PATTERN = re.compile(r'什么|哪些|怎样|如何|为什么|怎么|谁|哪里|多少|几|哪|吗|简述|分析|说明|指出|概括|阐述|列举')


def is_garbage_line(text):
    """判断是否是垃圾行"""
    for pattern in GARBAGE_PATTERNS:
        if re.match(pattern, text):
            return True
    return False


def match_score(doc, cached_texts=None):
    """
    计算文档与此模板的匹配度

    Args:
        doc: WPS 文档对象

    Returns:
        float: 匹配分数 (0-1)
    """
    total_lines = 0
    matched_lines = 0

    texts = cached_texts if cached_texts is not None else [
        p.Range.Text.strip() for p in doc.Paragraphs
    ]

    for text in texts:
        if not text:
            continue

        total_lines += 1

        for pattern in TEMPLATE_FEATURES["patterns"]:
            if re.search(pattern, text):
                matched_lines += 1
                break

    if total_lines == 0:
        return 0

    return matched_lines / total_lines


def set_font_format(doc):
    """
    设置文档字体格式
    - 小四（12号）
    - 黑色
    - 宋体（中文）/ Times New Roman（英文/数字）
    - 不加粗
    """
    try:
        font = doc.Content.Font
        font.Size = 12
        font.Color = 0
        font.Name = "Times New Roman"
        font.NameFarEast = "宋体"
        font.Bold = False
        font.Italic = False
        print("   ✓ 字体格式设置完成：小四、黑色、宋体/Times New Roman、不加粗")
    except Exception as e:
        print(f"   ! 字体设置失败: {e}")


def clean_document(doc):
    """
    清洗文档（WPS COM 对象）

    Args:
        doc: WPS 文档对象

    Returns:
        bool: 是否成功
    """
    print(f"   ▶ 使用模板C清洗: {doc.Name}")
    print(f"   📊 文档共 {doc.Paragraphs.Count} 个段落")

    # 1. 删除所有图片
    try:
        shapes_count = doc.Shapes.Count
        if shapes_count > 0:
            for i in range(shapes_count, 0, -1):
                doc.Shapes(i).Delete()
            print(f"   ✓ 删除 {shapes_count} 个悬浮图片")
        inline_count = doc.InlineShapes.Count
        if inline_count > 0:
            for i in range(inline_count, 0, -1):
                doc.InlineShapes(i).Delete()
            print(f"   ✓ 删除 {inline_count} 个嵌入图片")
    except Exception as e:
        print(f"   ! 删除图片失败: {e}")

    # 2. 第一遍遍历：收集所有有效内容
    paras = doc.Paragraphs
    questions_list = []    # 按扫描顺序存储
    current_question = None

    print(f"\n   🔍 开始解析文档内容...\n")

    i = 1
    while i <= paras.Count:
        p = paras(i)
        # 获取段落文本
        text = p.Range.Text.replace("\r", "").replace("\n", "").strip()
        
        # 获取 Word 列表格式标记（如自动编号 (1)(2)(3) 等）
        list_format = p.Range.ListFormat
        if list_format.ListType > 0:  # 有列表格式
            list_string = list_format.ListString
            if list_string:
                text = list_string + text
        
        if not text:
            i += 1
            continue

        debug_text = text[:60] + "..." if len(text) > 60 else text

        # 垃圾行过滤
        if is_garbage_line(text):
            print(f"   [行{i:3d}] 🚫 跳过垃圾行: {debug_text}")
            i += 1
            continue

        # ── 新格式：题号+小问编号连排（如 11（1）题文，无句号分隔）────────
        bq_sub_match = BIG_QUESTION_WITH_SUB_PATTERN.match(text)
        if bq_sub_match and not QUESTION_PATTERN.match(text):
            q_num = int(bq_sub_match.group(1))
            sub_num = bq_sub_match.group(2)
            sub_content = bq_sub_match.group(3).strip()
            # 创建新大题
            current_question = {
                'original_num': q_num,
                'type': 'big',
                'answer_lines': [],
                'analysis_lines': [],
            }
            questions_list.append(current_question)
            print(f"   [行{i:3d}] 📌 发现大题: 第{q_num}题（题号+小问连排）(顺序{len(questions_list)})")
            # 嵌入的小问视为问题行，跳过并向后查找【答案】
            sub_text = f"（{sub_num}）{sub_content}"
            print(f"   [行{i:3d}] ⏭️  跳过嵌入小问问题: {sub_text[:40]}")
            i += 1
            while i <= paras.Count:
                next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not next_text:
                    i += 1
                    continue
                if (SUB_QUESTION_PATTERN.match(next_text)
                        or QUESTION_PATTERN.match(next_text)
                        or BIG_QUESTION_WITH_SUB_PATTERN.match(next_text)):
                    break
                if next_text.startswith('【答案】'):
                    answer_body = next_text[4:]
                    if SUB_QUESTION_PATTERN.match(answer_body):
                        answer_content = answer_body
                    else:
                        answer_content = f"（{sub_num}）" + answer_body
                    current_question['answer_lines'].append(answer_content)
                    print(f"   [行{i:3d}]    └─ 大题答案: {answer_content[:40]}")
                    i += 1
                    # 收集续行，直到遇到小问编号、新题目或垃圾行
                    while i <= paras.Count:
                        cont_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                        if not cont_text:
                            i += 1
                            continue
                        if (SUB_QUESTION_PATTERN.match(cont_text) or QUESTION_PATTERN.match(cont_text)
                                or BIG_QUESTION_WITH_SUB_PATTERN.match(cont_text)
                                or is_garbage_line(cont_text) or cont_text.startswith('【子议题')):
                            break
                        current_question['answer_lines'][-1] += cont_text
                        print(f"   [行{i:3d}]       └─ 追加续行: {cont_text[:40]}")
                        i += 1
                    break
                else:
                    i += 1
            continue

        # ── 新题目（数字开头）──────────────────────────────────────────
        q_match = QUESTION_PATTERN.match(text)
        if q_match:
            q_num = int(q_match.group(1))
            current_question = {
                'original_num': q_num,
                'type': None,           # 'choice' 或 'big'
                'answer_lines': [],     # 选择题: ['A']；大题: ['选择小白：...', ...]
                'analysis_lines': [],   # 【解析】内容行列表（不含标记本身）
            }
            questions_list.append(current_question)
            print(f"   [行{i:3d}] 📌 发现题目: 第{q_num}题 (顺序{len(questions_list)})")
            i += 1
            continue

        # ── 选择题答案 【答案】A ────────────────────────────────────────
        choice_match = CHOICE_ANSWER_PATTERN.match(text)
        if choice_match and current_question:
            current_question['type'] = 'choice'
            current_question['answer_lines'] = [choice_match.group(1)]
            print(f"   [行{i:3d}] ✅ 选择题答案: {choice_match.group(1)}")
            i += 1
            continue

        # ── 大题答案 【答案】选择小白：... ────────────────────────────
        # 只在尚未识别类型且不是小问格式时触发（避免误匹配）
        big_match = BIG_ANSWER_PATTERN.match(text)
        if big_match and current_question and current_question['type'] is None:
            # 检查后面是否紧跟着小问编号，如果是则跳过（让小问逻辑处理）
            peek_i = i + 1
            has_immediate_sub_question = False
            while peek_i <= paras.Count:
                peek_text = paras(peek_i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not peek_text:
                    peek_i += 1
                    continue
                if SUB_QUESTION_PATTERN.match(peek_text):
                    has_immediate_sub_question = True
                    break
                # 如果遇到其他内容则停止检查
                if peek_text or QUESTION_PATTERN.match(peek_text):
                    break
                peek_i += 1
            
            if has_immediate_sub_question:
                # 后面有小问编号，跳过此【答案】让小问逻辑处理
                print(f"   [行{i:3d}] ⏭️  跳过【答案】（后面有小问编号）")
                i += 1
                continue
            
            # 真正的大题答案（无小问）
            current_question['type'] = 'big'
            first_ans = big_match.group(1)
            answer_lines = [first_ans]
            print(f"   [行{i:3d}] 📋 大题答案: {first_ans[:40]}")
            i += 1
            # 向后收集续行，直到遇到 【分析】/【解析】/【答案】/小问编号 或新题目
            while i <= paras.Count:
                next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not next_text:
                    i += 1
                    continue
                if (re.match(r'^【分析】|^【解析】|^【答案】', next_text)
                        or QUESTION_PATTERN.match(next_text)
                        or BIG_QUESTION_WITH_SUB_PATTERN.match(next_text)
                        or SUB_QUESTION_PATTERN.match(next_text)):
                    break
                answer_lines.append(next_text)
                print(f"   [行{i:3d}]    └─ 追加答案行: {next_text[:40]}")
                i += 1
            current_question['answer_lines'] = answer_lines
            continue

        # ── 【分析】处理：可能是元数据，也可能是唯一解析源 ──────────────
        if re.match(r'^【分析】', text) and current_question:
            # 先向后探查：本【分析】块之后是否有【解析】？
            has_analysis_followed_by_jiexi = False
            peek_i = i + 1
            while peek_i <= paras.Count:
                peek_text = paras(peek_i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not peek_text:
                    peek_i += 1
                    continue
                if re.match(r'^【解析】', peek_text):
                    has_analysis_followed_by_jiexi = True
                    break
                if QUESTION_PATTERN.match(peek_text) or re.match(r'^【答案】', peek_text):
                    break
                peek_i += 1

            if has_analysis_followed_by_jiexi:
                # 【分析】只是元数据，跳过
                print(f"   [行{i:3d}] ⏭️  跳过【分析】考点元数据（后续有【解析】）")
                i += 1
                while i <= paras.Count:
                    next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                    if not next_text:
                        i += 1
                        continue
                    if (re.match(r'^【解析】', next_text) or QUESTION_PATTERN.match(next_text)
                            or BIG_QUESTION_WITH_SUB_PATTERN.match(next_text)):
                        break
                    print(f"   [行{i:3d}]    └─ 跳过元数据: {next_text[:40]}")
                    i += 1
                continue
            else:
                # 无【解析】，把【分析】当作解析收集（替换标记为"解析："）
                first_line = re.sub(r'^【分析】\s*', '', text)
                analysis_lines = [first_line] if first_line else []
                print(f"   [行{i:3d}] 🔍 发现【分析】（无【解析】，收为解析）: {first_line[:40] if first_line else '(空)'}")
                i += 1
                while i <= paras.Count:
                    next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                    if not next_text:
                        i += 1
                        continue
                    if (QUESTION_PATTERN.match(next_text)
                            or BIG_QUESTION_WITH_SUB_PATTERN.match(next_text)
                            or re.match(r'^【答案】', next_text)
                            or is_garbage_line(next_text)):
                        break
                    analysis_lines.append(next_text)
                    i += 1
                current_question['analysis_lines'] = analysis_lines
                print(f"   [行{i:3d}]    └─ 解析共 {len(analysis_lines)} 行（来自【分析】）")
                continue

        # ── 【解析】收集 ────────────────────────────────────────────────
        analysis_match = ANALYSIS_PATTERN.match(text)
        if analysis_match and current_question:
            first_line = analysis_match.group(1)
            analysis_lines = [first_line] if first_line else []
            print(f"   [行{i:3d}] 🔍 发现【解析】: {first_line[:40] if first_line else '(空)'}")
            i += 1
            # 收集后续行，直到遇到下一题、下一个【答案】或垃圾行（如章节标题）
            while i <= paras.Count:
                next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not next_text:
                    i += 1
                    continue
                if (QUESTION_PATTERN.match(next_text)
                        or BIG_QUESTION_WITH_SUB_PATTERN.match(next_text)
                        or re.match(r'^【答案】', next_text)
                        or is_garbage_line(next_text)):
                    break
                analysis_lines.append(next_text)
                i += 1
            current_question['analysis_lines'] = analysis_lines
            print(f"   [行{i:3d}]    └─ 解析共 {len(analysis_lines)} 行")
            continue

        # ── 【详解】收集（部分道法文档用【详解】代替【解析】）────────────
        detail_match = DETAIL_PATTERN.match(text)
        if detail_match and current_question:
            first_line = detail_match.group(1)
            analysis_lines = [first_line] if first_line else []
            print(f"   [行{i:3d}] 🔍 发现【详解】: {first_line[:40] if first_line else '(空)'}")
            i += 1
            # 收集后续行，直到遇到下一题、下一个【答案】或垃圾行
            while i <= paras.Count:
                next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not next_text:
                    i += 1
                    continue
                if (QUESTION_PATTERN.match(next_text)
                        or BIG_QUESTION_WITH_SUB_PATTERN.match(next_text)
                        or re.match(r'^【答案】', next_text)
                        or is_garbage_line(next_text)):
                    break
                analysis_lines.append(next_text)
                i += 1
            current_question['analysis_lines'] = analysis_lines
            print(f"   [行{i:3d}]    └─ 解析共 {len(analysis_lines)} 行（来自【详解】）")
            continue

        # ── 大题小问编号处理 (1)(2)(3)... ────────────────────────────────
        sub_q_match = SUB_QUESTION_PATTERN.match(text)
        if sub_q_match and current_question:
            # 如果当前题目还没确定类型，检测到小问编号说明这是大题
            if current_question['type'] is None:
                current_question['type'] = 'big'
                print(f"   [行{i:3d}] 📋 检测到小问编号，标记为大题")
            
            # 检查小问问题行后面是否直接跟答案内容（没有【答案】标记）
            # 格式如: (2)孝悌忠信、礼义廉耻...
            question_content = text[sub_q_match.end():].strip()  # 去掉 (1)(2) 后的内容
            
            # 判断是问题还是答案：包含疑问词的是问题，不包含的可能是答案
            has_question_word = QUESTION_WORDS_PATTERN.search(question_content) is not None
            
            if question_content and not has_question_word:
                # 不包含疑问词，可能是答案直接跟在编号后面
                # 向后查看一行，找到下列任一情况都说明当前行是答案：
                # - 下一行是新小问编号
                # - 下一行是新题目
                # - 下一行是解析标记（解析是题目结束的标志）
                # - 已到文档末尾（最后一个小问答案）
                peek_i = i + 1
                is_direct_answer = False
                found_peek = False
                while peek_i <= paras.Count:
                    peek_text = paras(peek_i).Range.Text.replace("\r", "").replace("\n", "").strip()
                    if not peek_text:
                        peek_i += 1
                        continue
                    found_peek = True
                    # 下一个是小问编号、新题目、解析标记→ 当前行是答案
                    if (SUB_QUESTION_PATTERN.match(peek_text) or QUESTION_PATTERN.match(peek_text)
                            or BIG_QUESTION_WITH_SUB_PATTERN.match(peek_text)
                            or re.match(r'^【解析】|^【分析】|^【详解】|^【讲解】|^解析：|^详解：', peek_text)):
                        is_direct_answer = True
                    break
                # 已到文档末尾，也是直接答案
                if not found_peek:
                    is_direct_answer = True
                
                if is_direct_answer:
                    # 当前行就是答案，直接收集
                    answer_content = sub_q_match.group(0) + question_content
                    current_question['answer_lines'].append(answer_content)
                    print(f"   [行{i:3d}]    └─ 大题答案(直接): {answer_content[:40]}")
                    i += 1
                    continue
            
            # 标准格式：跳过小问问题本身，向后查找【答案】
            print(f"   [行{i:3d}] ⏭️  跳过小问问题: {text[:40]}")
            i += 1
            # 向后查找以"答案："或"【答案】"开头的行
            while i <= paras.Count:
                next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not next_text:
                    i += 1
                    continue
                # 如果遇到新的小问编号，停止当前收集（让外层循环处理新小问）
                if SUB_QUESTION_PATTERN.match(next_text):
                    break
                # 如果遇到新题目，停止收集
                if QUESTION_PATTERN.match(next_text) or BIG_QUESTION_WITH_SUB_PATTERN.match(next_text):
                    break
                # 如果以"答案："开头，收集为答案（替换前缀为小问编号）
                if next_text.startswith('答案：'):
                    # 使用当前小问编号（如 (1) (2)）替换"答案："
                    answer_content = sub_q_match.group(0) + next_text[3:]  # sub_q_match.group(0) 是 (1) 或 (2)
                    current_question['answer_lines'].append(answer_content)
                    print(f"   [行{i:3d}]    └─ 大题答案: {answer_content[:40]}")
                    i += 1
                    # 继续收集续行，直到遇到下一个小问、新题目或垃圾行
                    while i <= paras.Count:
                        cont_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                        if not cont_text:
                            i += 1
                            continue
                        # 遇到终止条件则停止（包括子议题标记和下一个小问）
                        if (SUB_QUESTION_PATTERN.match(cont_text) or QUESTION_PATTERN.match(cont_text)
                                or BIG_QUESTION_WITH_SUB_PATTERN.match(cont_text)
                                or is_garbage_line(cont_text) or cont_text.startswith('【子议题')):
                            break
                        # 追加续行到最后一行答案
                        current_question['answer_lines'][-1] += cont_text
                        print(f"   [行{i:3d}]       └─ 追加续行: {cont_text[:40]}")
                        i += 1
                    break  # 答案收集完成，跳出查找循环
                # 如果以"【答案】"开头，收集为答案（替换前缀为小问编号）
                elif next_text.startswith('【答案】'):
                    answer_body = next_text[4:]  # 去掉【答案】
                    # 如果答案内容已经以小问编号开头（如(1)诚信），直接使用，不重复拼接前缀
                    if SUB_QUESTION_PATTERN.match(answer_body):
                        answer_content = answer_body
                    else:
                        answer_content = sub_q_match.group(0) + answer_body  # 【答案】是4个字符
                    current_question['answer_lines'].append(answer_content)
                    print(f"   [行{i:3d}]    └─ 大题答案: {answer_content[:40]}")
                    i += 1
                    # 继续收集续行，直到遇到下一个小问、新题目或垃圾行
                    while i <= paras.Count:
                        cont_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                        if not cont_text:
                            i += 1
                            continue
                        # 遇到终止条件则停止（包括子议题标记和下一个小问）
                        if (SUB_QUESTION_PATTERN.match(cont_text) or QUESTION_PATTERN.match(cont_text)
                                or BIG_QUESTION_WITH_SUB_PATTERN.match(cont_text)
                                or is_garbage_line(cont_text) or cont_text.startswith('【子议题')):
                            break
                        # 追加续行到最后一行答案
                        current_question['answer_lines'][-1] += cont_text
                        print(f"   [行{i:3d}]       └─ 追加续行: {cont_text[:40]}")
                        i += 1
                    break  # 答案收集完成，跳出查找循环
                else:
                    i += 1
            continue

        # ── 其他行：若还未找到答案（题干/选项/图注等），直接跳过 ────────
        if current_question and current_question['type'] is None:
            # 【诊断】检查是否包含小问编号但被错过了
            if SUB_QUESTION_PATTERN.match(text):
                print(f"   [行{i:3d}] ⚠️  小问编号被跳过: {debug_text}")
            else:
                print(f"   [行{i:3d}] ⏭️  跳过题目内容: {debug_text}")
            i += 1
            continue

        # 未识别行
        print(f"   [行{i:3d}] ⚪ 未识别: {debug_text}")
        i += 1

    # 3. 统计
    print(f"\n   📊 解析完成统计:")
    print(f"      - 共发现 {len(questions_list)} 道题目（含重复/错序）")
    choice_found = sum(1 for q in questions_list if q['type'] == 'choice')
    big_found = sum(1 for q in questions_list if q['type'] == 'big')
    untyped_found = sum(1 for q in questions_list if q['type'] is None)
    print(f"      - 选择题: {choice_found} 道，大题: {big_found} 道")
    if untyped_found > 0:
        print(f"      - ⚠️  未识别类型（可能是原文顺序错乱导致的孤立题目）: {untyped_found} 道")

    # 4. 清空文档并重新写入（跳过无类型题目）
    doc.Content.Text = ""

    choice_count = 0
    big_count = 0
    output_idx = 0

    for q_data in questions_list:
        q_type = q_data['type']
        answer_lines = q_data['answer_lines']
        analysis_lines = q_data['analysis_lines']

        if q_type is None:
            print(f"   ⚠️  原第{q_data['original_num']}题类型未识别，跳过")
            continue

        output_idx += 1

        if q_type == 'choice':
            choice_count += 1
            letter = answer_lines[0] if answer_lines else ''
            doc.Content.InsertAfter(f"{output_idx}．{letter}\r")
            if analysis_lines:
                doc.Content.InsertAfter(f"解析：{analysis_lines[0]}\r")
                for line in analysis_lines[1:]:
                    doc.Content.InsertAfter(f"{line}\r")
            else:
                # 无解析时补空格，确保 F3 可正常跳转到下一题
                doc.Content.InsertAfter(f"解析： \r")

        elif q_type == 'big':
            big_count += 1
            doc.Content.InsertAfter(f"{output_idx}．\r")
            doc.Content.InsertAfter("答案：\r")
            for ans_line in answer_lines:
                doc.Content.InsertAfter(f"{ans_line}\r")
            if analysis_lines:
                doc.Content.InsertAfter(f"解析：{analysis_lines[0]}\r")
                for line in analysis_lines[1:]:
                    doc.Content.InsertAfter(f"{line}\r")
            else:
                doc.Content.InsertAfter(f"解析： \r")

    print(f"   ✓ 共处理 {choice_count} 道选择题, {big_count} 道大题，输出共 {output_idx} 题")

    # 5. 设置字体格式
    set_font_format(doc)

    return True


# 模块导出
TEMPLATE_INFO = TEMPLATE_FEATURES
