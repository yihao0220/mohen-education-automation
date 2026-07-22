# coding: utf-8
"""
格式模板B - 【分析】【详解】格式

适用文档特征：
- 选择题答案格式: `【答案】X` 如 `【答案】A`
- 选择题解析格式: `【分析】...` + `【详解】...` + `故选X。`
- 大题答案格式: `【答案】(1)xxx (2)xxx`
- 大题解析格式: `【分析】...` + `【详解】（1）...（2）...`
- 章节标题: `一、选择题`、`二、非选择题`
"""

import re
import os

from . import matches_garbage_pattern, pattern_match_score, set_standard_font

# 格式识别特征
TEMPLATE_FEATURES = {
    "name": "安乡金海 - 初一 - 生物",
    "patterns": [
        r"^【答案】",               # 答案标记
        r"^【分析】",               # 分析标记
        r"^【详解】",               # 详解标记
        r"故选[A-D]。$",           # 故选A。/故选B。等
    ],
    "match_threshold": 0.03,
}

# 需要删除的垃圾行模式
GARBAGE_PATTERNS = [
    r"^课时分层作业",
    r"^\(建议用时",
    r"^题组[一二三四五六七八九十]",
    r"^能力提升卷",
    r"^第\d+章",
    r"^一、选择题",
    r"^二、非选择题",
    r"^本题包括\d+小题",
    r"^每小题",
    r"^基础\+提升\+长句表达",
    r"^限时\d+层提升",
    r"^\d+\.\d+\.",                # 章节编号（如 3.7.2 食物链和食物网）
    r"^\d+[．.]\d+",               # 子标题：5.1 / 5．1 人要自强（数字.数字开头）
    r"^\s*$",                      # 空行
]

# 小问模式 (1) (2) (3) 等 - 用于识别大题的小问
SUB_QUESTION_PATTERN = re.compile(r"^\(\d+\)")

# 选择题题干模式（支持数字、中文、英文、书名号、括号、引号开头）
QUESTION_PATTERN = re.compile(r"^(\d+)[．.]([\s\"\'\u4e00-\u9fa5a-zA-Z0-9（(《].+)")

# 选择题答案模式: 【答案】A
CHOICE_ANSWER_PATTERN = re.compile(r"^【答案】\s*([A-D]+)$")

# 大题答案模式: 【答案】(1)xxx (2)xxx
BIG_QUESTION_ANSWER = re.compile(r"^【答案】\s*(.+)$")

# 分析标记
ANALYSIS_PATTERN = re.compile(r"^【分析】\s*(.*)$")

# 详解标记
DETAIL_PATTERN = re.compile(r"^【详解】\s*(.*)$")

# 故选X标记
ANSWER_END_PATTERN = re.compile(r"故选[A-D]。$")


def is_garbage_line(text):
    return matches_garbage_pattern(text, GARBAGE_PATTERNS)


def is_question_line(text):
    """判断是否是题目行"""
    return bool(QUESTION_PATTERN.match(text))


def is_option_line(text):
    """判断是否是选项行 A. B. C. D."""
    return bool(re.match(r"^[A-D][．.、]", text))


def match_score(doc, cached_texts=None):
    return pattern_match_score(doc, TEMPLATE_FEATURES["patterns"], cached_texts)


def set_font_format(doc):
    set_standard_font(doc)


def clean_document(doc):
    """
    清洗文档（WPS COM 对象）
    
    Args:
        doc: WPS 文档对象
    
    Returns:
        bool: 是否成功
    """
    print(f"   ▶ 使用模板B清洗: {doc.Name}")
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
    questions_list = []   # 按顺序存储，支持重复题号（多层试卷）
    current_question = None
    
    print(f"\n   🔍 开始解析文档内容...\n")
    
    i = 1
    while i <= paras.Count:
        p = paras(i)
        text = p.Range.Text.replace("\r", "").replace("\n", "").strip()
        
        if not text:
            i += 1
            continue
        
        debug_text = text[:60] + "..." if len(text) > 60 else text
        
        # 垃圾行过滤
        if is_garbage_line(text):
            print(f"   [行{i:3d}] 🚫 跳过垃圾行: {debug_text}")
            i += 1
            continue
        
        # 新题目（数字开头）
        q_match = QUESTION_PATTERN.match(text)
        if q_match:
            q_num = int(q_match.group(1))
            current_question = {
                'original_num': q_num,
                'type': None,          # 'choice' 或 'big'
                'answer_lines': [],    # 选择题: ['A']；大题: ['(1)xxx', '(2)xxx', ...]
                'analysis_lines': [],  # 【分析】内容行列表（不含标记本身）
                'detail_lines': [],    # 【详解】内容行列表（第一行含重建的【详解】标记）
            }
            questions_list.append(current_question)
            print(f"   [行{i:3d}] 📌 发现题目: 第{q_num}题 (顺序{len(questions_list)})")
            i += 1
            continue
        
        # 选项行 A. B. C. D. （题干选项，跳过）
        if is_option_line(text):
            print(f"   [行{i:3d}] ⏭️  跳过选项: {debug_text}")
            i += 1
            continue
        
        # 大题小问题干 (1)(2)(3) - 仅在【答案】收到前跳过
        if SUB_QUESTION_PATTERN.match(text) and current_question and current_question['type'] is None:
            print(f"   [行{i:3d}] ⏭️  跳过小问: {debug_text}")
            i += 1
            continue
        
        # ── 选择题答案 【答案】A ──────────────────────────────────────
        answer_match = CHOICE_ANSWER_PATTERN.match(text)
        if answer_match and current_question:
            current_question['type'] = 'choice'
            current_question['answer_lines'] = [answer_match.group(1)]
            print(f"   [行{i:3d}] ✅ 选择题答案: {answer_match.group(1)}")
            i += 1
            continue
        
        # ── 大题答案 【答案】(1)xxx ────────────────────────────────────
        big_answer_match = BIG_QUESTION_ANSWER.match(text)
        if big_answer_match and current_question:
            answer_text = big_answer_match.group(1)
            if re.search(r'\(\d+\)', answer_text):
                # 含 (n) 编号 → 大题，用内部循环收集后续 (n) 行
                current_question['type'] = 'big'
                answer_lines = [answer_text]
                print(f"   [行{i:3d}] 📋 大题答案: {answer_text[:40]}")
                i += 1
                while i <= paras.Count:
                    next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                    if not next_text:
                        i += 1
                        continue
                    if re.match(r'^\(\d+\)', next_text):
                        answer_lines.append(next_text)
                        print(f"   [行{i:3d}]    └─ 追加小问答案: {next_text[:40]}")
                        i += 1
                    else:
                        break   # 遇到非 (n) 行停止（指向【分析】或【详解】）
                current_question['answer_lines'] = answer_lines
                continue        # i 已在内循环更新
            else:
                # 单行大题答案（无小问编号）
                current_question['type'] = 'big'
                current_question['answer_lines'] = [answer_text]
                print(f"   [行{i:3d}] 📋 大题答案(单行): {answer_text[:40]}")
                i += 1
                continue
        
        # ── 【分析】收集 ──────────────────────────────────────────────
        analysis_match = ANALYSIS_PATTERN.match(text)
        if analysis_match and current_question:
            first_line = analysis_match.group(1)
            analysis_lines = [first_line] if first_line else []
            print(f"   [行{i:3d}] 🔍 发现【分析】: {first_line[:40] if first_line else '(空)'}")
            i += 1
            while i <= paras.Count:
                next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not next_text:
                    i += 1
                    continue
                # 遇到【详解】、【答案】 → 停止（不停在 \d+[．.] 编号行，防止分析内容的编号列表被误判为题目）
                if re.match(r'^【详解】|^【答案】', next_text):
                    break
                analysis_lines.append(next_text)
                i += 1
            current_question['analysis_lines'] = analysis_lines
            print(f"   [行{i:3d}]    └─ 分析共 {len(analysis_lines)} 行")
            continue
        
        # ── 【详解】收集 ──────────────────────────────────────────────
        detail_match = DETAIL_PATTERN.match(text)
        if detail_match and current_question:
            first_line = detail_match.group(1)
            # 第一行重建标记（【详解】 + 内容）
            detail_lines = [f"【详解】{first_line}"] if first_line else ["【详解】"]
            print(f"   [行{i:3d}] 🔍 发现【详解】: {first_line[:40] if first_line else '(空)'}")
            i += 1
            while i <= paras.Count:
                next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not next_text:
                    i += 1
                    continue
                # 遇到故选X。→ 包含并停止
                if ANSWER_END_PATTERN.search(next_text):
                    detail_lines.append(next_text)
                    i += 1
                    break
                # 遇到新题目编号、下一题【答案】、【分析】 → 停止
                # 注意：不停在 A./B./C./D. 行（那是选项解析，需要保留）
                if re.match(r'^\d+[．.]|^【答案】|^【分析】', next_text):
                    break
                detail_lines.append(next_text)
                i += 1
            current_question['detail_lines'] = detail_lines
            print(f"   [行{i:3d}]    └─ 详解共 {len(detail_lines)} 行")
            continue
        
        # 未识别行
        print(f"   [行{i:3d}] ⚪ 未识别: {debug_text}")
        i += 1
    
    # 3. 统计
    print(f"\n   📊 解析完成统计:")
    print(f"      - 共发现 {len(questions_list)} 道题目")
    choice_found = sum(1 for q in questions_list if q['type'] == 'choice')
    big_found = sum(1 for q in questions_list if q['type'] == 'big')
    untyped_found = sum(1 for q in questions_list if q['type'] is None)
    print(f"      - 选择题: {choice_found} 道，大题: {big_found} 道")
    if untyped_found > 0:
        print(f"      - ⚠️  未识别类型: {untyped_found} 道")
    
    # 4. 清空文档并重新写入（顺序编号）
    doc.Content.Text = ""
    
    choice_count = 0
    big_count = 0
    
    for idx, q_data in enumerate(questions_list, 1):
        q_type = q_data['type']
        answer_lines = q_data['answer_lines']
        analysis_lines = q_data['analysis_lines']
        detail_lines = q_data['detail_lines']
        
        if q_type == 'choice':
            choice_count += 1
            letter = answer_lines[0] if answer_lines else ''
            doc.Content.InsertAfter(f"{idx}．{letter}\r")
            # 解析：【分析】... → 第一行加 "解析：" 前缀
            if analysis_lines:
                doc.Content.InsertAfter(f"解析：【分析】{analysis_lines[0]}\r")
                for line in analysis_lines[1:]:
                    doc.Content.InsertAfter(f"{line}\r")
            # 【详解】逐行输出（第一行已含【详解】标记，后续含A./B./C./D.和故选X。）
            for line in detail_lines:
                doc.Content.InsertAfter(f"{line}\r")
        
        elif q_type == 'big':
            big_count += 1
            doc.Content.InsertAfter(f"{idx}．\r")
            for ans_line in answer_lines:
                doc.Content.InsertAfter(f"{ans_line}\r")
            if analysis_lines:
                doc.Content.InsertAfter(f"解析：【分析】{analysis_lines[0]}\r")
                for line in analysis_lines[1:]:
                    doc.Content.InsertAfter(f"{line}\r")
            for line in detail_lines:
                doc.Content.InsertAfter(f"{line}\r")
    
    print(f"   ✓ 共处理 {choice_count} 道选择题, {big_count} 道大题")
    
    # 5. 设置字体格式（与模板A一致）
    set_font_format(doc)
    
    return True


# 模块导出
TEMPLATE_INFO = TEMPLATE_FEATURES
