# coding: utf-8
"""
格式模板D - 布心中学物理格式

适用文档特征：
- 选择题答案格式: `1. A` 或 `1.C`（题号后直接跟答案字母，无【答案】标记）
- 选择题解析格式: 行内 `2. A 【解析】解析内容...`（解析紧跟答案）
- 大题答案格式: `5.(1)水平 调节平衡...` 或 `6. 地球`（无【答案】标记）
- 大题小问: `(1)` `(2)` 格式，答案可能跨多行
- 独立解析块: 部分大题有独立的 `【解析】` 行作为解析内容
- 章节标题: `第一讲 声现象`、`基础过关`、`能力提升`、`素养拔尖`
- 子课时: `第 1 课时 质量 密度`
"""

import re
import os

from . import matches_garbage_pattern, pattern_match_score, set_standard_font

# 格式识别特征
TEMPLATE_FEATURES = {
    "name": "布心中学 - 物理",
    "patterns": [
        r"^\d+\.\s*[A-D]\s*【解析】",  # 选择题行内解析: 1. A 【解析】...
        r"^\d+\.\s*[A-D](?!\w)",      # 选择题答案: 1. A / 1.C（后面不接字母）
        r"^【解析】",                   # 独立解析块
        r"^第[一二三四五六七八九十\d]+讲",  # 第X讲 章节标题
        r"^第\s*\d+\s*课时",           # 第 X 课时 子课时
        r"基础过关|能力提升|素养拔尖",  # 布心中学特有的小节标题
    ],
    "match_threshold": 0.03,
}

# 需要删除的垃圾行模式
# 注意：章节标题（第X讲、第X课时）保留以便检查，不作为垃圾行
GARBAGE_PATTERNS = [
    r"^课时作业$",                      # 课时作业
    r"^答:略。?$",                       # 答:略。
    r"^答：略。?$",                       # 答：略。
    r"^\s*$",                           # 空行
]

# 章节标题模式（保留用于检查，但识别为分隔标记）
SECTION_PATTERNS = [
    r"^第[一二三四五六七八九十\d]+讲",   # 第一讲、第二讲 等章节标题
    r"^第\s*\d+\s*课时",               # 第 1 课时、第 2 课时
    r"^基础过关$",                      # 基础过关
    r"^能力提升$",                      # 能力提升
    r"^素养拔尖$",                      # 素养拔尖
]

# 题号模式：支持半角点号（布心中学使用 1. 而非 1．）
# 注意：需要区分 5.(1) 这种大题小问格式和普通的 1. 题号
QUESTION_PATTERN = re.compile(r"^(\d+)\.\s*(.+)")

# 纯题号+答案格式（如 "1. A" 或 "1.C" 或 "1. D 【解析】..."）
CHOICE_ANSWER_PATTERN = re.compile(r"^(\d+)\.\s*([A-D])\b\s*(.*)")

# 大题小问编号模式: (1) (2) (3) 或 （1） （2） （3）等
SUB_QUESTION_PATTERN = re.compile(r"^[(（](\d+)[)）]\s*(.*)")

# 解析标记（独立行格式）
ANALYSIS_PATTERN = re.compile(r"^【解析】\s*(.*)$")

# 行内解析标记（用于分割答案和解析）
INLINE_ANALYSIS_PATTERN = re.compile(r"【解析】\s*(.*)$")


def is_garbage_line(text):
    return matches_garbage_pattern(text, GARBAGE_PATTERNS)


def is_section_title(text):
    """判断是否是章节标题（保留用于检查）"""
    for pattern in SECTION_PATTERNS:
        if re.match(pattern, text):
            return True
    return False


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
    print(f"   ▶ 使用模板D清洗: {doc.Name}")
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

        if not text:
            i += 1
            continue

        debug_text = text[:60] + "..." if len(text) > 60 else text

        # 垃圾行过滤
        if is_garbage_line(text):
            print(f"   [行{i:3d}] 🚫 跳过垃圾行: {debug_text}")
            i += 1
            continue

        # 章节标题保留（作为分隔标记）
        if is_section_title(text):
            # 将章节标题作为特殊题目保存（类型为section）
            current_question = {
                'original_num': 0,  # 章节标题无题号
                'type': 'section',
                'answer_lines': [text],
                'analysis_lines': [],
            }
            questions_list.append(current_question)
            print(f"   [行{i:3d}] 📑 保留章节标题: {debug_text}")
            i += 1
            continue

        # ── 选择题格式：1. A 【解析】... 或 1.C ─────────────────────────
        choice_match = CHOICE_ANSWER_PATTERN.match(text)
        if choice_match:
            q_num = int(choice_match.group(1))
            answer = choice_match.group(2)
            remaining = choice_match.group(3).strip()

            current_question = {
                'original_num': q_num,
                'type': 'choice',
                'answer_lines': [answer],
                'analysis_lines': [],
            }
            questions_list.append(current_question)
            print(f"   [行{i:3d}] 📌 发现选择题: 第{q_num}题 = {answer} (顺序{len(questions_list)})")

            # 检查是否有行内解析
            inline_analysis = INLINE_ANALYSIS_PATTERN.search(remaining)
            if inline_analysis:
                analysis_text = inline_analysis.group(1).strip()
                if analysis_text:
                    current_question['analysis_lines'].append(analysis_text)
                    print(f"   [行{i:3d}] 🔍 行内解析: {analysis_text[:40]}")
                    # 检查解析是否有续行（当前行解析被截断的情况）
                    if remaining.endswith('...') or remaining.endswith('…') or len(remaining) > 80:
                        # 可能需要收集续行
                        i += 1
                        while i <= paras.Count:
                            next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                            if not next_text:
                                i += 1
                                continue
                            # 如果遇到新题目或垃圾行，停止收集
                            if CHOICE_ANSWER_PATTERN.match(next_text) or QUESTION_PATTERN.match(next_text):
                                break
                            if is_garbage_line(next_text):
                                break
                            # 检查是否是解析续行（不以题号开头、不以【解析】开头）
                            if not next_text.startswith('【解析】') and not re.match(r'^\d+\.', next_text):
                                current_question['analysis_lines'][-1] += next_text
                                print(f"   [行{i:3d}]    └─ 解析续行: {next_text[:40]}")
                                i += 1
                            else:
                                break
                        continue
            i += 1
            continue

        # ── 独立【解析】行 ────────────────────────────────────────────
        analysis_match = ANALYSIS_PATTERN.match(text)
        if analysis_match and current_question:
            first_line = analysis_match.group(1)
            analysis_lines = [first_line] if first_line else []
            display_text = first_line[:40] if first_line else '(空)'
            print(f"   [行{i:3d}] 🔍 发现独立【解析】: {display_text}")
            i += 1
            # 收集后续行，直到遇到下一题或垃圾行
            while i <= paras.Count:
                next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                if not next_text:
                    i += 1
                    continue
                # 遇到新题目或垃圾行则停止
                if CHOICE_ANSWER_PATTERN.match(next_text) or is_garbage_line(next_text):
                    break
                # 如果是纯数字题号开头（如 "6. 地球"），停止
                if QUESTION_PATTERN.match(next_text) and not SUB_QUESTION_PATTERN.match(next_text):
                    # 检查是否是大题格式
                    q_match = QUESTION_PATTERN.match(next_text)
                    if q_match:
                        content = q_match.group(2)
                        # 如果内容不以小问编号开头，可能是新题目
                        if not SUB_QUESTION_PATTERN.match(content):
                            break
                analysis_lines.append(next_text)
                print(f"   [行{i:3d}]    └─ 追加解析行: {next_text[:40]}")
                i += 1
            current_question['analysis_lines'] = analysis_lines
            print(f"   [行{i:3d}]    └─ 解析共 {len(analysis_lines)} 行")
            continue

        # ── 大题格式：题号 + 小问编号 或 题号 + 直接答案 ───────────────
        q_match = QUESTION_PATTERN.match(text)
        if q_match:
            q_num = int(q_match.group(1))
            content = q_match.group(2).strip()

            # 检查是否包含小问编号 (1)(2)...
            sub_match = SUB_QUESTION_PATTERN.match(content)
            if sub_match:
                # 大题格式：5.(1)xxx
                current_question = {
                    'original_num': q_num,
                    'type': 'big',
                    'answer_lines': [],
                    'analysis_lines': [],
                }
                questions_list.append(current_question)
                print(f"   [行{i:3d}] 📌 发现大题: 第{q_num}题 (顺序{len(questions_list)})")

                # 处理可能包含多个小问的内容，如 "(1)振动 空气 音色 (2)510 340"
                # 使用正则拆分多个小问
                sub_questions = re.findall(r'[(（](\d+)[)）]([^（(]*?)(?=[(（]\d+[)）]|$)', content)
                if sub_questions:
                    for sub_num, sub_content in sub_questions:
                        sub_content = sub_content.strip()
                        if sub_content:
                            answer_line = f"({sub_num}){sub_content}"
                            current_question['answer_lines'].append(answer_line)
                            print(f"   [行{i:3d}]    └─ 小问({sub_num}): {sub_content[:40]}")
                else:
                    # 回退到原来的单小问处理
                    sub_num = sub_match.group(1)
                    sub_content = sub_match.group(2).strip()
                    answer_line = f"({sub_num}){sub_content}"
                    current_question['answer_lines'].append(answer_line)
                    print(f"   [行{i:3d}]    └─ 小问({sub_num}): {sub_content[:40]}")
                i += 1

                # 收集后续小问或续行
                while i <= paras.Count:
                    next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                    if not next_text:
                        i += 1
                        continue

                    # 检查是否是新的小问
                    next_sub = SUB_QUESTION_PATTERN.match(next_text)
                    if next_sub:
                        sub_num = next_sub.group(1)
                        sub_content = next_sub.group(2).strip()
                        answer_line = f"({sub_num}){sub_content}"
                        current_question['answer_lines'].append(answer_line)
                        print(f"   [行{i:3d}]    └─ 小问({sub_num}): {sub_content[:40]}")
                        i += 1
                        continue

                    # 检查是否是垃圾行或新题目
                    if is_garbage_line(next_text):
                        break
                    if CHOICE_ANSWER_PATTERN.match(next_text):
                        break
                    if QUESTION_PATTERN.match(next_text):
                        # 检查是否是纯新题（不是小问续行）
                        qm = QUESTION_PATTERN.match(next_text)
                        if qm:
                            next_content = qm.group(2)
                            if not SUB_QUESTION_PATTERN.match(next_content):
                                break

                    # 否则是上一小问的续行
                    if current_question['answer_lines']:
                        current_question['answer_lines'][-1] += next_text
                        print(f"   [行{i:3d}]       └─ 追加续行: {next_text[:40]}")
                    i += 1
                continue
            else:
                # 可能是大题直接答案（无小问编号）：6. 地球
                current_question = {
                    'original_num': q_num,
                    'type': 'big',
                    'answer_lines': [content],
                    'analysis_lines': [],
                }
                questions_list.append(current_question)
                print(f"   [行{i:3d}] 📌 发现大题(无小问): 第{q_num}题 (顺序{len(questions_list)})")
                print(f"   [行{i:3d}]    └─ 答案: {content[:40]}")
                i += 1

                # 收集续行
                while i <= paras.Count:
                    next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                    if not next_text:
                        i += 1
                        continue
                    if is_garbage_line(next_text) or CHOICE_ANSWER_PATTERN.match(next_text):
                        break
                    if QUESTION_PATTERN.match(next_text):
                        break
                    if ANALYSIS_PATTERN.match(next_text):
                        break
                    current_question['answer_lines'][-1] += next_text
                    print(f"   [行{i:3d}]       └─ 追加续行: {next_text[:40]}")
                    i += 1
                continue

        # ── 未识别行 ─────────────────────────────────────────────────
        print(f"   [行{i:3d}] ⚪ 未识别: {debug_text}")
        i += 1

    # 3. 统计
    print(f"\n   📊 解析完成统计:")
    print(f"      - 共发现 {len(questions_list)} 道题目（含重复/错序）")
    choice_found = sum(1 for q in questions_list if q['type'] == 'choice')
    big_found = sum(1 for q in questions_list if q['type'] == 'big')
    section_found = sum(1 for q in questions_list if q['type'] == 'section')
    print(f"      - 选择题: {choice_found} 道，大题: {big_found} 道，章节标题: {section_found} 个")

    # 4. 清空文档并重新写入
    doc.Content.Text = ""

    choice_count = 0
    big_count = 0
    section_count = 0
    output_idx = 0

    for q_data in questions_list:
        q_type = q_data['type']
        answer_lines = q_data['answer_lines']
        analysis_lines = q_data['analysis_lines']

        # 章节标题不占用题号
        if q_type == 'section':
            section_count += 1
            # 章节标题原样输出，不加题号
            for line in answer_lines:
                doc.Content.InsertAfter(f"{line}\r")
            continue

        output_idx += 1

        if q_type == 'choice':
            choice_count += 1
            letter = answer_lines[0] if answer_lines else ''
            if analysis_lines:
                # 有解析：答案和解析在同一行
                doc.Content.InsertAfter(f"{output_idx}．{letter}　解析：{analysis_lines[0]}\r")
                for line in analysis_lines[1:]:
                    doc.Content.InsertAfter(f"{line}\r")
            else:
                # 无解析：只有答案行
                doc.Content.InsertAfter(f"{output_idx}．{letter}　\r")

        elif q_type == 'big':
            big_count += 1
            doc.Content.InsertAfter(f"{output_idx}．\r")
            # 大题答案前加 "答案：" 标记
            if answer_lines:
                doc.Content.InsertAfter(f"答案：{answer_lines[0]}\r")
                for line in answer_lines[1:]:
                    doc.Content.InsertAfter(f"{line}\r")
            if analysis_lines:
                doc.Content.InsertAfter(f"解析：{analysis_lines[0]}\r")
                for line in analysis_lines[1:]:
                    doc.Content.InsertAfter(f"{line}\r")
            else:
                doc.Content.InsertAfter(f"解析： \r")

    print(f"   ✓ 共处理 {choice_count} 道选择题, {big_count} 道大题, {section_count} 个章节标题，输出共 {output_idx} 题")

    # 5. 设置字体格式
    set_font_format(doc)

    return True


# 模块导出
TEMPLATE_INFO = TEMPLATE_FEATURES
