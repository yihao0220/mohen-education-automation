# coding: utf-8
"""
格式模板A - 虎鹰-四年级-语文格式

适用文档特征：
- 章节标题: `1 古诗词三首`、`阅读素养练习(一)`、`语文园地`
- 选择题答案: `2. D`、`3. A`、`7. (1)B`（题号后直接跟答案字母）
- 大题答案: `1. 蜻蜓 蝶 稀 shū...`（题号后空格跟答案内容）
- 小问格式: `(1)xxx`、`(2)xxx`，可能多小问同行如 `(1)B (2)D (3)C`
- 无【解析】标记，解析内容直接跟在答案后或作为独立段落
- 答案续行：答案可能跨多行（如第6题答案分多行）

转换规则：
- 删除：章节标题、练习标题、空行
- 选择题：`2. D` → `题号．D`
- 大题：`1. 答案内容` → `题号．` + `答案内容`
- 小问：保留 `(1)(2)(3)` 格式
"""

import re
import os

from . import matches_garbage_pattern, pattern_match_score, set_standard_font

# 格式识别特征
TEMPLATE_FEATURES = {
    "name": "虎鹰-四年级-语文",
    "patterns": [
        r"^\d+\.\s*[A-D]\b",           # 选择题答案: 2. D / 3. A
        r"^\d+\.\s*[(（]\d+[)）]",     # 题号+小问: 7. (1)B
        r"^\d+\s+[^0-9]",              # 章节标题: 1 古诗词三首
        r"^阅读素养练习",              # 阅读素养练习(一)
        r"^单元素养练习",              # 单元素养练习(一)
        r"^语文园地",                  # 语文园地
    ],
    "match_threshold": 0.03,
}

# 需要删除的垃圾行模式
GARBAGE_PATTERNS = [
    r"^\s*$",                           # 空行
]

# 章节标题模式（识别为分隔标记，不占用题号）
SECTION_PATTERNS = [
    r"^\d+\s+[^0-9（(].*",             # 章节标题: 1 古诗词三首、2 乡下人家
    r"^\d+\*\s+.*",                   # 带星号的课文: 4* 三月桃花水、8* 千年梦圆在今朝
    r"^阅读素养练习[（(].*[)）]",      # 阅读素养练习(一)
    r"^单元素养练习[（(].*[)）]",      # 单元素养练习(一)
    r"^语文园地",                      # 语文园地
    r"^天窗$",                         # 纯文字课文标题（特定关键词）
]

# 题号模式：支持 1. 、2. 等
QUESTION_PATTERN = re.compile(r"^(\d+)\.\s*(.+)")

# 纯题号+答案字母（选择题）
CHOICE_ANSWER_PATTERN = re.compile(r"^(\d+)\.\s*([A-D])\b\s*(.*)")

# 题号+小问格式（如 7. (1)B）
QUESTION_WITH_SUB_PATTERN = re.compile(r"^(\d+)\.\s*[(（](\d+)[)）]\s*([A-D]?)\s*(.*)")

# 小问编号模式: (1) (2) (3) 或 （1）（2）（3）
SUB_QUESTION_PATTERN = re.compile(r"^[(（](\d+)[)）]\s*(.*)")

# 大题标题模式: 一、二、三、等中文数字标题
BIG_QUESTION_TITLE_PATTERN = re.compile(r"^([一二三四五六七八九十]+)、\s*(.*)")

# 多小问同行模式：提取所有 (1)xxx (2)xxx
MULTI_SUB_PATTERN = re.compile(r"[(（](\d+)[)）]([^（(]*?)(?=[(（]\d+[)）]|$)")


def is_garbage_line(text):
    return matches_garbage_pattern(text, GARBAGE_PATTERNS)


def is_section_title(text):
    """判断是否是章节标题"""
    for pattern in SECTION_PATTERNS:
        if re.match(pattern, text):
            return True
    return False


def match_score(doc, cached_texts=None):
    return pattern_match_score(doc, TEMPLATE_FEATURES["patterns"], cached_texts)


def set_font_format(doc):
    set_standard_font(doc)


def scan_sections(doc):
    """
    扫描文档中的所有章节标题

    Args:
        doc: WPS 文档对象

    Returns:
        list: 章节列表 [(行号, 章节标题), ...]
    """
    sections = []
    paras = doc.Paragraphs

    for i in range(1, paras.Count + 1):
        text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
        if text and is_section_title(text):
            sections.append((i, text))

    return sections


def select_sections(sections):
    """
    交互式选择要清洗的章节

    Args:
        sections: 章节列表 [(行号, 章节标题), ...]

    Returns:
        list: 选中的章节索引列表，如果全选返回 None
    """
    if not sections:
        print("   ⚠️ 未检测到章节标题，将清洗全部内容")
        return None

    print(f"\n   📑 检测到 {len(sections)} 个章节：\n")
    for idx, (line_num, title) in enumerate(sections, 1):
        print(f"   [{idx}] {title}")

    print(f"\n   [0] 清洗全部章节")
    print(f"   [a] 全选所有章节")

    while True:
        choice = input("\n   👉 请选择要清洗的章节（多选用逗号分隔，如 1,3,5）：").strip()

        if choice == '0':
            return None  # 清洗全部

        if choice.lower() == 'a':
            return list(range(len(sections)))  # 全选

        try:
            # 解析选择（支持 1,3,5 或 1-3 格式）
            selected = []
            for part in choice.split(','):
                part = part.strip()
                if '-' in part:
                    # 范围选择 1-3
                    start, end = part.split('-')
                    selected.extend(range(int(start) - 1, int(end)))
                else:
                    selected.append(int(part) - 1)

            # 验证选择范围
            if all(0 <= idx < len(sections) for idx in selected):
                return selected
            else:
                print(f"   ❌ 选择超出范围，请输入 1-{len(sections)} 之间的数字")
        except ValueError:
            print(f"   ❌ 输入格式错误，请重新输入")


def clean_document(doc, selected_sections=None, return_sections_info=False):
    """
    清洗文档（WPS COM 对象）

    Args:
        doc: WPS 文档对象
        selected_sections: 选中的章节索引列表，None 表示清洗全部
        return_sections_info: 是否返回章节信息（用于多章节分别保存）

    Returns:
        bool 或 dict: 是否成功，或返回章节信息字典
    """
    print(f"   ▶ 使用模板A清洗: {doc.Name}")
    print(f"   📊 文档共 {doc.Paragraphs.Count} 个段落")

    # 0. 扫描并选择章节
    sections = scan_sections(doc)
    if sections and selected_sections is None:
        selected_sections = select_sections(sections)

    # 确定清洗范围（行号范围列表）
    ranges = []
    if selected_sections is None:
        # 清洗全部
        ranges = [(1, doc.Paragraphs.Count)]
        print(f"\n   📝 清洗模式：全部章节")
    else:
        # 清洗选中章节
        for idx in selected_sections:
            start_line = sections[idx][0]
            # 结束行是下一个章节的开始，或者是文档末尾
            if idx + 1 < len(sections):
                end_line = sections[idx + 1][0] - 1
            else:
                end_line = doc.Paragraphs.Count
            ranges.append((start_line, end_line))
        print(f"\n   📝 清洗模式：选中 {len(selected_sections)} 个章节")
    
    # 如果需要返回章节信息（多章节分别保存模式）
    if return_sections_info and selected_sections and len(selected_sections) > 0:
        return {
            'sections': sections,
            'selected_sections': selected_sections,
            'ranges': ranges
        }

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

    # 2. 第一遍遍历：收集所有有效内容（只在选定范围内）
    paras = doc.Paragraphs
    questions_list = []    # 按扫描顺序存储
    current_question = None

    print(f"\n   🔍 开始解析文档内容...\n")

    # 根据范围列表处理
    for start_line, end_line in ranges:
        i = start_line
        while i <= end_line and i <= paras.Count:
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

            # 章节标题保留（作为分隔标记）
            if is_section_title(text):
                current_question = {
                    'original_num': 0,
                    'type': 'section',
                    'answer_lines': [text],
                    'analysis_lines': [],
                }
                questions_list.append(current_question)
                print(f"   [行{i:3d}] 📑 保留章节标题: {debug_text}")
                i += 1
                continue

            # ── 选择题格式：7. (1)B 或 2. D ─────────────────────────
            # 先检查是否包含小问编号的选择题（答案必须是字母A-D）
            sub_choice_match = QUESTION_WITH_SUB_PATTERN.match(text)
            if sub_choice_match:
                q_num = int(sub_choice_match.group(1))
                sub_num = sub_choice_match.group(2)
                answer = sub_choice_match.group(3)
                remaining = sub_choice_match.group(4).strip()

                # 如果答案部分是字母（A-D），则是选择题；否则是大题
                if answer and re.match(r'^[A-D]$', answer):
                    current_question = {
                        'original_num': q_num,
                        'type': 'choice',
                        'answer_lines': [f"({sub_num}){answer}"],
                        'analysis_lines': [],
                    }
                    questions_list.append(current_question)
                    print(f"   [行{i:3d}] 📌 发现选择题(带小问): 第{q_num}题 ({sub_num})={answer}")

                    # 检查同行是否还有更多小问，如 (2)D (3)C
                    if remaining:
                        more_subs = MULTI_SUB_PATTERN.findall(remaining)
                        for m_sub_num, m_content in more_subs:
                            m_content = m_content.strip()
                            if m_content:
                                current_question['answer_lines'].append(f"({m_sub_num}){m_content}")
                                print(f"   [行{i:3d}]    └─ 小问({m_sub_num}): {m_content[:30]}")
                    i += 1

                    # 收集后续小问（如 (2)xxx 独立成行的情况）
                    while i <= end_line and i <= paras.Count:
                        next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                        if not next_text:
                            i += 1
                            continue

                        # 检查是否是新的小问（纯小问格式，如 (2)duàn）
                        next_sub = SUB_QUESTION_PATTERN.match(next_text)
                        if next_sub:
                            sub_num = next_sub.group(1)
                            sub_content = next_sub.group(2).strip()
                            current_question['answer_lines'].append(f"({sub_num}){sub_content}")
                            print(f"   [行{i:3d}]    └─ 小问({sub_num}): {sub_content[:40]}")
                            i += 1
                            continue

                        # 遇到新题号或其他内容，停止收集
                        if QUESTION_PATTERN.match(next_text):
                            break
                        if CHOICE_ANSWER_PATTERN.match(next_text):
                            break
                        if is_garbage_line(next_text):
                            break

                        # 否则是上一小问的续行
                        if current_question['answer_lines']:
                            current_question['answer_lines'][-1] += next_text
                            print(f"   [行{i:3d}]       └─ 追加续行: {next_text[:40]}")
                        i += 1
                    continue
                else:
                    # 大题带小问格式，如 "6. (1)锄豆..."
                    current_question = {
                        'original_num': q_num,
                        'type': 'big',
                        'answer_lines': [],
                        'analysis_lines': [],
                    }
                    questions_list.append(current_question)
                    print(f"   [行{i:3d}] 📌 发现大题: 第{q_num}题")

                    # 处理第一个小问
                    sub_content = remaining if remaining else ""
                    if sub_content:
                        answer_line = f"({sub_num}){sub_content}"
                        current_question['answer_lines'].append(answer_line)
                        print(f"   [行{i:3d}]    └─ 小问({sub_num}): {sub_content[:40]}")
                    i += 1

                    # 收集后续小问或续行
                    while i <= end_line and i <= paras.Count:
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

                        # 检查是否是新题目（题号+小问格式，如 "4.(1)..."）
                        new_q_with_sub = QUESTION_WITH_SUB_PATTERN.match(next_text)
                        if new_q_with_sub:
                            # 这是新的大题，不是续行
                            break

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

            # 纯选择题格式：2. D
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
                print(f"   [行{i:3d}] 📌 发现选择题: 第{q_num}题 = {answer}")

                # 检查是否有同行多答案，如 "1. B 2. C"
                if remaining:
                    # 可能是多题同行，尝试匹配更多选择题
                    more_choices = re.findall(r"(\d+)\.\s*([A-D])\b", remaining)
                    for m_num, m_answer in more_choices:
                        extra_question = {
                            'original_num': int(m_num),
                            'type': 'choice',
                            'answer_lines': [m_answer],
                            'analysis_lines': [],
                        }
                        questions_list.append(extra_question)
                        print(f"   [行{i:3d}]    └─ 同行发现: 第{m_num}题 = {m_answer}")
                i += 1
                continue

            # ── 大题格式：题号 + 内容 ───────────────────────────────
            q_match = QUESTION_PATTERN.match(text)
            if q_match:
                q_num = int(q_match.group(1))
                content = q_match.group(2).strip()

                # 检查是否包含小问编号 (1)(2)...
                sub_match = SUB_QUESTION_PATTERN.match(content)
                if sub_match:
                    # 大题格式：4.(1)xxx 或 6.(1)xxx
                    current_question = {
                        'original_num': q_num,
                        'type': 'big',
                        'answer_lines': [],
                        'analysis_lines': [],
                    }
                    questions_list.append(current_question)
                    print(f"   [行{i:3d}] 📌 发现大题: 第{q_num}题")

                    # 处理可能包含多个小问的内容
                    sub_questions = MULTI_SUB_PATTERN.findall(content)
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
                    while i <= end_line and i <= paras.Count:
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
                    # 大题直接答案（无小问编号）：1. 蜻蜓 蝶...
                    current_question = {
                        'original_num': q_num,
                        'type': 'big',
                        'answer_lines': [content],
                        'analysis_lines': [],
                    }
                    questions_list.append(current_question)
                    print(f"   [行{i:3d}] 📌 发现大题(无小问): 第{q_num}题")
                    print(f"   [行{i:3d}]    └─ 答案: {content[:40]}")
                    i += 1

                    # 收集续行
                    while i <= end_line and i <= paras.Count:
                        next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                        if not next_text:
                            i += 1
                            continue
                        if is_garbage_line(next_text) or CHOICE_ANSWER_PATTERN.match(next_text):
                            break
                        if QUESTION_PATTERN.match(next_text):
                            break
                        # 检查是否是小问开始（如果是，属于当前题的答案）
                        if SUB_QUESTION_PATTERN.match(next_text):
                            sub_match = SUB_QUESTION_PATTERN.match(next_text)
                            sub_num = sub_match.group(1)
                            sub_content = sub_match.group(2).strip()
                            current_question['answer_lines'].append(f"({sub_num}){sub_content}")
                            print(f"   [行{i:3d}]    └─ 发现小问({sub_num}): {sub_content[:40]}")
                            i += 1
                            continue
                        # 否则是续行
                        current_question['answer_lines'][-1] += next_text
                        print(f"   [行{i:3d}]       └─ 追加续行: {next_text[:40]}")
                        i += 1
                    continue

            # ── 独立小问格式（无题号）────────────────────────────────
            sub_only_match = SUB_QUESTION_PATTERN.match(text)
            if sub_only_match and current_question and current_question['type'] == 'big':
                sub_num = sub_only_match.group(1)
                sub_content = sub_only_match.group(2).strip()
                current_question['answer_lines'].append(f"({sub_num}){sub_content}")
                print(f"   [行{i:3d}] 📌 小问续行({sub_num}): {sub_content[:40]}")
                i += 1

                # 收集此小问的续行
                while i <= end_line and i <= paras.Count:
                    next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                    if not next_text:
                        i += 1
                        continue
                    # 遇到新小问、新题或垃圾行则停止
                    if SUB_QUESTION_PATTERN.match(next_text):
                        break
                    if is_garbage_line(next_text):
                        break
                    if CHOICE_ANSWER_PATTERN.match(next_text):
                        break
                    if QUESTION_PATTERN.match(next_text):
                        break
                    # 追加续行
                    current_question['answer_lines'][-1] += next_text
                    print(f"   [行{i:3d}]       └─ 追加续行: {next_text[:40]}")
                    i += 1
                continue

            # ── 大题标题格式（一、二、三、等）────────────────────────
            big_title_match = BIG_QUESTION_TITLE_PATTERN.match(text)
            if big_title_match and current_question and current_question['type'] == 'big':
                title_num = big_title_match.group(1)
                title_content = big_title_match.group(2).strip()
                # 将大题标题作为答案的一部分保留
                if title_content:
                    current_question['answer_lines'].append(f"{title_num}、{title_content}")
                else:
                    current_question['answer_lines'].append(f"{title_num}、")
                print(f"   [行{i:3d}] 📌 大题标题({title_num}): {title_content[:40]}")
                i += 1

                # 收集此大题标题下的续行
                while i <= end_line and i <= paras.Count:
                    next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
                    if not next_text:
                        i += 1
                        continue
                    # 遇到新的大题标题、新小问、新题或垃圾行则停止
                    if BIG_QUESTION_TITLE_PATTERN.match(next_text):
                        break
                    if SUB_QUESTION_PATTERN.match(next_text):
                        break
                    if is_garbage_line(next_text):
                        break
                    if CHOICE_ANSWER_PATTERN.match(next_text):
                        break
                    if QUESTION_PATTERN.match(next_text):
                        break
                    # 追加续行到当前大题标题
                    current_question['answer_lines'][-1] += next_text
                    print(f"   [行{i:3d}]       └─ 追加续行: {next_text[:40]}")
                    i += 1
                continue

            # ── 未识别行 ─────────────────────────────────────────────
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

        # 章节标题不占用题号
        if q_type == 'section':
            section_count += 1
            for line in answer_lines:
                doc.Content.InsertAfter(f"{line}\r")
            continue

        output_idx += 1

        if q_type == 'choice':
            choice_count += 1
            # 选择题格式：题号．答案（可能包含多个小问如 (1)B(2)D）
            answer_text = "".join(answer_lines)
            doc.Content.InsertAfter(f"{output_idx}．{answer_text}\r")
            # 无解析时添加空解析行，方便后续录入
            doc.Content.InsertAfter(f"解析： \r")

        elif q_type == 'big':
            big_count += 1
            doc.Content.InsertAfter(f"{output_idx}．\r")
            # 大题答案直接输出，每行一个
            for line in answer_lines:
                doc.Content.InsertAfter(f"{line}\r")
            # 无解析时添加空解析行
            doc.Content.InsertAfter(f"解析： \r")

    print(f"   ✓ 共处理 {choice_count} 道选择题, {big_count} 道大题, {section_count} 个章节标题，输出共 {output_idx} 题")

    # 5. 设置字体格式
    set_font_format(doc)

    return True


# 模块导出
TEMPLATE_INFO = TEMPLATE_FEATURES
