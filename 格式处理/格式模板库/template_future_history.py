# coding: utf-8
"""
格式模板 - 未来高二历史选择性必修教师版答案

适用文档特征：
- 教师版文档前半部分包含题目，答案集中在后半部分
- 作业单后常见 `【答案】B` + `【解析】...`
- 部分课次使用 `参考答案` 后的 `1．D` + `【详解】...`
- 部分课次没有 `【答案】`，直接在题干后给 `1．D` 与详解
"""

import re


TEMPLATE_FEATURES = {
    "name": "未来高二 - 历史选择性必修教师版答案",
    "patterns": [
        r"第\d+课",
        r"【作业单】",
        r"参考答案",
        r"【答案】",
        r"【解析】",
        r"【详解】",
        r"高考真题",
    ],
    "match_threshold": 0.04,
}


ANSWER_MARKER_PATTERN = re.compile(r"^(?:【答案】|答案[：:])\s*(.*)$")
ANALYSIS_MARKER_PATTERN = re.compile(r"^(?:【(?:解析|详解|分析)】|(?:解析|详解|分析)[：:])\s*(.*)$")
INLINE_ANALYSIS_PATTERN = re.compile(r"^(.*?)(?:【(?:解析|详解|分析)】)\s*(.*)$")
QUESTION_PREFIX_PATTERN = re.compile(r"^(\d+)[．.、]\s*(.*)$")
CHOICE_ANSWER_PATTERN = re.compile(r"^[A-D]+$")
SUBANSWER_PREFIX_PATTERN = re.compile(r"^[（(]\s*1\s*[）)]|^①")
REFERENCE_ANSWER_PATTERN = re.compile(r"参考答案|答案解析|答案与解析")
HOMEWORK_PATTERN = re.compile(r"【作业单】")
LAYER_HEADING_PATTERN = re.compile(r"^【(?:基础巩固|能力提升|拓展延伸)】$")

HISTORY_KEYWORDS = (
    "历史",
    "政治制度",
    "民族关系",
    "对外交往",
    "民族政策",
    "法治",
    "外交",
    "文官制度",
    "货币",
    "赋税",
    "户籍",
    "基层治理",
    "社会保障",
)

CIRCLED_POINT_LABELS = {
    "①": "要点一：",
    "②": "要点二：",
    "③": "要点三：",
    "④": "要点四：",
    "⑤": "要点五：",
    "⑥": "要点六：",
    "⑦": "要点七：",
    "⑧": "要点八：",
    "⑨": "要点九：",
    "⑩": "要点十：",
}


def _clean_text(text):
    cleaned = (
        (text or "")
        .replace("\r", "")
        .replace("\n", "")
        .replace("\u00a0", " ")
        .strip()
    )
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", cleaned)


def _normalize_paragraph_texts(paragraph_texts):
    return [_clean_text(text) for text in paragraph_texts if _clean_text(text)]


def _looks_like_answer_payload(payload):
    stripped = _clean_text(payload)
    if not stripped:
        return False
    return bool(CHOICE_ANSWER_PATTERN.fullmatch(stripped) or SUBANSWER_PREFIX_PATTERN.match(stripped))


def _looks_like_question_stem(payload):
    stripped = _clean_text(payload)
    if not stripped:
        return False
    if _looks_like_answer_payload(stripped):
        return False
    if stripped.startswith("【答案】"):
        return False
    if re.search(r"(根据|结合).{0,12}(概括|指出|分析|说明|回答|简述|列举|归纳|谈谈)", stripped):
        return True
    if re.search(r"(概括|指出|分析|说明|回答|简述|列举|归纳).{0,20}(原因|特点|影响|意义|不同|问题)", stripped):
        return True
    if "高考真题" in stripped or "单元测试" in stripped:
        return True
    if "（" in stripped and "）" in stripped:
        return True
    return bool(re.search(r"[（(]\s*[ ]{0,2}[A-DＡ-Ｄ]?[ ]{0,2}[）)]", stripped) and len(stripped) > 25)


def _looks_like_instruction_question(payload):
    stripped = _clean_text(payload)
    return bool(
        re.search(r"(根据|结合).{0,12}(概括|指出|分析|说明|回答|简述|列举|归纳|谈谈)", stripped)
        or re.search(r"(概括|指出|分析|说明|回答|简述|列举|归纳).{0,20}(原因|特点|影响|意义|不同|问题)", stripped)
    )


def _clean_marker_payload(payload):
    return _clean_text(re.sub(r"^[：:\s]+", "", payload or ""))


def _split_answer_and_inline_analysis(payload):
    cleaned = _clean_marker_payload(payload)
    match = INLINE_ANALYSIS_PATTERN.match(cleaned)
    if not match:
        return cleaned, ""
    return _clean_text(match.group(1)), _clean_text(match.group(2))


def _has_leading_direct_answer_block(lines, end_index, *, loose=False):
    direct_count = 0
    for text in lines[:end_index]:
        q_match = QUESTION_PREFIX_PATTERN.match(text)
        if not q_match:
            continue
        payload = _clean_text(q_match.group(2))
        looks_like_question = _looks_like_instruction_question(payload) if loose else _looks_like_question_stem(payload)
        if payload and not looks_like_question:
            direct_count += 1
            if direct_count >= 2:
                return True
        elif direct_count or payload:
            return False
    return False


def _find_answer_zone_start(lines):
    reference_indexes = [idx for idx, text in enumerate(lines) if REFERENCE_ANSWER_PATTERN.search(text)]
    homework_indexes = [idx for idx, text in enumerate(lines) if HOMEWORK_PATTERN.search(text)]
    layer_indexes = [idx for idx, text in enumerate(lines) if LAYER_HEADING_PATTERN.fullmatch(text)]
    marker_indexes = [
        idx
        for idx, text in enumerate(lines)
        if ANSWER_MARKER_PATTERN.match(text) or ANALYSIS_MARKER_PATTERN.match(text)
    ]

    if marker_indexes:
        first_marker = marker_indexes[0]
        if not reference_indexes and not homework_indexes and _has_leading_direct_answer_block(lines, first_marker, loose=True):
            return 0
        previous_homework = [idx for idx in homework_indexes if idx < first_marker]
        if previous_homework:
            return previous_homework[-1] + 1
        previous_layer = [idx for idx in layer_indexes if idx < first_marker]
        if previous_layer:
            return previous_layer[-1] + 1
        previous_reference = [idx for idx in reference_indexes if idx < first_marker]
        if previous_reference:
            return previous_reference[-1] + 1
        return max(first_marker - 1, 0)

    if reference_indexes:
        return reference_indexes[-1] + 1
    if homework_indexes:
        return homework_indexes[-1] + 1
    if layer_indexes:
        return layer_indexes[0] + 1
    return 0


def _new_entry(qnum, answer_text=""):
    entry = {
        "qnum": str(qnum),
        "answer_lines": [],
        "analysis_lines": [],
    }
    if answer_text:
        entry["answer_lines"].append(answer_text)
    return entry


def _append_current(entries, current_entry):
    if current_entry is not None and (
        current_entry["answer_lines"] or current_entry["analysis_lines"]
    ):
        entries.append(current_entry)


def _entry_has_answer(entry):
    return bool(entry and any(line.strip() for line in entry["answer_lines"]))


def parse_paragraph_texts(paragraph_texts, *, renumber=False):
    lines = _normalize_paragraph_texts(paragraph_texts)
    start_index = _find_answer_zone_start(lines)
    layered_context = any(LAYER_HEADING_PATTERN.fullmatch(text) for text in lines)
    reference_context = any(idx < start_index for idx, text in enumerate(lines) if REFERENCE_ANSWER_PATTERN.search(text))
    direct_answer_mode = not layered_context and (
        reference_context
        or (start_index == 0 and _has_leading_direct_answer_block(lines[start_index:], min(len(lines) - start_index, 12), loose=True))
        or _has_leading_direct_answer_block(lines[start_index:], min(len(lines) - start_index, 12))
    )
    entries = []
    current_entry = None
    current_mode = None
    pending_qnum = None

    for text in lines[start_index:]:
        if LAYER_HEADING_PATTERN.fullmatch(text):
            current_mode = None
            continue

        q_match = QUESTION_PREFIX_PATTERN.match(text)
        if q_match:
            qnum, payload = q_match.group(1), _clean_text(q_match.group(2))
            inline_answer = ANSWER_MARKER_PATTERN.match(payload)
            inline_analysis = ANALYSIS_MARKER_PATTERN.match(payload)
            if inline_answer:
                _append_current(entries, current_entry)
                answer_text, analysis_text = _split_answer_and_inline_analysis(inline_answer.group(1))
                current_entry = _new_entry(qnum, answer_text)
                if analysis_text:
                    current_entry["analysis_lines"].append(analysis_text)
                current_mode = "analysis" if analysis_text else "answer"
                pending_qnum = qnum
                continue
            if inline_analysis:
                _append_current(entries, current_entry)
                current_entry = _new_entry(qnum)
                first_analysis = _clean_marker_payload(inline_analysis.group(1))
                if first_analysis:
                    current_entry["analysis_lines"].append(first_analysis)
                current_mode = "analysis"
                pending_qnum = qnum
                continue
            if _looks_like_answer_payload(payload):
                _append_current(entries, current_entry)
                current_entry = _new_entry(qnum, payload)
                current_mode = "answer"
                pending_qnum = qnum
                continue
            if payload and direct_answer_mode and not _looks_like_instruction_question(payload):
                _append_current(entries, current_entry)
                current_entry = _new_entry(qnum, payload)
                current_mode = "answer"
                pending_qnum = qnum
                continue
            if _looks_like_question_stem(payload):
                pending_qnum = qnum
                current_mode = None
                continue

        answer_match = ANSWER_MARKER_PATTERN.match(text)
        if answer_match:
            qnum = pending_qnum
            if not qnum and current_entry is not None and not _entry_has_answer(current_entry):
                qnum = current_entry["qnum"]
            if not qnum:
                continue
            _append_current(entries, current_entry)
            answer_text, analysis_text = _split_answer_and_inline_analysis(answer_match.group(1))
            current_entry = _new_entry(qnum, answer_text)
            if analysis_text:
                current_entry["analysis_lines"].append(analysis_text)
            current_mode = "analysis" if analysis_text else "answer"
            continue

        analysis_match = ANALYSIS_MARKER_PATTERN.match(text)
        if analysis_match:
            if current_entry is None:
                if not pending_qnum:
                    continue
                current_entry = _new_entry(pending_qnum)
            first_analysis = _clean_marker_payload(analysis_match.group(1))
            if first_analysis:
                current_entry["analysis_lines"].append(first_analysis)
            current_mode = "analysis"
            continue

        if current_entry is None:
            continue

        if current_mode == "answer":
            current_entry["answer_lines"].append(text)
        elif current_mode == "analysis":
            current_entry["analysis_lines"].append(text)

    _append_current(entries, current_entry)
    if renumber:
        for index, entry in enumerate(entries, 1):
            entry["source_qnum"] = entry["qnum"]
            entry["qnum"] = str(index)
    return entries


def render_standard_lines(entries):
    rendered = []
    for entry in entries:
        qnum = entry["qnum"]
        answer_text = " ".join(line.strip() for line in entry["answer_lines"] if line.strip()).strip()
        analysis_text = " ".join(line.strip() for line in entry["analysis_lines"] if line.strip()).strip()

        if answer_text.startswith("①"):
            for marker, label in CIRCLED_POINT_LABELS.items():
                answer_text = answer_text.replace(marker, label)

        if CHOICE_ANSWER_PATTERN.fullmatch(answer_text):
            rendered.append(f"{qnum}．{answer_text}")
        else:
            rendered.append(f"{qnum}．")
            rendered.append(f"答案：{answer_text}" if answer_text else "答案：")

        rendered.append(f"解析：{analysis_text}" if analysis_text else "解析：")
    return rendered


def match_score(doc, cached_texts=None):
    texts = cached_texts if cached_texts is not None else [
        _clean_text(p.Range.Text) for p in doc.Paragraphs
    ]
    cleaned_texts = [text for text in (_clean_text(text) for text in texts) if text]
    if not cleaned_texts:
        return 0

    matched = 0
    for text in cleaned_texts:
        if any(re.search(pattern, text) for pattern in TEMPLATE_FEATURES["patterns"]):
            matched += 1

    score = matched / len(cleaned_texts)
    joined_head = "\n".join(cleaned_texts[:120])
    keyword_hits = sum(1 for keyword in HISTORY_KEYWORDS if keyword in joined_head)
    if keyword_hits >= 2:
        score += 0.04
    try:
        doc_name = doc.Name
    except Exception:
        doc_name = ""
    if "教师版" in doc_name or "答案" in doc_name:
        score += 0.02
    if re.search(r"第\d+课", doc_name):
        score += 0.02
    return score


def set_font_format(doc):
    try:
        font = doc.Content.Font
        font.Size = 12
        font.Color = 0
        font.Name = "Times New Roman"
        font.NameFarEast = "宋体"
        font.Bold = False
        font.Italic = False
        print("   字体格式设置完成：小四、黑色、宋体/Times New Roman、不加粗")
    except Exception as exc:
        print(f"   字体设置失败: {exc}")


def clean_document(doc):
    print(f"   使用未来高二历史模板清洗: {doc.Name}")
    print(f"   文档共 {doc.Paragraphs.Count} 个段落")

    try:
        for index in range(doc.Shapes.Count, 0, -1):
            doc.Shapes(index).Delete()
        for index in range(doc.InlineShapes.Count, 0, -1):
            doc.InlineShapes(index).Delete()
    except Exception as exc:
        print(f"   删除图片失败: {exc}")

    paragraph_texts = [_clean_text(doc.Paragraphs(index).Range.Text) for index in range(1, doc.Paragraphs.Count + 1)]
    is_layered_document = sum(
        1 for text in paragraph_texts if LAYER_HEADING_PATTERN.fullmatch(_clean_text(text))
    ) >= 2
    entries = parse_paragraph_texts(paragraph_texts, renumber=is_layered_document)
    if not entries:
        print("   未识别到可清洗的历史答案区")
        return False

    rendered_lines = render_standard_lines(entries)
    doc.Content.Text = "\r".join(rendered_lines) + "\r"
    set_font_format(doc)
    print(f"   共提取 {len(entries)} 道答案，已忽略教师版前部题目区")
    return True


TEMPLATE_INFO = TEMPLATE_FEATURES
