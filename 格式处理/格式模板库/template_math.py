# coding: utf-8
"""
格式模板 - 安乡金海初二数学答案格式

适用文档特征：
- 选择题答案格式：`1．C`
- 选择题解析格式：下一行或同行 `【详解】...`
- 填空/解答题格式：`7．答案内容`、`13．【详解】（1）解...`
- 章节标题：`参考答案`、`一、选择题`、`二、填空题`、`三、解答题`
"""

import re
import unicodedata


TEMPLATE_FEATURES = {
    "name": "安乡金海 - 初二 - 数学",
    "patterns": [
        r"^\d+[．.]\s*[A-D](?:\s*$|\s*【(?:详解|解析)】)",
        r"^\d+[．.]\s*【(?:详解|解析|分析)】",
        r"^[一二三四五六七八九十]+、(?:选择题|填空题|解答题)",
        r"^参考答案$",
    ],
    "match_threshold": 0.02,
}


ANSWER_TITLE_PATTERN = re.compile(r"^\s*参考答案\s*$")
SECTION_PATTERN = re.compile(r"^\s*[一二三四五六七八九十]+、.*(?:选择题|填空题|解答题).*$")
QUESTION_PATTERN = re.compile(r"^(\d+)[．.]\s*(.*)$")
CHOICE_PATTERN = re.compile(r"^(\d+)[．.]\s*([A-D])(?:\s+(.*))?$")
ANALYSIS_MARKER_PATTERN = re.compile(r"^(?:【(?:详解|解析|分析)】|(?:详解|解析|分析)[：:]?)\s*(.*)$")
INLINE_ANALYSIS_SPLIT_PATTERN = re.compile(r"^(.*?)\s*【(?:详解|解析|分析)】\s*(.*)$")
QUESTION_PREFIX_PATTERN = re.compile(r"^(\d+)[．.]\s*")

GARBAGE_PATTERNS = [
    r"^\s*$",
    r"^>$",
]


def _clean_text(text):
    normalized = unicodedata.normalize("NFKC", text or "")
    return normalized.replace("\r", "").replace("\n", "").replace("\x07", "").replace("\u00a0", " ").strip()


def is_garbage_line(text):
    cleaned = _clean_text(text)
    for pattern in GARBAGE_PATTERNS:
        if re.match(pattern, cleaned):
            return True
    return False


def is_section_line(text):
    cleaned = _clean_text(text)
    return bool(ANSWER_TITLE_PATTERN.match(cleaned) or SECTION_PATTERN.match(cleaned))


def _is_question_line(text):
    return bool(QUESTION_PATTERN.match(_clean_text(text)))


def _strip_analysis_marker(text):
    match = ANALYSIS_MARKER_PATTERN.match(_clean_text(text))
    if not match:
        return None
    return match.group(1).strip()


def _split_inline_answer_and_analysis(text):
    cleaned = _clean_text(text)
    match = INLINE_ANALYSIS_SPLIT_PATTERN.match(cleaned)
    if not match:
        return cleaned, None
    answer_text = match.group(1).strip()
    analysis_text = match.group(2).strip()
    return answer_text, analysis_text


def _starts_new_block(text):
    cleaned = _clean_text(text)
    return is_section_line(cleaned) or _is_question_line(cleaned)


def _safe_inline_shape_count(rng):
    try:
        return rng.InlineShapes.Count
    except Exception:
        return 0


def _copy_source_document(doc):
    source_doc = doc.Application.Documents.Add()
    source_doc.Content.FormattedText = doc.Content.FormattedText
    return source_doc


def _paragraph_payload_start_offset(raw_text):
    match = QUESTION_PREFIX_PATTERN.match(raw_text or "")
    if not match:
        return 0
    return match.end()


def _is_formula_like_payload(raw_text, rng):
    return _safe_inline_shape_count(rng) > 0 or "\x01" in (raw_text or "")


def _collect_wps_entries(source_doc):
    entries = []
    current_question = None
    paras = source_doc.Paragraphs

    for index in range(1, paras.Count + 1):
        rng = paras(index).Range
        raw_text = (rng.Text or "").replace("\r", "").replace("\n", "").replace("\x07", "")
        text = raw_text.strip()
        if not text or is_garbage_line(text):
            continue

        if is_section_line(text):
            if current_question is not None:
                entries.append(current_question)
                current_question = None
            entries.append({"kind": "section", "text": text})
            continue

        question_match = QUESTION_PATTERN.match(text)
        if question_match:
            if current_question is not None:
                entries.append(current_question)

            qnum = question_match.group(1)
            payload = _clean_text(question_match.group(2))
            payload_start = _paragraph_payload_start_offset(raw_text)
            current_question = {
                "kind": "question",
                "qnum": qnum,
                "answer_text": "",
                "analysis_blocks": [],
            }

            if payload:
                marker_payload = _strip_analysis_marker(payload)
                if marker_payload is not None:
                    current_question["analysis_blocks"].append({"index": index, "start_offset": payload_start})
                elif _is_formula_like_payload(raw_text, rng):
                    current_question["analysis_blocks"].append({"index": index, "start_offset": payload_start})
                else:
                    answer_text, inline_analysis = _split_inline_answer_and_analysis(payload)
                    current_question["answer_text"] = answer_text.strip()
                    if inline_analysis:
                        current_question["analysis_blocks"].append({"index": index, "start_offset": payload_start + len(answer_text)})
            elif _is_formula_like_payload(raw_text, rng):
                current_question["analysis_blocks"].append({"index": index, "start_offset": payload_start})
            continue

        if current_question is None:
            continue

        current_question["analysis_blocks"].append({"index": index, "start_offset": 0})

    if current_question is not None:
        entries.append(current_question)

    return entries


def _append_formatted_block(target_doc, source_doc, block):
    para = source_doc.Paragraphs(block["index"])
    src_range = para.Range.Duplicate
    start_offset = max(0, int(block.get("start_offset", 0)))
    max_offset = max(0, src_range.End - src_range.Start - 1)
    if start_offset > max_offset:
        start_offset = max_offset
    src_range.Start = src_range.Start + start_offset
    dest_range = target_doc.Range(target_doc.Content.End - 1, target_doc.Content.End - 1)
    dest_range.FormattedText = src_range.FormattedText


def _render_wps_entries(target_doc, source_doc, entries):
    target_doc.Content.Text = ""
    for entry in entries:
        if entry["kind"] == "section":
            target_doc.Content.InsertAfter(f"{entry['text']}\r")
            continue

        answer_text = entry["answer_text"].strip()
        if answer_text:
            target_doc.Content.InsertAfter(f"{entry['qnum']}．{answer_text}\r")
        else:
            target_doc.Content.InsertAfter(f"{entry['qnum']}．\r")

        if entry["analysis_blocks"]:
            target_doc.Content.InsertAfter("解析：\r")
            for block in entry["analysis_blocks"]:
                _append_formatted_block(target_doc, source_doc, block)
        else:
            target_doc.Content.InsertAfter("解析：\r")


def parse_paragraph_texts(paragraph_texts):
    lines = [_clean_text(text) for text in paragraph_texts]
    entries = []
    i = 0

    while i < len(lines):
        text = lines[i]
        if not text or is_garbage_line(text):
            i += 1
            continue

        if is_section_line(text):
            entries.append({"kind": "section", "text": text})
            i += 1
            continue

        choice_match = CHOICE_PATTERN.match(text)
        if choice_match:
            qnum = choice_match.group(1)
            answer_text = choice_match.group(2)
            remaining = _clean_text(choice_match.group(3) or "")
            analysis_lines = []

            inline_analysis = _strip_analysis_marker(remaining)
            if inline_analysis is not None:
                if inline_analysis:
                    analysis_lines.append(inline_analysis)
            elif remaining:
                _, split_analysis = _split_inline_answer_and_analysis(remaining)
                if split_analysis:
                    analysis_lines.append(split_analysis)

            i += 1
            if not analysis_lines:
                while i < len(lines):
                    next_text = lines[i]
                    if not next_text or is_garbage_line(next_text):
                        i += 1
                        continue
                    if _starts_new_block(next_text):
                        break

                    marker_analysis = _strip_analysis_marker(next_text)
                    if marker_analysis is None:
                        break

                    if marker_analysis:
                        analysis_lines.append(marker_analysis)
                    i += 1
                    while i < len(lines):
                        follow_text = lines[i]
                        if not follow_text or is_garbage_line(follow_text):
                            i += 1
                            continue
                        if _starts_new_block(follow_text):
                            break
                        analysis_lines.append(follow_text)
                        i += 1
                    break

            entries.append(
                {
                    "kind": "question",
                    "qnum": qnum,
                    "answer_lines": [answer_text],
                    "analysis_lines": analysis_lines,
                }
            )
            continue

        question_match = QUESTION_PATTERN.match(text)
        if question_match:
            qnum = question_match.group(1)
            payload = _clean_text(question_match.group(2))
            answer_lines = []
            analysis_lines = []

            if payload:
                marker_payload = _strip_analysis_marker(payload)
                if marker_payload is not None:
                    if marker_payload:
                        answer_lines.append(marker_payload)
                else:
                    answer_text, inline_analysis = _split_inline_answer_and_analysis(payload)
                    if answer_text:
                        answer_lines.append(answer_text)
                    if inline_analysis:
                        analysis_lines.append(inline_analysis)

            i += 1
            collecting_analysis = bool(analysis_lines)
            while i < len(lines):
                next_text = lines[i]
                if not next_text or is_garbage_line(next_text):
                    i += 1
                    continue
                if _starts_new_block(next_text):
                    break

                marker_analysis = _strip_analysis_marker(next_text)
                if marker_analysis is not None:
                    collecting_analysis = True
                    if marker_analysis:
                        analysis_lines.append(marker_analysis)
                    i += 1
                    continue

                if collecting_analysis:
                    analysis_lines.append(next_text)
                else:
                    answer_lines.append(next_text)
                i += 1

            entries.append(
                {
                    "kind": "question",
                    "qnum": qnum,
                    "answer_lines": answer_lines,
                    "analysis_lines": analysis_lines,
                }
            )
            continue

        i += 1

    return entries


def render_standard_lines(entries):
    lines = []
    for entry in entries:
        if entry["kind"] == "section":
            lines.append(entry["text"])
            continue

        qnum = entry["qnum"]
        answer_lines = [line for line in entry["answer_lines"] if line is not None]
        analysis_lines = [line for line in entry["analysis_lines"] if line is not None]

        first_answer = answer_lines[0] if answer_lines else ""
        lines.append(f"{qnum}．{first_answer}" if first_answer else f"{qnum}．")
        for extra_line in answer_lines[1:]:
            lines.append(extra_line)

        if analysis_lines:
            first_analysis = analysis_lines[0].strip()
            lines.append(f"解析：{first_analysis}" if first_analysis else "解析：")
            for extra_line in analysis_lines[1:]:
                lines.append(extra_line)
        else:
            lines.append("解析：")

    return lines


def match_score(doc, cached_texts=None):
    total_lines = 0
    matched_lines = 0

    texts = cached_texts if cached_texts is not None else [
        _clean_text(p.Range.Text) for p in doc.Paragraphs
    ]

    for text in texts:
        cleaned = _clean_text(text)
        if not cleaned:
            continue

        total_lines += 1
        for pattern in TEMPLATE_FEATURES["patterns"]:
            if re.search(pattern, cleaned):
                matched_lines += 1
                break

    if total_lines == 0:
        return 0

    return matched_lines / total_lines


def set_font_format(doc):
    try:
        font = doc.Content.Font
        font.Size = 12
        font.Color = 0
        font.Name = "Times New Roman"
        font.NameFarEast = "宋体"
        font.Bold = False
        font.Italic = False
        print("   ✓ 字体格式设置完成：小四、黑色、宋体/Times New Roman、不加粗")
    except Exception as exc:
        print(f"   ! 字体设置失败: {exc}")


def clean_document(doc):
    print(f"   ▶ 使用数学模板清洗: {doc.Name}")
    print(f"   📊 文档共 {doc.Paragraphs.Count} 个段落")

    source_doc = None
    try:
        source_doc = _copy_source_document(doc)
        shapes_count = doc.Shapes.Count
        if shapes_count > 0:
            for index in range(shapes_count, 0, -1):
                doc.Shapes(index).Delete()
            print(f"   ✓ 删除 {shapes_count} 个悬浮图片")
    except Exception as exc:
        print(f"   ! 删除图片失败: {exc}")
        if source_doc is not None:
            try:
                source_doc.Close(SaveChanges=False)
            except Exception:
                pass
        return False

    entries = _collect_wps_entries(source_doc)
    _render_wps_entries(doc, source_doc, entries)

    question_count = sum(1 for entry in entries if entry["kind"] == "question")
    section_count = sum(1 for entry in entries if entry["kind"] == "section")
    print(f"   ✓ 共处理 {question_count} 道题，保留 {section_count} 个标题")

    set_font_format(doc)
    if source_doc is not None:
        try:
            source_doc.Close(SaveChanges=False)
        except Exception:
            pass
    return True


TEMPLATE_INFO = TEMPLATE_FEATURES
