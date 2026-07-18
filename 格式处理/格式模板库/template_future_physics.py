# coding: utf-8
"""
格式模板 - 未来高二物理总答案

适用文档特征：
- 章节标题：`课时跟踪检测(一)`、`课时跟踪检测(七十四)`
- 选择题答案格式：`1.选D 解析内容...`
- 选择题变体：`7.` 下一行 `选B 解析内容...`
- 主观题格式：`8.解析:(1)...`，末尾单独给出 `答案:` 行
"""

import re
import unicodedata


TEMPLATE_FEATURES = {
    "name": "未来高二 - 物理总答案",
    "patterns": [
        r"^课时跟踪检测[(（][^)）]+[)）]$",
        r"^\d+[．.]?\s*选[A-D]+",
        r"^\d+[．.]?\s*解析[：:]",
        r"^答案[：:]",
    ],
    "match_threshold": 0.02,
}


SECTION_PATTERN = re.compile(r"^课时跟踪检测[(（][^)）]+[)）]$")
QUESTION_WITH_PUNCT_PATTERN = re.compile(r"^(\d+)[．.]\s*(.*)$")
QUESTION_STANDALONE_PATTERN = re.compile(r"^(\d+)\s*$")
QUESTION_PREFIX_RAW_PATTERN = re.compile(r"^(\d+)[．.]\s*")
CHOICE_PATTERN = re.compile(r"^选\s*([A-D]+)\s*(.*)$")
CHOICE_RAW_PATTERN = re.compile(r"^\s*选\s*([A-D]+)\s*")
ANALYSIS_PATTERN = re.compile(r"^(?:解析|【解析】)[：:]?\s*(.*)$")
ANALYSIS_RAW_PATTERN = re.compile(r"^\s*(?:解析|【解析】)[：:]?\s*")
ANSWER_PATTERN = re.compile(r"^(?:答案|【答案】)[：:]?\s*(.*)$")
ANSWER_RAW_PATTERN = re.compile(r"^\s*(?:答案|【答案】)[：:]?\s*")
EMBEDDED_QUESTION_SPLIT_PATTERN = re.compile(
    r"(?<=[。；;！？!?）)])(?=\d+[．.](?:选|解析|答案))"
)
GARBAGE_PATTERNS = [
    r"^\s*$",
    r"^配套检测卷参考答案$",
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
    return bool(SECTION_PATTERN.match(_clean_text(text)))


def _match_question_line(text):
    cleaned = _clean_text(text)
    punct_match = QUESTION_WITH_PUNCT_PATTERN.match(cleaned)
    if punct_match:
        return punct_match.group(1), punct_match.group(2).strip()
    standalone_match = QUESTION_STANDALONE_PATTERN.match(cleaned)
    if standalone_match:
        return standalone_match.group(1), ""
    return None


def _match_choice_line(text):
    return CHOICE_PATTERN.match(_clean_text(text))


def _match_analysis_line(text):
    return ANALYSIS_PATTERN.match(_clean_text(text))


def _match_answer_line(text):
    return ANSWER_PATTERN.match(_clean_text(text))


def _finalize_entry(entries, current_entry):
    if current_entry is None:
        return
    entries.append(current_entry)


def _split_embedded_question_lines(raw_text):
    cleaned = _clean_text(raw_text)
    if not cleaned:
        return []

    split_points = [0]
    split_points.extend(match.start() for match in EMBEDDED_QUESTION_SPLIT_PATTERN.finditer(cleaned))
    split_points.append(len(cleaned))

    segments = []
    for start, end in zip(split_points, split_points[1:]):
        segment = cleaned[start:end].strip()
        if segment:
            segments.append(segment)
    return segments


def parse_paragraph_texts(paragraph_texts):
    entries = []
    current_entry = None
    current_mode = None

    for raw_text in paragraph_texts:
        for text in _split_embedded_question_lines(raw_text):
            if is_garbage_line(text):
                continue

            if is_section_line(text):
                _finalize_entry(entries, current_entry)
                current_entry = None
                current_mode = None
                entries.append({"kind": "section", "text": text})
                continue

            question_match = _match_question_line(text)
            if question_match:
                _finalize_entry(entries, current_entry)

                qnum, payload = question_match
                current_entry = {
                    "kind": "question",
                    "qnum": qnum,
                    "question_type": "unknown",
                    "answer_text": "",
                    "answer_lines": [],
                    "analysis_lines": [],
                }
                current_mode = None

                if payload:
                    choice_match = _match_choice_line(payload)
                    analysis_match = _match_analysis_line(payload)
                    answer_match = _match_answer_line(payload)
                    if choice_match:
                        current_entry["question_type"] = "choice"
                        current_entry["answer_text"] = choice_match.group(1)
                        if choice_match.group(2).strip():
                            current_entry["analysis_lines"].append(choice_match.group(2).strip())
                            current_mode = "analysis"
                    elif analysis_match:
                        current_entry["question_type"] = "big"
                        if analysis_match.group(1).strip():
                            current_entry["analysis_lines"].append(analysis_match.group(1).strip())
                        current_mode = "analysis"
                    elif answer_match:
                        current_entry["question_type"] = "big"
                        if answer_match.group(1).strip():
                            current_entry["answer_lines"].append(answer_match.group(1).strip())
                        current_mode = "answer"
                    else:
                        current_entry["question_type"] = "big"
                        current_entry["answer_lines"].append(payload)
                        current_mode = "answer"
                continue

            if current_entry is None:
                continue

            choice_match = _match_choice_line(text)
            if choice_match and current_entry["question_type"] == "unknown":
                current_entry["question_type"] = "choice"
                current_entry["answer_text"] = choice_match.group(1)
                if choice_match.group(2).strip():
                    current_entry["analysis_lines"].append(choice_match.group(2).strip())
                current_mode = "analysis"
                continue

            analysis_match = _match_analysis_line(text)
            if analysis_match:
                if current_entry["question_type"] == "unknown":
                    current_entry["question_type"] = "big"
                current_mode = "analysis"
                if analysis_match.group(1).strip():
                    current_entry["analysis_lines"].append(analysis_match.group(1).strip())
                continue

            answer_match = _match_answer_line(text)
            if answer_match:
                if current_entry["question_type"] == "unknown":
                    current_entry["question_type"] = "big"
                current_mode = "answer"
                if answer_match.group(1).strip():
                    current_entry["answer_lines"].append(answer_match.group(1).strip())
                continue

            if current_entry["question_type"] == "choice":
                current_entry["analysis_lines"].append(text)
                current_mode = "analysis"
                continue

            if current_entry["question_type"] == "unknown":
                current_entry["question_type"] = "big"
                current_mode = current_mode or "analysis"

            if current_mode == "answer":
                current_entry["answer_lines"].append(text)
            else:
                current_entry["analysis_lines"].append(text)
                current_mode = "analysis"

    _finalize_entry(entries, current_entry)
    return entries


def render_standard_lines(entries):
    lines = []
    for entry in entries:
        if entry["kind"] == "section":
            lines.append(entry["text"])
            continue

        qnum = entry["qnum"]
        if entry["question_type"] == "choice":
            lines.append(f"{qnum}．{entry['answer_text']}" if entry["answer_text"] else f"{qnum}．")
        else:
            lines.append(f"{qnum}．")
            if entry["answer_lines"]:
                first_answer = entry["answer_lines"][0].strip()
                lines.append(f"答案：{first_answer}" if first_answer else "答案：")
                for extra_line in entry["answer_lines"][1:]:
                    lines.append(extra_line)
            else:
                lines.append("答案：")

        if entry["analysis_lines"]:
            first_analysis = entry["analysis_lines"][0].strip()
            lines.append(f"解析：{first_analysis}" if first_analysis else "解析：")
            for extra_line in entry["analysis_lines"][1:]:
                lines.append(extra_line)
        else:
            lines.append("解析：")

    return lines


def _safe_inline_shape_count(rng):
    try:
        return rng.InlineShapes.Count
    except Exception:
        return 0


def _safe_omath_count(rng):
    try:
        return rng.OMaths.Count
    except Exception:
        return 0


def _copy_source_document(doc):
    source_doc = doc.Application.Documents.Add()
    source_doc.Content.FormattedText = doc.Content.FormattedText
    return source_doc


def _split_wps_paragraph_segments(raw_text):
    if not _clean_text(raw_text):
        return []

    split_points = [0]
    split_points.extend(match.start() for match in EMBEDDED_QUESTION_SPLIT_PATTERN.finditer(raw_text))
    split_points.append(len(raw_text))

    segments = []
    for start, end in zip(split_points, split_points[1:]):
        segment_raw = raw_text[start:end]
        segment_text = _clean_text(segment_raw)
        if segment_text:
            segments.append(
                {
                    "raw": segment_raw,
                    "text": segment_text,
                    "start": start,
                    "end": end,
                }
            )
    return segments


def _match_question_segment(segment):
    raw = segment["raw"]
    leading = len(raw) - len(raw.lstrip())
    stripped_raw = raw[leading:]

    punct_match = QUESTION_PREFIX_RAW_PATTERN.match(stripped_raw)
    if punct_match:
        payload_start_in_segment = leading + punct_match.end()
        payload_raw = raw[payload_start_in_segment:]
        return {
            "qnum": punct_match.group(1),
            "payload_raw": payload_raw,
            "payload_text": _clean_text(payload_raw),
            "payload_start": segment["start"] + payload_start_in_segment,
        }

    standalone_match = re.match(r"^(\d+)\s*$", stripped_raw)
    if standalone_match:
        return {
            "qnum": standalone_match.group(1),
            "payload_raw": "",
            "payload_text": "",
            "payload_start": segment["start"] + leading + len(standalone_match.group(1)),
        }

    return None


def _new_wps_question_entry(qnum):
    return {
        "kind": "question",
        "qnum": qnum,
        "question_type": "unknown",
        "answer_text": "",
        "analysis_blocks": [],
        "answer_blocks": [],
    }


def _append_wps_block(entry, block_kind, paragraph_index, start_offset=0, end_offset=None):
    if entry is None:
        return
    target_key = "answer_blocks" if block_kind == "answer" else "analysis_blocks"
    entry[target_key].append(
        {
            "index": paragraph_index,
            "start_offset": max(0, int(start_offset)),
            "end_offset": None if end_offset is None else max(0, int(end_offset)),
        }
    )


def _source_range_for_block(source_doc, block):
    para = source_doc.Paragraphs(block["index"])
    src_range = para.Range.Duplicate
    start_offset = max(0, int(block.get("start_offset", 0)))
    max_offset = max(0, src_range.End - src_range.Start - 1)
    if start_offset > max_offset:
        start_offset = max_offset
    src_range.Start = src_range.Start + start_offset

    end_offset = block.get("end_offset")
    if end_offset is not None:
        end_offset = min(max(0, int(end_offset)), max_offset)
        src_range.End = para.Range.Start + end_offset
    return src_range


def _block_has_rich_payload(source_doc, block):
    src_range = _source_range_for_block(source_doc, block)
    src_text = src_range.Text or ""
    return (
        _safe_inline_shape_count(src_range) > 0
        or _safe_omath_count(src_range) > 0
        or "\x01" in src_text
    )


def _entry_answer_requires_analysis_payload(source_doc, entry):
    return any(_block_has_rich_payload(source_doc, block) for block in entry.get("answer_blocks", []))


def _collect_wps_entries(source_doc):
    entries = []
    current_entry = None
    current_mode = None
    paras = source_doc.Paragraphs

    for paragraph_index in range(1, paras.Count + 1):
        para = paras(paragraph_index)
        rng = para.Range
        raw_text = (rng.Text or "").replace("\r", "").replace("\n", "").replace("\x07", "")
        segments = _split_wps_paragraph_segments(raw_text)

        if not segments:
            if current_entry is not None and _safe_inline_shape_count(rng) > 0:
                if current_entry["question_type"] == "unknown":
                    current_entry["question_type"] = "big"
                _append_wps_block(current_entry, current_mode or "analysis", paragraph_index, 0)
            continue

        for segment in segments:
            text = segment["text"]
            segment_end_offset = (
                None
                if len(segments) == 1 and segment["start"] == 0 and segment["end"] == len(raw_text)
                else segment["end"]
            )
            if is_garbage_line(text):
                continue

            if is_section_line(text):
                _finalize_entry(entries, current_entry)
                current_entry = None
                current_mode = None
                entries.append({"kind": "section", "text": text})
                continue

            question_match = _match_question_segment(segment)
            if question_match:
                _finalize_entry(entries, current_entry)
                current_entry = _new_wps_question_entry(question_match["qnum"])
                current_mode = None

                payload_raw = question_match["payload_raw"]
                payload_start = question_match["payload_start"]
                payload_end = segment["end"]
                if question_match["payload_text"]:
                    choice_match = CHOICE_RAW_PATTERN.match(payload_raw)
                    analysis_match = ANALYSIS_RAW_PATTERN.match(payload_raw)
                    answer_match = ANSWER_RAW_PATTERN.match(payload_raw)
                    if choice_match:
                        current_entry["question_type"] = "choice"
                        current_entry["answer_text"] = choice_match.group(1)
                        analysis_start = payload_start + choice_match.end()
                        if analysis_start < payload_end:
                            _append_wps_block(
                                current_entry,
                                "analysis",
                                paragraph_index,
                                analysis_start,
                                segment_end_offset,
                            )
                            current_mode = "analysis"
                    elif analysis_match:
                        current_entry["question_type"] = "big"
                        analysis_start = payload_start + analysis_match.end()
                        if analysis_start < payload_end:
                            _append_wps_block(
                                current_entry,
                                "analysis",
                                paragraph_index,
                                analysis_start,
                                segment_end_offset,
                            )
                        current_mode = "analysis"
                    elif answer_match:
                        current_entry["question_type"] = "big"
                        answer_start = payload_start + answer_match.end()
                        _append_wps_block(
                            current_entry,
                            "answer",
                            paragraph_index,
                            answer_start,
                            segment_end_offset,
                        )
                        current_mode = "answer"
                    else:
                        current_entry["question_type"] = "big"
                        _append_wps_block(
                            current_entry,
                            "answer",
                            paragraph_index,
                            payload_start,
                            segment_end_offset,
                        )
                        current_mode = "answer"
                continue

            if current_entry is None:
                continue

            segment_raw = segment["raw"]
            segment_start = segment["start"]
            segment_end = segment["end"]
            segment_end_offset = (
                None
                if len(segments) == 1 and segment_start == 0 and segment_end == len(raw_text)
                else segment_end
            )

            choice_match = CHOICE_RAW_PATTERN.match(segment_raw)
            if choice_match and current_entry["question_type"] == "unknown":
                current_entry["question_type"] = "choice"
                current_entry["answer_text"] = choice_match.group(1)
                analysis_start = segment_start + choice_match.end()
                if analysis_start < segment_end:
                    _append_wps_block(
                        current_entry,
                        "analysis",
                        paragraph_index,
                        analysis_start,
                        segment_end_offset,
                    )
                current_mode = "analysis"
                continue

            analysis_match = ANALYSIS_RAW_PATTERN.match(segment_raw)
            if analysis_match:
                if current_entry["question_type"] == "unknown":
                    current_entry["question_type"] = "big"
                analysis_start = segment_start + analysis_match.end()
                if analysis_start < segment_end:
                    _append_wps_block(
                        current_entry,
                        "analysis",
                        paragraph_index,
                        analysis_start,
                        segment_end_offset,
                    )
                current_mode = "analysis"
                continue

            answer_match = ANSWER_RAW_PATTERN.match(segment_raw)
            if answer_match:
                if current_entry["question_type"] == "unknown":
                    current_entry["question_type"] = "big"
                answer_start = segment_start + answer_match.end()
                _append_wps_block(current_entry, "answer", paragraph_index, answer_start, segment_end_offset)
                current_mode = "answer"
                continue

            if current_entry["question_type"] == "choice":
                _append_wps_block(current_entry, "analysis", paragraph_index, segment_start, segment_end_offset)
                current_mode = "analysis"
                continue

            if current_entry["question_type"] == "unknown":
                current_entry["question_type"] = "big"
                current_mode = current_mode or "analysis"

            _append_wps_block(current_entry, current_mode or "analysis", paragraph_index, segment_start, segment_end_offset)

    _finalize_entry(entries, current_entry)
    return entries


def _append_formatted_block(target_doc, source_doc, block):
    src_range = _source_range_for_block(source_doc, block)
    src_text = src_range.Text or ""
    dest_range = target_doc.Range(target_doc.Content.End - 1, target_doc.Content.End - 1)
    dest_range.FormattedText = src_range.FormattedText
    if src_text and not src_text.endswith("\r"):
        target_doc.Content.InsertAfter("\r")


def _render_wps_entries(target_doc, source_doc, entries):
    target_doc.Content.Text = ""
    for entry in entries:
        if entry["kind"] == "section":
            target_doc.Content.InsertAfter(f"{entry['text']}\r")
            continue

        if entry["question_type"] == "choice":
            answer_text = entry["answer_text"].strip()
            target_doc.Content.InsertAfter(f"{entry['qnum']}．{answer_text}\r" if answer_text else f"{entry['qnum']}．\r")
            analysis_blocks = entry["analysis_blocks"]
            answer_blocks = []
            answer_to_analysis = False
        else:
            answer_to_analysis = _entry_answer_requires_analysis_payload(source_doc, entry)
            target_doc.Content.InsertAfter(f"{entry['qnum']}．\r")
            target_doc.Content.InsertAfter("答案：\r")
            answer_blocks = [] if answer_to_analysis else entry["answer_blocks"]
            analysis_blocks = entry["analysis_blocks"]

        for block in answer_blocks:
            _append_formatted_block(target_doc, source_doc, block)

        target_doc.Content.InsertAfter("解析：\r")
        for block in analysis_blocks:
            _append_formatted_block(target_doc, source_doc, block)
        if answer_to_analysis and entry["answer_blocks"]:
            target_doc.Content.InsertAfter("答案：\r")
            for block in entry["answer_blocks"]:
                _append_formatted_block(target_doc, source_doc, block)


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
    print(f"   ▶ 使用未来高二物理模板清洗: {doc.Name}")
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
        print(f"   ! 创建富文本副本失败: {exc}")
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
    print(f"   ✓ 共处理 {question_count} 道题，保留 {section_count} 个课时标题")

    set_font_format(doc)
    if source_doc is not None:
        try:
            source_doc.Close(SaveChanges=False)
        except Exception:
            pass
    return True


TEMPLATE_INFO = TEMPLATE_FEATURES
