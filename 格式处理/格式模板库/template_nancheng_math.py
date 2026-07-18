# coding: utf-8
"""
格式模板 - 南城二小小学数学总答案

适用文档特征：
- 单个答案文档按 `第5单元`、`第6单元`、`第7单元`、`第8、9单元` 分段
- 每个题型段内从 `1．` 重新编号，需要按题目文档顺序映射为整卷题号
- 答案中可能出现 `附加题：` 与独立 `解析：` 续行
"""

from __future__ import annotations

import re
import unicodedata


TEMPLATE_FEATURES = {
    "name": "南城二小 - 小学数学总答案",
    "patterns": [
        r"^第\d+(?:[、,]\d+)?单元$",
        r"^[一二三四五六七八九十]+[、.．]\s*\d+\s*[.．]",
        r"^附加题[：:]",
        r"^解析[：:]",
        r"\d+\s*[.．]\s*[A-D](?=\s|$)",
    ],
    "match_threshold": 0.08,
}


UNIT_HEADING_PATTERN = re.compile(r"^第\d+(?:[、,]\d+)?单元$")
SECTION_PREFIX_PATTERN = re.compile(r"^\s*[一二三四五六七八九十]+[、.．]\s*")
LOCAL_QUESTION_START_PATTERN = re.compile(r"(?:(?<=^)|(?<=\s))(\d+)\s*(?:[.．]\s*|(?=[A-DX√×]))")
ANALYSIS_PATTERN = re.compile(r"^解析[：:]\s*(.*)$")
APPENDIX_PATTERN = re.compile(r"^附加题[：:]\s*(.*)$")
SUBQUESTION_MARKER_PATTERN = re.compile(r"^[（(]1[）)]|①")
ANY_SUBQUESTION_MARKER_PATTERN = re.compile(r"[（(]\s*\d+\s*[）)]")
LEADING_SUBQUESTION_MARKER_PATTERN = re.compile(r"^[（(]\s*(\d+)\s*[）)]")
GARBAGE_PATTERNS = [r"^\s*$"]


def _clean_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKC", text or "")
    return normalized.replace("\r", "").replace("\n", "").replace("\x07", "").replace("\u00a0", " ").strip()


def is_garbage_line(text: str) -> bool:
    cleaned = _clean_text(text)
    return any(re.match(pattern, cleaned) for pattern in GARBAGE_PATTERNS)


def is_unit_heading(text: str) -> bool:
    return bool(UNIT_HEADING_PATTERN.match(_clean_text(text)))


def _normalize_unit_title(text: str) -> str:
    return _clean_text(text).replace(",", "、")


def extract_unit_texts(paragraph_texts: list[str], unit_title: str) -> list[str]:
    target = _normalize_unit_title(unit_title)
    collecting = False
    lines: list[str] = []

    for raw_text in paragraph_texts:
        text = _clean_text(raw_text)
        if not text:
            continue

        if is_unit_heading(text):
            normalized_heading = _normalize_unit_title(text)
            if collecting and normalized_heading != target:
                break
            collecting = normalized_heading == target
            continue

        if collecting:
            lines.append(text)

    return lines


def _strip_section_prefix(text: str) -> tuple[str, bool]:
    cleaned = _clean_text(text)
    stripped = SECTION_PREFIX_PATTERN.sub("", cleaned, count=1).strip()
    return stripped, stripped != cleaned


def _split_local_question_segments(text: str, expected_local_number: int) -> tuple[list[tuple[str, str]], int, bool]:
    cleaned, had_section_prefix = _strip_section_prefix(text)
    if had_section_prefix:
        expected_local_number = 1

    candidate_matches = list(LOCAL_QUESTION_START_PATTERN.finditer(cleaned))
    matches = []
    for match in candidate_matches:
        number = int(match.group(1))
        if number != expected_local_number:
            continue
        matches.append(match)
        expected_local_number += 1

    if not matches:
        return [], expected_local_number, had_section_prefix

    segments: list[tuple[str, str]] = []
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(cleaned)
        payload = cleaned[start:end].strip()
        segments.append((match.group(1), payload))
    return segments, expected_local_number, had_section_prefix


def _next_question_id(question_units, position: int) -> str | None:
    if position >= len(question_units):
        return None
    return str(question_units[position].question_id)


def _question_expected_sub_count(question_units, position: int) -> int:
    if position >= len(question_units):
        return 0
    return len(getattr(question_units[position], "subquestions", []) or [])


def _append_answer(entry: dict | None, text: str) -> None:
    if entry is not None and text:
        entry["answer_lines"].append(text)


def _append_analysis(entry: dict | None, text: str) -> None:
    if entry is not None:
        entry["analysis_lines"].append(text)


def _append_subquestion_answer(entry: dict | None, local_qid: str, payload: str) -> None:
    if entry is not None:
        entry["answer_lines"].append(f"({local_qid}){payload.strip()}")


def _payload_starts_with_subquestion_marker(payload: str) -> bool:
    return bool(SUBQUESTION_MARKER_PATTERN.match((payload or "").strip()))


def _leading_subquestion_number(payload: str) -> int | None:
    match = LEADING_SUBQUESTION_MARKER_PATTERN.match((payload or "").strip())
    return int(match.group(1)) if match else None


def _ensure_subquestion_shape(entry: dict, expected_count: int) -> None:
    if expected_count <= 1:
        return

    joined_answer = " ".join(line for line in entry["answer_lines"] if line).strip()
    if not joined_answer:
        entry["answer_lines"] = [f"({index})" for index in range(1, expected_count + 1)]
        if entry["analysis_lines"]:
            first_analysis = entry["analysis_lines"][0].strip()
            if first_analysis and not ANY_SUBQUESTION_MARKER_PATTERN.search(first_analysis):
                entry["analysis_lines"][0] = f"(1){first_analysis}"
        else:
            entry["analysis_lines"].append("(1)原答案为图片或空白，请回原卷核对")
        return

    existing_markers = re.findall(r"[（(]\s*(\d+)\s*[）)]", joined_answer)
    if existing_markers:
        max_marker = max(int(marker) for marker in existing_markers)
        if max_marker < expected_count:
            entry["answer_lines"].extend(f"({index})" for index in range(max_marker + 1, expected_count + 1))

        answer_without_markers = ANY_SUBQUESTION_MARKER_PATTERN.sub("", joined_answer).strip()
        if not answer_without_markers:
            if entry["analysis_lines"]:
                first_analysis = entry["analysis_lines"][0].strip()
                if first_analysis and not ANY_SUBQUESTION_MARKER_PATTERN.search(first_analysis):
                    entry["analysis_lines"][0] = f"(1){first_analysis}"
            else:
                entry["analysis_lines"].append("(1)原答案为图片或空白，请回原卷核对")
            return
        return

    answer_lines = [f"(1){joined_answer}"]
    answer_lines.extend(f"({index})" for index in range(2, expected_count + 1))
    entry["answer_lines"] = answer_lines


def parse_unit_answer_texts(paragraph_texts: list[str], question_units) -> list[dict]:
    entries: list[dict] = []
    current_entry: dict | None = None
    current_mode = "answer"
    question_position = 0
    expected_local_number = 1
    active_subquestion_entry: dict | None = None
    active_subquestion_target_count = 0

    def start_next_entry(initial_answer: str = "") -> dict | None:
        nonlocal current_entry, current_mode, question_position, active_subquestion_entry, active_subquestion_target_count
        qid = _next_question_id(question_units, question_position)
        if qid is None:
            return None
        current_entry = {
            "kind": "question",
            "qnum": qid,
            "answer_lines": [],
            "analysis_lines": [],
            "expected_sub_count": _question_expected_sub_count(question_units, question_position),
        }
        if initial_answer:
            current_entry["answer_lines"].append(initial_answer)
        entries.append(current_entry)
        question_position += 1
        current_mode = "answer"
        active_subquestion_entry = None
        active_subquestion_target_count = 0
        return current_entry

    def consume_subquestion_segments(segments: list[tuple[str, str]]) -> bool:
        nonlocal active_subquestion_entry, active_subquestion_target_count
        if active_subquestion_entry is None:
            return False
        for local_qid, payload in segments:
            _append_subquestion_answer(active_subquestion_entry, local_qid, payload)
            if int(local_qid) >= active_subquestion_target_count:
                active_subquestion_entry = None
                active_subquestion_target_count = 0
        return True

    for raw_text in paragraph_texts:
        text = _clean_text(raw_text)
        if not text or is_garbage_line(text):
            continue

        appendix_match = APPENDIX_PATTERN.match(text)
        if appendix_match:
            start_next_entry(appendix_match.group(1).strip())
            continue

        analysis_match = ANALYSIS_PATTERN.match(text)
        if analysis_match:
            current_mode = "analysis"
            _append_analysis(current_entry, analysis_match.group(1).strip())
            continue

        segments, expected_local_number, had_section_prefix = _split_local_question_segments(text, expected_local_number)
        if segments:
            if consume_subquestion_segments(segments):
                continue

            expected_sub_count = _question_expected_sub_count(question_units, question_position)
            should_group_as_subanswers = (
                had_section_prefix
                and expected_sub_count > 1
                and len(segments) <= expected_sub_count
                and segments[0][0] == "1"
                and all(payload.strip() for _local_qid, payload in segments)
                and not _payload_starts_with_subquestion_marker(segments[0][1])
            )
            if should_group_as_subanswers:
                entry = start_next_entry()
                for local_qid, payload in segments:
                    _append_subquestion_answer(entry, local_qid, payload)
                if segments and int(segments[-1][0]) < expected_sub_count:
                    active_subquestion_entry = entry
                    active_subquestion_target_count = expected_sub_count
                continue

            for _local_qid, payload in segments:
                start_next_entry(payload)
            continue

        if current_mode == "analysis":
            _append_analysis(current_entry, text)
        else:
            leading_subquestion_number = _leading_subquestion_number(text)
            expected_sub_count = current_entry.get("expected_sub_count", 0) if current_entry else 0
            if leading_subquestion_number and expected_sub_count and leading_subquestion_number > expected_sub_count:
                _append_analysis(current_entry, text)
                continue
            _append_answer(current_entry, text)

    for entry in entries:
        _ensure_subquestion_shape(entry, entry.get("expected_sub_count", 0))

    return entries


def render_standard_lines(entries: list[dict]) -> list[str]:
    lines: list[str] = []
    for entry in entries:
        if entry.get("kind") != "question":
            continue

        qnum = entry["qnum"]
        answer_lines = [line for line in entry.get("answer_lines", []) if line is not None]
        analysis_lines = [line for line in entry.get("analysis_lines", []) if line is not None]

        first_answer = answer_lines[0].strip() if answer_lines else ""
        lines.append(f"{qnum}．{first_answer}" if first_answer else f"{qnum}．")
        for extra_answer in answer_lines[1:]:
            lines.append(extra_answer)

        first_analysis = analysis_lines[0].strip() if analysis_lines else ""
        lines.append(f"解析：{first_analysis}" if first_analysis else "解析：")
        for extra_analysis in analysis_lines[1:]:
            lines.append(extra_analysis)

    return lines


def match_score(doc, cached_texts=None):
    texts = cached_texts if cached_texts is not None else [_clean_text(p.Range.Text) for p in doc.Paragraphs]
    total_lines = 0
    matched_lines = 0

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
    print(f"   ▶ 使用南城小学数学模板清洗: {doc.Name}")
    print("   ⚠️ 该模板的总答案需要结合题目文档按单元映射；请优先使用离线拆分/审核流程。")
    set_font_format(doc)
    return True


TEMPLATE_INFO = TEMPLATE_FEATURES
