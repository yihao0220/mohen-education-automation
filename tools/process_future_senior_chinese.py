from __future__ import annotations

import argparse
from copy import deepcopy
from dataclasses import dataclass
from hashlib import sha256
import json
import os
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from lxml import etree as ET

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core import (
    build_answer_units_from_docx,
    derive_review_status_path,
    get_review_gate_result,
    initialize_review_status,
    update_review_status,
)
from tools.process_guanmei_geography import (
    DOCUMENT_XML,
    W,
    _assert_relationships_resolve,
    _canonical_children,
    _document_body,
    _element_text,
    _new_text_paragraph,
    _page_margins_twips,
    _parse_xml,
    _set_text,
    _slice_paragraph,
    _write_docx_package,
)


HEADING_ONE_STYLE_IDS = {"1", "Heading1"}
STYLES_XML = "word/styles.xml"
CATEGORY_TITLE_PATTERN = re.compile(r"真题精[炼练]\s*$")
QUESTION_ID_PATTERN = re.compile(r"^\s*(\d{1,2})\s*[.．、]")
ANSWER_START_PATTERN = re.compile(r"^\s*(\d{1,2})\s*[.．、]\s*(.*)$", re.S)
ANALYSIS_PATTERN = re.compile(r"^\s*(?:【\s*解析\s*】|解析\s*[:：])\s*(.*)$", re.S)
READING_GROUP_PATTERN = re.compile(
    r"完成\s*(\d{1,2})\s*(?:[~～\-—–]|至)\s*(\d{1,2})\s*(?:小)?题"
)


@dataclass(frozen=True)
class QuestionPartSpec:
    sequence: int
    category: str
    title: str
    start: int
    end: int

    @property
    def filename(self) -> str:
        return f"{self.sequence:02d} {self.title}.docx"

    @property
    def answer_filename(self) -> str:
        return f"{self.sequence:02d} {self.title}-答案_已清洗.docx"


@dataclass(frozen=True)
class AnswerSection:
    title: str
    normalized_title: str
    family: str
    start: int
    end: int
    source_ids: tuple[str, ...]
    entry_ranges: tuple[tuple[int, int], ...]
    ignored_trailing_ids: tuple[str, ...] = ()


def _digest(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def _safe_title(value: str) -> str:
    return _normalize_text(re.sub(r'[\\/:*?"<>|]', " ", value))


def normalize_exam_title(value: str) -> str:
    """把目录里的空格及易混罗马数字统一为匹配键，不改输出标题。"""

    normalized = _normalize_text(value)
    normalized = normalized.replace("Ⅱ", "II").replace("Ⅰ", "I")
    normalized = re.sub(r"\|\s*\|", "II", normalized)
    normalized = normalized.replace("|", "I")
    normalized = re.sub(r"\s+", "", normalized)
    return normalized


def _paragraph_style_id(paragraph) -> str | None:
    properties = paragraph.find(f"{W}pPr")
    if properties is None:
        return None
    style = properties.find(f"{W}pStyle")
    return style.get(f"{W}val") if style is not None else None


def _is_heading_one(paragraph, heading_style_ids: set[str]) -> bool:
    return (
        paragraph.tag == f"{W}p"
        and _paragraph_style_id(paragraph) in heading_style_ids
        and bool(_normalize_text(_element_text(paragraph)))
    )


def _load_document_root(path: Path):
    with ZipFile(path) as package:
        return _parse_xml(package, DOCUMENT_XML)


def _heading_one_style_ids(path: Path) -> set[str]:
    style_ids = set(HEADING_ONE_STYLE_IDS)
    with ZipFile(path) as package:
        if STYLES_XML not in package.namelist():
            return style_ids
        styles = _parse_xml(package, STYLES_XML)
    namespace = {"w": W[1:-1]}
    for style in styles.xpath("./w:style", namespaces=namespace):
        names = style.xpath("./w:name/@w:val", namespaces=namespace)
        style_id = style.get(f"{W}styleId")
        if not names or not style_id:
            continue
        normalized = re.sub(r"\s+", "", names[0]).lower()
        if normalized in {"heading1", "标题1"}:
            style_ids.add(style_id)
    return style_ids


def _heading_anchors(root, heading_style_ids: set[str]) -> list[tuple[int, str]]:
    return [
        (index, _normalize_text(_element_text(child)))
        for index, child in enumerate(_document_body(root))
        if _is_heading_one(child, heading_style_ids)
    ]


def discover_question_parts(
    source_docx: str | Path,
    *,
    expected_count: int | None = 14,
) -> list[QuestionPartSpec]:
    source = Path(source_docx)
    root = _load_document_root(source)
    anchors = _heading_anchors(root, _heading_one_style_ids(source))
    body_children = list(_document_body(root))
    specs: list[QuestionPartSpec] = []
    current_category: str | None = None
    pending_category_start: int | None = None

    for position, (index, title) in enumerate(anchors):
        if CATEGORY_TITLE_PATTERN.search(title):
            current_category = title
            pending_category_start = index
            continue
        if current_category is None:
            raise ValueError(f"套卷标题前缺少板块标题: {title}")
        next_index = anchors[position + 1][0] if position + 1 < len(anchors) else len(body_children)
        while next_index > index and body_children[next_index - 1].tag == f"{W}sectPr":
            next_index -= 1
        start = pending_category_start if pending_category_start is not None else index
        specs.append(
            QuestionPartSpec(
                sequence=len(specs) + 1,
                category=current_category,
                title=_safe_title(title),
                start=start,
                end=next_index,
            )
        )
        pending_category_start = None

    if expected_count is not None and len(specs) != expected_count:
        raise ValueError(f"题目目录应拆出 {expected_count} 份，实际为 {len(specs)} 份")
    return specs


def _question_ids(children: list) -> list[str]:
    ids: list[str] = []
    for child in children:
        if child.tag != f"{W}p":
            continue
        match = QUESTION_ID_PATTERN.match(_normalize_text(_element_text(child)))
        if match and match.group(1) not in ids:
            ids.append(match.group(1))
    if not ids:
        raise ValueError("套卷切片中未识别到题号")
    numbers = [int(value) for value in ids]
    expected = list(range(numbers[0], numbers[-1] + 1))
    if numbers != expected:
        raise ValueError(f"套卷题号不连续: 实际 {ids}")
    return ids


def _question_groups(children: list) -> tuple[tuple[str, ...], ...]:
    """按“完成1~5题”提示还原一次 F1 对应的阅读题组。"""

    question_ids = _question_ids(children)
    question_id_set = set(question_ids)
    groups: list[tuple[str, ...]] = []
    for child in children:
        if child.tag != f"{W}p":
            continue
        text = _normalize_text(_element_text(child))
        for match in READING_GROUP_PATTERN.finditer(text):
            start = int(match.group(1))
            end = int(match.group(2))
            if end < start:
                raise ValueError(f"阅读题组范围倒置: {match.group(0)}")
            group = tuple(str(number) for number in range(start, end + 1))
            if not set(group).issubset(question_id_set):
                continue
            groups.append(group)

    if not groups:
        return tuple((question_id,) for question_id in question_ids)

    flattened = [question_id for group in groups for question_id in group]
    if flattened != question_ids:
        raise ValueError(
            "阅读题组提示未完整覆盖套卷题号: "
            f"题号 {question_ids}，题组 {groups}"
        )
    return tuple(groups)


def inspect_question_document(
    source_docx: str | Path,
    *,
    expected_count: int | None = 14,
) -> list[dict]:
    """只读分析题目目录，供 answers-only 模式沿用 01/02 编号。"""

    source = Path(source_docx)
    root = _load_document_root(source)
    children = list(_document_body(root))
    results: list[dict] = []
    for spec in discover_question_parts(source, expected_count=expected_count):
        selected = children[spec.start : spec.end]
        results.append(
            {
                "spec": spec,
                "question_ids": _question_ids(selected),
                "question_groups": _question_groups(selected),
            }
        )
    return results


def split_question_document(
    source_docx: str | Path,
    output_dir: str | Path,
    *,
    expected_count: int | None = 14,
) -> list[dict]:
    source = Path(source_docx)
    output_root = Path(output_dir)
    source_hash = _digest(source)
    source_root = _load_document_root(source)
    source_body = _document_body(source_root)
    children = list(source_body)
    section = source_body.find(f"{W}sectPr")
    if section is None:
        raise ValueError("原题缺少节属性")
    source_margins = _page_margins_twips(source_root)
    specs = discover_question_parts(source, expected_count=expected_count)
    results: list[dict] = []

    for spec in specs:
        selected = children[spec.start : spec.end]
        target_root = deepcopy(source_root)
        target_body = _document_body(target_root)
        for child in list(target_body):
            target_body.remove(child)
        for child in selected:
            target_body.append(deepcopy(child))
        target_body.append(deepcopy(section))

        output = output_root / spec.filename
        _write_docx_package(source, output, target_root)
        _assert_relationships_resolve(output)
        written_root = _load_document_root(output)
        if _page_margins_twips(written_root) != source_margins:
            raise ValueError(f"{output.name} 页边距与原题不一致")
        if _canonical_children(written_root) != [
            ET.tostring(child, method="c14n") for child in selected
        ]:
            raise ValueError(f"{output.name} 正文 XML 与原文切片不一致")
        output_texts = [
            _normalize_text(_element_text(child))
            for child in _document_body(written_root)
            if child.tag == f"{W}p" and _normalize_text(_element_text(child))
        ]
        selected_first = _normalize_text(_element_text(selected[0]))
        expected_first = (
            spec.category
            if CATEGORY_TITLE_PATTERN.search(selected_first)
            else spec.title
        )
        if output_texts[0] != expected_first:
            raise ValueError(f"{output.name} 开头标题异常: {output_texts[0]}")
        results.append(
            {
                "spec": spec,
                "path": output,
                "question_ids": _question_ids(selected),
                "question_groups": _question_groups(selected),
                "sha256": _digest(output),
            }
        )

    if _digest(source) != source_hash:
        raise ValueError("题目拆分过程中原题哈希发生变化")
    return results


def _discover_answer_entries(children: list, start: int, end: int):
    markers: list[tuple[int, str]] = []
    for index in range(start, end):
        child = children[index]
        if child.tag != f"{W}p":
            continue
        match = ANSWER_START_PATTERN.match(_normalize_text(_element_text(child)))
        if match:
            markers.append((index, match.group(1)))
    if not markers:
        raise ValueError("答案标题下未识别到题号")

    accepted = [markers[0]]
    ignored: list[str] = []
    expected = int(markers[0][1]) + 1
    for marker_position, marker in enumerate(markers[1:], 1):
        number = int(marker[1])
        if number != expected:
            ignored.extend(value for _, value in markers[marker_position:])
            break
        accepted.append(marker)
        expected += 1

    ranges: list[tuple[int, int]] = []
    accepted_end = markers[len(accepted)][0] if len(markers) > len(accepted) else end
    for position, (index, _) in enumerate(accepted):
        next_index = accepted[position + 1][0] if position + 1 < len(accepted) else accepted_end
        ranges.append((index, next_index))
    return tuple(value for _, value in accepted), tuple(ranges), tuple(ignored)


def discover_answer_sections(source_docx: str | Path) -> list[AnswerSection]:
    source = Path(source_docx)
    root = _load_document_root(source)
    children = list(_document_body(root))
    anchors = _heading_anchors(root, _heading_one_style_ids(source))
    sections: list[AnswerSection] = []
    for position, (index, title) in enumerate(anchors):
        end = anchors[position + 1][0] if position + 1 < len(anchors) else len(children)
        while end > index and children[end - 1].tag == f"{W}sectPr":
            end -= 1
        source_ids, ranges, ignored = _discover_answer_entries(children, index + 1, end)
        sections.append(
            AnswerSection(
                title=title,
                normalized_title=normalize_exam_title(title),
                family="",
                start=index,
                end=end,
                source_ids=source_ids,
                entry_ranges=ranges,
                ignored_trailing_ids=ignored,
            )
        )
    return sections


def match_answer_sections(
    question_results: list[dict],
    answer_sections: list[AnswerSection],
) -> list[dict]:
    """按题目目录向后匹配同标题、同原题号的答案段。

    板块标题不参与语义判断。缺少某套答案时只跳过该套，游标不前移，
    因而后续已有答案仍可继续匹配。
    """

    matches: list[dict] = []
    answer_cursor = 0
    for result in question_results:
        spec: QuestionPartSpec = result["spec"]
        normalized_title = normalize_exam_title(spec.title)
        expected_ids = tuple(result["question_ids"])
        title_candidates = [
            (index, section)
            for index, section in enumerate(answer_sections[answer_cursor:], answer_cursor)
            if section.normalized_title == normalized_title
        ]
        exact_candidates = [
            (index, section)
            for index, section in title_candidates
            if section.source_ids == expected_ids
        ]
        if exact_candidates:
            matched_index, section = exact_candidates[0]
            answer_cursor = matched_index + 1
            status = "matched"
            reason = ""
        else:
            section = None
            status = "skipped"
            if title_candidates:
                found = ", ".join(
                    "/".join(candidate.source_ids) for _, candidate in title_candidates
                )
                reason = (
                    f"同名答案的原题号为 {found}，题目原题号为 "
                    f"{'/'.join(expected_ids)}"
                )
            else:
                reason = "答案目录中未找到后续同名套卷"
        matches.append(
            {
                "sequence": spec.sequence,
                "question": result,
                "answer_section": section,
                "status": status,
                "reason": reason,
            }
        )
    return matches


def _publish_clean_answer(
    source_docx: str | Path,
    output_dir: str | Path,
    match: dict,
) -> dict:
    """单份答案先在临时目录通过门禁，再发布到最终目录。"""

    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="future_chinese_answer_") as temp_dir:
        staged = clean_answer_section(source_docx, temp_dir, match)
        staged_doc = Path(staged["path"])
        staged_status = Path(derive_review_status_path(staged_doc))
        final_doc = output_root / staged_doc.name
        final_status = Path(derive_review_status_path(final_doc))
        os.replace(staged_doc, final_doc)
        os.replace(staged_status, final_status)
        staged["path"] = final_doc
        return staged


def _append_inline_payload(paragraph, element) -> None:
    if element is None:
        return
    if element.tag == f"{W}p":
        for child in element:
            if child.tag != f"{W}pPr":
                paragraph.append(deepcopy(child))
        return
    text = _normalize_text("".join(element.itertext()))
    if text:
        run = ET.SubElement(paragraph, f"{W}r")
        node = ET.SubElement(run, f"{W}t")
        _set_text(node, text)


def _renumber_answer_entry(children: list, start: int, end: int, sequence: int) -> list:
    source = children[start:end]
    if not source or source[0].tag != f"{W}p":
        raise ValueError("答案块开头不是段落")
    raw = _element_text(source[0])
    marker = ANSWER_START_PATTERN.match(raw)
    if marker is None:
        raise ValueError(f"答案块缺少题号: {_normalize_text(raw)}")
    reference = source[0]
    first = _new_text_paragraph(reference, f"{sequence}．")
    payload = _slice_paragraph(reference, marker.start(2), len(raw))
    _append_inline_payload(first, payload)
    output = [first]
    analysis_found = False

    for child in source[1:]:
        if child.tag != f"{W}p":
            output.append(deepcopy(child))
            continue
        text = _element_text(child)
        analysis = ANALYSIS_PATTERN.match(text)
        if analysis is None:
            output.append(deepcopy(child))
            continue
        normalized = _new_text_paragraph(child, "解析：")
        payload = _slice_paragraph(child, analysis.start(1), len(text))
        _append_inline_payload(normalized, payload)
        output.append(normalized)
        analysis_found = True

    if not analysis_found:
        output.append(_new_text_paragraph(reference, "解析："))
    return output


def _grouped_answer_entry(
    children: list,
    start: int,
    end: int,
    subquestion_sequence: int,
) -> tuple[list, list, object]:
    """拆出一个源答案的答案段与解析段，并改为独立 F4 小问段。"""

    source = children[start:end]
    if not source or source[0].tag != f"{W}p":
        raise ValueError("答案块开头不是段落")
    raw = _element_text(source[0])
    marker = ANSWER_START_PATTERN.match(raw)
    if marker is None:
        raise ValueError(f"答案块缺少题号: {_normalize_text(raw)}")

    reference = source[0]
    answer_first = _new_text_paragraph(reference, f"（{subquestion_sequence}）")
    answer_payload = _slice_paragraph(reference, marker.start(2), len(raw))
    _append_inline_payload(answer_first, answer_payload)
    answer_nodes = [answer_first]
    analysis_nodes: list = []
    analysis_reference = reference
    in_analysis = False

    for child in source[1:]:
        if not in_analysis and child.tag == f"{W}p":
            text = _element_text(child)
            analysis = ANALYSIS_PATTERN.match(text)
            if analysis is not None:
                in_analysis = True
                analysis_reference = child
                normalized = _new_text_paragraph(
                    child,
                    f"（{subquestion_sequence}）",
                )
                payload = _slice_paragraph(child, analysis.start(1), len(text))
                _append_inline_payload(normalized, payload)
                analysis_nodes.append(normalized)
                continue
        target = analysis_nodes if in_analysis else answer_nodes
        target.append(deepcopy(child))

    if not in_analysis:
        analysis_nodes.append(
            _new_text_paragraph(reference, f"（{subquestion_sequence}）")
        )
    return answer_nodes, analysis_nodes, analysis_reference


def _group_answer_entries(
    children: list,
    ranges: list[tuple[int, int]],
    group_sequence: int,
) -> list:
    """输出“顶层题号 → 全部 F4 答案 → 单一 F3 解析块”。"""

    if not ranges:
        raise ValueError("阅读题组没有答案块")
    first_reference = children[ranges[0][0]]
    output = [_new_text_paragraph(first_reference, f"{group_sequence}．")]
    grouped_answers: list[list] = []
    grouped_analyses: list[list] = []
    analysis_reference = first_reference
    for subquestion_sequence, (start, end) in enumerate(ranges, 1):
        answer_nodes, analysis_nodes, current_reference = _grouped_answer_entry(
            children,
            start,
            end,
            subquestion_sequence,
        )
        grouped_answers.append(answer_nodes)
        grouped_analyses.append(analysis_nodes)
        if subquestion_sequence == 1:
            analysis_reference = current_reference

    for answer_nodes in grouped_answers:
        output.extend(answer_nodes)
    output.append(_new_text_paragraph(analysis_reference, "解析："))
    for analysis_nodes in grouped_analyses:
        output.extend(analysis_nodes)
    return output


def clean_answer_section(
    source_docx: str | Path,
    output_dir: str | Path,
    match: dict,
) -> dict:
    if match["status"] != "matched":
        raise ValueError(match["reason"])
    source = Path(source_docx)
    output_root = Path(output_dir)
    section: AnswerSection = match["answer_section"]
    spec: QuestionPartSpec = match["question"]["spec"]
    source_root = _load_document_root(source)
    source_body = _document_body(source_root)
    source_children = list(source_body)
    section_properties = source_body.find(f"{W}sectPr")
    if section_properties is None:
        raise ValueError("总答案缺少节属性")

    target_root = deepcopy(source_root)
    target_body = _document_body(target_root)
    for child in list(target_body):
        target_body.remove(child)
    range_by_source_id = dict(zip(section.source_ids, section.entry_ranges))
    question_groups: tuple[tuple[str, ...], ...] = match["question"][
        "question_groups"
    ]
    for sequence, group in enumerate(question_groups, 1):
        try:
            ranges = [range_by_source_id[question_id] for question_id in group]
        except KeyError as exc:
            raise ValueError(f"阅读题组缺少答案题号: {exc.args[0]}") from exc
        if len(group) == 1:
            start, end = ranges[0]
            output_children = _renumber_answer_entry(
                source_children,
                start,
                end,
                sequence,
            )
        else:
            output_children = _group_answer_entries(
                source_children,
                ranges,
                sequence,
            )
        for child in output_children:
            target_body.append(child)
    target_body.append(deepcopy(section_properties))

    output = output_root / spec.answer_filename
    _write_docx_package(source, output, target_root)
    _assert_relationships_resolve(output)
    written_root = _load_document_root(output)
    if _page_margins_twips(written_root) != _page_margins_twips(source_root):
        raise ValueError(f"{output.name} 页边距与总答案不一致")
    parsed = build_answer_units_from_docx(output, preserve_source_positions=True)
    expected_ids = [str(index) for index in range(1, len(question_groups) + 1)]
    actual_ids = [unit.question_id for unit in parsed]
    if actual_ids != expected_ids:
        raise ValueError(f"{output.name} 清洗后题号异常: 期望 {expected_ids}，实际 {actual_ids}")
    for unit, group in zip(parsed, question_groups):
        if len(group) == 1:
            continue
        if unit.answer_mode != "subquestion":
            raise ValueError(f"{output.name} 第 {unit.question_id} 组未识别为 F4 小问")
        if len(unit.answer_items) != len(group):
            raise ValueError(
                f"{output.name} 第 {unit.question_id} 组 F4 数量异常: "
                f"期望 {len(group)}，实际 {len(unit.answer_items)}"
            )
    flags = [(unit.question_id, list(unit.review_flags)) for unit in parsed if unit.review_flags]
    if flags:
        raise ValueError(f"{output.name} 清洗后仍有审核标记: {flags}")
    initialize_review_status(output)
    update_review_status(
        output,
        status="approved",
        reviewer="Codex offline validation",
        note=(
            "未来高三语文：按题目目录编号，完整阅读题按一次 F1 对应的范围合并，"
            "组内答案连续 F4、整组解析单次 F3；套卷标题、题答数量、答案从 1 重排、"
            "解析边界、DOCX 结构与页边距已离线核验；Windows WPS 实机录入仍需人工确认。"
        ),
    )
    gate = get_review_gate_result(output)
    if not gate["allowed"]:
        raise ValueError(f"{output.name} 审核门禁未放行: {gate['reason']}")
    return {
        "path": output,
        "answer_ids": expected_ids,
        "question_groups": [list(group) for group in question_groups],
        "f2_answer_count": sum(1 for group in question_groups if len(group) == 1),
        "f4_answer_count": sum(len(group) for group in question_groups if len(group) > 1),
        "f3_analysis_count": len(question_groups),
        "ignored_trailing_ids": list(section.ignored_trailing_ids),
        "sha256": _digest(output),
    }


def _find_question_source(project_root: Path) -> Path:
    questions = project_root / "清华书 3年真题 .docx"
    if not questions.is_file():
        raise FileNotFoundError(questions)
    return questions


def _find_answer_source(project_root: Path, answer_source: str | Path | None) -> Path:
    answers = (
        Path(answer_source).expanduser().resolve()
        if answer_source is not None
        else project_root / "答案" / "语文清华书参考答案.docx"
    )
    if not answers.is_file():
        raise FileNotFoundError(answers)
    return answers


def process_project(
    project_root: str | Path,
    *,
    questions_only: bool = False,
    answers_only: bool = False,
    answer_source: str | Path | None = None,
) -> dict:
    if questions_only and answers_only:
        raise ValueError("questions_only 与 answers_only 不能同时启用")
    root = Path(project_root).resolve()
    question_source = _find_question_source(root)
    resolved_answer_source = (
        None if questions_only else _find_answer_source(root, answer_source)
    )
    source_hashes = {"question": _digest(question_source)}
    if resolved_answer_source is not None:
        source_hashes["answer"] = _digest(resolved_answer_source)

    if answers_only:
        question_results = inspect_question_document(question_source)
        question_documents_written = 0
    else:
        question_results = split_question_document(
            question_source,
            root / "题目（已拆分）",
        )
        question_documents_written = len(question_results)
    summary = {
        "project_root": str(root),
        "question_documents": len(question_results),
        "question_documents_written": question_documents_written,
        "question_order": [result["spec"].filename for result in question_results],
        "source_hashes_unchanged": False,
        "answer_documents": 0,
        "answer_preflight": [],
    }
    if not questions_only:
        assert resolved_answer_source is not None
        answer_sections = discover_answer_sections(resolved_answer_source)
        matches = match_answer_sections(question_results, answer_sections)
        answer_results: list[dict] = []
        for item in matches:
            record = {
                "sequence": item["sequence"],
                "question_file": item["question"]["spec"].filename,
                "status": item["status"],
                "reason": item["reason"],
                "source_answer_title": (
                    item["answer_section"].title if item["answer_section"] else None
                ),
            }
            if item["status"] == "matched":
                try:
                    answer_results.append(
                        _publish_clean_answer(
                            resolved_answer_source,
                            root / "答案" / "按题目目录拆分",
                            item,
                        )
                    )
                except Exception as exc:  # 单份失败不阻断后续套卷
                    record["status"] = "failed"
                    record["reason"] = str(exc)
            summary["answer_preflight"].append(record)
        summary["answer_documents"] = len(answer_results)

    current_hashes = {"question": _digest(question_source)}
    if resolved_answer_source is not None:
        current_hashes["answer"] = _digest(resolved_answer_source)
    summary["source_hashes_unchanged"] = source_hashes == current_hashes
    if not summary["source_hashes_unchanged"]:
        raise ValueError("处理过程中原题或原答案哈希发生变化")
    return summary


def main() -> int:
    parser = argparse.ArgumentParser(
        description="按题目目录顺序拆分未来高三语文，并用相同 01/02 编号清洗答案。"
    )
    parser.add_argument("project_root", help="未来-高三-语文业务目录")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--questions-only",
        action="store_true",
        help="只生成 01-14 题目拆分，不生成答案",
    )
    mode.add_argument(
        "--answers-only",
        action="store_true",
        help="只分析题目目录并生成当前能够对应的答案",
    )
    parser.add_argument(
        "--answer-source",
        help="显式指定答案 DOCX；省略时使用项目答案目录中的默认文件",
    )
    args = parser.parse_args()
    summary = process_project(
        args.project_root,
        questions_only=args.questions_only,
        answers_only=args.answers_only,
        answer_source=args.answer_source,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
