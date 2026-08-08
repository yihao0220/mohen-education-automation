from __future__ import annotations

import argparse
from copy import deepcopy
from dataclasses import dataclass, field
from hashlib import sha256
import json
from pathlib import Path
import re
import sys
from zipfile import ZipFile

from lxml import etree as ET

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core import (
    build_answer_units_from_docx,
    build_question_units_from_docx,
    get_review_gate_result,
    initialize_review_status,
    update_review_status,
)
from shared_core.answer_core import infer_grouped_question_ids
from shared_core.subject_overlay import is_question_input_excluded_for_context
from shared_core.subject_overlay import extract_additional_question_id_for_context
from tools.process_guanmei_geography import (
    DOCUMENT_XML,
    PAGE_MARGIN_TWIPS,
    W,
    _all_element_text,
    _assert_relationships_resolve,
    _canonical_children,
    _document_body,
    _element_text,
    _has_payload,
    _new_text_paragraph,
    _page_margins_twips,
    _parse_xml,
    _set_text,
    _slice_paragraph,
    _write_docx_package,
)


UNIT_ANCHOR_PATTERN = re.compile(
    r"^(?:"
    r"第[一二三四五六七八九十百]+框(?:\s|　)+.+|"
    r"能力提升练(?:\s.*)?|"
    r"专题专项练(?:对应主书\s*P?\d+)?|"
    r"五年高考练|"
    r"三年模拟练|"
    r"题型整合练"
    r")$"
)
OUTLINE_BOUNDARY_PATTERN = re.compile(
    r"^第[一二三四五六七八九十百]+(?:单元|课)(?:\s|　|$)"
)
QUESTION_ID_PATTERN = re.compile(r"^\s*(\d+)\s*[．.]")
CHOICE_CELL_PATTERN = re.compile(r"^\s*(\d+)\s*[．.]\s*([A-D])\s*$")
CHOICE_ANALYSIS_PATTERN = re.compile(
    r"^\s*(\d+)\s*[．.]\s*([A-D])(?:\s+|$)(.*)$"
)
SUBJECTIVE_ANSWER_PATTERN = re.compile(
    r"^\s*(\d+)\s*[．.]\s*答案\s*[:：]?\s*(.*)$"
)
ANALYSIS_MARKER_PATTERN = re.compile(r"^\s*解析\s*[:：]?\s*(.*)$")
COMPACT_CHOICE_PATTERN = re.compile(
    r"(\d+)\s*[．.]\s*([A-D])(?=\s*(?:\d+\s*[．.]\s*[A-D]|$))"
)
SPECIAL_SUBTITLE_PATTERN = re.compile(
    r"^(?:易混易错练(?:\s*[？?\d]+)?|专题强化练.*|专题专攻快速进阶)$"
)
ANSWER_NOISE_PATTERN = re.compile(
    r"(?:导师点睛|易混.*辨析|特别提醒|易错点拨|方法技巧|学习点拨|"
    r"知识拓展|易混易错|技巧|点拨|提醒|辨析)"
)
QUESTION_INPUT_HEADING_PATTERN = re.compile(
    r"^(?:基础过关练|考点\s*\d+.*|"
    r"题组[一二三四五六七八九十百\d]+(?:\s|　)+.+|"
    r"易错点\s*[一二三四五六七八九十百\d]+(?:\s|　)*.+)$"
)
HEADING_STYLE_ID = "1"
STANDALONE_FLOW_ARTIFACTS = {"续表"}


@dataclass(frozen=True)
class UnitSpec:
    sequence: int
    source_title: str
    output_title: str
    start: int
    end: int

    @property
    def question_filename(self) -> str:
        return f"{self.sequence:02d} {self.output_title}.docx"

    @property
    def answer_filename(self) -> str:
        return f"{self.sequence:02d} {self.output_title}-答案_已清洗.docx"


@dataclass
class AnswerEntry:
    choice: str | None = None
    answer_elements: list = field(default_factory=list)
    analysis_elements: list = field(default_factory=list)
    reference: object | None = None


def _digest(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def _safe_title(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]', " ", value)
    return _normalize_text(cleaned)


def _normalize_output_title(source_title: str, following_titles: list[str]) -> str:
    if source_title.startswith("能力提升练"):
        return "能力提升练"
    if source_title.startswith("专题专项练"):
        for candidate in following_titles:
            normalized = _normalize_text(candidate)
            if not SPECIAL_SUBTITLE_PATTERN.match(normalized):
                continue
            if normalized.startswith("易混易错练"):
                return "易混易错练"
            return _safe_title(normalized)
        return "专题专项练"
    return _safe_title(source_title)


def _paragraph_style_id(paragraph) -> str | None:
    properties = paragraph.find(f"{W}pPr")
    if properties is None:
        return None
    style = properties.find(f"{W}pStyle")
    return style.get(f"{W}val") if style is not None else None


def discover_units(source_docx: str | Path) -> list[UnitSpec]:
    source = Path(source_docx)
    with ZipFile(source) as package:
        root = _parse_xml(package, DOCUMENT_XML)
    children = list(_document_body(root))
    anchors: list[tuple[int, str]] = []
    for index, child in enumerate(children):
        if child.tag != f"{W}p":
            continue
        title = _normalize_text(_element_text(child))
        if UNIT_ANCHOR_PATTERN.match(title):
            anchors.append((index, title))

    specs: list[UnitSpec] = []
    for sequence, (start, title) in enumerate(anchors, 1):
        next_anchor = anchors[sequence][0] if sequence < len(anchors) else len(children)
        end = next_anchor
        following_titles: list[str] = []
        for child in children[start + 1 : next_anchor]:
            if child.tag != f"{W}p":
                continue
            text = _normalize_text(_element_text(child))
            if text and len(following_titles) < 5:
                following_titles.append(text)
            if OUTLINE_BOUNDARY_PATTERN.match(text):
                end = children.index(child)
                break
        while end > start and children[end - 1].tag == f"{W}sectPr":
            end -= 1
        specs.append(
            UnitSpec(
                sequence=sequence,
                source_title=title,
                output_title=_normalize_output_title(title, following_titles),
                start=start,
                end=end,
            )
        )
    return specs


def _set_two_cm_margins(root) -> None:
    body = _document_body(root)
    section = body.find(f"{W}sectPr")
    if section is None:
        raise ValueError("文档缺少节属性")
    page_margin = section.find(f"{W}pgMar")
    if page_margin is None:
        page_margin = ET.SubElement(section, f"{W}pgMar")
    for name in ("top", "bottom", "left", "right"):
        page_margin.set(f"{W}{name}", str(PAGE_MARGIN_TWIPS))


def _is_blank_page_break_paragraph(element) -> bool:
    if element.tag != f"{W}p":
        return False
    page_breaks = element.xpath('.//w:br[@w:type="page"]', namespaces={"w": W[1:-1]})
    if not page_breaks:
        return False
    if _normalize_text("".join(element.itertext())):
        return False
    return not any(
        ET.QName(node).localname in {"drawing", "pict", "object"}
        for node in element.iter()
    )


def _strip_trailing_blank_page_breaks(elements: list) -> tuple[list, int]:
    """删除会在拆分文档末尾制造空白页的纯分页符段落。"""

    kept = list(elements)
    removed = 0
    index = len(kept) - 1
    while index >= 0:
        element = kept[index]
        if _is_blank_page_break_paragraph(element):
            del kept[index]
            removed += 1
            index -= 1
            continue
        if (
            element.tag == f"{W}p"
            and not _normalize_text("".join(element.itertext()))
            and not element.xpath(".//w:br", namespaces={"w": W[1:-1]})
        ):
            index -= 1
            continue
        break
    return kept, removed


def _logical_question_ids(root) -> list[str]:
    ids: list[str] = []
    for child in _document_body(root):
        if child.tag != f"{W}p":
            continue
        text = _normalize_text(_element_text(child))
        match = QUESTION_ID_PATTERN.match(text)
        if match:
            ids.append(match.group(1))
            continue
        additional = extract_additional_question_id_for_context(
            text,
            "future_politics",
        )
        if additional:
            ids.append(additional)
    unique: list[str] = []
    for qid in ids:
        if qid not in unique:
            unique.append(qid)
    if not unique:
        raise ValueError("拆分单元未识别到题号")
    numbers = list(map(int, unique))
    expected = [str(index) for index in range(min(numbers), max(numbers) + 1)]
    if unique != expected:
        raise ValueError(f"题号不连续: 期望 {expected}，实际 {unique}")
    return unique


def split_question_document(
    source_docx: str | Path,
    output_dir: str | Path,
    *,
    expected_unit_count: int | None = 48,
) -> list[dict]:
    source = Path(source_docx)
    output_root = Path(output_dir)
    specs = discover_units(source)
    if expected_unit_count is not None and len(specs) != expected_unit_count:
        raise ValueError(
            f"题目拆分点应为 {expected_unit_count} 个，实际为 {len(specs)}"
        )
    with ZipFile(source) as package:
        source_root = _parse_xml(package, DOCUMENT_XML)
    if _page_margins_twips(source_root) != (PAGE_MARGIN_TWIPS,) * 4:
        raise ValueError(
            "原题四边页边距不是 2 cm（1134 twips），为避免改版已停止拆分"
        )
    source_body = _document_body(source_root)
    children = list(source_body)
    section = source_body.find(f"{W}sectPr")
    if section is None:
        raise ValueError("原题缺少节属性")

    results: list[dict] = []
    for spec in specs:
        selected, removed_page_breaks = _strip_trailing_blank_page_breaks(
            children[spec.start : spec.end]
        )
        target_root = deepcopy(source_root)
        target_body = _document_body(target_root)
        for child in list(target_body):
            target_body.remove(child)
        for child in selected:
            target_body.append(deepcopy(child))
        target_body.append(deepcopy(section))
        _set_two_cm_margins(target_root)

        output = output_root / spec.question_filename
        _write_docx_package(source, output, target_root)
        _assert_relationships_resolve(output)
        with ZipFile(output) as package:
            written_root = _parse_xml(package, DOCUMENT_XML)
        if _page_margins_twips(written_root) != (PAGE_MARGIN_TWIPS,) * 4:
            raise ValueError(f"{output.name} 四边页边距不是 2 cm")
        if _canonical_children(written_root) != [
            ET.tostring(child, method="c14n") for child in selected
        ]:
            raise ValueError(f"{output.name} 正文 XML 与原文切片不一致")
        logical_ids = _logical_question_ids(written_root)
        results.append(
            {
                "spec": spec,
                "path": output,
                "logical_question_ids": logical_ids,
                "removed_trailing_page_breaks": removed_page_breaks,
                "sha256": _digest(output),
            }
        )
    return results


def _table_choice_pairs(table) -> list[tuple[str, str]]:
    pairs: list[tuple[str, str]] = []
    for cell in table.iter(f"{W}tc"):
        text = _normalize_text(_all_element_text(cell))
        match = CHOICE_CELL_PATTERN.match(text)
        if match:
            pairs.append((match.group(1), match.group(2)))
    return pairs


def _compact_choice_pairs(text: str) -> list[tuple[str, str]]:
    pairs = COMPACT_CHOICE_PATTERN.findall(_normalize_text(text))
    return pairs if len(pairs) >= 2 else []


def _slice_payload(paragraph, match_end: int, text: str):
    start = match_end
    while start < len(text) and text[start].isspace():
        start += 1
    if start >= len(text):
        return None
    return _slice_paragraph(paragraph, start, len(text))


def _append_payload(target: list, element) -> None:
    if element is not None and _has_payload(element):
        target.append(deepcopy(element))


def parse_answer_unit(children: list) -> dict[str, AnswerEntry]:
    entries: dict[str, AnswerEntry] = {}
    current_qid: str | None = None
    state: str | None = None
    ignoring = False

    def entry(qid: str, reference=None) -> AnswerEntry:
        value = entries.setdefault(qid, AnswerEntry())
        if value.reference is None and reference is not None:
            value.reference = reference
        return value

    for child in children:
        if child.tag == f"{W}tbl":
            pairs = _table_choice_pairs(child)
            if pairs:
                for qid, answer in pairs:
                    entry(qid).choice = answer
                current_qid = None
                state = None
                ignoring = False
                continue
            if current_qid and state and not ignoring:
                target = (
                    entry(current_qid).answer_elements
                    if state == "answer"
                    else entry(current_qid).analysis_elements
                )
                _append_payload(target, child)
            continue
        if child.tag != f"{W}p":
            continue

        raw_text = _element_text(child)
        text = _normalize_text(raw_text)
        if not text:
            continue
        if text in STANDALONE_FLOW_ARTIFACTS:
            continue
        compact = _compact_choice_pairs(text)
        if compact:
            for qid, answer in compact:
                entry(qid, child).choice = answer
            current_qid = None
            state = None
            ignoring = False
            continue

        choice_match = CHOICE_ANALYSIS_PATTERN.match(raw_text)
        if choice_match:
            qid, answer = choice_match.group(1), choice_match.group(2)
            value = entry(qid, child)
            if value.choice is None:
                value.choice = answer
            payload = _slice_payload(child, choice_match.start(3), raw_text)
            _append_payload(value.analysis_elements, payload)
            current_qid = qid
            state = "analysis"
            ignoring = False
            continue

        answer_match = SUBJECTIVE_ANSWER_PATTERN.match(raw_text)
        if answer_match:
            qid = answer_match.group(1)
            value = entry(qid, child)
            payload = _slice_payload(child, answer_match.start(2), raw_text)
            _append_payload(value.answer_elements, payload)
            current_qid = qid
            state = "answer"
            ignoring = False
            continue

        analysis_match = ANALYSIS_MARKER_PATTERN.match(raw_text)
        if analysis_match and current_qid:
            value = entry(current_qid, child)
            payload = _slice_payload(child, analysis_match.start(1), raw_text)
            _append_payload(value.analysis_elements, payload)
            state = "analysis"
            ignoring = False
            continue

        if ANSWER_NOISE_PATTERN.search(text):
            current_qid = None
            state = None
            ignoring = True
            continue
        if (
            _paragraph_style_id(child) == HEADING_STYLE_ID
            and current_qid
            and entry(current_qid).choice is not None
            and not entry(current_qid).answer_elements
        ):
            current_qid = None
            state = None
            ignoring = True
            continue
        if ignoring:
            continue
        if current_qid and state:
            target = (
                entry(current_qid).answer_elements
                if state == "answer"
                else entry(current_qid).analysis_elements
            )
            _append_payload(target, child)

    return entries


def _append_inline_element(paragraph, element) -> None:
    if element.tag == f"{W}p":
        for child in element:
            if child.tag != f"{W}pPr":
                paragraph.append(deepcopy(child))
        return
    text = _table_text(element) if element.tag == f"{W}tbl" else _normalize_text(
        _all_element_text(element)
    )
    if text:
        run = ET.SubElement(paragraph, f"{W}r")
        node = ET.SubElement(run, f"{W}t")
        _set_text(node, text)


def _table_text(table) -> str:
    rows: list[str] = []
    for row in table.findall(f"{W}tr"):
        cells = [
            _normalize_text(_all_element_text(cell))
            for cell in row.findall(f"{W}tc")
        ]
        rows.append(" | ".join(cell for cell in cells if cell))
    return "；".join(row for row in rows if row)


def _payload_as_paragraph(payload, reference):
    if payload.tag == f"{W}p":
        return deepcopy(payload)
    text = _table_text(payload) if payload.tag == f"{W}tbl" else _normalize_text(
        _all_element_text(payload)
    )
    return _new_text_paragraph(reference, text)


def _fallback_reference(entries: dict[str, AnswerEntry]):
    for value in entries.values():
        if value.reference is not None:
            return value.reference
        for element in value.answer_elements + value.analysis_elements:
            if element.tag == f"{W}p":
                return element
    return None


def _validate_entry_ids(
    entries: dict[str, AnswerEntry],
    expected_count: int,
    context: str,
) -> list[str]:
    actual = sorted(entries, key=int)
    expected = [str(index) for index in range(1, expected_count + 1)]
    if actual != expected:
        raise ValueError(
            f"{context} 答案题号不连续: 期望 {expected}，实际 {actual}"
        )
    missing = [
        qid
        for qid, value in entries.items()
        if value.choice is None and not value.answer_elements
    ]
    if missing:
        raise ValueError(f"{context} 缺少答案内容: {missing}")
    return actual


def clean_answer_unit(
    source_docx: Path,
    source_root,
    answer_spec: UnitSpec,
    question_result: dict,
    output_dir: Path,
) -> dict:
    source_body = _document_body(source_root)
    source_children = list(source_body)
    selected = source_children[answer_spec.start : answer_spec.end]
    entries = parse_answer_unit(selected)
    source_question_ids = question_result["logical_question_ids"]
    answer_ids = _validate_entry_ids(
        entries,
        len(source_question_ids),
        answer_spec.output_title,
    )

    question_units = build_question_units_from_docx(
        question_result["path"],
        grade_hint="高二",
    )
    groups = [infer_grouped_question_ids(unit) for unit in question_units]
    covered_source_ids = [qid for group in groups for qid in group]
    if covered_source_ids != source_question_ids:
        raise ValueError(
            f"{question_result['path'].name} F1 分组未完整覆盖题号: "
            f"期望 {source_question_ids}，实际 {covered_source_ids}"
        )
    answer_id_by_source = dict(zip(source_question_ids, answer_ids))
    answer_groups = [
        [answer_id_by_source[qid] for qid in group]
        for group in groups
    ]

    target_root = deepcopy(source_root)
    target_body = _document_body(target_root)
    for child in list(target_body):
        target_body.remove(child)
    fallback = _fallback_reference(entries)
    placeholder_count = 0
    ignored_heading_count = sum(
        1
        for child in selected
        if child.tag == f"{W}p"
        and ANSWER_NOISE_PATTERN.search(_normalize_text(_element_text(child)))
    )

    for sequence, group in enumerate(answer_groups, 1):
        reference = (
            entries[group[0]].reference
            if entries[group[0]].reference is not None
            else fallback
        )
        if len(group) > 1:
            if any(entries[qid].choice is None for qid in group):
                raise ValueError(
                    f"{answer_spec.output_title} 材料题组 {group} 含非选择题，无法安全合并"
                )
            target_body.append(_new_text_paragraph(reference, f"{sequence}．"))
            for index, qid in enumerate(group, 1):
                target_body.append(
                    _new_text_paragraph(reference, f"（{index}）{entries[qid].choice}")
                )
            analysis = _new_text_paragraph(reference, "解析：")
            for index, qid in enumerate(group, 1):
                marker_run = ET.SubElement(analysis, f"{W}r")
                marker_text = ET.SubElement(marker_run, f"{W}t")
                _set_text(marker_text, f"（{index}）")
                payloads = entries[qid].analysis_elements
                if not payloads:
                    placeholder_count += 1
                for payload in payloads:
                    _append_inline_element(analysis, payload)
            target_body.append(analysis)
            continue

        qid = group[0]
        value = entries[qid]
        if value.choice is not None:
            target_body.append(
                _new_text_paragraph(reference, f"{sequence}．{value.choice}")
            )
        else:
            target_body.append(_new_text_paragraph(reference, f"{sequence}．"))
            for payload in value.answer_elements:
                target_body.append(_payload_as_paragraph(payload, reference))
        target_body.append(_new_text_paragraph(reference, "解析："))
        if value.analysis_elements:
            for payload in value.analysis_elements:
                target_body.append(_payload_as_paragraph(payload, reference))
        else:
            placeholder_count += 1

    section = source_body.find(f"{W}sectPr")
    if section is None:
        raise ValueError("总答案缺少节属性")
    target_body.append(deepcopy(section))
    _set_two_cm_margins(target_root)

    output = output_dir / question_result["spec"].answer_filename
    _write_docx_package(source_docx, output, target_root)
    _assert_relationships_resolve(output)
    with ZipFile(output) as package:
        written_root = _parse_xml(package, DOCUMENT_XML)
    if _page_margins_twips(written_root) != (PAGE_MARGIN_TWIPS,) * 4:
        raise ValueError(f"{output.name} 四边页边距不是 2 cm")
    answer_units = build_answer_units_from_docx(output, preserve_source_positions=True)
    parsed_ids = [unit.question_id for unit in answer_units]
    expected_output_ids = [str(index) for index in range(1, len(groups) + 1)]
    if parsed_ids != expected_output_ids:
        raise ValueError(
            f"{output.name} 清洗后答案块异常: 期望 {expected_output_ids}，实际 {parsed_ids}"
        )
    flags = [
        (unit.question_id, list(unit.review_flags))
        for unit in answer_units
        if unit.review_flags
    ]
    if flags:
        raise ValueError(f"{output.name} 清洗后仍有审核标记: {flags}")

    initialize_review_status(output)
    update_review_status(
        output,
        status="approved",
        reviewer="Codex offline validation",
        note=(
            "未来高二政治：题号、题答映射、干扰知识卡片、四边 2 cm 页边距、"
            "DOCX 结构与资源关系已离线核验；Windows WPS 实机录入仍需人工确认。"
        ),
    )
    gate = get_review_gate_result(output)
    if not gate["allowed"]:
        raise ValueError(f"{output.name} 审核门禁未放行: {gate['reason']}")
    return {
        "path": output,
        "f1_blocks": len(question_units),
        "logical_question_ids": len(source_question_ids),
        "answer_blocks": len(answer_units),
        "placeholder_analysis": placeholder_count,
        "ignored_noise_headings": ignored_heading_count,
        "sha256": _digest(output),
    }


def _find_sources(project_root: Path) -> tuple[Path, Path]:
    question = project_root / "高中政治选必1.docx"
    if not question.is_file():
        raise FileNotFoundError(question)
    candidates = sorted(
        path
        for path in (project_root / "答案").glob("高中政治选必1答案-*.docx")
        if not path.name.startswith(".~") and "_已清洗" not in path.stem
    )
    if len(candidates) != 1:
        raise ValueError(f"总答案匹配数量应为 1，实际为 {len(candidates)}: {candidates}")
    return question, candidates[0]


def _validate_question_input_headings(question_results: list[dict]) -> int:
    observed = 0
    for result in question_results:
        with ZipFile(result["path"]) as package:
            root = _parse_xml(package, DOCUMENT_XML)
        for child in _document_body(root):
            if child.tag != f"{W}p":
                continue
            text = _normalize_text(_element_text(child))
            if not QUESTION_INPUT_HEADING_PATTERN.match(text):
                continue
            observed += 1
            if not is_question_input_excluded_for_context(text, "future_politics"):
                raise ValueError(f"题内标题未进入忽略规则: {text}")
    if observed == 0:
        raise ValueError("真实拆分题目中未复核到题组/易错点/基础过关标题")
    return observed


def process_project(project_root: str | Path) -> dict:
    root = Path(project_root).resolve()
    question_source, answer_source = _find_sources(root)
    source_hashes = {
        str(question_source): _digest(question_source),
        str(answer_source): _digest(answer_source),
    }
    question_results = split_question_document(
        question_source,
        root / "题目（已拆分）",
    )
    answer_specs = discover_units(answer_source)
    if len(answer_specs) != len(question_results):
        raise ValueError(
            f"题答拆分点数量不一致: 题目 {len(question_results)}，答案 {len(answer_specs)}"
        )
    with ZipFile(answer_source) as package:
        answer_root = _parse_xml(package, DOCUMENT_XML)
    answer_results = []
    for question_result, answer_spec in zip(question_results, answer_specs):
        if answer_spec.sequence != question_result["spec"].sequence:
            raise ValueError("题答拆分顺序不一致")
        answer_results.append(
            clean_answer_unit(
                answer_source,
                answer_root,
                answer_spec,
                question_result,
                root / "答案" / "按最小单元拆分",
            )
        )

    ignored_question_headings = _validate_question_input_headings(question_results)
    after_hashes = {
        str(question_source): _digest(question_source),
        str(answer_source): _digest(answer_source),
    }
    if after_hashes != source_hashes:
        raise ValueError("处理过程中原题或原答案哈希发生变化")
    return {
        "project_root": str(root),
        "question_source_sha256": source_hashes[str(question_source)],
        "answer_source_sha256": source_hashes[str(answer_source)],
        "question_documents": len(question_results),
        "answer_documents": len(answer_results),
        "logical_question_ids": sum(
            len(item["logical_question_ids"]) for item in question_results
        ),
        "f1_blocks": sum(item["f1_blocks"] for item in answer_results),
        "answer_blocks": sum(item["answer_blocks"] for item in answer_results),
        "placeholder_analysis": sum(
            item["placeholder_analysis"] for item in answer_results
        ),
        "ignored_answer_noise_headings": sum(
            item["ignored_noise_headings"] for item in answer_results
        ),
        "ignored_question_input_headings_verified": ignored_question_headings,
        "removed_trailing_question_page_breaks": sum(
            item["removed_trailing_page_breaks"] for item in question_results
        ),
        "margins_twips": [PAGE_MARGIN_TWIPS] * 4,
        "source_hashes_unchanged": True,
        "windows_wps_plugin_execution": "pending",
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="拆分未来高二政治总题目并清洗、拆分总答案。"
    )
    parser.add_argument("project_root", help="未来-高二-政治业务目录")
    args = parser.parse_args()
    summary = process_project(args.project_root)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
