from __future__ import annotations

import hashlib
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from .models import DocNode, QuestionUnit
from .subject_overlay import (
    classify_media_hashes_for_context,
    choose_strategy_for_context,
    detect_subject_overlay,
    get_subject_overlay,
    is_leading_context_group_boundary_for_context,
    is_leading_context_start_for_context,
    is_numbered_intro_for_context,
    is_question_span_boundary_for_context,
    should_skip_question_start_for_context,
)
from .strategies import (
    BaseStrategy,
    ENGLISH_STRATEGY,
    LANGUAGE_STRATEGY,
    SCIENCE_STRATEGY,
    choose_strategy,
    extract_material_question_range,
    extract_numeric_question_id,
    extract_normalized_question_id,
)

EMBEDDED_DOCX_QUESTION_SPLIT_PATTERN = re.compile(
    r"(?<=[)）。；;！？!?])(?=\d+[．.](?:\(|（))"
)
INLINE_SUBQUESTION_SPLIT_PATTERN = re.compile(r"([（(]\s*\d{1,2}\s*[）)])")
INLINE_FIRST_SUBQUESTION_PATTERN = re.compile(r"[（(]\s*1\s*[）)]")


def _make_warning_slug(message: str) -> str:
    return message.replace(" ", "_")


def _split_inline_subquestion_segments(text: str) -> list[str]:
    cleaned = (text or "").strip()
    matches = list(INLINE_SUBQUESTION_SPLIT_PATTERN.finditer(cleaned))
    first_subquestion_index = next(
        (
            index
            for index, match in enumerate(matches)
            if re.search(r"\d+", match.group(1)).group(0) == "1"
        ),
        None,
    )
    if first_subquestion_index is None:
        return [cleaned] if cleaned else []

    matches = matches[first_subquestion_index:]

    segments: list[str] = []
    prefix = cleaned[: matches[0].start()].strip()
    if prefix:
        segments.append(prefix)
    for idx, match in enumerate(matches):
        start = match.start()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(cleaned)
        segment = cleaned[start:end].strip()
        if segment:
            segments.append(segment)
    return segments


def _choose_strategy_for_subject(
    doc_name: str,
    subject_name: str | None,
    sample_text: str,
    overlay_name: str | None = None,
) -> BaseStrategy:
    overlay_strategy = choose_strategy_for_context(subject_name, overlay_name)
    if overlay_strategy:
        return overlay_strategy
    if subject_name == "文科":
        return LANGUAGE_STRATEGY
    if subject_name == "理科":
        return SCIENCE_STRATEGY
    if subject_name == "英语":
        return ENGLISH_STRATEGY
    return choose_strategy(doc_name, sample_text)


def _resolve_unit_question_id(
    q_node: dict,
    nodes_by_index: dict[int, DocNode],
    start_idx: int,
    end_idx: int,
    strategy: BaseStrategy,
) -> str | None:
    if q_node.get("type") in {"READING", "LEADING_CONTEXT"}:
        current_text = q_node.get("text", "")
        current_normalized = (
            extract_normalized_question_id(current_text)
            if strategy.is_question_start(current_text)
            else None
        )
        if current_normalized:
            return current_normalized
        for idx in range(start_idx + 1, end_idx + 1):
            node = nodes_by_index.get(idx)
            if not node:
                continue
            text = (node.text or "").strip()
            if not text or strategy.is_option_line(text):
                continue
            if strategy.is_question_start(text):
                normalized = extract_normalized_question_id(text)
                if normalized:
                    return normalized
        return None

    normalized = extract_normalized_question_id(q_node.get("text", ""))
    if normalized:
        return normalized
    return strategy.extract_question_id(q_node.get("text", ""))


def scan_docx_nodes(docx_path: str | Path) -> list[DocNode]:
    def paragraph_media_sha256(para) -> list[str]:
        relationship_ids: list[str] = []
        for blip in para._p.iter(qn("a:blip")):
            relationship_id = blip.get(qn("r:embed"))
            if relationship_id and relationship_id not in relationship_ids:
                relationship_ids.append(relationship_id)
        for image_data in para._p.iter("{urn:schemas-microsoft-com:vml}imagedata"):
            relationship_id = image_data.get(qn("r:id"))
            if relationship_id and relationship_id not in relationship_ids:
                relationship_ids.append(relationship_id)

        hashes: list[str] = []
        for relationship_id in relationship_ids:
            relationship = doc.part.rels.get(relationship_id)
            target_part = getattr(relationship, "target_part", None)
            blob = getattr(target_part, "blob", None)
            if not blob:
                continue
            digest = hashlib.sha256(blob).hexdigest()
            if digest not in hashes:
                hashes.append(digest)
        return hashes

    def split_embedded_question_segments(text: str) -> list[str]:
        cleaned = (text or "").strip()
        if not cleaned:
            return [""]

        split_points = [0]
        split_points.extend(match.start() for match in EMBEDDED_DOCX_QUESTION_SPLIT_PATTERN.finditer(cleaned))
        split_points.append(len(cleaned))

        segments: list[str] = []
        for start, end in zip(split_points, split_points[1:]):
            segment = cleaned[start:end].strip()
            if segment:
                segments.append(segment)
        return segments or [""]

    doc = Document(docx_path)
    try:
        numbering_root = doc.part.numbering_part.element
    except (KeyError, NotImplementedError):
        numbering_root = None
    abstract_formats: dict[str, dict[str, tuple[str, str]]] = {}
    numbering_formats: dict[tuple[str, str], tuple[str, str]] = {}
    if numbering_root is not None:
        for abstract_num in numbering_root.findall(qn("w:abstractNum")):
            abstract_id = abstract_num.get(qn("w:abstractNumId"))
            level_formats = {}
            for level in abstract_num.findall(qn("w:lvl")):
                level_id = level.get(qn("w:ilvl"), "0")
                num_fmt = level.find(qn("w:numFmt"))
                if num_fmt is not None:
                    level_text = level.find(qn("w:lvlText"))
                    level_formats[level_id] = (
                        num_fmt.get(qn("w:val"), ""),
                        level_text.get(qn("w:val"), "") if level_text is not None else "",
                    )
            abstract_formats[abstract_id] = level_formats

        for numbering in numbering_root.findall(qn("w:num")):
            num_id = numbering.get(qn("w:numId"))
            abstract_ref = numbering.find(qn("w:abstractNumId"))
            if abstract_ref is None:
                continue
            abstract_id = abstract_ref.get(qn("w:val"))
            for level_id, numbering_definition in abstract_formats.get(abstract_id, {}).items():
                numbering_formats[(num_id, level_id)] = numbering_definition

    nodes: list[DocNode] = []
    virtual_index = 1
    last_explicit_question_number = 0
    subquestion_number = 0
    for idx, para in enumerate(doc.paragraphs, 1):
        xml = para._p.xml
        media_sha256 = paragraph_media_sha256(para)
        paragraph_text = para.text or ""
        explicit_match = re.match(r"^\s*(\d+)\s*[．.、]", paragraph_text)
        if explicit_match:
            last_explicit_question_number = int(explicit_match.group(1))
            subquestion_number = 0
        num_pr = para._p.pPr.numPr if para._p.pPr is not None else None
        num_id = str(num_pr.numId.val) if num_pr is not None and num_pr.numId is not None else ""
        level_id = str(num_pr.ilvl.val) if num_pr is not None and num_pr.ilvl is not None else "0"
        num_fmt, level_text = numbering_formats.get((num_id, level_id), ("", ""))
        is_top_level_decimal_numbering = (
            num_fmt == "decimal"
            and bool(re.fullmatch(r"%1\s*[．.、]", level_text))
        )
        if (
            not explicit_match
            and last_explicit_question_number
            and is_top_level_decimal_numbering
            and paragraph_text.strip()
        ):
            last_explicit_question_number += 1
            subquestion_number = 0
            paragraph_text = f"{last_explicit_question_number}．{paragraph_text.lstrip()}"
        elif (
            not explicit_match
            and num_fmt == "decimal"
            and bool(re.fullmatch(r"[（(]%1[）)]", level_text))
            and paragraph_text.strip()
        ):
            subquestion_number += 1
            marker = level_text.replace("%1", str(subquestion_number))
            paragraph_text = f"{marker}{paragraph_text.lstrip()}"

        segments = split_embedded_question_segments(paragraph_text)
        for segment_index, segment in enumerate(segments):
            nodes.append(
                DocNode(
                    index=virtual_index,
                    text=segment,
                    has_inline_media=segment_index == 0 and ("<wp:inline" in xml or "<w:drawing" in xml),
                    has_anchor_media=segment_index == 0 and "<wp:anchor" in xml,
                    page_break_before=segment_index == 0 and ("w:lastRenderedPageBreak" in xml or 'w:type="page"' in xml),
                    metadata={
                        "source_paragraph_index": idx,
                        "media_sha256": media_sha256 if segment_index == 0 else [],
                    },
                )
            )
            virtual_index += 1
    return nodes


def _is_node_span_boundary_for_context(
    node: DocNode,
    overlay_name: str | None,
) -> bool:
    if is_question_span_boundary_for_context(node.text, overlay_name):
        return True
    return bool(
        classify_media_hashes_for_context(
            node.metadata.get("media_sha256", []),
            overlay_name,
        )
    )


def scan_wps_nodes(doc, start_p: int, end_p: int) -> list[DocNode]:
    nodes: list[DocNode] = []
    paras = doc.Paragraphs
    for idx in range(start_p, end_p + 1):
        try:
            rng = paras(idx).Range
            text = (rng.Text or "").strip()
            list_label = ""
            try:
                list_label = str(rng.ListFormat.ListString or "").strip()
            except Exception:
                list_label = ""
            if (
                list_label
                and (
                    re.match(r"^\d+\s*[．.、]$", list_label)
                    or re.match(r"^[（(]\d+[）)]$", list_label)
                )
                and not text.startswith(list_label)
            ):
                text = f"{list_label}{text}"
            has_inline = bool(getattr(rng, "InlineShapes", None) and rng.InlineShapes.Count > 0)
            has_anchor = False
            try:
                has_anchor = bool(getattr(rng, "ShapeRange", None) and rng.ShapeRange.Count > 0)
            except Exception:
                has_anchor = False
            nodes.append(
                DocNode(
                    index=idx,
                    text=text,
                    has_inline_media=has_inline,
                    has_anchor_media=has_anchor,
                    metadata={
                        "in_table": bool(rng.Information(12)),
                        "list_label": list_label,
                    },
                )
            )
        except Exception:
            continue
    return nodes


def _collect_unit_parts(
    nodes_by_index: dict[int, DocNode],
    start_idx: int,
    end_idx: int,
    strategy: BaseStrategy,
    node_type: str,
) -> tuple[list[str], list[str], list[str], list[int], list[str], list[str], float]:
    stem_blocks: list[str] = []
    option_blocks: list[str] = []
    subquestions: list[str] = []
    media_blocks: list[int] = []
    material_blocks: list[str] = []
    warnings: list[str] = []
    confidence = 1.0

    all_text = []
    for idx in range(start_idx, end_idx + 1):
        node = nodes_by_index.get(idx)
        if not node:
            continue
        text = node.text.strip()
        if node.has_media:
            media_blocks.append(idx)
        if node.page_break_before and media_blocks:
            warnings.append("cross_page_media")
            confidence -= 0.15
        if not text:
            continue

        all_text.append(text)
        if node_type == "READING" and idx == start_idx:
            material_blocks.append(text)
            continue
        if strategy.is_option_line(text):
            option_blocks.append(text)
        elif strategy.is_subquestion_line(text) or INLINE_FIRST_SUBQUESTION_PATTERN.search(text):
            for subquestion_text in _split_inline_subquestion_segments(text):
                if strategy.is_subquestion_line(subquestion_text):
                    subquestions.append(subquestion_text)
                stem_blocks.append(subquestion_text)
        elif strategy.is_material_line(text) and idx != start_idx:
            material_blocks.append(text)
        else:
            stem_blocks.append(text)

    joined = " ".join(all_text)
    if strategy.has_figure_reference(joined):
        if not media_blocks:
            warnings.append("figure_reference_without_media")
            confidence -= 0.3
        else:
            warnings.append("image_related_question")
            confidence -= 0.05

    if option_blocks and len(option_blocks) < 2:
        warnings.append("sparse_options")
        confidence -= 0.25

    if node_type == "READING" and not any(extract_normalized_question_id(text) for text in stem_blocks + option_blocks + subquestions):
        warnings.append("material_without_question")
        confidence -= 0.35

    if media_blocks and option_blocks:
        first_option_idx = None
        for idx in range(start_idx, end_idx + 1):
            node = nodes_by_index.get(idx)
            if node and strategy.is_option_line(node.text):
                first_option_idx = idx
                break
        if first_option_idx and any(m < first_option_idx for m in media_blocks):
            warnings.append("image_between_stem_and_options")
            confidence -= 0.1

    unique_warnings = []
    for warning in warnings:
        if warning not in unique_warnings:
            unique_warnings.append(warning)

    return (
        stem_blocks,
        option_blocks,
        subquestions,
        media_blocks,
        material_blocks,
        unique_warnings,
        max(0.1, min(1.0, confidence)),
    )


def _detect_raw_starts(
    nodes: list[DocNode],
    strategy: BaseStrategy,
    overlay_name: str | None = None,
) -> list[dict]:
    raw_starts: list[dict] = []
    for node in nodes:
        if node.metadata.get("in_table"):
            continue
        text = (node.text or "").strip()
        if not text:
            continue
        if is_leading_context_start_for_context(text, overlay_name):
            raw_starts.append({"idx": node.index, "type": "LEADING_CONTEXT", "text": text})
            continue
        if is_leading_context_group_boundary_for_context(text, overlay_name):
            raw_starts.append({"idx": node.index, "type": "GROUP_BOUNDARY", "text": text})
            continue
        if should_skip_question_start_for_context(text, overlay_name):
            continue
        if strategy.is_material_line(text):
            raw_starts.append({"idx": node.index, "type": "READING", "text": text})
        elif strategy.is_question_start(text) and not strategy.is_option_line(text):
            node_type = "SUBQUESTION" if strategy.is_subquestion_line(text) else "STD"
            raw_starts.append({"idx": node.index, "type": node_type, "text": text})
    return raw_starts


def _merge_raw_starts(
    raw_starts: list[dict],
    strategy: BaseStrategy,
    group_leading_context_questions: bool = False,
) -> list[dict]:
    if not raw_starts:
        return []

    q_nodes: list[dict] = []
    pointer = 0
    while pointer < len(raw_starts):
        current = raw_starts[pointer]
        if current["type"] == "GROUP_BOUNDARY":
            pointer += 1
            continue
        if current["type"] == "LEADING_CONTEXT":
            pointer += 1
            while pointer < len(raw_starts) and raw_starts[pointer]["type"] == "SUBQUESTION":
                pointer += 1
            if pointer >= len(raw_starts) or raw_starts[pointer]["type"] == "LEADING_CONTEXT":
                continue
            q_nodes.append(
                {
                    **current,
                    "question_idx": raw_starts[pointer]["idx"],
                    "group_all_questions": group_leading_context_questions,
                }
            )
            pointer += 1
            if group_leading_context_questions:
                while pointer < len(raw_starts) and raw_starts[pointer]["type"] not in {
                    "LEADING_CONTEXT",
                    "GROUP_BOUNDARY",
                }:
                    pointer += 1
                continue
            while pointer < len(raw_starts) and raw_starts[pointer]["type"] == "SUBQUESTION":
                pointer += 1
            continue
        if current["type"] != "SUBQUESTION":
            normalized_type = "STD" if current["type"] == "SUBQUESTION" else current["type"]
            q_nodes.append({**current, "type": normalized_type})
        if current["type"] == "READING":
            current_text = current.get("text", "")
            current_has_question_id = bool(
                strategy.is_question_start(current_text) and extract_normalized_question_id(current_text)
            )
            current_range = extract_material_question_range(current_text)
            pointer += 1
            if current_has_question_id:
                while pointer < len(raw_starts) and raw_starts[pointer]["type"] == "SUBQUESTION":
                    pointer += 1
            elif current_range:
                range_start, range_end = current_range
                while pointer < len(raw_starts):
                    next_node = raw_starts[pointer]
                    if next_node["type"] == "READING":
                        break
                    if next_node["type"] == "SUBQUESTION":
                        pointer += 1
                        continue
                    next_qid = extract_numeric_question_id(next_node.get("text", ""))
                    if next_qid is None:
                        pointer += 1
                        continue
                    if range_start <= next_qid <= range_end:
                        pointer += 1
                        continue
                    break
            else:
                while pointer < len(raw_starts) and raw_starts[pointer]["type"] != "READING":
                    pointer += 1
            continue
        if current["type"] == "STD":
            pointer += 1
            while pointer < len(raw_starts) and raw_starts[pointer]["type"] == "SUBQUESTION":
                pointer += 1
            continue
        pointer += 1
    return q_nodes


def build_question_units_from_nodes(
    doc_name: str,
    subject_name: str,
    nodes: list[DocNode],
    grade_hint: str | None = None,
    overlay_name: str | None = None,
) -> list[QuestionUnit]:
    if not nodes:
        return []

    sample_text = " ".join(node.text for node in nodes[:20])
    if overlay_name is None:
        overlay_name = detect_subject_overlay(doc_name, sample_text, base_subject=subject_name)
    strategy = _choose_strategy_for_subject(doc_name, subject_name, sample_text, overlay_name=overlay_name)
    overlay = get_subject_overlay(overlay_name)
    q_nodes = _merge_raw_starts(
        _detect_raw_starts(nodes, strategy, overlay_name=overlay_name),
        strategy,
        group_leading_context_questions=bool(
            overlay and overlay.group_leading_context_questions
        ),
    )
    if not q_nodes:
        return []
    return build_question_units_from_wps_spans(
        doc_name=doc_name,
        subject_name=subject_name,
        nodes=nodes,
        q_nodes=q_nodes,
        end_p=nodes[-1].index,
        grade_hint=grade_hint,
        overlay_name=overlay_name,
    )


def build_question_units_from_wps_spans(
    doc_name: str,
    subject_name: str,
    nodes: list[DocNode],
    q_nodes: list[dict],
    end_p: int,
    grade_hint: str | None = None,
    overlay_name: str | None = None,
) -> list[QuestionUnit]:
    sample_text = " ".join(node.text for node in nodes[:20])
    if overlay_name is None:
        overlay_name = detect_subject_overlay(doc_name, sample_text, base_subject=subject_name)
    strategy = _choose_strategy_for_subject(doc_name, subject_name, sample_text, overlay_name=overlay_name)
    nodes_by_index = {node.index: node for node in nodes}
    units: list[QuestionUnit] = []

    for idx, q_node in enumerate(q_nodes):
        start_idx = q_node["idx"]
        end_idx = q_nodes[idx + 1]["idx"] - 1 if idx + 1 < len(q_nodes) else end_p
        boundary_start_idx = start_idx + 1
        if q_node.get("type") == "LEADING_CONTEXT":
            boundary_start_idx = q_node.get("question_idx", start_idx) + 1
        if q_node.get("group_all_questions"):
            while end_idx >= boundary_start_idx:
                node = nodes_by_index.get(end_idx)
                if not node or not _is_node_span_boundary_for_context(node, overlay_name):
                    break
                end_idx -= 1
        else:
            for boundary_idx in range(boundary_start_idx, end_idx + 1):
                node = nodes_by_index.get(boundary_idx)
                if node and _is_node_span_boundary_for_context(node, overlay_name):
                    end_idx = boundary_idx - 1
                    break

        qid = _resolve_unit_question_id(q_node, nodes_by_index, start_idx, end_idx, strategy)
        if not qid:
            continue

        content_start_idx = start_idx
        if q_node.get("type") == "LEADING_CONTEXT":
            content_start_idx = start_idx + 1
        if is_numbered_intro_for_context(q_node.get("text", ""), overlay_name):
            content_start_idx = 0
            for candidate_idx in range(start_idx + 1, end_idx + 1):
                node = nodes_by_index.get(candidate_idx)
                if not node:
                    continue
                if _is_node_span_boundary_for_context(node, overlay_name):
                    break
                if (node.text or "").strip() or node.has_media:
                    content_start_idx = candidate_idx
                    break
            if not content_start_idx:
                continue

        (
            stem_blocks,
            option_blocks,
            subquestions,
            media_blocks,
            material_blocks,
            warnings,
            confidence,
        ) = _collect_unit_parts(nodes_by_index, content_start_idx, end_idx, strategy, q_node.get("type", "STD"))
        question_type = "choice" if option_blocks else "subjective"
        if q_node.get("type") in {"READING", "LEADING_CONTEXT"}:
            question_type = "material_choice" if option_blocks else "material"
        elif "_" in " ".join(stem_blocks):
            question_type = "fill_blank"

        units.append(
            QuestionUnit(
                question_id=qid,
                subject=subject_name,
                subject_overlay=overlay_name,
                grade_hint=grade_hint,
                question_type=question_type,
                stem_blocks=stem_blocks,
                option_blocks=option_blocks,
                subquestions=subquestions,
                media_blocks=media_blocks,
                material_blocks=material_blocks,
                source_span=(content_start_idx, end_idx),
                confidence=confidence,
                warnings=warnings,
                node_type=q_node.get("type", "STD"),
            )
        )
    return units


def build_question_units_from_docx(docx_path: str | Path, grade_hint: str | None = None) -> list[QuestionUnit]:
    nodes = scan_docx_nodes(docx_path)
    if not nodes:
        return []
    doc_name = Path(docx_path).name
    sample_text = " ".join(node.text for node in nodes[:20])
    overlay_name = detect_subject_overlay(doc_name, sample_text, base_subject="文科")
    if overlay_name is None:
        overlay_name = detect_subject_overlay(doc_name, sample_text, base_subject="理科")
    if overlay_name:
        overlay = get_subject_overlay(overlay_name)
        subject_name = overlay.base_subject if overlay else choose_strategy(doc_name, sample_text).name
    else:
        subject_name = choose_strategy(doc_name, sample_text).name
    return build_question_units_from_nodes(
        doc_name,
        subject_name,
        nodes,
        grade_hint=grade_hint,
        overlay_name=overlay_name,
    )


def build_question_units_from_wps(
    doc_name: str,
    subject_name: str,
    doc,
    start_p: int,
    end_p: int,
    grade_hint: str | None = None,
    overlay_name: str | None = None,
) -> list[QuestionUnit]:
    nodes = scan_wps_nodes(doc, start_p, end_p)
    if not nodes:
        return []
    return build_question_units_from_nodes(
        doc_name,
        subject_name,
        nodes,
        grade_hint=grade_hint,
        overlay_name=overlay_name,
    )
