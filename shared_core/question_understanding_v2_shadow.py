from __future__ import annotations

import hashlib
from collections import Counter
from pathlib import Path
from typing import Any, Iterable, Sequence

from .contract_v2_shadow import build_legacy_node_source_refs
from .contracts_v2 import DecisionEvidence, SourceRef, canonical_json_bytes
from .models import DocNode, QuestionUnit
from .question_core import build_question_units_from_docx, scan_docx_nodes
from .question_understanding_v2 import (
    QUESTION_UNDERSTANDING_GENERATOR_VERSION,
    QUESTION_UNDERSTANDING_SCHEMA_VERSION,
    ProposedQuestionUnit,
    QuestionUnderstandingProposal,
    RefDecision,
    compile_question_understanding,
)
from .subject_overlay import (
    classify_media_hashes_for_context,
    is_question_input_excluded_for_context,
)


RULE_PROPOSAL_GENERATOR_VERSION = "mohen-question-rules-v2/1"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _unique_strings(values: Iterable[str]) -> tuple[str, ...]:
    result: list[str] = []
    for value in values:
        if value and value not in result:
            result.append(value)
    return tuple(result)


def _evidence(
    ref: SourceRef,
    *,
    rule_id: str,
    reason: str,
    confidence: float,
) -> DecisionEvidence:
    return DecisionEvidence(
        rule_id=rule_id,
        reason=reason,
        source_refs=(ref,),
        confidence=confidence,
        generator_version=RULE_PROPOSAL_GENERATOR_VERSION,
    )


def _paragraph_role(node: DocNode, unit: QuestionUnit) -> str:
    text = (node.text or "").strip()
    if text in {(item or "").strip() for item in unit.material_blocks}:
        return "material"
    if text in {(item or "").strip() for item in unit.option_blocks}:
        return "option"
    if text in {(item or "").strip() for item in unit.subquestions}:
        return "subquestion"
    return "stem"


def _media_hashes_for_ref(ref: SourceRef, node: DocNode) -> tuple[str, ...]:
    marker = "/media:"
    if marker in ref.native_id:
        return (ref.native_id.rsplit(marker, 1)[1],)
    return tuple(str(value) for value in node.metadata.get("media_sha256", []))


def _unit_ref_decision(
    ref: SourceRef,
    node: DocNode,
    unit: QuestionUnit,
) -> RefDecision:
    overlay = unit.subject_overlay
    confidence = unit.confidence
    if ref.native_kind == "paragraph":
        if is_question_input_excluded_for_context(node.text, overlay):
            disposition = "exclude"
            role = "internal_heading"
            reason = "项目覆盖层确认该题内标题不进入F1题块内容"
        else:
            disposition = "include"
            role = _paragraph_role(node, unit)
            reason = f"旧题块内容结构将该段识别为 {role}"
    elif ref.native_kind == "media":
        media_role = classify_media_hashes_for_context(
            _media_hashes_for_ref(ref, node),
            overlay,
        )
        if media_role:
            disposition = "exclude"
            role = "decorative_media"
            reason = f"项目覆盖层按已确认图片哈希识别为装饰媒体：{media_role}"
        else:
            disposition = "include"
            role = "question_media"
            reason = "旧题块范围内的未排除媒体作为题图候选"
    elif ref.native_kind == "table":
        disposition = "include"
        role = "table"
        reason = "旧题块范围内显式暴露的原生表格"
    else:
        disposition = "include"
        role = "formula"
        reason = "旧题块范围内显式暴露的公式"
    return RefDecision(
        source_ref=ref,
        disposition=disposition,
        role=role,
        evidence=_evidence(
            ref,
            rule_id=f"legacy-rule-candidate-{role}",
            reason=reason,
            confidence=confidence,
        ),
    )


def _document_ref_decision(
    ref: SourceRef,
    node: DocNode,
    overlay_names: Sequence[str | None],
) -> RefDecision:
    overlay = next((name for name in overlay_names if name), None)
    if ref.native_kind == "paragraph" and not (node.text or "").strip():
        disposition = "exclude"
        role = "top_level_boundary"
        reason = "题块范围外的空段落仅作为文档边界"
    elif ref.native_kind == "paragraph" and is_question_input_excluded_for_context(
        node.text,
        overlay,
    ):
        disposition = "exclude"
        role = "top_level_boundary"
        reason = "题块范围外且由项目覆盖层确认的结构标题"
    elif ref.native_kind == "media" and classify_media_hashes_for_context(
        _media_hashes_for_ref(ref, node),
        overlay,
    ):
        disposition = "exclude"
        role = "decorative_media"
        reason = "题块范围外且由项目覆盖层图片哈希确认的装饰媒体"
    else:
        disposition = "unknown"
        role = "unknown"
        reason = "旧题块没有覆盖该来源节点，规则层不得静默猜测"
    return RefDecision(
        source_ref=ref,
        disposition=disposition,
        role=role,
        evidence=_evidence(
            ref,
            rule_id=f"legacy-rule-document-{role}",
            reason=reason,
            confidence=1.0,
        ),
    )


def build_rule_question_understanding_proposal(
    document_sha256: str,
    legacy_units: Iterable[QuestionUnit],
    nodes: Iterable[DocNode],
    *,
    allowed_label_reset_before: tuple[int, ...] = (),
) -> tuple[QuestionUnderstandingProposal, tuple[SourceRef, ...]]:
    units = tuple(legacy_units)
    node_list = tuple(sorted(nodes, key=lambda node: node.index))
    if not units:
        raise ValueError("旧题块为空，规则候选层不能生成 QuestionUnitV2 建议")
    if not node_list:
        raise ValueError("DocNode 为空，规则候选层没有上游来源")
    nodes_by_index = {node.index: node for node in node_list}
    if len(nodes_by_index) != len(node_list):
        raise ValueError("DocNode.index 不能重复")
    refs_by_node = build_legacy_node_source_refs(document_sha256, node_list)
    source_refs = tuple(
        ref for node in node_list for ref in refs_by_node.get(node.index, ())
    )
    proposed_units: list[ProposedQuestionUnit] = []
    covered_ref_keys: set[str] = set()

    for sequence, unit in enumerate(units, 1):
        start, end = unit.source_span
        if start <= 0 or end < start:
            raise ValueError(f"旧 QuestionUnit.source_span 无效：{unit.source_span}")
        decisions: list[RefDecision] = []
        for node_index in range(start, end + 1):
            node = nodes_by_index.get(node_index)
            if node is None:
                continue
            for ref in refs_by_node.get(node_index, ()):
                decisions.append(_unit_ref_decision(ref, node, unit))
                covered_ref_keys.add(ref.key)
        if not decisions:
            raise ValueError(f"旧题块 {unit.question_id} 没有可引用来源")
        proposed_units.append(
            ProposedQuestionUnit(
                unit_key=f"legacy-unit-{sequence}",
                original_question_label=unit.question_id,
                member_question_labels=(unit.question_id,),
                subject=unit.subject,
                subject_overlay=unit.subject_overlay,
                question_type=unit.question_type,
                ref_decisions=tuple(decisions),
                review_flags=_unique_strings(unit.warnings),
            )
        )

    overlay_names = tuple(unit.subject_overlay for unit in units)
    document_decisions: list[RefDecision] = []
    for node in node_list:
        for ref in refs_by_node.get(node.index, ()):
            if ref.key not in covered_ref_keys:
                document_decisions.append(
                    _document_ref_decision(ref, node, overlay_names)
                )

    proposal = QuestionUnderstandingProposal(
        proposal_id="legacy-rules-question-understanding-v2",
        producer_kind="rules",
        generator_version=RULE_PROPOSAL_GENERATOR_VERSION,
        units=tuple(proposed_units),
        document_decisions=tuple(document_decisions),
        allowed_label_reset_before=allowed_label_reset_before,
    )
    return proposal, source_refs


def _build_diff(
    legacy_units: Sequence[QuestionUnit],
    proposal: QuestionUnderstandingProposal,
    result: dict[str, Any],
    *,
    source_sha256: str,
    source_name: str | None,
) -> dict[str, Any]:
    comparisons: list[dict[str, Any]] = []
    compiled_units = result["units"]
    for index, (legacy, proposed) in enumerate(
        zip(legacy_units, proposal.units),
        1,
    ):
        decisions = proposed.ref_decisions
        media_attribution = [
            {
                "source_ref_key": item.source_ref.key,
                "disposition": item.disposition,
                "role": item.role,
            }
            for item in decisions
            if item.source_ref.native_kind == "media"
        ]
        counts = Counter(item.disposition for item in decisions)
        comparisons.append(
            {
                "sequence": index,
                "legacy_question_label": legacy.question_id,
                "legacy_source_span": list(legacy.source_span),
                "proposed_unit_key": proposed.unit_key,
                "stable_unit_id": (
                    compiled_units[index - 1]["stable_unit_id"]["value"]
                    if index <= len(compiled_units)
                    else None
                ),
                "included_ref_count": counts.get("include", 0),
                "excluded_ref_count": counts.get("exclude", 0),
                "unknown_ref_count": counts.get("unknown", 0),
                "ref_decisions": [
                    {
                        "source_ref_key": item.source_ref.key,
                        "native_kind": item.source_ref.native_kind,
                        "disposition": item.disposition,
                        "role": item.role,
                    }
                    for item in decisions
                ],
                "media_attribution": media_attribution,
                "legacy_confidence": legacy.confidence,
                "legacy_warnings": list(legacy.warnings),
            }
        )
    return {
        "schema_version": QUESTION_UNDERSTANDING_SCHEMA_VERSION,
        "artifact_kind": "question_understanding_v2_diff",
        "mode": "shadow",
        "production_execution_enabled": False,
        "keypress_count": 0,
        "source_name": source_name,
        "source_sha256": source_sha256,
        "legacy_unit_count": len(legacy_units),
        "v2_unit_count": result["unit_count"],
        "status": result["status"],
        "comparisons": comparisons,
        "document_ref_decisions": [
            {
                "source_ref_key": item.source_ref.key,
                "native_kind": item.source_ref.native_kind,
                "disposition": item.disposition,
                "role": item.role,
            }
            for item in proposal.document_decisions
        ],
        "blocking_issues": result["blocking_issues"],
        "limitations": [
            "差异报告逐引用比较规则候选，不以题块数量相等代替语义一致。",
            "规则候选复用旧题块作为起点；未被旧题块覆盖的非空来源会进入unknown并阻断。",
            "仅比较上游已暴露的paragraph/table/media/formula，未暴露对象不能视为已覆盖。",
            "本报告不包含WPS Range，也不证明F1插件录入结果。",
        ],
    }


def build_question_understanding_v2_shadow(
    document_sha256: str,
    legacy_units: Iterable[QuestionUnit],
    nodes: Iterable[DocNode],
    *,
    source_name: str | None = None,
    allowed_label_reset_before: tuple[int, ...] = (),
) -> dict[str, dict[str, Any]]:
    units = tuple(legacy_units)
    node_list = tuple(nodes)
    proposal, source_refs = build_rule_question_understanding_proposal(
        document_sha256,
        units,
        node_list,
        allowed_label_reset_before=allowed_label_reset_before,
    )
    compiled = compile_question_understanding(
        source_refs,
        (proposal,),
        generator_version=QUESTION_UNDERSTANDING_GENERATOR_VERSION,
    )
    result = compiled.to_dict()
    candidates = {
        "schema_version": QUESTION_UNDERSTANDING_SCHEMA_VERSION,
        "artifact_kind": "question_understanding_v2_candidates",
        "mode": "shadow",
        "production_execution_enabled": False,
        "keypress_count": 0,
        "source_name": source_name,
        "source_sha256": document_sha256,
        "source_refs": [ref.to_dict() for ref in source_refs],
        "proposals": [proposal.to_dict()],
    }
    diff = _build_diff(
        units,
        proposal,
        result,
        source_sha256=document_sha256,
        source_name=source_name,
    )
    return {"candidates": candidates, "result": result, "diff": diff}


def write_question_understanding_v2_shadow(
    docx_path: str | Path,
    output_dir: str | Path,
    *,
    allowed_label_reset_before: tuple[int, ...] = (),
) -> dict[str, Path]:
    source = Path(docx_path).resolve()
    target_dir = Path(output_dir).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if source.suffix.lower() != ".docx":
        raise ValueError(f"只支持 DOCX 文件：{source}")
    if target_dir == source.parent or source.parent in target_dir.parents:
        raise ValueError("影子产物目录不能位于源文件所在目录内")

    hash_before = _sha256(source)
    nodes = scan_docx_nodes(source)
    legacy_units = build_question_units_from_docx(source)
    bundle = build_question_understanding_v2_shadow(
        hash_before,
        legacy_units,
        nodes,
        source_name=source.name,
        allowed_label_reset_before=allowed_label_reset_before,
    )
    hash_after = _sha256(source)
    if hash_after != hash_before:
        raise RuntimeError(f"源 DOCX 在题目理解影子运行期间发生变化：{source}")

    target_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "candidates_json": target_dir
        / f"{source.stem}_QuestionUnderstandingV2Candidates.json",
        "question_units_json": target_dir / f"{source.stem}_QuestionUnitV2.json",
        "diff_json": target_dir / f"{source.stem}_QuestionUnitV2Diff.json",
    }
    paths["candidates_json"].write_bytes(canonical_json_bytes(bundle["candidates"]))
    paths["question_units_json"].write_bytes(canonical_json_bytes(bundle["result"]))
    paths["diff_json"].write_bytes(canonical_json_bytes(bundle["diff"]))
    return paths
