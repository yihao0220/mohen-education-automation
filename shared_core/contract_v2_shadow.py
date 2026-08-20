from __future__ import annotations

import hashlib
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

from .contracts_v2 import (
    CONTRACT_GENERATOR_VERSION,
    CONTRACT_SCHEMA_VERSION,
    NATIVE_KINDS,
    DecisionEvidence,
    QuestionUnitV2,
    SourceRef,
    StableUnitId,
    build_source_partition,
    canonical_json_bytes,
)
from .models import DocNode, QuestionUnit
from .question_core import build_question_units_from_docx, scan_docx_nodes


SHADOW_ARTIFACT_VERSION = "1.0"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _text_anchor(text: str) -> str:
    normalized = " ".join((text or "").replace("\x07", "").split())
    if len(normalized) <= 160:
        return normalized
    digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:12]
    return f"{normalized[:147]}…#{digest}"


def _metadata_indexes(node: DocNode, field: str) -> tuple[str, ...]:
    values = node.metadata.get(field, [])
    if values is None:
        return ()
    if not isinstance(values, (list, tuple)):
        values = [values]
    return tuple(str(value) for value in values if str(value).strip())


def build_legacy_node_source_refs(
    document_sha256: str,
    nodes: Iterable[DocNode],
) -> dict[int, tuple[SourceRef, ...]]:
    ordered_nodes = sorted(nodes, key=lambda node: node.index)
    if len({node.index for node in ordered_nodes}) != len(ordered_nodes):
        raise ValueError("DocNode.index 不能重复")

    occurrences: Counter[tuple[str, str]] = Counter()
    refs_by_node: dict[int, tuple[SourceRef, ...]] = {}

    def create_ref(kind: str, native_id: str, anchor: str) -> SourceRef:
        occurrence_key = (kind, native_id)
        occurrences[occurrence_key] += 1
        return SourceRef(
            document_sha256=document_sha256,
            native_kind=kind,
            native_id=native_id,
            occurrence=occurrences[occurrence_key],
            text_anchor=anchor,
        )

    for node in ordered_nodes:
        paragraph_index = int(node.metadata.get("source_paragraph_index", node.index))
        anchor = _text_anchor(node.text)
        refs = [create_ref("paragraph", f"paragraph:{paragraph_index}", anchor)]

        media_hashes = _metadata_indexes(node, "media_sha256")
        for media_hash in media_hashes:
            refs.append(
                create_ref(
                    "media",
                    f"paragraph:{paragraph_index}/media:{media_hash}",
                    anchor,
                )
            )
        for table_index in _metadata_indexes(node, "table_indexes"):
            refs.append(create_ref("table", f"table:{table_index}", anchor))
        for formula_index in _metadata_indexes(node, "formula_indexes"):
            refs.append(
                create_ref(
                    "formula",
                    f"paragraph:{paragraph_index}/formula:{formula_index}",
                    anchor,
                )
            )
        refs_by_node[node.index] = tuple(refs)
    return refs_by_node


def _refs_for_span(
    refs_by_node: dict[int, tuple[SourceRef, ...]],
    source_span: tuple[int, int],
) -> tuple[SourceRef, ...]:
    start, end = source_span
    if start <= 0 or end < start:
        raise ValueError(f"QuestionUnit.source_span 无效：{source_span}")
    refs = tuple(
        ref
        for node_index in range(start, end + 1)
        for ref in refs_by_node.get(node_index, ())
    )
    if not refs:
        raise ValueError(f"QuestionUnit.source_span 没有可引用节点：{source_span}")
    return refs


def build_question_contract_v2_shadow(
    document_sha256: str,
    legacy_units: Iterable[QuestionUnit],
    nodes: Iterable[DocNode],
    *,
    source_name: str | None = None,
    generator_version: str = CONTRACT_GENERATOR_VERSION,
) -> dict[str, dict[str, Any]]:
    units = list(legacy_units)
    node_list = list(nodes)
    refs_by_node = build_legacy_node_source_refs(document_sha256, node_list)
    v2_units: list[QuestionUnitV2] = []
    comparisons: list[dict[str, Any]] = []

    for sequence, legacy_unit in enumerate(units, 1):
        refs = _refs_for_span(refs_by_node, legacy_unit.source_span)
        partition = build_source_partition(
            refs,
            included_refs=refs,
            excluded_refs=(),
            unknown_refs=(),
        )
        evidence = DecisionEvidence(
            rule_id="legacy-question-unit-shadow-adapter",
            reason=(
                "由旧 QuestionUnit 连续 source_span 生成，仅用于 V2 身份与来源覆盖对照；"
                "不代表已完成非连续排除或新题块判断。"
            ),
            source_refs=refs,
            confidence=legacy_unit.confidence,
            generator_version=generator_version,
        )
        stable_id = StableUnitId.create(
            document_sha256=document_sha256,
            unit_kind="question",
            occurrence=sequence,
            original_label=legacy_unit.question_id,
        )
        v2_unit = QuestionUnitV2(
            generator_version=generator_version,
            stable_unit_id=stable_id,
            original_question_label=legacy_unit.question_id,
            subject=legacy_unit.subject,
            subject_overlay=legacy_unit.subject_overlay,
            question_type=legacy_unit.question_type,
            source_partition=partition,
            decision_evidence=(evidence,),
        )
        v2_units.append(v2_unit)
        kind_counts = Counter(ref.native_kind for ref in refs)
        comparisons.append(
            {
                "sequence": sequence,
                "original_question_label": legacy_unit.question_id,
                "stable_unit_id": stable_id.value,
                "legacy_source_span": list(legacy_unit.source_span),
                "source_ref_count": len(refs),
                "native_kind_counts": {
                    kind: kind_counts.get(kind, 0) for kind in NATIVE_KINDS
                },
                "excluded_ref_count": 0,
                "unknown_ref_count": 0,
                "legacy_span_fully_represented": True,
            }
        )

    label_counts = Counter(unit.question_id for unit in units)
    duplicate_labels = []
    for label in sorted(label for label, count in label_counts.items() if count > 1):
        matching_ids = [
            unit.stable_unit_id.value
            for unit in v2_units
            if unit.original_question_label == label
        ]
        duplicate_labels.append(
            {"label": label, "count": label_counts[label], "stable_ids": matching_ids}
        )

    stable_values = [unit.stable_unit_id.value for unit in v2_units]
    contract = {
        "schema_version": CONTRACT_SCHEMA_VERSION,
        "artifact_kind": "question_contract_v2_bundle",
        "generator_version": generator_version,
        "mode": "shadow",
        "production_execution_enabled": False,
        "keypress_count": 0,
        "source_name": source_name,
        "source_sha256": document_sha256,
        "unit_count": len(v2_units),
        "units": [unit.to_dict() for unit in v2_units],
        "blocking_issues": [
            {
                "code": "SHADOW_CONTRACT_NOT_PRODUCTION",
                "severity": "blocker",
                "message": "V2 交接单当前只用于影子对照，不允许进入生产动作计划。",
            }
        ],
    }
    diff = {
        "schema_version": SHADOW_ARTIFACT_VERSION,
        "artifact_kind": "question_contract_v2_diff",
        "generator_version": generator_version,
        "mode": "shadow",
        "production_execution_enabled": False,
        "keypress_count": 0,
        "source_name": source_name,
        "source_sha256": document_sha256,
        "legacy_unit_count": len(units),
        "v2_unit_count": len(v2_units),
        "stable_id_unique": len(stable_values) == len(set(stable_values)),
        "duplicate_original_labels": duplicate_labels,
        "comparisons": comparisons,
        "limitations": [
            "本报告只比较身份和旧连续 source_span 的来源覆盖，不判断题块语义正确性。",
            "旧 QuestionUnit 无显式 excluded_refs；影子适配不会猜测中间标题排除。",
            "table/media/formula 只接收上游 DocNode metadata 已显式暴露的引用；未暴露对象不能视为已覆盖。",
            "本报告不包含 WPS Range 绑定，也不证明 F1 插件录入结果。",
        ],
    }
    return {"contract": contract, "diff": diff}


def write_question_contract_v2_shadow(
    docx_path: str | Path,
    output_dir: str | Path,
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
    units = build_question_units_from_docx(source)
    bundle = build_question_contract_v2_shadow(
        hash_before,
        units,
        nodes,
        source_name=source.name,
    )
    hash_after = _sha256(source)
    if hash_after != hash_before:
        raise RuntimeError(f"源 DOCX 在影子转换期间发生变化：{source}")

    target_dir.mkdir(parents=True, exist_ok=True)
    contract_path = target_dir / f"{source.stem}_QuestionContractV2.json"
    diff_path = target_dir / f"{source.stem}_QuestionContractV2Diff.json"
    contract_path.write_bytes(canonical_json_bytes(bundle["contract"]))
    diff_path.write_bytes(canonical_json_bytes(bundle["diff"]))
    return {"contract_json": contract_path, "diff_json": diff_path}
