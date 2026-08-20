from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from typing import Any, Mapping


CONTRACT_SCHEMA_VERSION = "2.0"
CONTRACT_GENERATOR_VERSION = "mohen-contract-v2/1"
NATIVE_KINDS = ("paragraph", "table", "media", "formula")

_SHA256_PATTERN = re.compile(r"^[0-9a-f]{64}$")
_UNIT_KIND_PATTERN = re.compile(r"^[a-z][a-z0-9_.-]*$")


def canonical_json_bytes(value: Any) -> bytes:
    return (
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        + "\n"
    ).encode("utf-8")


def _require_exact_fields(
    value: Mapping[str, Any],
    expected: set[str],
    label: str,
) -> None:
    actual = set(value)
    if actual != expected:
        missing = sorted(expected - actual)
        extra = sorted(actual - expected)
        raise ValueError(f"{label} 字段不一致：缺少 {missing}；多出 {extra}")


def _require_sha256(value: str, field: str = "document_sha256") -> None:
    if not isinstance(value, str) or not _SHA256_PATTERN.fullmatch(value):
        raise ValueError(f"{field} 必须是 64 位小写十六进制字符串")


def _require_nonempty_string(value: str, field: str) -> None:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} 必须是非空字符串")


@dataclass(frozen=True)
class SourceRef:
    document_sha256: str
    native_kind: str
    native_id: str
    occurrence: int
    text_anchor: str

    def __post_init__(self) -> None:
        _require_sha256(self.document_sha256)
        if self.native_kind not in NATIVE_KINDS:
            raise ValueError(f"native_kind 必须是 {NATIVE_KINDS} 之一")
        _require_nonempty_string(self.native_id, "native_id")
        if isinstance(self.occurrence, bool) or not isinstance(self.occurrence, int) or self.occurrence <= 0:
            raise ValueError("occurrence 必须是正整数")
        if not isinstance(self.text_anchor, str):
            raise ValueError("text_anchor 必须是字符串")

    @property
    def key(self) -> str:
        identity = {
            "native_id": self.native_id,
            "occurrence": self.occurrence,
            "text_anchor": self.text_anchor,
        }
        digest = hashlib.sha256(canonical_json_bytes(identity)).hexdigest()
        return f"{self.document_sha256}:{self.native_kind}:{digest}"

    def to_dict(self) -> dict[str, Any]:
        return {
            "document_sha256": self.document_sha256,
            "native_kind": self.native_kind,
            "native_id": self.native_id,
            "occurrence": self.occurrence,
            "text_anchor": self.text_anchor,
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "SourceRef":
        if not isinstance(value, Mapping):
            raise ValueError("SourceRef 必须是 JSON 对象")
        _require_exact_fields(
            value,
            {
                "document_sha256",
                "native_kind",
                "native_id",
                "occurrence",
                "text_anchor",
            },
            "SourceRef",
        )
        return cls(
            document_sha256=value["document_sha256"],
            native_kind=value["native_kind"],
            native_id=value["native_id"],
            occurrence=value["occurrence"],
            text_anchor=value["text_anchor"],
        )


@dataclass(frozen=True)
class StableUnitId:
    document_sha256: str
    unit_kind: str
    occurrence: int
    original_label: str
    value: str

    def __post_init__(self) -> None:
        _require_sha256(self.document_sha256)
        if not isinstance(self.unit_kind, str) or not _UNIT_KIND_PATTERN.fullmatch(self.unit_kind):
            raise ValueError("unit_kind 必须是小写英文标识符")
        if isinstance(self.occurrence, bool) or not isinstance(self.occurrence, int) or self.occurrence <= 0:
            raise ValueError("occurrence 必须是正整数")
        _require_nonempty_string(self.original_label, "original_label")
        expected = self._derive_value(
            self.document_sha256,
            self.unit_kind,
            self.occurrence,
            self.original_label,
        )
        if self.value != expected:
            raise ValueError("StableUnitId.value 与身份字段不一致")

    @staticmethod
    def _derive_value(
        document_sha256: str,
        unit_kind: str,
        occurrence: int,
        original_label: str,
    ) -> str:
        identity = {
            "document_sha256": document_sha256,
            "identity_version": CONTRACT_SCHEMA_VERSION,
            "occurrence": occurrence,
            "original_label": original_label,
            "unit_kind": unit_kind,
        }
        digest = hashlib.sha256(canonical_json_bytes(identity)).hexdigest()
        return f"mohen:v2:{digest}"

    @classmethod
    def create(
        cls,
        *,
        document_sha256: str,
        unit_kind: str,
        occurrence: int,
        original_label: str,
    ) -> "StableUnitId":
        value = cls._derive_value(
            document_sha256,
            unit_kind,
            occurrence,
            original_label,
        )
        return cls(document_sha256, unit_kind, occurrence, original_label, value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "document_sha256": self.document_sha256,
            "unit_kind": self.unit_kind,
            "occurrence": self.occurrence,
            "original_label": self.original_label,
            "value": self.value,
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "StableUnitId":
        if not isinstance(value, Mapping):
            raise ValueError("StableUnitId 必须是 JSON 对象")
        _require_exact_fields(
            value,
            {
                "document_sha256",
                "unit_kind",
                "occurrence",
                "original_label",
                "value",
            },
            "StableUnitId",
        )
        return cls(
            document_sha256=value["document_sha256"],
            unit_kind=value["unit_kind"],
            occurrence=value["occurrence"],
            original_label=value["original_label"],
            value=value["value"],
        )


@dataclass(frozen=True)
class DecisionEvidence:
    rule_id: str
    reason: str
    source_refs: tuple[SourceRef, ...]
    confidence: float
    generator_version: str

    def __post_init__(self) -> None:
        _require_nonempty_string(self.rule_id, "rule_id")
        _require_nonempty_string(self.reason, "reason")
        if not isinstance(self.source_refs, tuple) or not self.source_refs:
            raise ValueError("source_refs 必须是非空 SourceRef 元组")
        if any(not isinstance(ref, SourceRef) for ref in self.source_refs):
            raise ValueError("source_refs 只能包含 SourceRef")
        if len({ref.key for ref in self.source_refs}) != len(self.source_refs):
            raise ValueError("source_refs 不能包含重复引用")
        if (
            isinstance(self.confidence, bool)
            or not isinstance(self.confidence, (int, float))
            or not 0 <= self.confidence <= 1
        ):
            raise ValueError("confidence 必须位于 0 到 1")
        _require_nonempty_string(self.generator_version, "generator_version")
        object.__setattr__(self, "confidence", float(self.confidence))

    def to_dict(self) -> dict[str, Any]:
        return {
            "rule_id": self.rule_id,
            "reason": self.reason,
            "source_refs": [ref.to_dict() for ref in self.source_refs],
            "confidence": self.confidence,
            "generator_version": self.generator_version,
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "DecisionEvidence":
        if not isinstance(value, Mapping):
            raise ValueError("DecisionEvidence 必须是 JSON 对象")
        _require_exact_fields(
            value,
            {"rule_id", "reason", "source_refs", "confidence", "generator_version"},
            "DecisionEvidence",
        )
        refs = value["source_refs"]
        if not isinstance(refs, list):
            raise ValueError("DecisionEvidence.source_refs 必须是数组")
        return cls(
            rule_id=value["rule_id"],
            reason=value["reason"],
            source_refs=tuple(SourceRef.from_dict(ref) for ref in refs),
            confidence=value["confidence"],
            generator_version=value["generator_version"],
        )


def _require_ref_tuple(value: tuple[SourceRef, ...], field: str) -> None:
    if not isinstance(value, tuple):
        raise ValueError(f"{field} 必须是 SourceRef 元组")
    if any(not isinstance(ref, SourceRef) for ref in value):
        raise ValueError(f"{field} 只能包含 SourceRef")
    if len({ref.key for ref in value}) != len(value):
        raise ValueError(f"{field} 不能包含重复引用")


@dataclass(frozen=True)
class SourcePartition:
    included_refs: tuple[SourceRef, ...]
    excluded_refs: tuple[SourceRef, ...]
    unknown_refs: tuple[SourceRef, ...]

    def __post_init__(self) -> None:
        for field in ("included_refs", "excluded_refs", "unknown_refs"):
            _require_ref_tuple(getattr(self, field), field)
        all_refs = self.all_refs
        if not all_refs:
            raise ValueError("来源分区不能为空")
        if len({ref.key for ref in all_refs}) != len(all_refs):
            raise ValueError("来源引用被重复分类")
        if len({ref.document_sha256 for ref in all_refs}) != 1:
            raise ValueError("一个来源分区只能引用同一份文档")

    @property
    def all_refs(self) -> tuple[SourceRef, ...]:
        return self.included_refs + self.excluded_refs + self.unknown_refs

    @property
    def has_unknown_refs(self) -> bool:
        return bool(self.unknown_refs)

    def to_dict(self) -> dict[str, Any]:
        return {
            "included_refs": [ref.to_dict() for ref in self.included_refs],
            "excluded_refs": [ref.to_dict() for ref in self.excluded_refs],
            "unknown_refs": [ref.to_dict() for ref in self.unknown_refs],
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "SourcePartition":
        if not isinstance(value, Mapping):
            raise ValueError("SourcePartition 必须是 JSON 对象")
        _require_exact_fields(
            value,
            {"included_refs", "excluded_refs", "unknown_refs"},
            "SourcePartition",
        )
        parsed: dict[str, tuple[SourceRef, ...]] = {}
        for field in ("included_refs", "excluded_refs", "unknown_refs"):
            raw_refs = value[field]
            if not isinstance(raw_refs, list):
                raise ValueError(f"SourcePartition.{field} 必须是数组")
            parsed[field] = tuple(SourceRef.from_dict(ref) for ref in raw_refs)
        expected_refs = (
            parsed["included_refs"]
            + parsed["excluded_refs"]
            + parsed["unknown_refs"]
        )
        return build_source_partition(
            expected_refs,
            included_refs=parsed["included_refs"],
            excluded_refs=parsed["excluded_refs"],
            unknown_refs=parsed["unknown_refs"],
        )


def build_source_partition(
    expected_refs: tuple[SourceRef, ...],
    *,
    included_refs: tuple[SourceRef, ...],
    excluded_refs: tuple[SourceRef, ...],
    unknown_refs: tuple[SourceRef, ...],
) -> SourcePartition:
    _require_ref_tuple(expected_refs, "expected_refs")
    if not expected_refs:
        raise ValueError("expected_refs 不能为空")
    for field, refs in (
        ("included_refs", included_refs),
        ("excluded_refs", excluded_refs),
        ("unknown_refs", unknown_refs),
    ):
        _require_ref_tuple(refs, field)

    expected_keys = {ref.key for ref in expected_refs}
    partitions = {
        "included_refs": {ref.key for ref in included_refs},
        "excluded_refs": {ref.key for ref in excluded_refs},
        "unknown_refs": {ref.key for ref in unknown_refs},
    }
    classified_keys = set().union(*partitions.values())
    duplicate_keys = (
        (partitions["included_refs"] & partitions["excluded_refs"])
        | (partitions["included_refs"] & partitions["unknown_refs"])
        | (partitions["excluded_refs"] & partitions["unknown_refs"])
    )
    if duplicate_keys:
        raise ValueError(f"来源引用被重复分类：{sorted(duplicate_keys)}")
    missing = expected_keys - classified_keys
    if missing:
        raise ValueError(f"存在未分类来源引用：{sorted(missing)}")
    extra = classified_keys - expected_keys
    if extra:
        raise ValueError(f"存在范围外来源引用：{sorted(extra)}")

    refs_by_key = {ref.key: ref for ref in expected_refs}
    return SourcePartition(
        included_refs=tuple(
            refs_by_key[ref.key]
            for ref in expected_refs
            if ref.key in partitions["included_refs"]
        ),
        excluded_refs=tuple(
            refs_by_key[ref.key]
            for ref in expected_refs
            if ref.key in partitions["excluded_refs"]
        ),
        unknown_refs=tuple(
            refs_by_key[ref.key]
            for ref in expected_refs
            if ref.key in partitions["unknown_refs"]
        ),
    )


def _validate_contract_common(
    *,
    generator_version: str,
    stable_unit_id: StableUnitId,
    original_label: str,
    expected_unit_kind: str,
    source_partition: SourcePartition,
    decision_evidence: tuple[DecisionEvidence, ...],
) -> None:
    _require_nonempty_string(generator_version, "generator_version")
    if not isinstance(stable_unit_id, StableUnitId):
        raise ValueError("stable_unit_id 必须是 StableUnitId")
    if stable_unit_id.unit_kind != expected_unit_kind:
        raise ValueError(f"stable_unit_id.unit_kind 必须为 {expected_unit_kind}")
    _require_nonempty_string(original_label, "original_label")
    if stable_unit_id.original_label != original_label:
        raise ValueError("stable_unit_id.original_label 与对象标签不一致")
    if not isinstance(source_partition, SourcePartition):
        raise ValueError("source_partition 必须是 SourcePartition")
    if any(
        ref.document_sha256 != stable_unit_id.document_sha256
        for ref in source_partition.all_refs
    ):
        raise ValueError("source_partition 与 stable_unit_id 的文档 SHA256 不一致")
    if not isinstance(decision_evidence, tuple) or not decision_evidence:
        raise ValueError("decision_evidence 必须是非空元组")
    if any(not isinstance(evidence, DecisionEvidence) for evidence in decision_evidence):
        raise ValueError("decision_evidence 只能包含 DecisionEvidence")
    partition_keys = {ref.key for ref in source_partition.all_refs}
    for evidence in decision_evidence:
        if any(ref.key not in partition_keys for ref in evidence.source_refs):
            raise ValueError("decision_evidence 引用了来源分区之外的节点")


@dataclass(frozen=True)
class QuestionUnitV2:
    generator_version: str
    stable_unit_id: StableUnitId
    original_question_label: str
    subject: str
    subject_overlay: str | None
    question_type: str
    source_partition: SourcePartition
    decision_evidence: tuple[DecisionEvidence, ...]

    def __post_init__(self) -> None:
        _validate_contract_common(
            generator_version=self.generator_version,
            stable_unit_id=self.stable_unit_id,
            original_label=self.original_question_label,
            expected_unit_kind="question",
            source_partition=self.source_partition,
            decision_evidence=self.decision_evidence,
        )
        _require_nonempty_string(self.subject, "subject")
        if self.subject_overlay is not None:
            _require_nonempty_string(self.subject_overlay, "subject_overlay")
        _require_nonempty_string(self.question_type, "question_type")

    @property
    def has_unknown_refs(self) -> bool:
        return self.source_partition.has_unknown_refs

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": CONTRACT_SCHEMA_VERSION,
            "contract_kind": "question_unit_v2",
            "generator_version": self.generator_version,
            "stable_unit_id": self.stable_unit_id.to_dict(),
            "original_question_label": self.original_question_label,
            "subject": self.subject,
            "subject_overlay": self.subject_overlay,
            "question_type": self.question_type,
            "source_partition": self.source_partition.to_dict(),
            "decision_evidence": [item.to_dict() for item in self.decision_evidence],
        }

    def canonical_bytes(self) -> bytes:
        return canonical_json_bytes(self.to_dict())

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "QuestionUnitV2":
        if not isinstance(value, Mapping):
            raise ValueError("QuestionUnitV2 必须是 JSON 对象")
        _require_exact_fields(
            value,
            {
                "schema_version",
                "contract_kind",
                "generator_version",
                "stable_unit_id",
                "original_question_label",
                "subject",
                "subject_overlay",
                "question_type",
                "source_partition",
                "decision_evidence",
            },
            "QuestionUnitV2",
        )
        if value["schema_version"] != CONTRACT_SCHEMA_VERSION:
            raise ValueError(f"QuestionUnitV2.schema_version 必须为 {CONTRACT_SCHEMA_VERSION}")
        if value["contract_kind"] != "question_unit_v2":
            raise ValueError("QuestionUnitV2.contract_kind 无效")
        evidence = value["decision_evidence"]
        if not isinstance(evidence, list):
            raise ValueError("QuestionUnitV2.decision_evidence 必须是数组")
        return cls(
            generator_version=value["generator_version"],
            stable_unit_id=StableUnitId.from_dict(value["stable_unit_id"]),
            original_question_label=value["original_question_label"],
            subject=value["subject"],
            subject_overlay=value["subject_overlay"],
            question_type=value["question_type"],
            source_partition=SourcePartition.from_dict(value["source_partition"]),
            decision_evidence=tuple(DecisionEvidence.from_dict(item) for item in evidence),
        )


@dataclass(frozen=True)
class ContractTextItem:
    item_id: str
    text: str

    def __post_init__(self) -> None:
        _require_nonempty_string(self.item_id, "item_id")
        if not isinstance(self.text, str):
            raise ValueError("text 必须是字符串")

    def to_dict(self) -> dict[str, str]:
        return {"item_id": self.item_id, "text": self.text}

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "ContractTextItem":
        if not isinstance(value, Mapping):
            raise ValueError("ContractTextItem 必须是 JSON 对象")
        _require_exact_fields(value, {"item_id", "text"}, "ContractTextItem")
        return cls(item_id=value["item_id"], text=value["text"])


@dataclass(frozen=True)
class CanonicalAnswerUnit:
    generator_version: str
    stable_unit_id: StableUnitId
    original_question_label: str
    answer_items: tuple[ContractTextItem, ...]
    analysis_items: tuple[ContractTextItem, ...]
    rich_content_refs: tuple[SourceRef, ...]
    source_partition: SourcePartition
    decision_evidence: tuple[DecisionEvidence, ...]
    review_flags: tuple[str, ...]

    def __post_init__(self) -> None:
        _validate_contract_common(
            generator_version=self.generator_version,
            stable_unit_id=self.stable_unit_id,
            original_label=self.original_question_label,
            expected_unit_kind="answer",
            source_partition=self.source_partition,
            decision_evidence=self.decision_evidence,
        )
        for field in ("answer_items", "analysis_items"):
            items = getattr(self, field)
            if not isinstance(items, tuple) or any(
                not isinstance(item, ContractTextItem) for item in items
            ):
                raise ValueError(f"{field} 只能包含 ContractTextItem")
            item_ids = [item.item_id for item in items]
            if len(item_ids) != len(set(item_ids)):
                raise ValueError(f"{field} 不能包含重复 item_id")
        _require_ref_tuple(self.rich_content_refs, "rich_content_refs")
        partition_keys = {ref.key for ref in self.source_partition.all_refs}
        if any(ref.key not in partition_keys for ref in self.rich_content_refs):
            raise ValueError("rich_content_refs 引用了来源分区之外的节点")
        if not (self.answer_items or self.analysis_items or self.rich_content_refs):
            raise ValueError("CanonicalAnswerUnit 必须包含答案、解析或富内容引用")
        if not isinstance(self.review_flags, tuple) or any(
            not isinstance(flag, str) or not flag.strip() for flag in self.review_flags
        ):
            raise ValueError("review_flags 必须是非空字符串元组")
        if len(self.review_flags) != len(set(self.review_flags)):
            raise ValueError("review_flags 不能重复")

    @property
    def has_unknown_refs(self) -> bool:
        return self.source_partition.has_unknown_refs

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": CONTRACT_SCHEMA_VERSION,
            "contract_kind": "canonical_answer_unit",
            "generator_version": self.generator_version,
            "stable_unit_id": self.stable_unit_id.to_dict(),
            "original_question_label": self.original_question_label,
            "answer_items": [item.to_dict() for item in self.answer_items],
            "analysis_items": [item.to_dict() for item in self.analysis_items],
            "rich_content_refs": [ref.to_dict() for ref in self.rich_content_refs],
            "source_partition": self.source_partition.to_dict(),
            "decision_evidence": [item.to_dict() for item in self.decision_evidence],
            "review_flags": list(self.review_flags),
        }

    def canonical_bytes(self) -> bytes:
        return canonical_json_bytes(self.to_dict())

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "CanonicalAnswerUnit":
        if not isinstance(value, Mapping):
            raise ValueError("CanonicalAnswerUnit 必须是 JSON 对象")
        _require_exact_fields(
            value,
            {
                "schema_version",
                "contract_kind",
                "generator_version",
                "stable_unit_id",
                "original_question_label",
                "answer_items",
                "analysis_items",
                "rich_content_refs",
                "source_partition",
                "decision_evidence",
                "review_flags",
            },
            "CanonicalAnswerUnit",
        )
        if value["schema_version"] != CONTRACT_SCHEMA_VERSION:
            raise ValueError(
                f"CanonicalAnswerUnit.schema_version 必须为 {CONTRACT_SCHEMA_VERSION}"
            )
        if value["contract_kind"] != "canonical_answer_unit":
            raise ValueError("CanonicalAnswerUnit.contract_kind 无效")
        array_fields = (
            "answer_items",
            "analysis_items",
            "rich_content_refs",
            "decision_evidence",
            "review_flags",
        )
        for field in array_fields:
            if not isinstance(value[field], list):
                raise ValueError(f"CanonicalAnswerUnit.{field} 必须是数组")
        return cls(
            generator_version=value["generator_version"],
            stable_unit_id=StableUnitId.from_dict(value["stable_unit_id"]),
            original_question_label=value["original_question_label"],
            answer_items=tuple(ContractTextItem.from_dict(item) for item in value["answer_items"]),
            analysis_items=tuple(
                ContractTextItem.from_dict(item) for item in value["analysis_items"]
            ),
            rich_content_refs=tuple(
                SourceRef.from_dict(ref) for ref in value["rich_content_refs"]
            ),
            source_partition=SourcePartition.from_dict(value["source_partition"]),
            decision_evidence=tuple(
                DecisionEvidence.from_dict(item) for item in value["decision_evidence"]
            ),
            review_flags=tuple(value["review_flags"]),
        )
