from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Mapping, Sequence

from .contracts_v2 import (
    CONTRACT_SCHEMA_VERSION,
    DecisionEvidence,
    QuestionUnitV2,
    SourceRef,
    StableUnitId,
    build_source_partition,
    canonical_json_bytes,
)


QUESTION_UNDERSTANDING_SCHEMA_VERSION = "1.0"
QUESTION_UNDERSTANDING_GENERATOR_VERSION = "mohen-question-understanding-v2/1"
DEFAULT_MINIMUM_CONFIDENCE = 0.75

DISPOSITIONS = ("include", "exclude", "unknown")
INCLUDED_ROLES = (
    "material",
    "shared_material",
    "stem",
    "option",
    "subquestion",
    "table",
    "question_media",
    "formula",
)
EXCLUDED_ROLES = (
    "internal_heading",
    "top_level_boundary",
    "decorative_media",
)
UNKNOWN_ROLES = ("unknown",)

_ROLES_BY_DISPOSITION = {
    "include": set(INCLUDED_ROLES),
    "exclude": set(EXCLUDED_ROLES),
    "unknown": set(UNKNOWN_ROLES),
}
_ROLES_BY_NATIVE_KIND = {
    "paragraph": {
        "material",
        "shared_material",
        "stem",
        "option",
        "subquestion",
        "internal_heading",
        "top_level_boundary",
        "unknown",
    },
    "table": {"table", "unknown"},
    "media": {"question_media", "decorative_media", "unknown"},
    "formula": {"formula", "unknown"},
}
_INTEGER_LABEL = re.compile(r"^\d+$")


def _require_nonempty_string(value: str, field: str) -> None:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} 必须是非空字符串")


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


@dataclass(frozen=True)
class RefDecision:
    source_ref: SourceRef
    disposition: str
    role: str
    evidence: DecisionEvidence

    def __post_init__(self) -> None:
        if not isinstance(self.source_ref, SourceRef):
            raise ValueError("source_ref 必须是 SourceRef")
        if self.disposition not in DISPOSITIONS:
            raise ValueError(f"disposition 必须是 {DISPOSITIONS} 之一")
        if self.role not in _ROLES_BY_DISPOSITION[self.disposition]:
            raise ValueError(
                f"role={self.role!r} 与 disposition={self.disposition!r} 不兼容"
            )
        if self.role not in _ROLES_BY_NATIVE_KIND[self.source_ref.native_kind]:
            raise ValueError(
                f"role={self.role!r} 与 native_kind={self.source_ref.native_kind!r} 不兼容"
            )
        if not isinstance(self.evidence, DecisionEvidence):
            raise ValueError("evidence 必须是 DecisionEvidence")
        if self.source_ref.key not in {ref.key for ref in self.evidence.source_refs}:
            raise ValueError("evidence 必须引用当前 source_ref")

    def semantic_tuple(self) -> tuple[str, str, str]:
        return (self.source_ref.key, self.disposition, self.role)

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_ref": self.source_ref.to_dict(),
            "disposition": self.disposition,
            "role": self.role,
            "evidence": self.evidence.to_dict(),
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "RefDecision":
        if not isinstance(value, Mapping):
            raise ValueError("RefDecision 必须是 JSON 对象")
        _require_exact_fields(
            value,
            {"source_ref", "disposition", "role", "evidence"},
            "RefDecision",
        )
        return cls(
            source_ref=SourceRef.from_dict(value["source_ref"]),
            disposition=value["disposition"],
            role=value["role"],
            evidence=DecisionEvidence.from_dict(value["evidence"]),
        )


@dataclass(frozen=True)
class ProposedQuestionUnit:
    unit_key: str
    original_question_label: str
    member_question_labels: tuple[str, ...]
    subject: str
    subject_overlay: str | None
    question_type: str
    ref_decisions: tuple[RefDecision, ...]
    review_flags: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        for field in (
            "unit_key",
            "original_question_label",
            "subject",
            "question_type",
        ):
            _require_nonempty_string(getattr(self, field), field)
        if self.subject_overlay is not None:
            _require_nonempty_string(self.subject_overlay, "subject_overlay")
        if (
            not isinstance(self.member_question_labels, tuple)
            or not self.member_question_labels
            or any(
                not isinstance(label, str) or not label.strip()
                for label in self.member_question_labels
            )
        ):
            raise ValueError("member_question_labels 必须是非空题号元组")
        if not isinstance(self.ref_decisions, tuple) or not self.ref_decisions:
            raise ValueError("ref_decisions 必须是非空 RefDecision 元组")
        if any(not isinstance(item, RefDecision) for item in self.ref_decisions):
            raise ValueError("ref_decisions 只能包含 RefDecision")
        if not isinstance(self.review_flags, tuple) or any(
            not isinstance(flag, str) or not flag.strip() for flag in self.review_flags
        ):
            raise ValueError("review_flags 必须是非空字符串元组")
        if len(self.review_flags) != len(set(self.review_flags)):
            raise ValueError("review_flags 不能重复")

    def semantic_tuple(self) -> tuple[Any, ...]:
        return (
            self.unit_key,
            self.original_question_label,
            self.member_question_labels,
            self.subject,
            self.subject_overlay,
            self.question_type,
            tuple(item.semantic_tuple() for item in self.ref_decisions),
            self.review_flags,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "unit_key": self.unit_key,
            "original_question_label": self.original_question_label,
            "member_question_labels": list(self.member_question_labels),
            "subject": self.subject,
            "subject_overlay": self.subject_overlay,
            "question_type": self.question_type,
            "ref_decisions": [item.to_dict() for item in self.ref_decisions],
            "review_flags": list(self.review_flags),
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "ProposedQuestionUnit":
        if not isinstance(value, Mapping):
            raise ValueError("ProposedQuestionUnit 必须是 JSON 对象")
        _require_exact_fields(
            value,
            {
                "unit_key",
                "original_question_label",
                "member_question_labels",
                "subject",
                "subject_overlay",
                "question_type",
                "ref_decisions",
                "review_flags",
            },
            "ProposedQuestionUnit",
        )
        for field in ("member_question_labels", "ref_decisions", "review_flags"):
            if not isinstance(value[field], list):
                raise ValueError(f"ProposedQuestionUnit.{field} 必须是数组")
        return cls(
            unit_key=value["unit_key"],
            original_question_label=value["original_question_label"],
            member_question_labels=tuple(value["member_question_labels"]),
            subject=value["subject"],
            subject_overlay=value["subject_overlay"],
            question_type=value["question_type"],
            ref_decisions=tuple(
                RefDecision.from_dict(item) for item in value["ref_decisions"]
            ),
            review_flags=tuple(value["review_flags"]),
        )


@dataclass(frozen=True)
class QuestionUnderstandingProposal:
    proposal_id: str
    producer_kind: str
    generator_version: str
    units: tuple[ProposedQuestionUnit, ...]
    document_decisions: tuple[RefDecision, ...] = ()
    allowed_label_reset_before: tuple[int, ...] = ()

    def __post_init__(self) -> None:
        for field in ("proposal_id", "producer_kind", "generator_version"):
            _require_nonempty_string(getattr(self, field), field)
        if not isinstance(self.units, tuple) or not self.units:
            raise ValueError("units 必须是非空 ProposedQuestionUnit 元组")
        if any(not isinstance(unit, ProposedQuestionUnit) for unit in self.units):
            raise ValueError("units 只能包含 ProposedQuestionUnit")
        unit_keys = [unit.unit_key for unit in self.units]
        if len(unit_keys) != len(set(unit_keys)):
            raise ValueError("units 不能包含重复 unit_key")
        if not isinstance(self.document_decisions, tuple) or any(
            not isinstance(item, RefDecision) for item in self.document_decisions
        ):
            raise ValueError("document_decisions 只能包含 RefDecision")
        if any(item.disposition == "include" for item in self.document_decisions):
            raise ValueError("document_decisions 不能包含 include")
        resets = self.allowed_label_reset_before
        if (
            not isinstance(resets, tuple)
            or any(
                isinstance(index, bool) or not isinstance(index, int) or index < 2
                for index in resets
            )
            or tuple(sorted(set(resets))) != resets
        ):
            raise ValueError("allowed_label_reset_before 必须是从2开始的升序唯一整数元组")
        if any(index > len(self.units) for index in resets):
            raise ValueError("allowed_label_reset_before 超出题块范围")

    def semantic_tuple(self) -> tuple[Any, ...]:
        return (
            tuple(unit.semantic_tuple() for unit in self.units),
            tuple(item.semantic_tuple() for item in self.document_decisions),
            self.allowed_label_reset_before,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": QUESTION_UNDERSTANDING_SCHEMA_VERSION,
            "proposal_id": self.proposal_id,
            "producer_kind": self.producer_kind,
            "generator_version": self.generator_version,
            "units": [unit.to_dict() for unit in self.units],
            "document_decisions": [
                item.to_dict() for item in self.document_decisions
            ],
            "allowed_label_reset_before": list(self.allowed_label_reset_before),
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "QuestionUnderstandingProposal":
        if not isinstance(value, Mapping):
            raise ValueError("QuestionUnderstandingProposal 必须是 JSON 对象")
        _require_exact_fields(
            value,
            {
                "schema_version",
                "proposal_id",
                "producer_kind",
                "generator_version",
                "units",
                "document_decisions",
                "allowed_label_reset_before",
            },
            "QuestionUnderstandingProposal",
        )
        if value["schema_version"] != QUESTION_UNDERSTANDING_SCHEMA_VERSION:
            raise ValueError(
                "QuestionUnderstandingProposal.schema_version 必须为 "
                f"{QUESTION_UNDERSTANDING_SCHEMA_VERSION}"
            )
        for field in ("units", "document_decisions", "allowed_label_reset_before"):
            if not isinstance(value[field], list):
                raise ValueError(f"QuestionUnderstandingProposal.{field} 必须是数组")
        return cls(
            proposal_id=value["proposal_id"],
            producer_kind=value["producer_kind"],
            generator_version=value["generator_version"],
            units=tuple(ProposedQuestionUnit.from_dict(item) for item in value["units"]),
            document_decisions=tuple(
                RefDecision.from_dict(item) for item in value["document_decisions"]
            ),
            allowed_label_reset_before=tuple(value["allowed_label_reset_before"]),
        )


@dataclass(frozen=True)
class CompilationIssue:
    code: str
    message: str
    source_ref_keys: tuple[str, ...] = ()
    unit_keys: tuple[str, ...] = ()
    severity: str = "blocker"

    def __post_init__(self) -> None:
        _require_nonempty_string(self.code, "code")
        _require_nonempty_string(self.message, "message")
        if self.severity != "blocker":
            raise ValueError("第二部分编译问题当前只能是 blocker")

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "severity": self.severity,
            "message": self.message,
            "source_ref_keys": list(self.source_ref_keys),
            "unit_keys": list(self.unit_keys),
        }


@dataclass(frozen=True)
class QuestionUnderstandingResult:
    source_sha256: str
    generator_version: str
    status: str
    units: tuple[QuestionUnitV2, ...]
    blocking_issues: tuple[CompilationIssue, ...]
    proposal_ids: tuple[str, ...]

    @property
    def production_execution_enabled(self) -> bool:
        return False

    @property
    def keypress_count(self) -> int:
        return 0

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": QUESTION_UNDERSTANDING_SCHEMA_VERSION,
            "artifact_kind": "question_understanding_v2",
            "question_contract_schema_version": CONTRACT_SCHEMA_VERSION,
            "generator_version": self.generator_version,
            "mode": "shadow",
            "status": self.status,
            "production_execution_enabled": False,
            "keypress_count": 0,
            "source_sha256": self.source_sha256,
            "proposal_ids": list(self.proposal_ids),
            "unit_count": len(self.units),
            "units": [unit.to_dict() for unit in self.units],
            "blocking_issues": [issue.to_dict() for issue in self.blocking_issues],
            "limitations": [
                "结果仅供影子审核，不是F1动作计划。",
                "结果不包含WPS Range，也不会触发F1/F2/F3/F4。",
                "构造测试证明编译规则，不证明AI或所有真实DOCX判断正确。",
            ],
        }

    def canonical_bytes(self) -> bytes:
        return canonical_json_bytes(self.to_dict())


def _issue(
    code: str,
    message: str,
    *,
    refs: Sequence[SourceRef] = (),
    unit_keys: Sequence[str] = (),
) -> CompilationIssue:
    return CompilationIssue(
        code=code,
        message=message,
        source_ref_keys=tuple(ref.key for ref in refs),
        unit_keys=tuple(unit_keys),
    )


def _dedupe_issues(issues: Sequence[CompilationIssue]) -> tuple[CompilationIssue, ...]:
    unique: list[CompilationIssue] = []
    seen: set[tuple[Any, ...]] = set()
    for issue in issues:
        key = (
            issue.code,
            issue.message,
            issue.source_ref_keys,
            issue.unit_keys,
        )
        if key not in seen:
            seen.add(key)
            unique.append(issue)
    return tuple(unique)


def _numeric_members(unit: ProposedQuestionUnit) -> tuple[int, ...] | None:
    if not all(_INTEGER_LABEL.fullmatch(label) for label in unit.member_question_labels):
        return None
    return tuple(int(label) for label in unit.member_question_labels)


def _compile_blocked(
    source_sha256: str,
    proposals: Sequence[QuestionUnderstandingProposal],
    issues: Sequence[CompilationIssue],
    generator_version: str,
) -> QuestionUnderstandingResult:
    return QuestionUnderstandingResult(
        source_sha256=source_sha256,
        generator_version=generator_version,
        status="blocked",
        units=(),
        blocking_issues=_dedupe_issues(issues),
        proposal_ids=tuple(proposal.proposal_id for proposal in proposals),
    )


def compile_question_understanding(
    source_refs: Sequence[SourceRef],
    proposals: Sequence[QuestionUnderstandingProposal],
    *,
    minimum_confidence: float = DEFAULT_MINIMUM_CONFIDENCE,
    generator_version: str = QUESTION_UNDERSTANDING_GENERATOR_VERSION,
) -> QuestionUnderstandingResult:
    refs = tuple(source_refs)
    proposal_list = tuple(proposals)
    if not refs:
        return _compile_blocked(
            "",
            proposal_list,
            (CompilationIssue("EMPTY_SOURCE_REFS", "上游没有提供任何 SourceRef"),),
            generator_version,
        )
    source_sha256 = refs[0].document_sha256
    issues: list[CompilationIssue] = []
    if len({ref.key for ref in refs}) != len(refs):
        issues.append(CompilationIssue("DUPLICATE_SOURCE_REF", "上游 SourceRef 发生重复"))
    if any(ref.document_sha256 != source_sha256 for ref in refs):
        issues.append(
            CompilationIssue("MIXED_SOURCE_DOCUMENTS", "一次编译只能处理同一份原题")
        )
    if not proposal_list:
        issues.append(CompilationIssue("MISSING_PROPOSAL", "没有候选判断可供编译"))
        return _compile_blocked(source_sha256, proposal_list, issues, generator_version)
    if len({proposal.proposal_id for proposal in proposal_list}) != len(proposal_list):
        issues.append(CompilationIssue("DUPLICATE_PROPOSAL_ID", "proposal_id 不能重复"))

    semantic = proposal_list[0].semantic_tuple()
    if any(proposal.semantic_tuple() != semantic for proposal in proposal_list[1:]):
        issues.append(
            CompilationIssue(
                "PROPOSAL_CONFLICT",
                "规则、AI或其他候选判断对题块关系给出了不一致结果，必须人工裁决",
            )
        )
        return _compile_blocked(source_sha256, proposal_list, issues, generator_version)

    selected = proposal_list[0]
    refs_by_key = {ref.key: ref for ref in refs}
    source_position = {ref.key: index for index, ref in enumerate(refs)}
    decision_uses: dict[str, list[tuple[str, RefDecision]]] = {}

    def register_decision(owner: str, decision: RefDecision) -> None:
        if decision.source_ref.key not in refs_by_key:
            issues.append(
                _issue(
                    "SOURCE_REF_NOT_FOUND",
                    "候选判断引用了上游不存在的 SourceRef",
                    refs=(decision.source_ref,),
                    unit_keys=(() if owner == "__document__" else (owner,)),
                )
            )
        decision_uses.setdefault(decision.source_ref.key, []).append((owner, decision))
        if decision.disposition == "unknown":
            issues.append(
                _issue(
                    "UNKNOWN_SOURCE_REF",
                    "候选判断仍包含 unknown 节点，必须人工审核",
                    refs=(decision.source_ref,),
                    unit_keys=(() if owner == "__document__" else (owner,)),
                )
            )

    for proposal in proposal_list:
        for unit in proposal.units:
            unit_ref_keys = {
                decision.source_ref.key for decision in unit.ref_decisions
            }
            for decision in unit.ref_decisions:
                if decision.evidence.confidence < minimum_confidence:
                    issues.append(
                        _issue(
                            "LOW_CONFIDENCE_DECISION",
                            f"候选判断置信度低于 {minimum_confidence:.2f}",
                            refs=(decision.source_ref,),
                            unit_keys=(unit.unit_key,),
                        )
                    )
                outside_evidence = tuple(
                    ref
                    for ref in decision.evidence.source_refs
                    if ref.key not in unit_ref_keys
                )
                if outside_evidence:
                    issues.append(
                        _issue(
                            "EVIDENCE_OUTSIDE_UNIT",
                            "题块判断证据引用了该候选题块范围之外的节点",
                            refs=outside_evidence,
                            unit_keys=(unit.unit_key,),
                        )
                    )
        document_ref_keys = {
            decision.source_ref.key for decision in proposal.document_decisions
        }
        for decision in proposal.document_decisions:
            if decision.evidence.confidence < minimum_confidence:
                issues.append(
                    _issue(
                        "LOW_CONFIDENCE_DECISION",
                        f"文档级判断置信度低于 {minimum_confidence:.2f}",
                        refs=(decision.source_ref,),
                    )
                )
            outside_evidence = tuple(
                ref
                for ref in decision.evidence.source_refs
                if ref.key not in document_ref_keys
            )
            if outside_evidence:
                issues.append(
                    _issue(
                        "EVIDENCE_OUTSIDE_DOCUMENT_DECISIONS",
                        "文档级判断证据引用了文档级候选范围之外的节点",
                        refs=outside_evidence,
                    )
                )

    for unit in selected.units:
        seen_in_unit: set[str] = set()
        for decision in unit.ref_decisions:
            if decision.source_ref.key in seen_in_unit:
                issues.append(
                    _issue(
                        "DUPLICATE_UNIT_REF_DECISION",
                        "同一题块对同一 SourceRef 重复作出判断",
                        refs=(decision.source_ref,),
                        unit_keys=(unit.unit_key,),
                    )
                )
            seen_in_unit.add(decision.source_ref.key)
            register_decision(unit.unit_key, decision)
        if unit.review_flags:
            issues.append(
                CompilationIssue(
                    "CANDIDATE_REQUIRES_REVIEW",
                    "候选题块带有未解决的审核标记：" + "、".join(unit.review_flags),
                    unit_keys=(unit.unit_key,),
                )
            )
    seen_document: set[str] = set()
    for decision in selected.document_decisions:
        if decision.source_ref.key in seen_document:
            issues.append(
                _issue(
                    "DUPLICATE_DOCUMENT_REF_DECISION",
                    "文档级判断对同一 SourceRef 重复分类",
                    refs=(decision.source_ref,),
                )
            )
        seen_document.add(decision.source_ref.key)
        register_decision("__document__", decision)

    for ref in refs:
        if ref.key not in decision_uses:
            issues.append(
                _issue(
                    "UNCLASSIFIED_SOURCE_REF",
                    "来源节点未进入任何题块，也没有文档级排除或 unknown 判断",
                    refs=(ref,),
                )
            )

    for ref_key, uses in decision_uses.items():
        if len(uses) <= 1:
            continue
        shared_material = all(
            owner != "__document__"
            and decision.disposition == "include"
            and decision.role == "shared_material"
            for owner, decision in uses
        )
        distinct_owners = len({owner for owner, _ in uses}) == len(uses)
        if not (shared_material and distinct_owners):
            ref = refs_by_key.get(ref_key, uses[0][1].source_ref)
            issues.append(
                _issue(
                    "ILLEGAL_SOURCE_OVERLAP",
                    "同一来源节点被多个题块或文档级判断非法重复使用",
                    refs=(ref,),
                    unit_keys=tuple(owner for owner, _ in uses if owner != "__document__"),
                )
            )

    for unit in selected.units:
        excluded_headings = [
            decision.source_ref
            for decision in unit.ref_decisions
            if decision.disposition == "exclude"
            and decision.role == "internal_heading"
            and decision.source_ref.native_kind == "paragraph"
        ]
        unit_rich_refs = [
            decision.source_ref
            for decision in unit.ref_decisions
            if decision.source_ref.native_kind in {"table", "media", "formula"}
        ]
        for heading in excluded_headings:
            descendants = [
                ref
                for ref in unit_rich_refs
                if ref.native_id.startswith(f"{heading.native_id}/")
            ]
            if descendants:
                issues.append(
                    _issue(
                        "EXCLUDED_HEADING_HAS_NATIVE_CONTENT",
                        "拟排除的题内标题同时承载表格、图片或公式，不能安全删除整行",
                        refs=(heading, *descendants),
                        unit_keys=(unit.unit_key,),
                    )
                )

    previous_last_label: int | None = None
    previous_start_position = -1
    for unit_index, unit in enumerate(selected.units, 1):
        numeric_members = _numeric_members(unit)
        if numeric_members and len(numeric_members) > 1:
            expected = tuple(range(numeric_members[0], numeric_members[0] + len(numeric_members)))
            if numeric_members != expected:
                issues.append(
                    CompilationIssue(
                        "NONCONTIGUOUS_MERGED_LABELS",
                        "合并题块只能包含连续递增的数字题号",
                        unit_keys=(unit.unit_key,),
                    )
                )
        included = [
            decision
            for decision in unit.ref_decisions
            if decision.disposition == "include"
        ]
        if not included:
            issues.append(
                CompilationIssue(
                    "EMPTY_INCLUDED_REFS",
                    "题块没有任何 included_ref",
                    unit_keys=(unit.unit_key,),
                )
            )
        else:
            primary = [
                item for item in included if item.role != "shared_material"
            ] or included
            known_positions = [
                source_position[item.source_ref.key]
                for item in primary
                if item.source_ref.key in source_position
            ]
            if known_positions:
                start_position = min(known_positions)
                if start_position <= previous_start_position:
                    issues.append(
                        CompilationIssue(
                            "QUESTION_UNIT_ORDER_INVALID",
                            "题块顺序与来源节点顺序不一致",
                            unit_keys=(unit.unit_key,),
                        )
                    )
                previous_start_position = start_position
        if numeric_members:
            first_label = numeric_members[0]
            if (
                previous_last_label is not None
                and first_label <= previous_last_label
                and unit_index not in selected.allowed_label_reset_before
            ):
                issues.append(
                    CompilationIssue(
                        "QUESTION_LABEL_RESET_NOT_ALLOWED",
                        "题号重复或回退，但当前位置没有项目允许的重置声明",
                        unit_keys=(unit.unit_key,),
                    )
                )
            previous_last_label = numeric_members[-1]

    if issues:
        return _compile_blocked(source_sha256, proposal_list, issues, generator_version)

    compiled_units: list[QuestionUnitV2] = []
    for unit_index, proposed in enumerate(selected.units, 1):
        decisions_by_key = {
            decision.source_ref.key: decision for decision in proposed.ref_decisions
        }
        expected_refs = tuple(ref for ref in refs if ref.key in decisions_by_key)
        included_refs = tuple(
            ref
            for ref in expected_refs
            if decisions_by_key[ref.key].disposition == "include"
        )
        excluded_refs = tuple(
            ref
            for ref in expected_refs
            if decisions_by_key[ref.key].disposition == "exclude"
        )
        unknown_refs = tuple(
            ref
            for ref in expected_refs
            if decisions_by_key[ref.key].disposition == "unknown"
        )
        partition = build_source_partition(
            expected_refs,
            included_refs=included_refs,
            excluded_refs=excluded_refs,
            unknown_refs=unknown_refs,
        )
        evidence: list[DecisionEvidence] = []
        for proposal in proposal_list:
            matching = proposal.units[unit_index - 1]
            evidence.extend(item.evidence for item in matching.ref_decisions)
        compiled_units.append(
            QuestionUnitV2(
                generator_version=generator_version,
                stable_unit_id=StableUnitId.create(
                    document_sha256=source_sha256,
                    unit_kind="question",
                    occurrence=unit_index,
                    original_label=proposed.original_question_label,
                ),
                original_question_label=proposed.original_question_label,
                subject=proposed.subject,
                subject_overlay=proposed.subject_overlay,
                question_type=proposed.question_type,
                source_partition=partition,
                decision_evidence=tuple(evidence),
            )
        )

    return QuestionUnderstandingResult(
        source_sha256=source_sha256,
        generator_version=generator_version,
        status="compiled_for_review",
        units=tuple(compiled_units),
        blocking_issues=(),
        proposal_ids=tuple(proposal.proposal_id for proposal in proposal_list),
    )
