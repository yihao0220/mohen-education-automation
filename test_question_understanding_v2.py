from __future__ import annotations

import ast
import hashlib
from pathlib import Path

from shared_core.contracts_v2 import DecisionEvidence, QuestionUnitV2, SourceRef
from shared_core.question_understanding_v2 import (
    RefDecision,
    ProposedQuestionUnit,
    QuestionUnderstandingProposal,
    compile_question_understanding,
)


PROJECT_ROOT = Path(__file__).resolve().parent
DOCUMENT_SHA256 = hashlib.sha256("题目理解V2".encode("utf-8")).hexdigest()


def _ref(
    index: int,
    text: str,
    *,
    kind: str = "paragraph",
    native_id: str | None = None,
) -> SourceRef:
    return SourceRef(
        document_sha256=DOCUMENT_SHA256,
        native_kind=kind,
        native_id=native_id or f"paragraph:{index}",
        occurrence=1,
        text_anchor=text,
    )


def _decision(
    ref: SourceRef,
    disposition: str,
    role: str,
    *,
    confidence: float = 1.0,
    producer: str = "rules",
) -> RefDecision:
    return RefDecision(
        source_ref=ref,
        disposition=disposition,
        role=role,
        evidence=DecisionEvidence(
            rule_id=f"{producer}-{role}",
            reason=f"{producer} 将节点判断为 {role}",
            source_refs=(ref,),
            confidence=confidence,
            generator_version=f"{producer}/1",
        ),
    )


def _unit(
    key: str,
    label: str,
    decisions: tuple[RefDecision, ...],
    *,
    member_labels: tuple[str, ...] | None = None,
    review_flags: tuple[str, ...] = (),
) -> ProposedQuestionUnit:
    return ProposedQuestionUnit(
        unit_key=key,
        original_question_label=label,
        member_question_labels=member_labels or (label,),
        subject="文科",
        subject_overlay="chinese",
        question_type="subjective",
        ref_decisions=decisions,
        review_flags=review_flags,
    )


def _proposal(
    units: tuple[ProposedQuestionUnit, ...],
    *,
    document_decisions: tuple[RefDecision, ...] = (),
    allowed_resets: tuple[int, ...] = (),
    producer: str = "rules",
) -> QuestionUnderstandingProposal:
    return QuestionUnderstandingProposal(
        proposal_id=f"{producer}-proposal",
        producer_kind=producer,
        generator_version=f"{producer}/1",
        units=units,
        document_decisions=document_decisions,
        allowed_label_reset_before=allowed_resets,
    )


def test_compiles_ordinary_continuous_questions_in_source_order() -> None:
    q1 = _ref(1, "1．第一题")
    q1_option = _ref(2, "A．甲")
    q2 = _ref(3, "2．第二题")
    refs = (q1, q1_option, q2)
    proposal = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(q1, "include", "stem"),
                    _decision(q1_option, "include", "option"),
                ),
            ),
            _unit("q2", "2", (_decision(q2, "include", "stem"),)),
        )
    )

    result = compile_question_understanding(refs, (proposal,))

    assert result.status == "compiled_for_review"
    assert result.blocking_issues == ()
    assert [unit.original_question_label for unit in result.units] == ["1", "2"]
    assert result.production_execution_enabled is False
    assert result.keypress_count == 0


def test_leading_material_can_belong_only_to_first_question() -> None:
    material = _ref(1, "阅读材料")
    q1 = _ref(2, "1．第一题")
    q2 = _ref(3, "2．第二题")
    proposal = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(material, "include", "material"),
                    _decision(q1, "include", "stem"),
                ),
            ),
            _unit("q2", "2", (_decision(q2, "include", "stem"),)),
        )
    )

    result = compile_question_understanding((material, q1, q2), (proposal,))

    assert result.blocking_issues == ()
    assert material in result.units[0].source_partition.included_refs
    assert material not in result.units[1].source_partition.all_refs


def test_compiles_complete_reading_group_as_one_merged_unit() -> None:
    material = _ref(1, "阅读材料")
    q1 = _ref(2, "1．第一题")
    q2 = _ref(3, "2．第二题")
    q3 = _ref(4, "3．第三题")
    proposal = _proposal(
        (
            _unit(
                "reading-1-3",
                "1-3",
                (
                    _decision(material, "include", "material"),
                    _decision(q1, "include", "stem"),
                    _decision(q2, "include", "stem"),
                    _decision(q3, "include", "stem"),
                ),
                member_labels=("1", "2", "3"),
            ),
        )
    )

    result = compile_question_understanding((material, q1, q2, q3), (proposal,))

    assert result.blocking_issues == ()
    assert len(result.units) == 1
    assert result.units[0].original_question_label == "1-3"


def test_middle_heading_is_excluded_without_losing_surrounding_content() -> None:
    before = _ref(1, "材料正文")
    heading = _ref(2, "（一）高考专练")
    after = _ref(3, "1．题干")
    proposal = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(before, "include", "material"),
                    _decision(heading, "exclude", "internal_heading"),
                    _decision(after, "include", "stem"),
                ),
            ),
        )
    )

    result = compile_question_understanding((before, heading, after), (proposal,))

    assert result.blocking_issues == ()
    unit = result.units[0]
    assert unit.source_partition.included_refs == (before, after)
    assert unit.source_partition.excluded_refs == (heading,)
    assert QuestionUnitV2.from_dict(unit.to_dict()) == unit


def test_excluded_heading_with_native_table_is_blocked() -> None:
    heading = _ref(1, "专题表格")
    table = _ref(
        1,
        "专题表格",
        kind="table",
        native_id="paragraph:1/table:1",
    )
    stem = _ref(2, "1．题干")
    proposal = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(heading, "exclude", "internal_heading"),
                    _decision(table, "include", "table"),
                    _decision(stem, "include", "stem"),
                ),
            ),
        )
    )

    result = compile_question_understanding((heading, table, stem), (proposal,))

    assert result.units == ()
    assert {issue.code for issue in result.blocking_issues} == {
        "EXCLUDED_HEADING_HAS_NATIVE_CONTENT"
    }


def test_decorative_media_is_excluded_and_question_media_is_included() -> None:
    decorative_paragraph = _ref(1, "")
    decorative = _ref(
        1,
        "页眉装饰图",
        kind="media",
        native_id="paragraph:1/media:decorative",
    )
    stem = _ref(2, "1．根据图示作答")
    question_media = _ref(
        2,
        "题图",
        kind="media",
        native_id="paragraph:2/media:question",
    )
    proposal = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(stem, "include", "stem"),
                    _decision(question_media, "include", "question_media"),
                ),
            ),
        ),
        document_decisions=(
            _decision(decorative_paragraph, "exclude", "top_level_boundary"),
            _decision(decorative, "exclude", "decorative_media"),
        ),
    )

    result = compile_question_understanding(
        (decorative_paragraph, decorative, stem, question_media),
        (proposal,),
    )

    assert result.blocking_issues == ()
    assert question_media in result.units[0].source_partition.included_refs
    assert decorative not in result.units[0].source_partition.all_refs


def test_duplicate_labels_require_explicit_reset_and_keep_distinct_ids() -> None:
    first = _ref(1, "1．第一篇第一题")
    second = _ref(2, "1．第二篇第一题")
    units = (
        _unit("first-1", "1", (_decision(first, "include", "stem"),)),
        _unit("second-1", "1", (_decision(second, "include", "stem"),)),
    )

    blocked = compile_question_understanding((first, second), (_proposal(units),))
    allowed = compile_question_understanding(
        (first, second),
        (_proposal(units, allowed_resets=(2,)),),
    )

    assert "QUESTION_LABEL_RESET_NOT_ALLOWED" in {
        issue.code for issue in blocked.blocking_issues
    }
    assert allowed.blocking_issues == ()
    assert len({unit.stable_unit_id.value for unit in allowed.units}) == 2


def test_candidate_json_roundtrip_and_compiled_bytes_are_stable() -> None:
    stem = _ref(1, "1．题干")
    proposal = _proposal(
        (_unit("q1", "1", (_decision(stem, "include", "stem"),)),)
    )

    restored = QuestionUnderstandingProposal.from_dict(proposal.to_dict())
    first = compile_question_understanding((stem,), (proposal,))
    second = compile_question_understanding((stem,), (restored,))

    assert restored == proposal
    assert first.canonical_bytes() == second.canonical_bytes()


def test_agreeing_ai_low_confidence_evidence_still_blocks() -> None:
    stem = _ref(1, "1．题干")
    rules = _proposal(
        (_unit("q1", "1", (_decision(stem, "include", "stem"),)),)
    )
    ai = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(
                        stem,
                        "include",
                        "stem",
                        confidence=0.6,
                        producer="ai",
                    ),
                ),
            ),
        ),
        producer="ai",
    )

    result = compile_question_understanding((stem,), (rules, ai))

    assert result.units == ()
    assert "LOW_CONFIDENCE_DECISION" in {
        issue.code for issue in result.blocking_issues
    }


def test_unknown_missing_overlap_low_confidence_and_rule_ai_conflict_block_atomically() -> None:
    first = _ref(1, "1．第一题")
    second = _ref(2, "2．第二题")
    unknown = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(first, "include", "stem"),
                    _decision(second, "unknown", "unknown"),
                ),
            ),
        )
    )
    missing = _proposal((_unit("q1", "1", (_decision(first, "include", "stem"),)),))
    overlap = _proposal(
        (
            _unit("q1", "1", (_decision(first, "include", "stem"),)),
            _unit("q2", "2", (_decision(first, "include", "stem"),)),
        ),
        document_decisions=(_decision(second, "exclude", "top_level_boundary"),),
    )
    low = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(first, "include", "stem", confidence=0.6),
                    _decision(second, "include", "stem"),
                ),
            ),
        )
    )
    ai_conflict = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(first, "include", "stem", producer="ai"),
                    _decision(second, "include", "option", producer="ai"),
                ),
            ),
        ),
        producer="ai",
    )
    rules = _proposal(
        (
            _unit(
                "q1",
                "1",
                (
                    _decision(first, "include", "stem"),
                    _decision(second, "include", "stem"),
                ),
            ),
        )
    )

    cases = (
        (unknown,),
        (missing,),
        (overlap,),
        (low,),
        (rules, ai_conflict),
    )
    expected_codes = (
        "UNKNOWN_SOURCE_REF",
        "UNCLASSIFIED_SOURCE_REF",
        "ILLEGAL_SOURCE_OVERLAP",
        "LOW_CONFIDENCE_DECISION",
        "PROPOSAL_CONFLICT",
    )
    for proposals, expected_code in zip(cases, expected_codes):
        result = compile_question_understanding((first, second), proposals)
        assert result.units == ()
        assert expected_code in {issue.code for issue in result.blocking_issues}


def test_candidate_and_compiler_modules_do_not_import_execution_dependencies() -> None:
    forbidden = {
        "pyautogui",
        "win32com",
        "wps_helper",
        "answer_input",
        "core_parser",
        "document_render",
    }
    path = PROJECT_ROOT / "shared_core/question_understanding_v2.py"
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    imported: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported.add(node.module)
    assert not {
        name
        for name in imported
        if any(part in forbidden for part in name.split("."))
    }
