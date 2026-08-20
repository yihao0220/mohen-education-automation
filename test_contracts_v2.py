from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from shared_core.contracts_v2 import (
    CONTRACT_SCHEMA_VERSION,
    NATIVE_KINDS,
    CanonicalAnswerUnit,
    ContractTextItem,
    DecisionEvidence,
    QuestionUnitV2,
    SourceRef,
    StableUnitId,
    build_source_partition,
)


PROJECT_ROOT = Path(__file__).resolve().parent
SCHEMA_DIR = PROJECT_ROOT / "shared_core" / "schemas"
GOLDEN_PATH = PROJECT_ROOT / "test_contracts_v2" / "contract_fixture_golden.json"
DOCUMENT_SHA256 = hashlib.sha256("只读原题".encode("utf-8")).hexdigest()


def _ref(
    native_kind: str = "paragraph",
    native_id: str = "paragraph:1",
    occurrence: int = 1,
    text_anchor: str = "1．题干",
) -> SourceRef:
    return SourceRef(
        document_sha256=DOCUMENT_SHA256,
        native_kind=native_kind,
        native_id=native_id,
        occurrence=occurrence,
        text_anchor=text_anchor,
    )


def test_source_ref_validates_identity_fields() -> None:
    assert _ref().key.startswith(f"{DOCUMENT_SHA256}:paragraph:")

    with pytest.raises(ValueError, match="document_sha256"):
        SourceRef("not-a-sha", "paragraph", "paragraph:1", 1, "题干")
    with pytest.raises(ValueError, match="native_kind"):
        SourceRef(DOCUMENT_SHA256, "shape", "shape:1", 1, "题图")
    with pytest.raises(ValueError, match="native_id"):
        SourceRef(DOCUMENT_SHA256, "paragraph", "", 1, "题干")
    with pytest.raises(ValueError, match="occurrence"):
        SourceRef(DOCUMENT_SHA256, "paragraph", "paragraph:1", 0, "题干")


def test_native_kinds_have_distinct_reference_keys() -> None:
    refs = [
        _ref(kind, "native:1", text_anchor="同一锚点")
        for kind in ("paragraph", "table", "media", "formula")
    ]

    assert NATIVE_KINDS == ("paragraph", "table", "media", "formula")
    assert len({ref.key for ref in refs}) == 4


def test_repeated_question_labels_have_distinct_stable_ids() -> None:
    first = StableUnitId.create(
        document_sha256=DOCUMENT_SHA256,
        unit_kind="question",
        occurrence=1,
        original_label="1",
    )
    second = StableUnitId.create(
        document_sha256=DOCUMENT_SHA256,
        unit_kind="question",
        occurrence=2,
        original_label="1",
    )

    assert first.value != second.value
    assert StableUnitId.create(
        document_sha256=DOCUMENT_SHA256,
        unit_kind="question",
        occurrence=1,
        original_label="1",
    ) == first


def test_tampered_stable_id_is_rejected_on_read() -> None:
    stable_id = StableUnitId.create(
        document_sha256=DOCUMENT_SHA256,
        unit_kind="answer",
        occurrence=1,
        original_label="1",
    )
    payload = stable_id.to_dict()
    payload["value"] = "mohen:v2:" + "0" * 64

    with pytest.raises(ValueError, match="value"):
        StableUnitId.from_dict(payload)


def test_decision_evidence_validates_confidence_and_source_refs() -> None:
    evidence = DecisionEvidence(
        rule_id="legacy-question-shadow-adapter",
        reason="旧 QuestionUnit 连续范围仅用于影子对照",
        source_refs=(_ref(),),
        confidence=0.75,
        generator_version="mohen-contract-v2/1",
    )

    assert evidence.to_dict()["source_refs"][0]["native_kind"] == "paragraph"
    with pytest.raises(ValueError, match="confidence"):
        DecisionEvidence("rule", "reason", (_ref(),), 1.1, "generator")
    with pytest.raises(ValueError, match="rule_id"):
        DecisionEvidence("", "reason", (_ref(),), 0.5, "generator")


@pytest.mark.parametrize(
    "filename,expected_title",
    [
        ("question-unit-v2.schema.json", "QuestionUnitV2"),
        ("canonical-answer-unit-v2.schema.json", "CanonicalAnswerUnit"),
    ],
)
def test_json_schema_has_versioned_shared_definitions(
    filename: str,
    expected_title: str,
) -> None:
    schema = json.loads((SCHEMA_DIR / filename).read_text(encoding="utf-8"))

    assert schema["$schema"] == "https://json-schema.org/draft/2020-12/schema"
    assert schema["title"] == expected_title
    assert schema["properties"]["schema_version"]["const"] == CONTRACT_SCHEMA_VERSION
    assert schema["additionalProperties"] is False
    assert set(schema["$defs"]) >= {
        "sourceRef",
        "stableUnitId",
        "decisionEvidence",
        "sourcePartition",
    }
    assert schema["$defs"]["sourceRef"]["properties"]["native_kind"]["enum"] == list(
        NATIVE_KINDS
    )


def test_source_partition_preserves_explicit_middle_exclusion() -> None:
    first = _ref(native_id="paragraph:1", text_anchor="材料正文")
    heading = _ref(native_id="paragraph:2", text_anchor="（一）高考专练")
    last = _ref(native_id="paragraph:3", text_anchor="1．题干")

    partition = build_source_partition(
        (first, heading, last),
        included_refs=(first, last),
        excluded_refs=(heading,),
        unknown_refs=(),
    )

    assert partition.included_refs == (first, last)
    assert partition.excluded_refs == (heading,)
    assert partition.unknown_refs == ()
    assert partition.has_unknown_refs is False


def test_source_partition_rejects_overlap_missing_and_extra_refs() -> None:
    first = _ref(native_id="paragraph:1", text_anchor="第一段")
    second = _ref(native_id="paragraph:2", text_anchor="第二段")
    extra = _ref(native_id="paragraph:3", text_anchor="范围外段落")

    with pytest.raises(ValueError, match="重复分类"):
        build_source_partition(
            (first, second),
            included_refs=(first,),
            excluded_refs=(first,),
            unknown_refs=(second,),
        )
    with pytest.raises(ValueError, match="未分类"):
        build_source_partition(
            (first, second),
            included_refs=(first,),
            excluded_refs=(),
            unknown_refs=(),
        )
    with pytest.raises(ValueError, match="范围外"):
        build_source_partition(
            (first, second),
            included_refs=(first, second, extra),
            excluded_refs=(),
            unknown_refs=(),
        )


def test_question_unit_roundtrip_keeps_identity_refs_and_evidence() -> None:
    first = _ref(native_id="paragraph:1", text_anchor="材料正文")
    heading = _ref(native_id="paragraph:2", text_anchor="（一）高考专练")
    last = _ref(native_id="paragraph:3", text_anchor="1．题干")
    partition = build_source_partition(
        (first, heading, last),
        included_refs=(first, last),
        excluded_refs=(heading,),
        unknown_refs=(),
    )
    evidence = DecisionEvidence(
        rule_id="fixture",
        reason="人工构造非连续题块",
        source_refs=(first, heading, last),
        confidence=1.0,
        generator_version="test/1",
    )
    unit = QuestionUnitV2(
        generator_version="test/1",
        stable_unit_id=StableUnitId.create(
            document_sha256=DOCUMENT_SHA256,
            unit_kind="question",
            occurrence=1,
            original_label="1",
        ),
        original_question_label="1",
        subject="文科",
        subject_overlay="chinese",
        question_type="material",
        source_partition=partition,
        decision_evidence=(evidence,),
    )

    restored = QuestionUnitV2.from_dict(unit.to_dict())

    assert restored == unit
    assert restored.canonical_bytes() == unit.canonical_bytes()
    assert unit.canonical_bytes() == unit.canonical_bytes()


def test_question_unit_reports_explicit_unknown_refs() -> None:
    known = _ref(native_id="paragraph:1", text_anchor="已识别题干")
    unknown = _ref("formula", "formula:1", text_anchor="未知公式")
    partition = build_source_partition(
        (known, unknown),
        included_refs=(known,),
        excluded_refs=(),
        unknown_refs=(unknown,),
    )
    evidence = DecisionEvidence("fixture", "存在未知公式", (known, unknown), 0.5, "test/1")
    unit = QuestionUnitV2(
        generator_version="test/1",
        stable_unit_id=StableUnitId.create(
            document_sha256=DOCUMENT_SHA256,
            unit_kind="question",
            occurrence=1,
            original_label="1",
        ),
        original_question_label="1",
        subject="理科",
        subject_overlay=None,
        question_type="subjective",
        source_partition=partition,
        decision_evidence=(evidence,),
    )

    assert unit.has_unknown_refs is True


def test_canonical_answer_unit_does_not_encode_f2_or_f4_mode() -> None:
    paragraph = _ref(native_id="paragraph:8", text_anchor="1．（1）甲（2）乙")
    formula = _ref("formula", "paragraph:8/formula", text_anchor="x=1")
    partition = build_source_partition(
        (paragraph, formula),
        included_refs=(paragraph,),
        excluded_refs=(),
        unknown_refs=(formula,),
    )
    evidence = DecisionEvidence(
        "fixture-answer",
        "只记录答案结构，不决定快捷键",
        (paragraph, formula),
        0.8,
        "test/1",
    )
    answer = CanonicalAnswerUnit(
        generator_version="test/1",
        stable_unit_id=StableUnitId.create(
            document_sha256=DOCUMENT_SHA256,
            unit_kind="answer",
            occurrence=1,
            original_label="1",
        ),
        original_question_label="1",
        answer_items=(
            ContractTextItem("(1)", "甲"),
            ContractTextItem("(2)", "乙"),
        ),
        analysis_items=(ContractTextItem("analysis", "解析内容"),),
        rich_content_refs=(formula,),
        source_partition=partition,
        decision_evidence=(evidence,),
        review_flags=("公式待人工复核",),
    )

    payload = answer.to_dict()
    restored = CanonicalAnswerUnit.from_dict(payload)

    assert restored == answer
    assert "answer_mode" not in payload
    assert "F2" not in answer.canonical_bytes().decode("utf-8")
    assert "F4" not in answer.canonical_bytes().decode("utf-8")


def _golden_contract_bundle() -> dict:
    first = _ref(native_id="paragraph:1", text_anchor="材料正文")
    heading = _ref(native_id="paragraph:2", text_anchor="（一）高考专练")
    last = _ref(native_id="paragraph:3", text_anchor="1．题干")
    formula = _ref("formula", "paragraph:3/formula:1", text_anchor="x=1")
    question_partition = build_source_partition(
        (first, heading, last, formula),
        included_refs=(first, last),
        excluded_refs=(heading,),
        unknown_refs=(formula,),
    )
    question_evidence = DecisionEvidence(
        "golden-question",
        "固定题目契约样本",
        (first, heading, last, formula),
        0.9,
        "golden/1",
    )
    question = QuestionUnitV2(
        generator_version="golden/1",
        stable_unit_id=StableUnitId.create(
            document_sha256=DOCUMENT_SHA256,
            unit_kind="question",
            occurrence=1,
            original_label="1",
        ),
        original_question_label="1",
        subject="文科",
        subject_overlay="chinese",
        question_type="material",
        source_partition=question_partition,
        decision_evidence=(question_evidence,),
    )
    answer_partition = build_source_partition(
        (last, formula),
        included_refs=(last,),
        excluded_refs=(),
        unknown_refs=(formula,),
    )
    answer_evidence = DecisionEvidence(
        "golden-answer",
        "固定答案契约样本",
        (last, formula),
        0.8,
        "golden/1",
    )
    answer = CanonicalAnswerUnit(
        generator_version="golden/1",
        stable_unit_id=StableUnitId.create(
            document_sha256=DOCUMENT_SHA256,
            unit_kind="answer",
            occurrence=1,
            original_label="1",
        ),
        original_question_label="1",
        answer_items=(ContractTextItem("answer", "甲"),),
        analysis_items=(ContractTextItem("analysis", "解析"),),
        rich_content_refs=(formula,),
        source_partition=answer_partition,
        decision_evidence=(answer_evidence,),
        review_flags=("公式待人工复核",),
    )
    return {"answer": answer.to_dict(), "question": question.to_dict()}


def test_contract_bundle_matches_golden_json_bytes() -> None:
    from shared_core.contracts_v2 import canonical_json_bytes

    assert canonical_json_bytes(_golden_contract_bundle()) == GOLDEN_PATH.read_bytes()
