from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from shared_core.document_families import (
    analyze_document_profiles,
    cluster_document_ids,
    discover_profile_paths,
    select_representative,
    write_document_family_report,
)


def _profile(
    name: str,
    *,
    paragraphs: int = 100,
    actions: int = 20,
    headings: int = 5,
    tables: int = 0,
    media: int = 0,
    formulas: int = 0,
    role_counts: dict[str, int] | None = None,
    issues: list[dict] | None = None,
    heading_candidates: list[dict] | None = None,
) -> dict:
    counts = role_counts or {
        "document_title": 1,
        "section_heading": 4,
        "internal_heading": 5,
        "question_start": 20,
        "option": 40,
        "body": max(0, paragraphs - 70),
    }
    sha256 = hashlib.sha256(name.encode("utf-8")).hexdigest()
    return {
        "schema_version": "1.1",
        "generator": "test",
        "source": {
            "name": name,
            "path": f"C:/readonly/{name}",
            "sha256": sha256,
            "readonly_input": True,
            "hash_verified_unchanged": True,
        },
        "roles": {
            "schema_version": "1.0",
            "automatic_exclusion_enabled": False,
            "role_counts": counts,
            "paragraphs": [],
            "heading_candidates": heading_candidates or [],
        },
        "fingerprint": {
            "paragraph_count": paragraphs,
            "nonempty_paragraph_count": paragraphs,
            "table_count": tables,
            "media_node_count": media,
            "formula_count": formulas,
            "heading_count": headings,
            "question_action_count": actions,
            "role_counts": counts,
        },
        "issues": issues or [],
    }


def _stable_json(report: dict) -> str:
    return json.dumps(report, ensure_ascii=False, sort_keys=True)


def test_equivalent_structure_with_different_scale_is_grouped() -> None:
    first = _profile("第一章.docx", paragraphs=100, actions=20, headings=5)
    second = _profile(
        "第二章.docx",
        paragraphs=200,
        actions=40,
        headings=10,
        role_counts={
            "document_title": 2,
            "section_heading": 8,
            "internal_heading": 10,
            "question_start": 40,
            "option": 80,
            "body": 60,
        },
    )

    report = analyze_document_profiles([first, second])

    assert len(report["families"]) == 1
    assert report["families"][0]["member_count"] == 2
    pair = report["pairwise_similarities"][0]
    assert pair["similarity"] >= report["similarity_definition"]["family_threshold"]
    assert set(pair["components"]) == {
        "role_distribution",
        "action_density",
        "heading_density",
        "object_presence",
        "rich_object_density",
        "document_scale",
    }


def test_role_distribution_prevents_same_size_documents_from_false_merge() -> None:
    first = _profile(
        "题库.docx",
        role_counts={"question_start": 50, "option": 50},
        actions=50,
        headings=0,
    )
    second = _profile(
        "讲义.docx",
        role_counts={"body": 100},
        actions=0,
        headings=0,
    )

    report = analyze_document_profiles([first, second])

    assert len(report["families"]) == 2
    assert report["pairwise_similarities"][0]["similarity"] < 0.78
    assert len(report["unresolved_singletons"]) == 2
    assert report["outlier_candidates"] == []


def test_complete_link_prevents_chain_merge() -> None:
    similarities = {
        ("A", "B"): 0.90,
        ("A", "C"): 0.70,
        ("B", "C"): 0.90,
    }

    clusters = cluster_document_ids(["C", "B", "A"], similarities, threshold=0.78)

    assert clusters == [["A", "B"], ["C"]]


def test_representative_is_medoid_with_stable_tie_break() -> None:
    similarities = {
        ("A", "B"): 0.80,
        ("A", "C"): 0.70,
        ("B", "C"): 0.90,
    }

    representative, averages = select_representative(["C", "A", "B"], similarities)

    assert representative == "B"
    assert averages == {"A": 0.75, "B": 0.85, "C": 0.8}


def test_singleton_requires_three_documents_and_low_nearest_similarity_for_outlier() -> None:
    first = _profile("题库一.docx")
    second = _profile("讲义.docx", role_counts={"body": 100}, actions=0, headings=0)

    two_document_report = analyze_document_profiles([first, second])
    assert two_document_report["outlier_candidates"] == []

    third = _profile("题库二.docx")
    three_document_report = analyze_document_profiles([first, second, third])

    assert [item["source_name"] for item in three_document_report["outlier_candidates"]] == [
        "讲义.docx"
    ]
    assert three_document_report["outlier_candidates"][0]["nearest_similarity"] < 0.60


def test_report_is_deterministic_when_input_order_changes() -> None:
    profiles = [
        _profile("第一章.docx"),
        _profile("第二章.docx", paragraphs=200, actions=40, headings=10),
        _profile("讲义.docx", role_counts={"body": 100}, actions=0, headings=0),
    ]

    forward = analyze_document_profiles(profiles)
    reversed_report = analyze_document_profiles(list(reversed(profiles)))

    assert _stable_json(forward) == _stable_json(reversed_report)
    assert forward["classification_mode"] == "advisory_only"
    assert forward["automatic_rule_binding_enabled"] is False
    assert forward["production_execution_enabled"] is False


def test_invalid_profile_stops_the_whole_batch() -> None:
    invalid_schema = _profile("旧画像.docx")
    invalid_schema["schema_version"] = "1.0"
    with pytest.raises(ValueError, match="schema_version"):
        analyze_document_profiles([_profile("正常.docx"), invalid_schema])

    writable = _profile("可写画像.docx")
    writable["source"]["readonly_input"] = False
    with pytest.raises(ValueError, match="readonly_input"):
        analyze_document_profiles([writable])

    exclusion_enabled = _profile("自动排除.docx")
    exclusion_enabled["roles"]["automatic_exclusion_enabled"] = True
    with pytest.raises(ValueError, match="automatic_exclusion_enabled"):
        analyze_document_profiles([exclusion_enabled])


def test_same_source_sha_with_conflicting_profiles_stops_batch() -> None:
    first = _profile("同源一.docx")
    second = json.loads(json.dumps(first, ensure_ascii=False))
    second["source"]["name"] = "同源二.docx"
    second["source"]["path"] = "D:/other/同源二.docx"
    second["fingerprint"]["question_action_count"] = 99

    with pytest.raises(ValueError, match="SHA256"):
        analyze_document_profiles([first, second])


def test_duplicate_profile_is_deduplicated_and_profile_risks_enter_review_queue() -> None:
    risky = _profile(
        "风险画像.docx",
        role_counts={
            "document_title": 1,
            "question_start": 20,
            "body": 78,
            "future_role": 1,
        },
        issues=[{"severity": "warning", "code": "TEST_WARNING"}],
        heading_candidates=[
            {"paragraph_index": 8, "confidence": 0.70, "text": "专项提升"}
        ],
    )
    duplicate = json.loads(json.dumps(risky, ensure_ascii=False))
    duplicate["source"]["path"] = "D:/same-source/风险画像.docx"

    report = analyze_document_profiles([duplicate, risky])

    assert report["input_profile_count"] == 1
    assert any("重复 Profile 已去重" in warning for warning in report["warnings"])
    assert any("future_role" in warning for warning in report["warnings"])
    risk_items = [
        item for item in report["review_queue"] if item["review_type"] == "profile_risk"
    ]
    assert len(risk_items) == 1
    assert "TEST_WARNING" in risk_items[0]["reason"]
    assert "置信度 0.7" in risk_items[0]["reason"]
    assert "future_role" in risk_items[0]["reason"]


def test_empty_profile_stops_batch() -> None:
    empty = _profile("空画像.docx")
    empty["fingerprint"]["nonempty_paragraph_count"] = 0

    with pytest.raises(ValueError, match="nonempty_paragraph_count"):
        analyze_document_profiles([empty])


def test_write_report_uses_json_as_authority_and_discovers_only_profiles(
    tmp_path: Path,
) -> None:
    profile_dir = tmp_path / "profiles"
    profile_dir.mkdir()
    first_path = profile_dir / "第一章_DocumentProfile.json"
    second_path = profile_dir / "第二章_DocumentProfile.json"
    ignored_path = profile_dir / "第一章_ActionPlan.json"
    first_path.write_text(json.dumps(_profile("第一章.docx"), ensure_ascii=False), encoding="utf-8")
    second_path.write_text(json.dumps(_profile("第二章.docx"), ensure_ascii=False), encoding="utf-8")
    ignored_path.write_text("{}", encoding="utf-8")

    discovered = discover_profile_paths([profile_dir, first_path])
    assert discovered == [first_path.resolve(), second_path.resolve()]

    output_dir = tmp_path / "report"
    paths = write_document_family_report([profile_dir], output_dir)
    report = json.loads(paths["json"].read_text(encoding="utf-8"))
    markdown = paths["markdown"].read_text(encoding="utf-8")

    assert paths["json"].name == "DocumentFamilyReport.json"
    assert paths["markdown"].name == "DocumentFamilyReport.md"
    assert report["input_profile_count"] == 2
    assert f"候选文档族：{len(report['families'])}" in markdown
    assert "自动绑定文档族规则：否" in markdown
    assert "允许生产执行：否" in markdown


def test_existing_real_profiles_are_compatible_and_unchanged(tmp_path: Path) -> None:
    baseline_dir = Path("回归样本/预检基线")
    if not baseline_dir.is_dir():
        pytest.skip("真实预检基线只保留在本机，不进入 Git 仓库")
    profile_paths = discover_profile_paths([baseline_dir])
    assert len(profile_paths) == 3
    before = {path: hashlib.sha256(path.read_bytes()).hexdigest() for path in profile_paths}

    paths = write_document_family_report([baseline_dir], tmp_path / "report")
    report = json.loads(paths["json"].read_text(encoding="utf-8"))

    assert report["input_profile_count"] == 3
    assert {path: hashlib.sha256(path.read_bytes()).hexdigest() for path in profile_paths} == before
    assert report["production_execution_enabled"] is False

    action_plans = sorted(baseline_dir.rglob("*_ActionPlan.json"))
    plans = [json.loads(path.read_text(encoding="utf-8")) for path in action_plans]
    assert sorted(len(plan["actions"]) for plan in plans) == [1, 30, 119]
    assert all(plan["execution_enabled"] is False for plan in plans)
