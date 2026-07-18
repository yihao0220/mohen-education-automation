from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path

import pytest

from shared_core.document_family_calibration import (
    build_p1b_calibration_report,
    write_p1b_calibration_report,
)


def _sha(name: str) -> str:
    return hashlib.sha256(name.encode("utf-8")).hexdigest()


def _family_report(names: list[str], scores: dict[tuple[int, int], float]) -> dict:
    profiles = [
        {
            "document_id": _sha(name),
            "source_name": name,
            "source_sha256": _sha(name),
            "profile_path": f"C:/profiles/{name}_DocumentProfile.json",
        }
        for name in names
    ]
    pairs = []
    for (left, right), score in scores.items():
        pairs.append(
            {
                "first_document_id": profiles[left]["document_id"],
                "first_source_name": names[left],
                "second_document_id": profiles[right]["document_id"],
                "second_source_name": names[right],
                "similarity": score,
                "components": {},
            }
        )
    return {
        "schema_version": "1.0",
        "generator": "test",
        "classification_mode": "advisory_only",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "input_profile_count": len(profiles),
        "input_profiles": profiles,
        "similarity_definition": {"family_threshold": 0.78},
        "pairwise_similarities": pairs,
    }


def _ground_truth(names: list[str], families: list[str]) -> dict:
    return {
        "schema_version": "1.0",
        "classification_basis": "execution_rule_equivalence",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "documents": [
            {
                "source_name": name,
                "source_sha256": _sha(name),
                "rule_family_id": family,
                "document_shape": "worksheet",
                "review_status": "confirmed",
            }
            for name, family in zip(names, families)
        ],
    }


def _summaries(names: list[str]) -> list[dict]:
    return [
        {
            "source_name": name,
            "source_sha256": _sha(name),
            "source_hash_verified_unchanged": True,
            "production_page_truth": True,
            "region_count": 10,
            "confirmed_count": 10,
            "draft_count": 0,
            "rejected_count": 0,
            "unknown_count": 0,
            "pending_count": 0,
            "included_count": 8,
            "excluded_count": 2,
            "question_media_binding_count": 2,
            "question_media_without_binding": 0,
            "decorative_regions_in_actions": 0,
            "blocking_reasons": [],
            "gate_ready": True,
        }
        for name in names
    ]


def test_one_family_calibration_is_batch_scoped_and_selects_highest_zero_error() -> None:
    names = ["A.docx", "B.docx"]
    family_report = _family_report(names, {(0, 1): 0.9})

    report = build_p1b_calibration_report(
        family_report,
        _ground_truth(names, ["biology_v1", "biology_v1"]),
        _summaries(names),
    )

    assert report["classification_mode"] == "human_calibrated_advisory"
    assert report["automatic_rule_binding_enabled"] is False
    assert report["production_execution_enabled"] is False
    assert report["calibration_scope"] == "batch_scoped"
    assert report["recommended_threshold"] == 0.9
    assert report["gate"]["passed"] is True
    assert any("负例" in warning for warning in report["warnings"])
    by_threshold = {item["threshold"]: item for item in report["threshold_evaluations"]}
    assert by_threshold[1.0]["false_split_count"] == 1
    assert by_threshold[0.9]["total_error_count"] == 0


def test_multiple_ground_truth_families_report_false_merge_and_false_split() -> None:
    names = ["A.docx", "B.docx", "C.docx"]
    family_report = _family_report(
        names,
        {(0, 1): 0.9, (0, 2): 0.4, (1, 2): 0.4},
    )

    report = build_p1b_calibration_report(
        family_report,
        _ground_truth(names, ["family_1", "family_1", "family_2"]),
        _summaries(names),
    )

    assert report["calibration_scope"] == "multi_family_calibrated"
    assert report["recommended_threshold"] == 0.9
    by_threshold = {item["threshold"]: item for item in report["threshold_evaluations"]}
    assert by_threshold[1.0]["false_split_count"] == 1
    assert by_threshold[0.4]["false_merge_count"] == 2
    assert by_threshold[0.9]["total_error_count"] == 0


def test_batch_gate_reports_missing_and_visual_blockers() -> None:
    names = ["A.docx", "B.docx"]
    family_report = _family_report(names, {(0, 1): 0.9})
    ground_truth = _ground_truth(names, ["biology_v1", "biology_v1"])
    summaries = _summaries(names)
    summaries.pop()

    missing = build_p1b_calibration_report(family_report, ground_truth, summaries)
    assert missing["gate"]["passed"] is False
    assert any("视觉审核" in reason and "缺少" in reason for reason in missing["gate"]["blocking_reasons"])

    summaries = _summaries(names)
    summaries[0]["decorative_regions_in_actions"] = 1
    summaries[0]["blocking_reasons"] = ["装饰区域进入 F1"]
    summaries[0]["gate_ready"] = False
    contaminated = build_p1b_calibration_report(family_report, ground_truth, summaries)
    assert contaminated["gate"]["passed"] is False
    assert any("装饰" in reason for reason in contaminated["gate"]["blocking_reasons"])

    summaries = _summaries(names)
    summaries[0]["source_hash_verified_unchanged"] = False
    summaries[0]["gate_ready"] = False
    summaries[0]["blocking_reasons"] = ["哈希未确认"]
    hash_failed = build_p1b_calibration_report(family_report, ground_truth, summaries)
    assert hash_failed["gate"]["passed"] is False
    assert any("SHA256" in reason for reason in hash_failed["gate"]["blocking_reasons"])

    summaries = _summaries(names)
    summaries[0]["production_page_truth"] = False
    summaries[0]["gate_ready"] = False
    summaries[0]["blocking_reasons"] = ["Mac 开发预览"]
    preview_only = build_p1b_calibration_report(family_report, ground_truth, summaries)
    assert preview_only["gate"]["passed"] is False
    assert any("Windows WPS" in reason for reason in preview_only["gate"]["blocking_reasons"])


def test_unconfirmed_ground_truth_blocks_and_duplicate_sha_stops() -> None:
    names = ["A.docx", "B.docx"]
    family_report = _family_report(names, {(0, 1): 0.9})
    ground_truth = _ground_truth(names, ["biology_v1", "biology_v1"])
    ground_truth["documents"][0]["review_status"] = "draft"

    report = build_p1b_calibration_report(family_report, ground_truth, _summaries(names))
    assert report["gate"]["passed"] is False
    assert any("人工真值" in reason for reason in report["gate"]["blocking_reasons"])

    duplicate = _summaries(names)
    duplicate.append(copy.deepcopy(duplicate[0]))
    with pytest.raises(ValueError, match="重复"):
        build_p1b_calibration_report(
            family_report,
            _ground_truth(names, ["biology_v1", "biology_v1"]),
            duplicate,
        )


def test_enabled_production_switch_stops_calibration() -> None:
    names = ["A.docx", "B.docx"]
    family_report = _family_report(names, {(0, 1): 0.9})
    family_report["production_execution_enabled"] = True

    with pytest.raises(ValueError, match="production_execution_enabled"):
        build_p1b_calibration_report(
            family_report,
            _ground_truth(names, ["biology_v1", "biology_v1"]),
            _summaries(names),
        )


def test_write_p1b_report_uses_json_as_authority(tmp_path: Path) -> None:
    names = ["A.docx", "B.docx"]
    family_path = tmp_path / "DocumentFamilyReport.json"
    truth_path = tmp_path / "GroundTruth.json"
    summaries_path = tmp_path / "VisualReviewSummaries.json"
    family_path.write_text(
        json.dumps(_family_report(names, {(0, 1): 0.9}), ensure_ascii=False),
        encoding="utf-8",
    )
    truth_path.write_text(
        json.dumps(_ground_truth(names, ["biology_v1", "biology_v1"]), ensure_ascii=False),
        encoding="utf-8",
    )
    summaries_path.write_text(json.dumps(_summaries(names), ensure_ascii=False), encoding="utf-8")

    paths = write_p1b_calibration_report(
        family_path,
        truth_path,
        summaries_path,
        tmp_path / "output",
    )

    report = json.loads(paths["json"].read_text(encoding="utf-8"))
    markdown = paths["markdown"].read_text(encoding="utf-8")
    assert paths["json"].name == "P1bCalibrationReport.json"
    assert report["gate"]["passed"] is True
    assert "允许生产执行：否" in markdown
    assert f"建议阈值：{report['recommended_threshold']:.6f}" in markdown
