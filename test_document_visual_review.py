from __future__ import annotations

import copy
import json
from pathlib import Path

import pytest

from shared_core.document_visual_review import (
    build_visual_review_template,
    summarize_visual_review,
    validate_visual_review,
    write_visual_review_template,
)


def _manifest() -> dict:
    source_sha = "a" * 64
    return {
        "schema_version": "1.0",
        "generator": "test",
        "classification_mode": "render_facts_only",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "source": {
            "name": "sample.docx",
            "path": "C:/readonly/sample.docx",
            "sha256": source_sha,
            "readonly_input": True,
            "hash_verified_unchanged": True,
        },
        "render": {
            "provider": "wps_com",
            "page_truth_authority": True,
            "pdf_sha256": "b" * 64,
            "dpi": 144,
        },
        "page_count": 1,
        "pages": [
            {
                "page_number": 1,
                "width_pt": 595.0,
                "height_pt": 842.0,
                "width_px": 1190,
                "height_px": 1684,
                "png_path": "pages/page-1.png",
                "png_sha256": "c" * 64,
                "region_count": 2,
                "regions": [
                    {
                        "region_id": "page-0001-text_line-0001-a",
                        "region_type": "text_line",
                        "bbox_pdf": [10.0, 10.0, 100.0, 30.0],
                        "bbox_px": [20, 20, 200, 60],
                        "visible_text": "1. Question",
                        "object_name": None,
                        "source_size": None,
                        "crop_sha256": "d" * 64,
                    },
                    {
                        "region_id": "page-0001-image-0002-b",
                        "region_type": "image",
                        "bbox_pdf": [10.0, 40.0, 200.0, 100.0],
                        "bbox_px": [20, 80, 400, 200],
                        "visible_text": "",
                        "object_name": "Im1",
                        "source_size": [400, 120],
                        "crop_sha256": "e" * 64,
                    },
                ],
            }
        ],
        "issues": [],
    }


def _confirmed_review() -> tuple[dict, dict]:
    manifest = _manifest()
    review = build_visual_review_template(manifest)
    text, image = review["regions"]
    text.update(
        {
            "role": "question_text",
            "decision": "include",
            "belongs_to_question_ids": ["1"],
            "confidence": 1.0,
            "evidence": ["human_confirmed"],
            "review_status": "confirmed",
        }
    )
    image.update(
        {
            "role": "exercise_label",
            "decision": "exclude",
            "action_sequences": [1],
            "confidence": 1.0,
            "evidence": ["wps_visible", "user_confirmed"],
            "review_status": "confirmed",
        }
    )
    return manifest, review


def test_review_template_defaults_every_region_to_unknown_draft() -> None:
    manifest = _manifest()

    review = build_visual_review_template(manifest)

    assert review["schema_version"] == "1.0"
    assert review["source"]["sha256"] == manifest["source"]["sha256"]
    assert review["automatic_rule_binding_enabled"] is False
    assert review["production_execution_enabled"] is False
    assert len(review["regions"]) == 2
    assert all(item["role"] == "unknown" for item in review["regions"])
    assert all(item["decision"] == "review" for item in review["regions"])
    assert all(item["review_status"] == "draft" for item in review["regions"])
    assert summarize_visual_review(manifest, review)["gate_ready"] is False


def test_confirmed_review_reports_decorative_action_contamination() -> None:
    manifest, review = _confirmed_review()

    validated = validate_visual_review(manifest, review)
    summary = summarize_visual_review(manifest, validated)

    assert summary["confirmed_count"] == 2
    assert summary["unknown_count"] == 0
    assert summary["draft_count"] == 0
    assert summary["included_count"] == 1
    assert summary["excluded_count"] == 1
    assert summary["decorative_regions_in_actions"] == 1
    assert summary["gate_ready"] is False
    assert any("装饰" in reason for reason in summary["blocking_reasons"])


def test_macos_development_preview_cannot_pass_production_gate() -> None:
    manifest, review = _confirmed_review()
    manifest["render"].update(
        provider="macos_quicklook_webkit",
        page_truth_authority=False,
    )
    review = build_visual_review_template(manifest)
    for item in review["regions"]:
        item.update(
            role="question_text",
            decision="include",
            belongs_to_question_ids=["1"],
            confidence=1.0,
            evidence=["macos_development_preview"],
            review_status="confirmed",
        )

    summary = summarize_visual_review(manifest, review)

    assert summary["production_page_truth"] is False
    assert summary["gate_ready"] is False
    assert any("Windows WPS" in reason for reason in summary["blocking_reasons"])


@pytest.mark.parametrize(
    ("mutate", "message"),
    [
        (lambda review: review["source"].update(sha256="f" * 64), "source.sha256"),
        (lambda review: review.update(manifest_sha256="f" * 64), "manifest_sha256"),
        (lambda review: review["regions"].pop(), "区域集合"),
        (
            lambda review: review["regions"].append(copy.deepcopy(review["regions"][0])),
            "重复",
        ),
        (lambda review: review["regions"][0].update(role="bad_role"), "role"),
        (
            lambda review: review["regions"][0].update(
                role="unknown",
                review_status="confirmed",
                decision="review",
                evidence=["human_confirmed"],
            ),
            "unknown",
        ),
        (
            lambda review: review["regions"][1].update(
                role="exercise_label",
                decision="exclude",
                belongs_to_question_ids=["1"],
                review_status="confirmed",
                confidence=1.0,
                evidence=["human_confirmed"],
            ),
            "排除区域",
        ),
        (
            lambda review: review["regions"][0].update(
                role="question_text",
                decision="include",
                belongs_to_question_ids=[],
                review_status="confirmed",
                confidence=1.0,
                evidence=["human_confirmed"],
            ),
            "题号",
        ),
    ],
)
def test_invalid_review_contract_stops_the_document(mutate, message: str) -> None:
    manifest = _manifest()
    review = build_visual_review_template(manifest)
    mutate(review)

    with pytest.raises(ValueError, match=message):
        validate_visual_review(manifest, review)


def test_write_visual_review_template_uses_json_as_authority(tmp_path: Path) -> None:
    manifest_path = tmp_path / "PageRenderManifest.json"
    manifest_path.write_text(json.dumps(_manifest(), ensure_ascii=False), encoding="utf-8")

    paths = write_visual_review_template(manifest_path, tmp_path / "review")

    review = json.loads(paths["json"].read_text(encoding="utf-8"))
    markdown = paths["markdown"].read_text(encoding="utf-8")
    assert paths["json"].name == "VisualRoleReview.json"
    assert paths["markdown"].name == "VisualRoleReview.md"
    assert f"- 待确认区域：{len(review['regions'])}" in markdown
    assert "允许生产执行：否" in markdown
