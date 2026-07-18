from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


PROJECT_ROOT = Path(__file__).resolve().parent


def _run(*args: object) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "ascii:backslashreplace"
    return subprocess.run(
        [sys.executable, *(str(arg) for arg in args)],
        cwd=PROJECT_ROOT,
        env=env,
        text=True,
        encoding="utf-8",
        capture_output=True,
        timeout=60,
    )


def _pdf(path: Path) -> None:
    document = canvas.Canvas(str(path), pagesize=A4)
    document.drawString(72, 770, "Question 1")
    document.save()


def test_render_and_visual_review_cli_offline(tmp_path: Path) -> None:
    source = tmp_path / "source" / "sample.docx"
    source.parent.mkdir()
    source.write_bytes(b"readonly-docx")
    pdf = tmp_path / "wps.pdf"
    _pdf(pdf)
    render_dir = tmp_path / "render"

    rendered = _run(
        "tools/build_document_render.py",
        source,
        "--pdf-input",
        pdf,
        "--output-dir",
        render_dir,
    )
    assert rendered.returncode == 0, rendered.stderr
    result = json.loads(rendered.stdout)
    assert result["page_count"] == 1
    assert result["page_truth_authority"] is False
    assert result["production_execution_enabled"] is False
    manifest = render_dir / "PageRenderManifest.json"
    assert manifest.is_file()
    manifest_data = json.loads(manifest.read_text(encoding="utf-8"))
    assert manifest_data["render"]["provider"] == "external_pdf_preview"

    review_dir = tmp_path / "review"
    reviewed = _run(
        "tools/build_visual_review.py",
        manifest,
        "--output-dir",
        review_dir,
    )
    assert reviewed.returncode == 0, reviewed.stderr
    review_result = json.loads(reviewed.stdout)
    assert review_result["gate_ready"] is False
    assert (review_dir / "VisualRoleReview.json").is_file()
    assert (review_dir / "VisualReviewSummary.json").is_file()


def test_render_cli_requires_explicit_attestation_for_preexported_wps_pdf(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source" / "sample.docx"
    source.parent.mkdir()
    source.write_bytes(b"readonly-docx")
    pdf = tmp_path / "wps.pdf"
    _pdf(pdf)
    output_dir = tmp_path / "render"

    rendered = _run(
        "tools/build_document_render.py",
        source,
        "--pdf-input",
        pdf,
        "--attest-wps-export",
        "--output-dir",
        output_dir,
    )

    assert rendered.returncode == 0, rendered.stderr
    result = json.loads(rendered.stdout)
    manifest = json.loads(
        (output_dir / "PageRenderManifest.json").read_text(encoding="utf-8")
    )
    assert result["page_truth_authority"] is True
    assert manifest["render"]["provider"] == "user_attested_wps_pdf"


def test_render_cli_rejects_attestation_without_pdf_input(tmp_path: Path) -> None:
    source = tmp_path / "source" / "sample.docx"
    source.parent.mkdir()
    source.write_bytes(b"readonly-docx")

    result = _run(
        "tools/build_document_render.py",
        source,
        "--attest-wps-export",
        "--output-dir",
        tmp_path / "render",
    )

    assert result.returncode != 0
    assert "只能与 --pdf-input 同时使用" in result.stderr


def test_render_cli_rejects_source_pollution_and_wps_lock_files(tmp_path: Path) -> None:
    source = tmp_path / "source" / "sample.docx"
    source.parent.mkdir()
    source.write_bytes(b"readonly-docx")
    pdf = tmp_path / "wps.pdf"
    _pdf(pdf)
    polluted = _run(
        "tools/build_document_render.py",
        source,
        "--pdf-input",
        pdf,
        "--output-dir",
        source.parent / "artifacts",
    )
    assert polluted.returncode != 0
    assert "原题目录" in polluted.stderr

    lock = source.parent / "~$sample.docx"
    lock.write_bytes(b"lock")
    rejected = _run(
        "tools/build_document_render.py",
        lock,
        "--pdf-input",
        pdf,
        "--output-dir",
        tmp_path / "lock-output",
    )
    assert rejected.returncode != 0
    assert "WPS 临时文件" in rejected.stderr


@pytest.mark.skipif(sys.platform == "win32", reason="Windows 上不得启动真实 WPS COM")
def test_render_cli_explicit_wps_mode_reports_windows_blocker(tmp_path: Path) -> None:
    source = tmp_path / "source" / "sample.docx"
    source.parent.mkdir()
    source.write_bytes(b"readonly-docx")
    result = _run(
        "tools/build_document_render.py",
        source,
        "--renderer",
        "wps",
        "--output-dir",
        tmp_path / "render",
    )
    assert result.returncode != 0
    assert "WPS COM 导出只支持 Windows" in result.stderr


def test_calibration_cli_reads_review_directory(tmp_path: Path) -> None:
    document_id = "a" * 64
    family_report = {
        "schema_version": "1.0",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "input_profile_count": 1,
        "input_profiles": [{"document_id": document_id, "source_name": "sample.docx"}],
        "pairwise_similarities": [],
        "similarity_definition": {"family_threshold": 0.78},
    }
    ground_truth = {
        "schema_version": "1.0",
        "classification_basis": "execution_rule_equivalence",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "documents": [
            {
                "source_sha256": document_id,
                "source_name": "sample.docx",
                "rule_family_id": "biology-live-sheet",
                "review_status": "confirmed",
            }
        ],
    }
    summary = {
        "source_sha256": document_id,
        "source_name": "sample.docx",
        "source_hash_verified_unchanged": True,
        "production_page_truth": True,
        "decorative_regions_in_actions": 0,
        "question_media_without_binding": 0,
        "blocking_reasons": [],
        "gate_ready": True,
    }
    family_path = tmp_path / "DocumentFamilyReport.json"
    truth_path = tmp_path / "GroundTruth.json"
    reviews_dir = tmp_path / "reviews" / "sample"
    reviews_dir.mkdir(parents=True)
    family_path.write_text(json.dumps(family_report), encoding="utf-8")
    truth_path.write_text(json.dumps(ground_truth), encoding="utf-8")
    (reviews_dir / "VisualReviewSummary.json").write_text(json.dumps(summary), encoding="utf-8")

    result = _run(
        "tools/calibrate_document_families.py",
        family_path,
        truth_path,
        "--reviews",
        reviews_dir.parent,
        "--output-dir",
        tmp_path / "calibration",
    )
    assert result.returncode == 0, result.stderr
    payload = json.loads(result.stdout)
    assert payload["gate_passed"] is True
    assert payload["calibration_scope"] == "batch_scoped"
    assert payload["production_execution_enabled"] is False
