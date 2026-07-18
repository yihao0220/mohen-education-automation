from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from shared_core.document_render import (
    build_page_render_manifest,
    export_docx_to_pdf_with_wps,
    sha256_file,
    write_page_render_artifacts,
)


def _build_pdf(path: Path) -> None:
    image_path = path.with_suffix(".png")
    Image.new("RGB", (80, 40), color=(180, 180, 180)).save(image_path)
    pdf = canvas.Canvas(str(path), pagesize=A4)
    pdf.drawString(72, 770, "Question 1")
    pdf.drawImage(str(image_path), 72, 680, width=80, height=40)
    pdf.showPage()
    pdf.drawString(72, 770, "Question 2")
    pdf.save()
    image_path.unlink()


def _build_source(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"readonly-source-docx")


def test_build_page_render_manifest_records_pages_regions_and_hashes(tmp_path: Path) -> None:
    source = tmp_path / "source" / "sample.docx"
    pdf_path = tmp_path / "wps-render.pdf"
    output_dir = tmp_path / "artifacts"
    _build_source(source)
    _build_pdf(pdf_path)
    before = sha256_file(source)

    manifest = build_page_render_manifest(
        source,
        pdf_path,
        output_dir,
        renderer_metadata={"provider": "preexported_wps_pdf", "wps_version": "test"},
        dpi=144,
    )

    assert manifest["schema_version"] == "1.0"
    assert manifest["source"]["sha256"] == before
    assert manifest["source"]["hash_verified_unchanged"] is True
    assert sha256_file(source) == before
    assert manifest["render"]["rasterizer"] in {"pypdfium2", "pdftoppm"}
    assert manifest["page_count"] == 2
    assert manifest["automatic_rule_binding_enabled"] is False
    assert manifest["production_execution_enabled"] is False
    assert len(manifest["pages"]) == 2
    assert [page["page_number"] for page in manifest["pages"]] == [1, 2]

    first_page = manifest["pages"][0]
    png_path = output_dir / first_page["png_path"]
    assert png_path.is_file()
    assert first_page["png_sha256"] == sha256_file(png_path)
    assert first_page["width_px"] > 0
    assert first_page["height_px"] > 0
    assert first_page["width_pt"] > 0
    assert first_page["height_pt"] > 0
    assert any(region["region_type"] == "text_line" for region in first_page["regions"])
    assert any(region["region_type"] == "image" for region in first_page["regions"])
    assert all(region["crop_sha256"] for region in first_page["regions"])
    for page in manifest["pages"]:
        for region in page["regions"]:
            x0, top, x1, bottom = region["bbox_pdf"]
            assert 0 <= x0 < x1 <= page["width_pt"]
            assert 0 <= top < bottom <= page["height_pt"]


def test_write_page_render_artifacts_uses_json_as_authority(tmp_path: Path) -> None:
    source = tmp_path / "source" / "sample.docx"
    pdf_path = tmp_path / "wps-render.pdf"
    output_dir = tmp_path / "artifacts"
    _build_source(source)
    _build_pdf(pdf_path)

    paths = write_page_render_artifacts(
        source,
        pdf_path,
        output_dir,
        renderer_metadata={"provider": "preexported_wps_pdf", "wps_version": "test"},
    )

    manifest = json.loads(paths["json"].read_text(encoding="utf-8"))
    markdown = paths["markdown"].read_text(encoding="utf-8")
    assert paths["json"].name == "PageRenderManifest.json"
    assert paths["markdown"].name == "PageRenderManifest.md"
    assert f"- 页面数：{manifest['page_count']}" in markdown
    assert "允许生产执行：否" in markdown
    assert not list(output_dir.glob(".*.tmp"))


def test_render_rejects_source_directory_pollution_and_missing_inputs(tmp_path: Path) -> None:
    source = tmp_path / "source" / "sample.docx"
    pdf_path = tmp_path / "wps-render.pdf"
    _build_source(source)
    _build_pdf(pdf_path)

    with pytest.raises(ValueError, match="原题目录"):
        build_page_render_manifest(
            source,
            pdf_path,
            source.parent / "artifacts",
            renderer_metadata={"provider": "test"},
        )
    with pytest.raises(FileNotFoundError):
        build_page_render_manifest(
            source,
            tmp_path / "missing.pdf",
            tmp_path / "artifacts",
            renderer_metadata={"provider": "test"},
        )
    with pytest.raises(FileNotFoundError, match="pdftoppm"):
        build_page_render_manifest(
            source,
            pdf_path,
            tmp_path / "artifacts",
            renderer_metadata={"provider": "test"},
            pdftoppm_path=tmp_path / "missing-pdftoppm",
        )


def test_render_removes_stale_page_pngs(tmp_path: Path) -> None:
    source = tmp_path / "source" / "sample.docx"
    pdf_path = tmp_path / "wps-render.pdf"
    output_dir = tmp_path / "artifacts"
    _build_source(source)
    _build_pdf(pdf_path)
    pages_dir = output_dir / "pages"
    pages_dir.mkdir(parents=True)
    stale = pages_dir / "page-99.png"
    stale.write_bytes(b"stale")

    manifest = build_page_render_manifest(
        source,
        pdf_path,
        output_dir,
        renderer_metadata={"provider": "test"},
    )

    assert manifest["page_count"] == 2
    assert not stale.exists()


class _FakeDocument:
    def __init__(self, output_pdf: Path, *, export_supported: bool = True) -> None:
        self.output_pdf = output_pdf
        self.export_supported = export_supported
        self.closed_with = None
        self.saved = False
        self.export_method = None

    def ExportAsFixedFormat(self, *args, **kwargs) -> None:
        if not self.export_supported:
            raise AttributeError("ExportAsFixedFormat unavailable")
        target = kwargs.get("OutputFileName") or args[0]
        Path(target).write_bytes(b"%PDF-fake")
        self.export_method = "ExportAsFixedFormat"

    def SaveAs2(self, target: str, file_format: int) -> None:
        Path(target).write_bytes(b"%PDF-fake")
        self.export_method = f"SaveAs2:{file_format}"

    def Save(self) -> None:
        self.saved = True

    def Close(self, SaveChanges: int = -1) -> None:
        self.closed_with = SaveChanges


class _FakeDocuments:
    def __init__(self, document: _FakeDocument) -> None:
        self.document = document
        self.open_kwargs = None

    def Open(self, path: str, **kwargs):
        self.open_kwargs = {"path": path, **kwargs}
        return self.document


class _FakeApplication:
    def __init__(self, document: _FakeDocument) -> None:
        self.Documents = _FakeDocuments(document)
        self.DisplayAlerts = True
        self.Version = "12.1-test"
        self.quit_called = False

    def Quit(self) -> None:
        self.quit_called = True


def test_wps_export_is_read_only_and_closes_without_saving(tmp_path: Path) -> None:
    source = tmp_path / "source" / "source.docx"
    output_pdf = tmp_path / "render" / "source.pdf"
    _build_source(source)
    document = _FakeDocument(output_pdf)
    application = _FakeApplication(document)
    progids = []

    def dispatch(progid: str):
        progids.append(progid)
        return application

    metadata = export_docx_to_pdf_with_wps(
        source,
        output_pdf,
        dispatch_factory=dispatch,
    )

    assert progids == ["KWPS.Application"]
    assert application.Documents.open_kwargs == {
        "path": str(source.resolve()),
        "ConfirmConversions": False,
        "ReadOnly": True,
        "AddToRecentFiles": False,
    }
    assert document.saved is False
    assert document.closed_with == 0
    assert application.quit_called is True
    assert output_pdf.is_file()
    assert metadata["provider"] == "wps_com"
    assert metadata["source_hash_verified_unchanged"] is True
    assert metadata["export_method"] == "ExportAsFixedFormat"


def test_wps_export_falls_back_to_saveas2_and_cleans_failed_pdf(tmp_path: Path) -> None:
    source = tmp_path / "source" / "source.docx"
    output_pdf = tmp_path / "render" / "source.pdf"
    _build_source(source)
    document = _FakeDocument(output_pdf, export_supported=False)
    application = _FakeApplication(document)

    metadata = export_docx_to_pdf_with_wps(
        source,
        output_pdf,
        dispatch_factory=lambda _: application,
    )
    assert metadata["export_method"] == "SaveAs2"
    assert document.export_method == "SaveAs2:17"

    output_pdf.unlink()

    class BrokenDocument(_FakeDocument):
        def ExportAsFixedFormat(self, *args, **kwargs) -> None:
            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            output_pdf.write_bytes(b"partial")
            raise RuntimeError("export failed")

        def SaveAs2(self, target: str, file_format: int) -> None:
            raise RuntimeError("fallback failed")

    broken = BrokenDocument(output_pdf)
    broken_app = _FakeApplication(broken)
    with pytest.raises(RuntimeError, match="fallback failed"):
        export_docx_to_pdf_with_wps(
            source,
            output_pdf,
            dispatch_factory=lambda _: broken_app,
        )
    assert not output_pdf.exists()
    assert broken.closed_with == 0
    assert broken_app.quit_called is True


def test_wps_export_never_falls_back_to_microsoft_word(tmp_path: Path) -> None:
    source = tmp_path / "source" / "source.docx"
    output_pdf = tmp_path / "render" / "source.pdf"
    _build_source(source)
    attempted = []

    def unavailable(progid: str):
        attempted.append(progid)
        raise RuntimeError("not registered")

    with pytest.raises(RuntimeError, match="无法创建 WPS Writer COM 对象"):
        export_docx_to_pdf_with_wps(source, output_pdf, dispatch_factory=unavailable)

    assert attempted == ["KWPS.Application", "wps.Application"]
    assert "Word.Application" not in attempted
    assert not output_pdf.exists()


def test_sha256_helper_matches_hashlib(tmp_path: Path) -> None:
    path = tmp_path / "payload.bin"
    path.write_bytes(b"payload")
    assert sha256_file(path) == hashlib.sha256(b"payload").hexdigest()
