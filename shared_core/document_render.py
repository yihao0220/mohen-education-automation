from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
from io import BytesIO
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence


MANIFEST_SCHEMA_VERSION = "1.0"
GENERATOR_NAME = "mohen-wps-page-render-p1b"
DEFAULT_DPI = 144
PDF_FORMAT = 17
WPS_PROGIDS = ("KWPS.Application", "wps.Application")


def sha256_file(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _is_inside(path: Path, directory: Path) -> bool:
    try:
        path.relative_to(directory)
    except ValueError:
        return False
    return True


def _validate_source_and_output(source_path: Path, output_dir: Path) -> None:
    source = source_path.resolve()
    target = output_dir.resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if _is_inside(target, source.parent):
        raise ValueError("页面渲染产物不能写入原题目录或其子目录")


def discover_pdftoppm(explicit_path: str | Path | None = None) -> Path:
    if explicit_path is not None:
        path = Path(explicit_path).resolve()
        if not path.is_file():
            raise FileNotFoundError(f"pdftoppm 不存在：{path}")
        return path
    located = shutil.which("pdftoppm")
    if not located:
        raise FileNotFoundError("没有找到 pdftoppm，请安装 Poppler 或显式传入路径")
    return Path(located).resolve()


def discover_pdf_rasterizer(
    explicit_pdftoppm_path: str | Path | None = None,
) -> tuple[str, Path | None]:
    if explicit_pdftoppm_path is not None:
        return "pdftoppm", discover_pdftoppm(explicit_pdftoppm_path)
    try:
        import pypdfium2  # noqa: F401
    except ImportError:
        return "pdftoppm", discover_pdftoppm()
    return "pypdfium2", None


def _page_number(path: Path) -> int:
    suffix = path.stem.rsplit("-", 1)[-1]
    if not suffix.isdigit():
        raise ValueError(f"无法从页面文件名识别页码：{path.name}")
    return int(suffix)


def render_pdf_pages(
    pdf_path: str | Path,
    pages_dir: str | Path,
    *,
    dpi: int = DEFAULT_DPI,
    pdftoppm_path: str | Path | None = None,
) -> list[Path]:
    source_pdf = Path(pdf_path).resolve()
    if not source_pdf.is_file():
        raise FileNotFoundError(source_pdf)
    if not isinstance(dpi, int) or dpi <= 0:
        raise ValueError("dpi 必须是正整数")
    rasterizer, executable = discover_pdf_rasterizer(pdftoppm_path)
    target_dir = Path(pages_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    for stale in target_dir.glob("page-*.png"):
        stale.unlink()
    if rasterizer == "pypdfium2":
        try:
            import pypdfium2 as pdfium

            document = pdfium.PdfDocument(str(source_pdf))
            try:
                if len(document) <= 0:
                    raise ValueError(f"PDF 页数为 0：{source_pdf}")
                scale = dpi / 72
                for page_index in range(len(document)):
                    page = document[page_index]
                    try:
                        bitmap = page.render(scale=scale)
                        try:
                            image = bitmap.to_pil()
                            image.save(target_dir / f"page-{page_index + 1}.png", format="PNG")
                        finally:
                            bitmap.close()
                    finally:
                        page.close()
            finally:
                document.close()
        except (OSError, RuntimeError, ValueError) as exc:
            raise RuntimeError(f"PDF 页面渲染失败：{source_pdf.name}：{exc}") from exc
    else:
        prefix = target_dir / "page"
        command = [str(executable), "-png", "-r", str(dpi), str(source_pdf), str(prefix)]
        try:
            subprocess.run(
                command,
                check=True,
                capture_output=True,
                text=True,
                timeout=600,
            )
        except (OSError, subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
            raise RuntimeError(f"PDF 页面渲染失败：{source_pdf.name}：{exc}") from exc
    pages = sorted(target_dir.glob("page-*.png"), key=_page_number)
    if not pages:
        raise ValueError(f"PDF 没有生成任何页面：{source_pdf}")
    expected = list(range(1, len(pages) + 1))
    actual = [_page_number(path) for path in pages]
    if actual != expected:
        raise ValueError(f"页面编号不连续：期望 {expected}，实际 {actual}")
    return pages


def _round_coord(value: Any) -> float:
    return round(float(value), 3)


def _normalize_bbox(
    bbox: Sequence[Any],
    *,
    width: float,
    height: float,
    context: str,
) -> list[float]:
    if len(bbox) != 4:
        raise ValueError(f"{context} 的边界框必须包含 4 个坐标")
    x0, top, x1, bottom = (_round_coord(value) for value in bbox)
    epsilon = 0.01
    if x0 < -epsilon or top < -epsilon or x1 > width + epsilon or bottom > height + epsilon:
        raise ValueError(f"{context} 的边界框超出页面")
    x0 = max(0.0, x0)
    top = max(0.0, top)
    x1 = min(_round_coord(width), x1)
    bottom = min(_round_coord(height), bottom)
    if not (x0 < x1 and top < bottom):
        raise ValueError(f"{context} 的边界框无效")
    return [x0, top, x1, bottom]


def _pixel_bbox(
    bbox_pdf: Sequence[float],
    *,
    width_pt: float,
    height_pt: float,
    width_px: int,
    height_px: int,
) -> list[int]:
    scale_x = width_px / width_pt
    scale_y = height_px / height_pt
    x0, top, x1, bottom = bbox_pdf
    result = [
        max(0, min(width_px - 1, round(x0 * scale_x))),
        max(0, min(height_px - 1, round(top * scale_y))),
        max(1, min(width_px, round(x1 * scale_x))),
        max(1, min(height_px, round(bottom * scale_y))),
    ]
    if not (result[0] < result[2] and result[1] < result[3]):
        raise ValueError("页面区域转换为像素坐标后无效")
    return result


def _crop_sha256(image: Any, bbox_px: Sequence[int]) -> str:
    cropped = image.crop(tuple(bbox_px))
    if cropped.width <= 0 or cropped.height <= 0:
        raise ValueError("页面区域裁剪结果为空")
    buffer = BytesIO()
    cropped.save(buffer, format="PNG")
    return hashlib.sha256(buffer.getvalue()).hexdigest()


def _region_id(
    page_number: int,
    region_type: str,
    sequence: int,
    bbox_pdf: Sequence[float],
    visible_text: str,
) -> str:
    evidence = json.dumps(
        [page_number, region_type, sequence, list(bbox_pdf), visible_text],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    digest = hashlib.sha256(evidence.encode("utf-8")).hexdigest()[:10]
    return f"page-{page_number:04d}-{region_type}-{sequence:04d}-{digest}"


def _extract_text_lines(page: Any) -> list[dict[str, Any]]:
    try:
        lines = page.extract_text_lines(return_chars=False) or []
    except (AttributeError, TypeError):
        lines = page.extract_words(use_text_flow=True, keep_blank_chars=False) or []
    normalized = []
    for line in lines:
        text = str(line.get("text") or "").strip()
        if not text:
            continue
        normalized.append(
            {
                "text": text,
                "bbox": [line.get("x0"), line.get("top"), line.get("x1"), line.get("bottom")],
            }
        )
    return sorted(normalized, key=lambda item: (float(item["bbox"][1]), float(item["bbox"][0]), item["text"]))


def extract_pdf_layout(
    pdf_path: str | Path,
    rendered_pages: Sequence[str | Path],
    *,
    dpi: int = DEFAULT_DPI,
) -> list[dict[str, Any]]:
    try:
        import pdfplumber
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("缺少 pdfplumber 或 Pillow，无法生成页面布局画像") from exc

    source_pdf = Path(pdf_path).resolve()
    page_paths = [Path(path).resolve() for path in rendered_pages]
    with pdfplumber.open(source_pdf) as pdf:
        if not pdf.pages:
            raise ValueError(f"PDF 页数为 0：{source_pdf}")
        if len(pdf.pages) != len(page_paths):
            raise ValueError(
                f"PDF 页数与 PNG 数量不一致：PDF={len(pdf.pages)}，PNG={len(page_paths)}"
            )
        pages: list[dict[str, Any]] = []
        for page_number, (page, png_path) in enumerate(zip(pdf.pages, page_paths), start=1):
            if not png_path.is_file():
                raise FileNotFoundError(png_path)
            with Image.open(png_path) as image:
                image.load()
                width_px, height_px = image.size
                if width_px <= 0 or height_px <= 0:
                    raise ValueError(f"第 {page_number} 页 PNG 尺寸无效")
                width_pt = _round_coord(page.width)
                height_pt = _round_coord(page.height)
                if width_pt <= 0 or height_pt <= 0:
                    raise ValueError(f"第 {page_number} 页 PDF 尺寸无效")
                raw_regions: list[dict[str, Any]] = []
                for line in _extract_text_lines(page):
                    raw_regions.append(
                        {
                            "region_type": "text_line",
                            "visible_text": line["text"],
                            "bbox": line["bbox"],
                            "object_name": None,
                            "source_size": None,
                        }
                    )
                for item in page.images or []:
                    raw_regions.append(
                        {
                            "region_type": "image",
                            "visible_text": "",
                            "bbox": [item.get("x0"), item.get("top"), item.get("x1"), item.get("bottom")],
                            "object_name": str(item.get("name") or ""),
                            "source_size": list(item.get("srcsize")) if item.get("srcsize") else None,
                        }
                    )
                raw_regions.sort(
                    key=lambda item: (
                        float(item["bbox"][1]),
                        float(item["bbox"][0]),
                        item["region_type"],
                        item["visible_text"],
                    )
                )
                regions = []
                for sequence, item in enumerate(raw_regions, start=1):
                    bbox_pdf = _normalize_bbox(
                        item["bbox"],
                        width=width_pt,
                        height=height_pt,
                        context=f"第 {page_number} 页第 {sequence} 个区域",
                    )
                    bbox_px = _pixel_bbox(
                        bbox_pdf,
                        width_pt=width_pt,
                        height_pt=height_pt,
                        width_px=width_px,
                        height_px=height_px,
                    )
                    regions.append(
                        {
                            "region_id": _region_id(
                                page_number,
                                item["region_type"],
                                sequence,
                                bbox_pdf,
                                item["visible_text"],
                            ),
                            "region_type": item["region_type"],
                            "bbox_pdf": bbox_pdf,
                            "bbox_px": bbox_px,
                            "visible_text": item["visible_text"],
                            "object_name": item["object_name"],
                            "source_size": item["source_size"],
                            "crop_sha256": _crop_sha256(image, bbox_px),
                        }
                    )
            pages.append(
                {
                    "page_number": page_number,
                    "width_pt": width_pt,
                    "height_pt": height_pt,
                    "width_px": width_px,
                    "height_px": height_px,
                    "png_path": f"pages/{png_path.name}",
                    "png_sha256": sha256_file(png_path),
                    "region_count": len(regions),
                    "regions": regions,
                }
            )
    return pages


def build_page_render_manifest(
    source_path: str | Path,
    pdf_path: str | Path,
    output_dir: str | Path,
    *,
    renderer_metadata: Mapping[str, Any],
    dpi: int = DEFAULT_DPI,
    pdftoppm_path: str | Path | None = None,
) -> dict[str, Any]:
    source = Path(source_path).resolve()
    source_pdf = Path(pdf_path).resolve()
    target_dir = Path(output_dir).resolve()
    _validate_source_and_output(source, target_dir)
    if not source_pdf.is_file():
        raise FileNotFoundError(source_pdf)
    if not isinstance(renderer_metadata, Mapping) or not renderer_metadata.get("provider"):
        raise ValueError("renderer_metadata.provider 不能为空")
    before_hash = sha256_file(source)
    pages_dir = target_dir / "pages"
    rasterizer, _ = discover_pdf_rasterizer(pdftoppm_path)
    rendered_pages = render_pdf_pages(
        source_pdf,
        pages_dir,
        dpi=dpi,
        pdftoppm_path=pdftoppm_path,
    )
    pages = extract_pdf_layout(source_pdf, rendered_pages, dpi=dpi)
    after_hash = sha256_file(source)
    if before_hash != after_hash:
        raise RuntimeError("原题 SHA256 在页面渲染过程中发生变化，已停止")
    return {
        "schema_version": MANIFEST_SCHEMA_VERSION,
        "generator": GENERATOR_NAME,
        "classification_mode": "render_facts_only",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "source": {
            "name": source.name,
            "path": source.as_posix(),
            "sha256": before_hash,
            "readonly_input": True,
            "hash_verified_unchanged": True,
        },
        "render": {
            "pdf_path": source_pdf.as_posix(),
            "pdf_sha256": sha256_file(source_pdf),
            "dpi": dpi,
            **dict(renderer_metadata),
            "rasterizer": rasterizer,
        },
        "page_count": len(pages),
        "pages": pages,
        "issues": [],
    }


def render_page_manifest_markdown(manifest: Mapping[str, Any]) -> str:
    text_regions = sum(
        region["region_type"] == "text_line"
        for page in manifest["pages"]
        for region in page["regions"]
    )
    image_regions = sum(
        region["region_type"] == "image"
        for page in manifest["pages"]
        for region in page["regions"]
    )
    return "\n".join(
        [
            "# WPS 页面渲染清单",
            "",
            "> JSON 是机器权威事实；本文件仅供人工审核。",
            "",
            "## 安全状态",
            "",
            "- 自动绑定规则：否",
            "- 允许生产执行：否",
            f"- 原题哈希未改变：{'是' if manifest['source']['hash_verified_unchanged'] else '否'}",
            "",
            "## 摘要",
            "",
            f"- 原题：{manifest['source']['name']}",
            f"- 页面数：{manifest['page_count']}",
            f"- 文字区域：{text_regions}",
            f"- 图片区域：{image_regions}",
            f"- 渲染提供方：{manifest['render']['provider']}",
            f"- 生产页面真值：{'是' if manifest['render'].get('page_truth_authority') is True else '否'}",
            "",
        ]
    )


def write_page_render_artifacts(
    source_path: str | Path,
    pdf_path: str | Path,
    output_dir: str | Path,
    *,
    renderer_metadata: Mapping[str, Any],
    dpi: int = DEFAULT_DPI,
    pdftoppm_path: str | Path | None = None,
) -> dict[str, Path]:
    target_dir = Path(output_dir).resolve()
    manifest = build_page_render_manifest(
        source_path,
        pdf_path,
        target_dir,
        renderer_metadata=renderer_metadata,
        dpi=dpi,
        pdftoppm_path=pdftoppm_path,
    )
    target_dir.mkdir(parents=True, exist_ok=True)
    json_path = target_dir / "PageRenderManifest.json"
    markdown_path = target_dir / "PageRenderManifest.md"
    temp_json = target_dir / ".PageRenderManifest.json.tmp"
    temp_markdown = target_dir / ".PageRenderManifest.md.tmp"
    try:
        temp_json.write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        temp_markdown.write_text(render_page_manifest_markdown(manifest), encoding="utf-8")
        json.loads(temp_json.read_text(encoding="utf-8"))
        temp_markdown.read_text(encoding="utf-8")
        temp_json.replace(json_path)
        temp_markdown.replace(markdown_path)
    finally:
        temp_json.unlink(missing_ok=True)
        temp_markdown.unlink(missing_ok=True)
    return {"json": json_path, "markdown": markdown_path}


def _default_dispatch_factory(progid: str) -> Any:
    if sys.platform != "win32":
        raise RuntimeError("WPS COM 导出只支持 Windows；本机请使用已由 WPS 导出的 PDF")
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("缺少 pywin32，无法连接 WPS COM") from exc
    return win32com.client.Dispatch(progid)


def _create_wps_application(dispatch_factory: Callable[[str], Any]) -> tuple[Any, str]:
    errors = []
    for progid in WPS_PROGIDS:
        try:
            return dispatch_factory(progid), progid
        except Exception as exc:
            errors.append(f"{progid}: {exc}")
    raise RuntimeError("无法创建 WPS Writer COM 对象：" + "；".join(errors))


def export_docx_to_pdf_with_wps(
    source_path: str | Path,
    output_pdf: str | Path,
    *,
    dispatch_factory: Callable[[str], Any] | None = None,
) -> dict[str, Any]:
    source = Path(source_path).resolve()
    target = Path(output_pdf).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if _is_inside(target.parent, source.parent):
        raise ValueError("WPS 导出的 PDF 不能写入原题目录或其子目录")
    before_hash = sha256_file(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.unlink(missing_ok=True)
    factory = dispatch_factory or _default_dispatch_factory
    application = None
    document = None
    export_method = None
    progid = None
    application_version = "unknown"
    try:
        application, progid = _create_wps_application(factory)
        try:
            application_version = str(application.Version)
        except Exception:
            pass
        try:
            application.DisplayAlerts = False
        except Exception:
            pass
        document = application.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        export_error = None
        try:
            document.ExportAsFixedFormat(
                OutputFileName=str(target),
                ExportFormat=PDF_FORMAT,
            )
            export_method = "ExportAsFixedFormat"
        except Exception as exc:
            export_error = exc
            try:
                document.SaveAs2(str(target), PDF_FORMAT)
                export_method = "SaveAs2"
            except Exception:
                raise
        if not target.is_file() or target.stat().st_size <= 0:
            detail = f"；ExportAsFixedFormat 错误：{export_error}" if export_error else ""
            raise RuntimeError(f"WPS 未生成有效 PDF{detail}")
    except Exception:
        target.unlink(missing_ok=True)
        raise
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=0)
            except Exception:
                pass
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass
    after_hash = sha256_file(source)
    if before_hash != after_hash:
        target.unlink(missing_ok=True)
        raise RuntimeError("WPS 导出后原题 SHA256 发生变化，已删除 PDF 并停止")
    return {
        "provider": "wps_com",
        "renderer_role": "production_truth",
        "page_truth_authority": True,
        "layout_mode": "wps_pages",
        "progid": progid,
        "application_version": application_version,
        "export_method": export_method,
        "read_only_open": True,
        "source_hash_verified_unchanged": True,
    }
