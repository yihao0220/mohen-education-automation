from __future__ import annotations

import plistlib
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from shared_core.document_render import sha256_file


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SWIFT_HELPER = PROJECT_ROOT / "tools" / "macos" / "quicklook_webkit_pdf.swift"


def _run(command: list[str], *, label: str, timeout: int = 180) -> None:
    try:
        subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except (OSError, subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
        stderr = getattr(exc, "stderr", "") or ""
        detail = f"：{stderr.strip()}" if stderr.strip() else ""
        raise RuntimeError(f"{label}失败{detail}") from exc


def export_docx_to_pdf_with_macos_quicklook(
    source_path: str | Path,
    output_pdf: str | Path,
    *,
    helper_source: str | Path = DEFAULT_SWIFT_HELPER,
) -> dict[str, Any]:
    if sys.platform != "darwin":
        raise RuntimeError("macOS Quick Look 渲染只支持 Mac")
    source = Path(source_path).resolve()
    target = Path(output_pdf).resolve()
    helper = Path(helper_source).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if not helper.is_file():
        raise FileNotFoundError(f"Quick Look WebKit 辅助程序不存在：{helper}")
    if target.parent.is_relative_to(source.parent):
        raise ValueError("Quick Look 导出的 PDF 不能写入原题目录或其子目录")
    qlmanage = shutil.which("qlmanage")
    swiftc = shutil.which("swiftc")
    if not qlmanage:
        raise FileNotFoundError("Mac 系统没有找到 qlmanage")
    if not swiftc:
        raise FileNotFoundError("Mac 系统没有找到 swiftc，请安装 Xcode Command Line Tools")

    before_hash = sha256_file(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.unlink(missing_ok=True)
    try:
        with tempfile.TemporaryDirectory(prefix="mohen-p1b-quicklook-") as temp_name:
            temp_dir = Path(temp_name)
            preview_root = temp_dir / "preview"
            preview_root.mkdir()
            _run(
                [qlmanage, "-p", "-o", str(preview_root), str(source)],
                label="Quick Look DOCX 预览",
            )
            preview_dirs = sorted(preview_root.glob("*.qlpreview"))
            if len(preview_dirs) != 1:
                raise RuntimeError(
                    f"Quick Look 预览目录数量异常：期望 1，实际 {len(preview_dirs)}"
                )
            preview_dir = preview_dirs[0]
            preview_html = preview_dir / "Preview.html"
            properties_path = preview_dir / "PreviewProperties.plist"
            if not preview_html.is_file() or not properties_path.is_file():
                raise RuntimeError("Quick Look 没有生成 Preview.html/PreviewProperties.plist")
            properties = plistlib.loads(properties_path.read_bytes())
            if properties.get("AllowNetworkAccess") is not False:
                raise RuntimeError("Quick Look 预览未明确禁止网络访问，已停止")
            helper_binary = temp_dir / "quicklook_webkit_pdf"
            _run(
                [swiftc, str(helper), "-o", str(helper_binary)],
                label="WebKit 辅助程序编译",
            )
            _run(
                [str(helper_binary), str(preview_html), str(target)],
                label="WebKit 连续预览 PDF 生成",
            )
            attachment_count = len(list(preview_dir.glob("Attachment*.png")))
            if attachment_count <= 0:
                raise RuntimeError("Quick Look 预览没有任何图片附件，需要人工复核")
            logical_width = int(properties.get("Width") or 0)
            logical_height = int(properties.get("Height") or 0)
            can_have_pages = properties.get("CanHavePages") is True
    except Exception:
        target.unlink(missing_ok=True)
        raise
    if not target.is_file() or target.stat().st_size <= 0:
        target.unlink(missing_ok=True)
        raise RuntimeError("WebKit 没有生成有效 PDF")
    after_hash = sha256_file(source)
    if before_hash != after_hash:
        target.unlink(missing_ok=True)
        raise RuntimeError("Quick Look 渲染后原题 SHA256 发生变化，已删除 PDF 并停止")
    return {
        "provider": "macos_quicklook_webkit",
        "renderer_role": "development_preview",
        "page_truth_authority": False,
        "layout_mode": "continuous_preview",
        "quicklook_generator": properties.get("BaseBundlePath"),
        "quicklook_can_have_pages": can_have_pages,
        "logical_preview_width": logical_width,
        "logical_preview_height": logical_height,
        "visible_attachment_count": attachment_count,
        "read_only_open": True,
        "source_hash_verified_unchanged": True,
    }
