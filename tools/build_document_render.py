from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.document_render import (
    export_docx_to_pdf_with_wps,
    write_page_render_artifacts,
)
from shared_core.macos_quicklook_render import export_docx_to_pdf_with_macos_quicklook


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="用 WPS 页面真值生成只读渲染清单。")
    parser.add_argument("source", type=Path, help="只读 DOCX 原题")
    parser.add_argument("--output-dir", required=True, type=Path, help="原题目录之外的产物目录")
    parser.add_argument("--pdf-input", type=Path, help="已由 WPS 导出的 PDF；不传则使用 Windows WPS COM")
    parser.add_argument(
        "--renderer",
        choices=("auto", "wps", "macos-quicklook"),
        default="auto",
        help="默认 Mac 用 Quick Look 开发预览，Windows 用 WPS 生产真值",
    )
    parser.add_argument("--dpi", type=int, default=144, help="PNG 渲染 DPI（默认 144）")
    parser.add_argument("--pdftoppm", type=Path, help="pdftoppm 可执行文件路径")
    return parser.parse_args()


def _run(args: argparse.Namespace) -> dict[str, object]:
    source = args.source.resolve()
    if source.name.startswith(("~$", ".~")):
        raise ValueError(f"拒绝 WPS 临时文件：{source.name}")
    target = args.output_dir.resolve()
    if args.pdf_input is not None:
        pdf_path = args.pdf_input.resolve()
        renderer_metadata = {
            "provider": "preexported_wps_pdf",
            "renderer_role": "production_truth",
            "page_truth_authority": True,
            "layout_mode": "wps_pages",
            "read_only_open": True,
            "source_hash_verified_unchanged": True,
        }
    elif args.renderer == "wps" or (args.renderer == "auto" and sys.platform == "win32"):
        pdf_path = target / "wps-render.pdf"
        renderer_metadata = export_docx_to_pdf_with_wps(source, pdf_path)
    elif args.renderer == "macos-quicklook" or (
        args.renderer == "auto" and sys.platform == "darwin"
    ):
        pdf_path = target / "macos-quicklook-preview.pdf"
        renderer_metadata = export_docx_to_pdf_with_macos_quicklook(source, pdf_path)
    else:
        raise RuntimeError("当前平台无可用的自动渲染器；请传入 --pdf-input")
    paths = write_page_render_artifacts(
        source,
        pdf_path,
        target,
        renderer_metadata=renderer_metadata,
        dpi=args.dpi,
        pdftoppm_path=args.pdftoppm,
    )
    manifest = json.loads(paths["json"].read_text(encoding="utf-8"))
    return {
        "source": str(source),
        "page_count": manifest["page_count"],
        "render_provider": manifest["render"]["provider"],
        "page_truth_authority": manifest["render"].get("page_truth_authority") is True,
        "production_execution_enabled": manifest["production_execution_enabled"],
        "artifacts": {key: str(path) for key, path in paths.items()},
    }


def main() -> int:
    try:
        result = _run(parse_args())
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
