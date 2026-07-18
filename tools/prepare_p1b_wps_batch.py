from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.cli_output import configure_utf8_stdio
from shared_core.p1b_batch import write_wps_readonly_batch


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="为 P1b WPS 页面渲染准备外部只读副本。")
    parser.add_argument("batch_root", type=Path, help="包含两个活页目录的批次根目录")
    parser.add_argument("--output-dir", required=True, type=Path, help="原题目录之外的只读副本目录")
    parser.add_argument("--expected-count", type=int, help="预期正式 DOCX 数量")
    return parser.parse_args()


def main() -> int:
    configure_utf8_stdio()
    args = parse_args()
    try:
        path = write_wps_readonly_batch(
            args.batch_root,
            args.output_dir,
            expected_count=args.expected_count,
        )
        manifest = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    print(
        json.dumps(
            {
                "document_count": manifest["document_count"],
                "production_execution_enabled": manifest["production_execution_enabled"],
                "manifest": str(path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
