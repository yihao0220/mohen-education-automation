from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.cli_output import configure_utf8_stdio
from shared_core.document_visual_review import write_visual_review_artifacts


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成或验证页面视觉角色审核 JSON。")
    parser.add_argument("manifest", type=Path, help="PageRenderManifest.json")
    parser.add_argument("--output-dir", required=True, type=Path, help="审核产物目录")
    parser.add_argument("--review", type=Path, help="已人工填写的 VisualRoleReview.json")
    return parser.parse_args()


def main() -> int:
    configure_utf8_stdio()
    args = parse_args()
    try:
        paths = write_visual_review_artifacts(
            args.manifest,
            args.output_dir,
            review_path=args.review,
        )
        summary = json.loads(paths["summary_json"].read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    print(
        json.dumps(
            {
                "source": summary["source_name"],
                "gate_ready": summary["gate_ready"],
                "pending_count": summary["pending_count"],
                "artifacts": {key: str(path) for key, path in paths.items()},
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
