from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.document_preflight import write_preflight_artifacts


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="为 DOCX 生成只读文档画像与 F1 预检动作计划。")
    parser.add_argument("sources", nargs="+", type=Path, help="一个或多个只读 DOCX 路径")
    parser.add_argument("--output-dir", required=True, type=Path, help="源文档之外的产物目录")
    parser.add_argument("--no-docling", action="store_true", help="跳过 Docling，只运行原生扫描")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    summaries = []
    for source in args.sources:
        target_dir = args.output_dir / source.stem
        paths = write_preflight_artifacts(
            source,
            target_dir,
            include_docling=not args.no_docling,
        )
        profile = json.loads(paths["profile_json"].read_text(encoding="utf-8"))
        plan = json.loads(paths["plan_json"].read_text(encoding="utf-8"))
        summaries.append(
            {
                "source": str(source.resolve()),
                "source_sha256": profile["source"]["sha256"],
                "question_action_count": len(plan["actions"]),
                "execution_enabled": plan["execution_enabled"],
                "artifacts": {key: str(path.resolve()) for key, path in paths.items()},
            }
        )
    print(json.dumps({"results": summaries}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
