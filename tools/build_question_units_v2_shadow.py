from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.question_understanding_v2_shadow import (
    write_question_understanding_v2_shadow,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "为只读 DOCX 生成题目理解候选、QuestionUnitV2 和逐引用差异报告；"
            "不连接WPS，不执行快捷键。"
        )
    )
    parser.add_argument("sources", nargs="+", type=Path, help="一个或多个只读 DOCX 路径")
    parser.add_argument(
        "--output-dir",
        required=True,
        type=Path,
        help="源文档目录之外的影子产物根目录",
    )
    parser.add_argument(
        "--allow-label-reset-before",
        action="append",
        default=[],
        type=int,
        help="明确允许在第N个题块前发生题号重置；可重复指定",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    allowed_resets = tuple(sorted(set(args.allow_label_reset_before)))
    results = []
    for source in args.sources:
        target_dir = args.output_dir / source.stem
        paths = write_question_understanding_v2_shadow(
            source,
            target_dir,
            allowed_label_reset_before=allowed_resets,
        )
        result = json.loads(paths["question_units_json"].read_text(encoding="utf-8"))
        results.append(
            {
                "source": str(source.resolve()),
                "source_sha256": result["source_sha256"],
                "status": result["status"],
                "unit_count": result["unit_count"],
                "blocking_issue_count": len(result["blocking_issues"]),
                "mode": result["mode"],
                "production_execution_enabled": result[
                    "production_execution_enabled"
                ],
                "keypress_count": result["keypress_count"],
                "artifacts": {key: str(path) for key, path in paths.items()},
            }
        )
    print(
        json.dumps(
            {
                "message": (
                    "QuestionUnitV2 题目理解影子产物已生成；"
                    "旧生产入口未改变，按键数为0。"
                ),
                "results": results,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
