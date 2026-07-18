from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.document_families import write_document_family_report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="读取 DocumentProfile 1.1，生成只读文档族与异常候选报告。"
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help="一个或多个 *_DocumentProfile.json 文件或包含画像的目录",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        type=Path,
        help="报告输出目录，不得覆盖输入画像",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    paths = write_document_family_report(args.inputs, args.output_dir)
    report = json.loads(paths["json"].read_text(encoding="utf-8"))
    print(
        json.dumps(
            {
                "classification_mode": report["classification_mode"],
                "input_profile_count": report["input_profile_count"],
                "family_count": len(report["families"]),
                "outlier_candidate_count": len(report["outlier_candidates"]),
                "unresolved_singleton_count": len(report["unresolved_singletons"]),
                "automatic_rule_binding_enabled": report[
                    "automatic_rule_binding_enabled"
                ],
                "production_execution_enabled": report["production_execution_enabled"],
                "artifacts": {key: str(path) for key, path in paths.items()},
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
