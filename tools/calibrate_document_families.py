from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.document_family_calibration import write_p1b_calibration_report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="校准 P1a 文档族阈值并生成 P1b 整批门禁。")
    parser.add_argument("family_report", type=Path, help="DocumentFamilyReport.json")
    parser.add_argument("ground_truth", type=Path, help="GroundTruth.json")
    parser.add_argument("--reviews", required=True, type=Path, help="VisualReviewSummary.json 或其上层目录")
    parser.add_argument("--output-dir", required=True, type=Path, help="校准报告目录")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        paths = write_p1b_calibration_report(
            args.family_report,
            args.ground_truth,
            args.reviews,
            args.output_dir,
        )
        report = json.loads(paths["json"].read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    print(
        json.dumps(
            {
                "recommended_threshold": report["recommended_threshold"],
                "calibration_scope": report["calibration_scope"],
                "gate_passed": report["gate"]["passed"],
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
