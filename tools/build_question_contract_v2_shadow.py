from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.contract_v2_shadow import write_question_contract_v2_shadow


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="为只读 DOCX 生成 QuestionContractV2 影子对象和差异报告，不执行快捷键。"
    )
    parser.add_argument("sources", nargs="+", type=Path, help="一个或多个只读 DOCX 路径")
    parser.add_argument(
        "--output-dir",
        required=True,
        type=Path,
        help="源文档目录之外的影子产物根目录",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    results = []
    for source in args.sources:
        target_dir = args.output_dir / source.stem
        paths = write_question_contract_v2_shadow(source, target_dir)
        contract = json.loads(paths["contract_json"].read_text(encoding="utf-8"))
        results.append(
            {
                "source": str(source.resolve()),
                "source_sha256": contract["source_sha256"],
                "unit_count": contract["unit_count"],
                "mode": contract["mode"],
                "production_execution_enabled": contract["production_execution_enabled"],
                "keypress_count": contract["keypress_count"],
                "artifacts": {key: str(path) for key, path in paths.items()},
            }
        )
    print(
        json.dumps(
            {
                "message": "V2 交接单影子产物已生成；未改变生产入口，按键数为 0。",
                "results": results,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
