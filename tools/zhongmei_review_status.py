from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core import (
    build_answer_units_from_docx,
    build_question_units_from_docx,
    build_review_report,
    get_review_gate_result,
    initialize_review_status,
    map_answers,
    update_review_status,
)


DEFAULT_ANSWER_DIR = Path(
    r"D:\墨痕教育题目\众美-高三-语文\答案\对点练案答案"
)
CLEANED_FILE_PATTERN = re.compile(r"^对点练案(\d+)_已清洗\.docx$")


def infer_question_dir(answer_path: str | Path) -> Path:
    answer_path = Path(answer_path)
    return answer_path.parent.parent.parent / "对点练案"


def derive_question_path(
    answer_path: str | Path,
    question_dir: str | Path | None = None,
) -> Path:
    answer_path = Path(answer_path)
    stem = answer_path.stem
    if not stem.endswith("_已清洗"):
        raise ValueError(f"不是已清洗答案文档：{answer_path.name}")
    target_dir = Path(question_dir) if question_dir else infer_question_dir(answer_path)
    return target_dir / f"{stem[:-len('_已清洗')]}.docx"


def discover_cleaned_answers(answer_dir: str | Path) -> list[Path]:
    answer_dir = Path(answer_dir)
    files: list[tuple[int, Path]] = []
    for path in answer_dir.glob("对点练案*_已清洗.docx"):
        if path.name.startswith("~$"):
            continue
        match = CLEANED_FILE_PATTERN.match(path.name)
        if match:
            files.append((int(match.group(1)), path))
    files.sort(key=lambda item: item[0])
    return [path for _, path in files]


def refresh_review_statuses(
    answer_paths: list[str | Path],
    *,
    question_dir: str | Path | None = None,
    require_all_approved: bool = True,
) -> list[dict]:
    """重新执行题答检查，并为已清洗答案生成与文件签名绑定的审核状态。"""
    prepared: list[tuple[Path, Path, object]] = []
    for raw_answer_path in answer_paths:
        answer_path = Path(raw_answer_path)
        if not answer_path.exists():
            raise FileNotFoundError(f"未找到答案文档：{answer_path}")
        question_path = derive_question_path(answer_path, question_dir)
        if not question_path.exists():
            raise FileNotFoundError(
                f"未找到与 {answer_path.name} 配对的题目文档：{question_path}"
            )

        question_units = build_question_units_from_docx(question_path)
        raw_answer_units = build_answer_units_from_docx(answer_path)
        answer_units = map_answers(question_units, raw_answer_units)
        report = build_review_report(answer_path.name, question_units, answer_units)
        prepared.append((answer_path, question_path, report))

    blocked = [
        (answer_path, report)
        for answer_path, _, report in prepared
        if any(issue.severity == "error" for issue in report.issues)
    ]
    if blocked and require_all_approved:
        details = "；".join(
            f"{path.name}：{report.summary['high_risk_count']} 个高风险问题"
            for path, report in blocked
        )
        raise ValueError(f"自动检查未通过，未写入放行状态：{details}")

    results: list[dict] = []
    for answer_path, question_path, report in prepared:
        has_errors = any(issue.severity == "error" for issue in report.issues)
        status = "rejected" if has_errors else "approved"
        initialize_review_status(answer_path, report=report)
        update_review_status(
            answer_path,
            status=status,
            reviewer="system:zhongmei",
            note="众美题答自动检查未通过" if has_errors else "众美题答自动检查通过",
        )
        gate = get_review_gate_result(answer_path)
        results.append(
            {
                "answer_path": str(answer_path),
                "question_path": str(question_path),
                "status": status,
                "allowed": gate["allowed"],
                "summary": report.summary,
                "status_path": gate["status_path"],
            }
        )
    return results


def main() -> int:
    parser = argparse.ArgumentParser(description="刷新众美对点练案答案的自动检查状态")
    parser.add_argument("--answer-dir", type=Path, default=DEFAULT_ANSWER_DIR)
    parser.add_argument("--question-dir", type=Path)
    args = parser.parse_args()

    answer_paths = discover_cleaned_answers(args.answer_dir)
    if not answer_paths:
        raise FileNotFoundError(f"未找到已清洗答案文档：{args.answer_dir}")
    results = refresh_review_statuses(answer_paths, question_dir=args.question_dir)
    print(f"审核状态已刷新：{len(results)} 份，全部 approved。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
