from __future__ import annotations

import argparse
import os
from pathlib import Path
import sys


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
from shared_core.cli_output import configure_utf8_stdio


DEFAULT_PROJECT_ROOT = Path(
    os.environ.get(
        "MOHEN_GUANMEI_BIOLOGY_DIR",
        Path.home() / "Documents" / "墨痕教育" / "莞美-高二-生物",
    )
)
ANSWER_NAME_SUFFIX = "-答案_已清洗"
QUESTION_NAME_SUFFIX = "-排版终稿.docx"


def resolve_answer_root(
    project_root: str | Path,
    answer_root: str | Path | None = None,
) -> Path:
    root = Path(project_root)
    if answer_root is not None:
        resolved = Path(answer_root)
        if not resolved.is_dir():
            raise FileNotFoundError(f"未找到答案目录：{resolved}")
        return resolved

    candidates = (
        root / "答案" / "按章节拆分",
        root / "答案",
    )
    for candidate in candidates:
        if candidate.is_dir() and next(candidate.rglob("*_已清洗.docx"), None):
            return candidate
    raise FileNotFoundError(f"未找到莞美高二生物已清洗答案：{root / '答案'}")


def discover_question_answer_pairs(
    project_root: str | Path,
    *,
    answer_root: str | Path | None = None,
    expected_count: int | None = None,
) -> tuple[Path, list[tuple[Path, Path]]]:
    root = Path(project_root)
    if not root.is_dir():
        raise FileNotFoundError(f"未找到莞美高二生物项目目录：{root}")
    resolved_answer_root = resolve_answer_root(root, answer_root)
    answers = sorted(
        path
        for path in resolved_answer_root.rglob("*_已清洗.docx")
        if not path.name.startswith(("~$", ".~", "~"))
    )
    if expected_count is not None and len(answers) != expected_count:
        raise ValueError(
            f"已清洗答案数量不符合预期：预期 {expected_count}，实际 {len(answers)}"
        )
    if not answers:
        raise FileNotFoundError(f"未找到已清洗答案：{resolved_answer_root}")

    pairs: list[tuple[Path, Path]] = []
    for answer_path in answers:
        relative = answer_path.relative_to(resolved_answer_root)
        if not answer_path.stem.endswith(ANSWER_NAME_SUFFIX):
            raise ValueError(f"未知的莞美答案命名，整批停止：{relative}")
        base_name = answer_path.stem[: -len(ANSWER_NAME_SUFFIX)]
        question_path = root / relative.parent / f"{base_name}{QUESTION_NAME_SUFFIX}"
        if not question_path.is_file():
            raise FileNotFoundError(
                f"未找到配对原题，整批停止：{relative} -> {question_path}"
            )
        pairs.append((question_path, answer_path))
    return resolved_answer_root, pairs


def audit_question_answer_pairs(
    pairs: list[tuple[Path, Path]],
) -> list[tuple[Path, Path, object]]:
    prepared: list[tuple[Path, Path, object]] = []
    for question_path, answer_path in pairs:
        question_units = build_question_units_from_docx(question_path, grade_hint="高二")
        raw_answer_units = build_answer_units_from_docx(answer_path)
        answer_units = map_answers(question_units, raw_answer_units)
        report = build_review_report(answer_path.name, question_units, answer_units)
        prepared.append((question_path, answer_path, report))

    blocked = [
        (answer_path, report)
        for _, answer_path, report in prepared
        if any(issue.severity == "error" for issue in report.issues)
    ]
    if blocked:
        details = []
        for answer_path, report in blocked:
            errors = [
                f"{issue.question_id}:{issue.title}"
                for issue in report.issues
                if issue.severity == "error"
            ]
            details.append(f"{answer_path.name}：{'，'.join(errors)}")
        raise ValueError("自动检查未通过，未写入任何放行状态：\n" + "\n".join(details))
    return prepared


def refresh_guanmei_biology_review_statuses(
    project_root: str | Path,
    *,
    answer_root: str | Path | None = None,
    expected_count: int | None = 27,
) -> list[dict]:
    _, pairs = discover_question_answer_pairs(
        project_root,
        answer_root=answer_root,
        expected_count=expected_count,
    )
    prepared = audit_question_answer_pairs(pairs)

    results: list[dict] = []
    for question_path, answer_path, report in prepared:
        initialize_review_status(answer_path, report=report)
        update_review_status(
            answer_path,
            status="approved",
            reviewer="system:guanmei_biology",
            note="莞美高二生物题答自动检查通过",
        )
        gate = get_review_gate_result(answer_path)
        if not gate.get("allowed"):
            raise RuntimeError(
                f"审核状态写入后门禁仍未通过：{answer_path.name}：{gate.get('reason')}"
            )
        results.append(
            {
                "question_path": str(question_path),
                "answer_path": str(answer_path),
                "status_path": gate["status_path"],
                "status": gate["status"],
                "allowed": gate["allowed"],
                "summary": report.summary,
            }
        )
    return results


def main() -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description="刷新莞美高二生物答案自动检查状态")
    parser.add_argument("project_root", nargs="?", type=Path, default=DEFAULT_PROJECT_ROOT)
    parser.add_argument("--answer-root", type=Path)
    parser.add_argument("--expected-count", type=int, default=27)
    args = parser.parse_args()

    results = refresh_guanmei_biology_review_statuses(
        args.project_root,
        answer_root=args.answer_root,
        expected_count=args.expected_count,
    )
    question_count = sum(item["summary"]["question_count"] for item in results)
    print(f"审核状态已刷新：{len(results)} 份")
    print(f"核对题数：{question_count}")
    print("门禁结果：全部 approved")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
