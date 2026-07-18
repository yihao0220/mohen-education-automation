from __future__ import annotations

import re
from pathlib import Path

from .answer_core import infer_grouped_question_ids, question_expects_subquestion_answers
from .models import AnswerUnit, QuestionUnit, ReviewIssue, ReviewReport

EMBEDDED_ANSWER_SECTION_PATTERN = re.compile(r"(?:^|[\s，。；;:：])\d+\s*【答案】")
EMBEDDED_COMPACT_CHOICE_PATTERN = re.compile(
    r"(?<![0-9A-Za-z])(\d+)[．.]?\s*([A-D])(?=(?:\s*\d+[．.]?\s*[A-D])|$)"
)

QUESTION_WARNING_DETAILS = {
    "cross_page_media": "题目中的图片跨页或位置特殊，可能影响题块边界判断",
    "figure_reference_without_media": "题干提到了图或表，但当前没有检测到对应图片",
    "image_related_question": "这是图片/图表相关题，题干依赖图片或图表内容",
    "sparse_options": "选项行偏少，可能存在漏抓或选项未识别完整",
    "material_without_question": "材料块里没有识别到明确题号，可能是材料边界判断异常",
    "image_between_stem_and_options": "题干和选项之间夹有图片或表格，边界可能受影响",
}


def _join_answer_text(unit: AnswerUnit) -> str:
    return "\n".join(item.text.strip() for item in unit.answer_items if item.text and item.text.strip()).strip()


def _join_analysis_text(unit: AnswerUnit) -> str:
    return "\n".join(item.text.strip() for item in unit.analysis_items if item.text and item.text.strip()).strip()


def _analysis_only_subanswers(answer: AnswerUnit) -> bool:
    return bool(answer.metadata.get("analysis_only_subanswers"))


def _force_whole_answer_input(answer: AnswerUnit) -> bool:
    return bool(answer.metadata.get("force_whole_answer_input"))


def _answer_items_all_empty(answer: AnswerUnit) -> bool:
    return bool(answer.answer_items) and not any(item.text and item.text.strip() for item in answer.answer_items)


def _looks_like_contaminated_answer_block(unit: AnswerUnit) -> bool:
    payload = "\n".join(filter(None, [_join_answer_text(unit), _join_analysis_text(unit)])).strip()
    if not payload:
        return False
    if EMBEDDED_ANSWER_SECTION_PATTERN.search(payload):
        return True
    compact_pairs = EMBEDDED_COMPACT_CHOICE_PATTERN.findall(payload)
    return len(compact_pairs) >= 2


def format_question_warning_details(warnings: list[str]) -> list[str]:
    details: list[str] = []
    for warning in warnings:
        details.append(QUESTION_WARNING_DETAILS.get(warning, f"未分类风险：{warning}"))
    return details


def build_review_report(
    source_name: str,
    question_units: list[QuestionUnit],
    answer_units: list[AnswerUnit],
) -> ReviewReport:
    issues: list[ReviewIssue] = []
    ordered_occurrence_mapping = bool(answer_units) and any(
        unit.metadata.get("mapping_method") == "ordered_occurrence"
        for unit in answer_units
    ) and all(
        unit.metadata.get("mapping_method") in {"ordered_occurrence", "orphan"}
        for unit in answer_units
    )
    answer_map = {unit.question_id: unit for unit in answer_units}
    question_ids = {q.question_id for q in question_units}

    for question_index, question in enumerate(question_units):
        if question.requires_review:
            issues.append(
                ReviewIssue(
                    module="题目录入",
                    question_id=question.question_id,
                    severity="warning",
                    title="题块置信度较低",
                    detail=f"confidence={question.confidence:.2f}; warnings={', '.join(question.warnings) or '无'}",
                )
            )

        answer = (
            answer_units[question_index]
            if ordered_occurrence_mapping and question_index < len(answer_units)
            else answer_map.get(question.question_id)
        )
        if not answer:
            issues.append(
                ReviewIssue(
                    module="格式转换",
                    question_id=question.question_id,
                    severity="error",
                    title="未找到对应答案",
                    detail="题目结构已识别，但标准化答案中没有对应题号。",
                )
            )
            continue

        force_whole_answer_input = _force_whole_answer_input(answer)

        if question_expects_subquestion_answers(question) and answer.answer_mode != "subquestion":
            if _analysis_only_subanswers(answer):
                issues.append(
                    ReviewIssue(
                        module="格式转换",
                        question_id=question.question_id,
                        severity="warning",
                        title="小问答案转入解析承载",
                        detail="题目含小问，但答案区留空，解析里保留了带小问标记的内容；适用于公式不宜直接录入答案框的场景。",
                    )
                )
            elif force_whole_answer_input:
                pass
            else:
                issues.append(
                    ReviewIssue(
                        module="格式转换",
                        question_id=question.question_id,
                        severity="error",
                        title="题目含小问但答案未拆分",
                        detail="题目结构检测到小问，答案仍按整题处理，建议人工审核。",
                    )
                )
        elif question_expects_subquestion_answers(question) and answer.answer_mode == "subquestion":
            if _analysis_only_subanswers(answer):
                issues.append(
                    ReviewIssue(
                        module="格式转换",
                        question_id=question.question_id,
                        severity="warning",
                        title="小问答案转入解析承载",
                        detail="题目含小问，但答案区留空，解析里保留了带小问标记的内容；适用于公式不宜直接录入答案框的场景。",
                    )
                )
            elif (
                not force_whole_answer_input
                and not answer.metadata.get("allow_answer_defined_subquestions")
            ):
                grouped_question_ids = infer_grouped_question_ids(question)
                expected_sub_count = (
                    len(grouped_question_ids)
                    if len(grouped_question_ids) > 1
                    else len(getattr(question, "subquestions", []) or [])
                )
                actual_sub_count = len(answer.answer_items)
                if expected_sub_count and actual_sub_count != expected_sub_count:
                    issues.append(
                        ReviewIssue(
                            module="格式转换",
                            question_id=question.question_id,
                            severity="error",
                            title="题目与答案小问数量不一致",
                            detail=f"题目侧预期 {expected_sub_count} 个小问，答案侧识别为 {actual_sub_count} 个小问。",
                        )
                    )

        if "sequential_mapping" in answer.review_flags:
            issues.append(
                ReviewIssue(
                    module="格式转换",
                    question_id=question.question_id,
                    severity="error",
                    title="题答顺序映射，需人工确认",
                    detail=f"原答案题号={answer.metadata.get('original_question_id', '未知')}，当前按顺序兜底映射。",
                )
            )

        if "answer_split_but_question_whole" in answer.review_flags:
            issues.append(
                ReviewIssue(
                    module="格式转换",
                    question_id=question.question_id,
                    severity="error",
                    title="答案拆分与题目结构不一致",
                    detail="题目未识别为小问结构，但答案块已拆成多个小问。",
                )
            )

        suppress_force_whole_review = (
            force_whole_answer_input
            and not [flag for flag in answer.review_flags if flag not in {"empty_sub_answers", "question_has_subquestions_but_answer_whole"}]
            and (answer.answer_mode != "subquestion" or _answer_items_all_empty(answer) or _analysis_only_subanswers(answer))
        )
        suppress_analysis_only_review = (
            _analysis_only_subanswers(answer)
            and not [flag for flag in answer.review_flags if flag not in {"empty_sub_answers", "question_has_subquestions_but_answer_whole"}]
        )

        if answer.requires_review and not (suppress_force_whole_review or suppress_analysis_only_review):
            issues.append(
                ReviewIssue(
                    module="答案录入",
                    question_id=question.question_id,
                    severity="error" if answer.confidence < 0.75 or answer.review_flags else "warning",
                    title="答案块置信度较低",
                    detail=f"confidence={answer.confidence:.2f}; flags={', '.join(answer.review_flags) or '无'}",
                )
            )

        if _looks_like_contaminated_answer_block(answer):
            issues.append(
                ReviewIssue(
                    module="格式转换",
                    question_id=question.question_id,
                    severity="error",
                    title="答案块疑似串入其他题号",
                    detail="答案或解析中检测到其他题号答案片段，清洗结果可能混入后续题目内容。",
                )
            )

    if ordered_occurrence_mapping:
        orphan_answers = [
            unit.question_id
            for index, unit in enumerate(answer_units)
            if index >= len(question_units) or "orphan_answer" in unit.review_flags
        ]
    else:
        orphan_answers = [
            unit.question_id
            for unit in answer_units
            if unit.question_id not in question_ids or "orphan_answer" in unit.review_flags
        ]
    for qid in orphan_answers:
        issues.append(
            ReviewIssue(
                module="格式转换",
                question_id=qid,
                severity="error",
                title="存在孤立答案块",
                detail="标准化答案里有该题号，但题目结构中没有识别到对应题块。",
            )
        )

    summary = {
        "question_count": len(question_units),
        "answer_count": len(answer_units),
        "issue_count": len(issues),
        "high_risk_count": sum(1 for issue in issues if issue.severity == "error"),
    }
    return ReviewReport(source_name=source_name, summary=summary, issues=issues)


def export_review_report(report: ReviewReport, output_path: str | Path) -> str:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    lines = [
        f"# 审核清单：{report.source_name}",
        "",
        "## 摘要",
        f"- 题目数：{report.summary['question_count']}",
        f"- 答案块数：{report.summary['answer_count']}",
        f"- 问题数：{report.summary['issue_count']}",
        f"- 高风险数：{report.summary['high_risk_count']}",
        "",
        "## 问题列表",
    ]

    if not report.issues:
        lines.append("- 未发现需要人工审核的问题。")
    else:
        for issue in report.issues:
            lines.append(f"- [{issue.severity}] {issue.module} / 第{issue.question_id}题 / {issue.title}: {issue.detail}")

    output.write_text("\n".join(lines), encoding="utf-8")
    return str(output)
