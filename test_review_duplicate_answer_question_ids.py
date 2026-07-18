# coding: utf-8
from __future__ import annotations

import unittest

from shared_core.models import AnswerItem, AnswerUnit, QuestionUnit
from shared_core.review import build_review_report


def _question(question_id: str) -> QuestionUnit:
    return QuestionUnit(
        question_id=question_id,
        subject="文科",
        subject_overlay="zhongmei_chinese",
        grade_hint="高三",
        question_type="subjective",
        stem_blocks=[f"{question_id}．题目"],
        option_blocks=[],
        subquestions=[],
        media_blocks=[],
        material_blocks=[],
        source_span=(1, 1),
    )


def _answer(question_id: str, text: str, index: int) -> AnswerUnit:
    return AnswerUnit(
        question_id=question_id,
        answer_mode="whole",
        answer_items=[AnswerItem(item_id="whole", text=text)],
        analysis_items=[AnswerItem(item_id="whole", text="")],
        metadata={
            "mapping_method": "ordered_occurrence",
            "question_index": index,
        },
    )


class DuplicateAnswerReviewTests(unittest.TestCase):
    def test_review_pairs_duplicate_ids_by_occurrence(self):
        questions = [_question("1"), _question("2"), _question("1"), _question("2")]
        answers = [
            _answer("1", "甲一", 1),
            _answer("2", "甲二", 2),
            _answer("1", "乙一", 3),
            _answer("2", "乙二", 4),
        ]

        report = build_review_report("双篇文言文", questions, answers)

        self.assertEqual(report.summary["question_count"], 4)
        self.assertEqual(report.summary["answer_count"], 4)
        self.assertEqual(report.summary["high_risk_count"], 0)
        self.assertFalse(report.issues)

    def test_review_still_reports_extra_occurrence_as_orphan(self):
        questions = [_question("1"), _question("2")]
        answers = [
            _answer("1", "第一题", 1),
            _answer("2", "第二题", 2),
            AnswerUnit(
                question_id="1",
                answer_mode="whole",
                answer_items=[AnswerItem(item_id="whole", text="多余答案")],
                analysis_items=[AnswerItem(item_id="whole", text="")],
                review_flags=["orphan_answer"],
                metadata={"mapping_method": "orphan"},
            ),
        ]

        report = build_review_report("双篇文言文", questions, answers)

        self.assertEqual(report.summary["high_risk_count"], 1)
        self.assertEqual(report.issues[0].title, "存在孤立答案块")


if __name__ == "__main__":
    unittest.main()
