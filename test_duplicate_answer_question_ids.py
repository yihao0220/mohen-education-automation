# coding: utf-8
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx import Document

from shared_core import (
    build_answer_units_from_docx,
    build_answer_units_from_paragraph_texts,
    map_answers,
)
from shared_core.models import QuestionUnit


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


def _write_classical_answer_docx(
    temp_root: str,
    filename: str,
    paragraphs: list[str],
) -> Path:
    path = (
        Path(temp_root)
        / "众美-高三-语文"
        / "答案"
        / "文言文答案"
        / "选择性必修上册"
        / filename
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)
    return path


class DuplicateAnswerQuestionIdTests(unittest.TestCase):
    def test_classical_chinese_docx_keeps_duplicate_ids_in_source_order(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = _write_classical_answer_docx(
                temp_dir,
                "16《大学之道》《人皆有不忍人之心》_已清洗.docx",
                [
                    "《大学之道》",
                    "1．",
                    "(1)答案：第一篇第一问。",
                    "(2)答案：第一篇第二问。",
                    "解析：第一篇第一题解析。",
                    "2．",
                    "答案：第一篇第二题。",
                    "解析：第一篇第二题解析。",
                    "《人皆有不忍人之心》",
                    "1．",
                    "答案：第二篇第一题。",
                    "解析：第二篇第一题解析。",
                    "2．",
                    "答案：第二篇第二题。",
                    "解析：第二篇第二题解析。",
                ],
            )

            units = build_answer_units_from_docx(path)

        self.assertEqual([unit.question_id for unit in units], ["1", "2", "1", "2"])
        self.assertEqual(
            [unit.answer_items[0].text for unit in units],
            ["第一篇第一问。", "第一篇第二题。", "第二篇第一题。", "第二篇第二题。"],
        )
        self.assertEqual(units[0].answer_mode, "subquestion")
        self.assertTrue(units[0].metadata["allow_answer_defined_subquestions"])
        self.assertTrue(all("《" not in item.text for unit in units for item in unit.analysis_items))

        mapped = map_answers(
            [_question("1"), _question("2"), _question("1"), _question("2")],
            units,
        )
        self.assertTrue(mapped[0].metadata["answer_defined_subquestions"])
        self.assertNotIn("answer_split_but_question_whole", mapped[0].review_flags)

    def test_classical_title_boundary_is_not_enabled_for_other_contexts(self):
        units = build_answer_units_from_paragraph_texts(
            [
                "1．",
                "答案：第一题",
                "解析：正文引用书名。",
                "《普通引用》",
                "2．",
                "答案：第二题",
                "解析：第二题解析。",
            ],
            use_zhongmei_heading_boundaries=True,
        )

        self.assertEqual([unit.question_id for unit in units], ["1", "2"])
        self.assertIn("《普通引用》", units[0].analysis_items[0].text)

    def test_classical_chinese_answers_map_by_ordered_occurrence_with_duplicate_ids(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = _write_classical_answer_docx(
                temp_dir,
                "07《庖丁解牛》《烛之武退秦师》_已清洗.docx",
                [
                    "《庖丁解牛》",
                    "1．",
                    "答案：甲一",
                    "解析：甲一解析",
                    "2．",
                    "答案：甲二",
                    "解析：甲二解析",
                    "《烛之武退秦师》",
                    "1．",
                    "答案：乙一",
                    "解析：乙一解析",
                    "2．",
                    "答案：乙二",
                    "解析：乙二解析",
                ],
            )
            raw_answers = build_answer_units_from_docx(path)

        mapped = map_answers(
            [_question("1"), _question("2"), _question("1"), _question("2")],
            raw_answers,
        )

        self.assertEqual([unit.question_id for unit in mapped], ["1", "2", "1", "2"])
        self.assertEqual(
            [unit.answer_items[0].text for unit in mapped],
            ["甲一", "甲二", "乙一", "乙二"],
        )
        self.assertTrue(
            all(unit.metadata["mapping_method"] == "ordered_occurrence" for unit in mapped)
        )
        self.assertTrue(all("sequential_mapping" not in unit.review_flags for unit in mapped))

    def test_classical_chinese_ordered_mapping_accepts_source_id_mismatch(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = _write_classical_answer_docx(
                temp_dir,
                "20《苏武传》_已清洗.docx",
                [
                    "1．",
                    "答案：第一题",
                    "解析：第一题解析",
                    "2．",
                    "答案：第二题",
                    "解析：第二题解析",
                    "3．",
                    "答案：第三题",
                    "解析：第三题解析",
                    "4．",
                    "答案：第四题",
                    "解析：第四题解析",
                ],
            )
            raw_answers = build_answer_units_from_docx(path)

        mapped = map_answers(
            [_question("1"), _question("2"), _question("5"), _question("6")],
            raw_answers,
        )

        self.assertEqual([unit.question_id for unit in mapped], ["1", "2", "5", "6"])
        self.assertEqual(
            [unit.answer_items[0].text for unit in mapped],
            ["第一题", "第二题", "第三题", "第四题"],
        )
        self.assertEqual(
            [unit.metadata["original_question_id"] for unit in mapped],
            ["1", "2", "3", "4"],
        )
        self.assertTrue(all(unit.confidence == 1.0 for unit in mapped))
        self.assertTrue(all(not unit.review_flags for unit in mapped))

    def test_classical_single_explicit_subanswer_stays_in_f4_mode(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = _write_classical_answer_docx(
                temp_dir,
                "10《与妻书》《促织》_已清洗.docx",
                [
                    "《与妻书》",
                    "2．",
                    "(1)答案：唯一一条译文。",
                    "解析：",
                    "(1)得分点：关键字。",
                ],
            )
            unit = build_answer_units_from_docx(path)[0]

        self.assertEqual(unit.answer_mode, "subquestion")
        self.assertEqual([(item.item_id, item.text) for item in unit.answer_items], [("(1)", "唯一一条译文。")])
        self.assertEqual(len(unit.analysis_items), 1)
        self.assertTrue(unit.analysis_items[0].text.startswith("(1)得分点："))
        self.assertFalse(unit.metadata["force_whole_answer_input"])

    def test_classical_whole_answer_stays_f2_when_question_contains_child_items(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = _write_classical_answer_docx(
                temp_dir,
                "01《劝学》_已清洗.docx",
                [
                    "1．",
                    "答案：①木受绳则直；②虽有槁暴。",
                    "解析： ",
                ],
            )
            raw_answers = build_answer_units_from_docx(path)

        question = _question("1")
        question.subquestions = ["（1）第一处", "（2）第二处"]
        mapped = map_answers([question], raw_answers)

        self.assertEqual(mapped[0].answer_mode, "whole")
        self.assertTrue(mapped[0].metadata["force_whole_answer_input"])
        self.assertNotIn(
            "question_has_subquestions_but_answer_whole",
            mapped[0].review_flags,
        )


if __name__ == "__main__":
    unittest.main()
