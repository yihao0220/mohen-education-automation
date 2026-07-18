# coding: utf-8
from __future__ import annotations

import unittest
from pathlib import Path

from shared_core.question_core import build_question_units_from_docx
from shared_core.subject_overlay import should_use_native_table_input_for_context


QINGYAN_DOC = Path(r"D:\墨痕教育题目\清艳-五年级-数学-数拔-王逸豪\五年级暑假数拔教材.docx")


class QingyanMathQuestionScanTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QINGYAN_DOC.exists():
            raise unittest.SkipTest(f"缺少清艳真实样本: {QINGYAN_DOC}")
        cls.units = build_question_units_from_docx(QINGYAN_DOC)

    def test_qingyan_overlay_scans_real_textbook_questions(self):
        self.assertEqual(len(self.units), 119)
        self.assertEqual({unit.subject for unit in self.units}, {"理科"})
        self.assertEqual({unit.subject_overlay for unit in self.units}, {"qingyan_math"})

    def test_ignores_explanations_examples_and_chapter_boundaries(self):
        forbidden_keywords = (
            "基础知识讲解",
            "课堂例题",
            "例题",
            "随堂练习",
            "巩固练习",
            "出类拔萃",
            "实战演练",
            "结构特征",
            "公式展开证明",
            "公式图形证明",
            "完全平方公式变形",
            "第3章",
            "第4章",
            "第5章",
            "第6章",
            "第7章",
            "第8章",
        )

        for unit in self.units:
            with self.subTest(qid=unit.question_id, span=unit.source_span):
                text = " ".join(unit.stem_blocks + unit.option_blocks + unit.subquestions + unit.material_blocks)
                self.assertFalse(
                    [keyword for keyword in forbidden_keywords if keyword in text],
                    text[:160],
                )

    def test_numbered_section_labels_are_anchors_not_input_text(self):
        units_by_start = {unit.source_span[0]: unit for unit in self.units}

        self.assertEqual(units_by_start[585].question_id, "8")
        self.assertTrue(units_by_start[585].preview.startswith("（1）7÷0.125÷8"))
        self.assertEqual(units_by_start[620].question_id, "10")
        self.assertTrue(units_by_start[620].preview.startswith("（1）0.36×44"))
        self.assertEqual(units_by_start[664].question_id, "2")
        self.assertTrue(units_by_start[664].preview.startswith("（1）7÷0.125÷8"))

    def test_numbered_teaching_lines_are_not_questions(self):
        forbidden_starts = {634, 645, 1002, 1005, 1017, 1019}
        self.assertFalse(forbidden_starts & {unit.source_span[0] for unit in self.units})

    def test_uses_native_table_input_for_qingyan_textbook(self):
        self.assertTrue(
            should_use_native_table_input_for_context(
                QINGYAN_DOC.name,
                overlay_name="qingyan_math",
            )
        )


if __name__ == "__main__":
    unittest.main()
