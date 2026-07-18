# coding: utf-8
import unittest

from shared_core.answer_core import build_answer_units_from_paragraph_texts


class CompoundAnswerIdTests(unittest.TestCase):
    def test_compound_subquestion_labels_are_separate_answer_units(self):
        units = build_answer_units_from_paragraph_texts(
            [
                "3．（1）24（场）",
                "解析：4×3÷2=6（场），4×6=24（场）",
                "3．（2）7（场）",
                "解析：8-1=7（场）",
            ]
        )

        self.assertEqual([unit.question_id for unit in units], ["3．（1）", "3．（2）"])
        self.assertEqual(units[0].answer_items[0].text, "24（场）")
        self.assertEqual(units[1].answer_items[0].text, "7（场）")
        self.assertIn("4×6=24", units[0].analysis_items[0].text)

    def test_numbered_lines_inside_analysis_do_not_start_new_answers(self):
        units = build_answer_units_from_paragraph_texts(
            [
                "8．4",
                "解析：",
                "1.第一名(5 分):1 胜 2 平",
                "2.第二名(4 分):1 胜 1 平 1 负",
                "3.第三名(3 分):3 平",
                "4.第四名(2 分):2 平 1 负",
                "9．1 平 2 负",
                "解析：D 队 0 胜 1 平 2 负",
            ]
        )

        self.assertEqual([unit.question_id for unit in units], ["8", "9"])
        self.assertIn("2.第二名", units[0].analysis_items[0].text)
        self.assertEqual(units[1].answer_items[0].text, "1 平 2 负")

    def test_multiline_calculation_process_does_not_require_empty_analysis_label(self):
        units = build_answer_units_from_paragraph_texts(
            [
                "1．（1）=15×10000+15×1000+15×100+15×10+15×1",
                "=15×11111",
                "=166665",
                "1．（2）=40×(100+10+1)",
                "=40×111",
                "=4440",
            ]
        )

        self.assertEqual([unit.question_id for unit in units], ["1．（1）", "1．（2）"])
        self.assertFalse(units[0].review_flags)
        self.assertFalse(units[1].review_flags)

    def test_math_section_headings_do_not_enter_previous_answer(self):
        units = build_answer_units_from_paragraph_texts(
            [
                "6．（2）=2×20170000",
                "=4034000",
                "三、“运算律的复杂应用”思维冲浪：",
                "基础训练",
                "7．（1）=18×1000+18×100+18×10+18×1",
                "=19998",
            ]
        )

        self.assertEqual([unit.question_id for unit in units], ["6．（2）", "7．（1）"])
        self.assertNotIn("思维冲浪", units[0].answer_items[0].text)
        self.assertNotIn("基础训练", units[0].answer_items[0].text)


if __name__ == "__main__":
    unittest.main()
