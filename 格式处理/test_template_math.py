# coding: utf-8

import os
import sys
import unittest
import importlib.util
from pathlib import Path

from docx import Document


PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "格式处理"))
sys.path.insert(0, str(PROJECT_ROOT))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from shared_core.answer_core import build_answer_units_from_paragraph_texts, map_answers
from shared_core.models import AnswerItem, AnswerUnit, QuestionUnit
from shared_core.review import build_review_report
from 格式模板库 import template_b, template_d, template_e, template_math


MATH_ANSWER_DOC = Path(r"D:\墨痕教育题目\安乡金海-初二-数学\答案\第6次限时训练参考答案.docx")
MATH_DETAIL_DOC = Path(r"D:\墨痕教育题目\安乡金海-初二-数学\答案\限训7答案.docx")
FORMAT_MAIN_SPEC = importlib.util.spec_from_file_location("mohen_format_main_test", PROJECT_ROOT / "格式处理" / "main.py")
format_main = importlib.util.module_from_spec(FORMAT_MAIN_SPEC)
assert FORMAT_MAIN_SPEC and FORMAT_MAIN_SPEC.loader
FORMAT_MAIN_SPEC.loader.exec_module(format_main)


class FakeRange:
    def __init__(self, text):
        self.Text = text


class FakeParagraph:
    def __init__(self, text):
        self.Range = FakeRange(text)


class FakeParagraphCollection:
    def __init__(self, texts):
        self._items = [FakeParagraph(text) for text in texts]
        self.Count = len(self._items)

    def __call__(self, index):
        return self._items[index - 1]


class FakeDoc:
    def __init__(self, name, texts):
        self.Name = name
        self.Paragraphs = FakeParagraphCollection(texts)


def load_nonempty_paragraphs(path):
    doc = Document(path)
    return [paragraph.text.strip() for paragraph in doc.paragraphs if paragraph.text.strip()]


class MathTemplateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.sample_texts = load_nonempty_paragraphs(MATH_ANSWER_DOC)
        cls.detail_texts = load_nonempty_paragraphs(MATH_DETAIL_DOC)

    def test_math_template_wins_match_for_reference_answer_doc(self):
        fake_doc = FakeDoc(MATH_ANSWER_DOC.name, self.sample_texts)
        template, score = format_main.match_best_template(fake_doc, interactive=False, doc_name=MATH_ANSWER_DOC.name)
        self.assertIs(template, template_math)
        self.assertGreater(score, 0.05)

    def test_math_template_scores_above_existing_templates(self):
        math_score = template_math.match_score(None, cached_texts=self.sample_texts)
        self.assertGreater(math_score, template_b.match_score(None, cached_texts=self.sample_texts))
        self.assertGreater(math_score, template_d.match_score(None, cached_texts=self.sample_texts))
        self.assertGreater(math_score, template_e.match_score(None, cached_texts=self.sample_texts))

    def test_reference_answer_doc_renders_parseable_units(self):
        entries = template_math.parse_paragraph_texts(self.sample_texts)
        rendered_lines = template_math.render_standard_lines(entries)
        units = build_answer_units_from_paragraph_texts(rendered_lines)
        unit_map = {unit.question_id: unit for unit in units}

        self.assertEqual(len(units), 15)
        self.assertEqual(unit_map["1"].answer_items[0].text, "C")
        self.assertEqual(unit_map["6"].answer_items[0].text, "C")
        self.assertEqual(unit_map["13"].answer_mode, "subquestion")
        self.assertTrue(unit_map["13"].analysis_items)

    def test_detail_answer_doc_preserves_choice_analysis(self):
        entries = template_math.parse_paragraph_texts(self.detail_texts)
        rendered_lines = template_math.render_standard_lines(entries)
        units = build_answer_units_from_paragraph_texts(rendered_lines)
        unit_map = {unit.question_id: unit for unit in units}

        self.assertEqual(unit_map["1"].answer_items[0].text, "C")
        self.assertIn("样本", unit_map["1"].analysis_items[0].text)
        self.assertTrue(unit_map["6"].analysis_items[0].text)

    def test_math_italic_unicode_is_normalized_to_plain_text(self):
        entries = template_math.parse_paragraph_texts(["10．𝑦=2𝑥+3"])
        rendered_lines = template_math.render_standard_lines(entries)

        self.assertEqual(rendered_lines[0], "10．y=2x+3")

    def test_formula_subanswers_in_analysis_are_not_hard_blocked(self):
        question = QuestionUnit(
            question_id="13",
            subject="理科",
            subject_overlay=None,
            grade_hint=None,
            question_type="subjective",
            stem_blocks=["13．已知..."],
            option_blocks=[],
            subquestions=["(1) 求A", "(2) 求B"],
            media_blocks=[],
            material_blocks=[],
            source_span=(1, 4),
        )
        raw_answer = AnswerUnit(
            question_id="13",
            answer_mode="whole",
            answer_items=[AnswerItem(item_id="13", text="")],
            analysis_items=[AnswerItem(item_id="13", text="【详解】（1）解：...（2）解：...")],
        )

        mapped = map_answers([question], [raw_answer])
        report = build_review_report("math-formula.docx", [question], mapped)

        self.assertTrue(mapped[0].metadata.get("analysis_only_subanswers"))
        self.assertEqual(report.summary["high_risk_count"], 0)
        self.assertTrue(any(issue.title == "小问答案转入解析承载" for issue in report.issues))


if __name__ == "__main__":
    unittest.main()
