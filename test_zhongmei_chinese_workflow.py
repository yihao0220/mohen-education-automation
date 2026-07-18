# coding: utf-8
from __future__ import annotations

import importlib.util
import re
import sys
import unittest
from pathlib import Path

from shared_core import (
    build_question_units_from_docx,
    build_question_units_from_nodes,
    detect_subject_overlay,
)
from shared_core.answer_core import infer_grouped_question_ids, map_answers
from shared_core.models import AnswerItem, AnswerUnit, DocNode
from shared_core.question_core import scan_docx_nodes
from shared_core.subject_overlay import (
    is_leading_context_start_for_context,
    is_question_input_excluded_for_context,
)
from tools.clean_zhongmei_chinese_answers import parse_document


QUESTION_DIR = Path(
    r"D:\墨痕教育题目\众美-高三-语文\对点练案"
)
ANSWER_DIR = QUESTION_DIR.parent / "答案" / "对点练案答案"
PROJECT_ROOT = Path(__file__).resolve().parent

EXPECTED_QUESTION_COUNTS = {
    1: 12,
    2: 9,
    3: 12,
    4: 8,
    5: 8,
    6: 6,
    7: 6,
    8: 7,
    9: 6,
    10: 8,
    11: 6,
    12: 8,
    13: 5,
    14: 6,
    15: 4,
    16: 7,
    17: 10,
    18: 11,
    19: 10,
    20: 9,
    21: 9,
    22: 7,
    23: 7,
    24: 8,
    25: 6,
    26: 7,
    27: 6,
    28: 6,
    29: 6,
    30: 7,
    31: 30,
    32: 14,
    33: 14,
    34: 12,
    35: 9,
    36: 10,
    37: 13,
    38: 12,
    39: 10,
    40: 12,
    41: 9,
    42: 10,
    43: 9,
    44: 9,
}

EXPECTED_UNIT_COUNTS = {
    1: 8,
    2: 5,
    3: 8,
    4: 4,
    5: 4,
    6: 1,
    7: 1,
    8: 1,
    9: 1,
    10: 1,
    11: 1,
    12: 1,
    13: 1,
    14: 1,
    15: 1,
    16: 1,
    17: 9,
    18: 9,
    19: 10,
    20: 8,
    21: 7,
    22: 7,
    23: 7,
    24: 8,
    25: 3,
    26: 3,
    27: 3,
    28: 3,
    29: 3,
    30: 3,
    31: 30,
    32: 14,
    33: 14,
    34: 12,
    35: 9,
    36: 10,
    37: 13,
    38: 12,
    39: 10,
    40: 12,
    41: 9,
    42: 10,
    43: 9,
    44: 9,
}

USER_IGNORED_LINES = {
    "一、专项训练",
    "(一)请比对下列选项与原文，说明选项错在何处。",
    "(二)图文解读题",
    "(一)分析评价信息",
    "(二)逻辑推断",
    "(三)分析论证特点",
    "(二)梳理论述思路",
    "(一)分析理据关系",
    "(一)阅读下面的文字，完成文后题目。",
    "(三)阅读下面的文字，完成文后题目。",
}

SECTION_TITLE_PATTERN = re.compile(
    r"^\s*[一二三四五六七八九十]+[、．.]\s*[^。？！]{1,24}$"
)


def _lesson_number(path: Path) -> int:
    match = re.search(r"\d+", path.stem)
    if not match:
        raise AssertionError(f"文件名缺少编号: {path.name}")
    return int(match.group())


def _load_fast_blade_core_parser():
    fast_blade_dir = PROJECT_ROOT / "墨痕快刀"
    module_path = fast_blade_dir / "core_parser.py"
    spec = importlib.util.spec_from_file_location("zhongmei_core_parser", module_path)
    if not spec or not spec.loader:
        raise AssertionError(f"无法加载题目录入核心: {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.path.insert(0, str(fast_blade_dir))
    try:
        spec.loader.exec_module(module)
    finally:
        sys.path.remove(str(fast_blade_dir))
    return module


class ZhongmeiChineseQuestionInputTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.files = sorted(
            (
                path
                for path in QUESTION_DIR.glob("*.docx")
                if not path.name.startswith("~$")
            ),
            key=_lesson_number,
        )
        if len(cls.files) != 44:
            raise unittest.SkipTest(f"众美真实题目样本数量不是 44: {len(cls.files)}")

        cls.nodes = {path: scan_docx_nodes(path) for path in cls.files}
        cls.units = {
            path: build_question_units_from_docx(path) for path in cls.files
        }

    def test_all_44_documents_use_zhongmei_chinese_overlay(self):
        for path in self.files:
            sample_text = " ".join(node.text for node in self.nodes[path][:20])
            with self.subTest(path=path.name):
                self.assertEqual(
                    detect_subject_overlay(path.name, sample_text, base_subject="文科"),
                    "zhongmei_chinese",
                )
                self.assertEqual(
                    {unit.subject_overlay for unit in self.units[path]},
                    {"zhongmei_chinese"},
                )
                self.assertEqual({unit.subject for unit in self.units[path]}, {"文科"})

    def test_all_44_documents_group_complete_readings_without_losing_questions(self):
        self.assertEqual(sum(EXPECTED_QUESTION_COUNTS.values()), 400)
        self.assertEqual(sum(EXPECTED_UNIT_COUNTS.values()), 296)
        self.assertEqual(set(map(_lesson_number, self.files)), set(EXPECTED_UNIT_COUNTS))

        for path in self.files:
            lesson = _lesson_number(path)
            units = self.units[path]
            with self.subTest(path=path.name):
                self.assertEqual(len(units), EXPECTED_UNIT_COUNTS[lesson])
                self.assertEqual(
                    [
                        question_id
                        for unit in units
                        for question_id in infer_grouped_question_ids(unit)
                    ],
                    [str(index) for index in range(1, EXPECTED_QUESTION_COUNTS[lesson] + 1)],
                )

    def test_ignored_lines_and_section_titles_never_enter_question_spans(self):
        for path in self.files:
            texts_by_index = {
                node.index: (node.text or "").strip() for node in self.nodes[path]
            }
            for unit in self.units[path]:
                span_texts = [
                    texts_by_index.get(index, "")
                    for index in range(unit.source_span[0], unit.source_span[1] + 1)
                ]
                with self.subTest(
                    path=path.name,
                    question_id=unit.question_id,
                    span=unit.source_span,
                ):
                    self.assertFalse(USER_IGNORED_LINES.intersection(span_texts))
                    self.assertFalse(
                        [text for text in span_texts if SECTION_TITLE_PATTERN.match(text)]
                    )

    def test_internal_practice_headings_are_excluded_from_grouped_f1_content(self):
        affected_lessons = set()
        excluded_count = 0

        for path in self.files:
            lesson = _lesson_number(path)
            texts_by_index = {
                node.index: (node.text or "").strip() for node in self.nodes[path]
            }
            for unit in self.units[path]:
                excluded = [
                    texts_by_index.get(index, "")
                    for index in range(unit.source_span[0], unit.source_span[1] + 1)
                    if is_question_input_excluded_for_context(
                        texts_by_index.get(index, ""),
                        overlay_name=unit.subject_overlay,
                    )
                ]
                if excluded:
                    affected_lessons.add(lesson)
                    excluded_count += len(excluded)
                    self.assertEqual(unit.node_type, "LEADING_CONTEXT")
                    self.assertGreater(len(infer_grouped_question_ids(unit)), 1)

        self.assertEqual(affected_lessons, set(range(6, 16)))
        self.assertEqual(excluded_count, 20)

    def test_doc1_complete_reading_is_one_question_unit(self):
        path = QUESTION_DIR / "对点练案1.docx"
        units = self.units[path]
        self.assertEqual([unit.question_id for unit in units], [str(i) for i in range(1, 9)])
        self.assertEqual(units[-1].node_type, "LEADING_CONTEXT")
        self.assertEqual(units[-1].source_span, (53, 87))
        self.assertEqual(infer_grouped_question_ids(units[-1]), ["8", "9", "10", "11", "12"])

    def test_doc1_complete_reading_maps_all_five_answers_without_orphans(self):
        path = QUESTION_DIR / "对点练案1.docx"
        reading_unit = self.units[path][-1]
        raw_answers = [
            AnswerUnit(
                question_id=str(question_id),
                answer_mode="whole",
                answer_items=[AnswerItem(item_id="", text=f"答案{question_id}")],
                analysis_items=[AnswerItem(item_id="", text=f"解析{question_id}")],
            )
            for question_id in range(8, 13)
        ]

        mapped = map_answers([reading_unit], raw_answers)

        self.assertEqual(len(mapped), 1)
        self.assertEqual(mapped[0].question_id, "8")
        self.assertEqual(mapped[0].metadata["grouped_question_ids"], ["8", "9", "10", "11", "12"])
        self.assertEqual(mapped[0].metadata["original_question_ids"], ["8", "9", "10", "11", "12"])
        self.assertNotIn("orphan_answer", mapped[0].review_flags)

    def test_all_44_real_answers_map_to_grouped_units_without_orphans(self):
        answer_paths = {
            lesson: ANSWER_DIR / f"对点练案{lesson}.docx"
            for lesson in EXPECTED_QUESTION_COUNTS
        }
        if any(not path.exists() for path in answer_paths.values()):
            self.skipTest("缺少44份真实答案样本")

        for question_path in self.files:
            lesson = _lesson_number(question_path)
            clean_document = parse_document(answer_paths[lesson])
            raw_answers = []
            for question in clean_document.questions:
                raw_answers.append(
                    AnswerUnit(
                        question_id=str(question.number),
                        answer_mode="subquestion" if len(question.occurrences) > 1 else "whole",
                        answer_items=[
                            AnswerItem(
                                item_id=occurrence.label or "",
                                text=" ".join(occurrence.answer_lines),
                            )
                            for occurrence in question.occurrences
                        ],
                        analysis_items=[
                            AnswerItem(
                                item_id=occurrence.label or "",
                                text=" ".join(occurrence.analysis_lines),
                            )
                            for occurrence in question.occurrences
                            if occurrence.analysis_lines
                        ],
                    )
                )

            mapped = map_answers(self.units[question_path], raw_answers)
            consumed_question_ids = []
            for answer in mapped:
                consumed_question_ids.extend(
                    answer.metadata.get("original_question_ids")
                    or [answer.metadata.get("original_question_id", answer.question_id)]
                )

            with self.subTest(path=question_path.name):
                self.assertEqual(len(mapped), EXPECTED_UNIT_COUNTS[lesson])
                self.assertFalse(
                    [answer for answer in mapped if "orphan_answer" in answer.review_flags]
                )
                self.assertEqual(
                    consumed_question_ids,
                    [str(index) for index in range(1, EXPECTED_QUESTION_COUNTS[lesson] + 1)],
                )

    def test_future_top_level_section_stops_complete_reading_group(self):
        nodes = [
            DocNode(1, "阅读下面的文字，完成文后题目。"),
            DocNode(2, "这是阅读文章正文。"),
            DocNode(3, "1．阅读题一"),
            DocNode(4, "A．选项"),
            DocNode(5, "2．阅读题二"),
            DocNode(6, "二、语言文字运用"),
            DocNode(7, "3．后续独立题"),
            DocNode(8, "A．选项"),
        ]

        units = build_question_units_from_nodes(
            doc_name="对点练案后续样例.docx",
            subject_name="文科",
            nodes=nodes,
            overlay_name="zhongmei_chinese",
        )

        self.assertEqual([unit.question_id for unit in units], ["1", "3"])
        self.assertEqual(units[0].source_span, (2, 5))
        self.assertEqual(infer_grouped_question_ids(units[0]), ["1", "2"])
        self.assertEqual(units[1].source_span, (7, 8))

    def test_each_leading_reading_material_contains_all_following_questions(self):
        expected_spans = {
            "对点练案4.docx": {
                "1": ((5, 10), ["1"]),
                "2": ((12, 19), ["2"]),
                "3": ((21, 34), ["3"]),
                "4": ((37, 67), ["4", "5", "6", "7", "8"]),
            },
            "对点练案6.docx": {
                "1": ((4, 72), ["1", "2", "3", "4", "5", "6"]),
            },
            "对点练案17.docx": {"9": ((39, 50), ["9", "10"])},
            "对点练案30.docx": {
                "1": ((5, 32), ["1", "2", "3"]),
                "4": ((34, 49), ["4", "5"]),
                "6": ((52, 74), ["6", "7"]),
            },
        }

        for file_name, spans in expected_spans.items():
            path = QUESTION_DIR / file_name
            units_by_id = {unit.question_id: unit for unit in self.units[path]}
            for question_id, (expected_span, expected_question_ids) in spans.items():
                with self.subTest(path=file_name, question_id=question_id):
                    self.assertEqual(
                        units_by_id[question_id].source_span,
                        expected_span,
                    )
                    self.assertEqual(
                        infer_grouped_question_ids(units_by_id[question_id]),
                        expected_question_ids,
                    )

    def test_all_non_numbered_reading_prompts_anchor_the_next_question(self):
        prompt_count = 0
        for path in self.files:
            unit_starts = {unit.source_span[0] for unit in self.units[path]}
            for node in self.nodes[path]:
                text = (node.text or "").strip()
                if not (
                    "阅读下面" in text
                    and ("完成" in text or "回答" in text)
                    and not re.match(r"^\s*\d+\s*[．.、]", text)
                ):
                    continue
                prompt_count += 1
                with self.subTest(path=path.name, index=node.index, text=text[:80]):
                    self.assertTrue(
                        is_leading_context_start_for_context(
                            text,
                            overlay_name="zhongmei_chinese",
                        )
                    )
                    self.assertIn(node.index + 1, unit_starts)

        self.assertEqual(prompt_count, 46)

    def test_section_boundaries_stop_previous_question(self):
        path = QUESTION_DIR / "对点练案31.docx"
        units_by_id = {unit.question_id: unit for unit in self.units[path]}
        self.assertEqual(units_by_id["10"].source_span, (14, 14))
        self.assertEqual(units_by_id["20"].source_span, (26, 26))

    def test_wps_execution_does_not_retruncate_leading_context_units(self):
        core_parser = _load_fast_blade_core_parser()
        unit = type(
            "Unit",
            (),
            {
                "node_type": "LEADING_CONTEXT",
                "question_type": "material_choice",
                "warnings": [],
                "media_blocks": [],
                "subquestions": [],
            },
        )()

        self.assertTrue(
            core_parser.should_ignore_inline_obstacle(
                unit,
                "（一）高考专练",
                "Overlay_IgnoredQuestionStart",
            )
        )
        self.assertTrue(
            core_parser.should_ignore_inline_obstacle(
                unit,
                "二、综合训练",
                "Overlay_SpanBoundary",
            )
        )
        self.assertTrue(
            core_parser.should_ignore_inline_obstacle(
                unit,
                "",
                "Img_DecorativeHeader",
            )
        )

    def test_wps_execution_collects_internal_practice_headings_for_filtered_copy(self):
        core_parser = _load_fast_blade_core_parser()

        class FakeRange:
            def __init__(self, text):
                self.Text = text

        class FakeParagraph:
            def __init__(self, text):
                self.Range = FakeRange(text)

        class FakeParagraphs:
            def __init__(self, texts):
                self._paragraphs = [FakeParagraph(text) for text in texts]

            def __call__(self, index):
                return self._paragraphs[index - 1]

        paragraphs = FakeParagraphs(
            [
                "文章正文",
                "(一)典题专练",
                "1．第一题",
                "（二）高考专练",
                "2．第二题",
            ]
        )

        self.assertEqual(
            core_parser.collect_input_excluded_paragraph_indices(
                paragraphs,
                1,
                5,
                overlay_name="zhongmei_chinese",
            ),
            [2, 4],
        )

    def test_filtered_copy_allows_media_but_blocks_native_tables(self):
        core_parser = _load_fast_blade_core_parser()

        self.assertTrue(
            core_parser.is_filtered_input_copy_safe(
                contains_table=False,
                has_media=True,
            )
        )
        self.assertTrue(
            core_parser.is_filtered_input_copy_safe(
                contains_table=False,
                has_media=False,
            )
        )
        self.assertFalse(
            core_parser.is_filtered_input_copy_safe(
                contains_table=True,
                has_media=False,
            )
        )
        self.assertEqual(
            core_parser.describe_ignored_inline_obstacle(
                "（一）高考专练",
                overlay_name="zhongmei_chinese",
            ),
            "题内标题",
        )
        self.assertEqual(
            core_parser.describe_ignored_inline_obstacle(
                "",
                overlay_name="zhongmei_chinese",
            ),
            "图片/空白段",
        )


if __name__ == "__main__":
    unittest.main()
