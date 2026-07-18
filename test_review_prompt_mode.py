import os
import sys
import unittest


sys.path.insert(0, os.path.abspath("墨痕快刀"))

import config
from core_parser import (
    format_deferred_review_items,
    should_ignore_inline_obstacle,
    should_pause_for_review,
    should_preserve_question_prefix,
    sync_subject_overlay,
)


class _Unit:
    def __init__(
        self,
        question_id,
        confidence,
        warnings,
        question_type="choice",
        node_type="STD",
        media_blocks=None,
        subquestions=None,
    ):
        self.question_id = question_id
        self.confidence = confidence
        self.warnings = warnings
        self.question_type = question_type
        self.node_type = node_type
        self.media_blocks = media_blocks or []
        self.subquestions = subquestions or []


class ReviewPromptModeTests(unittest.TestCase):
    def test_low_confidence_units_do_not_pause_when_auto_continue_enabled(self):
        original = getattr(config, "PAUSE_ON_LOW_CONFIDENCE", False)
        config.PAUSE_ON_LOW_CONFIDENCE = False
        try:
            unit = _Unit("6", 0.9, ["image_between_stem_and_options"])
            self.assertFalse(should_pause_for_review(unit))
        finally:
            config.PAUSE_ON_LOW_CONFIDENCE = original

    def test_deferred_review_items_use_chinese_summary(self):
        items = format_deferred_review_items(
            [
                _Unit("6", 0.9, ["image_between_stem_and_options"]),
                _Unit("7", 0.85, ["image_related_question"]),
            ]
        )
        self.assertEqual(
            items,
            [
                "第 6 题：题干和选项之间夹有图片或表格，边界可能受影响",
                "第 7 题：这是图片/图表相关题，题干依赖图片或图表内容",
            ],
        )

    def test_question_prefix_is_preserved_only_for_allowlisted_docs(self):
        self.assertTrue(should_preserve_question_prefix("限训5.docx"))
        self.assertTrue(should_preserve_question_prefix("限训11：5.7.docx"))
        self.assertFalse(should_preserve_question_prefix("普通地理练习.docx"))
        self.assertFalse(should_preserve_question_prefix("高一英语.docx"))

    def test_sync_subject_overlay_sets_geography_context(self):
        class _Range:
            def __init__(self, text):
                self.Text = text

            def __call__(self, *_args):
                return self

            @property
            def End(self):
                return len(self.Text)

        class _Doc:
            Name = "限训4.docx"

            def __init__(self, text):
                self.Range = _Range(text)

        doc = _Doc("下图为世界部分区域海陆分布示意图，据图完成6-8小题。Y6.亚马孙河是世界上流量最大的河流。")
        original = getattr(config, "CURRENT_SUBJECT_OVERLAY", None)
        try:
            overlay_name = sync_subject_overlay(doc, "文科")
            self.assertEqual(overlay_name, "geography")
            self.assertEqual(config.CURRENT_SUBJECT_OVERLAY, "geography")
        finally:
            config.CURRENT_SUBJECT_OVERLAY = original

    def test_subjective_question_ignores_internal_bracket_heading_obstacle(self):
        unit = _Unit("14", 1.0, [], question_type="subjective", node_type="VIP")
        self.assertTrue(should_ignore_inline_obstacle(unit, "【独特的自然环境】", r"^【.*】$"))

    def test_vip_material_choice_question_ignores_internal_bracket_heading_obstacle(self):
        unit = _Unit("14", 1.0, [], question_type="material_choice", node_type="VIP")
        self.assertTrue(should_ignore_inline_obstacle(unit, "【独特的自然环境】", r"^【.*】$"))

    def test_choice_question_still_respects_bracket_heading_obstacle(self):
        unit = _Unit("6", 1.0, [], question_type="choice", node_type="READING")
        self.assertFalse(should_ignore_inline_obstacle(unit, "【独特的自然环境】", r"^【.*】$"))

    def test_subjective_question_ignores_internal_decorative_image_with_subquestions(self):
        unit = _Unit(
            "12",
            0.95,
            ["image_related_question"],
            question_type="subjective",
            node_type="STD",
            media_blocks=[54],
            subquestions=["(1)弹簧弹力F弹的大小", "(2)B对A的摩擦力", "(3)B与地面间的动摩擦因数"],
        )
        self.assertTrue(
            should_ignore_inline_obstacle(unit, "", "Img_DecorativeHeader_130x37")
        )

    def test_subjective_question_without_image_context_keeps_decorative_header_blocking(self):
        unit = _Unit("12", 1.0, [], question_type="subjective", node_type="STD")
        self.assertFalse(
            should_ignore_inline_obstacle(unit, "", "Img_DecorativeHeader_130x37")
        )


if __name__ == "__main__":
    unittest.main()
