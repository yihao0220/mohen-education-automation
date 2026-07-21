from __future__ import annotations

from hashlib import sha256
import os
from pathlib import Path
import unittest

from shared_core.models import DocNode
from shared_core.question_core import (
    build_question_units_from_docx,
    build_question_units_from_nodes,
    scan_docx_nodes,
)
from shared_core.subject_overlay import (
    classify_inter_question_media_boundary_for_context,
    classify_media_hashes_for_context,
    detect_subject_overlay,
)


GUANMEI_EXERCISE_LABEL_HASH = (
    "9cf693c46d970beea342457d6e13783cb9cc45c229f988a725f1f7e387f52463"
)
REAL_SAMPLE = Path(
    os.environ.get(
        "MOHEN_GUANMEI_BIOLOGY_SAMPLE",
        Path.home()
        / "Documents"
        / "墨痕教育"
        / "莞美-高二-生物"
        / "第1章　人体的内环境与稳态"
        / "第1节　细胞的生活环境-排版终稿.docx",
    )
)


def _sha256_file(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class GuanmeiBiologyQuestionInputTest(unittest.TestCase):
    def test_overlay_is_scoped_to_final_layout_documents(self) -> None:
        sample_text = "细胞通过内环境进行物质交换，组织液和血浆参与稳态调节"

        self.assertEqual(
            detect_subject_overlay(
                "第1节　细胞的生活环境-排版终稿.docx",
                sample_text,
                base_subject="理科",
            ),
            "guanmei_biology",
        )
        self.assertIsNone(
            detect_subject_overlay(
                "第1节　细胞的生活环境.docx",
                sample_text,
                base_subject="理科",
            )
        )
        self.assertIsNone(
            detect_subject_overlay(
                "其他学科-排版终稿.docx",
                "只有一个细胞关键词",
                base_subject="理科",
            )
        )

    def test_confirmed_exercise_label_is_excluded_between_questions(self) -> None:
        nodes = [
            DocNode(index=1, text="9．上一道题"),
            DocNode(index=2, text="A．选项A"),
            DocNode(index=3, text="D．选项D"),
            DocNode(
                index=4,
                text="",
                has_inline_media=True,
                metadata={"media_sha256": [GUANMEI_EXERCISE_LABEL_HASH]},
            ),
            DocNode(index=5, text="10．下一道题"),
            DocNode(index=6, text="题内图片", has_inline_media=True),
        ]

        units = build_question_units_from_nodes(
            "第1节　细胞的生活环境-排版终稿.docx",
            "理科",
            nodes,
            overlay_name="guanmei_biology",
        )

        self.assertEqual(
            classify_media_hashes_for_context(
                [GUANMEI_EXERCISE_LABEL_HASH], "guanmei_biology"
            ),
            "exercise_label",
        )
        self.assertEqual([unit.question_id for unit in units], ["9", "10"])
        self.assertEqual([unit.source_span for unit in units], [(1, 3), (5, 6)])
        self.assertEqual(units[0].media_blocks, [])
        self.assertEqual(units[1].media_blocks, [6])

    def test_wps_context_excludes_only_blank_header_between_questions(self) -> None:
        obstacle = "Img_DecorativeHeader_170x42"

        self.assertEqual(
            classify_inter_question_media_boundary_for_context(
                current_text="",
                next_text="10．下一道题",
                obstacle_type=obstacle,
                overlay_name="guanmei_biology",
            ),
            "exercise_label",
        )
        self.assertIsNone(
            classify_inter_question_media_boundary_for_context(
                current_text="如图所示",
                next_text="10．下一道题",
                obstacle_type=obstacle,
                overlay_name="guanmei_biology",
            )
        )
        self.assertIsNone(
            classify_inter_question_media_boundary_for_context(
                current_text="",
                next_text="题目续行",
                obstacle_type=obstacle,
                overlay_name="guanmei_biology",
            )
        )

    @unittest.skipUnless(REAL_SAMPLE.is_file(), "当前机器没有莞美高二生物真实样本")
    def test_real_sample_excludes_label_but_keeps_question_figures(self) -> None:
        source_hash = _sha256_file(REAL_SAMPLE)
        nodes = scan_docx_nodes(REAL_SAMPLE)
        sample_text = " ".join(node.text for node in nodes[:20])
        self.assertEqual(
            detect_subject_overlay(
                REAL_SAMPLE.name,
                sample_text,
                base_subject="理科",
            ),
            "guanmei_biology",
        )

        target_node = next(
            node
            for node in nodes
            if GUANMEI_EXERCISE_LABEL_HASH in node.metadata.get("media_sha256", [])
        )
        units = build_question_units_from_docx(REAL_SAMPLE, grade_hint="高二")
        question_nine = next(unit for unit in units if unit.question_id == "9")
        question_ten = next(unit for unit in units if unit.question_id == "10")

        self.assertEqual(len(units), 13)
        self.assertEqual(question_nine.source_span, (36, 40))
        self.assertEqual(question_nine.media_blocks, [36])
        self.assertEqual(target_node.index, 41)
        self.assertNotIn(
            target_node.index,
            range(question_nine.source_span[0], question_nine.source_span[1] + 1),
        )
        self.assertEqual(question_ten.source_span, (42, 45))
        self.assertEqual(question_ten.media_blocks, [43])
        self.assertEqual(_sha256_file(REAL_SAMPLE), source_hash)


if __name__ == "__main__":
    unittest.main()
