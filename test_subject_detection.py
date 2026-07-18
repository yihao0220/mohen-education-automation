import os
import re
import sys
import unittest


sys.path.insert(0, os.path.abspath("墨痕快刀"))

import config
from core_parser import _is_bracket_subquestion, _is_numbered_main_question, detect_subject
from shared_core.strategies import strip_question_noise_prefix


class MockRange:
    def __init__(self, text):
        self.Text = text

    def __call__(self, *args):
        return self

    @property
    def End(self):
        return len(self.Text)


class MockDoc:
    def __init__(self, name, text):
        self.Name = name
        self.Range = MockRange(text)


class SubjectDetectionTests(unittest.TestCase):
    def test_science_question_format_accepts_star_prefix(self):
        text = "★11.如图所示,将一个半径为R、质量为M的均匀大球"
        wps_text = "/11.如图所示,将一个半径为R、质量为M的均匀大球"

        self.assertTrue(
            any(re.match(pattern, text) for pattern in config.CONFIG_SCIENCE["formats"].values())
        )
        self.assertEqual(strip_question_noise_prefix(text), "11.如图所示,将一个半径为R、质量为M的均匀大球")
        self.assertTrue(
            any(re.match(pattern, wps_text) for pattern in config.CONFIG_SCIENCE["formats"].values())
        )
        self.assertEqual(strip_question_noise_prefix(wps_text), "11.如图所示,将一个半径为R、质量为M的均匀大球")
        self.assertTrue(_is_numbered_main_question(text))
        self.assertTrue(_is_numbered_main_question(wps_text))
        self.assertTrue(_is_bracket_subquestion("★(1)小球经过A点的速度"))
        self.assertTrue(_is_bracket_subquestion("/(1)小球经过A点的速度"))

    def test_detect_subject_marks_geography_content_as_arts(self):
        text = """Y1.下列四幅局部等高线地形图中，有可能发育河流的是（    ）
G3.在西藏发现鱼龙化石，说明（   ）
下图为世界部分区域海陆分布示意图，据图完成6-8小题。
Y6.亚马孙河是世界上流量最大的河流，刚果河是世界第二大河，两河都注入（   ）
G13.亚洲气候类型多样，各地的气候差异很大。读图，回答下列问题。
"""
        doc = MockDoc("限训4.docx", text)

        subject = detect_subject(doc)

        self.assertEqual(subject["name"], "文科")
        self.assertEqual(config.CURRENT_SUBJECT_OVERLAY, "geography")

    def test_detect_subject_marks_history_lesson_content_as_arts(self):
        text = """第10课 当代中国的法治与精神文明建设
【课程标准】
了解当代中国法治建设和社会主义精神文明建设的成就。
1．改革开放以来，中国特色社会主义法律体系不断完善。
2．社会主义精神文明建设推动社会风尚变化。
"""
        doc = MockDoc("第10课 当代中国的法治与精神文明建设 学生版.doc", text)

        subject = detect_subject(doc)

        self.assertEqual(subject["name"], "文科")
        self.assertEqual(config.CURRENT_SUBJECT_OVERLAY, "history")

    def test_arts_sections_include_history_work_sheet_boundaries(self):
        section_patterns = config.CONFIG_ARTS["sections"]

        self.assertTrue(any(re.match(pattern, "【作业单】") for pattern in section_patterns))
        self.assertTrue(any(re.match(pattern, "【学习单】") for pattern in section_patterns))


if __name__ == "__main__":
    unittest.main()
