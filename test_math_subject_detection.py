import os
import sys
import unittest


sys.path.insert(0, os.path.abspath("墨痕快刀"))

import config
from core_parser import detect_subject


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


class MathSubjectDetectionTests(unittest.TestCase):
    def test_detect_subject_marks_math_filename_as_science(self):
        original_overlay = getattr(config, "CURRENT_SUBJECT_OVERLAY", None)
        try:
            doc = MockDoc("八年级数学第6次限时训练.docx", "一次函数 二元一次方程组 直线交点")
            subject = detect_subject(doc)
            self.assertEqual(subject["name"], "理科")
            self.assertIsNone(config.CURRENT_SUBJECT_OVERLAY)
        finally:
            config.CURRENT_SUBJECT_OVERLAY = original_overlay


if __name__ == "__main__":
    unittest.main()
