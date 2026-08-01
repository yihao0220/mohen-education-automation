import io
import unittest
from contextlib import redirect_stdout

from 格式处理 import format_answers_deepseek
from 格式处理.格式模板库 import (
    template_a,
    template_b,
    template_c,
    template_chinese,
    template_d,
    template_e,
)


TEMPLATES = (template_a, template_b, template_c, template_d, template_e, template_chinese)


class FakeFont:
    pass


class FakeContent:
    Font = FakeFont()


class FakeDocument:
    Content = FakeContent()


class LowRiskCleanupTests(unittest.TestCase):
    def test_disabled_deepseek_entry_keeps_compatibility_result(self):
        output = io.StringIO()
        with redirect_stdout(output):
            result = format_answers_deepseek.run_llm_engine(object())

        self.assertFalse(result)
        self.assertIn("格式处理/main.py", output.getvalue())

    def test_templates_share_matching_and_garbage_helpers(self):
        for template in TEMPLATES:
            with self.subTest(template=template.__name__):
                self.assertTrue(template.is_garbage_line(""))
                self.assertEqual(template.match_score(None, cached_texts=["无匹配内容"]), 0)

        self.assertEqual(
            template_a.match_score(None, cached_texts=["2. D", "无匹配内容"]),
            0.5,
        )

    def test_shared_font_helper_preserves_template_output_format(self):
        document = FakeDocument()
        template_a.set_font_format(document)

        self.assertEqual(document.Content.Font.Size, 12)
        self.assertEqual(document.Content.Font.Color, 0)
        self.assertEqual(document.Content.Font.Name, "Times New Roman")
        self.assertEqual(document.Content.Font.NameFarEast, "宋体")
        self.assertFalse(document.Content.Font.Bold)
        self.assertFalse(document.Content.Font.Italic)


if __name__ == "__main__":
    unittest.main()
