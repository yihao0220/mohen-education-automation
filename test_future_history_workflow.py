from __future__ import annotations

import importlib.util
import shutil
import tempfile
import unittest
from pathlib import Path

from docx import Document
from shared_core import build_answer_units_from_docx


PROJECT_ROOT = Path(__file__).resolve().parent
HISTORY_ANSWER_DIR = Path(r"D:\墨痕教育题目\未来-高二-历史（王逸豪）\答案\学案·24-25选必一学案解析")


def _load_module(module_name: str, relative_path: str):
    module_path = PROJECT_ROOT / relative_path
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module


template = _load_module("template_future_history", "格式处理/格式模板库/template_future_history.py")


class FutureHistoryTemplateTests(unittest.TestCase):
    def test_reference_answer_section_ignores_front_questions(self):
        entries = template.parse_paragraph_texts(
            [
                "【作业单】",
                "1．（2022·高考真题）这是题干（     ）",
                "A．甲\tB．乙\tC．丙\tD．丁",
                "参考答案",
                "1．D",
                "【详解】第一题解析。",
                "2．C",
                "【详解】第二题解析。",
                "9．（1）印度：增强反抗殖民压迫的信心。",
                "（2）说明：天赋人权；人民主权。",
                "【详解】第九题解析。",
            ]
        )

        rendered = template.render_standard_lines(entries)

        self.assertEqual([entry["qnum"] for entry in entries], ["1", "2", "9"])
        self.assertIn("1．D", rendered)
        self.assertIn("9．", rendered)
        self.assertIn("答案：（1）印度：增强反抗殖民压迫的信心。 （2）说明：天赋人权；人民主权。", rendered)

    def test_inline_answer_markers_use_previous_question_number(self):
        entries = template.parse_paragraph_texts(
            [
                "【作业单】",
                "1．（2022·湖北·高考真题）如图是近代日记，箭头所指文字（     ）",
                "A．甲\tB．乙\tC．丙\tD．丁",
                "【答案】B",
                "【解析】第一题解析。故选B项。",
                "2．（2022·全国·高考真题）材料说明（       ）",
                "A．甲\tB．乙\tC．丙\tD．丁",
                "【答案】A",
                "【解析】第二题解析。",
            ]
        )

        rendered = template.render_standard_lines(entries)

        self.assertEqual([entry["qnum"] for entry in entries], ["1", "2"])
        self.assertEqual(rendered[:4], ["1．B", "解析：第一题解析。故选B项。", "2．A", "解析：第二题解析。"])

    def test_detail_only_answers_after_homework_are_collected(self):
        entries = template.parse_paragraph_texts(
            [
                "【作业单】",
                "1．（2019·天津·高考真题）题干内容",
                "A．甲\tB．乙\tC．丙\tD．丁",
                "1．D",
                "【详解】由材料可知，英国民主制度建设尚需完善。",
                "2．（2010·江苏·高考真题）题干内容",
                "A．甲\tB．乙\tC．丙\tD．丁",
                "2．C",
                "【详解】九品中正制导致门第影响仕途。",
            ]
        )

        rendered = template.render_standard_lines(entries)

        self.assertEqual([entry["qnum"] for entry in entries], ["1", "2"])
        self.assertEqual(rendered[0], "1．D")
        self.assertIn("英国民主制度建设尚需完善", rendered[1])

    def test_real_sample_parsing_counts(self):
        samples = {
            "第2课　西方国家古代和近代政治制度的演变.doc": 9,
            "第6课 西方的文官制度  教师版.doc": 9,
            "第15课 货币的使用与世界货币体系的形成 教师版.doc": 6,
        }
        missing = [name for name in samples if not (HISTORY_ANSWER_DIR / name).exists()]
        if missing:
            self.skipTest(f"缺少真实样本: {missing}")

        win32 = self._import_win32_or_skip()
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            for filename, expected_count in samples.items():
                paragraphs = self._read_word_paragraphs(word, HISTORY_ANSWER_DIR / filename)
                entries = template.parse_paragraph_texts(paragraphs)
                self.assertEqual(
                    len(entries),
                    expected_count,
                    f"{filename} 应提取 {expected_count} 道答案",
                )
        finally:
            word.Quit()

    def test_clean_real_doc_to_temp_docx_can_be_read_by_answer_core(self):
        sample_path = HISTORY_ANSWER_DIR / "第15课 货币的使用与世界货币体系的形成 教师版.doc"
        if not sample_path.exists():
            self.skipTest("缺少第15课真实样本")

        win32 = self._import_win32_or_skip()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_source = Path(temp_dir) / sample_path.name
            temp_output = Path(temp_dir) / "第15课历史答案_已清洗.docx"
            shutil.copy2(sample_path, temp_source)

            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                doc = word.Documents.Open(str(temp_source), ReadOnly=False, AddToRecentFiles=False)
                try:
                    self.assertTrue(template.clean_document(doc))
                    doc.SaveAs2(str(temp_output), FileFormat=16)
                finally:
                    doc.Close(False)
            finally:
                word.Quit()

            clean_doc = Document(temp_output)
            clean_lines = [paragraph.text.strip() for paragraph in clean_doc.paragraphs if paragraph.text.strip()]
            self.assertEqual(clean_lines[0], "1．B")
            self.assertTrue(clean_lines[1].startswith("解析："))

            units = build_answer_units_from_docx(temp_output)
            self.assertEqual([unit.question_id for unit in units], ["1", "2", "3", "4", "5", "6"])
            self.assertEqual(units[0].answer_items[0].text, "B")
            self.assertTrue(units[5].analysis_items)

    def test_user_markdown_sample_direct_subjective_answers(self):
        sample_path = Path(r"D:\墨痕教育题目\未来-高二-历史（王逸豪）\选必一\样版.md")
        if not sample_path.exists():
            self.skipTest("缺少用户给定样版.md")

        entries = template.parse_paragraph_texts(sample_path.read_text(encoding="utf-8").splitlines())
        rendered = template.render_standard_lines(entries)

        self.assertEqual([entry["qnum"] for entry in entries], [str(i) for i in range(1, 14)])
        self.assertTrue(entries[0]["answer_lines"][0].startswith("雅典：国家权力掌握在"))
        self.assertEqual(entries[4]["answer_lines"], ["D"])
        self.assertTrue(any(line.startswith("解析：根据材料可知") for line in rendered))

    def _import_win32_or_skip(self):
        try:
            import win32com.client as win32
        except Exception as exc:  # pragma: no cover - depends on local Windows COM
            self.skipTest(f"当前环境不可用 Word COM: {exc}")
        return win32

    def _read_word_paragraphs(self, word, path: Path) -> list[str]:
        doc = word.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        try:
            return [
                doc.Paragraphs(index).Range.Text
                for index in range(1, doc.Paragraphs.Count + 1)
            ]
        finally:
            doc.Close(False)


if __name__ == "__main__":
    unittest.main()
