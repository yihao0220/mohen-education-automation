from pathlib import Path

from shared_core.subject_overlay import should_use_native_table_input_for_context


def test_all_documents_use_native_table_input():
    assert should_use_native_table_input_for_context("任意学校语文试卷.docx") is True
    assert should_use_native_table_input_for_context("任意学校数学试卷.docx", None) is True


def test_question_input_runtime_has_no_table_to_text_fallback():
    source = Path("墨痕快刀/core_parser.py").read_text(encoding="utf-8")

    assert "[表格转文本]" not in source
    assert "_select_text_via_temp_document" not in source

