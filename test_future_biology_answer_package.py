from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

from shared_core.review_gate import (
    derive_review_status_path,
    get_review_gate_result,
    initialize_review_status,
    update_review_status,
)
from tools.package_future_biology_answers import package_future_biology_answers


def test_package_uses_utf8_names_and_portable_review_status(tmp_path: Path) -> None:
    project_root = tmp_path / "未来-高二-生物"
    cleaned_root = project_root / "答案" / "已清洗" / "选必一答案"
    report_root = project_root / "答案" / "审核清单" / "选必一答案"
    cleaned_root.mkdir(parents=True)
    report_root.mkdir(parents=True)
    answer_path = cleaned_root / "第1课_已清洗.docx"
    answer_path.write_bytes(b"portable-docx-content")
    initialize_review_status(answer_path)
    update_review_status(answer_path, status="approved", reviewer="system")
    (report_root / "第1课_已清洗_审核清单.md").write_text(
        "# 通过",
        encoding="utf-8",
    )
    (project_root / "答案" / "已清洗" / ".DS_Store").write_bytes(b"mac")

    archive_path, document_count, file_count = package_future_biology_answers(
        project_root,
        expected_docx_count=1,
    )

    assert document_count == 1
    assert file_count == 3
    assert archive_path.name == "future_biology_cleaned_answers_windows.zip"
    with ZipFile(archive_path) as package:
        infos = package.infolist()
        assert package.testzip() is None
        assert all(info.flag_bits & 0x800 for info in infos if not info.filename.isascii())
        assert not any(".DS_Store" in info.filename for info in infos)
        windows_root = tmp_path / "windows"
        package.extractall(windows_root)

    relocated = windows_root / answer_path.relative_to(project_root)
    relocated_status = Path(derive_review_status_path(relocated))
    assert relocated.is_file()
    assert relocated_status.is_file()
    assert get_review_gate_result(relocated)["allowed"] is True
