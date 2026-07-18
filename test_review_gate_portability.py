from __future__ import annotations

import json
import os
from pathlib import Path
import shutil

from shared_core.review_gate import (
    derive_review_status_path,
    get_review_gate_result,
    initialize_review_status,
    update_review_status,
)


def _approve(path: Path) -> Path:
    initialize_review_status(path)
    update_review_status(path, status="approved", reviewer="system")
    return Path(derive_review_status_path(path))


def test_approved_status_survives_cross_machine_relocation_and_mtime_change(
    tmp_path: Path,
) -> None:
    mac_root = tmp_path / "mac" / "答案" / "已清洗"
    windows_root = tmp_path / "windows" / "答案" / "已清洗"
    mac_root.mkdir(parents=True)
    windows_root.mkdir(parents=True)
    source = mac_root / "第1课_已清洗.docx"
    source.write_bytes(b"portable-answer-content")
    source_status = _approve(source)

    relocated = windows_root / source.name
    relocated_status = windows_root / source_status.name
    shutil.copyfile(source, relocated)
    shutil.copyfile(source_status, relocated_status)
    os.utime(relocated, ns=(1_800_000_000_000_000_000,) * 2)

    result = get_review_gate_result(relocated)
    assert result["allowed"] is True
    assert result["status"] == "approved"


def test_same_size_content_change_is_stale_even_if_mtime_is_restored(
    tmp_path: Path,
) -> None:
    answer = tmp_path / "第1课_已清洗.docx"
    answer.write_bytes(b"original-bytes")
    _approve(answer)
    original_stat = answer.stat()

    answer.write_bytes(b"modified-bytes")
    assert answer.stat().st_size == original_stat.st_size
    os.utime(answer, ns=(original_stat.st_atime_ns, original_stat.st_mtime_ns))

    result = get_review_gate_result(answer)
    assert result["allowed"] is False
    assert result["status"] == "stale"


def test_legacy_path_and_mtime_signature_cannot_be_silently_migrated(
    tmp_path: Path,
) -> None:
    answer = tmp_path / "第1课_已清洗.docx"
    answer.write_bytes(b"legacy-answer")
    status_path = Path(derive_review_status_path(answer))
    stat = answer.stat()
    status_path.write_text(
        json.dumps(
            {
                "status": "approved",
                "answer_signature": {
                    "path": str(answer.resolve()),
                    "size": stat.st_size,
                    "mtime_ns": stat.st_mtime_ns,
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    result = get_review_gate_result(answer)
    assert result["allowed"] is False
    assert result["status"] == "stale"
    assert "旧版文件签名" in result["reason"]
