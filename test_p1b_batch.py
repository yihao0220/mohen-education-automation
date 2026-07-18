from __future__ import annotations

import json
import stat
from pathlib import Path

import pytest

from shared_core.document_render import sha256_file
from shared_core.p1b_batch import discover_formal_docx, write_wps_readonly_batch


def _batch(tmp_path: Path) -> Path:
    root = tmp_path / "future-biology"
    for directory in ("选必一活页", "选必二活页"):
        (root / directory).mkdir(parents=True)
    (root / "选必一活页" / "a.docx").write_bytes(b"a")
    (root / "选必二活页" / "b.docx").write_bytes(b"b")
    (root / "选必一活页" / "~$a.docx").write_bytes(b"lock")
    (root / "选必二活页" / ".~b.docx").write_bytes(b"lock")
    answers = root / "选必一答案"
    answers.mkdir()
    (answers / "answer.docx").write_bytes(b"answer")
    return root


def test_prepare_batch_only_copies_formal_docs_read_only(tmp_path: Path) -> None:
    root = _batch(tmp_path)
    output = tmp_path / "outside" / "copies"
    sources = discover_formal_docx(root)
    before = {path: sha256_file(path) for path in sources}

    manifest_path = write_wps_readonly_batch(root, output, expected_count=2)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))

    assert manifest["document_count"] == 2
    assert manifest["automatic_rule_binding_enabled"] is False
    assert manifest["production_execution_enabled"] is False
    assert {item["source_name"] for item in manifest["documents"]} == {"a.docx", "b.docx"}
    assert all(item["copy_read_only"] for item in manifest["documents"])
    assert all(item["source_hash_verified_unchanged"] for item in manifest["documents"])
    assert all(not (Path(item["copy_path"]).stat().st_mode & stat.S_IWUSR) for item in manifest["documents"])
    assert {path: sha256_file(path) for path in sources} == before
    assert not list(output.glob("~$*.docx"))
    assert not list(output.glob(".~*.docx"))


def test_prepare_batch_rejects_source_pollution_count_mismatch_and_conflict(tmp_path: Path) -> None:
    root = _batch(tmp_path)
    with pytest.raises(ValueError, match="批次目录"):
        write_wps_readonly_batch(root, root / "artifacts", expected_count=2)
    with pytest.raises(ValueError, match="期望 3"):
        write_wps_readonly_batch(root, tmp_path / "outside", expected_count=3)

    output = tmp_path / "conflict"
    output.mkdir()
    (output / "选必一活页__a.docx").write_bytes(b"different")
    with pytest.raises(ValueError, match="与原题不同"):
        write_wps_readonly_batch(root, output, expected_count=2)
