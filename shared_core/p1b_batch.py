from __future__ import annotations

import json
import shutil
import stat
from pathlib import Path
from typing import Any, Sequence

from shared_core.document_render import sha256_file


BATCH_SCHEMA_VERSION = "1.0"
FORMAL_DIRECTORIES = ("选必一活页", "选必二活页")


def _is_inside(path: Path, directory: Path) -> bool:
    try:
        path.relative_to(directory)
    except ValueError:
        return False
    return True


def discover_formal_docx(
    batch_root: str | Path,
    *,
    formal_directories: Sequence[str] = FORMAL_DIRECTORIES,
) -> list[Path]:
    root = Path(batch_root).resolve()
    if not root.is_dir():
        raise FileNotFoundError(root)
    sources = []
    for directory_name in formal_directories:
        directory = root / directory_name
        if not directory.is_dir():
            raise FileNotFoundError(f"缺少正式原题目录：{directory}")
        sources.extend(
            path.resolve()
            for path in directory.rglob("*.docx")
            if not path.name.startswith(("~$", ".~"))
        )
    return sorted(sources, key=lambda path: (path.parent.name, path.name))


def prepare_wps_readonly_batch(
    batch_root: str | Path,
    output_dir: str | Path,
    *,
    expected_count: int | None = None,
    formal_directories: Sequence[str] = FORMAL_DIRECTORIES,
) -> dict[str, Any]:
    root = Path(batch_root).resolve()
    target = Path(output_dir).resolve()
    if _is_inside(target, root):
        raise ValueError("P1b 批次副本不能写入原题批次目录")
    sources = discover_formal_docx(root, formal_directories=formal_directories)
    if expected_count is not None and len(sources) != expected_count:
        raise ValueError(f"正式 DOCX 数量异常：期望 {expected_count}，实际 {len(sources)}")
    target.mkdir(parents=True, exist_ok=True)
    documents = []
    used_copy_names: set[str] = set()
    for source in sources:
        source_hash_before = sha256_file(source)
        copy_name = f"{source.parent.name}__{source.name}"
        if copy_name in used_copy_names:
            raise ValueError(f"批次副本名称冲突：{copy_name}")
        used_copy_names.add(copy_name)
        copy_path = target / copy_name
        if copy_path.exists():
            if not copy_path.is_file() or sha256_file(copy_path) != source_hash_before:
                raise ValueError(f"目标已存在且与原题不同：{copy_path}")
        else:
            shutil.copy2(source, copy_path)
        copy_path.chmod(stat.S_IRUSR | stat.S_IRGRP | stat.S_IROTH)
        copy_hash = sha256_file(copy_path)
        source_hash_after = sha256_file(source)
        if source_hash_before != source_hash_after or copy_hash != source_hash_before:
            raise RuntimeError(f"批次副本哈希校验失败：{source.name}")
        documents.append(
            {
                "source_name": source.name,
                "source_path": source.as_posix(),
                "source_sha256": source_hash_before,
                "source_parent": source.parent.name,
                "copy_name": copy_name,
                "copy_path": copy_path.as_posix(),
                "copy_sha256": copy_hash,
                "source_hash_verified_unchanged": True,
                "copy_read_only": not bool(copy_path.stat().st_mode & stat.S_IWUSR),
            }
        )
    return {
        "schema_version": BATCH_SCHEMA_VERSION,
        "classification_mode": "wps_readonly_copy_preparation",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "source_root": root.as_posix(),
        "formal_directories": list(formal_directories),
        "document_count": len(documents),
        "documents": documents,
    }


def write_wps_readonly_batch(
    batch_root: str | Path,
    output_dir: str | Path,
    *,
    expected_count: int | None = None,
) -> Path:
    target = Path(output_dir).resolve()
    manifest = prepare_wps_readonly_batch(
        batch_root,
        target,
        expected_count=expected_count,
    )
    path = target / "BatchSourceManifest.json"
    temp_path = target / ".BatchSourceManifest.json.tmp"
    try:
        temp_path.write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        json.loads(temp_path.read_text(encoding="utf-8"))
        temp_path.replace(path)
    finally:
        temp_path.unlink(missing_ok=True)
    return path
