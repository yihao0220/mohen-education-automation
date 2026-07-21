from __future__ import annotations

import argparse
from hashlib import sha256
import os
from pathlib import Path
import sys
import tempfile
import unicodedata
from zipfile import ZIP_DEFLATED, ZipFile


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.cli_output import configure_utf8_stdio
from shared_core.review_gate import derive_review_status_path, get_review_gate_result
from tools.refresh_guanmei_biology_review_status import (
    DEFAULT_PROJECT_ROOT,
    discover_question_answer_pairs,
)


DEFAULT_ARCHIVE_NAME = "guanmei_biology_cleaned_answers_windows_v2.zip"


def _sha256_file(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _collect_package_entries(
    answer_root: Path,
    pairs: list[tuple[Path, Path]],
) -> list[tuple[Path, str]]:
    entries: list[tuple[Path, str]] = []
    seen: set[str] = set()
    for _, answer_path in pairs:
        gate = get_review_gate_result(answer_path)
        if not gate.get("allowed"):
            raise ValueError(
                f"打包前门禁未通过：{answer_path.name}：{gate.get('reason', '未知原因')}"
            )
        status_path = Path(derive_review_status_path(answer_path))
        for source_path in (answer_path, status_path):
            relative = source_path.relative_to(answer_root)
            archive_name = unicodedata.normalize(
                "NFC",
                (Path("答案") / relative).as_posix(),
            )
            if archive_name in seen:
                raise ValueError(f"NFC归一化后路径重复：{archive_name}")
            seen.add(archive_name)
            entries.append((source_path, archive_name))
    return entries


def _validate_package(
    package_path: Path,
    entries: list[tuple[Path, str]],
    expected_document_count: int,
) -> None:
    with ZipFile(package_path) as package:
        if package.testzip() is not None:
            raise ValueError("Windows压缩包CRC校验失败")
        infos = package.infolist()
        if len(infos) != expected_document_count * 2:
            raise ValueError(
                f"压缩包文件数异常：预期 {expected_document_count * 2}，实际 {len(infos)}"
            )
        if any(
            not (info.flag_bits & 0x800)
            for info in infos
            if not info.filename.isascii()
        ):
            raise ValueError("压缩包中文路径未全部标记为UTF-8")
        if any(
            "__MACOSX" in info.filename
            or ".DS_Store" in info.filename
            or Path(info.filename).name.startswith("._")
            for info in infos
        ):
            raise ValueError("压缩包仍含Mac元数据")
        for source_path, archive_name in entries:
            if sha256(package.read(archive_name)).hexdigest() != _sha256_file(source_path):
                raise ValueError(f"压缩包内容校验失败：{archive_name}")

    with tempfile.TemporaryDirectory(prefix="guanmei_windows_extract_") as temporary_dir:
        extraction_root = Path(temporary_dir)
        with ZipFile(package_path) as package:
            package.extractall(extraction_root)
        extracted_answers = sorted((extraction_root / "答案").rglob("*_已清洗.docx"))
        if len(extracted_answers) != expected_document_count:
            raise ValueError("模拟Windows解压后的答案数量异常")
        for answer_path in extracted_answers:
            gate = get_review_gate_result(answer_path)
            if not gate.get("allowed"):
                raise ValueError(
                    f"模拟Windows解压后门禁未通过：{answer_path.name}：{gate.get('reason')}"
                )


def package_guanmei_biology_answers(
    project_root: str | Path,
    *,
    answer_root: str | Path | None = None,
    output_path: str | Path | None = None,
    expected_count: int | None = 27,
    overwrite: bool = False,
) -> tuple[Path, int, int]:
    root = Path(project_root)
    resolved_answer_root, pairs = discover_question_answer_pairs(
        root,
        answer_root=answer_root,
        expected_count=expected_count,
    )
    target = Path(output_path) if output_path else root / DEFAULT_ARCHIVE_NAME
    if target.exists() and not overwrite:
        raise FileExistsError(f"压缩包已存在，未覆盖：{target}")

    entries = _collect_package_entries(resolved_answer_root, pairs)
    target.parent.mkdir(parents=True, exist_ok=True)
    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=f"{target.stem}_",
        suffix=".zip",
        dir=target.parent,
    )
    os.close(file_descriptor)
    temporary_path = Path(temporary_name)
    try:
        with ZipFile(
            temporary_path,
            "w",
            compression=ZIP_DEFLATED,
            compresslevel=6,
        ) as package:
            for source_path, archive_name in entries:
                package.write(source_path, archive_name)
        _validate_package(temporary_path, entries, len(pairs))
        temporary_path.replace(target)
    finally:
        temporary_path.unlink(missing_ok=True)
    return target, len(pairs), len(entries)


def main() -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(
        description="生成莞美高二生物答案与审核状态的Windows兼容ZIP"
    )
    parser.add_argument("project_root", nargs="?", type=Path, default=DEFAULT_PROJECT_ROOT)
    parser.add_argument("--answer-root", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--expected-count", type=int, default=27)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    archive_path, document_count, file_count = package_guanmei_biology_answers(
        args.project_root,
        answer_root=args.answer_root,
        output_path=args.output,
        expected_count=args.expected_count,
        overwrite=args.overwrite,
    )
    print(f"已清洗答案：{document_count}")
    print(f"打包文件：{file_count}")
    print(f"Windows兼容压缩包：{archive_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
