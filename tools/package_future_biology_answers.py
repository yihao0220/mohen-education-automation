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
from shared_core.review_gate import get_review_gate_result


DEFAULT_ARCHIVE_NAME = "future_biology_cleaned_answers_windows.zip"
INCLUDED_DIRECTORIES = (Path("答案") / "已清洗", Path("答案") / "审核清单")


def _sha256_file(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _is_mac_metadata(path: Path) -> bool:
    return path.name == ".DS_Store" or path.name.startswith("._") or "__MACOSX" in path.parts


def _collect_files(project_root: Path) -> list[tuple[Path, str]]:
    entries: list[tuple[Path, str]] = []
    seen: set[str] = set()
    for relative_root in INCLUDED_DIRECTORIES:
        source_root = project_root / relative_root
        if not source_root.is_dir():
            raise FileNotFoundError(f"缺少打包目录: {source_root}")
        for source_path in sorted(source_root.rglob("*")):
            if not source_path.is_file() or _is_mac_metadata(source_path):
                continue
            archive_name = unicodedata.normalize(
                "NFC",
                source_path.relative_to(project_root).as_posix(),
            )
            if archive_name in seen:
                raise ValueError(f"NFC 归一化后路径重复: {archive_name}")
            seen.add(archive_name)
            entries.append((source_path, archive_name))
    return entries


def _validate_review_gates(
    project_root: Path,
    *,
    expected_docx_count: int | None,
) -> list[Path]:
    cleaned_root = project_root / "答案" / "已清洗"
    cleaned_documents = sorted(cleaned_root.rglob("*_已清洗.docx"))
    if expected_docx_count is not None and len(cleaned_documents) != expected_docx_count:
        raise ValueError(
            f"已清洗答案数量不符合预期: "
            f"预期 {expected_docx_count}，实际 {len(cleaned_documents)}"
        )
    if not cleaned_documents:
        raise ValueError("未找到已清洗答案")

    blocked: list[str] = []
    for document_path in cleaned_documents:
        gate = get_review_gate_result(document_path)
        if not gate.get("allowed"):
            blocked.append(f"{document_path.name}: {gate.get('reason', '未通过门禁')}")
    if blocked:
        raise ValueError("打包前审核门禁未通过:\n" + "\n".join(blocked))
    return cleaned_documents


def package_future_biology_answers(
    project_root: str | Path,
    *,
    output_path: str | Path | None = None,
    expected_docx_count: int | None = None,
    overwrite: bool = False,
) -> tuple[Path, int, int]:
    root = Path(project_root)
    if not root.is_dir():
        raise FileNotFoundError(f"未找到未来高二生物目录: {root}")
    target = Path(output_path) if output_path else root / DEFAULT_ARCHIVE_NAME
    if target.exists() and not overwrite:
        raise FileExistsError(f"压缩包已存在，未覆盖: {target}")

    cleaned_documents = _validate_review_gates(
        root,
        expected_docx_count=expected_docx_count,
    )
    entries = _collect_files(root)
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

        with ZipFile(temporary_path) as package:
            if package.testzip() is not None:
                raise ValueError("Windows 压缩包 CRC 校验失败")
            infos = package.infolist()
            if any(
                not (info.flag_bits & 0x800)
                for info in infos
                if not info.filename.isascii()
            ):
                raise ValueError("压缩包的中文路径未全部标记为 UTF-8")
            if any(
                "__MACOSX" in info.filename or ".DS_Store" in info.filename
                for info in infos
            ):
                raise ValueError("压缩包仍含 Mac 元数据")
            for source_path, archive_name in entries:
                if sha256(package.read(archive_name)).hexdigest() != _sha256_file(source_path):
                    raise ValueError(f"压缩包内容校验失败: {archive_name}")
        temporary_path.replace(target)
    finally:
        temporary_path.unlink(missing_ok=True)
    return target, len(cleaned_documents), len(entries)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="生成未来高二生物已清洗答案的 Windows UTF-8 兼容 ZIP。"
    )
    parser.add_argument("project_root", help="未来-高二-生物项目根目录")
    parser.add_argument("--output", help="输出 ZIP；默认放在项目根目录")
    parser.add_argument("--expected-docx-count", type=int)
    parser.add_argument("--overwrite", action="store_true", help="显式覆盖旧压缩包")
    return parser


def main() -> int:
    configure_utf8_stdio()
    args = _build_parser().parse_args()
    archive_path, document_count, file_count = package_future_biology_answers(
        args.project_root,
        output_path=args.output,
        expected_docx_count=args.expected_docx_count,
        overwrite=args.overwrite,
    )
    print(f"已清洗答案: {document_count}")
    print(f"打包文件: {file_count}")
    print(f"Windows 兼容压缩包: {archive_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
