from __future__ import annotations

import argparse
from dataclasses import dataclass
from hashlib import sha256
import json
from pathlib import Path
import re
import sys
import tempfile
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile

from docx import Document
from lxml import etree as ET

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.cli_output import configure_utf8_stdio


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
DOCUMENT_XML = "word/document.xml"
ANSWER_DIRECTORIES = ("选必一答案", "选必二答案")
MARKER_TEXT = "课时对点练"
VISIBLE_MARKER_PATTERN = re.compile(r"^\s*课时对点练\s*")
SCORE_SUFFIX_PATTERN = re.compile(r"^[\[【]分值[：:].+[\]】]$")
ANSWER_PATTERN = re.compile(r"^\s*答案(?:\s|[：:])")
ANALYSIS_PATTERN = re.compile(r"^\s*(?:解析|【解析】)(?:\s|[：:])")
MANIFEST_FILENAME = "FutureBiologyAnswerTrimManifest.json"

ET.register_namespace("w", W_NS)


@dataclass(frozen=True)
class AnswerDocumentScan:
    source_path: Path
    relative_path: Path
    source_sha256: str
    matched: bool
    marker_kind: str | None
    marker_child_index: int | None
    body_child_count: int
    answer_count: int
    analysis_count: int


@dataclass(frozen=True)
class TrimResult:
    source_path: Path
    output_path: Path
    marker_kind: str
    removed_body_child_count: int
    answer_count: int
    analysis_count: int
    source_sha256: str
    output_sha256: str


@dataclass(frozen=True)
class _MarkerLocation:
    child_index: int
    kind: str
    visible_prefix_length: int


def _sha256_file(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as file_obj:
        for chunk in iter(lambda: file_obj.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _is_lock_file(path: Path) -> bool:
    return path.name.startswith(("~$", ".~", "~"))


def discover_answer_documents(source_root: str | Path) -> list[Path]:
    root = Path(source_root)
    if not root.is_dir():
        raise FileNotFoundError(f"未找到未来高二生物目录: {root}")

    missing_directories = [name for name in ANSWER_DIRECTORIES if not (root / name).is_dir()]
    if missing_directories:
        raise FileNotFoundError(
            "缺少答案目录: " + "、".join(str(root / name) for name in missing_directories)
        )

    sources: list[Path] = []
    for directory_name in ANSWER_DIRECTORIES:
        sources.extend(
            path
            for path in sorted((root / directory_name).glob("*.docx"))
            if not _is_lock_file(path)
        )
    return sources


def _read_document_root(path: Path) -> ET._Element:
    try:
        with ZipFile(path, "r") as package:
            bad_member = package.testzip()
            if bad_member:
                raise ValueError(f"DOCX 压缩包损坏: {path.name} -> {bad_member}")
            return ET.fromstring(package.read(DOCUMENT_XML))
    except BadZipFile as exc:
        raise ValueError(f"不是有效的 DOCX 文件: {path}") from exc
    except KeyError as exc:
        raise ValueError(f"DOCX 缺少 {DOCUMENT_XML}: {path}") from exc


def _node_text(node: ET._Element) -> str:
    return "".join(text_node.text or "" for text_node in node.iter(f"{W}t"))


def _node_instruction_text(node: ET._Element) -> str:
    return "".join(text_node.text or "" for text_node in node.iter(f"{W}instrText"))


def _find_marker(body: ET._Element, source_path: Path) -> _MarkerLocation | None:
    hits: list[_MarkerLocation] = []
    for child_index, child in enumerate(body):
        visible_text = _node_text(child)
        instruction_text = _node_instruction_text(child)
        visible_match = VISIBLE_MARKER_PATTERN.match(visible_text)
        field_match = MARKER_TEXT in instruction_text
        if not visible_match and not field_match:
            continue

        if child.tag != f"{W}p":
            raise ValueError(
                f"{source_path.name} 的“{MARKER_TEXT}”必须位于正文顶层段落，"
                "不能位于表格或其他嵌套结构"
            )
        if visible_match and field_match:
            raise ValueError(f"{source_path.name} 的“{MARKER_TEXT}”同时命中文本和图片字段")

        if visible_match:
            suffix = visible_text[visible_match.end() :].strip()
            if suffix and not SCORE_SUFFIX_PATTERN.match(suffix):
                raise ValueError(
                    f"{source_path.name} 的“{MARKER_TEXT}”同行存在未知内容: {suffix[:80]}"
                )
            hits.append(
                _MarkerLocation(
                    child_index=child_index,
                    kind="visible_text",
                    visible_prefix_length=visible_match.end(),
                )
            )
            continue

        if visible_text.strip():
            raise ValueError(
                f"{source_path.name} 的“{MARKER_TEXT}”图片段落同时存在未知可见文字"
            )
        hits.append(
            _MarkerLocation(
                child_index=child_index,
                kind="field_image",
                visible_prefix_length=0,
            )
        )

    if len(hits) > 1:
        raise ValueError(f"{source_path.name} 的“{MARKER_TEXT}”标记数量异常: {len(hits)}")
    return hits[0] if hits else None


def _retained_elements(body: ET._Element, marker: _MarkerLocation) -> list[ET._Element]:
    children = list(body)
    if marker.kind == "field_image":
        return children[marker.child_index + 1 :]
    return children[marker.child_index:]


def _count_answer_analysis(elements: list[ET._Element]) -> tuple[int, int]:
    answers = 0
    analyses = 0
    for element in elements:
        for paragraph in element.iter(f"{W}p"):
            text = _node_text(paragraph)
            answers += bool(ANSWER_PATTERN.match(text))
            analyses += bool(ANALYSIS_PATTERN.match(text))
    return answers, analyses


def scan_answer_document(
    source_path: str | Path,
    *,
    source_root: str | Path | None = None,
) -> AnswerDocumentScan:
    source = Path(source_path)
    if not source.is_file():
        raise FileNotFoundError(f"未找到答案文档: {source}")

    root = _read_document_root(source)
    body = root.find(f"{W}body")
    if body is None:
        raise ValueError(f"文档缺少正文节点: {source.name}")

    marker = _find_marker(body, source)
    relative_path = source.relative_to(Path(source_root)) if source_root else Path(source.name)
    if marker is None:
        return AnswerDocumentScan(
            source_path=source,
            relative_path=relative_path,
            source_sha256=_sha256_file(source),
            matched=False,
            marker_kind=None,
            marker_child_index=None,
            body_child_count=len(body),
            answer_count=0,
            analysis_count=0,
        )

    retained = [element for element in _retained_elements(body, marker) if element.tag != f"{W}sectPr"]
    if not retained:
        raise ValueError(f"{source.name} 的“{MARKER_TEXT}”后没有可保留内容")
    answers, analyses = _count_answer_analysis(retained)
    if answers == 0 or analyses == 0:
        raise ValueError(
            f"{source.name} 的“{MARKER_TEXT}”后未同时识别到答案和解析: "
            f"答案 {answers}，解析 {analyses}"
        )

    return AnswerDocumentScan(
        source_path=source,
        relative_path=relative_path,
        source_sha256=_sha256_file(source),
        matched=True,
        marker_kind=marker.kind,
        marker_child_index=marker.child_index,
        body_child_count=len(body),
        answer_count=answers,
        analysis_count=analyses,
    )


def scan_answer_batch(source_root: str | Path) -> list[AnswerDocumentScan]:
    root = Path(source_root)
    return [
        scan_answer_document(source, source_root=root)
        for source in discover_answer_documents(root)
    ]


def _set_text(node: ET._Element, value: str) -> None:
    node.text = value
    if value.startswith(" ") or value.endswith(" "):
        node.set(XML_SPACE, "preserve")
    else:
        node.attrib.pop(XML_SPACE, None)


def _remove_visible_prefix(paragraph: ET._Element, prefix_length: int) -> None:
    text_nodes = list(paragraph.iter(f"{W}t"))
    if not text_nodes:
        raise ValueError("可见标记段落没有文本节点")

    cursor = 0
    for node in text_nodes:
        value = node.text or ""
        start = cursor
        end = start + len(value)
        cursor = end
        if end <= prefix_length:
            _set_text(node, "")
        elif start < prefix_length:
            _set_text(node, value[prefix_length - start :])


def _write_docx_with_document_xml(
    source_path: Path,
    output_path: Path,
    document_xml: bytes,
    *,
    overwrite: bool,
) -> None:
    if source_path.resolve() == output_path.resolve():
        raise ValueError("截取输出不能覆盖原答案文档")
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"输出已存在，未覆盖: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile(
        prefix=f"{output_path.stem}_",
        suffix=".docx",
        dir=output_path.parent,
        delete=False,
    ) as temporary_file:
        temporary_path = Path(temporary_file.name)

    try:
        with ZipFile(source_path, "r") as source_zip, ZipFile(
            temporary_path,
            "w",
            compression=ZIP_DEFLATED,
        ) as target_zip:
            for info in source_zip.infolist():
                payload = document_xml if info.filename == DOCUMENT_XML else source_zip.read(info.filename)
                target_zip.writestr(info, payload)

        with ZipFile(temporary_path, "r") as check_zip:
            bad_member = check_zip.testzip()
            if bad_member:
                raise ValueError(f"截取后的 DOCX 压缩包损坏: {bad_member}")
        Document(temporary_path)
        temporary_path.replace(output_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def trim_answer_document(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
) -> TrimResult:
    source = Path(source_path)
    output = Path(output_path)
    scan = scan_answer_document(source)
    if not scan.matched:
        raise ValueError(f"{source.name} 未识别到“{MARKER_TEXT}”，不会截取")

    root = _read_document_root(source)
    body = root.find(f"{W}body")
    if body is None:
        raise ValueError(f"文档缺少正文节点: {source.name}")
    marker = _find_marker(body, source)
    if marker is None:
        raise ValueError(f"{source.name} 的标记在写入前消失")

    for child in list(body)[: marker.child_index]:
        body.remove(child)

    marker_element = list(body)[0]
    if marker.kind == "field_image":
        body.remove(marker_element)
        removed_count = marker.child_index + 1
    else:
        _remove_visible_prefix(marker_element, marker.visible_prefix_length)
        removed_count = marker.child_index
        if not _node_text(marker_element).strip() and not _node_instruction_text(marker_element).strip():
            body.remove(marker_element)
            removed_count += 1

    document_xml = ET.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
        standalone=True,
    )
    _write_docx_with_document_xml(
        source,
        output,
        document_xml,
        overwrite=overwrite,
    )

    if _sha256_file(source) != scan.source_sha256:
        output.unlink(missing_ok=True)
        raise RuntimeError(f"原答案文档在截取过程中发生变化: {source}")

    output_scan_root = _read_document_root(output)
    output_body = output_scan_root.find(f"{W}body")
    if output_body is None or _find_marker(output_body, output) is not None:
        output.unlink(missing_ok=True)
        raise RuntimeError(f"截取结果仍包含“{MARKER_TEXT}”: {output}")
    answers, analyses = _count_answer_analysis(list(output_body))
    if (answers, analyses) != (scan.answer_count, scan.analysis_count):
        output.unlink(missing_ok=True)
        raise RuntimeError(
            f"截取结果答案/解析数量变化: {source.name} "
            f"{scan.answer_count}/{scan.analysis_count} -> {answers}/{analyses}"
        )

    return TrimResult(
        source_path=source,
        output_path=output,
        marker_kind=scan.marker_kind or "",
        removed_body_child_count=removed_count,
        answer_count=answers,
        analysis_count=analyses,
        source_sha256=scan.source_sha256,
        output_sha256=_sha256_file(output),
    )


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def _validate_output_directory(source_root: Path, output_dir: Path) -> None:
    resolved_output = output_dir.resolve()
    for directory_name in ANSWER_DIRECTORIES:
        answer_dir = (source_root / directory_name).resolve()
        if _is_relative_to(resolved_output, answer_dir):
            raise ValueError(f"输出目录不能放进原答案目录: {output_dir}")


def trim_answer_batch(
    source_root: str | Path,
    *,
    output_dir: str | Path | None = None,
    expected_source_count: int | None = None,
    expected_matched_count: int | None = None,
    overwrite: bool = False,
) -> Path:
    root = Path(source_root)
    output = Path(output_dir) if output_dir else root / "答案" / "按课时截取"
    _validate_output_directory(root, output)

    scans = scan_answer_batch(root)
    matched = [scan for scan in scans if scan.matched]
    if expected_source_count is not None and len(scans) != expected_source_count:
        raise ValueError(f"答案文档数量异常: 期望 {expected_source_count}，实际 {len(scans)}")
    if expected_matched_count is not None and len(matched) != expected_matched_count:
        raise ValueError(
            f"命中“{MARKER_TEXT}”的文档数量异常: "
            f"期望 {expected_matched_count}，实际 {len(matched)}"
        )
    if not matched:
        raise ValueError(f"整批未发现“{MARKER_TEXT}”标记")

    destinations = {scan.source_path: output / scan.relative_path for scan in matched}
    existing = [destination for destination in destinations.values() if destination.exists()]
    if existing and not overwrite:
        raise FileExistsError(f"已有 {len(existing)} 份输出，未覆盖: {existing[0]}")

    results = [
        trim_answer_document(source, destination, overwrite=overwrite)
        for source, destination in destinations.items()
    ]
    if any(_sha256_file(scan.source_path) != scan.source_sha256 for scan in scans):
        raise RuntimeError("整批处理后发现原答案文档哈希变化")

    output.mkdir(parents=True, exist_ok=True)
    manifest = {
        "schema_version": "1.0",
        "project": "未来-高二-生物",
        "rule": "保留课时对点练标记之后的内容",
        "source_root": str(root.resolve()),
        "output_dir": str(output.resolve()),
        "source_document_count": len(scans),
        "matched_document_count": len(matched),
        "skipped_document_count": len(scans) - len(matched),
        "original_sources_verified_unchanged": True,
        "documents": [],
    }
    results_by_source = {result.source_path: result for result in results}
    for scan in scans:
        item = {
            "relative_path": scan.relative_path.as_posix(),
            "source_sha256": scan.source_sha256,
            "status": "trimmed" if scan.matched else "skipped_no_marker",
            "marker_kind": scan.marker_kind,
            "answer_count": scan.answer_count,
            "analysis_count": scan.analysis_count,
        }
        if scan.matched:
            result = results_by_source[scan.source_path]
            item.update(
                {
                    "output_path": str(result.output_path.resolve()),
                    "output_sha256": result.output_sha256,
                    "removed_body_child_count": result.removed_body_child_count,
                }
            )
        manifest["documents"].append(item)

    manifest_path = output / MANIFEST_FILENAME
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return manifest_path


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="截取未来高二生物答案中“课时对点练”之后的内容，不覆盖原答案。"
    )
    parser.add_argument("source_root", help="未来-高二-生物项目目录")
    parser.add_argument(
        "--output-dir",
        help="输出目录，默认使用 项目目录/答案/按课时截取",
    )
    parser.add_argument("--expected-source-count", type=int, help="预期有效答案 DOCX 数量")
    parser.add_argument("--expected-matched-count", type=int, help="预期命中标记的 DOCX 数量")
    parser.add_argument("--dry-run", action="store_true", help="只做整批预检，不生成 DOCX")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已有截取结果")
    return parser


def main() -> None:
    configure_utf8_stdio()
    parser = build_arg_parser()
    args = parser.parse_args()
    root = Path(args.source_root)

    if args.dry_run:
        scans = scan_answer_batch(root)
        matched = [scan for scan in scans if scan.matched]
        if args.expected_source_count is not None and len(scans) != args.expected_source_count:
            parser.error(f"答案文档数量异常: 期望 {args.expected_source_count}，实际 {len(scans)}")
        if args.expected_matched_count is not None and len(matched) != args.expected_matched_count:
            parser.error(
                f"命中标记数量异常: 期望 {args.expected_matched_count}，实际 {len(matched)}"
            )
        field_count = sum(scan.marker_kind == "field_image" for scan in matched)
        text_count = sum(scan.marker_kind == "visible_text" for scan in matched)
        print(
            f"预检通过：有效答案 {len(scans)} 份，命中 {len(matched)} 份，"
            f"图片字段 {field_count} 份，可见文字 {text_count} 份，"
            f"跳过其他文档族 {len(scans) - len(matched)} 份。"
        )
        return

    manifest_path = trim_answer_batch(
        root,
        output_dir=args.output_dir,
        expected_source_count=args.expected_source_count,
        expected_matched_count=args.expected_matched_count,
        overwrite=args.overwrite,
    )
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    print(
        f"截取完成：生成 {manifest['matched_document_count']} 份，"
        f"跳过 {manifest['skipped_document_count']} 份其他文档族。"
    )
    print(f"批次清单：{manifest_path}")


if __name__ == "__main__":
    main()
