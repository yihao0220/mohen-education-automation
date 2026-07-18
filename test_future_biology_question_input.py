from __future__ import annotations

from collections import Counter
from hashlib import sha256
import os
from pathlib import Path
import re

import pytest

from shared_core.document_preflight import build_preflight_bundle
from shared_core.models import DocNode
from shared_core.question_core import (
    build_question_units_from_docx,
    build_question_units_from_nodes,
    scan_docx_nodes,
)
from shared_core.subject_overlay import (
    classify_inter_question_media_boundary_for_context,
    classify_media_hashes_for_context,
    detect_subject_overlay,
)


REAL_BATCH = Path(
    os.environ.get(
        "MOHEN_FUTURE_BIOLOGY_DIR",
        Path(__file__).resolve().parents[2] / "墨痕教育" / "未来-高二-生物",
    )
)
FORMAL_DIRECTORIES = ("选必一活页", "选必二活页")
BIOLOGY_MEDIA_HASHES = {
    "document_banner": {
        "363013449da1db958e2f66f789d87f7eede936259833c21c1b645b8dc3953e71",
    },
    "exercise_label": {
        "f379746c07eb677523e68d9032053391d5982b4b933addfbe5a29de84d6d62de",
        "8c22c23eb8612cae71029cd75df419b262b028da00d66be51a0829c22004773a",
        "169289e4b5356a198f95facc3778da4e755b303d99f3b6155332099d69ed9de5",
    },
}
GROUP_HEADING_PATTERN = re.compile(r"^\s*题组[一二三四五六七八九十\d]+(?:\s|　|[：:]).*$")
SECTION_HEADING_PATTERN = re.compile(
    r"^\s*[一二三四五六七八九十]+[、．.]\s*(?:选择题|非选择题)(?:\s*[（(].*)?$"
)
TOP_LEVEL_QUESTION_PATTERN = re.compile(r"^\s*(\d+)\s*[．.、]")


def _real_question_documents() -> list[Path]:
    paths: list[Path] = []
    for directory_name in FORMAL_DIRECTORIES:
        directory = REAL_BATCH / directory_name
        paths.extend(
            path
            for path in sorted(directory.glob("*.docx"))
            if not path.name.startswith(("~$", ".~", "~"))
        )
    return paths


def _sha256_file(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _selected_texts(nodes: list[DocNode], units) -> list[str]:
    nodes_by_index = {node.index: node for node in nodes}
    selected: list[str] = []
    for unit in units:
        for index in range(unit.source_span[0], unit.source_span[1] + 1):
            node = nodes_by_index.get(index)
            if node and node.text.strip():
                selected.append(node.text.strip())
    return selected


def test_biology_overlay_excludes_structure_headings_without_splitting_subquestions() -> None:
    nodes = [
        DocNode(index=1, text="题组一　基础知识"),
        DocNode(index=2, text="1．第一题"),
        DocNode(index=3, text="A．选项A"),
        DocNode(index=4, text="D．选项D"),
        DocNode(index=5, text="题组二　能力提升"),
        DocNode(index=6, text="2．第二题"),
        DocNode(index=7, text="二、非选择题(本题包括1小题，共20分)"),
        DocNode(index=8, text="3．主观题"),
        DocNode(index=9, text="(1)第一小问"),
        DocNode(index=10, text="(2)第二小问"),
    ]

    units = build_question_units_from_nodes(
        "作业1　细胞生活的环境.docx",
        "理科",
        nodes,
        overlay_name="future_biology",
    )

    assert [unit.question_id for unit in units] == ["1", "2", "3"]
    assert [unit.source_span for unit in units] == [(2, 4), (6, 6), (8, 10)]
    assert len(units[2].subquestions) == 2
    selected = _selected_texts(nodes, units)
    assert not [text for text in selected if GROUP_HEADING_PATTERN.match(text)]
    assert not [text for text in selected if SECTION_HEADING_PATTERN.match(text)]


def test_confirmed_biology_media_hash_becomes_an_exclusion_boundary() -> None:
    exercise_label_hash = next(iter(BIOLOGY_MEDIA_HASHES["exercise_label"]))
    nodes = [
        DocNode(index=1, text="1．第一题"),
        DocNode(index=2, text="A．选项A"),
        DocNode(index=3, text="D．选项D"),
        DocNode(
            index=4,
            text="",
            has_inline_media=True,
            metadata={"media_sha256": [exercise_label_hash]},
        ),
        DocNode(index=5, text="2．第二题"),
    ]

    units = build_question_units_from_nodes(
        "作业1　细胞生活的环境.docx",
        "理科",
        nodes,
        overlay_name="future_biology",
    )

    assert classify_media_hashes_for_context(
        [exercise_label_hash], "future_biology"
    ) == "exercise_label"
    assert [unit.source_span for unit in units] == [(1, 3), (5, 5)]
    assert all(4 not in unit.media_blocks for unit in units)


def test_wps_inter_question_role_uses_context_not_dimensions_alone() -> None:
    obstacle = "Img_DecorativeHeader_170x42"

    assert classify_inter_question_media_boundary_for_context(
        current_text="",
        next_text="9．下一道题",
        obstacle_type=obstacle,
        overlay_name="future_biology",
    ) == "exercise_label"
    assert classify_inter_question_media_boundary_for_context(
        current_text="",
        next_text="9．下一道题",
        obstacle_type=obstacle,
        overlay_name=None,
    ) is None
    assert classify_inter_question_media_boundary_for_context(
        current_text="如图所示",
        next_text="9．下一道题",
        obstacle_type=obstacle,
        overlay_name="future_biology",
    ) is None
    assert classify_inter_question_media_boundary_for_context(
        current_text="",
        next_text="题目续行",
        obstacle_type=obstacle,
        overlay_name="future_biology",
    ) is None


def test_real_batch_has_zero_structure_or_confirmed_media_contamination() -> None:
    if not REAL_BATCH.is_dir():
        pytest.skip("当前机器没有未来高二生物真实原题批次")

    paths = _real_question_documents()
    assert len(paths) == 52
    source_hashes = {path: _sha256_file(path) for path in paths}
    total_units = 0
    excluded_role_counts: Counter[str] = Counter()

    for path in paths:
        nodes = scan_docx_nodes(path)
        sample_text = " ".join(node.text for node in nodes[:20])
        assert detect_subject_overlay(
            path.name, sample_text, base_subject="理科"
        ) == "future_biology", path.name

        units = build_question_units_from_docx(path, grade_hint="高二")
        total_units += len(units)
        direct_ids = [
            match.group(1)
            for node in nodes
            if (match := TOP_LEVEL_QUESTION_PATTERN.match(node.text or ""))
        ]
        assert [unit.question_id for unit in units] == direct_ids, path.name
        assert all(not unit.question_id.startswith(("(", "（")) for unit in units)

        selected = _selected_texts(nodes, units)
        assert not [text for text in selected if GROUP_HEADING_PATTERN.match(text)], path.name
        assert not [text for text in selected if SECTION_HEADING_PATTERN.match(text)], path.name

        selected_indexes = {
            index
            for unit in units
            for index in range(unit.source_span[0], unit.source_span[1] + 1)
        }
        for node in nodes:
            role = classify_media_hashes_for_context(
                node.metadata.get("media_sha256", []),
                "future_biology",
            )
            if not role:
                continue
            excluded_role_counts[role] += 1
            assert node.index not in selected_indexes, (path.name, node.index, role)

    assert total_units == 764
    assert excluded_role_counts == Counter(
        {
            "document_banner": 1,
            "exercise_label": 88,
        }
    )
    assert all(_sha256_file(path) == digest for path, digest in source_hashes.items())


def test_real_sample_action_plan_uses_biology_overlay_and_clean_boundaries() -> None:
    sample = REAL_BATCH / "选必一活页" / "第1章　作业1　细胞生活的环境.docx"
    if not sample.is_file():
        pytest.skip("当前机器没有未来高二生物作业1样本")

    bundle = build_preflight_bundle(sample, include_docling=False)
    actions = bundle["plan"]["actions"]

    assert len(actions) == 15
    assert all(action["subject_overlay"] == "future_biology" for action in actions)
    assert not [
        action
        for action in actions
        if "题组" in action["preview"] or "非选择题" in action["preview"]
    ]
    action_eight = next(action for action in actions if action["question_ids"] == ["8"])
    assert 43 not in action_eight["source_ref"]["media_paragraphs"]
