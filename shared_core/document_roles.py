from __future__ import annotations

from collections import Counter
from pathlib import Path
import re
import unicodedata
from typing import Any

from docx import Document


ROLE_SCHEMA_VERSION = "1.0"
QUESTION_START_PATTERN = re.compile(
    r"^\s*(?:[★*/]\s*)?(?:【[A-Za-z]】|[A-Za-z])?\s*\d+\s*[．.、]"
)
OPTION_PATTERN = re.compile(r"^\s*[A-HＡ-Ｈ]\s*[．.、:：]")
SECTION_PREFIX_PATTERN = re.compile(
    r"^\s*(?:第[一二三四五六七八九十百0-9]+(?:部分|章|节|单元)|[一二三四五六七八九十]+[、.．])"
)
INTERNAL_HEADING_PATTERN = re.compile(
    r"^\s*[（(][一二三四五六七八九十]+[）)]\s*\S+"
)
TERMINAL_PUNCTUATION_PATTERN = re.compile(r"[。！？；：.!?;:]\s*$")


def _normalize_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKC", text or "")
    return re.sub(r"\s+", "", normalized)


def _style_name(paragraph) -> str:
    return str(getattr(getattr(paragraph, "style", None), "name", "") or "")


def _direct_bold_ratio(paragraph) -> float:
    total = 0
    bold = 0
    for run in paragraph.runs:
        length = len((run.text or "").strip())
        if not length:
            continue
        total += length
        if run.bold is True:
            bold += length
    return round(bold / total, 3) if total else 0.0


def _max_font_size_pt(paragraph) -> float | None:
    sizes = [
        float(run.font.size.pt)
        for run in paragraph.runs
        if run.font.size is not None
    ]
    return round(max(sizes), 2) if sizes else None


def _extract_features(docx_path: str | Path) -> list[dict[str, Any]]:
    document = Document(docx_path)
    features: list[dict[str, Any]] = []
    for paragraph_index, paragraph in enumerate(document.paragraphs, 1):
        text = (paragraph.text or "").strip()
        if not text:
            continue
        style_name = _style_name(paragraph)
        style_lower = style_name.lower()
        xml = paragraph._p.xml
        features.append(
            {
                "paragraph_index": paragraph_index,
                "text": text,
                "style": style_name,
                "is_title_style": style_lower == "title" or style_name == "标题",
                "is_heading_style": (
                    style_lower.startswith("heading")
                    or style_name.startswith("标题")
                ),
                "direct_bold_ratio": _direct_bold_ratio(paragraph),
                "max_font_size_pt": _max_font_size_pt(paragraph),
                "has_media_or_formula": (
                    "<w:drawing" in xml
                    or "<w:pict" in xml
                    or "<m:oMath" in xml
                ),
                "is_question_start": bool(QUESTION_START_PATTERN.match(text)),
                "is_option": bool(OPTION_PATTERN.match(text)),
                "has_section_prefix": bool(SECTION_PREFIX_PATTERN.match(text)),
                "has_internal_heading_prefix": bool(INTERNAL_HEADING_PATTERN.match(text)),
                "is_short": len(text) <= 30,
                "has_terminal_punctuation": bool(TERMINAL_PUNCTUATION_PATTERN.search(text)),
            }
        )
    return features


def _align_docling_labels(
    paragraph_features: list[dict[str, Any]],
    docling_items: list[dict[str, Any]],
) -> tuple[dict[int, str], dict[str, int]]:
    normalized_paragraphs = [_normalize_text(item["text"]) for item in paragraph_features]
    labels_by_paragraph: dict[int, str] = {}
    cursor = 0
    aligned_count = 0

    for item in sorted(docling_items, key=lambda value: int(value.get("order", 0))):
        normalized_item = _normalize_text(str(item.get("text") or ""))
        if not normalized_item:
            continue
        match_position = next(
            (
                position
                for position in range(cursor, len(normalized_paragraphs))
                if normalized_paragraphs[position] == normalized_item
            ),
            None,
        )
        if match_position is None:
            continue
        paragraph_index = paragraph_features[match_position]["paragraph_index"]
        labels_by_paragraph[paragraph_index] = str(item.get("label") or "unknown")
        aligned_count += 1
        cursor = match_position + 1

    return labels_by_paragraph, {
        "source_count": len(docling_items),
        "aligned_count": aligned_count,
        "unaligned_count": max(0, len(docling_items) - aligned_count),
    }


def _confidence(score: int) -> float:
    return round(min(0.99, 0.55 + max(score, 0) * 0.05), 2)


def _classify_feature(
    feature: dict[str, Any],
    *,
    position: int,
    first_question_position: int | None,
    next_is_question: bool,
    docling_label: str | None,
) -> dict[str, Any]:
    evidence: list[str] = []
    if feature["is_question_start"]:
        evidence.append("明确题号模式")
        if docling_label:
            evidence.append(f"Docling标签={docling_label}")
        return {"role": "question_start", "score": 100, "confidence": 0.99, "evidence": evidence}

    if feature["is_option"]:
        evidence.append("明确选项前缀")
        if docling_label:
            evidence.append(f"Docling标签={docling_label}")
        return {"role": "option", "score": 100, "confidence": 0.99, "evidence": evidence}

    before_first_question = first_question_position is None or position < first_question_position
    after_question_start = first_question_position is not None and position > first_question_position

    title_score = 0
    if feature["is_title_style"]:
        title_score += 5
        evidence.append("Title/标题样式")
    if position == 0:
        title_score += 2
        evidence.append("首个非空段落")
    if (
        position == 0
        and feature["is_heading_style"]
        and not feature["has_section_prefix"]
    ):
        title_score += 3
        evidence.append("首段Heading/标题样式")
    if feature["is_short"]:
        title_score += 1
    if docling_label == "title":
        title_score += 2
        evidence.append("Docling标签=title")
    if title_score >= 6:
        if feature["is_short"]:
            evidence.append("短段落")
        return {
            "role": "document_title",
            "score": title_score,
            "confidence": _confidence(title_score),
            "evidence": evidence,
        }

    section_evidence: list[str] = []
    section_score = 0
    if feature["is_heading_style"]:
        section_score += 4
        section_evidence.append("Heading/标题样式")
    if feature["has_section_prefix"]:
        section_score += 3
        section_evidence.append("章节编号前缀")
    if feature["is_short"]:
        section_score += 1
        section_evidence.append("短段落")
    if not feature["has_terminal_punctuation"]:
        section_score += 1
        section_evidence.append("无句末标点")
    if before_first_question:
        section_score += 1
        section_evidence.append("位于首题之前")
    if docling_label == "section_header":
        section_score += 2
        section_evidence.append("Docling标签=section_header")
    if section_score >= 5 and (
        before_first_question
        or feature["has_section_prefix"]
    ):
        return {
            "role": "section_heading",
            "score": section_score,
            "confidence": _confidence(section_score),
            "evidence": section_evidence,
        }

    internal_evidence: list[str] = []
    internal_score = 0
    if after_question_start:
        internal_score += 2
        internal_evidence.append("位于首题之后")
    if feature["is_short"]:
        internal_score += 1
        internal_evidence.append("短段落")
    if not feature["has_terminal_punctuation"]:
        internal_score += 1
        internal_evidence.append("无句末标点")
    if feature["direct_bold_ratio"] >= 0.6:
        internal_score += 2
        internal_evidence.append("直接粗体占比>=0.60")
    if feature["has_internal_heading_prefix"]:
        internal_score += 2
        internal_evidence.append("中文括号标题前缀")
    if next_is_question:
        internal_score += 1
        internal_evidence.append("下一非空段落为题号")
    if docling_label in {"section_header", "title"}:
        internal_score += 2
        internal_evidence.append(f"Docling标签={docling_label}")
    if feature["has_media_or_formula"]:
        internal_score -= 3
        internal_evidence.append("含媒体或公式，降低标题置信度")
    if len(feature["text"]) > 50:
        internal_score -= 3
        internal_evidence.append("长文本，降低标题置信度")
    has_heading_signal = bool(
        feature["is_heading_style"]
        or feature["direct_bold_ratio"] >= 0.6
        or feature["has_internal_heading_prefix"]
        or docling_label in {"section_header", "title"}
    )
    if after_question_start and has_heading_signal and internal_score >= 5:
        return {
            "role": "internal_heading",
            "score": internal_score,
            "confidence": _confidence(internal_score),
            "evidence": internal_evidence,
        }

    if feature["is_heading_style"] and section_score >= 4:
        return {
            "role": "section_heading",
            "score": section_score,
            "confidence": _confidence(section_score),
            "evidence": section_evidence,
        }

    body_evidence = ["未达到标题角色阈值"]
    if docling_label:
        body_evidence.append(f"Docling标签={docling_label}")
    return {"role": "body", "score": 0, "confidence": 0.7, "evidence": body_evidence}


def build_document_role_profile(
    docx_path: str | Path,
    docling_items: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    features = _extract_features(docx_path)
    labels_by_paragraph, alignment = _align_docling_labels(features, docling_items or [])
    first_question_position = next(
        (position for position, item in enumerate(features) if item["is_question_start"]),
        None,
    )

    paragraphs: list[dict[str, Any]] = []
    for position, feature in enumerate(features):
        next_is_question = position + 1 < len(features) and features[position + 1]["is_question_start"]
        docling_label = labels_by_paragraph.get(feature["paragraph_index"])
        result = _classify_feature(
            feature,
            position=position,
            first_question_position=first_question_position,
            next_is_question=next_is_question,
            docling_label=docling_label,
        )
        paragraphs.append(
            {
                "paragraph_index": feature["paragraph_index"],
                "role": result["role"],
                "confidence": result["confidence"],
                "score": result["score"],
                "evidence": result["evidence"],
                "text": feature["text"][:160],
                "style": feature["style"],
                "direct_bold_ratio": feature["direct_bold_ratio"],
                "max_font_size_pt": feature["max_font_size_pt"],
                "docling_label": docling_label,
            }
        )

    role_counts = Counter(item["role"] for item in paragraphs)
    heading_candidates = [
        item
        for item in paragraphs
        if item["role"] in {"document_title", "section_heading", "internal_heading"}
    ]
    return {
        "schema_version": ROLE_SCHEMA_VERSION,
        "status": "success",
        "role_counts": dict(sorted(role_counts.items())),
        "paragraphs": paragraphs,
        "heading_candidates": heading_candidates,
        "internal_heading_candidates": [
            item for item in paragraphs if item["role"] == "internal_heading"
        ],
        "docling_alignment": alignment,
        "automatic_exclusion_enabled": False,
    }
