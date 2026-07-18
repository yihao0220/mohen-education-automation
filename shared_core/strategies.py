from __future__ import annotations

import re
from dataclasses import dataclass, field

QUESTION_PREFIX_PATTERN = r"(?:[★☆/]\s*)?(?:【\s*[GYC]\s*】\s*|[GYC]\s*)?"
SECTION_QUESTION_PREFIX_PATTERN = r"(?:[一二三四五六七八九十]+[、．.]\s*)?"
QUESTION_ID_PATTERN = re.compile(
    rf"^\s*{SECTION_QUESTION_PREFIX_PATTERN}{QUESTION_PREFIX_PATTERN}([（\(]?\d+[）\)]?|[一二三四五六七八九十]+)\s*[．.、]?"
)


QUESTION_RANGE_SUFFIX_PATTERN = r"\d+\s*[~～\-—、至]\s*\d+\s*(?:小)?题"
QUESTION_RANGE_PATTERN = re.compile(r"(\d+)\s*[~～\-—、至]\s*(\d+)\s*(?:小)?题")

READING_PATTERNS = [
    rf"据此回答\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"据此完成\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"回答\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"完成\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"读.*图.*回答\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"读.*图.*完成\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"读.*材料.*回答\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"读.*材料.*完成\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"下图为.*完成\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"如下图.*完成\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"据图完成\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    rf"据下图完成\s*{QUESTION_RANGE_SUFFIX_PATTERN}",
    r"完成下面小题",
    r"回答下面小题",
    r"完成下面各题",
    r"回答下面各题",
    r"完成下列各题",
    r"回答下列各题",
    r"完成下列小题",
    r"回答下列小题",
    r"完成下面题",
    r"回答下面题",
    r"完成下题",
    r"回答下题",
    r"完成下列问题",
    r"回答下列问题",
]


@dataclass
class BaseStrategy:
    name: str
    question_patterns: list[str]
    material_patterns: list[str] = field(default_factory=list)
    option_pattern: str = r"^\s*[A-DＡ-Ｄ][．.\s、]"
    subquestion_pattern: str = r"^\s*[（\(]\s*\d+\s*[）\)]"
    figure_keywords: list[str] = field(default_factory=list)
    material_preferred: bool = False

    def is_question_start(self, text: str) -> bool:
        return any(re.match(pattern, text) for pattern in self.question_patterns)

    def extract_question_id(self, text: str) -> str:
        match = QUESTION_ID_PATTERN.match(text)
        return match.group(1) if match else text[:12]

    def is_option_line(self, text: str) -> bool:
        return bool(re.match(self.option_pattern, text))

    def is_subquestion_line(self, text: str) -> bool:
        return bool(re.match(self.subquestion_pattern, text))

    def is_material_line(self, text: str) -> bool:
        return any(re.search(pattern, text, re.IGNORECASE) for pattern in self.material_patterns)

    def has_figure_reference(self, text: str) -> bool:
        return any(keyword in text for keyword in self.figure_keywords)


SCIENCE_STRATEGY = BaseStrategy(
    name="理科",
    question_patterns=[
        rf"^[\s_—\-]*{SECTION_QUESTION_PREFIX_PATTERN}{QUESTION_PREFIX_PATTERN}\d+\s*[．.]",
        rf"^[\s_—\-]*{SECTION_QUESTION_PREFIX_PATTERN}{QUESTION_PREFIX_PATTERN}\d+\s*、",
        rf"^[\s_—\-]*{QUESTION_PREFIX_PATTERN}[(\[（]\d+[)\]）]",
    ],
    figure_keywords=["如图", "图示", "示意图", "装置图", "曲线图", "能量图", "主要过程如图所示"],
)

LANGUAGE_STRATEGY = BaseStrategy(
    name="文科",
    question_patterns=[
        rf"^[\s_—\-]*{SECTION_QUESTION_PREFIX_PATTERN}{QUESTION_PREFIX_PATTERN}\d+\s*[．.]",
        rf"^[\s_—\-]*{SECTION_QUESTION_PREFIX_PATTERN}{QUESTION_PREFIX_PATTERN}\d+\s*、",
        rf"^[\s_—\-]*{QUESTION_PREFIX_PATTERN}[（\(]\s*\d+\s*[）\)]",
    ],
    material_patterns=READING_PATTERNS,
    figure_keywords=["读图", "如下图", "下图为", "下图", "如图", "示意图", "材料", "阅读图文资料"],
    material_preferred=True,
)

ENGLISH_STRATEGY = BaseStrategy(
    name="英语",
    question_patterns=[
        r"^\s*(Passage|Part|Text)?\s*[\(（\[]?[A-EＡ-Ｅ][\)）\]]?[．.]?\s*$",
        r"^[\s_—\-]*(\d+|[（(]\s*[)）]\s*\d+)[．.]",
        r".*?(?<!\d)(\d+[．.]?)\s*_{3,}.*",
    ],
    figure_keywords=["Passage", "picture", "graph", "table"],
)


def extract_normalized_question_id(text: str) -> str | None:
    match = QUESTION_ID_PATTERN.match(text or "")
    if not match:
        return None
    return match.group(1)


def extract_numeric_question_id(text: str) -> int | None:
    normalized = extract_normalized_question_id(text)
    if not normalized:
        return None
    digits = re.sub(r"[^\d]", "", normalized)
    if not digits:
        return None
    return int(digits)


def extract_material_question_range(text: str) -> tuple[int, int] | None:
    match = QUESTION_RANGE_PATTERN.search(text or "")
    if not match:
        return None
    start = int(match.group(1))
    end = int(match.group(2))
    if start > end:
        start, end = end, start
    return start, end


def strip_question_noise_prefix(text: str) -> str:
    return re.sub(
        rf"^(\s*){SECTION_QUESTION_PREFIX_PATTERN}{QUESTION_PREFIX_PATTERN}(?=\d+\s*[．.、])",
        r"\1",
        text or "",
        count=1,
    )


def choose_strategy(doc_name: str, sample_text: str = "") -> BaseStrategy:
    if any(x in doc_name for x in ["英语", "English", "外语", "XB"]):
        return ENGLISH_STRATEGY
    if any(x in doc_name for x in ["物理", "化学", "生物", "科学", "理综", "数学"]):
        return SCIENCE_STRATEGY
    if any(x in doc_name for x in ["政治", "历史", "地理", "道德", "文综", "语文", "分层作业", "同步练习"]):
        return LANGUAGE_STRATEGY
    if "Passage" in sample_text or "UNIT" in sample_text:
        return ENGLISH_STRATEGY
    if any(x in sample_text for x in ["阅读图文资料", "据此完成", "读图", "材料"]):
        return LANGUAGE_STRATEGY
    return SCIENCE_STRATEGY
