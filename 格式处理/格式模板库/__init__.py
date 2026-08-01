# coding: utf-8
"""
格式模板库

包含所有已定义的文档格式模板
"""

import re

try:
    from ..common import set_document_font
except ImportError:
    from common import set_document_font


def matches_garbage_pattern(text, patterns):
    return any(re.match(pattern, text) for pattern in patterns)


def pattern_match_score(doc, patterns, cached_texts=None):
    texts = cached_texts if cached_texts is not None else [
        paragraph.Range.Text.strip() for paragraph in doc.Paragraphs
    ]
    nonempty_texts = [text for text in texts if text]
    if not nonempty_texts:
        return 0
    matched_lines = sum(
        any(re.search(pattern, text) for pattern in patterns)
        for text in nonempty_texts
    )
    return matched_lines / len(nonempty_texts)


def set_standard_font(doc):
    if set_document_font(doc):
        print("   ✓ 字体格式设置完成：小四、黑色、宋体/Times New Roman、不加粗")
    else:
        print("   ! 字体设置失败，详情见上方错误日志")

from . import template_a
from . import template_b
from . import template_c
from . import template_d
from . import template_e
from . import template_chinese
from . import template_math
from . import template_future_physics
from . import template_future_history
from . import template_nancheng_math

__all__ = [
    "matches_garbage_pattern",
    "pattern_match_score",
    "set_standard_font",
    "template_a",
    "template_b",
    "template_c",
    "template_d",
    "template_e",
    "template_chinese",
    "template_math",
    "template_future_physics",
    "template_future_history",
    "template_nancheng_math",
]
