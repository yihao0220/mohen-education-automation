from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class DocNode:
    index: int
    text: str
    has_inline_media: bool = False
    has_anchor_media: bool = False
    page_break_before: bool = False
    metadata: dict = field(default_factory=dict)

    @property
    def has_media(self) -> bool:
        return self.has_inline_media or self.has_anchor_media


@dataclass
class QuestionUnit:
    question_id: str
    subject: str
    subject_overlay: str | None
    grade_hint: str | None
    question_type: str
    stem_blocks: list[str]
    option_blocks: list[str]
    subquestions: list[str]
    media_blocks: list[int]
    material_blocks: list[str]
    source_span: tuple[int, int]
    confidence: float = 1.0
    warnings: list[str] = field(default_factory=list)
    node_type: str = "STD"

    @property
    def preview(self) -> str:
        text = " ".join((self.material_blocks + self.stem_blocks)[:2]).strip()
        return text[:80]

    @property
    def requires_review(self) -> bool:
        return self.confidence < 0.75 or bool(self.warnings)


@dataclass
class AnswerItem:
    item_id: str
    text: str


@dataclass
class AnswerUnit:
    question_id: str
    answer_mode: str
    answer_items: list[AnswerItem]
    analysis_items: list[AnswerItem]
    confidence: float = 1.0
    review_flags: list[str] = field(default_factory=list)
    source_span: tuple[int, int] = (0, 0)
    answer_span: tuple[int, int] = (0, 0)
    analysis_span: tuple[int, int] = (0, 0)
    metadata: dict = field(default_factory=dict)

    @property
    def requires_review(self) -> bool:
        return self.confidence < 0.75 or bool(self.review_flags)


@dataclass
class ReviewIssue:
    module: str
    question_id: str
    severity: str
    title: str
    detail: str


@dataclass
class ReviewReport:
    source_name: str
    summary: dict[str, int]
    issues: list[ReviewIssue] = field(default_factory=list)
