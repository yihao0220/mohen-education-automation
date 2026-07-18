from __future__ import annotations

from dataclasses import dataclass, field
import re

from .strategies import ENGLISH_STRATEGY, LANGUAGE_STRATEGY, SCIENCE_STRATEGY


GEOGRAPHY_KEYWORDS = (
    "等高线",
    "地形图",
    "海陆分布",
    "亚马孙河",
    "刚果河",
    "亚洲气候",
    "气候类型",
    "降水量",
    "地壳运动",
    "板块",
    "山脉",
    "海峡",
    "河流",
    "西藏",
    "鱼龙化石",
    "波斯湾",
    "地理",
)


HISTORY_KEYWORDS = (
    "政治制度",
    "中华民国",
    "人民代表大会",
    "法治建设",
    "精神文明",
    "民族关系",
    "对外交往",
    "民族政策",
    "外交方针",
    "文官制度",
    "官员的选拔",
    "赋税制度",
    "户籍制度",
    "基层治理",
    "社会保障",
    "币制",
    "货币",
    "国际金融秩序",
    "近代西方",
    "国际法",
    "中国历代",
    "变法",
    "改革",
    "官员",
    "选拔",
    "管理",
    "法律",
    "教化",
    "中国外交",
    "和平共处",
    "中美关系",
    "人类命运共同体",
    "制度",
)


HISTORY_ELECTIVE_TWO_DOC_KEYWORDS = (
    "从食物采集到食物生产",
    "食物物种交流",
    "食物的生产、储备与食品安全",
    "古代的生产工具与劳作",
    "工业革命与工厂制度",
    "现代技术进步与人类社会发展",
    "古代的商业贸易",
    "世界市场与商业贸易",
    "世界经济的发展",
    "古代的村落、集镇和城市",
    "近代以来的城市化进程",
    "水陆交通的变迁",
    "现代交通运输的新变化",
    "历史上的疫病与医学成就",
    "现代医疗卫生体系与社会生活",
)


@dataclass(frozen=True)
class SubjectOverlay:
    name: str
    base_subject: str
    required_doc_name_keywords: tuple[str, ...] = ()
    doc_name_keywords: tuple[str, ...] = ()
    content_keywords: tuple[str, ...] = ()
    min_keyword_hits: int = 1
    preserve_question_prefix_docs: tuple[str, ...] = ()
    native_table_input_doc_keywords: tuple[str, ...] = ()
    ignored_question_start_patterns: tuple[str, ...] = ()
    span_boundary_patterns: tuple[str, ...] = ()
    numbered_intro_patterns: tuple[str, ...] = ()
    leading_context_patterns: tuple[str, ...] = ()
    group_leading_context_questions: bool = False
    leading_context_group_boundary_patterns: tuple[str, ...] = ()
    question_input_excluded_patterns: tuple[str, ...] = ()
    excluded_media_sha256_by_role: dict[str, tuple[str, ...]] = field(default_factory=dict)
    inferred_inter_question_media_role: str | None = None
    warning_details: dict[str, str] = field(default_factory=dict)

    def matches(self, doc_name: str, sample_text: str = "", base_subject: str | None = None) -> bool:
        if base_subject and self.base_subject != base_subject:
            return False
        if self.required_doc_name_keywords and not any(
            keyword in (doc_name or "") for keyword in self.required_doc_name_keywords
        ):
            return False
        if doc_name in self.preserve_question_prefix_docs:
            return True
        if any(keyword in (doc_name or "") for keyword in self.doc_name_keywords):
            return True
        if not sample_text or not self.content_keywords:
            return False
        hits = sum(1 for keyword in self.content_keywords if keyword in sample_text)
        return hits >= self.min_keyword_hits

    def matches_any(self, patterns: tuple[str, ...], text: str) -> bool:
        return any(re.match(pattern, text or "") for pattern in patterns)


ZHONGMEI_CHINESE_GROUP_BOUNDARY_PATTERNS = (
    r"^\s*[一二三四五六七八九十]+[、．.]\s*[^。？！]{1,24}\s*$",
)


ZHONGMEI_CHINESE_IGNORED_PATTERNS = (
    r"^\s*一、专项训练\s*$",
    r"^\s*[（(]一[）)]\s*请比对下列选项与原文，说明选项错在何处。\s*$",
    r"^\s*[（(]二[）)]\s*图文解读题\s*$",
    r"^\s*[（(]一[）)]\s*分析评价信息\s*$",
    r"^\s*[（(]二[）)]\s*逻辑推断\s*$",
    r"^\s*[（(]三[）)]\s*分析论证特点\s*$",
    r"^\s*[（(]二[）)]\s*梳理论述思路\s*$",
    r"^\s*[（(]一[）)]\s*分析理据关系\s*$",
    *ZHONGMEI_CHINESE_GROUP_BOUNDARY_PATTERNS,
    r"^\s*[（(][一二三四五六七八九十]+[）)]\s*(?:高考专练|典题专练|专项训练|综合训练)\s*$",
)


ZHONGMEI_CHINESE_LEADING_CONTEXT_PATTERNS = (
    # 带阿拉伯题号的“1．阅读下面……”本身是题目，不能当成前置材料提示。
    r"^(?!\s*\d+\s*[．.、]).*阅读下面.*完成(?:文后|后面|下面|下列)?.*(?:题目|各题|小题)\s*[。.]?\s*$",
)


ZHONGMEI_CHINESE_QUESTION_INPUT_EXCLUDED_PATTERNS = (
    r"^\s*[（(][一二三四五六七八九十]+[）)]\s*(?:高考专练|典题专练)\s*$",
)


ZHONGMEI_CHINESE_OVERLAY = SubjectOverlay(
    name="zhongmei_chinese",
    base_subject="文科",
    required_doc_name_keywords=("对点练案",),
    content_keywords=(
        "分值",
        "文中",
        "阅读下面",
        "下列",
        "请简要",
        "选项：",
        "原文：",
        "补写",
        "翻译",
        "默写",
        "表达效果",
        "诗",
        "语句",
        "语段",
        "文章",
        "小说",
        "文言",
        "词语",
        "成语",
        "解说词",
    ),
    min_keyword_hits=2,
    ignored_question_start_patterns=ZHONGMEI_CHINESE_IGNORED_PATTERNS,
    span_boundary_patterns=ZHONGMEI_CHINESE_IGNORED_PATTERNS,
    leading_context_patterns=ZHONGMEI_CHINESE_LEADING_CONTEXT_PATTERNS,
    group_leading_context_questions=True,
    leading_context_group_boundary_patterns=ZHONGMEI_CHINESE_GROUP_BOUNDARY_PATTERNS,
    question_input_excluded_patterns=ZHONGMEI_CHINESE_QUESTION_INPUT_EXCLUDED_PATTERNS,
)


GEOGRAPHY_OVERLAY = SubjectOverlay(
    name="geography",
    base_subject="文科",
    doc_name_keywords=("地理",),
    content_keywords=GEOGRAPHY_KEYWORDS,
    min_keyword_hits=2,
    preserve_question_prefix_docs=(
        "限训4.docx",
        "限训5.docx",
        "限训11：5.7.docx",
        "限训12.docx",
        "限训16：6.14.docx",
    ),
)


HISTORY_OVERLAY = SubjectOverlay(
    name="history",
    base_subject="文科",
    doc_name_keywords=(
        "历史",
        "政治制度",
        "国际法",
        "民族关系",
        "民族政策",
        "外交",
        "货币",
        "赋税",
        "户籍",
        "基层治理",
        "社会保障",
        "变法",
        "改革",
        "官员",
        "文官制度",
        "法治",
        "教化",
    ) + HISTORY_ELECTIVE_TWO_DOC_KEYWORDS,
    content_keywords=HISTORY_KEYWORDS,
    min_keyword_hits=2,
    native_table_input_doc_keywords=HISTORY_ELECTIVE_TWO_DOC_KEYWORDS,
)


QINGYAN_MATH_OVERLAY = SubjectOverlay(
    name="qingyan_math",
    base_subject="理科",
    doc_name_keywords=(
        "五年级暑假数拔教材",
        "清艳",
        "数拔",
    ),
    content_keywords=(
        "比赛中的数学问题",
        "运算律的复杂运用",
        "小数的简便计算",
        "等差数列",
        "公式类计算",
        "特殊数列",
    ),
    min_keyword_hits=2,
    native_table_input_doc_keywords=(
        "五年级暑假数拔教材",
        "清艳",
        "数拔",
    ),
    ignored_question_start_patterns=(
        r"^\s*\d+[．.]\s*在进行相关简便计算",
        r"^\s*\d+[．.]\s*等差数列求和公式",
        r"^\s*\d+[．.]\s*结构特征[:：]?\s*$",
        r"^\s*\d+[．.]\s*公式展开证明[:：]?\s*$",
        r"^\s*\d+[．.]\s*公式图形证明[:：]?\s*$",
        r"^\s*\d+[．.]\s*完全平方公式变形[:：]?\s*$",
    ),
    span_boundary_patterns=(
        r"^\s*第\d+章\s+.*",
        r"^\s*[一二三四五六七八九十]+[、．.]\s*.*(?:基础知识讲解|课堂例题|思维冲浪|实战演练).*$",
        r"^\s*[一二三四五六七八九十]+[、．.]\s*(?:平方差公式|完全平方公式|错位相减)\s*$",
        r"^\s*(?:出类拔萃|实战演练|基础训练|能力提优|极限挑战)\s*$",
        r"^\s*随堂练习[:：]?\s*$",
        r"^\s*巩固练习\d*[:：]?\s*$",
        r"^\s*例\d+[:：].*",
        r"^\s*例题\d*[:：]?.*",
        r"^\s*\d+[．.]\s*在进行相关简便计算",
        r"^\s*\d+[．.]\s*等差数列求和公式",
        r"^\s*\d+[．.]\s*结构特征[:：]?\s*$",
        r"^\s*\d+[．.]\s*公式展开证明[:：]?\s*$",
        r"^\s*\d+[．.]\s*公式图形证明[:：]?\s*$",
        r"^\s*\d+[．.]\s*完全平方公式变形[:：]?\s*$",
    ),
    numbered_intro_patterns=(
        r"^\s*\d+\s*[、．.]\s*(?:随堂练习[:：]?|基础训练|能力提优|极限挑战)\s*$",
    ),
)


FUTURE_BIOLOGY_STRUCTURE_PATTERNS = (
    r"^\s*题组[一二三四五六七八九十\d]+(?:\s|　|[：:]).*$",
    r"^\s*[一二三四五六七八九十]+[、．.]\s*(?:选择题|非选择题)(?:\s*[（(].*)?$",
)


FUTURE_BIOLOGY_OVERLAY = SubjectOverlay(
    name="future_biology",
    base_subject="理科",
    required_doc_name_keywords=("作业", "检测试卷"),
    content_keywords=(
        "细胞",
        "内环境",
        "稳态",
        "神经",
        "反射",
        "兴奋",
        "激素",
        "免疫",
        "生长素",
        "植物",
        "种群",
        "群落",
        "生态",
        "生物",
        "酵母菌",
        "食物链",
        "能量流动",
    ),
    min_keyword_hits=1,
    ignored_question_start_patterns=FUTURE_BIOLOGY_STRUCTURE_PATTERNS,
    span_boundary_patterns=FUTURE_BIOLOGY_STRUCTURE_PATTERNS,
    excluded_media_sha256_by_role={
        "document_banner": (
            "363013449da1db958e2f66f789d87f7eede936259833c21c1b645b8dc3953e71",
        ),
        "exercise_label": (
            "f379746c07eb677523e68d9032053391d5982b4b933addfbe5a29de84d6d62de",
            "8c22c23eb8612cae71029cd75df419b262b028da00d66be51a0829c22004773a",
            "169289e4b5356a198f95facc3778da4e755b303d99f3b6155332099d69ed9de5",
        ),
    },
    inferred_inter_question_media_role="exercise_label",
)


REGISTERED_OVERLAYS: tuple[SubjectOverlay, ...] = (
    ZHONGMEI_CHINESE_OVERLAY,
    HISTORY_OVERLAY,
    GEOGRAPHY_OVERLAY,
    QINGYAN_MATH_OVERLAY,
    FUTURE_BIOLOGY_OVERLAY,
)


def get_subject_overlay(overlay_name: str | None) -> SubjectOverlay | None:
    if not overlay_name:
        return None
    for overlay in REGISTERED_OVERLAYS:
        if overlay.name == overlay_name:
            return overlay
    return None


def detect_subject_overlay(
    doc_name: str,
    sample_text: str = "",
    base_subject: str | None = None,
) -> str | None:
    for overlay in REGISTERED_OVERLAYS:
        if overlay.matches(doc_name, sample_text, base_subject=base_subject):
            return overlay.name
    return None


def choose_strategy_for_context(subject_name: str | None, overlay_name: str | None):
    overlay = get_subject_overlay(overlay_name)
    effective_subject = overlay.base_subject if overlay else subject_name
    if effective_subject == "文科":
        return LANGUAGE_STRATEGY
    if effective_subject == "理科":
        return SCIENCE_STRATEGY
    if effective_subject == "英语":
        return ENGLISH_STRATEGY
    return None


def should_skip_question_start_for_context(text: str, overlay_name: str | None = None) -> bool:
    overlay = get_subject_overlay(overlay_name)
    return bool(overlay and overlay.matches_any(overlay.ignored_question_start_patterns, text or ""))


def is_question_span_boundary_for_context(text: str, overlay_name: str | None = None) -> bool:
    overlay = get_subject_overlay(overlay_name)
    return bool(overlay and overlay.matches_any(overlay.span_boundary_patterns, (text or "").strip()))


def classify_media_hashes_for_context(
    media_sha256: list[str] | tuple[str, ...],
    overlay_name: str | None = None,
) -> str | None:
    """按项目覆盖层中已人工确认的图片哈希返回排除角色。"""
    overlay = get_subject_overlay(overlay_name)
    if not overlay or not media_sha256:
        return None
    observed = set(media_sha256)
    for role, known_hashes in overlay.excluded_media_sha256_by_role.items():
        if observed.intersection(known_hashes):
            return role
    return None


def classify_inter_question_media_boundary_for_context(
    current_text: str,
    next_text: str,
    obstacle_type: str | None,
    overlay_name: str | None = None,
) -> str | None:
    """识别 WPS 中位于两道顶层题之间的已知装饰图片区。"""
    overlay = get_subject_overlay(overlay_name)
    if not overlay or not overlay.inferred_inter_question_media_role:
        return None
    if (current_text or "").strip():
        return None
    if "DecorativeHeader" not in str(obstacle_type or ""):
        return None
    if not re.match(r"^\s*\d+\s*[．.、]", next_text or ""):
        return None
    return overlay.inferred_inter_question_media_role


def is_question_input_excluded_for_context(
    text: str,
    overlay_name: str | None = None,
) -> bool:
    overlay = get_subject_overlay(overlay_name)
    return bool(
        overlay
        and overlay.matches_any(
            overlay.question_input_excluded_patterns,
            (text or "").strip(),
        )
    )


def is_numbered_intro_for_context(text: str, overlay_name: str | None = None) -> bool:
    overlay = get_subject_overlay(overlay_name)
    return bool(overlay and overlay.matches_any(overlay.numbered_intro_patterns, (text or "").strip()))


def is_leading_context_start_for_context(text: str, overlay_name: str | None = None) -> bool:
    overlay = get_subject_overlay(overlay_name)
    return bool(overlay and overlay.matches_any(overlay.leading_context_patterns, (text or "").strip()))


def is_leading_context_group_boundary_for_context(
    text: str,
    overlay_name: str | None = None,
) -> bool:
    overlay = get_subject_overlay(overlay_name)
    return bool(
        overlay
        and overlay.matches_any(
            overlay.leading_context_group_boundary_patterns,
            (text or "").strip(),
        )
    )


def should_preserve_question_prefix_for_context(doc_name: str, overlay_name: str | None = None) -> bool:
    overlay = get_subject_overlay(overlay_name)
    if overlay and doc_name in overlay.preserve_question_prefix_docs:
        return True
    for registered in REGISTERED_OVERLAYS:
        if doc_name in registered.preserve_question_prefix_docs:
            return True
    return False


def should_use_native_table_input_for_context(doc_name: str, overlay_name: str | None = None) -> bool:
    """所有学科和项目的表格题都保留原生表格直接录入。"""
    return True
