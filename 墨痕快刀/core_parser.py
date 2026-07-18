# core_parser.py
import re
import time
import os
import sys
import pyautogui
import config
from wps_helper import get_active_wps
from debug_logger import logger  # ✨ 导入日志模块

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from shared_core import (
    build_question_units_from_wps,
    build_question_units_from_wps_spans,
    classify_inter_question_media_boundary_for_context,
    detect_subject_overlay,
    format_question_warning_details,
    is_numbered_intro_for_context,
    is_question_input_excluded_for_context,
    is_question_span_boundary_for_context,
    scan_wps_nodes,
    should_preserve_question_prefix_for_context,
    should_skip_question_start_for_context,
)
from shared_core.strategies import strip_question_noise_prefix

def _is_numbered_main_question(text):
    section_prefix = getattr(config, "SECTION_QUESTION_PREFIX", "")
    return bool(re.match(rf'^[\s_—\-]*{section_prefix}{config.QUESTION_PREFIX}\d+\s*[．.]', text or ""))


def _is_bracket_subquestion(text):
    return bool(re.match(rf'^[\s_—\-]*{config.QUESTION_PREFIX}[（\(]\s*\d+\s*[）\)]', text or ""))


def _range_contains_table(doc, start_idx, end_idx):
    for p_idx in range(start_idx, end_idx + 1):
        try:
            if doc.Paragraphs(p_idx).Range.Information(12):
                return True
        except Exception:
            continue
    return False


def collect_input_excluded_paragraph_indices(
    paragraphs,
    start_idx,
    end_idx,
    overlay_name=None,
):
    excluded = []
    for p_idx in range(start_idx, end_idx + 1):
        try:
            text = paragraphs(p_idx).Range.Text.strip()
        except Exception:
            continue
        if is_question_input_excluded_for_context(text, overlay_name):
            excluded.append(p_idx)
    return excluded


def create_filtered_input_document(doc, source_range, overlay_name=None):
    """复制当前题块到不保存的临时文档，并删除只供排版使用的题内标题。"""
    temp_doc = doc.Application.Documents.Add()
    try:
        temp_doc.Content.FormattedText = source_range.FormattedText
        excluded = collect_input_excluded_paragraph_indices(
            temp_doc.Paragraphs,
            1,
            temp_doc.Paragraphs.Count,
            overlay_name=overlay_name,
        )
        for p_idx in reversed(excluded):
            temp_doc.Paragraphs(p_idx).Range.Delete()
        return temp_doc, excluded
    except Exception:
        try:
            temp_doc.Close(False)
        except Exception:
            pass
        raise


def is_filtered_input_copy_safe(*, contains_table, has_media):
    """富文本副本可保留媒体；只有原生表格不能走临时副本。"""
    return not contains_table


def describe_ignored_inline_obstacle(text, overlay_name=None):
    if is_question_input_excluded_for_context(text, overlay_name):
        return "题内标题"
    return "图片/空白段"


def sync_subject_overlay(doc, subject_name=None, content_sample=None):
    if content_sample is None:
        content_sample = doc.Range(0, min(5000, doc.Range().End)).Text
    config.CURRENT_SUBJECT_OVERLAY = detect_subject_overlay(
        doc.Name,
        content_sample,
        base_subject=subject_name,
    )
    return config.CURRENT_SUBJECT_OVERLAY

def detect_subject(doc):
    name = doc.Name
    content_sample = doc.Range(0, min(5000, doc.Range().End)).Text

    if any(x in name for x in ["英语", "English", "外语", "XB"]):
        detected = config.CONFIG_ENGLISH
    elif any(x in name for x in ["物理", "化学", "生物", "科学", "理综", "数学"]):
        detected = config.CONFIG_SCIENCE
    elif any(x in name for x in ["政治", "历史", "地理", "道德", "文综", "语文", "分层作业", "同步练习"]):
        detected = config.CONFIG_ARTS
    elif detect_subject_overlay(name, content_sample, base_subject="文科"):
        detected = config.CONFIG_ARTS
    elif "UNIT" in content_sample or "Passage" in content_sample:
        detected = config.CONFIG_ENGLISH
    elif any(x in content_sample for x in ["现代文阅读", "文言文", "阅读下面的文字", "阅读下面的材料", "语言文字运用",
                                            "通假字", "词类活用", "诗经", "补写出下列句子", "古代诗文", "记叙文", "说明文", "议论文", "散文"]):
        detected = config.CONFIG_ARTS
    elif "第一讲" in content_sample or "实验" in content_sample:
        detected = config.CONFIG_SCIENCE
    elif "第一课" in content_sample or "材料分析" in content_sample:
        detected = config.CONFIG_ARTS
    else:
        detected = config.CONFIG_SCIENCE

    sync_subject_overlay(doc, detected["name"], content_sample)
    return detected

def detect_best_format(doc, start_p, end_p):
    paras = doc.Paragraphs
    formats = config.CURRENT_CONFIG["formats"]
    scores = {name: 0 for name in formats}
    scan_len = min(200, end_p - start_p) 
    
    for i in range(start_p, start_p + scan_len):
        try:
            rng = paras(i).Range
            if rng.Information(12): continue
            text = rng.Text.strip()
            if not text: continue
            for name, pattern in formats.items():
                if re.match(pattern, text):
                    scores[name] += 1
        except: continue
    
    if config.CURRENT_CONFIG == config.CONFIG_ENGLISH:
        passage_count = scores.get("Passage", 0)
        if passage_count > 0:
            return formats["Passage"], "Passage"
            
    # ✨ 优化：如果存在标准大题号（如 "1."），且数量大于等于 2 个，则强制优先选择大题号，
    # 避免被茫茫多的小题号（如 "(1)", "(2)"）带偏，导致大题干丢失。
    if "标准点号" in scores and scores["标准点号"] >= 2:
        return formats["标准点号"], "标准点号"

    if not any(scores.values()): return None, None
    best = max(scores, key=lambda k: scores[k])
    return formats[best], best

def is_obstacle(rng):
    try:
        text = rng.Text.strip()
    except:
        return False, None
        
    if not text and getattr(rng, 'InlineShapes', None) and rng.InlineShapes.Count == 0: 
        return False, None

    overlay_name = getattr(config, "CURRENT_SUBJECT_OVERLAY", None)
    if text and should_skip_question_start_for_context(text, overlay_name):
        return True, "Overlay_IgnoredQuestionStart"
    if (
        text
        and is_question_span_boundary_for_context(text, overlay_name)
        and not is_numbered_intro_for_context(text, overlay_name)
    ):
        return True, "Overlay_SpanBoundary"
    
    for pattern in config.CURRENT_CONFIG["obstacles"]:
        if text and __import__('re').match(pattern, text): return True, pattern
        
    try:
        if getattr(rng, 'InlineShapes', None) and rng.InlineShapes.Count > 0:
            for i in range(1, rng.InlineShapes.Count + 1):
                shape = rng.InlineShapes(i)
                alt = str(shape.AlternativeText or "") + str(shape.Title or "")
                if alt:
                    for kw in ["拓展", "延伸", "导引", "点评", "考点", "技巧", "警示", "点拨", "变式", "训练"]:
                        if kw in alt:
                            return True, f"Img_Alt_{kw}"
                            
                try:
                    w = shape.Width
                    h = shape.Height
                    aspect = w / h if h else 0
                    
                    cleaned = __import__('re').sub(r'[\s\x00-\x1F]+', '', text)
                    if h < 50 and w > 40 and aspect > 2.0 and len(cleaned) < 5:
                        return True, f"Img_DecorativeHeader_{w:.0f}x{h:.0f}"
                except:
                    pass
    except:
        pass
        
    return False, None

def is_vip_header(text):
    if not text: return False
    for pattern in config.CURRENT_CONFIG["vip_headers"]:
        if re.match(pattern, text, re.IGNORECASE): return True
    return False

def is_reading_material(text):
    """
    识别地理/历史等科目的阅读材料段落。
    特征：包含"据此回答X~Y题"、"读...图/材料，完成X~Y题"、
    "下图为...完成X~Y题"、"完成下面小题/下列各题" 等标记
    """
    if not text:
        return False

    range_part = r"\d+\s*[~～\-—、至]\s*\d+"
    patterns = [
        rf"据此回答\s*{range_part}\s*题",                 # 据此回答1~2题
        rf"据此完成\s*{range_part}\s*题",                 # 据此完成1~2题
        rf"读.*图.*回答\s*{range_part}\s*题",             # 读图回答1~2题
        rf"读.*图.*完成\s*{range_part}\s*题",             # 读图完成1~2题
        rf"读.*材料.*回答\s*{range_part}\s*题",           # 读材料回答1~2题
        rf"读.*材料.*完成\s*{range_part}\s*题",           # 读材料完成1~2题
        rf"下图为.*完成\s*{range_part}\s*题",             # 下图为...完成1~2题
        rf"如下图.*完成\s*{range_part}\s*题",             # 如下图...完成1~2题
        rf"据图完成\s*{range_part}\s*题",                 # 据图完成1~2题
        rf"据下图完成\s*{range_part}\s*题",               # 据下图完成1~2题
        rf"完成\s*{range_part}\s*题",                     # 完成1~2题
        rf"回答\s*{range_part}\s*题",                     # 回答1~2题
        r"据此完成下面小题",                               # 据此完成下面小题
        r"据此回答下面小题",                               # 据此回答下面小题
        r"完成下面小题",                                   # 完成下面小题
        r"回答下面小题",                                   # 回答下面小题
        r"据此完成下面各题",                               # 据此完成下面各题
        r"据此回答下面各题",                               # 据此回答下面各题
        r"完成下面各题",                                   # 完成下面各题
        r"回答下面各题",                                   # 回答下面各题
        r"据此完成下列各题",                               # 据此完成下列各题
        r"据此回答下列各题",                               # 据此回答下列各题
        r"完成下列各题",                                   # 完成下列各题
        r"回答下列各题",                                   # 回答下列各题
        r"据此完成下列小题",                               # 据此完成下列小题
        r"据此回答下列小题",                               # 据此回答下列小题
        r"完成下列小题",                                   # 完成下列小题
        r"回答下列小题",                                   # 回答下列小题
        r"据此完成下面题",                                 # 据此完成下面题
        r"据此回答下面题",                                 # 据此回答下面题
        r"完成下面题",                                     # 完成下面题
        r"回答下面题",                                     # 回答下面题
        r"据此完成下题",                                   # 据此完成下题
        r"完成下题",                                       # 完成下题
        r"据此回答下题",                                   # 据此回答下题
        r"回答下题",                                       # 回答下题
        r"据此完成下列问题",                               # 据此完成下列问题
        r"据此回答下列问题",                               # 据此回答下列问题
        r"完成下列问题",                                   # 完成下列问题
        r"回答下列问题",                                   # 回答下列问题
        r"读.*图.*完成下面小题",                           # 读图，完成下面小题
        r"读.*图.*回答下面小题",                           # 读图，回答下面小题
        r"读.*图.*完成下面题",                             # 读图，完成下面题
        r"读.*图.*回答下面题",                             # 读图，回答下面题
        r"读.*图.*完成下列各题",                           # 读图，完成下列各题
        r"读.*图.*回答下列各题",                           # 读图，回答下列各题
        r"读.*图.*回答下面各题",                           # 读图，回答下面各题
        r"读.*材料.*完成下面小题",                         # 读材料，完成下面小题
        r"读.*材料.*回答下面小题",                         # 读材料，回答下面小题
        r"读.*材料.*完成下面题",                           # 读材料，完成下面题
        r"读.*材料.*回答下面题",                           # 读材料，回答下面题
        r"读.*材料.*完成下列各题",                         # 读材料，完成下列各题
        r"读.*材料.*回答下列各题",                         # 读材料，回答下列各题
        r"读.*材料.*回答下面各题",                         # 读材料，回答下面各题
    ]
    for pattern in patterns:
        if re.search(pattern, text, re.IGNORECASE):
            return True
    return False


def should_ignore_inline_obstacle(unit, text, obs_type):
    """
    录入阶段的二次兜底：
    阅读材料块中如果夹着图片段/空白图片区，不应因此把整题提前截断。
    但纯文字障碍物（如题组、建议用时、小节标题）仍然要阻断。
    """
    if isinstance(unit, str):
        node_type = unit
        question_type = ""
        warnings = []
    else:
        node_type = unit.node_type
        question_type = unit.question_type
        warnings = unit.warnings
        media_blocks = getattr(unit, "media_blocks", [])
        subquestions = getattr(unit, "subquestions", [])

    cleaned_text = (text or "").strip()
    if cleaned_text in getattr(config, "IGNORED_LAYER_HEADERS", ()):
        return False
    if node_type in {"READING", "LEADING_CONTEXT"} and not cleaned_text:
        return True

    if node_type in {"READING", "LEADING_CONTEXT"} and obs_type and "Img_" in str(obs_type):
        return True

    # 前置阅读材料题块的范围已经由 shared_core 定位到对应首题。
    # 题号前可能夹有“（一）高考专练”等项目标题，执行层不应再次截断。
    if (
        node_type == "LEADING_CONTEXT"
        and obs_type in {"Overlay_IgnoredQuestionStart", "Overlay_SpanBoundary"}
    ):
        return True

    # 文科大题内部常见的板块小标题，如【独特的自然环境】、【一粒种子的远征】，
    # 已经处于当前题块 source_span 内时，不应把整题提前截断。
    if (
        isinstance(obs_type, str)
        and obs_type == r"^【.*】$"
        and (node_type == "VIP" or question_type in ["subjective", "material", "material_choice"])
    ):
        return True

    # 理科图文选择题：图片常被误判为横条装饰图，此时不应截断整题
    if (
        question_type in ["choice", "material_choice"]
        and "image_related_question" in warnings
        and obs_type
        and "DecorativeHeader" in str(obs_type)
    ):
        return True

    # 理科带图大题：共享内核已识别出完整 span，题内插图不应在执行层再次截断。
    if (
        question_type in ["subjective", "material"]
        and "image_related_question" in warnings
        and obs_type
        and "DecorativeHeader" in str(obs_type)
        and (media_blocks or subquestions)
    ):
        return True

    return False


def should_pause_for_review(unit):
    if not getattr(config, "PAUSE_ON_LOW_CONFIDENCE", False):
        return False
    return unit.confidence < getattr(config, "AUTO_INPUT_MIN_CONFIDENCE", 0.75) or bool(unit.warnings)


def should_mark_for_review(unit):
    return unit.confidence < getattr(config, "AUTO_INPUT_MIN_CONFIDENCE", 0.75) or bool(unit.warnings)


def should_preserve_question_prefix(doc_name):
    return should_preserve_question_prefix_for_context(
        doc_name,
        getattr(config, "CURRENT_SUBJECT_OVERLAY", None),
    )


def format_deferred_review_items(units):
    items = []
    for unit in units:
        warning_details = format_question_warning_details(unit.warnings)
        detail = "；".join(warning_details[:2]) if warning_details else f"置信度 {unit.confidence:.2f}"
        items.append(f"第 {unit.question_id} 题：{detail}")
    return items

def get_sections(doc):
    print(f"📑 扫描目录...")
    sections = [] 
    paras = doc.Paragraphs
    count = paras.Count
    current_title = "文档开头 (前言)"
    current_start = 1
    patterns = config.CURRENT_CONFIG["sections"]
    
    for i in range(1, count + 1):
        try:
            rng = paras(i).Range
            if rng.Information(12):
                continue
            text = rng.Text.strip()
            # 过滤目录页码格式：制表符+数字、省略号+数字、多个空格+数字
            if re.search(r'\t\d+$', text) or re.search(r'…\d+$', text) or re.search(r'\s{3,}\d+$', text): continue
            is_section = False
            for pattern in patterns:
                if re.match(pattern, text):
                    is_section = True
                    break
            if is_section:
                sections.append({"title": current_title, "start": current_start, "end": i - 1})
                current_title = text
                current_start = i
        except: continue
        
    sections.append({"title": current_title, "start": current_start, "end": count})
    return sections

def process_chapter(section_info):
    wps = get_active_wps()
    if not wps: return
    try:
        doc = wps.ActiveDocument
        wps.Activate() 
        doc.Activate()
    except: return

    logger.section_start(section_info['title']) # ✨ 使用日志模块
    start_p, end_p = section_info['start'], section_info['end']
    try:
        if end_p > doc.Paragraphs.Count: end_p = doc.Paragraphs.Count
    except: pass

    # --- 英语特供：智能吞并模式 ---
    if config.CURRENT_CONFIG == config.CONFIG_ENGLISH:
        logger.english_mode_start()
        raw_starts = []
        paras = doc.Paragraphs
        
        # 1. 扫描所有潜在节点
        for i in range(start_p, end_p + 1):
            try:
                rng = paras(i).Range
                if rng.Information(12): continue
                text = rng.Text.strip()
                if not text: continue

                is_obs, obs_pat = is_obstacle(rng)
                
                # ✨ 关键：先检查是否是VIP，如果是VIP则优先作为VIP处理（即使它也匹配obstacle）
                node_type = None
                if is_vip_header(text):
                    node_type = "VIP"
                elif is_obs: 
                    logger.scan_obstacle_filtered(i, text, obs_pat)
                    continue 
                elif re.match(config.CONFIG_ENGLISH["formats"]["Passage"], text):
                    node_type = "PASSAGE"
                elif re.match(config.CONFIG_ENGLISH["formats"]["Standard"], text):
                    node_type = "STD"
                
                if node_type:
                    logger.scan_node_found(i, node_type, text) # ✨ 记录发现节点
                    raw_starts.append({"idx": i, "type": node_type, "text": text})
            except: continue

        # 2. 吞并算法
        q_nodes = []
        skip_indices = set()
        greedy_keywords = ["完形", "Cloze", "七选五", "写作", "短文改错", "阅读", "大题", "部分", "节", "非选择题","材料","单句","短语","完成句子"]
        
        for k in range(len(raw_starts)):
            if k in skip_indices: continue
            curr = raw_starts[k]
            q_nodes.append(curr)
            curr_text = curr['text']
            is_greedy = any(kw in curr_text for kw in greedy_keywords)
            
            has_merged_passage = False
            has_merged_section = False
            has_merged_material = False

            for j in range(k+1, len(raw_starts)):
                nxt = raw_starts[j]
                nxt_text = nxt['text']
                dist = nxt['idx'] - curr['idx']
                should_merge = False
                reason = "未知"
                
                if curr['type'] == 'VIP':
                    if nxt['type'] == 'STD':
                        if is_greedy: 
                            should_merge, reason = True, "贪婪VIP吞小题"
                        elif j == k + 1: 
                            should_merge, reason = True, "紧邻小题"
                            
                    elif nxt['type'] == 'PASSAGE':
                        if dist < 50:
                            if any(x in curr_text for x in ["大题", "部分", "节", "非选择题", "阅读"]):
                                if not has_merged_passage:
                                    should_merge, reason = True, "VIP首篇吞并"
                                    has_merged_passage = True
                                else:
                                    reason = "拒绝:已吞过一篇"
                            elif j == k + 1: 
                                should_merge, reason = True, "紧邻首篇"
                            else:
                                reason = "拒绝:不在白名单且不紧邻"
                        else:
                            reason = f"拒绝:距离过远({dist}>50)"

                         # 情况 C: 后面跟着 VIP
                                        # 情况 C: 后面跟着 VIP
                    elif nxt['type'] == 'VIP' and dist < 20: 
                        # 大题/非选择题 吞并一切 VIP
                        if any(x in curr_text for x in ["大题", "非选择题"]): 
                            should_merge, reason = True, "大题吞子标题"
                            
                        # ✨ 修复：允许 "节" 吞并 "听第X段材料"，但仅限第一段！
                        elif "节" in curr_text and "材料" in nxt_text:
                            # 必须初始化 has_merged_material 标志位 (在 process_chapter 开头)
                            # 如果还没吞过听力材料，那就吞这一段
                            if not has_merged_material:
                                should_merge, reason = True, "节首段材料吞并"
                                has_merged_material = True # 🔒 锁定！后续材料不再吞并
                            else:
                                reason = "拒绝:已吞过一段材料"
                            
                        # 部分 只能吞并紧邻的第一个 节
                        elif "部分" in curr_text and "节" in nxt_text: 
                            if not has_merged_section:
                                should_merge, reason = True, "部分吞节"
                                has_merged_section = True
                            else:
                                reason = "拒绝:已吞过一节"
                        else:
                            reason = "拒绝:普通VIP不互吞"

                elif curr['type'] == 'PASSAGE':
                    if nxt['type'] == 'STD': 
                        should_merge, reason = True, "文章带小题"
                
                if should_merge:
                    logger.merge_decision(curr, nxt, True, reason) # ✨ 记录成功吞并
                    skip_indices.add(j)
                else:
                    if curr['type'] == 'VIP':
                        logger.merge_decision(curr, nxt, False, reason) # ✨ 记录拒绝原因
                    break

    # 📜 语文文科专属逻辑 (模块打包 + obstacle边界阻断)
    elif config.CURRENT_CONFIG == config.CONFIG_ARTS:
        logger.log("   📜 启用文科【模块打包】专属模式")
        pattern, fmt_name = detect_best_format(doc, start_p, end_p)
        if not pattern:
            print("   ⚠️ 跳过（无题目）。")
            return
        
        # ✨ 文科模式：同时支持括号小题号作为独立题目
        bracket_pattern = config.CONFIG_ARTS["formats"].get("括号小题")
        
        raw_starts = []
        paras = doc.Paragraphs
        for i in range(start_p, end_p + 1):
            try:
                rng = paras(i).Range
                if rng.Information(12): continue
                text = rng.Text.strip()
                if not text: continue
                
                is_obs, obs_pat = is_obstacle(rng)
                if is_obs: 
                    logger.scan_obstacle_filtered(i, text, obs_pat)
                    continue
                
                node_type = None
                if is_vip_header(text): 
                    node_type = "VIP"
                # ✨ 新增：识别地理/历史阅读材料（包含"据此回答X~Y题"等标记）
                elif is_reading_material(text):
                    node_type = "READING"
                elif re.match(pattern, text): 
                    node_type = "STD"
                # ✨ 新增：括号小题号也作为STD节点
                elif bracket_pattern and re.match(bracket_pattern, text):
                    node_type = "STD"
                
                if node_type:
                    logger.scan_node_found(i, node_type, text)
                    raw_starts.append({"idx": i, "type": node_type, "text": text})
            except: continue
            
        q_nodes = []
        skip_indices = set()
        
        for k in range(len(raw_starts)):
            if k in skip_indices: continue
            curr = raw_starts[k]
            q_nodes.append(curr)
            
            if curr['type'] in ['VIP', 'READING']:
                for j in range(k+1, len(raw_starts)):
                    nxt = raw_starts[j]
                    
                    # ✨ 关键：检查VIP/READING与下一个节点之间是否存在obstacle
                    # 如果存在，立即停止吞并，让obstacle后面的题目独立录入
                    has_obstacle_between = False
                    for mid_idx in range(curr['idx'] + 1, nxt['idx']):
                        try:
                            mid_text = paras(mid_idx).Range.Text.strip()
                            if mid_text:
                                mid_obs, _ = is_obstacle(paras(mid_idx).Range)
                                if mid_obs:
                                    has_obstacle_between = True
                                    break
                        except: continue
                    
                    if has_obstacle_between:
                        logger.merge_decision(curr, nxt, False, "中间存在obstacle边界")
                        break
                    
                    # VIP/READING节点：吞并后续STD节点（文科大题打包小题）
                    if nxt['type'] == 'STD':
                        skip_indices.add(j)
                        logger.merge_decision(curr, nxt, True, "文科大题打包小题" if curr['type'] == 'VIP' else "阅读材料打包小题")
                    elif nxt['type'] in ['VIP', 'READING']:
                        logger.merge_decision(curr, nxt, False, "遇到下一个阅读模块")
                        break
            
            # ✨ 新增：STD大题号（如"6." "7."）吞并后续的括号小题号
            elif curr['type'] == 'STD':
                curr_text = curr['text']
                # 检查当前节点是否是大题号格式（数字+点号）
                is_main_question = _is_numbered_main_question(curr_text)
                if is_main_question:
                    for j in range(k+1, len(raw_starts)):
                        nxt = raw_starts[j]
                        
                        # 检查中间是否有obstacle（装饰性图片除外）
                        has_obstacle_between = False
                        for mid_idx in range(curr['idx'] + 1, nxt['idx']):
                            try:
                                mid_text = paras(mid_idx).Range.Text.strip()
                                mid_obs, obs_type = is_obstacle(paras(mid_idx).Range)
                                if mid_obs:
                                    # 装饰性图片不阻断吞并
                                    if obs_type and "DecorativeHeader" in obs_type:
                                        continue
                                    has_obstacle_between = True
                                    break
                            except: continue
                        
                        if has_obstacle_between:
                            break
                        
                        # 只吞并括号小题号
                        if nxt['type'] == 'STD':
                            nxt_text = nxt['text']
                            is_bracket_subq = _is_bracket_subquestion(nxt_text)
                            if is_bracket_subq:
                                skip_indices.add(j)
                                logger.merge_decision(curr, nxt, True, "大题号吞并括号小题")
                            else:
                                # 遇到非括号小题号，停止吞并
                                break
                        else:
                            # 遇到VIP，停止吞并
                            break

    # --- 普通科目模式 ---
    else:
        pattern, fmt_name = detect_best_format(doc, start_p, end_p)
        if not pattern:
            print("   ⚠️ 跳过（无题目）。")
            return
        print(f"   ✅ 格式：{fmt_name}")
        q_nodes = []
        paras = doc.Paragraphs
        for i in range(start_p, end_p + 1):
            try:
                rng = paras(i).Range
                if rng.Information(12): continue
                text = rng.Text.strip()
                if not text: continue
                
                is_obs, obs_pat = is_obstacle(rng)
                if is_obs: 
                    logger.scan_obstacle_filtered(i, text, obs_pat)
                    continue
                    
                if re.match(pattern, text): 
                    logger.scan_node_found(i, "STD", text)
                    q_nodes.append({"idx": i, "type": "STD", "text": text})
            except: continue

        # --- 录入循环 ---
    if 'q_nodes' not in locals():
        q_nodes = []

    try:
        nodes = scan_wps_nodes(doc, start_p, end_p)
    except Exception:
        nodes = []

    question_units = []
    if nodes and config.CURRENT_CONFIG != config.CONFIG_ENGLISH:
        try:
            question_units = build_question_units_from_wps(
                doc_name=doc.Name,
                subject_name=config.CURRENT_CONFIG["name"],
                doc=doc,
                start_p=start_p,
                end_p=end_p,
                overlay_name=getattr(config, "CURRENT_SUBJECT_OVERLAY", None),
            )
        except Exception:
            question_units = []

    if not question_units and q_nodes and nodes:
        question_units = build_question_units_from_wps_spans(
            doc_name=doc.Name,
            subject_name=config.CURRENT_CONFIG["name"],
            nodes=nodes,
            q_nodes=q_nodes,
            end_p=end_p,
            overlay_name=getattr(config, "CURRENT_SUBJECT_OVERLAY", None),
        )

    if not question_units and q_nodes:
        question_units = []
        for i, node in enumerate(q_nodes):
            start_idx = node["idx"]
            unit_end = q_nodes[i + 1]["idx"] - 1 if i + 1 < len(q_nodes) else end_p
            question_units.append(
                type("FallbackUnit", (), {
                    "question_id": str(start_idx),
                    "question_type": "subjective",
                    "warnings": [],
                    "confidence": 1.0,
                    "node_type": node.get("type", "STD"),
                    "source_span": (start_idx, unit_end),
                    "preview": node.get("text", "")[:80],
                })()
            )

    total = len(question_units)
    logger.summary(total) 
    if total == 0: return
        
    if config.MANUAL_CONFIRM:
        user_choice = input("   ⌨️ 按回车开始 (输入 's' 跳过本章): ").strip().lower()
        if user_choice == 's':
            print("   ⏩ 已跳过。")
            return

    deferred_review_units = []

    for i, unit in enumerate(question_units):
        idx, real_end = unit.source_span
        node_type = unit.node_type
        current = i + 1
        overlay_name = (
            getattr(unit, "subject_overlay", None)
            or getattr(config, "CURRENT_SUBJECT_OVERLAY", None)
        )

        if should_pause_for_review(unit):
            print(f"   ⚠️ 第 {unit.question_id} 题置信度较低 ({unit.confidence:.2f})")
            print(f"      预览: {unit.preview}")
            if unit.warnings:
                warning_details = format_question_warning_details(unit.warnings)
                print(f"      风险说明: {'；'.join(warning_details)}")
            user_choice = input("      仅提醒，不会阻止录入。回车继续 / 输入 's' 跳过本题 / 输入 'q' 终止本章: ").strip().lower()
            if user_choice == "s":
                continue
            if user_choice == "q":
                print("   ⏹️ 已终止本章录入。")
                return
        elif should_mark_for_review(unit):
            deferred_review_units.append(unit)

        for p_idx in range(idx + 1, real_end + 1):
            try:
                rng = paras(p_idx).Range
                if rng.Information(12): continue # 表格内文本放行
                text = rng.Text.strip()
                
                # ✨ 终极极简逻辑：
                # 既然 next_q 已经是绝对正确的边界，在这个安全范围内，
                # 唯一需要我们提前刹车的，只有那些垃圾干扰项（如"建议用时"、"页码"等）
                is_obs, obs_type = is_obstacle(rng)
                if is_obs:
                    next_text = ""
                    if p_idx < real_end:
                        try:
                            next_text = paras(p_idx + 1).Range.Text.strip()
                        except Exception:
                            next_text = ""
                    excluded_media_role = classify_inter_question_media_boundary_for_context(
                        current_text=text,
                        next_text=next_text,
                        obstacle_type=obs_type,
                        overlay_name=overlay_name,
                    )
                    if excluded_media_role:
                        logger.log(
                            f"      ✂️ [媒体边界] 行{p_idx} [{node_type}] "
                            f"排除题间装饰图片 ({excluded_media_role})"
                        )
                        real_end = p_idx - 1
                        break
                    if should_ignore_inline_obstacle(unit, text, obs_type):
                        ignored_kind = describe_ignored_inline_obstacle(
                            text,
                            overlay_name=overlay_name,
                        )
                        ignored_icon = "🏷️" if ignored_kind == "题内标题" else "🖼️"
                        logger.log(
                            f"      {ignored_icon} [忽略截断] 行{p_idx} [{node_type}] {ignored_kind}不截断"
                        )
                        continue
                    logger.log(f"      ✂️ [提前截断] 行{p_idx} [{node_type}] 命中 obstacle: {str(obs_type)[:40]}")
                    real_end = p_idx - 1
                    break
            except: continue
        
        if real_end < idx: real_end = idx
        if real_end > end_p: real_end = end_p

        temp_doc = None
        try:
            contains_table = _range_contains_table(doc, idx, real_end)
            if contains_table:
                logger.log(f"      📋 [原生表格] 行{idx}-{real_end} 直接选中题干、表格和选项")
            start_pos = paras(idx).Range.Start

            # ✨ 新增：剥离题号前粘连的下划线、破折号和空白符
            first_text = paras(idx).Range.Text
            if first_text:
                match = __import__('re').match(r"^([\s_—\-]+)", first_text)
                if match:
                    start_pos += len(match.group(1))
                    first_text = first_text[len(match.group(1)):]

                if not should_preserve_question_prefix(doc.Name):
                    stripped_text = strip_question_noise_prefix(first_text)
                    if stripped_text != first_text:
                        start_pos += len(first_text) - len(stripped_text)

            end_pos = paras(real_end).Range.End
            rng = doc.Range(start_pos, end_pos)
            excluded_indices = collect_input_excluded_paragraph_indices(
                paras,
                idx,
                real_end,
                overlay_name=overlay_name,
            )
            if excluded_indices:
                if not is_filtered_input_copy_safe(
                    contains_table=contains_table,
                    has_media=bool(getattr(unit, "media_blocks", [])),
                ):
                    raise RuntimeError(
                        "题内专练标题与原生表格共存，无法安全生成过滤副本，已停止本题录入"
                    )
                temp_doc, removed_indices = create_filtered_input_document(
                    doc,
                    rng,
                    overlay_name=overlay_name,
                )
                if len(removed_indices) != len(excluded_indices):
                    raise RuntimeError(
                        "临时录入副本中的题内专练标题数量与原选区不一致"
                    )
                temp_doc.Activate()
                rng = temp_doc.Content
                logger.log(
                    f"      🧹 [临时副本] 已剔除 {len(removed_indices)} 个题内专练标题"
                )
            rng.Select()
            rng.Application.ActiveWindow.ScrollIntoView(rng)
            print(f"   ▶️ 录入 {current}{total}...")
            pyautogui.press(config.KEY_TO_PRESS)
            time.sleep(config.WAIT_TIME)
        except Exception as exc:
            logger.log(f"      ❌ [录入失败] 第 {unit.question_id} 题：{exc}")
            print(f"   ❌ 第 {unit.question_id} 题录入失败：{exc}")
        finally:
            if temp_doc is not None:
                try:
                    temp_doc.Close(False)
                except Exception:
                    pass
                try:
                    doc.Activate()
                except Exception:
                    pass

    print(f"   🎉 完成！")
    if deferred_review_units:
        print(f"   📝 本章已后台标记 {len(deferred_review_units)} 题待复核，不影响本次录入。")
        for item in format_deferred_review_items(deferred_review_units[:5]):
            print(f"      - {item}")
        if len(deferred_review_units) > 5:
            print(f"      - 其余 {len(deferred_review_units) - 5} 题已省略显示。")
