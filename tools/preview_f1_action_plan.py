from __future__ import annotations

import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Callable

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.cli_output import configure_utf8_stdio
from shared_core.document_preflight import build_preflight_bundle


def _normalize_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKC", text or "")
    return re.sub(r"[\s\x00-\x1f\x7f]+", "", normalized)


def _stable_anchor(preview: str) -> str:
    return _normalize_text(preview)[:12]


def _legacy_preview_anchor(preview: str) -> str:
    """旧计划缺少 start_preview 时，只截取明确的材料题范围句。"""

    normalized = unicodedata.normalize("NFKC", preview or "")
    material_match = re.match(
        r"^(?P<material>.*?(?:完成|回答)\s*\d+\s*(?:~|-|—|至)\s*\d+\s*题[。.!?！？])"
        r"\s+(?=(?:[/★]\s*)?\d{1,3}\s*\.)",
        normalized,
    )
    if material_match:
        return _stable_anchor(material_match.group("material"))
    return _stable_anchor(preview)


def _anchor_is_near_start(text: str, anchor: str) -> bool:
    position = text.find(anchor)
    return 0 <= position <= 3


def _paragraph_range(doc, paragraphs, start: int, end: int):
    start_pos = paragraphs(start).Range.Start
    end_pos = paragraphs(end).Range.End
    return doc.Range(start_pos, end_pos)


def _find_matching_range(doc, paragraphs, start: int, end: int, anchor: str):
    matches = []
    span = end - start

    # ponytail: 可扫描整篇，但只接受题目开头唯一命中的选区。
    for candidate_start in range(1, paragraphs.Count - span + 1):
        candidate_end = candidate_start + span
        candidate_range = _paragraph_range(
            doc, paragraphs, candidate_start, candidate_end
        )
        candidate_text = _normalize_text(getattr(candidate_range, "Text", ""))
        if _anchor_is_near_start(candidate_text, anchor):
            offset = candidate_start - start
            matches.append((candidate_start, candidate_end, candidate_range, offset))

    if len(matches) > 1:
        raise ValueError("整份 WPS 文档中有多个题目开头相同的选区，无法安全确定位置")
    return matches[0] if matches else None


def _find_unique_anchor_paragraph(paragraphs, anchor: str, *, start_index: int = 1):
    matches = []
    for index in range(start_index, paragraphs.Count + 1):
        paragraph_range = paragraphs(index).Range
        text = _normalize_text(getattr(paragraph_range, "Text", ""))
        if _anchor_is_near_start(text, anchor):
            matches.append((index, paragraph_range))
    if len(matches) > 1:
        raise ValueError(f"WPS 文档中有多个段落命中锚点“{anchor}”")
    return matches[0] if matches else None


def _select_table_action_range(doc, paragraphs, action: dict, source_ref: dict):
    start_anchor = _stable_anchor(source_ref.get("start_preview", ""))
    end_anchor = _stable_anchor(source_ref.get("end_preview", ""))
    if not start_anchor or not end_anchor:
        raise ValueError(f"第 {action['sequence']} 个表格动作缺少题目开头或结尾锚点")

    start_match = _find_unique_anchor_paragraph(paragraphs, start_anchor)
    if start_match is None:
        raise ValueError(f"第 {action['sequence']} 个表格动作找不到开头锚点“{start_anchor}”")
    wps_start, start_range = start_match
    end_match = _find_unique_anchor_paragraph(
        paragraphs,
        end_anchor,
        start_index=wps_start,
    )
    if end_match is None:
        raise ValueError(f"第 {action['sequence']} 个表格动作找不到结尾锚点“{end_anchor}”")
    wps_end, end_range = end_match

    selected_range = doc.Range(start_range.Start, end_range.End)
    expected_tables = len(source_ref.get("table_indexes") or [])
    actual_tables = int(getattr(getattr(selected_range, "Tables", None), "Count", 0))
    if actual_tables != expected_tables:
        raise ValueError(
            f"第 {action['sequence']} 个动作计划包含 {expected_tables} 个原生表格，"
            f"但 WPS 选区中的表格数量为 {actual_tables}，已停止"
        )
    return selected_range, wps_start, wps_end


def validate_preview_plan(plan: dict) -> None:
    if plan.get("mode") != "preview_only":
        raise ValueError("只允许预览 preview_only ActionPlan")
    if plan.get("execution_enabled") is not False:
        raise ValueError("ActionPlan.execution_enabled 必须为 false")

    actions = plan.get("actions")
    if not isinstance(actions, list) or not actions:
        raise ValueError("ActionPlan 没有可预览的 F1 动作")
    if [action.get("sequence") for action in actions] != list(range(1, len(actions) + 1)):
        raise ValueError("ActionPlan 动作序号不连续")
    if any(action.get("key") != "F1" for action in actions):
        raise ValueError("试点只允许预览 F1 动作")


def select_action_range(
    doc,
    action: dict,
    *,
    output: Callable[[str], None] = print,
):
    """把单个 ActionPlan 动作绑定到 WPS Range 并选中。"""

    paragraphs = doc.Paragraphs
    source_ref = action.get("source_ref") or {}
    start = source_ref.get("paragraph_start")
    end = source_ref.get("paragraph_end")
    if not isinstance(start, int) or not isinstance(end, int) or start < 1 or end < start:
        raise ValueError(f"第 {action['sequence']} 个动作的段落范围无效")

    wps_start = start
    wps_end = end
    if source_ref.get("table_indexes"):
        selected_range, wps_start, wps_end = _select_table_action_range(
            doc,
            paragraphs,
            action,
            source_ref,
        )
        selected_range.Select()
        try:
            selected_range.Application.ActiveWindow.ScrollIntoView(selected_range)
        except Exception:
            pass
        return selected_range, wps_start, wps_end

    start_preview = source_ref.get("start_preview")
    preview_anchor = (
        _stable_anchor(start_preview)
        if start_preview
        else _legacy_preview_anchor(action.get("preview", ""))
    )
    selected_range = None
    if wps_end <= paragraphs.Count:
        selected_range = _paragraph_range(doc, paragraphs, wps_start, wps_end)

    selected_text = _normalize_text(
        getattr(selected_range, "Text", "") if selected_range is not None else ""
    )
    if preview_anchor and not _anchor_is_near_start(selected_text, preview_anchor):
        nearby = _find_matching_range(
            doc, paragraphs, wps_start, wps_end, preview_anchor
        )
        if nearby is None:
            raise ValueError(
                f"第 {action['sequence']} 个动作在计划段落 {start}-{end} "
                f"和整份 WPS 文档中都没找到唯一的题目开头“{preview_anchor}”，已停止"
            )
        wps_start, wps_end, selected_range, offset = nearby
        offset_text = f"+{offset}" if offset > 0 else str(offset)
        output(
            f"⚠️ 第 {action['sequence']} 组在 WPS 中偏移 {offset_text} 段，"
            "已按题目开头找到整篇中的唯一位置。"
        )

    if selected_range is None:
        raise ValueError(
            f"第 {action['sequence']} 个动作需要 WPS 段落 {wps_start}-{wps_end}，"
            f"但当前文档只有 {paragraphs.Count} 段"
        )

    selected_range.Select()
    try:
        selected_range.Application.ActiveWindow.ScrollIntoView(selected_range)
    except Exception:
        pass
    return selected_range, wps_start, wps_end


def preview_actions(
    doc,
    plan: dict,
    *,
    confirm: Callable[[str], str] = input,
    output: Callable[[str], None] = print,
) -> dict[str, int | str]:
    """只在 WPS 中依次选区；永远不执行 F1。"""

    validate_preview_plan(plan)
    selected_count = 0
    for action in plan["actions"]:
        _selected_range, wps_start, wps_end = select_action_range(
            doc,
            action,
            output=output,
        )

        selected_count += 1
        question_ids = "、".join(str(item) for item in action.get("question_ids", []))
        output(
            f"\n[{action['sequence']}/{len(plan['actions'])}] "
            f"题号 {question_ids}，WPS 段落 {wps_start}-{wps_end}"
        )
        output(f"计划预览：{action.get('preview', '')}")
        choice = confirm("请查看 WPS 选区：正确按回车继续，输入 s 停止：").strip().lower()
        if choice == "s":
            return {
                "status": "stopped",
                "selected_actions": selected_count,
                "keypress_count": 0,
            }

    return {
        "status": "completed",
        "selected_actions": selected_count,
        "keypress_count": 0,
    }


def get_active_document():
    question_input_dir = PROJECT_ROOT / "墨痕快刀"
    if str(question_input_dir) not in sys.path:
        sys.path.insert(0, str(question_input_dir))
    from wps_helper import get_active_wps

    wps = get_active_wps()
    if not wps:
        raise RuntimeError("没有连接到 WPS，请先打开原题文档并加载墨痕题库插件")
    return wps.ActiveDocument


def main() -> int:
    configure_utf8_stdio()
    doc = get_active_document()
    if not bool(getattr(doc, "Saved", True)):
        raise RuntimeError("当前 WPS 文档有未保存修改，已停止；请关闭修改后重新打开原题")
    source_path = Path(doc.FullName)
    bundle = build_preflight_bundle(source_path, include_docling=False)
    plan = bundle["plan"]
    validate_preview_plan(plan)

    print(f"当前文档：{doc.Name}")
    print(f"检测到 {len(plan['actions'])} 个 F1 题块。")
    print("本程序只依次框选，不会按 F1，也不会修改原文档。")
    if input("按回车开始无按键预览，输入 q 退出：").strip().lower() == "q":
        return 0

    receipt = preview_actions(doc, plan)
    print("\n预览结果：")
    print(json.dumps(receipt, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"❌ 已停止：{exc}")
        raise SystemExit(1)
