from __future__ import annotations

import argparse
import ctypes
import json
import sys
import time
from pathlib import Path
from typing import Callable

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from shared_core.cli_output import configure_utf8_stdio
from shared_core.document_preflight import build_preflight_bundle
from tools.preview_f1_action_plan import (
    bind_action_plan_to_wps,
    get_active_document,
    select_action_range,
    validate_preview_plan,
)

APPROVED_PILOT_SHA256 = (
    "021afd8fc4a63ac69ca54f861fa29a0a8e95480c537f6f2acd1d9b18f2441921"
)


def validate_execution_profile(profile: dict) -> None:
    table_count = (profile.get("fingerprint") or {}).get("table_count", 0)
    if table_count:
        raise RuntimeError("当前受控 F1 执行器暂不支持原生表格文档")
    source_sha256 = (profile.get("source") or {}).get("sha256")
    if source_sha256 != APPROVED_PILOT_SHA256:
        raise RuntimeError("当前文档不是已完成 8/8 WPS 选区验收的试点文档")


def _reactivate_selection(doc, selected_range) -> None:
    try:
        doc.Application.Activate()
    except Exception:
        pass
    try:
        doc.Activate()
    except Exception:
        pass
    try:
        doc.ActiveWindow.Activate()
    except Exception:
        pass
    if sys.platform == "win32":
        try:
            hwnd = int(doc.ActiveWindow.Hwnd)
            ctypes.windll.user32.ShowWindow(hwnd, 5)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
        except Exception:
            pass
    selected_range.Select()


def execute_actions(
    doc,
    plan: dict,
    *,
    max_actions: int | None = 1,
    confirm: Callable[[str], str] = input,
    press_f1: Callable[[], None],
    sleep: Callable[[float], None] = time.sleep,
    wait_seconds: float = 0.5,
    output: Callable[[str], None] = print,
) -> dict[str, int | str]:
    """每题二次人工确认，并且只在精确输入 f1 后按一次 F1。"""

    validate_preview_plan(plan)
    if max_actions is not None and max_actions < 1:
        raise ValueError("max_actions 必须大于等于 1")

    actions = plan["actions"] if max_actions is None else plan["actions"][:max_actions]
    selected_count = 0
    executed_count = 0

    for action in actions:
        selected_range, wps_start, wps_end = select_action_range(
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
        choice = confirm(
            "请检查 WPS 选区：输入 f1 并回车才会录入，输入 s 停止："
        ).strip().lower()
        if choice != "f1":
            return {
                "status": "stopped",
                "selected_actions": selected_count,
                "executed_actions": executed_count,
                "keypress_count": executed_count,
            }

        # 控制台输入会抢走焦点；真正按键前必须重新激活 WPS 并恢复选区。
        _reactivate_selection(doc, selected_range)
        sleep(0.2)
        press_f1()
        executed_count += 1
        sleep(wait_seconds)

        is_canary_end = (
            max_actions is not None
            and executed_count == len(actions)
            and len(actions) < len(plan["actions"])
        )
        next_text = "结束本次首题试录" if is_canary_end else "继续"
        check = confirm(
            f"请检查插件中的临时题目：正确按回车{next_text}，"
            "输入 s 停止："
        ).strip().lower()
        if check == "s":
            status = (
                "completed"
                if executed_count == len(plan["actions"])
                else "stopped"
            )
            return {
                "status": status,
                "selected_actions": selected_count,
                "executed_actions": executed_count,
                "keypress_count": executed_count,
            }

    status = "completed" if max_actions is None else "canary_completed"
    return {
        "status": status,
        "selected_actions": selected_count,
        "executed_actions": executed_count,
        "keypress_count": executed_count,
    }


def _press_f1() -> None:
    import pyautogui

    pyautogui.press("f1")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按 ActionPlan 逐题确认并受控执行 F1。默认只试录第 1 题。"
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="执行所有动作；仍需每题输入 f1 确认",
    )
    return parser.parse_args()


def main() -> int:
    configure_utf8_stdio()
    args = _parse_args()
    doc = get_active_document()
    if not bool(getattr(doc, "Saved", True)):
        raise RuntimeError("当前 WPS 文档有未保存修改，已停止")

    source_path = Path(doc.FullName)
    bundle = build_preflight_bundle(source_path, include_docling=False)
    plan = bundle["plan"]
    validate_preview_plan(plan)
    validate_execution_profile(bundle["profile"])
    bound_plan = bind_action_plan_to_wps(doc, plan)

    action_limit = None if args.all else 1
    mode_text = "全部题组" if args.all else "第 1 个题组试录"
    phrase = "执行全部F1" if args.all else "执行首题F1"
    print(f"当前文档：{doc.Name}")
    print(f"已生成并绑定 {len(bound_plan['actions'])} 个 F1 动作，本次：{mode_text}。")
    print("不会保存 WPS 文档；插件中是否最终保存仍由人工决定。")
    if input(f"请输入“{phrase}”继续，其他输入退出：").strip() != phrase:
        print("已退出，未按 F1。")
        return 0

    receipt = execute_actions(
        doc,
        bound_plan,
        max_actions=action_limit,
        press_f1=_press_f1,
    )
    print("\n执行回执：")
    print(json.dumps(receipt, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"❌ 已停止：{exc}")
        raise SystemExit(1)
