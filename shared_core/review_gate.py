from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


def derive_review_status_path(answer_doc_path: str | Path) -> str:
    path = Path(answer_doc_path)
    return str(path.with_name(f"{path.stem}_审核状态.json"))


def _utc_now() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def _capture_file_signature(answer_doc_path: str | Path) -> dict:
    path = Path(answer_doc_path)
    stat = path.stat()
    return {
        "path": str(path.resolve()),
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }


def _read_status_file(status_path: Path) -> dict | None:
    if not status_path.exists():
        return None
    return json.loads(status_path.read_text(encoding="utf-8"))


def _write_status_file(status_path: Path, payload: dict) -> str:
    status_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(status_path)


def initialize_review_status(
    answer_doc_path: str | Path,
    *,
    report_path: str | None = None,
    report=None,
) -> str:
    answer_path = Path(answer_doc_path)
    status_path = Path(derive_review_status_path(answer_path))
    payload = {
        "status": "pending",
        "updated_at": _utc_now(),
        "answer_doc_path": str(answer_path.resolve()),
        "answer_signature": _capture_file_signature(answer_path),
        "report_path": str(Path(report_path).resolve()) if report_path else None,
        "report_summary": getattr(report, "summary", None),
        "reviewer": "",
        "note": "清洗后自动重置为待审核",
    }
    return _write_status_file(status_path, payload)


def update_review_status(
    answer_doc_path: str | Path,
    *,
    status: str,
    reviewer: str = "",
    note: str = "",
) -> str:
    if status not in {"approved", "rejected", "pending"}:
        raise ValueError(f"不支持的审核状态: {status}")

    answer_path = Path(answer_doc_path)
    status_path = Path(derive_review_status_path(answer_path))
    current = _read_status_file(status_path) or {}
    current.update(
        {
            "status": status,
            "updated_at": _utc_now(),
            "answer_doc_path": str(answer_path.resolve()),
            "answer_signature": _capture_file_signature(answer_path),
            "reviewer": reviewer.strip(),
            "note": note.strip(),
        }
    )
    return _write_status_file(status_path, current)


def get_review_gate_result(answer_doc_path: str | Path) -> dict:
    answer_path = Path(answer_doc_path)
    status_path = Path(derive_review_status_path(answer_path))
    if not answer_path.exists():
        return {
            "allowed": False,
            "status": "missing_file",
            "reason": "未找到已清洗答案文件，请先检查工作台文件路径或重新生成标准答案。",
            "status_path": str(status_path),
        }

    current_signature = _capture_file_signature(answer_path)
    status_payload = _read_status_file(status_path)

    if status_payload is None:
        return {
            "allowed": False,
            "status": "missing",
            "reason": "未找到自动检查状态，请先走答案清洗+自动检查。",
            "status_path": str(status_path),
        }

    stored_signature = status_payload.get("answer_signature") or {}
    if stored_signature != current_signature:
        return {
            "allowed": False,
            "status": "stale",
            "reason": "答案文件在自动检查后发生变更，原放行结果已失效，请重新清洗并自动检查。",
            "status_path": str(status_path),
        }

    current_status = str(status_payload.get("status") or "pending")
    if current_status != "approved":
        reason_map = {
            "pending": "答案仍处于待检查状态，系统尚未放行录入。",
            "rejected": "答案未通过自动检查，请先根据审核清单修正后重跑。",
        }
        return {
            "allowed": False,
            "status": current_status,
            "reason": reason_map.get(current_status, f"检查状态为 {current_status}，暂不允许录入。"),
            "status_path": str(status_path),
        }

    return {
        "allowed": True,
        "status": current_status,
        "reason": "自动检查已通过，可以继续录入。",
        "status_path": str(status_path),
        "status_payload": status_payload,
    }


def format_review_gate_message(result: dict) -> str:
    status_path = result.get("status_path", "")
    suffix = f" 状态文件: {status_path}" if status_path else ""
    return f"{result.get('reason', '自动检查门禁未通过。')}{suffix}"
