import json
import os
import re
import sys
import importlib.util
from pathlib import Path

from shared_core import (
    build_answer_units_from_docx,
    build_question_units_from_docx,
    build_question_units_from_wps,
    build_review_report,
    derive_review_status_path,
    format_review_gate_message,
    export_review_report,
    get_review_gate_result,
    initialize_review_status,
    map_answers,
    update_review_status,
)


PROJECT_ROOT = Path(__file__).resolve().parent
QUESTION_DIR = PROJECT_ROOT / "墨痕快刀"
FORMAT_DIR = PROJECT_ROOT / "格式处理"
ANSWER_DIR = PROJECT_ROOT / "答案录入"
REGRESSION_DIR = PROJECT_ROOT / "回归样本"
REGRESSION_MANIFEST_PATH = REGRESSION_DIR / "样本清单.json"
WORKSPACE_CONFIG_PATH = PROJECT_ROOT / "工作台路径配置.json"

WORKSPACE_CONFIG_FIELDS = {
    "question_dirs": "题目目录",
    "raw_answer_dirs": "原始答案目录",
    "clean_answer_dirs": "已清洗答案目录",
    "review_dirs": "审核单目录",
}

DEFAULT_WORKSPACE_CONFIG = {
    "question_dirs": ["墨痕快刀\\待录入文档"],
    "raw_answer_dirs": ["格式处理\\待清洗文件"],
    "clean_answer_dirs": ["格式处理\\待清洗文件", "格式处理\\已清洗文件"],
    "review_dirs": ["格式处理\\待清洗文件"],
}

WORKSPACE_STAGE_ORDER = {
    "可录答案": 0,
    "待清洗": 1,
    "自动检查未通过": 2,
    "清洗结果已过期": 3,
    "待自动检查": 4,
    "缺状态文件": 5,
    "缺题目": 6,
    "缺原答案": 7,
    "待补材料": 8,
}


def derive_cleaned_output_path(answer_doc_path: str | os.PathLike[str]) -> str:
    path = Path(answer_doc_path)
    stem = path.stem
    if stem.endswith("_已清洗"):
        return str(path)
    suffix = ".docx" if path.suffix.lower() == ".doc" else path.suffix
    return str(path.with_name(f"{stem}_已清洗{suffix}"))


def derive_review_report_path(answer_doc_path: str | os.PathLike[str]) -> str:
    standardized_path = Path(derive_cleaned_output_path(answer_doc_path))
    return str(standardized_path.with_name(f"{standardized_path.stem}_审核清单.md"))


def has_blocking_review_issues(report) -> bool:
    return any(issue.severity == "error" for issue in report.issues)


def apply_auto_review_decision(standardized_path: str, report, *, reviewer: str = "system") -> str:
    if has_blocking_review_issues(report):
        return update_review_status(
            standardized_path,
            status="rejected",
            reviewer=reviewer,
            note="自动检查未通过",
        )
    return update_review_status(
        standardized_path,
        status="approved",
        reviewer=reviewer,
        note="自动检查通过",
    )


def _normalize_workspace_key(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r"^\s*【空白试卷】\s*", "", stem)
    stem = re.sub(r"_已清洗_审核清单$", "", stem)
    stem = re.sub(r"_已清洗_审核状态$", "", stem)
    stem = re.sub(r"_审核清单$", "", stem)
    stem = re.sub(r"_审核状态$", "", stem)
    stem = re.sub(r"_已清洗$", "", stem)
    changed = True
    while changed:
        changed = False
        for suffix in ["标准答案", "原始答案", "原答案", "答案", "题目", "试题", "解析"]:
            if stem.endswith(suffix):
                stem = stem[: -len(suffix)]
                changed = True
    stem = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", stem).lower()
    return stem or "untitled"


def _normalize_workspace_config_data(config_data: dict | None) -> dict:
    normalized = dict(DEFAULT_WORKSPACE_CONFIG)
    for field_name in WORKSPACE_CONFIG_FIELDS:
        raw_value = (config_data or {}).get(field_name, normalized[field_name])
        if isinstance(raw_value, str):
            normalized[field_name] = [raw_value]
        elif isinstance(raw_value, list):
            normalized[field_name] = [str(item).strip() for item in raw_value if str(item).strip()]
        else:
            normalized[field_name] = list(DEFAULT_WORKSPACE_CONFIG[field_name])
        if not normalized[field_name]:
            normalized[field_name] = list(DEFAULT_WORKSPACE_CONFIG[field_name])
    return normalized


def _write_workspace_config(config_path: Path, config_data: dict) -> str:
    config_path.write_text(
        json.dumps(config_data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return str(config_path)


def ensure_workspace_config(config_path: Path | None = None) -> str:
    target_path = Path(config_path or WORKSPACE_CONFIG_PATH)
    if target_path.exists():
        return str(target_path)
    _write_workspace_config(target_path, DEFAULT_WORKSPACE_CONFIG)
    return str(target_path)


def _resolve_workspace_dir(path_text: str, *, project_root: Path | None = None) -> Path:
    root = Path(project_root or PROJECT_ROOT)
    candidate = Path(path_text)
    if not candidate.is_absolute():
        candidate = root / candidate
    return candidate.resolve()


def load_workspace_config(
    config_path: Path | None = None,
    *,
    project_root: Path | None = None,
) -> dict:
    target_path = Path(config_path or WORKSPACE_CONFIG_PATH)
    if not target_path.exists():
        ensure_workspace_config(target_path)

    raw_data = json.loads(target_path.read_text(encoding="utf-8"))
    normalized = _normalize_workspace_config_data(raw_data)
    resolved_dirs = {
        field_name: [_resolve_workspace_dir(path_text, project_root=project_root) for path_text in normalized[field_name]]
        for field_name in WORKSPACE_CONFIG_FIELDS
    }
    return {
        "config_path": str(target_path),
        "project_root": str(Path(project_root or PROJECT_ROOT)),
        "raw": normalized,
        "resolved": resolved_dirs,
    }


def _iter_visible_files(directory: Path):
    if not directory.exists() or not directory.is_dir():
        return []
    return [
        path for path in sorted(directory.iterdir())
        if path.is_file() and not path.name.startswith("~$")
    ]


def _match_workspace_files(all_files: list[Path], field_name: str) -> list[Path]:
    if field_name == "review_dirs":
        return [path for path in all_files if path.suffix.lower() == ".md" and path.stem.endswith("_审核清单")]
    if field_name == "clean_answer_dirs":
        return [
            path for path in all_files
            if path.suffix.lower() in {".doc", ".docx"} and "_已清洗" in path.stem
        ]
    if field_name == "raw_answer_dirs":
        return [
            path for path in all_files
            if path.suffix.lower() in {".doc", ".docx"}
            and "_已清洗" not in path.stem
            and "_审核清单" not in path.stem
        ]
    return [
        path for path in all_files
        if path.suffix.lower() in {".doc", ".docx"} and "答案" not in path.stem
    ]


def _collect_workspace_bucket(directory: Path, field_name: str) -> dict:
    all_files = _iter_visible_files(directory)
    matched = _match_workspace_files(all_files, field_name)

    return {
        "path": str(directory),
        "exists": directory.exists(),
        "count": len(matched),
        "files": [path.name for path in matched[:5]],
    }


def _iter_workspace_field_files(workspace_config: dict, field_name: str) -> list[Path]:
    files: list[Path] = []
    for directory in workspace_config["resolved"][field_name]:
        files.extend(_match_workspace_files(_iter_visible_files(directory), field_name))
    return sorted(files)


def collect_workspace_overview(config: dict | None = None) -> dict:
    workspace_config = config or load_workspace_config()
    buckets = {}
    for field_name in WORKSPACE_CONFIG_FIELDS:
        buckets[field_name] = [
            _collect_workspace_bucket(directory, field_name)
            for directory in workspace_config["resolved"][field_name]
        ]
    return {
        "config_path": workspace_config["config_path"],
        "buckets": buckets,
    }


def print_workspace_overview(config: dict | None = None) -> dict:
    overview = collect_workspace_overview(config)
    print("\n" + "=" * 60)
    print("🧭 工作台入口")
    print("=" * 60)
    print(f"配置文件: {overview['config_path']}")
    for field_name, label in WORKSPACE_CONFIG_FIELDS.items():
        print(f"\n[{label}]")
        for bucket in overview["buckets"][field_name]:
            status = "存在" if bucket["exists"] else "缺失"
            print(f"- {bucket['path']} | {status} | {bucket['count']} 个文件")
            for file_name in bucket["files"]:
                print(f"  · {file_name}")
    print("-" * 60)
    return overview


def _build_workspace_record(records_by_key: dict, key: str, display_name: str) -> dict:
    record = records_by_key.get(key)
    if record is not None:
        return record
    record = {
        "key": key,
        "display_name": display_name,
        "question_doc_path": None,
        "raw_answer_path": None,
        "clean_answer_path": None,
        "review_report_path": None,
        "review_status_path": None,
        "gate_status": None,
        "gate_reason": "",
        "stage": "待补材料",
        "next_action": "补充题目或答案文档",
    }
    records_by_key[key] = record
    return record


def _classify_workspace_record(record: dict) -> tuple[str, str]:
    if record.get("clean_answer_path"):
        status = record.get("gate_status") or "missing"
        stage_map = {
            "approved": ("可录答案", "直接录入标准答案"),
            "rejected": ("自动检查未通过", "查看审核清单并重新清洗"),
            "stale": ("清洗结果已过期", "重新执行答案清洗+自动检查"),
            "pending": ("待自动检查", "重新执行自动检查或手动覆盖"),
            "missing": ("缺状态文件", "重新执行答案清洗+自动检查"),
            "missing_file": ("待补材料", "补回已清洗答案文件"),
        }
        return stage_map.get(status, ("待补材料", f"检查状态异常：{status}"))
    if record.get("question_doc_path") and record.get("raw_answer_path"):
        return "待清洗", "执行答案清洗+自动检查"
    if record.get("raw_answer_path"):
        return "缺题目", "补充题目文档后再清洗"
    if record.get("question_doc_path"):
        return "缺原答案", "补充原始答案文档"
    return "待补材料", "补充题目或答案文档"


def collect_workspace_records(config: dict | None = None) -> list[dict]:
    workspace_config = config or load_workspace_config()
    records_by_key: dict[str, dict] = {}

    for path in _iter_workspace_field_files(workspace_config, "question_dirs"):
        key = _normalize_workspace_key(path.name)
        record = _build_workspace_record(records_by_key, key, path.stem)
        record["question_doc_path"] = record["question_doc_path"] or str(path)

    for path in _iter_workspace_field_files(workspace_config, "raw_answer_dirs"):
        key = _normalize_workspace_key(path.name)
        record = _build_workspace_record(records_by_key, key, path.stem)
        record["raw_answer_path"] = record["raw_answer_path"] or str(path)

    for path in _iter_workspace_field_files(workspace_config, "clean_answer_dirs"):
        key = _normalize_workspace_key(path.name)
        record = _build_workspace_record(records_by_key, key, path.stem)
        record["clean_answer_path"] = record["clean_answer_path"] or str(path)
        derived_report_path = Path(derive_review_report_path(path))
        derived_status_path = Path(derive_review_status_path(path))
        if derived_report_path.exists():
            record["review_report_path"] = str(derived_report_path)
        if derived_status_path.exists():
            record["review_status_path"] = str(derived_status_path)
        gate_result = get_review_gate_result(path)
        record["gate_status"] = gate_result["status"]
        record["gate_reason"] = gate_result["reason"]
        if gate_result.get("status_path"):
            record["review_status_path"] = gate_result["status_path"]

    for path in _iter_workspace_field_files(workspace_config, "review_dirs"):
        key = _normalize_workspace_key(path.name)
        record = _build_workspace_record(records_by_key, key, path.stem)
        record["review_report_path"] = record["review_report_path"] or str(path)

    records = list(records_by_key.values())
    for record in records:
        record["display_name"] = (
            Path(
                record.get("question_doc_path")
                or record.get("raw_answer_path")
                or record.get("clean_answer_path")
                or record.get("review_report_path")
                or record["display_name"]
            ).stem
        )
        stage, next_action = _classify_workspace_record(record)
        record["stage"] = stage
        record["next_action"] = next_action

    records.sort(
        key=lambda item: (
            WORKSPACE_STAGE_ORDER.get(item["stage"], 99),
            item["display_name"],
        )
    )
    return records


def summarize_workspace_records(records: list[dict]) -> dict:
    summary = {
        "total": len(records),
        "可录答案": 0,
        "待清洗": 0,
        "自动检查未通过": 0,
        "清洗结果已过期": 0,
        "待自动检查": 0,
        "缺状态文件": 0,
        "缺题目": 0,
        "缺原答案": 0,
        "待补材料": 0,
    }
    for record in records:
        summary[record["stage"]] = summary.get(record["stage"], 0) + 1
    return summary


def print_workspace_task_board(config: dict | None = None, *, limit: int | None = 12) -> dict:
    records = collect_workspace_records(config)
    summary = summarize_workspace_records(records)
    print("\n" + "=" * 60)
    print("📋 工作台任务面板")
    print("=" * 60)
    print(
        " | ".join(
            [
                f"总任务 {summary['total']}",
                f"可录答案 {summary['可录答案']}",
                f"待清洗 {summary['待清洗']}",
                f"未通过 {summary['自动检查未通过']}",
                f"已过期 {summary['清洗结果已过期']}",
                f"缺题目 {summary['缺题目']}",
                f"缺答案 {summary['缺原答案']}",
            ]
        )
    )
    if not records:
        print("- 当前工作台没有可识别任务")
        print("-" * 60)
        return {"records": records, "summary": summary}

    display_records = records if limit is None else records[:limit]
    for idx, record in enumerate(display_records, 1):
        marker_parts = []
        if record.get("question_doc_path"):
            marker_parts.append("题")
        if record.get("raw_answer_path"):
            marker_parts.append("原答")
        if record.get("clean_answer_path"):
            marker_parts.append("已清洗")
        gate_marker = record.get("gate_status")
        if gate_marker:
            marker_parts.append(f"状态:{gate_marker}")
        marker_text = " / ".join(marker_parts) if marker_parts else "无文件"
        print(f"[{idx}] {record['display_name']} | {record['stage']} | {marker_text}")
        print(f"    下一步：{record['next_action']}")
    if limit is not None and len(records) > len(display_records):
        print(f"... 其余 {len(records) - len(display_records)} 条可在选择时继续查看")
    print("-" * 60)
    return {"records": records, "summary": summary}


def _get_workspace_action_candidates(action: str, config: dict | None = None) -> list[dict]:
    records = collect_workspace_records(config)
    if action == "clean":
        return [record for record in records if record.get("question_doc_path") and record.get("raw_answer_path")]
    if action == "input":
        return [record for record in records if record.get("clean_answer_path")]
    return records


def prompt_select_workspace_record(action: str, config: dict | None = None) -> dict | None:
    candidates = _get_workspace_action_candidates(action, config)
    action_label = {"clean": "答案清洗", "input": "标准答案录入"}.get(action, "工作台任务")
    if not candidates:
        print(f"⚠️ 工作台里没有可用于{action_label}的候选文件。")
        return None

    print("\n" + "=" * 60)
    print(f"📌 选择{action_label}任务")
    print("=" * 60)
    for idx, record in enumerate(candidates, 1):
        print(f"[{idx}] {record['display_name']} | {record['stage']}")
        if action == "clean":
            print(f"    题目：{record.get('question_doc_path')}")
            print(f"    原答：{record.get('raw_answer_path')}")
        else:
            print(f"    已清洗：{record.get('clean_answer_path')}")
            if record.get("gate_status"):
                print(f"    状态：{record['gate_status']} | {record['gate_reason']}")
    print("-" * 60)
    raw = input("👉 输入任务编号（回车取消）: ").strip()
    if not raw:
        return None
    index = int(raw)
    if index < 1 or index > len(candidates):
        raise RuntimeError("任务编号超出范围。")
    return candidates[index - 1]


def find_workspace_record_by_question_path(question_doc_path: str, config: dict | None = None) -> dict | None:
    target = Path(question_doc_path).resolve()
    for record in collect_workspace_records(config):
        record_question_path = record.get("question_doc_path")
        if record_question_path and Path(record_question_path).resolve() == target:
            return record
    return None


def build_manifest_entry(
    *,
    file_path: Path,
    sample_kind: str,
    subject: str | None = None,
    question_count: int | None = None,
    tags: list[str] | None = None,
    requires_subquestion_split: bool | None = None,
) -> dict:
    return {
        "file_name": file_path.name,
        "relative_path": str(file_path),
        "sample_kind": sample_kind,
        "subject": subject or "未知",
        "question_count": question_count,
        "tags": tags or [],
        "requires_subquestion_split": requires_subquestion_split,
    }


def _guess_subject_from_name(name: str) -> str:
    if any(key in name for key in ["英语", "English", "外语"]):
        return "英语"
    if any(key in name for key in ["物理", "化学", "生物", "科学", "理综"]):
        return "理科"
    if any(key in name for key in ["地理", "历史", "政治", "语文", "文综", "道法"]):
        return "文科"
    return "未知"


def _guess_tags_from_name(name: str) -> list[str]:
    tags = []
    if any(key in name for key in ["图", "示意", "装置"]):
        tags.append("image")
    if any(key in name for key in ["材料", "阅读", "据此", "下题"]):
        tags.append("material")
    if any(key in name for key in ["答案", "解析"]):
        tags.append("answer")
    return tags


def sync_regression_manifest() -> str:
    REGRESSION_DIR.mkdir(parents=True, exist_ok=True)
    sample_specs = [
        (FORMAT_DIR / "待清洗文件", "raw_answer_doc"),
        (FORMAT_DIR / "原格式", "formatted_answer_reference"),
        (QUESTION_DIR / "待录入文档", "question_doc"),
    ]

    entries = []
    for folder, sample_kind in sample_specs:
        if not folder.exists():
            continue
        for path in sorted(folder.iterdir()):
            if path.suffix.lower() not in {".doc", ".docx"}:
                continue
            if path.name.startswith("~$"):
                continue
            entries.append(
                build_manifest_entry(
                    file_path=path.relative_to(PROJECT_ROOT),
                    sample_kind=sample_kind,
                    subject=_guess_subject_from_name(path.name),
                    tags=_guess_tags_from_name(path.name),
                    requires_subquestion_split=None,
                )
            )

    REGRESSION_MANIFEST_PATH.write_text(
        json.dumps({"entries": entries}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return str(REGRESSION_MANIFEST_PATH)


def _ensure_module_paths():
    for path in [str(QUESTION_DIR), str(FORMAT_DIR), str(ANSWER_DIR)]:
        if path not in sys.path:
            sys.path.insert(0, path)


def _get_wps():
    _ensure_module_paths()
    from common import get_active_wps

    return get_active_wps()


def _open_document_in_wps(wps, doc_path: str):
    target = os.path.abspath(doc_path)
    for doc in wps.Documents:
        try:
            if os.path.abspath(doc.FullName) == target:
                doc.Activate()
                return doc
        except Exception:
            continue
    return wps.Documents.Open(target)


def _load_question_modules():
    _ensure_module_paths()
    import config as question_config
    from core_parser import detect_subject, get_sections, process_chapter

    return question_config, detect_subject, get_sections, process_chapter


def _load_format_module():
    module_path = FORMAT_DIR / "main.py"
    spec = importlib.util.spec_from_file_location("mohen_format_main", module_path)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module


def _load_answer_module():
    _ensure_module_paths()
    import answer_input

    return answer_input


def _select_sections(sections):
    valid_sections = []
    for section in sections:
        if section["end"] < section["start"]:
            continue
        if "文档开头" in section["title"] and len(sections) > 1:
            continue
        valid_sections.append(section)

    if not valid_sections:
        return []

    print("\n可录入章节：")
    for idx, section in enumerate(valid_sections, 1):
        clean_title = " ".join(section["title"].split())[:50]
        print(f"  [{idx}] {clean_title}")

    raw = input("👉 输入章节序号/范围（如 1 或 1-3，直接回车=全部）: ").strip()
    if not raw:
        return valid_sections
    if "-" in raw:
        start_idx, end_idx = map(int, raw.split("-"))
        return valid_sections[start_idx - 1:end_idx]
    return [valid_sections[int(raw) - 1]]


def _build_question_units_for_sections(doc, subject_name: str, sections: list[dict], overlay_name: str | None = None):
    units = []
    for section in sections:
        units.extend(
            build_question_units_from_wps(
                doc_name=doc.Name,
                subject_name=subject_name,
                doc=doc,
                start_p=section["start"],
                end_p=section["end"],
                overlay_name=overlay_name,
            )
        )
    return units


def run_question_input_only():
    question_config, detect_subject, get_sections, process_chapter = _load_question_modules()
    wps = _get_wps()
    if not wps:
        raise RuntimeError("未找到运行中的 WPS。")

    doc = wps.ActiveDocument
    question_config.CURRENT_CONFIG = detect_subject(doc)
    print(f"✅ 当前题目文档：{doc.Name}")
    print(f"✅ 自动识别科目：{question_config.CURRENT_CONFIG['name']}")

    sections = _select_sections(get_sections(doc))
    if not sections:
        raise RuntimeError("未找到可录入章节。")

    question_units = _build_question_units_for_sections(
        doc,
        question_config.CURRENT_CONFIG["name"],
        sections,
        overlay_name=getattr(question_config, "CURRENT_SUBJECT_OVERLAY", None),
    )
    for section in sections:
        process_chapter(section)
    return {
        "question_doc_path": doc.FullName,
        "question_units": question_units,
        "subject_name": question_config.CURRENT_CONFIG["name"],
    }


def run_answer_clean_and_review(question_doc_path: str, question_units=None, answer_doc_path: str | None = None):
    if not answer_doc_path:
        matched_record = find_workspace_record_by_question_path(question_doc_path)
        if matched_record and matched_record.get("raw_answer_path"):
            answer_doc_path = matched_record["raw_answer_path"]
            print(f"🧭 已从工作台自动匹配答案文档：{answer_doc_path}")
        else:
            answer_doc_path = input("👉 请输入答案文档路径: ").strip()
    if not answer_doc_path:
        raise RuntimeError("答案文档路径不能为空。")

    wps = _get_wps()
    if not wps:
        raise RuntimeError("未找到运行中的 WPS。")

    format_main = _load_format_module()
    doc = _open_document_in_wps(wps, answer_doc_path)
    if not format_main.process_wps_document(doc, wps, question_doc_path=question_doc_path):
        raise RuntimeError("答案清洗失败。")

    standardized_path = derive_cleaned_output_path(answer_doc_path)
    if question_units is None:
        question_units = build_question_units_from_docx(question_doc_path)
    raw_answer_units = build_answer_units_from_docx(
        standardized_path,
        preserve_source_positions=True,
    )
    mapped_answer_units = map_answers(question_units, raw_answer_units)
    report = build_review_report(os.path.basename(standardized_path), question_units, mapped_answer_units)
    report_path = standardized_path.replace(".docx", "_审核清单.md").replace(".doc", "_审核清单.md")
    export_review_report(report, report_path)
    review_status_path = initialize_review_status(
        standardized_path,
        report_path=report_path,
        report=report,
    )
    review_status_path = apply_auto_review_decision(standardized_path, report)
    gate_result = get_review_gate_result(standardized_path)
    return {
        "standardized_path": standardized_path,
        "report": report,
        "report_path": report_path,
        "review_status_path": review_status_path,
        "gate_result": gate_result,
        "answer_units": mapped_answer_units,
    }


def run_standard_answer_input(standardized_path: str, answer_units=None):
    gate_result = get_review_gate_result(standardized_path)
    if not gate_result["allowed"]:
        raise RuntimeError(format_review_gate_message(gate_result))

    answer_module = _load_answer_module()
    wps = _get_wps()
    if not wps:
        raise RuntimeError("未找到运行中的 WPS。")

    doc = _open_document_in_wps(wps, standardized_path)
    if answer_units is None:
        answer_units, blocking_units = answer_module.ensure_standardized_answer_units(standardized_path)
    else:
        blocking_units = answer_module.find_review_required_units(answer_units)

    if blocking_units:
        raise RuntimeError(
            f"标准答案仍存在 {len(blocking_units)} 个高风险答案块，已拒绝录入。"
        )

    answer_module.execute_input_from_units(doc, wps, answer_units, 0, len(answer_units), strict=True)


def run_full_pipeline():
    question_result = run_question_input_only()
    review_result = run_answer_clean_and_review(
        question_doc_path=question_result["question_doc_path"],
        question_units=question_result["question_units"],
    )
    print(f"📝 审核清单：{review_result['report_path']}")
    print(f"🛂 审核状态：{review_result['review_status_path']}")
    if has_blocking_review_issues(review_result["report"]):
        print("⚠️ 自动检查未通过，已停止在答案录入前。请先处理审核清单。")
        return
    print("✅ 自动检查通过，继续答案录入。")
    run_standard_answer_input(
        review_result["standardized_path"],
        answer_units=review_result["answer_units"],
    )


def run_review_decision():
    standardized_path = input("👉 请输入标准答案文档路径: ").strip()
    if not standardized_path:
        raise RuntimeError("标准答案文档路径不能为空。")

    gate_result = get_review_gate_result(standardized_path)
    print(f"🛂 当前检查状态: {gate_result['status']}")
    print(f"📝 状态说明: {gate_result['reason']}")
    print(f"📄 状态文件: {gate_result['status_path']}")

    action = input("👉 输入 a=手动放行 / r=手动拦截 / p=重置待检查: ").strip().lower()
    if action not in {"a", "r", "p"}:
        raise RuntimeError("无效状态覆盖操作。")

    reviewer = input("👉 处理人姓名（可空）: ").strip()
    note = input("👉 备注（可空）: ").strip()
    status = {"a": "approved", "r": "rejected", "p": "pending"}[action]
    status_path = update_review_status(
        standardized_path,
        status=status,
        reviewer=reviewer,
        note=note,
    )
    final_result = get_review_gate_result(standardized_path)
    print(f"✅ 检查状态已更新: {status_path}")
    print(f"🧾 {final_result['reason']}")


def show_menu():
    print("\n" + "=" * 60)
    print("🚀 墨痕教育总控入口")
    print("=" * 60)
    print("  [1] 全流程录入")
    print("  [2] 只录题目")
    print("  [3] 只做答案清洗+自动检查")
    print("  [4] 只录标准答案")
    print("  [5] 手动覆盖检查状态")
    print("  [p] 查看工作台概览/任务面板")
    print("  [q] 退出")
    print("-" * 60)


def main():
    config_path = ensure_workspace_config()
    manifest_path = sync_regression_manifest()
    print(f"🧭 工作台路径配置: {config_path}")
    print(f"📚 已刷新回归样本清单: {manifest_path}")
    print_workspace_overview()
    print_workspace_task_board(limit=8)

    while True:
        show_menu()
        choice = input("👉 请选择模式: ").strip().lower()
        try:
            if choice == "1":
                run_full_pipeline()
            elif choice == "2":
                run_question_input_only()
            elif choice == "3":
                answer_doc_path = None
                question_doc_path = input("👉 请输入题目文档路径（回车则从工作台选择 / 自动读取当前题目文档）: ").strip()
                if not question_doc_path:
                    record = prompt_select_workspace_record("clean")
                    if record:
                        question_doc_path = record["question_doc_path"]
                        answer_doc_path = record["raw_answer_path"]
                    else:
                        wps = _get_wps()
                        if not wps:
                            raise RuntimeError("未找到运行中的 WPS。")
                        question_doc_path = wps.ActiveDocument.FullName
                review_result = run_answer_clean_and_review(question_doc_path, answer_doc_path=answer_doc_path)
                print(f"📝 审核清单已生成: {review_result['report_path']}")
                print(f"🛂 审核状态已生成: {review_result['review_status_path']}")
                if has_blocking_review_issues(review_result["report"]):
                    print("⚠️ 自动检查未通过，已阻止后续录入。")
                else:
                    print("✅ 自动检查通过，当前答案可直接用于后续录入。")
            elif choice == "4":
                standardized_path = input("👉 请输入标准答案文档路径（回车则从工作台选择）: ").strip()
                if not standardized_path:
                    record = prompt_select_workspace_record("input")
                    if not record:
                        raise RuntimeError("标准答案文档路径不能为空。")
                    standardized_path = record["clean_answer_path"]
                run_standard_answer_input(standardized_path)
            elif choice == "5":
                run_review_decision()
            elif choice == "p":
                print_workspace_overview()
                print_workspace_task_board(limit=20)
            elif choice == "q":
                break
            else:
                print("❌ 无效输入，请重新选择。")
        except Exception as exc:
            print(f"❌ {exc}")


if __name__ == "__main__":
    main()
