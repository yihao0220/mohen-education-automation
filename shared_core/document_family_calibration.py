from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Mapping, Sequence

from shared_core.document_families import cluster_document_ids


CALIBRATION_SCHEMA_VERSION = "1.0"
FAMILY_REPORT_SCHEMA_VERSION = "1.0"
GROUND_TRUTH_SCHEMA_VERSION = "1.0"
GENERATOR_NAME = "mohen-document-family-calibration-p1b"


def _require_mapping(value: Any, field: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise ValueError(f"{field} 必须是 JSON 对象")
    return value


def _validate_sha256(value: Any, field: str) -> str:
    if not isinstance(value, str) or len(value) != 64:
        raise ValueError(f"{field} 必须是 64 位 SHA256")
    try:
        int(value, 16)
    except ValueError as exc:
        raise ValueError(f"{field} 不是有效 SHA256") from exc
    return value.lower()


def _validate_family_report(
    report: Mapping[str, Any],
) -> tuple[dict[str, dict[str, Any]], dict[tuple[str, str], float], float]:
    if report.get("schema_version") != FAMILY_REPORT_SCHEMA_VERSION:
        raise ValueError(f"DocumentFamilyReport.schema_version 必须为 {FAMILY_REPORT_SCHEMA_VERSION}")
    if report.get("automatic_rule_binding_enabled") is not False:
        raise ValueError("DocumentFamilyReport.automatic_rule_binding_enabled 必须为 false")
    if report.get("production_execution_enabled") is not False:
        raise ValueError("DocumentFamilyReport.production_execution_enabled 必须为 false")
    profiles = report.get("input_profiles")
    if not isinstance(profiles, list) or not profiles:
        raise ValueError("DocumentFamilyReport.input_profiles 不能为空")
    by_id: dict[str, dict[str, Any]] = {}
    for profile in profiles:
        item = _require_mapping(profile, "DocumentFamilyReport.input_profiles[]")
        document_id = _validate_sha256(
            item.get("document_id") or item.get("source_sha256"),
            "DocumentFamilyReport.input_profiles[].document_id",
        )
        if document_id in by_id:
            raise ValueError(f"DocumentFamilyReport 出现重复 document_id：{document_id}")
        source_name = item.get("source_name")
        if not isinstance(source_name, str) or not source_name:
            raise ValueError("DocumentFamilyReport.input_profiles[].source_name 无效")
        by_id[document_id] = {"document_id": document_id, "source_name": source_name}
    if int(report.get("input_profile_count", -1)) != len(by_id):
        raise ValueError("DocumentFamilyReport.input_profile_count 与输入画像数量不一致")

    similarities: dict[tuple[str, str], float] = {}
    for pair in report.get("pairwise_similarities") or []:
        item = _require_mapping(pair, "DocumentFamilyReport.pairwise_similarities[]")
        first = _validate_sha256(item.get("first_document_id"), "first_document_id")
        second = _validate_sha256(item.get("second_document_id"), "second_document_id")
        if first == second or first not in by_id or second not in by_id:
            raise ValueError("DocumentFamilyReport 两两相似度引用了无效文档")
        value = item.get("similarity")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not 0 <= value <= 1:
            raise ValueError("DocumentFamilyReport.similarity 必须位于 0 到 1")
        key = tuple(sorted((first, second)))
        if key in similarities:
            raise ValueError(f"DocumentFamilyReport 出现重复文档对：{key}")
        similarities[key] = round(float(value), 6)
    expected_pairs = len(by_id) * (len(by_id) - 1) // 2
    if len(similarities) != expected_pairs:
        raise ValueError(
            f"DocumentFamilyReport 两两相似度不完整：期望 {expected_pairs}，实际 {len(similarities)}"
        )
    definition = _require_mapping(report.get("similarity_definition"), "similarity_definition")
    original_threshold = definition.get("family_threshold")
    if (
        isinstance(original_threshold, bool)
        or not isinstance(original_threshold, (int, float))
        or not 0 <= original_threshold <= 1
    ):
        raise ValueError("DocumentFamilyReport.family_threshold 无效")
    return by_id, similarities, round(float(original_threshold), 6)


def _validate_ground_truth(ground_truth: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    if ground_truth.get("schema_version") != GROUND_TRUTH_SCHEMA_VERSION:
        raise ValueError(f"GroundTruth.schema_version 必须为 {GROUND_TRUTH_SCHEMA_VERSION}")
    if ground_truth.get("classification_basis") != "execution_rule_equivalence":
        raise ValueError("GroundTruth.classification_basis 必须为 execution_rule_equivalence")
    if ground_truth.get("automatic_rule_binding_enabled") is not False:
        raise ValueError("GroundTruth.automatic_rule_binding_enabled 必须为 false")
    if ground_truth.get("production_execution_enabled") is not False:
        raise ValueError("GroundTruth.production_execution_enabled 必须为 false")
    documents = ground_truth.get("documents")
    if not isinstance(documents, list) or not documents:
        raise ValueError("GroundTruth.documents 不能为空")
    result: dict[str, dict[str, Any]] = {}
    for document in documents:
        item = _require_mapping(document, "GroundTruth.documents[]")
        sha256 = _validate_sha256(item.get("source_sha256"), "GroundTruth.source_sha256")
        if sha256 in result:
            raise ValueError(f"GroundTruth 出现重复 source_sha256：{sha256}")
        family_id = item.get("rule_family_id")
        source_name = item.get("source_name")
        if not isinstance(family_id, str) or not family_id:
            raise ValueError("GroundTruth.rule_family_id 不能为空")
        if not isinstance(source_name, str) or not source_name:
            raise ValueError("GroundTruth.source_name 不能为空")
        result[sha256] = {
            "source_sha256": sha256,
            "source_name": source_name,
            "rule_family_id": family_id,
            "document_shape": item.get("document_shape"),
            "review_status": item.get("review_status"),
        }
    return result


def _validate_visual_summaries(
    summaries: Sequence[Mapping[str, Any]],
) -> dict[str, dict[str, Any]]:
    if not isinstance(summaries, Sequence) or isinstance(summaries, (str, bytes)):
        raise ValueError("VisualReviewSummaries 必须是数组")
    result: dict[str, dict[str, Any]] = {}
    for summary in summaries:
        item = _require_mapping(summary, "VisualReviewSummaries[]")
        sha256 = _validate_sha256(item.get("source_sha256"), "VisualReviewSummary.source_sha256")
        if sha256 in result:
            raise ValueError(f"VisualReviewSummaries 出现重复 source_sha256：{sha256}")
        source_name = item.get("source_name")
        if not isinstance(source_name, str) or not source_name:
            raise ValueError("VisualReviewSummary.source_name 不能为空")
        result[sha256] = dict(item)
    return result


def _cluster_map(clusters: Sequence[Sequence[str]]) -> dict[str, int]:
    return {
        document_id: cluster_index
        for cluster_index, cluster in enumerate(clusters)
        for document_id in cluster
    }


def _threshold_evaluations(
    document_ids: Sequence[str],
    similarities: Mapping[tuple[str, str], float],
    original_threshold: float,
    ground_truth: Mapping[str, Mapping[str, Any]],
) -> list[dict[str, Any]]:
    candidates = sorted({0.0, 1.0, original_threshold, *similarities.values()})
    comparable_ids = sorted(set(document_ids) & set(ground_truth))
    results = []
    for threshold in candidates:
        clusters = (
            cluster_document_ids(document_ids, similarities, threshold=threshold)
            if len(document_ids) > 1
            else [list(document_ids)]
        )
        memberships = _cluster_map(clusters)
        false_splits = []
        false_merges = []
        for index, first in enumerate(comparable_ids):
            for second in comparable_ids[index + 1 :]:
                same_truth = (
                    ground_truth[first]["rule_family_id"]
                    == ground_truth[second]["rule_family_id"]
                )
                same_prediction = memberships[first] == memberships[second]
                pair = [first, second]
                if same_truth and not same_prediction:
                    false_splits.append(pair)
                elif not same_truth and same_prediction:
                    false_merges.append(pair)
        results.append(
            {
                "threshold": round(float(threshold), 6),
                "cluster_count": len(clusters),
                "clusters": [sorted(cluster) for cluster in clusters],
                "false_split_count": len(false_splits),
                "false_merge_count": len(false_merges),
                "total_error_count": len(false_splits) + len(false_merges),
                "false_split_pairs": false_splits,
                "false_merge_pairs": false_merges,
            }
        )
    return results


def build_p1b_calibration_report(
    family_report: Mapping[str, Any],
    ground_truth: Mapping[str, Any],
    visual_review_summaries: Sequence[Mapping[str, Any]],
) -> dict[str, Any]:
    profiles, similarities, original_threshold = _validate_family_report(family_report)
    truth = _validate_ground_truth(ground_truth)
    summaries = _validate_visual_summaries(visual_review_summaries)
    ids = sorted(profiles)
    evaluations = _threshold_evaluations(ids, similarities, original_threshold, truth)
    zero_error = [item["threshold"] for item in evaluations if item["total_error_count"] == 0]
    recommended_threshold = max(zero_error) if zero_error else None
    truth_family_ids = sorted({item["rule_family_id"] for item in truth.values()})
    calibration_scope = "batch_scoped" if len(truth_family_ids) <= 1 else "multi_family_calibrated"
    warnings = []
    if calibration_scope == "batch_scoped":
        warnings.append("当前人工真值只有一个执行规则族，缺少负例；建议阈值仅适用于本批次。")
    if recommended_threshold is None:
        warnings.append("没有找到同时满足零误拆和零误并的候选阈值。")

    expected = set(ids)
    blocking_reasons = []
    missing_truth = sorted(expected - set(truth))
    extra_truth = sorted(set(truth) - expected)
    if missing_truth:
        blocking_reasons.append(f"人工真值缺少 {len(missing_truth)} 份文档")
    if extra_truth:
        blocking_reasons.append(f"人工真值包含 {len(extra_truth)} 份批次外文档")
    missing_reviews = sorted(expected - set(summaries))
    extra_reviews = sorted(set(summaries) - expected)
    if missing_reviews:
        blocking_reasons.append(f"视觉审核缺少 {len(missing_reviews)} 份文档")
    if extra_reviews:
        blocking_reasons.append(f"视觉审核包含 {len(extra_reviews)} 份批次外文档")
    for document_id in sorted(expected & set(truth)):
        if truth[document_id].get("review_status") != "confirmed":
            blocking_reasons.append(
                f"{profiles[document_id]['source_name']}：人工真值尚未 confirmed"
            )
    for document_id in sorted(expected & set(summaries)):
        summary = summaries[document_id]
        source_name = profiles[document_id]["source_name"]
        if summary.get("source_hash_verified_unchanged") is not True:
            blocking_reasons.append(f"{source_name}：原题 SHA256 未确认保持不变")
        if summary.get("production_page_truth") is not True:
            blocking_reasons.append(f"{source_name}：尚未使用 Windows WPS 生产页面真值")
        if summary.get("decorative_regions_in_actions", 0):
            blocking_reasons.append(
                f"{source_name}：仍有 {summary['decorative_regions_in_actions']} 个装饰区域进入 F1"
            )
        if summary.get("question_media_without_binding", 0):
            blocking_reasons.append(
                f"{source_name}：仍有 {summary['question_media_without_binding']} 个题目图片未绑定题号"
            )
        if summary.get("gate_ready") is not True:
            details = summary.get("blocking_reasons") or ["视觉审核未通过"]
            blocking_reasons.append(f"{source_name}：视觉审核未通过：{'；'.join(map(str, details))}")
    if recommended_threshold is None:
        blocking_reasons.append("阈值校准没有零错误候选")
    blocking_reasons = list(dict.fromkeys(blocking_reasons))

    return {
        "schema_version": CALIBRATION_SCHEMA_VERSION,
        "generator": GENERATOR_NAME,
        "classification_mode": "human_calibrated_advisory",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "calibration_scope": calibration_scope,
        "input_document_count": len(ids),
        "ground_truth_family_count": len(truth_family_ids),
        "ground_truth_families": truth_family_ids,
        "original_threshold": original_threshold,
        "recommended_threshold": recommended_threshold,
        "threshold_evaluations": evaluations,
        "visual_review_summaries": [summaries[key] for key in sorted(summaries)],
        "gate": {
            "passed": not blocking_reasons,
            "blocking_reasons": blocking_reasons,
            "missing_ground_truth_document_ids": missing_truth,
            "missing_visual_review_document_ids": missing_reviews,
        },
        "warnings": warnings,
    }


def render_p1b_calibration_markdown(report: Mapping[str, Any]) -> str:
    recommended = report.get("recommended_threshold")
    threshold_text = f"{recommended:.6f}" if isinstance(recommended, (int, float)) else "无"
    lines = [
        "# P1b 文档族校准与整批门禁报告",
        "",
        "> JSON 是机器权威事实；本文件仅供人工审核。",
        "",
        "## 安全状态",
        "",
        f"- 分析模式：{report['classification_mode']}",
        "- 自动绑定文档族规则：否",
        "- 允许生产执行：否",
        "",
        "## 校准摘要",
        "",
        f"- 输入文档：{report['input_document_count']}",
        f"- 人工真值族：{report['ground_truth_family_count']}",
        f"- 校准范围：{report['calibration_scope']}",
        f"- 原阈值：{report['original_threshold']:.6f}",
        f"- 建议阈值：{threshold_text}",
        f"- 整批门禁：{'通过' if report['gate']['passed'] else '未通过'}",
        "",
        "## 阻断原因",
        "",
    ]
    if report["gate"]["blocking_reasons"]:
        lines.extend(f"- {reason}" for reason in report["gate"]["blocking_reasons"])
    else:
        lines.append("- 无")
    lines.extend(["", "## 阈值扫描", ""])
    for item in report["threshold_evaluations"]:
        lines.append(
            f"- {item['threshold']:.6f}：族 {item['cluster_count']}，"
            f"误拆 {item['false_split_count']}，误并 {item['false_merge_count']}"
        )
    lines.extend(["", "## 告警", ""])
    if report["warnings"]:
        lines.extend(f"- {warning}" for warning in report["warnings"])
    else:
        lines.append("- 无")
    return "\n".join(lines) + "\n"


def _read_json(path: str | Path, label: str) -> Any:
    source = Path(path).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    try:
        return json.loads(source.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法读取 {label}：{source}：{exc}") from exc


def read_visual_review_summaries(path: str | Path) -> list[Mapping[str, Any]]:
    source = Path(path).resolve()
    if source.is_file():
        value = _read_json(source, "VisualReviewSummary.json")
        return value if isinstance(value, list) else [value]
    if not source.is_dir():
        raise FileNotFoundError(source)
    files = sorted(source.rglob("VisualReviewSummary.json"))
    if not files:
        raise ValueError(f"审核目录中没有 VisualReviewSummary.json：{source}")
    summaries: list[Mapping[str, Any]] = []
    for file_path in files:
        value = _read_json(file_path, "VisualReviewSummary.json")
        if isinstance(value, list):
            summaries.extend(value)
        else:
            summaries.append(value)
    return summaries


def write_p1b_calibration_report(
    family_report_path: str | Path,
    ground_truth_path: str | Path,
    visual_summaries_path: str | Path,
    output_dir: str | Path,
) -> dict[str, Path]:
    report = build_p1b_calibration_report(
        _read_json(family_report_path, "DocumentFamilyReport.json"),
        _read_json(ground_truth_path, "GroundTruth.json"),
        read_visual_review_summaries(visual_summaries_path),
    )
    target_dir = Path(output_dir).resolve()
    inputs = {
        Path(family_report_path).resolve(),
        Path(ground_truth_path).resolve(),
        Path(visual_summaries_path).resolve(),
    }
    json_path = target_dir / "P1bCalibrationReport.json"
    markdown_path = target_dir / "P1bCalibrationReport.md"
    if json_path in inputs or markdown_path in inputs:
        raise ValueError("P1b 输出不能覆盖输入 JSON")
    target_dir.mkdir(parents=True, exist_ok=True)
    temp_json = target_dir / ".P1bCalibrationReport.json.tmp"
    temp_markdown = target_dir / ".P1bCalibrationReport.md.tmp"
    try:
        temp_json.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        temp_markdown.write_text(render_p1b_calibration_markdown(report), encoding="utf-8")
        json.loads(temp_json.read_text(encoding="utf-8"))
        temp_markdown.read_text(encoding="utf-8")
        temp_json.replace(json_path)
        temp_markdown.replace(markdown_path)
    finally:
        temp_json.unlink(missing_ok=True)
        temp_markdown.unlink(missing_ok=True)
    return {"json": json_path, "markdown": markdown_path}
