from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
from typing import Any, Mapping, Sequence


REVIEW_SCHEMA_VERSION = "1.0"
MANIFEST_SCHEMA_VERSION = "1.0"
GENERATOR_NAME = "mohen-visual-role-review-p1b"

EXCLUSION_ROLES = {
    "document_banner",
    "document_title",
    "score_instruction",
    "exercise_label",
    "group_heading",
}
QUESTION_ROLES = {"question_text", "question_media"}
KNOWN_ROLES = EXCLUSION_ROLES | QUESTION_ROLES | {"unknown"}
KNOWN_DECISIONS = {"include", "exclude", "review"}
KNOWN_REVIEW_STATUSES = {"draft", "confirmed", "rejected"}


def canonical_json_sha256(value: Mapping[str, Any]) -> str:
    encoded = json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _require_mapping(value: Any, field: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise ValueError(f"{field} 必须是 JSON 对象")
    return value


def _require_string_list(value: Any, field: str) -> list[str]:
    if not isinstance(value, list) or any(not isinstance(item, str) or not item for item in value):
        raise ValueError(f"{field} 必须是非空字符串组成的数组")
    if len(value) != len(set(value)):
        raise ValueError(f"{field} 不能包含重复值")
    return list(value)


def _require_action_sequences(value: Any, field: str) -> list[int]:
    if not isinstance(value, list):
        raise ValueError(f"{field} 必须是数组")
    if any(isinstance(item, bool) or not isinstance(item, int) or item <= 0 for item in value):
        raise ValueError(f"{field} 只能包含正整数")
    if len(value) != len(set(value)):
        raise ValueError(f"{field} 不能包含重复值")
    return list(value)


def _validate_manifest(manifest: Mapping[str, Any]) -> tuple[dict[str, dict[str, Any]], dict[int, dict[str, Any]]]:
    if manifest.get("schema_version") != MANIFEST_SCHEMA_VERSION:
        raise ValueError(f"PageRenderManifest.schema_version 必须为 {MANIFEST_SCHEMA_VERSION}")
    if manifest.get("automatic_rule_binding_enabled") is not False:
        raise ValueError("PageRenderManifest.automatic_rule_binding_enabled 必须为 false")
    if manifest.get("production_execution_enabled") is not False:
        raise ValueError("PageRenderManifest.production_execution_enabled 必须为 false")
    source = _require_mapping(manifest.get("source"), "PageRenderManifest.source")
    sha256 = source.get("sha256")
    if not isinstance(sha256, str) or len(sha256) != 64:
        raise ValueError("PageRenderManifest.source.sha256 无效")
    pages = manifest.get("pages")
    if not isinstance(pages, list) or not pages:
        raise ValueError("PageRenderManifest.pages 不能为空")
    by_id: dict[str, dict[str, Any]] = {}
    page_map: dict[int, dict[str, Any]] = {}
    for page in pages:
        if not isinstance(page, Mapping):
            raise ValueError("PageRenderManifest.pages 的元素必须是对象")
        page_number = page.get("page_number")
        if isinstance(page_number, bool) or not isinstance(page_number, int) or page_number <= 0:
            raise ValueError("PageRenderManifest.page_number 必须是正整数")
        if page_number in page_map:
            raise ValueError(f"PageRenderManifest 出现重复页码：{page_number}")
        page_map[page_number] = dict(page)
        regions = page.get("regions")
        if not isinstance(regions, list):
            raise ValueError(f"第 {page_number} 页 regions 必须是数组")
        for region in regions:
            if not isinstance(region, Mapping):
                raise ValueError(f"第 {page_number} 页区域必须是对象")
            region_id = region.get("region_id")
            if not isinstance(region_id, str) or not region_id:
                raise ValueError(f"第 {page_number} 页存在无效 region_id")
            if region_id in by_id:
                raise ValueError(f"PageRenderManifest 出现重复 region_id：{region_id}")
            normalized = copy.deepcopy(dict(region))
            normalized["page"] = page_number
            by_id[region_id] = normalized
    if int(manifest.get("page_count", -1)) != len(page_map):
        raise ValueError("PageRenderManifest.page_count 与 pages 数量不一致")
    return by_id, page_map


def build_visual_review_template(manifest: Mapping[str, Any]) -> dict[str, Any]:
    manifest_regions, _ = _validate_manifest(manifest)
    regions = []
    for region_id, region in sorted(
        manifest_regions.items(),
        key=lambda item: (item[1]["page"], item[1]["bbox_pdf"][1], item[1]["bbox_pdf"][0], item[0]),
    ):
        regions.append(
            {
                "region_id": region_id,
                "page": region["page"],
                "region_type": region.get("region_type"),
                "bbox_pdf": copy.deepcopy(region.get("bbox_pdf")),
                "bbox_px": copy.deepcopy(region.get("bbox_px")),
                "visible_text": region.get("visible_text") or "",
                "crop_sha256": region.get("crop_sha256"),
                "role": "unknown",
                "decision": "review",
                "belongs_to_question_ids": [],
                "action_sequences": [],
                "confidence": 0.0,
                "evidence": [],
                "review_status": "draft",
            }
        )
    return {
        "schema_version": REVIEW_SCHEMA_VERSION,
        "generator": GENERATOR_NAME,
        "classification_mode": "human_review_required",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "manifest_sha256": canonical_json_sha256(manifest),
        "source": copy.deepcopy(dict(manifest["source"])),
        "region_count": len(regions),
        "regions": regions,
        "notes": [],
    }


def validate_visual_review(
    manifest: Mapping[str, Any],
    review: Mapping[str, Any],
) -> dict[str, Any]:
    manifest_regions, page_map = _validate_manifest(manifest)
    if not isinstance(review, Mapping):
        raise ValueError("VisualRoleReview 必须是 JSON 对象")
    if review.get("schema_version") != REVIEW_SCHEMA_VERSION:
        raise ValueError(f"VisualRoleReview.schema_version 必须为 {REVIEW_SCHEMA_VERSION}")
    if review.get("automatic_rule_binding_enabled") is not False:
        raise ValueError("VisualRoleReview.automatic_rule_binding_enabled 必须为 false")
    if review.get("production_execution_enabled") is not False:
        raise ValueError("VisualRoleReview.production_execution_enabled 必须为 false")
    review_source = _require_mapping(review.get("source"), "VisualRoleReview.source")
    if review_source.get("sha256") != manifest["source"].get("sha256"):
        raise ValueError("VisualRoleReview.source.sha256 与 PageRenderManifest 不一致")
    expected_manifest_sha = canonical_json_sha256(manifest)
    if review.get("manifest_sha256") != expected_manifest_sha:
        raise ValueError("VisualRoleReview.manifest_sha256 与 PageRenderManifest 不一致")
    regions = review.get("regions")
    if not isinstance(regions, list):
        raise ValueError("VisualRoleReview.regions 必须是数组")
    by_id: dict[str, Mapping[str, Any]] = {}
    for item in regions:
        if not isinstance(item, Mapping):
            raise ValueError("VisualRoleReview.regions 的元素必须是对象")
        region_id = item.get("region_id")
        if region_id in by_id:
            raise ValueError(f"VisualRoleReview 出现重复 region_id：{region_id}")
        by_id[str(region_id)] = item
    if set(by_id) != set(manifest_regions):
        missing = sorted(set(manifest_regions) - set(by_id))
        extra = sorted(set(by_id) - set(manifest_regions))
        raise ValueError(f"VisualRoleReview 区域集合不一致：缺少 {missing}；多出 {extra}")

    normalized = copy.deepcopy(dict(review))
    normalized_regions = []
    for region_id in sorted(
        manifest_regions,
        key=lambda key: (
            manifest_regions[key]["page"],
            manifest_regions[key]["bbox_pdf"][1],
            manifest_regions[key]["bbox_pdf"][0],
            key,
        ),
    ):
        item = dict(by_id[region_id])
        fact = manifest_regions[region_id]
        for field in ("page", "region_type", "bbox_pdf", "bbox_px", "visible_text", "crop_sha256"):
            expected = fact["page"] if field == "page" else fact.get(field)
            if item.get(field) != expected:
                raise ValueError(f"{region_id}: {field} 与 PageRenderManifest 不一致")
        page = page_map[fact["page"]]
        x0, top, x1, bottom = item["bbox_pdf"]
        if not (0 <= x0 < x1 <= page["width_pt"] and 0 <= top < bottom <= page["height_pt"]):
            raise ValueError(f"{region_id}: bbox_pdf 超出页面")
        role = item.get("role")
        decision = item.get("decision")
        status = item.get("review_status")
        if role not in KNOWN_ROLES:
            raise ValueError(f"{region_id}: role 无效：{role!r}")
        if decision not in KNOWN_DECISIONS:
            raise ValueError(f"{region_id}: decision 无效：{decision!r}")
        if status not in KNOWN_REVIEW_STATUSES:
            raise ValueError(f"{region_id}: review_status 无效：{status!r}")
        confidence = item.get("confidence")
        if isinstance(confidence, bool) or not isinstance(confidence, (int, float)) or not 0 <= confidence <= 1:
            raise ValueError(f"{region_id}: confidence 必须位于 0 到 1")
        question_ids = _require_string_list(
            item.get("belongs_to_question_ids"),
            f"{region_id}.belongs_to_question_ids",
        )
        action_sequences = _require_action_sequences(
            item.get("action_sequences"),
            f"{region_id}.action_sequences",
        )
        evidence = item.get("evidence")
        if not isinstance(evidence, list) or any(not isinstance(value, str) or not value for value in evidence):
            raise ValueError(f"{region_id}.evidence 必须是字符串数组")
        if status == "confirmed" and not evidence:
            raise ValueError(f"{region_id}: confirmed 决策必须提供 evidence")
        if role == "unknown" and status == "confirmed":
            raise ValueError(f"{region_id}: unknown 角色不能标为 confirmed")
        if status == "confirmed" and role in EXCLUSION_ROLES:
            if decision != "exclude":
                raise ValueError(f"{region_id}: 排除角色必须使用 exclude 决策")
            if question_ids:
                raise ValueError(f"{region_id}: 排除区域不能绑定题号")
        if status == "confirmed" and role in QUESTION_ROLES:
            if decision != "include":
                raise ValueError(f"{region_id}: 题目区域必须使用 include 决策")
            if len(question_ids) != 1:
                raise ValueError(f"{region_id}: 题目区域必须绑定且仅绑定一个题号")
        item["belongs_to_question_ids"] = question_ids
        item["action_sequences"] = action_sequences
        item["confidence"] = float(confidence)
        normalized_regions.append(item)
    if int(review.get("region_count", -1)) != len(normalized_regions):
        raise ValueError("VisualRoleReview.region_count 与 regions 数量不一致")
    normalized["regions"] = normalized_regions
    return normalized


def summarize_visual_review(
    manifest: Mapping[str, Any],
    review: Mapping[str, Any],
) -> dict[str, Any]:
    validated = validate_visual_review(manifest, review)
    regions = validated["regions"]
    confirmed_count = sum(item["review_status"] == "confirmed" for item in regions)
    draft_count = sum(item["review_status"] == "draft" for item in regions)
    rejected_count = sum(item["review_status"] == "rejected" for item in regions)
    unknown_count = sum(item["role"] == "unknown" for item in regions)
    included_count = sum(item["decision"] == "include" for item in regions)
    excluded_count = sum(item["decision"] == "exclude" for item in regions)
    decorative_in_actions = sum(
        item["review_status"] == "confirmed"
        and item["role"] in EXCLUSION_ROLES
        and bool(item["action_sequences"])
        for item in regions
    )
    question_media_without_binding = sum(
        item["review_status"] == "confirmed"
        and item["role"] == "question_media"
        and len(item["belongs_to_question_ids"]) != 1
        for item in regions
    )
    pending_count = sum(
        item["review_status"] != "confirmed" or item["role"] == "unknown"
        for item in regions
    )
    blocking_reasons = []
    production_page_truth = manifest["render"].get("page_truth_authority") is True
    if manifest["source"].get("hash_verified_unchanged") is not True:
        blocking_reasons.append("原题 SHA256 未确认保持不变")
    if not production_page_truth:
        blocking_reasons.append("当前渲染不是 Windows WPS 生产页面真值")
    if draft_count:
        blocking_reasons.append(f"仍有 {draft_count} 个区域处于 draft")
    if rejected_count:
        blocking_reasons.append(f"仍有 {rejected_count} 个区域被拒绝，需重新审核")
    if unknown_count:
        blocking_reasons.append(f"仍有 {unknown_count} 个 unknown 区域")
    if decorative_in_actions:
        blocking_reasons.append(f"仍有 {decorative_in_actions} 个已确认装饰区域进入 F1 动作")
    if question_media_without_binding:
        blocking_reasons.append(f"仍有 {question_media_without_binding} 个题目图片未绑定题号")
    return {
        "source_name": manifest["source"]["name"],
        "source_sha256": manifest["source"]["sha256"],
        "source_hash_verified_unchanged": manifest["source"].get("hash_verified_unchanged") is True,
        "render_provider": manifest["render"].get("provider"),
        "production_page_truth": production_page_truth,
        "region_count": len(regions),
        "confirmed_count": confirmed_count,
        "draft_count": draft_count,
        "rejected_count": rejected_count,
        "unknown_count": unknown_count,
        "pending_count": pending_count,
        "included_count": included_count,
        "excluded_count": excluded_count,
        "question_media_binding_count": sum(
            item["review_status"] == "confirmed" and item["role"] == "question_media"
            for item in regions
        ),
        "question_media_without_binding": question_media_without_binding,
        "decorative_regions_in_actions": decorative_in_actions,
        "blocking_reasons": blocking_reasons,
        "gate_ready": not blocking_reasons,
    }


def render_visual_review_markdown(
    manifest: Mapping[str, Any],
    review: Mapping[str, Any],
) -> str:
    summary = summarize_visual_review(manifest, review)
    lines = [
        "# 页面视觉角色审核",
        "",
        "> JSON 是审核权威事实；本文件仅供人工查看。",
        "",
        "## 安全状态",
        "",
        "- 自动绑定规则：否",
        "- 允许生产执行：否",
        "",
        "## 摘要",
        "",
        f"- 原题：{summary['source_name']}",
        f"- 页面区域：{summary['region_count']}",
        f"- 已确认区域：{summary['confirmed_count']}",
        f"- 待确认区域：{summary['pending_count']}",
        f"- 装饰区域进入动作：{summary['decorative_regions_in_actions']}",
        f"- 门禁可通过：{'是' if summary['gate_ready'] else '否'}",
        "",
        "## 阻断原因",
        "",
    ]
    if not summary["blocking_reasons"]:
        lines.append("- 无")
    else:
        lines.extend(f"- {reason}" for reason in summary["blocking_reasons"])
    return "\n".join(lines) + "\n"


def write_visual_review_template(
    manifest_path: str | Path,
    output_dir: str | Path,
) -> dict[str, Path]:
    return write_visual_review_artifacts(manifest_path, output_dir)


def _read_json(path: str | Path, label: str) -> Any:
    source = Path(path).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    try:
        return json.loads(source.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法读取 {label}：{source}：{exc}") from exc


def write_visual_review_artifacts(
    manifest_path: str | Path,
    output_dir: str | Path,
    *,
    review_path: str | Path | None = None,
) -> dict[str, Path]:
    source_path = Path(manifest_path).resolve()
    manifest = _read_json(source_path, "PageRenderManifest.json")
    review_source = Path(review_path).resolve() if review_path is not None else None
    review = (
        validate_visual_review(manifest, _read_json(review_source, "VisualRoleReview.json"))
        if review_source is not None
        else build_visual_review_template(manifest)
    )
    summary = summarize_visual_review(manifest, review)
    target_dir = Path(output_dir).resolve()
    if target_dir == source_path.parent:
        raise ValueError("视觉审核产物目录不能覆盖页面清单目录")
    target_dir.mkdir(parents=True, exist_ok=True)
    json_path = target_dir / "VisualRoleReview.json"
    markdown_path = target_dir / "VisualRoleReview.md"
    summary_path = target_dir / "VisualReviewSummary.json"
    if review_source is not None and review_source in {json_path, markdown_path, summary_path}:
        raise ValueError("视觉审核产物不能覆盖输入审核 JSON")
    temp_json = target_dir / ".VisualRoleReview.json.tmp"
    temp_markdown = target_dir / ".VisualRoleReview.md.tmp"
    temp_summary = target_dir / ".VisualReviewSummary.json.tmp"
    try:
        temp_json.write_text(
            json.dumps(review, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        temp_markdown.write_text(
            render_visual_review_markdown(manifest, review),
            encoding="utf-8",
        )
        temp_summary.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        json.loads(temp_json.read_text(encoding="utf-8"))
        json.loads(temp_summary.read_text(encoding="utf-8"))
        temp_markdown.read_text(encoding="utf-8")
        temp_json.replace(json_path)
        temp_markdown.replace(markdown_path)
        temp_summary.replace(summary_path)
    finally:
        temp_json.unlink(missing_ok=True)
        temp_markdown.unlink(missing_ok=True)
        temp_summary.unlink(missing_ok=True)
    return {"json": json_path, "markdown": markdown_path, "summary_json": summary_path}
