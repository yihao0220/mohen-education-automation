from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


REPORT_SCHEMA_VERSION = "1.0"
GENERATOR_NAME = "mohen-document-family-analysis-poc"
PROFILE_SCHEMA_VERSION = "1.1"
FAMILY_THRESHOLD = 0.78
OUTLIER_THRESHOLD = 0.60
BOUNDARY_REVIEW_UPPER = 0.82
LOW_CONFIDENCE_THRESHOLD = 0.75

KNOWN_ROLES = (
    "document_title",
    "section_heading",
    "internal_heading",
    "question_start",
    "option",
    "body",
)

SIMILARITY_WEIGHTS = {
    "role_distribution": 0.35,
    "action_density": 0.20,
    "heading_density": 0.10,
    "object_presence": 0.15,
    "rich_object_density": 0.10,
    "document_scale": 0.10,
}


def _round(value: float) -> float:
    return round(float(value), 6)


def _require_mapping(value: Any, field: str, source_name: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise ValueError(f"{source_name}: {field} 必须是 JSON 对象")
    return value


def _require_nonnegative_number(value: Any, field: str, source_name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0:
        raise ValueError(f"{source_name}: {field} 必须是非负数")
    return float(value)


def _canonical_duplicate_view(profile: Mapping[str, Any]) -> str:
    view = copy.deepcopy(dict(profile))
    source = view.get("source")
    if isinstance(source, dict):
        source.pop("path", None)
    view.pop("_profile_path", None)
    return json.dumps(view, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _validate_profile(profile: Mapping[str, Any]) -> dict[str, Any]:
    if not isinstance(profile, Mapping):
        raise ValueError("输入 Profile 必须是 JSON 对象")

    source = _require_mapping(profile.get("source"), "source", "未知 Profile")
    source_name = source.get("name")
    if not isinstance(source_name, str) or not source_name.strip():
        raise ValueError("Profile 缺少有效的 source.name")
    if profile.get("schema_version") != PROFILE_SCHEMA_VERSION:
        raise ValueError(
            f"{source_name}: schema_version 必须为 {PROFILE_SCHEMA_VERSION}，"
            f"实际为 {profile.get('schema_version')!r}"
        )

    source_sha256 = source.get("sha256")
    if not isinstance(source_sha256, str) or len(source_sha256) != 64:
        raise ValueError(f"{source_name}: source.sha256 必须是 64 位十六进制字符串")
    try:
        int(source_sha256, 16)
    except ValueError as exc:
        raise ValueError(f"{source_name}: source.sha256 不是有效十六进制") from exc
    if source.get("readonly_input") is not True:
        raise ValueError(f"{source_name}: source.readonly_input 必须为 true")

    roles = _require_mapping(profile.get("roles"), "roles", source_name)
    if roles.get("automatic_exclusion_enabled") is not False:
        raise ValueError(f"{source_name}: roles.automatic_exclusion_enabled 必须为 false")

    fingerprint = _require_mapping(profile.get("fingerprint"), "fingerprint", source_name)
    required_numbers = (
        "nonempty_paragraph_count",
        "question_action_count",
        "heading_count",
        "table_count",
        "media_node_count",
        "formula_count",
    )
    numbers = {
        field: _require_nonnegative_number(fingerprint.get(field), f"fingerprint.{field}", source_name)
        for field in required_numbers
    }
    if numbers["nonempty_paragraph_count"] <= 0:
        raise ValueError(f"{source_name}: fingerprint.nonempty_paragraph_count 必须大于 0")

    role_counts = _require_mapping(fingerprint.get("role_counts"), "fingerprint.role_counts", source_name)
    normalized_role_counts: dict[str, float] = {}
    for role, value in role_counts.items():
        if not isinstance(role, str):
            raise ValueError(f"{source_name}: fingerprint.role_counts 的角色名必须是字符串")
        normalized_role_counts[role] = _require_nonnegative_number(
            value,
            f"fingerprint.role_counts.{role}",
            source_name,
        )

    normalized = copy.deepcopy(dict(profile))
    normalized["_document_id"] = source_sha256.lower()
    normalized["_source_name"] = source_name
    normalized["_numbers"] = numbers
    normalized["_role_counts"] = normalized_role_counts
    normalized["_unknown_roles"] = sorted(set(normalized_role_counts) - set(KNOWN_ROLES))
    normalized["_profile_path"] = str(profile.get("_profile_path") or "")
    return normalized


def _deduplicate_profiles(profiles: Sequence[Mapping[str, Any]]) -> tuple[list[dict[str, Any]], list[str]]:
    validated = [_validate_profile(profile) for profile in profiles]
    if not validated:
        raise ValueError("至少需要一个 DocumentProfile.json")

    by_sha: dict[str, dict[str, Any]] = {}
    canonical_by_sha: dict[str, str] = {}
    warnings: list[str] = []
    for profile in sorted(
        validated,
        key=lambda item: (item["_document_id"], item["_source_name"], item["_profile_path"]),
    ):
        document_id = profile["_document_id"]
        canonical = _canonical_duplicate_view(profile)
        if document_id not in by_sha:
            by_sha[document_id] = profile
            canonical_by_sha[document_id] = canonical
            continue
        if canonical_by_sha[document_id] != canonical:
            raise ValueError(
                f"源 SHA256 {document_id} 对应互相冲突的 Profile 内容，已停止整批分析"
            )
        warnings.append(
            f"重复 Profile 已去重：{profile['_source_name']}（源 SHA256 {document_id}）"
        )
    return list(by_sha.values()), warnings


def _features(profile: Mapping[str, Any]) -> dict[str, Any]:
    numbers = profile["_numbers"]
    paragraph_count = numbers["nonempty_paragraph_count"]
    role_counts = profile["_role_counts"]
    return {
        "role_distribution": {
            role: _round(role_counts.get(role, 0.0) / paragraph_count) for role in KNOWN_ROLES
        },
        "action_density": _round(numbers["question_action_count"] / paragraph_count),
        "heading_density": _round(numbers["heading_count"] / paragraph_count),
        "object_presence": {
            "table": numbers["table_count"] > 0,
            "media": numbers["media_node_count"] > 0,
            "formula": numbers["formula_count"] > 0,
        },
        "media_density": _round(numbers["media_node_count"] / paragraph_count),
        "formula_density": _round(numbers["formula_count"] / paragraph_count),
        "document_scale": int(paragraph_count),
    }


def _scalar_ratio_similarity(first: float, second: float) -> float:
    if first == 0 and second == 0:
        return 1.0
    if first == 0 or second == 0:
        return 0.0
    return min(first, second) / max(first, second)


def _pair_similarity(first: Mapping[str, Any], second: Mapping[str, Any]) -> dict[str, Any]:
    first_features = first["_features"]
    second_features = second["_features"]

    role_distance = sum(
        abs(first_features["role_distribution"][role] - second_features["role_distribution"][role])
        for role in KNOWN_ROLES
    )
    role_score = max(0.0, min(1.0, 1.0 - 0.5 * role_distance))
    presence_score = sum(
        first_features["object_presence"][key] == second_features["object_presence"][key]
        for key in ("table", "media", "formula")
    ) / 3
    rich_density_score = (
        _scalar_ratio_similarity(
            first_features["media_density"], second_features["media_density"]
        )
        + _scalar_ratio_similarity(
            first_features["formula_density"], second_features["formula_density"]
        )
    ) / 2
    scores = {
        "role_distribution": role_score,
        "action_density": _scalar_ratio_similarity(
            first_features["action_density"], second_features["action_density"]
        ),
        "heading_density": _scalar_ratio_similarity(
            first_features["heading_density"], second_features["heading_density"]
        ),
        "object_presence": presence_score,
        "rich_object_density": rich_density_score,
        "document_scale": _scalar_ratio_similarity(
            first_features["document_scale"], second_features["document_scale"]
        ),
    }
    components = {
        name: {
            "score": _round(scores[name]),
            "weight": weight,
            "contribution": _round(scores[name] * weight),
        }
        for name, weight in SIMILARITY_WEIGHTS.items()
    }
    similarity = _round(sum(item["contribution"] for item in components.values()))
    return {"similarity": similarity, "components": components}


def _pair_key(first_id: str, second_id: str) -> tuple[str, str]:
    if first_id == second_id:
        raise ValueError("不能为同一文档创建相似度键")
    return tuple(sorted((first_id, second_id)))


def _similarity_value(
    similarities: Mapping[tuple[str, str], Any],
    first_id: str,
    second_id: str,
) -> float:
    value = similarities[_pair_key(first_id, second_id)]
    if isinstance(value, Mapping):
        value = value["similarity"]
    return float(value)


def cluster_document_ids(
    document_ids: Iterable[str],
    similarities: Mapping[tuple[str, str], Any],
    *,
    threshold: float = FAMILY_THRESHOLD,
) -> list[list[str]]:
    clusters = [[document_id] for document_id in sorted(set(document_ids))]
    while True:
        candidates: list[tuple[float, tuple[str, ...], int, int]] = []
        for left_index in range(len(clusters)):
            for right_index in range(left_index + 1, len(clusters)):
                left = clusters[left_index]
                right = clusters[right_index]
                complete_link = min(
                    _similarity_value(similarities, left_id, right_id)
                    for left_id in left
                    for right_id in right
                )
                if complete_link >= threshold:
                    merged_key = tuple(sorted(left + right))
                    candidates.append((complete_link, merged_key, left_index, right_index))
        if not candidates:
            break
        _, _, left_index, right_index = min(
            candidates,
            key=lambda item: (-item[0], item[1]),
        )
        merged = sorted(clusters[left_index] + clusters[right_index])
        clusters = [
            cluster
            for index, cluster in enumerate(clusters)
            if index not in (left_index, right_index)
        ]
        clusters.append(merged)
        clusters.sort(key=lambda cluster: tuple(cluster))
    return clusters


def select_representative(
    document_ids: Iterable[str],
    similarities: Mapping[tuple[str, str], Any],
) -> tuple[str, dict[str, float]]:
    ids = sorted(set(document_ids))
    if not ids:
        raise ValueError("候选文档族不能为空")
    if len(ids) == 1:
        return ids[0], {ids[0]: 1.0}
    averages = {
        document_id: _round(
            sum(
                _similarity_value(similarities, document_id, other_id)
                for other_id in ids
                if other_id != document_id
            )
            / (len(ids) - 1)
        )
        for document_id in ids
    }
    representative = min(ids, key=lambda document_id: (-averages[document_id], document_id))
    return representative, averages


def _profile_risk_reasons(profile: Mapping[str, Any]) -> list[str]:
    reasons: list[str] = []
    for issue in profile.get("issues") or []:
        if isinstance(issue, Mapping) and str(issue.get("severity", "")).lower() in {
            "warning",
            "error",
        }:
            reasons.append(f"Profile issue: {issue.get('code') or issue.get('message') or '未知'}")
    roles = profile.get("roles") or {}
    for candidate in roles.get("heading_candidates") or []:
        if not isinstance(candidate, Mapping):
            continue
        confidence = candidate.get("confidence")
        if isinstance(confidence, (int, float)) and confidence < LOW_CONFIDENCE_THRESHOLD:
            reasons.append(
                f"低置信度标题候选：第 {candidate.get('paragraph_index', '?')} 段，"
                f"置信度 {_round(confidence)}"
            )
    if profile["_unknown_roles"]:
        reasons.append(f"未知角色：{', '.join(profile['_unknown_roles'])}")
    return reasons


def _member_view(profile: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "document_id": profile["_document_id"],
        "source_name": profile["_source_name"],
        "source_sha256": profile["_document_id"],
        "profile_path": profile["_profile_path"],
    }


def analyze_document_profiles(profiles: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    normalized, warnings = _deduplicate_profiles(profiles)
    normalized.sort(key=lambda item: (item["_document_id"], item["_source_name"]))
    for profile in normalized:
        profile["_features"] = _features(profile)

    by_id = {profile["_document_id"]: profile for profile in normalized}
    ids = sorted(by_id)
    pair_details: dict[tuple[str, str], dict[str, Any]] = {}
    pairwise_report: list[dict[str, Any]] = []
    for index, first_id in enumerate(ids):
        for second_id in ids[index + 1 :]:
            detail = _pair_similarity(by_id[first_id], by_id[second_id])
            pair_details[(first_id, second_id)] = detail
            pairwise_report.append(
                {
                    "first_document_id": first_id,
                    "first_source_name": by_id[first_id]["_source_name"],
                    "second_document_id": second_id,
                    "second_source_name": by_id[second_id]["_source_name"],
                    **detail,
                }
            )

    clusters = (
        cluster_document_ids(ids, pair_details, threshold=FAMILY_THRESHOLD)
        if len(ids) > 1
        else [ids]
    )
    families: list[dict[str, Any]] = []
    for cluster in clusters:
        representative_id, averages = select_representative(cluster, pair_details)
        pair_scores = [
            _similarity_value(pair_details, first_id, second_id)
            for index, first_id in enumerate(cluster)
            for second_id in cluster[index + 1 :]
        ]
        if pair_scores:
            minimum_similarity = _round(min(pair_scores))
            average_similarity = _round(sum(pair_scores) / len(pair_scores))
            component_averages = {
                component: _round(
                    sum(
                        pair_details[_pair_key(first_id, second_id)]["components"][component][
                            "score"
                        ]
                        for index, first_id in enumerate(cluster)
                        for second_id in cluster[index + 1 :]
                    )
                    / len(pair_scores)
                )
                for component in SIMILARITY_WEIGHTS
            }
            common_evidence = [
                {"component": name, "average_score": score}
                for name, score in sorted(
                    component_averages.items(),
                    key=lambda item: (-item[1], item[0]),
                )[:3]
            ]
        else:
            minimum_similarity = 1.0
            average_similarity = 1.0
            common_evidence = []
        family_digest = hashlib.sha256("|".join(cluster).encode("ascii")).hexdigest()[:12]
        families.append(
            {
                "family_id": f"family_{family_digest}",
                "member_count": len(cluster),
                "members": [_member_view(by_id[document_id]) for document_id in cluster],
                "representative": _member_view(by_id[representative_id]),
                "member_average_similarities": averages,
                "minimum_similarity": minimum_similarity,
                "average_similarity": average_similarity,
                "common_evidence": common_evidence,
            }
        )
    families.sort(key=lambda item: item["family_id"])

    singleton_ids = [cluster[0] for cluster in clusters if len(cluster) == 1]
    outlier_candidates: list[dict[str, Any]] = []
    unresolved_singletons: list[dict[str, Any]] = []
    for document_id in singleton_ids:
        nearest_id = None
        nearest_similarity = None
        if len(ids) > 1:
            nearest_id = min(
                (other_id for other_id in ids if other_id != document_id),
                key=lambda other_id: (
                    -_similarity_value(pair_details, document_id, other_id),
                    other_id,
                ),
            )
            nearest_similarity = _round(
                _similarity_value(pair_details, document_id, nearest_id)
            )
        item = {
            **_member_view(by_id[document_id]),
            "nearest_document_id": nearest_id,
            "nearest_source_name": by_id[nearest_id]["_source_name"] if nearest_id else None,
            "nearest_similarity": nearest_similarity,
        }
        if len(ids) >= 3 and nearest_similarity is not None and nearest_similarity < OUTLIER_THRESHOLD:
            item["reason"] = f"最高近邻相似度低于 {OUTLIER_THRESHOLD:.2f}"
            outlier_candidates.append(item)
        else:
            item["reason"] = "未达到候选族合并阈值，证据不足，等待人工确认"
            unresolved_singletons.append(item)
    outlier_candidates.sort(key=lambda item: (item["source_name"], item["document_id"]))
    unresolved_singletons.sort(key=lambda item: (item["source_name"], item["document_id"]))

    review_queue: list[dict[str, Any]] = []
    review_queue.extend(
        {"review_type": "outlier_candidate", **item} for item in outlier_candidates
    )
    review_queue.extend(
        {"review_type": "unresolved_singleton", **item} for item in unresolved_singletons
    )
    for family in families:
        if (
            family["member_count"] > 1
            and FAMILY_THRESHOLD <= family["minimum_similarity"] < BOUNDARY_REVIEW_UPPER
        ):
            review_queue.append(
                {
                    "review_type": "boundary_family",
                    "family_id": family["family_id"],
                    "source_name": family["representative"]["source_name"],
                    "reason": (
                        f"族内最低相似度处于 {FAMILY_THRESHOLD:.2f}—"
                        f"{BOUNDARY_REVIEW_UPPER:.2f} 边界复核区间"
                    ),
                }
            )
    for profile in normalized:
        reasons = _profile_risk_reasons(profile)
        if reasons:
            review_queue.append(
                {
                    "review_type": "profile_risk",
                    **_member_view(profile),
                    "reason": "；".join(reasons),
                }
            )
    for family in families:
        review_queue.append(
            {
                "review_type": "family_representative",
                "family_id": family["family_id"],
                **family["representative"],
                "reason": "候选文档族代表样本",
            }
        )

    for profile in normalized:
        if profile["_unknown_roles"]:
            warnings.append(
                f"{profile['_source_name']} 出现未知角色：{', '.join(profile['_unknown_roles'])}"
            )
    warnings.append(
        "P1a 阈值仅为保守启动值，必须使用同一真实项目批次校准后才能进入生产门禁。"
    )

    input_profiles = []
    for profile in normalized:
        issues = profile.get("issues") or []
        input_profiles.append(
            {
                **_member_view(profile),
                "profile_schema_version": profile["schema_version"],
                "unknown_roles": profile["_unknown_roles"],
                "issue_count": len(issues) if isinstance(issues, list) else 0,
                "features": profile["_features"],
            }
        )

    return {
        "schema_version": REPORT_SCHEMA_VERSION,
        "generator": GENERATOR_NAME,
        "classification_mode": "advisory_only",
        "automatic_rule_binding_enabled": False,
        "production_execution_enabled": False,
        "input_profile_count": len(normalized),
        "input_profiles": input_profiles,
        "feature_definition": {
            "known_roles": list(KNOWN_ROLES),
            "normalization": "counts_per_nonempty_paragraph",
            "low_confidence_threshold": LOW_CONFIDENCE_THRESHOLD,
        },
        "similarity_definition": {
            "method": "explainable_weighted_similarity_complete_link",
            "weights": SIMILARITY_WEIGHTS,
            "family_threshold": FAMILY_THRESHOLD,
            "outlier_threshold": OUTLIER_THRESHOLD,
            "boundary_review_upper": BOUNDARY_REVIEW_UPPER,
        },
        "pairwise_similarities": pairwise_report,
        "families": families,
        "outlier_candidates": outlier_candidates,
        "unresolved_singletons": unresolved_singletons,
        "review_queue": review_queue,
        "warnings": warnings,
    }


def discover_profile_paths(inputs: Iterable[str | Path]) -> list[Path]:
    discovered: set[Path] = set()
    for raw_input in inputs:
        path = Path(raw_input).resolve()
        if not path.exists():
            raise FileNotFoundError(path)
        if path.is_dir():
            discovered.update(item.resolve() for item in path.rglob("*_DocumentProfile.json"))
            continue
        if not path.name.endswith("_DocumentProfile.json"):
            raise ValueError(f"不是 DocumentProfile JSON：{path}")
        discovered.add(path)
    if not discovered:
        raise ValueError("没有发现 *_DocumentProfile.json")
    return sorted(discovered, key=lambda path: path.as_posix().casefold())


def _load_profiles(profile_paths: Sequence[Path]) -> list[dict[str, Any]]:
    profiles: list[dict[str, Any]] = []
    for path in profile_paths:
        try:
            profile = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError) as exc:
            raise ValueError(f"无法读取 Profile JSON：{path}：{exc}") from exc
        if not isinstance(profile, dict):
            raise ValueError(f"Profile JSON 顶层必须是对象：{path}")
        profile["_profile_path"] = path.as_posix()
        profiles.append(profile)
    return profiles


def render_document_family_markdown(report: Mapping[str, Any]) -> str:
    lines = [
        "# 文档族分析报告",
        "",
        "> JSON 是机器权威事实；本文件仅供人工审核，不允许作为执行输入。",
        "",
        "## 安全状态",
        "",
        f"- 分析模式：{report['classification_mode']}",
        "- 自动绑定文档族规则：否",
        "- 允许生产执行：否",
        "- 原始 DOCX：未读取、未写回",
        "",
        "## 批次摘要",
        "",
        f"- 输入画像：{report['input_profile_count']}",
        f"- 候选文档族：{len(report['families'])}",
        f"- 异常候选：{len(report['outlier_candidates'])}",
        f"- 未决单例：{len(report['unresolved_singletons'])}",
        f"- 人工复核项：{len(report['review_queue'])}",
        "",
        "## 人工复核队列",
        "",
    ]
    if not report["review_queue"]:
        lines.append("- 无")
    for index, item in enumerate(report["review_queue"], start=1):
        lines.append(
            f"{index}. `{item['review_type']}`｜{item.get('source_name') or item.get('family_id')}｜"
            f"{item.get('reason', '人工确认')}"
        )

    lines.extend(["", "## 候选文档族", ""])
    for family in report["families"]:
        lines.extend(
            [
                f"### {family['family_id']}",
                "",
                f"- 成员数：{family['member_count']}",
                f"- 代表样本：{family['representative']['source_name']}",
                f"- 族内最低/平均相似度：{family['minimum_similarity']:.6f} / {family['average_similarity']:.6f}",
                "- 成员：" + "、".join(member["source_name"] for member in family["members"]),
                "- 主要共同证据："
                + (
                    "、".join(
                        f"{item['component']}={item['average_score']:.6f}"
                        for item in family["common_evidence"]
                    )
                    or "单例，无族内比较证据"
                ),
                "",
            ]
        )

    lines.extend(["## 异常与未决单例", ""])
    exceptional = list(report["outlier_candidates"]) + list(report["unresolved_singletons"])
    if not exceptional:
        lines.append("- 无")
    for item in exceptional:
        nearest = (
            f"最近邻 {item['nearest_source_name']}，相似度 {item['nearest_similarity']:.6f}"
            if item.get("nearest_source_name") and item.get("nearest_similarity") is not None
            else "没有可比较的最近邻"
        )
        lines.append(f"- {item['source_name']}｜{nearest}｜{item['reason']}")

    lines.extend(["", "## 两两相似度", ""])
    if not report["pairwise_similarities"]:
        lines.append("- 仅一份画像，无两两比较")
    for pair in report["pairwise_similarities"]:
        components = "；".join(
            f"{name}={detail['score']:.6f}×{detail['weight']:.2f}"
            for name, detail in pair["components"].items()
        )
        lines.append(
            f"- {pair['first_source_name']} ↔ {pair['second_source_name']}："
            f"{pair['similarity']:.6f}｜{components}"
        )

    lines.extend(["", "## 告警与校准说明", ""])
    for warning in report["warnings"]:
        lines.append(f"- {warning}")
    return "\n".join(lines) + "\n"


def write_document_family_report(
    inputs: Iterable[str | Path],
    output_dir: str | Path,
) -> dict[str, Path]:
    profile_paths = discover_profile_paths(inputs)
    target_dir = Path(output_dir).resolve()
    json_path = target_dir / "DocumentFamilyReport.json"
    markdown_path = target_dir / "DocumentFamilyReport.md"
    input_set = set(profile_paths)
    if json_path in input_set or markdown_path in input_set:
        raise ValueError("输出文件不能覆盖输入 Profile")

    report = analyze_document_profiles(_load_profiles(profile_paths))
    json_text = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    markdown_text = render_document_family_markdown(report)
    target_dir.mkdir(parents=True, exist_ok=True)
    temp_json = target_dir / ".DocumentFamilyReport.json.tmp"
    temp_markdown = target_dir / ".DocumentFamilyReport.md.tmp"
    try:
        temp_json.write_text(json_text, encoding="utf-8")
        temp_markdown.write_text(markdown_text, encoding="utf-8")
        json.loads(temp_json.read_text(encoding="utf-8"))
        temp_markdown.read_text(encoding="utf-8")
        temp_json.replace(json_path)
        temp_markdown.replace(markdown_path)
    finally:
        temp_json.unlink(missing_ok=True)
        temp_markdown.unlink(missing_ok=True)
    return {"json": json_path, "markdown": markdown_path}
