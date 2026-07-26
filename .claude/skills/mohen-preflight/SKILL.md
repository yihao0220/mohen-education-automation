---
name: mohen-preflight
description: 墨痕教育 DOCX 只读预检链路（P0 画像 / P1a 文档族 / P1b 页面视觉）的安全边界与门禁规则。当涉及 build_document_preflight、analyze_document_families、build_document_render、prepare_p1b_wps_batch、DocumentProfile、ActionPlan、DocumentFamilyReport、PageRenderManifest、VisualRoleReview、page_truth_authority、execution_enabled 或任何"预检/画像/文档族/页面渲染/视觉审核/校准门禁"话题时使用。也用于回答"这个预检结果能不能直接驱动 WPS 录入"、"能不能进 P1c"、"Mac 渲染算不算真值"这类边界问题。
---

# 墨痕教育只读预检链路

预检链路**完全不接管生产流程**：不连接 WPS 执行按键，不改写原题。它的全部产出都是"建议"和"证据"，能不能升级为执行由安全开关决定。

## 三条旁路

| 阶段 | 入口 | 产出 | 当前定位 |
|---|---|---|---|
| P0 | `tools/build_document_preflight.py` | `DocumentProfile.json` (schema 1.1)、`ActionPlan.json`、派生 Markdown | PoC + P0 可用 |
| P1a | `tools/analyze_document_families.py` | `DocumentFamilyReport.json/.md` | 只作批次建议 |
| P1b | `tools/build_document_render.py`、`tools/prepare_p1b_wps_batch.py` | `PageRenderManifest.json`、页面 PNG、`VisualRoleReview.json`、校准报告、`BatchSourceManifest.json` | 跨平台框架可用，生产批次待校准 |

具体命令行参数以脚本 `--help` 和 `README.md` 为准，不要在本文件里维护第二份命令清单。

## 安全开关（全部为 false，不得绕过）

| 开关 | 值 | 含义 |
|---|---|---|
| `execution_enabled` | `false` | 动作计划不得改成真实 F1/F2/F3/F4 执行。要放开必须先完成 WPS Range 绑定、审核门禁和受控执行器。 |
| `automatic_exclusion_enabled` | `false` | Profile 1.1 的题内角色**只作证据**，不得绕过文档族阈值和审核门禁直接删段落。 |
| `classification_mode` | `advisory_only` | 候选文档族只是建议。 |
| `automatic_rule_binding_enabled` | `false` | 候选文档族不得直接绑定到生产规则。 |
| `production_execution_enabled` | `false` | 文档族不得直接驱动 WPS 执行。 |
| `page_truth_authority` | macOS 固定 `false` | Quick Look 只供开发预览，不是分页真值。 |

## 页面真值的唯一来源

只有以下两种算生产页面真值：

1. Windows 上真实的 WPS COM 渲染；
2. 与 `--pdf-input` **同时显式传入 `--attest-wps-export`** 的 WPS 导出 PDF。

未传 `--attest-wps-export` 的任意外部 PDF 一律视为开发预览。**Microsoft Word 不作为回退渲染器。**

## 阶段推进门槛

- 阈值必须用**同一个真实项目批次**校准后，才能进入后续审核门禁。
- P1b → P1c 的门槛是 **52/52 Windows WPS 真实批次人工审核完成**。未完成前不得视为生产分族完成，也不得形成 P1c 文档族规则快照。
- Windows 结果仍需人工填 `VisualRoleReview.json`，**整批**审核通过才允许生成校准报告。

## 产物存放

P1b 的 PDF、页面 PNG、页面/视觉/校准 JSON 及 `BatchSourceManifest.json` 必须放在**原题目录和 Git 仓库之外**，不得提交业务路径或真实文档派生事实。`prepare_p1b_wps_batch.py` 只复制正式 DOCX、排除锁文件、校验原件与副本哈希。

## 排查

预检相关异常的定位表见 skill `mohen-triage`；架构进度主口径见 `docs/墨痕教育架构问题工程思维分析拆解.md`（6 个根问题：1 已解决、5 部分解决、0 未解决）。
