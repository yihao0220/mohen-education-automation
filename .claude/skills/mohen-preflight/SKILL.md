---
name: mohen-preflight
description: 墨痕教育 DOCX 只读预检链路（P0 画像 / P1a 文档族 / P1b 页面视觉）的安全边界与门禁规则。当涉及 build_document_preflight、analyze_document_families、build_document_render、prepare_p1b_wps_batch、DocumentProfile、ActionPlan、DocumentFamilyReport、PageRenderManifest、VisualRoleReview、page_truth_authority、execution_enabled 或任何"预检/画像/文档族/页面渲染/视觉审核/校准门禁"话题时使用。也用于回答"这个预检结果能不能直接驱动 WPS 录入"、"能不能进 P1c"、"Mac 渲染算不算真值"这类边界问题。
---

# 墨痕教育只读预检链路

预检链路**完全不接管生产流程**：不连接 WPS 执行按键，不改写原题。它的全部产出都是"建议"和"证据"，能不能升级为执行由安全开关决定。

`tools/preview_f1_action_plan.py` 是预检之外的 WPS Range 无按键绑定试点：它可连接当前 WPS 原题并选区，但永远不执行 F1，回执必须保持 `keypress_count=0`。不得把它说成生产执行器。

`tools/execute_f1_action_plan.py` 是独立的单样本受控执行试点，不改动预检安全开关。它只允许已完成 WPS Range 验收的指定源哈希，默认只执行一题，并要求启动确认、每题精确 `f1` 确认和按键后人工检查。2026-08-01 莞美高二地理无表格样本完成 8/8 真实 F1；该证据不放开表格和其他文档家族。

## 三条旁路

| 阶段 | 入口 | 产出 | 当前定位 |
|---|---|---|---|
| P0 | `tools/build_document_preflight.py` | `DocumentProfile.json` (schema 1.1)、`ActionPlan.json`、派生 Markdown | PoC + P0 可用 |
| P1a | `tools/analyze_document_families.py` | `DocumentFamilyReport.json/.md` | 只作批次建议 |
| P1b | `tools/build_document_render.py`、`tools/prepare_p1b_wps_batch.py` | `PageRenderManifest.json`、页面 PNG、`VisualRoleReview.json`、校准报告、`BatchSourceManifest.json` | 跨平台框架可用，生产批次待校准 |

## F1 ActionPlan → WPS Range 无按键试点

- 只允许 `mode=preview_only` 且 `execution_enabled=false` 的 F1 ActionPlan。
- 只选区并等待人工确认，不导入或调用按键自动化；中途输入 `s` 可安全停止。
- Source ActionPlan 1.1 只描述“录什么”：保留题号、开头/结尾锚点、表格/媒体/公式数量等语义事实，不输出 DOCX 段落号、虚拟节点号或媒体段落号。解析器内部坐标只用于编译，不能控制 WPS。
- Windows 连接当前 WPS 后，必须在第一次人工确认前读取 `Document.Paragraphs` / `Tables`，一次性给全部动作生成 `wps_ref`；预览和受控执行只认 `wps_ref`。任一动作零命中、多命中、顺序异常或表格数量不符时，在开始前停止。
- 生成 ActionPlan 时，从 DOCX 向上查找最近一个明确包含学科名的项目文件夹（如“莞美-高二-地理”），直接以该目录名设定基础学科和 overlay；目录名没有明确学科时才回退到正文内容识别，不用每课关键词猜学科。
- WPS 全量绑定按每组开头和结尾锚点确定范围，并用下一组开头限制当前组搜索区间；无文字图片段不会再制造相邻的重复范围。旧 ActionPlan 1.0 兼容分支保留，但新计划不得回退使用 DOCX 段落号。
- 含原生表格的动作同样使用 WPS 开头/结尾锚点夹住真实 Range，并校验表格数量，不预估表格造成的段落偏移。2026-08-02 《课时分层作业5》（两张原生表格）已由用户在 Windows/WPS 确认全量绑定及后续 8/8 选区全部正确；该证据不扩展到 37 份全批生产。
- 2026-08-01 的单样本 Windows/WPS 验收为 8/8 选区通过，但没有按 F1；这只缩小了 WPS Range 绑定的缺口，不改变生产开关。

具体命令行参数以脚本 `--help` 和 `README.md` 为准，不要在本文件里维护第二份命令清单。

## F1 受控执行试点

- ActionPlan 仍保持 `mode=preview_only` 和 `execution_enabled=false`；这份 JSON 本身不授权按键。
- 只有执行器内置的已批准源 SHA256 可进入试录；不认文件名或路径。
- 默认只执行第 1 个动作；`--all` 仍要求每题手工输入 `f1`。
- 控制台输入后必须重新激活 WPS 窗口并恢复选区，再按一次 F1；按键后等待人工检查插件临时结果。
- 未保存文档、非批准哈希和原生表格文档继续阻断。
- 执行器不调用 WPS 保存；插件内容是否最终保存由人工决定。

## 安全开关（全部为 false，不得绕过）

| 开关 | 值 | 含义 |
|---|---|---|
| `execution_enabled` | `false` | ActionPlan 本身不授权真实 F1/F2/F3/F4。单样本只能经独立受控执行器、批准哈希和逐题人工确认试录；不得因此放开批量生产。 |
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
