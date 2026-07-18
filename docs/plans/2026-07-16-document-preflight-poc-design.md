# DOCX 只读画像与动作计划 PoC 设计

> 状态：已实现并复核（2026-07-16）。该文档保留为设计基线，后续 P0 角色证据层与 P1a 文档族分析均在其只读安全边界上演进。

## 目标

在不修改、不保存、不覆盖原始 DOCX，且不连接 WPS `F1/F2/F3/F4` 的前提下，建立一条可重复验证的旁路：

```text
原始 DOCX（只读）
→ 现有 shared_core 扫描（坐标真值）
→ Docling 结构对照（可选增强）
→ DocumentProfile.json（机器事实）
→ ActionPlan.json（预演契约）
→ 对应 Markdown（人工审核视图）
```

## 方案比较

### 方案 A：Docling 替换现有扫描器

优点是统一结构模型快；缺点是不能证明其元素能稳定映射回 WPS 原始段落与 Range，会破坏已经验证的原生表格和坐标边界。不采用。

### 方案 B：旁路 sidecar（采用）

现有 `scan_docx_nodes()` 和 `QuestionUnit.source_span` 继续作为原始坐标真值；Docling 只输出结构计数与标签分布，用来发现差异。所有产物写到源文档之外，原文件前后校验 SHA256。

### 方案 C：只写一次性比较脚本

改动最小，但结果无法作为后续规则组合、审核门禁和执行器的稳定输入。不采用。

## 数据契约

### DocumentProfile

- `schema_version`：画像 Schema 版本。
- `source`：名称、绝对路径、大小、SHA256、只读校验结果。
- `native`：python-docx/OOXML 得到的段落、表格、媒体、公式和标题计数。
- `docling`：Docling 版本、转换状态、结构元素计数和标签分布。
- `comparison`：两种扫描结果的可比字段和差异。
- `fingerprint`：用于后续文档分族的稳定结构指纹。
- `issues`：预检风险，不直接触发 WPS。

### ActionPlan

- `schema_version`：动作计划 Schema 版本。
- `source_sha256`：与画像绑定，源文件改变后计划失效。
- `execution_enabled=false`：PoC 固定关闭执行。
- `actions`：现有 `QuestionUnit` 编译出的 F1 预演动作。
- `source_ref`：虚拟节点范围、原始段落范围、表格索引、媒体节点索引。
- `blocking_issues`：未绑定 WPS Range、未识别题块等阻断原因。

JSON 是权威机器事实；Markdown 只能从 JSON 单向生成，不作为机器执行输入。

## 错误处理

- Docling 未安装或转换失败：现有扫描继续完成，画像记录警告。
- 源文件哈希前后不一致：立即失败，不写出可用计划。
- 无题块：生成画像，但计划包含阻断问题。
- 表格计数不一致：记录结构差异，禁止据此修改原题。
- PoC 阶段：无论识别结果如何，都不执行 WPS 动作。

## 样本与验证

1. `对点练案6.docx`：纯文本基线。
2. `五年级暑假数拔教材.docx`：原生表格、媒体和公式。
3. `对点练案31.docx`：众美完整阅读题分组和 VML 媒体。

验证要求：

- 三份源文件前后 SHA256 一致。
- JSON 可重新读取，UTF-8 中文无乱码。
- Markdown 与 JSON 中的题块数、阻断状态一致。
- pytest 覆盖哈希不变、表格锚点、题块坐标、JSON/MD 单向生成。
- pytest-regressions 固化一个最小文档的黄金 JSON。

## 非目标

- 不实现文档族自动聚类。
- 不接入 Hydra 规则组合。
- 不修改现有题目识别规则。
- 不连接或调用 WPS COM。
- 不保存 Docling 的完整文档副本。
