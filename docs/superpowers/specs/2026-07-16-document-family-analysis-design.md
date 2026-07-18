# P1a 文档族分组与异常样本发现设计

> 日期：2026-07-16  
> 状态：已批准并按规格实现  
> 范围：只读批次分析，不接管生产 F1

## 1. 目标

在已经生成 `DocumentProfile.json` 1.1 的前提下，为一批文档提供可解释、确定性的结构分析：

1. 判断哪些文档可以归入同一候选文档族；
2. 为每个候选文档族选择一份代表样本；
3. 标出离群文档、无法可靠归组的单例和需要人工优先检查的样本；
4. 生成 JSON 机器真值和由 JSON 单向派生的 Markdown 审核报告。

P1a 回答的是“这一批文档怎样分组”，不改变 P0 对单段角色的判断，也不让角色结果直接控制 WPS 选区。

## 2. 非目标与安全边界

- 不读取、修改、保存或写回原始 DOCX；输入仅为外部 `DocumentProfile.json`。
- 不修改现有 `DocumentProfile.json` 和 `ActionPlan.json`。
- 不连接 WPS COM，不触发 F1/F2/F3/F4。
- 不自动绑定文档族规则，不生成可执行规则快照。
- 不启用题内标题自动排除；`automatic_exclusion_enabled=false` 保持不变。
- 不引入 scikit-learn、向量数据库或不可解释的机器学习模型。
- 不把“单独成族”等同于“错误文档”；证据不足时进入人工复核队列。

输出必须固定包含：

```text
classification_mode = advisory_only
automatic_rule_binding_enabled = false
production_execution_enabled = false
```

## 3. 阶段关系

```text
P0：单文档段落角色证据
  ↓
P1a：只读批次分族、代表样本和异常候选（本次）
  ↓
P1b：真实批次阈值校准与整批审核门禁
  ↓
P1c：人工确认后的文档族规则快照
  ↓
后续阶段：受控生产 F1 消费已审核规则
```

因此，本次完成 P1a 后，“一个任务并不等于一种格式”仍按部分解决记录；只有真实批次校准、审核门禁和规则快照形成闭环后，才评估是否升级为已解决。

## 4. 组件边界

### 4.1 `shared_core/document_families.py`

纯 Python 分析模块，职责限定为：

- 校验并读取 Profile 1.1；
- 从已有指纹和角色证据生成归一化特征；
- 计算两两相似度及可解释分项；
- 进行保守的确定性分组；
- 选择代表样本；
- 生成异常候选和人工复核队列；
- 生成报告数据结构和 Markdown 文本。

不得导入 WPS、`pyautogui` 或题目录入执行模块。

### 4.2 `tools/analyze_document_families.py`

命令行薄入口：

```powershell
python .\tools\analyze_document_families.py `
  .\回归样本\预检基线 `
  --output-dir .\回归样本\文档族分析
```

输入路径可以是一个或多个 Profile JSON 文件或目录。目录只递归发现 `*_DocumentProfile.json`，按规范化绝对路径排序并去重。

输出：

```text
DocumentFamilyReport.json
DocumentFamilyReport.md
```

JSON 是机器权威事实；Markdown 必须完全由同一次运行中的 JSON 数据结构生成，不允许作为后续输入。

### 4.3 `test_document_families.py`

使用最小合成 Profile 验证算法边界，并使用现有三份真实预检基线验证输入兼容、安全开关和确定性。

## 5. 输入契约

每个输入必须满足：

- `schema_version == "1.1"`；
- 存在 `source.name`、`source.sha256`、`source.readonly_input`；
- 存在 `fingerprint.nonempty_paragraph_count`、`question_action_count` 和 `role_counts`；
- 规范化后相同的输入路径直接去重；不同 Profile 若引用相同源 SHA256，则比较移除 `source.path` 后的 Profile 内容：内容一致时保留排序靠前的一份并给出告警，内容冲突时整批失败；
- `source.readonly_input` 必须为 `true`；
- `roles.automatic_exclusion_enabled` 必须为 `false`。

发现不兼容输入时，整批停止并给出中文错误，不跳过坏文件后继续生成看似完整的报告。

## 6. 特征设计

只使用现有 Profile 中已落盘的机器事实，不重新解析 DOCX。

### 6.1 角色分布

对以下角色按非空段落数归一化：

```text
document_title
section_heading
internal_heading
question_start
option
body
```

缺失角色按 0 处理，但未知角色必须保留在报告的兼容性告警中，避免静默丢失新 schema 信息。

### 6.2 密度与存在性

- 题目动作密度：`question_action_count / nonempty_paragraph_count`；
- 标题密度：`heading_count / nonempty_paragraph_count`；
- 媒体、公式密度：各自计数除以非空段落数；
- 原生表格、媒体、公式是否存在；
- 文档规模：非空段落数，仅作低权重旁证，避免不同章节长度主导分族。

所有除法在分母为 0 时显式返回 0，并把空画像列入阻断错误。

## 7. 相似度模型

第一版采用固定、可解释的加权相似度，范围为 `[0, 1]`：

| 分项 | 权重 | 计算原则 |
|------|------|----------|
| 角色分布 | 0.35 | `1 - 0.5 × L1 距离` |
| 题目动作密度 | 0.20 | 标量比例相似度 |
| 标题密度 | 0.10 | 标量比例相似度 |
| 表格/媒体/公式存在性 | 0.15 | 三个布尔特征等权比较 |
| 媒体与公式密度 | 0.10 | 两个密度相似度均值 |
| 文档规模 | 0.10 | `min / max`，只作旁证 |

标量比例相似度规则：两值都为 0 时为 1；只有一值为 0 时为 0；其余为 `min / max`。

每一对文档都输出：总相似度、六个分项分数、权重和加权贡献。报告不能只给一个无法解释的总分。

## 8. 分组算法

采用保守的确定性完全链接聚合：

1. 每份文档先作为一个独立候选族；
2. 计算任意两个候选族之间所有跨族文档对的最小相似度；
3. 选择最小相似度最高的一对候选族；
4. 只有该分数 `>= 0.78` 才合并；
5. 重复直到没有候选族可以合并。

使用完全链接而不是“只要有一对相似就合并”，防止 A 像 B、B 像 C、但 A 与 C 差异很大时发生链式误并。

分数相同时，按候选族成员 SHA256 和文件名组成的稳定键排序，保证相同输入得到字节级稳定结果。

`0.78` 是 P1a 的保守启动阈值，不宣称是生产真值；后续必须使用真实同项目批次校准。

## 9. 代表样本与异常候选

### 9.1 代表样本

每个多成员候选族选择平均族内相似度最高的文档作为代表样本，即 medoid。并列时按 SHA256、文件名稳定排序。

单成员候选族的唯一成员只标为“单例代表”，不因此自动判为异常。

### 9.2 异常候选

批次至少有 3 份文档时，某个单例与所有其他文档的最高相似度 `< 0.60`，才标为 `outlier_candidate`。

其他未合并单例标为 `unresolved_singleton`，进入人工复核，但不宣称异常。

以下情况也进入人工复核队列：

- Profile 自带 `warning` 或 `error` 级 issue；
- 角色证据中存在置信度 `< 0.75` 的低置信度标题候选；
- 输入 schema 出现未知角色；
- 文档族内最小相似度处于固定的 P1a 边界复核区间 `[0.78, 0.82)`。

人工复核队列顺序固定为：异常候选 → 未决单例 → 阈值边界族 → 每族代表样本。

## 10. 输出契约

`DocumentFamilyReport.json` schema 1.0 至少包含：

```text
schema_version
generator
classification_mode
automatic_rule_binding_enabled
production_execution_enabled
input_profiles
feature_definition
similarity_definition
pairwise_similarities
families
outlier_candidates
unresolved_singletons
review_queue
warnings
```

每个候选族包含：稳定 `family_id`、成员、代表样本、族内最小/平均相似度、形成该族的主要共同证据。`family_id` 由排序后的成员 SHA256 摘要生成，不使用运行顺序编号作为机器身份。

Markdown 人工报告按以下顺序展示：

1. 安全状态；
2. 批次摘要；
3. 人工复核队列；
4. 每个候选族及代表样本；
5. 异常候选与最高相似邻项；
6. 两两相似度及分项证据；
7. 阈值仍需真实批次校准的说明。

## 11. 错误处理

以下情况整批失败，不写出半成品报告：

- 找不到 Profile JSON；
- JSON 无法解析或编码异常；
- schema 不是 1.1；
- 缺少必需字段；
- 非空段落数为 0；
- 输入不满足只读证据边界；
- 输出目录与任一输入文件路径冲突；
- 同一 SHA256 出现互相矛盾的 Profile 内容。

写文件时先在输出目录完整序列化并重新读取两个临时文件，验证成功后再替换正式文件；序列化或重读失败时不替换旧报告，并清理本次临时文件。若文件替换本身失败，命令返回失败并明确指出可能只替换了其中一个文件，不把该批次声明为成功。

## 12. 测试与验收

### 12.1 单元测试

- 两份结构等价但规模不同的 Profile 能归入同一族；
- 角色分布明显不同的 Profile 不会因为文件大小接近而误并；
- A-B、B-C 接近但 A-C 不接近时，完全链接阻止链式误并；
- 代表样本是平均族内相似度最高的 medoid；
- 单例不自动等于异常；
- 批次至少 3 份且最高近邻分数低于 0.60 时才产生异常候选；
- 相同输入顺序打乱后，JSON 稳定视图一致；
- Markdown 的数量、family ID、代表样本和安全开关与 JSON 一致；
- 不兼容 schema、重复冲突 SHA 和空画像均整批失败；
- 不产生 WPS 调用，不改变 ActionPlan。

### 12.2 真实基线

使用 `回归样本/预检基线/` 下现有三份 Profile：

- 能成功生成批次报告；
- 输入 Profile SHA256 和内容前后完全一致；
- 对应三份 ActionPlan 的动作数仍为 `1、119、30`；
- `execution_enabled=false` 全部保持不变；
- 输出 UTF-8 中文无乱码，JSON 可重新读取，Markdown 无断链。

三份跨项目、跨学科基线只能验证兼容性和安全边界，不能用于证明 `0.78/0.60` 阈值已经适合生产。阈值生产化必须另取同一真实项目批次校准。

### 12.3 回归命令

```powershell
python -m pytest -q -p no:cacheprovider .\test_document_families.py
python -m pytest -q -p no:cacheprovider .\test_document_preflight.py .\test_document_families.py
```

## 13. 预计变更范围

实现阶段只允许新增或定向修改：

- `shared_core/document_families.py`；
- `tools/analyze_document_families.py`；
- `test_document_families.py`；
- 必要的真实回归报告目录；
- `README.md`、`AGENTS.md`、架构进度文档；
- 项目问题归档和外部长记忆索引/记录。

不修改 `wps_helper.py`，不修改现有题目录入生产解析和执行逻辑。

## 14. 实施完成后的进度口径

若全部验收通过：

- P0：保持已完成；
- P1a：记为已完成并复核；
- 根问题 3“一个任务并不等于一种格式”：仍为部分解决，但缺口缩小为真实同项目批次阈值校准、整批审核门禁和文档族规则快照；
- 6 个根问题总账仍预计为 `1 个已解决、5 个部分解决、0 个未解决`；
- 细分问题“批量画像、结构指纹、自动分组”从部分完成推进到“分组 PoC 已完成，生产校准未完成”。
