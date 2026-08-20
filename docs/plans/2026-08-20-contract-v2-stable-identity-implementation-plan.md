# 墨痕教育 V2 交接单与稳定身份 Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 完成《墨痕教育智能分层与模块隔离执行规格》的第一部分，在不改变任何 F1/F2/F3/F4 生产行为的前提下，建立可序列化、可验证、可追溯、可影子对照的 V2 交接单。

**Architecture:** 在 `shared_core` 新增纯标准库契约层，定义 `SourceRef`、`StableUnitId`、`DecisionEvidence`、`QuestionUnitV2` 与 `CanonicalAnswerUnit`；所有来源通过 included/excluded/unknown 三分区表达，构造时强制完整且互斥。另设影子适配层，把现有 `QuestionUnit` 转为 V2 JSON 和差异报告；旧对象和旧入口保持不变，WPS、`pyautogui` 和快捷键模块不得被导入。

**Tech Stack:** Python 3、`dataclasses`、`json`、`hashlib`、JSON Schema Draft 2020-12、pytest、python-docx（仅影子 DOCX 入口沿用现有扫描器）。

---

## 本次施工部分

### 目标

只完成：

- 版本化 V2 契约对象及 JSON Schema；
- 同题号重复出现时仍唯一的稳定身份；
- paragraph/table/media/formula 四类原生引用；
- included/excluded/unknown 完整互斥分区；
- 稳定 JSON 字节视图和严格反序列化；
- 现有 `QuestionUnit` 的只读影子适配与差异报告；
- 不接 WPS 的命令行影子产物。

### 明确不做

- 不改 `QuestionUnit`、`AnswerUnit` 的现有字段和生产消费者；
- 不让 `QuestionUnitV2` 接管题块判断；
- 不实现 AI、F1 ActionPlan V2、F2/F3/F4 ActionPlan 或简单执行器；
- 不修改 `wps_helper.py`；
- 不触发真实快捷键；
- 不提交、不推送，除非用户另行明确授权。

### 当前输入契约

- 纯契约测试：手工构造的 `SourceRef` 集合；
- 影子适配：现有 `DocNode`、`QuestionUnit`、源 DOCX SHA256；
- 真实样本：只读 DOCX，运行前后必须核对 SHA256。

### 当前输出契约

- `QuestionContractV2.json`：`schema_version=2.0`，`mode=shadow`，`production_execution_enabled=false`；
- `QuestionContractV2Diff.json`：只比较身份、标签、来源覆盖和原生类型，不宣称题块语义正确；
- JSON 使用 UTF-8、稳定键序和确定性数组顺序；
- Markdown 不作为机器输入。

### 允许修改的文件

- 新建 `shared_core/contracts_v2.py`；
- 新建 `shared_core/contract_v2_shadow.py`；
- 新建 `shared_core/schemas/question-unit-v2.schema.json`；
- 新建 `shared_core/schemas/canonical-answer-unit-v2.schema.json`；
- 新建 `tools/build_question_contract_v2_shadow.py`；
- 新建 `test_contracts_v2.py`、`test_contract_v2_shadow.py`；
- 新建第一部分问题归档并更新两级索引；
- 必要时只补 README/规格的实施状态，不改生产规则。

### 禁止修改的模块

- `墨痕快刀/wps_helper.py`；
- `答案录入/answer_input.py`；
- `格式处理/main.py`；
- 现有 `shared_core/models.py` 数据结构；
- 现有 `DocumentProfile.json` 与 `ActionPlan.json` schema。

### 假上游

- 固定的 SourceRef 列表；
- 手工构造的旧 `QuestionUnit` 和 `DocNode`。

### 假下游

- 只读取和反序列化 V2 JSON；
- 不生成 ActionPlan，不连接 WPS。

### 回退开关

- 新工具只允许 `shadow`；
- 生产入口没有导入新模块，因此回退方式是停止运行影子工具；
- 影子失败不改旧 JSON、不改源文件、不改生产状态。

## 开发前复盘

- 任务类型：跨模块契约地基，属于核心架构施工。
- 相关历史经验：旧 `QuestionUnit`/`AnswerUnit` 先作为共享结果层落地；`DocumentProfile` 与 `ActionPlan` 均保持 JSON 为机器真值、Markdown 为派生说明；角色证据先旁路，不直接改变 F1。
- 主要风险：把题号当唯一 ID；用连续 `source_span` 冒充非连续来源；契约层反向依赖 WPS；影子产物被误当生产计划；为了接 V2 直接修改旧消费者。
- 验证方式：构造输入单元测试、JSON 往返与黄金字节、AST 导入隔离、只读真实 DOCX 影子运行、现有预检/共享内核回归。
- 归档要求：先更新 `问题归档/题目录入/INDEX.md`，再新增单问题记录；长期记忆只在用户触发“检查归档”时更新。

---

### Task 1: 契约原语与 JSON Schema

**Files:**

- Create: `shared_core/contracts_v2.py`
- Create: `shared_core/schemas/question-unit-v2.schema.json`
- Create: `shared_core/schemas/canonical-answer-unit-v2.schema.json`
- Test: `test_contracts_v2.py`

**Step 1: 写失败测试**

覆盖：

- SHA256、原生类型、occurrence、confidence 和空关键字段校验；
- paragraph/table/media/formula 的引用键互不混淆；
- 同一文档两次第 1 题因 occurrence 不同而稳定 ID 不同；
- StableUnitId 的派生值被篡改时反序列化失败；
- 两份 schema 的 `$id`、版本和共同 `$defs` 字段一致。

**Step 2: 验证测试先失败**

Run: `python -m pytest -q test_contracts_v2.py`

Expected: FAIL，原因是 `shared_core.contracts_v2` 尚不存在。

**Step 3: 实现最小纯契约层**

- 使用 frozen dataclass；
- 仅依赖 Python 标准库；
- `StableUnitId.value` 由版本化规范 JSON 的 SHA256 派生；
- `SourceRef.key` 包含文档哈希、native_kind、native_id、occurrence 和 text_anchor；
- 提供 `to_dict()` / `from_dict()`，不接受未知 schema 版本。

**Step 4: 写入两份 JSON Schema**

- 使用 Draft 2020-12；
- `additionalProperties=false`；
- 四类原生节点使用枚举；
- SHA256 使用 64 位小写十六进制 pattern；
- 题目和答案对象都引用同构的 SourceRef、StableUnitId、DecisionEvidence 定义。

**Step 5: 运行测试**

Run: `python -m pytest -q test_contracts_v2.py`

Expected: PASS。

### Task 2: V2 题目/答案对象和完整来源分区

**Files:**

- Modify: `shared_core/contracts_v2.py`
- Modify: `test_contracts_v2.py`

**Step 1: 写失败测试**

覆盖：

- 中间标题进入 `excluded_refs`，前后正文进入 `included_refs`；
- included/excluded/unknown 交叉时阻断；
- expected scope 中存在未分类节点时阻断；
- 多出的非 scope 引用时阻断；
- 序列化再读取后引用、证据和稳定 ID 不变；
- 同一输入重复生成 canonical bytes 完全一致；
- `CanonicalAnswerUnit` 保留答案项、解析项与 rich refs，但不决定 F2/F4。

**Step 2: 验证测试先失败**

Run: `python -m pytest -q test_contracts_v2.py -k 'partition or roundtrip or canonical or answer'`

Expected: FAIL，原因是 V2 对象或分区校验尚未实现。

**Step 3: 实现来源分区**

- `build_source_partition(expected_refs, included_refs, excluded_refs, unknown_refs)`；
- 先拒绝分区内部重复，再拒绝分区交叉；
- 精确比较 expected 与三分区并集；
- 不把 explicit unknown 当作静默丢失，但在对象上暴露 `has_unknown_refs`。

**Step 4: 实现 V2 对象和稳定 JSON**

- `QuestionUnitV2` 保存稳定 ID、原题标签、学科/覆盖层/题型、三分区和证据；
- `CanonicalAnswerUnit` 保存稳定 ID、原题标签、答案项、解析项、rich refs、三分区和证据；
- `canonical_json_bytes()` 固定 UTF-8、`sort_keys=True`、紧凑 separators 和尾换行；
- `from_dict()` 严格复核派生 ID 和所有分区。

**Step 5: 运行测试**

Run: `python -m pytest -q test_contracts_v2.py`

Expected: PASS。

### Task 3: 旧题目对象影子适配与差异报告

**Files:**

- Create: `shared_core/contract_v2_shadow.py`
- Create: `tools/build_question_contract_v2_shadow.py`
- Create: `test_contract_v2_shadow.py`

**Step 1: 写失败测试**

覆盖：

- 旧 `QuestionUnit.source_span` 只转换为 V2 included refs，不修改旧对象；
- 同题号按出现次序分配不同稳定 ID；
- 一个 DOCX 段落拆成多个虚拟节点时 occurrence 稳定；
- diff 报告逐题列出 legacy span、V2 ref 数、原生类型计数和重复标签；
- 所有产物固定 `mode=shadow`、`production_execution_enabled=false`；
- AST 扫描确认契约层/影子层不导入 WPS、`win32com`、`pyautogui` 或执行模块。

**Step 2: 验证测试先失败**

Run: `python -m pytest -q test_contract_v2_shadow.py`

Expected: FAIL，原因是影子模块和 CLI 尚不存在。

**Step 3: 实现影子适配**

- 复用现有 `scan_docx_nodes()` 和 `build_question_units_from_docx()`；
- 使用源文件 SHA256 和单位出现序号生成身份；
- 不改变旧 `QuestionUnit`，不写回 DOCX；
- 生成独立 V2 契约与差异 JSON，不改现有预检 bundle。

**Step 4: 实现只读 CLI**

Run: `python tools/build_question_contract_v2_shadow.py SOURCE.docx --output-dir OUTPUT`

- 输出目录必须显式提供；
- 拒绝把产物写入源文件所在目录；
- 运行前后校验源 SHA256；
- 控制台摘要明确“影子模式、按键数 0、生产行为未改变”。

**Step 5: 运行测试**

Run: `python -m pytest -q test_contract_v2_shadow.py`

Expected: PASS。

### Task 4: 真实样本、相邻契约与回归验证

**Files:**

- Modify: `test_contract_v2_shadow.py`（仅在需要固定最小 DOCX 样本时）
- No production files.

**Step 1: 最小真实 DOCX 只读验证**

- 在 pytest `tmp_path` 生成含重复题号、标题、表格、媒体/公式描述的最小输入；
- 运行影子 CLI；
- 核对源 SHA256 不变、JSON 可重开、无 U+FFFD、无快捷键依赖。

**Step 2: 运行相邻契约测试**

Run: `python -m pytest -q test_contracts_v2.py test_contract_v2_shadow.py test_document_preflight.py test_shared_core_flow.py`

Expected: 全部 PASS；若真实样本条件缺失，只允许与样本缺失直接相关的既有 skip。

**Step 3: 运行语法和导入隔离检查**

Run: `python -m py_compile shared_core/contracts_v2.py shared_core/contract_v2_shadow.py tools/build_question_contract_v2_shadow.py test_contracts_v2.py test_contract_v2_shadow.py`

Expected: 无输出，退出码 0。

**Step 4: 清理临时产物**

- 删除非 `tmp_path` 的临时 JSON/DOCX；
- `git status --short` 不得出现计划外生成物。

### Task 5: 归档与实施状态

**Files:**

- Modify: `问题归档/题目录入/INDEX.md`
- Create: `问题归档/题目录入/2026-08-20-V2统一交接单与稳定身份.md`
- Modify: `docs/superpowers/specs/2026-08-20-mohen-intelligence-release-execution-spec.md`

**Step 1: 先更新模块索引**

- 新增第一部分记录，状态必须区分离线契约完成与生产未接管。

**Step 2: 写单问题记录**

- 记录现象、根因、改动点、测试样例、测试命令、结论、下次排查入口；
- 明确离线测试不证明 WPS 定位和真实插件录入。

**Step 3: 回写规格状态**

- 只把第一部分标为“影子契约已实现并离线验证”；
- 第二至第十部分保持未实施；
- 不把 V2 契约写成生产已启用。

**Step 4: 最终复核**

Run: `git diff --check && git status --short`

Expected: 无空白错误；只出现本计划声明的文件。

## 证明了什么

- 交接单能表达重复题号、非连续来源和不同原生节点类型；
- 同一输入可以产生稳定、可往返的 JSON；
- 未分类或重复分类来源会被阻断；
- 旧题目对象可在不改变生产链的情况下生成 V2 影子结果；
- 纯契约和影子模块与 WPS 执行依赖隔离。

## 尚未证明什么

- 不证明题块判断本身正确；
- 不证明答案提取与题答映射正确；
- 不证明 SourceRef 已能绑定真实 WPS Range；
- 不证明 F1/F2/F3/F4 插件行为；
- 不证明 V2 已经可以替换任何生产入口。

## 归档位置

- 项目证据：`问题归档/题目录入/2026-08-20-V2统一交接单与稳定身份.md`；
- 外部长期记忆：本轮不更新，除非用户明确触发“检查归档”。
