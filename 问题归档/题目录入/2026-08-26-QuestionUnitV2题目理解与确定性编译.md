# QuestionUnitV2题目理解与确定性编译

## L1 摘要

- 模块：共享内核 / 题目录入影子旁路
- 日期：2026-08-26
- 状态：第二部分离线实现完成；生产 F1 未接管
- 结论：新增候选建议对象、确定性编译器、旧规则影子适配、只读 CLI 和逐引用差异报告。规则或未来 AI 只能提交来源关系；缺失、冲突、低置信度、未知、非法重叠、题号异常或富内容歧义均整批阻断。新链固定 `mode=shadow`、`production_execution_enabled=false`、`keypress_count=0`。

## L2 现象与根因

### 现象

- 旧 `QuestionUnit.source_span` 只能表达连续首尾范围，题块中间的标题无法在交接单中单独排除。
- 旧规则、项目覆盖层和 WPS 运行时共同承担题块判断，低置信度和 warning 不总是在同一层阻断。
- 材料、题图、装饰图、表格和公式即使被识别，也缺少统一的“属于哪个题块或为什么排除”候选格式。
- 只比较新旧题块数量会漏掉题块内部来源和媒体归属差异。

### 根因

- 第一部分只建立 V2 身份和来源契约，旧题块影子适配仍把连续范围全部视为 included，没有新的题目理解编译阶段。
- 规则或 AI 的建议若直接变成 WPS 范围，就跳过来源存在、覆盖、冲突、顺序和富内容安全检查。
- 真实 DOCX 原始节点提取仍有 table/formula 暴露限制，不能靠下游猜测补齐。

## L3 改动点

### 候选与确定性编译

- 新增 `shared_core/question_understanding_v2.py`：
  - `RefDecision` 明确记录每个 `SourceRef` 的 include/exclude/unknown、角色和 `DecisionEvidence`；
  - `ProposedQuestionUnit` 记录题号、连续合并成员、学科、题型、来源判断和审核标记；
  - `QuestionUnderstandingProposal` 作为规则、AI或人工候选的版本化 JSON 交接单；
  - 多候选语义不一致时触发 `PROPOSAL_CONFLICT`，AI不能覆盖规则或直接生成 WPS Range；
  - 编译器检查来源存在、完整覆盖、非法重叠、来源顺序、连续合并、显式题号重置、低置信度、unknown 和富内容归属；
  - 任一 blocker 出现时固定 `units=[]`，不返回部分成功题块。

### 影子适配与差异

- 新增 `shared_core/question_understanding_v2_shadow.py`：
  - 复用旧 `QuestionUnit` 作为规则候选起点，但把题内标题、装饰媒体和富内容转换为逐引用判断；
  - 旧题块范围外的非空未知节点显式进入 unknown 并阻断，不沉默丢失；
  - 差异报告逐题列出每个 ref 的 native kind、disposition、role 和媒体归属；
  - 运行前后核对原题 SHA256，产物必须写到源目录之外。
- 新增 `tools/build_question_units_v2_shadow.py`，只生成候选、`QuestionUnitV2` 和差异 JSON；不连接 WPS。
- 没有修改 `shared_core/question_core.py`、`shared_core/models.py`、`墨痕快刀/core_parser.py`、任何 `wps_helper.py` 或生产入口。

## L4 测试样例与命令

构造测试覆盖：

- 普通连续题；
- 前置材料只随首题；
- 完整阅读组按连续题号合并；
- 中间标题 excluded，前后正文 included；
- 标题同时承载原生表格时阻断；
- 装饰媒体 excluded、题图 included；
- 重复题号必须有明确 reset 声明，稳定 ID 仍不同；
- unknown、漏分、非法重叠、低置信度和规则/AI冲突均整批阻断；
- 真实最小 DOCX 影子运行前后哈希不变，三份 JSON 可按 UTF-8 重开且无 U+FFFD；
- AST 确认候选、编译和影子模块不导入 WPS、Renderer、答案录入或快捷键依赖。

验收命令见 `docs/plans/2026-08-26-question-understanding-v2-implementation-plan.md`。

- 第二部分、第一部分契约、现有预检和共享内核相邻回归：`47 passed in 4.90s`；
- 定向 Python 编译：退出码 0，无语法错误；
- CLI 最小真实 DOCX 实跑：`status=compiled_for_review`、`unit_count=2`、`production_execution_enabled=false`、`keypress_count=0`；
- CLI 生成候选、题块、逐引用差异三份 JSON，均可按 UTF-8 重开，无 U+FFFD；源文件运行前后 SHA256 由工具和测试双重核对；
- 临时 DOCX 和 JSON 已清理，仓库未残留业务产物。

## L5 结论、边界与回退

### 已证明

- 固定规则或 AI 候选可以无执行副作用地编译成 `QuestionUnitV2`，或在不安全时整批阻断；
- 每个上游来源必须被题块判断或文档级判断覆盖，表格、图片和公式不能静默消失；
- 题内标题可以在一个题块内部单独排除，保留其前后正文；
- 重复题号、显式重置、连续合并、来源顺序和媒体归属均有确定性检查；
- 新旧题块可以逐引用比较，旧生产入口保持不变。

### 未证明

- 不证明真实 AI 对所有新文档都能给出正确候选；本次未连接任何 AI 服务；
- 不证明现有 DOCX 扫描器已经暴露全部原生表格和公式；未暴露对象不能冒充已覆盖；
- 不证明新结果已经接管 F1、能绑定 Windows WPS Range 或能被插件正确保存；
- 未执行 Windows/WPS 测试，因为第二部分明确只输出语义对象，不生成或执行动作计划。

### 回退

- 生产入口没有导入第二部分模块；停止运行影子 CLI 即回到原状态。
- 影子文件使用独立名称，不覆盖第一部分契约、DocumentProfile、ActionPlan 或审核状态。

### 下次入口

- 候选与编译：`shared_core/question_understanding_v2.py`；
- 旧规则影子适配：`shared_core/question_understanding_v2_shadow.py`；
- CLI：`tools/build_question_units_v2_shadow.py`；
- 下一部分：第三部分 F1 ActionPlan V2，只能消费审核后的 `QuestionUnitV2`，不得重新判断题块。
