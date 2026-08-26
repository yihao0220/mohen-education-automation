# 墨痕教育 QuestionUnitV2 题目理解 Implementation Plan

**目标：** 完成《墨痕教育智能分层与模块隔离执行规格》第二部分，在不改变旧 F1 生产入口的前提下，把规则或可选 AI 的候选关系安全编译为只读 `QuestionUnitV2`，或在来源缺失、冲突、低置信度和未知节点出现时整批阻断。

**架构：** 新增纯数据候选层 `QuestionUnderstandingProposal`，每个来源节点以 `RefDecision` 明确记录 include/exclude/unknown、语义角色和证据；确定性编译器只接受上游 `SourceRef` 与候选 JSON，统一检查来源存在、全量覆盖、非法重叠、题号顺序/重置、富内容归属、置信度和规则/AI冲突。另设旧题块规则影子适配器和逐引用差异报告；旧 `QuestionUnit`、`question_core.py`、WPS 与执行器不改。

## 施工范围

### 实现

- 候选对象的严格序列化和反序列化；
- paragraph 的材料、题干、选项、小问、题内标题和顶层边界角色；
- table、question media、decorative media、formula 和 unknown 角色；
- 规则/AI 多候选一致性检查，冲突即阻断；
- 来源存在性、全量覆盖、共享材料例外、非法重叠和来源顺序检查；
- 连续题号合并、显式题号重置和重复题号稳定身份；
- 低于 0.75 的证据、旧题块 warning、unknown 与富内容歧义整批阻断；
- 旧规则候选、只读 DOCX 影子入口、`QuestionUnitV2.json` 与逐引用差异报告；
- 固定 `mode=shadow`、`production_execution_enabled=false`、`keypress_count=0`。

### 不实现

- 不接真实 AI 服务；AI 只能在未来提交同一候选 JSON；
- 不让新题块接管 `墨痕快刀/core_parser.py` 或旧 F1；
- 不生成 F1 ActionPlan V2，不绑定 WPS Range，不执行快捷键；
- 不扩写原始节点提取器；现有 metadata 未暴露的 table/media/formula 仍不能冒充已覆盖；
- 不修改原题、旧 `QuestionUnit`、审核状态或旧 ActionPlan。

## 执行任务

1. 先写构造测试，覆盖规格 4.8 的七类输入以及 unknown、漏分、重叠、低置信度和规则/AI冲突。
2. 实现候选数据对象和确定性编译器，保证失败时 `units=[]`，不产生部分成功结果。
3. 写影子适配测试，再实现旧题块规则候选、逐引用差异和只读 CLI。
4. 联调第一部分契约、现有预检和共享内核；做 AST 导入隔离、Python 编译、UTF-8、源哈希与工作树范围检查。
5. 回写执行规格和问题归档，明确离线证据不证明真实 DOCX 全适配、WPS Range 或 F1 插件录入。

## 验收命令

```bash
uv run --with-requirements requirements-dev.txt python -m pytest -q \
  test_contracts_v2.py \
  test_contract_v2_shadow.py \
  test_question_understanding_v2.py \
  test_question_understanding_v2_shadow.py \
  test_document_preflight.py \
  test_shared_core_flow.py

uv run --with-requirements requirements-dev.txt python -m py_compile \
  shared_core/question_understanding_v2.py \
  shared_core/question_understanding_v2_shadow.py \
  tools/build_question_units_v2_shadow.py \
  test_question_understanding_v2.py \
  test_question_understanding_v2_shadow.py

git diff --check
git status --short
```

## 证明边界

通过上述测试只能证明固定候选能被确定性编译或安全阻断、影子产物只读且执行依赖隔离。它不证明 AI 能理解所有文档，不证明所有真实 DOCX 富内容均已被上游暴露，也不证明 Windows WPS 选区或 F1 插件录入正确。
