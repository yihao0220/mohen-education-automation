# Document Role Evidence Implementation Plan

> 状态：全部任务已实现并复核（2026-07-16）。本计划保留为实施记录；后续工作已转入 P1b 同项目真实批次阈值校准与审核门禁。

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 为 DOCX 只读预检增加可解释、保守且不影响现有题目录入行为的段落角色证据层。

**Architecture:** 原生 python-docx/OOXML 特征继续作为坐标真值；Docling 的元素标签通过单调文本对齐成为可选辅助证据。角色结果只写入 `DocumentProfile` 和派生 Markdown，不进入现有 F1 解析或执行路径。

**Tech Stack:** Python 3.13、python-docx、Docling 2.113.0（可选）、pytest、pytest-regressions。

---

### Task 1: 固化角色分类契约

**Files:**
- Test: `test_document_preflight.py`

1. 构造包含文档标题、章节标题、题内标题、题号、选项和正文的最小 DOCX。
2. 写入角色、证据、保守回退和只读边界断言。
3. 运行 `python -m pytest -q test_document_preflight.py -k role`，确认因角色字段不存在而失败。

### Task 2: 实现原生角色证据层

**Files:**
- Create: `shared_core/document_roles.py`
- Modify: `shared_core/document_preflight.py`

1. 提取段落文本、样式、粗体、字号、编号、位置、上下文等特征。
2. 先识别强排他角色 `question_start` 和 `option`。
3. 对标题类角色进行保守加权评分；低置信度回退 `body/unknown`。
4. 将角色摘要和逐段候选写入 `profile.roles`，不修改 ActionPlan 行为。

### Task 3: 接入 Docling 辅助证据

**Files:**
- Modify: `shared_core/document_preflight.py`
- Modify: `shared_core/document_roles.py`

1. 保存 Docling 文本元素的顺序、文本预览和标签。
2. 使用规范化文本和单调游标绑定原生段落；歧义时放弃绑定。
3. 将 `title/section_header/list_item/paragraph` 转换为辅助证据，不覆盖原生强反证。

### Task 4: 派生审核视图与回归验证

**Files:**
- Modify: `shared_core/document_preflight.py`
- Test: `test_document_preflight.py`
- Test: `test_document_preflight/test_minimal_bundle_golden_json.json`

1. Markdown 显示角色计数和标题候选的证据、置信度。
2. 运行定向测试并更新黄金 JSON。
3. 运行共享内核相关测试，确认题块和动作计划没有回归。
4. 对三份真实样本生成临时产物、检查 SHA256/UTF-8/JSON/MD 后清理临时目录。

### Task 5: 更新进度账本与双轨归档

**Files:**
- Modify: `docs/墨痕教育架构问题工程思维分析拆解.md`
- Modify: `问题归档/题目录入/INDEX.md`
- Create: `问题归档/题目录入/2026-07-16-DOCX题内角色证据层.md`
- Modify: `../墨痕教育-记忆库/INDEX.md`
- Create or modify: `../墨痕教育-记忆库/模块问题/跨模块项目/...`

1. 记录现象、根因、实现、测试命令、真实样本和下一步入口。
2. 更新 6 个根问题与 14 条细粒度审计状态，但不把“只读候选”误记为生产完成。
3. 检查 Markdown 链接、乱码、临时产物和无关文件变更。

> 本计划不执行 git commit；只有用户明确要求时才提交。
