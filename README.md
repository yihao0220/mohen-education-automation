# 墨痕教育自动化录入系统

为 WPS「墨痕题库」插件定制的自动化工具链，覆盖题目录入、答案格式清洗、答案/解析录入三步流程。

## 当前状态

截至 2026-07-18，系统已形成四个业务入口，并新增三条不接管生产流程的 DOCX 只读分析旁路：

| 模块 | 入口 | 状态 | 说明 |
|------|------|------|------|
| 任务面板 | `main.py` | 稳定可用 | 自动配对题目 / 原始答案 / 已清洗答案 / 审核状态，并给出下一步动作 |
| 题目录入 | `main.py` / `墨痕快刀/main.py` | 稳定可用 | 英语、理科、文科三核 + 学科覆盖层 |
| 答案格式清洗 | `格式处理/main.py` | 稳定可用 | 多模板清洗，未来高二物理、未来高二历史与安乡金海初二数学模板已接入 |
| 答案录入 | `答案录入/answer_input.py` | 稳定可用 | F2/F3/F4 自动录入，依赖审核状态文件 |
| DOCX 只读预检 | `tools/build_document_preflight.py` | PoC + P0 可用 | 生成 Profile 1.1、题内角色证据、Docling 对照和 F1 预演计划；固定不连接 WPS、不执行按键 |
| 文档族分析 | `tools/analyze_document_families.py` | P1a 可用 | 读取一批 Profile 1.1，生成候选文档族、代表样本、异常候选和人工复核队列；固定只作建议 |
| 页面视觉预检 | `tools/build_document_render.py` | P1b 开发预览可用 | Mac 用 Quick Look 生成连续视觉预览；Windows 用 WPS COM 生成生产页面真值；两者均不执行 F1/F2/F3/F4 |

架构问题主口径共 6 个：当前 **1 个已解决、5 个部分解决、0 个未解决**。严格完成口径是 `1/6`，已进入工程解决口径是 `6/6`；详细进度见[《墨痕教育架构问题工程思维分析拆解》](./docs/墨痕教育架构问题工程思维分析拆解.md)。

## 两台电脑的职责

- Mac：开发、测试和 Quick Look 视觉预览。它能识别页面上的题目文字、题图与装饰图，但连续预览不是 WPS 分页真值。
- Windows：WPS 生产验证与最终录入。只有 `wps_com` 或明确传入的 WPS PDF 才能标记 `page_truth_authority=true`。
- 两端共用同一 Git 仓库；原题、答案、页面截图、JSON 批次产物和凭据不提交到仓库。

日常同步采用普通 Git 工作流：

```bash
# Mac 或 Windows 开工前
git pull --ff-only

# 修改完成并验证后
git add <本次修改文件>
git commit -m "说明本次修改"
git push
```

Windows 首次使用时，把私有仓库克隆到代码目录，不要克隆进题目/答案业务目录：

```powershell
git clone <你的私有仓库HTTPS地址> D:\CODEX.projection\墨痕教育
Set-Location D:\CODEX.projection\墨痕教育
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\setup_windows.ps1
```

源文档和 P1b 产物继续放在仓库外部，两台电脑各自通过 `工作台路径配置.json` 指向本机业务目录；仓库只提供不含个人路径的 `工作台路径配置.example.json`。

## 安装

Mac：

```bash
./scripts/setup_macos.sh
```

Windows PowerShell：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\setup_windows.ps1
```

## 生产运行前提

1. Windows 环境。
2. 已运行 `scripts/setup_windows.ps1`。
3. 先打开 WPS 文档，并加载「墨痕题库工具」侧边栏。
4. 插件快捷键需可用：`F1` 题目，`F2` 答案，`F3` 解析，`F4` 小题答案。

## 标准流程

```powershell
# Step 0: 查看任务面板（推荐）
python main.py

# Step 1: 题目录入
python .\墨痕快刀\main.py

# Step 2: 答案清洗
python .\格式处理\main.py

# Step 3: 答案/解析录入
python .\答案录入\answer_input.py
```

开发新批次或核对结构规则时，可在生产流程之外先运行只读预检：

```powershell
.\.venv\Scripts\python.exe .\tools\build_document_preflight.py "D:\...\原题.docx" --output-dir ".\回归样本\预检基线"
```

它只生成外部 `DocumentProfile.json`、`ActionPlan.json` 及派生 Markdown。当前 `DocumentProfile` schema 为 1.1，包含文档标题、板块标题、题内小标题、题号、选项、正文等角色及其置信度和证据。JSON 是机器权威事实，Markdown 只供人工审核；角色层固定 `automatic_exclusion_enabled=false`，动作计划固定 `execution_enabled=false`，均不能直接驱动 WPS。

已经生成一批 Profile 1.1 后，可以继续运行 P1a 只读文档族分析：

```powershell
.\.venv\Scripts\python.exe .\tools\analyze_document_families.py .\回归样本\预检基线 --output-dir .\回归样本\文档族分析
```

它生成 `DocumentFamilyReport.json/.md`，采用可解释加权相似度和完全链接分组，输出候选文档族、代表样本、异常候选与人工复核队列。当前固定 `classification_mode=advisory_only`、`automatic_rule_binding_enabled=false`、`production_execution_enabled=false`；阈值必须用同一真实项目批次校准后，才能进入后续审核门禁。

P1b 单文档视觉预览必须把产物写到原题目录之外。Mac 开发机运行：

```bash
.venv/bin/python tools/build_document_render.py \
  "/绝对路径/原题.docx" \
  --output-dir "/绝对路径/墨痕教育-P1b产物/文档名"
```

Windows 生产机运行同一入口，`auto` 会改用 WPS COM：

```powershell
.\.venv\Scripts\python.exe .\tools\build_document_render.py `
  "D:\原题\作业1.docx" `
  --output-dir "D:\墨痕教育-P1b产物\作业1"
```

输出的 `PageRenderManifest.json` 是机器事实；Mac 结果固定 `page_truth_authority=false`，不能通过 P1b 生产校准门禁。Windows 结果仍需人工完成 `VisualReview.json`，整批审核通过后才允许生成校准报告；当前所有生产执行开关仍为 `false`。

清洗后的答案文档必须先通过审核门禁；用户手改 `_已清洗.docx` 后，视为审核状态失效，需要重新审核。

所有学科、学校和项目中的表格题统一保留原生表格：直接整块框选题干、表格和选项后按 `F1` 录入，不再转成纯文本或创建临时文档。

`main.py` 的任务面板会把文档自动归到这些阶段之一：

- `可录答案`
- `待清洗`
- `自动检查未通过`
- `清洗结果已过期`
- `缺状态文件`

数学答案补充规则：

- 能稳定转成纯文本的数学内容保留在答案区。
- 不能稳定进入答案框的公式对象转入解析区承载。
- 数学斜体 Unicode 字母（如 `𝑥 / 𝑦 / 𝑏 / 𝑘`）会在清洗阶段归一化为普通字母，避免录题侧显示成方框。

## 目录速查

| 路径 | 用途 |
|------|------|
| `墨痕快刀/` | 题目录入核心，包含 WPS COM 扫描、题块识别、F1 录入 |
| `格式处理/` | 答案清洗模板系统 |
| `答案录入/` | 答案、解析、小题答案自动录入 |
| `shared_core/` | 题目/答案共享内核、审核门禁、学科策略；`document_preflight.py` 提供只读画像与预演动作计划，`document_roles.py` 提供题内角色证据，`document_families.py` 提供只读文档族分析 |
| `tools/` | 拆分、检查、批处理工具 |
| `docs/` | 架构分析与工程设计基线；入口见[《墨痕教育架构问题工程思维分析拆解》](./docs/墨痕教育架构问题工程思维分析拆解.md) |
| `问题归档/` | 每次代码问题修复的证据归档 |
| `回归样本/` | 固定回归样本入口；`预检基线/` 保存单文档 JSON/MD 基线，`文档族分析/` 保存 P1a 批次报告 |

## 重点项目

众美高三语文（对点练）题目录入已经完成专项开发：

- 项目归档：`../墨痕教育-记忆库/模块问题/题目录入/2026-07-13-众美对点练完整阅读题整组录入.md`
- 代码证据：`问题归档/题目录入/2026-07-13-众美对点练完整阅读题整组录入.md`
- 关键测试：`test_zhongmei_chinese_workflow.py`
- 当前成果：44 份题目文档稳定形成 296 个录入题块，完整阅读题按“文章 + 所属全部题目”整组录入，展开后仍覆盖 400 道原题
- 题答验证：44 份真实答案已与分组题块全量映射，无孤立答案块

众美对点练中，`leading_context_patterns` 负责定位阅读提示，`group_leading_context_questions` 负责将文章及下一个阅读提示或顶层大板块边界前的全部所属题归入同一题块。题块内部的“高考专练/典题专练”由 `question_input_excluded_patterns` 标记，F1执行层在不保存的临时富文本副本中删除标题后整组录入，原题不改写。新项目应由覆盖层显式选择是“只随首题”还是“完整阅读整组”，不要修改全局默认。

同批 44 份答案文档已完成专项关键词清洗：

- 清洗入口：`tools/clean_zhongmei_chinese_answers.py`
- 回归测试：`test_zhongmei_chinese_answer_cleaner.py`
- 答案录入证据：`问题归档/答案录入/2026-07-14-众美答案小标题不再录入解析.md`、`问题归档/答案录入/2026-07-14-众美题型标题过滤导致答案录入段号错位.md`、`问题归档/答案录入/2026-07-14-圆圈序号答案列点不再误触发F4.md`
- 运行方式：`python .\tools\clean_zhongmei_chinese_answers.py --preflight-only` 先只读预检；确认通过后去掉 `--preflight-only` 批量输出 `对点练案N_已清洗.docx`
- 当前成果：44 份文档、400 道题、409 个答案块；56 道缺少源解析的题已补空白 `解析：`
- 安全边界：整批写入前必须全部通过预检；遇到未知答案标记、表格答案或题号不连续时停止，不猜测清洗
- 验证结论：8 项自动化测试通过，44 份输出均回读一致，Word 导出的 139 页预览已逐页目检
- 答案录入：只有明确的 `(1)(2)…` / `（1）（2）…` 小问和完整阅读组连续使用 `F4`；`①②③…` 是单题答案内部列点，整段只使用一次 `F2`；同组各小题解析合并后只录入一次 `F3`
- 标题边界：训练标题、题型小标题及“阅读下面……完成题目”提示只供人工阅读，不进入 `F2/F3/F4` 选区；旧文档中粘在解析末段的标题也会被剥离
- 录入坐标：实际 WPS 选区必须保留清洗文档的原始段落位置；过滤标题和空段只能改变语义载荷，不能压缩 `F2/F3/F4` 使用的段落坐标
- 门禁状态：44 份 `_已清洗.docx` 的审核状态已刷新，当前全部可通过答案录入门禁

众美高三语文（文言文）答案已完成专项清洗：

- 清洗入口：`tools/clean_zhongmei_classical_chinese_answers.py`
- 运行方式：`python .\tools\clean_zhongmei_classical_chinese_answers.py --preflight-only` 先整批只读预检；通过后去掉 `--preflight-only` 正式生成
- 当前成果：28 份 `_已清洗.docx` 及 28 份审核状态，覆盖 154 个录入题块、196 个源答案标记
- 编号规则：双篇文档保留各篇原题号重置；仅 `20《苏武传》` 将原 `1、2、5、6` 修正为 `1、2、3、4`
- 内容规则：33 个翻译题块统一输出“全部 `(n)答案` → 单一 `解析：` → 全部 `(n)得分点`”；共 74 个小题连续使用 F4，每个翻译题块只使用一次 F3，单小题翻译也保留 `(1)` 和 F4；330 个得分点全部移入解析
- 活动文档纠偏：若 WPS COM 的活动文档仍是众美文言文原题，答案录入会先定位同名 `_已清洗.docx`，仅在审核门禁通过后自动打开并激活；缺文件或状态失效时继续阻断
- 验证结论：154 个答案块全部映射、审核问题为 0、审核门禁 28/28 放行；答案录入聚焦测试为 22 passed、108 subtests passed，众美与共享核心回归为 94 passed、591 subtests passed；59 页 Word 预览已目检，56 份原始题答文档未改动
- 代码证据：`问题归档/格式转换/2026-07-14-众美高三语文文言文答案清洗.md`、`问题归档/答案录入/2026-07-15-众美文言文答案录入自动纠正WPS活动原题.md`

未来高二历史选必一已经完成项目级归档：

- 总归档：`../墨痕教育-记忆库/模块问题/跨模块项目/2026-06-17-未来高二历史选必一开发总归档.md`
- 模板经验：`格式处理/模板开发经验记录/模板FutureHistory开发经验记录.md`
- 关键测试：`test_future_history_workflow.py`
- 当前成果：18 份 `_已清洗.docx` 及对应 `_审核状态.json` 已生成并通过门禁检查

接手未来高二历史时，答案教师版前半部分的题目必须忽略，只读取后部答案区；题目录入需保留 `【学习单】` / `【作业单】` 章节边界，并过滤 `B提升练` 等分层标题。

未来高二物理已经完成项目级归档：

- 总归档：`../墨痕教育-记忆库/模块问题/跨模块项目/2026-04-26-未来高二物理开发总归档.md`
- 模板经验：`格式处理/模板开发经验记录/模板F开发经验记录.md`
- 关键测试：`test_future_physics_workflow.py`

接手未来高二物理时，不要从零分析。先读总归档，再看项目内 `问题归档/` 对应证据。

## 开发规则

- 先读 `AGENTS.md`，它是 AI 接手的主规则文件。
- 原题文档属于不可变输入，禁止覆盖或写回；题目结构分析和动作计划必须放在原题之外，只有答案文档可以生成清洗版。
- 每次修复后必须测试，再报告结果。
- 每次问题修复或规则调整后必须写入 `问题归档/`。
- 长期项目记忆写入 `../墨痕教育-记忆库/`，不要塞进用户题目源目录。
- 不要修改 `wps_helper.py`，除非用户明确要求处理 WPS 连接层。
- 如果接手的是未来高二物理，不要先假设带 `(1)(2)(3)` 的解答题答案一定要拆成多个 `F4`；先确认插件实际是单答案框还是多答案框。
- 如果发现题目录入题数异常暴涨，优先排查表格里的小数是否被误判成题号。

## 常用验证

```powershell
.\.venv\Scripts\python.exe -m pytest -q -p no:cacheprovider .\test_document_preflight.py .\test_document_families.py
python -m pytest -q .\test_future_physics_workflow.py -k "AnswerInput"
python -m pytest -q .\test_future_history_workflow.py
python -m pytest -q .\test_zhongmei_chinese_workflow.py
python -m pytest -q .\test_zhongmei_chinese_answer_input.py
python -m pytest -q .\test_subject_detection.py
python -m pytest -q .\test_shared_core_flow.py
```

更多历史问题和修复证据见 `问题归档/INDEX.md`。
