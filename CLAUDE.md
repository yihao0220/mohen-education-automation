# 墨痕教育 — 项目规则

WPS「墨痕题库」插件的自动化录入工具链。三步流程：题目录入 → 答案格式清洗 → 答案/解析录入。

本文件只放**读代码看不出来的事实**。可查阅的知识在下方「按需加载」表里，不要一次性全读。

## 插件契约（由 WPS 插件侧定义，代码里推不出来）

| 按键 | 用途 |
|------|------|
| `F1` | 题目录入 |
| `F2` | 普通题整题答案 |
| `F3` | 解析录入 |
| `F4` | 大题小问答案 |

生产录入只能在 Windows + WPS COM 下执行。macOS 只做开发、离线测试和 Quick Look 只读预览；Quick Look 结果固定 `page_truth_authority=false`，不是 WPS 分页真值。

## shared_core/ 模块职责（文件名区分不出来）

- `document_preflight.py` — DOCX 只读画像、Docling 对照、F1 预演计划
- `document_roles.py` — 段落角色、置信度与可解释证据
- `document_families.py` — 文档族分组、代表样本与异常候选
- `document_render.py` — 跨平台 PDF 页面渲染；Windows WPS 为生产真值
- `macos_quicklook_render.py` — macOS 开发预览
- `document_visual_review.py` — 页面视觉角色与人工审核门禁
- `document_family_calibration.py` — P1b 阈值校准与整批门禁
- `p1b_batch.py` — Windows WPS 整批只读副本与哈希清单
- `cli_output.py` — 跨平台 CLI UTF-8 输出

## 协作方式（最高优先级）

用户可以不写代码，但必须理解产品怎么运转、问题在哪一层、方案为什么这样选、如何验收。AI 负责执行，用户保留判断权和最终验收权。

涉及架构、核心流程、数据或文档流转、跨模块修改和 Bug 修复时，先用非程序员能理解的中文依次说明：当前功能链 → 问题层级与证据 → 拟议改动与替代方案 → 影响范围与风险 → 验收方法。用户理解并确认后再实施核心修改。

不用术语、函数名或测试数字代替解释。报告结果时必须说明证据**证明了什么、没有证明什么**。用户说「没听懂」「不知道改了哪里」或无法自行验收时，回到五段式重讲，不堆技术细节。

证据：`问题归档/协作规范/2026-07-22-方案必须让用户理解后再实施.md`

## 铁律

1. **原题文档不可变。** 禁止覆盖、保存或写回原题，禁止改动其排版、字号、页边距、表格、图片、标题和段落顺序。原题只允许被只读扫描、建结构画像、生成外部题块与动作计划，以及在 WPS 中执行选区和按键。只有答案文档可以生成 `_已清洗.docx` 等派生产物。
2. **不要改 `wps_helper.py`。** 底层 WPS 连接层。`墨痕快刀/` 和 `答案录入/` 各有一份独立副本（历史遗留），保持各自独立，不要合并。
3. **新模板必须注册进 `格式处理/main.py` 的 `TEMPLATES` 列表。** 只更新 `__init__.py` 不够 —— `main.py` 不依赖 `__all__`。
4. **答案录入前必须有审核状态文件。** 签名固定为 `size + SHA256`，不得把绝对路径或修改时间作为相等条件。用户手改过 `_已清洗.docx` 即视为状态失效，必须重新过门禁；缺 SHA256 的旧状态不得静默放行。
5. **预检产物以 JSON 为机器真值。** `DocumentProfile.json` / `ActionPlan.json` 是执行契约，Markdown 只能由 JSON 派生供人工审核，不得反向作为执行输入。各阶段安全开关见 skill `mohen-preflight`。

## 学科逻辑架构

三大核心 `英语 / 理科 / 文科` + `subject_overlay` 覆盖层。学科差异优先放覆盖层，不要拆成一套套独立引擎，也不要散落进 `core_parser.py` 的临时 `if/else`。同一类问题修到第三次时，抽象成共享规则或覆盖层，不要继续补局部特判。

## 按需加载

项目知识全部在本仓库内，随 git 同步到两台机器。下表给的是**文件路径**，直接 Read 即可，不依赖 skill 是否被注册。
在项目目录内启动 Claude Code 时，`.claude/skills/` 下的三项也会自动注册为可调用 skill。

| 场景 | 入口 |
|------|------|
| 用户报 bug，要定位改哪个文件 | `.claude/skills/mohen-triage/SKILL.md` |
| 接手或继续某学校某学科项目 | `.claude/skills/mohen-project-rules/SKILL.md`（内含路由表，只读需要的那一个 reference） |
| P0/P1a/P1b 只读预检与安全开关 | `.claude/skills/mohen-preflight/SKILL.md` |
| 用户说「检查归档」 | `.claude/skills/mohen-archive-check/SKILL.md` |
| 模板开发经验 | `格式处理/模板开发经验记录/README.md` |
| 架构问题主口径 | `docs/墨痕教育架构问题工程思维分析拆解.md` |
| 文件分类约定（stable/sample/generated/local） | `工作区约定.md` |

**一次只读需要的那一个。** 把上表全部读进来，就回到了重构前的状态。

## 记忆与归档双轨

- 项目内代码证据：`问题归档/INDEX.md` — 现象、根因、改动点、测试样例、测试命令、结论
- 外部长期记忆：`../墨痕教育-记忆库/INDEX.md` — 可复用摘要、全局规则、工作偏好
- 手动触发口令：`检查归档`

## 路径

- Mac 开发：`/Users/xiaosheng/工作/全局工作区/40_事业与收入/墨痕教育`
- Windows 生产：`E:\CODEX.projection\墨痕教育`（独立仓库）。执行 `pull/reset/clean/stash` 前必须确认 `git rev-parse --show-toplevel` 严格等于该目录；旧目录 `E:\PYTHON\practice\墨痕教育` 属于上级仓库，不得在其中执行本项目的覆盖命令。日常更新用 `git pull --ff-only origin main`，`git status --short` 非空时停止并保留现场。
- 原题、答案、P1b 产物、页面截图、JSON 批次产物一律在仓库之外；两台机器各自通过 `工作台路径配置.json` 指向本机业务目录。
- `.qoder/repowiki/` 是本地生成快照，不是事实来源。与本文件冲突时以本文件为准。
