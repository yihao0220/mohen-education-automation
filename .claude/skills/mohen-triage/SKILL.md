---
name: mohen-triage
description: 墨痕教育录入系统的 Bug 定位速查。当用户报告题目录入、答案清洗、答案录入或预检环节出问题时使用，例如"答案录入乱位"、"大题小问识别错"、"格式清洗不干净"、"审核清单误拦截"、"题目录入范围错"、"科目识别错误"、"题数暴涨"、"WPS 连不上"、"预检画像异常"、"文档族分组异常"、"P1b 视觉门禁异常"。用于把用户描述的症状映射到具体文件和函数，缩短排查路径。不用于新功能开发（那走 mohen-project-rules）。
---

# 墨痕教育 Bug 定位

先按症状查表定位文件，再读代码确认；表只给起点，不是结论。

## 症状 → 排查入口

| 症状 | 排查文件 | 先看什么 |
|------|----------|----------|
| 答案录入乱位 | `答案录入/answer_input.py`、`shared_core/answer_core.py` | 先核对 `ans_start_p` / `ana_start_p` 与原文段号是否一致；确认是坐标错位还是按键落位过快，只有后者才调 `WAIT_TIME` |
| 大题小问识别错 | `答案录入/answer_input.py`、`shared_core/answer_core.py` | `find_subquestion_matches()`、`_extract_sub_answer_items()`、`should_split_subquestion_answers()` |
| 格式清洗不干净 | `格式处理/main.py`、`格式处理/格式模板库/*.py` | 对应模板的正则与清洗逻辑；确认模板是否真的命中 |
| 审核清单误拦截 | `shared_core/review.py`、`shared_core/answer_core.py` | 审核规则与题答映射规则 |
| 题目录入范围错 | `墨痕快刀/core_parser.py`、`shared_core/question_core.py` | 对应科目的范围与吞并规则 |
| 科目识别错误 | `墨痕快刀/config.py`、`墨痕快刀/core_parser.py` | `detect_subject()` 与 overlay 逻辑；文件名不含学科时是否走了内容嗅探 |
| 录入题数异常暴涨 | `墨痕快刀/core_parser.py` | **优先排查表格里的小数/百分数是否被误判成题号起点**；对比日志里"发现题号数量"与"准备录入题数" |
| 预检画像 / 动作计划异常 | `shared_core/document_preflight.py`、`test_document_preflight.py` | 源文件 SHA256、原生节点坐标、Docling 对照、`execution_enabled` |
| 文档族分组 / 异常候选异常 | `shared_core/document_families.py`、`test_document_families.py` | Profile schema、角色分布、分项相似度、完全链接阈值、三个安全开关 |
| P1b 页面 / 视觉门禁异常 | `shared_core/document_render.py`、`document_visual_review.py`、`document_family_calibration.py` | 渲染 provider、`page_truth_authority`、源哈希、未知区域、三个安全开关 |
| WPS 未连接 | — | 确认 WPS 已打开文档且「墨痕题库工具」侧边栏已加载；再查是否有残留 WPS 进程 |

## 用户报 bug 时的默认排查顺序

1. `答案录入/answer_input.py` — 当前录入行为问题
2. `格式处理/main.py` + `格式处理/格式模板库/` — 格式清洗与审核问题
3. `墨痕快刀/core_parser.py` + `shared_core/` — 题目录入与共享规则问题

## 排查纪律

- **先看是不是已知问题。** `问题归档/INDEX.md` → 对应模块 `INDEX.md` → 单条记录。同一现象反复出现时，历史记录里通常已有根因。
- **两层验证。** 离线测试覆盖核心逻辑 + 真实样本文档定向复核。优先补失败用例再改代码。
- **测试后清理。** 临时 docx、临时审核清单、临时输出目录必须删掉，避免污染下一轮判断。
- **同一类问题第三次出现时停下来。** 抽象成共享规则或 `subject_overlay`，不要继续补局部特判。
- 修复完成后按双轨制归档，走 skill `mohen-archive-check`。

## 用测试文件当规格说明

排查具体项目行为时，对应的回归测试比文字描述更准确：

- `test_zhongmei_chinese_workflow.py` / `test_zhongmei_chinese_answer_input.py` — 众美高三语文
- `test_future_physics_workflow.py` — 未来高二物理
- `test_future_history_workflow.py` — 未来高二历史
- `test_future_biology_*.py` — 未来高二生物题目/截取/清洗/打包
- `test_subject_detection.py` / `test_math_subject_detection.py` — 科目识别
- `test_shared_core_flow.py` — 共享内核主链路
- `test_document_preflight.py` / `test_document_families.py` — P0 / P1a
- `test_document_render.py` / `test_document_visual_review.py` / `test_document_family_calibration.py` / `test_p1b_*.py` — P1b
