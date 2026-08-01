# Serena 与 Ponytail 首轮低风险代码瘦身

## 现象

- 已停用的 `格式处理/format_answers_deepseek.py` 在无条件 `return False` 后仍保留约 330 行不可达实现。
- 6 个格式模板重复实现垃圾行判断、模板匹配评分和统一字体设置。
- `core_parser.py` 与 `answer_input.py` 在模块导入阶段强制加载 `pyautogui`、`wps_helper/win32com`，导致 Mac 纯解析测试无法启动。
- P1b 三个模块各自实现同一个“路径是否位于目录内”函数；格式处理与答案录入另有重复 UTF-8 输出逻辑。
- 离线基线包含四类噪音：地理测试仍期待旧题号格式；缺少 Windows 真实样本时，南城数学测试直接失败而非跳过；Mac 测试用 `Path` 构造 Windows 路径，导致文件名断言失真；预检黄金文件写死本机 Docling 安装状态和版本号。

## 根因

- 旧实验链路停用时只加了提前返回，没有删除其后实现。
- 早期模板以复制方式扩展，公共接口稳定后未回收完全相同的辅助代码。
- Windows 执行依赖与纯解析逻辑位于同一模块，但依赖仍放在顶层导入。
- P1b 开发阶段分别落地模块，形成可由 `Path.is_relative_to()` 直接替代的小型重复函数。
- 测试基线没有同步复合小题号规则，也没有统一真实样本缺失时的跳过策略。

## 改动点

- 把 DeepSeek 旧文件缩成兼容壳，保留 `process_document_llm()`、`run_llm_engine()` 和停用提示。
- 在 `格式处理/格式模板库/__init__.py` 集中三个已稳定的模板辅助函数；6 个模板保留原公开函数名，以一行包装调用共享实现。
- 将 `pyautogui`、`get_active_wps` 移入真正的 WPS 执行函数；未修改 `wps_helper.py`。
- 格式处理和答案录入统一调用 `shared_core.cli_output.configure_utf8_stdio()`。
- P1b 路径门禁直接使用 `Path.is_relative_to()`，门禁文案和阻断条件不变。
- 众美文言文路径映射使用 `PureWindowsPath` 解析 Windows 路径，使 Mac 离线测试可验证同一映射规则。
- 地理测试改为期待已生效的复合题号 `13．（1）`；南城真实样本测试在样本缺失时显式跳过。
- 跨平台清单测试改用 `PureWindowsPath` 表达 Windows 路径，不修改生产清单逻辑。
- 预检黄金文件把 Docling 可用性和版本归一化成占位符，继续严格比较稳定画像和动作计划字段。
- 新增 `test_low_risk_cleanup.py`，覆盖停用兼容入口、共享模板匹配/垃圾过滤和字体格式。

## 测试样例

- DeepSeek 兼容入口仍返回 `False`，并引导使用 `格式处理/main.py`。
- 6 个模板继续过滤空行；模板 A 对一条匹配、一条不匹配文本的得分仍为 `0.5`。
- Mac 不安装 `pyautogui`、`win32com` 时，纯解析、学科识别和答案映射测试可直接导入运行。
- 按键测试通过 `sys.modules["pyautogui"]` 注入替身，继续覆盖懒加载后的 F4/F3 顺序。
- `D:\\墨痕教育题目\\众美-高三-语文\\文言文\\...` 仍映射到同项目 `答案\\文言文答案\\..._已清洗.docx`。
- Document Render、P1b Batch、Quick Look 三个入口仍拒绝把派生产物写入原题目录或批次目录。

## 测试命令

```bash
PYTHON=/Users/xiaosheng/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3

$PYTHON -m unittest test_low_risk_cleanup
$PYTHON -m unittest test_subject_detection test_math_subject_detection test_review_prompt_mode
$PYTHON -m unittest test_geography_shared_core test_compound_answer_ids test_duplicate_answer_question_ids test_review_duplicate_answer_question_ids
$PYTHON -m unittest test_future_physics_workflow test_zhongmei_chinese_answer_input
$PYTHON -m unittest test_nancheng_math_workflow
$PYTHON -m compileall -q main.py shared_core tools 格式处理 答案录入 墨痕快刀
git diff --check
```

另将 20 个不依赖 pytest 的 unittest 模块逐个独立运行，避免测试模块之间的 `main` 名称污染；共收集 132 项，其中 91 项通过、41 项按真实样本缺失规则跳过、0 项失败。

pytest 与 `pytest-regressions` 仅临时安装到系统临时目录后运行，目录已清理；11 个 pytest 模块共 62 项通过、6 项按平台条件跳过、0 项失败。

## 结论

- 核心实现（不含测试与归档）新增 93 行、删除 796 行，净减少 703 行；未新增第三方依赖。
- 离线语法、模板接口、解析逻辑、跨平台路径映射和 P1b 输出目录门禁通过。
- unittest 与 pytest 专属套件均已执行；pytest 依赖只用于临时验证，没有写入项目依赖。
- 未发送 F1/F2/F3/F4，未连接 Windows WPS；本次结果不构成 Windows WPS 生产实机证明，也不改变 P1b 的 52/52 人工审核门禁。

## 下次排查入口

1. 模板匹配或字体表现异常：先查 `格式处理/格式模板库/__init__.py` 的三个共享辅助函数，再跑 `test_low_risk_cleanup.py` 和对应模板测试。
2. Mac 纯解析导入再次出现 Windows 依赖错误：检查 `core_parser.py::process_chapter` 与 `answer_input.py::execute_input/main`，不要修改 `wps_helper.py`。
3. P1b 输出目录保护异常：检查 `document_render.py`、`p1b_batch.py`、`macos_quicklook_render.py` 的 `Path.is_relative_to()` 调用和原题哈希测试。
4. Windows 生产验收：仍按现有 WPS 手册执行；不要把本次 Mac 离线结果升级为生产页面或按键证据。
