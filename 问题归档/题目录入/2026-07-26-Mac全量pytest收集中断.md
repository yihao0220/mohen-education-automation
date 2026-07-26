# Mac 全量 pytest 收集中断与平台依赖替身

## L1 摘要
- 模块：题目录入（跨三模块的测试基建）
- 日期：2026-07-26
- 状态：已完成并复核；Windows 全量回归待执行
- 结论：裸跑 `pytest` 此前在任意平台都会崩溃、0 个用例运行且看不到原因。修复后 macOS 上 192 passed；平台依赖替身从单个测试文件的副作用提升为全局 `conftest.py`，消除了"测试能否通过取决于收集顺序"的隐式依赖。

## L2 现象与根因

两层独立缺陷叠加，互相掩盖。

**第一层：pytest 进程崩溃（平台无关）**

- 现象：项目根执行 `pytest`，输出 `ValueError: I/O operation on closed file.`，`0 tests ran`，且看不到任何一条收集错误的原因。加 `--continue-on-collection-errors` 也无效。
- 根因：`答案录入/test_parse_logic.py` 与 `答案录入/test_subq_split.py` 在**模块级**执行
  `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")`。
  pytest 捕获模式下 `sys.stdout` 是 pytest 包装的临时文件流；新建 `TextIOWrapper` 接管底层
  `buffer` 后，原 wrapper 失去引用被 GC，**GC 时关闭了底层 tmpfile**。pytest 随后
  `tmpfile.seek(0)` 读取捕获内容即抛 `ValueError`，整个进程终止。
- 这与 macOS / Windows 无关。项目长期使用「手工列出文件」的方式跑测试，恰好绕开了触发链，因此从未暴露。

**第二层：Windows 专属依赖缺失（macOS）**

- `pyautogui`（模拟按键）与 `win32com`（WPS COM）只在 `requirements-windows.txt`，macOS 按设计不装。
- 全项目仅 2 处 `import pyautogui`（`墨痕快刀/core_parser.py:6`、`答案录入/answer_input.py:6`），
  4 处调用且全部是 `pyautogui.press(...)`。
- `test_zhongmei_physics_workflow.py:24` 已用 `sys.modules.setdefault` 注入轻量替身，但那是**单个测试文件的副作用**：
  其他测试能否通过，取决于 pytest 是否恰好先收集到该文件。单独执行
  `pytest test_zhongmei_chinese_answer_input.py` 会 22 errors，全量执行却只有 1 failed。

**附带查明：11 个文件名以 `test_` 开头但不是测试**

`墨痕快刀/` 下 8 个、`格式处理/test_template_d.py`、`答案录入/` 下 2 个，均为 0 个 test 函数、0 个 TestCase
的早年脚本式调试工具，设计上由人工 `python <文件>` 直接运行看输出，模块级即执行真实逻辑
（`test_word_com.py` 模块级直接连 WPS）。它们对 pytest 贡献 0 个用例，却制造收集错误。

## L3 改动与验证

- 改动文件：
  - `答案录入/test_parse_logic.py`、`答案录入/test_subq_split.py` — 模块级 stdout 劫持改为就地
    `sys.stdout.reconfigure(encoding="utf-8", errors="strict")`，不替换对象，因此不触发 GC 关闭。
  - `conftest.py`（新增）— 全局注入 `pyautogui` / `wps_helper` 替身；`collect_ignore` 排除 11 个脚本式调试工具。
  - `README.md` — 常用验证改为可直接跑全量，并说明既有失败性质。
  - 未改动任何生产代码：`core_parser.py`、`answer_input.py`、`wps_helper.py` 均未触碰。

- 测试命令：
  ```bash
  .venv/bin/python -m pytest -q
  ```

- 结果对比（macOS）：

  | | 裸跑 pytest | 加 `--continue-on-collection-errors` |
  |---|---|---|
  | 修复前 | 崩溃，0 passed | 165 passed / 7 failed / 16 errors |
  | 修复后 | **192 passed** / 7 failed / 55 skipped / 10 errors | 同左 |

- 恢复的 27 个用例来自 `test_future_physics_workflow.py`(13)、`test_review_prompt_mode.py`(9)、
  `test_subject_detection.py`(4)、`test_math_subject_detection.py`(1)。
- 两个被改的脚本，手工直接运行的输出与修改前**逐字节一致**（`diff` 验证），管道重定向下中文仍正常。
- 换行符未被破坏：两文件仍为全 CRLF 一致。

## L4 深层备注

- **证明了什么**：macOS 上全量回归可一条命令跑通；27 个此前无法执行的用例恢复；替身不改变任何生产执行路径。
- **没有证明什么**：Windows 全量 `pytest` 尚未执行；替身在 Windows 下是否与真实 `pyautogui` 存在行为差异未实测；
  本次未验证任何 WPS 生产录入行为。
- 替身不会掩盖 Windows 依赖缺失：依赖校验由 `scripts/verify_windows.ps1:15` 用独立子进程
  `import ... pyautogui, win32com.client` 完成，该子进程不加载 `conftest.py`。
- 生产运行（`python main.py` 等）不加载 `conftest.py`，替身仅在 pytest 进程内生效。
- 剩余 7 failed + 10 errors 属**既有第三层问题**，与本次改动无关（改动前后失败项逐项一致）：
  测试硬编码了 `D:\墨痕教育题目\...`、`E:\samples\...` 业务文档路径，真实样本按约定不入库；
  且 Windows 路径语义在 macOS 上不成立（`'E:\samples\x.docx'` 不被当作路径分隔）。未处理。
- 下次排查入口：根 `conftest.py`；如新增测试需要真实业务文档，应加存在性跳过而不是让其 error。
- 关联：`../协作规范/2026-07-22-方案必须让用户理解后再实施.md`
