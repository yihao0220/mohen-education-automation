# Windows 更新与未来高二生物验证手册

> 适用日期：2026-07-18 起
> Windows 独立代码目录：`E:\CODEX.projection\墨痕教育`
> GitHub：`https://github.com/yihao0220/mohen-education-automation`

## 先记住三条

1. 只在 `E:\CODEX.projection\墨痕教育` 更新和测试代码。
2. `E:\PYTHON\practice` 是旧的上级 Git 仓库，不得在里面运行项目级 `reset`、`clean` 或 `stash pop`。
3. 每条 PowerShell 命令单独粘贴、单独回车；只要实际 Git 顶层目录不是独立代码目录，立即停止。

## 下次更新只看这一节

打开 PowerShell，依次逐条运行。

### 1. 进入独立代码目录

```powershell
Set-Location "E:\CODEX.projection\墨痕教育"
```

### 2. 核对 Git 顶层目录

```powershell
git rev-parse --show-toplevel
```

必须显示：

```text
E:/CODEX.projection/墨痕教育
```

如果显示 `E:/PYTHON/practice` 或其他目录，不要继续。

### 3. 核对本机是否有未提交改动

```powershell
git status --short
```

正常情况：没有任何输出。只要出现内容，不要强制覆盖，先保留现场。

### 4. 下载 GitHub 最新代码

```powershell
git pull --ff-only origin main
```

### 5. 查看当前版本

```powershell
git rev-parse --short HEAD
```

版本号会随 GitHub 更新而变化，不要求永远是 `18b9b05`。只需确认 `git pull` 没有报错。

### 6. 更新 Windows 依赖

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

```powershell
.\scripts\setup_windows.ps1 -Development
```

预期结尾：

```text
Windows 开发与生产环境已就绪
```

### 7. 运行 Windows 安全自检

```powershell
.\scripts\verify_windows.ps1 -RunTests
```

该命令只检查依赖、WPS COM 注册和离线测试，不启动 WPS、不打开原题、不按 F1/F2/F3/F4。

预期结尾：

```text
Windows 环境自检通过；未启动 WPS、未打开文档、未执行任何按键。
```

## 第一次安装到一台新 Windows 电脑

先检查目标目录是否已经存在：

```powershell
Test-Path "E:\CODEX.projection\墨痕教育"
```

只有返回 `False` 时才继续：

```powershell
New-Item -ItemType Directory -Force "E:\CODEX.projection"
```

```powershell
git clone "https://github.com/yihao0220/mohen-education-automation.git" "E:\CODEX.projection\墨痕教育"
```

克隆完成后，回到“下次更新只看这一节”，从第 1 步继续执行。

如果 `Test-Path` 返回 `True`，不要再次克隆或覆盖，直接使用日常更新流程。

## 未来高二生物离线测试

### 1. 指向真实业务目录

把下面路径替换成 Windows 上“未来-高二-生物”的真实目录：

```powershell
$env:MOHEN_FUTURE_BIOLOGY_DIR = "D:\墨痕教育题目\未来-高二-生物"
```

核对路径：

```powershell
Test-Path $env:MOHEN_FUTURE_BIOLOGY_DIR
```

必须返回 `True`。

### 2. 运行题目、答案截取与答案清洗专项测试

```powershell
.\.venv\Scripts\python.exe -m pytest -q -p no:cacheprovider .\test_future_biology_question_input.py .\test_future_biology_answer_trim.py .\test_future_biology_answer_clean.py .\test_future_biology_answer_package.py .\test_review_gate_portability.py
```

路径和批次正确时，预期为：

```text
19 passed
```

如果显示 `15 passed, 4 skipped`，说明真实业务目录没有找到或路径填错；这不算完整通过。

专项测试会核对：

- 52 份正式题目、764 个 F1 题块；
- 结构标题污染为 0；
- 已确认装饰图片污染为 0；
- 73 份答案中命中“课时对点练”边界的 42 份；
- 42 份清洗答案共 580 题，每题都有题号、答案和解析；
- 题干、选项、题组标题、题干表格和装饰图均未进入清洗答案；
- 同段连写的明确 `(1)(2)(3)(4)` 小问会拆成 `F4 × 4 → F3 × 1`；
- 审核状态和 Windows 压缩包的跨机签名、UTF-8 路径保持有效；
- 原题和原答案只读，测试前后哈希不变。

## 未来高二生物 WPS 实机验证

离线测试通过后，再进行真实 WPS 验证。此步骤会实际触发墨痕插件的 `F1`，不能当作只读测试。

### 1. 先准备 WPS

1. 打开 WPS Writer。
2. 加载“墨痕题库工具”侧边栏。
3. 先打开一份代表性原题，不要打开答案文件。

优先测试：

- `选必一活页\第1章　作业1　细胞生活的环境.docx`
- `选必二活页\第1章　作业5　重点突破练(一).docx`
- `选必一活页\章末检测试卷(一).docx`

### 2. 启动题目录入

```powershell
.\.venv\Scripts\python.exe .\墨痕快刀\main.py
```

### 3. 人工验收清单

- 每个顶层阿拉伯题号只按一次 F1；
- 主观题 `(1)(2)(3)` 保持在同一道 F1 内；
- 题干、选项、题图、原生表格、公式完整；
- 试卷标题、分值说明、题组标题、“选择题/非选择题”标题不进入 F1；
- 章节横幅、“对点训练”、“综合强化”图片不进入 F1；
- 第 8 题保留自己的题图，但不带入后面的“综合强化”；
- 第 9 题从自己的顶层题号开始。

先验证一份，确认墨痕题库中的实际结果正确，再扩大到其他文档。

## 答案处理：先截取，再清洗

### 1. 按“课时对点练”截取

先只读预检：

```powershell
.\.venv\Scripts\python.exe .\tools\trim_future_biology_answers.py $env:MOHEN_FUTURE_BIOLOGY_DIR --dry-run --expected-source-count 73 --expected-matched-count 42
```

确认预检通过后才生成派生答案：

```powershell
.\.venv\Scripts\python.exe .\tools\trim_future_biology_answers.py $env:MOHEN_FUTURE_BIOLOGY_DIR --expected-source-count 73 --expected-matched-count 42
```

输出位于业务目录下的 `答案\按课时截取\`，不会覆盖原答案。遇到标记重复、标记位于表格、标记后缺答案等异常时，整批停止。

`按课时截取` 只是中间文档，里面仍有题干和选项，不能直接录答案。

### 2. 清洗为可录入答案

先只读预检：

```powershell
.\.venv\Scripts\python.exe .\tools\clean_future_biology_answers.py $env:MOHEN_FUTURE_BIOLOGY_DIR --preflight-only --expected-source-count 42 --expected-question-count 580
```

必须显示：

```text
预检文件: 42
预检题数: 580
需补空解析: 46
排除结构标题: 50
排除装饰媒体: 41
保留答案/解析媒体: 5
```

确认后再生成：

```powershell
.\.venv\Scripts\python.exe .\tools\clean_future_biology_answers.py $env:MOHEN_FUTURE_BIOLOGY_DIR --expected-source-count 42 --expected-question-count 580
```

预期结尾：

```text
清洗文件: 42
答案总数: 580
自动检查通过: 42/42
```

最终答案位于 `答案\已清洗\`，审核清单位于 `答案\审核清单\`，审核状态 JSON 与对应 `_已清洗.docx` 同目录。每题标准结构是：

```text
19．
答案：(1)…… (2)……
解析：……
```

如果输出已存在，脚本会停止，不会默认覆盖。确定要重建时才在生成命令末尾加 `--overwrite`。

### 3. 接收 Mac 生成的 Windows 兼容答案包

Mac 发来的文件必须叫：

```text
future_biology_cleaned_answers_windows.zip
```

不要使用 Mac Finder 直接生成的 `答案.zip`。兼容包内只有 `答案\已清洗\` 和 `答案\审核清单\`，中文路径已使用 UTF-8。

1. 先在独立代码仓库执行本手册开头的 `git pull --ff-only origin main`，确保 Windows 使用内容 SHA256 版审核门禁。
2. 把旧解压目录改名为备份，不要把新旧状态 JSON 混在一起。
3. 将新 ZIP 直接解压到 `D:\墨痕教育题目\未来-高二-生物\`；因压缩包自带 `答案\` 根目录，解压后会得到 `D:\墨痕教育题目\未来-高二-生物\答案\已清洗\`。
4. 路径和 ZIP 解压时间变化不会使门禁失效；若 DOCX 内容被手改，仍会正常变为 `stale`。

### 4. Windows WPS 抽样验证答案录入

1. 在 WPS 打开 `答案\已清洗\` 中一份 `_已清洗.docx`，加载墨痕题库侧边栏。
2. 运行 `.\.venv\Scripts\python.exe .\答案录入\answer_input.py`。
3. 先只验证一份，确认题号不进入 F2/F3/F4，题干和选项不存在，答案进 F2 或 F4，整题解析只进一次 F3。
4. 重点核对带 `(1)(2)(3)` 的主观题；只按清洗答案中实际存在的小问项录 F4，不因题干出现所有小问编号就自动补项。
5. 定向验证“第1章　第1节　细胞生活的环境”第 15 题：同一段中的 `(1)(2)(3)(4)` 必须分别选中，日志应显示 `F4 × 4 → F3 × 1`；任何一次 F4 都不得包含下一小问的编号或内容。

这一步会真实按 F2/F3/F4，必须等离线专项测试和清洗门禁通过后才执行。

## 出现异常时怎么做

### `git status --short` 出现文件

停止更新。不要运行：

```text
git reset --hard
git clean -fd
git stash pop
```

先记录当前目录和状态，避免覆盖本机改动。

### `origin does not appear to be a git repository`

先核对：

```powershell
git rev-parse --show-toplevel
```

只有输出严格等于 `E:/CODEX.projection/墨痕教育` 时，才继续检查：

```powershell
git remote -v
```

独立克隆正常应自带 `origin`。不要给 `E:\PYTHON\practice` 上级仓库添加本项目的 `origin`。

### 出现 `Unlink of file ... nul failed. Should I try again? (y/n)`

输入：

```text
n
```

然后停止。这个提示通常意味着正在错误的上级仓库执行覆盖或清理。

### 出现 `../课表网站/` 等无关目录

立即停止。这说明当前 Git 顶层目录过大，已经把其他项目纳入状态检查。

### PowerShell 提示多行粘贴

本手册已把关键命令拆开。每次只复制一个代码块，不要一次粘贴整章。

## 当前交接状态

截至 2026-07-18：

- GitHub 功能代码已克隆到独立 Windows 目录；首次安装时版本为 `18b9b05`，同段小问修复的功能基线为 `a9ad251`；
- Windows 开发与生产依赖已安装，用户确认安装脚本显示“Windows 开发与生产环境已就绪”；
- GitHub Windows runner 已通过 PR #2 的干净克隆、开发安装和离线回归；Mac 当前生物专项回归为 `19 passed`；
- 上述证据不等于生产 WPS 验证。Windows 真实业务目录专项结果仍应按本手册保存；
- 真实 WPS F1/F2/F3/F4 录入和 52/52 P1b 人工门禁仍未完成，不能仅凭安装成功或 GitHub CI 宣称生产验证通过。
