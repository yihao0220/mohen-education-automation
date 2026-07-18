# 模板C开发经验记录

> 适用文档类型：道法（思想品德）试卷，格式特征为 `【答案】` + `【解析】`

**开发时间**：2026-03-05  
**测试文档**：`1.1 青春的邀约.docx`、`1.3.docx`  
**最终状态**：✅ 已完成并修复

---

## 一、文档格式特征

```
1．题干（选择题）
A．选项  B．选项  C．选项  D．选项
【答案】A
【解析】本题考查XXX的相关知识。
①：逐项解释...；
②：...；
故本题选A。

9．大题题干（材料分析题）
(1)第一问
(2)第二问
【答案】(1)第一问答案
(2)第二问答案
【分析】考点考查：...（元数据，跳过）
      能力考查：...
【解析】（1）详细分析...
       （2）...
```

---

## 二、核心设计决策

### 1. 题目存储用 list，不用 dict
模板A/B 使用 `dict` 存题目（题号→内容），会在重复题号时覆盖。
此文档 Q7 出现两次（内容不同），Q6 乱序，改用 `list` 按扫描顺序存储，
输出时用独立的 `output_idx` 计数确保题号连续。

### 2. GARBAGE_PATTERNS 必须使用全宽句号
子标题 `1．3 学会自我保护` 中的 `.` 是 U+FF0E（全宽句号），
GARBAGE_PATTERNS 里必须用 `[．.]`（全宽+半角），只用 ASCII `.` 无法匹配。

正确写法：
```python
r"^\d+[．.]\d+[\s　]"   # ✅ 包含全宽 ．
r"^\d+[.]\d+[\s　]"     # ❌ 仅 ASCII，漏过全宽句号
```

### 3. QUESTION_PATTERN 宽松化
初版模式限制了首字符集合（汉字/字母/数字/括号），导致：
- `2．"你脸上的青春痘..."` 中的 `"` (U+201C 全角双引号) 无法匹配 → Q2 丢失
- `8． "五一"小长假` 开头有空格 → Q8 丢失

**修复**：改为 `r"^(\d+)[．.](.+)"` 任意内容，依赖 GARBAGE_PATTERNS 前置过滤。

### 4. 【分析】考点元数据跳过
大题中 `【分析】考点考查：...` 是教师侧元数据，不属于学生解析，
用内循环跳过整个 `【分析】` 块（至 `【解析】` 出现为止）。

---

## 三、问题记录

### 问题1：QUESTION_PATTERN 不兼容全角引号和空格前缀

| 字段 | 内容 |
|------|------|
| 症状 | Q2、Q8 未被识别，Q1 解析内容被 Q2 的解析覆盖 |
| 根因 | 初版正则 `^(\d+)[．.]([\u4e00-\u9fa5a-zA-Z0-9（(《「""¿].+)` 字符类不含 `"` (U+201C) 和空格 |
| 修复 | 改为 `^(\d+)[．.](.+)` 宽松匹配，依赖 GARBAGE_PATTERNS 过滤章节编号 |
| 验证 | 9 道题全部正确识别 |

---

### 问题2：模板注册遗漏导致清洗用错模板 ⭐

| 字段 | 内容 |
|------|------|
| 症状 | 解析只输出 1 行（`【解析】` 首行），`①②③④` 续行和 `故本题选X。` 全部丢失 |
| 根因 | **`格式处理/main.py` 的 `TEMPLATES` 列表没有注册 `template_c`**，只更新了 `__init__.py` 是不够的；main.py 不依赖 `__all__`，需要显式添加 |
| 现象 | 清洗时实际使用 template_b，template_b 对 `【解析】` 处理不同，续行显示为"未识别" |
| 修复 | 在 `main.py` 导入和 `TEMPLATES` 列表中加入 `template_c` |
| 教训 | **新模板开发完毕必须同步更新 `main.py`**（参见 AGENTS.md 规范第4条） |

---

### 问题3：AI 分析文档时未主动识别解析续行格式 ⭐

| 字段 | 内容 |
|------|------|
| 症状 | 用户发送待清洗文档供 AI 审阅，但 AI 没有主动指出 `①②③④` 须被收集 |
| 根因 | AI 在读文档内容时只验证了题号识别，未检查解析块的续行完整性 |
| 修复 | AGENTS.md 新增规范第5条：**收到 docx 样本必须主动检查解析续行模式** |
| 适用行格式 | `①②③④`、`A：B：C：D：`（逐项解释）、`故本题选X。`、`（1）第X步：` |
| 教训 | 用户发文档给 AI 看，目的就是让 AI 发现隐性格式风险，不只是验证已知逻辑 |

---

### 问题4：Word 自动编号 (1)(2)(3) 无法识别 ⭐

| 字段 | 内容 |
|------|------|
| 症状 | 大题小问 `(1)(2)(3)` 的答案丢失，第一小问缺少 `(1)` 前缀，后续小问答案混入前一小问 |
| 根因 | Word 的自动编号（列表格式）**不在段落文本中**，通过 `Paragraph.Range.Text` 获取不到 `(1)` 前缀，实际文本只有问题内容 |
| 发现 | 通过 `ListFormat.ListString` 可获取自动编号文本，如 `'(1)'`、`'(2)'`、`'（1）'`（含中文括号） |
| 修复 | 在读取段落文本时，拼接 `ListString` + `Text`：`text = list_string + text` |
| 代码 | ```python
list_format = p.Range.ListFormat
if list_format.ListType > 0:
    list_string = list_format.ListString
    if list_string:
        text = list_string + text
``` |
| 教训 | Word 文档中的编号可能是自动列表格式，不是纯文本，必须通过 COM 接口的 `ListFormat` 属性获取 |

---

### 问题5：BIG_ANSWER_PATTERN 与小问处理逻辑冲突

| 字段 | 内容 |
|------|------|
| 症状 | `(2)` 的问题内容被追加到 `(1)` 的答案中，`(2)` 的答案丢失 |
| 根因 | `BIG_ANSWER_PATTERN`（匹配 `【答案】`）的续行收集没有检测小问编号，导致 `(2)` 被当作续行 |
| 修复 | 1. 在 `BIG_ANSWER_PATTERN` 续行收集中添加 `SUB_QUESTION_PATTERN` 检测<br>2. 答案收集完成后 `break`，让外层循环处理下一个小问 |
| 代码 | ```python
if (re.match(r'^【分析】|^【解析】|^【答案】', next_text)
        or QUESTION_PATTERN.match(next_text)
        or SUB_QUESTION_PATTERN.match(next_text)):  # 新增
    break
``` |

---

## 问题6：【子议题】标记混入答案

| 字段 | 内容 |
|------|------|
| 症状 | `【子议题三：传承榜样之美】` 被追加到小问答案中 |
| 根因 | 小问答案的续行收集停止条件缺少对 `【子议题】` 的检测 |
| 修复 | 在续行收集停止条件中添加 `cont_text.startswith('【子议题')` |

---

### 问题7：小问问题行与答案行无法区分 ⭐

| 字段 | 内容 |
|------|------|
| 症状 | `(1)这个故事体现了中华民族的哪种传统美德？` 被判定为直接答案收集；`(2)孝悌忠信...` 等答案行反而被跳过 |
| 根因 | 原代码用"不以问号结尾"判断是否为答案，但 `(2)我们还要弘扬哪些中华传统美德？（写出两个即可）` 问号在中间而非结尾，导致误判 |
| 修复 | 新增 `QUESTION_WORDS_PATTERN`，检测内容是否含疑问词（什么/哪/怎样/如何/为什么/怎么/谁/哪里/多少/几/吗） |
| 规则 | 含疑问词 → 判定为问题，跳过并向后查找 `【答案】`；不含疑问词 → 判定为答案，直接收集 |
| 代码 | ```python
QUESTION_WORDS_PATTERN = re.compile(r'什么\|哪些\|怎样\|如何\|为什么\|怎么\|谁\|哪里\|多少\|几\|哪\|吗')
has_question_word = QUESTION_WORDS_PATTERN.search(question_content) is not None
if question_content and not has_question_word:
    # 判定为答案，直接收集
``` |

---

### 问题8：最后一个小问答案被跳过 ⭐

| 字段 | 内容 |
|------|------|
| 症状 | `(3)做自强不息的人；做敬业乐群的人；...` 被打印 `⏭️ 跳过小问问题`，答案丢失 |
| 根因 | peek 逻辑只检测下一行是否为"小问编号"或"新题目"。最后一个小问答案后接的是 `解析：`，不满足条件，`is_direct_answer` 保持 False，落入跳过分支 |
| 修复 | 扩展 peek 终止条件，以下情况均判定当前行为直接答案：<br>1. 下一行是新小问编号<br>2. 下一行是新题目<br>3. 下一行是解析标记（`解析：`/`【解析】`/`【分析】`/`【详解】`）<br>4. 已到文档末尾（`found_peek = False`） |
| 代码 | ```python
if (SUB_QUESTION_PATTERN.match(peek_text) or QUESTION_PATTERN.match(peek_text)
        or re.match(r'^[【解析】【分析】【详解】【讲解】]|^解析：|^详解：', peek_text)):
    is_direct_answer = True
if not found_peek:  # 文档末尾
    is_direct_answer = True
``` |

---

### 问题9：【答案】内含小问编号时被重复拼接前缀

| 字段 | 内容 |
|------|------|
| 症状 | 答案输出为 `（3）(1)诚信。`，多出外层小问编号 `（3）` |
| 根因 | 当处理 `（3）践行...` 小问时，代码向后找到 `【答案】(1)诚信。`，将外层编号 `（3）` 拼在内容前面，变成 `（3）(1)诚信。` |
| 格式说明 | 该文档大题格式为：三个问题 `（1）（2）（3）` 分别列出，然后统一跟一个 `【答案】(1)...(2)...(3)...` 块，答案内容本身已含小问编号 |
| 修复 | 处理 `【答案】` 时检测内容是否已以小问编号开头，如是则直接使用，不再拼接外层编号 |
| 代码 | ```python
answer_body = next_text[4:]  # 去掉【答案】
if SUB_QUESTION_PATTERN.match(answer_body):
    answer_content = answer_body       # 已含(1)，直接使用
else:
    answer_content = sub_q_match.group(0) + answer_body  # 无编号，拼接
``` |

---

## 四、template_c.py 核心实现备忘

### 1. Word 自动编号获取（关键！）

Word 的列表自动编号（如 `(1)(2)(3)`、`一、`）**不在段落文本中**，必须通过 `ListFormat` 获取：

```python
p = paras(i)
text = p.Range.Text.replace("\r", "").replace("\n", "").strip()

# 获取 Word 列表格式标记（自动编号）
list_format = p.Range.ListFormat
if list_format.ListType > 0:  # 有列表格式
    list_string = list_format.ListString
    if list_string:
        text = list_string + text  # 拼接编号和文本
```

常见 `ListString` 值：
- `'(1)'`、`'(2)'` — 英文括号自动编号
- `'（1）'`、`'（2）'` — 中文括号自动编号  
- `'一、'`、`'二、'` — 中文数字编号
- `'1.'`、`'2.'` — 阿拉伯数字编号

### 2. 内循环收集【解析】续行（选择题和大题通用）

```python
while i <= paras.Count:
    next_text = paras(i).Range.Text.strip()
    if not next_text: i += 1; continue
    # 停止条件：下一题题号 或 新的【答案】
    if QUESTION_PATTERN.match(next_text) or re.match(r'^【答案】', next_text):
        break
    analysis_lines.append(next_text)   # ①② 行、故本题选X 均会被收集
    i += 1
```

`①` (U+2460) 是 Unicode No 类别（数字-其他），**不是** Nd（十进制数字），
Python `\d` 不匹配 `①`，所以 `QUESTION_PATTERN` 不会误触发。

### 3. 大题小问答案收集（支持 (1)(2)(3) 多小问）

```python
# 检测小问编号
sub_q_match = SUB_QUESTION_PATTERN.match(text)
if sub_q_match and current_question:
    # 跳过问题，查找答案
    while i <= paras.Count:
        next_text = paras(i).Range.Text.strip()
        if SUB_QUESTION_PATTERN.match(next_text) or QUESTION_PATTERN.match(next_text):
            break  # 下一个小问或新题目
        if next_text.startswith('【答案】'):
            # 用小问编号替换【答案】前缀
            answer_content = sub_q_match.group(0) + next_text[4:]
            current_question['answer_lines'].append(answer_content)
            # 收集续行...
            break
        i += 1
```

### 4. BIG_ANSWER_PATTERN 续行收集（需检测小问编号）

```python
# 向后收集续行，直到遇到标记或小问编号
while i <= paras.Count:
    next_text = paras(i).Range.Text.strip()
    if (re.match(r'^【分析】|^【解析】|^【答案】', next_text)
            or QUESTION_PATTERN.match(next_text)
            or SUB_QUESTION_PATTERN.match(next_text)):  # 必须检测小问
        break
    answer_lines.append(next_text)
    i += 1
```

---

## 五、模板竞争得分（测试文档 1.1 青春的邀约.docx）

| 模板 | 得分 | 结果 |
|------|------|------|
| 模板A | 16.22% | — |
| 模板B | 9.01% | — |
| 模板C | **22.52%** | ✅ 胜出 |
