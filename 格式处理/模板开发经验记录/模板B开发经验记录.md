# 模板B开发经验记录

> 记录模板B（【分析】【详解】格式）的完整开发过程，包括问题上报、诊断、修复方案，供后续模板开发参考。

---

## 基本信息

| 项目 | 内容 |
|------|------|
| 模板名称 | 模板B_分析详解格式 |
| 适用格式 | 【答案】A + 【分析】... + 【详解】... + 故选X。 |
| 开发时间 | 2026-03-05 |
| 参考模板 | 无（全新开发） |
| 样本文档 | `3.7.2 食物链和食物网（分层作业）（解析版）.docx` |

---

## 问题记录与修复

### 问题1：模板B从未被加载

**上报时间**: 2026-03-05  
**问题描述**: 运行清洗时始终显示"⚪ 未识别"，【分析】【详解】标记未被识别  
**根本原因**: `__init__.py` 只导出了 template_a，template_b 从未被系统加载

**诊断过程**:
```
用户: 重新放入了新的格式待清洗文件，分析这个文件格式，然后生成模板B
AI: 分析文档格式特征...【答案】+【分析】+【详解】+故选X。
AI: 创建 template_b.py
用户: 【分析】【详解】系统没有识别成功
AI: 检查 __init__.py 发现只导出 template_a
```

**修复方案**:
```python
# 格式处理/格式模板库/__init__.py
from . import template_a
from . import template_b  # 新增

__all__ = ["template_a", "template_b"]  # 新增
```

**经验教训**:
- ✅ 创建新模板后必须检查 `__init__.py` 是否导出
- ✅ 调试时先确认模板是否被加载（打印匹配分数）

---

### 问题2：【详解】内容被截断

**上报时间**: 2026-03-05  
**问题描述**: 第6题详解只有第一句，后续内容丢失  
**根本原因**: 【详解】内循环的停止条件包含 `[A-D][．.、]`，遇到 `B．"清明雨涟涟..."` 就停止

**诊断过程**:
```
用户: 第6题详解只有第一句
AI: 检查 template_b.py 详解内循环
发现: if re.match(r'^\d+[．.]|[A-D][．.、]|^【答案】|^【分析】', next_text):
问题: [A-D][．.、] 把选项解析行误判为停止信号
```

**修复方案**:
```python
# 修改前
if re.match(r'^\d+[．.]|[A-D][．.、]|^【答案】|^【分析】', next_text):
    break

# 修改后
if re.match(r'^\d+[．.]|^【答案】|^【分析】', next_text):
    break
```

**经验教训**:
- ✅ 内循环停止条件要精确，避免过度匹配
- ✅ 选项行（A./B./C./D.）可能是解析内容的一部分，不能作为停止信号
- ✅ 测试时用包含多行解析的题目验证

---

### 问题3：大题答案续行丢失

**上报时间**: 2026-03-05  
**问题描述**: 大题答案只有 `(1)`，`(2)(3)(4)` 续行丢失  
**根本原因**: 识别 `(1)` 后没有向后查找连续的小问行

**诊断过程**:
```
用户: 大题答案只有(1)，后面的(2)(3)(4)没了
AI: 检查代码发现识别【答案】(1)xxx后直接continue，没有收集续行
```

**修复方案**:
```python
big_answer_match = BIG_QUESTION_ANSWER.match(text)
if big_answer_match and current_question:
    answer_text = big_answer_match.group(1)
    if re.search(r'\(\d+\)', answer_text):
        current_question['type'] = 'big'
        answer_lines = [answer_text]
        i += 1
        # 新增：向后查找连续小问
        while i <= paras.Count:
            next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
            if not next_text:
                i += 1
                continue
            if re.match(r'^\(\d+\)', next_text):
                answer_lines.append(next_text)
                i += 1
            else:
                break
        current_question['answer_lines'] = answer_lines
        continue
```

**经验教训**:
- ✅ 大题答案/解析通常有多行，必须实现向后查找内循环
- ✅ 小问格式 `(1)xxx` 是识别续行的关键特征

---

### 问题4：输出格式未保留【分析】【详解】标记

**上报时间**: 2026-03-05  
**问题描述**: 清洗后输出没有【分析】【详解】标记，只有纯文本  
**根本原因**: 输出时直接拼接 `analysis_lines` 和 `detail_lines`，没有添加标记

**修复方案**:
```python
# 选择题输出
if q_type == 'choice':
    doc.Content.InsertAfter(f"{idx}．{letter}\r")
    if analysis_lines:
        doc.Content.InsertAfter(f"解析：【分析】{analysis_lines[0]}\r")  # 保留【分析】
        for line in analysis_lines[1:]:
            doc.Content.InsertAfter(f"{line}\r")
    for line in detail_lines:
        doc.Content.InsertAfter(f"{line}\r")  # detail_lines[0] 已包含【详解】

# 大题输出
elif q_type == 'big':
    doc.Content.InsertAfter(f"{idx}．\r")
    for ans_line in answer_lines:
        doc.Content.InsertAfter(f"{ans_line}\r")
    if analysis_lines:
        doc.Content.InsertAfter(f"解析：【分析】{analysis_lines[0]}\r")
        for line in analysis_lines[1:]:
            doc.Content.InsertAfter(f"{line}\r")
    for line in detail_lines:
        doc.Content.InsertAfter(f"{line}\r")
```

**经验教训**:
- ✅ 输出格式必须与 `Standard output format.md` 一致
- ✅ 保留原始标记（【分析】【详解】）便于后续处理

---

### 问题5：章节编号误识别为题号

**上报时间**: 2026-03-05  
**问题描述**: 章节标题 `3.7.2 食物链和食物网` 中的 `3.` 被识别为第3题，导致第1题变成第2题  
**根本原因**: `QUESTION_PATTERN` 匹配了 `3.7.2` 中的 `3.`

**诊断过程**:
```
用户: 第1题变第2题了
AI: 检查日志发现先识别了"第3题"（来自3.7.2），然后才是真的第1题
```

**修复方案**:
```python
# GARBAGE_PATTERNS 新增
GARBAGE_PATTERNS = [
    # ... 其他模式 ...
    r"^\d+\.\d+\.",  # 章节编号（如 3.7.2 食物链和食物网）
    # ...
]
```

**经验教训**:
- ✅ 章节编号 `x.x.x` 是常见格式，必须过滤
- ✅ 题号识别前必须先过滤垃圾行

---

### 问题6：【分析】内容被截断 + 序号错乱

**上报时间**: 2026-03-05  
**问题描述**: 第16题解析不全，第17题变成第19题  
**根本原因**: 【分析】内循环的停止条件包含 `^\d+[．.]`，把分析内容中的编号列表 `2．观察法...` 误判为新题目

**诊断过程**:
```
用户: 第16题解析不全，第17题变第19题
AI: 检查【分析】内循环停止条件
发现: if re.match(r'^【详解】|^【答案】|^\d+[．.]', next_text):
问题: 分析内容中的编号列表被误判
```

**修复方案**:
```python
# 修改前
if re.match(r'^【详解】|^【答案】|^\d+[．.]', next_text):
    break

# 修改后（删除 ^\d+[．.] 条件）
if re.match(r'^【详解】|^【答案】', next_text):
    break
```

**经验教训**:
- ✅ 分析/解析内容中可能包含编号列表（1. 2. 3.），不能作为停止信号
- ✅ 停止条件应仅限于明确的标记（【详解】【答案】）
- ✅ 一个根因可能导致多个表面现象（解析截断 + 序号错乱）

---

## 开发流程优化建议

基于模板B的开发经验，后续模板开发应遵循：

### 第一阶段：文档分析
1. 提供 `task_instruction.md` 描述目标格式
2. 提供样本文档（至少1个完整文件）
3. 人工标注文档结构：
   - 题号格式（数字 + 全角/半角点号）
   - 答案格式（【答案】/ [答案] / 直接跟字母）
   - 解析格式（【分析】+【详解】/ [解析] / 直接跟内容）
   - 垃圾行类型（章节标题、题型说明等）

### 第二阶段：模板编写
1. 选择最接近的参考模板（A或B）
2. 复制并修改 `TEMPLATE_FEATURES` 和 `GARBAGE_PATTERNS`
3. 实现 `clean_document` 核心逻辑
4. **必须实现** `set_font_format` 统一字体

### 第三阶段：测试验证
1. **单文档测试**: 验证输出格式、字体、题号连续性
2. **边界测试**: 包含编号列表的解析、多小问大题
3. **批量测试**: 验证稳定性

### 第四阶段：问题归档
1. 将问题记录到本文件
2. 更新 `TEMPLATE_DEV_GUIDE.md`
3. 提交 git commit

---

## 关键代码片段

### 内循环收集模式（向后查找）

```python
# 收集连续的多行内容（如大题小问、多行解析）
lines = [first_line]
i += 1
while i <= paras.Count:
    next_text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
    if not next_text:
        i += 1
        continue
    # 停止条件：遇到明确的标记或新题目
    if re.match(r'^【详解】|^【答案】', next_text):  # 精确停止
        break
    lines.append(next_text)
    i += 1
```

### 字体格式统一

```python
def set_font_format(doc):
    try:
        for para in doc.Paragraphs:
            for run in para.Range.Words:
                run.Font.Size = 12
                run.Font.Color = 0
                run.Font.Name = "Times New Roman"
                run.Font.NameFarEast = "宋体"
                run.Font.Bold = False
        print("   ✓ 字体格式设置完成")
    except Exception as e:
        print(f"   ! 字体设置失败: {e}")
```

### 题号连续性校验

```python
# 重新编号确保连续
sorted_questions = sorted(questions_dict.items(), key=lambda x: x[1]['order'])
question_map = {}
new_num = 1
for original_num, _ in sorted_questions:
    question_map[original_num] = new_num
    new_num += 1
```

---

## 附录：调试技巧

### 1. 打印每行处理结果

```python
for i in range(1, paras.Count + 1):
    text = paras(i).Range.Text.replace("\r", "").replace("\n", "").strip()
    debug_text = text[:60] + "..." if len(text) > 60 else text
    print(f"   [行{i:3d}] {debug_text}")
```

### 2. 检查模板匹配分数

```python
# 在 main.py 的 match_best_template 中添加
print(f"   模板A匹配度: {template_a.match_score(doc):.2%}")
print(f"   模板B匹配度: {template_b.match_score(doc):.2%}")
```

### 3. 验证正则匹配

```python
import re
text = "【分析】1．科学探究的基本方法..."
pattern = re.compile(r"^【分析】\s*(.*)$")
match = pattern.match(text)
print(match.group(1))  # 验证提取内容
```

---

**记录创建**: 2026-03-05  
**最后更新**: 2026-03-05  
**维护者**: AI Assistant + 用户协作
