"""
离线测试 parse_answer_blocks 的解析逻辑（不依赖 WPS COM）
模拟文档结构: 1．B 2．C 3．A 4．D （选择题）+ 大题小问
"""
import re
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ——— 复制 answer_input.py 中的正则 ———
q_start_pattern = re.compile(r"^\s*(\d+)[．.]\s*")       # 支持 "1." 和 "1．B"
answer_prefix_pattern = re.compile(r"^\s*答案[：:]\s*")
big_question_pattern = re.compile(r"^\s*[（(][一二三四五六七八九十]+[）)]\s*$")
sub_question_pattern = re.compile(r"^\s*[（(](\d+)[）)]\s*")
analysis_start_pattern = re.compile(r"^\s*解析[：:]")

# ——— 模拟文档段落（邓稼先真实格式）———
# 普通题: "1." + "答案：B" + "解析：xxx"
# 大题: "（一）" + 小问内容 + "解析：xxx"
# 大题小问合并在一行: "答案：（1）xxx（2）xxx（3）xxx"
MOCK_PARAS = [
    "1.",
    "答案：B",
    "解析：这是第1题的解析。",
    "2.",
    "答案：C",
    "解析：这是第2题的解析。",
    "3．A",
    "解析：",
    "4．D",
    "解析：这是第4题的解析。",
    "（一）",
    "（1）这是大题第1小问的答案",
    "（2）这是大题第2小问的答案",
    "解析：这是大题的解析。",
    "（二）",
    "答案：（1）邓稼先 1924 年生于安徽。（2）“邓稼先始终站在第一线。”（3）这句话强调邓稼先的核心地位。",
    "解析：大题二的多行解析内容...",
]

# ——— 模拟解析逻辑 ———
blocks = []
i = 0
in_big_question = False
big_question_qnum = None
sub_q_index = 0
current_big_question_blocks = []
current_q = None

while i < len(MOCK_PARAS):
    text = MOCK_PARAS[i].strip()
    if not text:
        i += 1
        continue

    match_big = big_question_pattern.match(text)
    if match_big:
        if current_q is not None:
            current_q["end_p"] = i - 1
            blocks.append(current_q)
            current_q = None
        in_big_question = True
        big_question_qnum = text.strip()[1:-1]
        sub_q_index = 0
        i += 1
        continue

    match_sub = sub_question_pattern.match(text)
    if match_sub and in_big_question:
        sub_q_index += 1
        block = {
            "qnum": f"{big_question_qnum}.{sub_q_index}",
            "ans_start_p": i,
            "ana_start_p": None,
            "end_p": None,
            "is_sub_question": True,
        }
        current_big_question_blocks.append(block)
        i += 1
        continue

    match_q = q_start_pattern.match(text)
    if match_q:
        if current_big_question_blocks:
            blocks.extend(current_big_question_blocks)
            current_big_question_blocks = []
        if current_q is not None:
            current_q["end_p"] = i - 1
            blocks.append(current_q)
        in_big_question = False
        big_question_qnum = None

        ans_start_p = i
        q_content = text[len(match_q.group(0)):]  # 题号后面的内容
        if not q_content.strip():  # 题号行只有数字，后面是空白
            if i + 1 < len(MOCK_PARAS):
                next_text = MOCK_PARAS[i + 1].strip()
                if answer_prefix_pattern.match(next_text):
                    ans_start_p = i + 1

        current_q = {
            "qnum": match_q.group(1),
            "ans_start_p": ans_start_p,
            "ana_start_p": None,
            "end_p": None,
        }
        i += 1
        continue

    match_ana = analysis_start_pattern.match(text)
    if match_ana:
        if current_big_question_blocks:
            for block in current_big_question_blocks:
                if block["ana_start_p"] is None:
                    block["ana_start_p"] = i
        elif current_q and current_q["ana_start_p"] is None:
            current_q["ana_start_p"] = i

    i += 1

if current_big_question_blocks:
    blocks.extend(current_big_question_blocks)
if current_q is not None:
    current_q["end_p"] = len(MOCK_PARAS) - 1
    blocks.append(current_q)

blocks.sort(key=lambda x: x['ans_start_p'])

# ——— 打印结果 ———
print("=" * 50)
print("解析结果：")
print("=" * 50)
for b in blocks:
    ans_p = b['ans_start_p']
    ans_text = MOCK_PARAS[ans_p] if ans_p < len(MOCK_PARAS) else "(无)"
    ana_p = b['ana_start_p']
    ana_text = MOCK_PARAS[ana_p] if ana_p is not None and ana_p < len(MOCK_PARAS) else "(无解析)"
    print(f"题号: {b['qnum']:<8}  ans_line({ans_p}): {ans_text:<25}  analysis: {ana_text[:20]}")

print()
print(f"✅ 共识别到 {len(blocks)} 题")

# 验证
expected_qnums = ["1", "2", "3", "4", "一.1", "一.2"]
actual_qnums = [b['qnum'] for b in blocks]
if actual_qnums == expected_qnums:
    print("✅ 题号顺序正确：", actual_qnums)
else:
    print("❌ 题号不匹配！")
    print("  期望:", expected_qnums)
    print("  实际:", actual_qnums)
