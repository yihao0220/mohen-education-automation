import re

text = """22．(1)高压蒸汽灭菌　琼脂　选择　(2)104　(3)S的浓度超过某一值时会抑制菌株的生长　(4)取淤泥加入无菌水中，涂布(或稀释涂布)到乙培养基上，培养后计数　(5)水、碳源、氮源和无机盐"""

# 移除题号
m = re.match(r'^\s*\d+[．.]\s*', text)
start_offset = 0
if m:
    start_offset = len(m.group(0))

ans_text = text[start_offset:]
print("Answer text:", ans_text)

# 判断是否是小题格式：以 (1), （1）, ① 开头
subq_pattern = re.compile(r'(?:^|\s|\r|\n)([\(（]\d+[\)）]|[①-⑩])\s*')

# 查找所有小题标题
matches = list(subq_pattern.finditer(ans_text))
if matches and matches[0].group(1) in ['(1)', '（1）', '①']:
    print("大题，包含小题！")
    for i in range(len(matches)):
        m_curr = matches[i]
        # 内容的起点：小题号之后
        content_start = m_curr.end()
        # 内容的终点：下一个小题号之前，或者是结尾
        content_end = matches[i+1].start() if i + 1 < len(matches) else len(ans_text)
        
        content = ans_text[content_start:content_end]
        # 在原 text 中的绝对位置
        abs_start = start_offset + content_start
        abs_end = start_offset + content_end
        
        print(f"Sub-question {m_curr.group(1)}: [{content}] (abs: {abs_start}-{abs_end})")
else:
    print("普通单题")
