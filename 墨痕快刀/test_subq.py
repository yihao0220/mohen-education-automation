import re

text = "22．(1)高压蒸汽灭菌　琼脂　选择　(2)104　(3)S的浓度超过某一值时会抑制菌株的生长　(4)取淤泥加入无菌水中，涂布(或稀释涂布)到乙培养基上，培养后计数　(5)水、碳源、氮源和无机盐"

# Remove the main question number
m = re.match(r'^\s*\d+[．.]\s*', text)
if m:
    text = text[len(m.group(0)):]

print("Text without qnum:", text)

# Try to find sub-questions like (1), （1）, ①
subq_pattern = re.compile(r'(\([1-9]\)|（[1-9]）|[①②③④⑤⑥⑦⑧⑨⑩])')
parts = subq_pattern.split(text)

print(parts)
