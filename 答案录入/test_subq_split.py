"""
测试小题拆分逻辑：验证 （1）（2）（3） 能正确拆分
"""
import re
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

text = '（1）邓稼先 1924 年生于安徽。（2）"邓稼先始终站在第一线。"（3）这句话强调邓稼先的核心地位。'

# 新正则：不要求前面是空白
combined_pattern = re.compile(r'([(（]\d+[)）]|[①-⑩])')
first_match = combined_pattern.search(text)

print(f'first_match: {first_match.group(1) if first_match else None}')

matches = []
if first_match:
    matched_marker = first_match.group(1)
    print(f'matched_marker: [{matched_marker}]')
    if re.match(r'^[(（]1[)）]$', matched_marker):
        subq_pattern = re.compile(r'([(（]\d+[)）])')
        matches = list(subq_pattern.finditer(text))

print(f'\n找到 {len(matches)} 个小题号:')
for i, m in enumerate(matches):
    content_start = m.end()
    content_end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
    content = text[content_start:content_end].strip()
    print(f'  小题{m.group(1)}: {content[:40]}...')

# 验证
if len(matches) == 3:
    print('\n✅ 测试通过！小题拆分正确')
else:
    print(f'\n❌ 测试失败！期望3个小题，实际{len(matches)}个')
