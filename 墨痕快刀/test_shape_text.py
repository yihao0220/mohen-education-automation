import re

ans_text = "22．(1)高压蒸汽灭菌\r(2)104\r\r"
q_prefix_pattern = re.compile(r'^\s*\d+[．.]\s*')
m_prefix = q_prefix_pattern.match(ans_text)
offset_start = m_prefix.end() if m_prefix else 0
text_body = ans_text[offset_start:]

subq_pattern = re.compile(r'(?:^|[\s\r])([\(\（]\d+[\)\）]|[①-⑩])\s*')
matches = list(subq_pattern.finditer(text_body))

for i in range(len(matches)):
    m_curr = matches[i]
    content_start = m_curr.end()
    content_end = matches[i+1].start() if i + 1 < len(matches) else len(text_body)
    sub_text = text_body[content_start:content_end]
    stripped_sub = sub_text.rstrip()
    trailing_len = len(sub_text) - len(stripped_sub)
    print(f"[{stripped_sub}] trailing: {trailing_len}")
