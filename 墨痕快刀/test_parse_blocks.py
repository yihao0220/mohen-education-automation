import re

paras_text = [
    "1．A",
    "解析：这是第一题的解析",
    "第二段解析",
    "2．B",
    "3．C",
    "解析：这里是第3题解析"
]

q_start_pattern = re.compile(r'^\s*(\d+)[．.]')
analysis_start_pattern = re.compile(r'^\s*解析：')

blocks = []
current_q = None

for i, text in enumerate(paras_text, 1):
    match_q = q_start_pattern.match(text)
    if match_q:
        if current_q:
            current_q['end_p'] = i - 1
            blocks.append(current_q)
        current_q = {
            'qnum': match_q.group(1),
            'ans_start_p': i,
            'ana_start_p': None,
            'end_p': None
        }
        continue
        
    match_ana = analysis_start_pattern.match(text)
    if match_ana and current_q and current_q['ana_start_p'] is None:
        current_q['ana_start_p'] = i

if current_q:
    current_q['end_p'] = len(paras_text)
    blocks.append(current_q)

for b in blocks:
    ans_end = (b['ana_start_p'] - 1) if b['ana_start_p'] else b['end_p']
    print(f"Q{b['qnum']}: Ans[{b['ans_start_p']}-{ans_end}], Ana[{b['ana_start_p']}-{b['end_p']}]")
