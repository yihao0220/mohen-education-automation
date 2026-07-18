# coding: utf-8
"""
模板D测试脚本 - 离线验证清洗逻辑
"""

import sys
import re

# 导入模板D
sys.path.insert(0, '.')
from 格式模板库 import template_d

# 模拟文档段落
TEST_PARAGRAPHS = [
    # 章节标题（保留）
    "课时作业",
    "第一讲 声现象",
    "基础过关",
    # 选择题
    "1. A",
    "2. A 【解析】A. 声音的传播需要介质, 真空不能传声",
    "3. B",
    "4. B 【解析】A. 频率表示每秒钟内振动的次数",
    "5. 信息 传播过程中",
    # 大题
    "6.(1)振动 空气 音色 (2)510 340",
    "能力提升",
    "7.C 8.A",
    "素养拔尖",
    "9.(1)两个隔音盒厚度不同(2)响度",
    "(3)聚酯棉(4)寻找普遍规律",
    "(5)B(6)传播过程中",
    "第二讲 光现象",
    "基础过关",
    "1. D 2. C",
    "3. B 【解析】A. 图甲中，树荫下的亮斑是小孔成像现象",
    "",
    "【解析】独立解析行内容",
    "续行内容",
]

print("=" * 60)
print("模板D - 布心中学物理格式 - 离线测试")
print("=" * 60)

# 测试垃圾行过滤
print("\n【测试1】垃圾行过滤")
for text in ["课时作业", "第一讲 声现象", "基础过关", "1. A", "【解析】内容"]:
    is_garbage = template_d.is_garbage_line(text)
    status = "🚫 垃圾" if is_garbage else "✅ 保留"
    print(f"  {status}: {text[:30]}")

# 测试正则匹配
print("\n【测试2】正则匹配")
test_cases = [
    ("1. A", "选择题"),
    ("1.C", "选择题无空格"),
    ("2. A 【解析】解析内容", "选择题带解析"),
    ("5.(1)水平 调节平衡", "大题带小问"),
    ("6.(1)振动 空气 音色 (2)510 340", "大题多小问同行"),
    ("6. 地球", "大题无小问"),
    ("(1)振动 空气", "纯小问"),
    ("【解析】独立解析", "独立解析"),
]

for text, desc in test_cases:
    choice = template_d.CHOICE_ANSWER_PATTERN.match(text)
    analysis = template_d.ANALYSIS_PATTERN.match(text)
    question = template_d.QUESTION_PATTERN.match(text)
    sub = template_d.SUB_QUESTION_PATTERN.match(text)
    
    matched = []
    if choice:
        matched.append(f"选择题(题号:{choice.group(1)},答案:{choice.group(2)})")
    if analysis:
        matched.append("独立解析")
    if question:
        content = question.group(2)
        matched.append(f"题号(内容:{content[:30]})")
        # 测试多小问拆分
        sub_questions = re.findall(r'[(（](\d+)[)）]([^（(]*?)(?=[(（]\d+[)）]|$)', content)
        if sub_questions:
            matched.append(f"拆分出{len(sub_questions)}个小问")
    if sub:
        matched.append(f"小问({sub.group(1)})")
    
    result = " | ".join(matched) if matched else "未匹配"
    print(f"  {desc}: {result}")

print("\n【测试3】完整流程模拟")
questions_list = []
current_question = None

for i, text in enumerate(TEST_PARAGRAPHS, 1):
    if not text:
        continue
    
    # 垃圾行过滤
    if template_d.is_garbage_line(text):
        print(f"  [行{i:2d}] 🚫 垃圾: {text[:30]}")
        continue
    
    # 章节标题保留
    if template_d.is_section_title(text):
        current_question = {
            'original_num': 0,
            'type': 'section',
            'answer_lines': [text],
            'analysis_lines': [],
        }
        questions_list.append(current_question)
        print(f"  [行{i:2d}] 📑 章节标题: {text[:30]}")
        continue
    
    # 选择题
    choice_match = template_d.CHOICE_ANSWER_PATTERN.match(text)
    if choice_match:
        q_num = int(choice_match.group(1))
        answer = choice_match.group(2)
        remaining = choice_match.group(3).strip()
        
        current_question = {
            'original_num': q_num,
            'type': 'choice',
            'answer_lines': [answer],
            'analysis_lines': [],
        }
        questions_list.append(current_question)
        print(f"  [行{i:2d}] ✅ 选择题: 第{q_num}题 = {answer}")
        
        # 检查行内解析
        inline = template_d.INLINE_ANALYSIS_PATTERN.search(remaining)
        if inline:
            analysis_text = inline.group(1).strip()
            current_question['analysis_lines'].append(analysis_text)
            print(f"  [行{i:2d}]    └─ 解析: {analysis_text[:30]}")
        continue
    
    # 独立解析
    analysis_match = template_d.ANALYSIS_PATTERN.match(text)
    if analysis_match and current_question:
        first_line = analysis_match.group(1)
        if first_line:
            current_question['analysis_lines'].append(first_line)
        print(f"  [行{i:2d}] ✅ 独立解析: {first_line[:30] if first_line else '(空)'}")
        continue
    
    # 题号（大题）
    q_match = template_d.QUESTION_PATTERN.match(text)
    if q_match:
        q_num = int(q_match.group(1))
        content = q_match.group(2).strip()
        
        sub_match = template_d.SUB_QUESTION_PATTERN.match(content)
        if sub_match:
            # 大题带小问
            current_question = {
                'original_num': q_num,
                'type': 'big',
                'answer_lines': [],
                'analysis_lines': [],
            }
            questions_list.append(current_question)
            print(f"  [行{i:2d}] ✅ 大题: 第{q_num}题")
            
            # 使用多小问拆分逻辑（与template_d一致）
            import re
            sub_questions = re.findall(r'[(（](\d+)[)）]([^（(]*?)(?=[(（]\d+[)）]|$)', content)
            if sub_questions:
                for sub_num, sub_content in sub_questions:
                    sub_content = sub_content.strip()
                    if sub_content:
                        answer_line = f"({sub_num}){sub_content}"
                        current_question['answer_lines'].append(answer_line)
                        print(f"  [行{i:2d}]    └─ 小问({sub_num}): {sub_content[:30]}")
            else:
                # 回退到单小问处理
                sub_num = sub_match.group(1)
                sub_content = sub_match.group(2).strip()
                answer_line = f"({sub_num}){sub_content}"
                current_question['answer_lines'].append(answer_line)
                print(f"  [行{i:2d}]    └─ 小问({sub_num}): {sub_content[:30]}")
        else:
            # 大题无小问
            current_question = {
                'original_num': q_num,
                'type': 'big',
                'answer_lines': [content],
                'analysis_lines': [],
            }
            questions_list.append(current_question)
            print(f"  [行{i:2d}] ✅ 大题(无小问): 第{q_num}题 = {content[:30]}")
        continue
    
    # 纯小问续行
    sub_match = template_d.SUB_QUESTION_PATTERN.match(text)
    if sub_match and current_question and current_question['type'] == 'big':
        sub_num = sub_match.group(1)
        sub_content = sub_match.group(2).strip()
        answer_line = f"({sub_num}){sub_content}"
        current_question['answer_lines'].append(answer_line)
        print(f"  [行{i:2d}] ✅ 小问续行({sub_num}): {sub_content[:30]}")
        continue
    
    # 其他（可能是续行）
    if current_question and current_question['type'] == 'big' and current_question['answer_lines']:
        current_question['answer_lines'][-1] += text
        print(f"  [行{i:2d}]    └─ 追加续行: {text[:30]}")
        continue
    
    print(f"  [行{i:2d}] ⚪ 未识别: {text[:30]}")

# 统计
print("\n【测试结果统计】")
print(f"  共发现 {len(questions_list)} 个项目")
choice_count = sum(1 for q in questions_list if q['type'] == 'choice')
big_count = sum(1 for q in questions_list if q['type'] == 'big')
section_count = sum(1 for q in questions_list if q['type'] == 'section')
print(f"  选择题: {choice_count} 道")
print(f"  大题: {big_count} 道")
print(f"  章节标题: {section_count} 个")

print("\n【题目详情】")
for i, q in enumerate(questions_list, 1):
    if q['type'] == 'section':
        title = q['answer_lines'][0] if q['answer_lines'] else "(无)"
        print(f"  [{i}] 📑 章节: {title[:30]}")
    else:
        q_type = "选择题" if q['type'] == 'choice' else "大题"
        answer = q['answer_lines'][0] if q['answer_lines'] else "(无)"
        analysis = "有" if q['analysis_lines'] else "无"
        print(f"  第{i}题 ({q_type}): 答案={answer[:30]}, 解析={analysis}")

print("\n【标准格式输出预览】")
output_idx = 0
for q in questions_list:
    q_type = q['type']
    answer_lines = q['answer_lines']
    analysis_lines = q['analysis_lines']
    
    if q_type == 'section':
        # 章节标题原样输出
        for line in answer_lines:
            print(f"{line}")
        continue
    
    output_idx += 1
    
    if q_type == 'choice':
        letter = answer_lines[0] if answer_lines else ''
        if analysis_lines:
            print(f"{output_idx}．{letter}　解析：{analysis_lines[0][:50]}")
            for line in analysis_lines[1:]:
                print(f"{line}")
        else:
            print(f"{output_idx}．{letter}　")
    elif q_type == 'big':
        print(f"{output_idx}．")
        if answer_lines:
            print(f"答案：{answer_lines[0][:50]}")
            for line in answer_lines[1:]:
                print(f"{line}")
        if analysis_lines:
            print(f"解析：{analysis_lines[0][:50]}")
            for line in analysis_lines[1:]:
                print(f"{line}")

print("\n" + "=" * 60)
print("测试完成")
print("=" * 60)
