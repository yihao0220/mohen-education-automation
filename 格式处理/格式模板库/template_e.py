# coding: utf-8
"""
格式模板E - 纯答案格式（无详解）

适用文档特征：
- 选择题答案格式: `【答案】X` 如 `【答案】A`
- 无【详解】或【解析】标记
- 大题答案格式: `【答案】(1)xxx (2)xxx`，续行可能跨多行
- 答案内容直接跟在【答案】后，无额外解析

典型文档：高二化学限时训练答案（简化版）
"""

import re
import os

from . import matches_garbage_pattern, pattern_match_score, set_standard_font

# 格式识别特征
TEMPLATE_FEATURES = {
    "name": "高二化学 - 纯答案格式",
    "patterns": [
        r"^【答案】",               # 答案标记
        r"^\d+[．.]\s*【答案】",     # 题号+答案组合
    ],
    "match_threshold": 0.03,
}

# 需要删除的垃圾行模式
GARBAGE_PATTERNS = [
    r"^课时分层作业",
    r"^\(建议用时",
    r"^题组[一二三四五六七八九十]",
    r"^能力提升卷",
    r"^第\d+章",
    r"^一、选择题",
    r"^二、非选择题",
    r"^三、非选择题",
    r"^本题包括\d+小题",
    r"^每小题",
    r"^基础\+提升\+长句表达",
    r"^限时\d+层提升",
    r"^\d+\.\d+\.",                # 章节编号（如 3.7.2）
    r"^\d+[．.]\d+",               # 子标题：5.1 / 5．1
    r"^20\d\d-20\d\d学年",         # 学年标题
    r"^姓名：",
    r"^班级：",
    r"^满分",
    r"^用时：",
    r"^基础知识",
    r"^二\.选择题",
    r"^下列有关",                  # 题干开头（不应出现在答案文档中）
    r"^A[．.]",                     # 选项行
    r"^B[．.]",
    r"^C[．.]",
    r"^D[．.]",
    r"^\s*$",                      # 空行
]

# 小问模式 (1) (2) (3) 等 - 用于识别大题的小问
SUB_QUESTION_PATTERN = re.compile(r"^\(\d+\)")

# 题号模式（支持数字、中文、英文开头）
QUESTION_PATTERN = re.compile(r"^(\d+)[．.]\s*(.+)")

# 选择题答案模式: 【答案】A（可能无空格）
CHOICE_ANSWER_PATTERN = re.compile(r"^【答案】\s*([A-D]+)$")

# 大题答案模式: 【答案】(1)xxx 或 【答案】内容
BIG_QUESTION_ANSWER = re.compile(r"^【答案】\s*(.+)$")

# 题号+答案组合模式: 1．【答案】A 或 1【答案】A
QUESTION_ANSWER_COMBO = re.compile(r"^(\d+)[．.]?\s*【答案】\s*([A-D])$")


def is_garbage_line(text):
    return matches_garbage_pattern(text, GARBAGE_PATTERNS)


def is_option_line(text):
    """判断是否是选项行 A. B. C. D."""
    return bool(re.match(r"^[A-D][．.、]", text))


def match_score(doc, cached_texts=None):
    return pattern_match_score(doc, TEMPLATE_FEATURES["patterns"], cached_texts)


def set_font_format(doc):
    set_standard_font(doc)


def clean_document(doc):
    """
    清洗文档（WPS COM 对象）
    
    Args:
        doc: WPS 文档对象
    
    Returns:
        bool: 是否成功
    """
    print(f"   ▶ 使用模板E清洗: {doc.Name}")
    print(f"   📊 文档共 {doc.Paragraphs.Count} 个段落")
    
    # 1. 删除所有图片
    try:
        shapes_count = doc.Shapes.Count
        if shapes_count > 0:
            for i in range(shapes_count, 0, -1):
                doc.Shapes(i).Delete()
            print(f"   ✓ 删除 {shapes_count} 个悬浮图片")
        inline_count = doc.InlineShapes.Count
        if inline_count > 0:
            for i in range(inline_count, 0, -1):
                doc.InlineShapes(i).Delete()
            print(f"   ✓ 删除 {inline_count} 个嵌入图片")
    except Exception as e:
        print(f"   ! 删除图片失败: {e}")
    
    # 2. 第一遍遍历：收集所有有效内容
    paras = doc.Paragraphs
    questions_list = []   # 按顺序存储
    current_question = None
    
    print(f"\n   🔍 开始解析文档内容...\n")
    
    i = 1
    while i <= paras.Count:
        p = paras(i)
        text = p.Range.Text.replace("\r", "").replace("\n", "").strip()
        
        if not text:
            i += 1
            continue
        
        debug_text = text[:60] + "..." if len(text) > 60 else text
        
        # 垃圾行过滤
        if is_garbage_line(text):
            print(f"   [行{i:3d}] 🚫 跳过垃圾行: {debug_text}")
            i += 1
            continue
        
        # 检查是否是题号+答案组合格式: 1．【答案】A
        combo_match = QUESTION_ANSWER_COMBO.match(text)
        if combo_match:
            q_num = int(combo_match.group(1))
            answer = combo_match.group(2)
            current_question = {
                'original_num': q_num,
                'type': 'choice',
                'answer_lines': [answer],
            }
            questions_list.append(current_question)
            print(f"   [行{i:3d}] ✅ 选择题(组合格式): 第{q_num}题 = {answer}")
            i += 1
            continue
        
        # 新题目（数字开头，但不含【答案】）
        q_match = QUESTION_PATTERN.match(text)
        if q_match and '【答案】' not in text:
            # 这只是题干，跳过（答案文档中不应有题干）
            print(f"   [行{i:3d}] ⏭️  跳过题干: {debug_text}")
            i += 1
            continue
        
        # 新题目（数字开头，包含【答案】）- 大题格式
        if q_match and '【答案】' in text:
            q_num = int(q_match.group(1))
            remaining = q_match.group(2)
            
            # 提取【答案】后的内容
            answer_match = re.match(r'【答案】\s*(.+)', remaining)
            if answer_match:
                answer_text = answer_match.group(1)
                
                # 检查是否是选择题答案（单字母A-D）
                if re.match(r'^[A-D]$', answer_text):
                    current_question = {
                        'original_num': q_num,
                        'type': 'choice',
                        'answer_lines': [answer_text],
                    }
                    questions_list.append(current_question)
                    print(f"   [行{i:3d}] ✅ 选择题: 第{q_num}题 = {answer_text}")
                else:
                    # 大题答案
                    current_question = {
                        'original_num': q_num,
                        'type': 'big',
                        'answer_lines': [],
                    }
                    
                    # 检查是否包含小问编号 (1)(2)(3)
                    if re.search(r'\(\d+\)', answer_text):
                        # 第一行包含小问
                        current_question['answer_lines'].append(answer_text)
                        print(f"   [行{i:3d}] 📋 大题答案: 第{q_num}题 {answer_text[:40]}")
                    else:
                        # 单行大题答案
                        current_question['answer_lines'].append(answer_text)
                        print(f"   [行{i:3d}] 📋 大题答案(单行): 第{q_num}题 {answer_text[:40]}")
                    
                    questions_list.append(current_question)
                
                i += 1
                continue
        
        # 独立的选择题答案行 【答案】A
        choice_match = CHOICE_ANSWER_PATTERN.match(text)
        if choice_match and current_question is None:
            # 没有前置题号的独立答案行，尝试向前查找题号
            answer = choice_match.group(1)
            # 这种情况不应该发生，因为我们期望题号+答案的组合
            print(f"   [行{i:3d}] ⚠️  独立答案行(无题号): 【答案】{answer}")
            i += 1
            continue
        
        # 大题答案续行收集（当current_question是大题时）
        if current_question and current_question['type'] == 'big':
            # 检查是否是小问续行 (1) (2) (3)
            if SUB_QUESTION_PATTERN.match(text):
                current_question['answer_lines'].append(text)
                print(f"   [行{i:3d}]    └─ 追加小问答案: {text[:40]}")
                i += 1
                continue
            
            # 检查是否是答案内容的续行（不以题号开头，不是垃圾行）
            if not QUESTION_PATTERN.match(text) and not text.startswith('【答案】'):
                # 追加到最后一行答案
                if current_question['answer_lines']:
                    current_question['answer_lines'][-1] += " " + text
                    print(f"   [行{i:3d}]    └─ 追加续行: {text[:40]}")
                i += 1
                continue
        
        # 选项行 A. B. C. D. （跳过）
        if is_option_line(text):
            print(f"   [行{i:3d}] ⏭️  跳过选项: {debug_text}")
            i += 1
            continue
        
        # 未识别行
        print(f"   [行{i:3d}] ⚪ 未识别: {debug_text}")
        i += 1
    
    # 3. 统计
    print(f"\n   📊 解析完成统计:")
    print(f"      - 共发现 {len(questions_list)} 道题目")
    choice_found = sum(1 for q in questions_list if q['type'] == 'choice')
    big_found = sum(1 for q in questions_list if q['type'] == 'big')
    print(f"      - 选择题: {choice_found} 道，大题: {big_found} 道")
    
    # 4. 清空文档并重新写入（顺序编号）
    doc.Content.Text = ""
    
    choice_count = 0
    big_count = 0
    
    for idx, q_data in enumerate(questions_list, 1):
        q_type = q_data['type']
        answer_lines = q_data['answer_lines']
        
        if q_type == 'choice':
            choice_count += 1
            letter = answer_lines[0] if answer_lines else ''
            # 选择题格式: 1．A　（无解析）
            doc.Content.InsertAfter(f"{idx}．{letter}　\r")
        
        elif q_type == 'big':
            big_count += 1
            # 大题格式: 题号后换行，答案独立成行
            doc.Content.InsertAfter(f"{idx}．\r")
            for ans_line in answer_lines:
                doc.Content.InsertAfter(f"{ans_line}\r")
    
    print(f"   ✓ 共处理 {choice_count} 道选择题, {big_count} 道大题")
    
    # 5. 设置字体格式
    set_font_format(doc)
    
    return True


# 模块导出
TEMPLATE_INFO = TEMPLATE_FEATURES
