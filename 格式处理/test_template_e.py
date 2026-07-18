# coding: utf-8
"""
模板E测试脚本 - 离线验证（无需WPS）

测试内容：
1. 正则模式匹配测试
2. 垃圾行过滤测试
3. 完整流程模拟测试
"""

import re
import sys
import os

# 添加模板库路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入模板E的正则和函数
from 格式模板库.template_e import (
    TEMPLATE_FEATURES,
    GARBAGE_PATTERNS,
    QUESTION_PATTERN,
    QUESTION_ANSWER_COMBO,
    CHOICE_ANSWER_PATTERN,
    BIG_QUESTION_ANSWER,
    SUB_QUESTION_PATTERN,
    is_garbage_line,
    is_option_line,
)


def test_regex_patterns():
    """测试正则模式"""
    print("=" * 60)
    print("测试1: 正则模式匹配")
    print("=" * 60)
    
    test_cases = [
        # (输入, 期望匹配的模式, 描述)
        ("1．【答案】A", "QUESTION_ANSWER_COMBO", "选择题组合格式（全角点号）"),
        ("3【答案】B", "QUESTION_ANSWER_COMBO", "选择题组合格式（无点号）"),
        ("10【答案】C", "QUESTION_ANSWER_COMBO", "选择题组合格式（两位数题号）"),
        ("14．【答案】(1)2NA", "QUESTION_PATTERN", "大题带答案"),
        ("(1)2NA或1.204×1024", "SUB_QUESTION_PATTERN", "小问编号"),
        ("(2)1∶2", "SUB_QUESTION_PATTERN", "小问编号2"),
        ("【答案】A", "CHOICE_ANSWER_PATTERN", "独立答案行"),
        ("【答案】(1)xxx (2)yyy", "BIG_QUESTION_ANSWER", "大题答案行"),
    ]
    
    patterns = {
        "QUESTION_ANSWER_COMBO": QUESTION_ANSWER_COMBO,
        "QUESTION_PATTERN": QUESTION_PATTERN,
        "SUB_QUESTION_PATTERN": SUB_QUESTION_PATTERN,
        "CHOICE_ANSWER_PATTERN": CHOICE_ANSWER_PATTERN,
        "BIG_QUESTION_ANSWER": BIG_QUESTION_ANSWER,
    }
    
    passed = 0
    failed = 0
    
    for text, pattern_name, desc in test_cases:
        pattern = patterns.get(pattern_name)
        match = pattern.match(text) if pattern else None
        status = "✅" if match else "❌"
        if match:
            passed += 1
            groups = match.groups() if hasattr(match, 'groups') else ()
            print(f"  {status} {desc}: '{text[:30]}...' -> 匹配组: {groups}")
        else:
            failed += 1
            print(f"  {status} {desc}: '{text[:30]}...' -> 未匹配")
    
    print(f"\n  结果: {passed} 通过, {failed} 失败")
    return failed == 0


def test_garbage_filtering():
    """测试垃圾行过滤"""
    print("\n" + "=" * 60)
    print("测试2: 垃圾行过滤")
    print("=" * 60)
    
    garbage_cases = [
        ("2025-2026学年高二化学单元检测卷", "学年标题"),
        ("姓名：________", "姓名行"),
        ("班级：_______", "班级行"),
        ("满分100分", "满分行"),
        ("用时：40min", "用时行"),
        ("一、选择题", "题型标题"),
        ("二、非选择题", "题型标题2"),
        ("基础知识", "小节标题"),
        ("A．选项内容", "选项A"),
        ("B．选项内容", "选项B"),
    ]
    
    valid_cases = [
        ("1．【答案】A", "选择题答案"),
        ("14．【答案】(1)2NA", "大题答案"),
        ("(1)2NA或1.204×1024", "小问答案"),
        ("(2)1∶2", "小问答案2"),
    ]
    
    passed = 0
    failed = 0
    
    print("  垃圾行（应被过滤）:")
    for text, desc in garbage_cases:
        is_garbage = is_garbage_line(text)
        status = "✅" if is_garbage else "❌"
        if is_garbage:
            passed += 1
        else:
            failed += 1
        print(f"    {status} {desc}: '{text[:30]}...'")
    
    print("\n  有效行（应保留）:")
    for text, desc in valid_cases:
        is_garbage = is_garbage_line(text)
        status = "✅" if not is_garbage else "❌"
        if not is_garbage:
            passed += 1
        else:
            failed += 1
        print(f"    {status} {desc}: '{text[:30]}...'")
    
    print(f"\n  结果: {passed} 通过, {failed} 失败")
    return failed == 0


def test_choice_vs_big():
    """测试选择题与大题区分"""
    print("\n" + "=" * 60)
    print("测试3: 选择题与大题区分")
    print("=" * 60)
    
    test_cases = [
        # (【答案】后的内容, 期望类型)
        ("A", "choice", "单字母A"),
        ("B", "choice", "单字母B"),
        ("C", "choice", "单字母C"),
        ("D", "choice", "单字母D"),
        ("(1)2NA", "big", "小问格式"),
        ("(1)xxx (2)yyy", "big", "多小问"),
        ("氢键", "big", "文字答案"),
        ("N>O>C", "big", "符号答案"),
    ]
    
    passed = 0
    failed = 0
    
    for answer_text, expected_type, desc in test_cases:
        # 使用与模板相同的逻辑判断
        is_choice = bool(re.match(r'^[A-D]$', answer_text))
        actual_type = "choice" if is_choice else "big"
        
        status = "✅" if actual_type == expected_type else "❌"
        if actual_type == expected_type:
            passed += 1
        else:
            failed += 1
        
        type_name = "选择题" if actual_type == "choice" else "大题"
        print(f"  {status} {desc}: '{answer_text[:20]}...' -> {type_name}")
    
    print(f"\n  结果: {passed} 通过, {failed} 失败")
    return failed == 0


def test_full_simulation():
    """模拟完整处理流程"""
    print("\n" + "=" * 60)
    print("测试4: 完整流程模拟")
    print("=" * 60)
    
    # 模拟文档内容（来自高二限时训练1答案.docx）
    mock_document = [
        "2025-2026学年高二化学单元检测卷",
        "一、选择题(本大题共14个小题，每题3分，共42分，)",
        "1．【答案】A",
        "2．【答案】C",
        "3【答案】B",
        "4【答案】B",
        "5【答案】B",
        "6．【答案】C",
        "7．【答案】D",
        "8．【答案】A",
        "9．【答案】A",
        "10【答案】C",
        "11．【答案】D",
        "12【答案】B",
        "二、非选择题(本大题共4个小题，共58分)",
        "14．【答案】(1)2NA或1.204×1024     (2)1∶2      1∶1",
        "(3)3        (4)5∶1",
        "(5)6NA或3.612×1024      7",
        "(6)HF＞H2O＞NH3＞CH4",
        "15．【答案】(1)1s22s22p63s23p63d104s24p2(或[Ar]4s24p2)   (2)共价键",
        "(3)正四面体   sp3  非极性分子",
        "(4)GeCl4     组成和结构相似的分子，相对分子质量越大，分子间作用力越大，熔、沸点越高",
    ]
    
    questions_list = []
    current_question = None
    
    print("  模拟处理过程:")
    for i, text in enumerate(mock_document, 1):
        # 垃圾行过滤
        if is_garbage_line(text):
            print(f"    [行{i:2d}] 🚫 跳过: {text[:30]}...")
            continue
        
        # 检查是否是题号+答案组合格式
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
            print(f"    [行{i:2d}] ✅ 选择题: 第{q_num:2d}题 = {answer}")
            continue
        
        # 新题目（数字开头，包含【答案】）
        q_match = QUESTION_PATTERN.match(text)
        if q_match and '【答案】' in text:
            q_num = int(q_match.group(1))
            remaining = q_match.group(2)
            
            # 提取【答案】后的内容
            answer_match = re.match(r'【答案】\s*(.+)', remaining)
            if answer_match:
                answer_text = answer_match.group(1)
                
                # 检查是否是选择题答案
                if re.match(r'^[A-D]$', answer_text):
                    current_question = {
                        'original_num': q_num,
                        'type': 'choice',
                        'answer_lines': [answer_text],
                    }
                    questions_list.append(current_question)
                    print(f"    [行{i:2d}] ✅ 选择题: 第{q_num:2d}题 = {answer_text}")
                else:
                    # 大题答案
                    current_question = {
                        'original_num': q_num,
                        'type': 'big',
                        'answer_lines': [answer_text],
                    }
                    questions_list.append(current_question)
                    print(f"    [行{i:2d}] 📋 大题: 第{q_num:2d}题 {answer_text[:30]}...")
                
                continue
        
        # 大题答案续行
        if current_question and current_question['type'] == 'big':
            if SUB_QUESTION_PATTERN.match(text):
                current_question['answer_lines'].append(text)
                print(f"    [行{i:2d}]    └─ 追加小问: {text[:30]}...")
                continue
            
            # 答案内容续行
            if not QUESTION_PATTERN.match(text) and not text.startswith('【答案】'):
                if current_question['answer_lines']:
                    current_question['answer_lines'][-1] += " " + text
                    print(f"    [行{i:2d}]    └─ 追加续行: {text[:30]}...")
                continue
        
        print(f"    [行{i:2d}] ⚪ 未识别: {text[:30]}...")
    
    # 统计结果
    print(f"\n  统计结果:")
    print(f"    - 共发现 {len(questions_list)} 道题目")
    choice_count = sum(1 for q in questions_list if q['type'] == 'choice')
    big_count = sum(1 for q in questions_list if q['type'] == 'big')
    print(f"    - 选择题: {choice_count} 道")
    print(f"    - 大题: {big_count} 道")
    
    # 验证结果
    expected_questions = 14  # 12道选择题 + 2道大题
    passed = len(questions_list) == expected_questions
    
    print(f"\n  验证: 期望 {expected_questions} 题, 实际 {len(questions_list)} 题")
    print(f"  结果: {'✅ 通过' if passed else '❌ 失败'}")
    
    return passed


def main():
    """主测试函数"""
    print("\n" + "=" * 60)
    print("模板E测试套件")
    print("=" * 60)
    
    results = []
    
    results.append(("正则模式匹配", test_regex_patterns()))
    results.append(("垃圾行过滤", test_garbage_filtering()))
    results.append(("选择题与大题区分", test_choice_vs_big()))
    results.append(("完整流程模拟", test_full_simulation()))
    
    print("\n" + "=" * 60)
    print("测试汇总")
    print("=" * 60)
    
    for name, passed in results:
        status = "✅ 通过" if passed else "❌ 失败"
        print(f"  {status}: {name}")
    
    all_passed = all(r[1] for r in results)
    
    print("\n" + "=" * 60)
    if all_passed:
        print("🎉 所有测试通过！")
    else:
        print("⚠️  部分测试失败，请检查代码")
    print("=" * 60)
    
    return all_passed


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
