# main.py
import time
import sys
import os
import re

# ✨ 将 core 文件夹加入运行环境，强行保护内部相互调用
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "core"))

import config
from wps_helper import get_active_wps
from core_parser import detect_subject, get_sections, process_chapter, sync_subject_overlay
from debug_logger import logger


def main():
    while True:
        print("\n" + "=" * 50)
        print("🔌 连接 WPS...")
        wps = get_active_wps()
        if not wps:
            print("❌ 未找到 (5秒重试)...")
            time.sleep(5)
            continue

        try:
            doc = wps.ActiveDocument
            doc_name = doc.Name

            # 更新全局配置（自动检测）
            config.CURRENT_CONFIG = detect_subject(doc)
            print(f"✅ 自动检测模式：【{config.CURRENT_CONFIG['name']}】")
            
            # ✨ 新增：手动切换科目
            manual_switch = input("👉 按回车确认 / 输入 '1' 切换英语 / '2' 切换理科 / '3' 切换文科: ").strip()
            if manual_switch == '1':
                config.CURRENT_CONFIG = config.CONFIG_ENGLISH
                sync_subject_overlay(doc, config.CURRENT_CONFIG["name"])
                print(f"🔄 已手动切换为：【英语】")
            elif manual_switch == '2':
                config.CURRENT_CONFIG = config.CONFIG_SCIENCE
                sync_subject_overlay(doc, config.CURRENT_CONFIG["name"])
                print(f"🔄 已手动切换为：【理科】")
            elif manual_switch == '3':
                config.CURRENT_CONFIG = config.CONFIG_ARTS
                sync_subject_overlay(doc, config.CURRENT_CONFIG["name"])
                print(f"🔄 已手动切换为：【文科】")

            sections = get_sections(doc)
            valid_sections = []
            for i, s in enumerate(sections):
                if s["end"] < s["start"]:
                    continue
                if "文档开头" in s["title"]:
                    if len(sections) > 1:
                        continue
                valid_sections.append(s)

            if not valid_sections:
                valid_sections = [
                    {
                        "title": "全文档 (自动兜底)",
                        "start": 1,
                        "end": doc.Paragraphs.Count,
                    }
                ]

        except:
            time.sleep(2)
            continue

        while True:
            print("\n" + "-" * 50)
            print(f"📚 {doc_name}")
            print("-" * 50)

            for idx, sec in enumerate(valid_sections):
                # 简单清洗标题，避免过长
                clean_title = re.sub(r"\s+", " ", sec["title"])[:40]
                print(f"   [{idx + 1:<3}] {clean_title}...")

            print("-" * 50)
            user_input = input("👉 序号/范围（如 1 或 1-3，回车=全部） / 'r' 刷新 / 'q' 退出: ").strip().lower()

            if user_input == "q":
                sys.exit()
            if user_input == "r":
                break

            try:
                if not user_input:
                    targets = valid_sections
                elif "-" in user_input:
                    s, e = map(int, user_input.split("-"))
                    targets = valid_sections[s - 1 : e]
                else:
                    choice = int(user_input) - 1
                    targets = [valid_sections[choice]]

                for sec in targets:
                    process_chapter(sec)

                print("✅ 结束。")
            except Exception as e:
                print(f"❌ {e}")
                time.sleep(1)
if __name__ == "__main__":
    main()

