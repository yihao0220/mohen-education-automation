# debug_logger.py
import sys

class DebugLogger:
    def __init__(self):
        self.enabled = True  # 默认开启，会在 main.py 中根据 config.DEBUG_MODE 同步

    def log(self, message):
        """普通日志"""
        if self.enabled:
            print(message)

    def section_start(self, title):
        """章节开始"""
        self.log(f"\n🚀 处理：【{title}】")

    def english_mode_start(self):
        """英语模式启动"""
        self.log("   🇬🇧 启用英语【智能吞并】模式")

    def scan_node_found(self, idx, node_type, text):
        """发现潜在节点"""
        self.log(f"      ✅ [发现] 行{idx} [{node_type}]: {text[:20]}")

    def scan_obstacle_filtered(self, idx, text, pattern):
        """过滤干扰项"""
        self.log(f"      🚫 [过滤] 行{idx}: {text[:20]}... (匹配: {pattern[:30]})")

    def merge_decision(self, current_node, next_node, decision, reason):
        """记录吞并决策"""
        status = "➕ 吞并" if decision else "🛑 停止"
        self.log(f"         {status}: 行{next_node['idx']} [{next_node['type']}] {next_node['text'][:10]}... 原因: {reason}")

    def summary(self, total_questions):
        """总结"""
        self.log(f"   📍 准备录入 {total_questions} 题...")

# 全局单例
logger = DebugLogger()
