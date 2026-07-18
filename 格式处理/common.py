# coding: utf-8
"""
公共模块 - 存放答案格式清洗的通用功能

包含：
- 调试日志
- WPS 连接
- 文档字体设置
"""

import os
import time

# 尝试导入 win32com
try:
    import win32com.client as win32com_client
except ImportError:
    win32com_client = None  # type: ignore

DEBUG_MODE = True


def debug_log(msg, level="info"):
    """统一调试日志输出"""
    if not DEBUG_MODE:
        return
    
    prefixes = {
        "info": "   ℹ️ ",
        "filter": "   🚫 ",
        "split": "   ✂️ ",
        "match": "   ✅ ",
        "fill": "   📝 ",
        "llm": "   🤖 ",
        "json": "   📦 ",
        "action": "   🔨 ",
        "error": "   ❌ ",
    }
    prefix = prefixes.get(level, "   ℹ️ ")
    print(f"{prefix}{msg}")


def get_active_wps():
    """
    获取当前活动的 WPS 或 Word 应用程序对象
    尝试 5 次，每次间隔 1 秒
    """
    if win32com_client is None:
        return None
    
    for _ in range(5):
        try:
            try:
                app = win32com_client.GetActiveObject("Kwps.Application")
            except:
                app = win32com_client.GetActiveObject("Word.Application")
            _ = app.Name  # 测试对象是否存活
            return app
        except:
            time.sleep(1)
    return None


def set_document_font(doc):
    """
    统一设置文档字体样式
    - 颜色：纯黑色
    - 英文字体：Times New Roman
    - 中文字体：宋体
    - 字号：小四（12磅）
    - 取消加粗和斜体
    """
    try:
        font = doc.Content.Font
        font.Color = 0  # 纯黑色
        font.Name = "Times New Roman"
        font.NameFarEast = "宋体"
        font.Size = 12
        font.Bold = False
        font.Italic = False
        debug_log("全文已刷成黑色，并统一格式为：宋体/Times New Roman，小四，取消加粗/斜体", "info")
        return True
    except Exception as e:
        debug_log(f"设置字体样式失败: {e}", "error")
        return False


