# wps_helper.py
import win32com.client
import time

def get_active_wps():
    """
    尝试获取当前活动的 WPS 或 Word 应用程序对象。
    尝试 5 次，每次间隔 1 秒。
    """
    for i in range(5):
        try:
            try:
                # 优先尝试获取 WPS
                app = win32com.client.GetActiveObject("Kwps.Application")
            except:
                # 失败则尝试获取 Microsoft Word
                app = win32com.client.GetActiveObject("Word.Application")
            _ = app.Name  # 测试对象是否存活
            return app
        except:
            time.sleep(1)
    return None
