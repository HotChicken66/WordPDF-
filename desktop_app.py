"""
Word/PDF 转换器 - 桌面应用启动器
使用 pywebview 创建独立的桌面窗口
"""
import webview
import threading
import time
import sys
import os
from app import app

# 全局变量存储 Flask 服务器线程
flask_thread = None

def start_flask():
    """在后台线程中启动 Flask 服务"""
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)

class WindowAPI:
    def __init__(self):
        self.window = None

    def minimize(self):
        if self.window:
            self.window.minimize()

    def toggle_fullscreen(self):
        if self.window:
            self.window.toggle_fullscreen()

    def close(self):
        if self.window:
            self.window.destroy()

def main():
    """主函数：启动 Flask 并创建桌面窗口"""
    global flask_thread
    
    # 启动 Flask 服务器（后台线程）
    flask_thread = threading.Thread(target=start_flask, daemon=True)
    flask_thread.start()
    
    # 等待 Flask 启动
    time.sleep(2)
    
    # 创建 API 实例
    api = WindowAPI()
    
    # 创建桌面窗口 (无边框模式)
    window = webview.create_window(
        title='Word/PDF 转换器',
        url='http://127.0.0.1:5000',
        width=1000,
        height=700,
        resizable=True,
        min_size=(800, 600),
        frameless=True,
        js_api=api
    )
    
    # 将窗口实例绑定到 API
    api.window = window
    
    # 启动 webview（阻塞直到窗口关闭）
    webview.start()
    
    # 窗口关闭后，程序自动退出（因为 Flask 线程是 daemon）
    print("应用已关闭")

if __name__ == '__main__':
    main()
