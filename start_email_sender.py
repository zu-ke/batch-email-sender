import traceback
import os
import sys
import time
import tkinter as tk

try:
    # 导入主程序中的EmailSender类
    from email_sender_main import EmailSender

    # 主程序入口点
    def main():
        root = tk.Tk()
        app = EmailSender(root)
        root.mainloop()

    if __name__ == "__main__":
        main()

except Exception as e:
    # 捕获所有异常并打印详细信息
    error_message = f"发生错误:\n{str(e)}\n\n详细错误信息:\n{traceback.format_exc()}"
    print(error_message)

    # 创建错误日志
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    with open(os.path.join(log_dir, "error_log.txt"), "w", encoding="utf-8") as f:
        f.write(error_message)

    print("\n错误日志已保存到 logs/error_log.txt")
    input("\n按Enter键退出...") # 等待用户按键而不是自动关闭
