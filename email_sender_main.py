import pandas as pd
import smtplib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import threading
import time
import datetime
import os
import random
import re
import traceback
import sys

class EmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("批量邮件发送工具")
        self.root.geometry("700x650")
        self.root.resizable(True, True)

        self.email_data = None
        self.sender_accounts = None
        self.current_sender_index = 0

        self.is_sending = False
        self.current_job = None

        # 创建UI组件
        self.create_widgets()

    def create_widgets(self):
        # 发件人配置框架
        sender_frame = ttk.LabelFrame(self.root, text="发件人设置")
        sender_frame.pack(fill="x", padx=10, pady=5)

        # 单一发件人设置
        self.single_sender_var = tk.BooleanVar(value=True)
        ttk.Radiobutton(sender_frame, text="使用单一发件人", variable=self.single_sender_var,
                       value=True, command=self.toggle_sender_mode).grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.single_sender_frame = ttk.Frame(sender_frame)
        self.single_sender_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="w")

        ttk.Label(self.single_sender_frame, text="发件人邮箱:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.email_entry = ttk.Entry(self.single_sender_frame, width=30)
        self.email_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(self.single_sender_frame, text="@163.com").grid(row=0, column=2, padx=0, pady=5, sticky="w")

        ttk.Label(self.single_sender_frame, text="邮箱密码/授权码:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.password_entry = ttk.Entry(self.single_sender_frame, width=30, show="*")
        self.password_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # 多发件人设置
        ttk.Radiobutton(sender_frame, text="使用多个发件人轮询发送", variable=self.single_sender_var,
                       value=False, command=self.toggle_sender_mode).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.multi_sender_frame = ttk.Frame(sender_frame)
        self.multi_sender_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        self.multi_sender_frame.grid_remove()  # 默认隐藏

        ttk.Label(self.multi_sender_frame, text="发件人Excel文件:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.sender_file_entry = ttk.Entry(self.multi_sender_frame, width=40)
        self.sender_file_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(self.multi_sender_frame, text="浏览", command=self.browse_sender_file).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(self.multi_sender_frame, text="加载发件人", command=self.load_senders).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.multi_sender_frame, text="轮询方式:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.rotation_mode = tk.StringVar(value="batch")
        ttk.Radiobutton(self.multi_sender_frame, text="每批一个账号", variable=self.rotation_mode,
                      value="batch").grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(self.multi_sender_frame, text="每封一个账号", variable=self.rotation_mode,
                      value="email").grid(row=2, column=2, padx=5, pady=5, sticky="w")

        # 文件选择框架
        file_frame = ttk.LabelFrame(self.root, text="收件人设置")
        file_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(file_frame, text="收件人Excel文件:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.file_path_entry = ttk.Entry(file_frame, width=50)
        self.file_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(file_frame, text="浏览", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(file_frame, text="加载数据", command=self.load_data).grid(row=1, column=1, padx=5, pady=5)

        # 发送设置框架
        send_frame = ttk.LabelFrame(self.root, text="发送设置")
        send_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(send_frame, text="立即发送:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.send_now_var = tk.BooleanVar(value=True)
        ttk.Radiobutton(send_frame, text="是", variable=self.send_now_var, value=True,
                         command=self.toggle_schedule).grid(row=0, column=1, padx=5, pady=5)
        ttk.Radiobutton(send_frame, text="否", variable=self.send_now_var, value=False,
                         command=self.toggle_schedule).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(send_frame, text="定时发送:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.schedule_frame = ttk.Frame(send_frame)
        self.schedule_frame.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        self.schedule_frame.grid_remove()  # 默认隐藏

        ttk.Label(self.schedule_frame, text="日期:").pack(side=tk.LEFT, padx=2)
        self.date_entry = ttk.Entry(self.schedule_frame, width=10)
        self.date_entry.pack(side=tk.LEFT, padx=2)
        self.date_entry.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))

        ttk.Label(self.schedule_frame, text="时间:").pack(side=tk.LEFT, padx=2)
        self.time_entry = ttk.Entry(self.schedule_frame, width=8)
        self.time_entry.pack(side=tk.LEFT, padx=2)
        self.time_entry.insert(0, datetime.datetime.now().strftime("%H:%M"))

        ttk.Label(send_frame, text="批量设置:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        batch_frame = ttk.Frame(send_frame)
        batch_frame.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="w")

        ttk.Label(batch_frame, text="每批数量:").pack(side=tk.LEFT, padx=2)
        self.batch_size_entry = ttk.Entry(batch_frame, width=5)
        self.batch_size_entry.pack(side=tk.LEFT, padx=2)
        self.batch_size_entry.insert(0, "3")  # 改为默认3封

        ttk.Label(batch_frame, text="批次间隔(秒):").pack(side=tk.LEFT, padx=2)
        self.batch_interval_entry = ttk.Entry(batch_frame, width=5)
        self.batch_interval_entry.pack(side=tk.LEFT, padx=2)
        self.batch_interval_entry.insert(0, "300")  # 改为默认300秒

        # 数据预览框架
        preview_frame = ttk.LabelFrame(self.root, text="数据预览")
        preview_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # 创建标签页
        tab_control = ttk.Notebook(preview_frame)

        # 收件人标签页
        recipient_tab = ttk.Frame(tab_control)
        tab_control.add(recipient_tab, text="收件人列表")

        # 创建Treeview组件
        self.tree = ttk.Treeview(recipient_tab, columns=("email", "subject", "content"), show="headings")
        self.tree.heading("email", text="邮箱地址")
        self.tree.heading("subject", text="邮件主题")
        self.tree.heading("content", text="邮件内容")
        self.tree.column("email", width=150)
        self.tree.column("subject", width=150)
        self.tree.column("content", width=350)

        # 添加滚动条
        scrollbar_y = ttk.Scrollbar(recipient_tab, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(recipient_tab, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        # 发件人标签页
        sender_tab = ttk.Frame(tab_control)
        tab_control.add(sender_tab, text="发件人列表")

        # 创建发件人Treeview
        self.sender_tree = ttk.Treeview(sender_tab, columns=("email", "password", "status"), show="headings")
        self.sender_tree.heading("email", text="发件人邮箱")
        self.sender_tree.heading("password", text="密码/授权码")
        self.sender_tree.heading("status", text="状态")
        self.sender_tree.column("email", width=200)
        self.sender_tree.column("password", width=200)
        self.sender_tree.column("status", width=100)

        # 发件人列表滚动条
        sender_scrollbar_y = ttk.Scrollbar(sender_tab, orient="vertical", command=self.sender_tree.yview)
        sender_scrollbar_x = ttk.Scrollbar(sender_tab, orient="horizontal", command=self.sender_tree.xview)
        self.sender_tree.configure(yscrollcommand=sender_scrollbar_y.set, xscrollcommand=sender_scrollbar_x.set)

        sender_scrollbar_y.pack(side="right", fill="y")
        sender_scrollbar_x.pack(side="bottom", fill="x")
        self.sender_tree.pack(fill="both", expand=True)

        # 显示标签页
        tab_control.pack(fill="both", expand=True)

        # 发送按钮和进度条
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=5)

        self.send_button = ttk.Button(control_frame, text="开始发送", command=self.start_sending)
        self.send_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.stop_button = ttk.Button(control_frame, text="停止发送", command=self.stop_sending, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5, pady=5)

        # 添加状态栏
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(status_frame, text="状态:").pack(side=tk.LEFT, padx=5)
        self.status_label = ttk.Label(status_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT, padx=5)

        ttk.Label(status_frame, text="进度:").pack(side=tk.LEFT, padx=5)
        self.progress_bar = ttk.Progressbar(status_frame, length=300, mode="determinate")
        self.progress_bar.pack(side=tk.LEFT, padx=5, fill="x", expand=True)

        self.progress_label = ttk.Label(status_frame, text="0/0")
        self.progress_label.pack(side=tk.LEFT, padx=5)

        # 添加当前发件人显示
        self.current_sender_label = ttk.Label(status_frame, text="")
        self.current_sender_label.pack(side=tk.RIGHT, padx=5)

    def toggle_sender_mode(self):
        if self.single_sender_var.get():
            self.single_sender_frame.grid()
            self.multi_sender_frame.grid_remove()
        else:
            self.single_sender_frame.grid_remove()
            self.multi_sender_frame.grid()

    def toggle_schedule(self):
        if self.send_now_var.get():
            self.schedule_frame.grid_remove()
        else:
            self.schedule_frame.grid()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)

    def browse_sender_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.sender_file_entry.delete(0, tk.END)
            self.sender_file_entry.insert(0, file_path)

    def load_senders(self):
        file_path = self.sender_file_entry.get()
        if not file_path:
            messagebox.showerror("错误", "请选择发件人Excel文件")
            return

        try:
            # 加载Excel文件，使用Sheet1而不是sheet1
            excel_file = pd.ExcelFile(file_path)
            # 检查是否有Sheet1
            if "Sheet1" not in excel_file.sheet_names:
                messagebox.showerror("错误", f"Excel文件中没有名为'Sheet1'的工作表。找到的工作表: {', '.join(excel_file.sheet_names)}")
                return

            sender_data = pd.read_excel(file_path, sheet_name="Sheet1")

            # 检查列名
            required_columns = ["邮箱", "密钥"]
            if not all(col in sender_data.columns for col in required_columns):
                messagebox.showerror("错误", f"发件人Excel文件必须包含以下列: {', '.join(required_columns)}")
                return

            # 处理数据，添加域名后缀
            self.sender_accounts = []
            for _, row in sender_data.iterrows():
                email = row["邮箱"]
                if not "@" in str(email):  # 如果没有@，则添加@163.com
                    email = f"{email}@163.com"
                self.sender_accounts.append({
                    "email": email,
                    "password": row["密钥"],
                    "status": "就绪",
                    "sent_count": 0
                })

            # 清空现有发件人数据
            for item in self.sender_tree.get_children():
                self.sender_tree.delete(item)

            # 添加发件人数据到treeview
            for sender in self.sender_accounts:
                self.sender_tree.insert("", "end", values=(
                    sender["email"],
                    "*" * len(str(sender["password"])),  # 密码显示为星号
                    sender["status"]
                ))

            self.current_sender_index = 0
            messagebox.showinfo("成功", f"已加载 {len(self.sender_accounts)} 个发件人账号")

        except Exception as e:
            messagebox.showerror("错误", f"加载发件人文件时出错: {str(e)}\n\n详细错误: {traceback.format_exc()}")

    def load_data(self):
        file_path = self.file_path_entry.get()
        if not file_path:
            messagebox.showerror("错误", "请选择收件人Excel文件")
            return

        try:
            # 加载Excel文件，使用Sheet1而不是sheet1
            excel_file = pd.ExcelFile(file_path)
            # 检查是否有Sheet1
            if "Sheet1" not in excel_file.sheet_names:
                messagebox.showerror("错误", f"Excel文件中没有名为'Sheet1'的工作表。找到的工作表: {', '.join(excel_file.sheet_names)}")
                return

            self.email_data = pd.read_excel(file_path, sheet_name="Sheet1")

            # 检查列名
            required_columns = ["邮箱地址", "邮件主题", "邮件内容"]
            if not all(col in self.email_data.columns for col in required_columns):
                messagebox.showerror("错误", f"收件人Excel文件必须包含以下列: {', '.join(required_columns)}")
                return

            # 清空现有数据
            for item in self.tree.get_children():
                self.tree.delete(item)

            # 添加数据到treeview
            for _, row in self.email_data.iterrows():
                self.tree.insert("", "end", values=(
                    row["邮箱地址"],
                    row["邮件主题"],
                    row["邮件内容"][:50] + "..." if len(str(row["邮件内容"])) > 50 else row["邮件内容"]
                ))

            self.status_label.config(text=f"已加载 {len(self.email_data)} 条记录")
            messagebox.showinfo("成功", f"已加载 {len(self.email_data)} 条记录")

        except Exception as e:
            messagebox.showerror("错误", f"加载收件人文件时出错: {str(e)}\n\n详细错误: {traceback.format_exc()}")

    def get_next_sender(self):
        """获取下一个可用的发件人账号"""
        if not self.sender_accounts:
            return None, None

        # 从当前索引开始查找可用账号
        original_index = self.current_sender_index
        while True:
            sender = self.sender_accounts[self.current_sender_index]

            # 移动到下一个索引，准备下次使用
            self.current_sender_index = (self.current_sender_index + 1) % len(self.sender_accounts)

            # 如果状态是"就绪"或者"已使用"，则可用
            if sender["status"] in ["就绪", "已使用"]:
                return sender["email"], sender["password"]

            # 如果已经遍历了一圈还没找到可用账号，返回None
            if self.current_sender_index == original_index:
                return None, None

    def update_sender_status(self, email, status, sent_count=None):
        """更新发件人状态"""
        if not self.sender_accounts:
            return

        for i, sender in enumerate(self.sender_accounts):
            if sender["email"] == email:
                sender["status"] = status
                if sent_count is not None:
                    sender["sent_count"] = sent_count

                # 更新UI
                for item in self.sender_tree.get_children():
                    if self.sender_tree.item(item, "values")[0] == email:
                        self.sender_tree.item(item, values=(
                            email,
                            "*" * len(str(sender["password"])),
                            f"{status} ({sender.get('sent_count', 0)}封)"
                        ))
                        break
                break

    def start_sending(self):
        if self.is_sending:
            return

        if self.email_data is None or len(self.email_data) == 0:
            messagebox.showerror("错误", "没有要发送的数据")
            return

        # 获取发件人信息
        if self.single_sender_var.get():
            # 单一发件人模式
            sender_name = self.email_entry.get()
            if not sender_name:
                messagebox.showerror("错误", "请输入发件人邮箱")
                return

            sender_email = f"{sender_name}@163.com" if not "@" in sender_name else sender_name
            sender_password = self.password_entry.get()

            if not sender_password:
                messagebox.showerror("错误", "请输入邮箱密码或授权码")
                return

            # 创建一个单一发件人账号
            self.sender_accounts = [{
                "email": sender_email,
                "password": sender_password,
                "status": "就绪",
                "sent_count": 0
            }]
            self.current_sender_index = 0
        else:
            # 多发件人模式
            if not self.sender_accounts or len(self.sender_accounts) == 0:
                messagebox.showerror("错误", "请加载发件人账号")
                return

        # 获取批量发送设置
        try:
            batch_size = int(self.batch_size_entry.get())
            batch_interval = int(self.batch_interval_entry.get())
        except ValueError:
            messagebox.showerror("错误", "批量设置必须是数字")
            return

        # 如果是定时发送
        if not self.send_now_var.get():
            try:
                schedule_time = f"{self.date_entry.get()} {self.time_entry.get()}"
                schedule_dt = datetime.datetime.strptime(schedule_time, "%Y-%m-%d %H:%M")

                now = datetime.datetime.now()
                if schedule_dt <= now:
                    messagebox.showerror("错误", "定时发送时间必须是未来时间")
                    return

                delay_seconds = (schedule_dt - now).total_seconds()

                # 设置定时任务
                self.status_label.config(text=f"已设置定时任务，将在 {schedule_time} 开始发送")
                self.is_sending = True
                self.send_button.config(state=tk.DISABLED)
                self.stop_button.config(state=tk.NORMAL)

                self.current_job = threading.Timer(
                    delay_seconds,
                    self.send_emails_in_batches,
                    args=(batch_size, batch_interval)
                )
                self.current_job.start()

            except ValueError:
                messagebox.showerror("错误", "日期格式错误，请使用YYYY-MM-DD HH:MM格式")
                return
        else:
            # 立即发送
            self.is_sending = True
            self.send_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)

            # 在新线程中发送邮件
            self.current_job = threading.Thread(
                target=self.send_emails_in_batches,
                args=(batch_size, batch_interval)
            )
            self.current_job.start()

    def stop_sending(self):
        if not self.is_sending:
            return

        self.is_sending = False

        if self.current_job is not None and isinstance(self.current_job, threading.Timer):
            self.current_job.cancel()

        self.status_label.config(text="已停止发送")
        messagebox.showinfo("停止", "邮件发送已停止")
        self.reset_ui()

    def reset_ui(self):
        self.send_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_label.config(text="就绪")
        self.current_job = None

    def send_emails_in_batches(self, batch_size, batch_interval):
        total_emails = len(self.email_data)
        sent_count = 0
        failed_count = 0
        failed_emails = []

        self.progress_bar["maximum"] = total_emails
        self.progress_bar["value"] = 0
        self.progress_label.config(text=f"0/{total_emails}")

        # 创建日志目录
        log_dir = "email_logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)

        # 创建日志文件
        log_file = os.path.join(log_dir, f"email_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        with open(log_file, "w", encoding="utf-8") as f:
            f.write(f"邮件发送日志 - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"发件人模式: {'单一发件人' if self.single_sender_var.get() else '多发件人轮询'}\n")
            if not self.single_sender_var.get():
                f.write(f"轮询方式: {'每批一个账号' if self.rotation_mode.get() == 'batch' else '每封一个账号'}\n")
            f.write("-" * 50 + "\n\n")

        batches = [self.email_data.iloc[i:i+batch_size] for i in range(0, total_emails, batch_size)]

        for batch_index, batch in enumerate(batches):
            if not self.is_sending:
                break

            self.status_label.config(text=f"正在发送第 {batch_index+1}/{len(batches)} 批")

            # 每批选择发件人 (如果是每批轮询)
            sender_email = None
            sender_password = None

            if not self.single_sender_var.get() and self.rotation_mode.get() == "batch":
                sender_email, sender_password = self.get_next_sender()
                if not sender_email:
                    with open(log_file, "a", encoding="utf-8") as f:
                        f.write(f"[错误] 批次 {batch_index+1}: 没有可用的发件人账号\n")
                    self.root.after(0, lambda: messagebox.showerror("错误", "没有可用的发件人账号"))
                    break

                self.current_sender_label.config(text=f"当前发件人: {sender_email}")

                # 更新日志
                with open(log_file, "a", encoding="utf-8") as f:
                    f.write(f"[批次 {batch_index+1}] 使用发件人: {sender_email}\n")
            elif self.single_sender_var.get():
                # 单一发件人模式
                sender_email = self.sender_accounts[0]["email"]
                sender_password = self.sender_accounts[0]["password"]
                self.current_sender_label.config(text=f"当前发件人: {sender_email}")

            # 每批重新连接SMTP服务器
            smtp = None
            connection_active = False

            # 发送批次内的邮件
            for _, row in batch.iterrows():
                if not self.is_sending:
                    break

                recipient = row["邮箱地址"]
                subject = row["邮件主题"]
                content = row["邮件内容"]

                # 验证邮箱格式
                email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                if not re.match(email_regex, str(recipient)):
                    error_msg = f"邮箱格式不正确: {recipient}"
                    with open(log_file, "a", encoding="utf-8") as f:
                        f.write(f"[失败] {recipient} - {error_msg}\n")
                    failed_emails.append((recipient, subject, error_msg))
                    failed_count += 1
                    self.progress_bar["value"] = sent_count + failed_count
                    self.progress_label.config(text=f"{sent_count}/{total_emails} (失败: {failed_count})")
                    continue

                # 每封邮件选择发件人 (如果是每封轮询)
                if not self.single_sender_var.get() and self.rotation_mode.get() == "email":
                    # 关闭上一个连接
                    if smtp and connection_active:
                        try:
                            smtp.quit()
                        except:
                            pass
                        connection_active = False

                    sender_email, sender_password = self.get_next_sender()
                    if not sender_email:
                        with open(log_file, "a", encoding="utf-8") as f:
                            f.write(f"[错误] 邮件 {sent_count+failed_count+1}: 没有可用的发件人账号\n")
                        self.root.after(0, lambda: messagebox.showerror("错误", "没有可用的发件人账号"))
                        break

                    self.current_sender_label.config(text=f"当前发件人: {sender_email}")

                    # 更新日志
                    with open(log_file, "a", encoding="utf-8") as f:
                        f.write(f"[邮件 {sent_count+failed_count+1}] 使用发件人: {sender_email}\n")

                # 如果连接不活跃，则建立新连接
                if not connection_active:
                    try:
                        smtp = smtplib.SMTP_SSL("smtp.163.com", 465)
                        smtp.login(sender_email, sender_password)
                        connection_active = True

                        # 更新发件人状态
                        for sender in self.sender_accounts:
                            if sender["email"] == sender_email:
                                self.update_sender_status(sender_email, "已使用", sender.get("sent_count", 0))
                                break
                    except Exception as e:
                        error_msg = str(e)
                        with open(log_file, "a", encoding="utf-8") as f:
                            f.write(f"[连接失败] 发件人 {sender_email} - {error_msg}\n")

                        # 标记此发件人为不可用
                        self.update_sender_status(sender_email, "连接失败")

                        # 如果是每封轮询，尝试下一个发件人
                        if not self.single_sender_var.get() and self.rotation_mode.get() == "email":
                            continue
                        else:
                            # 如果是每批轮询或单一发件人，跳过整个批次
                            break

                try:
                    # 创建邮件
                    msg = MIMEMultipart()
                    msg["From"] = sender_email
                    msg["To"] = recipient
                    msg["Subject"] = Header(str(subject), "utf-8")

                    # 添加邮件内容
                    msg.attach(MIMEText(str(content), "plain", "utf-8"))

                    # 增加随机延迟（0.5-2秒），避免太过规律的发送
                    time.sleep(0.5 + random.random() * 1.5)

                    # 发送邮件
                    smtp.sendmail(sender_email, recipient, msg.as_string())

                    sent_count += 1

                    # 更新发件人发送计数
                    for sender in self.sender_accounts:
                        if sender["email"] == sender_email:
                            sender["sent_count"] = sender.get("sent_count", 0) + 1
                            self.update_sender_status(sender_email, "已使用", sender["sent_count"])
                            break

                    self.progress_bar["value"] = sent_count + failed_count
                    self.progress_label.config(text=f"{sent_count}/{total_emails} (失败: {failed_count})")
                    self.root.update_idletasks()

                    with open(log_file, "a", encoding="utf-8") as f:
                        f.write(f"[成功] {recipient} - 已发送 (发件人: {sender_email})\n")

                except Exception as e:
                    error_msg = str(e)
                    with open(log_file, "a", encoding="utf-8") as f:
                        f.write(f"[失败] {recipient} - {error_msg} (发件人: {sender_email})\n")
                    failed_emails.append((recipient, subject, error_msg))
                    failed_count += 1

                    # 如果是连接断开或认证失败
                    if "Connection unexpectedly closed" in error_msg or "please run connect() first" in error_msg:
                        connection_active = False

                        # 标记此发件人为不可用
                        self.update_sender_status(sender_email, "连接断开")

                        with open(log_file, "a", encoding="utf-8") as f:
                            f.write(f"[发件人断开] {sender_email} - 连接被服务器关闭\n")

                        # 如果是每封轮询，将在下一封邮件尝试新发件人
                        if not self.single_sender_var.get() and self.rotation_mode.get() == "email":
                            continue
                        else:
                            # 如果是每批轮询或单一发件人，跳过此批次剩余邮件
                            break

                    self.progress_bar["value"] = sent_count + failed_count
                    self.progress_label.config(text=f"{sent_count}/{total_emails} (失败: {failed_count})")
                    self.root.update_idletasks()

            # 无论成功或失败，每批结束都关闭SMTP连接
            try:
                if smtp and connection_active:
                    smtp.quit()
            except:
                pass

            # 如果不是最后一批且还要继续发送，则等待指定的间隔时间
            if batch_index < len(batches) - 1 and self.is_sending:
                wait_time = batch_interval + random.randint(-5, 10)  # 增加随机波动
                wait_time = max(30, wait_time)  # 确保至少有30秒间隔

                for i in range(wait_time):
                    if not self.is_sending:
                        break
                    self.status_label.config(text=f"等待下一批发送... {wait_time-i} 秒")
                    time.sleep(1)

        # 记录发送结果摘要
        with open(log_file, "a", encoding="utf-8") as f:
            f.write("\n" + "-" * 50 + "\n")
            f.write(f"发送摘要:\n")
            f.write(f"总邮件数: {total_emails}\n")
            f.write(f"成功发送: {sent_count}\n")
            f.write(f"发送失败: {failed_count}\n\n")

            if failed_count > 0:
                f.write("失败详情:\n")
                for email, subj, error in failed_emails:
                    f.write(f"{email} ({subj}): {error}\n")

        if self.is_sending:  # 只有在正常完成时才显示成功消息
            result_msg = f"发送完成！\n成功: {sent_count}/{total_emails}\n失败: {failed_count}/{total_emails}"
            if failed_count > 0:
                result_msg += f"\n\n详细日志已保存到: {log_file}"
            self.root.after(0, lambda: messagebox.showinfo("发送结果", result_msg))

        self.is_sending = False
        self.root.after(0, self.reset_ui)
