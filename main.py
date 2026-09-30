import tkinter as tk
import threading
import sys
import os
import time

from tkcalendar import DateEntry
from tkinter import ttk, messagebox
from method import check_email_inbox, get_email_folders
from datetime import datetime, date, time as datetime_time


EMAIL_SERVER_PRESETS = {
    "QQ邮箱": "imap.qq.com",
    "163邮箱": "imap.163.com",
    "Gmail": "imap.gmail.com",
    "自定义": "",
}


class EmailMonitorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("邮件监控系统")
        self.root.geometry("458x570")
        self.monitoring = False
        # 设置窗口图标
        if getattr(sys, 'frozen', False):
            # 打包后的程序：优先找 PyInstaller 数据目录，其次找 exe 同级目录。
            icon_candidates = [
                os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(sys.executable)), "icon.ico"),
                os.path.join(os.path.dirname(sys.executable), "icon.ico"),
            ]
        else:
            # 开发环境
            icon_candidates = [os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")]
        for icon_path in icon_candidates:
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                break

        # 初始化缓存文件
        self.processed_emails_path = self.init_processed_emails_file()
        # 创建配置框架
        self.create_config_frame()
        # 创建控制按钮
        self.create_control_buttons()
        # 创建日志文本框
        self.create_log_area()

    def init_processed_emails_file(self):
        """初始化已处理邮件记录文件"""
        import os
        import json

        # 获取程序运行路径
        if getattr(sys, 'frozen', False):
            # 如果是打包后的程序
            application_path = os.path.dirname(sys.executable)
        else:
            # 如果是开发环境
            application_path = os.path.dirname(os.path.abspath(__file__))

        processed_emails_path = os.path.join(application_path, 'processed_emails.json')

        # 如果文件不存在，创建空的JSON文件
        if not os.path.exists(processed_emails_path):
            with open(processed_emails_path, 'w', encoding='utf-8') as f:
                json.dump({}, f)

        return processed_emails_path

    def create_config_frame(self):
        config_frame = ttk.LabelFrame(self.root, text="配置信息", padding="10")
        config_frame.pack(fill="x", padx=10, pady=5)

        # 邮箱类型预设
        ttk.Label(config_frame, text="邮箱类型:").grid(row=0, column=0, sticky="w")
        self.provider_var = tk.StringVar(value="QQ邮箱")
        self.provider_combobox = ttk.Combobox(
            config_frame,
            textvariable=self.provider_var,
            values=list(EMAIL_SERVER_PRESETS.keys()),
            state="readonly",
            width=12
        )
        self.provider_combobox.grid(row=0, column=1, sticky="w")
        self.provider_combobox.bind("<<ComboboxSelected>>", self.on_provider_changed)

        # 邮箱服务器
        ttk.Label(config_frame, text="邮箱服务器:").grid(row=1, column=0, sticky="w")
        self.server_var = tk.StringVar(value="imap.qq.com")
        ttk.Entry(config_frame, textvariable=self.server_var, width=34).grid(row=1, column=1, sticky="w")

        # 邮箱账号
        ttk.Label(config_frame, text="邮箱账号:").grid(row=2, column=0, sticky="w")
        self.username_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.username_var, width=34).grid(row=2, column=1, sticky="w")

        # 邮箱密码
        ttk.Label(config_frame, text="邮箱密码:").grid(row=3, column=0, sticky="w")
        self.password_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.password_var, show="*", width=34).grid(row=3, column=1, sticky="w")

        # 添加登录按钮
        login_button = ttk.Button(config_frame, text="登录", command=self.login)
        login_button.grid(row=3, column=2, padx=(6, 0), sticky="ew")

        # 添加收件箱选择
        ttk.Label(config_frame, text="选择收件箱:").grid(row=4, column=0, sticky="w")
        self.folder_var = tk.StringVar(value="INBOX")
        self.folder_combobox = ttk.Combobox(config_frame, textvariable=self.folder_var, state="disabled", width=32)
        self.folder_combobox.grid(row=4, column=1, sticky="w")

        # 开始日期时间
        ttk.Label(config_frame, text="开始日期时间:").grid(row=8, column=0, sticky="w")
        datetime_frame = ttk.Frame(config_frame)
        datetime_frame.grid(row=8, column=1, columnspan=2, sticky="w")
        self.start_date = DateEntry(datetime_frame, width=12,
                                    background='darkblue', foreground='white',
                                    borderwidth=2, locale='zh_CN',
                                    date_pattern='yyyy-mm-dd')
        self.start_date.grid(row=0, column=0, sticky="w")
        self.start_hour_var = tk.StringVar(value="17")
        self.start_minute_var = tk.StringVar(value="00")
        tk.Spinbox(
            datetime_frame,
            from_=0,
            to=23,
            wrap=True,
            width=3,
            format="%02.0f",
            textvariable=self.start_hour_var
        ).grid(row=0, column=1, padx=(8, 0), sticky="w")
        ttk.Label(datetime_frame, text=":").grid(row=0, column=2, sticky="w")
        tk.Spinbox(
            datetime_frame,
            from_=0,
            to=59,
            wrap=True,
            width=3,
            format="%02.0f",
            textvariable=self.start_minute_var
        ).grid(row=0, column=3, sticky="w")

        # 关键词
        ttk.Label(config_frame, text="搜索关键词:").grid(row=5, column=0, sticky="w")
        self.keyword_var = tk.StringVar(value="发票")
        ttk.Entry(config_frame, textvariable=self.keyword_var, width=34).grid(row=5, column=1, sticky="w")

        ttk.Label(config_frame, text="检查间隔(秒):").grid(row=6, column=0, sticky="w")
        self.interval_var = tk.StringVar(value="300")
        ttk.Entry(config_frame, textvariable=self.interval_var, width=34).grid(row=6, column=1, sticky="w")

        ttk.Label(config_frame, text="下载路径:").grid(row=7, column=0, sticky="w")
        self.save_dir_var = tk.StringVar(value="attachments")
        save_dir_entry = ttk.Entry(config_frame, textvariable=self.save_dir_var, width=34)
        save_dir_entry.grid(row=7, column=1, sticky="w")

        browse_button = ttk.Button(config_frame, text="浏览", command=self.browse_directory)
        browse_button.grid(row=7, column=2, padx=(6, 0), sticky="ew")

        config_frame.columnconfigure(1, weight=0)
        config_frame.columnconfigure(2, weight=1)

    def on_provider_changed(self, event=None):
        """根据邮箱类型自动填充 IMAP 服务器；自定义时保留用户手工输入。"""
        server = EMAIL_SERVER_PRESETS.get(self.provider_var.get(), "")
        if server:
            self.server_var.set(server)

    def login(self):
        """登录邮箱并获取文件夹列表"""
        try:
            server = self.server_var.get()
            username = self.username_var.get()
            password = self.password_var.get()

            if not all([server, username, password]):
                messagebox.showerror("错误", "请填写服务器、账号和密码")
                return

            folders = get_email_folders(
                imap_server=server,
                username=username,
                password=password
            )

            if folders:
                self.folder_combobox['values'] = folders
                self.folder_combobox['state'] = 'readonly'
                self.folder_var.set(folders[0])  # 设置默认值
                self.log("登录成功，已获取收件箱列表")
            else:
                messagebox.showwarning("警告", "未找到可用的收件箱")

        except Exception as e:
            messagebox.showerror("错误", f"登录失败: {str(e)}")

    def browse_directory(self):
        """选择下载文件保存路径"""
        from tkinter import filedialog
        directory = filedialog.askdirectory()
        if directory:
            self.save_dir_var.set(directory)

    def get_start_datetime(self):
        """读取界面上的开始日期时间。"""
        start_date = self.start_date.get_date()
        hour = int(self.start_hour_var.get())
        minute = int(self.start_minute_var.get())
        if not (0 <= hour <= 23 and 0 <= minute <= 59):
            raise ValueError("开始时间超出范围")
        return datetime.combine(start_date, datetime_time(hour, minute))

    def create_control_buttons(self):
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=5)

        self.start_button = ttk.Button(button_frame, text="开始监控", command=self.start_monitoring)
        self.start_button.pack(side="left", padx=(6, 8))

        self.stop_button = ttk.Button(button_frame, text="停止监控", command=self.stop_monitoring, state="disabled")
        self.stop_button.pack(side="left", padx=0)

    def create_log_area(self):
        # 添加倒计时标签
        self.countdown_label = ttk.Label(self.root, text="等待开始监控...")
        self.countdown_label.pack(padx=10, pady=5)

        log_frame = ttk.LabelFrame(self.root, text="运行日志", padding="10")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=9, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)

    def start_monitoring(self):
        try:
            self.interval = int(self.interval_var.get())
            if not all([self.server_var.get(), self.username_var.get(),
                        self.password_var.get(), self.keyword_var.get()]):
                messagebox.showerror("错误", "请填写所有配置信息")
                return

            # 验证开始日期时间
            try:
                start_datetime = self.get_start_datetime()
                if start_datetime.date() > date.today():
                    messagebox.showerror("错误", "开始日期不能晚于今天")
                    return
            except ValueError:
                messagebox.showerror("错误", "开始日期时间错误，请检查小时和分钟")
                return

            self.monitoring = True
            self.start_button.config(state="disabled")
            self.stop_button.config(state="normal")

            # 启动监控线程
            self.monitor_thread = threading.Thread(
                target=self.run_monitor,
                daemon=True
            )
            self.monitor_thread.start()

            self.log("开始监控邮箱...")
            self.countdown_label.config(text="检查中...")

        except ValueError:
            messagebox.showerror("错误", "检查间隔必须是数字")

    def stop_monitoring(self):
        self.monitoring = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.countdown_label.config(text="下次检查: 等待开始监控...")
        self.log("停止监控")

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def show_download_summary(self, count, file_list):
        """下载完成后用带滚动条的窗口展示完整文件列表。"""
        unique_files = list(dict.fromkeys(file_list))
        self.log("下载文件列表：")
        for filepath in unique_files:
            self.log(f" - {filepath}")

        save_dir = os.path.abspath(self.save_dir_var.get())
        summary = (
            f"找到 {count} 封相关邮件。\n"
            f"成功处理文件数：{len(unique_files)}\n"
            f"保存目录：{save_dir}"
        )

        result_window = tk.Toplevel(self.root)
        result_window.title("下载完成")
        result_window.geometry("760x420")
        result_window.transient(self.root)

        summary_label = ttk.Label(result_window, text=summary, justify="left")
        summary_label.pack(fill="x", padx=10, pady=(10, 6))

        list_frame = ttk.LabelFrame(result_window, text="下载文件列表（双击打开所在文件夹）", padding="6")
        list_frame.pack(fill="both", expand=True, padx=10, pady=6)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        file_listbox = tk.Listbox(list_frame, height=12, yscrollcommand=scrollbar.set)
        scrollbar.config(command=file_listbox.yview)
        file_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        numbered_files = [f"{index:02d}. {filepath}" for index, filepath in enumerate(unique_files, start=1)]
        for display_text in numbered_files:
            file_listbox.insert("end", display_text)

        def open_selected_folder(event=None):
            selection = file_listbox.curselection()
            if not selection:
                return
            filepath = unique_files[selection[0]]
            if os.path.exists(filepath):
                os.startfile(os.path.dirname(filepath))
            else:
                messagebox.showwarning("文件不存在", f"文件不存在或已被移动：\n{filepath}")

        file_listbox.bind("<Double-Button-1>", open_selected_folder)

        button_frame = ttk.Frame(result_window)
        button_frame.pack(fill="x", padx=10, pady=(0, 10))

        def copy_all():
            self.root.clipboard_clear()
            self.root.clipboard_append("\n".join(numbered_files))

        ttk.Button(button_frame, text="复制全部路径", command=copy_all).pack(side="left")
        ttk.Button(button_frame, text="打开所选文件夹", command=open_selected_folder).pack(side="left", padx=8)
        ttk.Button(button_frame, text="确定", command=result_window.destroy).pack(side="right")

    def run_monitor(self):
        while self.monitoring:
            try:
                # 获取界面选择的开始日期时间
                start_datetime = self.get_start_datetime()
                print(f"开始时间: {start_datetime}")

                count, file_list = check_email_inbox(
                    subject_keyword=self.keyword_var.get(),
                    save_dir=self.save_dir_var.get(),
                    imap_server=self.server_var.get(),
                    username=self.username_var.get(),
                    password=self.password_var.get(),
                    folder=self.folder_var.get(),
                    today=start_datetime
                )
                self.log(f"检查完成")
                if count > 0:
                    self.log(f"找到 {count} 封相关邮件，时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    self.show_download_summary(count, file_list)
                elif count == 0 and file_list:
                    # 弹窗显示下载的文件
                    messagebox.showinfo("发生错误", f"错误原因{file_list[0]}")

                # 间隔时间
                # 开始倒计时
                start_time = time.time()
                interval = int(self.interval_var.get())
                while self.monitoring and (time.time() - start_time) < interval:
                    remaining = interval - int(time.time() - start_time)
                    minutes = remaining // 60
                    seconds = remaining % 60
                    countdown_text = f"下次检查: {minutes:02d}:{seconds:02d}"
                    # 使用 after 方法安全地更新GUI
                    self.root.after(0, lambda t=countdown_text: self.countdown_label.config(text=t))
                    time.sleep(1)

                if self.monitoring:
                    self.log("开始监控邮箱...")
                    countdown_text = "检查中..."
                    self.root.after(0, lambda t=countdown_text: self.countdown_label.config(text=t))
            except Exception as e:
                self.log(f"发生错误: {str(e)}")
                break


if __name__ == '__main__':
    root = tk.Tk()
    app = EmailMonitorGUI(root)
    root.mainloop()

