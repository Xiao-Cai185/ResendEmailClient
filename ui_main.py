import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, scrolledtext, font
from config import get_api_key, set_api_key_dialog
from history import get_input_history, add_input_history, remove_input_history, clear_input_history
from utils import get_resource_path, validate_email
from email_send import email_sender
import threading
from tkcalendar import DateEntry
import os
import sys
import base64
import tzlocal
from datetime import datetime, timedelta, timezone
try:
    from zoneinfo import ZoneInfo
except ImportError:
    from pytz import timezone as ZoneInfo
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

class ResendEmailClient:
    def __init__(self):
        self.api_key = get_api_key()
        # 启动时如无API Key先弹窗输入，取消则退出
        if not self.api_key:
            set_api_key_dialog()
            self.api_key = get_api_key()
            if not self.api_key:
                sys.exit(0)
        self.input_history = get_input_history()
        self.attachments = []
        self.setup_main_window()

    def setup_main_window(self):
        if TkinterDnD and DND_FILES:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        self.root.title("Resend邮件发送客户端")
        self.root.geometry("900x700")
        try:
            self.root.iconbitmap(get_resource_path("email.ico"))
        except Exception:
            pass
        # 菜单栏
        menubar = tk.Menu(self.root)
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="API Key设置", command=self.menu_set_api_key)
        menubar.add_cascade(label="设置", menu=settings_menu)
        self.root.config(menu=menubar)
        # 主界面布局
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        # 发件人信息
        sender_frame = ttk.LabelFrame(main_frame, text="发件人信息", padding="5")
        sender_frame.grid(row=0, column=0, columnspan=2, sticky="we", pady=5)
        sender_frame.columnconfigure(1, weight=1)
        sender_frame.columnconfigure(3, weight=1)
        ttk.Label(sender_frame, text="发件用户名:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.sender_name = ttk.Combobox(sender_frame, width=20, values=self.input_history["sender_names"])
        self.sender_name.grid(row=0, column=1, sticky="we", padx=5)
        self.sender_name.bind("<FocusOut>", lambda e: None)
        self.sender_name.bind("<Return>", lambda e: None)
        self.sender_name.bind("<Button-3>", lambda e: self.show_history_menu(e, "sender_names", self.sender_name))
        ttk.Label(sender_frame, text="发件邮箱:").grid(row=0, column=2, sticky=tk.W, padx=5)
        self.sender_email = ttk.Combobox(sender_frame, width=30, values=self.input_history["sender_emails"])
        self.sender_email.grid(row=0, column=3, sticky="we", padx=5)
        self.sender_email.bind("<FocusOut>", lambda e: self.validate_email_format(self.sender_email))
        self.sender_email.bind("<Return>", lambda e: None)
        self.sender_email.bind("<Button-3>", lambda e: self.show_history_menu(e, "sender_emails", self.sender_email))
        # 收件人
        recipient_frame = ttk.LabelFrame(main_frame, text="收件人", padding="5")
        recipient_frame.grid(row=1, column=0, columnspan=2, sticky="we", pady=5)
        recipient_frame.columnconfigure(0, weight=1)
        self.recipients = []
        self.recipient_frame_inner = ttk.Frame(recipient_frame)
        self.recipient_frame_inner.grid(row=0, column=0, sticky="we")
        self.recipient_frame_inner.columnconfigure(0, weight=1)
        self.add_recipient_row()
        add_btn = ttk.Button(recipient_frame, text="+ 添加收件人", command=self.add_recipient_row)
        add_btn.grid(row=1, column=0, sticky=tk.W, pady=5)
        # 主题
        subject_frame = ttk.Frame(main_frame)
        subject_frame.grid(row=5, column=0, columnspan=2, sticky="we", pady=5)
        subject_frame.columnconfigure(1, weight=1)
        ttk.Label(subject_frame, text="邮件主题:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.subject = ttk.Entry(subject_frame)
        self.subject.grid(row=0, column=1, sticky="we", padx=5)
        # 邮件内容编辑器
        content_frame = ttk.LabelFrame(main_frame, text="邮件内容", padding="5")
        content_frame.grid(row=6, column=0, columnspan=2, sticky="wenes", pady=5)
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        # 附件区块（补充上传、移除、拖拽等功能）
        self.attachment_display = tk.StringVar(value="附件上传区：")
        attachment_frame = ttk.Frame(content_frame)
        attachment_frame.grid(row=6, column=0, columnspan=2, sticky="we", pady=(0,2))
        ttk.Label(attachment_frame, textvariable=self.attachment_display).pack(side=tk.LEFT, padx=2)
        upload_btn = ttk.Button(attachment_frame, text="上传本地小附件", command=self.upload_attachments)
        upload_btn.pack(side=tk.LEFT, padx=2)
        remote_btn = ttk.Button(attachment_frame, text="导入远程大文件链接", command=self.add_remote_attachment)
        remote_btn.pack(side=tk.LEFT, padx=2)
        ttk.Label(attachment_frame, text="【支持拖拽上传】").pack(side=tk.LEFT, padx=2)
        self.attachment_list_frame = ttk.Frame(attachment_frame)
        self.attachment_list_frame.pack(side=tk.LEFT, padx=5)
        self.update_attachment_label()
        # 格式化工具栏
        toolbar = ttk.Frame(content_frame)
        toolbar.grid(row=0, column=0, sticky="we", pady=5)
        ttk.Button(toolbar, text="粗体", command=self.make_bold).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="斜体", command=self.make_italic).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="下划线", command=self.make_underline).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="清除格式", command=self.clear_format).pack(side=tk.LEFT, padx=2)
        # 文本编辑区
        text_frame = ttk.Frame(content_frame)
        text_frame.grid(row=1, column=0, sticky="wenes", pady=5)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        self.content_text = tk.Text(
            text_frame, 
            height=10, 
            wrap=tk.WORD,
            bg='white',
            fg='#2E5F8C',
            font=('Arial', 11)
        )
        scrollbar_content = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.content_text.yview)
        self.content_text.configure(yscrollcommand=scrollbar_content.set)
        self.content_text.grid(row=0, column=0, sticky="wenes")
        scrollbar_content.grid(row=0, column=1, sticky="ns")
        # 拖拽支持（修正：必须在self.content_text创建后注册）
        if TkinterDnD and DND_FILES:
            self.content_text.drop_target_register(DND_FILES)
            self.content_text.dnd_bind('<<Drop>>', self.on_drop_files)
        # 配置文本标签样式
        self.setup_text_tags()
        # 绑定快捷键
        self.content_text.bind("<Control-b>", lambda e: self.make_bold())
        self.content_text.bind("<Control-i>", lambda e: self.make_italic())
        self.content_text.bind("<Control-u>", lambda e: self.make_underline())
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text="发送邮件", command=self.send_email).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清空内容", command=self.clear_form).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="查看历史", command=self.show_history).pack(side=tk.LEFT, padx=5)
        # 高级设置区块
        adv_frame = ttk.LabelFrame(main_frame, text="高级设置", padding="5")
        adv_frame.grid(row=2, column=0, columnspan=2, sticky="we", pady=5)
        adv_frame.columnconfigure(0, weight=1)
        # 折叠控制
        self.adv_collapsed = True
        self.adv_toggle_btn = ttk.Button(adv_frame, text="展开", width=6, command=self.toggle_adv)
        self.adv_toggle_btn.grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.adv_inner_frame = ttk.Frame(adv_frame)
        self.adv_inner_frame.grid(row=1, column=0, sticky="w")
        for i in range(3):
            self.adv_inner_frame.columnconfigure(i, weight=0)
        # 抄送
        cc_label = ttk.Label(self.adv_inner_frame, text="抄送 (CC):")
        cc_label.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.cc_emails = []
        self.cc_frame_inner = ttk.Frame(self.adv_inner_frame)
        self.cc_frame_inner.grid(row=0, column=1, sticky="w", pady=2)
        self.cc_frame_inner.columnconfigure(0, weight=0)
        self.add_cc_row()
        add_cc_btn = ttk.Button(self.adv_inner_frame, text="+ 添加抄送", command=self.add_cc_row)
        add_cc_btn.grid(row=0, column=2, sticky="w", padx=5, pady=2)
        # 密送
        bcc_label = ttk.Label(self.adv_inner_frame, text="密送 (BCC):")
        bcc_label.grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.bcc_emails = []
        self.bcc_frame_inner = ttk.Frame(self.adv_inner_frame)
        self.bcc_frame_inner.grid(row=1, column=1, sticky="w", pady=2)
        self.bcc_frame_inner.columnconfigure(0, weight=0)
        self.add_bcc_row()
        add_bcc_btn = ttk.Button(self.adv_inner_frame, text="+ 添加密送", command=self.add_bcc_row)
        add_bcc_btn.grid(row=1, column=2, sticky="w", padx=5, pady=2)
        # 回复地址
        reply_label = ttk.Label(self.adv_inner_frame, text="回复地址:")
        reply_label.grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.reply_to = ttk.Entry(self.adv_inner_frame, width=30)
        self.reply_to.grid(row=2, column=1, sticky="w", padx=5, pady=2)
        self.reply_to.bind("<FocusOut>", lambda e: self.validate_email_format(self.reply_to))
        sync_btn = ttk.Button(self.adv_inner_frame, text="同步发件邮箱", command=self.sync_reply_to)
        sync_btn.grid(row=2, column=2, padx=5, pady=2, sticky="w")
        # 延迟发送区块
        delay_label = ttk.Label(self.adv_inner_frame, text="发送方式:")
        delay_label.grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.send_type = tk.StringVar(value="immediate")
        rb1 = ttk.Radiobutton(self.adv_inner_frame, text="立即发送", variable=self.send_type, value="immediate", command=self.toggle_delay)
        rb2 = ttk.Radiobutton(self.adv_inner_frame, text="延迟发送", variable=self.send_type, value="delay", command=self.toggle_delay)
        rb1.grid(row=3, column=1, sticky="w", pady=2)
        rb2.grid(row=3, column=2, sticky="w", padx=5, pady=2)
        # 延迟发送时间控件（默认隐藏）
        self.delay_time_frame = ttk.Frame(self.adv_inner_frame)
        self.delay_time_frame.grid(row=4, column=0, columnspan=3, sticky="w", pady=2)
        self.delay_time_frame.columnconfigure(1, weight=0)
        today = datetime.now().date()
        max_day = today + timedelta(days=30)
        self.scheduled_date = DateEntry(self.delay_time_frame, mindate=today, maxdate=max_day, date_pattern='yyyy-mm-dd', width=12)
        self.scheduled_date.grid(row=0, column=0, sticky="w", padx=(5,0))
        self.scheduled_hour = ttk.Combobox(self.delay_time_frame, width=3, values=[f"{i:02d}" for i in range(24)], state="readonly")
        self.scheduled_hour.set(f"{datetime.now().hour:02d}")
        self.scheduled_hour.grid(row=0, column=1, sticky="w", padx=(5,0))
        ttk.Label(self.delay_time_frame, text=":").grid(row=0, column=2, sticky="w")
        self.scheduled_minute = ttk.Combobox(self.delay_time_frame, width=3, values=[f"{i:02d}" for i in range(60)], state="readonly")
        self.scheduled_minute.set(f"{datetime.now().minute:02d}")
        self.scheduled_minute.grid(row=0, column=3, sticky="w", padx=(5,0))
        self.timezone_options = [
            ("UTC-12", -12), ("UTC-11", -11), ("UTC-10", -10), ("UTC-9", -9), ("UTC-8", -8), ("UTC-7", -7), ("UTC-6", -6), ("UTC-5", -5), ("UTC-4", -4), ("UTC-3", -3), ("UTC-2", -2), ("UTC-1", -1),
            ("UTC", 0), ("UTC+1", 1), ("UTC+2", 2), ("UTC+3", 3), ("UTC+4", 4), ("UTC+5", 5), ("UTC+5:30", 5.5), ("UTC+6", 6), ("UTC+7", 7), ("UTC+8", 8), ("UTC+9", 9), ("UTC+9:30", 9.5), ("UTC+10", 10), ("UTC+11", 11), ("UTC+12", 12)
        ]
        self.scheduled_tz = ttk.Combobox(self.delay_time_frame, width=8, state="readonly", values=[x[0] for x in self.timezone_options])
        self.scheduled_tz.set("UTC+8")
        self.scheduled_tz.grid(row=0, column=4, sticky="w", padx=(5,0))
        ttk.Label(self.delay_time_frame, text="(30天内)").grid(row=0, column=5, sticky="w", padx=(5,0))
        self.delay_time_frame.grid_remove()
        self.adv_inner_frame.grid_remove()

    def add_recipient_row(self):
        row = len(self.recipients)
        frame = ttk.Frame(self.recipient_frame_inner)
        frame.grid(row=row, column=0, sticky="we", pady=2)
        frame.columnconfigure(0, weight=1)
        entry = ttk.Combobox(frame, values=self.input_history["recipient_emails"])
        entry.grid(row=0, column=0, sticky="we", padx=5)
        def on_focus_out(e):
            self.validate_email_format(entry)
        entry.bind("<FocusOut>", on_focus_out)
        entry.bind("<Return>", lambda e: None)
        entry.bind("<Button-3>", lambda e: self.show_history_menu(e, "recipient_emails", entry))
        if row > 0:
            del_btn = ttk.Button(frame, text="删除", command=lambda: self.remove_recipient_row(frame, entry))
            del_btn.grid(row=0, column=1, padx=5)
        self.recipients.append(entry)

    def remove_recipient_row(self, frame, entry):
        if entry in self.recipients:
            self.recipients.remove(entry)
        frame.destroy()

    def show_history_menu(self, event, key, combobox):
        menu = tk.Menu(self.root, tearoff=0)
        for item in self.input_history[key]:
            menu.add_command(label=item, command=lambda v=item: combobox.set(v))
        if self.input_history[key]:
            menu.add_separator()
            menu.add_command(label="清空全部", command=lambda: self.clear_history_and_update(key, combobox))
        menu.tk_popup(event.x_root, event.y_root)

    def clear_history_and_update(self, key, combobox):
        clear_input_history(key)
        combobox['values'] = []
        combobox.set("")

    def validate_email_format(self, entry):
        email = entry.get().strip()
        if email and not validate_email(email):
            messagebox.showerror("邮箱格式错误", f"邮箱格式不正确: {email}")
            entry.focus_set()

    def menu_set_api_key(self):
        set_api_key_dialog(parent=self.root)
        self.api_key = get_api_key()

    def clear_form(self):
        self.subject.delete(0, tk.END)
        self.content_text.delete(1.0, tk.END)
        for recipient in self.recipients:
            recipient.delete(0, tk.END)

    def run(self):
        self.root.mainloop()

    def send_email(self):
        # 校验必填
        if not self.sender_email.get():
            messagebox.showerror("错误", "请输入发件邮箱")
            return
        if not self.subject.get():
            messagebox.showerror("错误", "请输入邮件主题")
            return
        to_emails = []
        for recipient in self.recipients:
            email = recipient.get().strip()
            if email:
                to_emails.append(email)
        if not to_emails:
            messagebox.showerror("错误", "请至少输入一个收件人")
            return
        # 抄送/密送
        cc_emails = [cc.get().strip() for cc in self.cc_emails if cc.get().strip()]
        bcc_emails = [bcc.get().strip() for bcc in self.bcc_emails if bcc.get().strip()]
        # 判断是否有本地附件且选择了延迟发送
        has_local_attachment = any('content' in att for att in self.attachments)
        if has_local_attachment and self.send_type.get() == "delay":
            messagebox.showerror("错误", "带本地附件的邮件不支持延迟发送！")
            return
        # 组装参数前先准备scheduled_at
        scheduled_at = None
        if self.send_type.get() == "delay":
            try:
                date_str = self.scheduled_date.get()
                hour = int(self.scheduled_hour.get())
                minute = int(self.scheduled_minute.get())
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                scheduled_dt = dt.replace(hour=hour, minute=minute, second=0, microsecond=0)
                now = datetime.now()
                tz_label = self.scheduled_tz.get()
                tz_offset = 0
                for label, offset in self.timezone_options:
                    if label == tz_label:
                        tz_offset = offset
                        break
                if isinstance(tz_offset, float):
                    hours = int(tz_offset)
                    minutes = int((abs(tz_offset) - abs(hours)) * 60)
                    tz = timezone(timedelta(hours=hours, minutes=minutes if tz_offset>=0 else -minutes))
                else:
                    tz = timezone(timedelta(hours=tz_offset))
                scheduled_dt = scheduled_dt.replace(tzinfo=tz)
                if scheduled_dt <= now.astimezone(tz):
                    messagebox.showerror("错误", "预约时间必须晚于当前时间！")
                    return
                if scheduled_dt > (now + timedelta(days=30)).astimezone(tz):
                    messagebox.showerror("错误", "预约时间不能超过30天！")
                    return
                scheduled_at = scheduled_dt.isoformat()
            except Exception:
                messagebox.showerror("错误", "延迟发送时间设置有误！")
                return
        self.show_sending_loading()
        # 只将本地小附件作为API附件上传
        api_attachments = [att for att in self.attachments if 'content' in att]
        def thread_target():
            try:
                sender_name = self.sender_name.get()
                sender_email = self.sender_email.get()
                subject = self.subject.get()
                html = self.get_html_content()
                reply_to = self.reply_to.get().strip()
                params = email_sender.build_params(
                    sender_name, sender_email, to_emails, subject, html,
                    cc_emails=cc_emails, bcc_emails=bcc_emails,
                    reply_to=reply_to if reply_to else None,
                    attachments=api_attachments,
                    scheduled_at=scheduled_at
                )
                resp = email_sender.send_email(params)
                from history import add_email_record, add_input_history
                is_scheduled = "scheduled_at" in params
                # 本地历史附件保存所有（本地和远程）
                local_attachments = []
                for att in (self.attachments or []):
                    if 'content' in att and 'filename' in att:
                        size_kb = int(len(att['content']) * 3 / 4 / 1024)
                        local_attachments.append({
                            'filename': att['filename'],
                            'path': att.get('local_path', ''),
                            'size_kb': size_kb
                        })
                    elif 'url' in att and 'filename' in att:
                        local_attachments.append({
                            'filename': att['filename'],
                            'path': att['url'],
                            'size_kb': 0
                        })
                email_record = {
                    "id": resp.get("id"),
                    "params": params,
                    "sent_at": datetime.now().isoformat(),
                    "status": "scheduled" if is_scheduled else "delivered",
                    "attachments": local_attachments
                }
                add_email_record(email_record)
                # 发送成功后，记录所有实际用到且格式正确的邮箱
                sender_name = self.sender_name.get().strip()
                sender_email = self.sender_email.get().strip()
                if sender_name:
                    add_input_history("sender_names", sender_name)
                if sender_email and validate_email(sender_email):
                    add_input_history("sender_emails", sender_email)
                for email in params["to"]:
                    if validate_email(email):
                        add_input_history("recipient_emails", email)
                if "cc" in params:
                    for email in params["cc"]:
                        if validate_email(email):
                            add_input_history("recipient_emails", email)
                if "bcc" in params:
                    for email in params["bcc"]:
                        if validate_email(email):
                            add_input_history("recipient_emails", email)
                self.root.after(0, lambda: self.on_email_sent(resp.get("id"), is_scheduled))
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda: messagebox.showerror("发送失败", f"邮件发送失败: {err_msg}"))
            finally:
                self.root.after(0, self.hide_sending_loading)
        threading.Thread(target=thread_target, daemon=True).start()

    def on_email_sent(self, email_id, is_scheduled):
        if is_scheduled:
            messagebox.showinfo("已创建计划任务", f"计划任务已创建！\n邮件ID: {email_id}")
        else:
            messagebox.showinfo("发送成功", f"邮件发送成功！\n邮件ID: {email_id}")

    def show_sending_loading(self):
        if hasattr(self, '_sending_mask') and self._sending_mask:
            self._sending_mask.place(relx=0.5, rely=0.5, anchor="center")
            self.root.update()
            return
        self._sending_mask = tk.Label(self.root, text="Sending...", bg="#E6F3FF", fg="#2E5F8C", font=("Arial", 24), bd=2, relief="groove")
        self._sending_mask.place(relx=0.5, rely=0.5, anchor="center")
        self.root.update()
    def hide_sending_loading(self):
        if hasattr(self, '_sending_mask') and self._sending_mask:
            self._sending_mask.place_forget()
            self.root.update()

    def upload_attachments(self):
        from tkinter import filedialog
        files = filedialog.askopenfilenames(title="选择附件", filetypes=[("所有文件", "*.*")])
        if not files:
            return
        for file_path in files:
            from utils import is_blacklisted_attachment
            ext = os.path.splitext(file_path)[1].lower()
            if is_blacklisted_attachment(file_path):
                messagebox.showerror("不支持的附件类型", f"{file_path} 为不支持的类型，禁止上传！")
                continue
            size_kb = int(os.path.getsize(file_path) / 1024)
            filename = os.path.basename(file_path)
            with open(file_path, "rb") as f:
                content = f.read()
            import base64
            b64_content = base64.b64encode(content).decode()
            self.attachments.append({"content": b64_content, "filename": filename, "local_path": os.path.abspath(file_path), "size_kb": size_kb})
        self.update_attachment_label()

    def on_drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            from utils import is_blacklisted_attachment
            ext = os.path.splitext(file_path)[1].lower()
            if is_blacklisted_attachment(file_path):
                messagebox.showerror("不支持的附件类型", f"{file_path} 为不支持的类型，禁止上传！")
                continue
            size_kb = int(os.path.getsize(file_path) / 1024)
            filename = os.path.basename(file_path)
            with open(file_path, "rb") as f:
                content = f.read()
            import base64
            b64_content = base64.b64encode(content).decode()
            self.attachments.append({"content": b64_content, "filename": filename, "local_path": os.path.abspath(file_path), "size_kb": size_kb})
        self.update_attachment_label()

    def add_remote_attachment(self):
        url = simpledialog.askstring("远程附件链接", "请输入远程附件URL：")
        if not url:
            return
        filename = simpledialog.askstring("文件名", "请输入附件文件名（含扩展名）：")
        if not filename:
            return
        from utils import is_blacklisted_attachment
        if is_blacklisted_attachment(filename):
            messagebox.showerror("不支持的附件类型", f"{filename} 为不支持的类型，禁止上传！")
            return
        # 只记录，不作为API附件上传
        self.attachments.append({"filename": filename, "url": url})
        self.update_attachment_label()

    def update_attachment_label(self):
        for widget in self.attachment_list_frame.winfo_children():
            widget.destroy()
        if not self.attachments:
            ttk.Label(self.attachment_list_frame, text="无附件").pack(side=tk.LEFT)
        else:
            for i, att in enumerate(self.attachments):
                name = att.get("filename", "")
                size_kb = att.get("size_kb", 0)
                lbl = ttk.Label(self.attachment_list_frame, text=f"{name} ({size_kb}kb)")
                lbl.pack(side=tk.LEFT, padx=2)
                btn = ttk.Button(self.attachment_list_frame, text="×", width=2, command=(lambda idx=i: lambda: self.remove_attachment(idx))())
                btn.pack(side=tk.LEFT, padx=1)

    def remove_attachment(self, idx):
        del self.attachments[idx]
        self.update_attachment_label()

    def get_html_content(self):
        content = self.content_text.get(1.0, tk.END)
        html_content = ""
        index = "1.0"
        while True:
            next_index = self.content_text.index(f"{index}+1c")
            if self.content_text.compare(index, ">=", tk.END):
                break
            char = self.content_text.get(index, next_index)
            tags = self.content_text.tag_names(index)
            open_tags = []
            close_tags = []
            if "bold_italic" in tags:
                open_tags.append("<b><i>")
                close_tags.insert(0, "</i></b>")
            elif "bold" in tags:
                open_tags.append("<b>")
                close_tags.insert(0, "</b>")
            elif "italic" in tags:
                open_tags.append("<i>")
                close_tags.insert(0, "</i>")
            if "underline" in tags:
                open_tags.append("<u>")
                close_tags.insert(0, "</u>")
            next_tags = self.content_text.tag_names(next_index) if next_index != tk.END else []
            html_content += "".join(open_tags) + char
            if set(tags) != set(next_tags):
                html_content += "".join(close_tags)
            index = next_index
        html_content = html_content.replace('\n', '<br>\n')
        # 追加远程大文件超链接
        remote_links = []
        for att in self.attachments:
            if 'url' in att:
                remote_links.append(f'<a href="{att["url"]}">{att["filename"]}（下载）</a>')
        if remote_links:
            html_content += '<br><br>' + '<br>'.join(remote_links)
        return html_content

    def add_cc_row(self):
        row = len(self.cc_emails)
        frame = ttk.Frame(self.cc_frame_inner)
        frame.grid(row=row, column=0, sticky="we", pady=2)
        frame.columnconfigure(0, weight=1)
        entry = ttk.Entry(frame)
        entry.grid(row=0, column=0, sticky="we", padx=5)
        entry.bind("<FocusOut>", lambda e: self.validate_email_format(entry))
        if row > 0:
            del_btn = ttk.Button(frame, text="删除", command=lambda: self.remove_cc_row(frame, entry))
            del_btn.grid(row=0, column=1, padx=5)
        self.cc_emails.append(entry)

    def remove_cc_row(self, frame, entry):
        if entry in self.cc_emails:
            self.cc_emails.remove(entry)
        frame.destroy()

    def add_bcc_row(self):
        row = len(self.bcc_emails)
        frame = ttk.Frame(self.bcc_frame_inner)
        frame.grid(row=row, column=0, sticky="we", pady=2)
        frame.columnconfigure(0, weight=1)
        entry = ttk.Entry(frame)
        entry.grid(row=0, column=0, sticky="we", padx=5)
        entry.bind("<FocusOut>", lambda e: self.validate_email_format(entry))
        if row > 0:
            del_btn = ttk.Button(frame, text="删除", command=lambda: self.remove_bcc_row(frame, entry))
            del_btn.grid(row=0, column=1, padx=5)
        self.bcc_emails.append(entry)

    def remove_bcc_row(self, frame, entry):
        if entry in self.bcc_emails:
            self.bcc_emails.remove(entry)
        frame.destroy()

    def sync_reply_to(self):
        self.reply_to.delete(0, tk.END)
        self.reply_to.insert(0, self.sender_email.get())

    def toggle_adv(self):
        if self.adv_collapsed:
            self.adv_inner_frame.grid()
            self.adv_toggle_btn.config(text="收起")
            self.adv_collapsed = False
        else:
            self.adv_inner_frame.grid_remove()
            self.adv_toggle_btn.config(text="展开")
            self.adv_collapsed = True

    def toggle_delay(self):
        if self.send_type.get() == "delay":
            self.delay_time_frame.grid()
        else:
            self.delay_time_frame.grid_remove()

    def show_history(self):
        # 只允许开一个历史窗口，已开时激活
        if hasattr(self, "history_window") and self.history_window and self.history_window.winfo_exists():
            self.history_window.lift()
            self.history_window.focus_force()
            return
        from ui_history import HistoryUI
        self.history_window = HistoryUI(self.root)

    def setup_text_tags(self):
        default_font = font.Font(family="Arial", size=11)
        bold_font = font.Font(family="Arial", size=11, weight="bold")
        italic_font = font.Font(family="Arial", size=11, slant="italic")
        bold_italic_font = font.Font(family="Arial", size=11, weight="bold", slant="italic")
        underline_font = font.Font(family="Arial", size=11, underline=True)
        self.content_text.tag_configure("bold", font=bold_font)
        self.content_text.tag_configure("italic", font=italic_font)
        self.content_text.tag_configure("bold_italic", font=bold_italic_font)
        self.content_text.tag_configure("underline", font=underline_font)
        self.content_text.tag_configure("link", foreground="blue", underline=True)

    def make_bold(self):
        self.apply_format("bold")
    def make_italic(self):
        self.apply_format("italic")
    def make_underline(self):
        self.apply_format("underline")
    def apply_format(self, format_type):
        try:
            start = self.content_text.index(tk.SEL_FIRST)
            end = self.content_text.index(tk.SEL_LAST)
            current_tags = self.content_text.tag_names(start)
            if format_type in current_tags:
                self.content_text.tag_remove(format_type, start, end)
            else:
                self.content_text.tag_add(format_type, start, end)
            tags = self.content_text.tag_names(start)
            if "bold" in tags and "italic" in tags:
                self.content_text.tag_remove("bold", start, end)
                self.content_text.tag_remove("italic", start, end)
                self.content_text.tag_add("bold_italic", start, end)
            elif "bold_italic" in tags:
                if format_type == "bold":
                    self.content_text.tag_remove("bold_italic", start, end)
                    self.content_text.tag_add("italic", start, end)
                elif format_type == "italic":
                    self.content_text.tag_remove("bold_italic", start, end)
                    self.content_text.tag_add("bold", start, end)
        except tk.TclError:
            pass
    def clear_format(self):
        try:
            start = self.content_text.index(tk.SEL_FIRST)
            end = self.content_text.index(tk.SEL_LAST)
            for tag in ["bold", "italic", "underline", "bold_italic", "link"]:
                self.content_text.tag_remove(tag, start, end)
        except tk.TclError:
            pass 