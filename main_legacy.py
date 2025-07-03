import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, scrolledtext, font
import resend
import json
import os
from datetime import datetime, timedelta, timezone
import threading
import uuid
from tkcalendar import DateEntry
import sys
import re
try:
    from zoneinfo import ZoneInfo
except ImportError:
    from pytz import timezone as ZoneInfo
import tzlocal
import base64
import tkinter.filedialog as filedialog
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

CONFIG_FILE = "config.json"
HISTORY_FILE = "history.json"

EMAIL_REGEX = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")

class ResendEmailClient:
    def __init__(self):
        self.api_key = ""
        self.email_history = []
        self.load_history()
        self.load_config()  # 新增：加载API Key
        self.load_input_history()  # 新增：加载输入历史
        
        # 附件黑名单，禁止上传的后缀
        self.attachment_blacklist = {
            '.adp','.app','.asp','.bas','.bat','.cer','.chm','.cmd','.com','.cpl','.crt','.csh','.der','.exe','.fxp','.gadget','.hlp','.hta','.inf','.ins','.isp','.its','.js','.jse','.ksh','.lib','.lnk','.mad','.maf','.mag','.mam','.maq','.mar','.mas','.mat','.mau','.mav','.maw','.mda','.mdb','.mde','.mdt','.mdw','.mdz','.msc','.msh','.msh1','.msh2','.mshxml','.msh1xml','.msh2xml','.msi','.msp','.mst','.ops','.pcd','.pif','.plg','.prf','.prg','.reg','.scf','.scr','.sct','.shb','.shs','.sys','.ps1','.ps1xml','.ps2','.ps2xml','.psc1','.psc2','.tmp','.url','.vb','.vbe','.vbs','.vps','.vsmacros','.vss','.vst','.vsw','.vxd','.ws','.wsc','.wsf','.wsh','.xnk'
        }
        
        # 启动时要求输入API Key（如果本地没有）
        if not self.api_key:
            self.setup_api_key()
        else:
            resend.api_key = self.api_key
        if not self.api_key:
            return
        self.setup_main_window()
        
    def load_config(self):
        """加载本地配置文件（API Key）"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    self.api_key = config.get("api_key", "")
            except Exception:
                self.api_key = ""
        else:
            self.api_key = ""

    def save_config(self):
        """保存API Key到本地配置文件"""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({"api_key": self.api_key}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def load_input_history(self):
        """加载输入历史"""
        self.input_history = {
            "sender_names": [],
            "sender_emails": [],
            "recipient_emails": []
        }
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    self.input_history = json.load(f)
            except Exception:
                pass

    def save_input_history(self):
        """保存输入历史"""
        try:
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self.input_history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def add_to_history(self, key, value):
        """添加历史项，key为sender_names/sender_emails/recipient_emails"""
        if value and value not in self.input_history[key]:
            self.input_history[key].insert(0, value)
            # 最多保存20条
            self.input_history[key] = self.input_history[key][:20]
            self.save_input_history()

    def remove_from_history(self, key, value):
        if value in self.input_history[key]:
            self.input_history[key].remove(value)
            self.save_input_history()

    def clear_history(self, key):
        self.input_history[key] = []
        self.save_input_history()
        
    def setup_api_key(self):
        """设置API Key"""
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        try:
            root.iconbitmap("email.ico")
        except Exception:
            pass
        root.title("Resend邮件客户端")
        api_key = simpledialog.askstring(
            "API Key设置", 
            "请输入您的Resend API Key:",
            show='*',
            parent=root
        )
        if api_key:
            self.api_key = api_key
            resend.api_key = api_key
            self.save_config()  # 保存到本地
        root.destroy()
        
    def menu_set_api_key(self):
        """菜单栏API Key设置"""
        api_key = simpledialog.askstring(
            "API Key设置", 
            "请输入新的Resend API Key:",
            show='*',
            parent=self.root
        )
        if api_key:
            self.api_key = api_key
            resend.api_key = api_key
            self.save_config()
            messagebox.showinfo("成功", "API Key已更新并保存！", parent=self.root)

    def setup_main_window(self):
        """创建主窗口"""
        # 修复拖拽：用TkinterDnD.Tk()替代tk.Tk()（如可用）
        if TkinterDnD and DND_FILES:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        self.root.title("Resend邮件发送客户端")
        self.root.geometry("900x700")
        self.root.configure(bg='#E6F3FF')  # 天蓝色背景
        try:
            self.root.iconbitmap("email.ico")
        except Exception:
            pass
        # 菜单栏
        menubar = tk.Menu(self.root)
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="API Key设置", command=self.menu_set_api_key)
        menubar.add_cascade(label="设置", menu=settings_menu)
        self.root.config(menu=menubar)
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', background='#E6F3FF', foreground='#2E5F8C')
        style.configure('TFrame', background='#E6F3FF')
        style.configure('TButton', background='#87CEEB', foreground='#2E5F8C')
        style.map('TButton', background=[('active', '#B0E0E6')])
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 发件人信息区域
        sender_frame = ttk.LabelFrame(main_frame, text="发件人信息", padding="5")
        sender_frame.grid(row=0, column=0, columnspan=2, sticky="we", pady=5)
        sender_frame.columnconfigure(1, weight=1)
        sender_frame.columnconfigure(3, weight=1)
        
        ttk.Label(sender_frame, text="发件用户名:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.sender_name = ttk.Combobox(sender_frame, width=20, values=self.input_history["sender_names"])
        self.sender_name.grid(row=0, column=1, sticky="we", padx=5)
        self.sender_name.bind("<FocusOut>", lambda e: self.add_to_history("sender_names", self.sender_name.get().strip()))
        self.sender_name.bind("<Return>", lambda e: self.add_to_history("sender_names", self.sender_name.get().strip()))
        self.sender_name.bind("<Button-3>", lambda e: self.show_history_menu(e, "sender_names", self.sender_name))
        
        ttk.Label(sender_frame, text="发件邮箱:").grid(row=0, column=2, sticky=tk.W, padx=5)
        self.sender_email = ttk.Combobox(sender_frame, width=30, values=self.input_history["sender_emails"])
        self.sender_email.grid(row=0, column=3, sticky="we", padx=5)
        self.sender_email.bind("<FocusOut>", lambda e: self.validate_email_format(self.sender_email))
        self.sender_email.bind("<Return>", lambda e: self.add_to_history("sender_emails", self.sender_email.get().strip()))
        self.sender_email.bind("<Button-3>", lambda e: self.show_history_menu(e, "sender_emails", self.sender_email))
        
        # 收件人区域
        recipient_frame = ttk.LabelFrame(main_frame, text="收件人", padding="5")
        recipient_frame.grid(row=1, column=0, columnspan=2, sticky="we", pady=5)
        recipient_frame.columnconfigure(0, weight=1)
        
        # 收件人列表
        self.recipients = []
        self.recipient_frame_inner = ttk.Frame(recipient_frame)
        self.recipient_frame_inner.grid(row=0, column=0, sticky="we")
        self.recipient_frame_inner.columnconfigure(0, weight=1)
        
        self.add_recipient_row()
        
        add_btn = ttk.Button(recipient_frame, text="+ 添加收件人", command=self.add_recipient_row)
        add_btn.grid(row=1, column=0, sticky=tk.W, pady=5)
        
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
        
        # 附件区块（放在内容编辑器上方）
        self.attachments = []
        self.attachment_display = tk.StringVar(value="无附件")
        attachment_frame = ttk.Frame(content_frame)
        attachment_frame.grid(row=6, column=0, columnspan=2, sticky="we", pady=(0,2))

        upload_btn = ttk.Button(attachment_frame, text="上传本地小附件", command=self.upload_attachments)
        upload_btn.pack(side=tk.LEFT, padx=2)
        remote_btn = ttk.Button(attachment_frame, text="导入远程大文件链接", command=self.add_remote_attachment)
        remote_btn.pack(side=tk.LEFT, padx=2)

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
        # 拖拽支持
        if TkinterDnD and DND_FILES:
            self.content_text.drop_target_register(DND_FILES)
            self.content_text.dnd_bind('<<Drop>>', self.on_drop_files)
        else:
            def show_dnd_tip():
                messagebox.showinfo("提示", "如需拖拽上传附件，请先安装tkinterDnD2库。\n可用pip install tkinterdnd2")
            self.content_text.bind('<Button-3>', lambda e: show_dnd_tip())
        
        # 添加滚动条
        scrollbar_content = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.content_text.yview)
        self.content_text.configure(yscrollcommand=scrollbar_content.set)
        
        self.content_text.grid(row=0, column=0, sticky="wenes")
        scrollbar_content.grid(row=0, column=1, sticky="ns")
        
        # 配置文本标签样式
        self.setup_text_tags()
        
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text="发送邮件", command=self.send_email).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清空内容", command=self.clear_form).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="查看历史", command=self.show_history).pack(side=tk.LEFT, padx=5)
        
    def add_recipient_row(self):
        """添加收件人输入行"""
        row = len(self.recipients)
        frame = ttk.Frame(self.recipient_frame_inner)
        frame.grid(row=row, column=0, sticky="we", pady=2)
        frame.columnconfigure(0, weight=1)
        entry = ttk.Combobox(frame, values=self.input_history["recipient_emails"])
        entry.grid(row=0, column=0, sticky="we", padx=5)
        def on_focus_out(e):
            self.validate_email_format(entry)
            self.add_to_history("recipient_emails", entry.get().strip())
        entry.bind("<FocusOut>", on_focus_out)
        entry.bind("<Return>", lambda e: self.add_to_history("recipient_emails", entry.get().strip()))
        entry.bind("<Button-3>", lambda e: self.show_history_menu(e, "recipient_emails", entry))
        if row > 0:
            del_btn = ttk.Button(frame, text="删除", command=lambda: self.remove_recipient_row(frame, entry))
            del_btn.grid(row=0, column=1, padx=5)
        self.recipients.append(entry)
        
    def add_cc_row(self):
        """添加抄送输入行"""
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
        """删除抄送行"""
        if entry in self.cc_emails:
            self.cc_emails.remove(entry)
        frame.destroy()
        
    def add_bcc_row(self):
        """添加密送输入行"""
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
        """删除密送行"""
        if entry in self.bcc_emails:
            self.bcc_emails.remove(entry)
        frame.destroy()
        
    def setup_text_tags(self):
        """设置文本标签样式"""
        # 创建字体对象
        default_font = font.Font(family="Arial", size=11)
        bold_font = font.Font(family="Arial", size=11, weight="bold")
        italic_font = font.Font(family="Arial", size=11, slant="italic")
        bold_italic_font = font.Font(family="Arial", size=11, weight="bold", slant="italic")
        underline_font = font.Font(family="Arial", size=11, underline=True)
        
        # 配置标签
        self.content_text.tag_configure("bold", font=bold_font)
        self.content_text.tag_configure("italic", font=italic_font)
        self.content_text.tag_configure("bold_italic", font=bold_italic_font)
        self.content_text.tag_configure("underline", font=underline_font)
        self.content_text.tag_configure("link", foreground="blue", underline=True)
        
    def remove_recipient_row(self, frame, entry):
        """删除收件人行"""
        if entry in self.recipients:
            self.recipients.remove(entry)
        frame.destroy()
        
        # 绑定键盘快捷键
        self.content_text.bind("<Control-b>", lambda e: self.make_bold())
        self.content_text.bind("<Control-i>", lambda e: self.make_italic())
        self.content_text.bind("<Control-u>", lambda e: self.make_underline())
        
    def sync_reply_to(self):
        """同步回复地址为发件邮箱"""
        self.reply_to.delete(0, tk.END)
        self.reply_to.insert(0, self.sender_email.get())
        
    def make_bold(self):
        """粗体格式"""
        self.apply_format("bold")
        
    def make_italic(self):
        """斜体格式"""
        self.apply_format("italic")
        
    def make_underline(self):
        """下划线格式"""
        self.apply_format("underline")
        
    def apply_format(self, format_type):
        """应用格式"""
        try:
            # 获取选中的文本范围
            start = self.content_text.index(tk.SEL_FIRST)
            end = self.content_text.index(tk.SEL_LAST)
            
            # 获取当前的标签
            current_tags = self.content_text.tag_names(start)
            
            # 切换格式
            if format_type in current_tags:
                # 如果已经有这个格式，则移除
                self.content_text.tag_remove(format_type, start, end)
            else:
                # 添加格式
                self.content_text.tag_add(format_type, start, end)
                
            # 处理组合格式 (粗体+斜体)
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
            # 没有选中文本时，设置插入点的格式
            current_pos = self.content_text.index(tk.INSERT)
            # 为将要输入的文本设置格式标记
            pass
            
    def clear_format(self):
        """清除选中文本的所有格式"""
        try:
            start = self.content_text.index(tk.SEL_FIRST)
            end = self.content_text.index(tk.SEL_LAST)
            
            # 移除所有格式标签
            for tag in ["bold", "italic", "underline", "bold_italic", "link"]:
                self.content_text.tag_remove(tag, start, end)
                
        except tk.TclError:
            pass
            
    def get_html_content(self):
        """获取带格式的HTML内容"""
        content = self.content_text.get(1.0, tk.END)
        html_content = ""
        
        # 获取所有文本和格式信息
        index = "1.0"
        while True:
            next_index = self.content_text.index(f"{index}+1c")
            if self.content_text.compare(index, ">=", tk.END):
                break
                
            char = self.content_text.get(index, next_index)
            tags = self.content_text.tag_names(index)
            
            # 构建HTML标签
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
            
            # 检查是否需要关闭标签
            next_tags = self.content_text.tag_names(next_index) if next_index != tk.END else []
            
            html_content += "".join(open_tags) + char
            
            # 如果下一个字符的标签不同，关闭当前标签
            if set(tags) != set(next_tags):
                html_content += "".join(close_tags)
            
            index = next_index
            
        # 处理换行
        html_content = html_content.replace('\n', '<br>\n')
        
        return html_content
            
    def clear_form(self):
        """清空表单"""
        self.subject.delete(0, tk.END)
        self.content_text.delete(1.0, tk.END)
        for recipient in self.recipients:
            recipient.delete(0, tk.END)
        for cc in self.cc_emails:
            cc.delete(0, tk.END)
        for bcc in self.bcc_emails:
            bcc.delete(0, tk.END)
        
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

    def send_email(self):
        """发送邮件"""
        # 验证必填字段
        if not self.sender_email.get():
            messagebox.showerror("错误", "请输入发件邮箱")
            return
            
        if not self.subject.get():
            messagebox.showerror("错误", "请输入邮件主题")
            return
            
        # 收集收件人
        to_emails = []
        for recipient in self.recipients:
            email = recipient.get().strip()
            if email:
                to_emails.append(email)
                
        if not to_emails:
            messagebox.showerror("错误", "请至少输入一个收件人")
            return
            
        # 构建发送参数
        sender_name = self.sender_name.get()
        sender_email = self.sender_email.get()
        
        if sender_name:
            from_address = f"{sender_name} <{sender_email}>"
        else:
            from_address = sender_email
            
        # 收集抄送和密送
        cc_emails = []
        for cc in self.cc_emails:
            email = cc.get().strip()
            if email:
                cc_emails.append(email)
                
        bcc_emails = []
        for bcc in self.bcc_emails:
            email = bcc.get().strip()
            if email:
                bcc_emails.append(email)
        
        # 保存历史
        self.add_to_history("sender_names", sender_name)
        self.add_to_history("sender_emails", sender_email)
        for email in to_emails:
            self.add_to_history("recipient_emails", email)
        
        params = {
            "from": from_address,
            "to": to_emails,
            "subject": self.subject.get(),
            "html": self.get_html_content()
        }
        
        # 添加可选字段
        if cc_emails:
            params["cc"] = cc_emails
        if bcc_emails:
            params["bcc"] = bcc_emails
        if self.reply_to.get():
            params["reply_to"] = [self.reply_to.get()]
        # 附件处理
        if hasattr(self, 'attachments') and self.attachments:
            params["attachments"] = self.attachments
        # 附件邮件不允许延迟发送
        if hasattr(self, 'attachments') and self.attachments and self.send_type.get() == "delay":
            messagebox.showerror("错误", "带附件的邮件不支持延迟发送！")
            return
        # 延迟发送参数
        if self.send_type.get() == "delay":
            try:
                date_str = self.scheduled_date.get()
                hour = int(self.scheduled_hour.get())
                minute = int(self.scheduled_minute.get())
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                scheduled_dt = dt.replace(hour=hour, minute=minute, second=0, microsecond=0)
                now = datetime.now()
                # 处理时区
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
                params["scheduled_at"] = scheduled_dt.isoformat()
            except Exception:
                messagebox.showerror("错误", "延迟发送时间设置有误！")
                return
            
        self.show_sending_loading()
        def thread_target():
            try:
                self.send_email_thread(params)
            finally:
                self.root.after(0, self.hide_sending_loading)
        threading.Thread(target=thread_target, daemon=True).start()
        
    def send_email_thread(self, params):
        """在后台线程发送邮件"""
        try:
            response = resend.Emails.send(params)
            # 保存到历史记录
            # 判断是否延迟发送
            is_scheduled = "scheduled_at" in params
            email_record = {
                "id": response.get("id"),
                "params": params,
                "sent_at": datetime.now().isoformat(),
                "status": "scheduled" if is_scheduled else "delivered"
            }
            self.email_history.append(email_record)
            self.save_history()
            # 动态插入到历史窗口
            if hasattr(self, "history_tree") and self.history_tree.winfo_exists():
                self.insert_history_row(email_record, self.history_tree, prepend=True)
            # 更新UI
            self.root.after(0, lambda: self.on_email_sent(response.get("id")))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("发送失败", f"邮件发送失败: {str(e)}"))
            
    def on_email_sent(self, email_id):
        """邮件发送成功回调"""
        # 判断是否为延迟发送
        is_scheduled = False
        for rec in self.email_history:
            if rec.get('id') == email_id:
                is_scheduled = rec.get('status') == 'scheduled'
                break
        if is_scheduled:
            messagebox.showinfo("已创建计划任务", f"计划任务已创建！\n邮件ID: {email_id}")
        else:
            messagebox.showinfo("发送成功", f"邮件发送成功！\n邮件ID: {email_id}")
            
    def show_history(self):
        """显示邮件历史（最多只能开一个窗口，已开时激活）"""
        if not hasattr(self, "history_window") or not self.history_window or not self.history_window.winfo_exists():
            self.selected_tz = getattr(self, 'selected_tz', 'local')
        if hasattr(self, "history_window") and self.history_window and self.history_window.winfo_exists():
            self.history_window.lift()
            self.history_window.focus_force()
            return
        history_window = tk.Toplevel(self.root)
        self.history_window = history_window
        history_window.title("邮件发送历史（列表最后更新时间：" + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "）")
        history_window.geometry("1100x650")
        history_window.configure(bg='#E6F3FF')
        try:
            history_window.iconbitmap("email.ico")
        except Exception:
            pass
        frame = ttk.Frame(history_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        columns = ("创建时间", "收件人", "主题", "状态", "递送/计划时间", "操作")
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=15)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150)
        tree.column("操作", width=220)
        tree.column("递送/计划时间", width=220)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_tree = tree  # 便于动态插入
        self.history_cancel_buttons = {}
        self.history_status_cache = {}
        # loading蒙版
        loading_mask = tk.Label(history_window, text="Loading...", bg="#E6F3FF", fg="#2E5F8C", font=("Arial", 24), bd=2, relief="groove")
        loading_mask.place_forget()
        def show_loading():
            loading_mask.place(relx=0.5, rely=0.5, anchor="center")
            history_window.update()
            # 同步更新时间到标题
            history_window.title(f"邮件发送历史（列表最后更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）")
        def hide_loading():
            loading_mask.place_forget()
            history_window.update()
        # 时区选择，菜单与主页一致
        tz_options = [("系统时区", 'local')] + [(label, label if isinstance(offset, str) else f'UTC{offset:+g}' if isinstance(offset, (int, float)) and offset != 0 else 'UTC') for label, offset in self.timezone_options]
        # 兼容Asia/Shanghai等
        def tz_code_to_zoneinfo(code):
            if code == 'local':
                try:
                    return tzlocal.get_localzone()
                except Exception:
                    return ZoneInfo('UTC')
            if code.startswith('UTC') and code != 'UTC':
                # 处理UTC+8等
                try:
                    hours = float(code[3:])
                    return timezone(timedelta(hours=hours))
                except Exception:
                    return ZoneInfo('UTC')
            try:
                return ZoneInfo(code)
            except Exception:
                return ZoneInfo('UTC')
        def get_display_tz():
            return tz_code_to_zoneinfo(self.selected_tz)
        def format_time(timestr):
            if not timestr or timestr == '-':
                return '-'
            try:
                # 支持 2025-07-02 15:07:00+00
                if '+' in timestr:
                    dt = datetime.fromisoformat(timestr.replace(' ', 'T'))
                else:
                    dt = datetime.fromisoformat(timestr)
                tz = get_display_tz()
                if hasattr(dt, 'astimezone'):
                    dt = dt.astimezone(tz)
                return dt.strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                return timestr
        def get_delivery_time(record, status, full_resp):
            # 优先本地保存
            if status == 'delivered':
                t = record.get('created_at') or full_resp.get('created_at', '-')
                return format_time(t)
            elif status == 'scheduled':
                t = record.get('scheduled_at') or full_resp.get('scheduled_at', '-')
                return format_time(t)
            elif status == 'canceled':
                return '-'
            else:
                return '-'
        def insert_history_row(record, tree, prepend=False):
            sent_time = datetime.fromisoformat(record["sent_at"]).strftime("%Y-%m-%d %H:%M:%S")
            recipients = ", ".join(record["params"]["to"])
            subject = record["params"]["subject"]
            status = record["status"]
            # 操作栏状态修正
            if status == "scheduled":
                op = "计划可修改，可双击此项修改"
            elif status == "canceled":
                op = "发送计划已取消"
            elif status == "delivered":
                op = "已成功投递"
            else:
                op = "-"
            # 递送/计划时间，优先本地，无则同步接口
            full_resp = self.history_status_cache.get(record.get('id'), {})
            delivery_time = None
            if status == 'delivered':
                t = record.get('created_at')
                if not t:
                    try:
                        resp = resend.Emails.get(email_id=record["id"])
                        t = resp.get('created_at')
                        if t:
                            record['created_at'] = t
                            self.history_status_cache[record['id']] = resp
                    except Exception:
                        pass
                delivery_time = format_time(t)
            elif status == 'scheduled':
                t = record.get('scheduled_at')
                if not t:
                    try:
                        resp = resend.Emails.get(email_id=record["id"])
                        t = resp.get('scheduled_at')
                        if t:
                            record['scheduled_at'] = t
                            self.history_status_cache[record['id']] = resp
                    except Exception:
                        pass
                delivery_time = format_time(t)
            elif status == 'canceled':
                delivery_time = '-'
            else:
                delivery_time = '-'
            values = (sent_time, recipients, subject, status, delivery_time, op)
            if prepend:
                item_id = tree.insert("", 0, values=values)
            else:
                item_id = tree.insert("", "end", values=values)
            self.history_cancel_buttons[item_id] = record
        self.insert_history_row = insert_history_row
        def fetch_status(record):
            try:
                if record.get("status") == "scheduled":
                    resp = resend.Emails.get(email_id=record["id"])
                    return resp.get("last_event", record.get("status", "unknown")), resp
                else:
                    return record.get("status", "unknown"), record
            except Exception:
                return record.get("status", "unknown"), record
        show_loading()
        self.root.after(100, lambda: self._finish_show_history(tree, insert_history_row, hide_loading))

    def _finish_show_history(self, tree, insert_history_row, hide_loading):
        for idx, record in enumerate(reversed(self.email_history)):
            insert_history_row(record, tree, prepend=False)
        hide_loading()
        # 重新绑定事件，防止异步后失效
        def show_detail_by_item(item_id):
            record = self.history_cancel_buttons.get(item_id)
            if not record:
                return
            email_id = record.get("id")
            detail = self.history_status_cache.get(email_id)
            if not detail:
                try:
                    detail = resend.Emails.get(email_id=email_id)
                except Exception as e:
                    messagebox.showerror("错误", f"获取详情失败: {str(e)}", parent=self.history_window)
                    return
            self.show_email_detail(detail)
        def on_tree_double_click(event):
            item_id = tree.identify_row(event.y)
            if item_id:
                show_detail_by_item(item_id)
        tree.bind("<Double-1>", on_tree_double_click)
        def on_tree_right_click(event):
            item_id = tree.identify_row(event.y)
            if item_id:
                tree.selection_set(item_id)
            menu = tk.Menu(self.history_window, tearoff=0)
            if item_id:
                menu.add_command(label="查看详情", command=lambda: show_detail_by_item(item_id))
                menu.add_command(label="刷新此项", command=lambda: self.refresh_one(item_id))
            menu.add_separator()
            menu.add_command(label="刷新所有邮件信息(较缓慢)", command=lambda: self.refresh_all(True))
            menu.add_command(label="仅刷新计划投递邮件信息", command=lambda: self.refresh_all(False))
            # 递送时间时区格式选择二级菜单，当前选中项加对钩
            tz_menu = tk.Menu(menu, tearoff=0)
            for label, code in [("系统时区", 'local')] + [(label, label if isinstance(offset, str) else f'UTC{offset:+g}' if isinstance(offset, (int, float)) and offset != 0 else 'UTC') for label, offset in self.timezone_options]:
                checked = (self.selected_tz == code)
                display_label = f'✔ {label}' if checked else label
                tz_menu.add_command(label=display_label, command=lambda c=code: self.on_tz_select(c))
            menu.add_cascade(label="递送时间时区格式选择", menu=tz_menu)
            menu.tk_popup(event.x_root, event.y_root)
        tree.bind("<Button-3>", on_tree_right_click)

    def show_email_detail(self, detail):
        """显示邮件详细信息，detail为接口返回的完整字典。每个邮件只允许开一个详情窗口。"""
        if not hasattr(self, 'detail_windows'):
            self.detail_windows = {}
        mail_id = detail.get('id')
        if mail_id in self.detail_windows and self.detail_windows[mail_id].winfo_exists():
            self.detail_windows[mail_id].lift()
            self.detail_windows[mail_id].focus_force()
            return
        detail_window = tk.Toplevel(self.root)
        self.detail_windows[mail_id] = detail_window
        detail_window.title("邮件详情")
        detail_window.geometry("600x500")
        detail_window.configure(bg='#E6F3FF')
        try:
            detail_window.iconbitmap("email.ico")
        except Exception:
            pass
        frame = ttk.Frame(detail_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        detail_text = scrolledtext.ScrolledText(frame, height=20, wrap=tk.WORD)
        detail_text.pack(fill=tk.BOTH, expand=True)
        info = ""
        for k, v in detail.items():
            info += f"{k}: {v}\n"
        detail_text.insert(1.0, info)
        detail_text.config(state='disabled')
        # scheduled状态显示操作按钮
        if detail.get("last_event") == "scheduled":
            btn_frame = ttk.Frame(frame)
            btn_frame.pack(fill=tk.X, pady=10)
            btn_cancel = ttk.Button(btn_frame, text="取消计划", command=lambda: self.cancel_scheduled_from_detail(detail, detail_window))
            btn_cancel.pack(side=tk.LEFT, padx=10)
            btn_update = ttk.Button(btn_frame, text="更新计划", command=lambda: self.update_scheduled_from_detail(detail, detail_window))
            btn_update.pack(side=tk.LEFT, padx=10)
        def on_close():
            if mail_id in self.detail_windows:
                del self.detail_windows[mail_id]
            detail_window.destroy()
        detail_window.protocol("WM_DELETE_WINDOW", on_close)

    def cancel_scheduled_from_detail(self, detail, win):
        try:
            response = resend.Emails.cancel(detail["id"])
            messagebox.showinfo("成功", f"邮件已取消: {response.get('id')}", parent=win)
            win.destroy()
            # 自动刷新该条记录（强制接口获取），并高亮
            if hasattr(self, 'history_window') and self.history_window and self.history_window.winfo_exists():
                self.history_window.lift()
                self.history_window.focus_force()
            if hasattr(self, 'history_tree') and self.history_tree.winfo_exists():
                for iid, rec in self.history_cancel_buttons.items():
                    if rec.get('id') == detail.get('id'):
                        def refresh_one(item_id):
                            try:
                                resp = resend.Emails.get(email_id=rec["id"])
                                last_event = resp.get("last_event", rec.get("status", "unknown"))
                                self.history_status_cache[rec["id"]] = resp
                                self.history_tree.set(iid, "状态", last_event)
                                if last_event == "scheduled":
                                    self.history_tree.set(iid, "操作", "计划可修改，可双击此项修改")
                                elif last_event == "canceled":
                                    self.history_tree.set(iid, "操作", "发送计划已取消")
                                elif last_event == "delivered":
                                    self.history_tree.set(iid, "操作", "已成功投递")
                                else:
                                    self.history_tree.set(iid, "操作", "-")
                                # 写入本地递送/计划时间
                                if last_event == 'delivered' and resp.get('created_at'):
                                    rec['created_at'] = resp.get('created_at')
                                if last_event == 'scheduled' and resp.get('scheduled_at'):
                                    rec['scheduled_at'] = resp.get('scheduled_at')
                                delivery_time = self.history_tree.set(iid, "递送/计划时间", resp.get('created_at') if last_event == 'delivered' else (resp.get('scheduled_at') if last_event == 'scheduled' else '-'))
                                rec["status"] = last_event
                                self.save_history()
                                # 高亮该行
                                self.history_tree.selection_set(iid)
                                self.history_tree.see(iid)
                            except Exception as e:
                                messagebox.showerror("取消失败", f"无法取消邮件: {str(e)}", parent=win)
                        refresh_one(iid)
                        break
        except Exception as e:
            messagebox.showerror("取消失败", f"无法取消邮件: {str(e)}", parent=win)

    def update_scheduled_from_detail(self, detail, win):
        # 只允许开一个更新计划弹窗
        if not hasattr(self, 'update_popup_windows'):
            self.update_popup_windows = {}
        mail_id = detail.get('id')
        if mail_id in self.update_popup_windows and self.update_popup_windows[mail_id].winfo_exists():
            self.update_popup_windows[mail_id].lift()
            self.update_popup_windows[mail_id].focus_force()
            return

        popup = tk.Toplevel(win)
        self.update_popup_windows[mail_id] = popup
        popup.title("更新计划时间")
        popup.geometry("520x140")

        for i in range(7):
            popup.grid_columnconfigure(i, weight=0)

        ttk.Label(popup, text="选择新计划时间:").grid(row=0, column=0, padx=(5,2), pady=5, sticky="w")

        today = datetime.now().date()
        max_day = today + timedelta(days=30)

        date_entry = DateEntry(popup, mindate=today, maxdate=max_day, date_pattern='yyyy-mm-dd', width=12)
        date_entry.grid(row=0, column=1, padx=(0,2), pady=5, sticky="w")

        # 用 Frame 包裹时间部分，避免空隙
        time_frame = ttk.Frame(popup)
        time_frame.grid(row=0, column=2, padx=(0,0), pady=5, sticky="w")

        hour_box = ttk.Combobox(time_frame, width=3, values=[f"{i:02d}" for i in range(24)], state="readonly")
        hour_box.set(f"{datetime.now().hour:02d}")
        hour_box.pack(side="left")

        colon = ttk.Label(time_frame, text=":")
        colon.pack(side="left", padx=(2,2))

        min_box = ttk.Combobox(time_frame, width=3, values=[f"{i:02d}" for i in range(60)], state="readonly")
        min_box.set(f"{datetime.now().minute:02d}")
        min_box.pack(side="left")

        tz_box = ttk.Combobox(popup, width=8, state="readonly", values=[x[0] for x in self.timezone_options])
        tz_box.set("UTC+8")
        tz_box.grid(row=0, column=3, padx=(4,0), pady=5, sticky="w")

        ttk.Label(popup, text="(30天内)").grid(row=0, column=4, padx=(4,0), pady=5, sticky="w")

        # 按钮单独一行左对齐
        btn_save = ttk.Button(popup, text="保存", command=lambda: do_save())
        btn_save.grid(row=1, column=1, pady=10, sticky="w")

        btn_cancel = ttk.Button(popup, text="取消", command=lambda: do_cancel())
        btn_cancel.grid(row=1, column=2, pady=10, sticky="w")

        def do_save():
            try:
                date_str = date_entry.get()
                hour = int(hour_box.get())
                minute = int(min_box.get())
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                scheduled_dt = dt.replace(hour=hour, minute=minute, second=0, microsecond=0)
                now = datetime.now()
                tz_label = tz_box.get()
                tz_offset = 0
                for label, offset in self.timezone_options:
                    if label == tz_label:
                        tz_offset = offset
                        break
                if isinstance(tz_offset, float):
                    hours = int(tz_offset)
                    minutes = int((abs(tz_offset) - abs(hours)) * 60)
                    tz = timezone(timedelta(hours=hours, minutes=minutes if tz_offset >= 0 else -minutes))
                else:
                    tz = timezone(timedelta(hours=tz_offset))
                scheduled_dt = scheduled_dt.replace(tzinfo=tz)
                if scheduled_dt <= now.astimezone(tz):
                    messagebox.showerror("错误", "预约时间必须晚于当前时间！", parent=popup)
                    return
                if scheduled_dt > (now + timedelta(days=30)).astimezone(tz):
                    messagebox.showerror("错误", "预约时间不能超过30天！", parent=popup)
                    return
                update_params = {"id": detail["id"], "scheduled_at": scheduled_dt.isoformat()}
                response = resend.Emails.update(params=update_params)
                messagebox.showinfo("成功", f"计划已更新: {response.get('id')}", parent=popup)
                popup.destroy()
                win.destroy()
                # 自动刷新该条记录（强制接口获取），并高亮
                if hasattr(self, 'history_window') and self.history_window and self.history_window.winfo_exists():
                    self.history_window.lift()
                    self.history_window.focus_force()
                if hasattr(self, 'history_tree') and self.history_tree.winfo_exists():
                    tree = self.history_tree
                    # 复用show_history作用域下的get_delivery_time
                    def get_delivery_time(record, status, full_resp):
                        # 优先本地保存
                        if status == 'delivered':
                            t = record.get('created_at') or full_resp.get('created_at', '-')
                            return format_time(t)
                        elif status == 'scheduled':
                            t = record.get('scheduled_at') or full_resp.get('scheduled_at', '-')
                            return format_time(t)
                        elif status == 'canceled':
                            return '-'
                        else:
                            return '-'
                    def format_time(timestr):
                        if not timestr or timestr == '-':
                            return '-'
                        try:
                            if '+' in timestr:
                                dt = datetime.fromisoformat(timestr.replace(' ', 'T'))
                            else:
                                dt = datetime.fromisoformat(timestr)
                            # 用当前所选时区
                            tz = None
                            if hasattr(self, 'selected_tz'):
                                tz = tzlocal.get_localzone() if self.selected_tz == 'local' else timezone(timedelta(hours=float(self.selected_tz[3:]))) if self.selected_tz.startswith('UTC') and self.selected_tz != 'UTC' else ZoneInfo(self.selected_tz) if self.selected_tz != 'UTC' else timezone.utc
                            else:
                                tz = timezone.utc
                            if hasattr(dt, 'astimezone'):
                                dt = dt.astimezone(tz)
                            return dt.strftime('%Y-%m-%d %H:%M:%S')
                        except Exception:
                            return timestr
                    for iid, rec in self.history_cancel_buttons.items():
                        if rec.get('id') == detail.get('id'):
                            def refresh_one(item_id):
                                try:
                                    resp = resend.Emails.get(email_id=rec["id"])
                                    last_event = resp.get("last_event", rec.get("status", "unknown"))
                                    self.history_status_cache[rec["id"]] = resp
                                    tree.set(iid, "状态", last_event)
                                    if last_event == "scheduled":
                                        tree.set(iid, "操作", "计划可修改，可双击此项修改")
                                    elif last_event == "canceled":
                                        tree.set(iid, "操作", "发送计划已取消")
                                    elif last_event == "delivered":
                                        tree.set(iid, "操作", "已成功投递")
                                    else:
                                        tree.set(iid, "操作", "-")
                                    # 写入本地递送/计划时间
                                    if last_event == 'delivered' and resp.get('created_at'):
                                        rec['created_at'] = resp.get('created_at')
                                    if last_event == 'scheduled' and resp.get('scheduled_at'):
                                        rec['scheduled_at'] = resp.get('scheduled_at')
                                    delivery_time = get_delivery_time(rec, last_event, resp)
                                    tree.set(iid, "递送/计划时间", delivery_time)
                                    rec["status"] = last_event
                                    self.save_history()
                                    # 高亮该行
                                    tree.selection_set(iid)
                                    tree.see(iid)
                                except Exception as e:
                                    messagebox.showerror("更新失败", f"无法更新计划: {str(e)}", parent=popup)
                            refresh_one(iid)
                            break
            except Exception as e:
                messagebox.showerror("更新失败", f"无法更新计划: {str(e)}", parent=popup)

        def do_cancel():
            popup.destroy()

        def on_close():
            if mail_id in self.update_popup_windows:
                del self.update_popup_windows[mail_id]
            popup.destroy()

        popup.protocol("WM_DELETE_WINDOW", on_close)
        
    def load_history(self):
        """加载邮件历史"""
        try:
            if os.path.exists("email_history.json"):
                with open("email_history.json", "r", encoding="utf-8") as f:
                    self.email_history = json.load(f)
        except Exception:
            self.email_history = []
            
    def save_history(self):
        """保存邮件历史"""
        try:
            with open("email_history.json", "w", encoding="utf-8") as f:
                json.dump(self.email_history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
            
    def show_history_menu(self, event, key, combobox):
        menu = tk.Menu(self.root, tearoff=0)
        for item in self.input_history[key]:
            menu.add_command(label=item, command=lambda v=item: combobox.set(v))
        if self.input_history[key]:
            menu.add_separator()
            menu.add_command(label="清空全部", command=lambda: self.clear_history_and_update(key, combobox))
        menu.tk_popup(event.x_root, event.y_root)

    def clear_history_and_update(self, key, combobox):
        self.clear_history(key)
        combobox['values'] = []
        combobox.set("")
            
    def run(self):
        """运行应用"""
        if hasattr(self, 'root'):
            self.root.mainloop()

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

    def upload_attachments(self):
        files = filedialog.askopenfilenames(title="选择附件", filetypes=[("所有文件", "*.*")])
        if not files:
            return
        for file_path in files:
            ext = os.path.splitext(file_path)[1].lower()
            if ext in self.attachment_blacklist:
                messagebox.showerror("不支持的附件类型", f"{file_path} 为不支持的类型，禁止上传！")
                continue
            self.add_attachment_from_path(file_path)
        self.update_attachment_label()

    def on_drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            ext = os.path.splitext(file_path)[1].lower()
            if ext in self.attachment_blacklist:
                messagebox.showerror("不支持的附件类型", f"{file_path} 为不支持的类型，禁止上传！")
                continue
            self.add_attachment_from_path(file_path)
        self.update_attachment_label()

    def add_attachment_from_path(self, file_path):
        max_size = 40 * 1024 * 1024
        size = os.path.getsize(file_path)
        filename = os.path.basename(file_path)
        if size <= max_size:
            with open(file_path, "rb") as f:
                content = f.read()
            b64_content = base64.b64encode(content).decode()
            self.attachments.append({"content": b64_content, "filename": filename})
        else:
            url = tk.simpledialog.askstring("大文件附件", f"{filename} 超过40MB，请输入远程附件URL：")
            if url:
                self.attachments.append({"path": url, "filename": filename})

    def update_attachment_label(self):
        # 清空原有
        for widget in self.attachment_list_frame.winfo_children():
            widget.destroy()
        if not self.attachments:
            ttk.Label(self.attachment_list_frame, text="无附件").pack(side=tk.LEFT)
        else:
            for i, att in enumerate(self.attachments):
                name = att.get("filename", "")
                lbl = ttk.Label(self.attachment_list_frame, text=name)
                lbl.pack(side=tk.LEFT, padx=2)
                btn = ttk.Button(self.attachment_list_frame, text="×", width=2, command=(lambda idx=i: lambda: self.remove_attachment(idx))())
                btn.pack(side=tk.LEFT, padx=1)

    def remove_attachment(self, idx):
        del self.attachments[idx]
        self.update_attachment_label()

    def add_remote_attachment(self):
        url = tk.simpledialog.askstring("远程附件链接", "请输入远程附件URL：")
        if not url:
            return
        filename = tk.simpledialog.askstring("文件名", "请输入附件文件名（含扩展名）：")
        if not filename:
            return
        ext = os.path.splitext(filename)[1].lower()
        if ext in self.attachment_blacklist:
            messagebox.showerror("不支持的附件类型", f"{filename} 为不支持的类型，禁止上传！")
            return
        self.attachments.append({"path": url, "filename": filename})
        self.update_attachment_label()

    def validate_email_format(self, entry):
        email = entry.get().strip()
        if email and not EMAIL_REGEX.match(email):
            messagebox.showerror("邮箱格式错误", f"邮箱格式不正确: {email}")
            entry.focus_set()

    def refresh_one(self, item_id):
        tree = self.history_tree
        record = self.history_cancel_buttons.get(item_id)
        if not record:
            return
        def show_loading():
            if hasattr(self, 'history_window') and self.history_window:
                for widget in self.history_window.winfo_children():
                    if isinstance(widget, tk.Label) and widget.cget('text') == 'Loading...':
                        widget.place(relx=0.5, rely=0.5, anchor="center")
                        self.history_window.update()
        def hide_loading():
            if hasattr(self, 'history_window') and self.history_window:
                for widget in self.history_window.winfo_children():
                    if isinstance(widget, tk.Label) and widget.cget('text') == 'Loading...':
                        widget.place_forget()
                        self.history_window.update()
        show_loading()
        try:
            resp = resend.Emails.get(email_id=record["id"])
            last_event = resp.get("last_event", record.get("status", "unknown"))
            self.history_status_cache[record["id"]] = resp
            tree.set(item_id, "状态", last_event)
            if last_event == "scheduled":
                tree.set(item_id, "操作", "计划可修改，可双击此项修改")
            elif last_event == "canceled":
                tree.set(item_id, "操作", "发送计划已取消")
            elif last_event == "delivered":
                tree.set(item_id, "操作", "已成功投递")
            else:
                tree.set(item_id, "操作", "-")
            if last_event == 'delivered' and resp.get('created_at'):
                record['created_at'] = resp.get('created_at')
            if last_event == 'scheduled' and resp.get('scheduled_at'):
                record['scheduled_at'] = resp.get('scheduled_at')
            delivery_time = self.get_delivery_time(record, last_event, resp)
            tree.set(item_id, "递送/计划时间", delivery_time)
            record["status"] = last_event
            self.save_history()
        except Exception:
            pass
        hide_loading()

    def refresh_all(self, refresh_all):
        tree = self.history_tree
        history_window = self.history_window
        # Loading遮罩
        loading_mask = None
        for widget in history_window.winfo_children():
            if isinstance(widget, tk.Label) and widget.cget('text') == 'Loading...':
                loading_mask = widget
                break
        if not loading_mask:
            loading_mask = tk.Label(history_window, text="Loading...", bg="#E6F3FF", fg="#2E5F8C", font=("Arial", 24), bd=2, relief="groove")
            loading_mask.place_forget()
        loading_mask.place(relx=0.5, rely=0.5, anchor="center")
        history_window.update()
        updated = False
        for iid in tree.get_children():
            record = self.history_cancel_buttons.get(iid)
            if not record:
                continue
            need_refresh = refresh_all or (tree.set(iid, "状态") == "scheduled")
            if need_refresh:
                try:
                    resp = resend.Emails.get(email_id=record["id"])
                    last_event = resp.get("last_event", record.get("status", "unknown"))
                    self.history_status_cache[record["id"]] = resp
                    tree.set(iid, "状态", last_event)
                    if last_event == "scheduled":
                        tree.set(iid, "操作", "计划可修改，可双击此项修改")
                    elif last_event == "canceled":
                        tree.set(iid, "操作", "发送计划已取消")
                    elif last_event == "delivered":
                        tree.set(iid, "操作", "已成功投递")
                    else:
                        tree.set(iid, "操作", "-")
                    if last_event == 'delivered' and resp.get('created_at'):
                        record['created_at'] = resp.get('created_at')
                    if last_event == 'scheduled' and resp.get('scheduled_at'):
                        record['scheduled_at'] = resp.get('scheduled_at')
                    delivery_time = self.get_delivery_time(record, last_event, resp)
                    tree.set(iid, "递送/计划时间", delivery_time)
                    record["status"] = last_event
                    updated = True
                except Exception:
                    pass
        loading_mask.place_forget()
        history_window.title(f"邮件发送历史（列表最后更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）")
        history_window.update()
        if updated:
            self.save_history()

    def on_tz_select(self, tz_code):
        self.selected_tz = tz_code
        self.refresh_all_delivery_time()

    def get_delivery_time(self, record, status, full_resp):
        # 复制show_history作用域下的get_delivery_time
        if status == 'delivered':
            t = record.get('created_at') or full_resp.get('created_at', '-')
            return self.format_time(t)
        elif status == 'scheduled':
            t = record.get('scheduled_at') or full_resp.get('scheduled_at', '-')
            return self.format_time(t)
        elif status == 'canceled':
            return '-'
        else:
            return '-'

    def format_time(self, timestr):
        if not timestr or timestr == '-':
            return '-'
        try:
            if '+' in timestr:
                dt = datetime.fromisoformat(timestr.replace(' ', 'T'))
            else:
                dt = datetime.fromisoformat(timestr)
            tz = None
            if hasattr(self, 'selected_tz'):
                tz = tzlocal.get_localzone() if self.selected_tz == 'local' else timezone(timedelta(hours=float(self.selected_tz[3:]))) if self.selected_tz.startswith('UTC') and self.selected_tz != 'UTC' else ZoneInfo(self.selected_tz) if self.selected_tz != 'UTC' else timezone.utc
            else:
                tz = timezone.utc
            if hasattr(dt, 'astimezone'):
                dt = dt.astimezone(tz)
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except Exception:
            return timestr

    def refresh_all_delivery_time(self):
        tree = self.history_tree
        for iid in tree.get_children():
            record = self.history_cancel_buttons.get(iid)
            if not record:
                continue
            status = tree.set(iid, "状态")
            full_resp = self.history_status_cache.get(record.get('id'), {})
            delivery_time = self.get_delivery_time(record, status, full_resp)
            tree.set(iid, "递送/计划时间", delivery_time)

# 运行应用
if __name__ == "__main__":
    app = ResendEmailClient()
    app.run()