# 这里只写骨架，具体实现可从原main.py迁移
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog
from history import get_email_history, add_email_record
from utils import format_time
from datetime import datetime, timedelta, timezone
import resend
import tzlocal
from config import get_api_key
try:
    from zoneinfo import ZoneInfo
except ImportError:
    from pytz import timezone as ZoneInfo
from tkcalendar import DateEntry

class HistoryUI(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("邮件发送历史（列表最后更新时间：" + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "）")
        self.geometry("1100x650")
        self.configure(bg='#E6F3FF')
        try:
            self.iconbitmap("email.ico")
        except Exception:
            pass
        frame = ttk.Frame(self, padding="10")
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
        self.tree = tree
        self.history_cancel_buttons = {}
        self.history_status_cache = {}
        self.selected_tz = 'local'
        self.loading_mask = tk.Label(self, text="Loading...", bg="#E6F3FF", fg="#2E5F8C", font=("Arial", 24), bd=2, relief="groove")
        self.loading_mask.place_forget()
        self.timezone_options = [
            ("系统时区", 'local'),
            ("UTC-12", 'UTC-12'), ("UTC-11", 'UTC-11'), ("UTC-10", 'UTC-10'), ("UTC-9", 'UTC-9'), ("UTC-8", 'UTC-8'), ("UTC-7", 'UTC-7'), ("UTC-6", 'UTC-6'), ("UTC-5", 'UTC-5'), ("UTC-4", 'UTC-4'), ("UTC-3", 'UTC-3'), ("UTC-2", 'UTC-2'), ("UTC-1", 'UTC-1'),
            ("UTC", 'UTC'), ("UTC+1", 'UTC+1'), ("UTC+2", 'UTC+2'), ("UTC+3", 'UTC+3'), ("UTC+4", 'UTC+4'), ("UTC+5", 'UTC+5'), ("UTC+5:30", 'UTC+5.5'), ("UTC+6", 'UTC+6'), ("UTC+7", 'UTC+7'), ("UTC+8", 'UTC+8'), ("UTC+9", 'UTC+9'), ("UTC+9:30", 'UTC+9.5'), ("UTC+10", 'UTC+10'), ("UTC+11", 'UTC+11'), ("UTC+12", 'UTC+12')
        ]
        self.detail_windows = {}
        self.update_popup_windows = {}
        self.load_history()
        tree.bind("<Double-1>", self.on_tree_double_click)
        tree.bind("<Button-3>", self.on_tree_right_click)

    def show_loading(self):
        self.loading_mask.place(relx=0.5, rely=0.5, anchor="center")
        self.update()
        self.title(f"邮件发送历史（列表最后更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）")
    def hide_loading(self):
        self.loading_mask.place_forget()
        self.update()

    def load_history(self):
        self.show_loading()
        self.tree.delete(*self.tree.get_children())
        email_history = get_email_history()
        need_api_update = []
        need_expired_refresh = []
        now = datetime.now(timezone.utc)
        for record in reversed(email_history):
            sent_time = datetime.fromisoformat(record["sent_at"]).strftime("%Y-%m-%d %H:%M:%S")
            recipients = ", ".join(record["params"]["to"])
            subject = record["params"]["subject"]
            status = record["status"]
            op = "计划可修改，可双击此项修改" if status == "scheduled" else ("发送计划已取消" if status == "canceled" else ("已成功投递" if status == "delivered" else "-"))
            full_resp = self.history_status_cache.get(record.get('id'), {})
            delivery_time = self.get_delivery_time(record, status, full_resp)
            values = (sent_time, recipients, subject, status, delivery_time, op)
            iid = self.tree.insert("", 0, values=values)
            self.history_cancel_buttons[iid] = record
            if (status == 'scheduled' and not record.get('scheduled_at')) or (status == 'delivered' and not record.get('created_at')):
                need_api_update.append((iid, record, status))
            if status == 'scheduled' and record.get('scheduled_at'):
                try:
                    scheduled_dt = datetime.fromisoformat(record['scheduled_at'].replace(' ', 'T'))
                    if scheduled_dt.tzinfo is None:
                        scheduled_dt = scheduled_dt.replace(tzinfo=timezone.utc)
                    if scheduled_dt < now:
                        need_expired_refresh.append((iid, record))
                except Exception:
                    pass
        self.hide_loading()
        if need_api_update:
            self.after(100, lambda: self._auto_update_delivery_time(need_api_update))
        if need_expired_refresh:
            self.after(200, lambda: self._auto_refresh_expired_scheduled(need_expired_refresh))

    def _auto_update_delivery_time(self, need_api_update):
        self.show_loading()
        from config import get_api_key
        import resend
        from history import history_manager
        resend.api_key = get_api_key()
        updated = False
        for iid, record, status in need_api_update:
            try:
                resp = resend.Emails.get(email_id=record["id"])
                if status == 'scheduled' and resp.get('scheduled_at'):
                    record['scheduled_at'] = resp['scheduled_at']
                    self.history_status_cache[record['id']] = resp
                    delivery_time = self.get_delivery_time(record, 'scheduled', resp)
                    self.tree.set(iid, "递送/计划时间", delivery_time)
                    updated = True
                if status == 'delivered' and resp.get('created_at'):
                    record['created_at'] = resp['created_at']
                    self.history_status_cache[record['id']] = resp
                    delivery_time = self.get_delivery_time(record, 'delivered', resp)
                    self.tree.set(iid, "递送/计划时间", delivery_time)
                    updated = True
            except Exception:
                pass
        if updated:
            history_manager.save_email_history()
        self.hide_loading()

    def _auto_refresh_expired_scheduled(self, need_expired_refresh):
        resend.api_key = get_api_key()
        for iid, record in need_expired_refresh:
            try:
                resp = resend.Emails.get(email_id=record["id"])
                last_event = resp.get("last_event", record.get("status", "unknown"))
                self.history_status_cache[record["id"]] = resp
                self.tree.set(iid, "状态", last_event)
                if last_event == "scheduled":
                    self.tree.set(iid, "操作", "计划可修改，可双击此项修改")
                elif last_event == "canceled":
                    self.tree.set(iid, "操作", "发送计划已取消")
                elif last_event == "delivered":
                    self.tree.set(iid, "操作", "已成功投递")
                else:
                    self.tree.set(iid, "操作", "-")
                if last_event == 'delivered' and resp.get('created_at'):
                    record['created_at'] = resp.get('created_at')
                if last_event == 'scheduled' and resp.get('scheduled_at'):
                    record['scheduled_at'] = resp.get('scheduled_at')
                delivery_time = self.get_delivery_time(record, last_event, resp)
                self.tree.set(iid, "递送/计划时间", delivery_time)
                record["status"] = last_event
            except Exception:
                pass

    def get_delivery_time(self, record, status, full_resp):
        if status == 'delivered':
            t = record.get('created_at') or full_resp.get('created_at', '-')
            return format_time(t, self.selected_tz)
        elif status == 'scheduled':
            t = record.get('scheduled_at') or full_resp.get('scheduled_at', '-')
            return format_time(t, self.selected_tz)
        elif status == 'canceled':
            return '-'
        else:
            return '-'

    def on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id:
            self.show_detail_by_item(item_id)

    def on_tree_right_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id:
            self.tree.selection_set(item_id)
        menu = tk.Menu(self, tearoff=0)
        if item_id:
            menu.add_command(label="查看详情", command=lambda: self.show_detail_by_item(item_id))
            menu.add_command(label="刷新此项", command=lambda: self.refresh_one(item_id))
        menu.add_separator()
        menu.add_command(label="刷新所有邮件信息(较缓慢)", command=lambda: self.refresh_all(True))
        menu.add_command(label="仅刷新计划投递邮件信息", command=lambda: self.refresh_all(False))
        # 时区切换
        tz_menu = tk.Menu(menu, tearoff=0)
        for label, code in self.timezone_options:
            checked = (self.selected_tz == code)
            display_label = f'✔ {label}' if checked else label
            tz_menu.add_command(label=display_label, command=lambda c=code: self.on_tz_select(c))
        menu.add_cascade(label="递送时间时区格式选择", menu=tz_menu)
        menu.tk_popup(event.x_root, event.y_root)

    def show_detail_by_item(self, item_id):
        record = self.history_cancel_buttons.get(item_id)
        if not record:
            return
        email_id = record.get("id")
        detail = self.history_status_cache.get(email_id)
        if not detail:
            try:
                resend.api_key = get_api_key()
                detail = resend.Emails.get(email_id=email_id)
            except Exception as e:
                messagebox.showerror("错误", f"获取详情失败: {str(e)}", parent=self)
                return
        self.show_email_detail(detail)

    def show_email_detail(self, detail):
        mail_id = detail.get('id')
        if mail_id in self.detail_windows and self.detail_windows[mail_id].winfo_exists():
            self.detail_windows[mail_id].lift()
            self.detail_windows[mail_id].focus_force()
            return
        detail_window = tk.Toplevel(self)
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
        # 追加本地附件信息
        local_attachments = None
        for rec in get_email_history():
            if rec.get('id') == mail_id:
                local_attachments = rec.get('attachments', None)
                break
        info += "\n附件信息：\n"
        if local_attachments and len(local_attachments) > 0:
            for att in local_attachments:
                name = att.get("filename", "")
                size_kb = att.get("size_kb", 0)
                path = att.get("path", "")
                info += f"{name}({size_kb}kb)[{path}]\n"
        else:
            info += "No attachments\n"
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
            resend.api_key = get_api_key()
            response = resend.Emails.cancel(detail["id"])
            messagebox.showinfo("成功", f"邮件已取消: {response.get('id')}", parent=win)
            win.destroy()
            for iid, rec in self.history_cancel_buttons.items():
                if rec.get('id') == detail.get('id'):
                    self.refresh_one(iid)
                    self.tree.selection_set(iid)
                    self.tree.see(iid)
                    break
        except Exception as e:
            messagebox.showerror("取消失败", f"无法取消邮件: {str(e)}", parent=win)

    def update_scheduled_from_detail(self, detail, win):
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
                for label, code in self.timezone_options:
                    if label == tz_label:
                        tz_offset = code
                        break
                if isinstance(tz_offset, str) and tz_offset.startswith('UTC') and tz_offset != 'UTC':
                    hours = float(tz_offset[3:])
                    tz = timezone(timedelta(hours=hours))
                elif tz_offset == 'UTC':
                    tz = timezone.utc
                elif tz_offset == 'local':
                    tz = tzlocal.get_localzone()
                else:
                    tz = timezone.utc
                scheduled_dt = scheduled_dt.replace(tzinfo=tz)
                if scheduled_dt <= now.astimezone(tz):
                    messagebox.showerror("错误", "预约时间必须晚于当前时间！", parent=popup)
                    return
                if scheduled_dt > (now + timedelta(days=30)).astimezone(tz):
                    messagebox.showerror("错误", "预约时间不能超过30天！", parent=popup)
                    return
                resend.api_key = get_api_key()
                response = resend.Emails.update(params={"id": detail["id"], "scheduled_at": scheduled_dt.isoformat()})
                messagebox.showinfo("成功", f"计划已更新: {response.get('id')}", parent=popup)
                popup.destroy()
                win.destroy()
                for iid, rec in self.history_cancel_buttons.items():
                    if rec.get('id') == detail.get('id'):
                        self.refresh_one(iid)
                        self.tree.selection_set(iid)
                        self.tree.see(iid)
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

    def refresh_one(self, item_id):
        self.show_loading()
        record = self.history_cancel_buttons.get(item_id)
        if not record:
            self.hide_loading()
            return
        try:
            resend.api_key = get_api_key()
            resp = resend.Emails.get(email_id=record["id"])
            last_event = resp.get("last_event", record.get("status", "unknown"))
            self.history_status_cache[record["id"]] = resp
            self.tree.set(item_id, "状态", last_event)
            if last_event == "scheduled":
                self.tree.set(item_id, "操作", "计划可修改，可双击此项修改")
            elif last_event == "canceled":
                self.tree.set(item_id, "操作", "发送计划已取消")
            elif last_event == "delivered":
                self.tree.set(item_id, "操作", "已成功投递")
            else:
                self.tree.set(item_id, "操作", "-")
            if last_event == 'delivered' and resp.get('created_at'):
                record['created_at'] = resp.get('created_at')
            if last_event == 'scheduled' and resp.get('scheduled_at'):
                record['scheduled_at'] = resp.get('scheduled_at')
            delivery_time = self.get_delivery_time(record, last_event, resp)
            self.tree.set(item_id, "递送/计划时间", delivery_time)
            record["status"] = last_event
        except Exception:
            pass
        self.hide_loading()

    def refresh_all(self, refresh_all):
        self.show_loading()
        for iid in self.tree.get_children():
            record = self.history_cancel_buttons.get(iid)
            if not record:
                continue
            need_refresh = refresh_all or (self.tree.set(iid, "状态") == "scheduled")
            if need_refresh:
                try:
                    resend.api_key = get_api_key()
                    resp = resend.Emails.get(email_id=record["id"])
                    last_event = resp.get("last_event", record.get("status", "unknown"))
                    self.history_status_cache[record["id"]] = resp
                    self.tree.set(iid, "状态", last_event)
                    if last_event == "scheduled":
                        self.tree.set(iid, "操作", "计划可修改，可双击此项修改")
                    elif last_event == "canceled":
                        self.tree.set(iid, "操作", "发送计划已取消")
                    elif last_event == "delivered":
                        self.tree.set(iid, "操作", "已成功投递")
                    else:
                        self.tree.set(iid, "操作", "-")
                    if last_event == 'delivered' and resp.get('created_at'):
                        record['created_at'] = resp.get('created_at')
                    if last_event == 'scheduled' and resp.get('scheduled_at'):
                        record['scheduled_at'] = resp.get('scheduled_at')
                    delivery_time = self.get_delivery_time(record, last_event, resp)
                    self.tree.set(iid, "递送/计划时间", delivery_time)
                    record["status"] = last_event
                except Exception:
                    pass
        self.hide_loading()
        self.title(f"邮件发送历史（列表最后更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）")

    def on_tz_select(self, tz_code):
        self.selected_tz = tz_code
        for iid in self.tree.get_children():
            record = self.history_cancel_buttons.get(iid)
            if not record:
                continue
            status = self.tree.set(iid, "状态")
            full_resp = self.history_status_cache.get(record.get('id'), {})
            delivery_time = self.get_delivery_time(record, status, full_resp)
            self.tree.set(iid, "递送/计划时间", delivery_time)
    # ...历史窗口相关方法... 