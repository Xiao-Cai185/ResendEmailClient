import os
import json
CONFIG_FILE = "config.json"

class ConfigManager:
    def __init__(self):
        self.api_key = ""
        self.load_config()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    self.api_key = config.get("api_key", "")
            except Exception:
                self.api_key = ""
        else:
            self.api_key = ""

    def save_config(self, api_key):
        self.api_key = api_key
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({"api_key": api_key}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def get_api_key(self):
        return self.api_key

    def set_api_key(self, key):
        self.save_config(key)
        self.api_key = key

    def set_api_key_dialog(self, parent=None):
        import tkinter as tk
        from tkinter import simpledialog, messagebox
        root = None
        if parent is None:
            root = tk.Tk()
            root.withdraw()
            try:
                root.iconbitmap("email.ico")
            except Exception:
                pass
        else:
            root = parent
        api_key = simpledialog.askstring(
            "API Key设置",
            "请输入您的Resend API Key:",
            show='*',
            parent=root
        )
        if api_key:
            self.set_api_key(api_key)
            if parent:
                messagebox.showinfo("成功", "API Key已更新并保存！", parent=parent)
        if parent is None and root:
            root.destroy()

# 单例
config_manager = ConfigManager()

def get_api_key():
    return config_manager.get_api_key()

def set_api_key(key):
    config_manager.set_api_key(key)

def set_api_key_dialog(parent=None):
    config_manager.set_api_key_dialog(parent=parent) 