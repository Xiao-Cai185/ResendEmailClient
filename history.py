import os
import json
HISTORY_FILE = "history.json"
EMAIL_HISTORY_FILE = "email_history.json"

class HistoryManager:
    def __init__(self):
        self.input_history = {
            "sender_names": [],
            "sender_emails": [],
            "recipient_emails": []
        }
        self.email_history = []
        self.load_input_history()
        self.load_email_history()

    # 输入历史相关
    def load_input_history(self):
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    self.input_history = json.load(f)
            except Exception:
                pass

    def save_input_history(self):
        try:
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self.input_history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def add_to_history(self, key, value):
        if value and value not in self.input_history[key]:
            self.input_history[key].insert(0, value)
            self.input_history[key] = self.input_history[key][:20]
            self.save_input_history()

    def remove_from_history(self, key, value):
        if value in self.input_history[key]:
            self.input_history[key].remove(value)
            self.save_input_history()

    def clear_history(self, key):
        self.input_history[key] = []
        self.save_input_history()

    # 邮件历史相关
    def load_email_history(self):
        if os.path.exists(EMAIL_HISTORY_FILE):
            try:
                with open(EMAIL_HISTORY_FILE, "r", encoding="utf-8") as f:
                    self.email_history = json.load(f)
            except Exception:
                self.email_history = []
        else:
            self.email_history = []

    def save_email_history(self):
        try:
            with open(EMAIL_HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self.email_history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def add_email_record(self, record):
        self.email_history.append(record)
        self.save_email_history()

    def get_email_history(self):
        return self.email_history

# 单例
history_manager = HistoryManager()

def get_input_history():
    return history_manager.input_history

def add_input_history(key, value):
    history_manager.add_to_history(key, value)

def remove_input_history(key, value):
    history_manager.remove_from_history(key, value)

def clear_input_history(key):
    history_manager.clear_history(key)

def get_email_history():
    return history_manager.get_email_history()

def add_email_record(record):
    history_manager.add_email_record(record) 