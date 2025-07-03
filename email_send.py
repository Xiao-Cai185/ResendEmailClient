import resend
import base64
import os
from utils import validate_email, file_to_base64, is_blacklisted_attachment
from config import get_api_key

class EmailSender:
    def __init__(self):
        pass

    def build_params(self, sender_name, sender_email, to_emails, subject, html, cc_emails=None, bcc_emails=None, reply_to=None, attachments=None, scheduled_at=None):
        params = {
            "from": f"{sender_name} <{sender_email}>" if sender_name else sender_email,
            "to": to_emails,
            "subject": subject,
            "html": html
        }
        if cc_emails:
            params["cc"] = cc_emails
        if bcc_emails:
            params["bcc"] = bcc_emails
        if reply_to:
            params["reply_to"] = [reply_to]
        if attachments:
            params["attachments"] = attachments
        if scheduled_at:
            params["scheduled_at"] = scheduled_at
        return params

    def send_email(self, params):
        resend.api_key = get_api_key()
        return resend.Emails.send(params)

    def prepare_attachment(self, file_path):
        # 40MB以内转base64，否则返回None（由UI层弹窗提示并处理远程链接）
        max_size = 40 * 1024 * 1024
        size = os.path.getsize(file_path)
        filename = os.path.basename(file_path)
        if is_blacklisted_attachment(filename):
            return None  # 由UI层弹窗提示
        if size <= max_size:
            b64_content = file_to_base64(file_path)
            return {"content": b64_content, "filename": filename}
        else:
            return None

    def cancel_scheduled(self, email_id):
        resend.api_key = get_api_key()
        return resend.Emails.cancel(email_id)

    def update_scheduled(self, email_id, scheduled_at):
        resend.api_key = get_api_key()
        return resend.Emails.update(params={"id": email_id, "scheduled_at": scheduled_at})

# 单例
email_sender = EmailSender() 