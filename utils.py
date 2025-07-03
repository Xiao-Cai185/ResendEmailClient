import re
import sys
import os
import base64
from datetime import datetime, timedelta, timezone
try:
    from zoneinfo import ZoneInfo
except ImportError:
    from pytz import timezone as ZoneInfo
import tzlocal
import tkinter as tk
from tkinter import messagebox

EMAIL_REGEX = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")

ATTACHMENT_BLACKLIST = {
    '.adp','.app','.asp','.bas','.bat','.cer','.chm','.cmd','.com','.cpl','.crt','.csh','.der','.exe','.fxp','.gadget','.hlp','.hta','.inf','.ins','.isp','.its','.js','.jse','.ksh','.lib','.lnk','.mad','.maf','.mag','.mam','.maq','.mar','.mas','.mat','.mau','.mav','.maw','.mda','.mdb','.mde','.mdt','.mdw','.mdz','.msc','.msh','.msh1','.msh2','.mshxml','.msh1xml','.msh2xml','.msi','.msp','.mst','.ops','.pcd','.pif','.plg','.prf','.prg','.reg','.scf','.scr','.sct','.shb','.shs','.sys','.ps1','.ps1xml','.ps2','.ps2xml','.psc1','.psc2','.tmp','.url','.vb','.vbe','.vbs','.vps','.vsmacros','.vss','.vst','.vsw','.vxd','.ws','.wsc','.wsf','.wsh','.xnk'
}

def validate_email(email):
    return EMAIL_REGEX.match(email)

def get_resource_path(filename):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return filename

def file_to_base64(filepath):
    with open(filepath, "rb") as f:
        return base64.b64encode(f.read()).decode()

def is_blacklisted_attachment(filename):
    ext = os.path.splitext(filename)[1].lower()
    return ext in ATTACHMENT_BLACKLIST

def show_error(msg, parent=None):
    messagebox.showerror("错误", msg, parent=parent)

def show_info(msg, parent=None):
    messagebox.showinfo("提示", msg, parent=parent)

def format_time(timestr, tz_code='local'):
    if not timestr or timestr == '-':
        return '-'
    try:
        if '+' in timestr:
            dt = datetime.fromisoformat(timestr.replace(' ', 'T'))
        else:
            dt = datetime.fromisoformat(timestr)
        if tz_code == 'local':
            tz = tzlocal.get_localzone()
        elif tz_code.startswith('UTC') and tz_code != 'UTC':
            hours = float(tz_code[3:])
            tz = timezone(timedelta(hours=hours))
        elif tz_code == 'UTC':
            tz = timezone.utc
        else:
            tz = ZoneInfo(tz_code)
        if hasattr(dt, 'astimezone'):
            dt = dt.astimezone(tz)
        return dt.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return timestr 