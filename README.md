# Resend 邮件客户端（Tkinter版）

## 项目简介

本项目是一个基于 Python Tkinter 的 Resend 邮件发送客户端，支持 Resend 邮件 API，适合日常邮件群发、定时投递、历史管理等场景。

**主要功能：**
- 支持 API Key 本地管理，随时切换
- 发件人、收件人、抄送、密送、回复地址等历史输入自动补全
- 邮件历史本地保存、详细信息弹窗、计划任务管理
- 支持本地小附件上传（≤40MB）和远程大文件链接导入
- 邮件发送支持立即/定时，带附件时自动禁用定时
- 所有邮箱输入框自动校验格式
- 历史窗口支持右键刷新、时区切换、计划任务修改/取消等
- 发送、刷新等操作均有Loading遮罩提示，体验流畅

## Windows下打包指南

1. 安装依赖（如未安装PyInstaller）：
   ```bash
   pip install pyinstaller
   ```
2. 确保 `email.ico` 与 `main.py` 在同一目录下。
3. 执行以下命令打包为单文件exe（无控制台窗口）：
   ```bash
   pyinstaller --onefile --noconsole --icon=email.ico --add-data "email.ico;." main.py
   ```
4. 打包完成后，exe文件在 `dist/` 目录下，无需额外ico文件即可显示自定义窗口图标。

## 安装依赖

推荐使用如下命令安装所有依赖：

```bash
pip install -r requirements.txt
```

如需支持拖拽上传附件功能，请额外安装tkinterdnd2：

```bash
pip install tkinterdnd2
```

---
如需更多功能或遇到问题，欢迎反馈！ 