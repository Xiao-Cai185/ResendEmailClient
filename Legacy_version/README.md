# Legacy Version 说明

 `main_legacy.py` 为本项目重构前的旧版本实现，包含了所有功能的单文件实现。

主要特点：
- 所有UI、逻辑、历史、配置等均在一个文件中实现，结构较为混杂。
- 支持基本的邮件发送、历史记录、附件管理等功能。
- 不支持模块化、分层管理，维护和扩展较为困难。

新版已将各功能模块拆分，结构更清晰，体验和健壮性大幅提升。 

运行方式：
```bash
python main_legacy.py
```

推荐使用如下命令安装所有依赖：

```bash
pip install -r requirements.txt
```