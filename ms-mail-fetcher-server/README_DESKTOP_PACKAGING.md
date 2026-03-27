# MS Mail Fetcher Desktop 打包说明（Windows）

## 1. 准备环境
在 `ms-mail-fetcher-server` 目录安装依赖：

- `requirements.txt`（已包含 `PyInstaller` 与 `pywebview`）

## 2. 确认前端资源
先在前端项目构建，然后把前端 `dist/` 同步到后端 `template/`。后端运行时读取：

- `template/index.html`

## 3. 本地运行桌面版（开发验证）
运行：

- `python desktop_main.py`

这会：
- 启动内置 FastAPI（127.0.0.1 自动可用端口）
- 打开 pywebview 桌面窗口

## 4. 打包桌面版
使用 spec 文件打包：

- `pyinstaller ms-mail-fetcher-desktop.spec`

打包结果目录：

- `dist/ms-mail-fetcher/`

可执行文件：

- `dist/ms-mail-fetcher/ms-mail-fetcher.exe`

## 5. 发布建议
发布时至少保留：

- `ms-mail-fetcher.exe`
- `server.config.json`
- `template/`（前端静态资源，已被打进 onedir 包）

## 6. 数据库存储位置
SQLite 数据库会写入用户目录：

- `%LOCALAPPDATA%/ms-mail-fetcher/ms_mail_fetcher.db`

这样升级程序不会覆盖用户数据。

## 7. 端口配置
程序会读取可执行文件同级的 `server.config.json`，核心字段：

- `port`
- `auto_port_fallback`
- `port_retry_count`

即使默认端口被占用，也会自动尝试后续端口。
