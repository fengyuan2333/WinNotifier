# PC Notifier 电脑开关机通知程序

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.6%2B-blue)](https://www.python.org/)

一个基于Python开发的Windows系统开关机通知程序，可以在电脑开机或关机时自动发送通知到指定的设备或邮箱。

## ✨ 功能特点

- 🔔 开机自动发送通知
- 🛑 关机自动发送通知
- 📱 支持Bark推送（iOS设备）
- 📧 支持邮件推送
- 🖥️ 图形界面配置，操作简单
- 🔄 支持最小化到系统托盘
- 🚀 自动添加开机启动项
- 📝 完整的日志记录系统
- 🔒 防止多实例运行

## 🛠️ 技术栈

- Python 3.6+
- tkinter (GUI界面)
- pystray (系统托盘)
- Pillow (图像处理)
- requests (网络请求)
- psutil (系统信息)
- pywin32 (Windows API)
- wmi (Windows管理规范)

## 📦 安装

### 方式一：直接下载

1. 从[Releases](../../releases)页面下载最新版本的可执行文件
2. 解压后直接运行`main.exe`

### 方式二：从源码安装

1. 克隆仓库：
```bash
git clone https://github.com/fengyuan2333/WinNotifier.git
cd pc-notifier
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

3. 运行程序：
```bash
python main.py
```

## 🚀 使用方法

1. 首次运行程序会显示配置界面
2. 在基本设置中选择是否启用开机通知和关机通知
3. 选择并配置推送方式：

   ### Bark推送配置
   - 服务器URL（例如：https://api.day.app/）
   - Token（在IOS Bark App中获取）

   ### 邮件推送配置
   - SMTP服务器地址
   - SMTP端口（通常为465，使用SSL）
   - 发件人邮箱
   - 邮箱密码/授权码
   - 收件人邮箱

4. 点击「测试推送」确认配置是否正确
5. 点击「保存配置」保存设置
6. 关闭窗口后程序会自动最小化到系统托盘

## 🔨 开发打包

### 使用PyInstaller打包

1. 安装PyInstaller：
```bash
pip install pyinstaller
```

2. 执行打包命令：
```bash
pyinstaller --noconsole --onefile --icon=icons/app_icon.png --add-data=icons/app_icon.png:icons --add-data=config.json:. --add-data=logs:logs main.py

```

打包参数说明：
- `--noconsole`：隐藏控制台窗口
- `--onefile`：打包为单个exe文件
- `--icon`：设置程序图标
- `--add-data`：包含资源文件

## 📝 注意事项

- 程序会在同目录下创建`config.json`配置文件和`logs`目录
- 启用开机通知功能后会自动添加到Windows开机启动项
- 程序采用互斥锁机制防止多个实例同时运行
- 关机监听支持多种实现方式，自动适配不同Windows系统版本

## 🤝 贡献

欢迎提交问题和改进建议！提交PR前请确保：

1. Fork本仓库
2. 创建特性分支
3. 提交更改
4. 推送到分支
5. 提交Pull Request

## 📄 开源协议

本项目采用MIT协议开源，详见[LICENSE](LICENSE)文件。
