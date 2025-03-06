import os
import sys
import time
import logging
import threading
import json
import ctypes
import winreg
from datetime import datetime
from pathlib import Path

import smtplib
from email.mime.text import MIMEText
from email.header import Header
import requests

import tkinter as tk
from tkinter import ttk, messagebox
import pystray
from PIL import Image, ImageDraw

# 确定程序运行方式
def get_run_mode():
    if getattr(sys, 'frozen', False):
        return 'exe'
    else:
        return 'py'

# 设置日志系统
def get_app_dir():
    # 获取应用程序目录
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        return os.path.dirname(sys.executable)
    else:
        # 如果是Python脚本
        return os.path.dirname(os.path.abspath(__file__))

def setup_logging():
    # 设置日志系统
    log_dir = Path(get_app_dir()) / 'logs'
    log_dir.mkdir(exist_ok=True)
    
    log_file = log_dir / f"pc_notifier_{datetime.now().strftime('%Y%m%d')}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    run_mode = get_run_mode()
    logging.info(f"程序以 {run_mode} 方式启动")

class Config:
    def __init__(self):
        self.config_path = Path(get_app_dir()) / 'config.json'
        self.default_config = {
            'startup_enabled': False,
            'shutdown_enabled': False,
            'notification_method': 'bark',  # 'bark' or 'email'
            'bark': {
                'server_url': '',
                'device_key': ''
            },
            'email': {
                'smtp_server': '',
                'smtp_port': 465,
                'sender': '',
                'password': '',
                'receiver': ''
            }
        }
        self.config = self.load_config()
    
    def load_config(self):
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                logging.info("配置文件加载成功")
                return config
            except Exception as e:
                logging.error(f"加载配置文件失败: {e}")
                return self.default_config
        else:
            logging.info("配置文件不存在，使用默认配置")
            return self.default_config
    
    def save_config(self):
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            logging.info("配置保存成功")
            return True
        except Exception as e:
            logging.error(f"保存配置失败: {e}")
            return False

# 启动项管理
class StartupManager:
    def __init__(self):
        self.app_name = "PC_Notifier"
        self.run_mode = get_run_mode()
        if self.run_mode == 'exe':
            self.app_path = os.path.abspath(sys.executable)
        else:
            self.app_path = f'pythonw "{os.path.abspath(__file__)}"'
    
    def add_to_startup(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, self.app_name, 0, winreg.REG_SZ, self.app_path)
            winreg.CloseKey(key)
            logging.info("已添加到开机启动项")
            return True
        except Exception as e:
            logging.error(f"添加开机启动项失败: {e}")
            return False
    
    def remove_from_startup(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, self.app_name)
            winreg.CloseKey(key)
            logging.info("已从开机启动项移除")
            return True
        except Exception as e:
            logging.error(f"移除开机启动项失败: {e}")
            return False
    
    def check_startup_status(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            try:
                value, _ = winreg.QueryValueEx(key, self.app_name)
                winreg.CloseKey(key)
                return True
            except:
                winreg.CloseKey(key)
                return False
        except Exception as e:
            logging.error(f"检查开机启动项状态失败: {e}")
            return False

# 消息推送
class Notifier:
    def __init__(self, config):
        self.config = config
    
    def send_bark_notification(self, title, content):
        try:
            bark_url = self.config['bark']['server_url']
            device_key = self.config['bark']['device_key']
            
            if not bark_url or not device_key:
                logging.error("Bark配置不完整")
                return False
            
            # 确保URL格式正确
            if not bark_url.endswith('/'):
                bark_url += '/'
            
            url = f"{bark_url}{device_key}/{title}/{content}"
            response = requests.get(url)
            
            if response.status_code == 200:
                logging.info("Bark消息发送成功")
                return True
            else:
                logging.error(f"Bark消息发送失败: {response.status_code}")
                return False
        except Exception as e:
            logging.error(f"Bark消息发送异常: {e}")
            return False
    
    def send_email_notification(self, title, content):
        try:
            email_config = self.config['email']
            
            if not all([email_config['smtp_server'], email_config['sender'], 
                       email_config['password'], email_config['receiver']]):
                logging.error("邮件配置不完整")
                return False
            
            message = MIMEText(content, 'plain', 'utf-8')
            message['From'] = Header(email_config['sender'])
            message['To'] = Header(email_config['receiver'])
            message['Subject'] = Header(title)
            
            server = smtplib.SMTP_SSL(email_config['smtp_server'], email_config['smtp_port'])
            server.login(email_config['sender'], email_config['password'])
            server.sendmail(email_config['sender'], [email_config['receiver']], message.as_string())
            server.quit()
            
            logging.info("邮件发送成功")
            return True
        except Exception as e:
            logging.error(f"邮件发送异常: {e}")
            return False
    
    def send_notification(self, title, content):
        if self.config['notification_method'] == 'bark':
            return self.send_bark_notification(title, content)
        elif self.config['notification_method'] == 'email':
            return self.send_email_notification(title, content)
        else:
            logging.error(f"不支持的通知方式: {self.config['notification_method']}")
            return False

# 关机监听
class ShutdownListener:
    # 类变量，用于跟踪窗口类是否已注册
    class_registered = False
    hwnd = None
    
    def __init__(self, notifier, config):
        self.notifier = notifier
        self.config = config
        self.is_running = False
        self.thread = None
    
    def start(self):
        if not self.is_running and self.config['shutdown_enabled']:
            self.is_running = True
            self.thread = threading.Thread(target=self._listen_for_shutdown)
            self.thread.daemon = True
            self.thread.start()
            logging.info("关机监听已启动")
    
    def stop(self):
        self.is_running = False
        logging.info("关机监听已停止")
    
    def _listen_for_shutdown(self):
        # 注册关机事件处理
        try:
            import win32api
            import win32con
            import win32gui
            
            def handle_system_command(hwnd, msg, wparam, lparam):
                if msg == win32con.WM_QUERYENDSESSION:
                    logging.info(f"检测到系统关机事件，消息ID: {msg}")
                    # 根据wParam判断是否为重启事件
                    is_restart = (wparam & 0x00000040) == 0x00000040
                    self._send_shutdown_notification(is_restart=is_restart)
                    return 0
                elif msg == win32con.WM_ENDSESSION:
                    # 忽略WM_ENDSESSION消息，避免重复通知
                    logging.info(f"收到WM_ENDSESSION消息，已忽略，消息ID: {msg}")
                    return 0
                elif msg == win32con.WM_SYSCOMMAND:
                    if wparam == win32con.SC_CLOSE:
                        logging.info("检测到窗口关闭事件")
                        return 0
                return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
            
            # 检查窗口类是否已注册
            if not ShutdownListener.class_registered:
                try:
                    # 尝试注册窗口类
                    wc = win32gui.WNDCLASS()
                    wc.lpfnWndProc = handle_system_command
                    wc.lpszClassName = "ShutdownListener"
                    wc.hInstance = win32api.GetModuleHandle(None)
                    
                    class_atom = win32gui.RegisterClass(wc)
                    ShutdownListener.class_registered = True
                    logging.info("成功注册关机监听窗口类")
                except Exception as e:
                    # 如果注册失败，可能是类已存在
                    if "类已存在" in str(e):
                        logging.info("关机监听窗口类已存在，继续使用")
                        ShutdownListener.class_registered = True
                    else:
                        raise e
            
            # 创建窗口（如果窗口已存在，则不创建新窗口）
            if ShutdownListener.hwnd is None or not win32gui.IsWindow(ShutdownListener.hwnd):
                ShutdownListener.hwnd = win32gui.CreateWindow(
                    "ShutdownListener", "ShutdownListener", 0, 0, 0, 0, 0, 0, 0, 
                    win32api.GetModuleHandle(None), None
                )
                logging.info(f"创建关机监听窗口，句柄: {ShutdownListener.hwnd}")
            
            while self.is_running:
                win32gui.PumpWaitingMessages()
                time.sleep(0.1)
                
        except ImportError:
            logging.error("无法导入win32api模块，使用备用方法监听关机事件")
            self._fallback_shutdown_listener()
    
    def _fallback_shutdown_listener(self):
        # 备用方法：使用WMI监听关机事件
        try:
            import wmi
            c = wmi.WMI()
            
            watcher = c.Win32_ComputerShutdownEvent.watch_for()
            while self.is_running:
                shutdown_event = watcher(timeout_ms=100)
                if shutdown_event:
                    logging.info("检测到系统关机事件(WMI)")
                    self._send_shutdown_notification()
                time.sleep(0.1)
        except ImportError:
            logging.error("无法导入wmi模块，使用最后备用方法监听关机事件")
            self._last_resort_shutdown_listener()
    
    def _last_resort_shutdown_listener(self):
        # 最后备用方法：使用系统钩子
        try:
            from ctypes import wintypes
            import win32con
            
            # 定义回调函数类型
            PHANDLER_ROUTINE = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.DWORD)
            
            # 设置控制处理函数
            SetConsoleCtrlHandler = ctypes.windll.kernel32.SetConsoleCtrlHandler
            SetConsoleCtrlHandler.argtypes = (PHANDLER_ROUTINE, wintypes.BOOL)
            SetConsoleCtrlHandler.restype = wintypes.BOOL
            
            @PHANDLER_ROUTINE
            def handle_shutdown(ctrl_type):
                if ctrl_type in (win32con.CTRL_SHUTDOWN_EVENT, win32con.CTRL_LOGOFF_EVENT):
                    logging.info(f"检测到系统事件: {ctrl_type}")
                    self._send_shutdown_notification()
                    return True
                return False
            
            SetConsoleCtrlHandler(handle_shutdown, True)
            
            # 保持线程运行
            while self.is_running:
                time.sleep(0.5)
                
        except Exception as e:
            logging.error(f"设置关机监听失败: {e}")
    
    def _send_shutdown_notification(self, is_restart=False):
        if self.config['shutdown_enabled']:
            computer_name = os.environ.get('COMPUTERNAME', '未知电脑')
            title = f"{computer_name} {'重启' if is_restart else '关机'}通知"
            content = f"电脑 {computer_name} 正在{'重启' if is_restart else '关机'}，时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
            # 尝试发送通知
            success = self.notifier.send_notification(title, content)
            if success:
                logging.info("关机通知发送成功")
            else:
                logging.error("关机通知发送失败")

# GUI界面
class NotifierApp:
    def __init__(self, root=None):
        # 检查是否已有实例运行
        if not self._ensure_single_instance():
            sys.exit(0)
            
        # 初始化日志
        setup_logging()
        
        # 加载配置
        self.config_manager = Config()
        self.config = self.config_manager.config
        
        # 初始化启动项管理器
        self.startup_manager = StartupManager()
        
        # 初始化通知器
        self.notifier = Notifier(self.config)
        
        # 初始化关机监听器
        self.shutdown_listener = ShutdownListener(self.notifier, self.config)
        
        # 检查是否是开机启动
        self.is_startup_launch = self._check_startup_launch()
        
        # 创建GUI
        self.root = root if root else tk.Tk()
        self.root.title("电脑开关机通知")
        self.root.geometry("500x450")
        self.root.resizable(False, False)
        
        # 设置图标
        self.icon_path = self._create_icon()
        
        # 创建托盘图标
        self.tray_icon = None
        
        # 创建界面
        self._create_ui()
        
        # 设置关闭事件处理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # 如果是开机启动，则直接最小化到托盘
        if self.is_startup_launch and self.config['startup_enabled']:
            self._send_startup_notification()
            self.root.withdraw()
            self._create_tray_icon()
        
        # 启动关机监听
        if self.config['shutdown_enabled']:
            self.shutdown_listener.start()
    
    def _ensure_single_instance(self):
        # 确保只有一个实例运行
        try:
            # 创建互斥锁
            mutex_name = "PC_Notifier_Mutex"
            self.mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
            last_error = ctypes.windll.kernel32.GetLastError()
            
            if last_error == 183:  # ERROR_ALREADY_EXISTS
                logging.warning("程序已经在运行中")
                messagebox.showwarning("警告", "程序已经在运行中！")
                return False
            return True
        except Exception as e:
            logging.error(f"检查程序实例失败: {e}")
            return True  # 出错时允许运行
    
    def _check_startup_launch(self):
        # 检查是否是开机启动
        import psutil
        try:
            current_process = psutil.Process()
            current_process_create_time = current_process.create_time()
            boot_time = psutil.boot_time()
            
            # 如果程序启动时间与系统启动时间相差不到2分钟，认为是开机启动
            return (current_process_create_time - boot_time) < 120
        except Exception as e:
            logging.error(f"检查开机启动失败: {e}")
            return False
    
    def _create_icon(self):
        # 创建应用图标
        icon_dir = Path(os.path.dirname(os.path.abspath(__file__))) / 'icons'
        icon_dir.mkdir(exist_ok=True)
        
        icon_path = icon_dir / 'app_icon.png'
        
        if not icon_path.exists():
            # 创建一个简单的图标
            img = Image.new('RGB', (64, 64), color=(73, 109, 137))
            d = ImageDraw.Draw(img)
            d.text((10, 10), "PC", fill=(255, 255, 255))
            d.text((10, 30), "通知", fill=(255, 255, 255))
            img.save(icon_path)
            logging.info(f"创建应用图标: {icon_path}")
        
        return icon_path
    
    def _create_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建选项卡
        tab_control = ttk.Notebook(main_frame)
        
        # 基本设置选项卡
        basic_tab = ttk.Frame(tab_control)
        tab_control.add(basic_tab, text="基本设置")
        
        # Bark设置选项卡
        bark_tab = ttk.Frame(tab_control)
        tab_control.add(bark_tab, text="Bark设置")
        
        # 邮件设置选项卡
        email_tab = ttk.Frame(tab_control)
        tab_control.add(email_tab, text="邮件设置")
        
        tab_control.pack(expand=True, fill=tk.BOTH)
        
        # 基本设置界面
        self._create_basic_settings(basic_tab)
        
        # Bark设置界面
        self._create_bark_settings(bark_tab)
        
        # 邮件设置界面
        self._create_email_settings(email_tab)
        
        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(bottom_frame, text="测试推送", command=self._test_notification).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="保存配置", command=self._save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="最小化到托盘", command=self._minimize_to_tray).pack(side=tk.RIGHT, padx=5)
    
    def _create_basic_settings(self, parent):
        # 基本设置界面
        frame = ttk.LabelFrame(parent, text="通知设置", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 开机通知设置
        self.startup_var = tk.BooleanVar(value=self.config['startup_enabled'])
        ttk.Checkbutton(frame, text="启用开机通知", variable=self.startup_var).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # 检查开机启动状态按钮
        ttk.Button(frame, text="检查开机启动状态", command=self._check_startup).grid(row=0, column=1, padx=10)
        
        # 关机通知设置
        self.shutdown_var = tk.BooleanVar(value=self.config['shutdown_enabled'])
        ttk.Checkbutton(frame, text="启用关机通知", variable=self.shutdown_var).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # 推送方式选择
        ttk.Label(frame, text="推送方式:").grid(row=2, column=0, sticky=tk.W, pady=10)
        
        self.notification_method_var = tk.StringVar(value=self.config['notification_method'])
        ttk.Radiobutton(frame, text="Bark", variable=self.notification_method_var, value="bark").grid(row=3, column=0, sticky=tk.W)
        ttk.Radiobutton(frame, text="邮件", variable=self.notification_method_var, value="email").grid(row=4, column=0, sticky=tk.W)
    
    def _create_bark_settings(self, parent):
        # Bark设置界面
        frame = ttk.LabelFrame(parent, text="Bark推送设置", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 服务器URL
        ttk.Label(frame, text="服务器URL:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.bark_server_var = tk.StringVar(value=self.config['bark']['server_url'])
        ttk.Entry(frame, textvariable=self.bark_server_var, width=40).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # 设备Key
        ttk.Label(frame, text="设备Key:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.bark_key_var = tk.StringVar(value=self.config['bark']['device_key'])
        ttk.Entry(frame, textvariable=self.bark_key_var, width=40).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # 说明
        ttk.Label(frame, text="说明: Bark是一款iOS应用，用于接收自定义通知。\n服务器URL格式如: https://api.day.app/").grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=10)
    
    def _create_email_settings(self, parent):
        # 邮件设置界面
        frame = ttk.LabelFrame(parent, text="邮件推送设置", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # SMTP服务器
        ttk.Label(frame, text="SMTP服务器:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.smtp_server_var = tk.StringVar(value=self.config['email']['smtp_server'])
        ttk.Entry(frame, textvariable=self.smtp_server_var, width=30).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # SMTP端口
        ttk.Label(frame, text="SMTP端口:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.smtp_port_var = tk.IntVar(value=self.config['email']['smtp_port'])
        ttk.Entry(frame, textvariable=self.smtp_port_var, width=10).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # 发件人邮箱
        ttk.Label(frame, text="发件人邮箱:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.sender_var = tk.StringVar(value=self.config['email']['sender'])
        ttk.Entry(frame, textvariable=self.sender_var, width=30).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # 邮箱密码/授权码
        ttk.Label(frame, text="邮箱密码/授权码:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar(value=self.config['email']['password'])
        password_entry = ttk.Entry(frame, textvariable=self.password_var, width=30, show="*")
        password_entry.grid(row=3, column=1, sticky=tk.W, padx=5)
        
        # 收件人邮箱
        ttk.Label(frame, text="收件人邮箱:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.receiver_var = tk.StringVar(value=self.config['email']['receiver'])
        ttk.Entry(frame, textvariable=self.receiver_var, width=30).grid(row=4, column=1, sticky=tk.W, padx=5)
    
    def _create_tray_icon(self):
        # 创建托盘图标
        if self.tray_icon is None:
            icon = Image.open(self.icon_path)
            menu = (pystray.MenuItem('显示', self._show_window),
                    pystray.MenuItem('退出', self._quit_app))
            self.tray_icon = pystray.Icon("PC_Notifier", icon, "电脑开关机通知", menu)
            self.tray_icon.on_double_click = self._show_window  # 添加双击事件处理
            self.tray_icon.run_detached()
            logging.info("创建托盘图标")
    
    def _show_window(self, icon=None, item=None):
        # 显示主窗口
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        logging.info("显示主窗口")
    
    def _minimize_to_tray(self):
        # 最小化到托盘
        self.root.withdraw()
        self._create_tray_icon()
        logging.info("最小化到托盘")
    
    def _on_close(self):
        # 关闭窗口事件
        self._minimize_to_tray()
    
    def _quit_app(self, icon=None, item=None):
        # 退出应用
        if self.tray_icon:
            self.tray_icon.stop()
        
        # 停止关机监听
        self.shutdown_listener.stop()
        
        logging.info("程序退出")
        # 使用after方法确保在主线程中执行销毁操作
        self.root.after(0, self._safe_destroy)
    
    def _safe_destroy(self):
        # 在主线程中安全地销毁窗口
        try:
            self.root.quit()
            self.root.destroy()
        except Exception as e:
            logging.error(f"销毁窗口时出错: {e}")
    
    def _check_startup(self):
        # 检查开机启动状态
        status = self.startup_manager.check_startup_status()
        if status:
            messagebox.showinfo("开机启动状态", "已设置为开机启动")
        else:
            messagebox.showinfo("开机启动状态", "未设置开机启动")
        logging.info(f"检查开机启动状态: {status}")
    
    def _save_settings(self):
        # 保存设置
        # 更新配置
        self.config['startup_enabled'] = self.startup_var.get()
        self.config['shutdown_enabled'] = self.shutdown_var.get()
        self.config['notification_method'] = self.notification_method_var.get()
        
        # 更新Bark设置
        self.config['bark']['server_url'] = self.bark_server_var.get()
        self.config['bark']['device_key'] = self.bark_key_var.get()
        
        # 更新邮件设置
        self.config['email']['smtp_server'] = self.smtp_server_var.get()
        self.config['email']['smtp_port'] = self.smtp_port_var.get()
        self.config['email']['sender'] = self.sender_var.get()
        self.config['email']['password'] = self.password_var.get()
        self.config['email']['receiver'] = self.receiver_var.get()
        
        # 保存配置
        if self.config_manager.save_config():
            messagebox.showinfo("保存成功", "配置已保存")
            
            # 根据开机启动设置更新注册表
            if self.config['startup_enabled']:
                self.startup_manager.add_to_startup()
            else:
                self.startup_manager.remove_from_startup()
            
            # 更新关机监听状态
            if self.config['shutdown_enabled']:
                self.shutdown_listener.start()
            else:
                self.shutdown_listener.stop()
                
            logging.info("配置已更新并保存")
        else:
            messagebox.showerror("保存失败", "配置保存失败")
    
    def _test_notification(self):
        # 测试推送
        computer_name = os.environ.get('COMPUTERNAME', '未知电脑')
        title = f"{computer_name} 测试通知"
        content = f"这是一条测试消息，发送时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        # 临时使用当前界面的配置进行测试
        test_config = self.config.copy()
        test_config['notification_method'] = self.notification_method_var.get()
        
        # 更新Bark设置
        test_config['bark']['server_url'] = self.bark_server_var.get()
        test_config['bark']['device_key'] = self.bark_key_var.get()
        
        # 更新邮件设置
        test_config['email']['smtp_server'] = self.smtp_server_var.get()
        test_config['email']['smtp_port'] = self.smtp_port_var.get()
        test_config['email']['sender'] = self.sender_var.get()
        test_config['email']['password'] = self.password_var.get()
        test_config['email']['receiver'] = self.receiver_var.get()
        
        # 创建临时通知器
        test_notifier = Notifier(test_config)
        
        # 发送测试通知
        success = test_notifier.send_notification(title, content)
        
        if success:
            messagebox.showinfo("测试成功", "测试消息发送成功")
            logging.info("测试消息发送成功")
        else:
            messagebox.showerror("测试失败", "测试消息发送失败，请检查配置")
            logging.error("测试消息发送失败")
    
    def _send_startup_notification(self):
        # 发送开机通知
        if self.config['startup_enabled']:
            computer_name = os.environ.get('COMPUTERNAME', '未知电脑')
            title = f"{computer_name} 开机通知"
            content = f"电脑 {computer_name} 已开机，时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
            # 尝试发送通知
            success = self.notifier.send_notification(title, content)
            if success:
                logging.info("开机通知发送成功")
            else:
                logging.error("开机通知发送失败")
    
    def run(self):
        # 运行应用
        self.root.mainloop()

# 程序入口
def main():
    app = NotifierApp()
    app.run()

if __name__ == "__main__":
    main()