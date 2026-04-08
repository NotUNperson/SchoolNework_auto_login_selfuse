import os
import sys
import time
import logging
import threading
import configparser
import ctypes
import shutil
import subprocess
from datetime import datetime
import webview
from ping3 import ping
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw

# ================= 全局日志初始化 =================
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    datefmt='%Y.%m.%d %H:%M:%S'
)

# ================= 环境隔离与延迟退出机制 =================
if getattr(sys, 'frozen', False):
    SCRIPT_PATH = sys.executable
else:
    SCRIPT_PATH = os.path.abspath(__file__)

WORK_DIR = os.path.dirname(SCRIPT_PATH)
CONFIG_FILE = os.path.join(WORK_DIR, 'config.ini')
APP_SIGNATURE = "campusnet_autologin_v1" # 程序的唯一标识

def delayed_exit(seconds=15, exit_code=1):
    """通用的延迟退出函数，防止控制台瞬间闪退导致报错无法阅读"""
    print(f"\n⏳ 为了让您看清以上提示，程序将在 {seconds} 秒后自动退出...")
    time.sleep(seconds)
    sys.exit(exit_code)

def migrate_and_exit(reason):
    """环境隔离核心逻辑：复制自身到新文件夹并退出"""
    print(f"\n❌ 环境异常拦截: {reason}")
    
    base_dir_name = "CampusNet_Isolated_Env"
    target_dir = os.path.join(WORK_DIR, base_dir_name)
    
    # 如果默认名称的文件夹已存在，则加上时间戳后缀
    if os.path.exists(target_dir):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        target_dir = os.path.join(WORK_DIR, f"{base_dir_name}_{timestamp}")
        
    try:
        os.makedirs(target_dir, exist_ok=True)
        new_script_path = os.path.join(target_dir, os.path.basename(SCRIPT_PATH))
        
        # 使用 copy2 避免 Windows 运行中文件锁定的限制
        shutil.copy2(SCRIPT_PATH, new_script_path)
        print(f"✅ 已自动为您创建了纯净的隔离文件夹：{target_dir}")
        print(f"👉 程序已自动将自身复制过去。请进入该文件夹运行新的程序，当前程序即将退出！")
        print(f"💡 提示：您可以随后安全地删除当前目录下的旧程序。")
    except Exception as e:
        print(f"⚠️ 尝试自动隔离环境失败 ({e})，请手动将本程序移动到一个新建的空白文件夹中运行！")
        
    delayed_exit(15)

# === A. 检查目录读写权限 ===
test_file = os.path.join(WORK_DIR, '.permission_test')
try:
    with open(test_file, 'w') as f:
        f.write('test')
    os.remove(test_file)
except PermissionError:
    print(f"❌ 致命错误：程序在当前目录 ({WORK_DIR}) 没有读写权限！无法生成配置文件。请更换目录。")
    delayed_exit(15)

# === B. 检查配置文件归属及目录洁癖 ===
if os.path.exists(CONFIG_FILE):
    # 存在配置文件，检查是不是自己的
    temp_config = configparser.ConfigParser()
    try:
        temp_config.read(CONFIG_FILE, encoding='utf-8')
        if not temp_config.has_section('System') or temp_config.get('System', 'app_signature', fallback='') != APP_SIGNATURE:
            migrate_and_exit("检测到当前目录下的 config.ini 并非本程序的配置，为防止污染您的其他项目，必须隔离运行。")
    except Exception:
        migrate_and_exit("检测到当前目录下的 config.ini 格式损坏或属于其他程序，必须隔离运行。")
else:
    # 不存在配置文件，初次运行，检查目录是否杂乱
    allowed_items = {os.path.basename(SCRIPT_PATH), 'config.ini', '__pycache__', 'build', 'dist'}
    try:
        current_items = set(os.listdir(WORK_DIR))
        unrelated_items = current_items - allowed_items
        if len(unrelated_items) > 2:
            migrate_and_exit("检测到当前文件夹内存在较多杂乱文件，为防止生成的配置文件弄脏您的文件夹，必须隔离运行。")
    except Exception:
        pass


# ================= 配置文件处理模块 =================

def load_or_create_config():
    """读取配置，自动修复缺失条目"""
    config = configparser.ConfigParser(interpolation=None) 
    config_needs_save = False
    is_first_run = False
    
    default_config = {
        'System': {
            'app_signature': APP_SIGNATURE
        },
        'Account': {
            'username': '',
            'password': ''
        },
        'Settings': {
            'logout_btn_id': 'logout',
            'login_url': '',
            'internal_test_url': '',  # 提取为可配置项
            'external_test_url': 'www.baidu.com',  # 提取为可配置项
            'check_interval': '45',
            'quick_retry_times': '2',
            'auto_install': 'false', 
            'max_retries': '3',
            'retry_delay': '180'
        }
    }
    
    if not os.path.exists(CONFIG_FILE):
        for section, options in default_config.items():
            config[section] = options
        config_needs_save = True
        is_first_run = True
    else:
        config.read(CONFIG_FILE, encoding='utf-8')
        for section, options in default_config.items():
            if not config.has_section(section):
                config.add_section(section)
                config_needs_save = True
            for key, default_val in options.items():
                if not config.has_option(section, key):
                    config.set(section, key, default_val)
                    config_needs_save = True

    if config_needs_save:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            config.write(f)
            
    if is_first_run:
        msg = f"已生成配置文件：\n{CONFIG_FILE}\n\n请打开填写账号密码后，再次运行。"
        ctypes.windll.user32.MessageBoxW(0, msg, "初次运行配置", 0x40)
        # 弹窗确认后退出，也给一点控制台保留时间
        delayed_exit(5, 0)
        
    username = config.get('Account', 'username', fallback='').strip()
    password = config.get('Account', 'password', fallback='').strip()
    if not username or not password:
        ctypes.windll.user32.MessageBoxW(0, "账号或密码为空！\n请打开 config.ini 填写完整。", "配置错误", 0x10)
        delayed_exit(5, 0)
        
    return config

app_config = load_or_create_config()

# ================= 特工潜入模块 (Dropper) =================

def is_already_installed():
    """检测当前程序是否已经位于电脑本机的潜伏基地中"""
    if not getattr(sys, 'frozen', False):
        return True # 源码运行环境当做已安装处理，不触发转移
        
    current_dir = os.path.dirname(sys.executable).lower()
    local_appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~\\AppData\\Local'))
    primary_target_dir = os.path.join(local_appdata, 'CampusNetDaemon').lower()
    fallback_target_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'data', 'CampusNetDaemon').lower()
    
    return current_dir == primary_target_dir or current_dir == fallback_target_dir

def perform_installation(icon_to_stop=None):
    """执行文件拷贝与潜伏逻辑"""
    if is_already_installed():
        if icon_to_stop:
            ctypes.windll.user32.MessageBoxW(0, "当前程序已经在电脑本机中运行，无需重复安装！", "提示", 0x40)
        return False
        
    current_exe = sys.executable
    exe_name = os.path.basename(current_exe)
    local_appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~\\AppData\\Local'))
    primary_target_dir = os.path.join(local_appdata, 'CampusNetDaemon')
    fallback_target_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'data', 'CampusNetDaemon')
    
    target_dir = primary_target_dir
    try:
        os.makedirs(target_dir, exist_ok=True)
        test_file = os.path.join(target_dir, '.test_write')
        with open(test_file, 'w') as f: f.write('1')
        os.remove(test_file)
    except Exception:
        target_dir = fallback_target_dir
        try:
            os.makedirs(target_dir, exist_ok=True)
        except Exception as e2:
            print(f"❌ 无法创建安装目录，权限不足: {e2}")
            if icon_to_stop:
                ctypes.windll.user32.MessageBoxW(0, "权限不足，无法安装到电脑本机！", "错误", 0x10)
            delayed_exit(15)
            return False

    target_exe = os.path.join(target_dir, exe_name)
    target_config = os.path.join(target_dir, 'config.ini')
    
    try:
        os.system(f'taskkill /f /im "{exe_name}" >nul 2>&1')
        time.sleep(0.5) 
        shutil.copy2(current_exe, target_exe)
        shutil.copy2(CONFIG_FILE, target_config)
        subprocess.Popen([target_exe], cwd=target_dir)
        
        logging.info(f"安装完成，程序已复制到: {target_dir}")
        if icon_to_stop:
            ctypes.windll.user32.MessageBoxW(0, f"已成功安装至: {target_dir}\n当前 U 盘进程将自动关闭，您可以安全拔出 U 盘！", "安装成功", 0x40)
            exit_action(icon_to_stop, None) 
        else:
            sys.exit(0)
            
    except Exception as e:
        print(f"❌ 安装/迁移过程中发生错误: {e}")
        if icon_to_stop:
            ctypes.windll.user32.MessageBoxW(0, f"安装失败: {e}", "错误", 0x10)
        delayed_exit(15)
        return False

# ================= 核心守护进程模块 =================

class CampusNetWebviewDaemon:
    def __init__(self):
        self.username = app_config.get('Account', 'username')
        self.password = app_config.get('Account', 'password')
        self.logout_btn_id = app_config.get('Settings', 'logout_btn_id')
        self.login_url = app_config.get('Settings', 'login_url')
        self.check_interval = app_config.getint('Settings', 'check_interval')
        self.quick_retry_times = app_config.getint('Settings', 'quick_retry_times', fallback=2)
        self.max_retries = app_config.getint('Settings', 'max_retries')
        self.retry_delay = app_config.getint('Settings', 'retry_delay')
        
        # 已动态读取配置文件中的测试地址
        self.internal_test_url = app_config.get('Settings', 'internal_test_url')
        self.external_test_url = app_config.get('Settings', 'external_test_url')

        self.is_connected = False
        self.error_count = 0
        self.success_count = 0
        self.last_error_time = 0
        self.page_loaded_event = threading.Event()

    def is_network_available(self, host):
        try:
            delay = ping(host, timeout=2)
            return delay is not None and delay is not False
        except Exception:
            return False

    def _on_page_loaded_callback(self):
        self.page_loaded_event.set()

    def execute_login(self, window):
        logging.info("启动浏览器进行登录操作")
        self.page_loaded_event.clear()  
        window.load_url(self.login_url) 

        if not self.page_loaded_event.wait(timeout=15):
            logging.error("登录页面加载超时")
            return False

        js_injector = f"""
            (function() {{
                const userInp = document.getElementById('username');
                const passInp = document.getElementById('password');
                const loginBtn = document.getElementById('login-account');
                
                if (userInp && passInp && loginBtn) {{
                    userInp.value = '{self.username}';
                    passInp.value = '{self.password}';
                    
                    const inputEvent = new Event('input', {{ bubbles: true }});
                    const changeEvent = new Event('change', {{ bubbles: true }});
                    userInp.dispatchEvent(inputEvent);
                    passInp.dispatchEvent(inputEvent);
                    userInp.dispatchEvent(changeEvent);
                    passInp.dispatchEvent(changeEvent);
                    
                    loginBtn.click();
                    return 'attempted';
                }}
                return 'not_found';
            }})();
        """
        
        attempt_result = 'not_found'
        for _ in range(10):
            attempt_result = window.evaluate_js(js_injector)
            if attempt_result == 'attempted':
                break
            time.sleep(1)
            
        if attempt_result != 'attempted':
            window.load_url('about:blank')
            return False

        check_success_js = f"document.getElementById('{self.logout_btn_id}') !== null;"
        for _ in range(12): 
            time.sleep(1)
            is_logged_in = window.evaluate_js(check_success_js)
            if is_logged_in:
                window.load_url('about:blank')
                return True

        window.load_url('about:blank')
        return False

    def daemon_worker(self, window):
        logging.info("校园网守护线程启动")
        window.events.loaded += self._on_page_loaded_callback

        while True:
            try:
                if self.is_network_available(self.external_test_url):
                    self.error_count = 0
                    if not self.is_connected:
                        self.is_connected = True
                        logging.info("网络连接正常，无需登录")
                else:
                    current_time = time.time()
                    if self.error_count >= self.max_retries:
                        if current_time - self.last_error_time < self.retry_delay:
                            time.sleep(self.check_interval)
                            continue
                        else:
                            self.error_count = 0

                    if self.is_network_available(self.internal_test_url):
                        login_success = False
                        total_attempts = self.quick_retry_times + 1 
                        
                        for attempt in range(1, total_attempts + 1):
                            logging.info(f"登录尝试 {attempt}/{total_attempts}")
                            if self.execute_login(window):
                                login_success = True
                                break 
                            else:
                                if attempt < total_attempts:
                                    time.sleep(3) 
                        
                        if login_success:
                            self.success_count += 1
                            self.error_count = 0
                            self.is_connected = True
                            logging.info(f"登录成功 (累计成功: {self.success_count} 次)")
                        else:
                            self.error_count += 1
                            self.last_error_time = current_time
                            self.is_connected = False
                            logging.warning("登录失败，进入重试等待")
                    else:
                        self.is_connected = False
                        logging.debug("内部网络不可用")
                
                time.sleep(self.check_interval)

            except Exception as e:
                logging.error(f"守护线程异常: {e}")
                time.sleep(self.check_interval)

# ================= 托盘图标与优雅退出控制区 =================

def create_tray_icon_image():
    image = Image.new('RGB', (64, 64), color=(255, 255, 255))
    draw = ImageDraw.Draw(image)
    draw.ellipse((16, 16, 48, 48), fill=(46, 204, 113))
    return image

def exit_action(icon, item):
    logging.info("收到退出指令，正在清理进程")
    
    try:
        if webview.windows:
            for window in webview.windows:
                window.destroy()
        icon.stop()
    except Exception as e:
        logging.error(f"清理资源时发生异常: {e}")
    finally:
        import sys
        logging.info("程序已退出")
        os._exit(0)

def manual_install_action(icon, item):
    perform_installation(icon_to_stop=icon)

def setup_tray():
    menu_items = []
    if not is_already_installed():
        menu_items.append(item('安装移动到data或桌面', manual_install_action))
    
    menu_items.append(item('退出守护程序', exit_action))
    
    menu = pystray.Menu(*menu_items)
    icon = pystray.Icon("CampusNetAutoLogin", create_tray_icon_image(), "校园网守护中...", menu)
    icon.run()

# ================= 主程序入口 =================

if __name__ == "__main__":
    # 1. 检测配置里的自动安装开关
    if app_config.getboolean('Settings', 'auto_install', fallback=False):
        perform_installation()
        
    # 2. 启动托盘图标
    tray_thread = threading.Thread(target=setup_tray, daemon=True)
    tray_thread.start()

    # 3. 启动网络守护主体
    daemon = CampusNetWebviewDaemon()
    
    hidden_window = webview.create_window(
        title='CampusNet_Daemon', 
        url='about:blank', 
        hidden=True, 
        min_size=(1, 1)
    )
    webview.start(daemon.daemon_worker, hidden_window)