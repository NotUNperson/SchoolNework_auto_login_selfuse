import os
import sys
import time
import logging
import threading
import configparser
import ctypes
import shutil
import subprocess
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

# ================= 配置文件处理模块 =================

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(application_path, 'config.ini')

def load_or_create_config():
    """读取配置，自动修复缺失条目"""
    config = configparser.ConfigParser(interpolation=None) 
    config_needs_save = False
    is_first_run = False
    
    default_config = {
        'Account': {
            'username': '',
            'password': ''
        },
        'Settings': {
            'logout_btn_id': 'logout',
            'login_url': 'http://192.168.57.33',
            'check_interval': '45',
            'quick_retry_times': '2',
            'auto_install': 'false'  # 新增：默认不自动潜入，留在原地运行
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
        sys.exit(0)
        
    username = config.get('Account', 'username', fallback='').strip()
    password = config.get('Account', 'password', fallback='').strip()
    if not username or not password:
        ctypes.windll.user32.MessageBoxW(0, "账号或密码为空！\n请打开 config.ini 填写完整。", "配置错误", 0x10)
        sys.exit(0)
        
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
            logging.error("无法创建安装目录，权限不足，取消安装")
            if icon_to_stop:
                ctypes.windll.user32.MessageBoxW(0, "权限不足，无法安装到电脑本机！", "错误", 0x10)
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
            exit_action(icon_to_stop, None) # 调用优雅退出销毁当前进程
        else:
            sys.exit(0)
            
    except Exception as e:
        logging.error(f"安装过程中发生错误: {e}")
        if icon_to_stop:
            ctypes.windll.user32.MessageBoxW(0, f"安装失败: {e}", "错误", 0x10)
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
        
        self.internal_test_url = "192.168.55.66"
        self.external_test_url = "www.baidu.com"
        self.max_retries = 3
        self.retry_delay = 180

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
        # 1. 先销毁所有 webview 窗口，触发 pywebview 的退出流程
        if webview.windows:
            for window in webview.windows:
                window.destroy()
        
        # 2. 停止托盘图标
        icon.stop()
        
    except Exception as e:
        logging.error(f"清理资源时发生异常: {e}")
    finally:
        # 3. 使用 sys.exit() 替代 os._exit()，让 Python 有机会清理现场
        # 如果是在子线程中，可能需要用到更复杂的退出策略
        import sys
        logging.info("程序已退出")
        os._exit(0) # 只有在 sys.exit 失效且确认需要强制杀掉进程时才用它

def manual_install_action(icon, item):
    """托盘菜单点击事件：手动触发安装"""
    perform_installation(icon_to_stop=icon)

def setup_tray():
    # 动态生成菜单：如果在 U 盘里，就加上安装选项
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