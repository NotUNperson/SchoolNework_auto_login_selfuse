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
# 将日志初始化提到最前，确保启动的所有状态都能被记录
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
    """读取配置，自动修复缺失条目，并严格校验账号密码"""
    config = configparser.ConfigParser(interpolation=None) 
    config_needs_save = False
    is_first_run = False
    
    # 定义标准配置模板
    default_config = {
        'Account': {
            'username': '',
            'password': ''
        },
        'Settings': {
            'logout_btn_id': 'logout',
            'login_url': 'http://192.168.57.33',
            'check_interval': '45',
            'quick_retry_times': '2'
        }
    }
    
    logging.info("正在加载配置文件...")
    
    if not os.path.exists(CONFIG_FILE):
        logging.warning("未检测到 config.ini，准备生成初始配置。")
        for section, options in default_config.items():
            config[section] = options
        config_needs_save = True
        is_first_run = True
    else:
        config.read(CONFIG_FILE, encoding='utf-8')
        
        # 逐条检测缺失的 Section 和 Key，进行无损修复
        for section, options in default_config.items():
            if not config.has_section(section):
                logging.warning(f"配置缺失节 [{section}]，已自动修复。")
                config.add_section(section)
                config_needs_save = True
            for key, default_val in options.items():
                if not config.has_option(section, key):
                    logging.warning(f"配置缺失条目 [{section}] -> {key}，已填充默认值 '{default_val}'。")
                    config.set(section, key, default_val)
                    config_needs_save = True

    # 如果有任何修复或新建动作，保存文件
    if config_needs_save:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            config.write(f)
            logging.info("配置变更已保存至 config.ini。")
            
    # 初次运行弹窗中断
    if is_first_run:
        msg = f"已在当前目录下生成配置文件：\n{CONFIG_FILE}\n\n请打开该文件，填写您的校园网账号和密码后，再次运行本程序。"
        logging.info("初次运行，等待用户填写账号信息，程序挂起退出。")
        ctypes.windll.user32.MessageBoxW(0, msg, "初次运行配置", 0x40)
        sys.exit(0)
        
    # 严格校验必须的账号密码
    username = config.get('Account', 'username', fallback='').strip()
    password = config.get('Account', 'password', fallback='').strip()
    if not username or not password:
        logging.error("账号或密码未填写！")
        ctypes.windll.user32.MessageBoxW(0, "配置文件中的账号或密码为空！\n请打开 config.ini 填写完整后再运行。", "配置错误", 0x10)
        sys.exit(0)
        
    logging.info("配置文件加载与校验完成，数据完整。")
    return config

# 模块加载时，先读取或生成当前目录的配置文件
app_config = load_or_create_config()

# ================= 特工潜入模块 (带权限容灾保护) =================

def self_install_and_run():
    """判断环境，执行潜伏拷贝，遇到权限不足自动降级到 Desktop/data 目录"""
    if not getattr(sys, 'frozen', False):
        logging.info("当前为 Python 源码运行环境，跳过 U 盘潜伏逻辑。")
        return
        
    current_exe = sys.executable
    exe_name = os.path.basename(current_exe)
    
    # 设定主基地：AppData\Local
    local_appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~\\AppData\\Local'))
    primary_target_dir = os.path.join(local_appdata, 'CampusNetDaemon')
    
    # 设定备用基地：桌面\data
    fallback_target_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'data', 'CampusNetDaemon')
    
    current_dir = os.path.dirname(current_exe).lower()
    
    # 检查是否已经在基地里了
    if current_dir == primary_target_dir.lower() or current_dir == fallback_target_dir.lower():
        logging.info(f"身份确认：程序已在潜伏基地运行 ({os.path.dirname(current_exe)})")
        return
        
    logging.info("检测到在外部环境 (如 U 盘) 运行，启动潜伏转移机制...")
    
    # 1. 测试主基地权限
    target_dir = primary_target_dir
    try:
        os.makedirs(target_dir, exist_ok=True)
        test_file = os.path.join(target_dir, '.test_write')
        with open(test_file, 'w') as f: f.write('1')
        os.remove(test_file)
        logging.info("主基地 AppData 目录权限测试通过。")
    except Exception as e:
        logging.warning(f"主基地没有写入权限或创建失败: {e}")
        logging.info(f"触发权限降级：尝试切换至备用基地 {fallback_target_dir}...")
        target_dir = fallback_target_dir
        try:
            os.makedirs(target_dir, exist_ok=True)
            logging.info("备用基地创建成功！")
        except Exception as e2:
            logging.error(f"备用基地也无法创建 ({e2})。放弃转移，将在当前位置原地运行。")
            return

    # 2. 开始转移
    target_exe = os.path.join(target_dir, exe_name)
    target_config = os.path.join(target_dir, 'config.ini')
    
    try:
        logging.info(f"正在清理系统内可能残留的旧版本进程: {exe_name} ...")
        os.system(f'taskkill /f /im "{exe_name}" >nul 2>&1')
        time.sleep(0.5) 
        
        logging.info(f"正在克隆核心文件至目标阵地...")
        shutil.copy2(current_exe, target_exe)
        shutil.copy2(CONFIG_FILE, target_config)
        
        logging.info("克隆完成，正在唤醒分身进程...")
        subprocess.Popen([target_exe], cwd=target_dir)
        
        logging.info("分身已接管，本体执行自我销毁退出。您可以安全拔出 U 盘。")
        sys.exit(0)
        
    except Exception as e:
        logging.error(f"潜伏转移过程中发生异常: {e}。程序将留在当前目录原地硬抗。")

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
        self.retry_delay = 300

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
        logging.info("正在唤醒隐身浏览器准备注入...")
        self.page_loaded_event.clear()  
        window.load_url(self.login_url) 

        if not self.page_loaded_event.wait(timeout=15):
            logging.error("登录页面加载超时。")
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
        
        logging.info("等待前端框架渲染并执行注入指令...")
        attempt_result = 'not_found'
        for _ in range(10):
            attempt_result = window.evaluate_js(js_injector)
            if attempt_result == 'attempted':
                logging.info("成功捕获登录组件，完成自动点击！")
                break
            time.sleep(1)
            
        if attempt_result != 'attempted':
            logging.error("等待 10 秒后仍未发现登录输入框，当前网页结构可能不匹配。")
            window.load_url('about:blank')
            return False

        logging.info("静默等待登录结果验证...")
        check_success_js = f"document.getElementById('{self.logout_btn_id}') !== null;"
        
        for _ in range(12): 
            time.sleep(1)
            is_logged_in = window.evaluate_js(check_success_js)
            if is_logged_in:
                logging.info("✅ 成功发现注销按钮，登录状态确认无误！")
                window.load_url('about:blank')
                return True

        logging.warning("超时 12 秒仍未检测到注销按钮，单次登录流程失败。")
        window.load_url('about:blank')
        return False

    def daemon_worker(self, window):
        logging.info(">>> 校园网原生幽灵守护线程正式启动 <<<")
        window.events.loaded += self._on_page_loaded_callback

        while True:
            try:
                if self.is_network_available(self.external_test_url):
                    self.error_count = 0
                    if not self.is_connected:
                        self.is_connected = True
                        logging.info("状态监测：外网连通，保持静默蛰伏。")
                else:
                    current_time = time.time()
                    if self.error_count >= self.max_retries:
                        if current_time - self.last_error_time < self.retry_delay:
                            time.sleep(self.check_interval)
                            continue
                        else:
                            self.error_count = 0

                    if self.is_network_available(self.internal_test_url):
                        logging.info("状态监测：发现内网，外网断开。准备接管网络...")
                        
                        login_success = False
                        total_attempts = self.quick_retry_times + 1 
                        
                        for attempt in range(1, total_attempts + 1):
                            if attempt > 1:
                                logging.info(f"🔁 启动快速重试协议 (第 {attempt-1} 次)...")
                                
                            if self.execute_login(window):
                                login_success = True
                                break 
                            else:
                                if attempt < total_attempts:
                                    logging.warning(f"本次注入未生效，休眠3秒后启动下一次尝试...")
                                    time.sleep(3) 
                        
                        if login_success:
                            self.success_count += 1
                            self.error_count = 0
                            self.is_connected = True
                            logging.info(f"🎉 网络接管成功！(本周期累计成功: {self.success_count} 次)")
                        else:
                            self.error_count += 1
                            self.last_error_time = current_time
                            self.is_connected = False
                            logging.error(f"❌ 连续 {total_attempts} 次突防均告失败，进入冷却周期。")

                    else:
                        self.is_connected = False
                
                time.sleep(self.check_interval)

            except Exception as e:
                logging.error(f"守护主线程发生未捕获异常: {str(e)}")
                time.sleep(self.check_interval)

# ================= 托盘图标控制区 =================

def create_tray_icon_image():
    image = Image.new('RGB', (64, 64), color=(255, 255, 255))
    draw = ImageDraw.Draw(image)
    draw.ellipse((16, 16, 48, 48), fill=(46, 204, 113))
    return image

def exit_action(icon, item):
    logging.info("接收到用户退出指令，结束进程。")
    icon.stop()
    os._exit(0)

def setup_tray():
    menu = pystray.Menu(item('退出守护程序', exit_action))
    icon = pystray.Icon("CampusNetAutoLogin", create_tray_icon_image(), "校园网幽灵守护中...", menu)
    icon.run()

# ================= 主程序入口 =================

if __name__ == "__main__":
    # 核心：环境甄别与潜入逻辑，必须置于首位
    self_install_and_run()
    
    # 启动托盘图标
    tray_thread = threading.Thread(target=setup_tray, daemon=True)
    tray_thread.start()

    # 启动网络守护主体
    daemon = CampusNetWebviewDaemon()
    
    hidden_window = webview.create_window(
        title='CampusNet_Daemon', 
        url='about:blank', 
        hidden=True, 
        min_size=(1, 1)
    )
    webview.start(daemon.daemon_worker, hidden_window)