import pyautogui
import pyperclip
import time
import subprocess
import psutil

# 定义检查进程是否在运行的函数
def is_process_running(process_name):
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == process_name:
            return True
    return False

# 检查企业微信是否已运行
if not is_process_running('WXWork.exe'):
    print("企业微信未运行，正在启动企业微信...")
    subprocess.Popen(r'D:\software\WXWork\WXWork.exe')
    time.sleep(5)
else:
    print("企业微信已在运行。")

# 尝试多次获取企业微信窗口
max_retries = 10         # 最大重试次数
retry_delay = 1          # 每次重试间隔时间（秒）

for attempt in range(max_retries):
    print(f"正在获取企业微信窗口...（尝试 {attempt + 1}/{max_retries}）")
    windows = [w for w in pyautogui.getAllWindows() if '企业微信' in w.title]
    if windows:
        print("企业微信窗口已找到。")
        break
    else:
        print("未找到企业微信窗口，等待一会儿再重试...")
        time.sleep(retry_delay)
else:
    print('超过最大重试次数，未找到企业微信窗口。')
    exit()

wxwork_window = windows[0]
wxwork_window.activate()
print("企业微信窗口已激活。")
time.sleep(1)

# 确保企业微信窗口获得焦点
print("确保企业微信窗口获得焦点...")
wxwork_window.maximize()
time.sleep(1)
wxwork_window.restore()
time.sleep(1)
wxwork_window.activate()
time.sleep(1)

# 使用快捷键打开搜索框，并确保焦点在搜索框内
print("正在打开搜索框...")
pyautogui.hotkey('ctrl', 'f')
time.sleep(1)

# 清空搜索框内容，确保焦点在搜索框内
print("清空搜索框...")
pyautogui.hotkey('ctrl', 'a')
time.sleep(0.5)
pyautogui.press('backspace')
time.sleep(0.5)

# 使用剪贴板粘贴群聊名称
group_name = '企业微信跳转测试'
print(f"正在输入群聊名称：{group_name}")
pyperclip.copy(group_name)
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)

# 进入群聊
print("进入群聊...")
pyautogui.press('enter')
time.sleep(1)

# 输入消息内容
message = '大家好，这是自动发送的消息。'
print(f"正在输入消息内容：{message}")
pyperclip.copy(message)
pyautogui.hotkey('ctrl', 'v')
time.sleep(0.5)

# 发送消息
print("正在发送消息...")
pyautogui.press('enter')
print("消息已发送。")
