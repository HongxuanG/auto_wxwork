import argparse
import pyautogui
import pyperclip
import time
import subprocess
import psutil
import win32gui
import win32con
import pygetwindow
import sys

def parse_arguments():
    parser = argparse.ArgumentParser(description='企业微信自动化工具')
    parser.add_argument('-p', '--path', 
                       help='企业微信可执行文件的路径',
                       default=None)
    parser.add_argument('-g', '--group', 
                       help='企业微信自动发送的目标群聊',
                       default=None)
    parser.add_argument('-t', '--text', 
                       help='企业微信自动发送的文本',
                       default=None)
    parser.add_argument('--image',
                        help='要发送的图片文件路径，多个图片用空格分隔',
                        nargs='+',  # 允许接收多个参数
                        default=None)
    return parser.parse_args()

def find_wxwork_path():
    # 常见的企业微信安装路径
    possible_paths = [
        r"C:\Program Files (x86)\WXWork",
        r"C:\Program Files\WXWork",
        r"D:\software\WXWork",
        os.path.join(os.environ['LOCALAPPDATA'], "WXWork"),
        os.path.join(os.environ['PROGRAMFILES'], "WXWork"),
        os.path.join(os.environ['PROGRAMFILES(X86)'], "WXWork")
    ]
    
    # 遍历所有可能的路径
    for base_path in possible_paths:
        wxwork_exe = os.path.join(base_path, "WXWork.exe")
        if os.path.exists(wxwork_exe):
            print(f"找到企业微信安装路径：{wxwork_exe}")
            return wxwork_exe
    
    print("未找到企业微信安装路径")
    return None

def take_debug_screenshot(x, y, width=100, height=100):
        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        screenshot.save('search_box_debug.png')
        print(f"已保存搜索框区域截图，请检查 search_box_debug.png")
# 定义检查进程是否在运行的函数
def is_process_running(process_name):
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == process_name:
            return True
    return False

# 定义根据窗口标题查找窗口句柄的函数
def find_window_handle(window_title):
    def callback(hwnd, extra):
        title = win32gui.GetWindowText(hwnd)
        if window_title in title:
            extra.append(hwnd)
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    if hwnds:
        return hwnds[0]
    else:
        return None

def clear_search_box():
    # 清空搜索框内容，确保焦点在搜索框内
    print("清空搜索框...")
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.3)
    pyautogui.press('backspace')
    time.sleep(0.5)


def send_message(message):
    pyperclip.copy(message)
    pyautogui.hotkey('ctrl', 'v')

def url_image_to_clipboard(image_urls):
    import requests
    import os
    from urllib.parse import urlparse

    # 使用session下载图片，增加超时时间
    session = requests.Session()
    session.trust_env = False  # 禁用环境变量中的代理设置
    # 获取用户的下载文件夹路径
    download_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
    for img_url in image_urls:
        print(f"正在处理图片：{img_url}")
        try:
            # 检查是否是URL
            if img_url.startswith(('http://', 'https://')):

                response = session.get(img_url, timeout=30)
                if response.status_code == 200:
                    # 从URL中提取文件名
                    filename = os.path.basename(urlparse(img_url).path)
                    temp_path = os.path.join(download_folder, 'temp_' + filename)
                    
                    # 保存临时文件
                    with open(temp_path, 'wb') as f:
                        f.write(response.content)
                    
                    # 使用临时文件路径
                    image_path = temp_path
                else:
                    err_msg = f"下载图片失败: HTTP {response.status_code}"
                    send_message(err_msg)
                    print(err_msg)
                    continue
            else:
                # 如果不是URL，直接使用本地路径
                image_path = img_url

            # 打开并处理图片
            from PIL import Image
            image = Image.open(image_path)
            
            # 将图片复制到剪贴板
            from io import BytesIO
            output = BytesIO()
            image.convert('RGB').save(output, 'BMP')
            data = output.getvalue()[14:]
            output.close()
            
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            
            print(f"已将图片复制到剪贴板: {temp_path}")
            # 粘贴图片
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            
            # 删除临时文件
            if img_url.startswith(('http://', 'https://')):
                try:
                    os.remove(temp_path)
                except:
                    pass
                    
            # 每张图片发送后等待一下
            if img_url != args.image[-1]:
                time.sleep(0.5)
                
        except Exception as e:
            err_msg = f"处理图片时出错: {str(e)}"
            send_message(err_msg)
            print(err_msg)
            continue


# 获取命令行参数
args = parse_arguments()

# 检查企业微信是否已运行
if not is_process_running('WXWork.exe'):
    print("企业微信未运行，正在启动企业微信...")
    
    # 优先使用命令行指定的路径
    wxwork_path = args.path if args.path and os.path.exists(args.path) else find_wxwork_path()
    if wxwork_path:
        print("正在启动企业微信...")
        subprocess.Popen(wxwork_path)
        time.sleep(5)
    else:
        print("无法找到企业微信程序，请确认是否已安装或使用 -p 参数指定正确的路径")
        sys.exit(1)
else:
    print("企业微信已在运行。")

# 获取窗口句柄
def get_window_active(window_name):
    # 尝试多次获取企业微信窗口
    max_retries = 3         # 最大重试次数
    retry_delay = 1          # 每次重试间隔时间（秒）

    for attempt in range(max_retries):
        print(f"正在获取{window_name}窗口...（尝试 {attempt + 1}/{max_retries}）")
        # 使用 win32gui 查找窗口句柄
        hwnd = find_window_handle(window_name)
        if hwnd:
            print(f"{window_name}窗口已找到（通过 win32gui），HWND: {hwnd}")
            # 恢复最小化或隐藏的窗口
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(1)
            # 将窗口句柄转换为窗口对象
            wxwork_window = pygetwindow.Window(hwnd)
            break
        else:
            print("未找到{window_name}窗口，等待一会儿再重试...")
            time.sleep(retry_delay)
    else:
        print('超过最大重试次数，未找到{window_name}窗口。')
        return None, None

    # 激活企业微信窗口
    wxwork_window.activate()
    print(f"{window_name}窗口已激活。")
    return wxwork_window, hwnd

def check_search_result(window_left, window_top, window_width, window_height): 
    # 计算搜索结果区域的位置（搜索框下方区域）
    print(f"搜索结果区域位置：left={window_left}, top={window_top}, width={window_width}, height={window_height}")
    search_area_x = window_left + int(window_width * 0.06)  # 左边距10%
    search_area_y = window_top + int(window_height * 0.058)  # 顶部距离10%
    search_area_width = int(window_width * 0.3)  # 搜索结果区域宽度
    search_area_height = int(window_height * 0.2)  # 搜索结果区域高度
    
    print("等待搜索结果加载...")
    time.sleep(1)  # 给搜索一些响应时间
    
    # 获取搜索结果区域的截图
    try:
        screenshot = pyautogui.screenshot(region=(
            search_area_x, 
            search_area_y, 
            search_area_width, 
            search_area_height
        ))
        
        # 保存截图用于调试（可选）
        screenshot.save('search_result.png')
        
        # 检查截图中的像素颜色变化来判断是否有搜索结果
        # 将图片转换为灰度值列表
        pixels = list(screenshot.convert('L').getdata())
        
        # 计算非白色像素的数量（假设背景是白色）
        non_white_pixels = sum(1 for p in pixels if p < 250)  # 250是阈值，可以调整
        
        # 如果非白色像素超过一定比例，认为有搜索结果
        threshold = len(pixels) * 0.1  # 10%的像素变化阈值，可以调整
        print(f"非白色像素数量: {non_white_pixels}, 阈值: {threshold}")
        if non_white_pixels > threshold:
            print("找到搜索结果")
            return True
        else:
            print("未找到搜索结果")
            return False
            
    except Exception as e:
        print(f"检查搜索结果时出错: {e}")
        return False

wxwork_window, hwnd = get_window_active('全局搜索')
if wxwork_window is None:
    print("未找到全局搜索窗口，程序即将通过激活企业微信窗口搜索")
    
    # 激活企业微信窗口
    wxwork_window, hwnd = get_window_active('企业微信')
    if wxwork_window is None:
        print("未找到企业微信窗口，程序即将退出。")
        sys.exit()
    time.sleep(1)
    # 调整窗口大小为 1200x900
    print("调整窗口大小...")
    window_x = wxwork_window.left  # 保持原来的位置
    window_y = wxwork_window.top   # 保持原来的位置
    win32gui.MoveWindow(hwnd, window_x, window_y, 1200, 900, True)
    time.sleep(1)

    # 获取窗口的位置和大小
    window_left = wxwork_window.left
    window_top = wxwork_window.top
    window_width = wxwork_window.width
    window_height = wxwork_window.height
    print(f"窗口位置和大小：left={window_left}, top={window_top}, width={window_width}, height={window_height}")

    # 确保企业微信窗口获得焦点
    print("确保企业微信窗口获得焦点...")
    wxwork_window.activate()
    time.sleep(0.5)
    pyautogui.press('alt')
    time.sleep(0.5)

    # 使用快捷键打开搜索框，并确保焦点在搜索框内
    print("正在打开搜索框...")
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(1)

    # 计算搜索框的位置（根据窗口大小的比例计算）
    search_box_x = window_left + int(window_width * 0.06)  # 左边距约6%
    search_box_y = window_top + int(window_height * 0.058)  # 顶部距离约5.8%
    print(f"搜索框位置：x={search_box_x}, y={search_box_y}")
    full_search_box_width = 230
    full_search_box_height = 50
    # 方法1：直接点击搜索框位置
    print("点击搜索框...")

    # 计算搜索框中心点位置
    center_x = search_box_x + (full_search_box_width // 2)
    center_y = search_box_y + (full_search_box_height // 2)

    pyautogui.moveTo(center_x, center_y)   # moves mouse to X of 100, Y of 200.
    pyautogui.click(center_x, center_y)
    time.sleep(1)


# 激活全局搜索窗口
full_search_window, hwnd = get_window_active('全局搜索')
time.sleep(1)
# 调整窗口大小为 1000x700
print("调整窗口大小...")
window_x = full_search_window.left  # 保持原来的位置
window_y = full_search_window.top   # 保持原来的位置
win32gui.MoveWindow(hwnd, window_x, window_y, 800, 800, True)
time.sleep(1)

# 获取窗口的位置和大小
window_left = full_search_window.left
window_top = full_search_window.top
window_width = full_search_window.width
window_height = full_search_window.height
print(f"窗口位置和大小：left={window_left}, top={window_top}, width={window_width}, height={window_height}")
print("确保企业微信窗口获得焦点...")
full_search_window.activate()
group_search_option_x = window_left + int(window_width * 0.13)  # 左边距约6%
group_search_option_y = window_top + int(window_height * 0.1)  # 顶部距离约5.8%
group_search_option_width = 40
group_search_option_height = 30

# take_debug_screenshot(group_search_option_x, group_search_option_y, group_search_option_width, group_search_option_height)
# 计算搜索框中心点位置
center_x = group_search_option_x + (group_search_option_width // 2)
center_y = group_search_option_y + (group_search_option_height // 2)

pyautogui.moveTo(center_x, center_y)   # moves mouse to X of 100, Y of 200.
pyautogui.click(center_x, center_y)
time.sleep(0.5)

pyautogui.moveTo(center_x, center_y - 50)   # moves mouse to X of 100, Y of 200.
pyautogui.click(center_x, center_y - 50)
time.sleep(0.5)

clear_search_box()

# 使用剪贴板粘贴群聊名称
group_name = args.group
print(f"正在输入群聊名称：{group_name}")
send_message(group_name)
time.sleep(1)

# 判断是否成功搜索到群聊
print("判断是否成功搜索到群聊...")


if check_search_result(window_left, window_top, window_width, window_height):
    # 进入群聊
    print("进入群聊...")
    pyautogui.press('enter')
    time.sleep(1)  # 等待群聊窗口加载
    print("已成功进入群聊，准备发送消息。")
    clear_search_box()
    
    # 根据参数决定发送文本还是图片
    if args.image:
        print(f"准备发送 {len(args.image)} 张图片...")
        url_image_to_clipboard(args.image)
    elif args.text:
        # 输入消息内容
        message = args.text
        print(f"正在输入消息内容：{message}")
        send_message(message)
        time.sleep(0.5)
    else:
        print("未指定要发送的内容，程序即将退出。")
        sys.exit()

    # 发送消息
    print("正在发送消息...")
    pyautogui.press('enter')
    print("消息已发送。")
else:
    print("未能确认找到指定的群聊，程序即将退出。")
    sys.exit()
