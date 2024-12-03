import shutil
import sys
import time
import PyInstaller.__main__
import os

import psutil

from protocol_handler import unregister_protocol

def kill_running_exe():
    """结束正在运行的exe进程"""
    exe_name = "企业微信自动化工具.exe"
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] == exe_name:
                print(f"正在结束进程: {exe_name}")
                proc.kill()
                proc.wait()  # 等待进程完全结束
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

def safe_remove_path(path):
    """安全地删除文件或目录"""
    max_retries = 5
    retry_delay = 1
    
    for i in range(max_retries):
        try:
            if os.path.isfile(path):
                os.unlink(path)
            elif os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
            return True
        except Exception as e:
            print(f"删除失败 (尝试 {i+1}/{max_retries}): {str(e)}")
            if i < max_retries - 1:
                time.sleep(retry_delay)
                continue
            return False


# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 确保图标文件存在
icon_path = os.path.join(current_dir, 'auto_wxwork.ico')
if not os.path.exists(icon_path):
    print(f"警告：图标文件 {icon_path} 不存在，将不使用自定义图标")
    icon_option = []
else:
    icon_option = ['--icon=' + icon_path]

# 结束正在运行的exe
kill_running_exe()

# 等待一会儿确保进程完全结束
time.sleep(1)


# 清理之前的构建文件
build_dir = os.path.join(current_dir, 'build')
dist_dir = os.path.join(current_dir, 'dist')

unregister_protocol()

print("正在清理旧文件...")
if not safe_remove_path(build_dir):
    print("警告：无法完全清理 build 目录")
if not safe_remove_path(dist_dir):
    print("警告：无法完全清理 dist 目录")

try:
    print("开始打包...")
    # 打包命令
    PyInstaller.__main__.run([
        'main.py',                    # 主程序文件
        '--name=企业微信自动化工具',  # 生成的exe名称
        '--onefile',                  # 打包成单个文件
        '--console',                 # 修改为 console 模式，而不是 noconsole
        '--add-data=README.md;.',    # 添加额外文件
        '--add-data=protocol_handler.py;.',  # 添加protocol_handler.py文件
        '--clean',                   # 清理临时文件
        '--workpath=build',          # 指定工作目录
        '--distpath=dist',           # 指定输出目录
        '--uac-admin',              # 请求管理员权限
        '--version-file=version.txt', # 版本信息文件
    ] + icon_option)
except Exception as e:
    print(f"打包失败: {str(e)}")
    sys.exit(1)

