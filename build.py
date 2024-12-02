import shutil
import PyInstaller.__main__
import os

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 确保图标文件存在
icon_path = os.path.join(current_dir, 'auto_wxwork.ico')
if not os.path.exists(icon_path):
    print(f"警告：图标文件 {icon_path} 不存在，将不使用自定义图标")
    icon_option = []
else:
    icon_option = ['--icon=' + icon_path]

# 清理之前的构建文件
build_dir = os.path.join(current_dir, 'build')
dist_dir = os.path.join(current_dir, 'dist')
if os.path.exists(build_dir):
    shutil.rmtree(build_dir)
if os.path.exists(dist_dir):
    shutil.rmtree(dist_dir)

# 打包命令
PyInstaller.__main__.run([
    'main.py',                    # 主程序文件
    '--name=企业微信自动化工具',  # 生成的exe名称
    '--onefile',                  # 打包成单个文件
    '--noconsole',               # 运行时不显示命令行窗口
    '--add-data=README.md;.',    # 添加额外文件
    '--clean',                   # 清理临时文件
    '--workpath=build',          # 指定工作目录
    '--distpath=dist',           # 指定输出目录
    '--uac-admin',              # 请求管理员权限
    '--version-file=version.txt', # 版本信息文件
] + icon_option)
