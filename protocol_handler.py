import ctypes
import winreg
import os
import sys
import urllib.parse

def is_admin():
    """检查是否具有管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def register_protocol(protocol_name="AutoWxwork"):
    if not is_admin():
        print("需要管理员权限来注册协议")
        # 如果不是管理员，尝试以管理员身份重新运行
        try:
            if sys.argv[-1] != "asadmin":
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, " ".join(sys.argv) + " asadmin", None, 1
                )
                return True
        except Exception as e:
            print(f"请以管理员身份运行程序: {str(e)}")
            return False
    """注册自定义协议到 Windows 注册表"""
    try:
        # 获取当前执行文件路径
        if getattr(sys, 'frozen', False):
            application_path = sys.executable
        else:
            application_path = os.path.abspath(__file__)

        # 创建协议键
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, protocol_name) as key:
            # 设置协议描述
            winreg.SetValue(key, "", winreg.REG_SZ, f"URL:{protocol_name} Protocol")
            winreg.SetValueEx(key, "URL Protocol", 0, winreg.REG_SZ, "")
            
            # 设置图标
            with winreg.CreateKey(key, "DefaultIcon") as icon_key:
                winreg.SetValue(icon_key, "", winreg.REG_SZ, application_path)
            
            # 设置命令
            with winreg.CreateKey(key, r"shell\open\command") as cmd_key:
                winreg.SetValue(cmd_key, "", winreg.REG_SZ, f'"{application_path}" "%1"')
        return True
    except Exception as e:
        print(f"注册协议失败: {str(e)}")
        return False

def unregister_protocol(protocol_name="AutoWxwork"):
    """注销自定义协议"""
    try:
        # 尝试删除注册表项
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f"{protocol_name}\\shell\\open\\command")
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f"{protocol_name}\\shell\\open")
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f"{protocol_name}\\shell")
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f"{protocol_name}\\DefaultIcon")
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, protocol_name)
        print(f"成功注销协议 {protocol_name}")
        return True
    except WindowsError as e:
        print(f"注销协议失败: {str(e)}")
        return False

def is_protocol_registered(protocol_name="AutoWxwork"):
    """检查协议是否已注册"""
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, protocol_name) as key:
            return True
    except WindowsError:
        return False

def handle_protocol_args():
    """处理协议参数"""
    if len(sys.argv) > 1:
        protocol_url = sys.argv[1]
        # 解析协议URL，格式如: autowxwork://action?param1=value1&param2=value2
        if protocol_url.startswith("autowxwork://"):
            try:
                # 移除协议前缀
                params_str = protocol_url.split("://", 1)[1]
                
                # URL解码参数
                decoded_params = urllib.parse.unquote(params_str)
                
                # 将参数转换为命令行格式
                cmd_params = []
                for param in decoded_params.split('&'):
                    if '=' in param:
                        key, value = param.split('=', 1)
                        cmd_params.extend([f'--{key}', value])
                
                return cmd_params
            except Exception as e:
                print(f"处理协议参数时出错: {str(e)}")
                return None
    return None
