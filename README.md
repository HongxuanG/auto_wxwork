# 企业微信自动化工具

这是一个用于自动化操作企业微信的Python工具，可以实现自动搜索群聊并发送消息和图片。

## 功能特点

- 自动启动企业微信（如未运行）
- 自动搜索并进入指定群聊
- 支持发送文本消息
- 支持发送本地图片或网络图片
- 支持发送多张图片
- 自动查找企业微信安装路径

## 安装依赖

```bash
pip install pyautogui
pip install pyperclip
pip install psutil
pip install pywin32
pip install pygetwindow
pip install Pillow
pip install requests
```

## 使用方法

### 基本命令格式

```bash
python main.py [选项]
```

### 命令行参数

- `-p, --path`: 企业微信可执行文件的路径（可选）
- `-g, --group`: 目标群聊名称（必需）
- `-t, --text`: 要发送的文本消息
- `--image`: 要发送的图片文件路径或URL，支持多个图片（用空格分隔）

### 使用示例

1. 发送文本消息：
```bash
python main.py -g "测试群" -t "你好"
```
2. 发送单张图片：
```bash
python main.py -g "测试群" --image "https://example.com/image.jpg"
```

3. 发送多张图片：
```bash
python main.py -g "测试群" --image "https://example.com/image1.jpg https://example.com/image2.jpg"
```
4. 指定企业微信路径：
```bash
python main.py -p "C:\Program Files\WXWork\WXWork.exe" --group "群聊名称" --text "消息内容"
```
## 注意事项

1. 运行程序时请确保：
   - 企业微信已安装
   - 已登录企业微信账号
   - 群聊名称输入正确
   - 有足够的发送权限

2. 图片发送：
   - 支持本地图片路径和网络图片URL
   - 网络图片会临时下载到用户的下载文件夹
   - 发送完成后会自动清理临时文件

3. 窗口操作：
   - 程序会自动调整企业微信窗口大小
   - 请勿在程序运行时手动操作鼠标和键盘

## 系统要求

- 操作系统：Windows
- Python 3.6+
- 企业微信客户端

## 常见问题

1. 如果找不到企业微信路径，可以使用 `-p` 参数手动指定
2. 如果群聊搜索失败，请检查群名称是否正确
3. 如果图片发送失败，请检查网络连接和图片URL是否有效

## License

MIT License
