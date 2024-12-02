# 企业微信自动化工具

这是一个用于自动化操作企业微信的工具，可以实现自动搜索群聊并发送消息和图片。

## 功能特点

- 自动启动企业微信（如未运行）
- 自动搜索并进入指定群聊
- 支持发送文本消息
- 支持发送本地图片或网络图片
- 支持发送多张图片
- 自动查找企业微信安装路径

## 开发者指南

### 安装依赖

```bash
pip install pyautogui
pip install pyperclip
pip install psutil
pip install pywin32
pip install pygetwindow
pip install Pillow
pip install requests
```

### Python 使用方法

#### 基本命令格式

```bash
python main.py [选项]
```

#### 命令行参数

- `-p, --path`: 企业微信可执行文件的路径（可选）
- `-g, --group`: 目标群聊名称（必需）
- `-t, --text`: 要发送的文本消息
- `--image`: 要发送的图片文件路径或URL，支持多个图片（用空格分隔）

#### 使用示例

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
python main.py -g "测试群" --image "https://example.com/image1.jpg" "https://example.com/image2.jpg"
```

4. 指定企业微信路径：

```bash
python main.py -p "C:\Program Files\WXWork\WXWork.exe" -g "群聊名称" -t "消息内容"
```

## 用户使用指南

### 命令行方式使用

1. **发送文本消息**：

```bash
企业微信自动化工具.exe -g "群聊名称" -t "要发送的文本消息"
```

2. **发送本地图片**：

```bash
企业微信自动化工具.exe -g "群聊名称" --image "C:\图片1.jpg" "D:\图片2.png"
```

3. **发送网络图片**：

```bash
企业微信自动化工具.exe -g "群聊名称" --image "http://example.com/image1.jpg" "http://example.com/image2.jpg"
```

### 批处理文件使用

1. **使用发送文本消息.bat**：
   - 双击运行 `发送文本消息.bat`
   - 按提示输入群聊名称
   - 按提示输入要发送的消息内容

2. **使用发送图片.bat**：
   - 双击运行 `发送图片.bat`
   - 按提示输入群聊名称
   - 按提示输入图片路径（多个图片用空格分隔）

## 注意事项

1. 运行程序前请确保：
   - 企业微信已安装并已登录
   - 群聊名称输入正确
   - 有足够的发送权限

2. 图片发送说明：
   - 支持本地图片路径和网络图片URL
   - 网络图片会临时下载到用户的下载文件夹
   - 发送完成后会自动清理临时文件

3. 使用注意：
   - 程序会自动调整企业微信窗口大小
   - 运行过程中请勿移动鼠标或操作键盘
   - 建议以管理员身份运行程序

## 常见问题

1. **找不到企业微信**
   - 检查企业微信是否已安装
   - 使用 `-p` 参数手动指定企业微信路径

2. **找不到群聊**
   - 确保群聊名称输入正确（需要完全匹配）
   - 确保企业微信已登录

3. **图片发送失败**
   - 检查图片路径是否正确
   - 检查网络连接是否正常
   - 确保图片格式受支持（支持 jpg、png 等常见格式）

## 系统要求

### 开发环境要求

- 操作系统：Windows
- Python 3.6+
- 企业微信客户端

### 运行环境要求

- 操作系统：Windows 7/8/10/11
- 企业微信客户端（已登录）
- 管理员权限（推荐）

## License

MIT License
