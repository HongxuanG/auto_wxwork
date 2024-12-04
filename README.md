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
pip install pyinstaller
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

### 打包说明

#### 1. 使用 PyInstaller 打包成 EXE

```bash
python build.py
```

打包后会在 `dist` 目录下生成 `企业微信自动化工具.exe`。

#### 2. 使用 Inno Setup 制作安装程序

1. **准备文件**
   - 确保 `dist` 目录下有打包好的 exe 文件
   - 准备以下文件：
     - `setup.iss`（Inno Setup 脚本）
     - `auto_wxwork.ico`（图标文件）
     - `README.md`（说明文档）
     - `发送文本消息.bat`
     - `发送图片.bat`

2. **修改版本号**
   - 更新 `version.txt` 中的版本信息
   - 更新 `setup.iss` 中的版本号

3. **编译安装程序**
   - 打开 Inno Setup Compiler
   - 打开 `setup.iss` 文件
   - 点击 "Build" -> "Compile" 或按 Ctrl+F9
   - 编译完成后会在 `Output` 目录生成安装程序

4. **发布新版本流程**
   1. 更新代码并测试
   2. 修改版本号（`version.txt` 和 `setup.iss`）
   3. 运行 `python build.py` 生成新的 exe
   4. 使用 Inno Setup 编译生成安装程序
   5. 测试安装程序
   6. 发布新版本

## 用户使用指南

### 网页调用方式

1. **引入 JS 文件**

```html
<script src="openAutoWxwork.js"></script>
```

2. **基本用法**

```html
<!-- 发送文本消息 -->
<button onclick="openAutoWxwork('测试群', '测试消息')">发送消息</button>

<!-- 发送图片 -->
<button onclick="openAutoWxwork('测试群', null, ['图片URL1', '图片URL2'])">发送图片</button>
```

3. **完整示例**

```html
<!DOCTYPE html>
<html>
<head>
    <title>企业微信自动化工具示例</title>
    <script src="openAutoWxwork.js"></script>
</head>
<body>
    <button onclick="openAutoWxwork('测试群', '测试消息')">发送文本</button>
    <button onclick="openAutoWxwork('测试群', null, ['http://example.com/image.jpg'])">发送图片</button>
</body>
</html>
```

4. **URL 协议调用**

```javascript
// 发送文本消息
window.open('autowxwork://group=测试群&text=测试消息')

// 发送图片
window.open('autowxwork://group=测试群&image=http://example.com/image.jpg')
```

首次使用时会弹出协议调用确认框：

![协议调用确认](https://github.com/user-attachments/assets/ac61f7fd-d91a-4a32-b19f-6e96139fe582)

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
- 企业微信客户端 4.1.31.6017
- Inno Setup（用于制作安装程序）

### 运行环境要求

- 操作系统：Windows 7/8/10/11
- 企业微信客户端（已登录）
- 管理员权限（推荐）

## License

MIT License
