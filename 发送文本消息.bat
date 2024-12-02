@echo off
chcp 65001 > nul
set /p "group_name=请输入群聊名称："
set /p "message=请输入要发送的消息："
企业微信自动化工具.exe -g "%group_name%" -t "%message%"
pause
