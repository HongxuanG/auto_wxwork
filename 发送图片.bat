@echo off
chcp 65001 > nul
set /p "group_name=请输入群聊名称："
set /p "image_path=请输入图片路径（多个图片用空格分隔）："
企业微信自动化工具.exe -g "%group_name%" --image %image_path%
pause
