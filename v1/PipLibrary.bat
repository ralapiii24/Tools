@echo off
echo 正在安装巡检依赖库...
pip install --upgrade pip
pip install pyyaml tqdm requests lxml paramiko playwright colorama
echo 安装 Playwright 浏览器（Chromium）...
playwright install chromium
echo 完成！
pause