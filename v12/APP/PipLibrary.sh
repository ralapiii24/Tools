#!/usr/bin/env bash
# Linux 环境下的依赖安装脚本（手动/备用）
set -euo pipefail

echo "=== 安装巡检依赖库 ==="
python3 -m pip install --upgrade pip
python3 -m pip install pyyaml tqdm requests lxml paramiko playwright openpyxl xlsxwriter

echo "=== 安装 Playwright 浏览器依赖 ==="
python3 -m playwright install-deps
python3 -m playwright install chromium

echo "=== 安装完成 ==="

