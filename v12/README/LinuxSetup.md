# Linux 运行与依赖安装指南

## 前提
- Python 3.9+（推荐使用 `python3`）；
- 系统需要可联网访问 PyPI 和 playwright 镜像；
- 对 Playwright 浏览器需安装依赖库（通常通过 `apt install` 或 `python -m playwright install-deps`）。

## 快速安装依赖（推荐）
1. 进入仓库根目录：`cd /path/to/Inspection/v12`
2. 手动执行以下命令（按顺序）：
```bash
python3 -m ensurepip --upgrade
python3 -m pip install --upgrade pip
python3 -m pip install pyyaml tqdm requests lxml paramiko playwright openpyxl xlsxwriter
python3 -m playwright install-deps
python3 -m playwright install chromium
```

## 启动巡检
1. 依赖安装完成后，运行 `python3 Main.py` 启动程序；
2. `Main.py` 直接封装核心巡检调度逻辑，部署时不再依赖 APP 目录，本地依赖需在安装阶段准备；
3. 如需排查 Playwright 依赖，可以单独运行 `python3 -m playwright install-deps` 确保系统库存在。

## 其他提示
- 若系统缺少 `pip`，可以 `python3 -m ensurepip --upgrade` 后再执行上述步骤；
- Playwright 下载 Chromium 后会缓存到 `~/.cache/ms-playwright`，可在该目录下手动清理；
- 推荐在项目目录下使用 `python3 -m venv venv` 创建虚拟环境，再激活后运行脚本，避免与系统 Python 冲突。

