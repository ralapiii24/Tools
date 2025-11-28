# 主程序入口文件：整合了依赖检查和任务调度功能
#
# 主要职责:
# Main.py:主程序入口文件，整合了所有核心功能
#
# 第一部分：运行时环境检查功能（原 RuntimeEnvCheck.py）
# - 环境预检：核查 pyyaml/tqdm/requests/urllib3/lxml/paramiko/playwright/openpyxl/xlsxwriter 等包，并确认已安装 Playwright Chromium
# - 失败会写入报告到 REPORT/<YYYYMMDD>_DependencyCheck.log，并退出（阻止巡检）
# - 支持多个pip镜像源（官方、阿里云、清华），网络可达时自动安装缺失依赖
# - 支持跨平台（Windows、Linux、macOS），自动识别平台并安装相应的系统依赖
# - 支持 --check 参数：只检查依赖，不执行巡检
# - 支持 --install 参数：安装所有依赖包
#
# 第二部分：任务调度功能（原 Core.py）
# - 多任务巡检（FXOS/FortiGate镜像/Oxidized配置抓取/各类Linux服务器SSH获取CPU内存硬盘指标/Kibana ESN9K日志扫描/Flow服务检查/ACL策略分析/服务检查），写入日报与任务明细
# - V10新结构：创建任务日志目录:LOG/<TaskName>/；日报目录:REPORT/
# - 依次运行任务并记录结果；按任务生成明细日志（LOG/任务类名/YYYYMMDD-任务显示名.log），并在 REPORT/<YYYYMMDD>巡检日报.log 写入汇总与异常摘要
# - 从 YAML/Config.yaml 读取配置；settings.show_progress 控制 tqdm 进度条；所有任务输出统一分级:OK/WARN/CRIT/ERROR
# - 每个任务为 BaseTask 子类（V11: BaseTask位于TaskBase.py）:实现 items()（要巡检的对象列表）与 run_single(item)（单对象巡检逻辑）
# - 结果对象 Result(level, message, meta) 可带 meta 附加信息（如样例、原始行等）
# - 配置验证重构:将具体任务的配置检查从任务调度逻辑迁移到各自任务类中，实现单一职责原则
#
# 执行流程:依赖预检 → 执行巡检任务
#
# REPORT文件格式优化:
# - 添加时间隔离符号：==================== YYYY-MM-DD HH:MM:SS 运行 ====================
# - 时间戳格式：#################### YYYY-MM-DD HH:MM:SS 运行 ####################
# - 任务列表格式化：启用的任务每行一个，禁用的任务每行一个
# - 最新运行追加到顶部：新内容在文件顶部，历史记录在底部
# - 保持历史记录：每次运行保留之前的记录，便于查看历史
#
# 输出:REPORT/日期巡检日报.log（每日汇总报告），LOG/任务类名/YYYYMMDD-任务显示名.log（单任务详细日志，V10新结构）
# 配置说明:支持task_switches任务开关控制，自动创建目录结构

# 导入标准库
import json
import os
import subprocess
import sys
import time
import traceback
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

# 设置标准输出编码为 UTF-8，解决 Linux 系统中文乱码问题
# 在导入其他模块之前设置，确保所有输出都使用 UTF-8
if sys.stdout.encoding != 'utf-8':
    try:
        import io
        # 重新包装 stdout 和 stderr，强制使用 UTF-8 编码
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    except (AttributeError, io.UnsupportedOperation):
        # 如果无法重新包装（例如在某些环境中），设置环境变量
        os.environ['PYTHONIOENCODING'] = 'utf-8'
        # 尝试设置 locale
        try:
            import locale
            locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
        except Exception:
            pass

# 导入第三方库
import yaml
from tqdm import tqdm

# 导入本地应用
from TASK import __all__ as TASK_CLASSES
from TASK.TaskBase import require_keys

# ============================================================================
# 第一部分：运行时环境检查功能（原 RuntimeEnvCheck.py）
# ============================================================================

# 平台检测辅助函数
def _get_platform_name() -> str:
    """获取平台名称"""
    platform = sys.platform
    if platform == "win32":
        return "Windows"
    elif platform == "darwin":
        return "macOS"
    elif platform.startswith("linux"):
        return "Linux"
    else:
        return platform

def _is_windows() -> bool:
    """判断是否为 Windows 平台"""
    return sys.platform == "win32"

def _needs_playwright_system_deps() -> bool:
    """判断是否需要安装 Playwright 系统依赖（Linux 和 macOS 需要，Windows 不需要）"""
    return sys.platform != "win32"

# 获取项目根目录路径
PROJECT_ROOT = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_ROOT / "YAML" / "Config.yaml"

# 读取配置文件
with open(CONFIG_FILE, "r", encoding="utf-8") as CONFIG_FILE_HANDLE:
    CONFIG = yaml.safe_load(CONFIG_FILE_HANDLE)

# 从配置中读取依赖包，转换为所需格式
DEPENDENCIES = CONFIG.get("dependencies", {})
REQUIRED_PY_PKGS = [
    (import_name, package_name, f"{package_name} 包")
    for import_name, package_name in DEPENDENCIES.items()
]

# 定义函数 _try_import
def _try_import(mod: str) -> Tuple[bool, Optional[str]]:
    try:
        __import__(mod)
        return True, None
    except Exception as e:
        return False, f"{e.__class__.__name__}: {e}"

# 检测pip服务器网络连通性：检查pip镜像源是否可达
def _check_pip_network() -> Tuple[bool, Optional[str]]:
    try:
        import urllib.request
        import urllib.error

        TEST_URLS = CONFIG.get("pip_mirrors", [])
        for URL in TEST_URLS:
            try:
                with urllib.request.urlopen(URL, timeout=5) as RESPONSE:
                    if RESPONSE.status == 200:
                        return True, f"网络连通正常，使用镜像源: {URL}"
            except Exception:
                continue

        return False, "所有pip镜像源均无法连接，网络不可达"
    except Exception as e:
        return False, f"网络检测失败: {e}"

def _install_packages(packages: List[str]) -> Tuple[bool, str]:
    if not packages:
        return True, "无需安装"

    COMMAND = [sys.executable, "-m", "pip", "install", "--upgrade"] + packages
    RESULT = subprocess.run(
        COMMAND,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='ignore'
    )

    if RESULT.returncode == 0:
        return True, "依赖包安装成功"
    return False, f"安装失败，返回码: {RESULT.returncode}\n错误信息: {RESULT.stderr}"

def _upgrade_pip() -> Tuple[bool, str]:
    """升级 pip 到最新版本"""
    COMMAND = [sys.executable, "-m", "pip", "install", "--upgrade", "pip"]
    RESULT = subprocess.run(
        COMMAND,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='ignore'
    )

    if RESULT.returncode == 0:
        return True, "pip 升级成功"
    return False, f"pip 升级失败，返回码: {RESULT.returncode}\n错误信息: {RESULT.stderr}"

def _install_playwright_deps() -> Tuple[bool, str]:
    """安装 Playwright 系统依赖（Linux 和 macOS 需要，Windows 不需要）"""
    if _is_windows():
        return True, "Windows 平台无需安装系统依赖"
    
    PLATFORM_NAME = _get_platform_name()
    COMMAND = [sys.executable, "-m", "playwright", "install-deps"]
    RESULT = subprocess.run(
        COMMAND,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='ignore'
    )

    if RESULT.returncode == 0:
        return True, f"Playwright 系统依赖安装成功 ({PLATFORM_NAME})"
    return False, f"Playwright 系统依赖安装失败 ({PLATFORM_NAME})，返回码: {RESULT.returncode}\n错误信息: {RESULT.stderr}"

def _install_playwright_chromium() -> Tuple[bool, str]:
    COMMAND = [sys.executable, "-m", "playwright", "install", "chromium"]
    RESULT = subprocess.run(
        COMMAND,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='ignore'
    )

    if RESULT.returncode == 0:
        return True, "chromium 浏览器安装成功"
    return False, f"chromium 安装失败，返回码: {RESULT.returncode}\n错误信息: {RESULT.stderr}"

def _check_playwright_chromium() -> Tuple[bool, Optional[str]]:
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            browser.close()
        return True, None
    except Exception as e:
        return False, f"{e.__class__.__name__}: {e}"

def _format_missing_details(missing: List[str], details: List[str], prefix: str) -> str:
    NOW = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = []
    lines.append(f"=== [{NOW}] {prefix} ===")
    lines.append("")
    lines.append("缺失项：")
    lines.extend(details)
    lines.append("")
    return "\n".join(lines)

def check_runtime_dependencies() -> Tuple[bool, str]:
    """检查运行时依赖：检查Python依赖包和Playwright浏览器"""
    MISSING: List[str] = []
    DETAILS: List[str] = []

    for mod, pkg, desc in REQUIRED_PY_PKGS:
        OK, ERR = _try_import(mod)
        if not OK:
            MISSING.append(pkg)
            DETAILS.append(f"[缺] {pkg:<12}  ({desc})  import '{mod}' 失败 → {ERR}")

    PLAYWRIGHT_BROWSER_MISSING = False
    if "playwright" not in MISSING:
        OK, ERR = _check_playwright_chromium()
        if not OK:
            PLAYWRIGHT_BROWSER_MISSING = True
            MISSING.append("playwright-browsers(chromium)")
            DETAILS.append(f"[缺] chromium 浏览器 (Playwright)  未安装 → {ERR}")

    if not MISSING:
        return True, "=== Python 依赖包预检通过 ==="

    print("=== 检测到缺失依赖包，正在检查网络连通性... ===")
    NETWORK_OK, NETWORK_MSG = _check_pip_network()
    if not NETWORK_OK:
        lines = _format_missing_details(MISSING, DETAILS, "依赖检查未通过，网络不可达")
        lines += "\n"
        lines += f"网络检测结果：{NETWORK_MSG}\n\n"
        lines += "修复指引：\n"
        lines += "  1) 检查网络连接后重新运行\n"
        lines += "  2) 或手动执行 pip install 命令\n"
        if any(pkg != "playwright-browsers(chromium)" for pkg in MISSING):
            pkg_list = sorted({pkg for pkg in MISSING if pkg != "playwright-browsers(chromium)"})
            lines += f"     pip install -U {' '.join(pkg_list)}\n"
        if "playwright-browsers(chromium)" in MISSING:
            lines += "     python -m playwright install chromium\n"
        return False, lines

    print(f"网络检测结果：{NETWORK_MSG}")
    print("=== 网络连通正常，开始自动安装依赖包... ===")
    install_pkgs = sorted({pkg for pkg in MISSING if pkg != "playwright-browsers(chromium)"})
    INSTALL_OK, INSTALL_MSG = _install_packages(install_pkgs)
    if not INSTALL_OK:
        lines = _format_missing_details(MISSING, DETAILS, "自动安装依赖包失败")
        lines += "\n"
        lines += f"网络检测结果：{NETWORK_MSG}\n"
        lines += f"安装结果：{INSTALL_MSG}\n"
        lines += "\n修复指引：\n"
        lines += "  1) 手动运行 pip install 命令\n"
        if install_pkgs:
            lines += f"     pip install -U {' '.join(install_pkgs)}\n"
        if "playwright-browsers(chromium)" in MISSING:
            lines += "     python -m playwright install chromium\n"
        return False, lines

    if PLAYWRIGHT_BROWSER_MISSING:
        OK, MSG = _install_playwright_chromium()
        if not OK:
            lines = _format_missing_details(MISSING, DETAILS, "Playwright Chromium 安装失败")
            lines += f"\n安装结果：{MSG}\n"
            lines += "\n修复指引：\n"
            lines += "  python -m playwright install chromium\n"
            return False, lines

    print("=== 依赖包安装完成，重新检查... ===")
    return check_runtime_dependencies()

# 写入依赖检查报告：将报告内容写入REPORT目录下的文件
def write_dependency_REPORT(REPORT: str) -> None:
    try:
        DATE = datetime.now().strftime("%Y%m%d")
        OUT_DIR = PROJECT_ROOT / "REPORT"
        os.makedirs(OUT_DIR, exist_ok=True)
        PATH = os.path.join(OUT_DIR, f"{DATE}_DependencyCheck.log")
        with open(PATH, "a", encoding="utf-8") as f:
            f.write(REPORT + "\n")
    except Exception:
        print("[预检] 写入报告文件失败：\n" + traceback.format_exc(), file=sys.stderr)

# 手动安装所有依赖：整合 PipLibrary.bat 和 PipLibrary.sh 的功能
def install_all_dependencies() -> bool:
    """
    手动安装所有巡检依赖库
    功能包括：
    1. 升级 pip
    2. 安装所有 Python 依赖包
    3. 安装 Playwright 系统依赖（Linux 和 macOS，Windows 不需要）
    4. 安装 Playwright Chromium 浏览器
    
    支持平台：Windows、Linux、macOS
    """
    PLATFORM_NAME = _get_platform_name()
    print(f"=== 正在安装巡检依赖库 ({PLATFORM_NAME}) ===")
    
    # 1. 升级 pip
    print("正在升级 pip...")
    OK, MSG = _upgrade_pip()
    if not OK:
        print(f"警告：pip 升级失败: {MSG}")
    else:
        print(MSG)
    
    # 2. 获取所有依赖包名称列表
    ALL_PACKAGES = sorted([package_name for _, package_name, _ in REQUIRED_PY_PKGS])
    print(f"\n正在安装 Python 依赖包: {', '.join(ALL_PACKAGES)}")
    OK, MSG = _install_packages(ALL_PACKAGES)
    if not OK:
        print(f"错误：依赖包安装失败: {MSG}")
        return False
    print(MSG)
    
    # 3. 安装 Playwright 系统依赖（Linux 和 macOS 需要，Windows 不需要）
    if _needs_playwright_system_deps():
        print(f"\n=== 安装 Playwright 浏览器依赖 ({PLATFORM_NAME}) ===")
        OK, MSG = _install_playwright_deps()
        if not OK:
            print(f"警告：Playwright 系统依赖安装失败: {MSG}")
            print("提示：某些功能可能无法正常工作，建议手动安装系统依赖")
        else:
            print(MSG)
    
    # 4. 安装 Playwright Chromium 浏览器
    print("\n正在安装 Playwright 浏览器（Chromium）...")
    OK, MSG = _install_playwright_chromium()
    if not OK:
        print(f"错误：Chromium 安装失败: {MSG}")
        return False
    print(MSG)
    
    print(f"\n=== 安装完成 ({PLATFORM_NAME}) ===")
    return True

# ============================================================================
# 第二部分：任务调度功能（原 Core.py）
# ============================================================================

# 动态导入所有任务类
TASK_MODULES = {}
for TASK_NAME in TASK_CLASSES:
    try:
        TASK_MODULE = __import__(f'TASK.{TASK_NAME}', fromlist=[TASK_NAME])
        TASK_MODULES[TASK_NAME] = getattr(TASK_MODULE, TASK_NAME)
    except Exception as e:
        print(f"警告：无法导入任务类 {TASK_NAME}: {e}")

# 只检查Core需要的通用配置
require_keys(CONFIG, ["settings"], "root")

SHOW_PROGRESS = bool(CONFIG["settings"].get("show_progress", True))

# 主调度器：执行所有启用的巡检任务
def run_inspection_tasks():
    """执行所有启用的巡检任务并生成日报"""
    TODAY = datetime.now().strftime("%Y%m%d")
    SETTINGS = (CONFIG.get("settings") or {})
    BASE_LOG_DIR = SETTINGS.get("log_dir", "LOG")
    # V10新结构：不再使用日期目录，改为任务目录（在任务执行时创建）
    REPORT_DIR = SETTINGS.get("report_dir", "REPORT")
    # 创建报告目录（如果目录已存在则不报错）
    os.makedirs(REPORT_DIR, exist_ok=True)
    DAILY_REPORT = os.path.join(REPORT_DIR, f"{TODAY}巡检日报.log")

    # 获取任务开关配置
    TASK_SWITCHES = CONFIG.get("task_switches", {})

    # 动态创建任务实例（根据开关配置）
    TASKS = []
    ENABLED_TASKS = []
    DISABLED_TASKS = []

    for TASK_NAME in TASK_CLASSES:
        # 检查任务开关，默认为 True（启用）
        IS_ENABLED = TASK_SWITCHES.get(TASK_NAME, True)

        if not IS_ENABLED:
            DISABLED_TASKS.append(TASK_NAME)
            continue

        try:
            TASK_CLASS = TASK_MODULES[TASK_NAME]
            # 所有任务类统一调用，特殊参数处理在各自类内部完成
            TASK_INSTANCE = TASK_CLASS()

            TASKS.append(TASK_INSTANCE)
            ENABLED_TASKS.append(TASK_NAME)
        except Exception as e:
            print(f"警告：无法创建任务实例 {TASK_NAME}: {e}")

    # 任务名称映射
    TASK_NAMES = {
        "FXOSWebTask": "FXOSWebTask-FXOS设备Web巡检",
        "MirrorFortiGateTask": "MirrorFortiGateTask-FortiGate设备镜像巡检", 
        "OxidizedTask": "OxidizedTask-Oxidized配置备份巡检",
        "ESLogstashTask": "ESLogstashTask-Logstash服务器巡检",
        "ESBaseTask": "ESBaseTask-Elasticsearch基础巡检",
        "ESN9KLOGInspectTask": "ESN9KLOGInspectTask-ES N9K日志检查",
        "ESFlowTask": "ESFlowTask-Flow服务器巡检",
        "DeviceBackupTask": "DeviceBackupTask-设备备份任务",
        "DeviceDIFFTask": "DeviceDIFFTask-设备差异检查",
        "ASACompareTask": "ASACompareTask-ASA防火墙对比检查",
        "ACLDupCheckTask": "ACLDupCheckTask-ACL重复检查任务",
        "ACLArpCheckTask": "ACLArpCheckTask-ACL无ARP匹配检查任务",
        "ACLCrossCheckTask": "ACLCrossCheckTask-N9K&LINKAS ACL交叉检查任务",
        "ASADomainCheckTask": "ASADomainCheckTask-ASA域名提取和检测任务",
        "ASATempnetworkCheckTask": "ASATempnetworkCheckTask-ASA临时出网地址检查",
        "ServiceCheckTask": "ServiceCheckTask-服务检查任务(NTP TACACS+)",
        "LogRecyclingTask": "LogRecyclingTask-日志回收任务（月底最后一天执行）"
    }
    
    # 打印任务开关状态
    if ENABLED_TASKS:
        print(f"启用的任务: {', '.join(ENABLED_TASKS)}")
    if DISABLED_TASKS:
        print(f"禁用的任务: {', '.join(DISABLED_TASKS)}")
    print(f"本次巡检将执行 {len(TASKS)} 个任务")
    print()  # 空行分隔

    # 记录开始时间
    START_TIME = time.time()

    # 检查REPORT文件是否存在，如果存在则读取现有内容
    EXISTING_CONTENT = ""
    if os.path.exists(DAILY_REPORT):
        with open(DAILY_REPORT, "r", encoding="utf-8") as f:
            EXISTING_CONTENT = f.read()
    
    # 生成时间戳
    TIMESTAMP = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    with open(DAILY_REPORT, "w", encoding="utf-8") as REPORT:
        # 写入时间隔离符号和任务开关状态
        REPORT.write(f"==================== {TIMESTAMP} 运行 ====================\n")
        REPORT.write("启用的任务:\n")
        if ENABLED_TASKS:
            for TASK in ENABLED_TASKS:
                REPORT.write(f"{TASK_NAMES.get(TASK, TASK)}\n")
        else:
            REPORT.write("无\n")
        REPORT.write("禁用的任务:\n")
        if DISABLED_TASKS:
            for TASK in DISABLED_TASKS:
                REPORT.write(f"{TASK_NAMES.get(TASK, TASK)}\n")
        else:
            REPORT.write("无\n")
        REPORT.write(f"本次巡检执行任务数: {len(TASKS)}\n\n")

        ALL_SUMMARY: dict = {}
        TOTAL_COUNTER = {"OK": 0, "WARN": 0, "CRIT": 0, "ERROR": 0}

        for TASK in TASKS:
            HEADER = f"\n=== 执行 {TASK.NAME} ==="
            if SHOW_PROGRESS:
                tqdm.write(HEADER)
            else:
                print(HEADER, flush=True)

            TASK.run()

            LEVEL_COUNT = {"OK": 0, "WARN": 0, "CRIT": 0, "ERROR": 0}
            for RESULT in TASK.RESULTS:
                LEVEL_COUNT[RESULT.level] = LEVEL_COUNT.get(RESULT.level, 0) + 1
                TOTAL_COUNTER[RESULT.level] = TOTAL_COUNTER.get(RESULT.level, 0) + 1

            REPORT.write(
                f"{TASK.NAME}：CRIT {LEVEL_COUNT['CRIT']}, WARN {LEVEL_COUNT['WARN']}, "
                f"ERROR {LEVEL_COUNT['ERROR']}, OK {LEVEL_COUNT['OK']}\n")
            for RESULT in TASK.RESULTS:
                if RESULT.level != "OK":
                    REPORT.write(f"  - [{RESULT.level}] {RESULT.message}\n")

            # V10新结构：日志文件保存在任务目录下，文件名格式为 YYYYMMDD-任务显示名.log
            # 目录名使用类名，文件名使用任务显示名（例如：FXOS WEB 巡检）
            TASK_CLASS_NAME = TASK.__class__.__name__
            TASK_LOG_DIR = os.path.join(BASE_LOG_DIR, TASK_CLASS_NAME)
            os.makedirs(TASK_LOG_DIR, exist_ok=True)
            TASK_LOG_PATH = os.path.join(TASK_LOG_DIR, f"{TODAY}-{TASK.NAME}.log")
            with open(TASK_LOG_PATH, "w", encoding="utf-8") as DETAIL_FILE:
                for RESULT in TASK.RESULTS:
                    LINE = RESULT.message
                    if RESULT.meta:
                        LINE += f" | {json.dumps(RESULT.meta, ensure_ascii=False)}"
                    DETAIL_FILE.write(f"[{RESULT.level}] {LINE}\n")

            ALL_SUMMARY[TASK.NAME] = {
                "SUMMARY": LEVEL_COUNT,
                "RESULTS": [asdict(RESULT) for RESULT in TASK.RESULTS],
            }

        # 计算总耗时
        END_TIME = time.time()
        TOTAL_ELAPSED = END_TIME - START_TIME
        TOTAL_MINUTES = int(TOTAL_ELAPSED // 60)
        TOTAL_SECONDS = int(TOTAL_ELAPSED % 60)
        
        # 格式化总耗时字符串
        if TOTAL_MINUTES > 0:
            ELAPSED_STR = f"{TOTAL_MINUTES}分{TOTAL_SECONDS:02d}秒"
        else:
            ELAPSED_STR = f"{TOTAL_SECONDS}秒"

        REPORT.write("\n=== 巡检总汇 ===\n")
        REPORT.write(
            f"严重 {TOTAL_COUNTER['CRIT']}, 告警 {TOTAL_COUNTER['WARN']}, "
            f"错误 {TOTAL_COUNTER['ERROR']}, 正常 {TOTAL_COUNTER['OK']}\n")
        REPORT.write(f"{TODAY} 全部任务完成\n")
        REPORT.write(f"总耗时: {ELAPSED_STR}\n")
        # V10新结构：日志输出在各任务目录下，此处记录根目录
        REPORT.write(f"日志目录: {BASE_LOG_DIR}\n")
        
        # 如果有现有内容，追加到文件末尾
        if EXISTING_CONTENT:
            REPORT.write("\n")
            REPORT.write(EXISTING_CONTENT)
        
        # 终端输出总耗时
        if SHOW_PROGRESS:
            tqdm.write(f"\n=== 全部任务完成 ===")
            tqdm.write(f"总耗时: {ELAPSED_STR}")
        else:
            print(f"\n=== 全部任务完成 ===")
            print(f"总耗时: {ELAPSED_STR}")

# ============================================================================
# 第三部分：主程序入口
# ============================================================================

# 运行预检或退出：检查依赖并处理预检结果
def _run_preflight_or_exit():
    """运行依赖预检，如果失败则退出程序"""
    try:
        OK, REPORT = check_runtime_dependencies()
        print(REPORT)
        if not OK:
            sys.exit(2)
    except SystemExit:
        raise
    except Exception as e:
        SAFE_REPORT = f"[预检异常] {e.__class__.__name__}: {e}\n{traceback.format_exc()}"
        print(SAFE_REPORT, file=sys.stderr)
        sys.exit(2)

# 主程序入口：执行预检和巡检任务
def main():
    """主程序入口：执行依赖预检和巡检任务"""
    # 1) 先跑依赖预检（这一步会有 print）
    _run_preflight_or_exit()

    # 2) 预检通过后，直接执行巡检任务
    run_inspection_tasks()

if __name__ == "__main__":
    # 支持两种运行模式：
    # 1. 直接运行：执行完整流程（依赖检查 + 巡检任务）
    # 2. --check 参数：只检查依赖，不执行巡检
    # 3. --install 参数：安装所有依赖包
    
    if len(sys.argv) > 1:
        if sys.argv[1] == "--check":
            # 检查模式：只检查不安装
            PLATFORM_NAME = _get_platform_name()
            print(f"=== 依赖检查模式 ({PLATFORM_NAME}) ===\n")
            OK, MSG = check_runtime_dependencies()
            if not OK:
                print(MSG)
                sys.exit(1)
            print(MSG)
        elif sys.argv[1] == "--install":
            # 安装模式：执行完整安装流程
            SUCCESS = install_all_dependencies()
            if not SUCCESS:
                print(f"\n错误：安装过程中出现错误，请检查上述输出信息")
                sys.exit(1)
            
            # 安装完成后进行验证
            print("\n=== 验证安装结果 ===")
            OK, MSG = check_runtime_dependencies()
            if not OK:
                print(f"\n警告：安装后仍有问题:\n{MSG}")
                sys.exit(1)
            print(MSG)
        else:
            print(f"未知参数: {sys.argv[1]}")
            print("用法: python Main.py [--check|--install]")
            sys.exit(1)
    else:
        # 默认模式：执行完整流程
        main()
