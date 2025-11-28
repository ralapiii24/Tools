# 运行时环境检查：检查Python版本和依赖包

# 导入标准库
import os
import sys
import subprocess
import traceback
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

# 导入第三方库
import yaml

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
PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_FILE = PROJECT_ROOT / "YAML" / "Config.yaml"

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

# 写入依赖检查报告：将报告内容写入REPORT目录下的文件，不依赖现有日志框架
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

if __name__ == "__main__":
    # 支持两种运行模式：
    # 1. 直接运行：执行完整安装流程（整合 PipLibrary 功能）
    # 2. 作为模块导入：提供 check_runtime_dependencies() 函数供其他模块调用
    # 
    # 支持平台：Windows、Linux、macOS
    
    PLATFORM_NAME = _get_platform_name()
    
    if len(sys.argv) > 1 and sys.argv[1] == "--check":
        # 检查模式：只检查不安装
        print(f"=== 依赖检查模式 ({PLATFORM_NAME}) ===\n")
        OK, MSG = check_runtime_dependencies()
        if not OK:
            print(MSG)
            sys.exit(1)
        print(MSG)
    else:
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
