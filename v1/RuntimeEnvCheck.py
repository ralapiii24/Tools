from __future__ import annotations
import os
import sys
import traceback
from datetime import datetime

REQUIRED_PY_PKGS = [
    # (import模块名, 安装包名, 说明)
    ("yaml", "pyyaml", "YAML 解析"),
    ("tqdm", "tqdm", "进度条"),
    ("requests", "requests", "HTTP 请求"),
    ("lxml", "lxml", "XML/HTML 解析"),
    ("paramiko", "paramiko", "SSH"),
    # 与Playwright浏览器相关的操作
    ("playwright.sync_api", "playwright", "Web 自动化"),
    ("colorama", "colorama", "Windows 控制台着色"),
]

# 定义函数 _try_import
def _try_import(mod: str) -> tuple[bool, str | None]:
    try:
        __import__(mod)
        return True, None
    except Exception as e:
        return False, f"{e.__class__.__name__}: {e}"

# 定义函数 _check_playwright_chromium
def _check_playwright_chromium() -> tuple[bool, str | None]:
    """
    通过尝试 launch 的方式检查 chromium 是否已安装。
    若浏览器未安装，Playwright 会抛出“请运行 'playwright install'”之类的异常。
    注意：这里不会下载任何东西，只做存在性验证。
    """
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            # 仅启动/关闭以验证二进制存在；失败会抛 Missing browsers 相关异常
            browser = p.chromium.launch(headless=True)
            browser.close()
        return True, None
    except Exception as e:
        # 常见报错包含 'Please install' / 'playwright install'
        return False, f"{e.__class__.__name__}: {e}"

# 定义函数 check_windows_dependencies
def check_windows_dependencies() -> tuple[bool, str]:
    """
    返回 (ok, report)。ok=False 表示缺失，不允许继续更新与巡检。
    """
    if os.name != "nt":
        return True, "非 Windows 系统，跳过 Windows 依赖预检。"

    missing = []
    details = []

    # 1) 常规 Python 包
    for mod, pkg, desc in REQUIRED_PY_PKGS:
        ok, err = _try_import(mod)
        if not ok:
            missing.append(pkg)
            details.append(f"[缺] {pkg:<12}  ({desc})  import '{mod}' 失败 → {err}")

    # 2) 特殊：playwright 的 chromium 浏览器是否已安装
    if "playwright" not in missing:
        ok, err = _check_playwright_chromium()
        if not ok:
            missing.append("playwright-browsers(chromium)")
            details.append(f"[缺] chromium 浏览器 (Playwright)  未安装 → {err}")

    if not missing:
        return True, "=== Windows Python依赖包预检通过 ==="

    # 组装报告（含安装指引）
    lines = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines.append("=== " + f"[{now}] Windows 依赖检查未通过，阻止自更新与巡检 ===")
    lines.append("")
    lines.append("缺失项：")
    lines.extend(details)
    lines.append("")
    lines.append("修复指引（任选其一）：")
    lines.append("  1) 直接运行你仓库里的 PipLibrary.bat（推荐）")
    lines.append("  2) 或手动执行：")
    if any(p != "playwright-browsers(chromium)" for p in missing):
        # 仅列出真正的 pip 包
        pip_pkgs = sorted({p for p in missing if p != "playwright-browsers(chromium)"})
        if pip_pkgs:
            lines.append(f"     pip install -U {' '.join(pip_pkgs)}")
    if "playwright" in missing:
        lines.append("     （装完后再执行）python -m playwright install chromium")
    else:
        if "playwright-browsers(chromium)" in missing:
            lines.append("     python -m playwright install chromium")

    report = "\n".join(lines)
    return False, report

def write_dependency_report(report: str) -> None:
    """
    将预检结果落盘，方便查阅（不依赖现有日志框架，避免依赖链反噬）。
    """
    try:
        date = datetime.now().strftime("%Y%m%d")
        out_dir = os.path.join("REPORT")
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(out_dir, f"{date}_DependencyCheck.log")
        with open(path, "a", encoding="utf-8") as f:
            f.write(report + "\n")
    except Exception:
        # 仅兜底打印，不影响主流程退出
        print("[预检] 写入报告文件失败：\n" + traceback.format_exc(), file=sys.stderr)
