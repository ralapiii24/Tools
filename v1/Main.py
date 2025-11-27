# -*- coding: utf-8 -*-
"""
main.py
入口说明：
1) 先做 Windows 运行环境预检（Python 库 + Playwright 的 chromium）。
   —— 任一项缺失：直接退出（exit code=2），阻止自更新与巡检。
2) 预检通过后，进入原有流程：调用 Updater.check_update_then_run()。
"""

import os
import sys
import traceback


def _run_preflight_or_exit():
    """
    运行 Windows 依赖预检。失败则打印报告并以退出码 2 终止进程。
    """
    try:
        from RuntimeEnvCheck import check_windows_dependencies
    except Exception as e:
        safe_report = (
            "[预检致命] 无法导入 RuntimeEnvCheck 模块："
            f"{e.__class__.__name__}: {e}\n{traceback.format_exc()}"
        )
        print(safe_report, file=sys.stderr)
        sys.exit(2)

    try:
        ok, report = check_windows_dependencies()
        print(report)  # 这里会打印“Windows 依赖预检通过。”或缺失清单
        if not ok:
            sys.exit(2)
    except SystemExit:
        raise
    except Exception as e:
        safe_report = f"[预检异常] {e.__class__.__name__}: {e}\n{traceback.format_exc()}"
        print(safe_report, file=sys.stderr)
        sys.exit(2)


def main():
    # 1) 先跑依赖预检（这一步会有 print）
    _run_preflight_or_exit()

    # 2) 预检通过后，按原有逻辑进入自更新与巡检
    try:
        from Updater import check_update_then_run
    except Exception as e:
        err = (
            "[启动失败] 无法导入 Updater.check_update_then_run："
            f"{e.__class__.__name__}: {e}\n{traceback.format_exc()}"
        )
        print(err, file=sys.stderr)
        sys.exit(1)

    check_update_then_run()


if __name__ == "__main__":
    main()
