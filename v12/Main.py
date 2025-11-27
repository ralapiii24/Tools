# 主程序入口文件

# 导入标准库
import sys
import traceback

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
# (动态导入，在函数内部导入)

# 运行预检或退出：检查Windows依赖并处理预检结果
def _run_preflight_or_exit():
    try:
        from APP.RuntimeEnvCheck import check_runtime_dependencies
    except Exception as e:
        SAFE_REPORT = (
            "[预检致命] 无法导入 RuntimeEnvCheck 模块："
            f"{e.__class__.__name__}: {e}\n{traceback.format_exc()}"
        )
        print(SAFE_REPORT, file=sys.stderr)
        sys.exit(2)

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

# 主程序入口：执行预检、更新检查和巡检任务
def main():
    # 1) 先跑依赖预检（这一步会有 print）
    _run_preflight_or_exit()

    # 2) 预检通过后，按原有逻辑进入自更新与巡检
    try:
        from APP.Updater import check_update_then_run
    except Exception as e:
        ERR = (
            "[启动失败] 无法导入 Updater.check_update_then_run："
            f"{e.__class__.__name__}: {e}\n{traceback.format_exc()}"
        )
        print(ERR, file=sys.stderr)
        sys.exit(1)

    check_update_then_run()

if __name__ == "__main__":
    main()
