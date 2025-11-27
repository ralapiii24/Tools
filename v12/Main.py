# 主程序入口文件

# 导入标准库
import sys
import traceback

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
# (动态导入，在函数内部导入)

# 主程序入口：执行所有启用的巡检任务并生成日报

# 导入标准库
import concurrent.futures
import json
import logging
import os
import sys
import time
from dataclasses import asdict
from datetime import datetime

# 导入第三方库
import yaml
from tqdm import tqdm

# 导入本地应用
from TASK import __all__ as TASK_CLASSES
from TASK.TaskBase import require_keys

PARALLEL_FLAG = "--parallel"

# 动态导入所有任务类
TASK_MODULES = {}
for TASK_NAME in TASK_CLASSES:
    try:
        TASK_MODULE = __import__(f'TASK.{TASK_NAME}', fromlist=[TASK_NAME])
        TASK_MODULES[TASK_NAME] = getattr(TASK_MODULE, TASK_NAME)
    except Exception as e:
        print(f"警告：无法导入任务类 {TASK_NAME}: {e}")

# 读取Config.yaml配置
with open("YAML/Config.yaml", "r", encoding="utf-8") as f:
    CONFIG = yaml.safe_load(f)

# 只检查Main.py/巡检核心需要的通用配置
require_keys(CONFIG, ["settings"], "root")

SHOW_PROGRESS = bool(CONFIG["settings"].get("show_progress", True))
logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(message)s")
LOGGER = logging.getLogger(__name__)


def _get_task_cls(name: str):
    return TASK_MODULES[name]


def _run_task(name: str) -> None:
    task_cls = _get_task_cls(name)
    LOGGER.info("启动任务：%s", name)
    task = task_cls()
    task.run()
    LOGGER.info("任务完成：%s", name)


def _run_parallel(names: list[str]) -> None:
    if not names:
        return
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(names)) as executor:
        futures = {executor.submit(_run_task, name): name for name in names}
        for future in concurrent.futures.as_completed(futures):
            future.result()


def _parallel_inspection() -> None:
    _run_task("OxidizedTask")
    _run_task("DeviceBackupTask")

    _run_parallel([
        "ASACompareTask",
        "DeviceDIFFTask",
        "ACLCrossCheckTask",
        "ACLDupCheckTask",
        "ASADomainCheckTask",
    ])

    _run_task("ACLArpCheckTask")

    _run_task("FXOSWebTask")
    _run_task("MirrorFortiGateTask")

    for name in ("ESLogstashTask", "ESBaseTask", "ESFlowTask", "ESN9KLOGInspectTask", "ServiceCheckTask"):
        _run_task(name)

    _run_task("LogRecyclingTask")


def _report_header():
    TODAY = datetime.now().strftime("%Y%m%d")
    SETTINGS = CONFIG.get("settings") or {}
    BASE_LOG_DIR = SETTINGS.get("log_dir", "LOG")
    REPORT_DIR = SETTINGS.get("report_dir", "REPORT")
    os.makedirs(REPORT_DIR, exist_ok=True)
    DAILY_REPORT = os.path.join(REPORT_DIR, f"{TODAY}巡检日报.log")
    return TODAY, SETTINGS, BASE_LOG_DIR, REPORT_DIR, DAILY_REPORT


def _gather_tasks():
    TASK_SWITCHES = CONFIG.get("task_switches", {})
    TASKS = []
    ENABLED_TASKS = []
    DISABLED_TASKS = []

    for TASK_NAME in TASK_CLASSES:
        IS_ENABLED = TASK_SWITCHES.get(TASK_NAME, True)
        if not IS_ENABLED:
            DISABLED_TASKS.append(TASK_NAME)
            continue

        try:
            TASK_CLASS = TASK_MODULES[TASK_NAME]
            TASK_INSTANCE = TASK_CLASS()
            TASKS.append(TASK_INSTANCE)
            ENABLED_TASKS.append(TASK_NAME)
        except Exception as e:
            print(f"警告：无法创建任务实例 {TASK_NAME}: {e}")

    return TASKS, ENABLED_TASKS, DISABLED_TASKS


def main():
    if PARALLEL_FLAG in sys.argv[1:]:
        _parallel_inspection()
        return

    TODAY, SETTINGS, BASE_LOG_DIR, REPORT_DIR, DAILY_REPORT = _report_header()
    TASKS, ENABLED_TASKS, DISABLED_TASKS = _gather_tasks()

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
        "ServiceCheckTask": "ServiceCheckTask-服务检查任务(NTP TACACS+)",
        "LogRecyclingTask": "LogRecyclingTask-日志回收任务（月底最后一天执行）"
    }

    if ENABLED_TASKS:
        print(f"启用的任务: {', '.join(ENABLED_TASKS)}")
    if DISABLED_TASKS:
        print(f"禁用的任务: {', '.join(DISABLED_TASKS)}")
    print(f"本次巡检将执行 {len(TASKS)} 个任务")
    print()

    START_TIME = time.time()
    TIMESTAMP = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ALL_SUMMARY = {}
    TOTAL_COUNTER = {"OK": 0, "WARN": 0, "CRIT": 0, "ERROR": 0}

    with open(DAILY_REPORT, "w", encoding="utf-8") as REPORT:
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

        END_TIME = time.time()
        TOTAL_ELAPSED = END_TIME - START_TIME
        TOTAL_MINUTES = int(TOTAL_ELAPSED // 60)
        TOTAL_SECONDS = int(TOTAL_ELAPSED % 60)
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
        REPORT.write(f"日志目录: {BASE_LOG_DIR}\n")

        if SHOW_PROGRESS:
            tqdm.write(f"\n=== 全部任务完成 ===")
            tqdm.write(f"总耗时: {ELAPSED_STR}")
        else:
            print(f"\n=== 全部任务完成 ===")
            print(f"总耗时: {ELAPSED_STR}")


if __name__ == "__main__":
    main()
