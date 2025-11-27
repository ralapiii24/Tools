# 主调度器：执行所有启用的巡检任务并生成日报

# 导入标准库
import json
import os
import sys
import time
from dataclasses import asdict
from datetime import datetime

# 导入第三方库
import yaml
from tqdm import tqdm

# 导入本地应用
# (正常模式才加载)
from TASK import __all__ as TASK_CLASSES
from TASK.TaskBase import require_keys

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

# 只检查Core.py需要的通用配置
require_keys(CONFIG, ["settings"], "root")

SHOW_PROGRESS = bool(CONFIG["settings"].get("show_progress", True))

# 主调度器

# 主程序入口：执行所有启用的巡检任务
def main():
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
        
        # 终端输出总耗时
        if SHOW_PROGRESS:
            tqdm.write(f"\n=== 全部任务完成 ===")
            tqdm.write(f"总耗时: {ELAPSED_STR}")
        else:
            print(f"\n=== 全部任务完成 ===")
            print(f"总耗时: {ELAPSED_STR}")

if __name__ == "__main__":
    main()
