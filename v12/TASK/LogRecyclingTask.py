# 日志回收任务：定期清理和压缩历史日志文件
#
# 技术栈:Python, datetime, os, shutil, zipfile, 文件系统操作
# 目标:定期清理和压缩历史日志文件，自动管理LOG、REPORT、UPGRADELOG目录（V10: ACL/DOMAIN/CONFIGURATION已迁移至LOG下对应任务目录）
#
# 执行时机:月底最后一天自动执行（默认规则写死在代码中），或通过Config.yaml配置force_run_date指定执行日期
# 处理逻辑:日期判断 → 180天清理 → 规则1删除 → 规则2压缩 → 统计汇总
#
# 180天清理规则（V10新增）:
# - 执行优先级:最高优先级，优先于规则1和规则2执行
# - 清理范围:递归遍历LOG目录下的所有文件（包括所有任务子目录）
# - 清理条件:删除文件名以YYYYMMDD-开头的文件，且文件日期超过180天
# - 文件类型:支持所有文件类型（.log、.xlsx、.zip等）
# - 保护机制:排除今天和昨天的日志，避免与写入冲突
#
# 规则1（删除规则）:
# - 执行条件:月底最后一天执行
# - 保留策略:保留周一、周日、每月1号、31号的文件/目录
# - 删除策略:删除其他日期的文件/目录
# - 保护机制:排除今天和昨天的日志，避免与写入冲突
# - 处理目录（V10新结构）:
#   * LOG目录下的任务子目录中的所有文件（递归遍历）
#   * REPORT目录下的文件
#
# 规则2（压缩规则）:
# - 保留策略:保留本月和上个月的文件/目录
# - 压缩策略:压缩上上个月及更早的文件/目录，同月文件压缩到一个zip包
# - 压缩格式:YYYYMM-月份归档.zip（如202508-月份归档.zip）
# - 压缩位置:压缩包保存在原文件/目录的父目录中（与原文件/目录同级）
# - 删除策略:压缩成功后删除已被压缩的原文件/目录
# - 安全保护:明确排除.zip文件，防止误删过往月份的压缩包
#
# UPGRADELOG特殊处理:
# - 无条件删除UPGRADELOG目录下的所有文件
# - 不执行规则1和规则2，直接删除
#
# 配置说明:
# - Config.yaml中LogRecyclingTask配置节:
#   * force_run_date: 指定运行日期（YYYYMMDD格式，字符串或数字）
#     - null或不设置: 使用默认规则（月底最后一天执行）
#     - 指定日期: 只在指定日期执行
# - 从settings读取log_dir和report_dir配置
#
# 输出:统计删除项数、压缩项数、错误项数，记录到巡检日报

# 导入标准库
import datetime as dt
import os
import shutil
import zipfile

from typing import Optional

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
from .TaskBase import BaseTask, Level, CONFIG

# 日志回收任务类：根据规则删除和压缩历史文件/目录
class LogRecyclingTask(BaseTask):
    """日志回收任务

    定期清理和压缩历史日志文件，自动管理LOG、REPORT、UPGRADELOG目录
    """

    # 初始化日志回收任务：设置任务名称和配置
    def __init__(self):
        super().__init__("日志回收任务")
        # 从配置读取目录设置（如果存在）
        SETTINGS = CONFIG.get("settings", {})
        self.LOG_DIR = SETTINGS.get("log_dir", "LOG")
        self.REPORT_DIR = SETTINGS.get("report_dir", "REPORT")


        # 固定目录路径
        self.ACL_DIR = "ACL"
        self.DOMAIN_DIR = "DOMAIN"
        self.CONFIGURATION_DIR = "CONFIGURATION"
        self.UPGRADELOG_DIR = "UPGRADELOG"


        # 从配置读取指定运行日期
        TASK_CONFIG = CONFIG.get("LogRecyclingTask", {})
        FORCE_RUN_DATE_RAW = TASK_CONFIG.get("force_run_date")


        TODAY = dt.date.today()
        TODAY_STRING = TODAY.strftime("%Y%m%d")


        # 判断是否应该执行
        if FORCE_RUN_DATE_RAW is not None:
            # 如果配置了指定日期，先转换为字符串（YAML可能将纯数字解析为整数）
            try:
                FORCE_RUN_DATE_STR = str(FORCE_RUN_DATE_RAW)
                # 确保是8位数字字符串
                if len(FORCE_RUN_DATE_STR) != 8 or not FORCE_RUN_DATE_STR.isdigit():
                    raise ValueError(f"日期格式错误: {FORCE_RUN_DATE_STR}")
                FORCE_RUN_DATE = dt.datetime.strptime(FORCE_RUN_DATE_STR, "%Y%m%d").date()
                self.IS_MONTH_END = (TODAY == FORCE_RUN_DATE)
                if self.IS_MONTH_END:
                    self.add_result(Level.OK, f"配置指定运行日期: {FORCE_RUN_DATE_STR}，当前日期匹配，执行回收任务")
                else:
                    self.add_result(
                        Level.WARN,
                        f"配置指定运行日期: {FORCE_RUN_DATE_STR}，"
                        f"当前日期 {TODAY_STRING} 不匹配，跳过日志回收任务"
                    )
            except (ValueError, TypeError) as ERROR:
                self.add_result(
                    Level.ERROR,
                    f"配置的日期格式错误: {FORCE_RUN_DATE_RAW}，"
                    f"应为YYYYMMDD格式（字符串或数字），跳过日志回收任务"
                )
                self.IS_MONTH_END = False
        else:
            # 如果未配置指定日期，使用默认规则（月底最后一天）
            TOMORROW = TODAY + dt.timedelta(days=1)
            self.IS_MONTH_END = (TOMORROW.month != TODAY.month)
            if not self.IS_MONTH_END:
                self.add_result(Level.WARN, f"当前不是月底最后一天（当前日期: {TODAY_STRING}），跳过日志回收任务")


    # 尝试安全删除文件：支持 Windows 文件占用的重试与待删除标记
    def _safe_remove_file(self, file_path: str) -> bool:
        import time
        # 优先多次重试删除
        for _ in range(5):
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                return True
            except PermissionError:
                time.sleep(0.5)
            except Exception:
                break
        # 无法直接删除时，尝试标记为待删除
        try:
            if os.path.exists(file_path):
                pending_path = file_path + ".pending_delete"
                if not os.path.exists(pending_path):
                    os.replace(file_path, pending_path)
                self.add_result(Level.WARN, f"文件被占用，已标记待删除: {pending_path}")
            return True
        except Exception as error:
            self.add_result(Level.ERROR, f"删除失败 {file_path}: {error}")
            return False

    # 获取要处理的任务项列表：返回需要处理的目录类型列表（用于分组处理）
    def items(self):
        """返回需要处理的目录类型列表

        Returns:
            目录类型列表，只在月底最后一天返回非空列表
        """
        if not self.IS_MONTH_END:
            return []


        # 返回目录类型列表，用于分组处理
        CATEGORIES = []
        if os.path.exists(self.LOG_DIR):
            CATEGORIES.append("LOG_DIR")
        if os.path.exists(self.REPORT_DIR):
            CATEGORIES.append("REPORT_FILE")
        # V10: ACL/DOMAIN/CONFIGURATION 已迁移至 LOG 下对应任务目录，不再单独处理
        if os.path.exists(self.UPGRADELOG_DIR):
            CATEGORIES.append("UPGRADELOG_FILE")


        return CATEGORIES


    # 收集指定类型的所有路径：收集某个目录类型下的所有文件/目录路径
    # 收集指定类型的所有路径
    def _collect_items_by_category(self, category: str) -> list[tuple[str, str]]:
        ITEMS = []


        if category == "LOG_DIR":
            # LOG目录下的任务目录，收集每个任务目录下的日志文件
            if os.path.exists(self.LOG_DIR):
                for TASK_DIR_NAME in os.listdir(self.LOG_DIR):
                    TASK_DIR_PATH = os.path.join(self.LOG_DIR, TASK_DIR_NAME)
                    if os.path.isdir(TASK_DIR_PATH):
                        # 遍历任务目录下的所有文件
                        for FILE_NAME in os.listdir(TASK_DIR_PATH):
                            FILE_PATH = os.path.join(TASK_DIR_PATH, FILE_NAME)
                            if os.path.isfile(FILE_PATH):
                                ITEMS.append(("LOG_FILE", FILE_PATH))


        elif category == "REPORT_FILE":
            # REPORT目录下的文件
            if os.path.exists(self.REPORT_DIR):
                for ITEM_NAME in os.listdir(self.REPORT_DIR):
                    ITEM_PATH = os.path.join(self.REPORT_DIR, ITEM_NAME)
                    if os.path.isfile(ITEM_PATH):
                        ITEMS.append(("REPORT_FILE", ITEM_PATH))


        # V10: ACL/DOMAIN/CONFIGURATION 已迁移至 LOG 下对应任务目录，不再单独处理


        elif category == "UPGRADELOG_FILE":
            # UPGRADELOG目录下的所有文件（无条件删除）
            if os.path.exists(self.UPGRADELOG_DIR):
                for ITEM_NAME in os.listdir(self.UPGRADELOG_DIR):
                    ITEM_PATH = os.path.join(self.UPGRADELOG_DIR, ITEM_NAME)
                    if os.path.isfile(ITEM_PATH):
                        ITEMS.append(("UPGRADELOG_FILE", ITEM_PATH))


        return ITEMS


    # 从路径中提取日期：从文件名或目录名中提取YYYYMMDD格式的日期
    # 从路径中提取日期（支持YYYYMMDD格式）
    def _extract_date_from_path(self, path: str) -> Optional[dt.date]:
        import re
        BASENAME = os.path.basename(path)


        # 尝试匹配YYYYMMDD格式（可能在文件名开头或目录名）
        DATE_MATCH = re.match(r"^(\d{8})", BASENAME)
        if DATE_MATCH:
            try:
                DATE_STR = DATE_MATCH.group(1)
                return dt.datetime.strptime(DATE_STR, "%Y%m%d").date()
            except ValueError:
                pass


        # 如果目录名本身就是日期格式（8位数字）
        if len(BASENAME) == 8 and BASENAME.isdigit():
            try:
                return dt.datetime.strptime(BASENAME, "%Y%m%d").date()
            except ValueError:
                pass


        return None


    # 判断日期是否应该保留（规则1）：周一、周日、1号、31号保留
    # 判断日期是否应该保留（周一、周日、1号、31号）
    def _should_keep_date(self, date: dt.date) -> bool:
        # 周一：weekday() == 0
        # 周日：weekday() == 6
        # 1号：day == 1
        # 31号：day == 31
        return (date.weekday() == 0 or  # 周一
                date.weekday() == 6 or  # 周日
                date.day == 1 or        # 1号
                date.day == 31)         # 31号


    # 判断日期是否应该压缩（规则2）：上上个月及更早的日期
    # 判断日期是否应该压缩（上上个月及更早）
    def _should_compress_date(self, date: dt.date) -> bool:
        TODAY = dt.date.today()


        # 本月
        if date.year == TODAY.year and date.month == TODAY.month:
            return False


        # 上个月
        LAST_MONTH = (TODAY.month - 1) if TODAY.month > 1 else 12
        LAST_YEAR = TODAY.year if TODAY.month > 1 else (TODAY.year - 1)
        if date.year == LAST_YEAR and date.month == LAST_MONTH:
            return False


        # 上上个月及更早：需要压缩
        return True


    # 判断文件是否超过180天
    # 判断文件日期是否超过180天
    def _is_older_than_180_days(self, date: dt.date) -> bool:
        TODAY = dt.date.today()
        DAYS_AGO = TODAY - dt.timedelta(days=180)
        return date < DAYS_AGO


    # 递归收集LOG目录下的所有文件（包括子目录）
    def _collect_all_files_in_log_dir(self) -> list[tuple[str, str]]:
        ITEMS = []
        if not os.path.exists(self.LOG_DIR):
            return ITEMS


        for ROOT, DIRS, FILES in os.walk(self.LOG_DIR):
            for FILE_NAME in FILES:
                FILE_PATH = os.path.join(ROOT, FILE_NAME)
                if os.path.isfile(FILE_PATH):
                    ITEMS.append(("LOG_FILE", FILE_PATH))


        return ITEMS


    # 执行180天清理：删除超过180天的文件（包括所有文件类型：.log、.xlsx、.zip等）
    def _apply_180_days_cleanup(self) -> None:
        ALL_FILES = self._collect_all_files_in_log_dir()
        DELETED_COUNT = 0


        for ITEM_TYPE, ITEM_PATH in ALL_FILES:
            DATE = self._extract_date_from_path(ITEM_PATH)
            if not DATE:
                # 无法提取日期，跳过（文件名不符合 YYYYMMDD-文件名后缀 格式）
                continue


            # 检查是否超过180天
            if self._is_older_than_180_days(DATE):
                # 删除超过180天的文件（所有文件类型，包括 .log、.xlsx、.zip 等）
                if self._safe_remove_file(ITEM_PATH):
                    if not ITEM_PATH.endswith('.pending_delete') and not os.path.exists(ITEM_PATH):
                        DELETED_COUNT += 1
                        self.add_result(Level.OK, f"删除超过180天的文件: {ITEM_PATH}")


        if DELETED_COUNT > 0:
            self.add_result(Level.OK, f"180天清理完成：删除 {DELETED_COUNT} 个文件")


    # 执行规则1：删除不应该保留的文件/目录
    def _apply_rule1_delete(self, item_type: str, item_path: str) -> bool:
        # 排除压缩包文件（避免误删已生成的压缩包）
        if item_path.lower().endswith('.zip'):
            return False
        # 跳过当天的巡检日报，避免与写入冲突
        try:
            BASENAME = os.path.basename(item_path)
            TODAY_STRING = dt.date.today().strftime("%Y%m%d")
            if (item_type == "REPORT_FILE" and
                    BASENAME.startswith(TODAY_STRING) and
                    BASENAME.endswith("巡检日报.log")):
                return False
        except Exception:
            pass


        DATE = self._extract_date_from_path(item_path)
        if not DATE:
            # 无法提取日期，跳过
            return False


        # 排除今天和昨天的日志，避免与写入冲突
        TODAY = dt.date.today()
        YESTERDAY = TODAY - dt.timedelta(days=1)
        if DATE == TODAY or DATE == YESTERDAY:
            return False


        if self._should_keep_date(DATE):
            # 应该保留，不删除
            return False


        # 应该删除
        try:
            if item_type in ["LOG_DIR", "CONFIGURATION_DIR"]:
                # 删除目录
                shutil.rmtree(item_path)
                self.add_result(Level.OK, f"删除目录: {item_path}")
            else:
                # 删除文件
                if self._safe_remove_file(item_path):
                    # _safe_remove_file 内部已记录 WARN/ERROR；仅在真正删除时记 OK
                    if not item_path.endswith('.pending_delete') and not os.path.exists(item_path):
                        self.add_result(Level.OK, f"删除文件: {item_path}")
            return True
        except Exception as ERROR:
            self.add_result(Level.ERROR, f"删除失败 {item_path}: {ERROR}")
            return False


    # 执行规则2：按月份批量压缩上上个月及更早的文件/目录
    # （同一月的文件/目录压缩到一个zip包中）
    def _apply_rule2_compress_by_month(
        self, category: str, items_by_month: dict[str, list[tuple[str, str]]],
        parent_dir: str = None
    ) -> None:
        for MONTH_KEY, ITEMS in items_by_month.items():
            if not ITEMS:
                continue


            # 获取压缩文件存放位置
            if parent_dir is None:
                # 如果没有传入parent_dir，使用第一个项的父目录
                FIRST_ITEM_TYPE, FIRST_ITEM_PATH = ITEMS[0]
                PARENT_DIR = os.path.dirname(FIRST_ITEM_PATH)
            else:
                PARENT_DIR = parent_dir


            # 压缩文件名：YYYYMM-月份归档.zip
            ARCHIVE_NAME = f"{MONTH_KEY}-月份归档.zip"
            ARCHIVE_PATH = os.path.join(PARENT_DIR, ARCHIVE_NAME)


            try:
                # 第一步：先压缩所有文件/目录到zip包
                COMPRESSED_ITEMS = []  # 记录已成功压缩的项目，用于后续删除
                with zipfile.ZipFile(ARCHIVE_PATH, 'w', zipfile.ZIP_DEFLATED) as ZIPF:
                    for ITEM_TYPE, ITEM_PATH in ITEMS:
                        # 再次检查文件/目录是否仍然存在
                        if not os.path.exists(ITEM_PATH):
                            continue


                        try:
                            if ITEM_TYPE in ["LOG_DIR", "CONFIGURATION_DIR"]:
                                # 压缩目录：保持目录结构
                                BASE_NAME = os.path.basename(ITEM_PATH)
                                for ROOT, DIRS, FILES in os.walk(ITEM_PATH):
                                    for FILE in FILES:
                                        FILE_PATH = os.path.join(ROOT, FILE)
                                        # 保持相对路径结构：BASE_NAME/相对路径
                                        # 使用相对于ITEM_PATH的路径，而不是相对于父目录的路径
                                        REL_PATH = os.path.relpath(ROOT, ITEM_PATH)
                                        if REL_PATH == '.':
                                            ARCNAME = os.path.join(BASE_NAME, FILE)
                                        else:
                                            ARCNAME = os.path.join(BASE_NAME, REL_PATH, FILE)
                                        ZIPF.write(FILE_PATH, ARCNAME)
                                # 记录已压缩的目录
                                COMPRESSED_ITEMS.append((ITEM_TYPE, ITEM_PATH))
                            else:
                                # 压缩文件
                                ZIPF.write(ITEM_PATH, os.path.basename(ITEM_PATH))
                                # 记录已压缩的文件
                                COMPRESSED_ITEMS.append((ITEM_TYPE, ITEM_PATH))
                        except Exception as ITEM_ERROR:
                            # 单个项目压缩失败，记录错误但继续处理其他项目
                            self.add_result(Level.ERROR, f"压缩项目失败 {ITEM_PATH}: {ITEM_ERROR}")


                # 第二步：压缩成功后，删除已被压缩的原文件/目录
                DELETED_COUNT = 0
                for ITEM_TYPE, ITEM_PATH in COMPRESSED_ITEMS:
                    if not os.path.exists(ITEM_PATH):
                        continue
                    try:
                        if ITEM_TYPE in ["LOG_DIR", "CONFIGURATION_DIR"]:
                            # 删除目录
                            shutil.rmtree(ITEM_PATH)
                            DELETED_COUNT += 1
                        else:
                            # 删除文件
                            if self._safe_remove_file(ITEM_PATH):
                                if (not ITEM_PATH.endswith('.pending_delete') and
                                    not os.path.exists(ITEM_PATH)):
                                    DELETED_COUNT += 1
                    except Exception as DELETE_ERROR:
                        self.add_result(Level.ERROR, f"删除已压缩文件/目录失败 {ITEM_PATH}: {DELETE_ERROR}")


                if DELETED_COUNT > 0:
                    self.add_result(
                        Level.OK,
                        f"压缩并删除 {MONTH_KEY} 月 {DELETED_COUNT} 项到: "
                        f"{ARCHIVE_PATH}"
                    )
            except Exception as ERROR:
                self.add_result(Level.ERROR, f"压缩失败 {MONTH_KEY} 月归档: {ERROR}")


    # 处理单个目录类型：先执行规则1删除，再执行规则2压缩
    def run_single(self, category: str) -> None:
        """处理单个目录类型

        根据目录类型执行清理或压缩操作

        Args:
            category: 目录类型（如"LOG_DIR"、"REPORT_DIR"等）
        """
        if not self.IS_MONTH_END:
            return


        # UPGRADELOG特殊处理：无条件删除所有文件
        if category == "UPGRADELOG_FILE":
            ALL_ITEMS = self._collect_items_by_category(category)
            for ITEM_TYPE, ITEM_PATH in ALL_ITEMS:
                if os.path.exists(ITEM_PATH):
                    try:
                        if self._safe_remove_file(ITEM_PATH):
                            if (not ITEM_PATH.endswith('.pending_delete') and
                                    not os.path.exists(ITEM_PATH)):
                                self.add_result(Level.OK, f"删除文件: {ITEM_PATH}")
                    except Exception as ERROR:
                        self.add_result(Level.ERROR, f"删除失败 {ITEM_PATH}: {ERROR}")
            return


        # LOG_DIR特殊处理：先执行180天清理
        if category == "LOG_DIR":
            self._apply_180_days_cleanup()


        # 收集该类型的所有路径
        ALL_ITEMS = self._collect_items_by_category(category)
        if not ALL_ITEMS:
            return


        # 第一步：执行规则1删除
        REMAINING_ITEMS = []
        for ITEM_TYPE, ITEM_PATH in ALL_ITEMS:
            # 检查文件/目录是否仍然存在（可能已被删除）
            if not os.path.exists(ITEM_PATH):
                continue


            # 执行规则1删除
            if not self._apply_rule1_delete(ITEM_TYPE, ITEM_PATH):
                # 如果未删除，加入剩余列表
                if os.path.exists(ITEM_PATH):  # 再次检查，确保仍然存在
                    REMAINING_ITEMS.append((ITEM_TYPE, ITEM_PATH))


        # 第二步：对剩余的文件/目录执行规则2压缩（按目录和月份分组）
        # 按目录和月份分组需要压缩的项：{parent_dir: {YYYYMM: [(item_type, item_path), ...]}}
        ITEMS_BY_DIR_MONTH = {}  # {parent_dir: {YYYYMM: [(item_type, item_path), ...]}}
        for ITEM_TYPE, ITEM_PATH in REMAINING_ITEMS:
            # 再次检查文件/目录是否仍然存在
            if not os.path.exists(ITEM_PATH):
                continue


            # 排除压缩包文件（避免处理已生成的压缩包）
            if ITEM_PATH.lower().endswith('.zip'):
                continue


            DATE = self._extract_date_from_path(ITEM_PATH)
            if not DATE:
                # 无法提取日期，跳过
                continue


            # 排除今天和昨天的日志，避免与写入冲突
            TODAY = dt.date.today()
            YESTERDAY = TODAY - dt.timedelta(days=1)
            if DATE == TODAY or DATE == YESTERDAY:
                continue


            if not self._should_compress_date(DATE):
                # 不需要压缩（本月或上个月），跳过
                continue


            # 获取父目录（用于分组）
            PARENT_DIR = os.path.dirname(ITEM_PATH)


            # 按月份分组（YYYYMM格式）
            MONTH_KEY = DATE.strftime("%Y%m")


            if PARENT_DIR not in ITEMS_BY_DIR_MONTH:
                ITEMS_BY_DIR_MONTH[PARENT_DIR] = {}
            if MONTH_KEY not in ITEMS_BY_DIR_MONTH[PARENT_DIR]:
                ITEMS_BY_DIR_MONTH[PARENT_DIR][MONTH_KEY] = []
            ITEMS_BY_DIR_MONTH[PARENT_DIR][MONTH_KEY].append((ITEM_TYPE, ITEM_PATH))


        # 按目录和月份批量压缩
        for PARENT_DIR, ITEMS_BY_MONTH in ITEMS_BY_DIR_MONTH.items():
            if ITEMS_BY_MONTH:
                self._apply_rule2_compress_by_month(category, ITEMS_BY_MONTH, PARENT_DIR)


    # 重写run方法：只在月底最后一天执行
    def run(self) -> None:
        """执行日志回收任务

        只在月底最后一天执行回收任务，调用父类的run方法执行回收任务
        """
        if not self.IS_MONTH_END:
            return


        # 调用父类的run方法执行回收任务（会遍历items并调用run_single）
        super().run()


        # 执行完成后输出统计信息
        DELETE_COUNT = sum(
            1 for R in self.RESULTS
            if "删除" in R.message and "压缩并删除" not in R.message
        )
        COMPRESS_COUNT = sum(1 for R in self.RESULTS if "压缩并删除" in R.message)
        ERROR_COUNT = sum(1 for R in self.RESULTS if R.level == Level.ERROR.value)


        self.add_result(
            Level.OK,
            f"日志回收完成：删除 {DELETE_COUNT} 项，"
            f"压缩 {COMPRESS_COUNT} 项，错误 {ERROR_COUNT} 项"
        )

