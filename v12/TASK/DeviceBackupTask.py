# 关键设备配置备份输出EXCEL基础任务

# 导入标准库
import os
import re
from datetime import datetime
from typing import Optional

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
from .TaskBase import BaseTask, Level, CONFIG, extract_site_from_filename, safe_sheet_name, require_keys
from .CiscoBase import classify_text_as_cat, get_cat_classification_patterns

# 关键设备配置备份任务类：将设备配置备份并输出为Excel格式
class DeviceBackupTask(BaseTask):
    
    # 初始化关键设备配置备份任务：设置输出目录和数据结构
    def __init__(self):
        super().__init__("关键设备配置备份输出EXCEL基础任务")
        # 从配置文件读取log_dir（必须配置）
        require_keys(CONFIG, ["settings"], "root")
        require_keys(CONFIG["settings"], ["log_dir"], "settings")
        self.LOG_DIR = CONFIG["settings"]["log_dir"]
        # V10新结构：直接输出到 LOG/DeviceBackupTask/
        self.OUTPUT_DIR = os.path.join(self.LOG_DIR, "DeviceBackupTask")

        # ↓↓↓ 为 1/N 进度新增的最小状态（其余逻辑不变）
        self._TODAY = None
        self._IN_DIR = None
        self._XLSX_PATH = None
        self._GROUPED_PATHS: dict[str, dict[str, list[str]]] = {}  # site -> {cat1:[path...], cat2:[...], cat3:[...]}
        self._SITES_ORDER: list[str] = []
        self._WB = None
        self._TOTAL_DEVICES = 0

    # 扫描LOG目录获取站点列表：按站点分组N9K ASA设备配置文件
    def items(self):
        from openpyxl import Workbook

        self._TODAY = datetime.now().strftime("%Y%m%d")
        # V10新结构：从 LOG/OxidizedTask/OxidizedTaskBackup/ 读取
        self._IN_DIR = os.path.join(self.LOG_DIR, "OxidizedTask", "OxidizedTaskBackup")
        if not os.path.isdir(self._IN_DIR):
            self.add_result(Level.ERROR, f"未找到当日 Oxidized 日志目录: {self._IN_DIR}")
            return []

        # 创建输出目录（如果目录已存在则不报错）
        os.makedirs(self.OUTPUT_DIR, exist_ok=True)
        self._XLSX_PATH = os.path.join(
            self.OUTPUT_DIR,
            f"{self._TODAY}-关键设备配置备份输出EXCEL基础任务.xlsx"
        )

        # 分组处理：根据设备分类规则进行分组
        GROUPED_PATHS: dict[str, dict[str, list[str]]] = {}
        RULES = self._get_device_classification_rules()
        
        # 只处理当天的文件：检查文件名是否以当天日期开头
        TODAY_STR = self._TODAY  # 格式：YYYYMMDD
        # 只匹配格式: YYYYMMDD-设备名.log（日期在文件名开头）
        DATE_PATTERN = re.compile(r'^' + re.escape(TODAY_STR) + r'-')
        
        for NAME in os.listdir(self._IN_DIR):
            if not NAME.lower().endswith(".log"):
                continue
            FULL_PATH = os.path.join(self._IN_DIR, NAME)
            if not os.path.isfile(FULL_PATH):
                continue
            
            # 检查文件名是否以当天日期开头
            if not DATE_PATTERN.match(NAME):
                continue  # 跳过不以当天日期开头的文件

            CAT = self._CLASSIFY(NAME)  # 保持原分类
            if not CAT:
                continue
            
            # 获取该分类的分组策略
            CAT_RULE = RULES.get(CAT, {})
            GROUP_BY_SITE = CAT_RULE.get("group_by_site", True)
            
            if GROUP_BY_SITE:
                # 按站点分组
                SITE = self._EXTRACT_SITE(NAME)
                SITE_MAP = GROUPED_PATHS.setdefault(SITE, {})
                SITE_MAP.setdefault(CAT, []).append(FULL_PATH)
            else:
                # 不按站点分组，使用指定的工作表名称
                SHEET_NAME = CAT_RULE.get("sheet_name", CAT.upper())
                SHEET_MAP = GROUPED_PATHS.setdefault(SHEET_NAME, {})
                SHEET_MAP.setdefault(CAT, []).append(FULL_PATH)

        if not GROUPED_PATHS:
            self.add_result(Level.OK, f"未匹配到目标设备日志，未生成 Excel（目录：{self._IN_DIR}）")
            return []

        # 站点内保证稳定顺序：按分类规则定义的顺序，再按文件名排序
        for SITE, CATS in GROUPED_PATHS.items():
            for CATEGORY in RULES.keys():
                if CATEGORY in CATS:
                    CATS[CATEGORY].sort(key=lambda PATH: os.path.basename(PATH))

        # 初始化工作簿（不立即保存）
        WORKBOOK = Workbook()
        try:
            WORKBOOK.remove(WORKBOOK.active)
        except Exception:
            pass

        self._WB = WORKBOOK
        self._GROUPED_PATHS = GROUPED_PATHS
        self._SITES_ORDER = sorted(GROUPED_PATHS.keys())  # 进度总数 N
        self._TOTAL_DEVICES = 0

        # 返回每个站点一个 item，实现进度 1/N
        return self._SITES_ORDER

    # 判断文件名是否为目标设备：检查文件名是否匹配CS+N9K或LINK+AS模式
    @staticmethod
    def _is_target(OxidizedBackup_FILENAME: str) -> bool:
        import re
        FILENAME_LOWER = OxidizedBackup_FILENAME.lower()
        # 类别1：CS + N9K + (01|02|03|04)，兼容连写
        CONTAINS_CS = ("cs" in FILENAME_LOWER) or re.search(r"\bcs\b", FILENAME_LOWER)
        CONTAINS_N9K = ("n9k" in FILENAME_LOWER) or re.search(r"\bn9k\b", FILENAME_LOWER)
        CONTAINS_DEVICE_NUMBER = re.search(r"(?:^|[^0-9])0?[1-4](?:[^0-9]|$)", FILENAME_LOWER)
        CS_JOIN = re.search(r"cs0?[1-4]", FILENAME_LOWER)
        CLASS1 = (CONTAINS_CS and CONTAINS_N9K and CONTAINS_DEVICE_NUMBER) or (CS_JOIN and CONTAINS_N9K)

        # 类别2：LINK + AS + (01|02)，兼容连写
        CONTAINS_LINK = ("link" in FILENAME_LOWER) or re.search(r"\blink\b", FILENAME_LOWER)
        CONTAINS_AS = ("as" in FILENAME_LOWER) or re.search(r"\bas\b", FILENAME_LOWER)
        CONTAINS_01_02 = re.search(r"(?:^|[^0-9])0?[12](?:[^0-9]|$)", FILENAME_LOWER)
        LINKAS_JOIN = re.search(r"link[-_]*as0?[12]", FILENAME_LOWER)
        AS_JOIN = re.search(r"as0?[12]", FILENAME_LOWER)
        CLASS2 = LINKAS_JOIN or (CONTAINS_LINK and CONTAINS_AS and CONTAINS_01_02) or (CONTAINS_LINK and AS_JOIN)
        return bool(CLASS1 or CLASS2)

    # 设备分类规则配置：定义各种设备类型的匹配规则
    @staticmethod
    # 返回设备分类规则字典，包含分组策略
    def _get_device_classification_rules():
        PATTERNS = get_cat_classification_patterns()
        return {
            "cat1": {
                "name": "N9K核心交换机",
                "group_by_site": True,
                "patterns": PATTERNS["cat1"],
            },
            "cat2": {
                "name": "LINKAS接入交换机",
                "group_by_site": True,
                "patterns": PATTERNS["cat2"],
            },
            "cat3": {
                "name": "ASA防火墙",
                "group_by_site": True,
                "patterns": PATTERNS["cat3"],
            },
            "cat4": {
                "name": "LINK-DS交换机",
                "group_by_site": False,
                "sheet_name": "LINKDS",
                "patterns": PATTERNS["cat4"],
            },
            "cat5": {
                "name": "BGP设备",
                "group_by_site": False,
                "sheet_name": "BGP",
                "patterns": PATTERNS["cat5"],
            },
            "cat6": {
                "name": "OOB-DS交换机",
                "group_by_site": True,
                "patterns": PATTERNS["cat6"],
            },
        }

    # 对目标文件进行分类：根据文件名模式将文件分为不同类别
    @staticmethod
    def _CLASSIFY(FILENAME: str) -> Optional[str]:
        RULES = get_cat_classification_patterns()
        return classify_text_as_cat(FILENAME, RULES)

    # 从文件名提取站点名：解析文件名获取站点标识
    @staticmethod
    def _EXTRACT_SITE(FILENAME: str) -> str:
        return extract_site_from_filename(FILENAME)

    # 生成安全的Excel工作表名称：确保工作表名称符合Excel规范
    @staticmethod
    def _SAFE_SHEET_NAME(NAME: str) -> str:
        return safe_sheet_name(NAME)

    # 根据设备类别截取配置内容：从指定起始行开始截取配置
    # 根据设备类别截取配置内容
    def _EXTRACT_CONFIG(self, LINES: list[str], CAT: str, FNAME: str) -> list[str]:
        if CAT == "cat1":  # NXOS N9K
            START_MARKER = "! show running-config"
            for LINE_INDEX, LINE in enumerate(LINES):
                if LINE.strip().startswith(START_MARKER):
                    return LINES[LINE_INDEX:]  # 从该行开始截取到文件末尾
        elif CAT == "cat2":  # LINKAS接入交换机（IOS-XE）
            # LINKAS接入交换机使用IOS-XE配置格式
            IOSXE_MARKER = "! Last configuration change"
            for LINE_INDEX, LINE in enumerate(LINES):
                if LINE.strip().startswith(IOSXE_MARKER):
                    return LINES[LINE_INDEX:]  # IOS-XE 配置
        elif CAT == "cat3":  # ASA防火墙
            # ASA防火墙配置提取
            ASA_MARKER = "ASA Version"
            for LINE_INDEX, LINE in enumerate(LINES):
                if LINE.strip().startswith(ASA_MARKER):
                    return LINES[LINE_INDEX:]  # ASA 配置
        elif CAT == "cat4":  # LINK-DS交换机
            # LINK-DS交换机配置提取（支持N9K和IOS-XE）
            N9K_MARKER = "! show running-config"
            IOSXE_MARKER = "! Last configuration change"
            for LINE_INDEX, LINE in enumerate(LINES):
                if LINE.strip().startswith(N9K_MARKER):
                    return LINES[LINE_INDEX:]  # N9K 配置
                elif LINE.strip().startswith(IOSXE_MARKER):
                    return LINES[LINE_INDEX:]  # IOS-XE 配置
        elif CAT == "cat5":  # BGP设备
            # BGP设备配置提取（支持N9K和IOS-XE）
            N9K_MARKER = "! show running-config"
            IOSXE_MARKER = "! Last configuration change"
            for LINE_INDEX, LINE in enumerate(LINES):
                if LINE.strip().startswith(N9K_MARKER):
                    return LINES[LINE_INDEX:]  # N9K 配置
                elif LINE.strip().startswith(IOSXE_MARKER):
                    return LINES[LINE_INDEX:]  # IOS-XE 配置
        elif CAT == "cat6":  # OOB-DS交换机（IOS-XE）
            # OOB-DS交换机使用IOS-XE配置格式
            IOSXE_MARKER = "! Last configuration change"
            for LINE_INDEX, LINE in enumerate(LINES):
                if LINE.strip().startswith(IOSXE_MARKER):
                    return LINES[LINE_INDEX:]  # IOS-XE 配置
        
        # 如果没找到标记，返回原始内容
        return LINES

    # 处理单个站点/工作表的配置备份：读取配置文件并生成Excel工作表
    # 参数SITE可能是站点名（对于cat1/cat2/cat3）或工作表名（对于cat4/cat5）
    def run_single(self, SITE: str):
        from openpyxl.styles import Alignment
        from openpyxl.utils import get_column_letter

        if not self._WB or SITE not in self._GROUPED_PATHS:
            self.add_result(Level.WARN, f"站点/工作表 {SITE} 跳过（未初始化或未找到文件）")
            return

        WORKSHEET = self._WB.create_sheet(title=self._SAFE_SHEET_NAME(SITE))
        WRAP = Alignment(wrap_text=True, vertical="top")

        COLUMN_INDEX = 1
        SITE_CATEGORY_MAP = self._GROUPED_PATHS[SITE]

        # 列顺序：按分类规则定义的顺序显示
        RULES = self._get_device_classification_rules()
        for CATEGORY in RULES.keys():
            for FULL_PATH in SITE_CATEGORY_MAP.get(CATEGORY, []):
                FILE_NAME = os.path.basename(FULL_PATH)
                try:
                    with open(FULL_PATH, "r", encoding="utf-8", errors="ignore") as FILE_HANDLE:
                        LINES = FILE_HANDLE.read().splitlines()
                    
                    # 根据设备类别截取配置内容
                    CONFIG_LINES = self._EXTRACT_CONFIG(LINES, CATEGORY, FILE_NAME)
                    
                except Exception as EXCEPTION:
                    self.add_result(Level.ERROR, f"读取失败 {FILE_NAME}: {EXCEPTION}")
                    # 仍占一列留痕
                    WORKSHEET.cell(row=1, column=COLUMN_INDEX, value=f"{FILE_NAME}（读取失败：{EXCEPTION}）").alignment = WRAP
                    WORKSHEET.column_dimensions[get_column_letter(COLUMN_INDEX)].width = 80
                    COLUMN_INDEX += 1
                    continue

                WORKSHEET.cell(row=1, column=COLUMN_INDEX, value=FILE_NAME).alignment = WRAP
                ROW_INDEX = 2
                for LINE in CONFIG_LINES:
                    WORKSHEET.cell(row=ROW_INDEX, column=COLUMN_INDEX, value=LINE).alignment = WRAP
                    WORKSHEET.row_dimensions[ROW_INDEX].height = 15
                    ROW_INDEX += 1
                WORKSHEET.column_dimensions[get_column_letter(COLUMN_INDEX)].width = 80
                COLUMN_INDEX += 1

        # 统计：该站点的设备列数
        SITE_DEVICE_COUNT = COLUMN_INDEX - 1
        self._TOTAL_DEVICES += SITE_DEVICE_COUNT

        # 输出每个站点/工作表的结果
        self.add_result(Level.OK, f"站点/工作表 {SITE} 处理完成，设备数={SITE_DEVICE_COUNT}")

        # 检查是否是最后一个站点/工作表（通过其在列表中的位置判断）
        if SITE not in self._SITES_ORDER:
            # 如果SITE不在列表中，说明有逻辑错误，但不影响保存
            return
        CURRENT_INDEX = self._SITES_ORDER.index(SITE)
        if CURRENT_INDEX == len(self._SITES_ORDER) - 1:
            # 这是最后一个站点，保存Excel文件
            try:
                self._WB.save(self._XLSX_PATH)
            except Exception as EXCEPTION:
                self.add_result(Level.ERROR, f"Excel 保存失败: {self._XLSX_PATH} -> {EXCEPTION}")
                return

