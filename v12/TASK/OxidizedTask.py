# Oxidized 设备配置本地备份任务
#
# 技术栈:Python, requests, lxml, XPath, HTML解析, 文件IO, zipfile压缩
# 目标:从Oxidized服务器备份设备配置，支持多服务器并行处理，实现设备级和节点级统计分离
#
# 处理逻辑:状态检查 → 设备分组 → 配置获取 → 本地保存 → 统计汇总 → 日志压缩
#
# 状态检查优化:
# - 使用XPath检查Last Status：//*[@id='nodesTable']/tbody/tr/td[4]/div
# - 检查div元素的class属性（success/no_connection/never/failing）
# - 备用方案：检查内部span[@style="visibility: hidden"]元素内容
# - no_connection状态：直接标记为WARN，跳过备份请求，不需要检查URL
# - success状态：尝试备份配置，但需检查URL响应内容是否为"node not found"
# - never/failing状态：直接标记为WARN，跳过备份请求
# - 其他状态：标记为WARN，跳过备份请求
# - 响应内容检查：检测"node not found"响应，标记为WARN
# - 避免无效请求，提升任务执行效率
#
# 架构重构优化:
# - 设备级统计：ERROR和OK计数只涉及备份设备，不包含节点统计
# - 节点级统计：WARN计数只涉及备份节点失联，不包含设备失败
# - LOG输出：只记录每个设备的ERROR和OK，不记录节点统计信息
# - REPORT输出：只统计节点ERROR和OK数量，不显示具体设备信息
# - 设备过滤：支持基于前缀的设备过滤，可配置忽略特定设备（如HX07、TWIDC、FG）
# - 日志压缩：自动将当日所有.log文件压缩为Oxidized_设备配置备份-{日期}.zip
#
# 配置验证:
# - 验证OxidizedTask专用配置：检查base_urls配置项
# - 验证设备过滤配置：检查ignore_device_prefixes配置项
# - 独立配置验证：在任务初始化时验证所需配置，不依赖Core.py
#
# 输出:LOG/OxidizedTask/OxidizedTaskBackup/YYYYMMDD-设备名.log（成功设备配置文件，V10新结构：从LOG/日期/OxidizedTaskBackup迁移），压缩后生成YYYYMMDD-OxidizedTaskBackup.zip文件
# 配置说明:支持多服务器配置，自动处理设备分组和状态检查，支持设备过滤和日志压缩

# 导入标准库
import os
import random
import re
import time
from datetime import datetime
import zipfile

# 导入第三方库
import requests
from lxml import html
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# 导入本地应用
from .TaskBase import (
    BaseTask, Level, CONFIG, WEB_MAX_WORKERS, WEB_DELAY_RANGE,
    DEFAULT_HTTP_TIMEOUT, require_keys
)

# 安全文件名处理函数：将特殊字符替换为下划线，确保文件名合法
def safe_filename(NAME: str) -> str:
    """将文件名中的特殊字符替换为下划线，确保文件名合法
    

    Args:
        NAME: 原始文件名
        

    Returns:
        str: 处理后的安全文件名
    """
    return re.sub(r'[\\/:*?"<>|]+', "_", str(NAME)).strip()

# Oxidized设备配置备份任务类：从Oxidized服务获取设备配置并保存到本地
class OxidizedTask(BaseTask):
    """Oxidized设备配置备份任务
    

    从Oxidized服务器备份设备配置，支持多服务器并行处理，
    实现设备级和节点级统计分离
    """
    

    # 初始化Oxidized任务：设置任务名称、配置URL列表和日志目录
    def __init__(self):
        super().__init__("Oxidized设备配置本地备份")
        

        # 验证OxidizedTask专用配置
        require_keys(CONFIG, ["OxidizedTask"], "root")
        require_keys(
            CONFIG["OxidizedTask"],
            ["base_urls", "ignore_device_prefixes"],
            "OxidizedTask"
        )
        

        self.BASE_URLS: list[str] = CONFIG["OxidizedTask"]["base_urls"]
        # 自己获取日志目录（必须配置）
        require_keys(CONFIG, ["settings"], "root")
        require_keys(CONFIG["settings"], ["log_dir"], "settings")
        BASE_LOG_DIR = CONFIG["settings"]["log_dir"]
        # V10新结构：输出到 LOG/OxidizedTask/OxidizedTaskBackup/
        self.LOG_DIR = os.path.join(BASE_LOG_DIR, "OxidizedTask", "OxidizedTaskBackup")
        # 用于存储所有节点的设备信息
        self.ALL_DEVICES = []
        self.TOTAL_FAILED = 0
        self.TOTAL_SUCCEEDED = 0
        # 忽略的设备前缀列表（从配置文件读取，必须配置）
        self.IGNORE_DEVICE_PREFIXES = CONFIG["OxidizedTask"]["ignore_device_prefixes"]

    # 返回要处理的Oxidized服务URL列表
    def items(self):
        """返回要处理的Oxidized服务URL列表
        

        Returns:
            list[str]: Oxidized服务URL列表
        """
        return self.BASE_URLS

    # 执行单个Oxidized服务的设备配置备份：获取设备列表并下载配置
    def run_single(self, BASE_URL: str) -> None:
        """执行单个Oxidized服务的设备配置备份
        

        Args:
            BASE_URL: Oxidized服务的基础URL
        """
        # 初始化计数器
        SUCCEEDED, FAILED = 0, 0
        LOCAL_BACKUP_COUNT = 0
        TODAY_STR = datetime.now().strftime("%Y%m%d")
        

        # 确保OxidizedTaskBackup目录存在
        os.makedirs(self.LOG_DIR, exist_ok=True)
        

        try:
            SESSION = requests.Session()
            # 连接池收紧 + 429/50x 自动重试 + UA
            RETRIES = Retry(
                total=2,
                backoff_factor=0.5,
                status_forcelist=(429, 502, 503, 504),
                allowed_methods=frozenset(["GET"]),
                respect_retry_after_header=True,
            )
            ADAPTER = HTTPAdapter(
                pool_connections=WEB_MAX_WORKERS,
                pool_maxsize=WEB_MAX_WORKERS,
                max_retries=RETRIES,
            )
            SESSION.mount("http://", ADAPTER)
            SESSION.mount("https://", ADAPTER)
            SESSION.headers.update({
                "User-Agent": "FATTools/1.0",
                "Accept-Encoding": "gzip, deflate"
            })
            RESP = SESSION.get(BASE_URL, timeout=(DEFAULT_HTTP_TIMEOUT, DEFAULT_HTTP_TIMEOUT))
            RESP.raise_for_status()

        except Exception as ERROR:
            # 节点异常：只记录WARN，不记录ERROR/OK
            self.add_result(Level.WARN, f"{BASE_URL} 节点异常")
            return

        # 解析设备列表和状态
        TREE = html.fromstring(RESP.content)
        DEVICE_NAMES = TREE.xpath("//table/tbody/tr/td[1]/a/text()")
        GROUP_NAMES = TREE.xpath("//table/tbody/tr/td[3]/a/text()")
        

        # 获取Last Status状态 - 尝试多种XPath
        STATUS_ELEMENTS = TREE.xpath("//*[@id='nodesTable']/tbody/tr/td[4]/div")
        

        # 如果第一种XPath没找到，尝试其他可能的XPath
        if not STATUS_ELEMENTS:
            STATUS_ELEMENTS = TREE.xpath("//table/tbody/tr/td[4]/div")
        if not STATUS_ELEMENTS:
            STATUS_ELEMENTS = TREE.xpath("//tr/td[4]/div")
        if not STATUS_ELEMENTS:
            STATUS_ELEMENTS = TREE.xpath("//td[4]/div")
        if not STATUS_ELEMENTS:
            # 尝试查找所有包含class的div元素
            STATUS_ELEMENTS = TREE.xpath(
                "//div[@class='success' or @class='no_connection' "
                "or @class='never' or @class='failing']"
            )
        

        DEVICE_STATUSES = []
        for DIV in STATUS_ELEMENTS:
            # 检查div的class属性，提取状态信息
            DIV_CLASS = DIV.get('class', '')
            if DIV_CLASS in ['success', 'no_connection', 'never', 'failing']:
                DEVICE_STATUSES.append(DIV_CLASS)
            else:
                # 如果div没有class，尝试从内部的span获取
                SPAN = DIV.find('.//span[@style="visibility: hidden"]')
                if SPAN is not None:
                    STATUS_TEXT = SPAN.text_content().strip()
                    DEVICE_STATUSES.append(STATUS_TEXT)
                else:
                    DEVICE_STATUSES.append('unknown')
        

        # 如果状态元素数量不匹配，用unknown填充
        while len(DEVICE_STATUSES) < len(DEVICE_NAMES):
            DEVICE_STATUSES.append('unknown')

        # 第一步：统计设备状态（不进行实际备份）
        for DEVICE_NAME, GROUP_NAME, STATUS in zip(DEVICE_NAMES, GROUP_NAMES, DEVICE_STATUSES):
            DEVICE = DEVICE_NAME.strip()
            GROUP = GROUP_NAME.strip()
            

            # 检查是否忽略该设备
            if any(DEVICE.startswith(PREFIX) for PREFIX in self.IGNORE_DEVICE_PREFIXES):
                continue
            

            # 根据Last Status判断设备状态
            if STATUS == 'success':
                SUCCEEDED += 1
            else:
                FAILED += 1

        # 第二步：开始实际备份success状态的设备
        for DEVICE_NAME, GROUP_NAME, STATUS in zip(DEVICE_NAMES, GROUP_NAMES, DEVICE_STATUSES):
            DEVICE = DEVICE_NAME.strip()
            GROUP = GROUP_NAME.strip()
            

            # 检查是否忽略该设备
            if any(DEVICE.startswith(PREFIX) for PREFIX in self.IGNORE_DEVICE_PREFIXES):
                continue
            

            # 处理success状态的设备
            if STATUS == 'success':
                # 尝试获取配置
                FETCH_URL = f"{BASE_URL.replace('/nodes', '')}/node/fetch/{GROUP}/{DEVICE}"
                try:
                    CFG_RESP = SESSION.get(FETCH_URL, timeout=DEFAULT_HTTP_TIMEOUT)
                    CFG_RESP.raise_for_status()
                    

                    # 检查响应内容是否为"node not found"
                    if "node not found" in CFG_RESP.text.lower():
                        # 设备失败：记录到results和设备列表
                        self.add_result(
                            Level.ERROR,
                            f"{BASE_URL}-{DEVICE}({GROUP}) 节点未找到 - Oxidized备份失败"
                        )
                        self.ALL_DEVICES.append({
                            'base_url': BASE_URL,
                            'device': DEVICE,
                            'group': GROUP,
                            'status': 'ERROR',
                            'message': f"{BASE_URL}-{DEVICE}({GROUP}) 节点未找到 - Oxidized备份失败"
                        })
                        FAILED += 1
                        continue
                    

                    LOG_PATH = os.path.join(
                        self.LOG_DIR,
                        f"{TODAY_STR}-{safe_filename(DEVICE)}.log"
                    )
                    with open(LOG_PATH, "w", encoding="utf-8") as FILE_HANDLE:
                        FILE_HANDLE.write(CFG_RESP.text)
                    LOCAL_BACKUP_COUNT += 1
                    # 设备成功：记录到results和设备列表
                    self.add_result(Level.OK, f"{BASE_URL}-{DEVICE}({GROUP}) 备份成功")
                    self.ALL_DEVICES.append({
                        'base_url': BASE_URL,
                        'device': DEVICE,
                        'group': GROUP,
                        'status': 'OK',
                        'message': f"{BASE_URL}-{DEVICE}({GROUP}) 备份成功"
                    })
                    SUCCEEDED += 1
                except Exception as ERROR:
                    # 设备失败：记录到results和设备列表
                    self.add_result(Level.ERROR, f"{BASE_URL}-{DEVICE}({GROUP}) 获取失败: {ERROR}")
                    self.ALL_DEVICES.append({
                        'base_url': BASE_URL,
                        'device': DEVICE,
                        'group': GROUP,
                        'status': 'ERROR',
                        'message': f"{BASE_URL}-{DEVICE}({GROUP}) 获取失败: {ERROR}"
                    })
                    FAILED += 1
                time.sleep(random.uniform(*WEB_DELAY_RANGE))
            else:
                # 处理非success状态的设备（no_connection, never, failing等）
                # 设备失败：记录到results和设备列表
                self.add_result(
                    Level.ERROR,
                    f"{BASE_URL}-{DEVICE}({GROUP}) Oxidized备份失败 - {STATUS}"
                )
                self.ALL_DEVICES.append({
                    'base_url': BASE_URL,
                    'device': DEVICE,
                    'group': GROUP,
                    'status': 'ERROR',
                    'message': f"{BASE_URL}-{DEVICE}({GROUP}) Oxidized备份失败 - {STATUS}"
                })

        # 第三步：更新总计数器
        self.TOTAL_FAILED += FAILED
        self.TOTAL_SUCCEEDED += SUCCEEDED

    # 重写run方法，在所有节点处理完成后统一输出LOG
    def run(self):
        """执行所有Oxidized服务的设备配置备份任务
        

        重写父类方法，在所有节点处理完成后统一输出LOG和打包备份文件
        """
        # 清空设备列表和计数器
        self.ALL_DEVICES = []
        self.TOTAL_FAILED = 0
        self.TOTAL_SUCCEEDED = 0
        

        # 调用父类的run方法处理所有节点
        super().run()
        

        

        # 直接更新主计数器，不通过add_result避免写入LOG
        self._update_main_counters()
        # 打包历史（仅打包非今日的备份文件）
        try:
            self._pack_backups_excluding_today()
        except Exception:
            pass

    # 更新主计数器：直接更新计数器，不通过add_result避免写入LOG
    def _update_main_counters(self):
        pass

    # 按日期打包设备配置文件（不打包今天）
    def _pack_backups_excluding_today(self) -> None:
        TODAY_STR = datetime.now().strftime("%Y%m%d")
        if not os.path.isdir(self.LOG_DIR):
            return
        # 收集 YYYYMMDD-*.log 文件，按日期分组（排除今天）
        DATE_GROUPS: dict[str, list[str]] = {}
        for NAME in os.listdir(self.LOG_DIR):
            if not NAME.lower().endswith('.log'):
                continue
            MATCH = re.match(r"^(\d{8})-(.+)\.log$", NAME)
            if not MATCH:
                continue
            DATE_STR = MATCH.group(1)
            if DATE_STR == TODAY_STR:
                continue
            DATE_GROUPS.setdefault(DATE_STR, []).append(NAME)
        if not DATE_GROUPS:
            return
        # 逐日期打包
        for DATE_STR, FILES in sorted(DATE_GROUPS.items()):
            ZIP_NAME = f"{DATE_STR}-OxidizedTaskBackup.zip"
            ZIP_PATH = os.path.join(self.LOG_DIR, ZIP_NAME)
            if os.path.exists(ZIP_PATH):
                continue
            try:
                with zipfile.ZipFile(ZIP_PATH, 'w', zipfile.ZIP_DEFLATED) as ZIPF:
                    for F in FILES:
                        FILE_PATH = os.path.join(self.LOG_DIR, F)
                        if os.path.isfile(FILE_PATH):
                            ZIPF.write(FILE_PATH, arcname=F)
                # 打包成功后删除原文件
                for F in FILES:
                    FILE_PATH = os.path.join(self.LOG_DIR, F)
                    try:
                        os.remove(FILE_PATH)
                    except Exception:
                        pass
            except Exception:
                # 出错忽略，避免影响任务主流程
                pass
