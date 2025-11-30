# ES N9K 异常日志巡检任务（Kibana 上 Cisco N9K 日志异常扫描）
#
# 技术栈:Python, requests, JSON, 正则表达式
# 目标:在一组 ESN9KLOGInspectTask.kibana_bases 中挑选可用的 Kibana（调用 /api/status），然后经 Kibana 反代查询 ES
#
# 索引匹配:index_pattern（从配置文件读取，如 "*-n9k-*-*"）
# 时间范围:time_gte 到 time_lt（从配置文件读取，如 now-3d 到 now）
# 取数方式:_search?scroll 滚动查询，字段 @timestamp 与 message
#
# 告警提取:从 message 中解析 Cisco 标准样式的严重级别（形如 %...-<sev>-... 里的 <sev> 数字）；
# 映射:sev<=2 → CRITICAL，sev==3 → ERROR，sev==4 → WARN，sev>=5 或无 sev → OK
#
# 忽略列表:按 ESN9KLOGInspectTask.ignore_alarm_file 配置文件的 esn9k.message_contains 与 esn9k.message_regex 过滤噪声
#
# 输出:统计 scanned/matched，给出最严重等级与最多 10 条样例（时间戳、sev、sev 文本、等级、截断消息）

# 导入标准库
import json
import os
import re
from typing import Optional, Tuple

# 导入第三方库
import requests

# 导入本地应用
from .TaskBase import BaseTask, Level, CONFIG, require_keys

# ES N9K 异常日志巡检（通过 Kibana Console 代理）
SEVERITY_REGULAR_EXPRESSION = re.compile(r"%[^-\s]+-(\d)-")
SEVERITY_TEXT = {
    0: "EMERG", 1: "ALERT", 2: "CRIT", 3: "ERR",
    4: "WARN", 5: "NOTICE", 6: "INFO", 7: "DEBUG"
}
LEVEL_ORDER_ELASTICSEARCH = ["CRITICAL", "ERROR", "WARN", "OK"]

# 提取N9K日志消息中的最低严重级别：从日志消息中解析所有严重级别并返回最低值
def _esn9k_minimum_severity(MSG: str) -> Optional[int]:
    """提取N9K日志消息中的最低严重级别


    Args:
        MSG: 日志消息


    Returns:
        Optional[int]: 最低严重级别，如果未找到则返回None
    """
    if not MSG:
        return None
    VALS = []
    for MATCH in SEVERITY_REGULAR_EXPRESSION.finditer(MSG):
        try:
            VALS.append(int(MATCH.group(1)))
        except Exception:
            pass
    return min(VALS) if VALS else None

# ES N9K 异常日志巡检（实现匹配忽略告警）
# 将N9K严重级别转换为系统级别：根据严重级别数值映射到系统告警级别
def _esn9k_sev_to_level(MIN_SEV: Optional[int]) -> str:
    """将N9K严重级别转换为系统级别


    Args:
        MIN_SEV: 最低严重级别数值


    Returns:
        str: 系统告警级别（CRITICAL/ERROR/WARN/OK）
    """
    if MIN_SEV is None or MIN_SEV >= 5:
        return "OK"
    if MIN_SEV <= 2:
        return "CRITICAL"
    if MIN_SEV == 3:
        return "ERROR"
    if MIN_SEV == 4:
        return "WARN"
    return "OK"

# 比较两个告警级别的严重程度：返回更严重的告警级别
def _esn9k_worse(A: str, B: str) -> str:
    """比较两个告警级别的严重程度


    Args:
        A: 告警级别1
        B: 告警级别2


    Returns:
        str: 更严重的告警级别
    """
    ORDER = {LEVEL: INDEX for INDEX, LEVEL in enumerate(LEVEL_ORDER_ELASTICSEARCH)}
    return A if ORDER.get(A, 99) < ORDER.get(B, 99) else B

# 加载忽略告警配置：从YAML文件读取需要忽略的告警规则
def _esn9k_load_ignores() -> dict:
    """加载忽略告警配置


    从YAML文件读取需要忽略的告警规则


    Returns:
        dict: 包含contains和regex规则的字典
    """
    try:
        # 从配置文件读取ignore_alarm_file（必须配置）
        require_keys(CONFIG, ["ESN9KLOGInspectTask"], "root")
        require_keys(CONFIG["ESN9KLOGInspectTask"], ["ignore_alarm_file"], "ESN9KLOGInspectTask")
        FILE_PATH = CONFIG["ESN9KLOGInspectTask"]["ignore_alarm_file"]
        if not os.path.exists(FILE_PATH):
            return {"contains": [], "regex": []}
        import yaml as _yaml, re as _re
        with open(FILE_PATH, "r", encoding="utf-8") as FILE_HANDLE:
            YAML_DATA = _yaml.safe_load(FILE_HANDLE) or {}
        YAML_NODE = (YAML_DATA.get("esn9k") or {}) if isinstance(YAML_DATA, dict) else {}
        CONTAINS = YAML_NODE.get("message_contains") or []
        REGEXPS = YAML_NODE.get("message_regex") or []
        COMPILED = []
        for PATTERN in REGEXPS:
            try:
                COMPILED.append(_re.compile(PATTERN, _re.IGNORECASE))
            except Exception:
                pass

        # 标准化字符串：去除多余空格并转换为小写
        def _norm(S: str) -> str:
            return " ".join(str(S).split()).lower()

        return {"contains": [_norm(ITEM) for ITEM in CONTAINS], "regex": COMPILED}
    except Exception:
        return {"contains": [], "regex": []}

_ESN9K_IGNORES = _esn9k_load_ignores()

# 检查告警消息是否应该被忽略：根据忽略规则判断告警是否应该被过滤
def _esn9k_should_ignore(MSG: str) -> bool:
    """检查告警消息是否应该被忽略


    Args:
        MSG: 告警消息


    Returns:
        bool: 如果应该被忽略则返回True
    """
    if not MSG:
        return False
    try:
        NORM = " ".join(MSG.split()).lower()
        for SUBSTRING in _ESN9K_IGNORES.get("contains", []) or []:
            if SUBSTRING and SUBSTRING in NORM:
                return True
        for REGEX in _ESN9K_IGNORES.get("regex", []) or []:
            if REGEX.search(MSG):
                return True
    except Exception:
        return False
    return False

# 选择可用的Kibana实例：从配置的Kibana实例中选择一个可用的
def _esn9k_pick_kibana(SESSION: requests.Session) -> tuple[str, str]:
    """选择可用的Kibana实例


    Args:
        SESSION: requests会话对象


    Returns:
        tuple[str, str]: (Kibana名称, Kibana基础URL)元组


    Raises:
        RuntimeError: 如果没有可用的Kibana实例
    """
    # 从配置文件读取ESN9KLOGInspectTask的配置（必须配置）
    require_keys(CONFIG, ["ESN9KLOGInspectTask"], "root")
    require_keys(CONFIG["ESN9KLOGInspectTask"], ["kibana_bases"], "ESN9KLOGInspectTask")
    TASK_CONFIGURATION = CONFIG["ESN9KLOGInspectTask"]
    BASES: dict = TASK_CONFIGURATION["kibana_bases"]
    LAST_ERROR = None
    for NAME, BASE in BASES.items():
        try:
            UNIFORM_RESOURCE_LOCATOR = f"{BASE}/api/status"
            HTTP_RESPONSE = SESSION.get(
                UNIFORM_RESOURCE_LOCATOR, headers={"kbn-xsrf": "true"}, timeout=10
            )
            HTTP_RESPONSE.raise_for_status()
            return NAME, BASE
        except Exception as ERROR:
            LAST_ERROR = ERROR
    raise RuntimeError(f"没有可用的 Kibana（esn9k.kibana_bases）。最后错误: {LAST_ERROR}")

# 获取Kibana版本信息：通过API获取Kibana实例的版本号
def _esn9k_kbn_version(SESSION: requests.Session, BASE: str) -> str:
    """获取Kibana版本信息


    Args:
        SESSION: requests会话对象
        BASE: Kibana基础URL


    Returns:
        str: Kibana版本号
    """
    URL = f"{BASE}/api/status"
    HEADERS = {"kbn-xsrf": "true"}
    RESPONSE = SESSION.get(URL, headers=HEADERS, timeout=30)
    RESPONSE.raise_for_status()
    try:
        DATA = RESPONSE.json()
        VERSION_NUMBER = DATA.get("version", {}).get("number")
        if VERSION_NUMBER:
            return VERSION_NUMBER
    except Exception:
        pass
    return (RESPONSE.headers.get("kbn-version") or
            RESPONSE.headers.get("x-kibana-version") or "7.17.0")

# 通过Kibana代理执行ES查询：使用Kibana Console代理功能执行Elasticsearch查询
def _esn9k_kbn_proxy(
    SESSION: requests.Session, BASE: str, KIBANA_VERSION: str,
    METHOD: str, PATH: str, BODY: Optional[dict]
):
    """通过Kibana代理执行ES查询


    Args:
        SESSION: requests会话对象
        BASE: Kibana基础URL
        KIBANA_VERSION: Kibana版本
        METHOD: HTTP方法
        PATH: ES查询路径
        BODY: 请求体


    Returns:
        dict: ES查询结果
    """
    from urllib.parse import quote as _quote
    QUERY_PATH = _quote(PATH, safe="")
    UNIFORM_RESOURCE_LOCATOR = (
        f"{BASE}/api/console/proxy?method={METHOD}&path={QUERY_PATH}"
    )
    HEADERS = {
        "kbn-xsrf": "true",
        "kbn-version": KIBANA_VERSION,
        "Content-Type": "application/json"
    }
    JSON_DATA = json.dumps(BODY) if BODY is not None else None
    HTTP_RESPONSE = SESSION.post(
        UNIFORM_RESOURCE_LOCATOR, headers=HEADERS, data=JSON_DATA, timeout=120
    )
    HTTP_RESPONSE.raise_for_status()
    return HTTP_RESPONSE.json()

# 执行ES N9K日志巡检探测：查询Elasticsearch获取N9K设备日志并分析告警
def run_esn9k_probe(TARGET: Optional[Tuple[str, str]] = None) -> dict:
    """执行ES N9K日志巡检探测


    查询Elasticsearch获取N9K设备日志并分析告警级别


    Args:
        TARGET: 可选的(Kibana名称, Kibana基础URL)元组


    Returns:
        dict: 巡检结果，包含扫描数量、匹配数量、最严重级别和样例
    """
    global _ESN9K_IGNORES
    _ESN9K_IGNORES = _esn9k_load_ignores()
    require_keys(CONFIG, ["ESN9KLOGInspectTask"], "root")
    require_keys(
        CONFIG["ESN9KLOGInspectTask"],
        ["index_pattern", "time_gte", "time_lt"],
        "ESN9KLOGInspectTask"
    )
    TASK_CONFIGURATION = CONFIG["ESN9KLOGInspectTask"]
    INDEX_PATTERN: str = TASK_CONFIGURATION["index_pattern"]
    TIME_FIELD: str = TASK_CONFIGURATION.get("time_field", "@timestamp")
    TIME_GTE: str = TASK_CONFIGURATION["time_gte"]
    TIME_LT: str = TASK_CONFIGURATION["time_lt"]
    PAGE_SIZE: int = int(TASK_CONFIGURATION.get("page_size", 1000))
    SCROLL_KEEPALIVE: str = TASK_CONFIGURATION.get("scroll_keepalive", "2m")

    SESSION = requests.Session()
    # 如需鉴权，可从环境变量读取 KIBANA_USER/PASS（保密）
    if TARGET is not None:
        KIBANA_NAME, BASE = TARGET
    else:
        KIBANA_NAME, BASE = _esn9k_pick_kibana(SESSION)
    KIBANA_VERSION = _esn9k_kbn_version(SESSION, BASE)

    BASE_QUERY = {
        "size": PAGE_SIZE,
        "_source": [TIME_FIELD, "message"],
        "sort": [{TIME_FIELD: "asc"}],
        "track_total_hits": True,
        "query": {"range": {TIME_FIELD: {"gte": TIME_GTE, "lt": TIME_LT}}}
    }

    FIRST_PATH = f"/{INDEX_PATTERN}/_search?scroll={SCROLL_KEEPALIVE}"
    FIRST = _esn9k_kbn_proxy(SESSION, BASE, KIBANA_VERSION, "POST", FIRST_PATH, BASE_QUERY)
    SCROLL_IDENTIFIER = FIRST.get("_scroll_id")

    if not SCROLL_IDENTIFIER:
        return {"kibana": {"name": KIBANA_NAME, "base": BASE, "version": KIBANA_VERSION},
                "scanned": 0, "matched": 0, "worst_level": "OK", "samples": [],
                "note": "未获得 _scroll_id，可能索引无数据或权限不足"}

    SCANNED = MATCHED = 0
    WORST = "OK"
    SAMPLES = []
    SAMPLE_LIMIT = 10

    try:
        HTTP_RESPONSE = FIRST
        while True:
            HITS = HTTP_RESPONSE.get("hits", {}).get("hits", [])
            if not HITS:
                break
            for HIT in HITS:
                SCANNED += 1
                SOURCE_DATA = HIT.get("_source", {}) or {}
                TIMESTAMP = SOURCE_DATA.get(TIME_FIELD, "")
                MESSAGE = SOURCE_DATA.get("message", "") or ""
                if _esn9k_should_ignore(MESSAGE):
                    continue
                SEVERITY = _esn9k_minimum_severity(MESSAGE)
                if SEVERITY is not None and SEVERITY <= 4:
                    MATCHED += 1
                    LEVEL = _esn9k_sev_to_level(SEVERITY)
                    WORST = _esn9k_worse(WORST, LEVEL)
                    if len(SAMPLES) < SAMPLE_LIMIT:
                        SAMPLES.append({"timestamp": TIMESTAMP, "severity": SEVERITY,
                                        "severity_text": SEVERITY_TEXT.get(SEVERITY, "?"),
                                        "level": LEVEL, "message": MESSAGE[:800]})
            HTTP_RESPONSE = _esn9k_kbn_proxy(
                SESSION, BASE, KIBANA_VERSION, "POST", "/_search/scroll",
                {"scroll": SCROLL_KEEPALIVE, "scroll_id": SCROLL_IDENTIFIER}
            )
            SCROLL_IDENTIFIER = HTTP_RESPONSE.get("_scroll_id")
            if not SCROLL_IDENTIFIER:
                break
    finally:
        try:
            _esn9k_kbn_proxy(
                SESSION, BASE, KIBANA_VERSION, "DELETE", "/_search/scroll",
                {"scroll_id": [SCROLL_IDENTIFIER]}
            )
        except Exception:
            pass

    return {"kibana": {"name": KIBANA_NAME, "base": BASE, "version": KIBANA_VERSION},
            "scanned": SCANNED, "matched": MATCHED, "worst_level": WORST, "samples": SAMPLES}

# ESLogN9KInspectTask
# ES N9K日志巡检任务类：通过Elasticsearch查询N9K设备日志并分析告警级别
class ESN9KLOGInspectTask(BaseTask):
    """ES N9K日志巡检任务


    通过Elasticsearch查询N9K设备日志并分析告警级别
    """


    # 初始化ES N9K日志巡检任务：设置任务名称
    def __init__(self):
        super().__init__("ES服务器CS-N9K异常日志巡检")

    # 返回要巡检的Kibana实例列表
    def items(self):
        """返回要巡检的Kibana实例列表


        Returns:
            list: (Kibana名称, Kibana基础URL)元组列表
        """
        require_keys(CONFIG, ["ESN9KLOGInspectTask"], "root")
        require_keys(CONFIG["ESN9KLOGInspectTask"], ["kibana_bases"], "ESN9KLOGInspectTask")
        BASES = CONFIG["ESN9KLOGInspectTask"]["kibana_bases"]
        return list(BASES.items())

    # 执行单个Kibana实例的N9K日志巡检：查询日志并分析告警级别
    def run_single(self, ITEM):
        """执行单个Kibana实例的N9K日志巡检


        查询日志并分析告警级别


        Args:
            ITEM: (Kibana名称, Kibana基础URL)元组
        """
        try:
            NAME, BASE = ITEM
            # 从设备名中提取站点名（如HX00-ES -> HX00）
            SITE_NAME = NAME.split('-')[0] if '-' in NAME else NAME
            RESULT = run_esn9k_probe((NAME, BASE))
            WORST = RESULT.get("worst_level", "OK")
            LEVEL_MAP = {
                "CRITICAL": Level.CRIT, "ERROR": Level.ERROR,
                "WARN": Level.WARN, "OK": Level.OK
            }
            LEVEL = LEVEL_MAP.get(WORST, Level.OK)
            SCANNED = RESULT.get("scanned", 0)
            MATCHED = RESULT.get("matched", 0)
            MSG = (
                f"站点{SITE_NAME}扫描日志系统ES关于N9K全等级日志数量={SCANNED} "
                f"命中WARN级别以上数量={MATCHED} 巡检状态={WORST}"
            )
            # 将 meta 改为直接传 samples 列表，避免输出中出现 {"...": ...} 的花括号格式
            self.add_result(LEVEL, MSG, RESULT.get("samples", []))
        except Exception as ERROR:
            self.add_result(Level.ERROR, f"ESN9K 巡检失败: {ERROR}")
