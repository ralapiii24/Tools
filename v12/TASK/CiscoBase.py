# Cisco ACL 解析基础模块

# 导入标准库
import re
import socket
from dataclasses import dataclass
from ipaddress import IPv4Address, IPv4Network
from typing import Callable, Dict, List, Optional, Set, Tuple

CatPattern = Callable[[str], bool]
CatPatternDict = Dict[str, List[CatPattern]]


def _normalize_identifier(TEXT: Optional[str]) -> str:
    return str(TEXT or "").strip().lower()


def get_cat_classification_patterns() -> CatPatternDict:
    return {
        "cat1": [
            lambda text: ("n9k" in text) and ("cs" in text) and re.search(r"(?:^|[^0-9])0?[1-4](?:[^0-9]|$)", text),
            lambda text: re.search(r"cs0?[1-4]", text) and ("n9k" in text),
        ],
        "cat2": [
            lambda text: ("link" in text) and ("as" in text) and re.search(r"(?:^|[^0-9])0?[12](?:[^0-9]|$)", text),
            lambda text: re.search(r"link[-_]*as0?[12]", text),
            lambda text: ("link" in text) and re.search(r"as0?[12]", text),
        ],
        "cat3": [
            lambda text: "fw01-frp" in text or "fw02-frp" in text,
        ],
        "cat4": [
            lambda text: "link-ds" in text and re.search(r"0?[12]", text),
            lambda text: re.search(r"link[-_]?ds0?[12]", text),
        ],
        "cat5": [
            lambda text: "bgp" in text,
        ],
        "cat6": [
            lambda text: re.search(r"\boob[-_]?ds0?[12]\b", text) or (
                "oob" in text and "ds" in text and re.search(r"(?:^|[^0-9])0?[12](?:[^0-9]|$)", text)
            ),
        ],
    }


def classify_text_as_cat(TEXT: str, RULES: CatPatternDict) -> Optional[str]:
    LOWERED = _normalize_identifier(TEXT)
    for CAT_ID, PATTERNS in RULES.items():
        for PATTERN_FUNC in PATTERNS:
            if PATTERN_FUNC(LOWERED):
                return CAT_ID
    return None


def text_matches_cat(TEXT: str, CAT_ID: str, RULES: CatPatternDict) -> bool:
    return classify_text_as_cat(TEXT, RULES) == CAT_ID


def detect_cat_columns(WORKSHEET, RULES: CatPatternDict, TARGET_CATS: Optional[List[str]] = None) -> Dict[str, List[int]]:
    RESULT: Dict[str, List[int]] = {}
    TARGET_SET = set(TARGET_CATS) if TARGET_CATS else None
    for COLUMN in range(1, WORKSHEET.max_column + 1):
        CELL = WORKSHEET.cell(row=1, column=COLUMN).value
        if not CELL or not isinstance(CELL, str):
            continue
        CAT_ID = classify_text_as_cat(CELL, RULES)
        if not CAT_ID:
            continue
        if TARGET_SET and CAT_ID not in TARGET_SET:
            continue
        RESULT.setdefault(CAT_ID, []).append(COLUMN)
    return RESULT



# 导入第三方库
try:
    from openpyxl.worksheet.worksheet import Worksheet
except ImportError:
    Worksheet = None  # 类型提示用

# 导入本地应用
# (无本地应用依赖)

# 解析辅助函数
# 仅数字或系统services数据库解析；失败返回None（视为任意端口）
def service_to_port(SERVICE_STRING: str) -> Optional[int]:
    # 将服务名或端口号字符串转换为端口号
    if SERVICE_STRING is None:
        return None
    SERVICE_STRING = SERVICE_STRING.strip().lower()
    if not SERVICE_STRING:
        return None
    if SERVICE_STRING.isdigit():
        PORT_NUMBER = int(SERVICE_STRING)
        return PORT_NUMBER if 0 <= PORT_NUMBER <= 65535 else None
    try:
        return socket.getservbyname(SERVICE_STRING)
    except (OSError, socket.gaierror):
        return None

# IOS-XE：ip + wildcard转网络（假设通配位连续）
def ip_and_wildcard_to_network(IP_STRING: str, WILDCARD_STRING: str) -> Optional[IPv4Network]:
    # 将IP地址和通配符转换为IPv4Network对象
    try:
        IP_ADDRESS = int(IPv4Address(IP_STRING))
        WILDCARD_ADDRESS = int(IPv4Address(WILDCARD_STRING))
        NETMASK_INTEGER = (~WILDCARD_ADDRESS) & 0xFFFFFFFF
        PREFIX_LENGTH = 32 - bin(WILDCARD_ADDRESS).count("1")
        NETWORK_INTEGER = IP_ADDRESS & NETMASK_INTEGER
        return IPv4Network((IPv4Address(NETWORK_INTEGER), PREFIX_LENGTH), strict=False)
    except (ValueError, TypeError):
        return None

# 将IP地址转换为/32网络
def host_to_network(IP_STRING: str) -> Optional[IPv4Network]:
    # 将IP地址字符串转换为/32网络
    try:
        return IPv4Network(f"{IP_STRING}/32", strict=False)
    except (ValueError, TypeError):
        return None

# 将CIDR字符串转换为网络对象
def cidr_to_network(CIDR_STRING: str) -> Optional[IPv4Network]:
    # 将CIDR字符串转换为IPv4Network对象
    try:
        return IPv4Network(CIDR_STRING, strict=False)
    except (ValueError, TypeError):
        return None

# 将'any'关键字转换为0.0.0.0/0网络
def any_to_network() -> IPv4Network:
    # 返回表示任意地址的网络对象
    return IPv4Network("0.0.0.0/0", strict=False)

# 从端口字符串中提取所有端口：支持多个端口（用空格分隔），例如"22 22222"或"domain ntp"，返回端口号集合
def _extract_all_ports(port_str: str) -> Set[int]:
    # 从端口字符串中提取所有端口号
    if not port_str:
        return set()
    
    PORTS = set()
    PORT_ITEMS = port_str.strip().split()
    
    for ITEM in PORT_ITEMS:
        ITEM = ITEM.strip()
        if not ITEM:
            continue
        
        # 尝试转换为端口号
        if ITEM.isdigit():
            PORTS.add(int(ITEM))
        else:
            # 尝试将服务名转换为端口号
            try:
                PORT = service_to_port(ITEM)
                if PORT is not None:
                    PORTS.add(PORT)
            except (ValueError, TypeError, OSError):
                pass
    
    return PORTS

# ============================================================================
# 正则表达式定义 - 支持所有Cisco ACL格式
# ============================================================================

# NXOS格式 - 标准格式（目标端口）
# 示例: 10 permit tcp 10.10.0.0/16 10.20.0.0/16 eq 80
# 注意: 使用宽松匹配 (?:\s+\S+)* 以兼容log-input、time-range等关键字（与ACLCrossCheckTask保持一致）
NXOS_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE>\d+\.\d+\.\d+\.\d+/\d+)\s+
    (?P<DESTINATION>\d+\.\d+\.\d+\.\d+/\d+)
    (?:\s+eq\s+(?P<PORT>\S+))?
    (?:\s+\S+)*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# NXOS格式 - 源端口格式
# 示例: 162 permit tcp 10.10.100.31/32 eq 55888 10.10.108.63/32
NXOS_SRC_PORT_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE>\d+\.\d+\.\d+\.\d+/\d+)\s+
    eq\s+(?P<PORT>\S+)\s+
    (?P<DESTINATION>\d+\.\d+\.\d+\.\d+/\d+)
    (?:\s+eq\s+(?P<DST_PORT>\S+))?
    (?:\s+log)?
    \s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# NXOS格式 - 端口范围格式（两种格式）
# 格式1: 源地址 目标地址 range 起始端口 结束端口
# 格式2: 源地址 range 起始端口 结束端口 目标地址
# 示例: 182 permit tcp 10.10.106.40/32 range 8001 8002 10.10.62.32/31
NXOS_RANGE_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE>\d+\.\d+\.\d+\.\d+/\d+)\s+
    (?:
        (?P<DESTINATION_BEFORE>\d+\.\d+\.\d+\.\d+/\d+)\s+range\s+(?P<PORT_START>\d+)\s+(?P<PORT_END>\d+) |
        range\s+(?P<PORT_START2>\d+)\s+(?P<PORT_END2>\d+)\s+(?P<DESTINATION_AFTER>\d+\.\d+\.\d+\.\d+/\d+)
    )
    (?:\s+log)?
    \s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# IOS-XE格式 - Wildcard格式（标准）
# 示例: 10 permit tcp 10.10.0.0 0.0.255.255 10.20.0.0 0.0.255.255 eq 80
IOSXE_WC_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<SOURCE_WILDCARD>\d+\.\d+\.\d+\.\d+)
    (?:\s+eq\s+(?P<PORT_A>\S+))?
    \s+(?P<DESTINATION_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<DESTINATION_WILDCARD>\d+\.\d+\.\d+\.\d+)
    (?:\s+eq\s+(?P<PORT_B>\S+))?
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# IOS-XE格式 - Host格式
# 示例: 10 permit tcp host 10.10.0.1 host 10.20.0.1 eq 80
IOSXE_HOST_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    host\s+(?P<SOURCE_IP>\d+\.\d+\.\d+\.\d+)
    (?:\s+eq\s+(?P<PORT_A>\S+))?
    \s+host\s+(?P<DESTINATION_IP>\d+\.\d+\.\d+\.\d+)
    (?:\s+eq\s+(?P<PORT_B>\S+))?
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# IOS-XE格式 - 混合格式（host和wildcard混合）
# 示例: 10 permit tcp host 10.10.0.1 10.20.0.0 0.0.255.255 eq 80
IOSXE_MIX_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?:
        host\s+(?P<SOURCE_IP_HOST>\d+\.\d+\.\d+\.\d+) |
        (?P<SOURCE_IP_WILDCARD>\d+\.\d+\.\d+\.\d+)\s+(?P<SOURCE_WILDCARD_WILDCARD>\d+\.\d+\.\d+\.\d+)
    )
    (?:\s+eq\s+(?P<PORT_A>[\w\s]+))?
    \s+
    (?:
        host\s+(?P<DESTINATION_IP_HOST>\d+\.\d+\.\d+\.\d+) |
        (?P<DESTINATION_IP_WILDCARD>\d+\.\d+\.\d+\.\d+)\s+(?P<DESTINATION_WILDCARD_WILDCARD>\d+\.\d+\.\d+\.\d+)
    )
    (?:\s+eq\s+(?P<PORT_B>[\w\s]+))?
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# IOS-XE格式 - host <ip> eq <port1> <port2> ... <dst_ip> <wildcard>
# 示例: 4190 permit udp host 10.65.16.53 eq domain ntp 10.70.130.0 0.0.0.255 log
IOSXE_HOST_MULTI_EQ_WILDCARD_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    host\s+(?P<SOURCE_IP>\d+\.\d+\.\d+\.\d+)
    \s+eq\s+(?P<PORT_MULTI>[\w\s]+)
    \s+(?P<DESTINATION_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<DESTINATION_WILDCARD>\d+\.\d+\.\d+\.\d+)
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# IOS-XE格式 - 支持range和多个eq端口（源IP+通配符，目的host）
# 示例: 290 permit tcp 10.70.130.0 0.0.0.31 range 7180 8088 host 10.65.63.55 log
# 示例: 300 permit tcp 10.70.130.0 0.0.0.31 eq 8888 9000 9010 9020 9030 9083 9870 10000 host 10.65.63.55 log
IOSXE_RANGE_MULTI_EQ_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<SOURCE_WILDCARD>\d+\.\d+\.\d+\.\d+)
    \s+(?:range\s+(?P<PORT_RANGE_START>\d+)\s+(?P<PORT_RANGE_END>\d+)|eq\s+(?P<PORT_MULTI>[\w\s]+))
    \s+host\s+(?P<DESTINATION_IP>\d+\.\d+\.\d+\.\d+)
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# IOS-XE格式 - 支持range端口（源host，目的IP+通配符）
# 示例: 3570 permit tcp host 10.65.130.233 range 6446 6447 10.66.231.8 0.0.0.7
IOSXE_RANGE_HOST_SRC_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    host\s+(?P<SOURCE_IP>\d+\.\d+\.\d+\.\d+)
    \s+range\s+(?P<PORT_RANGE_START>\d+)\s+(?P<PORT_RANGE_END>\d+)
    \s+(?P<DESTINATION_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<DESTINATION_WILDCARD>\d+\.\d+\.\d+\.\d+)
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# IOS-XE格式 - 支持range端口（源IP+通配符，目的IP+通配符）
# 示例: 4580 permit tcp 10.62.110.96 0.0.0.31 range 9091 9093 10.66.130.0 0.0.0.255
# 示例: 640 permit tcp 10.10.0.0 0.0.255.255 range 8000 9999 10.104.166.0 0.0.0.255
IOSXE_RANGE_WILDCARD_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<SOURCE_WILDCARD>\d+\.\d+\.\d+\.\d+)
    \s+range\s+(?P<PORT_RANGE_START>\d+)\s+(?P<PORT_RANGE_END>\d+)
    \s+(?P<DESTINATION_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<DESTINATION_WILDCARD>\d+\.\d+\.\d+\.\d+)
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# IOS-XE格式 - 支持多个eq端口（源IP+通配符，目的IP+通配符）
# 示例: 180 permit tcp 10.65.88.192 0.0.0.63 eq www 443 8400 10.62.80.0 0.0.0.7
# 示例: 260 permit tcp 10.70.130.64 0.0.0.3 eq 3366 8030 9030 10.66.130.0 0.0.0.15 log
# 示例: permit tcp 10.10.15.0 0.0.0.255 eq www 443 5480 5900 10.10.62.32 0.0.0.1 log
IOSXE_MULTI_EQ_WILDCARD_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<SOURCE_WILDCARD>\d+\.\d+\.\d+\.\d+)
    \s+eq\s+(?P<PORT_MULTI>[\w\s]+)
    \s+(?P<DESTINATION_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<DESTINATION_WILDCARD>\d+\.\d+\.\d+\.\d+)
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# IOS-XE格式 - Any关键字格式
# 示例: 10 permit tcp any any eq 80
IOSXE_ANY_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?:
        any |
        (?P<SRC_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<SRC_WC>\d+\.\d+\.\d+\.\d+)
    )
    (?:\s+eq\s+.*?)?
    (?:\s+range\s+.*?)?
    \s+
    (?:
        any |
        (?P<DST_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<DST_WC>\d+\.\d+\.\d+\.\d+)
    )
    (?:\s+eq\s+.*?)?
    \s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# ASA格式 - 支持CIDR和IP地址
# 示例: permit tcp 10.10.0.0/16 10.20.0.0/16 eq www
# 示例: permit tcp any any
ASA_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SRC>any|\d+\.\d+\.\d+\.\d+(?:/\d+)?)
    (?:\s+(?:eq\s+)?(?P<PORT_A>\S+))?
    (?:\s+range\s+(?P<PORT_RANGE_START>\S+)\s+(?P<PORT_RANGE_END>\S+))?
    \s+
    (?P<DST>any|\d+\.\d+\.\d+\.\d+(?:/\d+)?)
    (?:\s+(?:eq\s+)?(?P<PORT_B>\S+))?
    \s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# 混合格式 - 源地址wildcard + 目的地址CIDR
# 示例: 99 permit ip 10.6.26.0 0.0.0.255 10.6.26.254/32
IOSXE_WC_SRC_CIDR_DST_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<SOURCE_WILDCARD>\d+\.\d+\.\d+\.\d+)
    \s+(?P<DESTINATION>\d+\.\d+\.\d+\.\d+/\d+)
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# 混合格式 - 源地址CIDR + 目的地址wildcard
# 示例: 94 permit tcp 10.12.8.43/32 eq 28800 10.12.17.80 0.0.0.7
IOSXE_CIDR_SRC_WC_DST_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE>\d+\.\d+\.\d+\.\d+/\d+)
    (?:\s+eq\s+(?P<PORT>\S+))?
    \s+(?P<DESTINATION_IP>\d+\.\d+\.\d+\.\d+)\s+(?P<DESTINATION_WILDCARD>\d+\.\d+\.\d+\.\d+)
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# NXOS格式 - 源端口 + 多端口（源端口后跟多个端口）
# 示例: 需要进一步确认是否存在这种格式
# 暂时不添加，等待实际数据验证

# ============================================================================
# ACL规则数据类
# ============================================================================

# ACL规则数据类：存储解析后的ACL规则的所有信息，包括原始文本、动作、协议、源/目的网络、端口信息、规则格式
@dataclass
class ACLRule:
    # ACL规则数据类
    raw: str
    action: str
    proto: str
    src: IPv4Network
    dst: IPv4Network
    port: Optional[int] = None  # None表示任意端口（用于兼容，通常是第一个端口）
    src_port: Optional[int] = None  # 源端口
    dst_port: Optional[int] = None  # 目标端口
    style: str = ""  # 'NXOS' / 'IOS-XE' / 'ASA'
    ports: Optional[Set[int]] = None  # 多个端口的集合（用于存储eq 22 22222这样的多个端口）

# ============================================================================
# 完整ACL解析函数
# ============================================================================

# 完整解析ACL规则行，支持所有Cisco ACL格式（包括any规则）
# 这是最底层的解析函数，支持所有格式和关键字。如果需要忽略any规则，请使用parse_acl()函数。
# Args: ACL_LINE - ACL规则文本行（可包含行号前缀）
# Returns: (ACLRule对象, None) 成功，或 (None, 错误信息) 失败
# Supported Formats: NXOS、IOS-XE、ASA等所有格式，支持any关键字、log、log-input、time-range等
def parse_acl_full(ACL_LINE: str) -> Tuple[Optional[ACLRule], Optional[str]]:
    CLEANED_LINE = (ACL_LINE or "").strip()
    if not CLEANED_LINE:
        return None, "empty"
    
    # 去除行号前缀（如果存在）
    line_number_pattern = re.compile(r'^\s*\d+\s+(permit|deny)', re.IGNORECASE)
    if line_number_pattern.match(CLEANED_LINE):
        match = re.search(r'\b(permit|deny)\b', CLEANED_LINE, re.IGNORECASE)
        if match:
            CLEANED_LINE = CLEANED_LINE[match.start():]
    
    # 按优先级尝试匹配各种格式（从最具体到最通用）
    
    # 1. NXOS源端口格式（优先，因为更具体）
    MATCH_RESULT = NXOS_SRC_PORT_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT")) if MATCH_RESULT.group("PORT") else None
        DST_PORT = service_to_port(MATCH_RESULT.group("DST_PORT")) if MATCH_RESULT.group("DST_PORT") else None
        SOURCE_NETWORK = cidr_to_network(MATCH_RESULT.group("SOURCE"))
        DESTINATION_NETWORK = cidr_to_network(MATCH_RESULT.group("DESTINATION"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                PORT_NUMBER,  # 源端口
                DST_PORT,  # 目标端口
                "NXOS"
            ), None
        return None, "nxos_src_port_network_parse_fail"
    
    # 2. NXOS端口范围格式
    MATCH_RESULT = NXOS_RANGE_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        SOURCE_NETWORK = cidr_to_network(MATCH_RESULT.group("SOURCE"))
        if MATCH_RESULT.group("DESTINATION_BEFORE"):
            DESTINATION_NETWORK = cidr_to_network(MATCH_RESULT.group("DESTINATION_BEFORE"))
        elif MATCH_RESULT.group("DESTINATION_AFTER"):
            DESTINATION_NETWORK = cidr_to_network(MATCH_RESULT.group("DESTINATION_AFTER"))
        else:
            DESTINATION_NETWORK = None
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                None,  # 端口范围，设为None
                None,
                None,
                "NXOS"
            ), None
        return None, "nxos_range_network_parse_fail"
    
    # 3. NXOS标准格式
    MATCH_RESULT = NXOS_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT")) if MATCH_RESULT.group("PORT") else None
        SOURCE_NETWORK = cidr_to_network(MATCH_RESULT.group("SOURCE"))
        DESTINATION_NETWORK = cidr_to_network(MATCH_RESULT.group("DESTINATION"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                None,
                PORT_NUMBER,
                "NXOS"
            ), None
        return None, "nxos_network_parse_fail"
    
    # 4. IOS-XE多端口格式（源IP+通配符，目的IP+通配符）
    MATCH_RESULT = IOSXE_MULTI_EQ_WILDCARD_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_MULTI_STR = MATCH_RESULT.group("PORT_MULTI").strip()
        ALL_PORTS = _extract_all_ports(PORT_MULTI_STR)
        SOURCE_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("SOURCE_IP"),
            MATCH_RESULT.group("SOURCE_WILDCARD")
        )
        DESTINATION_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("DESTINATION_IP"),
            MATCH_RESULT.group("DESTINATION_WILDCARD")
        )
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            PORT_NUMBER = min(ALL_PORTS) if ALL_PORTS else None
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                None,
                None,
                "IOS-XE",
                ALL_PORTS if len(ALL_PORTS) > 1 else None
            ), None
        return None, "iosxe_multi_eq_wildcard_network_parse_fail"
    
    # 5. IOS-XE端口范围格式（源IP+通配符，目的IP+通配符）
    MATCH_RESULT = IOSXE_RANGE_WILDCARD_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        SOURCE_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("SOURCE_IP"),
            MATCH_RESULT.group("SOURCE_WILDCARD")
        )
        DESTINATION_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("DESTINATION_IP"),
            MATCH_RESULT.group("DESTINATION_WILDCARD")
        )
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                None,
                None,
                None,
                "IOS-XE"
            ), None
        return None, "iosxe_range_wildcard_network_parse_fail"
    
    # 6. IOS-XE host + 多端口 + wildcard格式
    MATCH_RESULT = IOSXE_HOST_MULTI_EQ_WILDCARD_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_MULTI_STR = MATCH_RESULT.group("PORT_MULTI").strip()
        ALL_PORTS = _extract_all_ports(PORT_MULTI_STR)
        SOURCE_NETWORK = host_to_network(MATCH_RESULT.group("SOURCE_IP"))
        DESTINATION_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("DESTINATION_IP"),
            MATCH_RESULT.group("DESTINATION_WILDCARD")
        )
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            PORT_NUMBER = min(ALL_PORTS) if ALL_PORTS else None
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                PORT_NUMBER,
                None,
                "IOS-XE",
                ALL_PORTS if len(ALL_PORTS) > 1 else None
            ), None
        return None, "iosxe_host_multi_eq_wildcard_network_parse_fail"
    
    # 7. IOS-XE range + multi eq格式（源IP+通配符，目的host）
    MATCH_RESULT = IOSXE_RANGE_MULTI_EQ_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        SOURCE_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("SOURCE_IP"),
            MATCH_RESULT.group("SOURCE_WILDCARD")
        )
        DESTINATION_NETWORK = host_to_network(MATCH_RESULT.group("DESTINATION_IP"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                None,
                None,
                None,
                "IOS-XE"
            ), None
        return None, "iosxe_range_multi_eq_network_parse_fail"
    
    # 8. IOS-XE range格式（源host，目的IP+通配符）
    MATCH_RESULT = IOSXE_RANGE_HOST_SRC_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        SOURCE_NETWORK = host_to_network(MATCH_RESULT.group("SOURCE_IP"))
        DESTINATION_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("DESTINATION_IP"),
            MATCH_RESULT.group("DESTINATION_WILDCARD")
        )
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                None,
                None,
                None,
                "IOS-XE"
            ), None
        return None, "iosxe_range_host_src_network_parse_fail"
    
    # 9. IOS-XE混合格式（host和wildcard混合）
    MATCH_RESULT = IOSXE_MIX_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_A_STR = MATCH_RESULT.group("PORT_A")
        PORT_B_STR = MATCH_RESULT.group("PORT_B")
        
        # 判断端口位置
        port_a_before_dst = False
        if PORT_A_STR and not PORT_B_STR:
            eq_pos = CLEANED_LINE.lower().find("eq " + PORT_A_STR.lower().split()[0])
            dst_host_pos = CLEANED_LINE.lower().find("host", eq_pos if eq_pos >= 0 else 0)
            dst_wildcard_pos = -1
            if MATCH_RESULT.group("DESTINATION_IP_WILDCARD"):
                dst_ip_wildcard = MATCH_RESULT.group("DESTINATION_IP_WILDCARD")
                dst_wildcard_pos = CLEANED_LINE.lower().find(dst_ip_wildcard.lower(), eq_pos if eq_pos >= 0 else 0)
            if eq_pos >= 0:
                if dst_host_pos > eq_pos or (dst_wildcard_pos > eq_pos and dst_wildcard_pos > 0):
                    port_a_before_dst = True
        
        ALL_PORTS = set()
        SOURCE_PORT = None
        DESTINATION_PORT = None
        
        if port_a_before_dst:
            if PORT_A_STR:
                ALL_PORTS.update(_extract_all_ports(PORT_A_STR))
                PORT_A_LIST = PORT_A_STR.strip().split()
                if PORT_A_LIST:
                    DESTINATION_PORT = service_to_port(PORT_A_LIST[0])
            if PORT_B_STR:
                ALL_PORTS.update(_extract_all_ports(PORT_B_STR))
                PORT_B_LIST = PORT_B_STR.strip().split()
                if PORT_B_LIST:
                    SOURCE_PORT = service_to_port(PORT_B_LIST[0])
        else:
            if PORT_A_STR:
                ALL_PORTS.update(_extract_all_ports(PORT_A_STR))
                PORT_A_LIST = PORT_A_STR.strip().split()
                if PORT_A_LIST:
                    SOURCE_PORT = service_to_port(PORT_A_LIST[0])
            if PORT_B_STR:
                ALL_PORTS.update(_extract_all_ports(PORT_B_STR))
                PORT_B_LIST = PORT_B_STR.strip().split()
                if PORT_B_LIST:
                    DESTINATION_PORT = service_to_port(PORT_B_LIST[0])
        
        if SOURCE_PORT is not None and DESTINATION_PORT is not None and SOURCE_PORT != DESTINATION_PORT:
            return None, "conflicting_ports"
        
        PORT_NUMBER = DESTINATION_PORT if DESTINATION_PORT is not None else SOURCE_PORT
        SOURCE_NETWORK = host_to_network(MATCH_RESULT.group("SOURCE_IP_HOST")) if MATCH_RESULT.group("SOURCE_IP_HOST") else ip_and_wildcard_to_network(MATCH_RESULT.group("SOURCE_IP_WILDCARD"), MATCH_RESULT.group("SOURCE_WILDCARD_WILDCARD"))
        DESTINATION_NETWORK = host_to_network(MATCH_RESULT.group("DESTINATION_IP_HOST")) if MATCH_RESULT.group("DESTINATION_IP_HOST") else ip_and_wildcard_to_network(MATCH_RESULT.group("DESTINATION_IP_WILDCARD"), MATCH_RESULT.group("DESTINATION_WILDCARD_WILDCARD"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            PORTS_SET = ALL_PORTS if len(ALL_PORTS) > 1 else (ALL_PORTS if ALL_PORTS else None)
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                SOURCE_PORT,
                DESTINATION_PORT,
                "IOS-XE",
                PORTS_SET
            ), None
        return None, "iosxe_mix_network_parse_fail"
    
    # 10. IOS-XE Wildcard格式（标准）
    MATCH_RESULT = IOSXE_WC_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT_A")) or service_to_port(MATCH_RESULT.group("PORT_B"))
        SOURCE_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("SOURCE_IP"),
            MATCH_RESULT.group("SOURCE_WILDCARD")
        )
        DESTINATION_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("DESTINATION_IP"),
            MATCH_RESULT.group("DESTINATION_WILDCARD")
        )
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                None,
                None,
                "IOS-XE"
            ), None
        return None, "iosxe_wc_network_parse_fail"
    
    # 11. IOS-XE Host格式
    MATCH_RESULT = IOSXE_HOST_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT_A")) or service_to_port(MATCH_RESULT.group("PORT_B"))
        SOURCE_NETWORK = host_to_network(MATCH_RESULT.group("SOURCE_IP"))
        DESTINATION_NETWORK = host_to_network(MATCH_RESULT.group("DESTINATION_IP"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                None,
                None,
                "IOS-XE"
            ), None
        return None, "iosxe_host_network_parse_fail"
    
    # 12. 混合格式 - 源地址wildcard + 目的地址CIDR
    MATCH_RESULT = IOSXE_WC_SRC_CIDR_DST_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        SOURCE_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("SOURCE_IP"),
            MATCH_RESULT.group("SOURCE_WILDCARD")
        )
        DESTINATION_NETWORK = cidr_to_network(MATCH_RESULT.group("DESTINATION"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                None,
                None,
                None,
                "IOS-XE"
            ), None
        return None, "iosxe_wc_src_cidr_dst_network_parse_fail"
    
    # 13. 混合格式 - 源地址CIDR + 目的地址wildcard
    MATCH_RESULT = IOSXE_CIDR_SRC_WC_DST_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT")) if MATCH_RESULT.group("PORT") else None
        SOURCE_NETWORK = cidr_to_network(MATCH_RESULT.group("SOURCE"))
        DESTINATION_NETWORK = ip_and_wildcard_to_network(
            MATCH_RESULT.group("DESTINATION_IP"),
            MATCH_RESULT.group("DESTINATION_WILDCARD")
        )
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                PORT_NUMBER,  # 源端口
                None,
                "IOS-XE"
            ), None
        return None, "iosxe_cidr_src_wc_dst_network_parse_fail"
    
    # 14. IOS-XE Any关键字格式（host + any）
    # 示例: 10 permit tcp host 10.10.80.1 any eq 22 log
    IOSXE_HOST_ANY_RE = re.compile(
        r"""
        ^\s*(?P<NUMBER>\d+)?\s*
        (?P<ACTION>permit|deny)\s+
        (?P<PROTOCOL>\S+)\s+
        host\s+(?P<SOURCE_IP>\d+\.\d+\.\d+\.\d+)
        (?:\s+eq\s+(?P<PORT>\S+))?
        \s+any
        (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
        """,
        re.IGNORECASE | re.VERBOSE,
    )
    MATCH_RESULT = IOSXE_HOST_ANY_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        SOURCE_NETWORK = host_to_network(MATCH_RESULT.group("SOURCE_IP"))
        DESTINATION_NETWORK = any_to_network()
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT")) if MATCH_RESULT.group("PORT") else None
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                None,
                PORT_NUMBER,
                "IOS-XE"
            ), None
        return None, "iosxe_host_any_network_parse_fail"
    
    # 15. IOS-XE Any关键字格式（any + host）
    # 示例: 10 permit tcp any host 10.10.80.1 eq 22 log
    IOSXE_ANY_HOST_RE = re.compile(
        r"""
        ^\s*(?P<NUMBER>\d+)?\s*
        (?P<ACTION>permit|deny)\s+
        (?P<PROTOCOL>\S+)\s+
        any
        (?:\s+eq\s+(?P<PORT>\S+))?
        \s+host\s+(?P<DESTINATION_IP>\d+\.\d+\.\d+\.\d+)
        (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
        """,
        re.IGNORECASE | re.VERBOSE,
    )
    MATCH_RESULT = IOSXE_ANY_HOST_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        SOURCE_NETWORK = any_to_network()
        DESTINATION_NETWORK = host_to_network(MATCH_RESULT.group("DESTINATION_IP"))
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT")) if MATCH_RESULT.group("PORT") else None
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                None,
                PORT_NUMBER,
                "IOS-XE"
            ), None
        return None, "iosxe_any_host_network_parse_fail"
    
    # 16. IOS-XE Any关键字格式（wildcard + any 或 any + wildcard）
    MATCH_RESULT = IOSXE_ANY_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        if MATCH_RESULT.group("SRC_IP") and MATCH_RESULT.group("SRC_WC"):
            SOURCE_NETWORK = ip_and_wildcard_to_network(
                MATCH_RESULT.group("SRC_IP"),
                MATCH_RESULT.group("SRC_WC")
            )
        else:
            SOURCE_NETWORK = any_to_network()
        
        if MATCH_RESULT.group("DST_IP") and MATCH_RESULT.group("DST_WC"):
            DESTINATION_NETWORK = ip_and_wildcard_to_network(
                MATCH_RESULT.group("DST_IP"),
                MATCH_RESULT.group("DST_WC")
            )
        else:
            DESTINATION_NETWORK = any_to_network()
        
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                None,
                None,
                None,
                "IOS-XE"
            ), None
        return None, "iosxe_any_network_parse_fail"
    
    # 17. ASA格式
    MATCH_RESULT = ASA_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        src_str = MATCH_RESULT.group("SRC")
        dst_str = MATCH_RESULT.group("DST")
        
        if src_str.lower() == "any":
            SOURCE_NETWORK = any_to_network()
        elif "/" in src_str:
            SOURCE_NETWORK = cidr_to_network(src_str)
        else:
            SOURCE_NETWORK = host_to_network(src_str)
        
        if dst_str.lower() == "any":
            DESTINATION_NETWORK = any_to_network()
        elif "/" in dst_str:
            DESTINATION_NETWORK = cidr_to_network(dst_str)
        else:
            DESTINATION_NETWORK = host_to_network(dst_str)
        
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT_A")) or service_to_port(MATCH_RESULT.group("PORT_B"))
        
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(
                CLEANED_LINE,
                MATCH_RESULT.group("ACTION").lower(),
                MATCH_RESULT.group("PROTOCOL").lower(),
                SOURCE_NETWORK,
                DESTINATION_NETWORK,
                PORT_NUMBER,
                None,
                None,
                "ASA"
            ), None
        return None, "asa_network_parse_fail"
    
    return None, "no_pattern_match"

# ============================================================================
# ACL规则判断函数
# ============================================================================

# 判断文本是否为有效的ACL规则：通过检查关键字和格式，排除明显的非ACL配置行（如as-path、route-map、prefix-list等配置命令）
# Args: text - 待检查的文本行
# Returns: True表示是ACL规则，False表示不是
def is_acl_rule(text: str) -> bool:
    text = text.strip().lower()
    
    # 必须包含permit或deny
    if 'permit' not in text and 'deny' not in text:
        return False
    
    # 排除配置命令（以这些关键字开头的行，但不是ACL规则）
    exclude_patterns = [
        r'^ip\s+access-list\s+',  # ACL定义行
        r'^ip\s+as-path\s+access-list',  # as-path access-list
        r'^ip\s+prefix-list',  # prefix-list
        r'^ip\s+community-list',  # community-list
        r'^route-map\s+',  # route-map
        r'^logging\s+host',  # logging配置
        r'^certificate',  # certificate
        r'^crypto',  # crypto
        r'^interface',  # interface
        r'^router',  # router
        r'^version',  # version
        r'^hostname',  # hostname
        r'^enable',  # enable
        r'^password',  # password
        r'^username',  # username
        r'^line\s+',  # line
        r'^service\s+',  # service
        r'^logging\s+',  # logging
        r'^ntp\s+',  # ntp
        r'^snmp\s+',  # snmp
        r'^tacacs',  # tacacs
        r'^radius',  # radius
        r'^no\s+arp',  # no arp命令
    ]
    
    for pattern in exclude_patterns:
        if re.match(pattern, text, re.IGNORECASE):
            return False
    
    # 必须包含协议或IP地址
    if not (re.search(r'\b(tcp|udp|ip|icmp|ospf|eigrp|gre|esp|ah)\b', text) or 
            re.search(r'\d+\.\d+\.\d+\.\d+', text)):
        return False
    
    return True

# ============================================================================
# 通用ACL解析函数（忽略any规则）
# ============================================================================

# 解析ACL规则行，返回完整的ACLRule对象（忽略any规则）
# 这是最常用的ACL解析函数，适用于需要完整规则信息的场景。会自动忽略包含'any'关键字的规则。
# Args: ACL_LINE - ACL规则文本行（可包含行号前缀）
# Returns: (ACLRule对象, None) 成功，或 (None, 错误信息) 失败，或 (None, "contains_any") any规则
# Note: 自动去除行号前缀，自动忽略log-input、time-range等关键字，any规则会被过滤
def parse_acl(ACL_LINE: str) -> Tuple[Optional[ACLRule], Optional[str]]:
    CLEANED_LINE = (ACL_LINE or "").strip()
    if not CLEANED_LINE:
        return None, "empty"
    
    # 去除行号前缀（如果存在）：行号通常是开头的数字，后跟空格
    # 例如："464 permit ip ..." -> "permit ip ..."
    # 匹配开头的数字（可能有多位）后跟空格，然后才是permit/deny
    line_number_pattern = re.compile(r'^\s*\d+\s+(permit|deny)', re.IGNORECASE)
    if line_number_pattern.match(CLEANED_LINE):
        # 找到第一个permit或deny的位置
        match = re.search(r'\b(permit|deny)\b', CLEANED_LINE, re.IGNORECASE)
        if match:
            CLEANED_LINE = CLEANED_LINE[match.start():]
    
    # 忽略any规则
    if "any" in CLEANED_LINE.lower():
        return None, "contains_any"
    
    # 调用parse_acl_full进行完整解析
    # 注意：由于NXOS_RE已统一为宽松匹配，log-input/time-range等关键字会自动被忽略
    result, error = parse_acl_full(ACL_LINE)
    if error == "contains_any":
        return None, "contains_any"
    return result, error

# ============================================================================
# 简化ACL解析函数（只提取网络信息）
# ============================================================================

# 解析ACL规则行，只提取源/目的网络信息（忽略any规则）
# 适用于只需要网络信息的场景（如网络匹配、ARP检查等），不需要完整的ACLRule对象，性能更好。
# Args: ACL_LINE - ACL规则文本行（可包含行号前缀）
# Returns: ((源网络, 目的网络), None) 成功，或 (None, 错误信息) 失败，或 (None, "contains_any") any规则
def parse_acl_network_only(ACL_LINE: str) -> Tuple[Optional[Tuple[IPv4Network, IPv4Network]], Optional[str]]:
    rule, error = parse_acl(ACL_LINE)
    if rule:
        return (rule.src, rule.dst), None
    return None, error

# ============================================================================
# Excel ACL块处理函数
# ============================================================================

# 在指定列中找到ACL块：以'ip access-list'开始，以包含'vty'和'ip'的'ip access-list'行结束（登录ACL结束标记）
# Args: worksheet - openpyxl Worksheet对象, col - 列号（从1开始）
# Returns: ACL块列表，每个元素为(start_row, end_row)
def find_acl_blocks_in_column(worksheet, col: int) -> List[Tuple[int, int]]:
    ACL_BLOCKS = []
    CURRENT_START = None
    FOUND_VTY = False  # 标记是否遇到登录ACL结束标记
    
    for ROW in range(1, worksheet.max_row + 1):
        CELL_VALUE = worksheet.cell(row=ROW, column=col).value
        if CELL_VALUE and isinstance(CELL_VALUE, str):
            TEXT = str(CELL_VALUE).strip()
            TEXT_LOWER = TEXT.lower()
            
            # 业务ACL开始（排除登录ACL）
            if TEXT.startswith('ip access-list '):
                # 检查是否是登录ACL结束标记（包含vty和ip，忽略大小写）
                # 匹配：ip access-list VTY-ACL-IP 或 ip access-list extended vty-access-IP
                if 'vty' in TEXT_LOWER and 'ip' in TEXT_LOWER:
                    # 登录ACL标记 - 结束当前ACL块，不再处理后续ACL
                    if CURRENT_START is not None:
                        ACL_BLOCKS.append((CURRENT_START, ROW - 1))
                    FOUND_VTY = True
                    break  # 登录ACL及以下的不分析
                else:
                    # 业务ACL开始
                    if CURRENT_START is not None:
                        # 结束上一个ACL块
                        ACL_BLOCKS.append((CURRENT_START, ROW - 1))
                    CURRENT_START = ROW
    
    # 处理最后一个ACL块（只有在没有遇到登录ACL结束标记时才处理）
    if not FOUND_VTY and CURRENT_START is not None:
        ACL_BLOCKS.append((CURRENT_START, worksheet.max_row))
    
    return ACL_BLOCKS

# 从指定列的ACL块中提取ACL规则：忽略any规则，只提取有效的ACL规则
# Args: worksheet - openpyxl Worksheet对象, col - 列号（从1开始）, start_row - ACL块起始行号, end_row - ACL块结束行号
# Returns: 解析后的ACL规则列表
def extract_acl_rules_from_column(worksheet, col: int, start_row: int, end_row: int) -> List[ACLRule]:
    ACL_RULES = []
    for ROW in range(start_row, end_row + 1):
        CELL_VALUE = worksheet.cell(row=ROW, column=col).value
        if CELL_VALUE is None:
            continue
        CLEANED_TEXT = str(CELL_VALUE).strip()
        if not CLEANED_TEXT:
            continue
        
        # 只处理真正的ACL规则，忽略证书等数据
        if not is_acl_rule(CLEANED_TEXT):
            continue
        
        # 解析ACL规则（会自动忽略any规则）
        PARSED_RULE, PARSE_ERROR = parse_acl(CLEANED_TEXT)
        if PARSED_RULE:
            ACL_RULES.append(PARSED_RULE)
    
    return ACL_RULES

