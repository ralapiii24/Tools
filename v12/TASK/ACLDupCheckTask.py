# ACL 重复检查任务

# 导入标准库
import os
import re
import socket
from dataclasses import dataclass
from datetime import datetime
from ipaddress import IPv4Address, IPv4Network
from typing import Dict, Iterable, List, Optional, Set, Tuple

# 导入第三方库
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

# 导入本地应用
from .TaskBase import BaseTask, Level, CONFIG

# 解析辅助函数
# 仅数字或系统services数据库解析；失败返回None（视为任意端口）
def service_to_port(SERVICE_STRING: str) -> Optional[int]:
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
    except Exception:
        return None

# IOS-XE：ip + wildcard转网络（假设通配位连续）
def ip_and_wildcard_to_network(IP_STRING: str, WILDCARD_STRING: str) -> Optional[IPv4Network]:
    try:
        IP_ADDRESS = int(IPv4Address(IP_STRING))
        WILDCARD_ADDRESS = int(IPv4Address(WILDCARD_STRING))
        NETMASK_INTEGER = (~WILDCARD_ADDRESS) & 0xFFFFFFFF
        PREFIX_LENGTH = 32 - bin(WILDCARD_ADDRESS).count("1")
        NETWORK_INTEGER = IP_ADDRESS & NETMASK_INTEGER
        return IPv4Network((IPv4Address(NETWORK_INTEGER), PREFIX_LENGTH), strict=False)
    except Exception:
        return None

# 将IP地址转换为/32网络
def host_to_network(IP_STRING: str) -> Optional[IPv4Network]:
    try:
        return IPv4Network(f"{IP_STRING}/32", strict=False)
    except Exception:
        return None

# 将CIDR字符串转换为网络对象
def cidr_to_network(CIDR_STRING: str) -> Optional[IPv4Network]:
    try:
        return IPv4Network(CIDR_STRING, strict=False)
    except Exception:
        return None

# 正则表达式（支持NX-OS/IOS-XE）
NXOS_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE>\d+\.\d+\.\d+\.\d+\/\d+)\s+
    (?P<DESTINATION>\d+\.\d+\.\d+\.\d+\/\d+)
    (?:\s+eq\s+(?P<PORT>\S+))?
    \s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# NXOS格式：序号 permit 协议 源地址 eq 端口 目标地址
NXOS_SRC_PORT_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?P<SOURCE>\d+\.\d+\.\d+\.\d+\/\d+)\s+
    eq\s+(?P<PORT>\S+)\s+
    (?P<DESTINATION>\d+\.\d+\.\d+\.\d+\/\d+)
    \s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

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

IOSXE_MIX_RE = re.compile(
    r"""
    ^\s*(?P<NUMBER>\d+)?\s*
    (?P<ACTION>permit|deny)\s+
    (?P<PROTOCOL>\S+)\s+
    (?:
        host\s+(?P<SOURCE_IP_HOST>\d+\.\d+\.\d+\.\d+) |
        (?P<SOURCE_IP_WILDCARD>\d+\.\d+\.\d+\.\d+)\s+(?P<SOURCE_WILDCARD_WILDCARD>\d+\.\d+\.\d+\.\d+)
    )
    (?:\s+eq\s+(?P<PORT_A>\S+))?
    \s+
    (?:
        host\s+(?P<DESTINATION_IP_HOST>\d+\.\d+\.\d+\.\d+) |
        (?P<DESTINATION_IP_WILDCARD>\d+\.\d+\.\d+\.\d+)\s+(?P<DESTINATION_WILDCARD_WILDCARD>\d+\.\d+\.\d+\.\d+)
    )
    (?:\s+eq\s+(?P<PORT_B>\S+))?
    (?:\s+(?:log|log-input|time-range\s+\S+|\S+))*\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

# ACL规则数据类：存储解析后的ACL规则信息
@dataclass
class ACLRule:
    raw: str
    action: str
    proto: str
    src: IPv4Network
    dst: IPv4Network
    port: Optional[int]  # None表示任意端口
    src_port: Optional[int]  # 源端口
    dst_port: Optional[int]  # 目标端口
    style: str           # 'NXOS' / 'IOS-XE'

# 解析ACL规则行，支持多种格式
def parse_acl(ACL_LINE: str) -> Tuple[Optional[ACLRule], Optional[str]]:
    CLEANED_LINE = (ACL_LINE or "").strip()
    if not CLEANED_LINE:
        return None, "empty"
    
    # 忽略带有any关键字的ACL规则
    if "any" in CLEANED_LINE.lower():
        return None, "contains_any"

    MATCH_RESULT = NXOS_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT")) if MATCH_RESULT.group("PORT") else None
        SOURCE_NETWORK = cidr_to_network(MATCH_RESULT.group("SOURCE"))
        DESTINATION_NETWORK = cidr_to_network(MATCH_RESULT.group("DESTINATION"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(CLEANED_LINE, MATCH_RESULT.group("ACTION").lower(), MATCH_RESULT.group("PROTOCOL").lower(), SOURCE_NETWORK, DESTINATION_NETWORK, PORT_NUMBER, None, PORT_NUMBER, "NXOS"), None
        return None, "nxos_network_parse_fail"

    MATCH_RESULT = NXOS_SRC_PORT_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT")) if MATCH_RESULT.group("PORT") else None
        SOURCE_NETWORK = cidr_to_network(MATCH_RESULT.group("SOURCE"))
        DESTINATION_NETWORK = cidr_to_network(MATCH_RESULT.group("DESTINATION"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(CLEANED_LINE, MATCH_RESULT.group("ACTION").lower(), MATCH_RESULT.group("PROTOCOL").lower(), SOURCE_NETWORK, DESTINATION_NETWORK, PORT_NUMBER, PORT_NUMBER, None, "NXOS"), None
        return None, "nxos_src_port_network_parse_fail"

    MATCH_RESULT = IOSXE_WC_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT_A")) or service_to_port(MATCH_RESULT.group("PORT_B"))
        SOURCE_NETWORK = ip_and_wildcard_to_network(MATCH_RESULT.group("SOURCE_IP"), MATCH_RESULT.group("SOURCE_WILDCARD"))
        DESTINATION_NETWORK = ip_and_wildcard_to_network(MATCH_RESULT.group("DESTINATION_IP"), MATCH_RESULT.group("DESTINATION_WILDCARD"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(CLEANED_LINE, MATCH_RESULT.group("ACTION").lower(), MATCH_RESULT.group("PROTOCOL").lower(), SOURCE_NETWORK, DESTINATION_NETWORK, PORT_NUMBER, None, None, "IOS-XE"), None
        return None, "iosxe_wc_network_parse_fail"

    MATCH_RESULT = IOSXE_HOST_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        PORT_NUMBER = service_to_port(MATCH_RESULT.group("PORT_A")) or service_to_port(MATCH_RESULT.group("PORT_B"))
        SOURCE_NETWORK = host_to_network(MATCH_RESULT.group("SOURCE_IP"))
        DESTINATION_NETWORK = host_to_network(MATCH_RESULT.group("DESTINATION_IP"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(CLEANED_LINE, MATCH_RESULT.group("ACTION").lower(), MATCH_RESULT.group("PROTOCOL").lower(), SOURCE_NETWORK, DESTINATION_NETWORK, PORT_NUMBER, None, None, "IOS-XE"), None
        return None, "iosxe_host_network_parse_fail"

    MATCH_RESULT = IOSXE_MIX_RE.match(CLEANED_LINE)
    if MATCH_RESULT:
        # 正确处理端口：PORT_A是源端口，PORT_B是目标端口
        SOURCE_PORT = service_to_port(MATCH_RESULT.group("PORT_A"))
        DESTINATION_PORT = service_to_port(MATCH_RESULT.group("PORT_B"))
        # 如果两个端口都存在且不同，这是无效的规则
        if SOURCE_PORT is not None and DESTINATION_PORT is not None and SOURCE_PORT != DESTINATION_PORT:
            return None, "conflicting_ports"
        # 使用目标端口作为主要端口（ACL通常关注目标端口）
        PORT_NUMBER = DESTINATION_PORT if DESTINATION_PORT is not None else SOURCE_PORT
        SOURCE_NETWORK = host_to_network(MATCH_RESULT.group("SOURCE_IP_HOST")) if MATCH_RESULT.group("SOURCE_IP_HOST") else ip_and_wildcard_to_network(MATCH_RESULT.group("SOURCE_IP_WILDCARD"), MATCH_RESULT.group("SOURCE_WILDCARD_WILDCARD"))
        DESTINATION_NETWORK = host_to_network(MATCH_RESULT.group("DESTINATION_IP_HOST")) if MATCH_RESULT.group("DESTINATION_IP_HOST") else ip_and_wildcard_to_network(MATCH_RESULT.group("DESTINATION_IP_WILDCARD"), MATCH_RESULT.group("DESTINATION_WILDCARD_WILDCARD"))
        if SOURCE_NETWORK and DESTINATION_NETWORK:
            return ACLRule(CLEANED_LINE, MATCH_RESULT.group("ACTION").lower(), MATCH_RESULT.group("PROTOCOL").lower(), SOURCE_NETWORK, DESTINATION_NETWORK, PORT_NUMBER, SOURCE_PORT, DESTINATION_PORT, "IOS-XE"), None
        return None, "iosxe_mix_network_parse_fail"

    return None, "no_pattern_match"

# 覆盖关系逻辑
# 检查协议A是否覆盖协议B
def proto_covers(PROTO_A: str, PROTO_B: str) -> bool:
    return PROTO_A.lower() == "ip" or PROTO_A.lower() == PROTO_B.lower()

# 检查端口A是否覆盖端口B
def port_covers(PORT_A: Optional[int], PORT_B: Optional[int]) -> bool:
    # A None => 任意端口；B None 且 A 有端口 => 不覆盖；都有端口需相等
    if PORT_A is None:
        return True
    if PORT_B is None:
        return False
    return PORT_A == PORT_B

# 检查源端口A是否覆盖源端口B
def src_port_covers(SRC_PORT_A: Optional[int], SRC_PORT_B: Optional[int]) -> bool:
    if SRC_PORT_A is None:
        return True
    if SRC_PORT_B is None:
        return False
    return SRC_PORT_A == SRC_PORT_B

# 检查目标端口A是否覆盖目标端口B
def dst_port_covers(DST_PORT_A: Optional[int], DST_PORT_B: Optional[int]) -> bool:
    if DST_PORT_A is None:
        return True  # 任意端口覆盖特定端口
    if DST_PORT_B is None:
        return False  # 特定端口不覆盖任意端口
    return DST_PORT_A == DST_PORT_B

# 检查规则A是否覆盖规则B
def rule_covers(RULE_A: ACLRule, RULE_B: ACLRule) -> bool:
    if RULE_A.action != RULE_B.action:
        return False
    if not proto_covers(RULE_A.proto, RULE_B.proto):
        return False
    
    # 对于NXOS格式，使用port字段；对于IOS-XE格式，使用src_port和dst_port字段
    if RULE_A.style == "NXOS" and RULE_B.style == "NXOS":
        # NXOS格式：检查源端口和目标端口
        if not src_port_covers(RULE_A.src_port, RULE_B.src_port):
            return False
        if not dst_port_covers(RULE_A.dst_port, RULE_B.dst_port):
            return False
    else:
        # IOS-XE格式：检查源端口和目标端口
        # 如果src_port和dst_port都为None，则使用port字段
        if RULE_A.src_port is None and RULE_A.dst_port is None and RULE_B.src_port is None and RULE_B.dst_port is None:
            # 都使用port字段
            if not port_covers(RULE_A.port, RULE_B.port):
                return False
        else:
            # 使用src_port和dst_port字段
            # 但是要确保端口信息的一致性
            if not src_port_covers(RULE_A.src_port, RULE_B.src_port):
                return False
            if not dst_port_covers(RULE_A.dst_port, RULE_B.dst_port):
                return False
            # 如果两个规则都有port字段，也要检查port字段
            if RULE_A.port is not None and RULE_B.port is not None:
                if not port_covers(RULE_A.port, RULE_B.port):
                    return False
    
    if not RULE_B.src.subnet_of(RULE_A.src):
        return False
    if not RULE_B.dst.subnet_of(RULE_A.dst):
        return False
    return True

# 图论算法辅助函数
# 使用图论算法计算连通分量
def connected_components(NODES: Iterable[int], UNDIRECTED_EDGES: List[Tuple[int, int]]) -> List[Set[int]]:
    GRAPH: Dict[int, Set[int]] = {NODE: set() for NODE in NODES}
    for NODE_U, NODE_V in UNDIRECTED_EDGES:
        GRAPH.setdefault(NODE_U, set()).add(NODE_V)
        GRAPH.setdefault(NODE_V, set()).add(NODE_U)
    VISITED_NODES: Set[int] = set()
    COMPONENTS: List[Set[int]] = []
    for NODE in NODES:
        if NODE in VISITED_NODES:
            continue
        STACK = [NODE]
        COMPONENT = set()
        while STACK:
            CURRENT_NODE = STACK.pop()
            if CURRENT_NODE in VISITED_NODES:
                continue
            VISITED_NODES.add(CURRENT_NODE)
            COMPONENT.add(CURRENT_NODE)
            STACK.extend(GRAPH.get(CURRENT_NODE, []))
        COMPONENTS.append(COMPONENT)
    return COMPONENTS

# 严格连通分量算法：只考虑直接覆盖关系，避免间接连接
# 严格连通分量算法：只将直接有覆盖关系的规则归为一组，避免通过中间规则间接连接不相关的规则
def _strict_connected_components(NODES: Iterable[int], UNDIRECTED_EDGES: List[Tuple[int, int]]) -> List[Set[int]]:
    if not UNDIRECTED_EDGES:
        return []
    
    # 构建邻接表
    GRAPH: Dict[int, Set[int]] = {NODE: set() for NODE in NODES}
    for NODE_U, NODE_V in UNDIRECTED_EDGES:
        GRAPH.setdefault(NODE_U, set()).add(NODE_V)
        GRAPH.setdefault(NODE_V, set()).add(NODE_U)
    
    VISITED_NODES: Set[int] = set()
    COMPONENTS: List[Set[int]] = []
    
    for NODE in NODES:
        if NODE in VISITED_NODES:
            continue
        
        # 使用BFS找到所有直接连接的节点
        QUEUE = [NODE]
        COMPONENT = set()
        
        while QUEUE:
            CURRENT_NODE = QUEUE.pop(0)
            if CURRENT_NODE in VISITED_NODES:
                continue
            VISITED_NODES.add(CURRENT_NODE)
            COMPONENT.add(CURRENT_NODE)
            
            # 只添加直接相邻的节点
            for NEIGHBOR in GRAPH.get(CURRENT_NODE, []):
                if NEIGHBOR not in VISITED_NODES:
                    QUEUE.append(NEIGHBOR)
        
        if COMPONENT:
            COMPONENTS.append(COMPONENT)
    
    return COMPONENTS

# 智能连通分量算法：避免通过被多个规则覆盖的中间节点创建不合理的间接连接
# 智能连通分量算法：避免通过被多个规则覆盖的中间节点创建不合理的间接连接，如果某个节点被多个规则覆盖，则将其从连通图中移除，避免不合理的间接连接
def _smart_connected_components(NODES: Iterable[int], UNDIRECTED_EDGES: List[Tuple[int, int]], DIRECTED_EDGES: List[Tuple[int, int]]) -> List[Set[int]]:
    if not UNDIRECTED_EDGES:
        return []
    
    # 统计每个节点的入度（被多少个规则覆盖）
    INDEGREE = {}
    for NODE in NODES:
        INDEGREE[NODE] = 0
    
    for FROM_NODE, TO_NODE in DIRECTED_EDGES:
        if TO_NODE in INDEGREE:
            INDEGREE[TO_NODE] += 1
    
    # 找出被多个规则覆盖的节点（入度 > 1）
    MULTI_COVERED_NODES = {NODE for NODE, DEGREE in INDEGREE.items() if DEGREE > 1}
    
    # 为被多个规则覆盖的节点找到最靠前的覆盖者
    NODE_TO_GROUP = {}
    for NODE in MULTI_COVERED_NODES:
        # 找到所有覆盖这个节点的规则
        COVERERS = []
        for FROM_NODE, TO_NODE in DIRECTED_EDGES:
            if TO_NODE == NODE:
                COVERERS.append(FROM_NODE)
        # 选择行号最小的覆盖者
        if COVERERS:
            NODE_TO_GROUP[NODE] = min(COVERERS)
    
    # 构建邻接表，将被多个规则覆盖的节点连接到最靠前的覆盖者
    GRAPH: Dict[int, Set[int]] = {NODE: set() for NODE in NODES}
    for NODE_U, NODE_V in UNDIRECTED_EDGES:
        # 如果节点V被多个规则覆盖，连接到其最靠前的覆盖者
        if NODE_V in MULTI_COVERED_NODES and NODE_V in NODE_TO_GROUP:
            TARGET_GROUP = NODE_TO_GROUP[NODE_V]
            if NODE_U == TARGET_GROUP:
                GRAPH.setdefault(NODE_U, set()).add(NODE_V)
                GRAPH.setdefault(NODE_V, set()).add(NODE_U)
        # 如果节点U被多个规则覆盖，连接到其最靠前的覆盖者
        elif NODE_U in MULTI_COVERED_NODES and NODE_U in NODE_TO_GROUP:
            TARGET_GROUP = NODE_TO_GROUP[NODE_U]
            if NODE_V == TARGET_GROUP:
                GRAPH.setdefault(NODE_U, set()).add(NODE_V)
                GRAPH.setdefault(NODE_V, set()).add(NODE_U)
        # 正常连接
        else:
            GRAPH.setdefault(NODE_U, set()).add(NODE_V)
            GRAPH.setdefault(NODE_V, set()).add(NODE_U)
    
    VISITED_NODES: Set[int] = set()
    COMPONENTS: List[Set[int]] = []
    
    for NODE in NODES:
        if NODE in VISITED_NODES:
            continue
        
        # 使用BFS找到所有连接的节点
        QUEUE = [NODE]
        COMPONENT = set()
        
        while QUEUE:
            CURRENT_NODE = QUEUE.pop(0)
            if CURRENT_NODE in VISITED_NODES:
                continue
            VISITED_NODES.add(CURRENT_NODE)
            COMPONENT.add(CURRENT_NODE)
            
            # 添加相邻的节点
            for NEIGHBOR in GRAPH.get(CURRENT_NODE, []):
                if NEIGHBOR not in VISITED_NODES:
                    QUEUE.append(NEIGHBOR)
        
        if COMPONENT:
            COMPONENTS.append(COMPONENT)
    
    return COMPONENTS


# 基础颜色调色板 - 精选20种高对比度颜色
FILL_COLORS = [
    "FFFDE68A",  # amber-300
    "FFA7F3D0",  # teal-200  
    "FFFCA5A5",  # red-300
    "FF93C5FD",  # blue-300
    "FFA5B4FC",  # violet-300
    "FF86EFAC",  # green-300
    "FFFBCFE8",  # pink-200
    "FFE9D5FF",  # purple-200
    "FF67E8F9",  # cyan-300
    "FFFDE1AF",  # custom light
    "FFDBEAFE",  # blue-100
    "FFDCFCE7",  # green-100
    "FFFEE2E2",  # red-100
    "FFFEFCE8",  # yellow-100
    "FFF3E8FF",  # purple-100
    "FFFFF7ED",  # orange-100
    "FFF9FAFB",  # gray-100
    "FFECFEFF",  # cyan-100
    "FFFDF2F8",  # pink-100
    "FFEEF2FF",  # indigo-100
]

# 根据索引生成填充颜色
def fill_for_index(idx: int) -> PatternFill:
    color = FILL_COLORS[idx % len(FILL_COLORS)]
    return PatternFill(start_color=color, end_color=color, fill_type="solid")

# 红字样式（仅颜色，如需加粗可设 bold=True）
RED_FONT = Font(color="FFFF0000")

# 判断文本是否为ACL规则
def _is_acl_rule(text: str) -> bool:
    text = text.strip().lower()
    # 排除证书、配置等非ACL数据
    if any(keyword in text for keyword in [
        'certificate', 'crypto', 'quit', 'exit', 'config', 'interface',
        'router', 'version', 'hostname', 'enable', 'password', 'username',
        'line', 'service', 'logging', 'ntp', 'snmp', 'tacacs', 'radius'
    ]):
        return False
    
    # 只处理包含permit/deny的规则
    return 'permit' in text or 'deny' in text


# 从DeviceBackupTask.py复制的设备分类规则
# 返回设备分类规则字典，包含分组策略
def _get_device_classification_rules():
    return {
        "cat1": {
            "name": "N9K核心交换机",
            "patterns": [
                # CS + N9K + (01|02|03|04)，兼容连写
                lambda filenameLower: (("cs" in filenameLower) or re.search(r"\bcs\b", filenameLower)) and 
                         (("n9k" in filenameLower) or re.search(r"\bn9k\b", filenameLower)) and 
                         re.search(r"(?:^|[^0-9])0?[1-4](?:[^0-9]|$)", filenameLower),
                # CS连写模式
                lambda filenameLower: re.search(r"cs0?[1-4]", filenameLower) and (("n9k" in filenameLower) or re.search(r"\bn9k\b", filenameLower))
            ]
        },
        "cat2": {
            "name": "LINKAS接入交换机",
            "patterns": [
                # LINK + AS + (01|02)，兼容连写
                lambda filenameLower: (("link" in filenameLower) or re.search(r"\blink\b", filenameLower)) and 
                         (("as" in filenameLower) or re.search(r"\bas\b", filenameLower)) and 
                         re.search(r"(?:^|[^0-9])0?[12](?:[^0-9]|$)", filenameLower),
                # LINKAS连写模式
                lambda filenameLower: re.search(r"link[-_]*as0?[12]", filenameLower),
                # AS连写模式
                lambda filenameLower: (("link" in filenameLower) or re.search(r"\blink\b", filenameLower)) and re.search(r"as0?[12]", filenameLower)
            ]
        }
    }

# 检查文本是否匹配cat1设备（N9K核心交换机）
def _is_cat1_device(text: str) -> bool:
    text_lower = text.lower()
    rules = _get_device_classification_rules()
    for pattern_func in rules["cat1"]["patterns"]:
        if pattern_func(text_lower):
            return True
    return False

# 检查文本是否匹配cat2设备（LINKAS接入交换机）
def _is_cat2_device(text: str) -> bool:
    text_lower = text.lower()
    rules = _get_device_classification_rules()
    for pattern_func in rules["cat2"]["patterns"]:
        if pattern_func(text_lower):
            return True
    return False

# 分析第一行，识别cat1/cat2列
def analyze_first_row_for_cat1_cat2(worksheet):
    cat1_cols = []
    cat2_cols = []
    
    for COLUMN in range(1, worksheet.max_column + 1):
        first_cell = worksheet.cell(row=1, column=COLUMN).value
        if first_cell and isinstance(first_cell, str):
            if _is_cat1_device(first_cell):
                cat1_cols.append(COLUMN)
            elif _is_cat2_device(first_cell):
                cat2_cols.append(COLUMN)
    
    return cat1_cols, cat2_cols

# 在指定列中找到ACL块，以ip access-list开始，以包含vty和ip的ip access-list行结束
def find_acl_blocks_in_column(worksheet, col):
    acl_blocks = []
    current_start = None
    found_vty = False  # 标记是否遇到登录ACL结束标记
    
    for ROW in range(1, worksheet.max_row + 1):
        cell_value = worksheet.cell(row=ROW, column=col).value
        if cell_value and isinstance(cell_value, str):
            text = str(cell_value).strip()
            text_lower = text.lower()
            
            # 业务ACL开始（排除登录ACL）
            if text.startswith('ip access-list '):
                # 检查是否是登录ACL结束标记（包含vty和ip，忽略大小写）
                # 匹配：ip access-list VTY-ACL-IP 或 ip access-list extended vty-access-IP
                if 'vty' in text_lower and 'ip' in text_lower:
                    # 登录ACL标记 - 结束当前ACL块，不再处理后续ACL
                    if current_start is not None:
                        acl_blocks.append((current_start, ROW - 1))
                    found_vty = True
                    break  # 登录ACL及以下的不分析
                else:
                    # 业务ACL开始
                    if current_start is not None:
                        # 结束上一个ACL块
                        acl_blocks.append((current_start, ROW - 1))
                    current_start = ROW
    
    # 处理最后一个ACL块（只有在没有遇到登录ACL结束标记时才处理）
    if not found_vty and current_start is not None:
        acl_blocks.append((current_start, worksheet.max_row))
    
    return acl_blocks

# 处理单个ACL块内的覆盖关系
def process_acl_block(worksheet, col, start_row, end_row):
    # 采集ACL块内的所有规则
    col_rules: Dict[int, ACLRule] = {}
    for ROW in range(start_row, end_row + 1):
        cell_value = worksheet.cell(row=ROW, column=col).value
        if cell_value is None:
            continue
        cleaned_text = str(cell_value).strip()
        if not cleaned_text:
            continue
        
        # 只处理真正的ACL规则，忽略证书等数据
        if not _is_acl_rule(cleaned_text):
            continue
        
        # 解析ACL规则
        parsed_rule, parse_error = parse_acl(cleaned_text)
        if parsed_rule:
            col_rules[ROW] = parsed_rule

    # 检查是否有足够的规则
    if len(col_rules) < 2:
        return 0, 0, 0, 0  # groups, keep, recycle, total_in_groups

    # 覆盖关系（同列内）- 只考虑直接覆盖关系，避免间接连接
    directed_edges: List[Tuple[int, int]] = []    # A->B 表示 A 覆盖 B
    undirected_edges: List[Tuple[int, int]] = []  # 分组用
    rows = sorted(col_rules.keys())
    
    for rule_i in range(len(rows)):
        for rule_j in range(rule_i + 1, len(rows)):
            row_i, row_j = rows[rule_i], rows[rule_j]
            rule_a, rule_b = col_rules[row_i], col_rules[row_j]
            a_covers_b = rule_covers(rule_a, rule_b)
            b_covers_a = rule_covers(rule_b, rule_a)
            # 只有当两个规则之间有直接覆盖关系时才连接
            if a_covers_b or b_covers_a:
                undirected_edges.append((row_i, row_j))
                if a_covers_b:
                    directed_edges.append((row_i, row_j))
                if b_covers_a:
                    directed_edges.append((row_j, row_i))

    if not undirected_edges:
        return 0, 0, 0, 0  # 当前列没有覆盖关系

    # 仅取参与覆盖关系的节点
    nodes_in_edges: Set[int] = set()
    for node_u, node_v in undirected_edges:
        nodes_in_edges.add(node_u)
        nodes_in_edges.add(node_v)

    # 分量划分（组）- 使用智能分组算法
    # 被多个规则覆盖的节点归属到最靠前的组，避免不合理的间接连接
    comps = _smart_connected_components(sorted(nodes_in_edges), undirected_edges, directed_edges)

    groups_count = 0
    keep_count = 0
    recycle_count = 0
    total_in_groups = 0

    # 每个分量着色 & 最大规则标红 & 统计
    for idx, comp in enumerate(comps):
        groups_count += 1
        fill = fill_for_index(idx)

        # 计算入度（只看分量内部的有向边）
        indeg: Dict[int, int] = {row_number: 0 for row_number in comp}
        for edge_u, edge_v in directed_edges:
            if edge_u in comp and edge_v in comp:
                indeg[edge_v] += 1

        # 最大规则：每个组只有一条，选择入度为0且行号最小的规则
        # 如果同时被多个规则覆盖，归属到靠前的组（行号小的）
        zero_indegree_nodes = [row_number for row_number, degree in indeg.items() if degree == 0]
        if zero_indegree_nodes:
            # 选择行号最小的作为最大规则
            maxima = {min(zero_indegree_nodes)}
        else:
            # 如果没有入度为0的节点，选择行号最小的
            maxima = {min(comp)}
        group_total = len(comp)
        group_keep = len(maxima)
        group_recycle = group_total - group_keep

        total_in_groups += group_total
        keep_count += group_keep
        recycle_count += group_recycle

        # 上底色 - 应用到当前列
        for ROW_NUMBER in comp:
            worksheet.cell(row=ROW_NUMBER, column=col).fill = fill

        # 标红字体 - 应用到当前列的最大规则
        for ROW_NUMBER in maxima:
            cell = worksheet.cell(row=ROW_NUMBER, column=col)
            current_font = cell.font or Font()
            # 强制使用宋体字体
            cell.font = Font(
                name="宋体",
                sz=current_font.sz, 
                b=current_font.b, 
                i=current_font.i,
                underline=current_font.u, 
                strike=current_font.strike, 
                color="FFFF0000"
            )

    return groups_count, keep_count, recycle_count, total_in_groups

# 主处理流程
# 处理Excel文件，执行ACL覆盖检查（优化版本）
def process_file(input_path: str, output_path: str) -> Dict[str, Dict[str, int]]:
    inputWorkbook = load_workbook(input_path)
    per_sheet_stats: Dict[str, Dict[str, int]] = {}  # sheet -> {groups, keep, recycle, total_in_groups}

    for worksheet in inputWorkbook.worksheets:
        sheet_name = worksheet.title

        groups_count = 0
        keep_count = 0       # 红色（保留）
        recycle_count = 0    # 默认字体（可回收）
        total_in_groups = 0  # 参与覆盖组的规则总数

        # 1. 分析第一行，识别cat1/cat2列
        cat1_cols, cat2_cols = analyze_first_row_for_cat1_cat2(worksheet)
        target_cols = cat1_cols + cat2_cols
        
        # 2. 对每个目标列处理ACL块
        for COLUMN in target_cols:
            acl_blocks = find_acl_blocks_in_column(worksheet, COLUMN)
            
            for start_row, end_row in acl_blocks:
                # 3. 处理ACL块内的覆盖关系
                block_groups, block_keep, block_recycle, block_total = process_acl_block(worksheet, COLUMN, start_row, end_row)
                groups_count += block_groups
                keep_count += block_keep
                recycle_count += block_recycle
                total_in_groups += block_total

        per_sheet_stats[sheet_name] = {
            "groups": groups_count,
            "keep": keep_count,
            "recycle": recycle_count,
            "total_in_groups": total_in_groups,
        }

    # 删除 Report 工作表（如果存在）
    if "Report" in inputWorkbook.sheetnames:
        del inputWorkbook["Report"]

    inputWorkbook.save(output_path)
    return per_sheet_stats

# ACL重复检查任务类：分析ACL规则覆盖关系并生成可视化报告
# ACL重复检查任务：分析ACL规则间的覆盖关系，识别可回收的重复策略，使用图论算法进行分组，最大规则标红显示，生成带颜色标注的Excel报告和统计汇总
class ACLDupCheckTask(BaseTask):

    # 初始化ACL重复检查任务：设置固定配置参数
    def __init__(self):
        super().__init__("大段覆盖包含明细ACL检查")
        # 固定配置参数
        today = datetime.now().strftime("%Y%m%d")
        # V10新结构：从 LOG/DeviceBackupTask/ 读取（ACL/SourceACL已迁移）
        self.INPUT_PATH = f"LOG/DeviceBackupTask/{today}-关键设备配置备份输出EXCEL基础任务.xlsx"
        # V10新结构：直接输出到 LOG/ACLDupCheckTask/
        self.OUTPUT_DIR = "LOG/ACLDupCheckTask"
        self.NAME = "大段覆盖包含明细ACL检查"

    # 返回要处理的Sheet列表
    def items(self):
        if not os.path.exists(self.INPUT_PATH):
            return []
        try:
            inputWorkbook = load_workbook(self.INPUT_PATH)
            sheet_names = [worksheet.title for worksheet in inputWorkbook.worksheets if worksheet.title != 'Report']
            inputWorkbook.close()
            return sheet_names
        except Exception:
            return []

    # 处理单个Sheet：执行ACL覆盖检查并生成报告
    def run_single(self, sheet_name: str):
        try:
            # 确保输出目录存在
            os.makedirs(self.OUTPUT_DIR, exist_ok=True)

            # 生成输出文件名
            today = datetime.now().strftime("%Y%m%d")
            output_filename = f"{today}-大段覆盖包含明细ACL检查.xlsx"
            output_path = os.path.join(self.OUTPUT_DIR, output_filename)

            # 使用原始 acl_dup.py 的处理逻辑
            per_sheet_stats = process_file(self.INPUT_PATH, output_path)

            if sheet_name in per_sheet_stats:
                stats = per_sheet_stats[sheet_name]
                self.add_result(
                    Level.OK,
                    f"Sheet {sheet_name} 处理完成：可回收组数={stats['groups']}，"
                    f"保留规则={stats['keep']}，可回收规则={stats['recycle']}"
                )
            else:
                self.add_result(Level.WARN, f"Sheet {sheet_name} 无有效 ACL 规则")

        except Exception as EXCEPTION:
            self.add_result(Level.ERROR, f"处理 Sheet {sheet_name} 失败: {EXCEPTION}")

    # 重写run方法：处理所有Sheet并生成最终报告
    def run(self) -> None:
        task_items = list(self.items())
        if not task_items:
            self.add_result(Level.ERROR, "未找到可处理的 Sheet")
            return

        # 使用父类的进度条处理
        from tqdm import tqdm
        from .TaskBase import BAR_FORMAT, SHOW_PROGRESS

        progress = tqdm(
            total=len(task_items),
            desc=self.NAME,
            position=0,
            leave=True,
            dynamic_ncols=True,
            bar_format=BAR_FORMAT,
        ) if SHOW_PROGRESS else None

        try:
            # 收集所有 Sheet 的统计信息
            all_sheet_stats = {}

            for sheet_name in task_items:
                try:
                    # 处理单个 Sheet（这会调用 run_single）
                    self.run_single(sheet_name)

                    # 从结果中提取统计信息
                    for result in self.RESULTS:
                        if (result.meta and
                            result.meta.get("sheet_name") == sheet_name and
                            "stats" in result.meta):
                            all_sheet_stats[sheet_name] = result.meta["stats"]
                            break

                except Exception as EXCEPTION:
                    self.add_result(Level.ERROR, f"Sheet {sheet_name} 运行异常: {EXCEPTION!r}")

                if progress:
                    progress.update(1)

            # 生成最终汇总报告
            if all_sheet_stats:
                self._generate_final_report(all_sheet_stats)

        finally:
            if progress:
                progress.close()

    # 生成最终汇总报告：创建Report工作表并保存Excel文件
    def _generate_final_report(self, all_sheet_stats: Dict[str, Dict[str, int]]) -> None:
        try:
            # 生成输出文件名
            today = datetime.now().strftime("%Y%m%d")
            output_filename = f"{today}-大段覆盖包含明细ACL检查.xlsx"
            output_path = os.path.join(self.OUTPUT_DIR, output_filename)

            # 创建最终报告 - 使用已经处理过的Excel文件
            outputWorkbook = load_workbook(output_path)

            # 删除 Report 工作表（如果存在）
            if "Report" in outputWorkbook.sheetnames:
                del outputWorkbook["Report"]

            outputWorkbook.save(output_path)

            # 汇总统计（通过Config.yaml的enable_summary_output开关控制，仅输出到LOG文件）
            try:
                from .TaskBase import require_keys
                require_keys(CONFIG, ["ACLDupCheckTask"], "root")
                enable_summary = CONFIG["ACLDupCheckTask"].get("enable_summary_output", False)
            except Exception:
                enable_summary = False
            
            if enable_summary:
                total_groups = sum(stats.get("groups", 0) for stats in all_sheet_stats.values())
                total_keep = sum(stats.get("keep", 0) for stats in all_sheet_stats.values())
                total_recycle = sum(stats.get("recycle", 0) for stats in all_sheet_stats.values())
                total_in_groups_all = sum(stats.get("total_in_groups", 0) for stats in all_sheet_stats.values())
                self.add_result(
                    Level.OK,
                    f"ACL覆盖检查全部完成：处理{len(all_sheet_stats)}个站点，发现{total_groups}个覆盖组，"
                    f"建议保留{total_keep}条规则（红色标记），可回收{total_recycle}条重复规则，"
                    f"共{total_in_groups_all}条规则参与覆盖分析"
                )

        except Exception as EXCEPTION:
            self.add_result(Level.ERROR, f"生成最终报告失败: {EXCEPTION}")