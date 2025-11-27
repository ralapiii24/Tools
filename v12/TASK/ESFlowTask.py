# FLOW 服务器巡检任务

# 导入标准库
import re
from typing import Tuple

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
from .LinuxServerBase import BaseLinuxServerTask
from .TaskBase import Level, CONFIG, ssh_exec, require_keys

# 字节单位转换函数：将字符串格式的存储大小转换为字节数
def to_bytes(SIZE_STR: str) -> int:
    if not SIZE_STR:
        return -1
    NORMALIZED = SIZE_STR.strip().lower()
    try:
        if NORMALIZED.endswith("gb"):
            return int(float(NORMALIZED[:-2]) * 1024 ** 3)
        if NORMALIZED.endswith("mb"):
            return int(float(NORMALIZED[:-2]) * 1024 ** 2)
        if NORMALIZED.endswith("kb"):
            return int(float(NORMALIZED[:-2]) * 1024)
        if NORMALIZED.endswith("b"):
            return int(float(NORMALIZED[:-1]))
        return int(float(NORMALIZED))
    except Exception:
        return -1

# FLOW服务器巡检任务类：专门用于FLOW服务器的巡检，包括容器、端口和ES索引检查
class ESFlowTask(BaseLinuxServerTask):
    # 初始化FLOW服务器巡检任务：设置内存阈值和专项检查配置
    def __init__(self):
        # 验证ESFlowTask专用配置
        require_keys(CONFIG, ["ESServer"], "root")
        require_keys(CONFIG["ESServer"], ["thresholds", "ESFlowTask_CustomParameters"], "ESServer")
        require_keys(CONFIG["ESServer"]["thresholds"], ["mem_percent"], "ESServer.thresholds")
        require_keys(CONFIG["ESServer"]["thresholds"]["mem_percent"], ["ESFlowTask"], "ESServer.thresholds.mem_percent")
        require_keys(CONFIG["ESServer"]["ESFlowTask_CustomParameters"], 
                    ["require_ports", "require_containers", "index_prefix", "index_size_limit_bytes", "segment_max_non_recent"], 
                    "ESServer.ESFlowTask_CustomParameters")
        
        # 从配置文件读取ESFlowTask的内存阈值配置
        MEM_THRESHOLDS = CONFIG["ESServer"]["thresholds"]["mem_percent"]["ESFlowTask"]
        super().__init__("FLOW服务器巡检", "ESFlowTask", 
                        MEM_THRESHOLDS["warn"], MEM_THRESHOLDS["crit"])
        # 从配置文件读取ESFlowTask的自定义参数（必须配置）
        self.FC = CONFIG["ESServer"]["ESFlowTask_CustomParameters"]

    # 执行单个FLOW服务器的专项巡检：检查端口、容器、ES索引和段信息
    def run_single(self, ITEM: Tuple[str, str]) -> None:
        super().run_single(ITEM)

        SERVER_NAME, IP_ADDR = ITEM
        try:
            SSH = self._ssh(IP_ADDR, self.PORT, self.USERNAME, self.PASSWORD)

            _, NETSTAT_STDOUT, _ = ssh_exec(SSH, "netstat -tulnp", label="ports")
            DOCKER_EC, DOCKER_STDOUT, DOCKER_STDERR = ssh_exec(SSH, 'docker ps --format "{{.Names}} {{.Status}}"',
                                                               label="docker ps")
            _, INDICES_STDOUT, _ = ssh_exec(SSH, "curl -s 'http://localhost:9200/_cat/indices?v'", label="es indices")
            _, SEGMENTS_STDOUT, _ = ssh_exec(SSH, "curl -s 'http://localhost:9200/_cat/segments?v'",
                                             label="es segments")
        except Exception as ERROR:
            self.add_result(Level.ERROR, f"站点{SERVER_NAME}流量分析系统FLOW {IP_ADDR} FLOW专项巡检失败：{ERROR}")
            return
        finally:
            try:
                if SSH:
                    SSH.close()
            except Exception:
                pass

        for REQUIRED_PORT in self.FC["require_ports"]:
            PATTERN = rf":{REQUIRED_PORT}\b.*LISTEN"
            if not re.search(PATTERN, NETSTAT_STDOUT):
                self.add_result(Level.CRIT, f"站点{SERVER_NAME}流量分析系统FLOW {IP_ADDR} 端口 {REQUIRED_PORT} 未监听")

        # 容器检查：优先使用 docker ps --format 的精确名称匹配；若 docker 失败，则按端口占用做兜底
        if DOCKER_EC == 0 and DOCKER_STDOUT.strip():
            RUNNING = set()
            for LINE in DOCKER_STDOUT.splitlines():
                PARTS = LINE.strip().split(None, 1)
                if not PARTS:
                    continue
                NAME = PARTS[0].strip()
                STATUS = PARTS[1] if len(PARTS) > 1 else ""
                if re.search(r"\bUp\b", STATUS):
                    RUNNING.add(NAME)
            for CONTAINER_NAME in self.FC["require_containers"]:
                if CONTAINER_NAME not in RUNNING:
                    self.add_result(Level.CRIT, f"站点{SERVER_NAME}流量分析系统FLOW {IP_ADDR} 容器 {CONTAINER_NAME} 未运行(或STATUS非Up)")
        else:
            # docker ps 执行失败：检查关键端口是否被占用，若有则视为通过，否则失败
            FALLBACK_PORTS = [5601, 9600, 9300, 9200, 4739, 2055, 6343]
            # 从 netstat 原始输出中过滤匹配到的行，兼容空格/制表符分隔
            FILTERED_LINES = []
            for LINE in NETSTAT_STDOUT.splitlines():
                for PORT in FALLBACK_PORTS:
                    if f":{PORT} " in LINE or f":{PORT}    " in LINE:
                        FILTERED_LINES.append(LINE.rstrip())
                        break
            if FILTERED_LINES:
                self.add_result(Level.OK,
                                f"站点{SERVER_NAME}流量分析系统FLOW {IP_ADDR} docker ps 失败，但端口命中如下：\n" + "\n".join(FILTERED_LINES))
            else:
                self.add_result(Level.ERROR, f"站点{SERVER_NAME}流量分析系统FLOW {IP_ADDR} docker ps 失败且关键端口未占用")
        INDEX_PREFIX = self.FC["index_prefix"]
        SIZE_LIMIT_BYTES = int(self.FC["index_size_limit_bytes"])
        DATE_REGEX = re.compile(re.escape(INDEX_PREFIX) + r"(\d{4}\.\d{2}\.\d{2})")

        INDEX_LINES = [LINE.strip() for LINE in INDICES_STDOUT.splitlines() if INDEX_PREFIX in LINE and LINE.strip()]
        DATE_SET = set()
        OVERSIZE_LIST = []
        for LINE in INDEX_LINES:
            MATCH = DATE_REGEX.search(LINE)
            if not MATCH:
                continue
            DATE_STR = MATCH.group(1)
            DATE_SET.add(DATE_STR)
            COLS = LINE.split()
            if not COLS:
                continue
            LAST_SIZE_FIELD = COLS[-1].lower()
            if to_bytes(LAST_SIZE_FIELD) > SIZE_LIMIT_BYTES:
                OVERSIZE_LIST.append(f"{DATE_STR} 大小 {LAST_SIZE_FIELD}")

        if len(DATE_SET) > 31:
            self.add_result(Level.WARN, f"站点{SERVER_NAME}流量分析系统FLOW {IP_ADDR} 索引日期数量 {len(DATE_SET)} 超过 31")
        if OVERSIZE_LIST:
            self.add_result(Level.WARN, f"站点{SERVER_NAME}流量分析系统FLOW {IP_ADDR} 索引大小超过1G: " + "；".join(OVERSIZE_LIST))

        SEGMENT_COUNTER: dict[str, int] = {}
        SEGMENT_DATES = set()
        for LINE in SEGMENTS_STDOUT.splitlines():
            STRIPED = LINE.strip()
            if not STRIPED or STRIPED.lower().startswith("index"):
                continue
            PARTS = STRIPED.split()
            if not PARTS:
                continue
            INDEX_NAME = PARTS[0]
            if not INDEX_NAME.startswith(INDEX_PREFIX):
                continue
            SEGMENT_COUNTER[INDEX_NAME] = SEGMENT_COUNTER.get(INDEX_NAME, 0) + 1
            MATCH = DATE_REGEX.search(INDEX_NAME)
            if MATCH:
                SEGMENT_DATES.add(MATCH.group(1))

        TODAY_STR = None
        YEST_STR = None
        if SEGMENT_DATES:
            try:
                from datetime import datetime, timedelta
                AVAILABLE_DATES = sorted([datetime.strptime(DATE_STR, "%Y.%m.%d") for DATE_STR in SEGMENT_DATES])
                MAX_DATE = AVAILABLE_DATES[-1]
                TODAY_STR = MAX_DATE.strftime("%Y.%m.%d")
                YEST_STR = (MAX_DATE - timedelta(days=1)).strftime("%Y.%m.%d")
            except Exception:
                pass

        LIMIT = int(self.FC["segment_max_non_recent"])
        OVERS_SEGMENT = []
        for INDEX_NAME, COUNT in SEGMENT_COUNTER.items():
            MATCH = DATE_REGEX.search(INDEX_NAME)
            if not MATCH:
                continue
            DATE_STR = MATCH.group(1)
            if TODAY_STR and YEST_STR:
                if DATE_STR in (TODAY_STR, YEST_STR):
                    continue
            else:
                from datetime import datetime, timedelta
                FALLBACK_TODAY = datetime.now().strftime("%Y.%m.%d")
                FALLBACK_YEST = (datetime.now() - timedelta(days=1)).strftime("%Y.%m.%d")
                if DATE_STR in (FALLBACK_TODAY, FALLBACK_YEST):
                    continue
            if COUNT > LIMIT:
                OVERS_SEGMENT.append(f"{INDEX_NAME} 行数 {COUNT}")

        if OVERS_SEGMENT:
            self.add_result(Level.WARN,
                            f"站点{SERVER_NAME}流量分析系统FLOW {IP_ADDR} segments 行数超过{LIMIT}: " + "；".join(OVERS_SEGMENT))
