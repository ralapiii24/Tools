# 镜像飞塔防火墙巡检任务

# 导入标准库
import re
import unicodedata
from typing import Optional, Tuple

# 导入第三方库
import paramiko

# 导入本地应用
from .TaskBase import BaseTask, Level, CONFIG, create_ssh_connection, ssh_exec, grade_percent, require_keys, decrypt_password

# FortiGate 专用正则：仅本任务使用
RE_CPU_IDLE = re.compile(r"CPU\s+states:\s*.*?(\d+)\s*%\s*idle", re.IGNORECASE)
RE_MEM_USED = re.compile(r"Memory:\s*.*?used\s*\(\s*([\d\.]+)\s*%\s*\)", re.IGNORECASE)
RE_UPTIME_LINE = re.compile(r"(?im)^[^\S\r\n]*Uptime[:：]?\s*(.+)$")
RE_UPTIME_UNITS = re.compile(r"(\d+)\s*d(?:ay)?s?.{0,20}?(\d+)\s*h(?:our)?s?.{0,20}?(\d+)\s*m(?:in(?:ute)?)?s?",
                             re.IGNORECASE)

# 镜像飞塔防火墙巡检任务类：通过SSH连接检查FortiGate防火墙的性能指标和运行状态
class MirrorFortiGateTask(BaseTask):
    
    # 初始化镜像飞塔防火墙巡检任务：设置SSH连接参数和性能阈值
    def __init__(self):
        super().__init__("镜像飞塔防火墙巡检")
        
        # 验证MirrorFortiGateTask专用配置
        require_keys(CONFIG, ["MirrorFortiGateTask"], "root")
        require_keys(CONFIG["MirrorFortiGateTask"], ["username", "password", "port", "hosts", "thresholds"], "MirrorFortiGateTask")
        
        FORTI_CONFIG = CONFIG["MirrorFortiGateTask"]
        self.USERNAME = FORTI_CONFIG["username"]
        self.PASSWORD = decrypt_password(FORTI_CONFIG["password"])
        self.PORT: int = int(FORTI_CONFIG["port"])
        self.HOSTS: list[str] = FORTI_CONFIG["hosts"]

        # 加载性能阈值配置
        THRESHOLDS = FORTI_CONFIG["thresholds"]
        self.DISK_WARN = int(THRESHOLDS["disk_percent"]["warn"])
        self.DISK_CRIT = int(THRESHOLDS["disk_percent"]["crit"])
        self.CPU_WARN = int(THRESHOLDS["cpu_percent"]["warn"])
        self.CPU_CRIT = int(THRESHOLDS["cpu_percent"]["crit"])
        self.MEM_WARN = int(THRESHOLDS["mem_percent"]["warn"])
        self.MEM_CRIT = int(THRESHOLDS["mem_percent"]["crit"])
        self.MIN_UPTIME_DAYS = int(THRESHOLDS["uptime_days"]["min"])

    # 返回要巡检的FortiGate主机列表
    def items(self):
        return self.HOSTS

    # 解析FortiGate磁盘使用率：从命令输出中提取磁盘使用百分比
    @staticmethod
    def PARSE_DISK_PERCENT(OUTPUT_TEXT: str) -> Optional[float]:
        for LINE in OUTPUT_TEXT.splitlines():
            if 'HD logging space usage for vdom "root"' in LINE:
                MATCH = re.search(r'(\d+)MB(?:\(\d+MiB\))?\s*/\s*(\d+)MB', LINE)
                if MATCH:
                    USED_MB, TOTAL_MB = int(MATCH.group(1)), int(MATCH.group(2))
                    return round(USED_MB / TOTAL_MB * 100, 2) if TOTAL_MB > 0 else None
        return None

    # 处理FortiGate分页命令输出：自动处理--More--分页，获取完整命令结果
    @staticmethod
    # 处理 FortiGate 开启分页（--More--）的命令输出：通过 invoke_shell 发送命令，遇到 '--More--' 自动回空格取下一页，直到回到提示符或超时
    def SSH_EXEC_PAGED(SSH: paramiko.SSHClient, CMD: str, TIMEOUT: int = 20) -> str:
        CHANNEL = SSH.invoke_shell(width=200, height=5000)
        CHANNEL.settimeout(1.0)

        # 清空欢迎/提示符残留
        try:
            import time
            time.sleep(0.2)
            while CHANNEL.recv_ready():
                _ = CHANNEL.recv(65535)
        except Exception:
            pass

        CHANNEL.send(CMD + "\n")
        # 用于累积接收到的命令输出数据
        OUTPUT_BUFFER = ""
        START_TIME = time.time()
        IDLE_START_TIME = None

        # 简单提示符判断：行尾出现 '#' 或 '>'，且屏幕上没有 '--More--'
        PROMPT_REGEX = re.compile(r"(?m)[#>]\s*$")

        while time.time() - START_TIME < TIMEOUT:
            GOT_DATA = False
            try:
                while CHANNEL.recv_ready():
                    DATA_PIECE = CHANNEL.recv(65535).decode("utf-8", "ignore")
                    if not DATA_PIECE:
                        break
                    OUTPUT_BUFFER += DATA_PIECE
                    GOT_DATA = True
            except Exception:
                pass

            if GOT_DATA:
                IDLE_START_TIME = None
            else:
                if IDLE_START_TIME is None:
                    IDLE_START_TIME = time.time()
                elif time.time() - IDLE_START_TIME > 2.0:
                    break

            # 处理分页 --More--
            if "--More--" in OUTPUT_BUFFER:
                CHANNEL.send(" ")
                continue

            # 若看到提示符且未出现分页，认为结束
            if PROMPT_REGEX.search(OUTPUT_BUFFER) and "--More--" not in OUTPUT_BUFFER:
                break

            time.sleep(0.05)

        # 去掉可能的控制字符与 --More-- 残留
        OUTPUT_BUFFER = re.sub(r"\x1b\[[0-9;?]*[ -/]*[@-~]", "", OUTPUT_BUFFER)
        OUTPUT_BUFFER = OUTPUT_BUFFER.replace("\r", "")
        OUTPUT_BUFFER = OUTPUT_BUFFER.replace("--More--", "")
        return OUTPUT_BUFFER

    # 解析FortiGate性能状态：从命令输出中提取CPU、内存使用率和运行时间
    @staticmethod
    def PARSE_PERF_STATUS(OUTPUT_TEXT: str) -> Tuple[Optional[float], Optional[float], Optional[float]]:
        # 解析FortiGate性能状态
        CLEANED_TEXT = (OUTPUT_TEXT or "").replace("\r", "")
        CLEANED_TEXT = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", " ", CLEANED_TEXT)
        CLEANED_TEXT = re.sub(r"\x1b\[[0-9;?]*[ -/]*[@-~]", "", CLEANED_TEXT)

        CPU_USED_PCT: Optional[float] = None
        MEM_USED_PCT: Optional[float] = None
        UPTIME_DAYS: Optional[float] = None

        MATCH_IDLE = RE_CPU_IDLE.search(CLEANED_TEXT)
        if MATCH_IDLE:
            try:
                IDLE_PCT = float(MATCH_IDLE.group(1))
                CPU_USED_PCT = round(100.0 - IDLE_PCT, 2)
            except Exception:
                pass

        MATCH_MEM = RE_MEM_USED.search(CLEANED_TEXT)
        if MATCH_MEM:
            try:
                MEM_USED_PCT = float(MATCH_MEM.group(1))
            except Exception:
                pass

        # 解析uptime文本
        def _parse_uptime_text(TEXT: str) -> Optional[float]:
            # 解析uptime文本，返回天数
            if not TEXT:
                return None
            
            # 标准化字符串
            try:
                NORMALIZED_TEXT = unicodedata.normalize("NFKC", TEXT)
            except Exception:
                NORMALIZED_TEXT = TEXT
            NORMALIZED_TEXT = NORMALIZED_TEXT.replace("\xa0", " ").strip()
            NORMALIZED_TEXT = re.sub(r"\s+", " ", NORMALIZED_TEXT)
            
            # 尝试匹配标准格式
            MATCH_DAYS = re.search(r"(\d+)\s*d(?:ay)?s?", NORMALIZED_TEXT, re.IGNORECASE)
            MATCH_HOURS = re.search(r"(\d+)\s*h(?:our)?s?", NORMALIZED_TEXT, re.IGNORECASE)
            MATCH_MINUTES = re.search(r"(\d+)\s*m(?:in(?:ute)?)?s?", NORMALIZED_TEXT, re.IGNORECASE)
            
            if MATCH_DAYS and MATCH_HOURS and MATCH_MINUTES:
                DAYS_VAL = int(MATCH_DAYS.group(1))
                HOURS_VAL = int(MATCH_HOURS.group(1))
                MINUTES_VAL = int(MATCH_MINUTES.group(1))
            else:
                # 尝试数字序列
                NUMBERS = re.findall(r"(\d+)", NORMALIZED_TEXT)
                if len(NUMBERS) >= 3:
                    DAYS_VAL, HOURS_VAL, MINUTES_VAL = int(NUMBERS[0]), int(NUMBERS[1]), int(NUMBERS[2])
                else:
                    return None
            
            try:
                return DAYS_VAL + HOURS_VAL / 24.0 + MINUTES_VAL / (24.0 * 60.0)
            except Exception:
                return None

        # 查找uptime文本
        TAIL_TEXT = None
        MATCH_LINE = RE_UPTIME_LINE.search(CLEANED_TEXT)
        if MATCH_LINE:
            TAIL_TEXT = MATCH_LINE.group(1)
        else:
            MATCH_POS = re.search(r"uptime\s*[:：]?", CLEANED_TEXT, re.IGNORECASE)
            if MATCH_POS:
                SEGMENT = CLEANED_TEXT[MATCH_POS.end(): MATCH_POS.end() + 200]
                SEGMENT = SEGMENT.split("\n")[0]
                SEGMENT = re.split(r"[#>]\s*$", SEGMENT)[0]
                TAIL_TEXT = SEGMENT
        
        # 解析uptime
        if UPTIME_DAYS is None:
            # 先尝试正则匹配
            ALL_UNITS = list(RE_UPTIME_UNITS.finditer(CLEANED_TEXT))
            if ALL_UNITS:
                DAYS_VAL, HOURS_VAL, MINUTES_VAL = map(int, ALL_UNITS[-1].groups())
                try:
                    UPTIME_DAYS = DAYS_VAL + HOURS_VAL / 24.0 + MINUTES_VAL / (24.0 * 60.0)
                except Exception:
                    UPTIME_DAYS = None
            
            # 如果正则失败，尝试文本解析
            if UPTIME_DAYS is None and TAIL_TEXT:
                UPTIME_DAYS = _parse_uptime_text(TAIL_TEXT)

        return CPU_USED_PCT, MEM_USED_PCT, UPTIME_DAYS

    # 获取FortiGate主机名：通过SSH执行命令获取设备主机名
    @staticmethod
    def GET_HOSTNAME(SSH: paramiko.SSHClient) -> str:
        # 获取FortiGate主机名
        EXIT_CODE, STDOUT_TEXT, _ = ssh_exec(SSH, "get system status | grep Hostname", label="get hostname")
        MATCH = re.search(r"Hostname:\s*(\S+)", STDOUT_TEXT)
        return MATCH.group(1) if MATCH else "Unknown"

    # 性能指标评级：根据阈值判断性能指标等级
    @staticmethod
    def GRADE(VALUE: Optional[float], WARN: int, CRIT: int) -> Level:
        # 性能指标评级
        return grade_percent(VALUE, WARN, CRIT)
    
    # 检查性能指标：统一的性能指标检查逻辑
    def _check_performance_metric(self, HOSTNAME: str, HOST: str, METRIC_NAME: str, 
                                 VALUE: Optional[float], WARN_THRESHOLD: int, CRIT_THRESHOLD: int) -> None:
        # 检查性能指标并添加结果
        # 从主机名中提取站点名（如HX00-Sniffer-FortiGate-1801F -> HX00）
        SITE_NAME = HOSTNAME.split('-')[0] if '-' in HOSTNAME else HOSTNAME
        LEVEL = self.GRADE(VALUE, WARN_THRESHOLD, CRIT_THRESHOLD)
        if LEVEL == Level.ERROR:
            self.add_result(LEVEL, f"站点{SITE_NAME}镜像飞塔防火墙{METRIC_NAME} 使用率解析失败")
        else:
            self.add_result(LEVEL, f"站点{SITE_NAME}镜像飞塔防火墙{METRIC_NAME} {VALUE}%（预警WARN:{WARN_THRESHOLD}/严重CRITICAL:{CRIT_THRESHOLD}%）")

    # 执行单个FortiGate设备的巡检：检查CPU、内存、磁盘和运行时间
    def run_single(self, HOST: str) -> None:
        SSH: Optional[paramiko.SSHClient] = None
        try:
            SSH = create_ssh_connection(HOST, self.PORT, self.USERNAME, self.PASSWORD)

            HOSTNAME = self.GET_HOSTNAME(SSH)

            # 1) 日志盘
            _, DISK_OUTPUT, _ = ssh_exec(SSH, "diagnose sys logdisk usage", label="logdisk usage")
            DISK_PERCENT = self.PARSE_DISK_PERCENT(DISK_OUTPUT)
            # 从主机名中提取站点名（如HX00-Sniffer-FortiGate-1801F -> HX00）
            SITE_NAME = HOSTNAME.split('-')[0] if '-' in HOSTNAME else HOSTNAME
            if DISK_PERCENT is None:
                self.add_result(Level.ERROR, f"站点{SITE_NAME}镜像飞塔防火墙无日志盘信息")
            elif DISK_PERCENT >= self.DISK_CRIT:
                self.add_result(Level.CRIT, f"站点{SITE_NAME}镜像飞塔防火墙日志盘 {DISK_PERCENT}%（预警WARN:{self.DISK_WARN}/严重CRITICAL:{self.DISK_CRIT}%）")
            elif DISK_PERCENT >= self.DISK_WARN:
                self.add_result(Level.WARN, f"站点{SITE_NAME}镜像飞塔防火墙日志盘 {DISK_PERCENT}%（预警WARN:{self.DISK_WARN}/严重CRITICAL:{self.DISK_CRIT}%）")
            else:
                self.add_result(Level.OK, f"站点{SITE_NAME}镜像飞塔防火墙日志盘 {DISK_PERCENT}%（预警WARN:{self.DISK_WARN}/严重CRITICAL:{self.DISK_CRIT}%）")

            # 2) 性能状态（CPU/内存/Uptime）
            # 先用普通 exec_command 拿一把；若检测到分页或缺少 Uptime，再用交互式分页读取
            _, PERF_OUTPUT, _ = ssh_exec(SSH, "get system performance status", label="perf status")
            NEED_PAGED = ("--More--" in PERF_OUTPUT) or ("Uptime" not in PERF_OUTPUT and "uptime" not in PERF_OUTPUT)
            if NEED_PAGED:
                PERF_OUTPUT = self.SSH_EXEC_PAGED(SSH, "get system performance status")

            CPU_USED, MEM_USED, UPTIME_DAYS = self.PARSE_PERF_STATUS(PERF_OUTPUT)

            # 调试：打印 perf 回显与初次解析结果（仅目标主机）
            # 兜底：若 Uptime 仍为空，再从 get system status 抓一次（同样考虑分页）
            if UPTIME_DAYS is None:
                # 尽量直接读完整的 system status（不要带 grep，部分设备不支持管道/同样会分页）
                SYS_STATUS_FULL = self.SSH_EXEC_PAGED(SSH, "get system status")
                _, _, UPTIME_TRY = self.PARSE_PERF_STATUS(SYS_STATUS_FULL)
                if UPTIME_TRY is not None:
                    UPTIME_DAYS = UPTIME_TRY

            # 检查性能指标
            self._check_performance_metric(HOSTNAME, HOST, "CPU", CPU_USED, self.CPU_WARN, self.CPU_CRIT)
            self._check_performance_metric(HOSTNAME, HOST, "内存", MEM_USED, self.MEM_WARN, self.MEM_CRIT)
            
            # 检查运行时间
            if UPTIME_DAYS is None:
                self.add_result(Level.ERROR, f"站点{SITE_NAME}镜像飞塔防火墙启动时间 解析失败")
            elif UPTIME_DAYS < self.MIN_UPTIME_DAYS:
                self.add_result(Level.CRIT, f"站点{SITE_NAME}镜像飞塔防火墙启动时间 仅 {UPTIME_DAYS:.2f} 天（< {self.MIN_UPTIME_DAYS} 天）")
            else:
                self.add_result(Level.OK, f"站点{SITE_NAME}镜像飞塔防火墙启动时间 {UPTIME_DAYS:.2f} 天")

        except Exception as ERROR:
            self.add_result(Level.ERROR, f"{HOST} 巡检失败: {ERROR}")
        finally:
            try:
                if SSH:
                    SSH.close()
            except Exception:
                pass
