# 服务检查任务
#
# 技术栈:Python, SSH, 正则表达式, 系统服务检查
# 目标:检查NTP Chronyd和TACACS+ tac_plus服务状态，确保关键服务正常运行
#
# 处理逻辑:SSH连接 → 服务状态检查 → 进程检查 → 端口检查 → 状态解析 → 结果汇总
#
# Chronyd NTP服务检查:
# - systemctl status chronyd: 检查服务运行状态和运行时间
# - chronyc tracking: 检查NTP同步状态（Reference ID、Stratum、Last offset）
# - ps -ef | grep chronyd: 检查chronyd进程
# - ss -ulpn | grep chronyd: 检查UDP 123端口监听
# - 检查要点: Reference ID不为0.0.0.0，Stratum不为16，Last offset小于1秒
#
# TACACS+服务检查:
# - systemctl status tac_plus: 检查服务运行状态和运行时间
# - ps -ef | grep tac_plus: 检查tac_plus进程
# - ss -tulnp | grep 49: 检查TCP 49端口监听
# - 检查要点: 服务运行正常，进程存在，端口监听正常
#
# 输出:LOG/ServiceCheckTask/YYYYMMDD-服务检查任务.log（详细检查日志）
# 配置说明:复用ESLogstashTask配置，使用相同的SSH认证和服务器列表，避免配置重复
# 特点:支持多服务器并行检查，确保关键服务正常运行

# 导入标准库
import re

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
from .TaskBase import BaseTask, Level, CONFIG, ssh_exec, create_ssh_connection, decrypt_password

# 服务检查任务类：检查NTP Chronyd和TACACS+ tac_plus服务状态，确保关键服务正常运行
class ServiceCheckTask(BaseTask):
    """服务检查任务
    

    检查NTP Chronyd和TACACS+ tac_plus服务状态，确保关键服务正常运行
    """
    

    # 初始化服务检查任务：设置SSH连接参数和主机列表
    def __init__(self):
        super().__init__("服务检查任务")
        # 复用ESLogstashTask的配置
        self.SERVICE_CONFIG = CONFIG["ESServer"]["ESLogstashTask"]
        self.USERNAME = self.SERVICE_CONFIG["username"]
        self.PASSWORD = decrypt_password(self.SERVICE_CONFIG["password"])
        self.HOSTS = self.SERVICE_CONFIG["hosts"]
        

    # 返回所有需要检查的主机列表：从配置中获取主机名列表
    def items(self):
        """返回所有需要检查的主机列表
        

        Returns:
            list[str]: 主机名列表
        """
        return list(self.HOSTS.keys())
    

    # 检查单个主机的服务状态：检查Chronyd和TACACS+服务
    def run_single(self, hostname: str) -> None:
        """检查单个主机的服务状态
        

        检查Chronyd和TACACS+服务
        

        Args:
            hostname: 主机名
        """
        IP = self.HOSTS[hostname]
        

        try:
            # 检查Chronyd服务
            self._check_chronyd_service(hostname, IP)
            

            # 检查TACACS+服务
            self._check_tacplus_service(hostname, IP)
            

        except Exception as error:
            self.add_result(Level.ERROR, f"{hostname}({IP}) 服务检查异常: {error!r}")
    

    # 检查Chronyd NTP服务状态：检查服务运行状态、NTP同步状态、进程和端口监听
    def _check_chronyd_service(self, hostname: str, IP: str) -> None:
        """检查Chronyd NTP服务状态
        

        检查服务运行状态、NTP同步状态、进程和端口监听
        

        Args:
            hostname: 主机名
            IP: IP地址
        """
        SSH = None
        try:
            # 建立SSH连接
            SSH = create_ssh_connection(IP, 22, self.USERNAME, self.PASSWORD)
            

            # 检查chronyd服务状态
            STATUS_CMD = "systemctl status chronyd"
            EXIT_CODE, STATUS_OUTPUT, STDERR = ssh_exec(SSH, STATUS_CMD)
            

            # 解析服务状态
            if "Active: active (running)" in STATUS_OUTPUT:
                self.add_result(Level.OK, f"{hostname} Chronyd服务运行正常")
                

                # 检查运行时间
                SINCE_MATCH = re.search(r'since (.+?); (\d+) days? ago', STATUS_OUTPUT)
                if SINCE_MATCH:
                    DAYS_RUNNING = int(SINCE_MATCH.group(2))
                    if DAYS_RUNNING >= 1:
                        self.add_result(Level.OK, f"{hostname} Chronyd已运行{DAYS_RUNNING}天，状态稳定")
                    else:
                        self.add_result(Level.WARN, f"{hostname} Chronyd运行时间较短({DAYS_RUNNING}天)")
            else:
                self.add_result(Level.CRIT, f"{hostname} Chronyd服务未运行")
                return
            

            # 检查chronyc tracking状态
            TRACKING_CMD = "chronyc tracking"
            EXIT_CODE, TRACKING_OUTPUT, STDERR = ssh_exec(SSH, TRACKING_CMD)
            

            # 解析tracking信息
            self._parse_chronyc_tracking(hostname, TRACKING_OUTPUT)
            

            # 检查chronyd进程
            PS_CMD = "ps -ef | grep chronyd"
            EXIT_CODE, PS_OUTPUT, STDERR = ssh_exec(SSH, PS_CMD)
            

            if "/usr/sbin/chronyd" in PS_OUTPUT:
                self.add_result(Level.OK, f"{hostname} Chronyd进程运行正常")
            else:
                self.add_result(Level.WARN, f"{hostname} Chronyd进程未找到")
            

            # 检查UDP 123端口
            SS_CMD = "ss -ulpn | grep chronyd"
            EXIT_CODE, SS_OUTPUT, STDERR = ssh_exec(SSH, SS_CMD)
            

            if ":123" in SS_OUTPUT and "chronyd" in SS_OUTPUT:
                self.add_result(Level.OK, f"{hostname} Chronyd UDP 123端口监听正常")
            else:
                self.add_result(Level.WARN, f"{hostname} Chronyd UDP 123端口未监听")
            

                

        except Exception as error:
            self.add_result(Level.ERROR, f"{hostname} Chronyd检查失败: {error!r}")
        finally:
            if SSH:
                try:
                    SSH.close()
                except Exception:
                    pass
    

    # 解析chronyc tracking输出：提取Reference ID、Stratum和Last offset信息
    def _parse_chronyc_tracking(self, hostname: str, TRACKING_OUTPUT: str) -> None:
        """解析chronyc tracking输出
        

        提取Reference ID、Stratum和Last offset信息
        

        Args:
            hostname: 主机名
            TRACKING_OUTPUT: chronyc tracking命令的输出
        """
        try:
            # 检查Reference ID
            REF_ID_MATCH = re.search(r'Reference ID\s*:\s*([0-9A-F]+)\s*\(([^)]+)\)', TRACKING_OUTPUT)
            if REF_ID_MATCH:
                REF_IP = REF_ID_MATCH.group(2)
                if REF_IP != "0.0.0.0":
                    self.add_result(Level.OK, f"{hostname} Chronyd参考服务器: {REF_IP}")
                else:
                    self.add_result(Level.CRIT, f"{hostname} Chronyd参考服务器为0.0.0.0，同步异常")
            else:
                self.add_result(Level.WARN, f"{hostname} 无法获取Chronyd参考服务器信息")
            

            # 检查Stratum
            STRATUM_MATCH = re.search(r'Stratum\s*:\s*(\d+)', TRACKING_OUTPUT)
            if STRATUM_MATCH:
                STRATUM = int(STRATUM_MATCH.group(1))
                if STRATUM != 16:
                    self.add_result(Level.OK, f"{hostname} Chronyd层级: {STRATUM}")
                else:
                    self.add_result(Level.CRIT, f"{hostname} Chronyd层级为16，同步异常")
            else:
                self.add_result(Level.WARN, f"{hostname} 无法获取Chronyd层级信息")
            

            # 检查Last offset
            OFFSET_MATCH = re.search(r'Last offset\s*:\s*([+-]?\d+\.?\d*)\s*seconds', TRACKING_OUTPUT)
            if OFFSET_MATCH:
                OFFSET = float(OFFSET_MATCH.group(1))
                if abs(OFFSET) < 1.0:
                    self.add_result(Level.OK, f"{hostname} Chronyd时间偏移: {OFFSET}秒")
                else:
                    self.add_result(Level.WARN, f"{hostname} Chronyd时间偏移过大: {OFFSET}秒")
            else:
                self.add_result(Level.WARN, f"{hostname} 无法获取Chronyd时间偏移信息")
                

        except Exception as error:
            self.add_result(Level.ERROR, f"{hostname} Chronyd tracking解析失败: {error!r}")
    

    # 检查TACACS+ tac_plus服务状态：检查服务运行状态、进程和TCP 49端口监听
    def _check_tacplus_service(self, hostname: str, IP: str) -> None:
        """检查TACACS+ tac_plus服务状态
        

        检查服务运行状态、进程和TCP 49端口监听
        

        Args:
            hostname: 主机名
            IP: IP地址
        """
        SSH = None
        try:
            # 建立SSH连接
            SSH = create_ssh_connection(IP, 22, self.USERNAME, self.PASSWORD)
            

            # 检查tac_plus服务状态
            STATUS_CMD = "systemctl status tac_plus"
            EXIT_CODE, STATUS_OUTPUT, STDERR = ssh_exec(SSH, STATUS_CMD)
            

            # 解析服务状态
            if "Active: active (running)" in STATUS_OUTPUT:
                self.add_result(Level.OK, f"{hostname} TACACS+服务运行正常")
                

                # 检查运行时间
                SINCE_MATCH = re.search(r'since (.+?); (\d+) days? ago', STATUS_OUTPUT)
                if SINCE_MATCH:
                    DAYS_RUNNING = int(SINCE_MATCH.group(2))
                    if DAYS_RUNNING >= 1:
                        self.add_result(Level.OK, f"{hostname} TACACS+已运行{DAYS_RUNNING}天，状态稳定")
                    else:
                        self.add_result(Level.WARN, f"{hostname} TACACS+运行时间较短({DAYS_RUNNING}天)")
            else:
                self.add_result(Level.CRIT, f"{hostname} TACACS+服务未运行")
                return
            

            # 检查tac_plus进程
            PS_CMD = "ps -ef | grep tac_plus"
            EXIT_CODE, PS_OUTPUT, STDERR = ssh_exec(SSH, PS_CMD)
            

            if "/usr/sbin/tac_plus" in PS_OUTPUT:
                self.add_result(Level.OK, f"{hostname} TACACS+进程运行正常")
            else:
                self.add_result(Level.WARN, f"{hostname} TACACS+进程未找到")
            

            # 检查TCP 49端口
            SS_CMD = "ss -tulnp | grep 49"
            EXIT_CODE, SS_OUTPUT, STDERR = ssh_exec(SSH, SS_CMD)
            

            if ":49" in SS_OUTPUT and "tac_plus" in SS_OUTPUT:
                self.add_result(Level.OK, f"{hostname} TACACS+ TCP 49端口监听正常")
            else:
                self.add_result(Level.WARN, f"{hostname} TACACS+ TCP 49端口未监听")
            

                

        except Exception as error:
            self.add_result(Level.ERROR, f"{hostname} TACACS+检查失败: {error!r}")
        finally:
            if SSH:
                try:
                    SSH.close()
                except Exception:
                    pass
