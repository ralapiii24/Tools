# Linux服务器巡检基类（V11从Base.py拆分）
#
# 技术栈:Python, SSH, Paramiko, 正则表达式
# 目标:提供 Linux 服务器巡检的通用功能，供多个服务器巡检任务共用
#
# 适用:ESLogstashTask / ESBaseTask / ESFlowTask
#
# 通用检查:
# 根分区占用:df -h 解析 / 的 used%，与 servers.thresholds.disk_percent 比较（默认 WARN≥50，CRIT≥80）
# 内存占用:free -m 解析 used/total 推算占用率，各类型有不同阈值:
# - LOGSTASH:WARN≥50，CRIT≥80
# - ES/FLOW:WARN≥80，CRIT≥90

# 导入标准库
from typing import Dict, Optional, Tuple

# 导入第三方库
import paramiko

# 导入本地应用
from .TaskBase import BaseTask, Level, CONFIG, decrypt_password, create_ssh_connection, ssh_exec, grade_percent, require_keys

# 验证ESServer配置：检查Linux服务器任务所需的配置项
require_keys(CONFIG, ["ESServer"], "root")
require_keys(CONFIG["ESServer"], ["port", "thresholds", "ESLogstashTask", "ESBaseTask", "ESFlowTask"], "ESServer")

# Linux 服务器巡检基类（df -h + free -m）
# Linux服务器巡检基类：专门用于Linux服务器的巡检任务，提供内存和磁盘检查功能
class BaseLinuxServerTask(BaseTask):
    # 初始化Linux服务器巡检任务：设置SSH连接参数和性能阈值
    def __init__(self, name: str, section_key: str, mem_warn: int, mem_crit: int):
        super().__init__(name)
        servers_cfg = CONFIG["ESServer"]
        section_cfg = servers_cfg[section_key]
        self.USERNAME = section_cfg["username"]
        self.PASSWORD = decrypt_password(section_cfg["password"])
        self.PORT: int = int(servers_cfg["port"])
        self.HOSTS_MAP: Dict[str, str] = section_cfg["hosts"]

        disk_thr = servers_cfg["thresholds"]["disk_percent"]
        self.DISK_WARN = int(disk_thr["warn"])
        self.DISK_CRIT = int(disk_thr["crit"])

        self.MEM_WARN = int(mem_warn)
        self.MEM_CRIT = int(mem_crit)
        
        # 根据任务类型设置不同的描述
        if section_key == "ESLogstashTask":
            self.SERVICE_TYPE = "日志系统LOGSTASH"
        elif section_key == "ESBaseTask":
            self.SERVICE_TYPE = "日志系统ES"
        elif section_key == "ESFlowTask":
            self.SERVICE_TYPE = "流量分析系统FLOW"
        else:
            self.SERVICE_TYPE = "服务器"

    # 返回要巡检的主机列表：从配置中获取主机名和IP地址映射
    def items(self):
        return list(self.HOSTS_MAP.items())

    # 创建SSH连接：建立到指定主机的SSH连接
    @staticmethod
    def _ssh(ip: str, port: int, user: str, pwd: str) -> paramiko.SSHClient:
        return create_ssh_connection(ip, port, user, pwd)

    # 解析free -m命令输出：从内存使用信息中提取总内存、已用内存和使用率
    @staticmethod
    def _parse_free_m(output_text: str) -> Tuple[Optional[int], Optional[int], Optional[float]]:
        for line in output_text.splitlines():
            if line.lower().startswith("mem:"):
                parts = line.split()
                if len(parts) >= 3:
                    total_mb, used_mb = int(parts[1]), int(parts[2])
                    pct = round(used_mb / total_mb * 100, 2) if total_mb > 0 else None
                    return total_mb, used_mb, pct
        return None, None, None

    # 解析df -h命令输出：从磁盘使用信息中提取根分区使用率
    @staticmethod
    def _parse_df_root(output_text: str) -> Tuple[Optional[int], Optional[str]]:
        lines = output_text.strip().splitlines()
        if len(lines) <= 1:
            return None, None
        for line in lines[1:]:
            parts = line.split()
            if len(parts) < 2:
                continue
            mount_point = parts[-1]
            used_percent_field = parts[-2]
            try:
                used_percent_val = int(used_percent_field.strip().rstrip('%'))
            except Exception:
                continue
            if mount_point == '/':
                return used_percent_val, line
        return None, None

    # 执行单个服务器的巡检：检查内存和磁盘使用情况
    def run_single(self, item: Tuple[str, str]) -> None:
        server_name, ip_addr = item
        ssh: Optional[paramiko.SSHClient] = None
        try:
            ssh = self._ssh(ip_addr, self.PORT, self.USERNAME, self.PASSWORD)

            _, df_stdout, _ = ssh_exec(ssh, "df -h", label="df -h")
            disk_used_pct, raw_df_line = self._parse_df_root(df_stdout)

            _, free_stdout, _ = ssh_exec(ssh, "free -m", label="free -m")
            total_mb, used_mb, mem_used_pct = self._parse_free_m(free_stdout)
        except Exception as exc:
            self.add_result(Level.ERROR, f"站点{server_name}{self.SERVICE_TYPE} {ip_addr} 巡检失败：{exc}")
            return
        finally:
            try:
                if ssh:
                    ssh.close()
            except Exception:
                pass

        if mem_used_pct is None:
            self.add_result(Level.ERROR, f"站点{server_name}{self.SERVICE_TYPE} {ip_addr} 内存信息解析失败")
        else:
            mem_level = grade_percent(mem_used_pct, self.MEM_WARN, self.MEM_CRIT)
            self.add_result(mem_level,
                            f"站点{server_name}{self.SERVICE_TYPE} {ip_addr} 内存{mem_used_pct}%（预警WARN:{self.MEM_WARN}/严重CRITICAL:{self.MEM_CRIT}%）")

        if disk_used_pct is None:
            self.add_result(Level.ERROR, f"站点{server_name}{self.SERVICE_TYPE} {ip_addr} 未找到根分区 / 磁盘信息")
        else:
            disk_level = grade_percent(disk_used_pct, self.DISK_WARN, self.DISK_CRIT)
            message = f"站点{server_name}{self.SERVICE_TYPE} {ip_addr} 磁盘{disk_used_pct}%（预警WARN:{self.DISK_WARN}/严重CRITICAL:{self.DISK_CRIT}%）"
            if disk_level != Level.OK and raw_df_line:
                message += f"；原行: {raw_df_line}"
            self.add_result(disk_level, message)

