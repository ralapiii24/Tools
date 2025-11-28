# LOGSTASH 服务器巡检任务
#
# 技术栈:Python, SSH, Paramiko, 正则表达式
# 目标:检查 Logstash 服务器的健康状态
# 继承自 LinuxServerBase，包含 ESBaseTask 的所有通用检查
#
# 通用检查（来自 LinuxServerBase）:
# 根分区占用:df -h 解析 / 的 used%，与 servers.thresholds.disk_percent 比较（默认 WARN≥50，CRIT≥80）
# 内存占用:free -m 解析 used/total 推算占用率，LOGSTASH 阈值: WARN≥50，CRIT≥80

# 导入标准库
# (无标准库依赖)

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
from .LinuxServerBase import BaseLinuxServerTask
from .TaskBase import CONFIG, require_keys

# Logstash服务器巡检任务类：专门用于Logstash服务器的巡检，检查内存和磁盘使用情况
class ESLogstashTask(BaseLinuxServerTask):
    """Logstash服务器巡检任务
    

    专门用于Logstash服务器的巡检，检查内存和磁盘使用情况
    继承自LinuxServerBase，包含所有通用检查功能
    """
    # 初始化Logstash服务器巡检任务：设置内存阈值和巡检参数
    def __init__(self):
        # 验证ESLogstashTask专用配置
        require_keys(CONFIG, ["ESServer"], "root")
        require_keys(CONFIG["ESServer"], ["thresholds"], "ESServer")
        require_keys(CONFIG["ESServer"]["thresholds"], ["mem_percent"], "ESServer.thresholds")
        require_keys(
            CONFIG["ESServer"]["thresholds"]["mem_percent"],
            ["ESLogstashTask"],
            "ESServer.thresholds.mem_percent"
        )
        

        # 从配置文件读取ESLogstashTask的内存阈值配置
        MEM_THRESHOLDS = CONFIG["ESServer"]["thresholds"]["mem_percent"]["ESLogstashTask"]
        super().__init__("LOGSTASH服务器巡检", "ESLogstashTask", 

                        MEM_THRESHOLDS["warn"], MEM_THRESHOLDS["crit"])
