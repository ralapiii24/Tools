# ES 服务器巡检任务（三类 Linux 主机健康）
#
# 技术栈:Python, SSH, Paramiko, 正则表达式
# 目标:检查 Elasticsearch 服务器的健康状态
# 需调用LinuxServerBase.py（Linux服务器巡检基类，V11从Base.py拆分）
#
# 适用:ESLogstashTask / ESBaseTask / ESFlowTask
#
# 通用检查:
# 根分区占用:df -h 解析 / 的 used%，与 servers.thresholds.disk_percent 比较（默认 WARN≥50，CRIT≥80）
# 内存占用:free -m 解析 used/total 推算占用率，各类型有不同阈值:
# - LOGSTASH:WARN≥50，CRIT≥80
# - ES/FLOW:WARN≥80，CRIT≥90

# 导入标准库
# (无标准库依赖)

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
from .LinuxServerBase import BaseLinuxServerTask
from .TaskBase import CONFIG, require_keys

# Elasticsearch服务器巡检任务类：专门用于Elasticsearch服务器的巡检，检查内存和磁盘使用情况
class ESBaseTask(BaseLinuxServerTask):
    """Elasticsearch服务器巡检任务
    

    专门用于Elasticsearch服务器的巡检，检查内存和磁盘使用情况
    """
    # 初始化Elasticsearch服务器巡检任务：设置内存阈值和巡检参数
    def __init__(self):
        # 验证ESBaseTask专用配置
        require_keys(CONFIG, ["ESServer"], "root")
        require_keys(CONFIG["ESServer"], ["thresholds"], "ESServer")
        require_keys(CONFIG["ESServer"]["thresholds"], ["mem_percent"], "ESServer.thresholds")
        require_keys(
            CONFIG["ESServer"]["thresholds"]["mem_percent"],
            ["ESBaseTask"],
            "ESServer.thresholds.mem_percent"
        )
        

        # 从配置文件读取ESBaseTask的内存阈值配置
        MEM_THRESHOLDS = CONFIG["ESServer"]["thresholds"]["mem_percent"]["ESBaseTask"]
        super().__init__("ES服务器巡检", "ESBaseTask", 

                        MEM_THRESHOLDS["warn"], MEM_THRESHOLDS["crit"])
