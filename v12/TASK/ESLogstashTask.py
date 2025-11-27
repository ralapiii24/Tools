# LOGSTASH 服务器巡检任务

# 导入标准库
# (无标准库依赖)

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
from .LinuxServerBase import BaseLinuxServerTask
from .TaskBase import CONFIG, require_keys

# Logstash服务器巡检任务类：专门用于Logstash服务器的巡检，检查内存和磁盘使用情况
class ESLogstashTask(BaseLinuxServerTask):
    # 初始化Logstash服务器巡检任务：设置内存阈值和巡检参数
    def __init__(self):
        # 验证ESLogstashTask专用配置
        require_keys(CONFIG, ["ESServer"], "root")
        require_keys(CONFIG["ESServer"], ["thresholds"], "ESServer")
        require_keys(CONFIG["ESServer"]["thresholds"], ["mem_percent"], "ESServer.thresholds")
        require_keys(CONFIG["ESServer"]["thresholds"]["mem_percent"], ["ESLogstashTask"], "ESServer.thresholds.mem_percent")
        
        # 从配置文件读取ESLogstashTask的内存阈值配置
        MEM_THRESHOLDS = CONFIG["ESServer"]["thresholds"]["mem_percent"]["ESLogstashTask"]
        super().__init__("LOGSTASH服务器巡检", "ESLogstashTask", 
                        MEM_THRESHOLDS["warn"], MEM_THRESHOLDS["crit"])
