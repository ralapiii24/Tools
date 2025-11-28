# ASA 临时出网地址检查任务
#
# 技术栈:Python, 正则表达式, 文件IO
# 目标:检查ASA防火墙中临时出网地址配置，识别非默认的临时出网地址并告警
#
# 处理逻辑:
# - 从LOG/OxidizedTask/OxidizedTaskBackup/读取cat3（ASA防火墙）的01设备配置
# - 解析object-group network ServTemp-To-Internet配置块
# - 忽略默认配置network-object host 1.1.1.1
# - 检测其他network-object host地址，作为临时出网地址，WARN等级输出对应站点
#
# 输入文件:LOG/OxidizedTask/OxidizedTaskBackup/（V10新结构：从LOG/日期/OxidizedTaskBackup迁移）
# 输出:LOG/ASATempnetworkCheckTask/YYYYMMDD-ASA临时出网地址检查.log（详细检查日志）
# 配置说明:自动扫描当日配置，按站点输出临时出网地址

# 导入标准库
import os
import re
from datetime import datetime

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
from .TaskBase import BaseTask, Level, extract_site_from_device, CONFIG

# ASA临时出网地址检查任务类：检查ASA防火墙中临时出网地址配置
class ASATempnetworkCheckTask(BaseTask):
    
    # 初始化ASA临时出网地址检查任务：设置任务名称和路径
    def __init__(self):
        super().__init__("ASA临时出网地址检查")
        # 从配置文件读取log_dir（必须配置）
        SETTINGS = CONFIG.get("settings", {})
        self.LOG_DIR = SETTINGS.get("log_dir", "LOG")
        # V10新结构：直接输出到 LOG/ASATempnetworkCheckTask/
        self.OUTPUT_DIR = os.path.join(self.LOG_DIR, "ASATempnetworkCheckTask")
        self.LOG_DIR_PATH = os.path.join(self.LOG_DIR, "OxidizedTask", "OxidizedTaskBackup")
        self._TODAY = None
        self._DEFAULT_IP = "1.1.1.1"  # 默认配置IP，忽略
        self._OBJECT_GROUP_NAME = "ServTemp-To-Internet"  # 要检查的对象组名称

    # 扫描LOG目录获取站点列表：查找cat3的01设备配置
    def items(self):
        self._TODAY = datetime.now().strftime("%Y%m%d")
        
        if not os.path.isdir(self.LOG_DIR_PATH):
            self.add_result(Level.ERROR, f"未找到当日日志目录: {self.LOG_DIR_PATH}")
            return []

        # 创建输出目录（如果目录已存在则不报错）
        os.makedirs(self.OUTPUT_DIR, exist_ok=True)

        # 扫描LOG目录，查找cat3的01设备（ASA防火墙，设备编号为01）
        sites = []
        
        for filename in os.listdir(self.LOG_DIR_PATH):
            if not filename.lower().endswith('.log'):
                continue

            # 只处理当天的文件
            if not filename.startswith(self._TODAY + '-'):
                continue

            device_name = filename[len(self._TODAY) + 1:-4]  # 去掉日期前缀和.log后缀
            device_lower = device_name.lower()

            # 检查是否为cat3的01设备（ASA防火墙，设备编号为01）
            # cat3设备特征：包含fw01-frp或fw02-frp，但这里只检查01设备
            if 'fw01-frp' in device_lower:
                site = self._extract_site_from_device(device_name)
                if site:
                    file_path = os.path.join(self.LOG_DIR_PATH, filename)
                    sites.append((site, file_path, device_name))

        if not sites:
            self.add_result(Level.WARN, "未找到cat3的01设备配置")
            return []

        return sites

    # 从设备名中提取站点名：解析设备名称获取站点标识，如HX03-FW01-FRP2140-JPIDC -> HX03
    @staticmethod
    def _extract_site_from_device(device_name: str) -> str:
        return extract_site_from_device(device_name)

    # 检查单个站点配置：解析object-group network ServTemp-To-Internet并检查临时出网地址
    def run_single(self, item: tuple) -> None:
        site, file_path, device_name = item
        
        try:
            # 读取配置文件
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 检查配置采集是否失败
            if self._check_collection_failure(content, device_name):
                self.add_result(Level.ERROR, f"站点 {site} 设备 {device_name} 配置采集失败")
                return

            # 提取object-group network ServTemp-To-Internet配置块
            temp_ips = self._extract_temp_network_objects(content)
            
            if temp_ips:
                for ip in temp_ips:
                    self.add_result(Level.WARN, f"站点 {site} 发现临时出网地址: {ip}")
            else:
                # 如果没有找到临时出网地址，输出OK（可选，根据需求决定是否输出）
                pass

        except Exception as e:
            self.add_result(Level.ERROR, f"站点 {site} 设备 {device_name} 处理失败: {e}")

    # 检查配置采集是否失败：通过内容分析判断配置采集是否成功
    @staticmethod
    def _check_collection_failure(content: str, device_name: str) -> bool:
        if not content or len(content.strip()) < 500:
            return True

        content_lower = content.lower()
        
        # 检查明确的采集失败标识
        failure_indicators = [
            'node not found',
            'connection failed',
            'unable to connect',
            'authentication failed',
            'access denied',
            'no such file',
            'device unreachable',
            'oxidized error',
            'collection failed',
            'ssh connection failed',
            'telnet connection failed',
            'login failed',
            'permission denied'
        ]

        for indicator in failure_indicators:
            if indicator in content_lower:
                return True

        # 检查是否包含ASA配置的典型标识
        asa_indicators = [
            'hostname',
            'interface',
            'ip address',
            'access-list',
            'nat',
            'route',
            'crypto',
            'object',
            'object-group'
        ]

        has_asa_indicator = False
        for indicator in asa_indicators:
            if indicator in content_lower:
                has_asa_indicator = True
                break

        return not has_asa_indicator

    # 提取临时出网地址：从配置中提取object-group network ServTemp-To-Internet的network-object host地址
    def _extract_temp_network_objects(self, content: str) -> list:
        temp_ips = []
        
        # 匹配object-group network ServTemp-To-Internet配置块
        # 支持两种格式：
        # 格式1: network-object host 1.1.1.1
        # 格式2: network-object 1.1.1.1 255.255.255.255
        
        # 查找object-group network ServTemp-To-Internet的位置
        start_pattern = r'object-group\s+network\s+ServTemp-To-Internet'
        start_match = re.search(start_pattern, content, re.IGNORECASE)
        
        if not start_match:
            return temp_ips
        
        # 从匹配位置开始，提取后续的network-object行
        start_pos = start_match.end()
        remaining_content = content[start_pos:]
        
        # 匹配所有以空格或tab开头的network-object行
        # 直到遇到不以空格或tab开头的行（下一个顶级配置项）或文件末尾
        lines = remaining_content.split('\n')
        for line in lines:
            line_stripped = line.strip()
            # 如果遇到不以空格开头的非空行，说明到了下一个配置项，停止解析
            if line_stripped and not line.startswith((' ', '\t')):
                break
            
            # 匹配两种格式：
            # 格式1: network-object host 1.1.1.1
            # 格式2: network-object 1.1.1.1 255.255.255.255
            network_object_match = re.match(
                r'\s+network-object\s+(?:host\s+)?(\d+\.\d+\.\d+\.\d+)(?:\s+\d+\.\d+\.\d+\.\d+)?',
                line,
                re.IGNORECASE
            )
            if network_object_match:
                ip = network_object_match.group(1)
                # 忽略默认配置IP
                if ip != self._DEFAULT_IP:
                    temp_ips.append(ip)
        
        return temp_ips

