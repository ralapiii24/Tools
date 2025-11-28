# TASK模块初始化文件

# 包含所有任务类的导入
from .FXOSWebTask import FXOSWebTask
from .MirrorFortiGateTask import MirrorFortiGateTask
from .OxidizedTask import OxidizedTask
from .ESLogstashTask import ESLogstashTask
from .ESBaseTask import ESBaseTask
from .ESN9KLOGInspectTask import ESN9KLOGInspectTask
from .ESFlowTask import ESFlowTask
from .DeviceBackupTask import DeviceBackupTask
from .DeviceDIFFTask import DeviceDIFFTask
from .ASACompareTask import ASACompareTask
from .ACLDupCheckTask import ACLDupCheckTask
from .ACLArpCheckTask import ACLArpCheckTask
from .ACLCrossCheckTask import ACLCrossCheckTask
from .ASADomainCheckTask import ASADomainCheckTask
from .ASATempnetworkCheckTask import ASATempnetworkCheckTask
from .ServiceCheckTask import ServiceCheckTask
from .LogRecyclingTask import LogRecyclingTask

__all__ = [
    'FXOSWebTask',
    'MirrorFortiGateTask',
    'OxidizedTask',
    'ESLogstashTask',
    'ESBaseTask',
    'ESN9KLOGInspectTask',
    'ESFlowTask',
    'DeviceBackupTask',
    'DeviceDIFFTask',
    'ASACompareTask',
    'ACLDupCheckTask',
    'ACLArpCheckTask',
    'ACLCrossCheckTask',
    'ASADomainCheckTask',
    'ASATempnetworkCheckTask',
    'ServiceCheckTask',
    'LogRecyclingTask'
]
