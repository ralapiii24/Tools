# 通用基础类和工具函数

# 导入标准库
import base64
import os
from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, Iterable, Optional, Tuple

# 导入第三方库
import paramiko
import yaml
from progress import create_progress, tqdm

# 读取配置
with open("YAML/Config.yaml", "r", encoding="utf-8") as fileHandle:
    CONFIG = yaml.safe_load(fileHandle)

_DEFAULT_SYNC_TOKEN = "inspection-v9-default-2024"

# 使用密钥解密密码：XOR解密算法，将enc:前缀的加密密码解密为明文
def _decrypt_with_key(encrypted_password: str, key: bytes) -> str:
    if not encrypted_password.startswith("enc:"):
        return encrypted_password
    try:
        ENCRYPTED_BYTES = base64.b64decode(encrypted_password[4:])
        DECRYPTED_BYTES = bytearray()
        for I, BYTE in enumerate(ENCRYPTED_BYTES):
            DECRYPTED_BYTES.append(BYTE ^ key[I % len(key)])
        return DECRYPTED_BYTES.decode('utf-8')
    except Exception:
        return None

# 使用密钥加密密码：XOR加密算法，将明文密码加密并添加enc:前缀
def _encrypt_with_key(plain_password: str, key: bytes) -> str:
    PLAIN_BYTES = plain_password.encode('utf-8')
    ENCRYPTED_BYTES = bytearray()
    for I, BYTE in enumerate(PLAIN_BYTES):
        ENCRYPTED_BYTES.append(BYTE ^ key[I % len(key)])
    ENCRYPTED_B64 = base64.b64encode(bytes(ENCRYPTED_BYTES)).decode('ascii')
    return f"enc:{ENCRYPTED_B64}"

# 重新加密所有密码：递归遍历配置字典，使用新密钥重新加密所有密码字段
def _re_encrypt_all_passwords(old_key: str, new_key: str, config_dict: dict, path: str = ""):
    UPDATED = False
    if isinstance(config_dict, dict):
        for KEY, VALUE in config_dict.items():
            CURRENT_PATH = f"{path}.{KEY}" if path else KEY
            if KEY == "password" and isinstance(VALUE, str) and VALUE.startswith("enc:"):
                OLD_KEY_BYTES = old_key.encode()
                PLAIN = _decrypt_with_key(VALUE, OLD_KEY_BYTES)
                if PLAIN:
                    NEW_KEY_BYTES = new_key.encode()
                    NEW_ENCRYPTED = _encrypt_with_key(PLAIN, NEW_KEY_BYTES)
                    config_dict[KEY] = NEW_ENCRYPTED
                    UPDATED = True
            elif isinstance(VALUE, dict):
                if _re_encrypt_all_passwords(old_key, new_key, VALUE, CURRENT_PATH):
                    UPDATED = True
            elif isinstance(VALUE, list):
                for I, ITEM in enumerate(VALUE):
                    if isinstance(ITEM, dict):
                        if _re_encrypt_all_passwords(old_key, new_key, ITEM, f"{CURRENT_PATH}[{I}]"):
                            UPDATED = True
    return UPDATED

# 同步加密密钥到配置：检查并更新配置中的加密密钥，必要时重新加密所有密码
def _sync_encrypt_key_to_config():
    try:
        SETTINGS = CONFIG.get("settings", {})
        OLD_CONFIG_KEY = SETTINGS.get("config_version")
        NEW_CONFIG_KEY = _DEFAULT_SYNC_TOKEN
        
        if OLD_CONFIG_KEY and OLD_CONFIG_KEY != NEW_CONFIG_KEY:
            NEED_RE_ENCRYPT = _re_encrypt_all_passwords(OLD_CONFIG_KEY, NEW_CONFIG_KEY, CONFIG)
            if NEED_RE_ENCRYPT:
                if "config_version" in SETTINGS:
                    del CONFIG["settings"]["config_version"]
                with open("YAML/Config.yaml", "w", encoding="utf-8") as FILE_HANDLE:
                    yaml.dump(CONFIG, FILE_HANDLE, allow_unicode=True, default_flow_style=False, sort_keys=False)
        elif "config_version" in SETTINGS:
            del CONFIG["settings"]["config_version"]
            with open("YAML/Config.yaml", "w", encoding="utf-8") as FILE_HANDLE:
                yaml.dump(CONFIG, FILE_HANDLE, allow_unicode=True, default_flow_style=False, sort_keys=False)
    except Exception:
        pass

_sync_encrypt_key_to_config()
with open("YAML/Config.yaml", "r", encoding="utf-8") as FILE_HANDLE:
    CONFIG = yaml.safe_load(FILE_HANDLE)

_ENCRYPT_KEY = os.environ.get("ENCRYPT_KEY", _DEFAULT_SYNC_TOKEN).encode()

# 加密密码：使用默认密钥加密明文密码，返回enc:前缀的加密字符串
def encrypt_password(plain_password: str) -> str:
    PLAIN_BYTES = plain_password.encode('utf-8')
    KEY_BYTES = _ENCRYPT_KEY
    ENCRYPTED_BYTES = bytearray()
    for I, BYTE in enumerate(PLAIN_BYTES):
        ENCRYPTED_BYTES.append(BYTE ^ KEY_BYTES[I % len(KEY_BYTES)])
    ENCRYPTED_B64 = base64.b64encode(bytes(ENCRYPTED_BYTES)).decode('ascii')
    return f"enc:{ENCRYPTED_B64}"

# 解密密码：使用默认密钥解密enc:前缀的加密密码，返回明文密码
def decrypt_password(encrypted_password: str) -> str:
    if not encrypted_password.startswith("enc:"):
        return encrypted_password
    try:
        ENCRYPTED_BYTES = base64.b64decode(encrypted_password[4:])
        KEY_BYTES = _ENCRYPT_KEY
        DECRYPTED_BYTES = bytearray()
        for I, BYTE in enumerate(ENCRYPTED_BYTES):
            DECRYPTED_BYTES.append(BYTE ^ KEY_BYTES[I % len(KEY_BYTES)])
        return DECRYPTED_BYTES.decode('utf-8')
    except Exception as ERROR:
        raise ValueError(f"处理失败: {ERROR}")

# 从配置读取常量
# 从YAML配置中读取超时设置（移除硬编码默认值，只从配置文件读取）
DEFAULT_SSH_TIMEOUT = CONFIG["timeouts"]["ssh"]
DEFAULT_HTTP_TIMEOUT = CONFIG["timeouts"]["http"]
DEFAULT_PAGE_GOTO_TIMEOUT = CONFIG["timeouts"]["page_goto"]
DEFAULT_SELECTOR_TIMEOUT = CONFIG["timeouts"]["selector"]

# 从YAML配置中读取网络设置（移除硬编码默认值，只从配置文件读取）
WEB_MAX_WORKERS = CONFIG["network"]["max_workers"]
WEB_DELAY_RANGE = tuple(CONFIG["network"]["delay_range"])
BLOCK_RES_TYPES = set(CONFIG["network"]["block_resources"])

# 根据配置参数动态生成进度条格式：从YAML配置构建自定义进度条显示格式
def _build_progress_format():
    PROGRESS_CONFIG = CONFIG["progress"]

    # 获取配置参数（移除硬编码默认值，只从配置文件读取）
    SHOW_PERCENTAGE = PROGRESS_CONFIG["show_percentage"]
    SHOW_BAR = PROGRESS_CONFIG["show_bar"]
    SHOW_COUNT = PROGRESS_CONFIG["show_count"]
    SHOW_ELAPSED = PROGRESS_CONFIG["show_elapsed"]
    SHOW_REMAINING = PROGRESS_CONFIG["show_remaining"]
    BAR_LENGTH = PROGRESS_CONFIG["bar_length"]
    PREFIX = PROGRESS_CONFIG["prefix"]
    SUFFIX = PROGRESS_CONFIG["suffix"]

    # 构建格式字符串
    PARTS = []

    if SHOW_PERCENTAGE:
        PARTS.append("{percentage:3.0f}%")

    if SHOW_BAR:
        PARTS.append(f"|{{bar:{BAR_LENGTH}}}|")

    if SHOW_COUNT:
        PARTS.append("{n:>3d}/{total:>3d}")

    if SHOW_ELAPSED or SHOW_REMAINING:
        TIME_PARTS = []
        ELAPSED_LABEL = PROGRESS_CONFIG["elapsed_label"]
        REMAINING_LABEL = PROGRESS_CONFIG["remaining_label"]
        if SHOW_ELAPSED:
            TIME_PARTS.append(f"{ELAPSED_LABEL}:{{elapsed}}")
        if SHOW_REMAINING:
            TIME_PARTS.append(f"{REMAINING_LABEL}:{{remaining}}")
        PARTS.append(f"[{'  '.join(TIME_PARTS)}]")

    # 组合最终格式
    FORMAT_STR = " ".join(PARTS)
    if PREFIX:
        FORMAT_STR = f"{PREFIX} {FORMAT_STR}"
    if SUFFIX:
        FORMAT_STR = f"{FORMAT_STR} {SUFFIX}"

    return FORMAT_STR

BAR_FORMAT = _build_progress_format()

# 检查配置键是否存在：验证配置字典中是否包含必需的键
def require_keys(d: Dict[str, Any], keys: Iterable[str], ctx: str) -> None:
    for KEY in keys:
        if KEY not in d:
            raise ValueError(f"配置缺失: {ctx}.{KEY}")

require_keys(CONFIG, ["settings", "MirrorFortiGateTask", "OxidizedTask", "ESServer"], "root")

SHOW_PROGRESS = bool(CONFIG["settings"]["show_progress"])

# 分级结果模型
# 告警级别枚举：定义巡检结果的严重程度等级
class Level(Enum):
    OK = "OK"
    WARN = "WARN"
    CRIT = "CRIT"
    ERROR = "ERROR"

# 巡检结果数据类：存储单个巡检任务的结果信息
@dataclass
class Result:
    level: str
    message: str
    meta: Optional[dict] = None

# 通用工具
# 百分比分级函数：根据阈值将数值映射到告警级别
def grade_percent(value: Optional[float], warn: int, crit: int) -> Level:
    if value is None:
        return Level.ERROR
    if value >= crit:
        return Level.CRIT
    if value >= warn:
        return Level.WARN
    return Level.OK

# 从文件名提取站点名：使用正则表达式从文件名中提取站点标识
def extract_site_from_filename(filename: str, pattern: str = r"^\d{8}-([^-]+)-") -> str:
    import re
    BASE = os.path.basename(filename)
    MATCH = re.match(pattern, BASE)
    return MATCH.group(1) if MATCH else "DEFAULT"

# 从设备名提取站点名：从设备名称中提取站点标识
def extract_site_from_device(device_name: str) -> str:
    import re
    # 支持 HX 和 P 开头的站点
    MATCH = re.match(r'^(HX\d+|P\d+)', device_name)
    return MATCH.group(1) if MATCH else None

# 创建SSH连接：建立到远程主机的SSH连接
def create_ssh_connection(host: str, port: int, username: str, password: str, timeout: int = DEFAULT_SSH_TIMEOUT) -> paramiko.SSHClient:
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(host, port=port, username=username, password=password, timeout=timeout)
    return ssh

# 生成安全的Excel工作表名称：确保工作表名称符合Excel规范
def safe_sheet_name(name: str) -> str:
    import re
    safe_name = re.sub(r'[:\\/*?\[\]]', "_", name)[:31]
    return safe_name or "Sheet1"

# 执行SSH命令：通过SSH连接执行远程命令并返回结果
def ssh_exec(ssh: paramiko.SSHClient, cmd: str, timeout: int = DEFAULT_SSH_TIMEOUT, label: str = "") -> Tuple[
    int, str, str]:
    try:
        stdin, stdout, stderr = ssh.exec_command(cmd, timeout=timeout)
        stdout_text = stdout.read().decode("utf-8", "ignore")
        stderr_text = stderr.read().decode("utf-8", "ignore")
        exit_code = stdout.channel.recv_exit_status()
        return exit_code, stdout_text, stderr_text
    except Exception as error:
        raise RuntimeError(f"SSH 命令执行失败 {label or cmd}: {error}")

# 任务框架
# 基础任务类：所有巡检任务的基类，提供通用的任务执行框架
class BaseTask:
    # 初始化基础任务：设置任务名称和结果列表
    def __init__(self, name: str):
        self.NAME = name
        self.RESULTS: list[Result] = []
        # 从配置读取是否压降OK级别日志（默认False，不抑制）
        try:
            self.SUPPRESS_OK_LOGS: bool = bool(CONFIG.get("settings", {}).get("suppress_ok_logs", False))
        except Exception:
            self.SUPPRESS_OK_LOGS = False

    # 获取任务项目列表：子类必须实现，返回要处理的项目列表
    def items(self):
        raise NotImplementedError

    # 处理单个项目：子类必须实现，处理单个任务项目
    def run_single(self, item):
        raise NotImplementedError

    # 添加结果记录：将巡检结果添加到结果列表
    def add_result(self, level: Level, message: str, meta: Optional[dict] = None) -> None:
        # 可选抑制OK级别的非关键日志
        if level == Level.OK and getattr(self, "SUPPRESS_OK_LOGS", False):
            return
        entry = Result(level=level.value, message=message, meta=meta)
        self.RESULTS.append(entry)

    # 执行任务：遍历所有项目并执行巡检，显示进度条
    def run(self) -> None:
        task_items = list(self.items())
        progress = self.create_progress(
            total=len(task_items),
            position_offset=0,
        ) if SHOW_PROGRESS else None

        try:
            for single_item in task_items:
                try:
                    self.run_single(single_item)
                except Exception as error:
                    self.add_result(Level.ERROR, f"{single_item} 运行异常: {error!r}")
                if progress:
                    progress.update(1)
        finally:
            if progress:
                progress.close()

    def create_progress(self, total: int, position_offset: int = 0, **kwargs):
        kwargs.setdefault("desc", self.NAME)
        kwargs.setdefault("leave", True)
        kwargs.setdefault("dynamic_ncols", True)
        kwargs.setdefault("bar_format", BAR_FORMAT)
        return create_progress(total=total, position_offset=position_offset, **kwargs)


