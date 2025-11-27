import os
import re
import json
import time
import yaml
import paramiko
import requests
from lxml import html
from enum import Enum
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from typing import Tuple, Dict, Any, Optional, Iterable
from tqdm import tqdm
from playwright.sync_api import sync_playwright

# ======================================================================== 默认常量 ========================================================================
DEFAULT_SSH_TIMEOUT = 12
DEFAULT_HTTP_TIMEOUT = 12
DEFAULT_PAGE_GOTO_TIMEOUT = 60000
DEFAULT_SELECTOR_TIMEOUT = 30000

# 仅对这些主机进行 Uptime 调试输出（一次性块状打印 + 落盘）
# 固定 25 格的进度条样式
BAR_FORMAT = "{l_bar}{bar:25}| {n_fmt}/{total_fmt} [当前耗时:{elapsed}  任务预期剩余耗时:{remaining}]"

# ======================================================================== 正则预编译 ========================================================================
RE_CPU_IDLE = re.compile(r"CPU\s+states:\s*.*?(\d+)\s*%\s*idle", re.IGNORECASE)
RE_MEM_USED = re.compile(r"Memory:\s*.*?used\s*\(\s*([\d\.]+)\s*%\s*\)", re.IGNORECASE)
RE_UPTIME_LINE = re.compile(r"(?im)^[^\S\r\n]*Uptime[:：]?\s*(.+)$")
RE_UPTIME_UNITS = re.compile(
    r"(\d+)\s*d(?:ay)?s?.{0,20}?(\d+)\s*h(?:our)?s?.{0,20}?(\d+)\s*m(?:in(?:ute)?)?s?",
    re.IGNORECASE,
)

def safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", str(name)).strip()

# ======================================================================== 读取配置 ========================================================================
with open("YAML/Config.yaml", "r", encoding="utf-8") as f:
    CONFIG = yaml.safe_load(f)

def require_keys(d: Dict[str, Any], keys: Iterable[str], ctx: str) -> None:
    for key in keys:
        if key not in d:
            raise ValueError(f"配置缺失: {ctx}.{key}")

require_keys(CONFIG, ["settings", "fortigate", "oxidized", "servers"], "root")
require_keys(CONFIG["servers"], ["port", "thresholds", "logstash", "es", "flow"], "servers")

SHOW_PROGRESS = bool(CONFIG["settings"].get("show_progress", True))
# ======================================================================== 分级结果模型 ========================================================================
class Level(Enum):
    OK = "OK"
    WARN = "WARN"
    CRIT = "CRIT"
    ERROR = "ERROR"

@dataclass
class Result:
    level: str
    message: str
    meta: Optional[dict] = None

# ======================================================================== 通用工具 ========================================================================
def clip(text: str, max_len: int = 4000) -> str:
    return text if len(text) <= max_len else text[:max_len] + f"...[+{len(text)-max_len} chars]"

def ssh_exec(ssh: paramiko.SSHClient, cmd: str, timeout: int = DEFAULT_SSH_TIMEOUT, label: str = "") -> Tuple[int, str, str]:
    try:
        stdin, stdout, stderr = ssh.exec_command(cmd, timeout=timeout)
        stdout_text = stdout.read().decode("utf-8", "ignore")
        stderr_text = stderr.read().decode("utf-8", "ignore")
        exit_code = stdout.channel.recv_exit_status()
        return exit_code, stdout_text, stderr_text
    except Exception as exc:
        raise RuntimeError(f"SSH 命令执行失败 {label or cmd}: {exc}")

def ssh_exec_paged(ssh: paramiko.SSHClient, cmd: str, timeout: int = 20) -> str:
    """
    处理 FortiGate 开启分页（--More--）的命令输出。
    通过 invoke_shell 发送命令，遇到 '--More--' 自动回空格取下一页，直到回到提示符或超时。
    """
    chan = ssh.invoke_shell(width=200, height=5000)
    chan.settimeout(1.0)

    # 清空欢迎/提示符残留
    try:
        time.sleep(0.2)
        while chan.recv_ready():
            _ = chan.recv(65535)
    except Exception:
        pass

    chan.send(cmd + "\n")
    buf = ""
    start = time.time()
    idle_start = None

    # 简单提示符判断：行尾出现 '#' 或 '>'，且屏幕上没有 '--More--'
    prompt_re = re.compile(r"(?m)[#>]\s*$")

    while time.time() - start < timeout:
        got = False
        try:
            if chan.recv_ready():
                chunk = chan.recv(65535).decode("utf-8", "ignore")
                if not chunk:
                    continue
                buf += chunk
                got = True

                if "--More--" in chunk or "--More--" in buf:
                    # 清除上一屏的 --More-- 提示（有时前面会有回退控制符）
                    chan.send(" ")
                    # 继续循环读下一批
                    continue
        except Exception:
            pass

        if got:
            idle_start = None
            # 没有更多 --More--，且出现提示符，基本可认为结束
            if "--More--" not in buf and prompt_re.search(buf):
                break
        else:
            # 没数据：若也没有 --More--，累计空闲一段时间后退出
            if "--More--" not in buf:
                if idle_start is None:
                    idle_start = time.time()
                elif time.time() - idle_start > 1.0:
                    break
            time.sleep(0.1)

    try:
        chan.close()
    except Exception:
        pass

    # 去掉命令回显的首行（若存在）
    lines = buf.splitlines()
    if lines and lines[0].strip().startswith(cmd.split()[0]):
        lines = lines[1:]
    return "\n".join(lines)

def to_bytes(size_str: str) -> int:
    if not size_str:
        return -1
    normalized = size_str.strip().lower()
    try:
        if normalized.endswith("gb"):
            return int(float(normalized[:-2]) * 1024**3)
        if normalized.endswith("mb"):
            return int(float(normalized[:-2]) * 1024**2)
        if normalized.endswith("kb"):
            return int(float(normalized[:-2]) * 1024)
        if normalized.endswith("b"):
            return int(float(normalized[:-1]))
        return int(float(normalized))
    except Exception:
        return -1

# ======================================================================== 任务框架 ========================================================================
class BaseTask:
    def __init__(self, name: str):
        self.name = name
        self.results: list[Result] = []

    def items(self):
        raise NotImplementedError

    def run_single(self, item):
        raise NotImplementedError

    def add_result(self, level: Level, message: str, meta: dict | None = None) -> None:
        entry = Result(level=level.value, message=message, meta=meta)
        self.results.append(entry)
    def run(self) -> None:
        task_items = list(self.items())
        progress = tqdm(
            total=len(task_items),
            desc=self.name,
            position=0,
            leave=True,
            dynamic_ncols=True,
            bar_format=BAR_FORMAT,
        ) if SHOW_PROGRESS else None

        try:
            for single_item in task_items:
                try:
                    self.run_single(single_item)
                except Exception as exc:
                    self.add_result(Level.ERROR, f"{single_item} 运行异常: {exc!r}")
                if progress:
                    progress.update(1)
        finally:
            if progress:
                progress.close()

# ======================================================================== FXOS WEB 巡检（带可选“自动回车/按钮兜底”）========================================================================
class FXOSWebTask(BaseTask):
    def __init__(self):
        super().__init__("FXOS WEB 巡检")
        fxos_cfg = CONFIG["fxos"]
        self.username = fxos_cfg["username"]
        self.password = fxos_cfg["password"]
        self.device_urls: Dict[str, str] = fxos_cfg["devices"]
        self.expected_xpath = (
            '/html/body/div[6]/div/div/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/'
            'table/tbody/tr/td[1]/table/tbody/tr/td[5]/div'
        )
        self.auto_press_enter: bool = bool(fxos_cfg.get("auto_press_enter", False))
        self.enter_retries: int = int(fxos_cfg.get("enter_retries", 5))
        self.enter_interval_ms: int = int(fxos_cfg.get("enter_interval_ms", 400))

    def items(self):
        return list(self.device_urls.items())

    def _nudge_continue(self, page) -> None:
        if not self.auto_press_enter:
            return
        try:
            def _on_dialog(dialog):
                try:
                    dialog.accept()
                except Exception:
                    pass
            page.once("dialog", _on_dialog)
        except Exception:
            pass

        for _ in range(self.enter_retries):
            try:
                page.keyboard.press("Enter")
            except Exception:
                pass
            page.wait_for_timeout(self.enter_interval_ms)
            try:
                page.keyboard.press("Tab")
                page.keyboard.press("Enter")
            except Exception:
                pass
            page.wait_for_timeout(self.enter_interval_ms)

            for selector in (
                "text=Continue", "text=Proceed", "text=OK", "text=Confirm",
                "text=继续", "text=确认", "text=确定",
                "xpath=//button[contains(.,'Continue') or contains(.,'Proceed') or contains(.,'OK') or contains(.,'确认') or contains(.,'继续') or contains(.,'确定')]",
                "xpath=//a[contains(.,'Continue') or contains(.,'Proceed') or contains(.,'OK') or contains(.,'确认') or contains(.,'继续') or contains(.,'确定')]",
            ):
                try:
                    el = page.query_selector(selector)
                    if el:
                        el.click()
                        page.wait_for_timeout(self.enter_interval_ms)
                except Exception:
                    pass

    def run_single(self, item):
        device_name, url = item
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(ignore_https_errors=True)
            page = context.new_page()
            try:
                page.goto(url, timeout=DEFAULT_PAGE_GOTO_TIMEOUT)
                try:
                    page.wait_for_selector('xpath=/html/body/center/div/form/a[1]', timeout=5000)
                    page.click('xpath=/html/body/center/div/form/a[1]')
                except Exception:
                    pass

                self._nudge_continue(page)

                page.fill('xpath=/html/body/center/div/form/div[3]/input[1]', self.username)
                page.fill('xpath=/html/body/center/div/form/div[3]/input[2]', self.password)
                page.click('xpath=/html/body/center/div/form/a[2]')

                self._nudge_continue(page)

                page.wait_for_selector(f'xpath={self.expected_xpath}', timeout=DEFAULT_SELECTOR_TIMEOUT)
                self.add_result(Level.OK, f"{device_name}({url}) 登录成功")
            except Exception as exc:
                self.add_result(Level.ERROR, f"{device_name}({url}) 登录失败: {exc}")
            finally:
                try:
                    context.close()
                except Exception:
                    pass
                browser.close()

# ======================================================================== 镜像飞塔防火墙 巡检（含 CPU/内存/Uptime）========================================================================
class MirrorFortiGateTask(BaseTask):
    def __init__(self):
        super().__init__("镜像飞塔防火墙巡检")
        forti_cfg = CONFIG["fortigate"]
        self.username = forti_cfg["username"]
        self.password = forti_cfg["password"]
        self.port: int = int(forti_cfg.get("port", 22))
        self.hosts: list[str] = forti_cfg["hosts"]

        thr_disk = forti_cfg.get("thresholds", {}).get("disk_percent", {})
        self.disk_warn = int(thr_disk.get("warn", 60))
        self.disk_crit = int(thr_disk.get("crit", 80))

        thr_cpu = forti_cfg.get("thresholds", {}).get("cpu_percent", {})
        self.cpu_warn = int(thr_cpu.get("warn", 50))
        self.cpu_crit = int(thr_cpu.get("crit", 80))
        thr_mem = forti_cfg.get("thresholds", {}).get("mem_percent", {})
        self.mem_warn = int(thr_mem.get("warn", 50))
        self.mem_crit = int(thr_mem.get("crit", 80))
        thr_uptime = forti_cfg.get("thresholds", {}).get("uptime_days", {})
        self.min_uptime_days = int(thr_uptime.get("min", 3))

    def items(self):
        return self.hosts

    @staticmethod
    def parse_disk_percent(output_text: str) -> Optional[float]:
        for line in output_text.splitlines():
            if 'HD logging space usage for vdom "root"' in line:
                match = re.search(r'(\d+)MB(?:\(\d+MiB\))?\s*/\s*(\d+)MB', line)
                if match:
                    used_mb, total_mb = int(match.group(1)), int(match.group(2))
                    return round(used_mb / total_mb * 100, 2) if total_mb > 0 else None
        return None

    @staticmethod
    def parse_perf_status(output_text: str) -> Tuple[Optional[float], Optional[float], Optional[float]]:
        cleaned = (output_text or "").replace("\r", "")
        cleaned = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", " ", cleaned)
        cleaned = re.sub(r"\x1b\[[0-9;?]*[ -/]*[@-~]", "", cleaned)

        cpu_used_pct: Optional[float] = None
        mem_used_pct: Optional[float] = None
        uptime_days: Optional[float] = None

        match_idle = RE_CPU_IDLE.search(cleaned)
        if match_idle:
            try:
                idle_pct = float(match_idle.group(1))
                cpu_used_pct = round(100.0 - idle_pct, 2)
            except Exception:
                pass

        match_mem = RE_MEM_USED.search(cleaned)
        if match_mem:
            try:
                mem_used_pct = float(match_mem.group(1))
            except Exception:
                pass

        def _norm(s: str) -> str:
            try:
                import unicodedata
                s = unicodedata.normalize("NFKC", s)
            except Exception:
                pass
            s = s.replace("\xa0", " ")
            s = re.sub(r"\s+", " ", s).strip()
            return s

        tail_text: Optional[str] = None
        match_line = RE_UPTIME_LINE.search(cleaned)
        if match_line:
            tail_text = match_line.group(1)
        else:
            match_pos = re.search(r"uptime\s*[:：]?", cleaned, re.IGNORECASE)
            if match_pos:
                segment = cleaned[match_pos.end(): match_pos.end() + 200]
                segment = segment.split("\n")[0]
                segment = re.split(r"[#>]\s*$", segment)[0]
                tail_text = segment
            else:
                all_units = list(RE_UPTIME_UNITS.finditer(cleaned))
                if all_units:
                    days_val, hours_val, minutes_val = map(int, all_units[-1].groups())
                    try:
                        uptime_days = days_val + hours_val / 24.0 + minutes_val / (24.0 * 60.0)
                    except Exception:
                        uptime_days = None

        if tail_text and uptime_days is None:
            normalized = _norm(tail_text)
            match_days = re.search(r"(\d+)\s*d(?:ay)?s?", normalized, re.IGNORECASE)
            match_hours = re.search(r"(\d+)\s*h(?:our)?s?", normalized, re.IGNORECASE)
            match_minutes = re.search(r"(\d+)\s*m(?:in(?:ute)?)?s?", normalized, re.IGNORECASE)

            if match_days and match_hours and match_minutes:
                days_val = int(match_days.group(1))
                hours_val = int(match_hours.group(1))
                minutes_val = int(match_minutes.group(1))
            else:
                nums = re.findall(r"(\d+)", normalized)
                if len(nums) >= 3:
                    days_val, hours_val, minutes_val = int(nums[0]), int(nums[1]), int(nums[2])
                else:
                    days_val = hours_val = minutes_val = None

            if days_val is not None and hours_val is not None and minutes_val is not None:
                try:
                    uptime_days = days_val + hours_val / 24.0 + minutes_val / (24.0 * 60.0)
                except Exception:
                    uptime_days = None

        return cpu_used_pct, mem_used_pct, uptime_days

    @staticmethod
    def get_hostname(ssh: paramiko.SSHClient) -> str:
        exit_code, stdout_text, _ = ssh_exec(ssh, "get system status | grep Hostname", label="get hostname")
        match = re.search(r"Hostname:\s*(\S+)", stdout_text)
        return match.group(1) if match else "Unknown"

    @staticmethod
    def grade(value: Optional[float], warn: int, crit: int) -> Level:
        if value is None:
            return Level.ERROR
        if value >= crit:
            return Level.CRIT
        if value >= warn:
            return Level.WARN
        return Level.OK

    def run_single(self, host: str) -> None:
        ssh: Optional[paramiko.SSHClient] = None
        try:
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(host, port=self.port, username=self.username, password=self.password, timeout=DEFAULT_SSH_TIMEOUT)

            hostname = self.get_hostname(ssh)

            # 1) 日志盘
            _, disk_output, _ = ssh_exec(ssh, "diagnose sys logdisk usage", label="logdisk usage")
            disk_percent = self.parse_disk_percent(disk_output)
            if disk_percent is None:
                self.add_result(Level.ERROR, f"{hostname}({host}) 无日志盘信息")
            elif disk_percent >= self.disk_crit:
                self.add_result(Level.CRIT, f"{hostname}({host}) 日志盘 {disk_percent}% (>= {self.disk_crit}%)")
            elif disk_percent >= self.disk_warn:
                self.add_result(Level.WARN, f"{hostname}({host}) 日志盘 {disk_percent}% (>= {self.disk_warn}%)")
            else:
                self.add_result(Level.OK, f"{hostname}({host}) 日志盘 {disk_percent}%")

            # 2) 性能状态（CPU/内存/Uptime）
            # 先用普通 exec_command 拿一把；若检测到分页或缺少 Uptime，再用交互式分页读取
            _, perf_output, _ = ssh_exec(ssh, "get system performance status", label="perf status")
            need_paged = ("--More--" in perf_output) or ("Uptime" not in perf_output and "uptime" not in perf_output)
            if need_paged:
                perf_output = ssh_exec_paged(ssh, "get system performance status")

            cpu_used, mem_used, uptime_days = self.parse_perf_status(perf_output)

            # 调试：打印 perf 回显与初次解析结果（仅目标主机）
            # 兜底：若 Uptime 仍为空，再从 get system status 抓一次（同样考虑分页）
            if uptime_days is None:
                # 尽量直接读完整的 system status（不要带 grep，部分设备不支持管道/同样会分页）
                sys_status_full = ssh_exec_paged(ssh, "get system status")
                _, _, uptime_try = self.parse_perf_status(sys_status_full)
                if uptime_try is not None:
                    uptime_days = uptime_try

            # CPU
            level_cpu = self.grade(cpu_used, self.cpu_warn, self.cpu_crit)
            if level_cpu == Level.ERROR:
                self.add_result(level_cpu, f"{hostname}({host}) CPU 使用率解析失败")
            else:
                self.add_result(level_cpu, f"{hostname}({host}) CPU {cpu_used}%（阈{self.cpu_warn}/{self.cpu_crit}%）")

            # Memory
            level_mem = self.grade(mem_used, self.mem_warn, self.mem_crit)
            if level_mem == Level.ERROR:
                self.add_result(level_mem, f"{hostname}({host}) 内存使用率解析失败")
            else:
                self.add_result(level_mem, f"{hostname}({host}) 内存 {mem_used}%（阈{self.mem_warn}/{self.mem_crit}%）")

            # Uptime（必须 >= 3 天）
            if uptime_days is None:
                self.add_result(Level.ERROR, f"{hostname}({host}) Uptime 解析失败")
            else:
                if uptime_days < self.min_uptime_days:
                    self.add_result(Level.CRIT, f"{hostname}({host}) Uptime 仅 {uptime_days:.2f} 天（< {self.min_uptime_days} 天）")
                else:
                    self.add_result(Level.OK, f"{hostname}({host}) Uptime {uptime_days:.2f} 天")

        except Exception as exc:
            self.add_result(Level.ERROR, f"{host} 巡检失败: {exc}")
        finally:
            try:
                if ssh:
                    ssh.close()
            except Exception:
                pass

# ======================================================================== Oxidized 设备配置本地备份 ========================================================================
class OxidizedTask(BaseTask):
    def __init__(self, log_dir: str):
        super().__init__("Oxidized 设备配置本地备份")
        self.base_urls: list[str] = CONFIG["oxidized"]["base_urls"]
        self.log_dir = log_dir

    def items(self):
        return self.base_urls

    def run_single(self, base_url: str) -> None:
        try:
            session = requests.Session()
            resp = session.get(base_url, timeout=DEFAULT_HTTP_TIMEOUT)
            resp.raise_for_status()
        except Exception as exc:
            self.add_result(Level.ERROR, f"{base_url} 无法访问: {exc}")
            return

        tree = html.fromstring(resp.content)
        device_names = tree.xpath("//table/tbody/tr/td[1]/a/text()")
        group_names = tree.xpath("//table/tbody/tr/td[3]/a/text()")

        succeeded, failed = 0, 0
        today_str = datetime.now().strftime("%Y%m%d")

        for device_name, group_name in zip(device_names, group_names):
            device = device_name.strip()
            group = group_name.strip()
            fetch_url = f"{base_url.replace('/nodes','')}/node/fetch/{group}/{device}"
            try:
                cfg_resp = session.get(fetch_url, timeout=DEFAULT_HTTP_TIMEOUT)
                cfg_resp.raise_for_status()
                log_path = os.path.join(self.log_dir, f"{today_str}-{safe_filename(device)}.log")
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write(cfg_resp.text)
                succeeded += 1
                self.add_result(Level.OK, f"{device}({group}) 备份成功")
            except Exception as exc:
                failed += 1
                self.add_result(Level.ERROR, f"{device}({group}) 获取失败: {exc}")

        summary_level = Level.OK if failed == 0 else Level.WARN
        self.add_result(summary_level, f"{base_url} 处理完成，成功 {succeeded}，失败 {failed}")

# ======================================================================== Linux 服务器巡检基类（df -h + free -m）========================================================================
class BaseLinuxServerTask(BaseTask):
    def __init__(self, name: str, section_key: str, mem_warn: int, mem_crit: int):
        super().__init__(name)
        servers_cfg = CONFIG["servers"]
        section_cfg = servers_cfg[section_key]
        self.username = section_cfg["username"]
        self.password = section_cfg["password"]
        self.port: int = int(servers_cfg.get("port", 22))
        self.hosts_map: Dict[str, str] = section_cfg["hosts"]

        disk_thr = servers_cfg["thresholds"]["disk_percent"]
        self.disk_warn = int(disk_thr["warn"])
        self.disk_crit = int(disk_thr["crit"])

        self.mem_warn = int(mem_warn)
        self.mem_crit = int(mem_crit)

    def items(self):
        return list(self.hosts_map.items())

    @staticmethod
    def _ssh(ip: str, port: int, user: str, pwd: str) -> paramiko.SSHClient:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(ip, port=port, username=user, password=pwd, timeout=DEFAULT_SSH_TIMEOUT)
        return ssh

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

    @staticmethod
    def _grade_percent(value: Optional[float], warn: int, crit: int) -> Level:
        if value is None:
            return Level.ERROR
        if value >= crit:
            return Level.CRIT
        if value >= warn:
            return Level.WARN
        return Level.OK

    def run_single(self, item: Tuple[str, str]) -> None:
        server_name, ip_addr = item
        ssh: Optional[paramiko.SSHClient] = None
        try:
            ssh = self._ssh(ip_addr, self.port, self.username, self.password)

            _, df_stdout, _ = ssh_exec(ssh, "df -h", label="df -h")
            disk_used_pct, raw_df_line = self._parse_df_root(df_stdout)

            _, free_stdout, _ = ssh_exec(ssh, "free -m", label="free -m")
            total_mb, used_mb, mem_used_pct = self._parse_free_m(free_stdout)
        except Exception as exc:
            self.add_result(Level.ERROR, f"{server_name} {ip_addr} 巡检失败：{exc}")
            return
        finally:
            try:
                if ssh:
                    ssh.close()
            except Exception:
                pass

        if mem_used_pct is None:
            self.add_result(Level.ERROR, f"{server_name} {ip_addr} 内存信息解析失败")
        else:
            mem_level = self._grade_percent(mem_used_pct, self.mem_warn, self.mem_crit)
            self.add_result(mem_level, f"{server_name} {ip_addr} 内存{mem_used_pct}%（阈{self.mem_warn}/{self.mem_crit}%）")

        if disk_used_pct is None:
            self.add_result(Level.ERROR, f"{server_name} {ip_addr} 未找到根分区 / 磁盘信息")
        else:
            disk_level = self._grade_percent(disk_used_pct, self.disk_warn, self.disk_crit)
            message = f"{server_name} {ip_addr} 磁盘{disk_used_pct}%（阈{self.disk_warn}/{self.disk_crit}%）"
            if disk_level != Level.OK and raw_df_line:
                message += f"；原行: {raw_df_line}"
            self.add_result(disk_level, message)

# ======================================================================== LOGSTASH / ES / FLOW 具体任务 ========================================================================
class LogstashServerTask(BaseLinuxServerTask):
    def __init__(self):
        thr = CONFIG["servers"]["thresholds"]["mem_percent"]["LOGSTASH"]
        super().__init__("LOGSTASH 服务器巡检", "logstash", thr["warn"], thr["crit"])

class ElasticsearchServerTask(BaseLinuxServerTask):
    def __init__(self):
        thr = CONFIG["servers"]["thresholds"]["mem_percent"]["ES"]
        super().__init__("ES 服务器巡检", "es", thr["warn"], thr["crit"])

class FlowServerTask(BaseLinuxServerTask):
    def __init__(self):
        thr = CONFIG["servers"]["thresholds"]["mem_percent"]["FLOW"]
        super().__init__("FLOW 服务器巡检", "flow", thr["warn"], thr["crit"])
        self.fc = CONFIG["servers"]["flow_checks"]

    def run_single(self, item: Tuple[str, str]) -> None:
        super().run_single(item)

        server_name, ip_addr = item
        try:
            ssh = self._ssh(ip_addr, self.port, self.username, self.password)

            _, netstat_stdout, _ = ssh_exec(ssh, "netstat -tulnp", label="ports")
            docker_ec, docker_stdout, docker_stderr = ssh_exec(ssh, 'docker ps --format "{{.Names}} {{.Status}}"', label="docker ps")
            _, indices_stdout, _ = ssh_exec(ssh, "curl -s 'http://localhost:9200/_cat/indices?v'", label="es indices")
            _, segments_stdout, _ = ssh_exec(ssh, "curl -s 'http://localhost:9200/_cat/segments?v'", label="es segments")
        except Exception as exc:
            self.add_result(Level.ERROR, f"{server_name} {ip_addr} FLOW专项巡检失败：{exc}")
            return
        finally:
            try:
                if ssh:
                    ssh.close()
            except Exception:
                pass

        for required_port in self.fc.get("require_ports", []):
            pattern = rf":{required_port}\b.*LISTEN"
            if not re.search(pattern, netstat_stdout):
                self.add_result(Level.CRIT, f"{server_name} {ip_addr} 端口 {required_port} 未监听")

        # 容器检查：优先使用 docker ps --format 的精确名称匹配；若 docker 失败，则按端口占用做兜底
        if docker_ec == 0 and docker_stdout.strip():
            running = set()
            for line in docker_stdout.splitlines():
                parts = line.strip().split(None, 1)
                if not parts:
                    continue
                name = parts[0].strip()
                status = parts[1] if len(parts) > 1 else ""
                if re.search(r"\bUp\b", status):
                    running.add(name)
            for container_name in self.fc.get("require_containers", []):
                if container_name not in running:
                    self.add_result(Level.CRIT, f"{server_name} {ip_addr} 容器 {container_name} 未运行(或STATUS非Up)")
        else:
            # docker ps 执行失败：检查关键端口是否被占用，若有则视为通过，否则失败
            fallback_ports = [5601, 9600, 9300, 9200, 4739, 2055, 6343]
            # 从 netstat 原始输出中过滤匹配到的行，兼容空格/制表符分隔
            filtered_lines = []
            for line in netstat_stdout.splitlines():
                for p in fallback_ports:
                    if f":{p} " in line or f":{p}	" in line:
                        filtered_lines.append(line.rstrip())
                        break
            if filtered_lines:
                self.add_result(Level.OK, f"{server_name} {ip_addr} docker ps 失败，但端口命中如下：\n" + "\n".join(filtered_lines))
            else:
                self.add_result(Level.ERROR, f"{server_name} {ip_addr} docker ps 失败且关键端口未占用")
        index_prefix = self.fc.get("index_prefix", "elastiflow-4.0.1-")
        size_limit_bytes = int(self.fc.get("index_size_limit_bytes", 1024**3))
        date_regex = re.compile(re.escape(index_prefix) + r"(\d{4}\.\d{2}\.\d{2})")

        index_lines = [line.strip() for line in indices_stdout.splitlines() if index_prefix in line and line.strip()]
        date_set = set()
        oversize_list = []
        for line in index_lines:
            match = date_regex.search(line)
            if not match:
                continue
            date_str = match.group(1)
            date_set.add(date_str)
            cols = line.split()
            if not cols:
                continue
            last_size_field = cols[-1].lower()
            if to_bytes(last_size_field) > size_limit_bytes:
                oversize_list.append(f"{date_str} 大小 {last_size_field}")

        if len(date_set) > 31:
            self.add_result(Level.WARN, f"{server_name} {ip_addr} 索引日期数量 {len(date_set)} 超过 31")
        if oversize_list:
            self.add_result(Level.WARN, f"{server_name} {ip_addr} 索引大小超过1G: " + "；".join(oversize_list))

        segment_counter: Dict[str, int] = {}
        segment_dates = set()
        for line in segments_stdout.splitlines():
            striped = line.strip()
            if not striped or striped.lower().startswith("index"):
                continue
            parts = striped.split()
            if not parts:
                continue
            index_name = parts[0]
            if not index_name.startswith(index_prefix):
                continue
            segment_counter[index_name] = segment_counter.get(index_name, 0) + 1
            match = date_regex.search(index_name)
            if match:
                segment_dates.add(match.group(1))

        today_str = None
        yest_str = None
        if segment_dates:
            try:
                available_dates = sorted([datetime.strptime(d, "%Y.%m.%d") for d in segment_dates])
                max_date = available_dates[-1]
                today_str = max_date.strftime("%Y.%m.%d")
                yest_str = (max_date - timedelta(days=1)).strftime("%Y.%m.%d")
            except Exception:
                pass

        limit = int(self.fc.get("segment_max_non_recent", 3))
        overs_segment = []
        for index_name, count in segment_counter.items():
            match = date_regex.search(index_name)
            if not match:
                continue
            date_str = match.group(1)
            if today_str and yest_str:
                if date_str in (today_str, yest_str):
                    continue
            else:
                fallback_today = datetime.now().strftime("%Y.%m.%d")
                fallback_yest = (datetime.now() - timedelta(days=1)).strftime("%Y.%m.%d")
                if date_str in (fallback_today, fallback_yest):
                    continue
            if count > limit:
                overs_segment.append(f"{index_name} 行数 {count}")

        if overs_segment:
            self.add_result(Level.WARN, f"{server_name} {ip_addr} segments 行数超过{limit}（非今昨）: " + "；".join(overs_segment))



# ======================================================================== ES N9K 异常日志巡检（通过 Kibana Console 代理）========================================================================
SEV_RE = re.compile(r"%[^-\s]+-(\d)-")
SEV_TEXT = {0:"EMERG",1:"ALERT",2:"CRIT",3:"ERR",4:"WARN",5:"NOTICE",6:"INFO",7:"DEBUG"}
LEVEL_ORDER_ES = ["CRITICAL", "ERROR", "WARN", "OK"]

def _esn9k_min_sev(msg: str) -> Optional[int]:
    if not msg:
        return None
    vals = []
    for m in SEV_RE.finditer(msg):
        try:
            vals.append(int(m.group(1)))
        except Exception:
            pass
    return min(vals) if vals else None

def _esn9k_sev_to_level(min_sev: Optional[int]) -> str:
    if min_sev is None or min_sev >= 5:
        return "OK"
    if min_sev <= 2:
        return "CRITICAL"
    if min_sev == 3:
        return "ERROR"
    if min_sev == 4:
        return "WARN"
    return "OK"

def _esn9k_worse(a: str, b: str) -> str:
    order = {lvl: i for i, lvl in enumerate(LEVEL_ORDER_ES)}
    return a if order.get(a, 99) < order.get(b, 99) else b

# ===== ESN9K ignore list support (v5) =====
def _esn9k_load_ignores() -> dict:
    """
    Load ignore patterns for ES N9K messages.
    Priority: esn9k.ignore_file -> settings.ignore_alarm_file -> "YAML/Ignore_alarm.yaml"
    Schema:
    esn9k:
      message_contains: [ "substring1", "substring2" ]
      message_regex:    [ "regex1", "regex2" ]
    """
    try:
        es_cfg = CONFIG.get("esn9k", {}) or {}
        st_cfg = CONFIG.get("settings", {}) or {}
        path = es_cfg.get("ignore_file") or st_cfg.get("ignore_alarm_file") or "YAML/Ignore_alarm.yaml"
        if not os.path.exists(path):
            return {"contains": [], "regex": []}
        import yaml as _yaml, re as _re
        with open(path, "r", encoding="utf-8") as f:
            data = _yaml.safe_load(f) or {}
        node = (data.get("esn9k") or {}) if isinstance(data, dict) else {}
        contains = node.get("message_contains") or []
        regexps  = node.get("message_regex") or []
        compiled = []
        for pat in regexps:
            try:
                compiled.append(_re.compile(pat, _re.IGNORECASE))
            except Exception:
                pass
        def _norm(s: str) -> str:
            return " ".join(str(s).split()).lower()
        return {"contains": [_norm(x) for x in contains], "regex": compiled}
    except Exception:
        return {"contains": [], "regex": []}

_ESN9K_IGNORES = _esn9k_load_ignores()

def _esn9k_should_ignore(msg: str) -> bool:
    if not msg:
        return False
    try:
        norm = " ".join(msg.split()).lower()
        for sub in _ESN9K_IGNORES.get("contains", []) or []:
            if sub and sub in norm:
                return True
        for rgx in _ESN9K_IGNORES.get("regex", []) or []:
            if rgx.search(msg):
                return True
    except Exception:
        return False
    return False
# ======================================================================== end ignore support ========================================================================

def _esn9k_pick_kibana(session: requests.Session) -> tuple[str, str]:
    cfg = CONFIG.get("esn9k", {})
    bases: dict = cfg.get("kibana_bases", {})
    last_err = None
    for name, base in (bases or {}).items():
        try:
            url = f"{base}/api/status"
            r = session.get(url, headers={"kbn-xsrf": "true"}, timeout=10)
            r.raise_for_status()
            return name, base
        except Exception as e:
            last_err = e
    raise RuntimeError(f"没有可用的 Kibana（esn9k.kibana_bases）。最后错误: {last_err}")

def _esn9k_kbn_version(session: requests.Session, base: str) -> str:
    url = f"{base}/api/status"
    headers = {"kbn-xsrf": "true"}
    r = session.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    try:
        data = r.json()
        ver = data.get("version", {}).get("number")
        if ver:
            return ver
    except Exception:
        pass
    return r.headers.get("kbn-version") or r.headers.get("x-kibana-version") or "7.17.0"

def _esn9k_kbn_proxy(session: requests.Session, base: str, kbn_version: str, method: str, path: str, body: dict | None):
    from urllib.parse import quote as _quote
    q_path = _quote(path, safe="")
    url = f"{base}/api/console/proxy?method={method}&path={q_path}"
    headers = {"kbn-xsrf": "true", "kbn-version": kbn_version, "Content-Type": "application/json"}
    data = json.dumps(body) if body is not None else None
    r = session.post(url, headers=headers, data=data, timeout=120)
    r.raise_for_status()
    return r.json()

def run_esn9k_probe(target: tuple[str, str] | None = None) -> dict:
    global _ESN9K_IGNORES
    _ESN9K_IGNORES = _esn9k_load_ignores()
    cfg = CONFIG.get("esn9k", {}) or {}
    index_pattern: str = cfg.get("index_pattern", "*-n9k-*-*")
    time_field: str = cfg.get("time_field", "@timestamp")
    time_gte: str = cfg.get("time_gte", "now-7d")
    time_lt:  str = cfg.get("time_lt", "now")
    page_size: int = int(cfg.get("page_size", 1000))
    scroll_keepalive: str = cfg.get("scroll_keepalive", "2m")

    session = requests.Session()
    # 如需鉴权，可从环境变量读取 KIBANA_USER/PASS（保密）
    user = os.getenv("KIBANA_USER")
    pwd = os.getenv("KIBANA_PASS")
    if user and pwd:
        session.auth = (user, pwd)

    if target is not None:
        kbn_name, base = target
    else:
        kbn_name, base = _esn9k_pick_kibana(session)
    kbn_version = _esn9k_kbn_version(session, base)

    base_query = {
        "size": page_size,
        "_source": [time_field, "message"],
        "sort": [{time_field: "asc"}],
        "track_total_hits": True,
        "query": {"range": { time_field: { "gte": time_gte, "lt": time_lt } } }
    }

    first_path = f"/{index_pattern}/_search?scroll={scroll_keepalive}"
    first = _esn9k_kbn_proxy(session, base, kbn_version, "POST", first_path, base_query)
    scroll_id = first.get("_scroll_id")

    if not scroll_id:
        return {"kibana": {"name": kbn_name, "base": base, "version": kbn_version},
                "scanned": 0, "matched": 0, "worst_level": "OK", "samples": [],
                "note": "未获得 _scroll_id，可能索引无数据或权限不足"}

    scanned = matched = 0
    worst = "OK"
    samples = []
    sample_limit = 10

    try:
        resp = first
        while True:
            hits = resp.get("hits", {}).get("hits", [])
            if not hits:
                break
            for h in hits:
                scanned += 1
                src = h.get("_source", {}) or {}
                ts = src.get(time_field, "")
                msg = src.get("message", "") or ""
                if _esn9k_should_ignore(msg):
                    continue
                sev = _esn9k_min_sev(msg)
                if sev is not None and sev <= 4:
                    matched += 1
                    lv = _esn9k_sev_to_level(sev)
                    worst = _esn9k_worse(worst, lv)
                    if len(samples) < sample_limit:
                        samples.append({"timestamp": ts, "severity": sev,
                                        "severity_text": SEV_TEXT.get(sev,"?"),
                                        "level": lv, "message": msg[:800]})
            resp = _esn9k_kbn_proxy(session, base, kbn_version, "POST", "/_search/scroll",
                                    {"scroll": scroll_keepalive, "scroll_id": scroll_id})
            scroll_id = resp.get("_scroll_id")
            if not scroll_id:
                break
    finally:
        try:
            _esn9k_kbn_proxy(session, base, kbn_version, "DELETE", "/_search/scroll", {"scroll_id": [scroll_id]})
        except Exception:
            pass

    return {"kibana": {"name": kbn_name, "base": base, "version": kbn_version},
            "scanned": scanned, "matched": matched, "worst_level": worst, "samples": samples}

class ESN9KLogInspectTask(BaseTask):
    def __init__(self):
        super().__init__("ES 服务器CS-N9K异常日志巡检")

    def items(self):
        bases = (CONFIG.get("esn9k", {}) or {}).get("kibana_bases", {}) or {}
        return list(bases.items())

    def run_single(self, item):
        try:
            name, base = item
            result = run_esn9k_probe((name, base))
            worst = result.get("worst_level", "OK")
            level_map = {"CRITICAL": Level.CRIT, "ERROR": Level.ERROR, "WARN": Level.WARN, "OK": Level.OK}
            level = level_map.get(worst, Level.OK)
            msg = (f"{name}({base}) 扫描={result['scanned']} 命中(sev<=4)={result['matched']} 等级={worst}")
            self.add_result(level, msg, {"samples": result.get("samples", [])})
        except Exception as exc:
            self.add_result(Level.ERROR, f"ESN9K 巡检失败: {exc}")

# ======================================================================== 主调度器 ========================================================================
def main():
    today = datetime.now().strftime("%Y%m%d")
    settings = (CONFIG.get("settings") or {})
    base_log_dir = settings.get("log_dir", "LOG")
    log_dir = os.path.join(base_log_dir, today)
    os.makedirs(log_dir, exist_ok=True)
    report_dir = settings.get("report_dir", "REPORT")
    os.makedirs(report_dir, exist_ok=True)
    daily_report = os.path.join(report_dir, f"{today}巡检日报.log")

    tasks: list[BaseTask] = [
        FXOSWebTask(),
        MirrorFortiGateTask(),
        OxidizedTask(log_dir=log_dir),
        LogstashServerTask(),
        ElasticsearchServerTask(),
        ESN9KLogInspectTask(),
        FlowServerTask(),
    ]
    
    with open(daily_report, "a", encoding="utf-8") as report:
        all_summary: Dict[str, Any] = {}
        total_counter = {"OK": 0, "WARN": 0, "CRIT": 0, "ERROR": 0}

        for task in tasks:
            header = f"\n=== 执行 {task.name} ==="
            if SHOW_PROGRESS:
                tqdm.write(header)
            else:
                print(header, flush=True)

            task.run()

            level_count = {"OK": 0, "WARN": 0, "CRIT": 0, "ERROR": 0}
            for result in task.results:
                level_count[result.level] = level_count.get(result.level, 0) + 1
                total_counter[result.level] = total_counter.get(result.level, 0) + 1

            report.write(f"{task.name}：CRIT {level_count['CRIT']}, WARN {level_count['WARN']}, ERROR {level_count['ERROR']}, OK {level_count['OK']}\n")
            for result in task.results:
                if result.level != "OK":
                    report.write(f"  - [{result.level}] {result.message}\n")

            task_log_path = os.path.join(log_dir, f"{safe_filename(task.name)}-{today}.log")
            with open(task_log_path, "w", encoding="utf-8") as detail_file:
                for result in task.results:
                    line = result.message
                    if result.meta:
                        line += f" | {json.dumps(result.meta, ensure_ascii=False)}"
                    detail_file.write(f"[{result.level}] {line}\n")

            all_summary[task.name] = {
                "summary": level_count,
                "results": [asdict(result) for result in task.results],
            }

        report.write("\n=== 巡检总汇 ===\n")
        report.write(f"CRIT {total_counter['CRIT']}, WARN {total_counter['WARN']}, ERROR {total_counter['ERROR']}, OK {total_counter['OK']}\n")
        report.write(f"{today} 全部任务完成\n")
        report.write(f"日志目录: {log_dir}\n")
if __name__ == "__main__":
    main()