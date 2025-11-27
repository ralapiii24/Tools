# -*- coding: utf-8 -*-
"""
Updater (robust FTP dir download + merge apply + preserve paths)

修复/增强：
1) 递归下载：先 cwd 到目录，用相对名 nlst；判断目录/文件采用 cwd/size 双保险；
2) 应用更新：按文件“合并覆盖”，相对路径白名单保护（Main.py / Updater.py / YAML/Version.yaml）；
3) 远端缺少 Version.yaml 时，本地 YAML/Version.yaml 不会丢失。
"""

import ftplib
import io
import os
import sys
import time
import shutil
import tempfile
import subprocess
import hashlib
import re
from pathlib import Path

import yaml
import runpy

# ----------------------------
# 常量与全局
# ----------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_ROOT / "YAML" / "Version.yaml"
LOCAL_VERSION_FILE = CONFIG_FILE
CORE_ENTRY = PROJECT_ROOT / "APP" / "Core.py"

# 使用“相对路径”保护：相对于项目根目录 PROJECT_ROOT
PRESERVE_PATHS = {
    "Main.py",
    "Updater.py",
    "YAML/Version.yaml",
}

def _is_preserved(rel_path: str) -> bool:
    rel = str(rel_path).replace("\\", "/").lstrip("./")
    return rel in PRESERVE_PATHS

_VERBOSE_LINES = []
def _v(msg: str):
    _VERBOSE_LINES.append(msg)
    try:
        print(msg, flush=True)
    except Exception:
        pass

# ----------------------------
# 配置读取
# ----------------------------
def _load_cfg():
    if not CONFIG_FILE.exists():
        raise SystemExit("缺少配置文件: " + str(CONFIG_FILE))
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    return cfg.get("updater") or {}

def _require(up_cfg, key, env_key=None):
    val = up_cfg.get(key, None)
    if (val is None or str(val).strip() == "") and env_key:
        env_val = os.getenv(env_key, "").strip()
        if env_val:
            return env_val
    if val is None or str(val).strip() == "":
        raise SystemExit("配置缺失: updater." + key)
    return val

up_cfg = _load_cfg()

FTP_HOST = _require(up_cfg, "ftp_host", "UPDATER_FTP_HOST")
FTP_USER = _require(up_cfg, "ftp_user", "UPDATER_FTP_USER")
FTP_PASS = _require(up_cfg, "ftp_pass", "UPDATER_FTP_PASS")
FTP_TIMEOUT = int(_require(up_cfg, "ftp_timeout", "UPDATER_FTP_TIMEOUT"))
FTP_READ_TIMEOUT = int(up_cfg.get("ftp_read_timeout", up_cfg.get("ftp_timeout", 30)))
REMOTE_LATEST_FILE = _require(up_cfg, "remote_latest_file", "UPDATER_REMOTE_LATEST_FILE")
REMOTE_VERSIONS_ROOT = _require(up_cfg, "remote_versions_root", "UPDATER_REMOTE_VERSIONS_ROOT")

FTP_PORT = int(up_cfg.get("ftp_port", 21))
FTP_PASSIVE = bool(up_cfg.get("ftp_passive", True))
RETRY_TIMES = int(up_cfg.get("retry_times", 3))
RETRY_DELAY_SEC = float(up_cfg.get("retry_delay_sec", 2.0))

# ----------------------------
# 版本处理
# ----------------------------
def _version_tuple(s):
    nums = re.findall(r"\d+", str(s))
    if not nums:
        return (0,)
    return tuple(int(x) for x in nums)

def _compare_versions(a, b):
    ta, tb = _version_tuple(a), _version_tuple(b)
    maxlen = max(len(ta), len(tb))
    ta = ta + (0,) * (maxlen - len(ta))
    tb = tb + (0,) * (maxlen - len(tb))
    if ta > tb:
        return 1
    if ta < tb:
        return -1
    return 0

def _read_local_version():
    try:
        with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        return str(data.get("version") or "").strip()
    except Exception:
        return None

# ----------------------------
# FTP helpers
# ----------------------------
def _keepalive(ftp):
    try:
        ftp.voidcmd("NOOP")
    except Exception:
        pass

def _ftp_connect(passive=True):
    _v("[Updater] 尝试连接 FTP: " + str(FTP_HOST) + ":" + str(FTP_PORT) + " (passive=" + str(passive) + ")")
    last_err = None
    for enc in ("utf-8", "gbk", "latin-1"):
        _v("[Updater] 使用响应编码尝试: " + enc)
        try:
            ftp = ftplib.FTP()
            ftp.encoding = enc
            ftp.connect(str(FTP_HOST), int(FTP_PORT), timeout=int(FTP_TIMEOUT))
            ftp.login(str(FTP_USER), str(FTP_PASS))
            _v("[Updater] 登录成功")
            ftp.set_pasv(bool(passive))
            _v("[Updater] 已设置 PASV=" + str(passive))
            try:
                if getattr(ftp, "sock", None):
                    ftp.sock.settimeout(int(FTP_READ_TIMEOUT))
            except Exception:
                pass
            _v("[Updater] 连接建立完成")
            return ftp
        except UnicodeDecodeError as e:
            last_err = e
            continue
        except Exception as e:
            last_err = e
    raise SystemExit(f"[Updater] FTP 连接失败: {last_err}")

def _ftp_pwd(ftp):
    try:
        return ftp.pwd()
    except Exception:
        return "/"

def _ftp_join(a, b):
    a = (a or "/").rstrip("/")
    b = (b or "").lstrip("/")
    return (a + "/" + b) if a else ("/" + b)

def _ftp_cwd(ftp, path):
    ftp.cwd(path)

def _ftp_nlst_rel(ftp):
    """
    已经 cwd 到目标目录后，返回当前目录下的相对名字列表（不含路径）。
    """
    _keepalive(ftp)
    try:
        names = ftp.nlst()
    except Exception:
        names = []
    # 去掉 '.' '..' 空名，去重
    out, seen = [], set()
    for n in names:
        n = (n or "").strip()
        if not n or n in (".", ".."):
            continue
        # 有些服务器会返回绝对路径，取最后一段
        n = n.replace("\\", "/").split("/")[-1]
        if n in seen:
            continue
        seen.add(n)
        out.append(n)
    return out

def _ftp_is_dir_safely(ftp, child_name):
    """
    尝试判断 child_name（相对当前目录）是目录还是文件：
    1) 先尝试 cwd(child) 成功 => 目录（随后切回）
    2) 再尝试 size(child) 成功 => 文件
    3) 都失败：再试 LIST 判断；最后保守认为“文件”
    """
    cwd0 = _ftp_pwd(ftp)
    try:
        ftp.cwd(child_name)
        isdir = True
    except Exception:
        isdir = False
    finally:
        try:
            ftp.cwd(cwd0)
        except Exception:
            pass
    if isdir:
        return True

    try:
        sz = ftp.size(child_name)
        if isinstance(sz, int) or (isinstance(sz, (str, bytes)) and str(sz).isdigit()):
            return False  # 有 size 的一般是文件
    except Exception:
        pass

    # 尝试 LIST 判断
    try:
        lines = []
        ftp.retrlines("LIST " + child_name, lines.append)
        # 简单启发式：以 'd' 开头的类 UNIX 列表为目录
        if lines:
            first = lines[0].lower()
            if first.startswith("d"):
                return True
    except Exception:
        pass

    # 保守认为文件
    return False

def _ftp_download_file(ftp, remote_file_abs, local_file):
    _v(f"[Updater] 下载文件: {remote_file_abs} -> {local_file}")
    local_file.parent.mkdir(parents=True, exist_ok=True)
    with open(local_file, "wb") as f:
        ftp.retrbinary("RETR " + str(remote_file_abs), f.write)

def _ftp_download_dir_recursive(ftp, remote_dir_abs, local_dir, downloaded_list):
    """
    进入 remote_dir_abs，列出相对名，逐个判断目录/文件并递归下载。
    """
    local_dir.mkdir(parents=True, exist_ok=True)
    _v("[Updater] 列举目录: " + str(remote_dir_abs))

    # 记录与还原当前目录
    cwd0 = _ftp_pwd(ftp)
    ftp.cwd(remote_dir_abs)
    try:
        names = _ftp_nlst_rel(ftp)
        dirs, files = [], []
        for name in names:
            if _ftp_is_dir_safely(ftp, name):
                dirs.append(name)
            else:
                files.append(name)
        _v(f"[Updater] 目录 {remote_dir_abs} -> 子目录: {len(dirs)} / 文件: {len(files)}")

        # 下载文件
        for name in files:
            remote_abs = _ftp_join(remote_dir_abs, name)
            target_path = local_dir / name
            _ftp_download_file(ftp, remote_abs, target_path)
            downloaded_list.append(target_path)

        # 递归子目录
        for name in dirs:
            sub_remote_abs = _ftp_join(remote_dir_abs, name)
            sub_local = local_dir / name
            _ftp_download_dir_recursive(ftp, sub_remote_abs, sub_local, downloaded_list)

    finally:
        # 切回
        try:
            ftp.cwd(cwd0)
        except Exception:
            pass

def _remote_md5_via_command(ftp, remote_path):
    for cmd in ("XMD5", "MD5", "HASH MD5"):
        try:
            resp = ftp.sendcmd(cmd + " " + str(remote_path))
            parts = resp.strip().split()
            for tok in reversed(parts):
                tok = tok.strip().lower()
                if re.fullmatch(r"[0-9a-f]{32}", tok):
                    return tok
        except Exception:
            continue
    return None

def _remote_md5(ftp, remote_path):
    return _remote_md5_via_command(ftp, remote_path)

def _ftp_read_text_file(ftp, remote_path):
    _keepalive(ftp)
    buf = io.BytesIO()
    ftp.retrbinary("RETR " + str(remote_path), buf.write)
    data = buf.getvalue()
    try:
        return data.decode(ftp.encoding or "utf-8", "ignore")
    except Exception:
        return data.decode("utf-8", "ignore")

def _local_md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

# ----------------------------
# 应用远端版本（合并覆盖 + 保留白名单）
# ----------------------------
def _apply_remote_version_payload(payload_root: Path):
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    backup_root = PROJECT_ROOT / ("_backup_" + timestamp)
    backup_root.mkdir(parents=True, exist_ok=True)

    _v("[Updater] 开始应用远端版本(合并覆盖): " + str(payload_root) + " -> " + str(PROJECT_ROOT))

    # 1) 确保目录结构存在（不整体替换）
    for d in sorted([p for p in payload_root.rglob("*") if p.is_dir()], key=lambda p: len(p.parts)):
        rel = d.relative_to(payload_root)
        if _is_preserved(str(rel)):
            continue
        target_dir = PROJECT_ROOT / rel
        target_dir.mkdir(parents=True, exist_ok=True)

    # 2) 逐文件备份并覆盖
    for f in sorted([p for p in payload_root.rglob("*") if p.is_file()], key=lambda p: len(p.parts)):
        rel = f.relative_to(payload_root)
        rel_str = str(rel).replace("\\", "/")
        if _is_preserved(rel_str):
            _v("[Updater] 跳过保留文件: " + rel_str)
            continue

        target = PROJECT_ROOT / rel
        target.parent.mkdir(parents=True, exist_ok=True)

        if target.exists():
            backup_target = backup_root / rel
            backup_target.parent.mkdir(parents=True, exist_ok=True)
            try:
                target.rename(backup_target)  # 逐文件备份
            except Exception as e:
                _v("[Updater] 备份失败(继续覆盖): " + str(rel_str) + " -> " + str(e))

        # 尝试移动；跨设备失败则复制
        try:
            f.rename(target)
        except Exception:
            try:
                shutil.copy2(str(f), str(target))
            except Exception as e:
                _v("[Updater] 覆盖失败: " + str(rel_str) + " -> " + str(e))
            try:
                f.unlink()
            except Exception:
                pass

    # 清理远端缓存目录
    try:
        shutil.rmtree(payload_root, ignore_errors=True)
    except Exception:
        pass

    # 如果备份目录是空的就删掉；否则保留
    try:
        if backup_root.exists() and not any(backup_root.rglob("*")):
            shutil.rmtree(backup_root, ignore_errors=True)
    except Exception:
        pass

def _write_upgrade_log(latest_ver):
    base_dir = PROJECT_ROOT / "UPGRADELOG"
    base_dir.mkdir(parents=True, exist_ok=True)
    log_path = base_dir / ("upgrade_" + str(latest_ver) + ".log")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("版本: " + str(latest_ver) + "\n")
        f.write("时间: " + time.strftime("%Y-%m-%d %H:%M:%S") + "\n")
        f.write("\n--- 详细过程 ---\n")
        for line in _VERBOSE_LINES:
            try:
                f.write(line + "\n")
            except Exception:
                pass

# ----------------------------
# 运行本地 Core（无论是否更新）
# ----------------------------
def _run_local_core():
    print("=== 开始运行巡检! ===")
    time.sleep(5)
    if not (CORE_ENTRY.exists() and CORE_ENTRY.is_file()):
        print("=== 未找到本地巡检主文件Core.py ===")
        return
    args = [sys.executable, "-u", str(CORE_ENTRY), *sys.argv[1:]]
    try:
        if os.name == "nt":
            subprocess.call(args, cwd=str(PROJECT_ROOT))
            return
        else:
            os.execv(sys.executable, args)
    except Exception:
        runpy.run_path(str(CORE_ENTRY), run_name="__main__")
        return

# ----------------------------
# 主流程
# ----------------------------
def check_update_then_run():
    print("=== 正在连接版本服务器... ===")
    local_ver = _read_local_version() or "未知"
    latest_ver = "未知"

    try:
        ftp = _ftp_connect(passive=FTP_PASSIVE)
    except SystemExit:
        print("=== 版本服务器连接失败，不进行版本更新，执行本地版本巡检 ===")
        _run_local_core()
        return
    except Exception:
        print("=== 版本服务器连接失败，不进行版本更新，执行本地版本巡检 ===")
        _run_local_core()
        return

    try:
        print("=== 连接建立完成，登录成功! ===")
        latest_text = _ftp_read_text_file(ftp, str(REMOTE_LATEST_FILE))
        latest_info = yaml.safe_load(latest_text) or {}
        latest_ver = str(latest_info.get("latest") or "").strip() or "未知"
        print("=== 正在读取最新版本文件，本地版本" + str(local_ver) + "，服务器版本" + str(latest_ver) + " ===")

        if latest_ver == "未知":
            print("=== 无更新 ===")
            _run_local_core()
            return

        remote_ver_dir = str(REMOTE_VERSIONS_ROOT).rstrip("/") + "/" + str(latest_ver)
        cmp_val = _compare_versions(latest_ver, local_ver)

        need_refetch = False
        if cmp_val == 0:
            if not (CORE_ENTRY.exists() and CORE_ENTRY.is_file()):
                need_refetch = True
        else:
            need_refetch = True

        if not need_refetch:
            print("=== 无需更新（版本一致） ===")
            _run_local_core()
            return

        with tempfile.TemporaryDirectory() as td:
            tmp_root = Path(td)
            payload_root = tmp_root / "payload"
            payload_root.mkdir(parents=True, exist_ok=True)

            print("=== 正在下载版本目录: " + remote_ver_dir + " ===")
            downloaded = []
            _ftp_download_dir_recursive(ftp, remote_ver_dir, payload_root, downloaded)

            print("=== 下载完成，开始应用更新 ===")
            _apply_remote_version_payload(payload_root)

        update_label = f"=== 版本更新完成: {local_ver} -> {latest_ver} ==="
        print(update_label)
        _write_upgrade_log(latest_ver)

    except Exception as e:
        print("=== 版本更新过程出现异常，不进行版本更新，执行本地版本巡检 ===")
        _v(f"[Updater] 异常: {e}")
    finally:
        try:
            ftp.quit()
        except Exception:
            pass

    _run_local_core()

if __name__ == "__main__":
    check_update_then_run()
