# 更新器：从FTP服务器下载并安装更新包

# 导入标准库
import datetime
import ftplib
import io
import json
import os
import posixpath as _pp
import runpy
import shutil
import sys
import tempfile
import time
import traceback
from pathlib import Path
from typing import Dict, List, Any, Optional, Set

# 导入第三方库
import yaml

# 导入本地应用
# (无本地应用依赖)

# 路径配置
PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOCAL_VERSION_FILE = PROJECT_ROOT / "YAML" / "Version.yaml"
CONFIG_FILE = PROJECT_ROOT / "YAML" / "Config.yaml"
CORE_ENTRY = PROJECT_ROOT / "APP" / "Core.py"

# 加载本地版本配置
def _load_version_yaml() -> Dict[str, Any]:
    if not LOCAL_VERSION_FILE.exists():
        return {}
    try:
        with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    except Exception:
        return {}

# 加载主配置文件
def _load_config_yaml() -> Dict[str, Any]:
    if not CONFIG_FILE.exists():
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    except Exception:
        return {}

_CFG = _load_version_yaml()
_UPD: Dict[str, Any] = _CFG.get("updater", _CFG) if isinstance(_CFG, dict) else {}

# 加载主配置
_MAIN_CFG = _load_config_yaml()

# 控制台日志开关
VERBOSE: bool = bool(_UPD.get("verbose", False))
WRITE_UPGRADE_LOG: bool = bool(_UPD.get("write_upgrade_log", True))

# FTP配置
FTP_HOST = str(_UPD.get("ftp_host", _UPD.get("host", "")))
FTP_PORT = int(_UPD.get("ftp_port", _UPD.get("port", 21)))
FTP_USER = str(_UPD.get("ftp_user", _UPD.get("user", "")))
FTP_PASS = str(_UPD.get("ftp_pass", _UPD.get("password", "")))
FTP_PASSIVE = bool(_UPD.get("ftp_passive", _UPD.get("passive", False)))

# 超时配置
FTP_TIMEOUT = int(_UPD.get("ftp_timeout", _UPD.get("timeout", 10)))
FTP_READ_TIMEOUT = int(_UPD.get("ftp_read_timeout", _UPD.get("read_timeout", 30)))

# 远程版本目录
REMOTE_VERSIONS_ROOT = str(_UPD.get("remote_versions_root", "/"))
REMOTE_LATEST_FILE = str(_UPD.get("remote_latest_file", "latest.yaml"))  # 可为绝对路径（以 / 开头）

# 重试配置
RETRY_TIMES = int(_UPD.get("retry_times", 3))
RETRY_DELAY_SEC = int(_UPD.get("retry_delay_sec", 2))

# 日志控制
_VERBOSE_LINES: List[str] = []

# 受控调试输出：verbose=true才打印，但总会进入_VERBOSE_LINES以便落盘
def _v(msg: str) -> None:
    try:
        _VERBOSE_LINES.append(str(msg))
    except Exception:
        pass
    if VERBOSE:
        try:
            print(msg, flush=True)
        except Exception:
            pass

UPGRADELOG_DIR = PROJECT_ROOT / "UPGRADELOG"

# 写入升级日志
def _write_upgrade_log(tag: str = "", local_ver: Optional[str] = None, latest_ver: Optional[str] = None) -> None:
    if not WRITE_UPGRADE_LOG:
        return
    try:
        # 创建升级日志目录（如果目录已存在则不报错）
        UPGRADELOG_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_tag = (tag or "run").replace(" ", "_").replace("/", "_")
        fp = UPGRADELOG_DIR / f"updater_{ts}_{safe_tag}.log"
        header = []
        header.append(f"[time] {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
        if local_ver is not None:
            header.append(f"[local_version] v{str(local_ver).lstrip('vV')}")
        if latest_ver is not None:
            header.append(f"[remote_version] v{str(latest_ver).lstrip('vV')}")
        header.append("")  # 空行
        with open(fp, "w", encoding="utf-8") as f:
            f.write("\n".join(header))
            if _VERBOSE_LINES:
                f.write("\n".join(_VERBOSE_LINES))
        if VERBOSE:
            print(f"[Updater] 写入升级日志: {fp}", flush=True)
    except Exception as e:
        if VERBOSE:
            print(f"[Updater] 写入升级日志失败: {e}", flush=True)

# 工具函数
# 读取本地版本号
def _read_local_version() -> str:
    v = str(_CFG.get("version", "")).strip()
    return v or "0"

# 比较版本号：返回>0表示a>b，0表示相等，<0表示a<b
def _compare_versions(a: str, b: str) -> int:

    # 标准化版本号
    def norm(s: str) -> List[int]:
        s = str(s).strip().lstrip("vV")
        if not s:
            return [0]
        out = []
        for PART in s.split("."):
            try:
                out.append(int(PART))
            except Exception:
                out.append(0)
        return out

    aa, bb = norm(a), norm(b)
    for INDEX in range(max(len(aa), len(bb))):
        va = aa[INDEX] if INDEX < len(aa) else 0
        vb = bb[INDEX] if INDEX < len(bb) else 0
        if va != vb:
            return 1 if va > vb else -1
    return 0

# 标准化远程路径：POSIX规范化远端路径，去掉'.'、'..'、重复斜杠，并确保以'/'开头
def _norm_remote(path: str) -> str:
    if not path:
        return "/"
    p = path.replace("\\", "/")
    p = _pp.normpath(p)
    if not p.startswith("/"):
        p = "/" + p
    return p

# 连接远程路径：拼接远端路径，若path是以/开头的绝对路径，则直接返回path（再规范化）
def _join_remote(root: str, path: str) -> str:
    if not path:
        return _norm_remote(root or "/")
    if path.startswith("/"):
        return _norm_remote(path)
    r = (root or "/").rstrip("/")
    return _norm_remote(f"{r}/{path}")

# 建立FTP连接
def _ftp_connect(passive: bool = True) -> ftplib.FTP:
    # 带重试
    last_err = None
    for attempt in range(1, max(1, RETRY_TIMES) + 1):
        _v(f"[Updater] 尝试连接 FTP: {FTP_HOST}:{FTP_PORT} (passive={passive}) [attempt {attempt}/{RETRY_TIMES}]")
        for enc in ("utf-8", "gbk", "latin-1"):
            _v(f"[Updater] 使用响应编码尝试: {enc}")
            try:
                ftp = ftplib.FTP()
                ftp.encoding = enc
                ftp.connect(FTP_HOST, FTP_PORT, timeout=FTP_TIMEOUT)
                ftp.login(FTP_USER, FTP_PASS)
                _v("[Updater] 登录成功")
                ftp.set_pasv(bool(passive))
                _v(f"[Updater] 已设置 PASV={passive}")
                try:
                    if getattr(ftp, "sock", None):
                        ftp.sock.settimeout(FTP_READ_TIMEOUT)
                except Exception:
                    pass
                return ftp
            except Exception as e:
                last_err = e
        # 这一轮所有编码都失败，延时后重试
        if attempt < RETRY_TIMES:
            time.sleep(max(0, RETRY_DELAY_SEC))
    if last_err:
        raise last_err
    raise SystemExit("无法连接 FTP")

# 从FTP读取二进制数据
def _ftp_read_bytes(ftp: ftplib.FTP, remote_file: str) -> bytes:
    rf = _norm_remote(remote_file)
    _v(f"[Updater] 读取文件: {rf}")
    bio = io.BytesIO()
    ftp.retrbinary(f"RETR {rf}", bio.write)
    return bio.getvalue()

# 从FTP读取文本数据
def _ftp_read_text(ftp: ftplib.FTP, remote_file: str) -> str:
    return _ftp_read_bytes(ftp, remote_file).decode("utf-8", errors="ignore")

# 检查FTP路径是否为目录
def _ftp_is_dir(ftp: ftplib.FTP, path: str) -> bool:
    p = _norm_remote(path)
    cwd = ftp.pwd()
    try:
        ftp.cwd(p)
        ftp.cwd(cwd)
        return True
    except Exception:
        return False

# 列出FTP目录内容：列举base下的条目，返回绝对规范化路径，过滤'.'和'..'，只保留以'base/'开头的条目
def _ftp_list(ftp: ftplib.FTP, base: str) -> List[str]:
    base = _norm_remote(base)
    _v(f"[Updater] 列举目录: {base}")
    try:
        raw = ftp.nlst(base)
    except ftplib.error_perm:
        try:
            raw = ftp.nlst((base.rstrip("/") + "/*").replace("//", "/"))
        except Exception:
            return []
    except Exception:
        return []

    out: List[str] = []
    seen: Set[str] = set()
    prefix = base.rstrip("/") + "/"
    for it in raw or []:
        if not it:
            continue
        # 规范化项，允许服务器返回裸名
        if "/" not in it or it in (".", ".."):
            cand = _norm_remote(prefix + it)
        else:
            cand = _norm_remote(it)
        bn = _pp.basename(cand)
        # 过滤 '.', '..' 与 base 自身
        if bn in (".", ".."):
            continue
        if cand == base:
            continue
        # 强约束：必须在 base 子树内
        if not cand.startswith(prefix):
            continue
        if cand not in seen:
            seen.add(cand)
            out.append(cand)
    return out

# 递归下载FTP目录
def _ftp_download_dir_recursive(ftp: ftplib.FTP, remote_dir: str, local_dir: Path, downloaded: List[str],
                                root_prefix: str, visited: Optional[Set[str]] = None,
                                depth: int = 0, max_depth: int = 32) -> None:
    rdir = _norm_remote(remote_dir)
    root_prefix = _norm_remote(root_prefix).rstrip("/") + "/"
    root_base = root_prefix[:-1]  # e.g. '/v2'
    # 允许进入根目录本身('/v2')，以及其子路径('/v2/...')
    if not (rdir == root_base or rdir.startswith(root_prefix)):
        _v(f"[Updater] 跳过越界目录: {rdir}")
        return
    if visited is None:
        visited = set()
    if rdir in visited:
        _v(f"[Updater] 跳过已访问目录: {rdir}")
        return
    visited.add(rdir)
    if depth > max_depth:
        _v(f"[Updater] 达到最大目录深度 {max_depth}，跳过: {rdir}")
        return

    items = _ftp_list(ftp, rdir)
    # 创建本地目录（如果目录已存在则不报错）
    local_dir.mkdir(parents=True, exist_ok=True)

    for child in items:
        child_norm = _norm_remote(child)
        if not child_norm.startswith(root_prefix):
            _v(f"[Updater] 跳过越界项: {child_norm}")
            continue
        name = _pp.basename(child_norm)
        if _ftp_is_dir(ftp, child_norm):
            _ftp_download_dir_recursive(
                ftp,
                child_norm,
                local_dir / name,
                downloaded,
                root_prefix=root_prefix,
                visited=visited,
                depth=depth + 1,
                max_depth=max_depth,
            )
        else:
            dst = local_dir / name
            _v(f"[Updater] 下载文件: {child_norm} -> {dst}")
            # 创建目标目录（如果目录已存在则不报错）
            dst.parent.mkdir(parents=True, exist_ok=True)
            with open(dst, "wb") as f:
                ftp.retrbinary(f"RETR {child_norm}", f.write)
            downloaded.append(child_norm)

# 创建更新前备份：在升级前创建备份目录
def _create_backup_before_update() -> Optional[Path]:
    try:
        backup_dir = PROJECT_ROOT / f"backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        _v(f"[Updater] 创建升级前备份: {backup_dir}")

        # 备份关键目录和文件
        backup_items = [
            "APP", "TASK", "YAML", "Main.py", "Updater.py", "RuntimeEnvCheck.py"
        ]

        for item in backup_items:
            src = PROJECT_ROOT / item
            if src.exists():
                dst = backup_dir / item
                if src.is_dir():
                    shutil.copytree(src, dst, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
                else:
                    # 创建目标目录（如果目录已存在则不报错）
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, dst)

        _v(f"[Updater] 备份完成: {backup_dir}")
        return backup_dir
    except Exception as e:
        _v(f"[Updater] 备份失败: {e}")
        return None

# 应用远程版本更新
def _apply_remote_version_payload(payload_root: Path) -> None:
    for root, dirs, files in os.walk(payload_root):
        rel = Path(root).relative_to(payload_root)
        target_dir = PROJECT_ROOT / rel
        # 创建目标目录（如果目录已存在则不报错）
        target_dir.mkdir(parents=True, exist_ok=True)
        for fn in files:
            src = Path(root) / fn
            dst = target_dir / fn
            shutil.copy2(src, dst)

# 更新本地版本号
def _bump_local_version_yaml(local_yaml: Path, new_version: str) -> None:
    try:
        data = {}
        if local_yaml.exists():
            with open(local_yaml, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
        # 保留 updater 节
        if "updater" in _CFG:
            data["updater"] = _CFG["updater"]
        data["version"] = str(new_version)
        with open(local_yaml, "w", encoding="utf-8") as f:
            yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)
    except Exception as e:
        _v(f"[Updater] 更新本地版本号失败: {e}")

# 检查版本兼容性：检查版本兼容性和特殊升级需求
def _check_version_compatibility(local_ver: str, target_ver: str) -> bool:
    # 检查是否需要TASK目录（支持TASK架构的版本）
    # 使用更通用的版本检查方式
    try:
        # 提取主版本号进行比较
        target_major = int(target_ver.lstrip('vV').split('.')[0]) \
            if target_ver.lstrip('vV').split('.')[0].isdigit() else 0

        # V6及以上版本需要TASK目录
        if target_major >= 6:
            return _check_task_directory_required(local_ver, target_ver)
    except Exception:
        # 如果版本解析失败，回退到字符串匹配
        if target_ver.startswith('v6') or target_ver.startswith('v7') or target_ver.startswith('v8'):
            return _check_task_directory_required(local_ver, target_ver)

    # 未来可以添加其他版本的兼容性检查
    return False

# 检查是否需要TASK目录：检查是否需要TASK目录（支持TASK架构的版本都需要TASK目录）
def _check_task_directory_required(local_ver: str, target_ver: str) -> bool:
    # 任何版本升级到支持TASK架构的版本都需要检查TASK目录
    task_dir = PROJECT_ROOT / "TASK"
    task_init = PROJECT_ROOT / "TASK" / "__init__.py"
    
    # 检查基础文件
    task_base = PROJECT_ROOT / "TASK" / "TaskBase.py"

    # 检查基础文件
    if not task_dir.exists() or not task_init.exists() or not task_base.exists():
        _v(f"[Updater] 升级检测：TASK目录不完整，需要升级")
        return True

    # 检查所有任务文件是否完整（动态检测__init__.py中的所有任务）
    try:
        import sys
        sys.path.insert(0, str(PROJECT_ROOT))
        from TASK import __all__
        _v(f"[Updater] 升级检测：发现{len(__all__)}个任务模块")

        # 动态导入所有任务模块
        for task_name in __all__:
            task_module = __import__(f'TASK.{task_name}', fromlist=[task_name])
            getattr(task_module, task_name)

        _v(f"[Updater] 升级检测：所有{len(__all__)}个任务模块导入成功")
    except ImportError as e:
        _v(f"[Updater] 升级检测：任务模块导入失败，需要升级: {e}")
        return True
    except Exception as e:
        _v(f"[Updater] 升级检测：任务模块检查异常，需要升级: {e}")
        return True
    return False

# 检查本地缺失的必要文件
def _local_missing_required(latest_info: Optional[Dict[str, Any]] = None) -> List[str]:
    req = []
    if isinstance(latest_info, dict):
        req = latest_info.get("required_files") or latest_info.get("required") or []
    if not req:
        # 支持TASK架构的版本需要检查TASK目录的关键文件
        req = [
            "APP/Core.py",
            "Main.py",
            "Updater.py",
            "YAML/Version.yaml",
            "TASK/__init__.py",  # 检查TASK目录是否存在
            "TASK/TaskBase.py"   # 检查关键任务基础文件
        ]
    missing = []
    for rel in req:
        try:
            p = (PROJECT_ROOT / str(rel)).resolve()
            if (not p.exists()) or (p.is_file() and p.stat().st_size == 0):
                missing.append(str(rel))
        except Exception:
            missing.append(str(rel))
    return missing

# 防重复提示
_PRINTED_LOGIN_SUCCESS = False

# 运行本地Core程序
def _run_local_core() -> bool:
    print("=== 开始运行巡检! ===", flush=True)
    # 确保在项目根目录执行，并可相对导入
    try:
        os.chdir(str(PROJECT_ROOT))
    except Exception:
        pass
    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))

    if not (CORE_ENTRY.exists() and CORE_ENTRY.is_file()):
        print("=== 未找到本地巡检主文件Core.py ===", flush=True)
        return False

    try:
        # 以 __main__ 方式运行，兼容 if __name__ == '__main__' 分支
        runpy.run_path(str(CORE_ENTRY), run_name="__main__")
        return True
    except SystemExit as e:
        code = getattr(e, "code", 1)
        if code in (0, None):
            return True
        print("=== 巡检主程序 Core.py 运行失败（非零退出码） ===", flush=True)
        _v(f"[Updater] Core.py SystemExit: {e}")
        return False
    except Exception as e:
        trace = traceback.format_exc()
        print("=== 巡检主程序 Core.py 运行失败，请查看 UPGRADELOG 或将 verbose:true 后重试 ===", flush=True)
        _v(f"[Updater] 本地 Core.py 执行失败: {type(e).__name__}: {e}\n{trace}")
        return False

# 尝试读取最新版本信息
def _try_read_latest_info(ftp) -> tuple[dict, str]:
    root = REMOTE_VERSIONS_ROOT or "/"
    name = REMOTE_LATEST_FILE or "latest.yaml"

    # 可能的后缀
    import os as _os
    base, ext = _os.path.splitext(name.lstrip("/"))
    suffixes = [ext] if ext else [".yaml"]
    for SUFFIX in (".yaml", ".yml", ".json"):
        if SUFFIX not in suffixes:
            suffixes.append(SUFFIX)

    candidates = []
    for SUFFIX in suffixes:
        nm = (base if base else "latest") + SUFFIX
        # 原样（保持绝对/相对）
        raw = name if name.endswith(SUFFIX) else (("/" + nm) if name.startswith("/") else nm)
        candidates.append(_norm_remote(raw))
        # 与 root 拼接
        candidates.append(_join_remote(root, nm))

    # 去重保序
    seen = set()
    ordered = []
    for CANDIDATE in candidates:
        cc = _norm_remote(CANDIDATE)
        if cc not in seen:
            seen.add(cc)
            ordered.append(cc)

    for PATH in ordered:
        try:
            _v(f"[Updater] 尝试 latest: {PATH}")
            t = _ftp_read_text(ftp, PATH)
            if PATH.lower().endswith((".yaml", ".yml")):
                data = yaml.safe_load(t) or {}
            else:
                data = json.loads(t or "{}") or {}
            _v(f"[Updater] 使用 latest: {PATH}")
            return data, PATH
        except Exception as e:
            _v(f"[Updater] 读取失败 {PATH}: {e}")

    raise RuntimeError("latest 文件不可用（多路径/多后缀均失败）")

# 主流程
def check_update_then_run() -> None:
    global _PRINTED_LOGIN_SUCCESS

    # 检查本地升级开关
    settings = _MAIN_CFG.get("settings", {})
    enable_local_upgrade = settings.get("enable_local_upgrade", False)
    
    if not enable_local_upgrade:
        print("=== 本地版本升级已禁用，直接执行本地版本巡检 ===", flush=True)
        _run_local_core()
        return

    print("=== 正在连接版本服务器... ===", flush=True)
    local_ver = _read_local_version()
    latest_ver = None

    # 连接 FTP（带重试）
    try:
        ftp = _ftp_connect(passive=FTP_PASSIVE)
        if not _PRINTED_LOGIN_SUCCESS:
            print("=== 连接建立完成，登录成功! ===", flush=True)
            _PRINTED_LOGIN_SUCCESS = True
    except SystemExit:
        if not (CORE_ENTRY.exists() and CORE_ENTRY.is_file()):
            print("=== 与版本服务通信失败且本地缺失关键文件Core.py，请检查版本服务器FTP运行状态及网络状态 ===",
                  flush=True)
            _write_upgrade_log(tag="conn_fail_no_core", local_ver=local_ver, latest_ver=None)
            return
        print("=== 版本服务器连接失败，不进行版本更新，执行本地版本巡检 ===", flush=True)
        _run_local_core()
        _write_upgrade_log(tag="conn_fail", local_ver=local_ver, latest_ver=None)
        return
    except Exception:
        if not (CORE_ENTRY.exists() and CORE_ENTRY.is_file()):
            print("=== 与版本服务通信失败且本地缺失关键文件Core.py，请检查版本服务器FTP运行状态及网络状态 ===",
                  flush=True)
            _write_upgrade_log(tag="conn_fail_exc_no_core", local_ver=local_ver, latest_ver=None)
            return
        print("=== 版本服务器连接失败，不进行版本更新，执行本地版本巡检 ===", flush=True)
        _run_local_core()
        _write_upgrade_log(tag="conn_fail_exc", local_ver=local_ver, latest_ver=None)
        return

    try:
        # 读取 latest（容错）
        latest_info, used_path = _try_read_latest_info(ftp)
        latest_ver_raw = str(latest_info.get("latest") or latest_info.get("version") or "").strip()
        latest_ver = ("v" + latest_ver_raw.lstrip("vV")) if latest_ver_raw else "0"

        print(f"=== 正在读取最新版本文件，本地版本v{local_ver.lstrip('vV')}，服务器版本v{latest_ver.lstrip('vV')} ===",
              flush=True)

        cmp_val = _compare_versions(latest_ver, local_ver or "0")
        missing_required = _local_missing_required(latest_info)
        version_compatibility_required = _check_version_compatibility(local_ver or "0", latest_ver)
        need_refetch = (cmp_val == 0 and bool(missing_required)) or version_compatibility_required

        if cmp_val == 0 and not need_refetch:
            print("=== 本地与服务端版本一致，无需更新 ===", flush=True)
            _run_local_core()
            # 在"无需更新"分支也写一次日志，包含前面 _v 明细
            _write_upgrade_log(tag="no_update_end", local_ver=local_ver, latest_ver=latest_ver)
            return

        if cmp_val == 0 and need_refetch:
            if version_compatibility_required:
                print("=== 检测到版本升级，需要获取新组件，重新获取版本 ===", flush=True)
            else:
                print("=== 本地与服务端版本一致，但本地缺失关键文件，重新获取版本 ===", flush=True)

        with tempfile.TemporaryDirectory() as td:
            tmp_root = Path(td)
            payload_root = tmp_root / "payload"
            # 创建载荷根目录（如果目录已存在则不报错）
            payload_root.mkdir(parents=True, exist_ok=True)

            # 严格只下载 /vX 子树
            base_dir = _join_remote(REMOTE_VERSIONS_ROOT, latest_ver)
            print(f"=== 正在下载版本目录: {base_dir} ===", flush=True)
            downloaded: List[str] = []
            _ftp_download_dir_recursive(ftp, base_dir, payload_root, downloaded, root_prefix=base_dir)

            print("=== 下载完成，开始应用更新 ===", flush=True)

            # 创建升级前备份
            backup_dir = _create_backup_before_update()
            if backup_dir:
                print(f"=== 升级前备份已创建: {backup_dir.name} ===", flush=True)

            _apply_remote_version_payload(payload_root)
            _bump_local_version_yaml(LOCAL_VERSION_FILE, latest_ver)

            print(f"=== {base_dir} 版本更新成功！===", flush=True)
            print(f"=== 版本更新完成: {local_ver} -> {latest_ver} ===", flush=True)
            _write_upgrade_log(tag=f"update_{local_ver}_to_{latest_ver}", local_ver=local_ver, latest_ver=latest_ver)

    except Exception as e:
        print("=== 版本更新过程出现异常，不进行版本更新，执行本地版本巡检 ===", flush=True)
        _v(f"[Updater] 升级异常: {type(e).__name__}: {e}")
        _write_upgrade_log(tag="update_exception", local_ver=local_ver, latest_ver=latest_ver)
    finally:
        try:
            ftp.quit()
        except Exception:
            pass

    _run_local_core()
    _write_upgrade_log(tag="end", local_ver=local_ver, latest_ver=latest_ver)

if __name__ == "__main__":
    check_update_then_run()
