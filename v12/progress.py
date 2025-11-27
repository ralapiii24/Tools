"""
进度条辅助模块：提供带 position 控制的 tqdm 实例，并统一导出 write/set_lock。
"""

import sys
import threading
from typing import Any, Iterable, Optional

from tqdm import tqdm as _tqdm_orig

_POSITION = threading.local()


def get_position() -> Optional[int]:
    return getattr(_POSITION, "value", None)


def set_position(value: int) -> None:
    _POSITION.value = value


def clear_position() -> None:
    if hasattr(_POSITION, "value"):
        del _POSITION.value


def tqdm(*args: Iterable[Any], **kwargs: Any):
    pos = get_position()
    if pos is not None and "position" not in kwargs:
        kwargs["position"] = pos
    kwargs.setdefault("file", sys.__stdout__)
    return _tqdm_orig(*args, **kwargs)


def write(message: str) -> None:
    _tqdm_orig.write(message)


def set_lock(lock) -> None:
    _tqdm_orig.set_lock(lock)


def patch_tqdm_module() -> None:
    import tqdm as module

    module.tqdm = tqdm
    module.tqdm.write = write
    module.tqdm.set_lock = set_lock

