# -*- coding: utf-8 -*-
"""管線防並行：同專案同時只允許一個寫入流程。"""

from __future__ import annotations

import atexit
import os
import sys
from typing import Optional

LOCK_FILENAME = ".mlb_pipeline.lock"
_held_lock_path: Optional[str] = None


def lock_path(base_dir: str | None = None) -> str:
    root = base_dir or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(root, LOCK_FILENAME)


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


def _read_lock_pid(path: str) -> int | None:
    try:
        with open(path, "r", encoding="utf-8") as f:
            text = f.read().strip().splitlines()[0].strip()
        return int(text)
    except (OSError, ValueError, IndexError):
        return None


def acquire_pipeline_lock(base_dir: str | None = None, *, owner: str = "") -> None:
    """取得鎖；若其他存活行程占用則 SystemExit。過期鎖（PID 已死）會清掉重取。"""
    global _held_lock_path
    path = lock_path(base_dir)

    while True:
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            other = _read_lock_pid(path)
            if other == os.getpid():
                # 同行程已持鎖（可重入／重複呼叫）
                _held_lock_path = path
                return
            if other is not None and _pid_alive(other):
                label = f"（PID {other}）" if other else ""
                raise SystemExit(
                    "偵測到已有任務執行中"
                    f"{label}，請等目前視窗結束後再開。\n"
                    f"鎖檔：{path}"
                )
            try:
                os.remove(path)
            except OSError as e:
                raise SystemExit(f"無法清除過期鎖檔：{path}\n{e}") from e
            continue

        payload = f"{os.getpid()}\n{owner or sys.argv[0]}\n"
        try:
            os.write(fd, payload.encode("utf-8"))
        finally:
            os.close(fd)
        _held_lock_path = path
        atexit.register(release_pipeline_lock)
        return


def release_pipeline_lock() -> None:
    global _held_lock_path
    path = _held_lock_path
    if not path:
        return
    _held_lock_path = None
    try:
        if os.path.exists(path) and _read_lock_pid(path) in (None, os.getpid()):
            os.remove(path)
    except OSError:
        pass
