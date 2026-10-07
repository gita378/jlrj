"""Local smoke check for the native UI and the packaged EXE."""

from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parent


def check_source_ui():
    from desktop import DesktopApp, NAV

    app = DesktopApp()
    try:
        app.update()
        for key, _ in NAV:
            app.show(key)
            app.update()
        print("原生界面各页面打开正常")
    finally:
        app.destroy()


def check_exe():
    exe = (ROOT / "dist" / "计量设备管理" / "计量设备管理.exe"
           if "--onedir" in sys.argv else ROOT / "dist" / "计量设备管理.exe")
    if not exe.exists():
        raise FileNotFoundError(exe)
    data_root = ROOT / ".exe-smoke-data"
    data_root.mkdir(exist_ok=True)
    env = os.environ.copy()
    env["LOCALAPPDATA"] = str(data_root)
    env["TEMP"] = env["TMP"] = str(data_root)
    process = subprocess.Popen([str(exe)], cwd=ROOT, env=env)
    try:
        deadline = time.monotonic() + 60
        db = data_root / "CNAS计量台账" / "register.sqlite3"
        while time.monotonic() < deadline:
            if process.poll() is not None:
                raise RuntimeError(f"EXE 启动后退出，返回码 {process.returncode}")
            if db.exists() and db.stat().st_size:
                time.sleep(2)
                if process.poll() is None:
                    print("独立 EXE 启动正常，数据库已创建")
                    return
            time.sleep(0.5)
        raise RuntimeError("EXE 未在 60 秒内创建数据库")
    finally:
        process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()


if __name__ == "__main__":
    check_source_ui()
    if "--exe" in sys.argv:
        check_exe()
