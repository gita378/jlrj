"""Build the self-contained Windows desktop executable from local wheels."""

from __future__ import annotations

import os
import subprocess
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent
TOOLS = ROOT / ".build_tools"


def main():
    if sys.platform != "win32":
        raise SystemExit("请在 Windows 上构建 EXE")
    wheels = list((ROOT / "wheelhouse").glob("*.whl"))
    if not any(p.name.startswith("pyinstaller-") for p in wheels):
        raise SystemExit("缺少 PyInstaller wheel：先运行 pip download pyinstaller -d wheelhouse")
    TOOLS.mkdir(exist_ok=True)
    for wheel in wheels:
        with zipfile.ZipFile(wheel) as package:
            package.extractall(TOOLS)
    env = os.environ.copy()
    env["PYTHONPATH"] = str(TOOLS) + os.pathsep + env.get("PYTHONPATH", "")
    temp = ROOT / ".build_temp"
    temp.mkdir(exist_ok=True)
    env["TEMP"] = env["TMP"] = str(temp)
    mode = "--onedir" if "--onedir" in sys.argv else "--onefile"
    args = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
            mode, "--windowed", "--name", "计量设备管理",
            "--distpath", str(ROOT / "dist"), "--workpath", str(ROOT / "build"),
            "--specpath", str(ROOT / "build"), str(ROOT / "desktop.py")]
    subprocess.run(args, cwd=ROOT, env=env, check=True)
    if mode == "--onedir":
        package = ROOT / "dist" / "计量设备管理-免安装版.zip"
        folder = ROOT / "dist" / "计量设备管理"
        with zipfile.ZipFile(package, "w", compression=zipfile.ZIP_DEFLATED,
                             compresslevel=9) as archive:
            for item in folder.rglob("*"):
                if item.is_file():
                    archive.write(item, item.relative_to(ROOT / "dist"))
        print(f"免安装包：{package}")


if __name__ == "__main__":
    main()
