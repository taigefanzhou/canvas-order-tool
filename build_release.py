"""
Build a versioned release package.

Run on the Windows computer:
    python build_release.py

Output:
    release/丽群帆布纺织电商统计系统_v1.8.zip
"""

import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent
APP_FILE = ROOT / "帆布订单整理.py"
DIST_DIR = ROOT / "dist"
BUILD_DIR = ROOT / "build"
RELEASE_DIR = ROOT / "release"


def read_constant(name):
    text = APP_FILE.read_text(encoding="utf-8")
    match = re.search(rf'^{name}\s*=\s*"([^"]+)"', text, re.MULTILINE)
    if not match:
        raise RuntimeError(f"未找到 {name}")
    return match.group(1)


def main():
    app_name = read_constant("APP_NAME")
    app_version = read_constant("APP_VERSION")
    release_name = f"{app_name}_{app_version}"

    for folder in (DIST_DIR, BUILD_DIR):
        if folder.exists():
            shutil.rmtree(folder)
    RELEASE_DIR.mkdir(exist_ok=True)

    add_data_sep = ";" if os.name == "nt" else ":"
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        release_name,
        "--icon",
        str(ROOT / "logo.ico"),
        "--add-data",
        f"{ROOT / 'logo.png'}{add_data_sep}.",
        "--add-data",
        f"{ROOT / 'logo.ico'}{add_data_sep}.",
        str(APP_FILE),
    ]

    print("开始打包：", release_name)
    subprocess.check_call(cmd, cwd=ROOT)

    exe_suffix = ".exe" if os.name == "nt" else ""
    app_path = DIST_DIR / f"{release_name}{exe_suffix}"
    if not app_path.exists():
        raise RuntimeError(f"打包完成但没有找到文件：{app_path}")

    zip_path = RELEASE_DIR / f"{release_name}.zip"
    if zip_path.exists():
        zip_path.unlink()

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(app_path, arcname=app_path.name)

    print("打包完成：", zip_path)


if __name__ == "__main__":
    main()
