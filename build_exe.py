# -*- coding: utf-8 -*-
"""
公文自动排版工具 v2.0 打包脚本
运行：python build_exe.py
"""
import subprocess
import sys
import os

APP_NAME = "公文自动排版工具"
ENTRY = "app.py"
ADD_DATA = [
    "templates/黄金模板.docx;templates",
    "preprocess_md.py;.",
    "format_expert.py;.",
    "run_pipeline.py;.",
    "icon.ico;.",
]

def build():
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--icon=icon.ico",
        f"--name={APP_NAME}",
    ]
    for d in ADD_DATA:
        cmd += ["--add-data", d]
    cmd.append(ENTRY)

    print("开始打包...")
    result = subprocess.run(cmd)
    if result.returncode == 0:
        exe_path = os.path.join("dist", f"{APP_NAME}.exe")
        size_mb = os.path.getsize(exe_path) / 1024 / 1024
        print(f"\n[OK] 打包成功：{exe_path}（{size_mb:.1f} MB）")
    else:
        print("\n[ERR] 打包失败，请检查错误信息")

if __name__ == "__main__":
    # 先安装 PyInstaller
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller", "--break-system-packages"])
    build()
