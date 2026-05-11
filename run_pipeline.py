import os
import subprocess
import sys
import json
import shutil

# ---- 项目根目录（脚本所在目录）----
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

if len(sys.argv) > 1:
    input_full_path = sys.argv[1]
    base_dir = os.path.dirname(input_full_path)
    input_filename = os.path.basename(input_full_path)
else:
    print("Usage: python run_pipeline.py <absolute_path_to_md_file>")
    sys.exit(1)

# ---- 前置检查 ----
if not shutil.which("pandoc"):
    print("[ERR] 未找到 pandoc，请先安装 Pandoc 并加入 PATH")
    sys.exit(1)

file_md = os.path.abspath(input_full_path)

# ---- 衍生路径 ----
base_name = os.path.splitext(input_filename)[0]
file_escaped = os.path.join(base_dir, "temp_escaped.md")
file_raw = os.path.join(base_dir, "temp_raw.docx")
file_final = os.path.join(base_dir, f"{base_name}_标准排版.docx")

file_template = os.path.join(PROJECT_DIR, "templates", "黄金模板.docx")
script_preprocess = os.path.join(PROJECT_DIR, "preprocess_md.py")
script_format = os.path.join(PROJECT_DIR, "format_expert.py")
config_file_dynamic = os.path.join(PROJECT_DIR, "config_dynamic.json")

python_exe = sys.executable


def run_step(desc, cmd_args):
    print(f"\n--- Running {desc} ---")
    try:
        subprocess.check_call(cmd_args)
        print(f"Success: {desc}")
    except subprocess.CalledProcessError as e:
        print(f"Error executing {desc}: {e}")
        sys.exit(1)


# ---- Step 0: 生成动态配置 ----
config_data = {
    "source": file_raw,
    "output": file_final,
    "log_file": os.path.join(PROJECT_DIR, "format_log.txt"),
    "add_pagenum": True
}
with open(config_file_dynamic, 'w', encoding='utf-8') as f:
    json.dump(config_data, f, indent=4, ensure_ascii=False)
print(f"Created dynamic config: {config_file_dynamic}")

# ---- Step 1: 预处理 Markdown ----
run_step("Preprocess Markdown",
         [python_exe, script_preprocess, file_md, file_escaped])

# ---- Step 2: Pandoc 转换 ----
run_step("Pandoc Conversion",
         ["pandoc", file_escaped, "-o", file_raw, f"--reference-doc={file_template}"])

# ---- Step 3: Python 格式化 ----
run_step("Python Formatting",
         [python_exe, script_format, config_file_dynamic])

print(f"\nAll steps completed successfully. Output: {file_final}")
