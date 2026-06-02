# HSE-Doc-Formatter v2.0 升级改造计划书

**项目路径：** `E:\My-HSE\My-HSE\Gemini\chinese-docx-auto-typesetting`
**参考开源：** https://github.com/KaguraNanaga/docformat-gui（v1.8.1）
**当前版本：** v1.1.0（纯命令行，仅排版引擎）
**目标版本：** v2.0.0（GUI + 标点引擎 + 批量处理 + exe 打包）
**执行方式：** 分三个阶段依次完成，每阶段完成后截图验收再进入下一阶段

---

## 现有文件清单（改造基础）

```
chinese-docx-auto-typesetting/
├── format_expert.py       ← 核心排版引擎（v1.1.0，约677行）
├── run_pipeline.py        ← MD→Word 三步管道
├── preprocess_md.py       ← Markdown 预处理
├── templates/
│   └── 黄金模板.docx      ← Pandoc 参考模板
├── requirements.txt       ← 依赖列表
└── README.md              ← 当前文档
```

**重要原则：以上所有文件只增量修改，不允许重写或删除现有逻辑。**

---

## 阶段一：format_expert.py 核心能力升级

> 目标：在现有排版引擎基础上，新增标点修正、格式清洗、处理模式三大能力。
> 改动文件：仅 `format_expert.py`。
> 禁止：不允许修改现有 `DocumentFormatter.format()` 的主流程和标题识别逻辑。

---

### 任务 1-A：新增标点符号修正模块

在 `format_expert.py` 的 `# ===== 主标题智能识别 =====` 区域之前，插入以下完整函数：

```python
# ===== 标点符号修正模块（v2.0.0 新增） =====

# 全半角标点映射表
PUNCT_MAP = {
    # 引号
    '"': '\u201c',   # 左双引号
    '"': '\u201d',   # 右双引号
    "'": '\u2018',   # 左单引号
    "'": '\u2019',   # 右单引号
    # 括号
    '(': '\uff08',   # 全角左括号
    ')': '\uff09',   # 全角右括号
    # 标点
    ',': '\uff0c',   # 全角逗号
    ';': '\uff1b',   # 全角分号
    ':': '\uff1a',   # 全角冒号
    '!': '\uff01',   # 全角感叹号
    '?': '\uff1f',   # 全角问号
}

# 不替换的安全上下文：英文字母/数字前后的标点保留半角
import unicodedata

def _is_cjk(char):
    """判断是否为中文字符"""
    try:
        return unicodedata.name(char).startswith('CJK')
    except ValueError:
        return False

def fix_punctuation_in_text(text):
    """
    对单段文本执行标点修正。
    规则：仅当标点两侧至少一侧是中文字符或文本端点时，才替换为全角。
    英文句子内部的标点（两侧均为英文/数字）保留半角，避免误伤代码和英文。
    """
    if not text:
        return text
    result = list(text)
    length = len(text)
    for i, ch in enumerate(text):
        if ch not in PUNCT_MAP:
            continue
        left = text[i-1] if i > 0 else None
        right = text[i+1] if i < length - 1 else None
        left_is_en = left and (left.isascii() and (left.isalnum() or left in '._-'))
        right_is_en = right and (right.isascii() and (right.isalnum() or right in '._-'))
        # 两侧都是英文/数字：不替换（保护英文句子和代码）
        if left_is_en and right_is_en:
            continue
        result[i] = PUNCT_MAP[ch]
    return ''.join(result)

def fix_punctuation(doc):
    """
    遍历文档所有段落和表格单元格，对每个 run 的文本执行标点修正。
    跳过：纯英文段落、代码块（等宽字体段落）。
    """
    fixed_count = 0
    for para in doc.paragraphs:
        # 跳过纯英文段落（段落文本中中文字符占比低于 10%）
        full_text = para.text
        if not full_text.strip():
            continue
        cjk_ratio = sum(1 for c in full_text if _is_cjk(c)) / len(full_text)
        if cjk_ratio < 0.1:
            continue
        for run in para.runs:
            original = run.text
            fixed = fix_punctuation_in_text(original)
            if fixed != original:
                run.text = fixed
                fixed_count += 1
    # 同步处理表格单元格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        original = run.text
                        fixed = fix_punctuation_in_text(original)
                        if fixed != original:
                            run.text = fixed
                            fixed_count += 1
    return fixed_count
```

---

### 任务 1-B：新增格式深度清洗模块

紧接任务 1-A 的代码之后，继续插入：

```python
# ===== 格式深度清洗模块（v2.0.0 新增） =====

from docx.shared import RGBColor

# 允许保留的字体颜色（黑色系），其余一律清除
_ALLOWED_COLORS = {None, RGBColor(0, 0, 0), RGBColor(0x1F, 0x1F, 0x1F)}

def clean_dirty_format(doc):
    """
    深度清洗文档格式残留：
    1. 清除段落首尾多余空格
    2. 清除非黑色字体颜色（颜色标注等）
    3. 清除正文段落的非必要下划线、斜体
    4. 不处理表格内容（保留表格原始样式）
    """
    cleaned_count = 0
    BODY_STYLE_NAMES = {'Normal', 'Body Text', '正文', 'BodyText', 'Body', 'FirstParagraph', '2'}

    for para in doc.paragraphs:
        # 1. 清除首尾空格
        for run in para.runs:
            if run.text != run.text.strip():
                # 只清除纯空格的 run，避免破坏首行缩进等
                if run.text.strip() == '':
                    run.text = ''
                cleaned_count += 1

        # 2 & 3. 仅对正文段落清洗颜色/下划线/斜体
        style_name = para.style.name if para.style else ''
        is_body = style_name in BODY_STYLE_NAMES or get_paragraph_type(para.text.strip()) == 'body'
        if not is_body:
            continue

        for run in para.runs:
            # 清除非黑色字体颜色
            if run.font.color.rgb not in _ALLOWED_COLORS:
                run.font.color.rgb = RGBColor(0, 0, 0)
                cleaned_count += 1
            # 清除下划线（公文正文不应有下划线）
            if run.font.underline:
                run.font.underline = False
                cleaned_count += 1
            # 清除斜体（公文正文不应有斜体）
            if run.font.italic:
                run.font.italic = False
                cleaned_count += 1

    return cleaned_count
```

---

### 任务 1-C：新增 `--mode` 参数和处理模式

**修改 `main()` 函数**，在现有 `parser.add_argument('--no-pagenum', ...)` 这一行之后，插入：

```python
    parser.add_argument(
        '--mode',
        choices=['full', 'diagnose', 'punct'],
        default='full',
        help=(
            'full（默认）：标点修正+清洗+完整排版；'
            'diagnose：仅检测问题，输出诊断报告，不修改文件；'
            'punct：仅做标点修正，保留原有排版格式'
        )
    )
```

**在 `main()` 函数末尾**，找到 `formatter = DocumentFormatter(...)` 这一行，在它之前插入以下 `diagnose` 和 `punct` 模式的处理逻辑：

```python
    # ===== 处理模式路由（v2.0.0 新增） =====
    mode = getattr(args, 'mode', 'full')

    if mode == 'diagnose':
        # 诊断模式：只读文档，输出问题报告，不写文件
        logger.info("="*60)
        logger.info("【诊断模式】仅检测问题，不修改文件")
        logger.info("="*60)
        doc = Document(source)
        issues = []
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text:
                continue
            # 检测标点问题
            fixed = fix_punctuation_in_text(text)
            if fixed != text:
                issues.append(f"  第{i+1}段 [标点问题] {text[:30]}...")
            # 检测首行缩进缺失
            pf = para.paragraph_format
            if pf.first_line_indent is None or pf.first_line_indent == 0:
                ptype = get_paragraph_type(text)
                if ptype == 'body':
                    issues.append(f"  第{i+1}段 [缺首行缩进] {text[:30]}...")
        logger.info(f"\n共发现 {len(issues)} 个问题：")
        for iss in issues:
            logger.info(iss)
        logger.info("\n【诊断完成】文件未被修改。")
        return

    if mode == 'punct':
        # 仅标点模式：只做标点修正，保留所有原有排版
        logger.info("="*60)
        logger.info("【标点修正模式】仅修正标点，保留原排版")
        logger.info("="*60)
        doc = Document(source)
        count = fix_punctuation(doc)
        doc.save(output)
        logger.info(f"[OK] 标点修正完成，共修正 {count} 处，输出：{output}")
        return

    # full 模式：先做标点+清洗，再走完整排版流程
    if mode == 'full':
        logger.info("[v2.0] 执行标点修正与格式清洗...")
        doc_pre = Document(source)
        punct_count = fix_punctuation(doc_pre)
        clean_count = clean_dirty_format(doc_pre)
        # 保存为临时文件，再传给 DocumentFormatter
        import tempfile
        tmp_fd, tmp_path = tempfile.mkstemp(suffix='.docx')
        os.close(tmp_fd)
        doc_pre.save(tmp_path)
        logger.info(f"[OK] 标点修正 {punct_count} 处，格式清洗 {clean_count} 处")
        source = tmp_path  # 替换 source，后续流程使用清洗后的临时文件
```

并在 `format()` 执行之后，追加临时文件清理：

```python
    # 清理 full 模式的临时文件
    if mode == 'full' and 'tmp_path' in locals():
        try:
            os.remove(tmp_path)
        except Exception:
            pass
```

**最后将版本号改为：**
```python
__version__ = "2.0.0"
```

---

### 任务 1-D：更新 `requirements.txt`

确认 `requirements.txt` 包含以下内容（如缺少则追加）：

```
python-docx>=0.8.11
lxml>=4.9.0
```

---

### 阶段一完成验收标准

完成后用以下三条命令测试，截图日志给我看：

```bash
# 验收1：full 模式（默认）
python format_expert.py 测试文档.docx -o 输出_full.docx

# 验收2：diagnose 模式
python format_expert.py 测试文档.docx --mode diagnose

# 验收3：punct 模式
python format_expert.py 测试文档.docx --mode punct -o 输出_punct.docx
```

---

## 阶段二：新增 GUI 界面（新建 app.py）

> 目标：新建桌面 GUI 程序，调用阶段一改造后的 format_expert.py。
> 改动文件：新建 `app.py`，其他文件不动。

---

### 任务 2-A：新建 app.py

在项目根目录新建 `app.py`，完整内容如下：

```python
# -*- coding: utf-8 -*-
"""
HSE-Doc-Formatter v2.0 GUI 启动程序
"""

import os
import sys
import json
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

APP_TITLE = "HSE 公文自动排版系统 v2.0"
CONFIG_FILE = Path(__file__).parent / "custom_settings.json"
SCRIPT_DIR = Path(__file__).parent
FORMAT_SCRIPT = SCRIPT_DIR / "format_expert.py"
PIPELINE_SCRIPT = SCRIPT_DIR / "run_pipeline.py"
PYTHON_EXE = sys.executable

# ===== 配置读写 =====
DEFAULT_CONFIG = {
    "last_input_dir": "",
    "last_output_dir": "",
    "mode": "full",
    "add_pagenum": True,
    "open_after_done": False,
}

def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)
            return cfg
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)

def save_config(cfg):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

# ===== 主窗口 =====
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x600")
        self.resizable(True, True)
        self.config_data = load_config()
        self._input_files = []
        self._build_ui()

    def _build_ui(self):
        # ── 顶栏：文件选择 ──
        top = tk.Frame(self, padx=16, pady=10)
        top.pack(fill='x')

        tk.Label(top, text="输入文件：", font=('', 10)).grid(row=0, column=0, sticky='w')
        self.input_var = tk.StringVar()
        tk.Entry(top, textvariable=self.input_var, width=58, state='readonly').grid(row=0, column=1, padx=6)
        tk.Button(top, text="选择文件", command=self._browse_input).grid(row=0, column=2, padx=4)
        tk.Button(top, text="选择文件夹", command=self._browse_folder).grid(row=0, column=3)

        tk.Label(top, text="输出目录：", font=('', 10)).grid(row=1, column=0, sticky='w', pady=(6, 0))
        self.output_var = tk.StringVar()
        tk.Entry(top, textvariable=self.output_var, width=58, state='readonly').grid(row=1, column=1, padx=6, pady=(6, 0))
        tk.Button(top, text="选择目录", command=self._browse_output).grid(row=1, column=2, padx=4, pady=(6, 0))

        # ── 模式选择 ──
        mode_frame = tk.LabelFrame(self, text="处理模式", padx=12, pady=8)
        mode_frame.pack(fill='x', padx=16, pady=(0, 6))
        self.mode_var = tk.StringVar(value=self.config_data.get('mode', 'full'))
        modes = [
            ("full",      "🪄 智能一键处理（推荐）：标点修正 + 格式清洗 + 公文排版全套"),
            ("diagnose",  "🩺 格式诊断：仅检测问题，输出报告，不修改文件"),
            ("punct",     "🩹 仅标点修复：只修正标点符号，保留原有排版格式"),
        ]
        for val, label in modes:
            tk.Radiobutton(mode_frame, text=label, variable=self.mode_var,
                           value=val, anchor='w').pack(fill='x', pady=1)

        # ── 选项 ──
        opt_frame = tk.Frame(self, padx=16)
        opt_frame.pack(fill='x', pady=(0, 6))
        self.pagenum_var = tk.BooleanVar(value=self.config_data.get('add_pagenum', True))
        tk.Checkbutton(opt_frame, text="自动添加页码", variable=self.pagenum_var).pack(side='left')
        self.open_var = tk.BooleanVar(value=self.config_data.get('open_after_done', False))
        tk.Checkbutton(opt_frame, text="处理完成后自动打开文件", variable=self.open_var).pack(side='left', padx=20)

        # ── 开始按钮 ──
        self.start_btn = tk.Button(
            self, text="▶  开始处理", font=('', 13, 'bold'),
            bg='#2196F3', fg='white', activebackground='#1565C0',
            padx=24, pady=8, command=self._start
        )
        self.start_btn.pack(pady=8)

        # ── 日志区域 ──
        log_frame = tk.LabelFrame(self, text="处理日志", padx=8, pady=6)
        log_frame.pack(fill='both', expand=True, padx=16, pady=(0, 12))
        self.log_text = tk.Text(log_frame, height=14, state='disabled',
                                font=('Consolas', 9), bg='#1e1e1e', fg='#d4d4d4',
                                wrap='word', relief='flat')
        scroll = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.log_text.pack(fill='both', expand=True)

        # ── 底部状态栏 ──
        self.status_var = tk.StringVar(value="就绪 — 请选择文件后点击开始处理")
        tk.Label(self, textvariable=self.status_var, anchor='w',
                 font=('', 9), fg='#555').pack(fill='x', padx=16, pady=(0, 6))

    # ── 文件选择 ──
    def _browse_input(self):
        init_dir = self.config_data.get('last_input_dir', '')
        files = filedialog.askopenfilenames(
            title="选择 Word 文档",
            initialdir=init_dir or None,
            filetypes=[("Word文档", "*.docx *.doc"), ("所有文件", "*.*")]
        )
        if files:
            self._input_files = list(files)
            self.config_data['last_input_dir'] = str(Path(files[0]).parent)
            self.input_var.set(f"{len(files)} 个文件：{Path(files[0]).name}" if len(files) > 1 else files[0])

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="选择包含 Word 文档的文件夹")
        if folder:
            found = list(Path(folder).rglob("*.docx")) + list(Path(folder).rglob("*.doc"))
            found = [str(p) for p in found if '_processed' not in p.stem and '_formatted' not in p.stem]
            self._input_files = found
            self.input_var.set(f"文件夹：{folder}（找到 {len(found)} 个文档）")

    def _browse_output(self):
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.output_var.set(folder)
            self.config_data['last_output_dir'] = folder

    # ── 日志写入 ──
    def _log(self, msg):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    # ── 开始处理 ──
    def _start(self):
        if not self._input_files:
            messagebox.showwarning("提示", "请先选择输入文件或文件夹")
            return
        self.start_btn.configure(state='disabled', text="处理中…")
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')
        self.status_var.set("处理中，请稍候…")
        # 保存当前配置
        self.config_data['mode'] = self.mode_var.get()
        self.config_data['add_pagenum'] = self.pagenum_var.get()
        self.config_data['open_after_done'] = self.open_var.get()
        save_config(self.config_data)
        # 后台线程执行
        t = threading.Thread(target=self._run_all, daemon=True)
        t.start()

    def _run_all(self):
        mode = self.mode_var.get()
        add_pagenum = self.pagenum_var.get()
        output_dir = self.output_var.get() or None
        open_after = self.open_var.get()
        success_count, fail_count = 0, 0

        for fpath in self._input_files:
            p = Path(fpath)
            out_name = p.stem + '_processed' + p.suffix
            out_path = str((Path(output_dir) if output_dir else p.parent) / out_name)
            cmd = [PYTHON_EXE, str(FORMAT_SCRIPT), fpath, '-o', out_path, '--mode', mode]
            if not add_pagenum:
                cmd.append('--no-pagenum')

            self.after(0, self._log, f"\n{'='*50}")
            self.after(0, self._log, f"处理：{p.name}")
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
                for line in (result.stdout + result.stderr).splitlines():
                    self.after(0, self._log, line)
                if result.returncode == 0:
                    success_count += 1
                    self.after(0, self._log, f"✅ 完成：{out_path}")
                    if open_after and mode != 'diagnose':
                        os.startfile(out_path)
                else:
                    fail_count += 1
                    self.after(0, self._log, f"❌ 失败：{p.name}")
            except Exception as e:
                fail_count += 1
                self.after(0, self._log, f"❌ 异常：{e}")

        summary = f"全部完成 — 成功 {success_count} 个，失败 {fail_count} 个"
        self.after(0, self._log, f"\n{'='*50}\n{summary}")
        self.after(0, self.status_var.set, summary)
        self.after(0, self.start_btn.configure, {'state': 'normal', 'text': '▶  开始处理'})

if __name__ == '__main__':
    app = App()
    app.mainloop()
```

---

### 任务 2-B：检测 Pandoc 并给出友好提示

在 `app.py` 的 `App.__init__` 方法最后，追加以下 Pandoc 检测逻辑：

```python
        # Pandoc 检测
        import shutil
        if not shutil.which('pandoc'):
            self.after(500, self._warn_pandoc)

    def _warn_pandoc(self):
        ans = messagebox.askokcancel(
            "提示：未检测到 Pandoc",
            "Markdown → Word 管道模式需要 Pandoc。\n\n"
            "直接排版 .docx 文件无需 Pandoc，可正常使用。\n\n"
            "如需 Markdown 转换功能，请安装 Pandoc：\n"
            "  在命令行运行：winget install jgm.pandoc\n\n"
            "点击「确定」复制安装命令到剪贴板。"
        )
        if ans:
            self.clipboard_clear()
            self.clipboard_append('winget install jgm.pandoc')
```

---

### 任务 2-C：创建启动脚本

在项目根目录新建 `启动排版系统GUI.bat`：

```bat
@echo off
chcp 65001 >nul
cd /d "%~dp0"
python app.py
pause
```

---

### 阶段二完成验收标准

截图给我看以下内容：
1. GUI 窗口正常启动
2. 选择一个 .docx 文件，点击开始处理，日志窗口有输出
3. 输出目录出现 `_processed.docx` 文件

---

## 阶段三：打包 exe + 更新文档

> 目标：将系统打包为单文件 exe，同步更新 README。
> 改动文件：新建 `build_exe.py`，更新 `README.md`。

---

### 任务 3-A：打包脚本

在项目根目录新建 `build_exe.py`：

```python
"""
HSE-Doc-Formatter v2.0 打包脚本
运行：python build_exe.py
"""
import subprocess
import sys
import os

APP_NAME = "HSE公文排版系统"
ENTRY = "app.py"
ADD_DATA = [
    "templates/黄金模板.docx;templates",
    "preprocess_md.py;.",
    "format_expert.py;.",
    "run_pipeline.py;.",
]

def build():
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
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
        print(f"\n✅ 打包成功：{exe_path}（{size_mb:.1f} MB）")
    else:
        print("\n❌ 打包失败，请检查错误信息")

if __name__ == "__main__":
    # 先安装 PyInstaller
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller", "--break-system-packages"])
    build()
```

执行打包：

```bash
python build_exe.py
```

---

### 任务 3-B：更新 README.md

将 `README.md` 完全替换为以下内容：

```markdown
# HSE-Doc-Formatter

> 中文公文 Word 文档自动排版工具，严格遵循 GB/T 9704-2012 标准

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-v2.0.0-orange.svg)](CHANGELOG.md)

---

## 快速开始（推荐）

**直接下载 exe，双击运行，无需安装 Python。**

前往 [Releases 页面](../../releases) 下载最新版本 `HSE公文排版系统.exe`，双击即用。

---

## 功能特性

- **三种处理模式**：智能一键处理 / 格式诊断 / 仅标点修复
- **标点符号自动修正**：全半角混用自动修正（v2.0.0 新增）
- **格式深度清洗**：清除杂乱颜色、下划线、多余空格（v2.0.0 新增）
- **四级标题体系**：完整支持 `一、→（一）→1.→(1)` 自动识别与格式化
- **批量文件处理**：多选文件或整个文件夹批量处理（v2.0.0 新增）
- **原文件安全**：输出文件独立命名 `_processed`，原文件永不覆盖
- 应用 GB/T 9704-2012 中文公文标准字体与字号
- 智能识别无编号标题并自动补全编号
- 保留原有表格与图片，不丢失内容
- 支持页码自动生成
- 支持 Markdown → Word 管道模式（需 Pandoc）

---

## 三种处理模式说明

| 模式 | 适用场景 | 是否修改文件 |
|------|---------|------------|
| 🪄 智能一键处理 | 日常通用，全套修正 | ✅ 是（输出新文件） |
| 🩺 格式诊断 | 检查文档有哪些问题 | ❌ 否 |
| 🩹 仅标点修复 | 只改标点，保留排版 | ✅ 是（输出新文件） |

---

## 命令行使用（进阶）

```bash
pip install -r requirements.txt

# 智能一键处理（默认）
python format_expert.py 你的文档.docx

# 格式诊断
python format_expert.py 你的文档.docx --mode diagnose

# 仅标点修复
python format_expert.py 你的文档.docx --mode punct

# 指定输出路径
python format_expert.py 你的文档.docx -o 排版完成.docx

# Markdown 管道模式（需安装 Pandoc）
python run_pipeline.py input.md
```

---

## 排版标准

严格遵循 GB/T 9704-2012《党政机关公文格式》：

| 参数 | 值 |
|------|-----|
| 页面天头 | 37mm |
| 页面地脚 | 35mm |
| 左边距 | 28mm |
| 右边距 | 26mm |
| 行距 | 固定 30pt |
| 大标题 | 方正小标宋简体 22pt 加粗居中 |
| 一级标题 `一、` | 黑体 16pt 加粗 |
| 二级标题 `（一）` | 楷体 16pt |
| 三级标题 `1.` | 仿宋 16pt 加粗 |
| 四级标题 `(1)` | 仿宋 16pt |
| 正文 | 仿宋_GB2312 16pt 首行缩进 |

---

## 系统要求

- Windows 10/11（exe 版本，无需额外安装）
- Python 3.8+（源码运行模式）
- [Pandoc](https://pandoc.org/)（仅 Markdown 管道模式需要）

---

## 版本历史

| 版本 | 日期 | 说明 |
|------|------|------|
| v1.0.0 | 2026-01-16 | 开源发布，四级标题，图片表格保护 |
| v1.1.0 | 2026-05-21 | 四级标题(1)支持；主标题误判修复 |
| **v2.0.0** | **2026-06-xx** | **GUI界面；标点修正引擎；格式清洗；批量处理；三种模式；exe打包** |

---

## 许可证

[MIT License](LICENSE)
```

---

### 阶段三完成验收标准

1. `dist/HSE公文排版系统.exe` 存在，大小在 30–80 MB 之间
2. 双击 exe，GUI 窗口正常弹出（无需 Python 环境）
3. 选一个 .docx 测试，处理成功，截图给我看

---

## 文件变更总览

| 文件 | 操作 | 说明 |
|------|------|------|
| `format_expert.py` | 修改 | 新增标点引擎、清洗模块、--mode 参数，版本升 2.0.0 |
| `app.py` | 新建 | GUI 主程序，调用 format_expert.py |
| `启动排版系统GUI.bat` | 新建 | Windows 双击启动脚本 |
| `build_exe.py` | 新建 | PyInstaller 打包脚本 |
| `custom_settings.json` | 自动生成 | 用户配置持久化（首次运行后生成） |
| `README.md` | 替换 | 全面更新为 v2.0.0 文档 |
| `requirements.txt` | 追加 | 确认 python-docx、lxml 版本 |
| `run_pipeline.py` | 不动 | 保留现有 MD 管道逻辑 |
| `preprocess_md.py` | 不动 | 保留现有预处理逻辑 |
| `templates/黄金模板.docx` | 不动 | 保留现有模板 |

---

## 执行顺序提示

**严格按阶段执行，每阶段截图验收后再进入下一阶段：**

```
阶段一（format_expert.py 升级）
  → 验收三条命令行测试截图
    ↓
阶段二（app.py GUI 界面）
  → 验收 GUI 窗口 + 处理成功截图
    ↓
阶段三（build_exe.py 打包 + README 更新）
  → 验收 exe 双击运行截图
```

**禁止事项（避免破坏现有逻辑）：**
- 不得修改 `run_pipeline.py` 和 `preprocess_md.py`
- 不得删除 `format_expert.py` 中任何现有函数
- 不得修改 `DocumentFormatter._format_paragraph()` 和 `format()` 的主流程，只在外部包裹新逻辑
- `templates/黄金模板.docx` 不得移动或重命名
