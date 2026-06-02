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

APP_TITLE = "公文自动排版工具 v2.0"
CONFIG_FILE = Path(__file__).parent / "custom_settings.json"
SCRIPT_DIR = Path(__file__).parent
FORMAT_SCRIPT = SCRIPT_DIR / "format_expert.py"
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

# ===== 日志重定向流（用于在进程内执行时捕获标准输出并实时更新 GUI） =====
class RedirectedStdout:
    def __init__(self, callback):
        self.callback = callback
        self.buffer = ""
        self.old_stdout = sys.stdout
        self.old_stderr = sys.stderr

    def write(self, string):
        self.buffer += string
        if '\n' in self.buffer:
            lines = self.buffer.split('\n')
            for line in lines[:-1]:
                self.callback(line)
            self.buffer = lines[-1]
        if self.old_stdout is not None:
            try:
                self.old_stdout.write(string)
            except Exception:
                pass

    def flush(self):
        if self.old_stdout is not None:
            try:
                self.old_stdout.flush()
            except Exception:
                pass

    def __enter__(self):
        sys.stdout = self
        sys.stderr = self
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.buffer:
            self.callback(self.buffer)
        sys.stdout = self.old_stdout
        sys.stderr = self.old_stderr

# ===== 主窗口 =====
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        
        # 图标加载适配（兼容 PyInstaller 运行环境）
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, 'icon.ico')
        try:
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
            else:
                self.iconbitmap('icon.ico')
        except Exception:
            pass

        self.geometry("780x620")
        self.resizable(True, True)
        self.configure(bg='white')
        self.config_data = load_config()
        self._input_files = []
        self._build_ui()
        
        # 居中显示窗口
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        extra_w = self.winfo_screenwidth() - w
        extra_h = self.winfo_screenheight() - h
        self.geometry(f"{w}x{h}+{extra_w//2}+{extra_h//2}")

        # Pandoc 延时检测
        import shutil
        if not shutil.which('pandoc'):
            self.after(500, self._warn_pandoc)

    def _build_ui(self):
        # ── 顶栏：文件选择 ──
        top = tk.Frame(self, bg='white', padx=16, pady=10)
        top.pack(fill='x')

        # 第一行：输入文件
        tk.Label(top, text="输入文件：", font=('Microsoft YaHei', 10), bg='white', fg='#333333').grid(row=0, column=0, sticky='w')
        self.input_var = tk.StringVar()
        self.input_entry = tk.Entry(top, textvariable=self.input_var, width=58, state='readonly', 
                                    readonlybackground='#F5F5F5', fg='#555555', relief='flat', bd=1)
        self.input_entry.grid(row=0, column=1, padx=6)
        
        btn_select_file = tk.Button(top, text="选择文件", bg='#F5F5F5', fg='#333333', activebackground='#E5E5E5',
                                    relief='groove', bd=1, cursor='hand2', padx=8, command=self._browse_input)
        btn_select_file.grid(row=0, column=2, padx=4)
        
        btn_select_folder = tk.Button(top, text="选择文件夹", bg='#F5F5F5', fg='#333333', activebackground='#E5E5E5',
                                      relief='groove', bd=1, cursor='hand2', padx=8, command=self._browse_folder)
        btn_select_folder.grid(row=0, column=3)

        # 第二行：输出目录
        tk.Label(top, text="输出目录：", font=('Microsoft YaHei', 10), bg='white', fg='#333333').grid(row=1, column=0, sticky='w', pady=(8, 0))
        self.output_var = tk.StringVar()
        self.output_entry = tk.Entry(top, textvariable=self.output_var, width=58, state='readonly',
                                     readonlybackground='#F5F5F5', fg='#555555', relief='flat', bd=1)
        self.output_entry.grid(row=1, column=1, padx=6, pady=(8, 0))
        
        btn_select_out = tk.Button(top, text="选择目录", bg='#F5F5F5', fg='#333333', activebackground='#E5E5E5',
                                   relief='groove', bd=1, cursor='hand2', padx=8, command=self._browse_output)
        btn_select_out.grid(row=1, column=2, padx=4, pady=(8, 0))

        # ── 模式选择（自定义卡片） ──
        mode_container = tk.LabelFrame(self, text=" 处理模式 ", bg='white', fg='#333333',
                                       padx=12, pady=8, font=('Microsoft YaHei', 10, 'bold'))
        mode_container.pack(fill='x', padx=16, pady=(4, 6))
        self._build_mode_cards(mode_container)

        # ── 选项 ──
        opt_frame = tk.Frame(self, bg='white', padx=16)
        opt_frame.pack(fill='x', pady=(0, 6))
        
        self.pagenum_var = tk.BooleanVar(value=self.config_data.get('add_pagenum', True))
        chk_page = tk.Checkbutton(opt_frame, text="自动添加页码", variable=self.pagenum_var, 
                                  bg='white', fg='#333333', activebackground='white', font=('Microsoft YaHei', 9))
        chk_page.pack(side='left')
        
        self.open_var = tk.BooleanVar(value=self.config_data.get('open_after_done', False))
        chk_open = tk.Checkbutton(opt_frame, text="处理完成后自动打开文件", variable=self.open_var,
                                  bg='white', fg='#333333', activebackground='white', font=('Microsoft YaHei', 9))
        chk_open.pack(side='left', padx=30)

        # ── 开始按钮 ──
        self.start_btn = tk.Button(
            self, text="▶  开始处理", font=('Microsoft YaHei', 13, 'bold'),
            bg='#C84B1A', fg='white', activebackground='#A33A10', activeforeground='white',
            relief='flat', cursor='hand2', bd=0, pady=10, command=self._start
        )
        self.start_btn.pack(fill='x', padx=16, pady=8)
        
        # Hover 效果
        self.start_btn.bind('<Enter>', lambda e: self.start_btn.configure(bg='#A33A10') if self.start_btn['state'] == 'normal' else None)
        self.start_btn.bind('<Leave>', lambda e: self.start_btn.configure(bg='#C84B1A') if self.start_btn['state'] == 'normal' else None)

        # ── 日志区域 ──
        log_frame = tk.LabelFrame(self, text=" 处理日志 ", bg='white', fg='#333333',
                                   padx=8, pady=6, font=('Microsoft YaHei', 10, 'bold'))
        log_frame.pack(fill='both', expand=True, padx=16, pady=(0, 8))
        self.log_text = tk.Text(log_frame, height=10, state='disabled',
                                font=('Consolas', 9), bg='#1E1E1E', fg='#D4D4D4',
                                wrap='word', relief='flat')
        scroll = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.log_text.pack(fill='both', expand=True)

        # ── 底部状态栏 ──
        self.status_var = tk.StringVar(value="就绪 — 请选择文件后点击开始处理")
        self.status_lbl = tk.Label(self, textvariable=self.status_var, anchor='w', bg='#FFFFFF',
                                   font=('Microsoft YaHei', 9), fg='#666666')
        self.status_lbl.pack(fill='x', padx=16, pady=(0, 6))

    def _build_mode_cards(self, parent):
        self.mode_var = tk.StringVar(value='full')
        self._mode_cards = {}

        # 推荐大卡（智能一键处理）
        self._make_card(
            parent, 'full',
            icon='✦', title='智能一键处理',
            badge='推荐', sub='标点修正 + 格式清洗 + 公文排版，一步完成',
            full_width=True
        )

        # 两个小卡并列
        row = tk.Frame(parent, bg='white')
        row.pack(fill='x', pady=(0, 4))
        self._make_card(
            row, 'diagnose',
            icon='⊙', title='格式诊断',
            sub='仅检测问题，不修改文件', side='left'
        )
        self._make_card(
            row, 'punct',
            icon='✂', title='标点修复',
            sub='只修正标点，保留排版', side='right'
        )

    def _make_card(self, parent, mode, icon, title, sub,
                   badge=None, full_width=False, side=None):
        card = tk.Frame(parent, bg='white', cursor='hand2',
                        relief='flat', highlightthickness=2)
        if full_width:
            card.pack(fill='x', pady=(0, 6))
        else:
            card.pack(side=side, fill='both', expand=True,
                      padx=(0, 4) if side == 'left' else (4, 0))
            
        self._mode_cards[mode] = card

        # 标题行容器
        title_frame = tk.Frame(card, bg='white')
        title_frame.pack(fill='x', anchor='w', padx=12, pady=(10, 2))

        lbl_title = tk.Label(title_frame, text=icon + ' ' + title,
                             font=('Microsoft YaHei', 11, 'bold'), bg='white', fg='#333333')
        lbl_title.pack(side='left')

        if badge:
            lbl_badge = tk.Label(title_frame, text=badge, font=('Microsoft YaHei', 8, 'bold'),
                                 fg='white', bg='#C84B1A', padx=4, pady=1)
            lbl_badge.pack(side='left', padx=6)

        # 副标题
        lbl_sub = tk.Label(card, text=sub, font=('Microsoft YaHei', 9),
                           fg='#888888', bg='white')
        lbl_sub.pack(anchor='w', padx=12, pady=(0, 8))

        # 绑定点击事件到 Frame 及其所有子控件
        def bind_click(widget):
            widget.bind('<Button-1>', lambda e, m=mode: self._select_mode(m))
            for child in widget.winfo_children():
                bind_click(child)

        bind_click(card)
        self._select_mode(self.config_data.get('mode', 'full'))

    def _select_mode(self, mode):
        self.mode_var.set(mode)
        ACTIVE = {
            'bg': '#FAECE7',
            'highlightbackground': '#C84B1A',
            'highlightcolor': '#C84B1A'
        }
        INACTIVE = {
            'bg': 'white',
            'highlightbackground': '#DDDDDD',
            'highlightcolor': '#DDDDDD'
        }

        for m, card in self._mode_cards.items():
            cfg = ACTIVE if m == mode else INACTIVE
            card.configure(
                bg=cfg['bg'],
                highlightbackground=cfg['highlightbackground'],
                highlightcolor=cfg['highlightcolor']
            )

            # 递归更新子控件背景，跳过“推荐”标记的红色背景
            def cascade_bg(widget, bg_color):
                try:
                    if widget.cget('text') == '推荐':
                        return
                except Exception:
                    pass
                widget.configure(bg=bg_color)
                for child in widget.winfo_children():
                    cascade_bg(child, bg_color)

            cascade_bg(card, cfg['bg'])

    # ── Pandoc 检测弹窗 ──
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

    # ── 文件选择逻辑 ──
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
            self.config_data['last_input_dir'] = folder

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
        self.start_btn.configure(state='disabled', text="⏳  处理中…")
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

        import format_expert

        for fpath in self._input_files:
            p = Path(fpath)
            out_name = p.stem + '_processed' + p.suffix
            out_path = str((Path(output_dir) if output_dir else p.parent) / out_name)

            self.after(0, self._log, f"\n{'='*50}")
            self.after(0, self._log, f"正在处理文件：{p.name}")

            # 准备排版引擎的命令行参数（在进程内覆盖 sys.argv）
            old_argv = sys.argv
            new_argv = [str(FORMAT_SCRIPT), fpath, '--mode', mode]
            if mode != 'diagnose':
                new_argv += ['-o', out_path]
            if not add_pagenum:
                new_argv.append('--no-pagenum')
            sys.argv = new_argv

            # 日志回调函数
            def log_line(line):
                self.after(0, self._log, line)

            # 在进程内运行排版引擎，并捕获标准输出以更新 GUI 日志窗口
            try:
                with RedirectedStdout(log_line):
                    format_expert.main()
                success_count += 1
                self.after(0, self._log, f"✅ 处理完成：{out_path if mode != 'diagnose' else '诊断报告输出完毕'}")
                if open_after and mode != 'diagnose' and os.path.exists(out_path):
                    os.startfile(out_path)
            except SystemExit as se:
                if se.code == 0:
                    success_count += 1
                    self.after(0, self._log, f"✅ 处理完成：{out_path if mode != 'diagnose' else '诊断报告输出完毕'}")
                    if open_after and mode != 'diagnose' and os.path.exists(out_path):
                        os.startfile(out_path)
                else:
                    fail_count += 1
                    self.after(0, self._log, f"❌ 处理失败：排版引擎返回错误代码 {se.code}")
            except Exception as e:
                fail_count += 1
                self.after(0, self._log, f"❌ 进程异常：{e}")
            finally:
                sys.argv = old_argv

        summary = f"全部完成 — 成功 {success_count} 个，失败 {fail_count} 个"
        self.after(0, self._log, f"\n{'='*50}\n{summary}")
        self.after(0, self.status_var.set, summary)
        self.after(0, self.start_btn.configure, {'state': 'normal', 'text': '▶  开始处理'})

if __name__ == '__main__':
    app = App()
    app.mainloop()
