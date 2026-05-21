# HSE-Doc-Formatter

> 自动将中文公文 Word 文档按 GB/T 9704-2012 格式进行排版的命令行工具

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-v1.1.0-orange.svg)](CHANGELOG.md)

---

## 功能特性

- **四级标题体系**：完整支持 `一、→（一）→1.→(1)` 四级标题自动识别与格式化（v1.1.0 新增）
- 自动识别标题层级（一、（一）、1.、(1) 等）
- 智能检测无编号标题并自动补全
- 应用中文公文标准字体与字号（GB/T 9704-2012）
- 中文标题编号自动矫正（支持 0-99）
- 支持页码自动生成
- 保留原有表格与图片
- 支持 Markdown → Word 管道模式

---

## 安装

```bash
pip install -r requirements.txt
```

---

## 使用方法

### 直接排版 .docx 文件

```bash
# 基本用法
python format_expert.py 你的文档.docx

# 指定输出路径
python format_expert.py 你的文档.docx -o 排版完成.docx

# 不添加页码
python format_expert.py 你的文档.docx --no-pagenum
```

### Markdown 管道模式

```bash
# 输入 Markdown，自动转换并排版输出为 Word
python run_pipeline.py input.md
```

**管道流程：**

```
input.md
  │
  ▼  [Step 1] preprocess_md.py（转义数字列表，防 Pandoc 误解析）
  │
  ▼  [Step 2] Pandoc + 黄金模板（Markdown → raw .docx）
  │
  ▼  [Step 3] format_expert.py（GB/T 9704-2012 精排引擎）
  │
  ▼
output_标准排版.docx
```

### 配置文件模式（兼容旧用法）

```bash
python format_expert.py config.json
```

`config.example.json` 模板：

```json
{
  "source": "input.docx",
  "output": "output.docx",
  "add_pagenum": true,
  "log_file": "run.log"
}
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
| 三级标题 `1.` | 仿宋 16pt 加粗（紧凑模式） |
| **四级标题 `(1)`** | **仿宋 16pt（v1.1.0 新增）** |
| 正文 | 仿宋_GB2312 16pt 首行缩进 |
| 表头 | 黑体 10pt 加粗 |
| 表格内容 | 仿宋 10pt |

---

## 标题识别规则

### 自动识别并矫正编号

```
一、标题      → 一级标题（黑体 16pt 加粗）
（一）标题    → 二级标题（楷体 16pt）
1. 标题      → 三级标题（仿宋加粗，紧凑模式）
(1) 标题     → 四级标题（仿宋 16pt）      ← v1.1.0 新增
```

### 智能识别无编号标题

短小（≤25字）、无句末标点、无冒号结尾的段落，自动识别为二级标题：

```
检查范围    → （一）检查范围    ← 自动识别补全
重点内容    → （二）重点内容    ← 自动编号
```

### 内容 100% 保护

以下内容保持原样，不做任何修改：

- 图片（位置、大小、格式不变）
- 表格（结构不变，仅格式化字体）
- 阿拉伯数字编号（`1.`、`21、`、`(1)` 等绝不修改原始编号）

---

## 字体说明

本工具默认使用 Windows 系统内置字体（方正小标宋、黑体、楷体、仿宋）。

在 Linux/macOS 上运行时，请确保已安装对应字体，工具会自动切换为等效字体（`WenQuanYi`、`Noto Sans CJK SC`）。

---

## 系统要求

- Python 3.8+
- [Pandoc](https://pandoc.org/)（仅 Markdown 管道模式需要）
- Microsoft Word 字体（或等效中文字体）

---

## 故障排查

| 错误 | 解决方案 |
|------|---------|
| `ModuleNotFoundError` | `pip install -r requirements.txt` |
| `FileNotFoundError` | 检查输入文件路径是否正确 |
| `PermissionError` | 确认输出目录可写，且目标文件未被 Word 打开 |
| 字体显示不正确 | 确保系统已安装所需中文字体 |
| 图片丢失 | 已修复（v1.0.0） |
| 主标题误识别 | 已修复（v1.1.0） |

---

## 版本历史

| 版本 | 日期 | 说明 |
|------|------|------|
| V3.0 | 2026-01-15 | 基础版本（已废弃） |
| V4.0 | 2026-01-15 | 图片/表格丢失问题（已废弃） |
| V5.0 | 2026-01-16 | 首次实现内容完整保护 |
| V5.1 | 2026-01-16 | 增加无编号标题自动识别 |
| v1.0.0 | 2026-01-16 | 合并 V5.0+V5.1，中文数字支持至 99，开源发布 |
| **v1.1.0** | **2026-05-21** | **新增四级标题(1)支持；修复主标题误判逻辑；修复错误日志变量名；临时文件自动清理** |

---

## 许可证

[MIT License](LICENSE)
