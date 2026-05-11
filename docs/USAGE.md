# HSE-Doc-Formatter 使用指南

## 安装

```bash
pip install -r requirements.txt
```

## 快速开始

```bash
# 直接使用
python format_expert.py 待处理文件/你的文档.docx

# 指定输出路径
python format_expert.py 你的文档.docx -o 排版完成.docx

# 不添加页码
python format_expert.py 你的文档.docx --no-pagenum

# 兼容旧用法（配置文件模式）
python format_expert.py config.json
```

## 排版标准

严格遵循 GB/T 9704-2012《党政机关公文格式》：

| 参数 | 值 |
|------|-----|
| 天头（上边距） | 37mm |
| 地脚（下边距） | 35mm |
| 左边距 | 28mm |
| 右边距 | 26mm |
| 行距 | 固定 30pt |
| 大标题 | 方正小标宋简体 22pt 加粗居中 |
| 一级标题 | 黑体 16pt 加粗 |
| 二级标题 | 楷体 16pt 加粗 |
| 三级标题 | 楷体 16pt 不加粗 |
| 正文 | 仿宋_GB2312 16pt 首行缩进 |
| 表头 | 黑体 10pt 加粗 |
| 表格内容 | 仿宋 10pt |

## 标题识别规则

### 自动识别并矫正编号

```
一、标题    → 一级标题（黑体16pt加粗）
（一）标题  → 二级标题（楷体16pt加粗）
1. 标题    → 三级标题（楷体16pt，保留原编号）
```

### 智能识别无编号标题

短小（≤25字）、无句末标点、无冒号结尾的段落，自动识别为二级标题并补全编号：

```
检查范围   → （一）检查范围    ← 自动识别
重点内容   → （二）重点内容    ← 自动补全
```

### 自动排除

```
组长：张三    → 正文（职位信息）
2026年1月    → 正文（日期开头）
这是一段完整的句子。 → 正文（句末标点）
```

### 内容保护

以下内容 100% 保持原样，不做任何修改：

- 图片（位置、大小、格式不变）
- 表格（结构、样式不变，仅格式化字体）
- 阿拉伯数字编号（1.、21、、(1)等绝不修改）

## 配置文件（可选）

```json
{
  "source": "待处理文件/input.docx",
  "output": "已处理文件/output.docx",
  "add_pagenum": true,
  "log_file": "run.log"
}
```

## 管道模式（Markdown 输入）

```bash
python run_pipeline.py input.md
```

## 字体说明

本工具默认使用 Windows 系统内置字体（方正小标宋、黑体、楷体、仿宋）。
在 Linux/macOS 上运行时，请确保已安装对应字体。

## 故障排查

| 错误 | 解决方案 |
|------|---------|
| ModuleNotFoundError | `pip install -r requirements.txt` |
| FileNotFoundError | 检查输入文件路径是否正确 |
| PermissionError | 检查输出目录是否可写 |
| 字体显示不正确 | 确保系统已安装所需中文字体 |
| 图片丢失 | 已修复（v1.0.0） |
| 表格样式异常 | 已修复（v1.0.0） |

## 版本历史

| 版本 | 日期 | 说明 |
|------|------|------|
| V5.0 | 2026-01-16 | 首次实现内容完整保护 |
| V5.1 | 2026-01-16 | 增加无编号标题自动识别 |
| v1.0.0 | 当前 | 合并 V5.0+V5.1，中文数字支持至 99，CLI 直传，开源发布 |

## 管道模式流程图

```
input.md → preprocess_md.py → Pandoc (黄金模板) → format_expert.py → output.docx
```
