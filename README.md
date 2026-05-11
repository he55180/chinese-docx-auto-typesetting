# HSE-Doc-Formatter

> 自动将中文公文 Word 文档按 GB/T 9704-2012 格式进行排版的命令行工具

## 功能特性
- 自动识别标题层级（一、（一）、1. 等）
- 智能检测无编号标题并自动补全
- 应用中文公文标准字体与字号
- 中文标题编号自动矫正（支持 0-99）
- 支持页码自动生成
- 保留原有表格与图片

## 安装

```bash
pip install -r requirements.txt
```

## 使用方法

```bash
# 复制配置模板并修改
cp config.example.json config.json
# 编辑 config.json 后执行：
python format_expert.py config.json

# Markdown管道模式
python run_pipeline.py input.md
```

### config.example.json 模板
```json
{
  "source": "待处理文件/input.docx",
  "output": "已处理文件/output.docx",
  "add_pagenum": true,
  "log_file": "run.log"
}
```

## 字体说明

本工具默认使用 Windows 系统内置字体（方正小标宋、黑体、楷体、仿宋）。
在 Linux/macOS 上运行时，请确保已安装对应字体，或修改源码中的字体配置。

## 系统要求
- Python 3.8+
- Microsoft Word 字体（或等效中文字体）

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
| 一级标题 | 黑体 16pt 加粗 |
| 二级标题 | 楷体 16pt 加粗 |
| 正文 | 仿宋_GB2312 16pt 首行缩进 |

## 版本历史

| 版本 | 日期 | 说明 |
|------|------|------|
| V3.0 | 2026-01-15 | 基础版本（已废弃） |
| V4.0 | 2026-01-15 | 图片/表格丢失问题（已废弃） |
| V5.0 | 2026-01-16 | 首次实现内容完整保护 |
| V5.1 | 2026-01-16 | 增加无编号标题自动识别 |
| v1.0.0 | 当前 | 合并 V5.0+V5.1，中文数字支持至 99 |

## 许可证

MIT License
