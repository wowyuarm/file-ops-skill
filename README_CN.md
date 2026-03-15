# File Ops

English | [简体中文](README_CN.md)

本地文件操作技能，面向 AI agent。通过内置 Python 脚本对本机文件进行转换、检查、压缩和文本提取，所有操作返回结构化 JSON。

## 功能一览

### 转换 (convert)
- 常见图片格式互转：`png`、`jpg`、`webp`、`gif`、`bmp`、`tiff`
- `pdf -> docx`
- `xlsx -> csv`
- `xls -> csv`（需要 xlrd，默认不安装）
- `docx -> pdf`（需要 LibreOffice）
- `html/md/markdown -> pdf`（需要 wkhtmltopdf）

### 检查 (inspect)
- 基础元数据：文件大小、MIME 类型、权限、修改时间
- 图片：尺寸、色彩模式、格式
- PDF：页数
- Excel：工作表名、行数

### 压缩 (archive)
- 从文件和目录创建 zip 压缩包
- 解压 zip 到指定目录（含 ZipSlip 防护）

### 文本提取 (extract-text)
- 支持 PDF、DOCX、XLSX、HTML、Markdown、图片 EXIF 元数据

## 快速安装

```bash
bash skills/file-ops/scripts/setup.sh
```

## 使用方法

健康检查：
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py health
```

转换：
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py convert --input /路径/input.pdf --to docx
```

检查：
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py inspect --input /路径/file.pdf
```

压缩/解压：
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py archive --input /路径/目录 --output /路径/out.zip
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py archive --input /路径/archive.zip --output /路径/dest --extract
```

文本提取：
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py extract-text --input /路径/document.pdf
```

转换时如不想输出到输入文件旁边，可加 `--output /路径/output.ext`。

## 安装为 Skill

```bash
INSTALL_TO_CODEX=1 bash skills/file-ops/scripts/setup.sh
```
