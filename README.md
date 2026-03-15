# File Ops

[简体中文](README_CN.md) | English

A local file operations skill for AI agents. Convert, inspect, archive, and extract text from files on the local machine using a bundled Python script with structured JSON output.

## What It Does

### Convert
- Image formats: `png`, `jpg`, `webp`, `gif`, `bmp`, `tiff`
- `pdf -> docx`
- `xlsx -> csv`
- `xls -> csv` (requires xlrd, not installed by default)
- `docx -> pdf` (requires LibreOffice)
- `html/md/markdown -> pdf` (requires wkhtmltopdf)

### Inspect
- Basic metadata: size, MIME type, permissions, last modified
- Image: dimensions, color mode, format
- PDF: page count
- Excel: sheet names, row counts

### Archive
- Create zip from files and directories
- Extract zip to directory (with ZipSlip protection)

### Extract Text
- PDF, DOCX, XLSX, HTML, Markdown, image EXIF metadata

## Quick Setup

```bash
bash skills/file-ops/scripts/setup.sh
```

## Usage

Health check:
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py health
```

Convert:
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py convert --input /path/input.pdf --to docx
```

Inspect:
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py inspect --input /path/file.pdf
```

Archive:
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py archive --input /path/dir --output /path/out.zip
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py archive --input /path/archive.zip --output /path/dest --extract
```

Extract text:
```bash
skills/file-ops/.venv/bin/python skills/file-ops/scripts/file_ops.py extract-text --input /path/document.pdf
```

Use `--output /path/output.ext` with convert if you do not want the output written next to the input file.

## Install As A Skill

```bash
INSTALL_TO_CODEX=1 bash skills/file-ops/scripts/setup.sh
```
