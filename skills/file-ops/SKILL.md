---
name: file-ops
description: >-
  Local file operations on the user's machine: format conversion (image, PDF, Excel, DOCX, HTML, Markdown),
  file inspection (metadata, dimensions, page counts, sheet names), zip archive creation and extraction,
  and text extraction from PDF, DOCX, XLSX, HTML, Markdown, and image EXIF. Use when the user asks to
  convert files, check file details, compress/decompress files, or pull text content from documents.
  Runs a bundled Python script that returns structured JSON — no external API calls.
---

# File Ops

## Workflow

1. Run health check to discover available operations and missing dependencies:
   ```
   .venv/bin/python <SKILL_DIR>/scripts/file_ops.py health
   ```
2. Run the appropriate command with explicit paths:
   ```
   .venv/bin/python <SKILL_DIR>/scripts/file_ops.py <command> --input /path/to/file [options]
   ```
3. Read the JSON result and report the output path, metadata, or extracted text to the user.
4. If a dependency is missing, surface the exact gap from the health report or error message.

Replace `<SKILL_DIR>` with the absolute path to this skill directory. The script must be invoked
with the Python interpreter from the skill's own `.venv` (located at `<SKILL_DIR>/.venv`).
Run `<SKILL_DIR>/scripts/setup.sh` if the `.venv` does not exist.

## Commands

| Command | Purpose | Key flags |
|---------|---------|-----------|
| `health` | Report runtime capabilities | _(none)_ |
| `convert` | Convert file format | `--input`, `--to`, `--output`, `--sheet`, `--overwrite` |
| `inspect` | Return file metadata as JSON | `--input` |
| `archive` | Create or extract zip | `--input` (multi), `--output`, `--extract` |
| `extract-text` | Extract readable text | `--input` |

## Supported Conversions

- Image formats: `png`, `jpg`, `jpeg`, `webp`, `gif`, `bmp`, `tiff`
- `pdf -> docx`
- `xlsx -> csv`
- `xls -> csv` (requires xlrd, not installed by default)
- `docx -> pdf` (requires LibreOffice)
- `html/md/markdown -> pdf` (requires wkhtmltopdf)

## Notes

- All operations are path-based. Do not wrap file content in base64 unless the user explicitly asks.
- Run `health` before promising support for optional conversions or operations.
- For Excel conversion, `--sheet` accepts a sheet name or zero-based index.
- If a conversion or operation is unavailable, surface the exact dependency gap instead of falling back silently.
