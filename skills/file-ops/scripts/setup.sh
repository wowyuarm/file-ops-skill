#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
VENV_DIR="${SKILL_DIR}/.venv"
PYTHON_BIN="${PYTHON_BIN:-python3}"

PACKAGES=(
  pillow
  pandas
  pdf2docx
  pdfkit
  markdown
  openpyxl
  pymupdf
  python-docx
)

echo "Setting up File Ops Skill..."
echo "Skill dir: ${SKILL_DIR}"

if [[ -x "${VENV_DIR}/bin/python" ]] && "${VENV_DIR}/bin/python" -c "import importlib.util; raise SystemExit(0 if all(importlib.util.find_spec(name) for name in ['PIL', 'pandas', 'pdf2docx', 'pdfkit', 'markdown', 'openpyxl', 'fitz', 'docx']) else 1)"; then
  echo "Existing virtual environment already has the required Python packages"
else
  if command -v uv >/dev/null 2>&1; then
    echo "Using uv to create and populate ${VENV_DIR}"
    uv venv "${VENV_DIR}"
    uv pip install --python "${VENV_DIR}/bin/python" "${PACKAGES[@]}"
  else
    echo "uv not found, falling back to python -m venv"
    "${PYTHON_BIN}" -m venv "${VENV_DIR}"
    "${VENV_DIR}/bin/python" -m pip install --upgrade pip
    "${VENV_DIR}/bin/pip" install "${PACKAGES[@]}"
  fi
fi

if [[ "${INSTALL_TO_CODEX:-0}" == "1" ]]; then
  CODEX_HOME_DIR="${CODEX_HOME:-${HOME}/.codex}"
  TARGET_DIR="${CODEX_HOME_DIR}/skills/file-ops"
  mkdir -p "${CODEX_HOME_DIR}/skills"

  if [[ -e "${TARGET_DIR}" && ! -L "${TARGET_DIR}" ]]; then
    echo "Refusing to overwrite existing non-symlink skill at ${TARGET_DIR}"
    exit 1
  fi

  if [[ -L "${TARGET_DIR}" ]]; then
    CURRENT_TARGET="$(readlink "${TARGET_DIR}")"
    if [[ "${CURRENT_TARGET}" != "${SKILL_DIR}" ]]; then
      echo "Refusing to repoint existing skill symlink at ${TARGET_DIR}"
      exit 1
    fi
  else
    ln -s "${SKILL_DIR}" "${TARGET_DIR}"
    echo "Installed skill symlink at ${TARGET_DIR}"
  fi
fi

echo
echo "Health check:"
echo "  ${VENV_DIR}/bin/python ${SKILL_DIR}/scripts/file_ops.py health"
echo
echo "Example conversion:"
echo "  ${VENV_DIR}/bin/python ${SKILL_DIR}/scripts/file_ops.py convert --input /abs/path/input.pdf --to docx"
echo
echo "Example inspection:"
echo "  ${VENV_DIR}/bin/python ${SKILL_DIR}/scripts/file_ops.py inspect --input /abs/path/file.pdf"
