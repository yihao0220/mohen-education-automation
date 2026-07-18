#!/usr/bin/env bash
set -euo pipefail

project_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
python_bin="${PYTHON_BIN:-python3}"

command -v "$python_bin" >/dev/null 2>&1 || {
  echo "错误：没有找到 $python_bin，请先安装 Python 3.11 或更高版本。" >&2
  exit 1
}
command -v qlmanage >/dev/null 2>&1 || {
  echo "错误：当前系统没有 macOS Quick Look 的 qlmanage。" >&2
  exit 1
}
command -v swiftc >/dev/null 2>&1 || {
  echo "错误：没有找到 swiftc，请先安装 Xcode Command Line Tools。" >&2
  exit 1
}

"$python_bin" -m venv "$project_root/.venv"
venv_python="$project_root/.venv/bin/python"
"$venv_python" -m pip install --upgrade pip
"$venv_python" -m pip install -r "$project_root/requirements-dev.txt"
"$venv_python" -c "import pdfplumber, pypdfium2, PIL, docx, docling, pytest"

echo "Mac 开发环境已就绪：$venv_python"
echo "说明：Quick Look 仅供开发预览，不能替代 Windows WPS 的生产页面真值。"
