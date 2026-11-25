#!/bin/bash
# Instala o ambiente da versão Linux Mint x86_64.
set -e
cd "$(dirname "$0")"

echo "=== Instalação Linux x86_64 ==="

PYTHON_BASE="$(command -v python3 2>/dev/null || true)"
if [ -z "$PYTHON_BASE" ]; then
    echo "Python 3 não encontrado. Instale: sudo apt install python3 python3-venv python3-pip"
    exit 1
fi

python3 -m venv .venv-linux
export PLAYWRIGHT_BROWSERS_PATH="$HOME/.cache/ms-playwright-linux"

.venv-linux/bin/python -m pip install --upgrade pip
.venv-linux/bin/python -m pip install -r requirements.txt
.venv-linux/bin/python -m playwright install chromium
.venv-linux/bin/python -m playwright install-deps chromium || true

echo "=== Pronto ==="
echo "Ambiente: $(pwd)/.venv-linux"
echo "Browsers: $PLAYWRIGHT_BROWSERS_PATH"
