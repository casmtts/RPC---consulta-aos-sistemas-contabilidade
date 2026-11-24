#!/bin/bash
# Instala o ambiente x86_64 (Rosetta) em Apple Silicon.
# Em Mac Intel roda nativamente. Cria .venv-x86 e browsers separados.
set -e
cd "$(dirname "$0")"

echo "=== Instalação x86_64 ==="

PYTHON_BASE=""
for candidato in /opt/homebrew/bin/python3 /usr/local/bin/python3 "$(command -v python3 2>/dev/null)" /usr/bin/python3; do
    [ -n "$candidato" ] || continue
    [ -x "$candidato" ] || continue
    if arch -x86_64 "$candidato" -c 'import sys' >/dev/null 2>&1; then
        PYTHON_BASE="$candidato"
        break
    fi
done

if [ -z "$PYTHON_BASE" ]; then
    echo "Não foi possível executar Python em x86_64."
    echo "Em Apple Silicon instale o Rosetta 2:"
    echo "  softwareupdate --install-rosetta --agree-to-license"
    exit 1
fi

echo "Python base x86: $PYTHON_BASE"
arch -x86_64 "$PYTHON_BASE" -m venv .venv-x86

export PLAYWRIGHT_BROWSERS_PATH="$HOME/Library/Caches/ms-playwright-x86"
.venv-x86/bin/python -m pip install --upgrade pip
.venv-x86/bin/python -m pip install -r requirements-x86.txt
.venv-x86/bin/python -m playwright install chromium

echo "=== Pronto ==="
echo "Ambiente: $(pwd)/.venv-x86"
echo "Browsers: $PLAYWRIGHT_BROWSERS_PATH"
