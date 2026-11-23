#!/bin/bash
# Alternativa ao .app: duplo clique abre o programa automaticamente.
# Não é preciso digitar nada — esta janela fecha sozinha ao sair do programa.

cd "/Users/casmtts/Documents/TRIBUTOS E TERRA/RPC - Lista 1 aba todos debitos" || exit 1

PYTHON=""
for candidato in /opt/homebrew/bin/python3 /usr/local/bin/python3 "$(command -v python3 2>/dev/null)" /usr/bin/python3; do
    [ -n "$candidato" ] || continue
    [ -x "$candidato" ] || continue
    if "$candidato" -c 'import PyQt6, playwright, openpyxl, cryptography' >/dev/null 2>&1; then
        PYTHON="$candidato"
        break
    fi
done
[ -n "$PYTHON" ] || PYTHON="/opt/homebrew/bin/python3"

if ! "$PYTHON" -c 'import PyQt6, playwright, openpyxl, cryptography' >/dev/null 2>&1; then
    echo "Instalando dependências..."
    "$PYTHON" -m pip install -r requirements.txt
    "$PYTHON" -m playwright install chromium
fi

exec "$PYTHON" main.py
