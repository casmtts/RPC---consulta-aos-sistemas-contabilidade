#!/bin/bash
# Abre a versão x86_64 com duplo clique. Instala o ambiente se faltar.
cd "$(dirname "$0")" || exit 1

export PLAYWRIGHT_BROWSERS_PATH="$HOME/Library/Caches/ms-playwright-x86"

if [ ! -x .venv-x86/bin/python ]; then
    echo "Ambiente x86 não encontrado. Instalando..."
    bash install_x86.sh || exit 1
fi

exec .venv-x86/bin/python main.py
