#!/bin/bash
# Abre a versão Linux com duplo clique (permissão de execução).
cd "$(dirname "$0")" || exit 1

export PLAYWRIGHT_BROWSERS_PATH="$HOME/.cache/ms-playwright-linux"

if [ ! -x .venv-linux/bin/python ]; then
    echo "Ambiente Linux não encontrado. Instalando..."
    bash install_linux.sh || exit 1
fi

exec .venv-linux/bin/python main.py
