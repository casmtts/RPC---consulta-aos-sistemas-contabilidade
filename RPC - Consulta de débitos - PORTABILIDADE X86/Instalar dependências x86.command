#!/bin/bash
# Duplo clique: instala as dependências x86 e fecha esta janela ao terminar.
cd "$(dirname "$0")" || exit 1
bash ./install_x86.sh
echo ""
echo "Pressione qualquer tecla para fechar..."
read -n 1 -s -r
