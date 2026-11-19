#!/bin/bash

# Script de instalação do sistema de consulta de débitos

echo "=== Instalando dependências do sistema de consulta de débitos ==="

# Instalar dependências Python
echo "1. Instalando dependências Python..."
pip install -r requirements.txt

# Instalar Playwright
echo "2. Instalando Playwright..."
playwright install

# Instalar browsers
echo "3. Instalando browsers..."
playwright install chromium

echo ""
echo "=== Instalação concluída! ==="
echo "Para rodar o sistema: python main.py"
