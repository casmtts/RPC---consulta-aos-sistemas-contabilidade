# Consulta de Débitos — versão x86 (Intel)

Esta pasta é para uso em **computadores com arquitetura x86_64 (Intel)**.
Não é necessária em Macs Apple Silicon — nesses, use a pasta
`RPC - Consulta de débitos - REVISADO`.

## Por que uma pasta separada

Bibliotecas como **PyQt6** e os browsers do **Playwright** são específicos da
arquitetura. Instalar a versão errada gera erros de CPU ("bad CPU type").
Aqui tudo é instalado em um **virtualenv próprio** (`.venv-x86`) e os
browsers ficam em `~/Library/Caches/ms-playwright-x86`, separados de qualquer
instalação arm64.

## Como instalar no PC Intel

1. Dê duplo clique em **Instalar dependências x86.command**
   (ou rode `bash install_x86.sh` no terminal).
2. Ele cria o virtualenv e instala PyQt6, Playwright, openpyxl e
   cryptography (versão 42.0.8, fixada porque a 50 falha no build x86).
3. Também baixa o Chromium x86 do Playwright.

Se o macOS bloquear o `.command`: botão direito → **Abrir**.

## Como usar

Duplo clique em **Consulta de Débitos (x86).app**.

O programa abre direto, sem terminal. Se o macOS pedir confirmação:
botão direito → **Abrir**.

Logs: `~/Library/Logs/ConsultaDebitos-x86.log`.

## Conteúdo

- `main.py`, `config.py`, `brave_cookie.py`, `consulta_http.py` — programa
- `requirements-x86.txt` — dependências com versões corretas para Intel
- `install_x86.sh` / `Instalar dependências x86.command` — instalação
- `Consulta de Débitos (x86).app` — executável
- `.venv-x86/` — ambiente isolado (criado na instalação)
- documentação (README, guia e relatório)

## Observações

- Em macOS, o executável procura primeiro `/usr/local/bin/python3`, que é o
  caminho usual do Python em Macs Intel.
- A verificação `arch -x86_64` presente no script só entra em ação se a
  máquina for Apple Silicon; em Intel ela é ignorada.
- Cookies do Brave e conexão CDP (porta 9222) funcionam normalmente.
