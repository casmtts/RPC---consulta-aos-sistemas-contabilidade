#!/bin/bash
# Instala dependências e navegador em um virtualenv ISOLADO para x86_64.
# Isso evita misturar bibliotecas arm64 e x86 no mesmo site-packages.
set -e

DIRETORIO="$(cd "$(dirname "$0")" && pwd)"
export PLAYWRIGHT_BROWSERS_PATH="${HOME}/Library/Caches/ms-playwright-x86"
VENV="$DIRETORIO/.venv-x86"

CANDIDATOS="/usr/local/bin/python3 /usr/local/bin/python3.12 /usr/local/bin/python3.11 /opt/homebrew/bin/python3 /usr/bin/python3"
PYTHON=""
PREFIX=""

arquitetura_do_python() {
    $PREFIX "$1" -c 'import platform; print(platform.machine())' 2>/dev/null
}

# Sobrescrita opcional: PYTHON_X86=/caminho/do/python
if [ -n "${PYTHON_X86:-}" ] && [ -x "${PYTHON_X86}" ] &&
   [ "$(arquitetura_do_python "$PYTHON_X86")" = "x86_64" ]; then
    PYTHON="$PYTHON_X86"
fi

if [ -z "$PYTHON" ]; then
    for candidato in $CANDIDATOS; do
        [ -x "$candidato" ] || continue
        if [ "$(arquitetura_do_python "$candidato")" = "x86_64" ]; then
            PYTHON="$candidato"
            break
        fi
    done
fi

if [ -z "$PYTHON" ] && [ "$(uname -m)" = "arm64" ] && command -v arch >/dev/null 2>&1; then
    PREFIX="arch -x86_64"
    for candidato in $CANDIDATOS; do
        [ -x "$candidato" ] || continue
        if [ "$(arquitetura_do_python "$candidato")" = "x86_64" ]; then
            PYTHON="$candidato"
            break
        fi
    done
    [ -n "$PYTHON" ] || PREFIX=""
fi

if [ -z "$PYTHON" ]; then
    echo "ERRO: nenhum Python x86_64 encontrado."
    echo "Em Apple Silicon instale o Homebrew Intel/Rosetta:"
    echo '  arch -x86_64 /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
    echo "  arch -x86_64 /usr/local/bin/brew install python@3.12"
    exit 1
fi

echo "Base x86: $PREFIX $PYTHON"

# Cria o virtualenv se ainda não existir.
if [ ! -x "$VENV/bin/python" ]; then
    echo "Criando virtualenv x86 em $VENV ..."
    $PREFIX "$PYTHON" -m venv "$VENV"
fi

VENV_PY="$VENV/bin/python"

# O binário do venv é universal: em Apple Silicon ele nasce arm64 se chamado
# direto. Precisamos do prefixo arch -x86_64 também para o venv.
VPREFIX=""
if [ "$("$VENV_PY" -c 'import platform; print(platform.machine())' 2>/dev/null)" != "x86_64" ]; then
    if [ "$(uname -m)" = "arm64" ] && command -v arch >/dev/null 2>&1; then
        VPREFIX="arch -x86_64"
    fi
fi
if [ "$($VPREFIX "$VENV_PY" -c 'import platform; print(platform.machine())' 2>/dev/null)" != "x86_64" ]; then
    echo "ERRO: o virtualenv não ficou x86_64. Apague $VENV e tente novamente."
    exit 1
fi

$VPREFIX "$VENV_PY" -m pip install --upgrade pip
$VPREFIX "$VENV_PY" -m pip install -r "$DIRETORIO/requirements-x86.txt"
$VPREFIX "$VENV_PY" -m playwright install chromium
echo "Instalação x86 concluída em $VENV (prefixo: ${VPREFIX:-nenhum})"
