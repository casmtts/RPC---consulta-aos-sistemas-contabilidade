#!/bin/bash
# Instala o ambiente da versão Linux Mint x86_64.
set -e
set -o pipefail
cd "$(dirname "$0")"

echo "=== Instalação Linux x86_64 ==="

# Pré-requisitos que precisam de root (janela de senha do sistema, se necessário).
if [ "$(id -u)" -eq 0 ]; then
    ELEVAR=""
elif command -v pkexec >/dev/null 2>&1; then
    ELEVAR="pkexec"
else
    ELEVAR="sudo"
fi

echo "--- Bibliotecas do sistema (Qt XCB + keyring) ---"
$ELEVAR bash -c 'export DEBIAN_FRONTEND=noninteractive; apt-get update -qq && apt-get install -y -qq libxcb-cursor0 libsecret-tools' \
    || echo "AVISO: não foi possível instalar as bibliotecas do sistema."

echo "--- Navegador Brave ---"
if command -v brave-browser >/dev/null 2>&1 || command -v brave >/dev/null 2>&1; then
    echo "Brave já instalado."
else
    TMP_BRAVE="$(mktemp)"
    cat > "$TMP_BRAVE" <<'EOF'
#!/bin/bash
set -e
export DEBIAN_FRONTEND=noninteractive
apt-get update -qq
apt-get install -y -qq apt-transport-https curl gnupg
curl -fsSLo /usr/share/keyrings/brave-browser-archive-keyring.gpg \
    https://brave-browser-apt-release.s3.brave.com/brave-browser-archive-keyring.gpg
echo "deb [signed-by=/usr/share/keyrings/brave-browser-archive-keyring.gpg arch=amd64] https://brave-browser-apt-release.s3.brave.com/ stable main" \
    > /etc/apt/sources.list.d/brave-browser-release.list
apt-get update -qq
apt-get install -y -qq brave-browser
EOF
    $ELEVAR bash "$TMP_BRAVE" \
        || echo "AVISO: não foi possível instalar o Brave (instale manualmente depois)."
    rm -f "$TMP_BRAVE"
fi

PYTHON_BASE="$(command -v python3 2>/dev/null || true)"
if [ -z "$PYTHON_BASE" ]; then
    echo "Python 3 não encontrado. Instale: sudo apt install python3 python3-venv python3-pip"
    exit 1
fi

python3 -m venv .venv-linux

# Recupera o pip caso o ambiente virtual tenha sido criado sem ele.
.venv-linux/bin/python -m ensurepip --upgrade || true

export PLAYWRIGHT_BROWSERS_PATH="$HOME/.cache/ms-playwright-linux"

.venv-linux/bin/python -m pip install --upgrade pip
.venv-linux/bin/python -m pip install -r requirements.txt
# No Linux Mint o Playwright sempre avisa que o sistema "não é oficialmente
# suportado" e usa a build do Ubuntu 24.04. Isso é esperado e inofensivo;
# abaixo o aviso é trocado por uma mensagem clara em português.
.venv-linux/bin/python -m playwright install chromium 2>&1 | sed \
    -e 's/^BEWARE: your OS is not officially supported by Playwright; downloading fallback build for .*/Aviso: Linux Mint usa a build compatível do Ubuntu 24.04 (normal; o Chromium funciona)./'
.venv-linux/bin/python -m playwright install-deps chromium 2>&1 | sed \
    -e 's/^BEWARE: your OS is not officially supported by Playwright; installing dependencies for .*/Aviso: dependências instaladas como no Ubuntu 24.04 (normal no Linux Mint)./' || true

echo "--- Guest Additions (arrastar e soltar / área de transferência) ---"
# Necessário para drag-and-drop e área de transferência dentro da VM.
bash "$(dirname "$0")/instalar_guest_additions.sh" \
    || echo "AVISO: não foi possível configurar os Guest Additions."

echo "=== Pronto ==="
echo "Ambiente: $(pwd)/.venv-linux"
echo "Browsers: $PLAYWRIGHT_BROWSERS_PATH"
