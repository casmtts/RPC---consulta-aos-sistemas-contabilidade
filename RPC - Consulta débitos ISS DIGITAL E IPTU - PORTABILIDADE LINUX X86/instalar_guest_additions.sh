#!/usr/bin/env bash
# Instala a parte de área de trabalho dos VirtualBox Guest Additions.
#
# Nesta VM os módulos do kernel (vboxguest/vboxsf) e o VBoxService já existem,
# mas falta o VBoxClient — responsável por arrastar e soltar, área de
# transferência e integração de tela. No Ubuntu/Mint ele vem em
# virtualbox-guest-x11 (virtualbox-guest-utils traz só os módulos/serviço).
set -e
cd "$(dirname "$0")"

USUARIO="$(id -un)"
if [ "$(id -u)" -eq 0 ]; then
    ELEVAR=""
elif command -v pkexec >/dev/null 2>&1; then
    ELEVAR="pkexec"
else
    ELEVAR="sudo"
fi

echo "=== VirtualBox Guest Additions — arrastar e soltar ==="

# 1) VBoxClient (arrastar/soltar, clipboard, tela)
if command -v apt-get >/dev/null 2>&1; then
    $ELEVAR bash -c 'export DEBIAN_FRONTEND=noninteractive; apt-get update -qq && apt-get install -y -qq virtualbox-guest-x11' \
        || echo "AVISO: não foi possível instalar virtualbox-guest-x11."
else
    echo "AVISO: apt-get não encontrado — instale o VBoxClient manualmente."
fi

# 2) Permissão de escrita na pasta compartilhada
if ! id -nG "$USUARIO" | grep -qw vboxsf; then
    echo "Adicionando $USUARIO ao grupo vboxsf..."
    $ELEVAR usermod -aG vboxsf "$USUARIO" \
        || echo "AVISO: não foi possível adicionar $USUARIO ao grupo vboxsf."
fi

# 3) Iniciar o VBoxClient em cada sessão gráfica
AUTOSTART="$HOME/.config/autostart"
mkdir -p "$AUTOSTART"
cat > "$AUTOSTART/vboxclient.sh" <<'EOF'
#!/usr/bin/env bash
# Um processo por recurso: o VBoxClient só aceita uma opção por chamada.
command -v VBoxClient >/dev/null 2>&1 || exit 0
for recurso in display seamless draganddrop clipboard; do
    VBoxClient "--$recurso" >/dev/null 2>&1 &
done
exit 0
EOF
chmod +x "$AUTOSTART/vboxclient.sh"
cat > "$AUTOSTART/vboxclient.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=VBoxClient (arrastar e soltar / área de transferência)
Comment=Integração de área de trabalho dos VirtualBox Guest Additions
Exec=$AUTOSTART/vboxclient.sh
Hidden=false
NoDisplay=true
Terminal=false
X-GNOME-Autostart-enabled=true
EOF

# 4) Aplicar já na sessão atual, se o binário existir
if command -v VBoxClient >/dev/null 2>&1; then
    "$AUTOSTART/vboxclient.sh"
    echo "VBoxClient iniciado nesta sessão."
else
    echo "VBoxClient ainda não disponível — reinicie a sessão após a instalação."
fi

echo
echo "No HOST do VirtualBox habilite: Configurações → Geral →"
echo "  Arrastar e Soltar: Bidirecional"
echo "  Área de Transferência: Bidirecional"
echo "Se $USUARIO foi adicionado a vboxsf agora, reentre na sessão para valer."
echo "=== Concluído ==="
