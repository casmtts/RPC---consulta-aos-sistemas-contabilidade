#!/usr/bin/env bash
# Remonta o compartilhamento do VirtualBox "departamento d.a" com permissão
# de leitura/escrita para o usuário que executa o programa.
# Pode ser executado manualmente ou pelo autostart instalado abaixo.

set -e

NOME_PADRAO="DEPARTAMENTO_D.A"
PONTO_PADRAO="/media/sf_DEPARTAMENTO_D.A"

# Descobre o nome e o ponto de montagem atuais (o VirtualBox usa maiúsculas).
linha_mount="$(mount | awk '/type vboxsf/ {print; exit}' || true)"
if [ -n "$linha_mount" ]; then
    NOME="$(printf '%s' "$linha_mount" | awk '{print $1}')"
    PONTO="$(printf '%s' "$linha_mount" | awk '{print $3}')"
else
    NOME="$NOME_PADRAO"
    PONTO="$PONTO_PADRAO"
fi

UID_USUARIO="$(id -u)"
GID_USUARIO="$(id -g)"
OPCAO="uid=${UID_USUARIO},gid=${GID_USUARIO},dmode=0775,fmode=0775"

if mountpoint -q "$PONTO" 2>/dev/null; then
    pkexec mount -t vboxsf -o "remount,${OPCAO}" "$NOME" "$PONTO"
else
    pkexec mkdir -p "$PONTO"
    pkexec mount -t vboxsf -o "$OPCAO" "$NOME" "$PONTO"
fi
