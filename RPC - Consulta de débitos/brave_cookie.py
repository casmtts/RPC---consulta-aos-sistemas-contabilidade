"""
Le UNICAMENTE os dois cookies de sessao do sistema de debitos no Brave (macOS):
  - __Secure-DA : token JWT de autenticacao
  - JSESSIONID  : sessao JSF (guarda o perfil selecionado no servidor)

Ambos sao necessarios: sem o JSESSIONID o site nao reconhece o perfil e a
pesquisa e redirecionada para uma pagina com erro 403.

Escopo de seguranca (principio do menor privilegio):
  - Nenhuma copia do banco de cookies e gravada em disco (abertura somente-leitura).
  - A consulta SQL filtra por NOMES exatos e HOST exato — nenhum outro cookie
    e lido, nem mesmo de outros subdominios.
  - Nenhum historico, favorito, senha, configuracao ou dado de outro site e
    acessado.
  - A chave do Keychain e lida apenas do item "Brave Safe Storage" e fica
    somente na memoria; nunca e gravada nem exibida.
  - Os valores sao usados apenas para autenticar as consultas ao sistema de
    debitos; nada e enviado a terceiros.

Nao altera o perfil do navegador nem a sessao aberta.
"""

from __future__ import annotations

import base64
import hashlib
import json
import sqlite3
import subprocess
import time
from pathlib import Path

from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

# Unico arquivo e unicos registros que este modulo esta autorizado a acessar.
_COOKIES_DB = (
    Path.home()
    / "Library/Application Support/BraveSoftware/Brave-Browser/Default/Cookies"
)
_HOST = "santanadoaraguaia-pa.desenvolvecidade.com.br"
_NOMES = ("__Secure-DA", "JSESSIONID")
_SERVICO_KEYCHAIN = "Brave Safe Storage"

# Dominios/paths usados ao injetar os cookies no navegador do Playwright.
DOMINIO_JWT = ".desenvolvecidade.com.br"
DOMINIO_SESSAO = "santanadoaraguaia-pa.desenvolvecidade.com.br"


def _chave_derivada() -> bytes:
    """Le somente o item de Keychain indicado e deriva a chave em memoria."""
    senha = subprocess.check_output(
        ["security", "find-generic-password", "-ws", _SERVICO_KEYCHAIN],
        text=True,
    ).strip()
    return hashlib.pbkdf2_hmac("sha1", senha.encode(), b"saltysalt", 1003, 16)


def _descriptografar(valor: bytes, chave: bytes) -> str:
    """Descriptografa um valor de cookie v10/v11 do Chromium no macOS."""
    if not valor or valor[:3] not in (b"v10", b"v11"):
        return ""
    bruto = valor[3:]
    if len(bruto) < 48:
        return ""
    iv, corpo = bruto[16:32], bruto[32:]
    dec = Cipher(algorithms.AES(chave), modes.CBC(iv)).decryptor()
    claro = dec.update(corpo) + dec.finalize()
    return claro[:-claro[-1]].decode("utf-8", errors="replace")


def buscar_cookies_no_brave() -> dict[str, str]:
    """Retorna apenas os dois cookies de sessao do host do sistema.

    Raises:
        FileNotFoundError: perfil do Brave nao encontrado.
        RuntimeError: algum cookie inexistente, vazio ou invalido.
    """
    if not _COOKIES_DB.exists():
        raise FileNotFoundError(f"Banco de cookies do Brave nao encontrado: {_COOKIES_DB}")

    # Conexao somente-leitura; nenhum dado e copiado ou gravado.
    con = sqlite3.connect(f"file:{_COOKIES_DB}?mode=ro", uri=True)
    try:
        con.execute("PRAGMA query_only=ON")
        registros = con.execute(
            "SELECT name, encrypted_value FROM cookies "
            "WHERE host_key=? AND name IN (?, ?)",
            (_HOST, *_NOMES),
        ).fetchall()
    finally:
        con.close()

    encontrados = {nome: valor for nome, valor in registros}
    faltando = [nome for nome in _NOMES if not encontrados.get(nome)]
    if faltando:
        raise RuntimeError(
            f"Cookie(s) nao encontrado(s) no Brave: {', '.join(faltando)}. "
            "Verifique se a aba do sistema esta aberta e logada."
        )

    chave = _chave_derivada()
    resultado: dict[str, str] = {}
    for nome in _NOMES:
        valor = _descriptografar(encontrados[nome], chave)
        if not valor:
            raise RuntimeError(f"Nao foi possivel descriptografar o cookie {nome}.")
        resultado[nome] = valor

    # Validacao minima: o JWT deve comecar com 'eyJ' e ter 3 partes.
    jwt = resultado["__Secure-DA"]
    if not jwt.startswith("eyJ") or jwt.count(".") != 2:
        raise RuntimeError("Cookie __Secure-DA nao e um token de sessao valido.")
    return resultado


def segundos_para_expirar(jwt: str) -> int | None:
    """Segundos restantes da claim 'exp' do JWT (sem verificar assinatura).

    Usado apenas para decidir quando reler o cookie do Brave e avisar o
    usuario; a validacao definitiva e feita pelo servidor.
    """
    try:
        parte = jwt.split(".")[1]
        parte += "=" * (-len(parte) % 4)
        dados = json.loads(base64.urlsafe_b64decode(parte).decode("utf-8"))
        return int(dados["exp"]) - int(time.time())
    except Exception:  # noqa: BLE001
        return None
