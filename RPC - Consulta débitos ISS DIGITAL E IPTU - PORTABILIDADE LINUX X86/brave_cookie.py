"""
Le UNICAMENTE os dois cookies de sessao do sistema em navegadores Chromium no
Linux (Brave, Chrome, Chromium, Edge e Vivaldi):
  - __Secure-DA : token JWT de autenticacao
  - JSESSIONID  : sessao JSF (perfil selecionado no servidor)

No Linux o Chromium guarda a chave no keyring (secret-tool) ou usa a senha
fixa "peanuts". Ambos sao tentados, em memoria, sem gravar nada em disco.

Se nao for possivel ler (perfil bloqueado, outro keyring, navegador fechado),
o programa continua: basta informar JWT e JSESSIONID manualmente na aba
Consulta.

Escopo de seguranca: conexao somente-leitura ao banco, filtro por NOMES e
HOST exatos; nenhum outro cookie, historico, senha ou configuracao e lido.
"""

from __future__ import annotations

import hashlib
import sqlite3
import subprocess
import time
from pathlib import Path

from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

# Navegadores Chromium no Linux: nome, caminho do banco e chave do keyring.
_NAVEGADORES = (
    ("Brave", ".config/BraveSoftware/Brave-Browser/Default/Cookies", "brave"),
    ("Google Chrome", ".config/google-chrome/Default/Cookies", "chrome"),
    ("Chromium", ".config/chromium/Default/Cookies", "chromium"),
    ("Microsoft Edge", ".config/microsoft-edge/Default/Cookies", "microsoft-edge"),
    ("Vivaldi", ".config/vivaldi/Default/Cookies", "vivaldi"),
)
ULTIMO_NAVEGADOR: str = ""

_HOST = "santanadoaraguaia-pa.desenvolvecidade.com.br"
_NOMES = ("__Secure-DA", "JSESSIONID")

DOMINIO_JWT = ".desenvolvecidade.com.br"
DOMINIO_SESSAO = "santanadoaraguaia-pa.desenvolvecidade.com.br"

_chave_cache: bytes | None = None
_chave_cache_time: float = 0.0
_chave_cache_app: str = ""
_CACHE_TTL = 300


def _senha_keyring(app: str) -> bytes:
    """Tenta a senha do keyring; usa 'peanuts' quando nao ha keyring."""
    for servico in (app, "Chromium Safe Storage", "Chrome Safe Storage"):
        try:
            senha = subprocess.check_output(
                ["secret-tool", "lookup", "application", servico],
                text=True, stderr=subprocess.DEVNULL, timeout=4,
            ).strip()
            if senha:
                return senha.encode("utf-8")
        except Exception:  # noqa: BLE001
            continue
    return b"peanuts"


def _chave_derivada(app: str) -> bytes:
    """Deriva a chave AES (Linux usa 1 iteracao do PBKDF2)."""
    global _chave_cache, _chave_cache_time, _chave_cache_app
    now = time.time()
    if (
        _chave_cache is not None
        and _chave_cache_app == app
        and (now - _chave_cache_time) < _CACHE_TTL
    ):
        return _chave_cache
    senha = _senha_keyring(app)
    _chave_cache = hashlib.pbkdf2_hmac(
        "sha1", senha, b"saltysalt", 1, 16
    )
    _chave_cache_time = now
    _chave_cache_app = app
    return _chave_cache


def _descriptografar(valor: bytes, chave: bytes) -> str:
    """Descriptografa cookie v10/v11 do Chromium."""
    if not valor or valor[:3] not in (b"v10", b"v11"):
        return ""
    bruto = valor[3:]
    if len(bruto) < 48:
        return ""
    iv, corpo = bruto[16:32], bruto[32:]
    corpo = corpo[: len(corpo) // 16 * 16]
    if not corpo:
        return ""
    dec = Cipher(algorithms.AES(chave), modes.CBC(iv)).decryptor()
    claro = dec.update(corpo) + dec.finalize()
    if not claro:
        return ""
    return claro[:-claro[-1]].decode("utf-8", errors="replace")


def _ler_cookies(caminho_db: Path, chave: bytes) -> dict[str, str]:
    """Le e descriptografa apenas os dois cookies do host."""
    con = sqlite3.connect(f"file:{caminho_db}?mode=ro", uri=True)
    try:
        con.execute("PRAGMA query_only=ON")
        registros = con.execute(
            "SELECT name, value, encrypted_value, host_key FROM cookies "
            "WHERE host_key IN (?, ?) AND name IN (?, ?) "
            "ORDER BY CASE WHEN host_key=? THEN 0 ELSE 1 END",
            (_HOST, "." + _HOST, *_NOMES, _HOST),
        ).fetchall()
    finally:
        con.close()

    encontrados: dict[str, tuple[str, bytes]] = {}
    for nome, valor_texto, valor_criptografado, _host in registros:
        if nome in encontrados:
            continue
        if valor_texto or valor_criptografado:
            encontrados[nome] = (valor_texto, valor_criptografado or b"")

    resultado: dict[str, str] = {}
    for nome in _NOMES:
        valor_texto, valor_criptografado = encontrados.get(nome, ("", b""))
        valor = valor_texto.strip() if valor_texto else ""
        if not valor:
            valor = _descriptografar(valor_criptografado, chave)
        if valor:
            resultado[nome] = valor
    return resultado


def buscar_cookies_no_brave() -> dict[str, str]:
    """Tenta cada navegador; devolve o JWT (e o JSESSIONID, se existir).

    O JSESSIONID é criado pelo próprio site ao abrir a home — ele NÃO precisa
    já estar no perfil do navegador. Exigir os dois cookies fazia o programa
    rejeitar uma sessão perfeitamente válida (JWT presente, JSESSIONID ainda
    não gravado) e bloquear toda a consulta.
    """
    global ULTIMO_NAVEGADOR
    tentados: list[str] = []
    for nome, relativo, app in _NAVEGADORES:
        caminho = Path.home() / relativo
        if not caminho.exists():
            continue
        tentados.append(nome)
        try:
            chave = _chave_derivada(app)
            cookies = _ler_cookies(caminho, chave)
            if _NOMES[0] in cookies:  # __Secure-DA: credencial de login
                ULTIMO_NAVEGADOR = nome
                return cookies
        except Exception:  # noqa: BLE001
            continue

    raise RuntimeError(
        "Não foi possível ler o token de autenticação (__Secure-DA) do "
        "navegador. Entre no sistema no Brave (deixe a sessão logada) e tente "
        "novamente; se ainda falhar, informe o JWT manualmente na aba Consulta. "
        f"Navegadores verificados: {', '.join(tentados) or 'nenhum'}."
    )


buscar_cookies_do_navegador = buscar_cookies_no_brave


def segundos_para_expirar(jwt: str) -> int | None:
    """Segundos restantes do JWT (sem verificar assinatura)."""
    import base64
    import json

    try:
        parte = jwt.split(".")[1]
        parte += "=" * (-len(parte) % 4)
        dados = json.loads(base64.urlsafe_b64decode(parte).decode("utf-8"))
        return int(dados["exp"]) - int(time.time())
    except Exception:  # noqa: BLE001
        return None
