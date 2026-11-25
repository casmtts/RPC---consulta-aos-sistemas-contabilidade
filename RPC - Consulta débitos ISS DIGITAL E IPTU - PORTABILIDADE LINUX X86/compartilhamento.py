#!/usr/bin/env python3
"""Cópia automática dos relatórios para a pasta compartilhada do VirtualBox.

A pasta ``departamento d.a`` existe no host e é compartilhada com a máquina
virtual. Dentro dela os relatórios são separados por departamento:

* ``Contabilidade`` — relatório da OBSERVAÇÃO (1 linha) e saída clássica;
* ``Tributação`` — verificações de pagos e de CIFs.

O programa localiza o compartilhamento automaticamente (montado pelo
VirtualBox em ``/media/sf_...``) e copia somente os arquivos novos ou
alterados. Se o compartilhamento não estiver montado, usa a pasta local
``departamento d.a`` da raiz do programa, para os relatórios não se perderem.
"""

from __future__ import annotations

import os
import re
import shutil
import time
from pathlib import Path

from config import (
    DEPARTAMENTOS, NOME_COMPARTILHAMENTO, PASTA_COMPARTILHADA,
)

RAIZ_PROGRAMA = Path(__file__).resolve().parent
PASTA_LOCAL = RAIZ_PROGRAMA / NOME_COMPARTILHAMENTO

ERRO_SEM_COMPARTILHAMENTO = (
    "compartilhamento 'departamento d.a' não encontrado ou sem permissão "
    "de escrita. No VirtualBox o nome costuma ser 'DEPARTAMENTO_D.A' e o "
    "ponto de montagem '/media/sf_DEPARTAMENTO_D.A'. Verifique se o usuário "
    "está no grupo vboxsf (execute 'newgrp vboxsf' ou reentre na sessão) e "
    "rode o script 'montar-compartilhamento.sh' da pasta do programa."
)

EXTENSOES_RELATORIO = {".xlsx", ".xlsm", ".csv"}
# Saídas da verificação (cobrança) ficam com Tributação; o resto é contábil.
SUFIXOS_TRIBUTACAO = (
    "_VERIFICACAO_PAGOS", "_DETALHE_PAGOS",
    "_VERIFICACAO_CIFS", "_DETALHE_CIFS",
)


def _normalizar(texto: object) -> str:
    """Remove acentos, pontos, espaços e caixa para comparar nomes."""
    return re.sub(r"[^a-z0-9]", "", str(texto).lower())


def localizar_pasta_compartilhada() -> tuple[Path, str] | None:
    """Devolve ``(pasta, origem)`` onde a pasta compartilhada foi encontrada.

    Ordem de procura: config.py → variável de ambiente → montagens do
    VirtualBox (/media, /mnt, /run/media) → pasta local de fallback.
    """
    alvo = _normalizar(NOME_COMPARTILHAMENTO)

    candidatos: list[tuple[Path, str]] = []
    if PASTA_COMPARTILHADA:
        candidatos.append((Path(str(PASTA_COMPARTILHADA)).expanduser(), "config.py"))
    for variavel in ("PASTA_COMPARTILHADA", "PASTA_DEPARTAMENTO"):
        valor = os.environ.get(variavel)
        if valor:
            candidatos.append((Path(valor).expanduser(), f"variável {variavel}"))

    raizes = [Path("/media"), Path("/mnt")]
    usuario = os.environ.get("USER") or os.environ.get("LOGNAME")
    if usuario:
        raizes.append(Path("/run/media") / usuario)
    for raiz in raizes:
        if not raiz.is_dir():
            continue
        try:
            entradas = list(raiz.iterdir())
        except OSError:
            continue
        for entrada in entradas:
            nome = entrada.name
            # O VirtualBox monta como /media/sf_<nome do compartilhamento>.
            if nome.startswith("sf_"):
                nome = nome[3:]
            if _normalizar(nome) == alvo and entrada.is_dir():
                candidatos.append((entrada, "compartilhamento VirtualBox"))

    for pasta, origem in candidatos:
        try:
            pasta.mkdir(parents=True, exist_ok=True)
            if pasta.is_dir() and os.access(pasta, os.W_OK):
                return pasta, origem
        except OSError:
            continue

    # Fallback: mantém a estrutura organizada mesmo fora da máquina virtual.
    try:
        PASTA_LOCAL.mkdir(parents=True, exist_ok=True)
        return PASTA_LOCAL, "pasta local (compartilhamento não montado ou sem permissão)"
    except OSError:
        return None


def _departamento_do_arquivo(nome: str) -> str:
    if any(sufixo in nome.upper() for sufixo in SUFIXOS_TRIBUTACAO):
        return DEPARTAMENTOS[1]
    return DEPARTAMENTOS[0]


def _mesmo_arquivo(origem: Path, destino: Path) -> bool:
    """Evita recopiar (e sobrescrever) um arquivo idêntico."""
    try:
        a, b = origem.stat(), destino.stat()
    except OSError:
        return False
    return (
        a.st_size == b.st_size
        and abs(a.st_mtime_ns - b.st_mtime_ns) <= 1_000_000
    )


def sincronizar_relatorios(
    pasta_relatorios: Path | str,
    pasta_destino: Path | str | None = None,
    idade_minima: float = 1.0,
) -> tuple[list[str], str | None]:
    """Copia os relatórios para a pasta compartilhada, por departamento.

    Devolve ``(copias, erro)``. ``copias`` contém rótulos como
    ``Contabilidade/relatorio.xlsx``. Nunca lança exceção: se o
    compartilhamento estiver indisponível, o erro é devolvido em ``erro``.
    """
    try:
        if pasta_destino:
            destino = Path(pasta_destino).expanduser()
            origem_destino = "destino informado"
            destino.mkdir(parents=True, exist_ok=True)
        else:
            encontrado = localizar_pasta_compartilhada()
            if encontrado is None:
                return [], ERRO_SEM_COMPARTILHAMENTO
            destino, origem_destino = encontrado

        # Avisa quando a cópia foi para a pasta local (o host não enxerga).
        aviso = (
            ERRO_SEM_COMPARTILHAMENTO
            if origem_destino.startswith("pasta local") else None
        )

        for departamento in DEPARTAMENTOS:
            (destino / departamento).mkdir(parents=True, exist_ok=True)

        origem = Path(pasta_relatorios)
        if not origem.is_dir():
            falha = f"pasta de relatórios não encontrada: {origem}"
            return [], f"{falha} | {aviso}" if aviso else falha

        copias: list[str] = []
        agora = time.time()
        for arquivo in sorted(origem.iterdir()):
            if not arquivo.is_file():
                continue
            if arquivo.name.startswith("~$"):
                continue  # arquivo de bloqueio temporário do Excel
            if arquivo.suffix.lower() not in EXTENSOES_RELATORIO:
                continue  # progresso .json, logs etc. não são relatórios
            info = arquivo.stat()
            # Não copia um arquivo que ainda pode estar sendo gravado.
            if agora - info.st_mtime < idade_minima:
                continue
            departamento = _departamento_do_arquivo(arquivo.name)
            alvo = destino / departamento / arquivo.name
            if _mesmo_arquivo(arquivo, alvo):
                continue
            shutil.copy2(arquivo, alvo)
            copias.append(f"{departamento}/{arquivo.name}")
        return copias, aviso
    except Exception as exc:  # noqa: BLE001
        falha = f"{type(exc).__name__}: {exc}"
        aviso = locals().get("aviso")
        return [], f"{falha} | {aviso}" if aviso else falha
