#!/usr/bin/env python3
"""Verifica se os débitos da coluna OBSERVAÇÃO estão com situação PAGA.

Este módulo complementa ``main.py``: ele lê os débitos já listados na coluna
``OBSERVAÇÃO (débitos)`` do relatório (formato ``Tipo – Competência; ...``),
reconsulta cada CPF/CNPJ no site com os filtros de situação **DÉBITO + PAGO**
e grava:

* uma cópia do relatório com a coluna ``PAGO?`` (PAGO / EM ABERTO /
  NÃO CONFIRMADO), na mesma ordem dos débitos da observação;
* um relatório detalhado, uma linha por débito, com a situação encontrada
  no site e os valores/identificadores do débito.

A consulta em si continua em ``main.consultar_debito`` (função injetada no
worker), para não duplicar a lógica de sessão/cookies/JSF.
"""

from __future__ import annotations

import random
import re
import unicodedata
from pathlib import Path

from PyQt6.QtCore import QThread, pyqtSignal
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from disclaimer import adicionar_aviso
from regras_debito import STATUS_PENDENTE
from preencher_observacao import digitos, localizar_cabecalho, normalizar

SEPARADOR_ITENS = "; "
COLUNA_PAGO = "PAGO?"
COLUNA_COBRAVEL = "COBRÁVEL?"
COLUNA_SITUACAO = "SITUAÇÃO (SITE)"
COLUNA_NIVEL = "NÍVEL DE ALERTA"

NIVEL_COBRAR = "VERDE — COBRÁVEL"
NIVEL_NAO_COBRAR = "VERMELHO — NÃO COBRAR"
NIVEL_INDETERMINADO = "LARANJA — INDETERMINADO"
NIVEL_PENDENTE = "AMARELO — PENDENTE"

CORES_NIVEL = {
    NIVEL_COBRAR: ("FFC6EFCE", "FF006100"),
    NIVEL_NAO_COBRAR: ("FFFFC7CE", "FF9C0006"),
    NIVEL_INDETERMINADO: ("FFFFD9A0", "FF974706"),
    NIVEL_PENDENTE: ("FFFFEB9C", "FF9C6500"),
}

# Situações que encerram ou suspendem a exigibilidade — NÃO cobrar.
# A ordem importa: "AJUIZADO PAGO" e "DÍVIDA ATIVA PAGA" precisam cair em
# PAGO, não em AJUIZADO/DÍVIDA. A lista cobre TODAS as variações que o
# sistema mostra (inclusive PAGO MANUAL, PAGO POR COMPENSAÇÃO/DAÇÃO, etc.).
SITUACOES_ENCERRADAS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("PRESCRITO", ("prescrit",)),
    ("PAGO", ("pago", "paga", "quitad", "remid", "baixad")),
    # "RETIDO_DENTRO_MUNICIPIO_DEBITO" contém "debito", mas o tributo foi
    # retido na fonte: cobrar o contribuinte seria indevido. "retido" vem
    # ANTES das situações de dívida viva.
    ("RETIDO", ("retido",)),
    ("CANCELADO", ("cancelad", "anulad", "extint")),
    ("SUBSTITUÍDO", ("substituid",)),
    ("SUSPENSO", ("suspens",)),
    ("ISENTO", ("isent", "imun", "nao tributavel")),
)
# Somente estas mantêm a dívida viva — o gestor PODE cobrar.
# Tudo o mais (parcelamento revertido/descumprido, boleto gerado, em
# fiscalização, lançamento em AINF...) fica NÃO CONFIRMADO: na dúvida,
# nunca virar cobrança.
SITUACOES_COBRAVEIS = (
    "debito", "divida", "vencid", "a vencer", "consolidado",
    "parcelado", "ajuizado",
)

# Listas completas lidas dos dois sistemas (para auditoria/transparência).
# ISS Digital — valores exatos dos checkboxes "situaçãoDebito" do JSF.
SITUACOES_ISS_DIGITAL = (
    "DEBITO", "CONSOLIDADO", "LANCAMENTO_AINF", "PAGO", "DIVIDA_ATIVA",
    "DIVIDA_ATIVA_PARCELADA", "DIVIDA_ATIVA_PAGA", "PAGO_PARCELAMENTO",
    "PAGO_COMPENSACAO", "PAGO_RESTITUIDO", "PAGO_DACAO", "MEI",
    "SIMPLES_NACIONAL", "NOTA_AVULSA_GUIA_GERADA",
    "RETIDO_DENTRO_MUNICIPIO_DEBITO", "RETIDO_DENTRO_MUNICIPIO_PAGO",
    "RETIDO_FORA_MUNICIPIO", "NAO_TRIBUTAVEL", "ISENTO", "PARCELADO",
    "PARCELADO_OFFLINE", "PARCELAMENTO_REVERTIDO", "CANCELADO",
    "PRESCRITO", "REMIDO", "EM_FISCALIZACAO", "EXIGIBILIDADE_SUSPENSA",
    "SUSPENSO_ADMINISTRATIVO", "SUSPENSO_JUDICIAL", "AJUIZADO",
    "AJUIZADO_PAGO", "AJUIZADO_PARCELADO", "IMUNE",
)
# Conjunto usado na CONSULTA: situações de dívida e de pagamento/encerramento.
# As demais (MEI, SIMPLES_NACIONAL, NOTA_AVULSA, RETIDO...) não são dívidas a
# cobrar e só atrapalham o filtro do site. A classificação/auditoria continua
# cobrindo as 60 situações completas.
SITUACOES_VERIFICACAO_ESSENCIAIS = (
    "DEBITO", "CONSOLIDADO", "LANCAMENTO_AINF", "PAGO", "DIVIDA_ATIVA",
    "DIVIDA_ATIVA_PARCELADA", "DIVIDA_ATIVA_PAGA", "PAGO_PARCELAMENTO",
    "PAGO_COMPENSACAO", "PAGO_RESTITUIDO", "PAGO_DACAO", "NAO_TRIBUTAVEL",
    "ISENTO", "PARCELADO", "PARCELADO_OFFLINE", "PARCELAMENTO_REVERTIDO",
    "CANCELADO", "PRESCRITO", "REMIDO", "EM_FISCALIZACAO",
    "EXIGIBILIDADE_SUSPENSA", "SUSPENSO_ADMINISTRATIVO", "SUSPENSO_JUDICIAL",
    "AJUIZADO", "AJUIZADO_PAGO", "AJUIZADO_PARCELADO", "IMUNE",
)

# IPTU Digital — valores exatos do filtro "Situação Pagamento" do SPA.
SITUACOES_IPTU_DIGITAL = (
    "DÉBITO", "PAGO", "PAGO MANUAL", "NÃO TRIBUTÁVEL", "PARCELADO",
    "PARCELAMENTO REVERTIDO", "SUBSTITUÍDO", "ISENTO", "EM FISCALIZAÇÃO",
    "INSCRITO EM DÍVIDA ATIVA", "DÍVIDA ATIVA EM ACORDO DE PARCELAMENTO",
    "DÍVIDA ATIVA PAGA", "PARCELAMENTO DESCUMPRIDO", "PAGO PARCELAMENTO",
    "PRESCRITO", "EXIGIBILIDADE SUSPENSA", "SUSPENSO POR DECISÃO JUDICIAL",
    "SUSPENSO POR DECISÃO ADMINISTRATIVA", "ITBI (BOLETO GERADO)",
    "AJUIZADO", "AJUIZADO PAGO", "AJUIZADO PARCELADO",
    "PAGO POR COMPENSAÇÃO", "PAGO POR DAÇÃO", "REMIDO",
    "LANÇAMENTO EM AINF", "IMUNE",
)

COLUNAS_DETALHE = [
    "Aba", "Nome", "CPF/CNPJ", "Tipo de Débito", "Competência",
    "PAGO?", "COBRÁVEL?", "Situação no site", "Valor (R$)", "Pago (R$)",
    "CIF", "ID Dívida", "Vencimento", COLUNA_NIVEL, "Critério", "Observação",
]

CORES_STATUS = {
    "PAGO": ("FFDFF0DF", "FF1B5E20"),
    "EM ABERTO": ("FFFFE0B2", "FF8A4B00"),
    "CANCELADO": ("FFE1E1E1", "FF555555"),
    "PRESCRITO": ("FFE1E1E1", "FF555555"),
    "SUBSTITUÍDO": ("FFE1E1E1", "FF555555"),
    "RETIDO": ("FFE1E1E1", "FF555555"),
    "SUSPENSO": ("FFE1E1E1", "FF555555"),
    "ISENTO": ("FFE1E1E1", "FF555555"),
    "NÃO CONFIRMADO": ("FFFFF2CC", "FF7A6000"),
    "NÃO VERIFICADO": ("FFE1E1E1", "FF555555"),
}

CORES_COBRAVEL = {
    "SIM": ("FFDFF0DF", "FF1B5E20"),
    "NÃO": ("FFE1E1E1", "FF555555"),
    "NÃO CONFIRMADO": ("FFFFF2CC", "FF7A6000"),
    STATUS_PENDENTE: ("FFFFF2CC", "FF7A6000"),
}

MESES = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


# ---------------------------------------------------------------------------
# Normalização / comparação
# ---------------------------------------------------------------------------

def slug(texto: object) -> str:
    """Minúsculas sem acentos, mantendo números, barras e parênteses."""
    sem_acento = unicodedata.normalize("NFKD", str(texto or ""))
    sem_acento = "".join(c for c in sem_acento if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", sem_acento.lower()).strip()


# O site alterna nomes descritivos ("TAXA DE ALVARÁ DE OBRAS RESIDENCIAL")
# e códigos ("TAXA_ALVARA_OBRAS_RESIDENCIAL"). Palavras de ligação e a ordem
# dos termos não distinguem o tipo; remover/ordenar permite casar os dois.
PALAVRAS_LIGACAO = {
    "de", "do", "da", "dos", "das", "e", "para", "por", "a", "o", "as",
    "os", "no", "na", "nos", "nas", "ao", "aos", "em", "com",
}


def chave_tipo(texto: object) -> str:
    """Chave canônica do tipo: sem acentos, '(M)', pontuação ou artigos."""
    valor = slug(texto)
    valor = re.sub(r"\(m\)", " ", valor)
    valor = re.sub(r"[^a-z0-9]+", " ", valor)
    termos = [
        termo for termo in valor.split()
        if termo and termo not in PALAVRAS_LIGACAO
    ]
    return " ".join(sorted(termos))


def chave_competencia(texto: object) -> tuple[int, int]:
    """Devolve (ano, mês) de 'fev/22', '1/2026', '2022' etc."""
    valor = slug(texto)
    mes = None
    ano = None
    nome_mes = re.search(
        r"\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\b", valor
    )
    numerico = re.search(r"\b(\d{1,2})\s*[/\-]\s*(\d{2}|\d{4})\b", valor)
    if nome_mes:
        mes = MESES[nome_mes.group(1)]
        encontrado = re.search(r"\b(\d{2}|\d{4})\b", valor[nome_mes.end():])
        if encontrado:
            ano = int(encontrado.group(1))
    elif numerico:
        mes = int(numerico.group(1))
        ano = int(numerico.group(2))
    else:
        encontrado = re.search(r"\b(20\d{2})\b", valor)
        if encontrado:
            ano = int(encontrado.group(1))
    if ano is not None and ano < 100:
        ano += 2000
    return (ano or 0, mes or 0)


def chave_ordem_ano(item: dict) -> tuple[int, str]:
    """Ordena débitos do mais novo para o mais velho (ano decrescente)."""
    competencia = item.get("competencia") or item.get("Competência", "")
    ano = chave_competencia(competencia)[0]
    tipo = item.get("tipo") or item.get("Tipo de Débito", "")
    return (-ano, slug(tipo))


def competencia_equivalente(a: object, b: object) -> bool:
    """Compara competências; quando uma não tem mês, compara só o ano."""
    ano_a, mes_a = chave_competencia(a)
    ano_b, mes_b = chave_competencia(b)
    if not ano_a or not ano_b or ano_a != ano_b:
        return False
    if mes_a and mes_b:
        return mes_a == mes_b
    if not mes_a and not mes_b:
        return True
    # Um tem mês e o outro não: não há como saber se são a mesma parcela.
    return False


def parse_observacao(texto: object) -> list[dict]:
    """Transforma 'IPTU – fev/22; ITBI – 1/2025' em lista de débitos."""
    partes = re.split(r"\s*[;]\s*", str(texto or "").strip())
    debitos: list[dict] = []
    vistos: set[tuple[str, tuple[int, int]]] = set()
    for parte in partes:
        if not parte.strip():
            continue
        campos = re.split(r"\s+[–—-]\s+", parte.strip(), maxsplit=1)
        if len(campos) != 2:
            continue
        tipo, competencia = (campo.strip() for campo in campos)
        if not tipo or not competencia:
            continue
        chave = (chave_tipo(tipo), chave_competencia(competencia))
        if chave in vistos:
            continue
        vistos.add(chave)
        debitos.append({"tipo": tipo, "competencia": competencia})
    return sorted(debitos, key=chave_ordem_ano)


# ---------------------------------------------------------------------------
# Leitura dos documentos com OBSERVAÇÃO preenchida
# ---------------------------------------------------------------------------

def registros_com_observacao(caminho: Path, aba: str | None = None) -> list[dict]:
    """Devolve um registro por documento que tenha débitos na OBSERVAÇÃO.

    O mesmo documento pode aparecer em várias guias; os débitos são unidos e
    deduplicados. Se ``aba`` for informada e existir, somente ela é lida.
    """
    caminho = Path(caminho)
    wb = load_workbook(caminho, read_only=True, data_only=True)
    registros: dict[str, dict] = {}
    try:
        abas = [wb[aba]] if aba and aba in wb.sheetnames else wb.worksheets
        for ws in abas:
            if normalizar(ws.title) == "resumo":
                continue
            local = localizar_cabecalho(ws)
            if local is None:
                continue
            linha_cab, col_doc, col_obs = local
            cabecalho = next(
                ws.iter_rows(
                    min_row=linha_cab, max_row=linha_cab, values_only=True
                ),
                (),
            )
            col_nome = next(
                (i for i, valor in enumerate(cabecalho, start=1)
                 if "nome" in normalizar(valor)),
                None,
            )
            for linha in ws.iter_rows(min_row=linha_cab + 1, values_only=True):
                doc = digitos(linha[col_doc - 1] if col_doc <= len(linha) else "")
                if len(doc) not in (11, 14):
                    continue
                observacao = (
                    str(linha[col_obs - 1] or "").strip()
                    if col_obs <= len(linha) else ""
                )
                debitos = parse_observacao(observacao)
                if not debitos:
                    continue
                nome = (
                    str(linha[col_nome - 1] or "").strip()
                    if col_nome and col_nome <= len(linha) else ""
                )
                registro = registros.setdefault(doc, {
                    "Documento": doc,
                    "Nome": nome,
                    "Aba": ws.title,
                    "CPF": doc if len(doc) == 11 else "",
                    "CNPJ": doc if len(doc) == 14 else "",
                    "Tipo de Pessoa": "FISICA" if len(doc) == 11 else "JURIDICA",
                    "Debitos": [],
                })
                if not registro["Nome"] and nome:
                    registro["Nome"] = nome
                vistos = {
                    (chave_tipo(d["tipo"]), chave_competencia(d["competencia"]))
                    for d in registro["Debitos"]
                }
                for debito in debitos:
                    chave = (
                        chave_tipo(debito["tipo"]),
                        chave_competencia(debito["competencia"]),
                    )
                    if chave not in vistos:
                        vistos.add(chave)
                        registro["Debitos"].append(debito)
    finally:
        wb.close()
    return list(registros.values())


def registros_para_analise(resultados: list[dict]) -> list[dict]:
    """Monta a entrada do módulo novo a partir dos resultados já obtidos.

    Inclui os débitos vindos do snapshot legado — marcados com
    ``_pendente_confirmacao`` — para que o módulo novo os analise e classifique.
    Somente "Sem débitos" e marcadores "-"" ficam de fora.
    """
    registros: dict[str, dict] = {}
    for item in resultados or []:
        documento = digitos(item.get("CPF", ""))
        if len(documento) not in (11, 14):
            continue
        tipo = str(item.get("Tipo de Débito", "") or "").strip()
        competencia = str(item.get("Competência", "") or "").strip()
        if not tipo or not competencia or re.fullmatch(r"-{1,}", competencia):
            continue
        tipo_slug = slug(tipo)
        if tipo_slug in ("-",) or tipo_slug.startswith("sem debito"):
            continue
        registro = registros.setdefault(documento, {
            "Documento": documento,
            "Nome": str(item.get("Nome", "") or "").strip(),
            "Aba": str(item.get("Aba", "") or ""),
            "CPF": documento if len(documento) == 11 else "",
            "CNPJ": documento if len(documento) == 14 else "",
            "Tipo de Pessoa": "FISICA" if len(documento) == 11 else "JURIDICA",
            "Debitos": [],
        })
        chave = (chave_tipo(tipo), chave_competencia(competencia))
        if any(
            (chave_tipo(d["tipo"]), chave_competencia(d["competencia"])) == chave
            for d in registro["Debitos"]
        ):
            continue
        legado = bool(item.get("_pendente_confirmacao")) or (
            "relatório antigo" in str(item.get("Erro", "") or "").lower()
        )
        registro["Debitos"].append({
            "tipo": tipo, "competencia": competencia,
            "id_divida": str(item.get("ID Dívida", "") or "").strip(),
            "cif": str(item.get("CIF", "") or "").strip(),
            "vencimento": str(item.get("Vencimento", "") or "").strip(),
            "_pendente_confirmacao": legado,
        })
    for registro in registros.values():
        registro["Debitos"].sort(key=chave_ordem_ano)
    return list(registros.values())


# ---------------------------------------------------------------------------
# Classificação dos débitos encontrados no site
# ------------------------------------------------------------------------

def classificar_situacao(situacao: object, pago_valor: float = 0.0,
                         valor: float = 0.0) -> tuple[str, str]:
    """Devolve ``(PAGO?, COBRÁVEL?)`` para a situação informada pelo site.

    Regra de ouro contra cobrança indevida: só devolve ``SIM`` quando a
    situação é explicitamente de dívida viva. Qualquer pagamento (inclusive
    PAGO MANUAL, PAGO POR COMPENSAÇÃO/DAÇÃO, AJUIZADO PAGO, DÍVIDA ATIVA
    PAGA), prescrição, cancelamento, isenção, imunidade ou suspensão
    encerra a exigibilidade. Situação desconhecida nunca vira cobrança.
    """
    texto = slug(situacao)

    # Dívida integralmente quitada: não resta valor devido e há valor pago.
    if valor <= 0 and pago_valor > 0:
        return ("PAGO", "NÃO")

    for rotulo, termos in SITUACOES_ENCERRADAS:
        if any(termo in texto for termo in termos):
            return (rotulo, "NÃO")

    if any(termo in texto for termo in SITUACOES_COBRAVEIS):
        # Sem valor devido positivo não há o que cobrar, mesmo que a
        # situação diga DÉBITO (dados antigos/incompletos do site).
        if valor > 0:
            return ("EM ABERTO", "SIM")
        return ("NÃO CONFIRMADO", STATUS_PENDENTE)

    # Valor pago sem situação clara: também não pode virar cobrança.
    if pago_valor > 0:
        return ("PAGO", "NÃO")

    return ("NÃO CONFIRMADO", STATUS_PENDENTE)


def ano_de(texto: object) -> int:
    """Ano de uma data/vencimento (ex.: '31/12/2021' -> 2021)."""
    encontrado = re.search(r"(20\d{2})", str(texto or ""))
    return int(encontrado.group(1)) if encontrado else 0


def _cor_nivel_conjunto(valores: list[str]) -> tuple[str, str]:
    """Cor mais severa prevalece: vermelho > laranja > amarelo > verde."""
    for nivel in (
        NIVEL_NAO_COBRAR, NIVEL_INDETERMINADO, NIVEL_PENDENTE, NIVEL_COBRAR
    ):
        if nivel in valores:
            return CORES_NIVEL[nivel]
    return ("FFFFFFFF", "FF000000")


def nivel_alerta(cobravel: str, mensagem: str = "",
                 pendente: bool = False, inconsistente: bool = False) -> str:
    """Nível de alerta da cobrança.

    - VERDE: confirmado no site e em aberto → pode cobrar;
    - VERMELHO: pago/cancelado/prescrito/etc. → não cobrar (mesma cor para
      todos os encerramentos, como definido);
    - LARANJA: o sistema não conseguiu decidir (sem correspondência exata,
      conflito, ID não localizado) → exige decisão humana;
    - AMARELO: pendente de análise/falha na consulta → não cobrar ainda.
    """
    if cobravel == "NÃO":
        return NIVEL_NAO_COBRAR
    # Dados migrados com competência posterior ao vencimento (ex.: 2022 x
    # 2021) não identificam o débito com segurança: mesmo que o site diga
    # que está em aberto, vai para análise antes de qualquer cobrança.
    if inconsistente:
        return NIVEL_INDETERMINADO
    if cobravel == "SIM":
        return NIVEL_COBRAR
    if mensagem or pendente:
        return NIVEL_PENDENTE
    return NIVEL_INDETERMINADO


def classificar_debitos(esperados: list[dict], encontrados: list[dict],
                        mensagem: str = "") -> list[dict]:
    """Cruza cada débito da OBSERVAÇÃO com o que o site informa hoje.

    Resultado: ``PAGO?`` (situação do débito) e ``COBRÁVEL?`` (SIM / NÃO /
    NÃO CONFIRMADO). Sem correspondência segura, o débito é marcado como
    ``NÃO CONFIRMADO`` — nunca como pago ou cobrável por chute.
    """
    resultados: list[dict] = []
    usados: set[int] = set()

    # Índices pré-construídos: sem eles cada débito esperado percorria toda a
    # lista encontrada duas vezes (O(n²) em milhares de débitos).
    from collections import defaultdict

    indice_tipo_comp: dict[tuple[str, tuple[int, int]], list[tuple[int, dict]]] = (
        defaultdict(list)
    )
    indice_id: dict[str, list[tuple[int, dict]]] = defaultdict(list)
    for indice, candidato in enumerate(encontrados):
        chave = (
            chave_tipo(candidato.get("tipo_debito", "")),
            chave_competencia(candidato.get("competencia", "")),
        )
        indice_tipo_comp[chave].append((indice, candidato))
        indice_id[str(candidato.get("id_divida", "") or "").strip()].append(
            (indice, candidato)
        )

    for esperado in esperados:
        # Com dados migrados, tipo/competência não identificam a dívida com
        # segurança. Prioriza-se o ID Dívida; depois CIF + competência; e só
        # por fim tipo + competência (com conferência de mês).
        chave_esperada = (
            chave_tipo(esperado["tipo"]),
            chave_competencia(esperado["competencia"]),
        )
        vagos = [
            (i, c) for i, c in indice_tipo_comp.get(chave_esperada, [])
            if i not in usados
        ]
        id_divida = str(esperado.get("id_divida", "") or "").strip()
        cif = str(esperado.get("cif", "") or "").strip()
        candidatos: list[tuple[int, dict]] = []
        criterio = "Nenhum"
        if id_divida and id_divida not in ("-", "0"):
            exatos = [
                (i, c) for i, c in indice_id.get(id_divida, [])
                if i not in usados
            ]
            if exatos:
                candidatos, criterio = exatos[:1], "ID Dívida"
            else:
                # ID conhecido e ausente no site: não aceitar um "parecido".
                criterio = "ID não localizado"
        elif cif and cif not in ("-", "0"):
            por_cif = [
                (i, c) for i, c in vagos
                if str(c.get("cif", "") or "").strip() == cif
            ]
            candidatos = por_cif or vagos
            criterio = "CIF + Competência" if por_cif else "Tipo + Competência"
        else:
            candidatos = vagos
            criterio = "Tipo + Competência" if vagos else "Nenhum"

        encontrado = None
        observacao_extra = ""
        if not candidatos:
            situacao = ""
            pago_valor = valor = 0.0
        else:
            classificacoes = []
            for i, candidato in candidatos:
                sit = str(candidato.get("situacao", "") or "").strip()
                pago = float(candidato.get("pago", 0) or 0)
                val = float(candidato.get("valor", 0) or 0)
                classificacoes.append(
                    (i, candidato, sit, pago, val, classificar_situacao(sit, pago, val))
                )
            classificacoes_unicas = {item[5] for item in classificacoes}
            if len(classificacoes_unicas) == 1:
                i, encontrado, situacao, pago_valor, valor, classificacao = classificacoes[0]
                usados.add(i)
                situacao_resultado, cobravel = classificacao
            else:
                # Conflito entre débitos do mesmo tipo/competência.
                observacao_extra = (
                    "Mais de um débito com o mesmo tipo/competência e "
                    "situações diferentes — não foi possível identificar qual "
                    "pertence a esta guia."
                )
                situacao = "; ".join(sorted({item[2] for item in classificacoes if item[2]}))
                pago_valor = max(item[3] for item in classificacoes)
                valor = max(item[4] for item in classificacoes)
                situacao_resultado, cobravel = "NÃO CONFIRMADO", STATUS_PENDENTE

        if mensagem:
            situacao_resultado, cobravel = "NÃO VERIFICADO", STATUS_PENDENTE
        elif not candidatos:
            situacao_resultado, cobravel = "NÃO CONFIRMADO", STATUS_PENDENTE

        # Inconsistência típica da migração: competência depois do vencimento.
        ano_comp = chave_competencia(esperado["competencia"])[0]
        vencimento = (
            str(encontrado.get("vencimento", "") or "").strip()
            if encontrado else ""
        )
        ano_venc = ano_de(vencimento)
        inconsistente = bool(ano_comp and ano_venc and ano_venc < ano_comp)
        if inconsistente and cobravel == "SIM":
            situacao_resultado, cobravel = "NÃO CONFIRMADO", STATUS_PENDENTE
        if inconsistente:
            aviso_data = (
                f"Competência {ano_comp} posterior ao vencimento "
                f"{ano_venc} — conferir débito migrado."
            )
            observacao_extra = (
                (observacao_extra + " " + aviso_data) if observacao_extra
                else aviso_data
            )

        resultados.append({
            "Tipo de Débito": esperado["tipo"],
            "Competência": esperado["competencia"],
            "PAGO?": situacao_resultado,
            "COBRÁVEL?": cobravel,
            "Situação no site": situacao,
            "Valor (R$)": valor,
            "Pago (R$)": pago_valor,
            "CIF": str(encontrado.get("cif", "") or "") if encontrado else "",
            "ID Dívida": str(encontrado.get("id_divida", "") or "") if encontrado else "",
            "Vencimento": str(encontrado.get("vencimento", "") or "") if encontrado else "",
            COLUNA_NIVEL: nivel_alerta(
                cobravel, mensagem,
                bool(esperado.get("_pendente_confirmacao")),
                inconsistente,
            ),
            "Critério": (
                "NÃO VERIFICADO" if mensagem else
                "ID não localizado" if criterio == "ID não localizado" else criterio
            ),
            "Observação": mensagem or observacao_extra or (
                "Débito confirmado no site."
                if esperado.get("_pendente_confirmacao") and encontrado is not None
                else "Débito pendente de confirmação no site."
                if esperado.get("_pendente_confirmacao")
                else "Débito não localizado na consulta ao site."
                if encontrado is None else ""
            ),
        })
    return sorted(resultados, key=chave_ordem_ano)


# ---------------------------------------------------------------------------
# Gravação das saídas
# ---------------------------------------------------------------------------

def _estilo_cabecalho(ws, numero_linha: int, quantidade_colunas: int) -> None:
    for coluna in range(1, quantidade_colunas + 1):
        celula = ws.cell(row=numero_linha, column=coluna)
        celula.fill = PatternFill("solid", start_color="FF1F4E79", end_color="FF1F4E79")
        celula.font = Font(color="FFFFFFFF", bold=True, size=11)
        celula.alignment = Alignment(horizontal="center", vertical="center")


def _cor_conjunto(valores: list[str], tabela: dict[str, tuple[str, str]]) -> tuple[str, str]:
    """Escolhe a cor de um conjunto de status (o mais incerto prevalece)."""
    preferencia = ("NÃO CONFIRMADO", STATUS_PENDENTE, "NÃO VERIFICADO", "CANCELADO",
                   "PRESCRITO", "PAGO", "EM ABERTO", "NÃO", "SIM")
    for candidato in preferencia:
        if candidato in valores:
            return tabela.get(candidato, ("FFFFFFFF", "FF000000"))
    return ("FFFFFFFF", "FF000000")


def _cabecalho_novo(ws, linha: int, coluna: int, titulo: str) -> None:
    celula = ws.cell(row=linha, column=coluna, value=titulo)
    celula.fill = PatternFill("solid", start_color="FF1F4E79", end_color="FF1F4E79")
    celula.font = Font(color="FFFFFFFF", bold=True, size=11)
    celula.alignment = Alignment(horizontal="center", vertical="center")


def gravar_resultado_pagos(origem: Path, destino: Path, detalhe: Path,
                           resultados: list[dict]) -> tuple[Path, Path]:
    """Grava PAGO?/COBRÁVEL? no relatório e o detalhe em planilha separada."""
    origem, destino, detalhe = Path(origem), Path(destino), Path(detalhe)
    base = destino if destino.exists() else origem
    wb = load_workbook(base)

    por_documento: dict[str, list[dict]] = {}
    for item in sorted(resultados, key=chave_ordem_ano):
        doc = digitos(item.get("Documento", ""))
        if doc:
            por_documento.setdefault(doc, []).append(item)

    for ws in wb.worksheets:
        if normalizar(ws.title) == "resumo":
            continue
        local = localizar_cabecalho(ws)
        if local is None:
            continue
        linha_cab, col_doc, col_obs = local
        nomes = [
            normalizar(celula.value)
            for celula in ws[linha_cab][: max(ws.max_column, col_obs)]
        ]
        def achar(termo: str) -> int | None:
            return next(
                (i for i, nome in enumerate(nomes, start=1) if termo in nome),
                None,
            )

        col_pago = achar("pago")
        col_cobravel = achar("cobravel")
        # A planilha-base já tem "Situação Guia"/"Situação CNPJ": só casa
        # o cabeçalho exato da coluna nova, para não sobrescrever dados.
        col_sit = achar("situacao site") or achar("situacao no site")
        col_nivel = achar("nivel de alerta")
        faltando = [
            titulo for titulo, coluna in (
                (COLUNA_PAGO, col_pago), (COLUNA_COBRAVEL, col_cobravel),
                (COLUNA_SITUACAO, col_sit), (COLUNA_NIVEL, col_nivel),
            ) if coluna is None
        ]
        if faltando:
            existentes = [
                c for c in (
                    col_pago, col_cobravel, col_sit, col_nivel, col_obs
                ) if c
            ]
            inicio = max(existentes) + 1
            ws.insert_cols(inicio, amount=len(faltando))
            for deslocamento, titulo in enumerate(faltando):
                coluna = inicio + deslocamento
                _cabecalho_novo(ws, linha_cab, coluna, titulo)
                if titulo == COLUNA_PAGO:
                    col_pago = coluna
                elif titulo == COLUNA_COBRAVEL:
                    col_cobravel = coluna
                elif titulo == COLUNA_SITUACAO:
                    col_sit = coluna
                else:
                    col_nivel = coluna

        ws.column_dimensions[get_column_letter(col_pago)].width = 34
        ws.column_dimensions[get_column_letter(col_cobravel)].width = 22
        ws.column_dimensions[get_column_letter(col_sit)].width = 30
        ws.column_dimensions[get_column_letter(col_nivel)].width = 26

        for numero in range(linha_cab + 1, ws.max_row + 1):
            doc = digitos(ws.cell(row=numero, column=col_doc).value)
            itens = por_documento.get(doc)
            pagos = [str(item.get("PAGO?", "")) for item in itens or []]
            cobraveis = [str(item.get("COBRÁVEL?", "")) for item in itens or []]
            situacoes = [
                str(item.get("Situação no site", "") or "") for item in itens or []
            ]
            niveis = [
                str(item.get(COLUNA_NIVEL, "") or "") for item in itens or []
            ]
            celula_sit = ws.cell(row=numero, column=col_sit)
            celula_sit.value = (
                SEPARADOR_ITENS.join(s for s in situacoes if s) if situacoes else ""
            )
            celula_sit.alignment = Alignment(vertical="center", wrap_text=True)
            celula_nivel = ws.cell(row=numero, column=col_nivel)
            celula_nivel.value = SEPARADOR_ITENS.join(niveis) if niveis else ""
            celula_nivel.alignment = Alignment(vertical="center", wrap_text=True)
            fundo, texto = _cor_nivel_conjunto(niveis)
            celula_nivel.fill = PatternFill("solid", start_color=fundo, end_color=fundo)
            celula_nivel.font = Font(color=texto, bold=True, size=10)
            for coluna, valores, tabela in (
                (col_pago, pagos, CORES_STATUS),
                (col_cobravel, cobraveis, CORES_COBRAVEL),
            ):
                celula = ws.cell(row=numero, column=coluna)
                celula.value = SEPARADOR_ITENS.join(valores) if valores else ""
                celula.alignment = Alignment(vertical="center", wrap_text=True)
                fundo, texto = _cor_conjunto(valores, tabela)
                celula.fill = PatternFill("solid", start_color=fundo, end_color=fundo)
                celula.font = Font(color=texto, size=10)

    destino.parent.mkdir(parents=True, exist_ok=True)
    adicionar_aviso(wb)
    wb.save(destino)

    detalhe_wb = Workbook()
    detalhe_ws = detalhe_wb.active
    detalhe_ws.title = "Detalhe Pagos"
    detalhe_ws.sheet_view.showGridLines = False
    detalhe_ws.append(COLUNAS_DETALHE)
    _estilo_cabecalho(detalhe_ws, 1, len(COLUNAS_DETALHE))
    detalhe_ws.row_dimensions[1].height = 30

    colunas = {nome: i + 1 for i, nome in enumerate(COLUNAS_DETALHE)}
    for item in sorted(
        resultados,
        key=lambda r: (digitos(r.get("Documento", "")), chave_ordem_ano(r)),
    ):
        detalhe_ws.append([item.get(coluna, "") for coluna in COLUNAS_DETALHE])
        numero = detalhe_ws.max_row
        for nome, tabela in ((COLUNA_PAGO, CORES_STATUS),
                             (COLUNA_COBRAVEL, CORES_COBRAVEL)):
            valor = str(item.get(nome, ""))
            fundo, texto = tabela.get(valor, ("FFFFFFFF", "FF000000"))
            celula = detalhe_ws.cell(row=numero, column=colunas[nome])
            celula.fill = PatternFill("solid", start_color=fundo, end_color=fundo)
            celula.font = Font(color=texto, bold=True, size=10)
        fundo, texto = CORES_NIVEL.get(
            str(item.get(COLUNA_NIVEL, "")), ("FFFFFFFF", "FF000000")
        )
        celula_nivel = detalhe_ws.cell(
            row=numero, column=colunas[COLUNA_NIVEL]
        )
        celula_nivel.fill = PatternFill("solid", start_color=fundo, end_color=fundo)
        celula_nivel.font = Font(color=texto, bold=True, size=10)
        for nome in ("Valor (R$)", "Pago (R$)"):
            celula = detalhe_ws.cell(row=numero, column=colunas[nome])
            # Texto formatado: formato numérico com vírgula muda com o
            # idioma do Excel e pode mostrar "R$ 253,910".
            valor = float(item.get(nome, 0) or 0)
            celula.value = f"R$ {valor:.2f}".replace(".", ",")
            celula.number_format = "General"
            celula.alignment = Alignment(horizontal="right")

    ultima_coluna = get_column_letter(len(COLUNAS_DETALHE))
    detalhe_ws.auto_filter.ref = f"A1:{ultima_coluna}{max(2, detalhe_ws.max_row)}"
    detalhe_ws.freeze_panes = "A2"
    for indice, largura in enumerate(
        (18, 32, 20, 42, 14, 18, 18, 20, 16, 16, 10, 14, 14, 26, 20, 50), start=1
    ):
        detalhe_ws.column_dimensions[get_column_letter(indice)].width = largura
    # Auditoria: todas as situações dos dois sistemas e como cada uma é
    # tratada, para conferir que nenhuma forma de pagamento virou cobrança.
    auditoria = detalhe_wb.create_sheet("Situações (auditoria)")
    auditoria.append(["Sistema", "Situação", COLUNA_PAGO, COLUNA_COBRAVEL])
    _estilo_cabecalho(auditoria, 1, 4)
    for sistema, situacoes in (
        ("ISS Digital", SITUACOES_ISS_DIGITAL),
        ("IPTU Digital", SITUACOES_IPTU_DIGITAL),
    ):
        for situacao in situacoes:
            pago, cobravel = classificar_situacao(situacao, 0.0, 100.0)
            auditoria.append([sistema, situacao, pago, cobravel])
            numero = auditoria.max_row
            for coluna, valor, tabela in (
                (3, pago, CORES_STATUS), (4, cobravel, CORES_COBRAVEL)
            ):
                fundo, texto = tabela.get(valor, ("FFFFFFFF", "FF000000"))
                celula = auditoria.cell(row=numero, column=coluna)
                celula.fill = PatternFill("solid", start_color=fundo, end_color=fundo)
                celula.font = Font(color=texto, bold=True, size=10)
    auditoria.freeze_panes = "A2"
    auditoria.auto_filter.ref = f"A1:D{max(2, auditoria.max_row)}"
    for coluna, largura in ((1, 18), (2, 48), (3, 20), (4, 20)):
        auditoria.column_dimensions[get_column_letter(coluna)].width = largura

    detalhe.parent.mkdir(parents=True, exist_ok=True)
    adicionar_aviso(detalhe_wb)
    detalhe_wb.save(detalhe)
    return destino, detalhe


# ---------------------------------------------------------------------------
# Worker (roda em segundo plano na GUI)
# ---------------------------------------------------------------------------

class VerificarPagosWorker(QThread):
    progresso = pyqtSignal(int, str)
    aviso = pyqtSignal(str)
    documento_concluido = pyqtSignal(dict)
    finalizado = pyqtSignal(str)

    def __init__(self, registros: list[dict], consultar_func, cookie: str = "",
                 cookie_sessao: str = "", destino: Path | None = None,
                 detalhe: Path | None = None, origem: Path | None = None,
                 pausa: float = 1.5, variar_pausa: bool = False,
                 controle: dict | None = None, parent=None):
        super().__init__(parent)
        self.registros = registros
        self.consultar_func = consultar_func
        self.cookie = cookie
        self.cookie_sessao = cookie_sessao
        self.destino = Path(destino) if destino else None
        self.detalhe = Path(detalhe) if detalhe else None
        self.origem = Path(origem) if origem else None
        self.pausa = max(0.0, float(pausa))
        self.variar_pausa = variar_pausa
        self.controle = controle or {}
        self.resultados: list[dict] = []
        self._parar = False

    def stop(self) -> None:
        self._parar = True

    def parar_solicitado(self) -> bool:
        return self._parar

    def run(self) -> None:
        total = len(self.registros)
        erros = 0
        for indice, registro in enumerate(self.registros, start=1):
            if self._parar:
                break
            documento = str(registro.get("Documento", ""))
            nome = str(registro.get("Nome", ""))
            aba = str(registro.get("Aba", ""))
            rotulo = nome or documento or f"linha {indice}"
            self.progresso.emit(
                int((indice - 1) * 100 / max(1, total)),
                f"Verificando pagos {indice}/{total}: {rotulo}",
            )
            try:
                encontrados = self.consultar_func(
                    documento, nome, self.cookie, 30,
                    registro.get("Tipo de Pessoa", ""), self.cookie_sessao,
                    incluir_pagos=True,
                    deve_parar=self.parar_solicitado,
                )
                mensagem = ""
            except Exception as exc:  # noqa: BLE001
                if self._parar:
                    break
                encontrados = []
                mensagem = str(exc)
                erros += 1
                self.aviso.emit(
                    f"{documento}: não foi possível confirmar a situação "
                    f"({exc})."
                )

            itens = classificar_debitos(
                registro.get("Debitos", []), encontrados, mensagem
            )
            # Débitos que o site mostrou e que NÃO estavam na planilha também
            # entram na classificação — é assim a consulta completa descobre
            # pagamentos feitos em dívidas que a migração não listou.
            if not mensagem:
                ja_classificados = {
                    str(item.get("ID Dívida", "") or "").strip()
                    for item in itens
                    if str(item.get("ID Dívida", "") or "").strip()
                }
                for debito in encontrados:
                    if (
                        str(debito.get("id_divida", "") or "").strip()
                        in ja_classificados
                    ):
                        continue
                    situacao = str(debito.get("situacao", "") or "").strip()
                    pago_valor = float(debito.get("pago", 0) or 0)
                    valor = float(debito.get("valor", 0) or 0)
                    sit, cobravel = classificar_situacao(
                        situacao, pago_valor, valor
                    )
                    itens.append({
                        "Tipo de Débito": str(debito.get("tipo_debito", "") or ""),
                        "Competência": str(debito.get("competencia", "") or ""),
                        "PAGO?": sit,
                        "COBRÁVEL?": cobravel,
                        "Situação no site": situacao,
                        "Valor (R$)": valor,
                        "Pago (R$)": pago_valor,
                        "CIF": str(debito.get("cif", "") or ""),
                        "ID Dívida": str(debito.get("id_divida", "") or ""),
                        "Vencimento": str(debito.get("vencimento", "") or ""),
                        COLUNA_NIVEL: nivel_alerta(cobravel),
                        "Critério": "Somente no site",
                        "Observação": (
                            "Débito encontrado no site e não listado na "
                            "entrada — conferir origem."
                        ),
                    })
            itens.sort(key=chave_ordem_ano)
            for item in itens:
                item.update({
                    "Aba": aba,
                    "Nome": nome,
                    "Documento": documento,
                    "CPF/CNPJ": documento,
                })
            self.resultados.extend(itens)
            self.documento_concluido.emit({
                "Documento": documento,
                "Nome": nome,
                "Aba": aba,
                "Itens": itens,
            })

            pausa = float(self.controle.get("pausa", self.pausa))
            if self.variar_pausa or self.controle.get("variar", self.variar_pausa):
                pausa *= random.uniform(0.7, 1.3)
            if indice < total and not self._parar and pausa > 0:
                self.msleep(int(pausa * 1000))

        interrompido = self._parar
        if self.resultados and self.destino and self.detalhe:
            try:
                origem = self.origem or self.destino
                gravar_resultado_pagos(
                    origem, self.destino, self.detalhe, self.resultados
                )
            except Exception as exc:  # noqa: BLE001
                self.finalizado.emit(
                    f"Erro ao gravar relatórios de pagos: {exc}"
                )
                return

        if interrompido:
            mensagem = (
                f"Verificação de pagos interrompida: {len(self.resultados)} "
                f"débito(s) processado(s). Os resultados parciais foram salvos."
            )
        else:
            cobraveis = sum(
                1 for item in self.resultados
                if item.get("COBRÁVEL?") == "SIM"
            )
            mensagem = (
                f"Concluído: {len(self.registros)} documento(s) verificado(s), "
                f"{len(self.resultados)} débito(s) cruzado(s), "
                f"{cobraveis} cobrável(is)"
            )
            if erros:
                mensagem += f", {erros} consulta(s) sem confirmação"
            mensagem += "."
        self.progresso.emit(100, mensagem)
        self.finalizado.emit(mensagem)


def caminhos_saida(caminho_entrada: Path | None, pasta_relatorios: Path,
                   stem_base: str) -> tuple[Path, Path]:
    """Nomes padronizados dos dois arquivos gerados pela verificação."""
    stem = caminho_entrada.stem if caminho_entrada else stem_base
    return (
        pasta_relatorios / f"{stem}_VERIFICACAO_PAGOS.xlsx",
        pasta_relatorios / f"{stem}_DETALHE_PAGOS.xlsx",
    )
