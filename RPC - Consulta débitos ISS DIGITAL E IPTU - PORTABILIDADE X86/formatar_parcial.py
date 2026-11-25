"""Formatação final do relatório parcial: Status, MOTIVO e duplicatas.

Chamado por ``gravar_planilha`` quando o destino é ``*_resultado_parcial.xlsx``,
para que uma regravação pelo programa não desfaça a classificação feita.
"""

from __future__ import annotations

import collections
import re

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from regras_debito import sem_acento
from verificar_pagos import chave_competencia, classificar_situacao

VERDE = ("FFC6EFCE", "FF006100")
VERMELHO = ("FFFFC7CE", "FF9C0006")
LARANJA = ("FFFFF2CC", "FF7A6000")
AMARELO = ("FFFFEB9C", "FF9C6500")
DUPLICADO = PatternFill("solid", start_color="FFFFEB9C", end_color="FFFFEB9C")
ZEBRA = ("FFFFFFFF", "FFF2F2F2")


def _digitos(valor: object) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def _ano(valor: object) -> int:
    ano = chave_competencia(valor)[0]
    if ano:
        return ano
    encontrado = re.search(r"(20\d{2})", str(valor or ""))
    return int(encontrado.group(1)) if encontrado else 0


def _maior_sequencia(anos: list[int]) -> list[int]:
    anos = sorted(set(a for a in anos if a))
    if not anos:
        return []
    melhor: list[int] = []
    atual = [anos[0]]
    for ano in anos[1:]:
        if ano == atual[-1] + 1:
            atual.append(ano)
        else:
            if len(atual) > len(melhor):
                melhor = atual
            atual = [ano]
    return atual if len(atual) > len(melhor) else melhor


def formatar_parcial(ws, linhas: list[dict]) -> None:
    """Aplica Status colorido, MOTIVO e destaque de débitos duplicados."""
    cabecalho = [celula.value for celula in ws[1]]
    if "MOTIVO" not in cabecalho:
        col_motivo = ws.max_column + 1
        celula = ws.cell(row=1, column=col_motivo, value="MOTIVO")
        celula.fill = PatternFill(
            "solid", start_color="FF1F4E79", end_color="FF1F4E79"
        )
        celula.font = Font(color="FFFFFFFF", bold=True, size=11)
        celula.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_motivo)].width = 70
    else:
        col_motivo = cabecalho.index("MOTIVO") + 1

    col_status = cabecalho.index("Status") + 1
    col_nome = cabecalho.index("Nome") + 1
    col_cpf = cabecalho.index("CPF") + 1
    col_cif = cabecalho.index("CIF") + 1
    col_tipo = cabecalho.index("Tipo de Débito") + 1
    col_comp = cabecalho.index("Competência") + 1
    col_valor = cabecalho.index("Valor (R$)") + 1
    col_sit = (
        cabecalho.index("Situação (site)") + 1
        if "Situação (site)" in cabecalho else None
    )

    # Sequência de IPTU por documento.
    por_documento: dict[str, list[int]] = collections.defaultdict(list)
    for item in linhas:
        doc = _digitos(item.get("CPF"))
        if (
            len(doc) in (11, 14)
            and sem_acento(item.get("Tipo de Débito", "")).startswith("iptu")
        ):
            por_documento[doc].append(
                _ano(item.get("Competência")) or _ano(item.get("Vencimento"))
            )
    sequencia = {
        doc: ", ".join(str(a) for a in _maior_sequencia(anos))
        for doc, anos in por_documento.items()
        if len(_maior_sequencia(anos)) >= 3
    }

    vistos: collections.Counter[tuple] = collections.Counter()
    pessoa_atual = None
    indice_pessoa = 0

    for numero, item in enumerate(linhas, start=2):
        doc = _digitos(item.get("CPF"))
        situacao = str(item.get("Situação (site)", "") or "")
        _, cobravel = classificar_situacao(situacao, 0, 0)
        status_antigo = str(item.get("Status", "") or "").strip().upper()

        chave = (
            doc,
            str(item.get("CIF", "") or "").strip(),
            sem_acento(item.get("Tipo de Débito", "")),
            str(item.get("Competência", "") or "").strip(),
            str(item.get("Valor (R$)", "") or "").strip(),
        )
        vistos[chave] += 1
        duplicado = vistos[chave] > 1

        if duplicado:
            status = "DUPLICADO — CONFERIR"
            motivo = "Débito aparenta ser duplicado — conferir antes de cobrar."
            cor = AMARELO
        elif status_antigo == "COBRÁVEL" and cobravel != "NÃO":
            # Uma reconsulta que não trouxe confirmação (sessão expirada,
            # site fora, situação vazia) não pode desfazer uma classificação
            # verde já aprovada. Só uma situação encerrada informada pelo
            # site reverte a cobrança.
            status = "COBRÁVEL"
            motivo = (
                "Classificação anterior mantida: o site não informou "
                "pagamento/encerramento deste débito."
            )
            cor = VERDE
        elif (
            doc in sequencia
            and cobravel != "NÃO"
            and status_antigo != "DOCUMENTO INVÁLIDO"
        ):
            status = "COBRÁVEL"
            motivo = f"IPTU em anos consecutivos ({sequencia[doc]})."
            cor = VERDE
        elif (
            cobravel == "NÃO"
            or status_antigo in ("NÃO COBRAR", "DOCUMENTO INVÁLIDO")
        ):
            status = "NÃO COBRAR"
            if cobravel == "NÃO":
                motivo = (
                    f"Situação informada pelo site: "
                    f"{situacao or 'pagamento/encerramento'}."
                )
            elif status_antigo == "DOCUMENTO INVÁLIDO":
                motivo = "Documento inválido — conferir a planilha."
            else:
                motivo = (
                    "Classificação anterior mantida: débito já marcado "
                    "como não cobrável."
                )
            cor = VERMELHO
        else:
            status = "CLASSIFICADO — INDETERMINADO"
            motivo = (
                "Sem confirmação de pagamento no site; "
                "aguardando decisão do gestor."
            )
            cor = LARANJA

        pessoa = str(item.get("Nome", "") or "").strip().upper()
        if pessoa != pessoa_atual:
            pessoa_atual = pessoa
            indice_pessoa += 1
        fundo = (
            DUPLICADO if duplicado else PatternFill(
                "solid", start_color=ZEBRA[indice_pessoa % 2],
                end_color=ZEBRA[indice_pessoa % 2],
            )
        )
        for coluna in range(1, ws.max_column + 1):
            celula = ws.cell(row=numero, column=coluna)
            celula.fill = fundo
            celula.font = Font(size=10)
            celula.alignment = Alignment(vertical="center")

        celula_status = ws.cell(row=numero, column=col_status, value=status)
        celula_status.fill = PatternFill(
            "solid", start_color=cor[0], end_color=cor[0]
        )
        celula_status.font = Font(color=cor[1], bold=True, size=10)
        celula_status.alignment = Alignment(
            horizontal="center", vertical="center"
        )
        celula_motivo = ws.cell(row=numero, column=col_motivo, value=motivo)
        celula_motivo.alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(row=numero, column=col_valor).alignment = Alignment(
            horizontal="right", vertical="center"
        )
        ws.row_dimensions[numero].height = 20

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = (
        f"A1:{get_column_letter(ws.max_column)}{max(2, ws.max_row)}"
    )
