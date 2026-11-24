"""Aviso sobre origem, cruzamento de dados e responsabilidade dos relatórios.

O texto é gravado como primeira planilha ("AVISO") de todo arquivo gerado,
sem alterar as abas de dados nem os cabeçalhos usados pelo programa.
"""

from __future__ import annotations

import textwrap
from datetime import datetime

from openpyxl.styles import Alignment, Font, PatternFill

NOME_ABA = "AVISO"
NOME_LEGENDA = "LEGENDA"

ITENS_LEGENDA = (
    ("COR / DESTAQUE", "SIGNIFICADO", "O QUE O ANALISTA DEVE FAZER"),
    ("COBRÁVEL", "Verde: débito classificado como cobrável",
     "Conferir no sistema oficial antes de qualquer cobrança"),
    ("CLASSIFICADO — INDETERMINADO", "Laranja: análise sem confirmação do site",
     "Aguardar decisão do gestor; não cobrar automaticamente"),
    ("NÃO COBRAR", "Vermelho: pagamento ou encerramento informado",
     "Descartar da cobrança"),
    ("DUPLICADO — CONFERIR",
     "Amarelo na linha inteira e texto no MOTIVO",
     "Conferir e desconsiderar as repetições"),
    ("R$ 253,91", "Valor em reais com duas casas decimais",
     "Conferir o valor no sistema oficial"),
    ("MOTIVO", "Justificativa da classificação do débito",
     "Ler antes de decidir; não é comprovação de pagamento"),
)

CORES_LEGENDA = (
    ("FFC6EFCE", "FF006100"),
    ("FFFFF2CC", "FF7A6000"),
    ("FFFFC7CE", "FF9C0006"),
    ("FFFFEB9C", "FF9C6500"),
    (None, None),
    (None, None),
)


TEXTO_AVISO = (
    "AVISO SOBRE A ORIGEM, O CRUZAMENTO DE DADOS E O USO DESTE RELATÓRIO",
    "",
    "1. FONTE DOS DADOS — Os dados aqui contidos foram obtidos por consulta "
    "aos sistemas oficiais ISS Digital e IPTU Digital, mantidos "
    "por Desenvolve Cidade — Prefeitura Municipal de Santana do Araguaia/PA "
    "(doravante \"SISTEMA OFICIAL\").",
    "",
    "2. CRUZAMENTO E CLASSIFICAÇÃO — Todas as comparações, classificações e "
    "validações registradas neste arquivo, inclusive a situação de pagamento "
    "de cada débito (DÉBITO, PAGO, PAGO MANUAL, CANCELADO, PRESCRITO, AJUIZADO "
    "etc.), a indicação de cobrabilidade, o ID Dívida, o CIF e o endereço dos "
    "imóveis, derivam exclusivamente das informações disponíveis na base de "
    "dados do SISTEMA OFICIAL na data da consulta.",
    "",
    "3. NATUREZA DO DOCUMENTO — Este arquivo é relatório auxiliar de apoio à "
    "análise e à gestão. Não é certidão, comprovante "
    "de pagamento, lançamento tributário nem ato administrativo, e não "
    "substitui a consulta direta ao SISTEMA OFICIAL.",
    "",
    "4. ATUALIDADE — As informações podem ter sido alteradas no SISTEMA "
    "OFICIAL após a data de extração. Registros marcados como "
    "\"CLASSIFICADO — INDETERMINADO\", \"NÃO CONFIRMADO\" ou "
    "\"NÃO VERIFICADO\" indicam que a consulta não foi suficiente para "
    "confirmar o dado e NÃO devem ser tratados como débito exigível.",
    "",
    "5. RESPONSABILIDADE — A decisão final sobre qualquer cobrança, "
    "suspensão, baixa ou cancelamento de débito compete ao gestor público "
    "competente, com base no SISTEMA OFICIAL e nos autos do processo "
    "administrativo. A geração deste arquivo não constitui certificação, "
    "declaração ou ato administrativo por parte de quem o operou. "
    "O responsável pela geração não responde por decisões tomadas sem a "
    "conferência no SISTEMA OFICIAL, nem por divergências entre este "
    "arquivo e os sistemas oficiais.",
    "",
    "6. USO INTERNO — Destinado exclusivamente ao uso interno da Secretaria "
    "de Tributos e Terras. Não deve ser divulgado a terceiros nem utilizado "
    "para finalidade diversa da aqui descrita.",
    "",
    "7. LEGENDA DAS CORES E DAS DUPLICATAS — Verde: débito classificado como "
    "COBRÁVEL; laranja: CLASSIFICADO — INDETERMINADO; vermelho: NÃO COBRAR "
    "(pagamento ou encerramento informado); roxo na linha inteira: débito "
    "repetido na listagem, com status 'DUPLICADO — CONFERIR' e linha amarela; o MOTIVO diz "
    "'Débito aparenta ser duplicado'. Débitos duplicados devem ser conferidos e desconsiderados nas "
    "repetições antes de qualquer cobrança.",
)


def adicionar_aviso(wb) -> None:
    """Insere/substitui a aba AVISO na posição 0 do workbook."""
    existente = next(
        (ws for ws in wb.worksheets if ws.title.strip().upper() == NOME_ABA),
        None,
    )
    if existente is not None:
        wb.remove(existente)

    ws = wb.create_sheet(NOME_ABA, 0)
    ws.sheet_view.showGridLines = False
    # Coluna única e estreita: o texto fica em linhas verticais curtas, uma
    # frase por linha, fácil de ler sem rolagem horizontal.
    ws.column_dimensions["A"].width = 95
    ws.freeze_panes = "A2"

    titulo = ws.cell(row=1, column=1, value=TEXTO_AVISO[0])
    titulo.font = Font(bold=True, size=13, color="FFFFFFFF")
    titulo.fill = PatternFill("solid", start_color="FF1F4E79", end_color="FF1F4E79")
    titulo.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    linha = 2
    for texto in TEXTO_AVISO[1:]:
        if not texto:
            ws.row_dimensions[linha].height = 9
            linha += 1
            continue
        for parte in textwrap.wrap(
            texto, width=95, break_long_words=True, break_on_hyphens=False
        ):
            celula = ws.cell(row=linha, column=1, value=parte)
            celula.font = Font(size=11)
            celula.alignment = Alignment(vertical="center", wrap_text=False)
            ws.row_dimensions[linha].height = 16
            linha += 1

    ws.row_dimensions[linha].height = 9
    linha += 1
    gerado = ws.cell(
        row=linha,
        column=1,
        value=(
            "Relatório emitido em "
            + datetime.now().strftime("%d/%m/%Y às %H:%M")
            + "."
        ),
    )
    gerado.font = Font(italic=True, size=10, color="FF555555")
    ws.row_dimensions[linha].height = 16
    adicionar_legenda(wb)


def adicionar_legenda(wb) -> None:
    """Cria a aba LEGENDA com amostras visuais das cores de alerta."""
    existente = next(
        (ws for ws in wb.worksheets
         if ws.title.strip().upper() == NOME_LEGENDA),
        None,
    )
    if existente is not None:
        wb.remove(existente)

    # Posição 1: depois do AVISO e antes dos dados.
    ws = wb.create_sheet(NOME_LEGENDA, 1)
    ws.sheet_view.showGridLines = False
    for coluna, largura in (("A", 34), ("B", 52), ("C", 58)):
        ws.column_dimensions[coluna].width = largura

    for numero, itens in enumerate(ITENS_LEGENDA, start=1):
        fundo, texto = (
            CORES_LEGENDA[numero - 2] if 0 <= numero - 2 < len(CORES_LEGENDA)
            else (None, None)
        )
        for coluna, valor in enumerate(itens, start=1):
            celula = ws.cell(row=numero, column=coluna, value=valor)
            celula.alignment = Alignment(vertical="center", wrap_text=True)
            celula.font = Font(size=11)
            if numero == 1:
                celula.fill = PatternFill(
                    "solid", start_color="FF1F4E79", end_color="FF1F4E79"
                )
                celula.font = Font(bold=True, color="FFFFFFFF", size=11)
            elif fundo:
                celula.fill = PatternFill(
                    "solid", start_color=fundo, end_color=fundo
                )
                celula.font = Font(bold=True, color=texto, size=11)
        ws.row_dimensions[numero].height = 34

    aviso = ws.cell(
        row=len(ITENS_LEGENDA) + 2,
        column=1,
        value=(
            "As cores são apoio visual. A decisão final de cobrar ou baixar "
            "compete ao gestor público, com conferência no sistema oficial."
        ),
    )
    aviso.font = Font(italic=True, size=10, color="FF555555")
    aviso.alignment = Alignment(wrap_text=True)
