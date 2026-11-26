#!/usr/bin/env python3
"""Preenche a coluna OBSERVAÇÃO (débitos) do relatório contábil.

Para cada CPF/CNPJ, a coluna recebe TODOS os débitos daquela pessoa em
UMA ÚNICA LINHA, listando somente o tipo de débito e a competência:

    IPTU (M) – fev/22; TAXA DE LICENÇA PARA LOCALIZAÇÃO E FUNCIONAMENTO (M) – fev/22

Débitos repetidos (parcelas com o mesmo tipo/competência) aparecem uma
única vez, já que a observação mostra apenas esses dois campos.

O arquivo de saída é sempre o mesmo:
    relatorios/RELATORIO_CONTABILIDADE_2022_ANALISTAS_1_LINHA_TODOS_OS_DEBITOS.xlsx
Se ele já existir, é atualizado — nunca é criado outro arquivo.

Uso:
    python3 preencher_observacao.py
    python3 preencher_observacao.py --relatorio planilha.xlsx --debitos resultado.xlsx --saida out.xlsx
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from collections import OrderedDict
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from disclaimer import adicionar_aviso
from regras_debito import eh_debito_cobravel

BASE = Path(__file__).resolve().parent
RELATORIO_PADRAO = BASE / "RELATORIO_CONTABILIDADE_2022_ANALISTAS.xlsx"
DEBITOS_PADRAO = (
    BASE
    / "relatorios"
    / "RELATORIO_CONTABILIDADE_2022_ANALISTAS_resultado.xlsx"
)
SAIDA_PADRAO = (
    BASE
    / "relatorios"
    / "RELATORIO_CONTABILIDADE_2022_ANALISTAS_1_LINHA_TODOS_OS_DEBITOS.xlsx"
)

SEPARADOR = " – "  # tipo de débito – competência
SEPARADOR_ITENS = "; "  # separa os débitos dentro da MESMA linha
LARGURA_OBSERVACAO = 120  # deixa a linha toda visível no Excel
ABA_DEBITOS = "Resultado"


def normalizar(texto: object) -> str:
    """Remove acentos, maiúsculas e pontuação para comparar cabeçalhos."""
    if texto is None:
        return ""
    sem_acento = unicodedata.normalize("NFKD", str(texto))
    sem_acento = "".join(c for c in sem_acento if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", sem_acento.lower()).strip()


def digitos(valor: object) -> str:
    """Só os números do CPF/CNPJ (ignora pontos, barras e espaços)."""
    return re.sub(r"\D", "", str(valor or ""))


def tem_coluna_observacao(caminho: Path) -> bool:
    """Verifica se a planilha possui a coluna OBSERVAÇÃO (débitos)."""
    wb = load_workbook(caminho, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            if localizar_cabecalho(ws) is not None:
                return True
    finally:
        wb.close()
    return False


def ano_debito(competencia: str) -> int:
    """Ano da competência (ex.: 2024, 1/2026, fev/22)."""
    texto = str(competencia or "")
    match = re.search(r"(20\d{2})", texto)
    if match:
        return int(match.group(1))
    match = re.search(r"/\s*(\d{2})\b", texto)
    return 2000 + int(match.group(1)) if match else 0


_MESES = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


def chave_ordem_competencia(texto: str) -> tuple[int, int]:
    """(ano, mês) para ordenar débitos; competência final fica por último."""
    valor = str(texto or "").lower()
    mes = None
    nome = re.search(
        r"\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\b", valor
    )
    numerico = re.search(r"\b(\d{1,2})\s*[/\-]\s*(\d{2}|\d{4})\b", valor)
    if nome:
        mes = _MESES[nome.group(1)]
        ano = re.search(r"\b(\d{2}|\d{4})\b", valor[nome.end():])
    elif numerico:
        mes = int(numerico.group(1))
        ano = re.search(r"\d{2}|\d{4}", numerico.group(2))
    else:
        ano = re.search(r"\b(20\d{2})\b", valor)
    ano = int(ano.group(0)) if ano else 0
    if ano and ano < 100:
        ano += 2000
    return (ano, mes if mes else 13)


def carregar_debitos(caminho: Path) -> dict[str, list[str]]:
    """Lê a planilha de débitos e devolve {CPF/CNPJ: [tipo – competência]}."""
    wb = load_workbook(caminho, read_only=True, data_only=True)
    try:
        ws = wb[ABA_DEBITOS] if ABA_DEBITOS in wb.sheetnames else wb.active
        linhas = ws.iter_rows(values_only=True)
        cabecalho = [normalizar(c) for c in next(linhas, ())]

        def indice(*termos: str) -> int | None:
            for i, nome in enumerate(cabecalho):
                if all(termo in nome for termo in termos):
                    return i
            return None

        col_doc = indice("cpf")
        if col_doc is None:
            col_doc = indice("cnpj")
        if col_doc is None:
            raise ValueError("Planilha de débitos sem coluna CPF/CNPJ.")
        col_tipo = indice("tipo", "debito")
        col_comp = indice("compet")
        if col_tipo is None or col_comp is None:
            raise ValueError("Planilha de débitos sem tipo ou competência.")
        col_status = indice("status")
        col_erro = indice("erro")

        debitos: dict[str, OrderedDict[str, int]] = {}
        for linha in linhas:
            if col_doc >= len(linha):
                continue
            doc = digitos(linha[col_doc])
            if len(doc) not in (11, 14):
                continue

            status = str(linha[col_status] or "") if col_status is not None and col_status < len(linha) else ""
            erro = str(linha[col_erro] or "").strip() if col_erro is not None and col_erro < len(linha) else ""
            tipo = str(linha[col_tipo] or "").strip()
            comp = str(linha[col_comp] or "").strip()
            # Regra compartilhada: rejeita "Sem débitos", "-", status
            # diverso de OK/VERIFICADO, erro/ressalva e legado não confirmado.
            if not eh_debito_cobravel({
                "Status": status, "Erro": erro,
                "Tipo de Débito": tipo, "Competência": comp,
            }):
                continue
            debitos.setdefault(doc, OrderedDict())[
                tipo + SEPARADOR + comp
            ] = ano_debito(comp)
    finally:
        wb.close()

    return {
        doc: [texto for texto, _ in sorted(
            pares.items(), key=lambda par: (-par[1], par[0])
        )]
        for doc, pares in debitos.items()
    }


def localizar_cabecalho(ws) -> tuple[int, int, int] | None:
    """Devolve (linha, coluna CPF/CNPJ, coluna OBSERVAÇÃO), 1-based."""
    for numero, linha in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        nomes = [normalizar(c) for c in linha]
        col_doc = next(
            (i for i, n in enumerate(nomes, start=1) if "cpf" in n or "cnpj" in n),
            None,
        )
        col_obs = next(
            (i for i, n in enumerate(nomes, start=1) if "observa" in n),
            None,
        )
        if col_doc is not None and col_obs is not None:
            return numero, col_doc, col_obs
    return None


def registros_observacao_vazia(caminho: Path, aba: str | None = None) -> list[dict]:
    """Devolve os documentos das linhas com OBSERVAÇÃO (débitos) vazia.

    Se ``aba`` for informada e existir no arquivo, varre SOMENTE ela (para
    não misturar CPFs de uma aba com CNPJs de outra). Caso contrário, varre
    todas as abas, menos Resumo. Devolve no máximo um registro por documento.
    """
    wb = load_workbook(caminho, read_only=True, data_only=True)
    vistos: set[str] = set()
    registros: list[dict] = []
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
                ws.iter_rows(min_row=linha_cab, max_row=linha_cab, values_only=True),
                (),
            )
            col_nome = next(
                (i for i, valor in enumerate(cabecalho, start=1)
                 if "nome" in normalizar(valor)),
                None,
            )
            for linha in ws.iter_rows(min_row=linha_cab + 1, values_only=True):
                doc = digitos(linha[col_doc - 1] if col_doc <= len(linha) else "")
                if len(doc) not in (11, 14) or doc in vistos:
                    continue
                observacao = (
                    str(linha[col_obs - 1] or "").strip()
                    if col_obs <= len(linha) else ""
                )
                if observacao:
                    continue
                vistos.add(doc)
                nome = (
                    str(linha[col_nome - 1] or "").strip()
                    if col_nome and col_nome <= len(linha) else ""
                )
                registros.append({
                    "Nome": nome,
                    "CPF": doc if len(doc) == 11 else "",
                    "CNPJ": doc if len(doc) == 14 else "",
                    "Tipo de Pessoa": "FISICA" if len(doc) == 11 else "JURIDICA",
                })
    finally:
        wb.close()
    return registros


def preencher(relatorio: Path, debitos: dict[str, list[str]], destino: Path) -> dict[str, int]:
    """Preenche a OBSERVAÇÃO no arquivo de destino (atualiza se já existir)."""
    relatorio, destino = Path(relatorio), Path(destino)
    origem = destino if destino.exists() else relatorio
    wb = load_workbook(origem)
    contagens: dict[str, int] = {}

    for ws in wb.worksheets:
        if normalizar(ws.title) == "resumo":
            continue
        local = localizar_cabecalho(ws)
        if local is None:
            continue
        linha_cab, col_doc, col_obs = local
        preenchidas = 0
        ws.column_dimensions[get_column_letter(col_obs)].width = LARGURA_OBSERVACAO

        for numero in range(linha_cab + 1, ws.max_row + 1):
            doc = digitos(ws.cell(row=numero, column=col_doc).value)
            observacoes = debitos.get(doc)
            if not observacoes:
                continue
            celula = ws.cell(row=numero, column=col_obs)
            celula.value = SEPARADOR_ITENS.join(observacoes)
            celula.alignment = Alignment(wrap_text=False, vertical="center")
            preenchidas += 1

        contagens[ws.title] = preenchidas

    # Atualiza a coluna "Com observação" do painel Resumo, se existir.
    for ws in wb.worksheets:
        if normalizar(ws.title) != "resumo":
            continue
        for numero in range(1, min(ws.max_row, 20) + 1):
            valores = [normalizar(ws.cell(row=numero, column=c).value) for c in range(1, ws.max_column + 1)]
            if "planilha" not in valores or not any("observa" in v for v in valores):
                continue
            col_planilha = valores.index("planilha") + 1
            col_obs = next((i + 1 for i, v in enumerate(valores) if "observa" in v), None)
            if col_obs is None:
                break
            for linha_dados in range(numero + 1, ws.max_row + 1):
                nome = str(ws.cell(row=linha_dados, column=col_planilha).value or "").strip()
                if nome in contagens:
                    ws.cell(row=linha_dados, column=col_obs).value = contagens[nome]
                elif normalizar(nome) == "total":
                    ws.cell(row=linha_dados, column=col_obs).value = sum(contagens.values())
            break

    destino.parent.mkdir(parents=True, exist_ok=True)
    adicionar_aviso(wb)
    wb.save(destino)
    return contagens


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--relatorio", type=Path, default=RELATORIO_PADRAO)
    parser.add_argument("--debitos", type=Path, default=DEBITOS_PADRAO)
    parser.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    args = parser.parse_args()

    debitos = carregar_debitos(args.debitos)
    contagens = preencher(args.relatorio, debitos, args.saida)

    print(f"CPFs/CNPJs com débito: {len(debitos)}")
    for aba, total in contagens.items():
        print(f"  {aba}: {total} guia(s) com observação")
    print(f"Total de guias preenchidas: {sum(contagens.values())}")
    print(f"Arquivo gerado: {args.saida}")


if __name__ == "__main__":
    main()
