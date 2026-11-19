"""Leitura e gravação de planilhas em .xlsx e .csv."""

from __future__ import annotations

import csv
from pathlib import Path

EXTENSOES_SUPORTADAS = (".xlsx", ".csv")
SEPARADOR_CSV_SAIDA = ";"  # Excel em português usa ponto e vírgula


def ler_planilha(caminho: str | Path) -> tuple[list[str], list[dict[str, str]]]:
    """Lê a planilha e devolve (cabeçalho, linhas). Linhas vazias são descartadas."""
    p = Path(caminho)
    ext = p.suffix.lower()
    if ext == ".csv":
        return _ler_csv(p)
    if ext == ".xlsx":
        return _ler_xlsx(p)
    raise ValueError(f"Formato não suportado: {ext or 'sem extensão'}. Use .xlsx ou .csv")


def gravar_planilha(
    caminho_saida: str | Path,
    cabecalho: list[str],
    linhas: list[dict[str, str]],
) -> Path:
    """Grava no formato indicado pela extensão de caminho_saida."""
    p = Path(caminho_saida)
    ext = p.suffix.lower()
    if ext == ".csv":
        _gravar_csv(p, cabecalho, linhas)
    elif ext == ".xlsx":
        _gravar_xlsx(p, cabecalho, linhas)
    else:
        raise ValueError(f"Formato de saída não suportado: {ext}")
    return p


def caminho_saida_padrao(caminho_entrada: str | Path) -> Path:
    p = Path(caminho_entrada)
    return p.with_name(f"{p.stem}_resultado{p.suffix.lower()}")


# ---------- CSV ----------

def _ler_csv(p: Path) -> tuple[list[str], list[dict[str, str]]]:
    # utf-8-sig remove o BOM que o Excel costuma colocar no início do arquivo
    with p.open(newline="", encoding="utf-8-sig") as f:
        amostra = f.read(4096)
        f.seek(0)
        try:
            dialeto = csv.Sniffer().sniff(amostra, delimiters=",;\t")
        except csv.Error:
            dialeto = csv.excel
        leitor = csv.reader(f, dialect=dialeto)
        cabecalho_raw = next(leitor, [])
        cabecalho = [c.strip() for c in cabecalho_raw]
        linhas = []
        for valores in leitor:
            if not any(v.strip() for v in valores):
                continue
            linhas.append({
                h: (valores[i].strip() if i < len(valores) else "")
                for i, h in enumerate(cabecalho) if h
            })
    return [h for h in cabecalho if h], linhas


def _gravar_csv(p: Path, cabecalho: list[str], linhas: list[dict[str, str]]) -> None:
    with p.open("w", newline="", encoding="utf-8-sig") as f:
        escritor = csv.DictWriter(
            f, fieldnames=cabecalho, delimiter=SEPARADOR_CSV_SAIDA,
            extrasaction="ignore",
        )
        escritor.writeheader()
        escritor.writerows(linhas)


# ---------- XLSX ----------

def _ler_xlsx(p: Path) -> tuple[list[str], list[dict[str, str]]]:
    from openpyxl import load_workbook  # import local: só exigido se houver .xlsx

    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        ws = wb.active
        iterador = ws.iter_rows(values_only=True)
        cabecalho_raw = next(iterador, None) or ()
        cabecalho = ["" if c is None else str(c).strip() for c in cabecalho_raw]
        linhas = []
        for valores in iterador:
            if valores is None or all(v in (None, "") for v in valores):
                continue
            linhas.append({
                h: ("" if i >= len(valores) or valores[i] is None else str(valores[i]).strip())
                for i, h in enumerate(cabecalho) if h
            })
    finally:
        wb.close()
    return [h for h in cabecalho if h], linhas


def _gravar_xlsx(p: Path, cabecalho: list[str], linhas: list[dict[str, str]]) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"
    ws.append(cabecalho)
    for linha in linhas:
        ws.append([linha.get(h, "") for h in cabecalho])
    wb.save(p)
