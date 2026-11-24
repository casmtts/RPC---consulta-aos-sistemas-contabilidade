#!/usr/bin/env python3
"""
Consulta a situação cadastral de CNPJs na Receita Federal.

Usa a BrasilAPI (espelho dos dados da Receita Federal) e, se ela falhar,
a ReceitaWS como alternativa. Usa somente a biblioteca padrão do Python.

Exemplos:
  python consulta_cnpj.py 11.222.333/0001-81
  python consulta_cnpj.py 11222333000181 00000000000191
  python consulta_cnpj.py -f cnpjs.txt -o resultado.csv

O arquivo de entrada tem um CNPJ por linha (linhas vazias e iniciadas com # são ignoradas).
"""

import argparse
import csv
import json
import re
import sys
import time
import urllib.error
import urllib.request

BRASILAPI_URL = "https://brasilapi.com.br/api/cnpj/v1/{}"
RECEITAWS_URL = "https://www.receitaws.com.br/v1/cnpj/{}"
HEADERS = {"User-Agent": "consulta-cnpj-py/1.0", "Accept": "application/json"}

CAMPOS_CSV = [
    "cnpj", "situacao", "data_situacao", "razao_social", "nome_fantasia",
    "municipio", "uf", "fonte", "erro",
]


def limpar(cnpj: str) -> str:
    """Remove pontos, barras e traços. Aceita o formato alfanumérico (a partir de 2026)."""
    return re.sub(r"[^A-Za-z0-9]", "", cnpj).upper()


def cnpj_valido(cnpj: str) -> bool:
    """Confere o tamanho e os dois dígitos verificadores."""
    c = limpar(cnpj)
    if len(c) != 14 or not c[:12].isalnum() or not c[12:].isdigit():
        return False
    if len(set(c)) == 1:  # ex.: 00000000000000 não é válido
        return False

    valores = [ord(ch) - 48 for ch in c]  # dígitos valem 0-9; letras valem 17-42
    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    pesos2 = [6] + pesos1

    def digito(n: int, pesos: list[int]) -> int:
        soma = sum(v * p for v, p in zip(valores[:n], pesos))
        resto = 11 - (soma % 11)
        return 0 if resto > 9 else resto

    return digito(12, pesos1) == valores[12] and digito(13, pesos2) == valores[13]


def requisitar(url: str, timeout: int = 15):
    """Faz a requisição GET e devolve (status_http, json). Não lança erro em HTTP 4xx/5xx."""
    req = urllib.request.Request(url, headers=HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            corpo = json.loads(e.read().decode("utf-8"))
        except Exception:
            corpo = {"message": str(e.reason)}
        return e.code, corpo


def consultar_brasilapi(cnpj: str) -> dict:
    status, d = requisitar(BRASILAPI_URL.format(cnpj))
    if status != 200:
        raise RuntimeError(f"HTTP {status}: {d.get('message', d)}")
    return {
        "fonte": "BrasilAPI",
        "cnpj": cnpj,
        "situacao": d.get("descricao_situacao_cadastral"),
        "data_situacao": d.get("data_situacao_cadastral"),
        "razao_social": d.get("razao_social"),
        "nome_fantasia": d.get("nome_fantasia"),
        "municipio": d.get("municipio"),
        "uf": d.get("uf"),
    }


def consultar_receitaws(cnpj: str) -> dict:
    status, d = requisitar(RECEITAWS_URL.format(cnpj))
    if status != 200 or d.get("status") == "ERROR":
        raise RuntimeError(d.get("message", f"HTTP {status}"))
    return {
        "fonte": "ReceitaWS",
        "cnpj": cnpj,
        "situacao": d.get("situacao"),
        "data_situacao": d.get("data_situacao"),
        "razao_social": d.get("nome"),
        "nome_fantasia": d.get("fantasia"),
        "municipio": d.get("municipio"),
        "uf": d.get("uf"),
    }


def consultar(cnpj_bruto: str) -> dict:
    cnpj = limpar(cnpj_bruto)
    if not cnpj_valido(cnpj):
        return {"cnpj": cnpj_bruto, "erro": "CNPJ inválido (dígitos verificadores não conferem)"}

    erros = []
    for consulta in (consultar_brasilapi, consultar_receitaws):
        try:
            return consulta(cnpj)
        except Exception as e:  # tenta o próximo provedor
            erros.append(f"{consulta.__name__}: {e}")
    return {"cnpj": cnpj, "erro": " | ".join(erros)}


def ler_cnpjs(caminho: str) -> list[str]:
    with open(caminho, encoding="utf-8") as f:
        return [l.strip() for l in f if l.strip() and not l.lstrip().startswith("#")]


def main() -> int:
    parser = argparse.ArgumentParser(description="Consulta situação cadastral de CNPJs (Receita Federal).")
    parser.add_argument("cnpjs", nargs="*", help="CNPJs a consultar")
    parser.add_argument("-f", "--arquivo", help="arquivo .txt com um CNPJ por linha")
    parser.add_argument("-o", "--saida", help="salva o resultado em CSV")
    parser.add_argument("-i", "--intervalo", type=float, default=1.5,
                        help="segundos entre consultas (a ReceitaWS aceita ~3 por minuto no plano gratuito)")
    args = parser.parse_args()

    lista = list(args.cnpjs)
    if args.arquivo:
        lista += ler_cnpjs(args.arquivo)
    if not lista:
        parser.error("informe CNPJs ou um arquivo com -f")

    resultados = []
    for i, cnpj in enumerate(lista):
        r = consultar(cnpj)
        resultados.append(r)
        if "erro" in r:
            print(f"[ERRO] {r['cnpj']}: {r['erro']}")
        else:
            print(f"{r['cnpj']} | {r['situacao']} desde {r['data_situacao']} | "
                  f"{r['razao_social']} ({r['uf']}) [{r['fonte']}]")
        if i < len(lista) - 1 and "fonte" in r:
            time.sleep(args.intervalo)

    if args.saida:
        with open(args.saida, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=CAMPOS_CSV, extrasaction="ignore")
            w.writeheader()
            w.writerows(resultados)
        print(f"\nResultado salvo em {args.saida}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
