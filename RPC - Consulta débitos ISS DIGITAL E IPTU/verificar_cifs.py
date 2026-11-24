#!/usr/bin/env python3
"""Verifica se os CIFs dos débitos são reais (existem com endereço).

Etapa complementar a ``verificar_pagos.py``: para cada CPF/CNPJ com débitos,
consulta o **Cadastro Imobiliário** do IPTU Digital
(``#/cadastroImobiliario``), onde o CPF revela os imóveis com CIF e endereço.
O CIF só é considerado real quando aparece no cadastro **com endereço**.

Saídas (em ``relatorios/``):
* ``<planilha>_VERIFICACAO_CIFS.xlsx`` — cópia do resultado com as colunas
  ``CIF REAL?`` e ``ENDEREÇO (CIF)``;
* ``<planilha>_DETALHE_CIFS.xlsx`` — um CIF por linha, com o endereço e a
  situação do cadastro informados pelo site.
"""

from __future__ import annotations

import re
import urllib.request
from pathlib import Path

from PyQt6.QtCore import QThread, pyqtSignal
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import brave_cookie
from config import (
    BASE_URL, HOME_URL, IPTU_URL, CDP_URL, COOKIE_NAME, EXTRACTION_DELAY,
)
from disclaimer import adicionar_aviso
from preencher_observacao import digitos
from verificar_pagos import slug, chave_ordem_ano, _cabecalho_novo

COLUNA_REAL = "CIF REAL?"
COLUNA_ENDERECO = "ENDEREÇO (CIF)"

COLUNAS_DETALHE = [
    "CPF/CNPJ", "Nome", "CIF", "CIF REAL?", "Endereço",
    "Inscrição Imobiliária", "Situação do Cadastro", "Situação do Imóvel",
    "Tipo de Débito", "Competência", "Observação",
]

CORES_RESULTADO = {
    "SIM": ("FFDFF0DF", "FF1B5E20"),
    "NÃO": ("FFE1E1E1", "FF555555"),
    "NÃO CONFIRMADO": ("FFFFF2CC", "FF7A6000"),
    "NÃO VERIFICADO": ("FFE1E1E1", "FF555555"),
}


class ConsultaCifNaoVerificada(Exception):
    """Não foi possível confirmar os imóveis daquele documento."""


class ConsultaCifInterrompida(Exception):
    """Usuário pediu para parar no meio da consulta."""


# ---------------------------------------------------------------------------
# Leitura dos CIFs já obtidos nas consultas de débito
# ---------------------------------------------------------------------------

def carregar_cifs(resultados: list[dict] | None = None,
                  caminho_resultado: Path | None = None) -> dict[str, list[dict]]:
    """Agrupa por documento os CIFs presentes nos débitos.

    Usa os resultados em memória quando disponíveis; caso contrário lê o
    arquivo ``*_resultado.xlsx`` gravado pelo programa.
    """
    linhas = list(resultados or [])
    if not linhas and caminho_resultado and Path(caminho_resultado).exists():
        wb = load_workbook(caminho_resultado, read_only=True, data_only=True)
        try:
            ws = wb["Resultado"] if "Resultado" in wb.sheetnames else wb.active
            linhas_iter = ws.iter_rows(values_only=True)
            cabecalho = [slug(c) for c in next(linhas_iter, ())]

            def coluna(*termos: str) -> int | None:
                for i, nome in enumerate(cabecalho):
                    if all(termo in nome for termo in termos):
                        return i
                return None

            col_cpf = coluna("cpf")
            col_cif = coluna("cif")
            col_nome = coluna("nome")
            col_tipo = coluna("tipo", "debito")
            col_comp = coluna("compet")
            if col_cpf is None or col_cif is None:
                return {}
            for linha in linhas_iter:
                if col_cif >= len(linha):
                    continue
                linhas.append({
                    "CPF": linha[col_cpf] if col_cpf < len(linha) else "",
                    "CIF": linha[col_cif],
                    "Nome": linha[col_nome] if col_nome is not None and col_nome < len(linha) else "",
                    "Tipo de Débito": linha[col_tipo] if col_tipo is not None and col_tipo < len(linha) else "",
                    "Competência": linha[col_comp] if col_comp is not None and col_comp < len(linha) else "",
                })
        finally:
            wb.close()

    por_documento: dict[str, dict[str, dict]] = {}
    for item in linhas:
        documento = digitos(item.get("CPF", ""))
        if len(documento) not in (11, 14):
            continue
        cif = re.sub(r"\D", "", str(item.get("CIF", "") or ""))
        if not cif or cif == "0":
            continue
        grupo = por_documento.setdefault(documento, {})
        if cif not in grupo:
            grupo[cif] = {
                "Documento": documento,
                "Nome": str(item.get("Nome", "") or "").strip(),
                "CIF": cif,
                "Tipo de Débito": str(item.get("Tipo de Débito", "") or "").strip(),
                "Competência": str(item.get("Competência", "") or "").strip(),
            }
    return {
        doc: sorted(cifs.values(), key=chave_ordem_ano)
        for doc, cifs in por_documento.items()
    }


# ---------------------------------------------------------------------------
# Consulta ao Cadastro Imobiliário (página do IPTU Digital)
# ---------------------------------------------------------------------------

def _verificar_parada(deve_parar) -> None:
    if deve_parar is not None and deve_parar():
        raise ConsultaCifInterrompida()


def _esperar(pagina, milissegundos: int, deve_parar=None) -> None:
    restante = max(0, int(milissegundos))
    while restante > 0:
        _verificar_parada(deve_parar)
        passo = min(250, restante)
        pagina.wait_for_timeout(passo)
        restante -= passo
    _verificar_parada(deve_parar)


def consultar_imoveis(documento: str, cookie: str = "", cookie_sessao: str = "",
                      timeout: int = 30, deve_parar=None) -> list[dict]:
    """Consulta um CPF/CNPJ no Cadastro Imobiliário e devolve os imóveis.

    Cada imóvel traz CIF, endereço, inscrição imobiliária e situações do
    cadastro. Levanta ``ConsultaCifNaoVerificada`` se a pesquisa não puder ser
    confirmada — nesse caso os CIFs ficam NÃO VERIFICADO, nunca "reais".
    """
    _verificar_parada(deve_parar)
    # O Playwright Sync não pode rodar com um event loop asyncio ativo.
    try:
        import asyncio
        try:
            asyncio.get_running_loop()
            asyncio.set_event_loop(None)
        except RuntimeError:
            pass
    except Exception:  # noqa: BLE001
        pass

    from playwright.sync_api import sync_playwright

    usar_cdp = False
    try:
        with urllib.request.urlopen(f"{CDP_URL}/json/version", timeout=2) as resp:
            usar_cdp = resp.status == 200
    except Exception:  # noqa: BLE001
        usar_cdp = False

    with sync_playwright() as playwright:
        pagina_cdp = None
        if usar_cdp:
            navegador = playwright.chromium.connect_over_cdp(CDP_URL)
            pagina = navegador.contexts[0].new_page()
            pagina_cdp = pagina
        else:
            navegador = playwright.chromium.launch(headless=True)
            contexto = navegador.new_context()
            cookies = []
            if cookie.strip():
                cookies.append({
                    "name": COOKIE_NAME, "value": cookie.strip(),
                    "domain": brave_cookie.DOMINIO_JWT, "path": "/",
                    "secure": True,
                })
            if cookie_sessao.strip():
                cookies.append({
                    "name": "JSESSIONID", "value": cookie_sessao.strip(),
                    "domain": brave_cookie.DOMINIO_SESSAO, "path": "/nfsd",
                    "secure": True,
                })
            if cookies:
                contexto.add_cookies(cookies)
            pagina = contexto.new_page()

        try:
            pagina.goto(
                HOME_URL, wait_until="domcontentloaded", timeout=timeout * 1000
            )
            pagina.goto(
                IPTU_URL, wait_until="domcontentloaded", timeout=timeout * 1000
            )
            _esperar(pagina, 900, deve_parar)
            campo = pagina.locator("#documentoFiscal")
            try:
                campo.wait_for(state="visible", timeout=10000)
            except Exception as exc:  # noqa: BLE001
                raise ConsultaCifNaoVerificada(
                    "Campo CPF/CNPJ do Cadastro Imobiliário não apareceu."
                ) from exc

            campo.fill(digitos(documento))
            _esperar(pagina, 250, deve_parar)
            botao = pagina.get_by_role("button", name="Pesquisar")
            if botao.count() == 0:
                raise ConsultaCifNaoVerificada(
                    "Botão Pesquisar do Cadastro Imobiliário não encontrado."
                )
            botao.first.click()
            _esperar(pagina, 1200, deve_parar)

            # Aguarda o paginador mostrar a contagem de registros.
            for _ in range(30):
                _verificar_parada(deve_parar)
                texto = pagina.evaluate("""() => {
                    const el = document.querySelector('.p-paginator-current');
                    return el ? el.innerText : '';
                }""")
                if "registro" in (texto or "").lower():
                    break
                _esperar(pagina, 300, deve_parar)

            imoveis: list[dict] = []
            vistos: set[str] = set()
            paginas = 0
            while paginas < 30:
                _verificar_parada(deve_parar)
                linhas = pagina.evaluate("""() => {
                    const t = document.querySelector('table');
                    if (!t) return [];
                    const cabecalhos = Array.from(t.querySelectorAll('thead th'))
                        .map(h => (h.innerText || '').trim());
                    const indice = termo => cabecalhos.findIndex(h => h.includes(termo));
                    const col = {
                        cif: indice('CIF'), inscricao: indice('Inscri'),
                        responsavel: indice('Respons'), endereco: indice('Endere'),
                        englobado: indice('Englobado'),
                        tributaria: indice('Situação Tributária'),
                        natureza: indice('Natureza'),
                        cadastro: indice('Situação do Cadastro'),
                        imovel: indice('Situação do Imóvel'),
                    };
                    const valor = (c, i) => i >= 0 && i < c.length ? (c[i] || '').trim() : '';
                    return Array.from(t.querySelectorAll('tbody tr')).map(tr => {
                        const celulas = Array.from(tr.querySelectorAll('td'))
                            .map(td => (td.innerText || '').trim());
                        return {
                            cif: valor(celulas, col.cif),
                            inscricao: valor(celulas, col.inscricao),
                            responsavel: valor(celulas, col.responsavel),
                            endereco: valor(celulas, col.endereco),
                            englobado: valor(celulas, col.englobado),
                            situacao_tributaria: valor(celulas, col.tributaria),
                            natureza: valor(celulas, col.natureza),
                            situacao_cadastro: valor(celulas, col.cadastro),
                            situacao_imovel: valor(celulas, col.imovel),
                        };
                    });
                }""")
                for linha in linhas:
                    cif = digitos(linha.get("cif", ""))
                    if cif and cif not in vistos:
                        vistos.add(cif)
                        linha["cif"] = cif
                        imoveis.append(linha)
                proximo = pagina.locator("button.p-paginator-next")
                if proximo.count() == 0 or "disabled" in (
                    proximo.first.get_attribute("class") or ""
                ):
                    break
                anterior = pagina.evaluate("""() => {
                    const el = document.querySelector('.p-paginator-current');
                    return el ? el.innerText : '';
                }""")
                proximo.first.click()
                for _ in range(20):
                    _verificar_parada(deve_parar)
                    atual = pagina.evaluate("""() => {
                        const el = document.querySelector('.p-paginator-current');
                        return el ? el.innerText : '';
                    }""")
                    if atual != anterior:
                        break
                    _esperar(pagina, 250, deve_parar)
                paginas += 1
            return imoveis
        except ConsultaCifInterrompida:
            raise
        except Exception as exc:  # noqa: BLE001
            raise ConsultaCifNaoVerificada(str(exc)) from exc
        finally:
            if usar_cdp:
                try:
                    if pagina_cdp is not None:
                        pagina_cdp.close()
                except Exception:  # noqa: BLE001
                    pass
            else:
                navegador.close()


# ---------------------------------------------------------------------------
# Classificação e gravação
# ---------------------------------------------------------------------------

def endereco_tem_identificador(endereco: object) -> bool:
    """CIF real exige endereço com NÚMERO ou LOTE preenchidos.

    O cadastro devolve textos como:
      'RUA 02, BAIRRO CENTRO, N° 26, QUADRA: 29-B, LOTE: 26' -> real;
      'RUA 27, BAIRRO CARAJAS II, N° , QUADRA: 41, LOTE: 03' -> real (lote);
      'AVENIDA X, BAIRRO Y, N° , QUADRA: , LOTE:' -> fictícia;
      '' -> fictícia.
    """
    texto = str(endereco or "").strip()
    if not texto:
        return False
    # Número: "N° 26", "Nº 26", "N 26", "NUM 26", "n° 26-B" etc.
    if re.search(
        r"(?:n[°ºo]?\.?\s*|num(?:ero)?\.?\s*)([0-9]+[A-Za-z0-9/\-]*)",
        texto, re.I,
    ):
        return True
    # Lote: "LOTE: 03", "LOTE 03", "LOTE: 26-A"
    if re.search(r"lote\s*:?\s*([0-9]+[A-Za-z0-9/\-]*)", texto, re.I):
        return True
    return False


def classificar_cifs(esperados: list[dict], encontrados: list[dict],
                     mensagem: str = "") -> list[dict]:
    """CIF real só com endereço que tenha número ou lote preenchidos."""
    por_cif = {digitos(item.get("cif", "")): item for item in encontrados}
    resultados: list[dict] = []
    for esperado in esperados:
        cif = str(esperado.get("CIF", ""))
        item = por_cif.get(digitos(cif))
        if mensagem:
            situacao = "NÃO VERIFICADO"
        elif item is None:
            situacao = "NÃO CONFIRMADO"
        elif not endereco_tem_identificador(item.get("endereco", "")):
            situacao = "NÃO"
        else:
            situacao = "SIM"
        documento = str(esperado.get("Documento", ""))
        resultados.append({
            "Documento": documento,
            "CPF/CNPJ": documento,
            "Nome": str(esperado.get("Nome", "")),
            "CIF": cif,
            "CIF REAL?": situacao,
            "Endereço": str(item.get("endereco", "") or "").strip() if item else "",
            "Inscrição Imobiliária": str(item.get("inscricao", "") or "").strip() if item else "",
            "Situação do Cadastro": str(item.get("situacao_cadastro", "") or "").strip() if item else "",
            "Situação do Imóvel": str(item.get("situacao_imovel", "") or "").strip() if item else "",
            "Tipo de Débito": str(esperado.get("Tipo de Débito", "")),
            "Competência": str(esperado.get("Competência", "")),
            "Observação": mensagem or (
                "CIF não localizado no Cadastro Imobiliário."
                if item is None else
                "Endereço ausente ou sem número e sem lote — CIF considerada fictícia."
                if situacao == "NÃO" else ""
            ),
        })
    return resultados


def gravar_resultado_cifs(origem: Path, destino: Path, detalhe: Path,
                          resultados: list[dict]) -> tuple[Path, Path]:
    """Grava CIF REAL?/ENDEREÇO no resultado e o detalhe em planilha separada."""
    origem, destino, detalhe = Path(origem), Path(destino), Path(detalhe)
    base = destino if destino.exists() else origem
    wb = load_workbook(base)

    por_chave: dict[tuple[str, str], dict] = {}
    for item in resultados:
        chave = (digitos(item.get("Documento", "")), digitos(item.get("CIF", "")))
        if chave not in por_chave:
            por_chave[chave] = item

    for ws in wb.worksheets:
        headers = [slug(c.value) for c in ws[1]]
        col_cif = next((i + 1 for i, nome in enumerate(headers) if nome == "cif"), None)
        col_cpf = next((i + 1 for i, nome in enumerate(headers) if "cpf" in nome), None)
        if col_cif is None:
            continue
        col_real = next((i + 1 for i, nome in enumerate(headers) if "cif real" in nome), None)
        col_end = next((i + 1 for i, nome in enumerate(headers) if "endere" in nome), None)
        if col_real is None and col_end is None:
            col_real, col_end = col_cif + 1, col_cif + 2
            ws.insert_cols(col_real, amount=2)
            _cabecalho_novo(ws, 1, col_real, COLUNA_REAL)
            _cabecalho_novo(ws, 1, col_end, COLUNA_ENDERECO)
        elif col_real is None:
            col_real = col_end + 1
            ws.insert_cols(col_real)
            _cabecalho_novo(ws, 1, col_real, COLUNA_REAL)
        elif col_end is None:
            col_end = col_real + 1
            ws.insert_cols(col_end)
            _cabecalho_novo(ws, 1, col_end, COLUNA_ENDERECO)
        ws.column_dimensions[get_column_letter(col_real)].width = 18
        ws.column_dimensions[get_column_letter(col_end)].width = 70

        for numero in range(2, ws.max_row + 1):
            cif = digitos(ws.cell(row=numero, column=col_cif).value)
            doc = (
                digitos(ws.cell(row=numero, column=col_cpf).value)
                if col_cpf else ""
            )
            # Somente o par (documento, CIF) exato. Reaproveitar o resultado
            # de outro CPF com o mesmo CIF daria "CIF REAL?" de outra pessoa.
            item = por_chave.get((doc, cif))
            celula = ws.cell(row=numero, column=col_real)
            celula.value = item.get(COLUNA_REAL, "") if item else ""
            celula.alignment = Alignment(horizontal="center", vertical="center")
            fundo, texto = CORES_RESULTADO.get(
                str(celula.value), ("FFFFFFFF", "FF000000")
            )
            celula.fill = PatternFill("solid", start_color=fundo, end_color=fundo)
            celula.font = Font(color=texto, bold=True, size=10)
            ws.cell(row=numero, column=col_end).value = (
                item.get("Endereço", "") if item else ""
            )

    destino.parent.mkdir(parents=True, exist_ok=True)
    adicionar_aviso(wb)
    wb.save(destino)

    detalhe_wb = Workbook()
    detalhe_ws = detalhe_wb.active
    detalhe_ws.title = "Detalhe CIFs"
    detalhe_ws.sheet_view.showGridLines = False
    detalhe_ws.append(COLUNAS_DETALHE)
    for celula in detalhe_ws[1]:
        celula.fill = PatternFill("solid", start_color="FF1F4E79", end_color="FF1F4E79")
        celula.font = Font(color="FFFFFFFF", bold=True, size=11)
        celula.alignment = Alignment(horizontal="center", vertical="center")
    detalhe_ws.row_dimensions[1].height = 30

    colunas = {nome: i + 1 for i, nome in enumerate(COLUNAS_DETALHE)}
    for item in resultados:
        detalhe_ws.append([item.get(coluna, "") for coluna in COLUNAS_DETALHE])
        numero = detalhe_ws.max_row
        fundo, texto = CORES_RESULTADO.get(
            str(item.get(COLUNA_REAL, "")), ("FFFFFFFF", "FF000000")
        )
        celula = detalhe_ws.cell(row=numero, column=colunas[COLUNA_REAL])
        celula.fill = PatternFill("solid", start_color=fundo, end_color=fundo)
        celula.font = Font(color=texto, bold=True, size=10)

    ultima_coluna = get_column_letter(len(COLUNAS_DETALHE))
    detalhe_ws.auto_filter.ref = f"A1:{ultima_coluna}{max(2, detalhe_ws.max_row)}"
    detalhe_ws.freeze_panes = "A2"
    for indice, largura in enumerate(
        (20, 32, 12, 18, 70, 22, 22, 18, 30, 14, 45), start=1
    ):
        detalhe_ws.column_dimensions[get_column_letter(indice)].width = largura
    detalhe.parent.mkdir(parents=True, exist_ok=True)
    adicionar_aviso(detalhe_wb)
    detalhe_wb.save(detalhe)
    return destino, detalhe


# ---------------------------------------------------------------------------
# Worker
# ---------------------------------------------------------------------------

class VerificarCifsWorker(QThread):
    progresso = pyqtSignal(int, str)
    aviso = pyqtSignal(str)
    finalizado = pyqtSignal(str)

    def __init__(self, documentos: dict[str, list[dict]], consultar_func,
                 cookie: str = "", cookie_sessao: str = "",
                 destino: Path | None = None, detalhe: Path | None = None,
                 origem: Path | None = None, pausa: float = EXTRACTION_DELAY,
                 variar_pausa: bool = False, controle: dict | None = None,
                 parent=None):
        super().__init__(parent)
        self.documentos = documentos
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
        itens = list(self.documentos.items())
        total = len(itens)
        erros = 0
        for indice, (documento, cifs) in enumerate(itens, start=1):
            if self._parar:
                break
            rotulo = cifs[0].get("Nome") or documento if cifs else documento
            self.progresso.emit(
                int((indice - 1) * 100 / max(1, total)),
                f"Verificando CIFs {indice}/{total}: {rotulo}",
            )
            mensagem = ""
            try:
                imoveis = self.consultar_func(
                    documento, self.cookie, self.cookie_sessao, 30,
                    deve_parar=self.parar_solicitado,
                )
            except ConsultaCifInterrompida:
                break
            except Exception as exc:  # noqa: BLE001
                imoveis = []
                mensagem = str(exc)
                erros += 1
                self.aviso.emit(
                    f"{documento}: não foi possível consultar o cadastro "
                    f"imobiliário ({exc})."
                )
            self.resultados.extend(
                classificar_cifs(cifs, imoveis, mensagem)
            )
            pausa = float(self.controle.get("pausa", self.pausa))
            if self.variar_pausa or self.controle.get("variar", self.variar_pausa):
                import random
                pausa *= random.uniform(0.7, 1.3)
            if indice < total and not self._parar and pausa > 0:
                self.msleep(int(pausa * 1000))

        interrompido = self._parar
        if self.resultados and self.destino and self.detalhe:
            try:
                gravar_resultado_cifs(
                    self.origem or self.destino, self.destino, self.detalhe,
                    self.resultados,
                )
            except Exception as exc:  # noqa: BLE001
                self.finalizado.emit(
                    f"Erro ao gravar relatórios de CIFs: {exc}"
                )
                return
        if interrompido:
            mensagem_final = (
                f"Verificação de CIFs interrompida: {len(self.resultados)} "
                "CIF(s) processado(s). Resultados parciais salvos."
            )
        else:
            reais = sum(
                1 for item in self.resultados
                if item.get(COLUNA_REAL) == "SIM"
            )
            mensagem_final = (
                f"Concluído: {len(self.documentos)} documento(s), "
                f"{len(self.resultados)} CIF(s) verificado(s), "
                f"{reais} com endereço (reais)"
            )
            if erros:
                mensagem_final += f", {erros} consulta(s) sem confirmação"
            mensagem_final += "."
        self.progresso.emit(100, mensagem_final)
        self.finalizado.emit(mensagem_final)


def caminhos_saida_cifs(caminho_entrada: Path | None, pasta_relatorios: Path,
                        stem_base: str) -> tuple[Path, Path]:
    """Nomes padronizados dos arquivos da verificação de CIFs."""
    stem = caminho_entrada.stem if caminho_entrada else stem_base
    return (
        pasta_relatorios / f"{stem}_VERIFICACAO_CIFS.xlsx",
        pasta_relatorios / f"{stem}_DETALHE_CIFS.xlsx",
    )
