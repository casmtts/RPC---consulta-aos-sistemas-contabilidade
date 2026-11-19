"""
Sistema de Consulta de Débitos - GUI completa (PyQt6).

Arquivo principal: concentra interface, leitura/gravação de planilhas,
consulta ao site e processamento em segundo plano.
"""

from __future__ import annotations

import csv
import re
import sys
import threading
import unicodedata
from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import Qt, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import QBrush, QColor, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QStatusBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

# Garante que os módulos locais sejam encontrados
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import BASE_URL, HOME_URL, COOKIE_NAME, EXTRACTION_DELAY, CDP_URL
from models import PLANILLA_CABECALOS
from brave_cookie import (
    buscar_cookies_no_brave,
    segundos_para_expirar,
    DOMINIO_JWT,
    DOMINIO_SESSAO,
)

# Relê o token do Brave quando falta menos de 1 minuto para expirar.
MARGEM_RENOVACAO_TOKEN = 60

COLUNAS_SAIDA = list(PLANILLA_CABECALOS) + ["Erro"]
SEPARADOR_CSV = ";"

StatusConsulta = {
    "PENDENTE": "pendente",
    "PROCESSANDO": "processando",
    "COMPLETO": "completo",
    "ERRO": "erro",
}


# ============================================================
# Planilhas
# ============================================================

def _detectar_cabecalho(linhas: list[tuple]) -> int | None:
    """Encontra a linha de cabeçalho (ignora títulos e linhas vazias do topo)."""
    for indice, linha in enumerate(linhas[:20]):
        preenchidas = [str(cel).strip() for cel in linha if cel not in (None, "")]
        if len(preenchidas) < 2:
            continue
        texto = " ".join(preenchidas).lower()
        if any(palavra in texto for palavra in
               ("nome", "cpf", "cnpj", "valor", "origem", "compet", "situa")):
            return indice
    for indice, linha in enumerate(linhas[:20]):
        if sum(cel not in (None, "") for cel in linha) >= 2:
            return indice
    return None


def listar_abas(caminho: str) -> list[str]:
    """Lista as abas de um arquivo Excel."""
    from openpyxl import load_workbook

    wb = load_workbook(caminho, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _selecionar_aba(wb, nome_aba: str | None = None):
    """Escolhe a aba explicitamente ou a que tiver mais dados com cabeçalho."""
    if nome_aba:
        if nome_aba not in wb.sheetnames:
            raise ValueError(f"Aba não encontrada: {nome_aba}")
        return wb[nome_aba]

    melhor = None
    maior = -1
    for ws in wb.worksheets:
        try:
            amostra = list(ws.iter_rows(min_row=1, max_row=8, values_only=True))
        except Exception:  # noqa: BLE001
            continue
        if _detectar_cabecalho(amostra) is None:
            continue
        tamanho = ws.max_row or 0
        if tamanho > maior:
            melhor, maior = ws, tamanho
    return melhor or wb.active


def aba_padrao(caminho: str) -> str | None:
    """Nome da aba que será usada por padrão no arquivo Excel."""
    from openpyxl import load_workbook

    wb = load_workbook(caminho, read_only=True, data_only=True)
    try:
        return _selecionar_aba(wb).title
    finally:
        wb.close()


def ler_planilha(caminho: str, nome_aba: str | None = None) -> tuple[list[str], list[dict]]:
    """Lê .csv ou .xlsx e devolve (cabeçalhos, linhas).

    Em Excel, detecta o cabeçalho mesmo abaixo de títulos e permite escolher a aba.
    """
    caminho = Path(caminho)
    ext = caminho.suffix.lower()

    if ext == ".csv":
        with caminho.open("r", encoding="utf-8-sig", newline="") as arquivo:
            amostra = arquivo.read(4096)
            arquivo.seek(0)
            try:
                dialeto = csv.Sniffer().sniff(amostra, delimiters=",;\t")
            except csv.Error:
                dialeto = csv.excel
            leitor = csv.reader(arquivo, dialect=dialeto)
            cabecalho = [coluna.strip() for coluna in next(leitor, [])]
            linhas = []
            for valores in leitor:
                if not any(valor.strip() for valor in valores):
                    continue
                linhas.append({
                    col: (valores[i].strip() if i < len(valores) else "")
                    for i, col in enumerate(cabecalho)
                    if col
                })

    elif ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook

        wb = load_workbook(caminho, read_only=True, data_only=True)
        try:
            ws = _selecionar_aba(wb, nome_aba)
            brutas = list(ws.iter_rows(values_only=True))
            indice = _detectar_cabecalho(brutas)
            if indice is None:
                raise ValueError(f"Não foi possível encontrar o cabeçalho na aba '{ws.title}'.")
            cabecalho = ["" if cel is None else str(cel).strip() for cel in brutas[indice]]
            linhas = []
            for valores in brutas[indice + 1:]:
                if valores is None or all(cel in (None, "") for cel in valores):
                    continue
                linhas.append({
                    col: ("" if i >= len(valores) or valores[i] is None else str(valores[i]).strip())
                    for i, col in enumerate(cabecalho)
                    if col
                })
        finally:
            wb.close()

    else:
        raise ValueError(
            f"Formato não suportado: {ext or 'sem extensão'}. Use .csv ou .xlsx"
        )

    return [col for col in cabecalho if col], linhas


COR_COM_DEBITOS = "FFFFFFFF"   # branco
COR_SEM_DEBITOS = "FFD9D9D9"   # cinza-claro


def _tem_debitos(linha: dict) -> bool:
    """Identifica se a linha representa uma pessoa com débitos confirmados."""
    if str(linha.get("Status", "")).strip().upper() not in ("OK", "VERIFICADO"):
        return False
    descricao = str(linha.get("Tipo de Débito", "")).strip().lower()
    return bool(descricao) and "sem débitos" not in descricao and descricao != "-"


def _arquivo_progresso(caminho_entrada: Path, aba: str = "") -> Path:
    """Arquivo de retomada ao lado da planilha de entrada."""
    sufixo = f"_{_sem_acento(aba)}" if aba else ""
    return caminho_entrada.with_name(
        f"{caminho_entrada.stem}_progresso{sufixo}.json"
    )


def _pasta_relatorios() -> Path:
    """Pasta 'relatorios' dentro da pasta do programa (criada se preciso)."""
    pasta = Path(__file__).resolve().parent / "relatorios"
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def salvar_progresso(caminho: Path, processados: int, total: int,
                     resultados: list[dict]) -> None:
    """Grava quantos registros já foram processados e os resultados parciais."""
    import json
    dados = {
        "processados": processados,
        "total": total,
        "resultados": resultados,
        "atualizado": datetime.now().isoformat(timespec="seconds"),
    }
    try:
        caminho.write_text(
            json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:  # noqa: BLE001
        pass


def ler_progresso(caminho: Path) -> dict | None:
    """Lê o progresso salvo, se existir e for válido."""
    import json
    if not caminho.exists():
        return None
    try:
        dados = json.loads(caminho.read_text(encoding="utf-8"))
        if not isinstance(dados.get("processados"), int):
            return None
        if not isinstance(dados.get("resultados"), list):
            return None
        return dados
    except Exception:  # noqa: BLE001
        return None


def gravar_planilha(caminho: str, linhas: list[dict]) -> Path:
    """Grava o resultado em .xlsx ou .csv, com branco para com débitos e cinza para sem."""
    caminho = Path(caminho)
    if caminho.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill

        preenchimento_com = PatternFill(
            fill_type="solid", start_color=COR_COM_DEBITOS, end_color=COR_COM_DEBITOS
        )
        preenchimento_sem = PatternFill(
            fill_type="solid", start_color=COR_SEM_DEBITOS, end_color=COR_SEM_DEBITOS
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "Resultado"
        ws.append(COLUNAS_SAIDA)
        for linha in linhas:
            ws.append([linha.get(coluna, "") for coluna in COLUNAS_SAIDA])
            preenchimento = preenchimento_com if _tem_debitos(linha) else preenchimento_sem
            for coluna in range(1, len(COLUNAS_SAIDA) + 1):
                ws.cell(row=ws.max_row, column=coluna).fill = preenchimento
        wb.save(caminho)
    else:
        with caminho.open("w", encoding="utf-8-sig", newline="") as arquivo:
            escritor = csv.DictWriter(
                arquivo,
                fieldnames=COLUNAS_SAIDA,
                delimiter=SEPARADOR_CSV,
                extrasaction="ignore",
            )
            escritor.writeheader()
            escritor.writerows(linhas)
    return caminho


# ============================================================
# Consulta ao site
# ============================================================

def _sem_acento(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto).lower())
        if unicodedata.category(c) != "Mn"
    )


def _campo(registro: dict, nome: str) -> str:
    """Busca o valor de uma coluna pelo nome, ignorando acentos/maiúsculas."""
    alvo = _sem_acento(nome)
    for chave, valor in registro.items():
        if alvo in _sem_acento(chave):
            return str(valor).strip()
    return ""


def _tipo_debito_com_migracao(tipo_site: str, registro: dict) -> str:
    """Acrescenta '(M)' quando a origem da planilha indica débito migrado."""
    tipo = str(tipo_site or "").strip()
    origem = _campo(registro, "origem")
    if tipo and "(m)" in origem.lower() and "(m)" not in tipo.lower():
        return f"{tipo} (M)"
    return tipo


def _valor_por_chaves(item: dict, chaves: list[str]) -> str:
    for chave, valor in item.items():
        chave_norm = _sem_acento(chave)
        if any(_sem_acento(opcao) in chave_norm for opcao in chaves):
            return str(valor).strip()
    return ""


def _para_numero(texto: str) -> float:
    """Converte 'R$ 1.234,56' ou '1234.56' em float."""
    limpo = re.sub(r"[^0-9,.\-]", "", str(texto))
    if "," in limpo and "." in limpo:
        if limpo.rfind(",") > limpo.rfind("."):
            limpo = limpo.replace(".", "").replace(",", ".")
        else:
            limpo = limpo.replace(",", "")
    elif "," in limpo:
        limpo = limpo.replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return 0.0


def _extrair_tabelas(page) -> list[dict]:
    """Extrai débitos das tabelas da página."""
    debitos: list[dict] = []
    for tabela in page.query_selector_all("table"):
        linhas = tabela.query_selector_all("tr")
        if len(linhas) < 2:
            continue
        cabecalhos = [
            (cel.inner_text() or "").strip()
            for cel in linhas[0].query_selector_all("th, td")
        ]
        # Ignora tabelas de apoio (paginação, competências etc.), sem dados de débito.
        cabecalho_texto = _sem_acento(" ".join(cabecalhos))
        if not any(palavra in cabecalho_texto for palavra in
                   ("competencia", "divida", "situacao", "origem")):
            continue
        for tr in linhas[1:]:
            celulas = [
                (cel.inner_text() or "").strip()
                for cel in tr.query_selector_all("td")
            ]
            if not any(celulas):
                continue
            item = {
                cabecalhos[i]: valor
                for i, valor in enumerate(celulas)
                if i < len(cabecalhos)
            }
            debitos.append({
                "tipo_debito": _valor_por_chaves(
                    item, ["origem", "tipo", "debito", "tributo"]
                ),
                "competencia": _valor_por_chaves(item, ["competencia", "ano", "exercicio"]),
                "situacao": _valor_por_chaves(item, ["situacao", "status"]),
                "valor": _para_numero(_valor_por_chaves(
                    item, ["divida (r$)", "valor", "atualizado (r$)", "lancado (r$)"]
                )),
            })
    return debitos


class ConsultaNaoVerificada(Exception):
    """A consulta não pôde ser confirmada. Nunca deve virar 'Sem débitos'."""


class Erro404Site(ConsultaNaoVerificada):
    """O site respondeu 404: sessão/viewstate inválido ou fora do ar."""


def _estado_resultado(pagina) -> tuple[bool, bool, bool]:
    """Depois da pesquisa: (tem_tabela, tem_linhas_de_dados, diz_que_nao_ha_registros)."""
    try:
        texto = (pagina.inner_text("body") or "").lower()
    except Exception:  # noqa: BLE001
        texto = ""
    sem_registros = any(
        trecho in texto
        for trecho in (
            "nenhum registro", "nenhum debito", "nenhuma guia",
            "sem resultados", "nao foram encontrados", "não foram encontrados",
            "nenhum lancamento", "nenhum lançamento",
            "registro nao encontrado", "registro não encontrado",
            "no records found", "nenhum registro encontrado",
        )
    )
    tem_tabela = tem_dados = False
    for tabela in pagina.query_selector_all("table"):
        if not tabela.query_selector("tr th, tr td"):
            continue
        tem_tabela = True
        linhas = tabela.query_selector_all("tr")
        for linha in linhas[1:]:
            if linha.query_selector("td"):
                tem_dados = True
                break
    return tem_tabela, tem_dados, sem_registros


def _tipo_e_campo_documento(documento: str, tipo: str = "") -> tuple[str, str]:
    """Resolve o tipo de pessoa e o campo a preencher no formulário do site.

    No ISS Digital, selecionar FÍSICA/JURÍDICA/IMÓVEL revela, via AJAX, um
    campo diferente: CPF, CNPJ ou CIF, respectivamente.
    """
    digitos = re.sub(r"\D", "", documento or "")
    tipo_norm = _sem_acento(tipo).strip().upper()

    if not tipo_norm:
        if len(digitos) == 14:
            tipo_norm = "JURIDICA"
        elif len(digitos) == 11:
            tipo_norm = "FISICA"
        else:
            raise ConsultaNaoVerificada(
                "Não foi possível inferir FÍSICA/JURÍDICA do documento informado "
                "(esperado 11 dígitos para CPF ou 14 para CNPJ)."
            )

    mapa = {
        "FISICA": ("FISICA", "#documentoContribuinteCpf"),
        "FISICO": ("FISICA", "#documentoContribuinteCpf"),
        "CPF": ("FISICA", "#documentoContribuinteCpf"),
        "JURIDICA": ("JURIDICA", "#documentoContribuinteCnpj"),
        "JURIDICO": ("JURIDICA", "#documentoContribuinteCnpj"),
        "CNPJ": ("JURIDICA", "#documentoContribuinteCnpj"),
        "IMOVEL": ("IMOVEL", "#documentoContribuinteCif"),
    }
    if tipo_norm not in mapa:
        raise ConsultaNaoVerificada(f"Tipo de pessoa desconhecido: {tipo!r}.")
    return mapa[tipo_norm]


def _formatar_documento(documento: str) -> str:
    """Formata CPF/CNPJ para o padrão aceito pelo campo com máscara."""
    digitos = "".join(caracter for caracter in (documento or "") if caracter.isdigit())
    if "." in documento or "/" in documento or "-" in documento:
        return documento.strip()
    if len(digitos) == 11:
        return f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    if len(digitos) == 14:
        return f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:]}"
    return documento.strip()


def _aplicar_tipo_pesquisa(pagina) -> str:
    """Escolhe 'Competência' no select 'Pesquisar por' do site, se existir."""
    try:
        for select in pagina.query_selector_all("select"):
            opcoes = []
            for opt in select.query_selector_all("option"):
                texto = (opt.inner_text() or "").strip()
                valor = opt.get_attribute("value") or ""
                opcoes.append((texto, valor))
            normal = lambda s: _sem_acento(s).strip().lower()
            textos = " ".join(normal(t) for t, _ in opcoes)
            if "competencia" not in textos or "vencimento" not in textos:
                continue
            for texto, valor in opcoes:
                if "competencia" not in normal(texto):
                    continue
                try:
                    select.select_option(label=texto)
                except Exception:  # noqa: BLE001
                    try:
                        select.select_option(value=valor)
                    except Exception:  # noqa: BLE001
                        return ""
                pagina.wait_for_timeout(800)
                return "competencia"
    except Exception:  # noqa: BLE001
        pass
    return ""


def _aplicar_periodo(pagina) -> None:
    """Amplia o filtro de período para os últimos 10 anos.

    O relatório vem com ano inicial/final iguais ao ano corrente; sem ampliar,
    débitos de exercícios anteriores ficam de fora da consulta.
    """
    # O site aceita no máximo 10 anos inclusivos; 9 anos atrás até o ano atual
    # resulta em exatamente 10 anos (ex.: 2017-2026).
    ano_inicial = datetime.now().year - 9
    try:
        # O spinner do PrimeFaces ignora fill(): é preciso digitar com teclado.
        # Definir o ANO primeiro e só depois o MÊS — ordem inversa faz o AJAX
        # do mês sobrescrever o ano e a consulta retorna vazia.
        campo = pagina.locator("#anoInicial_input")
        campo.click()
        campo.press("ControlOrMeta+A")
        campo.press_sequentially(str(ano_inicial), delay=30)
        campo.press("Tab")
        pagina.wait_for_timeout(1000)

        mes_inicial = pagina.query_selector("#mesInicial_input")
        if mes_inicial:
            mes_inicial.select_option("JANEIRO")
            pagina.wait_for_timeout(500)
    except Exception:  # noqa: BLE001
        pass


def _executar_com_playwright(funcao, *args, **kwargs):
    """Roda ``funcao`` numa thread nova, sem event loop asyncio ativo.

    O Playwright Sync (greenlets) marca um loop asyncio como ativo na thread
    que o usa. Quando a consulta acontece dentro de um loop — ou depois de
    outra sessão na mesma thread — ``sync_playwright()`` falha com "Sync API
    inside the asyncio loop". Uma thread limpa por sessão evita o problema.
    """
    resultado: dict[str, object] = {}

    def alvo() -> None:
        try:
            resultado["valor"] = funcao(*args, **kwargs)
        except BaseException as exc:  # reerguido na thread chamadora
            resultado["erro"] = exc

    thread = threading.Thread(target=alvo, daemon=True)
    thread.start()
    thread.join()
    if "erro" in resultado:
        raise resultado["erro"]
    return resultado.get("valor")


def _limpar_loop_asyncio() -> None:
    """Remove um loop asyncio pendente deixado pelo Playwright na thread.

    O Playwright Sync marca um loop como ativo na thread; se outra sessão
    começar nela, o ``sync_playwright()`` reclama do "Async API". Esta
    função zera essa marcação antes de abrir a sessão.
    """
    try:
        import asyncio

        asyncio._set_running_loop(None)
    except Exception:  # noqa: BLE001
        pass


def _renovar_token(cookie_atual: str, sessao_atual: str) -> tuple[str, str]:
    """Relê os cookies do Brave (onde o navegador já renovou o JWT)."""
    try:
        novos = buscar_cookies_no_brave()
    except Exception:  # noqa: BLE001
        return cookie_atual, sessao_atual
    return (
        novos.get(COOKIE_NAME, cookie_atual),
        novos.get("JSESSIONID", sessao_atual),
    )


def _aviso_expiracao(cookie: str) -> str | None:
    """Aviso quando o JWT está vencido ou prestes a vencer."""
    restante = segundos_para_expirar(cookie)
    if restante is None:
        return None
    if restante <= 0:
        return (
            "Token de autenticação do Brave VENCIDO — entre no sistema "
            "novamente no navegador."
        )
    if restante <= MARGEM_RENOVACAO_TOKEN:
        return (
            f"Token do Brave expira em ~{max(1, restante // 60)} min — "
            "mantenha a sessão ativa no navegador."
        )
    return None


def consultar_debito(
    cpf: str,
    nome: str = "",
    cookie_value: str = "",
    timeout: int = 30,
    tipo: str = "",
    cookie_sessao: str = "",
    _tentativa: int = 0,
) -> list[dict]:
    """Consulta um documento (CPF/CNPJ/CIF) no site.

    O formulário exige selecionar primeiro o tipo de pessoa: FÍSICA revela o
    campo de CPF, JURÍDICA o de CNPJ e IMÓVEL o de CIF. Quando ``tipo`` não é
    informado, ele é inferido pela quantidade de dígitos do documento.

    Só devolve lista vazia quando a pesquisa foi executada e confirmadamente
    não há débitos. Qualquer dúvida levanta ConsultaNaoVerificada.
    """
    import urllib.request

    from playwright.sync_api import sync_playwright

    _limpar_loop_asyncio()
    tipo_pesquisa = ""

    # Preferir a sessão real do Brave (via CDP): é onde o "Pesquisar" funciona.
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
            contexto = navegador.contexts[0]
            pagina = contexto.new_page()
            pagina_cdp = pagina
        else:
            navegador = playwright.chromium.launch(headless=True)
            contexto = navegador.new_context()
            pagina = None
        try:
            if not usar_cdp:
                cookies = []
                if cookie_value.strip():
                    cookies.append({
                        "name": COOKIE_NAME,
                        "value": cookie_value.strip(),
                        "domain": DOMINIO_JWT,
                        "path": "/",
                        "secure": True,
                    })
                if cookie_sessao.strip():
                    cookies.append({
                        "name": "JSESSIONID",
                        "value": cookie_sessao.strip(),
                        "domain": DOMINIO_SESSAO,
                        "path": "/nfsd",
                        "secure": True,
                    })
                if cookies:
                    contexto.add_cookies(cookies)
                pagina = contexto.new_page()

            # O site (JSF/PrimeFaces) redireciona para a home na primeira
            # requisição de uma sessão nova, antes de criar o JSESSIONID.
            # Acessar a home primeiro estabelece a sessão autenticada; sem
            # isso, o goto abaixo cairia em home.jsf e geraria falso erro.
            pagina.goto(HOME_URL, wait_until="domcontentloaded", timeout=timeout * 1000)

            resposta = pagina.goto(BASE_URL, wait_until="domcontentloaded", timeout=timeout * 1000)
            if resposta is None or resposta.status >= 400:
                status = resposta.status if resposta else "sem resposta"
                raise Erro404Site(f"Site retornou status {status}.")

            if "home.jsf" in pagina.url.lower() and "relatorio" not in pagina.url.lower():
                raise ConsultaNaoVerificada(
                    "Site redirecionou para a home (sem contribuinte selecionado na sessão)."
                )

            pagina.wait_for_timeout(800)

            # "Informações do Débito" vem recolhido; abrir antes de mexer no
            # formulário evita que o AJAX do radio esconda o filtro de Situação.
            try:
                if pagina.query_selector("#j_idt379 .ui-fieldset-toggler"):
                    pagina.query_selector("#j_idt379 .ui-fieldset-toggler").click()
                    pagina.wait_for_timeout(250)
            except Exception:  # noqa: BLE001
                pass

            tipo_pessoa, seletor_documento = _tipo_e_campo_documento(cpf, tipo)

            radio = pagina.query_selector(
                f"input[name='tipoPessoaContribuinte'][value='{tipo_pessoa}']"
            )
            if not radio:
                raise ConsultaNaoVerificada(
                    f"Radio de tipo de pessoa {tipo_pessoa!r} não encontrado — "
                    "pesquisa não executada."
                )
            radio.check(force=True)
            pagina.wait_for_timeout(400)

            campo = pagina.locator(seletor_documento)
            campo.click()
            campo.press_sequentially(cpf, delay=20)
            campo.press("Tab")
            pagina.wait_for_timeout(900)

            # Se houver "Pesquisar por: competência/lançamento/vencimento",
            # escolhe competência antes de ajustar o período.
            tipo_pesquisa = _aplicar_tipo_pesquisa(pagina)

            # Consultar sempre os últimos 10 anos de competência (após o CPF,
            # para o AJAX do período não re-renderizar o campo com máscara).
            _aplicar_periodo(pagina)

            # Em "Informações do Débito", deixar marcada SOMENTE a Situação
            # "DÉBITO". A sessão pode trazer filtros de consultas anteriores,
            # então todos os checkboxes são normalizados antes de pesquisar.
            normalizar = """() => {
                document.querySelectorAll('input[name="situacaoDebito"]')
                    .forEach(campo => {
                        const deveEstarMarcado = campo.value === 'DEBITO';
                        if (campo.checked !== deveEstarMarcado) {
                            campo.checked = deveEstarMarcado;
                            campo.dispatchEvent(
                                new Event('change', { bubbles: true })
                            );
                        }
                    });
            }"""
            try:
                pagina.evaluate(normalizar)
                pagina.wait_for_timeout(300)
                marcados = pagina.eval_on_selector_all(
                    'input[name="situacaoDebito"]:checked',
                    "els => els.map(e => e.value)",
                )
                if marcados != ["DEBITO"]:
                    raise ConsultaNaoVerificada(
                        f"Não foi possível isolar a Situação = DÉBITO: {marcados}"
                    )
            except ConsultaNaoVerificada:
                raise
            except Exception as exc:  # noqa: BLE001
                raise ConsultaNaoVerificada(
                    f"Não foi possível configurar o filtro de Situação: {exc}"
                ) from exc

            botao = pagina.query_selector("#botaoAcaoPesquisar")
            if not botao:
                raise ConsultaNaoVerificada(
                    "Botão 'Pesquisar' não encontrado — pesquisa não executada."
                )
            botao.click()
            pagina.wait_for_timeout(5000)
            tem_tabela, tem_dados, sem_registros = _estado_resultado(pagina)

            if sem_registros or (tem_tabela and not tem_dados):
                debitos: list[dict] = []
            elif not tem_tabela:
                raise ConsultaNaoVerificada(
                    "Após a pesquisa não apareceu tabela de resultados nem mensagem de "
                    "'nenhum registro' — não é seguro afirmar que não há débitos."
                )
            else:
                debitos = _extrair_tabelas(pagina)
                for debito in debitos:
                    debito["_pesquisa_por"] = tipo_pesquisa
                if not debitos:
                    raise ConsultaNaoVerificada(
                        "A tabela de resultados tem linhas, mas não foi possível ler os débitos."
                    )
        finally:
            if usar_cdp:
                # Fecha apenas a aba criada, nunca o navegador do usuário.
                try:
                    if pagina_cdp is not None:
                        pagina_cdp.close()
                except Exception:  # noqa: BLE001
                    pass
            else:
                navegador.close()

    # Resultado vazio pode ser condição de corrida do JSF. Refazer a consulta
    # numa sessão nova antes de aceitar que o CPF não tem débitos.
    if not debitos and _tentativa == 0:
        return _executar_com_playwright(
            consultar_debito, cpf, nome, cookie_value, timeout, tipo,
            cookie_sessao, _tentativa=1
        )
    return debitos


# ============================================================
# Worker (processamento em segundo plano)
# ============================================================

class ConsultaWorker(QThread):
    progresso = pyqtSignal(int, str)
    aviso = pyqtSignal(str)
    item_concluido = pyqtSignal(dict)
    documento_revisado = pyqtSignal(str, list)
    finalizado = pyqtSignal(str)

    def __init__(self, registros: list[dict], cookie: str = "", timeout: int = 30,
                 cookie_sessao: str = "", indice_inicial: int = 0,
                 arquivo_progresso: Path | None = None,
                 resultados_iniciais: list[dict] | None = None,
                 modo_revisao: bool = False,
                 documentos_ignorar: set[str] | None = None, parent=None):
        super().__init__(parent)
        self.registros = registros
        self.cookie = cookie
        self.cookie_sessao = cookie_sessao
        self.timeout = timeout
        self.indice_inicial = indice_inicial
        self.arquivo_progresso = arquivo_progresso
        self.resultados = list(resultados_iniciais or [])
        self.modo_revisao = modo_revisao
        # Na execução paralela, estes documentos ficam com a revisão.
        self.documentos_ignorar = documentos_ignorar or set()
        self._parar = False

    def stop(self) -> None:
        self._parar = True

    def _publicar(self, documento: str, novos: list[dict]) -> None:
        """Publica os itens de um documento (normal ou revisão)."""
        if self.modo_revisao:
            # Na revisão, substituir as linhas antigas do mesmo CPF.
            self.documento_revisado.emit(documento, novos)
        else:
            for item in novos:
                self.resultados.append(item)
                self.item_concluido.emit(item)

    def run(self) -> None:
        total = len(self.registros)
        ok = erros = processados = 0

        for indice, registro in enumerate(self.registros, start=1):
            if self._parar:
                break
            # Retomada: pular os registros já processados numa execução anterior.
            if indice <= self.indice_inicial:
                continue

            # Se o token estiver a menos de 1 minuto de vencer, relê o
            # cookie do Brave antes da próxima consulta — sem parar nada.
            restante = segundos_para_expirar(self.cookie)
            if restante is None or restante <= MARGEM_RENOVACAO_TOKEN:
                self.cookie, self.cookie_sessao = _renovar_token(
                    self.cookie, self.cookie_sessao
                )

            documento = _campo(registro, "cpf") or _campo(registro, "cnpj")
            # Na execução paralela, estes documentos ficam com a revisão.
            if documento.strip() in self.documentos_ignorar:
                continue
            nome = _campo(registro, "nome")
            tipo = _campo(registro, "tipo de pessoa") or _campo(registro, "tipo pessoa")
            if not tipo and _campo(registro, "cnpj"):
                tipo = "JURIDICA"
            rotulo = nome or documento or f"linha {indice}"
            # Competência da época: prioriza a planilha de entrada.
            competencia_original = (
                _campo(registro, "competencia debito")
                or _campo(registro, "competencia guia")
                or _campo(registro, "competencia")
            )
            self.progresso.emit(
                int((indice - 1) * 100 / total), f"Consultando {indice}/{total}: {rotulo}"
            )

            try:
                # Se o token expirar e a consulta der erro, pega um token
                # novo do Brave e refaz a consulta (até 3 tentativas).
                for tentativa in range(3):
                    if tentativa:
                        self.cookie, self.cookie_sessao = _renovar_token(
                            self.cookie, self.cookie_sessao
                        )
                    try:
                        debitos = _executar_com_playwright(
                            consultar_debito,
                            documento, nome, self.cookie, self.timeout, tipo,
                            self.cookie_sessao,
                        )
                        break
                    except ConsultaNaoVerificada as exc:
                        mensagem = str(exc).lower()
                        # Falhas de sessão, navigation abortada e formulário
                        # ausente disparam uma nova tentativa em sessão nova.
                        repetir = any(indicio in mensagem for indicio in (
                            "404", "página de erro", "erro (código",
                            "sessão expirada", "sessao expirada",
                            "token", "autenticação", "autenticacao",
                            "pesquisar", "pesquisa não executada",
                            "não foi possível acessar", "err_aborted",
                            "net::", "não apareceu tabela",
                            "radio de tipo", "selecionar o tipo",
                        ))
                        if not repetir or tentativa == 2:
                            raise
                        self.aviso.emit(
                            f"{documento}: {exc} — abrindo sessão nova no "
                            f"Brave e refazendo a consulta ({tentativa + 2}/3)."
                        )
                        self.progresso.emit(
                            max(0, int((indice - 1) * 100 / total)),
                            f"Refazendo {documento} com sessão nova "
                            f"({tentativa + 2}/3)",
                        )
                        self.msleep(1000)
                if not debitos:
                    novos = [{
                        "Nome": nome, "CPF": documento,
                        "Tipo de Débito": "Sem débitos (confirmado)", "Competência": competencia_original or "-",
                        "Valor (R$)": "0,00", "Status": "VERIFICADO", "Erro": "",
                    }]
                else:
                    novos = [
                        {
                            "Nome": nome, "CPF": documento,
                            "Tipo de Débito": _tipo_debito_com_migracao(
                                debito["tipo_debito"], registro
                            ),
                            "Competência": (
                                debito["competencia"]
                                if debito.get("_pesquisa_por") == "competencia"
                                else (competencia_original or debito["competencia"])
                            ),
                            "_competencia_site": debito["competencia"],
                            "Valor (R$)": f"{debito['valor']:.2f}".replace(".", ","),
                            "Status": "VERIFICADO", "Erro": "",
                        }
                        for debito in debitos
                    ]
                self._publicar(documento, novos)
                ok += 1
            except ConsultaNaoVerificada as exc:
                erros += 1
                self._publicar(documento, [{
                    "Nome": nome, "CPF": documento,
                    "Tipo de Débito": "-", "Competência": competencia_original or "-",
                    "Valor (R$)": "0,00", "Status": "NAO VERIFICADO", "Erro": str(exc),
                }])
            except Exception as exc:  # noqa: BLE001
                erros += 1
                self._publicar(documento, [{
                    "Nome": nome, "CPF": documento,
                    "Tipo de Débito": "-", "Competência": competencia_original or "-",
                    "Valor (R$)": "0,00", "Status": "ERRO", "Erro": str(exc),
                }])

            processados = indice
            if self.arquivo_progresso is not None:
                salvar_progresso(self.arquivo_progresso, indice, total, self.resultados)
            self.progresso.emit(int(indice * 100 / total), f"Processado {indice}/{total}")

        if self._parar:
            self.finalizado.emit(
                f"Processamento interrompido: {ok} OK, {erros} com erro "
                f"({processados}/{total} linhas processadas)."
            )
        else:
            self.finalizado.emit(f"Concluído: {ok} OK, {erros} com erro.")


# ============================================================
# Interface gráfica
# ============================================================

def _configurar_tabela(tabela: QTableWidget, colunas: list[str]) -> None:
    tabela.setColumnCount(len(colunas))
    tabela.setHorizontalHeaderLabels(colunas)
    tabela.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    tabela.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    tabela.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
    tabela.verticalHeader().setVisible(False)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Consulta de Débitos")
        self.resize(1100, 720)
        self.setMinimumSize(900, 600)

        self.caminho_entrada: Path | None = None
        self.cabecalho: list[str] = []
        self.registros: list[dict] = []
        self.resultados: list[dict] = []
        # Revisão após uma parada: consultar somente o que ficou pendente.
        self._consulta_interrompida = False
        self._registros_revisao: list[dict] = []
        self._revisao_concluida: set[str] = set()
        # Revisão em paralelo: worker e resultados isolados até a mesclagem.
        self.worker_revisao: ConsultaWorker | None = None
        self.resultados_revisao: list[dict] = []
        self._revisao_paralela_ativa = False
        self.worker: ConsultaWorker | None = None

        self._montar_interface()
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Pronto — escolha uma planilha para começar.")

        # Tecla ESC (2x) para parar as consultas: na 1ª vez apenas muda a cor
        # como aviso; na 2ª vez interrompe de fato.
        self.atalho_parar = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        self.atalho_parar.setContext(Qt.ShortcutContext.ApplicationShortcut)
        self.atalho_parar.activated.connect(self._parar_por_esc)
        self._esc_armado = False
        self._timer_esc = QTimer(self)
        self._timer_esc.setSingleShot(True)
        self._timer_esc.setInterval(3000)
        self._timer_esc.timeout.connect(self._resetar_aviso_esc)

        # Renovação preventiva: relê o token do Brave a cada 2 minutos e
        # avisa se ele estiver perto de vencer.
        self._timer_token = QTimer(self)
        self._timer_token.setInterval(120000)
        self._timer_token.timeout.connect(self._renovar_token_periodicamente)
        self._timer_token.start()

        self._log('Aplicação iniciada. Dica: pressione "ESC" 2x para parar as consultas.')

    def _parar_por_esc(self) -> None:
        """Para com ESC (2x): 1º aperto muda a cor, 2º aperto interrompe."""
        from PyQt6.QtWidgets import QApplication

        # Não interferir com diálogos modais (onde o ESC fecha a janela).
        if QApplication.activeModalWidget() is not None:
            return
        if self.worker is None or not self.worker.isRunning():
            return

        if not self._esc_armado:
            self._esc_armado = True
            self._mostrar_aviso_esc(True)
            self._timer_esc.start()
            self.statusBar().showMessage(
                "Pressione ESC novamente para parar as consultas."
            )
            self._log("ESC pressionado — pressione novamente para parar as consultas.")
            return

        # Segundo aperto: interrompe e restaura o visual normal.
        self._resetar_aviso_esc()
        self.parar()

    def _mostrar_aviso_esc(self, ativo: bool) -> None:
        """Realça (ou restaura) os controles de parada no aviso do ESC."""
        if ativo:
            estilo_botao = (
                "QPushButton { background: #B71C1C; color: #FFFFFF; "
                "font-weight: 600; border: 1px solid #7F0000; "
                "border-radius: 6px; padding: 4px 10px; }"
            )
            estilo_etiqueta = "QLabel { color: #D32F2F; font-weight: 700; }"
            texto = "ESC NOVAMENTE PARA PARAR"
        else:
            estilo_botao = ""
            estilo_etiqueta = ""
            texto = 'PRESSIONE "ESC" 2x PARA PARAR CONSULTA'

        self.botao_parar.setStyleSheet(estilo_botao)
        self.botao_parar_resultados.setStyleSheet(estilo_botao)
        for etiqueta in self.findChildren(QLabel, "dicaEspaco"):
            etiqueta.setStyleSheet(estilo_etiqueta)
            etiqueta.setText(texto)

    def _resetar_aviso_esc(self) -> None:
        """Cancela o aviso do ESC e devolve os controles ao visual normal."""
        self._esc_armado = False
        self._timer_esc.stop()
        self._mostrar_aviso_esc(False)

    @staticmethod
    def _documento_registro(registro: dict) -> str:
        """Documento consultado (CPF, CNPJ, documento ou CIF) do registro."""
        return (
            _campo(registro, "cpf") or _campo(registro, "cnpj")
            or _campo(registro, "documento") or _campo(registro, "cif")
        ).strip()

    def _registros_faltantes(self) -> list[dict]:
        """Registros da planilha que ainda não produziram nenhum resultado."""
        com_resultado = {
            str(r.get("CPF", "")).strip()
            for r in self.resultados
            if str(r.get("CPF", "")).strip()
        }
        return [
            r for r in self.registros
            if self._documento_registro(r)
            and self._documento_registro(r) not in com_resultado
        ]

    def _atualizar_botao_revisar(self) -> None:
        """Habilita Revisar assim que houver planilha ou resultados carregados."""
        if not hasattr(self, "botao_revisar"):
            return
        rodando = (
            (self.worker is not None and self.worker.isRunning())
            or self._revisao_paralela_ativa
        )
        self.botao_revisar.setEnabled(
            not rodando and (bool(self.registros) or bool(self.resultados))
        )

    def _carregar_resultados_salvos(self) -> None:
        """Mostra os resultados já gravados ao abrir a planilha."""
        if self.worker is not None and self.worker.isRunning():
            return
        if not self.caminho_entrada:
            return
        aba = self.combo_aba.currentText() if self.combo_aba.isEnabled() else ""
        arquivo = _arquivo_progresso(self.caminho_entrada, aba)
        progresso = ler_progresso(arquivo) if arquivo else None
        if not progresso:
            return
        resultados = list(progresso.get("resultados") or [])
        if not resultados:
            return
        self.resultados = resultados
        self.tabela_resultados.setRowCount(0)
        for item in self.resultados:
            posicao = self.tabela_resultados.rowCount()
            self.tabela_resultados.insertRow(posicao)
            cor = QColor(COR_COM_DEBITOS if _tem_debitos(item) else COR_SEM_DEBITOS)
            pincel = QBrush(cor)
            for coluna, nome in enumerate(COLUNAS_SAIDA):
                celula = QTableWidgetItem(str(item.get(nome, "")))
                celula.setBackground(pincel)
                self.tabela_resultados.setItem(posicao, coluna, celula)
        processados = int(progresso.get("processados") or 0)
        if self.registros:
            self.barra_progresso.setValue(
                min(100, int(processados * 100 / len(self.registros)))
            )
        self.botao_salvar.setEnabled(True)
        self._log(f"Resultados anteriores carregados: {len(self.resultados)} linha(s).")

    # ---------- Montagem ----------

    def _montar_interface(self) -> None:
        central = QWidget(self)
        self.setCentralWidget(central)
        layout_principal = QVBoxLayout(central)

        abas = QTabWidget()
        layout_principal.addWidget(abas)

        # --- Aba Consulta ---
        aba_consulta = QWidget()
        consulta_layout = QVBoxLayout(aba_consulta)

        grupo_entrada = QGroupBox("Planilha de entrada")
        entrada_layout = QVBoxLayout(grupo_entrada)
        linha_arquivo = QHBoxLayout()
        self.campo_arquivo = QLineEdit()
        self.campo_arquivo.setReadOnly(True)
        self.campo_arquivo.setPlaceholderText("Nenhuma planilha selecionada")
        self.botao_escolher = QPushButton("Escolher planilha...")
        self.botao_escolher.clicked.connect(self.escolher_planilha)
        linha_arquivo.addWidget(self.campo_arquivo)
        linha_arquivo.addWidget(self.botao_escolher)
        entrada_layout.addLayout(linha_arquivo)

        linha_aba = QHBoxLayout()
        linha_aba.addWidget(QLabel("Aba:"))
        self.combo_aba = QComboBox()
        self.combo_aba.setEnabled(False)
        self.combo_aba.currentTextChanged.connect(self._ao_mudar_aba)
        linha_aba.addWidget(self.combo_aba, stretch=1)
        entrada_layout.addLayout(linha_aba)

        self.label_contagem = QLabel("0 registros carregados")
        entrada_layout.addWidget(self.label_contagem)
        self.tabela_entrada = QTableWidget()
        _configurar_tabela(self.tabela_entrada, ["Coluna"])
        entrada_layout.addWidget(self.tabela_entrada)
        consulta_layout.addWidget(grupo_entrada)

        grupo_sessao = QGroupBox("Sessão (opcional)")
        sessao_layout = QHBoxLayout(grupo_sessao)
        sessao_layout.addWidget(QLabel(f"Cookie ({COOKIE_NAME}):"))
        self.campo_cookie = QLineEdit()
        self.campo_cookie.setPlaceholderText("Cookie preenchido automaticamente do Brave")
        sessao_layout.addWidget(self.campo_cookie)
        self.botao_cookie_brave = QPushButton("Buscar no Brave")
        self.botao_cookie_brave.clicked.connect(self._carregar_cookie_brave)
        sessao_layout.addWidget(self.botao_cookie_brave)
        consulta_layout.addWidget(grupo_sessao)

        grupo_execucao = QGroupBox("Execução")
        execucao_layout = QVBoxLayout(grupo_execucao)
        linha_botoes = QHBoxLayout()
        self.botao_executar = QPushButton("Executar")
        self.botao_executar.setEnabled(False)
        self.botao_executar.clicked.connect(self.iniciar)
        self.botao_parar = QPushButton("Parar")
        self.botao_parar.setEnabled(False)
        self.botao_parar.clicked.connect(self.parar)
        self.botao_salvar = QPushButton("Salvar resultado...")
        self.botao_salvar.setEnabled(False)
        self.botao_salvar.clicked.connect(self.salvar_resultado)
        linha_botoes.addWidget(self.botao_executar)
        linha_botoes.addWidget(self.botao_parar)
        linha_botoes.addStretch()
        linha_botoes.addWidget(self.botao_salvar)
        execucao_layout.addLayout(linha_botoes)
        self.campo_revisar_paralelo = QCheckBox(
            "Revisar pendências em paralelo (segunda consulta)"
        )
        self.campo_revisar_paralelo.setToolTip(
            "Enquanto a consulta principal roda, reconfulta em paralelo os itens "
            "com erro, NAO VERIFICADO ou 'Sem débitos'. Os resultados revisados "
            "entram na tabela somente quando as duas consultas terminam."
        )
        execucao_layout.addWidget(self.campo_revisar_paralelo)
        self.barra_progresso = QProgressBar()
        self.barra_progresso.setValue(0)
        execucao_layout.addWidget(self.barra_progresso)
        self.label_status = QLabel("Pronto")
        execucao_layout.addWidget(self.label_status)
        consulta_layout.addWidget(grupo_execucao)

        abas.addTab(aba_consulta, "Consulta")

        # --- Aba Resultados ---
        aba_resultados = QWidget()
        resultados_layout = QVBoxLayout(aba_resultados)
        self.tabela_resultados = QTableWidget()
        _configurar_tabela(self.tabela_resultados, COLUNAS_SAIDA)
        resultados_layout.addWidget(self.tabela_resultados)
        linha_resultados = QHBoxLayout()
        self.botao_revisar = QPushButton("Revisar")
        self.botao_revisar.setEnabled(False)
        self.botao_revisar.setToolTip(
            "Após parar: consulta somente o que ficou faltando. "
            "Depois de concluído: reconulta sem débitos, erros e não verificados"
        )
        self.botao_revisar.clicked.connect(self.revisar)
        linha_resultados.addWidget(self.botao_revisar)
        linha_resultados.addStretch()
        self.botao_parar_resultados = QPushButton("Parar consultas")
        self.botao_parar_resultados.setEnabled(False)
        self.botao_parar_resultados.clicked.connect(self.parar)
        linha_resultados.addWidget(self.botao_parar_resultados)
        resultados_layout.addLayout(linha_resultados)
        abas.addTab(aba_resultados, "Resultados")

        # --- Aba Log ---
        aba_log = QWidget()
        log_layout = QVBoxLayout(aba_log)
        self.campo_log = QPlainTextEdit()
        self.campo_log.setReadOnly(True)
        self.campo_log.setMaximumBlockCount(2000)
        log_layout.addWidget(self.campo_log)
        abas.addTab(aba_log, "Log")
        abas.currentChanged.connect(self._ao_mudar_aba_principal)

        # Tenta carregar automaticamente os cookies da sessão aberta no Brave.
        self.cookie_sessao = ""
        self._carregar_cookie_brave(silent=True)

    # ---------- Ações ----------

    def _carregar_cookie_brave(self, silent: bool = False) -> None:
        """Lê os cookies de sessão do Brave e preenche o campo de cookie."""
        try:
            cookies = buscar_cookies_no_brave()
        except Exception as exc:  # noqa: BLE001
            if not silent:
                QMessageBox.warning(self, "Cookie do Brave", str(exc))
            self._log(f"Não foi possível ler os cookies do Brave: {exc}")
            return
        self.campo_cookie.setText(cookies["__Secure-DA"])
        self.cookie_sessao = cookies["JSESSIONID"]
        self._log("Cookies de sessão obtidos do Brave (JWT + JSESSIONID).")
        aviso = _aviso_expiracao(cookies["__Secure-DA"])
        if aviso:
            self._log(f"ATENÇÃO: {aviso}")
            if not silent:
                self.statusBar().showMessage(aviso)
        if not silent and not aviso:
            self.statusBar().showMessage("Cookies do Brave atualizados.")

    def _renovar_token_periodicamente(self) -> None:
        """Relê o token do Brave periodicamente e avisa antes de expirar."""
        if self.worker is not None and self.worker.isRunning():
            return  # o worker já renova o token durante a consulta
        cookie_atual = self.campo_cookie.text()
        cookie, sessao = _renovar_token(cookie_atual, self.cookie_sessao)
        if cookie != cookie_atual or sessao != self.cookie_sessao:
            self.campo_cookie.setText(cookie)
            self.cookie_sessao = sessao
            self._log("Token de autenticação renovado automaticamente do Brave.")
        aviso = _aviso_expiracao(cookie)
        if aviso:
            self._log(f"ATENÇÃO: {aviso}")
            self.statusBar().showMessage(aviso)

    def _log(self, mensagem: str) -> None:
        hora = datetime.now().strftime("%H:%M:%S")
        self.campo_log.appendPlainText(f"[{hora}] {mensagem}")

    def escolher_planilha(self) -> None:
        caminho, _ = QFileDialog.getOpenFileName(
            self,
            "Escolha a planilha",
            "",
            "Planilhas (*.xlsx *.xlsm *.csv);;Excel (*.xlsx *.xlsm);;CSV (*.csv)",
        )
        if not caminho:
            return

        self.caminho_entrada = Path(caminho)
        self.campo_arquivo.setText(str(self.caminho_entrada))

        extensao = self.caminho_entrada.suffix.lower()
        self.combo_aba.blockSignals(True)
        self.combo_aba.clear()
        if extensao in (".xlsx", ".xlsm"):
            try:
                abas = listar_abas(caminho)
                padrao = aba_padrao(caminho)
                self.combo_aba.addItems(abas)
                if padrao in abas:
                    self.combo_aba.setCurrentText(padrao)
                self.combo_aba.setEnabled(True)
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, "Erro ao ler planilha", str(exc))
                self.combo_aba.blockSignals(False)
                return
        else:
            self.combo_aba.setEnabled(False)
        self.combo_aba.blockSignals(False)
        self._carregar_planilha()

    def _ao_mudar_aba_principal(self, _indice: int) -> None:
        """Ao abrir a aba Resultados, mantém a opção de revisar disponível."""
        self._atualizar_botao_revisar()

    def _ao_mudar_aba(self, _nome: str) -> None:
        if self.caminho_entrada:
            self._carregar_planilha()

    def _carregar_planilha(self) -> None:
        if not self.caminho_entrada:
            return
        aba = self.combo_aba.currentText() if self.combo_aba.isEnabled() else None
        try:
            cabecalho, registros = ler_planilha(str(self.caminho_entrada), aba)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Erro ao ler planilha", str(exc))
            return

        if not registros:
            QMessageBox.warning(self, "Planilha vazia", "Nenhuma linha com dados foi encontrada.")
            return

        self.cabecalho = cabecalho
        self.registros = registros
        self.label_contagem.setText(f"{len(registros)} registro(s) carregado(s)")
        self._preencher_tabela(self.tabela_entrada, cabecalho, registros)
        self.botao_executar.setEnabled(True)
        self.statusBar().showMessage(f"Planilha carregada: {self.caminho_entrada.name}")
        self._log(
            f"Planilha carregada: {self.caminho_entrada.name}"
            + (f" (aba {aba})" if aba else "")
            + f" — {len(registros)} registros"
        )

        # Resultados de uma execução anterior já aparecem ao abrir a
        # planilha, e Revisar fica disponível na aba Resultados.
        self._carregar_resultados_salvos()
        self._atualizar_botao_revisar()

    def _preencher_tabela(
        self, tabela: QTableWidget, colunas: list[str], registros: list[dict]
    ) -> None:
        tabela.setRowCount(0)
        _configurar_tabela(tabela, colunas)
        for linha in registros:
            posicao = tabela.rowCount()
            tabela.insertRow(posicao)
            for coluna, nome in enumerate(colunas):
                tabela.setItem(posicao, coluna, QTableWidgetItem(str(linha.get(nome, ""))))

    def iniciar(self) -> None:
        if self.worker is not None and self.worker.isRunning():
            return
        if not self.registros:
            return

        # Uma execução nova recomeça do zero; só uma parada habilita a
        # revisão automática do que ficou pendente.
        self._consulta_interrompida = False
        self._registros_revisao = []
        self._revisao_concluida.clear()

        # Renova o cookie a partir da sessão aberta no Brave antes de consultar.
        self._carregar_cookie_brave(silent=True)

        aba = self.combo_aba.currentText() if self.combo_aba.isEnabled() else ""
        arquivo_progresso = (
            _arquivo_progresso(self.caminho_entrada, aba) if self.caminho_entrada else None
        )

        indice_inicial = 0
        resultados_iniciais: list[dict] = []
        progresso = ler_progresso(arquivo_progresso) if arquivo_progresso else None
        if progresso and 0 < progresso["processados"] < len(self.registros):
            resposta = QMessageBox.question(
                self,
                "Retomar consulta",
                f"Existe um progresso salvo de {progresso['processados']} de "
                f"{len(self.registros)} registros.\n\nDeseja continuar de onde parou?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if resposta == QMessageBox.StandardButton.Yes:
                indice_inicial = progresso["processados"]
                resultados_iniciais = progresso["resultados"]
            else:
                salvar_progresso(arquivo_progresso, 0, len(self.registros), [])
        elif progresso and progresso["processados"] >= len(self.registros):
            salvar_progresso(arquivo_progresso, 0, len(self.registros), [])

        self.resultados = list(resultados_iniciais)
        self.tabela_resultados.setRowCount(0)
        for item in self.resultados:
            self._ao_item(item)
        self.barra_progresso.setValue(
            int(indice_inicial * 100 / len(self.registros)) if self.registros else 0
        )
        self.label_status.setText("Executando...")
        self.botao_executar.setEnabled(False)
        self.botao_parar.setEnabled(True)
        self.botao_parar_resultados.setEnabled(True)
        self.botao_salvar.setEnabled(False)
        self.botao_escolher.setEnabled(False)
        self.botao_revisar.setEnabled(False)

        # Na revisão paralela, estes documentos ficam só com o segundo worker.
        documentos_revisao: set[str] = set()
        if self.campo_revisar_paralelo.isChecked():
            documentos_revisao = self._documentos_pendentes()

        self.worker = ConsultaWorker(
            self.registros,
            self.campo_cookie.text(),
            cookie_sessao=self.cookie_sessao,
            indice_inicial=indice_inicial,
            arquivo_progresso=arquivo_progresso,
            resultados_iniciais=resultados_iniciais,
            documentos_ignorar=documentos_revisao,
            parent=self,
        )
        self.worker.progresso.connect(self._ao_progresso)
        self.worker.aviso.connect(self._log)
        self.worker.item_concluido.connect(self._ao_item)
        self.worker.finalizado.connect(self._ao_finalizar)
        self.worker.finished.connect(self._limpar_worker)
        self.worker.start()
        if indice_inicial:
            self._log(
                f"Retomando consulta do registro {indice_inicial + 1} "
                f"de {len(self.registros)}."
            )
        else:
            self._log(f"Iniciando consulta de {len(self.registros)} débito(s).")

        if self.campo_revisar_paralelo.isChecked():
            self._iniciar_revisao_paralela(documentos_revisao)

    def parar(self) -> None:
        self._resetar_aviso_esc()
        self._consulta_interrompida = True
        if self.worker_revisao is not None and self.worker_revisao.isRunning():
            self.worker_revisao.stop()
            self._log("Parada solicitada na revisão paralela.")
        if self.worker is not None:
            self.worker.stop()
            self.label_status.setText("Parando após o item atual...")
            self._log("Parada solicitada pelo usuário.")

    def _ao_progresso(self, percentual: int, mensagem: str) -> None:
        self.barra_progresso.setValue(percentual)
        self.label_status.setText(mensagem)
        self.statusBar().showMessage(mensagem)

    def _ao_item(self, item: dict) -> None:
        self.resultados.append(item)
        posicao = self.tabela_resultados.rowCount()
        self.tabela_resultados.insertRow(posicao)
        cor = QColor(COR_COM_DEBITOS if _tem_debitos(item) else COR_SEM_DEBITOS)
        pincel = QBrush(cor)
        for coluna, nome in enumerate(COLUNAS_SAIDA):
            celula = QTableWidgetItem(str(item.get(nome, "")))
            celula.setBackground(pincel)
            self.tabela_resultados.setItem(posicao, coluna, celula)
        if item.get("Status") in ("ERRO", "NAO VERIFICADO"):
            self._log(f"{item['Status']} {item.get('CPF', '')}: {item.get('Erro', '')}")

    def _ao_documento_revisado(self, documento: str, novos: list[dict]) -> None:
        """Substitui, na tabela e na lista, as linhas antigas de um documento."""
        self._revisao_concluida.add(str(documento).strip())
        # Remover linhas antigas (de baixo para cima).
        for linha in range(self.tabela_resultados.rowCount() - 1, -1, -1):
            item = self.tabela_resultados.item(linha, COLUNAS_SAIDA.index("CPF"))
            if item is not None and item.text().strip() == documento.strip():
                self.tabela_resultados.removeRow(linha)
        self.resultados = [
            r for r in self.resultados if str(r.get("CPF", "")).strip() != documento.strip()
        ]
        for item in novos:
            self.resultados.append(item)
            self._ao_item(item)

    def _documentos_pendentes(self) -> set[str]:
        """Documentos com resultado que ainda precisa de revisão."""
        documentos = {
            str(r.get("CPF", "")).strip()
            for r in self.resultados
            if str(r.get("Status", "")).strip().upper() not in ("OK", "VERIFICADO")
            or "sem débitos" in str(r.get("Tipo de Débito", "")).lower()
        }
        documentos.discard("")
        return documentos

    def _iniciar_revisao_paralela(self, documentos: set[str] | None = None) -> None:
        """Revisa pendências em outro worker enquanto a consulta principal roda."""
        if self.worker_revisao is not None and self.worker_revisao.isRunning():
            return
        if documentos is None:
            documentos = self._documentos_pendentes()
        if not documentos:
            self._log("Revisão paralela: não há pendências para revisar.")
            return
        registros = [
            r for r in self.registros
            if self._documento_registro(r) in documentos
        ]
        if not registros:
            self._log("Revisão paralela: registros originais não encontrados.")
            return

        self.resultados_revisao = []
        self.worker_revisao = ConsultaWorker(
            registros,
            self.campo_cookie.text(),
            cookie_sessao=self.cookie_sessao,
            modo_revisao=True,
            parent=self,
        )
        self.worker_revisao.documento_revisado.connect(
            self._ao_documento_revisado_paralelo
        )
        self.worker_revisao.finalizado.connect(self._ao_finalizar_revisao_paralela)
        self.worker_revisao.finished.connect(self._limpar_worker_revisao)
        self.worker_revisao.aviso.connect(self._log)
        self._revisao_paralela_ativa = True
        self.botao_revisar.setEnabled(False)
        self.worker_revisao.start()
        self._log(f"Revisão paralela iniciada: {len(registros)} pendente(s).")

    def _ao_documento_revisado_paralelo(self, documento: str, novos: list[dict]) -> None:
        """Guarda a revisão paralela à parte; a tabela só muda na mesclagem."""
        self.resultados_revisao.extend(novos)

    def _ao_finalizar_revisao_paralela(self, mensagem: str) -> None:
        self._log(f"Revisão paralela: {mensagem}")
        self._revisao_paralela_ativa = False
        self._mesclar_revisao_paralela()

    def _limpar_worker_revisao(self) -> None:
        self.worker_revisao = None

    def _mesclar_revisao_paralela(self) -> None:
        """Aplica ao resultado final o que foi revisado em paralelo."""
        if not self.resultados_revisao:
            self._atualizar_botao_revisar()
            return
        if self.worker is not None and self.worker.isRunning():
            return
        documentos = {
            str(item.get("CPF", "")).strip()
            for item in self.resultados_revisao
            if str(item.get("CPF", "")).strip()
        }
        for documento in documentos:
            for linha in range(self.tabela_resultados.rowCount() - 1, -1, -1):
                celula = self.tabela_resultados.item(
                    linha, COLUNAS_SAIDA.index("CPF")
                )
                if celula is not None and celula.text().strip() == documento:
                    self.tabela_resultados.removeRow(linha)
            self.resultados = [
                r for r in self.resultados
                if str(r.get("CPF", "")).strip() != documento
            ]
        for item in self.resultados_revisao:
            self._ao_item(item)
        total = len(self.resultados_revisao)
        self.resultados_revisao = []
        self._log(f"Revisão paralela mesclada: {total} linha(s) atualizada(s).")
        self.botao_salvar.setEnabled(bool(self.resultados))
        self._atualizar_botao_revisar()

    def revisar(self) -> None:
        """Após uma parada, consulta só o que falta; senão, os itens pendentes."""
        if self.worker is not None and self.worker.isRunning():
            return
        if self.worker_revisao is not None and self.worker_revisao.isRunning():
            return

        continuar = self._consulta_interrompida
        self._consulta_interrompida = False
        if not continuar:
            self._revisao_concluida.clear()

        if continuar:
            # A consulta foi interrompida: rever SOMENTE o que não tem
            # resultado. Se a própria revisão foi interrompida, continuar
            # pelo que não chegou a ser revisado nesta rodada.
            if self._registros_revisao:
                base = self._registros_revisao
                concluidos = self._revisao_concluida
            else:
                base = self.registros
                # Resultados pendentes (erros e 'Sem débitos') também voltam
                # para a revisão, junto com os que não têm resultado.
                pendentes = {
                    str(item.get("CPF", "")).strip()
                    for item in self.resultados
                    if str(item.get("Status", "")).strip().upper() not in ("OK", "VERIFICADO")
                    or "sem débitos" in str(item.get("Tipo de Débito", "")).lower()
                }
                pendentes.discard("")
                concluidos = {
                    str(r.get("CPF", "")).strip()
                    for r in self.resultados
                    if str(r.get("CPF", "")).strip()
                } - pendentes
            registros = [
                r for r in base if self._documento_registro(r) not in concluidos
            ]
            if not registros:
                QMessageBox.information(
                    self, "Revisão", "Não faltam consultas: tudo já foi processado."
                )
                self._registros_revisao = []
                return
        else:
            documentos = {
                str(r.get("CPF", "")).strip()
                for r in self.resultados
                if str(r.get("Status", "")).strip().upper() not in ("OK", "VERIFICADO")
                or "sem débitos" in str(r.get("Tipo de Débito", "")).lower()
            }
            documentos.discard("")
            # Inclui também os registros que ficaram sem resultado após uma
            # parada (ou que nunca chegaram a ser consultados).
            for registro in self._registros_faltantes():
                documento = self._documento_registro(registro)
                if documento:
                    documentos.add(documento)
            if not documentos:
                QMessageBox.information(
                    self, "Revisão", "Não há itens para revisar (todos têm débitos confirmados)."
                )
                return
            registros = [
                r for r in self.registros
                if self._documento_registro(r) in documentos
            ]
            if not registros:
                QMessageBox.warning(
                    self, "Revisão", "Não foi possível localizar os registros originais."
                )
                return

        # Guarda a seleção: se parar de novo, Revisar continua exatamente daqui.
        self._registros_revisao = list(registros)

        self._carregar_cookie_brave(silent=True)
        self.barra_progresso.setValue(0)
        self.label_status.setText("Revisando...")
        self.botao_executar.setEnabled(False)
        self.botao_parar.setEnabled(True)
        self.botao_parar_resultados.setEnabled(True)
        self.botao_salvar.setEnabled(False)
        self.botao_revisar.setEnabled(False)
        self.botao_escolher.setEnabled(False)

        self.worker = ConsultaWorker(
            registros,
            self.campo_cookie.text(),
            cookie_sessao=self.cookie_sessao,
            modo_revisao=True,
            parent=self,
        )
        self.worker.progresso.connect(self._ao_progresso)
        self.worker.aviso.connect(self._log)
        self.worker.documento_revisado.connect(self._ao_documento_revisado)
        self.worker.finalizado.connect(self._ao_finalizar)
        self.worker.finished.connect(self._limpar_worker)
        self.worker.start()
        if continuar:
            self._log(
                f"Revisando {len(registros)} documento(s) que faltavam "
                "após a parada."
            )
        else:
            self._log(f"Revisando {len(registros)} documento(s).")

    def _ao_finalizar(self, mensagem: str) -> None:
        self._resetar_aviso_esc()
        self._log(mensagem)
        self.label_status.setText("Concluído" if "Concluído" in mensagem else mensagem)
        self.botao_executar.setEnabled(True)
        self.botao_parar.setEnabled(False)
        self.botao_parar_resultados.setEnabled(False)
        self.botao_escolher.setEnabled(True)
        self.botao_salvar.setEnabled(bool(self.resultados))
        # Revisar volta a ficar disponível (inclusive após uma parada).
        self._atualizar_botao_revisar()
        self.barra_progresso.setValue(100)
        self.statusBar().showMessage(mensagem)
        if "interrompido" in mensagem.lower():
            if self._registros_revisao:
                faltantes = [
                    r for r in self._registros_revisao
                    if self._documento_registro(r) not in self._revisao_concluida
                ]
            else:
                faltantes = self._registros_faltantes()
            if faltantes:
                aviso = (
                    f"{len(faltantes)} consulta(s) faltando — clique em "
                    "Revisar se quiser revisar somente elas."
                )
                self.label_status.setText(aviso)
                self.statusBar().showMessage(aviso)
                self._log(aviso)
        if "interrompido" in mensagem.lower() and self.resultados:
            self._salvar_automatico(parcial=True)
        # Mescla a revisão paralela quando as duas consultas terminaram.
        if self.resultados_revisao and not self._revisao_paralela_ativa:
            self._mesclar_revisao_paralela()
        if "Concluído" in mensagem and self.caminho_entrada:
            arquivo = _arquivo_progresso(
                self.caminho_entrada,
                self.combo_aba.currentText() if self.combo_aba.isEnabled() else "",
            )
            try:
                if arquivo.exists():
                    arquivo.unlink()
            except Exception:  # noqa: BLE001
                pass

    def _salvar_automatico(self, parcial: bool = False) -> None:
        """Salva o que foi processado até aqui sem abrir diálogo."""
        if not self.resultados:
            return
        if self.caminho_entrada:
            nome = f"{self.caminho_entrada.stem}_resultado"
        else:
            nome = "resultado_debitos"
        if parcial:
            nome += "_parcial"
        nome += ".xlsx"
        destino = _pasta_relatorios() / nome
        try:
            gravar_planilha(str(destino), self.resultados)
        except Exception as exc:  # noqa: BLE001
            self._log(f"Não foi possível salvar o relatório parcial: {exc}")
            QMessageBox.warning(self, "Relatório parcial", f"Não foi possível salvar: {exc}")
            return
        self._log(f"Relatório {'parcial ' if parcial else ''}salvo automaticamente em: {destino}")
        self.statusBar().showMessage(f"Relatório salvo em: {destino}")

    def _limpar_worker(self) -> None:
        self.worker = None

    def salvar_resultado(self) -> None:
        if not self.resultados:
            return
        nome = "resultado_debitos.xlsx"
        if self.caminho_entrada:
            nome = f"{self.caminho_entrada.stem}_resultado.xlsx"
        sugestao = str(_pasta_relatorios() / nome)
        caminho, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar resultado",
            sugestao,
            "Excel (*.xlsx);;CSV (*.csv)",
        )
        if not caminho:
            return
        try:
            destino = gravar_planilha(caminho, self.resultados)
        except PermissionError:
            QMessageBox.critical(
                self, "Erro ao salvar",
                "Não foi possível gravar. Feche a planilha de saída se ela estiver aberta.",
            )
            return
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Erro ao salvar", str(exc))
            return
        self._log(f"Resultado salvo em: {destino}")
        self.statusBar().showMessage(f"Resultado salvo em: {destino}")

    def closeEvent(self, event) -> None:  # noqa: N802
        if self.worker_revisao is not None and self.worker_revisao.isRunning():
            self.worker_revisao.stop()
            self.worker_revisao.wait(3000)
        if self.worker is not None and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(3000)
        event.accept()


def main() -> None:
    aplicacao = QApplication(sys.argv)
    aplicacao.setApplicationName("Consulta de Débitos")
    janela = MainWindow()
    janela.show()
    sys.exit(aplicacao.exec())


if __name__ == "__main__":
    main()
