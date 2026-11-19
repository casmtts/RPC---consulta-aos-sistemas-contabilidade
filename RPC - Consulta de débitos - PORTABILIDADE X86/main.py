"""
Sistema de Consulta de Débitos - GUI completa (PyQt6).

Arquivo principal: concentra interface, leitura/gravação de planilhas,
consulta ao site e processamento em segundo plano.
"""

from __future__ import annotations

import csv
import random
import re
import shutil
import subprocess
import sys
import threading
import unicodedata
from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import Qt, QSettings, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import (
    QBrush, QColor, QKeySequence, QPainter, QPen, QPixmap, QShortcut,
)
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
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
from brave_cookie import (
    buscar_cookies_no_brave,
    segundos_para_expirar,
    DOMINIO_JWT,
    DOMINIO_SESSAO,
)

# Relê o token do Brave quando falta menos de 1 minuto para expirar.
MARGEM_RENOVACAO_TOKEN = 60
# O modo HTTP é opcional: se o módulo estiver indisponível (por exemplo, uma
# dependência não instalada), o programa principal continua abrindo normalmente.
class ConsultaHttpNaoVerificada(Exception):
    """A consulta HTTP não pôde ser confirmada."""


def consultar_debito_http(*args, **kwargs) -> list[dict]:
    """Carrega o modo HTTP somente quando ele for realmente usado."""
    try:
        from consulta_http import (
            consultar_debito_http as _consultar,
            ConsultaNaoVerificada as _ErroHttp,
            ConsultaInterrompidaHttp as _Interrompida,
        )
    except Exception as exc:  # noqa: BLE001
        raise ConsultaHttpNaoVerificada(
            f"Módulo de consulta HTTP indisponível: {exc}"
        ) from exc
    try:
        return _consultar(*args, **kwargs)
    except _Interrompida as exc:
        raise ConsultaInterrompida() from exc
    except _ErroHttp as exc:
        raise ConsultaHttpNaoVerificada(str(exc)) from exc


PLANILLA_CABECALOS = [
    "Nome", "CPF", "Tipo de Débito", "Competência", "Valor (R$)", "Status",
    "CIF", "ID Dívida", "Vencimento",
]
COLUNAS_SAIDA = PLANILLA_CABECALOS + ["Erro"]
SEPARADOR_CSV = ";"

# Pincel em cache para destacar linhas que estão sendo revisadas.
_PINCEL_REVISAO: QBrush | None = None


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


# Paletas dos temas (claro / escuro). A preferência fica salva no QSettings.
TEMAS = {
    "claro": {
        "app_fundo": "#F4F7FA",
        "painel": "#FFFFFF",
        "texto": "#102A43",
        "borda": "#DCE4EC",
        "realce": "#5B9BD5",
        "cabecalho": "#1F4E79",
        "cabecalho_texto": "#FFFFFF",
        "zebra": "#F7FAFC",
        "selecao": "#D6E9FF",
        "botao": "#FFFFFF",
        "botao_hover": "#EDF4FA",
        "botao_press": "#DCEAF6",
        "botao_desativado_texto": "#9AA5B1",
        "botao_desativado_fundo": "#F4F6F8",
        "campo_fundo": "#FFFFFF",
        "aba_fundo": "#E9EEF3",
        "aba_texto": "#486581",
        "aba_selecionada": "#1F4E79",
        "aba_selecionada_texto": "#FFFFFF",
        "progresso_fundo": "#EDF2F7",
        "progresso_barra": "#2E86C1",
        "status_fundo": "#FFFFFF",
        "status_texto": "#486581",
        "titulo_grupo": "#1F4E79",
        "com_debitos": "#E8F5E9",
        "sem_debitos": "#F0F2F5",
        "erro": "#FFEBEE",
        "nao_verificado": "#FFF3E0",
        "texto_ok": "#1B5E20",
        "texto_erro": "#B71C1C",
        "texto_atencao": "#E65100",
        "texto_neutro": "#52606D",
        "linha_borda": "#EEF2F6",
    },
    "escuro": {
        # Preto + laranja + letras brancas.
        "app_fundo": "#000000",
        "painel": "#0A0A0A",
        "texto": "#FFFFFF",
        "borda": "#333333",
        "realce": "#FF8C00",
        "cabecalho": "#FF8C00",
        "cabecalho_texto": "#FFFFFF",
        "zebra": "#141414",
        "selecao": "#FF8C00",
        "botao": "#141414",
        "botao_hover": "#FF8C00",
        "botao_press": "#CC7000",
        "botao_desativado_texto": "#777777",
        "botao_desativado_fundo": "#0A0A0A",
        "campo_fundo": "#0D0D0D",
        "aba_fundo": "#141414",
        "aba_texto": "#FFFFFF",
        "aba_selecionada": "#FF8C00",
        "aba_selecionada_texto": "#FFFFFF",
        "progresso_fundo": "#1A1A1A",
        "progresso_barra": "#FF8C00",
        "status_fundo": "#000000",
        "status_texto": "#FFFFFF",
        "titulo_grupo": "#FF8C00",
        "com_debitos": "#102210",
        "sem_debitos": "#141414",
        "erro": "#2A1010",
        "nao_verificado": "#2A1E0A",
        "texto_ok": "#FFFFFF",
        "texto_erro": "#FFFFFF",
        "texto_atencao": "#FFFFFF",
        "texto_neutro": "#FFFFFF",
        "linha_borda": "#222222",
    },
}

TEMA_ATUAL = "claro"
PALETA = TEMAS[TEMA_ATUAL]

# Compatibilidade com o nome usado em outras partes do código
# (cores das planilhas exportadas permanecem as do tema claro).
COR_CABECALHO = TEMAS["claro"]["cabecalho"]
COR_COM_DEBITOS = TEMAS["claro"]["com_debitos"]
COR_SEM_DEBITOS = TEMAS["claro"]["sem_debitos"]


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
                     resultados: list[dict], em_andamento: bool = False,
                     indice_atual: int = 0, entrada: str = "", aba: str = "") -> None:
    """Grava o progresso, inclusive se uma consulta estava em andamento.

    ``em_andamento=True`` sinaliza que a pesquisa do ``indice_atual`` começou
    mas ainda não terminou — na retomada esse índice será sempre refeito.
    """
    import json
    dados = {
        "processados": processados,
        "total": total,
        "resultados": resultados,
        "em_andamento": em_andamento,
        "indice_atual": indice_atual,
        "entrada": entrada,
        "aba": aba,
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


def _cor_argb(cor: str) -> str:
    """Converte '#RRGGBB' para 'AARRGGBB' (formato exigido pelo openpyxl)."""
    texto = str(cor).lstrip("#").upper()
    if len(texto) == 6:
        return "FF" + texto
    return texto


def gravar_planilha(caminho: str, linhas: list[dict]) -> Path:
    """Grava o resultado em XLSX/CSV.

    No XLSX, todas as linhas de uma mesma pessoa recebem a mesma cor: branco
    para a primeira, cinza-claro para a segunda, alternando a cada novo nome.
    Assim fica fácil separar visualmente os débitos de cada devedor.
    """
    caminho = Path(caminho)
    if caminho.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        ws = wb.active
        ws.title = "Resultado"
        ws.sheet_view.showGridLines = False

        # Cabeçalho: azul, texto branco em negrito e filtro automático.
        ws.append(COLUNAS_SAIDA)
        cor_cabecalho = _cor_argb(COR_CABECALHO)
        for celula in ws[1]:
            celula.fill = PatternFill("solid", start_color=cor_cabecalho,
                                      end_color=cor_cabecalho)
            celula.font = Font(color="FFFFFFFF", bold=True, size=11)
            celula.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "A2"

        # Uma cor por PESSOA (todas as linhas do mesmo nome ficam iguais).
        cores_pessoas = ("FFFFFFFF", "FFF2F2F2")  # branco / cinza-claro
        cor_por_nome: dict[str, str] = {}
        proxima_cor = 0
        linha_borda = Side(style="hair", color="FFD9D9D9")

        for linha in linhas:
            nome = str(linha.get("Nome", "")).strip()
            chave = nome or str(linha.get("CPF", "")).strip() or "(sem nome)"
            if chave not in cor_por_nome:
                cor_por_nome[chave] = cores_pessoas[proxima_cor % len(cores_pessoas)]
                proxima_cor += 1
            cor = cor_por_nome[chave]

            ws.append([linha.get(coluna, "") for coluna in COLUNAS_SAIDA])
            numero_linha = ws.max_row
            for indice, nome_coluna in enumerate(COLUNAS_SAIDA, start=1):
                celula = ws.cell(row=numero_linha, column=indice)
                celula.fill = PatternFill("solid", start_color=cor, end_color=cor)
                celula.border = Border(bottom=linha_borda)
                celula.font = Font(color="FF000000", size=10)
                celula.alignment = Alignment(vertical="center")

                if nome_coluna == "Valor (R$)":
                    # Arredonda a 2 casas e usa o símbolo entre aspas: sem
                    # aspas o Excel pode interpretar a letra R e mostrar RR$.
                    celula.value = round(
                        _para_numero(linha.get(nome_coluna, "")), 2
                    )
                    celula.number_format = '"R$" #.##0,00'
                    celula.alignment = Alignment(
                        horizontal="right", vertical="center"
                    )
                elif nome_coluna in ("CPF", "Competência", "Status"):
                    celula.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )
            ws.row_dimensions[numero_linha].height = 20

        ultima_coluna = get_column_letter(len(COLUNAS_SAIDA))
        ws.auto_filter.ref = f"A1:{ultima_coluna}{max(2, ws.max_row)}"
        larguras = (26, 20, 46, 14, 16, 22, 12, 14, 14, 42)
        for indice, largura in enumerate(larguras, start=1):
            ws.column_dimensions[get_column_letter(indice)].width = largura

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
            # Normaliza o valor para exatamente 2 casas, com vírgula decimal.
            linhas_csv = []
            for linha in linhas:
                copia = dict(linha)
                if "Valor (R$)" in copia:
                    copia["Valor (R$)"] = (
                        f"{_para_numero(copia['Valor (R$)']):.2f}"
                        .replace(".", ",")
                    )
                linhas_csv.append(copia)
            escritor.writerows(linhas_csv)
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
    """Busca o valor pela ordem de prioridade das chaves informadas.

    A prioridade é das OPÇÕES, não das colunas da tabela: sem isso, a coluna
    'Lançado (R$)' (valor original) era escolhida antes de 'Atualizado (R$)'
    (valor devido hoje), subestimando a dívida.
    """
    # 1) correspondência exata (normalizada), na ordem de prioridade
    for opcao in chaves:
        alvo = _sem_acento(opcao)
        for chave, valor in item.items():
            if _sem_acento(chave) == alvo:
                return str(valor).strip()
    # 2) correspondência parcial, na ordem de prioridade
    for opcao in chaves:
        alvo = _sem_acento(opcao)
        for chave, valor in item.items():
            if alvo in _sem_acento(chave):
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
                # Prioriza o valor ATUALIZADO (valor devido hoje). "Lançado"
                # é o valor original e subestima a dívida.
                "valor": _para_numero(_valor_por_chaves(
                    item, [
                        "atualizado (r$)", "dívida (r$)", "divida (r$)",
                        "total (r$)", "valor", "lancado (r$)",
                    ]
                )),
                # Campos antes descartados, agora aproveitados da tabela.
                "cif": _valor_por_chaves(item, ["cif"]),
                "id_divida": _valor_por_chaves(
                    item, ["id dívida", "id divida"]
                ),
                "vencimento": _valor_por_chaves(item, ["vencimento"]),
                "pago": _para_numero(_valor_por_chaves(item, ["pago (r$)"])),
            })

    # Só entram dívidas REALMENTE em aberto: situação exatamente DÉBITO,
    # valor devido maior que zero e nenhum valor pago. Dívida ativa,
    # ajuizado, pago, cancelado etc. ficam de fora — mesmo que o site
    # restaure esses filtros na sessão depois do AJAX.
    debitos = [
        debito for debito in debitos
        if _sem_acento(debito.get("situacao", "")).strip() == "debito"
        and debito["valor"] > 0
        and debito["pago"] <= 0
    ]
    return debitos


class ConsultaNaoVerificada(Exception):
    """A consulta não pôde ser confirmada. Nunca deve virar 'Sem débitos'."""


class ErroSiteNaoVerificada(ConsultaNaoVerificada):
    """O site redirecionou para a página de erro (erro-page.jsf)."""


class Erro404Site(ConsultaNaoVerificada):
    """O site respondeu 404: sessão/viewstate inválido ou fora do ar."""


class ConsultaInterrompida(Exception):
    """O usuário clicou em Parar no meio da consulta."""


def _esperar(pagina, milissegundos: int, deve_parar=None) -> None:
    """Espera em fatias curtas para poder interromper a qualquer momento."""
    restante = max(0, int(milissegundos))
    while restante > 0:
        if deve_parar is not None and deve_parar():
            raise ConsultaInterrompida()
        passo = min(250, restante)
        pagina.wait_for_timeout(passo)
        restante -= passo
    if deve_parar is not None and deve_parar():
        raise ConsultaInterrompida()


def _goto_com_retry(pagina, url: str, timeout_ms: int, deve_parar=None):
    """Navega com algumas tentativas; ERR_ABORTED é comum via CDP."""
    ultimo_erro: Exception | None = None
    for tentativa in range(3):
        if deve_parar is not None and deve_parar():
            raise ConsultaInterrompida()
        try:
            return pagina.goto(
                url, wait_until="domcontentloaded", timeout=timeout_ms
            )
        except Exception as exc:  # noqa: BLE001
            ultimo_erro = exc
            if deve_parar is not None and deve_parar():
                raise ConsultaInterrompida()
            try:
                pagina.wait_for_timeout(1000)
            except Exception:  # noqa: BLE001
                break
    raise ConsultaNaoVerificada(
        f"Não foi possível acessar {url}: {ultimo_erro}"
    )


def _erro_site(pagina) -> str | None:
    """Retorna o código do erro se o site estiver na página de erro."""
    url = (pagina.url or "").lower()
    if "erro-page.jsf" in url:
        encontrado = re.search(r"erro=(\d+)", url)
        return encontrado.group(1) if encontrado else "desconhecido"
    return None


# O usuário quer SOMENTE DÉBITO. Dívida ativa e ajuizado ficam de fora,
# assim como qualquer situação paga/cancelada/prescrita.
SITUACOES_ATIVAS = ("DEBITO",)
SITUACOES_PAGAS_OU_ENCERRADAS = (
    "PAGO", "CANCELADO", "PRESCRITO", "REMIDO", "BAIXADO", "QUITADO",
)


def _tabela_resultados(pagina):
    """Encontra a tabela de débitos pelo cabeçalho, não qualquer tabela da página."""
    for tabela in pagina.query_selector_all("table"):
        linhas = tabela.query_selector_all("tr")
        if not linhas:
            continue
        cabecalhos = [
            (celula.inner_text() or "").strip()
            for celula in linhas[0].query_selector_all("th, td")
        ]
        texto = _sem_acento(" ".join(cabecalhos))
        if any(palavra in texto for palavra in
               ("competencia", "origem", "divida", "situacao", "lancado")):
            return tabela
    return None


def _estado_resultado(pagina) -> tuple[bool, bool, bool]:
    """(tem tabela de resultados, tem linhas de dados, mensagem visível de 'sem registros').

    A verificação de "sem registros" considera apenas texto VISÍVEL e nunca
    anula uma tabela com linhas: mensagens ocultas de painéis antigos causavam
    falso "sem débitos".
    """
    tabela = _tabela_resultados(pagina)
    tem_tabela = tabela is not None
    tem_dados = False
    if tabela is not None:
        for linha in tabela.query_selector_all("tr")[1:]:
            celulas = linha.query_selector_all("td")
            if celulas and any((celula.inner_text() or "").strip() for celula in celulas):
                tem_dados = True
                break

    sem_registros = pagina.evaluate("""() => {
        const trechos = [
            'nenhum registro', 'nenhum debito', 'nenhuma guia', 'sem resultados',
            'nao foram encontrados', 'não foram encontrados',
            'nenhum lancamento', 'nenhum lançamento', 'no records found',
        ];
        for (const elemento of document.querySelectorAll('div, span, p, td, li')) {
            const estilo = window.getComputedStyle(elemento);
            if (estilo.display === 'none' || estilo.visibility === 'hidden' ||
                    elemento.getClientRects().length === 0) {
                continue;
            }
            const texto = (elemento.textContent || '').trim().toLowerCase();
            if (!texto || texto.length > 200) {
                continue;
            }
            if (trechos.some(trecho => texto.includes(trecho))) {
                return true;
            }
        }
        return false;
    }""")
    return tem_tabela, tem_dados, sem_registros


def _sessao_expirada(pagina, incluir_oculto: bool = False) -> bool:
    """Detecta mensagem de sessão expirada na página.

    Por padrão considera só texto visível. Com ``incluir_oculto=True`` lê
    ``body.textContent`` inteiro: é usado apenas quando NÃO há tabela de
    resultados, caso em que a tela de login é a única explicação.
    """
    if incluir_oculto:
        try:
            texto = pagina.evaluate(
                "() => document.body ? document.body.textContent : ''"
            ) or ""
            return "sua sessao expirou" in _sem_acento(texto)
        except Exception:  # noqa: BLE001
            return False

    return pagina.evaluate("""() => {
        for (const elemento of document.querySelectorAll('div, span, p, td')) {
            const estilo = window.getComputedStyle(elemento);
            if (estilo.display === 'none' || estilo.visibility === 'hidden' ||
                    elemento.getClientRects().length === 0) {
                continue;
            }
            const texto = (elemento.textContent || '').trim().toLowerCase();
            if (texto.includes('sua sessão expirou') || texto.includes('sua sessao expirou')) {
                return true;
            }
        }
        return false;
    }""")


def _tipo_e_campo_documento(documento: str, tipo: str = "") -> tuple[str, str]:
    """Resolve o tipo de pessoa e o campo a preencher no formulário do site.

    No ISS Digital, selecionar FÍSICA/JURÍDICA/IMÓVEL revela, via AJAX, um
    campo diferente: CPF, CNPJ ou CIF, respectivamente.
    """
    digitos = re.sub(r"\D", "", documento or "")
    tipo_norm = _sem_acento(tipo).strip().upper()

    # A quantidade de dígitos é a fonte da verdade: CPF tem 11, CNPJ tem 14.
    # Um cabeçalho de planilha enganoso (por exemplo, uma coluna "CNPJ"
    # contendo CPFs) jamais pode fazer o programa selecionar JURÍDICA e
    # digitar o CPF no campo de CNPJ — isso gerava "sem débitos" falso.
    if tipo_norm in ("IMOVEL", "IMÓVEL", "CIF"):
        return "IMOVEL", "#documentoContribuinteCif"
    if len(digitos) == 11:
        return "FISICA", "#documentoContribuinteCpf"
    if len(digitos) == 14:
        return "JURIDICA", "#documentoContribuinteCnpj"

    mapa = {
        "FISICA": ("FISICA", "#documentoContribuinteCpf"),
        "FISICO": ("FISICA", "#documentoContribuinteCpf"),
        "CPF": ("FISICA", "#documentoContribuinteCpf"),
        "JURIDICA": ("JURIDICA", "#documentoContribuinteCnpj"),
        "JURIDICO": ("JURIDICA", "#documentoContribuinteCnpj"),
        "CNPJ": ("JURIDICA", "#documentoContribuinteCnpj"),
    }
    if tipo_norm in mapa:
        return mapa[tipo_norm]
    raise ConsultaNaoVerificada(
        "Não foi possível inferir FÍSICA/JURÍDICA do documento informado "
        "(esperado 11 dígitos para CPF ou 14 para CNPJ)."
    )


def _aplicar_tipo_pesquisa(pagina) -> str:
    """Escolhe 'Competência' no select 'Pesquisar por' do site, se existir.

    Quando a tela permite pesquisar por competência, lançamento ou vencimento,
    escolher competência faz o relatório trazer o período original do débito,
    em vez da referência da cobrança atual. Retorna "competencia" se aplicou.
    """
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
        campo.click(timeout=5000)
        campo.press("ControlOrMeta+A", timeout=5000)
        campo.press_sequentially(str(ano_inicial), delay=30, timeout=5000)
        campo.press("Tab", timeout=5000)
        pagina.wait_for_timeout(1000)

        mes_inicial = pagina.query_selector("#mesInicial_input")
        if mes_inicial:
            mes_inicial.select_option("JANEIRO", timeout=5000)
            pagina.wait_for_timeout(500)
    except Exception:  # noqa: BLE001
        pass


def _normalizar_situacao(pagina) -> None:
    """Marca somente a situação DÉBITO no formulário.

    Dívida ativa, ajuizado e quaisquer situações pagas/canceladas são
    desmarcadas. Quando existe um ``<label>`` associado, ele é clicado para
    que o PrimeFaces atualize o estado no servidor (não apenas o DOM).
    """
    normalizar = """(ativos) => {
        const campos = Array.from(
            document.querySelectorAll('input[name="situacaoDebito"]')
        );
        for (const campo of campos) {
            const deveEstarMarcado = ativos.includes(campo.value);
            if (campo.checked === deveEstarMarcado) {
                continue;
            }
            const rotulo = campo.id
                ? document.querySelector('label[for="' + campo.id + '"]')
                : null;
            if (rotulo) {
                rotulo.click();
            } else {
                campo.checked = deveEstarMarcado;
                campo.dispatchEvent(new Event('change', { bubbles: true }));
            }
        }
    }"""
    try:
        pagina.evaluate(normalizar, list(SITUACOES_ATIVAS))
        pagina.wait_for_timeout(300)
    except Exception as exc:  # noqa: BLE001
        raise ConsultaNaoVerificada(
            f"Não foi possível configurar o filtro de Situação: {exc}"
        ) from exc


def _valor_campo_js(pagina, seletor_documento: str) -> str | None:
    """Lê o valor do campo por JavaScript, sem esperar 'actionability'.

    Após clicar em Pesquisar, o JSF/PrimeFaces re-renderiza o formulário e o
    ``Locator.input_value()`` pode esperar 30 s por um elemento que foi
    temporariamente removido. ``page.evaluate`` não faz essa espera.
    """
    identificador = seletor_documento.lstrip("#")
    script = """(id) => {
        const campo = document.getElementById(id)
            || Array.from(document.querySelectorAll('input'))
                .find(elemento => (elemento.id || '').includes(id));
        return campo ? String(campo.value || '') : null;
    }"""
    try:
        return pagina.evaluate(script, identificador)
    except Exception:  # noqa: BLE001
        return None


def _digitar_documento(pagina, seletor_documento: str, documento: str) -> None:
    """Preenche o campo com retry e timeouts curtos.

    O PrimeFaces re-renderiza o formulário via AJAX; um ``Locator.press`` com
    timeout padrão de 30 s pode travar se o nó for trocado no meio da digitação.
    """
    digitos = "".join(caracter for caracter in documento if caracter.isdigit())
    ultimo_erro: Exception | None = None
    for tentativa in range(4):
        try:
            campo = pagina.locator(seletor_documento)
            campo.click(timeout=5000)
            campo.press("ControlOrMeta+A", timeout=5000)
            campo.press("Backspace", timeout=5000)
            campo.press_sequentially(digitos, delay=20, timeout=10000)
            campo.press("Tab", timeout=5000)
            return
        except Exception as exc:  # noqa: BLE001
            ultimo_erro = exc
            codigo = _erro_site(pagina)
            if codigo:
                raise ErroSiteNaoVerificada(
                    f"O site redirecionou para a página de erro "
                    f"(código {codigo}) ao preencher o documento."
                ) from exc
            pagina.wait_for_timeout(1000)
    raise ConsultaNaoVerificada(
        "Não foi possível preencher o documento no formulário: "
        f"{ultimo_erro}"
    )


def _verificar_pesquisa(pagina, documento: str, seletor_documento: str,
                        tipo_pessoa: str) -> bool | None:
    """Confere documento, tipo e filtros após a pesquisa.

    Retorna ``True`` se o documento foi confirmado, ``None`` se não foi
    possível lê-lo (indeterminado) e levanta ``ConsultaNaoVerificada`` somente
    quando há divergência concreta. Nunca transforma dúvida em "sem débitos".
    """
    documento_confirmado: bool | None = None
    valor = _valor_campo_js(pagina, seletor_documento)
    if valor is not None:
        digitos = "".join(caracter for caracter in valor if caracter.isdigit())
        esperado = "".join(caracter for caracter in documento if caracter.isdigit())
        if digitos != esperado:
            raise ConsultaNaoVerificada(
                f"O site está com o documento {digitos or 'vazio'}, mas foi "
                f"pedido {esperado} — resultado descartado."
            )
        documento_confirmado = True

    # As duas conferências abaixo usam JavaScript (sem espera de elemento).
    try:
        radios = pagina.eval_on_selector_all(
            "input[name='tipoPessoaContribuinte']:checked",
            "els => els.map(e => e.value)",
        )
        if radios and radios != [tipo_pessoa]:
            raise ConsultaNaoVerificada(
                f"Tipo de pessoa mudou após a pesquisa: {radios} "
                f"(esperado {tipo_pessoa})."
            )
    except ConsultaNaoVerificada:
        raise
    except Exception:  # noqa: BLE001
        pass

    try:
        situacoes = pagina.eval_on_selector_all(
            'input[name="situacaoDebito"]:checked',
            "els => els.map(e => e.value).filter(v => v)",
        )
        # O PrimeFaces pode restaurar na resposta os filtros antigos da sessão
        # (DIVIDA_ATIVA / AJUIZADO). Isso não invalida os débitos encontrados.
        # Só rejeitamos situações claramente pagas/encerradas, que esconderiam
        # débitos em aberto.
        indevidas = [
            situacao for situacao in situacoes
            if any(
                situacao.startswith(prefixo)
                for prefixo in SITUACOES_PAGAS_OU_ENCERRADAS
            )
        ]
        if indevidas:
            raise ConsultaNaoVerificada(
                "A pesquisa incluiu situações pagas/encerradas: "
                + ", ".join(indevidas)
            )
    except ConsultaNaoVerificada:
        raise
    except Exception:  # noqa: BLE001
        pass

    return documento_confirmado


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
    deve_parar=None,
    sem_confirmacao: list | None = None,
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

    def verificar_parada() -> None:
        if deve_parar is not None and deve_parar():
            raise ConsultaInterrompida()

    verificar_parada()
    debitos: list[dict] = []
    erro_site_msg: str | None = None
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
            _goto_com_retry(
                pagina, HOME_URL, timeout * 1000, deve_parar
            )

            def verificar_erro_site() -> None:
                codigo = _erro_site(pagina)
                if codigo:
                    raise ErroSiteNaoVerificada(
                        f"O site redirecionou para a página de erro "
                        f"(código {codigo}). Atualize os cookies do Brave e "
                        "tente novamente."
                    )

            resposta = _goto_com_retry(
                pagina, BASE_URL, timeout * 1000, deve_parar
            )
            verificar_erro_site()
            if resposta is None or resposta.status >= 400:
                status = resposta.status if resposta else "sem resposta"
                raise Erro404Site(f"Site retornou status {status}.")

            if "home.jsf" in pagina.url.lower() and "relatorio" not in pagina.url.lower():
                raise ConsultaNaoVerificada(
                    "Site redirecionou para a home (sem contribuinte selecionado na sessão)."
                )

            _esperar(pagina, 800, deve_parar)
            verificar_parada()

            # "Informações do Débito" vem recolhido; abrir antes de mexer no
            # formulário evita que o AJAX do radio esconda o filtro de Situação.
            try:
                if pagina.query_selector("#j_idt379 .ui-fieldset-toggler"):
                    pagina.query_selector("#j_idt379 .ui-fieldset-toggler").click()
                    _esperar(pagina, 250, deve_parar)
            except Exception:  # noqa: BLE001
                pass

            tipo_pessoa, seletor_documento = _tipo_e_campo_documento(cpf, tipo)

            try:
                radio = pagina.query_selector(
                    f"input[name='tipoPessoaContribuinte'][value='{tipo_pessoa}']"
                )
                if not radio:
                    raise ConsultaNaoVerificada(
                        f"Radio de tipo de pessoa {tipo_pessoa!r} não encontrado — "
                        "pesquisa não executada."
                    )
                try:
                    radio.check(force=True, timeout=5000)
                except Exception as exc:  # noqa: BLE001
                    verificar_erro_site()
                    raise ConsultaNaoVerificada(
                        f"Não foi possível selecionar o tipo de pessoa: {exc}"
                    ) from exc
                _esperar(pagina, 400, deve_parar)
                verificar_parada()
                verificar_erro_site()

                _digitar_documento(pagina, seletor_documento, cpf)
                _esperar(pagina, 900, deve_parar)
                verificar_parada()
                verificar_erro_site()

                # Se o site tiver "Pesquisar por: competência/lançamento/
                # vencimento", escolhe competência antes de ajustar o período.
                tipo_pesquisa = _aplicar_tipo_pesquisa(pagina)
                verificar_erro_site()

                # Consultar sempre os últimos 10 anos de competência (após o CPF,
                # para o AJAX do período não re-renderizar o campo com máscara).
                _aplicar_periodo(pagina)
                verificar_erro_site()

                # Em "Informações do Débito", manter marcadas somente as
                # situações de DÉBITO; pagas/canceladas ficam de fora.
                _normalizar_situacao(pagina)
                verificar_erro_site()
            except ErroSiteNaoVerificada:
                raise
            except Exception as exc:  # noqa: BLE001
                # Se o site foi para erro-page.jsf durante a espera do campo,
                # a mensagem do Playwright é só um timeout — troque pela causa.
                verificar_erro_site()
                raise

            # O PrimeFaces pode re-renderizar o formulário depois do CPF, do
            # período e dos checkboxes; o botão some e volta. Esperar e tentar
            # novamente evita falso "botão não encontrado".
            # Em páginas CDP, wait_for_selector pode travar mesmo com o botão
            # presente; query_selector é confiável aqui. Tenta também clicar
            # novamente se o AJAX trocar o nó entre a busca e o clique.
            botao = None
            for tentativa in range(10):
                verificar_parada()
                botao = (
                    pagina.query_selector("#botaoAcaoPesquisar")
                    or pagina.query_selector("button:has-text('Pesquisar')")
                    or pagina.query_selector("input[value='Pesquisar']")
                )
                if botao is not None:
                    try:
                        botao.click()
                        break
                    except Exception:  # noqa: BLE001
                        botao = None
                _esperar(pagina, 500, deve_parar)
            if botao is None:
                raise ConsultaNaoVerificada(
                    "Botão 'Pesquisar' não encontrado após aguardar o site — "
                    "pesquisa não executada."
                )
            _esperar(pagina, 2500, deve_parar)
            verificar_parada()
            verificar_erro_site()

            # Espera até a tabela de resultados (ou a mensagem de "nenhum
            # registro") aparecer. O AJAX do JSF pode demorar mais que 5 s.
            js_resultado_pronto = """() => {
                const tabelas = Array.from(document.querySelectorAll('table'));
                const comResultados = tabelas.some(tabela => {
                    const linhas = tabela.querySelectorAll('tr');
                    if (linhas.length < 2) return false;
                    const cabecalho = Array.from(
                        linhas[0].querySelectorAll('th, td')
                    ).map(celula => celula.textContent || '').join(' ').toLowerCase();
                    if (!cabecalho.includes('compet') &&
                        !cabecalho.includes('origem') &&
                        !cabecalho.includes('situa')) {
                        return false;
                    }
                    // Espera linhas de DADOS preenchidas, não só o cabeçalho.
                    for (let i = 1; i < linhas.length; i++) {
                        const celulas = Array.from(linhas[i].querySelectorAll('td'));
                        if (celulas.some(c => (c.textContent || '').trim())) {
                            return true;
                        }
                    }
                    return false;
                });
                const texto = document.body
                    ? (document.body.innerText || '').toLowerCase() : '';
                const semRegistros = texto.includes('nenhum registro') ||
                    texto.includes('nenhum debito') ||
                    texto.includes('sem resultados') ||
                    texto.includes('nao foram encontrados');
                return comResultados || semRegistros;
            }"""
            # Em vez de wait_for_function (que não pode ser interrompido), faz
            # polling em fatias curtas e verifica o botão Parar a cada volta.
            for _ in range(40):
                verificar_parada()
                try:
                    if pagina.evaluate(js_resultado_pronto):
                        break
                except Exception:  # noqa: BLE001
                    pass
                _esperar(pagina, 400, deve_parar)
            _esperar(pagina, 800, deve_parar)
            verificar_parada()

            # Confere documento/tipo/filtros. O campo pode estar temporariamente
            # indisponível após o AJAX: nesse caso o retorno é None.
            documento_confirmado = _verificar_pesquisa(
                pagina, cpf, seletor_documento, tipo_pessoa
            )

            tem_tabela, tem_dados, sem_registros = _estado_resultado(pagina)

            if tem_tabela and tem_dados:
                # Há linhas: a pesquisa obviamente executou com o documento
                # digitado e conferido antes do clique. Aceita os débitos mesmo
                # que não tenha sido possível reler o campo depois do AJAX.
                debitos: list[dict] = _extrair_tabelas(pagina)
                for debito in debitos:
                    debito["_pesquisa_por"] = tipo_pesquisa
                # Linhas existiam, mas nenhuma era DÉBITO (o site pode ter
                # trazido só dívida ativa/ajuizado): sem débitos confirmado.
            elif tem_tabela:
                if documento_confirmado is not True and sem_confirmacao is None:
                    raise ConsultaNaoVerificada(
                        "Não foi possível confirmar o documento após a pesquisa; "
                        "não é seguro afirmar que não há débitos."
                    )
                if documento_confirmado is not True:
                    sem_confirmacao.append(True)
                # Tabela de resultados presente, mas sem linhas.
                debitos = []
            elif sem_registros and not _sessao_expirada(pagina):
                if documento_confirmado is not True and sem_confirmacao is None:
                    raise ConsultaNaoVerificada(
                        "Não foi possível confirmar o documento após a pesquisa; "
                        "não é seguro afirmar que não há débitos."
                    )
                if documento_confirmado is not True:
                    sem_confirmacao.append(True)
                debitos = []
            elif _sessao_expirada(pagina):
                # Só confiamos na mensagem VISÍVEL: o site mantém um painel
                # oculto com "sua sessão expirou" até em páginas normais.
                if _tentativa == 0:
                    try:
                        novos_cookies = buscar_cookies_no_brave()
                        return _executar_com_playwright(
                            consultar_debito, cpf, nome,
                            novos_cookies["__Secure-DA"], timeout, tipo,
                            novos_cookies["JSESSIONID"], _tentativa=1,
                            deve_parar=deve_parar,
                            sem_confirmacao=sem_confirmacao,
                        )
                    except ConsultaInterrompida:
                        raise
                    except Exception:  # noqa: BLE001
                        pass
                raise ConsultaNaoVerificada(
                    "Sessão expirada no Brave. Entre no sistema no navegador e "
                    "use a aba Ambiente → Atualizar cookies."
                )
            else:
                # Sem tabela e sem mensagem: pode ter sido apenas lentidão do
                # AJAX. Repete uma vez em uma página nova antes de desistir.
                if _tentativa == 0:
                    return _executar_com_playwright(
                        consultar_debito, cpf, nome, cookie_value, timeout,
                        tipo, cookie_sessao, _tentativa=1,
                        deve_parar=deve_parar,
                        sem_confirmacao=sem_confirmacao,
                    )
                raise ConsultaNaoVerificada(
                    "Após a pesquisa não apareceu tabela de resultados nem mensagem "
                    "de 'nenhum registro' — não é seguro afirmar que não há débitos."
                )
        except ErroSiteNaoVerificada as exc:
            erro_site_msg = str(exc)
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

    # Página de erro do site: tenta renovar os cookies do Brave uma vez.
    if erro_site_msg is not None:
        if _tentativa == 0:
            try:
                novos_cookies = buscar_cookies_no_brave()
                return _executar_com_playwright(
                    consultar_debito, cpf, nome,
                    novos_cookies["__Secure-DA"], timeout, tipo,
                    novos_cookies["JSESSIONID"], _tentativa=1,
                    deve_parar=deve_parar,
                    sem_confirmacao=sem_confirmacao,
                )
            except ConsultaInterrompida:
                raise
            except Exception:  # noqa: BLE001
                pass
        raise ErroSiteNaoVerificada(
            f"{erro_site_msg} Se persistir, entre no sistema no Brave e "
            "depois clique em Revisar para reconquistar este item."
        )

    # Resultado vazio pode ser condição de corrida do JSF. Refazer a consulta
    # numa sessão nova antes de aceitar que o CPF não tem débitos.
    if not debitos and _tentativa == 0:
        return _executar_com_playwright(
            consultar_debito, cpf, nome, cookie_value, timeout, tipo,
            cookie_sessao, _tentativa=1, deve_parar=deve_parar,
            sem_confirmacao=sem_confirmacao
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
                 modo_revisao: bool = False, pausa: float = EXTRACTION_DELAY,
                 sem_navegador: bool = False, entrada: str = "", aba: str = "",
                 variar_pausa: bool = False, controle: dict | None = None,
                 documentos_ignorar: set[str] | None = None,
                 parent=None):
        super().__init__(parent)
        self.entrada = entrada
        self.aba = aba
        self.registros = registros
        self.cookie = cookie
        self.cookie_sessao = cookie_sessao
        self.timeout = timeout
        self.indice_inicial = indice_inicial
        self.arquivo_progresso = arquivo_progresso
        self.resultados = list(resultados_iniciais or [])
        self.modo_revisao = modo_revisao
        self.pausa = max(0.0, float(pausa))
        self.sem_navegador = sem_navegador
        self.variar_pausa = variar_pausa
        # Dicionário compartilhado com a interface: permite alterar a pausa
        # DURANTE a execução, sem precisar consultar novamente.
        self.controle = controle if controle is not None else {}
        # Documentos que ficam de fora desta execução porque a revisão
        # paralela está cuidando deles (evita consultar o mesmo CPF 2x).
        self.documentos_ignorar = documentos_ignorar or set()
        self._parar = False

    def stop(self) -> None:
        self._parar = True

    def parar_solicitado(self) -> bool:
        """Consultado pela função de consulta para interromper no meio."""
        return self._parar

    def _publicar(self, documento: str, novos: list[dict], indice: int = 0) -> None:
        """Publica os itens de um documento (normal ou revisão)."""
        for item in novos:
            # Marca a linha da planilha de origem. Usado na retomada para
            # descartar resultados do último item que será refeito.
            item["_indice"] = indice
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
            if not self.sem_navegador:
                restante = segundos_para_expirar(self.cookie)
                if restante is None or restante <= MARGEM_RENOVACAO_TOKEN:
                    self.cookie, self.cookie_sessao = _renovar_token(
                        self.cookie, self.cookie_sessao
                    )

            documento = (
                _campo(registro, "cpf") or _campo(registro, "cnpj")
                or _campo(registro, "documento") or _campo(registro, "cif")
            )
            # Na execução paralela, estes documentos ficam com a revisão.
            if documento.strip() in self.documentos_ignorar:
                continue
            nome = _campo(registro, "nome")
            tipo = _campo(registro, "tipo de pessoa") or _campo(registro, "tipo pessoa")

            # NUNCA inferir o tipo pelo nome da coluna: uma planilha com
            # cabeçalho "CNPJ" contendo CPF faria o programa marcar JURÍDICA
            # e digitar o CPF no campo errado, gerando "sem débitos" falso.
            digitos_documento = re.sub(r"\D", "", documento)
            tipo_planilha = _sem_acento(tipo).strip().upper()
            if tipo_planilha in ("IMOVEL", "CIF"):
                tipo = "IMOVEL"
            elif len(digitos_documento) == 11:
                tipo = "FISICA"
            elif len(digitos_documento) == 14:
                tipo = "JURIDICA"
            elif not tipo:
                tipo = ""
            rotulo = nome or documento or f"linha {indice}"
            # Competência da época: prioriza a planilha de entrada. O site
            # informa a competência de referência da cobrança atual, que pode
            # ser diferente do período original do débito.
            competencia_original = (
                _campo(registro, "competencia debito")
                or _campo(registro, "competencia guia")
                or _campo(registro, "competencia")
            )
            self.progresso.emit(
                int((indice - 1) * 100 / total), f"Consultando {indice}/{total}: {rotulo}"
            )
            # Avisa no arquivo de progresso que esta consulta COMEÇOU mas ainda
            # não terminou. Se o programa for fechado agora, a retomada refaz
            # este registro em vez de considerá-lo concluído.
            if self.arquivo_progresso is not None:
                salvar_progresso(
                    self.arquivo_progresso, indice - 1, total, self.resultados,
                    em_andamento=True, indice_atual=indice,
                    entrada=self.entrada, aba=self.aba,
                )

            try:
                sem_confirmacao: list[bool] = []
                if self.sem_navegador:
                    debitos = consultar_debito_http(
                        documento, nome, self.cookie, self.cookie_sessao,
                        self.timeout, tipo,
                        deve_parar=self.parar_solicitado,
                    )
                else:
                    # 404 e a página de erro do site (ex.: código 13429)
                    # normalmente são viewstate/sessão expirada: renova os
                    # cookies do Brave, abre sessão nova e tenta de novo.
                    for tentativa in range(3):
                        sem_confirmacao.clear()
                        if tentativa:
                            try:
                                novos = buscar_cookies_no_brave()
                                self.cookie = novos.get(COOKIE_NAME, self.cookie)
                                self.cookie_sessao = novos.get(
                                    "JSESSIONID", self.cookie_sessao
                                )
                            except Exception:  # noqa: BLE001
                                pass
                        try:
                            debitos = _executar_com_playwright(
                                consultar_debito,
                                documento, nome, self.cookie, self.timeout,
                                tipo, self.cookie_sessao,
                                deve_parar=self.parar_solicitado,
                                sem_confirmacao=sem_confirmacao,
                            )
                            if (
                                debitos
                                and any(
                                    d.get("_pesquisa_por") == "competencia"
                                    for d in debitos
                                )
                                and not getattr(self, "_filtro_comp_logado", False)
                            ):
                                self._filtro_comp_logado = True
                                self.aviso.emit(
                                    "Filtro do site 'Pesquisar por: Competência' "
                                    "aplicado — competência original dos débitos."
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
                            if self._parar:
                                raise ConsultaInterrompida() from exc
                            self.aviso.emit(
                                f"{documento}: {exc} — abrindo sessão nova no "
                                f"Brave e refazendo a consulta "
                                f"({tentativa + 2}/3)."
                            )
                            self.progresso.emit(
                                max(0, int((indice - 1) * 100 / total)),
                                f"Refazendo {documento} com sessão nova "
                                f"({tentativa + 2}/3)",
                            )
                            self.msleep(1000)
                if not debitos:
                    descricao = (
                        "Sem débitos (NÃO CONFIRMADO)" if sem_confirmacao
                        else "Sem débitos"
                    )
                    novos = [{
                        "Nome": nome, "CPF": documento,
                        "Tipo de Débito": descricao, "Competência": competencia_original or "-",
                        "Valor (R$)": "0,00", "Status": "VERIFICADO", "Erro": "",
                        "CIF": "-", "ID Dívida": "-", "Vencimento": "-",
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
                            "CIF": debito.get("cif", ""),
                            "ID Dívida": debito.get("id_divida", ""),
                            "Vencimento": debito.get("vencimento", ""),
                        }
                        for debito in debitos
                    ]
                self._publicar(documento, novos, indice)
                ok += 1
            except ConsultaInterrompida:
                # O usuário clicou em Parar no meio da consulta: não publica
                # resultado para este item e encerra o loop imediatamente.
                break
            except (ConsultaNaoVerificada, ConsultaHttpNaoVerificada) as exc:
                erros += 1
                self._publicar(documento, [{
                    "Nome": nome, "CPF": documento,
                    "Tipo de Débito": "-", "Competência": competencia_original or "-",
                    "Valor (R$)": "0,00", "Status": "NAO VERIFICADO", "Erro": str(exc),
                }], indice)
            except Exception as exc:  # noqa: BLE001
                erros += 1
                self._publicar(documento, [{
                    "Nome": nome, "CPF": documento,
                    "Tipo de Débito": "-", "Competência": competencia_original or "-",
                    "Valor (R$)": "0,00", "Status": "ERRO", "Erro": str(exc),
                }], indice)

            processados = indice
            if self.arquivo_progresso is not None:
                salvar_progresso(
                    self.arquivo_progresso, indice, total, self.resultados,
                    em_andamento=False, indice_atual=indice,
                    entrada=self.entrada, aba=self.aba,
                )
            self.progresso.emit(int(indice * 100 / total), f"Processado {indice}/{total}")

            # Intervalo entre consultas, lido a cada registro do controle
            # compartilhado: mudar o campo na interface vale JÁ para a
            # próxima pausa, sem reiniciar a consulta.
            pausa_atual = float(self.controle.get("pausa", self.pausa))
            variar_atual = bool(self.controle.get("variar", self.variar_pausa))
            if indice < total and not self._parar and pausa_atual > 0:
                pausa = pausa_atual
                if variar_atual:
                    pausa *= random.uniform(0.7, 1.3)
                self.msleep(int(pausa * 1000))

        if self._parar:
            self.finalizado.emit(
                f"Processamento interrompido: {ok} OK, {erros} com erro "
                f"({processados}/{total} linhas processadas)."
            )
        else:
            self.finalizado.emit(f"Concluído: {ok} OK, {erros} com erro.")


# ============================================================
# Configuração do ambiente (executada pela interface)
# ============================================================

def _cdp_ativo() -> bool:
    """Verifica se o Brave está com depuração remota na porta configurada."""
    import urllib.request

    try:
        with urllib.request.urlopen(f"{CDP_URL}/json/version", timeout=1) as resposta:
            return resposta.status == 200
    except Exception:  # noqa: BLE001
        return False


class ComandoWorker(QThread):
    """Executa comandos (pip/playwright) em segundo plano, sem travar a GUI."""

    saida = pyqtSignal(str)
    finalizado = pyqtSignal(bool, str)

    def __init__(self, comandos: list[list[str]], diretorio: str | None = None,
                 parent=None):
        super().__init__(parent)
        self.comandos = comandos
        self.diretorio = diretorio

    def run(self) -> None:
        for comando in self.comandos:
            self.saida.emit(f"$ {' '.join(comando)}")
            try:
                processo = subprocess.Popen(
                    comando,
                    cwd=self.diretorio,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                )
            except Exception as exc:  # noqa: BLE001
                self.finalizado.emit(False, f"Falha ao iniciar o comando: {exc}")
                return
            if processo.stdout is not None:
                for linha in processo.stdout:
                    self.saida.emit(linha.rstrip())
            codigo = processo.wait()
            if codigo != 0:
                self.finalizado.emit(
                    False,
                    f"Comando falhou (código {codigo}): {' '.join(comando)}",
                )
                return
        self.finalizado.emit(True, "Instalação concluída com sucesso.")


class BraveWorker(QThread):
    """Abre o Brave com depuração remota (CDP) para reaproveitar a sessão."""

    saida = pyqtSignal(str)
    finalizado = pyqtSignal(bool, str)

    def __init__(self, fechar: bool = False, parent=None):
        super().__init__(parent)
        self.fechar = fechar

    @staticmethod
    def _caminho_brave() -> str | None:
        candidatos = [
            str(
                Path.home()
                / "Applications/Brave Browser.app/Contents/MacOS/Brave Browser"
            ),
            "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
        ]
        for candidato in candidatos:
            if Path(candidato).exists():
                return candidato
        for nome in ("brave-browser", "brave", "Brave Browser"):
            encontrado = shutil.which(nome)
            if encontrado:
                return encontrado
        return None

    def run(self) -> None:
        if _cdp_ativo():
            self.finalizado.emit(
                True, "Brave já está com a depuração ativa na porta 9222."
            )
            return

        if self.fechar:
            self.saida.emit("Fechando o Brave para reabri-lo com depuração...")
            try:
                if sys.platform == "darwin":
                    subprocess.run(
                        ["osascript", "-e", 'tell application "Brave Browser" to quit'],
                        capture_output=True, text=True, timeout=15,
                    )
                else:
                    subprocess.run(
                        ["pkill", "-x", "Brave Browser"],
                        capture_output=True, timeout=10,
                    )
            except Exception as exc:  # noqa: BLE001
                self.saida.emit(f" Não foi possível fechar automaticamente: {exc}")
            self.msleep(1500)

        binario = self._caminho_brave()
        if not binario:
            self.finalizado.emit(
                False,
                "Brave não encontrado. Instale o navegador em /Applications.",
            )
            return

        porta = CDP_URL.rsplit(":", 1)[-1]
        self.saida.emit(f"Abrindo Brave com depuração na porta {porta}...")
        try:
            subprocess.Popen(
                [binario, f"--remote-debugging-port={porta}", HOME_URL],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
        except Exception as exc:  # noqa: BLE001
            self.finalizado.emit(False, f"Não foi possível abrir o Brave: {exc}")
            return

        # Aguarda até 30 s pelo endpoint de depuração.
        for _ in range(60):
            if _cdp_ativo():
                self.finalizado.emit(
                    True, "Brave aberto e conectado (porta 9222)."
                )
                return
            self.msleep(500)
        self.finalizado.emit(
            False,
            "O Brave abriu, mas a porta 9222 não respondeu. "
            "Feche todas as janelas do Brave e tente novamente.",
        )


# ============================================================
# Interface gráfica
# ============================================================

def _estilo_aplicacao() -> str:
    """Folha de estilo global construída com a paleta do tema ativo."""
    p = PALETA
    return f"""
        QWidget {{
            font-family: "Inter", "SF Pro Text", "Segoe UI", sans-serif;
            font-size: 13px;
            color: {p["texto"]};
        }}
        QMainWindow {{ background: {p["app_fundo"]}; }}
        QGroupBox {{
            background: {p["painel"]};
            border: 1px solid {p["borda"]};
            border-radius: 8px;
            margin-top: 12px;
            padding: 10px 12px 12px 12px;
            font-weight: 600;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 6px;
            color: {p["titulo_grupo"]};
        }}
        QPushButton {{
            background: {p["botao"]};
            border: 1px solid {p["borda"]};
            border-radius: 6px;
            padding: 7px 14px;
            min-height: 18px;
            color: {p["texto"]};
        }}
        QPushButton:hover {{ background: {p["botao_hover"]}; border-color: {p["realce"]}; }}
        QPushButton:pressed {{ background: {p["botao_press"]}; }}
        QPushButton:disabled {{
            color: {p["botao_desativado_texto"]};
            background: {p["botao_desativado_fundo"]};
        }}
        QLineEdit, QComboBox, QPlainTextEdit, QDoubleSpinBox {{
            background: {p["campo_fundo"]};
            color: {p["texto"]};
            border: 1px solid {p["borda"]};
            border-radius: 6px;
            padding: 6px 8px;
            selection-background-color: {p["selecao"]};
        }}
        QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus, QDoubleSpinBox:focus {{
            border-color: {p["realce"]};
        }}
        QTabWidget::pane {{
            border: 1px solid {p["borda"]};
            border-radius: 8px;
            background: {p["painel"]};
            top: -1px;
        }}
        QTabBar::tab {{
            background: {p["aba_fundo"]};
            color: {p["aba_texto"]};
            padding: 8px 18px;
            border-top-left-radius: 6px;
            border-top-right-radius: 6px;
            margin-right: 2px;
        }}
        QTabBar::tab:selected {{
            background: {p["aba_selecionada"]};
            color: {p["aba_selecionada_texto"]};
            font-weight: 600;
        }}
        QProgressBar {{
            border: 1px solid {p["borda"]};
            border-radius: 6px;
            background: {p["progresso_fundo"]};
            text-align: center;
            height: 18px;
            color: {p["texto"]};
        }}
        QProgressBar::chunk {{ background: {p["progresso_barra"]}; border-radius: 5px; }}
        QLabel {{ color: {p["texto"]}; }}
        QLabel#dicaEspaco {{
            color: {p["titulo_grupo"]};
            font-weight: 600;
            font-size: 12px;
            padding: 2px 0px;
        }}
        QScrollBar:vertical {{
            background: {p["app_fundo"]};
            width: 12px;
            border: none;
            margin: 0px;
        }}
        QScrollBar::handle:vertical {{
            background: {p["borda"]};
            border-radius: 6px;
            min-height: 24px;
        }}
        QScrollBar::handle:vertical:hover {{ background: {p["realce"]}; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
        QScrollBar:horizontal {{
            background: {p["app_fundo"]};
            height: 12px;
            border: none;
            margin: 0px;
        }}
        QScrollBar::handle:horizontal {{
            background: {p["borda"]};
            border-radius: 6px;
            min-width: 24px;
        }}
        QScrollBar::handle:horizontal:hover {{ background: {p["realce"]}; }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0px; }}
        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{ background: transparent; }}
        QStatusBar {{ background: {p["status_fundo"]}; color: {p["status_texto"]}; }}
        QCheckBox {{ color: {p["texto"]}; }}
    """


def _configurar_tabela(tabela: QTableWidget, colunas: list[str],
                       zebra: bool = False,
                       larguras: list[int | None] | None = None,
                       scroll_sempre: bool = False) -> None:
    """Aplica o visual moderno das tabelas (cabeçalho, zebra e espaçamento)."""
    tabela.setColumnCount(len(colunas))
    tabela.setHorizontalHeaderLabels(colunas)
    tabela.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    tabela.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    tabela.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
    tabela.setShowGrid(False)
    tabela.setWordWrap(False)
    tabela.setAlternatingRowColors(zebra)
    tabela.verticalHeader().setVisible(False)

    # Barra vertical sempre à vista na prévia e nos resultados.
    tabela.setVerticalScrollBarPolicy(
        Qt.ScrollBarPolicy.ScrollBarAlwaysOn if scroll_sempre
        else Qt.ScrollBarPolicy.ScrollBarAsNeeded
    )
    tabela.setHorizontalScrollBarPolicy(
        Qt.ScrollBarPolicy.ScrollBarAsNeeded
    )

    cabecalho = tabela.horizontalHeader()
    cabecalho.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
    cabecalho.setMinimumSectionSize(95)
    cabecalho.setFixedHeight(38)

    # Larguras específicas (None = última coluna ocupa o espaço restante).
    if larguras:
        for indice, largura in enumerate(larguras):
            if indice >= len(colunas):
                break
            if largura is None:
                cabecalho.setSectionResizeMode(
                    indice, QHeaderView.ResizeMode.Stretch
                )
            else:
                cabecalho.setSectionResizeMode(
                    indice, QHeaderView.ResizeMode.Interactive
                )
                tabela.setColumnWidth(indice, largura)

    p = PALETA
    tabela.setStyleSheet(f"""
        QTableWidget {{
            background-color: {p["painel"]};
            color: {p["texto"]};
            alternate-background-color: {p["zebra"]};
            border: 1px solid {p["borda"]};
            border-radius: 8px;
            gridline-color: transparent;
            font-family: "Inter", "SF Pro Text", "Segoe UI", sans-serif;
            font-size: 13px;
        }}
        QTableWidget::item {{
            padding: 6px 10px;
            border-bottom: 1px solid {p["linha_borda"]};
        }}
        QTableWidget::item:selected {{
            background-color: {p["selecao"]};
            color: {p["texto"]};
            border-bottom: 1px solid transparent;
        }}
        QHeaderView::section {{
            background-color: {p["cabecalho"]};
            color: {p["cabecalho_texto"]};
            font-family: "Inter", "SF Pro Text", "Segoe UI", sans-serif;
            font-size: 12px;
            font-weight: 600;
            letter-spacing: 0.3px;
            padding: 9px 10px;
            border: none;
            border-right: 1px solid rgba(255, 255, 255, 0.16);
            text-align: left;
        }}
        QHeaderView::section:last {{
            border-right: none;
        }}
        QTableCornerButton::section {{
            background-color: {p["cabecalho"]};
            border: none;
        }}
    """)


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
        self._cores_nomes: dict[str, int] = {}
        # Revisão após uma parada: consultar somente o que ficou pendente.
        self._consulta_interrompida = False
        self._registros_revisao: list[dict] = []
        self._revisao_concluida: set[str] = set()
        # Revisão em paralelo: worker e resultados isolados até a mesclagem.
        self.worker_revisao: ConsultaWorker | None = None
        self.resultados_revisao: list[dict] = []
        self._revisao_paralela_ativa = False
        self._documentos_em_revisao: set[str] = set()
        self.worker: ConsultaWorker | None = None
        self.config_worker: QThread | None = None
        self._config = QSettings("ConsultaDebitos", "ConsultaDebitos")

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

        # Reabre automaticamente a última planilha usada, para facilitar a
        # retomada de uma consulta interrompida.
        self.botao_reabrir.setEnabled(bool(self._config.value("ultima_planilha", "")))
        self.reabrir_ultima_planilha(silent=True)

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
        """Mostra os resultados já gravados ao reabrir a planilha."""
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
            self._inserir_linha_tabela(item)
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
        self.abas = abas
        layout_principal.addWidget(abas)

        # Botão no canto superior direito para alternar claro/escuro.
        self.botao_tema = QPushButton()
        self.botao_tema.clicked.connect(self._alternar_tema)
        self.abas.setCornerWidget(self.botao_tema, Qt.Corner.TopRightCorner)
        self._atualizar_botao_tema()

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
        self.botao_reabrir = QPushButton("Reabrir última")
        self.botao_reabrir.setToolTip(
            "Reabre a última planilha usada e retoma o progresso salvo"
        )
        self.botao_reabrir.clicked.connect(self.reabrir_ultima_planilha)
        self.botao_reabrir.setEnabled(False)
        linha_arquivo.addWidget(self.campo_arquivo)
        linha_arquivo.addWidget(self.botao_escolher)
        linha_arquivo.addWidget(self.botao_reabrir)
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
        # Mostra no máximo ~9 linhas; o restante fica acessível pela barra
        # vertical que sempre aparece.
        self.tabela_entrada.setMinimumHeight(320)
        self.tabela_entrada.setMaximumHeight(340)
        _configurar_tabela(
            self.tabela_entrada, ["Coluna"], zebra=True, scroll_sempre=True
        )
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
        self.botao_ambiente = QPushButton("Configurar ambiente")
        self.botao_ambiente.clicked.connect(
            lambda: self.abas.setCurrentWidget(self.aba_ambiente)
        )
        sessao_layout.addWidget(self.botao_ambiente)
        consulta_layout.addWidget(grupo_sessao)

        grupo_execucao = QGroupBox("Execução")
        execucao_layout = QVBoxLayout(grupo_execucao)
        linha_botoes = QHBoxLayout()
        self.botao_executar = QPushButton("Executar")
        self.botao_executar.setEnabled(False)
        self.botao_executar.clicked.connect(self.iniciar)
        self.botao_parar = QPushButton("Parar")
        self.botao_parar.setEnabled(False)
        self.botao_parar.setToolTip('Parar as consultas (ou pressione "ESC" 2x...)')
        self.botao_parar.clicked.connect(self.parar)
        self.botao_salvar = QPushButton("Salvar resultado...")
        self.botao_salvar.setEnabled(False)
        self.botao_salvar.clicked.connect(self.salvar_resultado)
        linha_botoes.addWidget(self.botao_executar)
        linha_botoes.addWidget(self.botao_parar)
        linha_botoes.addStretch()
        linha_botoes.addWidget(self.botao_salvar)
        execucao_layout.addLayout(linha_botoes)

        self.label_dica_espaco = QLabel(
            'PRESSIONE "ESC" 2x PARA PARAR CONSULTA'
        )
        self.label_dica_espaco.setObjectName("dicaEspaco")
        execucao_layout.addWidget(self.label_dica_espaco)

        #Alterna entre consulta com navegador (Playwright) e HTTP direto.
        self.campo_sem_navegador = QCheckBox(
            "Tentar SEM navegador (HTTP direto)"
        )
        self.campo_sem_navegador.setToolTip(
            "Marcado: envia o formulário por HTTP, sem abrir navegador "
            "(modo experimental).\n"
            "Desmarcado: usa Playwright — modo padrão e mais confiável.\n"
            "Se a resposta não puder ser confirmada, o item fica como "
            "NAO VERIFICADO, nunca 'sem débitos'."
        )
        execucao_layout.addWidget(self.campo_sem_navegador)

        self.campo_revisar_paralelo = QCheckBox(
            "Revisar pendências em paralelo (segunda consulta)"
        )
        self.campo_revisar_paralelo.setToolTip(
            "Enquanto a consulta principal roda, reconfulta em paralelo os itens "
            "com erro, NAO VERIFICADO ou 'Sem débitos'. Os resultados revisados "
            "entram na tabela somente quando as duas consultas terminam."
        )
        execucao_layout.addWidget(self.campo_revisar_paralelo)

        # Controle de velocidade: pausa entre consultas (anti-detecção).
        linha_velocidade = QHBoxLayout()
        linha_velocidade.addWidget(QLabel("Pausa entre consultas:"))
        self.campo_velocidade = QDoubleSpinBox()
        self.campo_velocidade.setRange(0.0, 60.0)
        self.campo_velocidade.setSingleStep(0.5)
        self.campo_velocidade.setDecimals(1)
        self.campo_velocidade.setSuffix(" s")
        try:
            pausa_inicial = float(
                self._config.value("pausa_consultas", EXTRACTION_DELAY)
            )
        except (TypeError, ValueError):
            pausa_inicial = EXTRACTION_DELAY
        self.campo_velocidade.setValue(pausa_inicial)
        self.campo_velocidade.setToolTip(
            "Intervalo entre um CPF e outro. Aumente se o site começar a "
            "bloquear ou a ficar lento. 0 = sem pausa."
        )
        self.campo_velocidade.valueChanged.connect(self._salvar_velocidade)
        linha_velocidade.addWidget(self.campo_velocidade)
        self.campo_variar_pausa = QCheckBox("Variar a pausa (anti-detecção)")
        self.campo_variar_pausa.setChecked(
            bool(self._config.value("variar_pausa", False))
        )
        self.campo_variar_pausa.setToolTip(
            "Aplica uma variação aleatória de ±30% na pausa para não criar "
            "um padrão exato de tempo entre as consultas."
        )
        self.campo_variar_pausa.stateChanged.connect(self._salvar_velocidade)
        linha_velocidade.addWidget(self.campo_variar_pausa)
        linha_velocidade.addStretch()
        execucao_layout.addLayout(linha_velocidade)

        # Valores atuais da pausa, compartilhados com o worker em execução.
        self.controle_consulta = {
            "pausa": self.campo_velocidade.value(),
            "variar": self.campo_variar_pausa.isChecked(),
        }

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
        self.tabela_resultados.setMinimumHeight(220)
        _configurar_tabela(
            self.tabela_resultados,
            COLUNAS_SAIDA,
            larguras=[185, 150, 225, 105, 120, 215, 90, 100, 100, None],
            scroll_sempre=True,
        )
        resultados_layout.addWidget(self.tabela_resultados)
        linha_resultados = QHBoxLayout()
        self.botao_revisar = QPushButton("Revisar")
        self.botao_revisar.setEnabled(False)
        self.botao_revisar.setToolTip(
            "Após parar: consulta somente o que ficou faltando. "
            "Depois de concluído: reconulta erros, NAO VERIFICADO, "
            "'Sem débitos' e linhas com CIF, ID Dívida ou Vencimento em branco"
        )
        self.botao_revisar.clicked.connect(self.revisar)
        linha_resultados.addWidget(self.botao_revisar)
        linha_resultados.addStretch()
        dica_resultados = QLabel('PRESSIONE "ESC" 2x PARA PARAR CONSULTA')
        dica_resultados.setObjectName("dicaEspaco")
        linha_resultados.addWidget(dica_resultados)
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

        # --- Aba Ambiente (tudo sem precisar do terminal) ---
        self.aba_ambiente = QWidget()
        ambiente_layout = QVBoxLayout(self.aba_ambiente)
        ambiente_layout.addWidget(QLabel(
            "Prepare e verifique o ambiente sem sair deste programa. "
            "As instalações aparecem no log abaixo."
        ))
        self.campo_ambiente = QPlainTextEdit()
        self.campo_ambiente.setReadOnly(True)
        self.campo_ambiente.setMaximumBlockCount(2000)
        ambiente_layout.addWidget(self.campo_ambiente)

        self.botao_instalar_tudo = QPushButton("Instalar tudo")
        self.botao_instalar_tudo.setToolTip(
            "Instala as bibliotecas Python e o navegador Chromium do Playwright"
        )
        self.botao_instalar_tudo.clicked.connect(self.instalar_tudo)
        self.botao_instalar_deps = QPushButton("Instalar dependências")
        self.botao_instalar_deps.clicked.connect(self.instalar_dependencias)
        self.botao_instalar_navegador = QPushButton("Instalar navegador")
        self.botao_instalar_navegador.clicked.connect(self.instalar_navegador)
        linha_instalar = QHBoxLayout()
        for botao in (self.botao_instalar_tudo, self.botao_instalar_deps,
                      self.botao_instalar_navegador):
            linha_instalar.addWidget(botao)
        ambiente_layout.addLayout(linha_instalar)

        self.botao_abrir_brave = QPushButton("Abrir Brave com depuração (9222)")
        self.botao_abrir_brave.clicked.connect(self.abrir_brave)
        self.botao_abrir_sistema = QPushButton("Abrir sistema no Brave")
        self.botao_abrir_sistema.clicked.connect(self.abrir_sistema)
        self.botao_atualizar_cookies = QPushButton("Atualizar cookies")
        self.botao_atualizar_cookies.clicked.connect(
            lambda: self._carregar_cookie_brave()
        )
        linha_brave = QHBoxLayout()
        for botao in (self.botao_abrir_brave, self.botao_abrir_sistema,
                      self.botao_atualizar_cookies):
            linha_brave.addWidget(botao)
        ambiente_layout.addLayout(linha_brave)

        self.botao_verificar = QPushButton("Verificar novamente")
        self.botao_verificar.clicked.connect(self._verificar_ambiente)
        linha_verificar = QHBoxLayout()
        linha_verificar.addWidget(self.botao_verificar)
        linha_verificar.addStretch()
        ambiente_layout.addLayout(linha_verificar)

        self.botoes_config = [
            self.botao_instalar_tudo, self.botao_instalar_deps,
            self.botao_instalar_navegador, self.botao_abrir_brave,
            self.botao_abrir_sistema, self.botao_atualizar_cookies,
        ]
        abas.addTab(self.aba_ambiente, "Ambiente")
        abas.currentChanged.connect(self._ao_mudar_aba_principal)

        # Tenta carregar automaticamente os cookies da sessão aberta no Brave.
        self.cookie_sessao = ""
        self._carregar_cookie_brave(silent=True)
        self._verificar_ambiente()

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

    def _log_ambiente(self, mensagem: str) -> None:
        self._log(mensagem)
        hora = datetime.now().strftime("%H:%M:%S")
        self.campo_ambiente.appendPlainText(f"[{hora}] {mensagem}")

    def _salvar_velocidade(self, _valor=None) -> None:
        """Guarda a pausa e já aplica ao worker que estiver em execução."""
        pausa = self.campo_velocidade.value()
        variar = self.campo_variar_pausa.isChecked()
        self._config.setValue("pausa_consultas", pausa)
        self._config.setValue("variar_pausa", variar)
        if hasattr(self, "controle_consulta"):
            self.controle_consulta["pausa"] = pausa
            self.controle_consulta["variar"] = variar

    @staticmethod
    def _pergunta_sim_nao(titulo: str, texto: str,
                          padrao_sim: bool = False) -> bool:
        """Caixa de confirmação com botões explicitamente em português."""
        caixa = QMessageBox()
        caixa.setWindowTitle(titulo)
        caixa.setText(texto)
        caixa.setIcon(QMessageBox.Icon.Question)
        botao_sim = caixa.addButton("Sim", QMessageBox.ButtonRole.YesRole)
        botao_nao = caixa.addButton("Não", QMessageBox.ButtonRole.NoRole)
        caixa.setDefaultButton(botao_sim if padrao_sim else botao_nao)
        caixa.exec()
        return caixa.clickedButton() is botao_sim

    # ---------- Tema ----------

    def _atualizar_botao_tema(self) -> None:
        if TEMA_ATUAL == "escuro":
            self.botao_tema.setText("☀️ Claro")
            self.botao_tema.setToolTip("Mudar para o tema claro")
        else:
            self.botao_tema.setText("🌙 Escuro")
            self.botao_tema.setToolTip("Mudar para o tema escuro")

    def _alternar_tema(self) -> None:
        """Alterna entre tema claro e escuro e reaplica em toda a interface."""
        global TEMA_ATUAL, PALETA

        TEMA_ATUAL = "escuro" if TEMA_ATUAL == "claro" else "claro"
        PALETA = TEMAS[TEMA_ATUAL]
        self._config.setValue("tema", TEMA_ATUAL)

        aplicacao = QApplication.instance()
        if aplicacao is not None:
            aplicacao.setStyleSheet(_estilo_aplicacao())

        # Tabela de entrada (zebra é controlada pelo stylesheet da tabela).
        colunas_entrada = self.cabecalho or ["Coluna"]
        _configurar_tabela(
            self.tabela_entrada, colunas_entrada, zebra=True, scroll_sempre=True
        )
        if self.cabecalho and self.registros:
            self._preencher_tabela(
                self.tabela_entrada, self.cabecalho, self.registros
            )

        # Tabela de resultados: reaplica as cores de cada linha.
        _configurar_tabela(
            self.tabela_resultados, COLUNAS_SAIDA,
            larguras=[185, 150, 225, 105, 120, 215, 90, 100, 100, None],
            scroll_sempre=True,
        )
        for posicao, item in enumerate(self.resultados):
            self._aplicar_estilo_linha_resultado(posicao, item)

        self._atualizar_botao_tema()
        self._log(f"Tema {TEMA_ATUAL} ativado.")

    # ---------- Ambiente ----------

    def _verificar_ambiente(self) -> None:
        """Mostra na aba Ambiente o estado de cada pré-requisito."""
        import importlib.util

        self.campo_ambiente.appendPlainText("")
        self.campo_ambiente.appendPlainText("--- Verificação do ambiente ---")

        def item(ok: bool, descricao: str) -> None:
            self.campo_ambiente.appendPlainText(
                f"{'OK  ' if ok else 'FALTA'}  {descricao}"
            )

        item(True, f"Python {sys.version.split()[0]}")
        for pacote in ("PyQt6", "playwright", "openpyxl", "cryptography"):
            item(importlib.util.find_spec(pacote) is not None, f"{pacote} instalado")

        diretorio = Path(__file__).resolve().parent
        item(
            (diretorio / "requirements.txt").exists(),
            "requirements.txt encontrado",
        )

        def caminho_chromium() -> str:
            from playwright.sync_api import sync_playwright

            with sync_playwright() as playwright:
                return str(playwright.chromium.executable_path)

        try:
            executavel = Path(_executar_com_playwright(caminho_chromium))
            item(
                executavel.exists(),
                f"Chromium do Playwright: {executavel}"
                if executavel.exists()
                else "Chromium do Playwright não instalado",
            )
        except Exception as exc:  # noqa: BLE001
            item(False, f"Chromium do Playwright: {exc}")

        item(_cdp_ativo(), f"Brave com depuração na porta {CDP_URL.rsplit(':', 1)[-1]}")

        try:
            cookies = buscar_cookies_no_brave()
            item(True, "Cookies do Brave: JWT + JSESSIONID encontrados")
            self.cookie_sessao = cookies["JSESSIONID"]
        except Exception as exc:  # noqa: BLE001
            item(False, f"Cookies do Brave: {exc}")

    def _iniciar_config_worker(self, worker: QThread) -> bool:
        if self.config_worker is not None and self.config_worker.isRunning():
            QMessageBox.information(
                self, "Configuração", "Já existe uma instalação em andamento."
            )
            return False
        self.config_worker = worker
        for botao in self.botoes_config:
            botao.setEnabled(False)
        worker.saida.connect(self._log_ambiente)
        worker.finalizado.connect(self._ao_fim_config)
        worker.finished.connect(self._limpar_config_worker)
        worker.start()
        return True

    def _ao_fim_config(self, sucesso: bool, mensagem: str) -> None:
        self._log_ambiente(mensagem)
        self.statusBar().showMessage(mensagem)
        if sucesso:
            self._carregar_cookie_brave(silent=True)
        self._verificar_ambiente()
        if not sucesso:
            QMessageBox.warning(self, "Configuração", mensagem)

    def _limpar_config_worker(self) -> None:
        self.config_worker = None
        for botao in self.botoes_config:
            botao.setEnabled(True)

    def instalar_dependencias(self) -> None:
        """Instala as bibliotecas do requirements.txt usando este Python."""
        diretorio = str(Path(__file__).resolve().parent)
        comandos = [[
            sys.executable, "-m", "pip", "install", "-r",
            str(Path(diretorio) / "requirements.txt"),
        ]]
        self._iniciar_config_worker(
            ComandoWorker(comandos, diretorio=diretorio, parent=self)
        )

    def instalar_navegador(self) -> None:
        """Instala o Chromium usado pelo Playwright."""
        diretorio = str(Path(__file__).resolve().parent)
        comandos = [[sys.executable, "-m", "playwright", "install", "chromium"]]
        self._iniciar_config_worker(
            ComandoWorker(comandos, diretorio=diretorio, parent=self)
        )

    def instalar_tudo(self) -> None:
        """Instala dependências e navegador em sequência."""
        diretorio = str(Path(__file__).resolve().parent)
        comandos = [
            [sys.executable, "-m", "pip", "install", "-r",
             str(Path(diretorio) / "requirements.txt")],
            [sys.executable, "-m", "playwright", "install", "chromium"],
        ]
        self._iniciar_config_worker(
            ComandoWorker(comandos, diretorio=diretorio, parent=self)
        )

    def abrir_brave(self) -> None:
        """Fecha (se necessário) e abre o Brave com a porta de depuração ativa."""
        if _cdp_ativo():
            QMessageBox.information(
                self, "Brave",
                "O Brave já está com a depuração ativa na porta 9222.",
            )
            return
        if self._pergunta_sim_nao(
            "Abrir Brave com depuração",
            "Para usar a sessão real do Brave é necessário fechar o navegador "
            "e reabri-lo com a depuração ativa (porta 9222).\n\n"
            "Deseja fechar o Brave e abrir novamente?",
        ):
            self._iniciar_config_worker(BraveWorker(fechar=True, parent=self))

    def abrir_sistema(self) -> None:
        """Abre a página do sistema de débitos no Brave."""
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", "-a", "Brave Browser", HOME_URL])
            else:
                import webbrowser

                webbrowser.open(HOME_URL)
            self._log_ambiente(f"Sistema aberto no navegador: {HOME_URL}")
            if not _cdp_ativo():
                self._log_ambiente(
                    "Dica: use 'Abrir Brave com depuração (9222)' para que o "
                    "programa reaproveite a sessão logada."
                )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Abrir sistema", str(exc))

    def escolher_planilha(self) -> None:
        caminho, _ = QFileDialog.getOpenFileName(
            self,
            "Escolha a planilha",
            "",
            "Planilhas (*.xlsx *.xlsm *.csv);;Excel (*.xlsx *.xlsm);;CSV (*.csv)",
        )
        if caminho:
            self._abrir_planilha(caminho)

    def reabrir_ultima_planilha(self, silent: bool = False) -> None:
        """Reabre a última planilha usada (e a aba escolhida)."""
        caminho = str(self._config.value("ultima_planilha", "") or "")
        if not caminho or not Path(caminho).exists():
            self.botao_reabrir.setEnabled(False)
            if not silent:
                QMessageBox.information(
                    self, "Reabrir planilha",
                    "A última planilha não existe mais. Escolha uma planilha.",
                )
            return
        self._abrir_planilha(caminho, str(self._config.value("ultima_aba", "") or ""))
        if not silent:
            self.statusBar().showMessage(
                f"Planilha reaberta: {Path(caminho).name}"
            )

    def _abrir_planilha(self, caminho: str, aba_salva: str = "") -> None:
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
                escolhida = aba_salva if aba_salva in abas else padrao
                if escolhida in abas:
                    self.combo_aba.setCurrentText(escolhida)
                self.combo_aba.setEnabled(True)
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, "Erro ao ler planilha", str(exc))
                self.combo_aba.blockSignals(False)
                return
        else:
            self.combo_aba.setEnabled(False)
        self.combo_aba.blockSignals(False)
        self._carregar_planilha()

        # Guarda a planilha e a aba para poder reabrir/retomar depois.
        self._config.setValue("ultima_planilha", str(self.caminho_entrada))
        self._config.setValue(
            "ultima_aba",
            self.combo_aba.currentText() if self.combo_aba.isEnabled() else "",
        )
        self.botao_reabrir.setEnabled(True)

    def _ao_mudar_aba_principal(self, _indice: int) -> None:
        """Ao abrir a aba Resultados, mantém a opção de revisar disponível."""
        self._atualizar_botao_revisar()

    def _ao_mudar_aba(self, _nome: str) -> None:
        if self.caminho_entrada:
            self._carregar_planilha()
            self._config.setValue("ultima_aba", self.combo_aba.currentText())

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

        # Resultados de uma execução anterior já aparecem ao abrir o
        # programa, e Revisar fica disponível na aba Resultados.
        self._carregar_resultados_salvos()
        self._atualizar_botao_revisar()

    def _preencher_tabela(
        self, tabela: QTableWidget, colunas: list[str], registros: list[dict]
    ) -> None:
        tabela.setRowCount(0)
        _configurar_tabela(tabela, colunas, zebra=True, scroll_sempre=True)
        for linha in registros:
            posicao = tabela.rowCount()
            tabela.insertRow(posicao)
            tabela.setRowHeight(posicao, 30)
            for coluna, nome in enumerate(colunas):
                celula = QTableWidgetItem(str(linha.get(nome, "")))
                alvo = _sem_acento(nome)
                if "valor" in alvo:
                    celula.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                elif "cpf" in alvo or "cnpj" in alvo:
                    celula.setTextAlignment(
                        Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
                    )
                tabela.setItem(posicao, coluna, celula)

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
        entrada = str(self.caminho_entrada) if self.caminho_entrada else ""
        if progresso and 0 < progresso["processados"] < len(self.registros):
            processados = progresso["processados"]
            em_andamento = bool(progresso.get("em_andamento"))
            indice_em_andamento = int(progresso.get("indice_atual") or 0)

            # Sempre refaz o ÚLTIMO registro: se a busca foi interrompida no
            # meio, o resultado dele não pode ser considerado confiável.
            if em_andamento and 1 <= indice_em_andamento <= len(self.registros):
                refazer = indice_em_andamento
                detalhe = " (a última consulta foi interrompida antes de terminar)"
            else:
                refazer = processados
                detalhe = " (o último registro será checado novamente)"
            refazer = max(1, min(refazer, len(self.registros)))

            resposta = self._pergunta_sim_nao(
                "Retomar consulta",
                f"Existe um progresso salvo de {processados} de "
                f"{len(self.registros)} registros.\n\n"
                f"Deseja continuar? O registro {refazer} será consultado "
                f"novamente{detalhe}.",
                padrao_sim=True,
            )
            if resposta:
                indice_inicial = refazer - 1
                # Descarta os resultados do registro que será refeito.
                resultados_iniciais = [
                    item for item in progresso["resultados"]
                    if int(item.get("_indice") or 0) < refazer
                ]
                self._log(
                    f"Retomando no registro {refazer} de {len(self.registros)}"
                    f"{detalhe}."
                )
            else:
                salvar_progresso(
                    arquivo_progresso, 0, len(self.registros), [],
                    entrada=entrada, aba=aba,
                )
        elif progresso and progresso["processados"] >= len(self.registros):
            salvar_progresso(
                arquivo_progresso, 0, len(self.registros), [],
                entrada=entrada, aba=aba,
            )

        self.resultados = list(resultados_iniciais)
        self._cores_nomes = {}
        self.tabela_resultados.setRowCount(0)
        for item in self.resultados:
            self._inserir_linha_tabela(item)
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
            sem_navegador=self.campo_sem_navegador.isChecked(),
            pausa=self.campo_velocidade.value(),
            variar_pausa=self.campo_variar_pausa.isChecked(),
            controle=self.controle_consulta,
            documentos_ignorar=documentos_revisao,
            entrada=entrada, aba=aba,
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

    def _inserir_linha_tabela(self, item: dict) -> None:
        """Insere um item na tabela sem alterar a lista ``resultados``.

        Usado ao popular a tabela com resultados já existentes (retomada),
        evitando que eles sejam adicionados uma segunda vez à lista.
        """
        posicao = self.tabela_resultados.rowCount()
        self.tabela_resultados.insertRow(posicao)
        self.tabela_resultados.setRowHeight(posicao, 30)
        for coluna, nome in enumerate(COLUNAS_SAIDA):
            self.tabela_resultados.setItem(
                posicao, coluna, QTableWidgetItem(str(item.get(nome, "")))
            )
        self._aplicar_estilo_linha_resultado(posicao, item)

    @staticmethod
    def _chave_pessoa(item: dict) -> str:
        """Identifica a pessoa de forma estável para agrupar as linhas."""
        return (
            str(item.get("Nome", "")).strip()
            or str(item.get("CPF", "")).strip()
            or "(sem nome)"
        )

    def _cor_fundo_resultado(self, item: dict) -> str:
        """Branco para a 1ª pessoa, cinza-claro para a 2ª, alternando."""
        chave = self._chave_pessoa(item)
        if chave not in self._cores_nomes:
            self._cores_nomes[chave] = len(self._cores_nomes) % 2
        return (
            PALETA["painel"] if self._cores_nomes[chave] == 0
            else PALETA["sem_debitos"]
        )

    @staticmethod
    def _cor_texto_status(item: dict) -> str:
        """Cor do texto do Status (o fundo é a cor da pessoa)."""
        p = PALETA
        status = str(item.get("Status", "")).strip().upper()
        if status == "ERRO":
            return p["texto_erro"]
        if status == "NAO VERIFICADO":
            return p["texto_atencao"]
        if _tem_debitos(item):
            return p["texto_ok"]
        return p["texto_neutro"]

    def _aplicar_estilo_linha_resultado(self, posicao: int, item: dict) -> None:
        """Aplica cores e alinhamentos; chamado também ao trocar de tema."""
        cor_fundo = self._cor_fundo_resultado(item)
        cor_texto = self._cor_texto_status(item)
        pincel_fundo = QBrush(QColor(cor_fundo))
        pincel_texto = QBrush(QColor(cor_texto))
        for coluna, nome in enumerate(COLUNAS_SAIDA):
            celula = self.tabela_resultados.item(posicao, coluna)
            if celula is None:
                continue
            celula.setBackground(pincel_fundo)
            if nome == "Status":
                celula.setForeground(pincel_texto)
                fonte = celula.font()
                fonte.setBold(True)
                celula.setFont(fonte)
                celula.setTextAlignment(
                    Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
                )
            elif nome == "Valor (R$)":
                celula.setTextAlignment(
                    Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                )
            elif nome in ("CPF", "Competência"):
                celula.setTextAlignment(
                    Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
                )

    def _ao_item(self, item: dict) -> None:
        self.resultados.append(item)
        self._inserir_linha_tabela(item)
        if item.get("Status") in ("ERRO", "NAO VERIFICADO"):
            self._log(
                f"{item.get('Status')} {item.get('CPF', '')}: {item.get('Erro', '')}"
            )

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
            self._ao_item(item)

    @staticmethod
    def _precisa_revisao(item: dict) -> bool:
        """Revisa erros, sem débitos e linhas com campos novos em branco.

        Linhas antigas podem ter vindo sem CIF (número do lote do IPTU), sem
        ID Dívida ou sem Vencimento. Elas voltam para a consulta a fim de
        preencher esses campos.
        """
        if str(item.get("Status", "")).strip().upper() not in ("OK", "VERIFICADO"):
            return True
        descricao = _sem_acento(item.get("Tipo de Débito", ""))
        if "sem debito" in descricao:
            return True
        return any(
            str(item.get(campo, "")).strip() in ("", "-")
            for campo in ("CIF", "ID Dívida", "Vencimento")
        )

    @staticmethod
    def _pincel_em_revisao() -> QBrush:
        """Pincel laranja com faixas diagonais (aspecto de 'em obra')."""
        global _PINCEL_REVISAO
        if _PINCEL_REVISAO is None:
            tecido = QPixmap(16, 16)
            tecido.fill(QColor("#FFE0B2"))
            pintor = QPainter(tecido)
            caneta = QPen(QColor("#FF8C00"))
            caneta.setWidth(4)
            pintor.setPen(caneta)
            for deslocamento in range(-16, 32, 8):
                pintor.drawLine(deslocamento, 16, deslocamento + 16, 0)
            pintor.end()
            _PINCEL_REVISAO = QBrush(tecido)
        return _PINCEL_REVISAO

    def _marcar_em_revisao(self, documentos: set[str]) -> None:
        """Pinta de laranja listrado as linhas dos documentos em revisão."""
        if not documentos:
            return
        pincel = self._pincel_em_revisao()
        coluna_cpf = COLUNAS_SAIDA.index("CPF")
        self._documentos_em_revisao |= set(documentos)
        for linha in range(self.tabela_resultados.rowCount()):
            celula = self.tabela_resultados.item(linha, coluna_cpf)
            if celula is None or celula.text().strip() not in documentos:
                continue
            for coluna in range(self.tabela_resultados.columnCount()):
                item = self.tabela_resultados.item(linha, coluna)
                if item is not None:
                    item.setBackground(pincel)
                    item.setForeground(QBrush(QColor("#000000")))

    def _desmarcar_em_revisao(self) -> None:
        """Restaura as cores normais das linhas que estavam em revisão."""
        if not self._documentos_em_revisao:
            return
        coluna_cpf = COLUNAS_SAIDA.index("CPF")
        for linha in range(self.tabela_resultados.rowCount()):
            celula = self.tabela_resultados.item(linha, coluna_cpf)
            if celula is None:
                continue
            documento = celula.text().strip()
            if documento not in self._documentos_em_revisao:
                continue
            for resultado in self.resultados:
                if str(resultado.get("CPF", "")).strip() == documento:
                    self._aplicar_estilo_linha_resultado(linha, resultado)
                    break
        self._documentos_em_revisao.clear()

    def _documentos_pendentes(self) -> set[str]:
        """Documentos com resultado que ainda precisa de revisão."""
        documentos = {
            str(r.get("CPF", "")).strip()
            for r in self.resultados
            if self._precisa_revisao(r)
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

        self._marcar_em_revisao(documentos)
        self.resultados_revisao = []
        self.worker_revisao = ConsultaWorker(
            registros,
            self.campo_cookie.text(),
            cookie_sessao=self.cookie_sessao,
            modo_revisao=True,
            sem_navegador=self.campo_sem_navegador.isChecked(),
            pausa=max(0.5, self.campo_velocidade.value()),
            variar_pausa=self.campo_variar_pausa.isChecked(),
            controle=self.controle_consulta,
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
        self._desmarcar_em_revisao()

    def _limpar_worker_revisao(self) -> None:
        self.worker_revisao = None

    def _mesclar_revisao_paralela(self) -> None:
        """Aplica ao resultado final o que foi revisado em paralelo."""
        if not self.resultados_revisao:
            self._atualizar_botao_revisar()
            return
        if self.worker is not None and self.worker.isRunning():
            return  # mescla quando a consulta principal terminar
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

        cifs_por_cpf: dict[str, set[str]] = {}
        nomes_por_cpf: dict[str, str] = {}
        if continuar:
            # A consulta foi interrompida: rever o que não tem resultado E
            # os itens pendentes (erros, NAO VERIFICADO e, sobretudo,
            # 'Sem débitos (NÃO CONFIRMADO)'). Se a própria revisão foi
            # interrompida, continuar pelo que faltou nesta rodada.
            if self._registros_revisao:
                base = self._registros_revisao
                concluidos = self._revisao_concluida
            else:
                base = self.registros
                pendentes = {
                    str(item.get("CPF", "")).strip()
                    for item in self.resultados
                    if self._precisa_revisao(item)
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
                if self._precisa_revisao(r)
            }
            documentos.discard("")

            # Inclui também os registros que ficaram sem resultado após uma
            # parada (ou que nunca chegaram a ser consultados).
            for registro in self._registros_faltantes():
                documento = self._documento_registro(registro)
                if documento:
                    documentos.add(documento)

            # Coleta os CIFs das dívidas (principalmente IPTU/IMÓVEL) para que
            # a revisão também consulte o imóvel e traga campos que faltavam.
            for resultado in self.resultados:
                if not self._precisa_revisao(resultado):
                    continue
                cpf = str(resultado.get("CPF", "")).strip()
                cif = str(resultado.get("CIF", "")).strip()
                if not cpf or cif in ("", "-"):
                    continue
                cifs_por_cpf.setdefault(cpf, set()).add(cif)
                nomes_por_cpf.setdefault(cpf, str(resultado.get("Nome", "")).strip())
            if not documentos:
                QMessageBox.information(
                    self, "Revisão", "Não há itens para revisar (todos têm débitos confirmados)."
                )
                return

            registros = [
                r for r in self.registros
                if self._documento_registro(r) in documentos
            ]
            if not registros and not cifs_por_cpf:
                QMessageBox.warning(
                    self, "Revisão", "Não foi possível localizar os registros originais."
                )
                return

        # Acrescenta consultas por CIF (tipo IMÓVEL) para buscar os campos que
        # não existiam antes, sem repetir CIFs já presentes na lista.
        existentes = {
            (_campo(registro, "cpf") or _campo(registro, "cnpj")
             or _campo(registro, "cif")).strip()
            for registro in registros
        }
        consultas_cif = 0
        for cpf, cifs in cifs_por_cpf.items():
            nome = nomes_por_cpf.get(cpf, "")
            for cif in sorted(cifs):
                if cif in existentes:
                    continue
                rotulo = f"{nome} (CIF {cif})".strip() if nome else f"CIF {cif}"
                registros.append({
                    "Nome": rotulo,
                    "CIF": cif,
                    "Tipo de Pessoa": "IMOVEL",
                })
                existentes.add(cif)
                consultas_cif += 1

        # Guarda a seleção: se parar de novo, Revisar continua exatamente daqui.
        self._registros_revisao = list(registros)
        self._marcar_em_revisao({
            self._documento_registro(r)
            for r in registros
            if self._documento_registro(r)
        })

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
            sem_navegador=self.campo_sem_navegador.isChecked(),
            pausa=self.campo_velocidade.value(),
            variar_pausa=self.campo_variar_pausa.isChecked(),
            controle=self.controle_consulta,
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
            self._log(
                f"Revisando {len(registros)} documento(s) — incluindo os sem "
                f"débitos, os com campos em branco (CIF/ID/Vencimento) e "
                f"{consultas_cif} consulta(s) por CIF."
            )

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
        # Se a revisão paralela terminou primeiro, mescla agora que a
        # consulta principal também acabou.
        if self.resultados_revisao and not self._revisao_paralela_ativa:
            self._mesclar_revisao_paralela()
        if self._documentos_em_revisao and not (
            self.worker_revisao is not None and self.worker_revisao.isRunning()
        ):
            self._desmarcar_em_revisao()
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
        if self.config_worker is not None and self.config_worker.isRunning():
            self.config_worker.requestInterruption()
            self.config_worker.wait(2000)
        event.accept()


def main() -> None:
    global TEMA_ATUAL, PALETA

    configuracoes = QSettings("ConsultaDebitos", "ConsultaDebitos")
    tema_salvo = str(configuracoes.value("tema", "claro") or "claro")
    TEMA_ATUAL = tema_salvo if tema_salvo in TEMAS else "claro"
    PALETA = TEMAS[TEMA_ATUAL]

    aplicacao = QApplication(sys.argv)
    aplicacao.setApplicationName("Consulta de Débitos")
    aplicacao.setStyleSheet(_estilo_aplicacao())
    janela = MainWindow()
    janela.show()
    sys.exit(aplicacao.exec())


if __name__ == "__main__":
    main()
