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
    QSpinBox,
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
    QScrollArea,
    QStatusBar,
    QStyledItemDelegate,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

# Garante que os módulos locais sejam encontrados
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import BASE_URL, HOME_URL, COOKIE_NAME, EXTRACTION_DELAY, CDP_URL
from preencher_observacao import (
    preencher as gravar_relatorio_com_observacao,
    tem_coluna_observacao,
    registros_observacao_vazia,
    localizar_cabecalho,
)
import brave_cookie
from brave_cookie import (
    buscar_cookies_no_brave,
    segundos_para_expirar,
    DOMINIO_JWT,
    DOMINIO_SESSAO,
)
from verificar_pagos import (
    VerificarPagosWorker,
    caminhos_saida,
    registros_com_observacao,
    registros_para_analise,
    chave_tipo,
    chave_competencia,
    classificar_situacao,
    SITUACOES_VERIFICACAO_ESSENCIAIS,
)
from verificar_cifs import (
    VerificarCifsWorker,
    caminhos_saida_cifs,
    carregar_cifs,
    consultar_imoveis,
)
from disclaimer import adicionar_aviso
from compartilhamento import sincronizar_relatorios
from regras_debito import (
    eh_debito_cobravel, situacao_eh_debito, STATUS_PENDENTE,
    STATUS_ERRO_CONSULTA, mensagem_amigavel_erro,
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
COLUNAS_SAIDA = PLANILLA_CABECALOS + ["Erro", "Situação (site)"]
SEPARADOR_CSV = ";"

# Planilha base desta versão: o relatório contábil copiado para esta pasta.
# Ela é carregada automaticamente ao abrir o programa.
PLANILHA_BASE = (
    Path(__file__).resolve().parent / "RELATORIO_CONTABILIDADE_2022_ANALISTAS.xlsx"
)

# Pincel em cache para destacar linhas que estão sendo revisadas.
_PINCEL_REVISAO: QBrush | None = None

# Os dois workers (consulta principal e revisão paralela) usam o MESMO
# JSESSIONID do Brave. Se consultarem ao mesmo tempo, a sessão JSF do site
# troca o contribuinte selecionado e um vê o documento do outro.
_SESSAO_SITE_LOCK = threading.Lock()

# Papel da marcação "em revisão" nas células da tabela.
_MARCA_REVISAO = Qt.ItemDataRole.UserRole + 20

# A tela de Resultados mostra no maximo este numero de linhas: manter um
# QTableWidgetItem por debito (dezenas de milhares) consumia gigabytes de RAM.
MAX_LINHAS_TABELA = 3000
# O progresso e reescrito no maximo a cada poucos segundos (o arquivo chega
# a dezenas de MB); o sinalizador em_andamento garante a retomagem segura.
INTERVALO_SALVAMENTO_PROGRESSO = 2.0
_ULTIMO_SALVAMENTO_PROGRESSO = 0.0


def _pincel_revisao() -> QBrush:
    """Pincel laranja com faixas diagonais (aspecto de 'em obra')."""
    global _PINCEL_REVISAO
    if _PINCEL_REVISAO is None:
        tecido = QPixmap(16, 16)
        tecido.fill(QColor("#FFE082"))
        pintor = QPainter(tecido)
        caneta = QPen(QColor("#F57C00"))
        caneta.setWidth(6)
        pintor.setPen(caneta)
        for deslocamento in range(-16, 32, 8):
            pintor.drawLine(deslocamento, 16, deslocamento + 16, 0)
        pintor.end()
        _PINCEL_REVISAO = QBrush(tecido)
    return _PINCEL_REVISAO


class _DelegadoRevisao(QStyledItemDelegate):
    """Pinta a faixa laranja por cima da célula, mesmo com stylesheet ativo."""

    def paint(self, painter, option, index) -> None:
        super().paint(painter, option, index)
        if not index.data(_MARCA_REVISAO):
            return
        painter.save()
        painter.setOpacity(0.8)
        painter.fillRect(option.rect, _pincel_revisao())
        painter.restore()


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
    return _pasta_relatorios() / f"{caminho_entrada.stem}_progresso{sufixo}.json"


def _arquivo_progresso_revisao(caminho_entrada: Path, aba: str = "") -> Path:
    """Progresso exclusivo da revisão paralela (separado da consulta)."""
    sufixo = f"_{_sem_acento(aba)}" if aba else ""
    return _pasta_relatorios() / f"{caminho_entrada.stem}_revisao_progresso{sufixo}.json"


def _arquivo_progresso_revisao_completa(caminho_entrada: Path, aba: str = "") -> Path:
    """Progresso da revisão completa, para retomar no documento seguinte."""
    sufixo = f"_{_sem_acento(aba)}" if aba else ""
    return _pasta_relatorios() / (
        f"{caminho_entrada.stem}_revisao_completa_progresso{sufixo}.json"
    )


def _pasta_relatorios() -> Path:
    """Pasta 'relatorios' dentro da pasta do programa (criada se preciso)."""
    pasta = Path(__file__).resolve().parent / "relatorios"
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def _novo_nome_relatorio(stem: str, sufixo: str = "resultado") -> str:
    """Nome inédito no modelo ``RELATORIO_..._DD.MM.YYYY_HHhMMmSs``.

    A data da planilha consultada é substituída pela data/hora da geração e,
    se ainda houver colisão, um número é acrescentado. Assim cada execução
    cria um relatório novo: nunca sobrescreve a planilha consultada nem um
    relatório anterior.
    """
    base = re.sub(r"_\d{2}\.\d{2}\.\d{4}$", "", stem)
    if sufixo:
        # "resultado" já contempla "resultado_parcial" (evita duplicar).
        sufixos = {sufixo} | ({"resultado_parcial"} if sufixo == "resultado" else set())
        if not any(base.lower().endswith(f"_{s.lower()}") for s in sufixos):
            base = f"{base}_{sufixo}"
    agora = datetime.now()
    nome = f"{base}_{agora.strftime('%d.%m.%Y')}_{agora.strftime('%Hh%Mm%Ss')}"
    candidato = nome
    numero = 2
    while (_pasta_relatorios() / f"{candidato}.xlsx").exists():
        candidato = f"{nome}_{numero}"
        numero += 1
    return candidato


def salvar_progresso(caminho: Path, processados: int, total: int,
                     resultados: list[dict], em_andamento: bool = False,
                     indice_atual: int = 0, entrada: str = "", aba: str = "",
                     forcado: bool = False) -> None:
    """Grava o progresso, inclusive se uma consulta estava em andamento.

    ``em_andamento=True`` sinaliza que a pesquisa do ``indice_atual`` começou
    mas ainda não terminou — na retomada esse índice será sempre refeito.
    """
    import json
    import time

    global _ULTIMO_SALVAMENTO_PROGRESSO
    agora = time.monotonic()
    if (
        not forcado
        and agora - _ULTIMO_SALVAMENTO_PROGRESSO < INTERVALO_SALVAMENTO_PROGRESSO
    ):
        return
    _ULTIMO_SALVAMENTO_PROGRESSO = agora
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
    """Lê o progresso salvo, se existir e for válido.

    Também aproveita arquivos de progresso antigos que ficaram na raiz da
    pasta, antes de o destino passar a ser ``relatorios/``.
    """
    import json
    if not caminho.exists():
        antigo = caminho.parent.parent / caminho.name
        if not antigo.exists():
            return None
        caminho = antigo
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


def _ano_debito(item: dict) -> int:
    """Ano de referência do débito: competência, com fallback no vencimento."""
    from datetime import date, datetime

    competencia = str(item.get("Competência", "") or "").strip()
    match = re.search(r"(20\d{2})", competencia)
    if match:
        return int(match.group(1))
    match = re.search(r"/\s*(\d{2})\b", competencia)
    if match:
        return 2000 + int(match.group(1))
    vencimento = item.get("Vencimento", "")
    if isinstance(vencimento, (datetime, date)):
        return vencimento.year
    match = re.search(r"(?:19|20)\d{2}", str(vencimento or ""))
    return int(match.group(0)) if match else 0


def _ordenar_resultados(resultados: list[dict]) -> list[dict]:
    """Contribuintes em ordem alfabética; débitos do mais novo ao mais velho."""
    return sorted(
        resultados,
        key=lambda item: (
            _sem_acento(str(item.get("Nome", "") or "")).strip().lower(),
            -_ano_debito(item),
            int(item.get("_indice") or 0),
            str(item.get("CPF", "") or ""),
        ),
    )


_DEBITOS_ANTIGOS: dict[str, list[dict]] | None = None
_DEBITOS_ANTIGOS_CHAVE = None


# Um snapshot com mais de 30 dias não pode alimentar cobrança: dívidas
# pagas/canceladas nesse intervalo apareceriam como em aberto.
IDADE_MAXIMA_SNAPSHOT_DIAS = 30


def _caminho_relatorio_antigo() -> Path | None:
    """Localiza o relatório MAIS RECENTE de débitos do sistema legado."""
    import time

    padrao = "relatorio-lancamentos-debito-divida-ativa-*.csv"
    bases = (
        Path(__file__).resolve().parent,
        Path(__file__).resolve().parent.parent,
        Path.cwd(),
    )
    encontrados: list[Path] = []
    for base in bases:
        encontrados.extend(base.glob(padrao))
    if not encontrados:
        return None
    mais_recente = max(encontrados, key=lambda caminho: caminho.stat().st_mtime)
    idade_dias = (time.time() - mais_recente.stat().st_mtime) / 86400
    if idade_dias > IDADE_MAXIMA_SNAPSHOT_DIAS:
        return None
    return mais_recente


def _debitos_relatorio_antigo() -> dict[str, list[dict]]:
    """Débitos em aberto do relatório antigo, indexados pelos dígitos do CPF.

    Os CPFs fictícios do sistema legado não são aceitos pelo site, mas os
    débitos deles estão nesse relatório. Carregado uma vez por processo.
    """
    global _DEBITOS_ANTIGOS, _DEBITOS_ANTIGOS_CHAVE
    caminho = _caminho_relatorio_antigo()
    chave = (str(caminho), caminho.stat().st_mtime) if caminho else None
    if _DEBITOS_ANTIGOS is not None and _DEBITOS_ANTIGOS_CHAVE == chave:
        return _DEBITOS_ANTIGOS
    dados: dict[str, list[dict]] = {}
    if caminho is not None:
        try:
            with open(
                caminho, encoding="utf-8-sig", errors="replace", newline=""
            ) as arquivo:
                leitor = csv.reader(arquivo, delimiter=";")
                cabecalho: list[str] | None = None
                for linha in leitor:
                    if cabecalho is None:
                        if any("CPF/CNPJ" in str(coluna) for coluna in linha):
                            cabecalho = [str(coluna).strip() for coluna in linha]
                        continue
                    if len(linha) != len(cabecalho):
                        continue
                    registro = dict(zip(cabecalho, linha))
                    if (
                        str(registro.get("Situação Atual do Débito", ""))
                        .strip()
                        .upper()
                        != "DÉBITO"
                    ):
                        continue
                    # Se a dívida ativa foi paga/cancelada/prescrita, o
                    # débito não pode virar cobrança mesmo constando no
                    # snapshot como DÉBITO.
                    situacao_divida = _sem_acento(
                        registro.get("Situação Atual da Dívida Ativa", "")
                    ).strip()
                    if any(
                        termo in situacao_divida
                        for termo in (
                            "pago", "paga", "cancelad", "prescrit",
                            "remid", "baixad", "quitad",
                        )
                    ):
                        continue
                    digitos = "".join(
                        caracter
                        for caracter in str(registro.get("CPF/CNPJ", ""))
                        if caracter.isdigit()
                    )
                    if not digitos:
                        continue
                    dados.setdefault(digitos, []).append({
                        "tipo_debito": str(
                            registro.get("Origem do Débito", "")
                        ).strip(),
                        "competencia": str(registro.get("Exercício", "")).strip(),
                        "valor": _para_numero(
                            registro.get("Valor Atualizado Débito (R$)", "")
                        ),
                        "cif": str(registro.get("CIF", "")).strip(),
                        "id_divida": str(
                            registro.get("Número do Débito", "")
                        ).strip(),
                        "vencimento": str(
                            registro.get("Vencimento do Débito", "")
                        ).strip(),
                        "situacao": str(
                            registro.get("Situação Atual do Débito", "")
                        ).strip(),
                        "_pendente_confirmacao": True,
                    })
        except Exception:  # noqa: BLE001
            pass
    _DEBITOS_ANTIGOS = dados
    _DEBITOS_ANTIGOS_CHAVE = chave
    return dados


def mapa_observacoes(linhas: list[dict]) -> dict[str, list[str]]:
    """Agrupa os débitos por CPF/CNPJ para a coluna OBSERVAÇÃO.

    Cada débito vira um item "Tipo de Débito – Competência". Ao gravar, todos
    os itens são unidos em UMA ÚNICA LINHA, separados por "; ". Parcelas com
    o mesmo tipo/competência aparecem uma única vez, já que a observação
    mostra somente esses dois campos.
    """
    from collections import OrderedDict

    mapa: dict[str, OrderedDict[str, int]] = {}
    for item in linhas:
        documento = re.sub(r"\D", "", str(item.get("CPF", "")))
        if len(documento) not in (11, 14):
            continue
        if not eh_debito_cobravel(item):
            continue
        tipo = str(item.get("Tipo de Débito", "") or "").strip()
        competencia = str(item.get("Competência", "") or "").strip()
        mapa.setdefault(documento, OrderedDict())[
            f"{tipo} – {competencia}"
        ] = _ano_debito(item)
    return {
        doc: [texto for texto, _ in sorted(
            pares.items(), key=lambda par: (-par[1], par[0])
        )]
        for doc, pares in mapa.items()
    }


def gravar_planilha(caminho: str, linhas: list[dict]) -> Path:
    """Grava o resultado em XLSX/CSV.

    No XLSX, todas as linhas de uma mesma pessoa recebem a mesma cor: branco
    para a primeira, cinza-claro para a segunda, alternando a cada novo nome.
    Assim fica fácil separar visualmente os débitos de cada devedor.

    CNPJs (documentos de 14 dígitos) são descartados desta planilha.
    """
    linhas = [
        item for item in linhas
        if len(re.sub(r"\D", "", str(item.get("CPF", "")))) != 14
    ]
    linhas = _ordenar_resultados(linhas)
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

        # Uma cor por CONTRIBUINTE (CPF/CNPJ): alterna ao mudar de pessoa.
        cores_pessoas = ("FFFFFFFF", "FFF2F2F2")  # branco / cinza-claro
        cor_por_pessoa: dict[str, str] = {}
        proxima_cor = 0
        linha_borda = Side(style="hair", color="FFD9D9D9")

        for linha in linhas:
            chave = re.sub(r"\D", "", str(linha.get("CPF", "")))
            if not chave:
                chave = str(linha.get("Nome", "")).strip() or "(sem nome)"
            if chave not in cor_por_pessoa:
                cor_por_pessoa[chave] = cores_pessoas[proxima_cor % len(cores_pessoas)]
                proxima_cor += 1
            cor = cor_por_pessoa[chave]

            ws.append([linha.get(coluna, "") for coluna in COLUNAS_SAIDA])
            numero_linha = ws.max_row
            for indice, nome_coluna in enumerate(COLUNAS_SAIDA, start=1):
                celula = ws.cell(row=numero_linha, column=indice)
                celula.fill = PatternFill("solid", start_color=cor, end_color=cor)
                celula.border = Border(bottom=linha_borda)
                celula.font = Font(color="FF000000", size=10)
                celula.alignment = Alignment(vertical="center")

                if nome_coluna == "Valor (R$)":
                    # Grava como TEXTO já formatado. Um formato numérico
                    # "#.##0,00" muda com o idioma do Excel e pode exibir
                    # "R$ 253,910" em vez de "R$ 253,91".
                    valor = _para_numero(linha.get(nome_coluna, ""))
                    celula.value = f"R$ {valor:.2f}".replace(".", ",")
                    celula.number_format = "General"
                    celula.alignment = Alignment(
                        horizontal="right", vertical="center"
                    )
                elif nome_coluna == "Erro":
                    celula.value = mensagem_amigavel_erro(celula.value)
                elif nome_coluna in ("CPF", "Competência", "Status"):
                    celula.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )
            ws.row_dimensions[numero_linha].height = 20

        ultima_coluna = get_column_letter(len(COLUNAS_SAIDA))
        ws.auto_filter.ref = f"A1:{ultima_coluna}{max(2, ws.max_row)}"
        larguras = (26, 20, 46, 14, 16, 22, 12, 14, 14, 42, 26)
        for indice, largura in enumerate(larguras, start=1):
            ws.column_dimensions[get_column_letter(indice)].width = largura

        # O relatório parcial é a lista de trabalho do gestor: já sai com
        # Status colorido, MOTIVO e duplicatas destacadas, para uma regravação
        # pelo programa não desfazer essa classificação.
        if "resultado_parcial" in caminho.name:
            try:
                from formatar_parcial import formatar_parcial
                formatar_parcial(ws, linhas)
            except Exception:  # noqa: BLE001
                pass
        adicionar_aviso(wb)
        wb.save(caminho)
    else:
        with caminho.open("w", encoding="utf-8-sig", newline="") as arquivo:
            escritor = csv.DictWriter(
                arquivo,
                fieldnames=COLUNAS_SAIDA + ["AVISO"],
                delimiter=SEPARADOR_CSV,
                extrasaction="ignore",
            )
            escritor.writeheader()
            # Normaliza o valor para exatamente 2 casas, com vírgula decimal.
            linhas_csv = []
            for linha in linhas:
                copia = dict(linha)
                # CSV não tem aba AVISO: o aviso acompanha cada linha.
                copia["AVISO"] = (
                    "Dados do ISS Digital/Desenvolve Cidade — relatório "
                    "auxiliar, não é certidão nem prova de pagamento; "
                    "decisão final do gestor público."
                )
                if "Erro" in copia:
                    copia["Erro"] = mensagem_amigavel_erro(copia["Erro"])
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


def _cpf_valido(valor: str) -> bool:
    """Valida os dígitos verificadores de um CPF."""
    digitos = "".join(caracter for caracter in valor if caracter.isdigit())
    if len(digitos) != 11 or len(set(digitos)) == 1:
        return False
    for tamanho in (9, 10):
        soma = sum(
            int(digitos[i]) * (tamanho + 1 - i) for i in range(tamanho)
        )
        resto = (soma * 10) % 11
        esperado = 0 if resto == 10 else resto
        if int(digitos[tamanho]) != esperado:
            return False
    return True


def _cnpj_valido(valor: str) -> bool:
    """Valida os dígitos verificadores de um CNPJ."""
    digitos = "".join(caracter for caracter in valor if caracter.isdigit())
    if len(digitos) != 14 or len(set(digitos)) == 1:
        return False
    pesos = (
        (5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2),
        (6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2),
    )
    for indice, tabela in enumerate(pesos):
        soma = sum(int(digitos[i]) * tabela[i] for i in range(12))
        resto = soma % 11
        esperado = 0 if resto < 2 else 11 - resto
        if int(digitos[12 + indice]) != esperado:
            return False
    return True


def _documento_valido(documento: str, tipo: str = "") -> bool:
    """CPF/CNPJ com dígitos inválidos não chegam a ser consultados no site."""
    if _sem_acento(tipo).strip().upper() in ("IMOVEL", "CIF"):
        return True  # CIF não tem dígito verificador conhecido.
    digitos = "".join(caracter for caracter in documento if caracter.isdigit())
    if len(digitos) == 11:
        return _cpf_valido(digitos)
    if len(digitos) == 14:
        return _cnpj_valido(digitos)
    return True  # formato desconhecido: deixa o site decidir.


def _campo(registro: dict, nome: str) -> str:
    """Busca o valor de uma coluna pelo nome, ignorando acentos/maiúsculas."""
    alvo = _sem_acento(nome)
    for chave, valor in registro.items():
        if alvo in _sem_acento(chave):
            return str(valor).strip()
    return ""


def _config_bool(config, chave: str, padrao: bool) -> bool:
    """Lê um bool do QSettings sem cair na armadilha bool('false') == True.

    O QSettings grava booleanos como texto ('true'/'false'); converter com
    bool() direto faz 'false' virar True e a preferência volta sozinha.
    """
    valor = config.value(chave, padrao, type=bool)
    if isinstance(valor, bool):
        return valor
    return str(valor).strip().lower() in ("true", "1", "yes", "on")


def _competencia_resultado(
    competencia_original: str, debito: dict, modo_revisao: bool
) -> str:
    """Competência exibida no resultado.

    Na revisão de itens já consultados, usa sempre a competência da planilha
    (data em que o débito foi gerado), pois a consulta antiga registrou a
    referência da cobrança atual, e não o período original.
    """
    if modo_revisao or debito.get("_pesquisa_por") != "competencia":
        return competencia_original or debito["competencia"]
    return debito["competencia"]


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
        # Valores monetários: no máximo 2 casas decimais. Sem isso, um texto
        # como "693,910" podia virar "R$ 693,910" na saída.
        return round(float(limpo), 2)
    except ValueError:
        return 0.0


def _extrair_tabelas(page, incluir_pagos: bool = False) -> list[dict]:
    """Extrai débitos das tabelas da página.

    Por padrão, só entram dívidas em aberto (situação DÉBITO). No modo
    ``incluir_pagos`` — usado pela verificação de pagos — todas as linhas
    são preservadas, com a situação original, para que cada débito da
    OBSERVAÇÃO possa ser comparado ao que o site informa hoje.
    """
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

    if incluir_pagos:
        # Verificação de cobrabilidade: preserva TODAS as situações (paga,
        # cancelada, prescrita, dívida ativa...) para cruzar com a
        # OBSERVAÇÃO e decidir se o gestor ainda pode cobrar cada débito.
        return [
            debito for debito in debitos
            if debito.get("tipo_debito") and debito.get("competencia")
        ]

    # Só entram dívidas REALMENTE em aberto: situação exatamente DÉBITO,
    # valor devido maior que zero e nenhum valor pago. Dívida ativa,
    # ajuizado, pago, cancelado etc. ficam de fora — mesmo que o site
    # restaure esses filtros na sessão depois do AJAX.
    debitos = [
        debito for debito in debitos
        if situacao_eh_debito(debito.get("situacao", ""))
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
# Na verificação de cobrabilidade, o site precisa devolver TAMBÉM as
# situações que encerram a dívida; senão não há como cruzar os dados.
SITUACOES_VERIFICACAO = SITUACOES_VERIFICACAO_ESSENCIAIS


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

    # innerText contém somente texto visível e abrange qualquer elemento
    # (th, h1, label...), ao contrário da lista antiga de tags.
    return pagina.evaluate("""() => {
        const texto = (document.body ? document.body.innerText : '').toLowerCase();
        return texto.includes('sua sessão expirou') ||
               texto.includes('sua sessao expirou') ||
               texto.includes('faça o login novamente') ||
               texto.includes('faca o login novamente');
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


def _normalizar_situacao(pagina, situacoes: tuple[str, ...] | None = None) -> None:
    """Ajusta o filtro de situação do formulário.

    Por padrão marca somente DÉBITO (dívida ativa, ajuizado e situações
    pagas/canceladas ficam de fora). A verificação de cobrabilidade passa
    ``SITUACOES_VERIFICACAO`` para ver também as situações que encerram a
    dívida. Quando existe um ``<label>`` associado, ele é clicado para que o
    PrimeFaces atualize o estado no servidor (não apenas o DOM).
    """
    # Reaconsulta os checkboxes a cada volta: o PrimeFaces re-renderiza o
    # painel depois de cada clique e os nós antigos ficam soltos no DOM.
    normalizar = """(ativos) => {
        const campos = () => Array.from(
            document.querySelectorAll('input[name="situacaoDebito"]')
        );
        for (let volta = 0; volta < 3; volta++) {
            let mudou = false;
            for (const campo of campos()) {
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
                mudou = true;
            }
            if (!mudou) break;
        }
    }"""
    try:
        pagina.evaluate(
            normalizar, list(situacoes or SITUACOES_ATIVAS)
        )
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
            # Confere se o valor digitado permaneceu: o PrimeFaces pode
            # re-renderizar e restaurar um documento anterior da sessão.
            valor = _valor_campo_js(pagina, seletor_documento)
            if valor is None or "".join(c for c in valor if c.isdigit()) == digitos:
                return
            ultimo_erro = RuntimeError(f"campo ficou com {valor!r}")
            pagina.wait_for_timeout(500)
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
                        tipo_pessoa: str,
                        incluir_pagos: bool = False) -> bool | None:
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

    if not incluir_pagos:
        try:
            situacoes = pagina.eval_on_selector_all(
                'input[name="situacaoDebito"]:checked',
                "els => els.map(e => e.value).filter(v => v)",
            )
            # O PrimeFaces pode restaurar na resposta os filtros antigos da
            # sessão (DIVIDA_ATIVA / AJUIZADO). Isso não invalida os débitos
            # encontrados. Só rejeitamos situações claramente pagas/encerradas,
            # que esconderiam débitos em aberto. Na verificação de
            # cobrabilidade essas situações são justamente o que queremos ver.
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
            "Token de autenticação VENCIDO — entre no sistema "
            "novamente no navegador."
        )
    if restante <= MARGEM_RENOVACAO_TOKEN:
        return (
            f"Token de autenticação expira em ~{max(1, restante // 60)} min — "
            "mantenha a sessão ativa no navegador."
        )
    return None


def _garantir_documento(pagina, seletor_documento: str, documento: str,
                        deve_parar=None) -> None:
    """Reafirma o documento depois dos filtros AJAX, antes de pesquisar.

    O site pode re-renderizar o campo e restaurar o CPF de uma consulta
    anterior; nesse caso o valor é digitado novamente.
    """
    digitos = "".join(caracter for caracter in documento if caracter.isdigit())
    valor = _valor_campo_js(pagina, seletor_documento)
    if valor is None:
        return
    if "".join(caracter for caracter in valor if caracter.isdigit()) == digitos:
        return
    _digitar_documento(pagina, seletor_documento, documento)
    _esperar(pagina, 400, deve_parar)


def consultar_debito(
    cpf: str,
    nome: str = "",
    cookie_value: str = "",
    timeout: int = 60,
    tipo: str = "",
    cookie_sessao: str = "",
    _tentativa: int = 0,
    deve_parar=None,
    sem_confirmacao: list | None = None,
    incluir_pagos: bool = False,
) -> list[dict]:
    """Consulta um documento (CPF/CNPJ/CIF) no site.

    O formulário exige selecionar primeiro o tipo de pessoa: FÍSICA revela o
    campo de CPF, JURÍDICA o de CNPJ e IMÓVEL o de CIF. Quando ``tipo`` não é
    informado, ele é inferido pela quantidade de dígitos do documento.

    Só devolve lista vazia quando a pesquisa foi executada e confirmadamente
    não há débitos. Qualquer dúvida levanta ConsultaNaoVerificada.

    Com ``incluir_pagos=True`` (verificação de cobrabilidade), o filtro do
    site passa a incluir pagas/canceladas/prescritas e todas as linhas são
    devolvidas com a situação original.
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
                        f"(código {codigo}). Atualize os cookies do navegador e "
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

                # Em "Informações do Débito": na consulta normal, somente
                # DÉBITO; na verificação, todas as situações relevantes.
                if incluir_pagos:
                    _normalizar_situacao(pagina, SITUACOES_VERIFICACAO)
                else:
                    _normalizar_situacao(pagina)
                verificar_erro_site()

                # Garante que o documento não foi trocado pelos AJAX acima.
                _garantir_documento(
                    pagina, seletor_documento, cpf, deve_parar
                )
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
                pagina, cpf, seletor_documento, tipo_pessoa,
                incluir_pagos=incluir_pagos,
            )

            tem_tabela, tem_dados, sem_registros = _estado_resultado(pagina)

            if tem_tabela and tem_dados:
                # NUNCA publicar débitos sem confirmar o documento. Se o
                # campo não pôde ser relido ( sessão CDP trocada, AJAX ), o
                # resultado pode ser de OUTRO contribuinte.
                if documento_confirmado is not True:
                    raise ConsultaNaoVerificada(
                        "Não foi possível confirmar o documento após a pesquisa; "
                        "os débitos exibidos podem ser de outro contribuinte."
                    )
                debitos: list[dict] = _extrair_tabelas(
                    pagina, incluir_pagos=incluir_pagos
                )
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
            elif _sessao_expirada(pagina, incluir_oculto=True):
                # Só confiamos na mensagem VISÍVEL: o site mantém um painel
                # oculto com "sua sessão expirou" até em páginas normais.
                if _tentativa == 0:
                    try:
                        novos_cookies = buscar_cookies_no_brave()
                        return _executar_com_playwright(
                            consultar_debito, cpf, nome,
                            novos_cookies["__Secure-DA"], timeout, tipo,
                            novos_cookies.get("JSESSIONID", cookie_sessao),
                            _tentativa=1,
                            deve_parar=deve_parar,
                            sem_confirmacao=sem_confirmacao,
                            incluir_pagos=incluir_pagos,
                        )
                    except ConsultaInterrompida:
                        raise
                    except Exception:  # noqa: BLE001
                        pass
                raise ConsultaNaoVerificada(
                    "Sessão expirada no site — os cookies do navegador são "
                    "renovados automaticamente antes de cada nova tentativa, "
                    "mas a consulta não pôde ser confirmada."
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
                        incluir_pagos=incluir_pagos,
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
                    novos_cookies.get("JSESSIONID", cookie_sessao),
                    _tentativa=1,
                    deve_parar=deve_parar,
                    sem_confirmacao=sem_confirmacao,
                    incluir_pagos=incluir_pagos,
                )
            except ConsultaInterrompida:
                raise
            except Exception:  # noqa: BLE001
                pass
        raise ErroSiteNaoVerificada(
            f"{erro_site_msg} Se persistir, entre no sistema novamente e "
            "depois clique em Revisar para reconquistar este item."
        )

    # Resultado vazio pode ser condição de corrida do JSF. Refazer a consulta
    # numa sessão nova antes de aceitar que o CPF não tem débitos.
    if not debitos and _tentativa == 0:
        return _executar_com_playwright(
            consultar_debito, cpf, nome, cookie_value, timeout, tipo,
            cookie_sessao, _tentativa=1, deve_parar=deve_parar,
            sem_confirmacao=sem_confirmacao,
            incluir_pagos=incluir_pagos,
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
    documento_iniciado = pyqtSignal(str)
    finalizado = pyqtSignal(str)

    def __init__(self, registros: list[dict], cookie: str = "", timeout: int = 60,
                 cookie_sessao: str = "", indice_inicial: int = 0,
                 arquivo_progresso: Path | None = None,
                 resultados_iniciais: list[dict] | None = None,
                 modo_revisao: bool = False, pausa: float = EXTRACTION_DELAY,
                 sem_navegador: bool = False, entrada: str = "", aba: str = "",
                 variar_pausa: bool = False, controle: dict | None = None,
                 documentos_ignorar: set[str] | None = None,
                 incluir_pagos: bool = False, parent=None):
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
        self.incluir_pagos = incluir_pagos
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
        self._falhas_rede = 0

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

    def _publicar_fallback_antigo(
        self, documento: str, nome: str, competencia: str,
        indice: int, mensagem_erro: str,
    ) -> bool:
        """Mostra débitos do snapshot legado, mas SEM confirmar cobrança.

        O arquivo é um snapshot com data: um débito pago/cancelado depois
        disso continuaria aparecendo como em aberto. Por isso o status nunca
        é VERIFICADO e a OBSERVAÇÃO não aceita estas linhas.
        """
        digitos = "".join(
            caracter for caracter in documento if caracter.isdigit()
        )
        antigos = _debitos_relatorio_antigo().get(digitos, [])
        if not antigos:
            return False
        novos = [
            {
                "Nome": nome, "CPF": documento,
                "Tipo de Débito": debito["tipo_debito"],
                "Competência": debito["competencia"],
                "Valor (R$)": f"{debito['valor']:.2f}".replace(".", ","),
                "Status": STATUS_PENDENTE,
                # A origem fica no STATUS e na flag _pendente_confirmacao: a coluna
                # Erro permanece limpa (não é uma falha de consulta).
                "Erro": "",
                "CIF": debito["cif"],
                "ID Dívida": debito["id_divida"],
                "Vencimento": debito["vencimento"],
                "Situação (site)": str(debito.get("situacao", "") or "").strip(),
                "_pendente_confirmacao": True,
            }
            for debito in antigos
        ]
        self._publicar(documento, novos, indice)
        return True

    def _linhas_fallback_entrada(
        self, registro: dict, mensagem: str
    ) -> list[dict]:
        """Recupera a linha original da planilha quando a consulta falhou.

        A falha continua registrada na coluna Erro, mas a classificação já
        aprovada (COBRÁVEL / NÃO COBRAR) não é rebaixada só porque a sessão
        expirou ou o site não respondeu. Sem isso, um problema de sessão
        desfaz todo o trabalho de classificação anterior.
        """
        tipo = _campo(registro, "tipo de debito")
        competencia = _campo(registro, "competencia")
        if not tipo or not competencia:
            # Sem tipo/competência úteis na entrada, o snapshot legado decide.
            return []
        return [{
            "Nome": _campo(registro, "nome"),
            "CPF": (
                _campo(registro, "cpf") or _campo(registro, "cnpj")
                or _campo(registro, "documento") or _campo(registro, "cif")
            ),
            "Tipo de Débito": tipo,
            "Competência": competencia,
            "Valor (R$)": _campo(registro, "valor") or "0,00",
            # A consulta não terminou: nunca manter COBRÁVEL aqui.
            "Status": STATUS_ERRO_CONSULTA,
            "Erro": mensagem_amigavel_erro(mensagem),
            "CIF": _campo(registro, "cif"),
            "ID Dívida": _campo(registro, "id divida"),
            "Vencimento": _campo(registro, "vencimento"),
            "Situação (site)": (
                _campo(registro, "situacao (site)")
                or _campo(registro, "situacao site")
                or _campo(registro, "situacao")
            ),
            "MOTIVO": _campo(registro, "motivo"),
        }]

    def _publicar_falha(
        self, documento: str, nome: str, competencia: str, indice: int,
        mensagem: str, registro: dict | None = None,
    ) -> bool:
        """Publica a classificação anterior em vez de um '-' indeterminado."""
        if registro is not None:
            linhas = self._linhas_fallback_entrada(registro, mensagem)
            if linhas:
                self._publicar(documento, linhas, indice)
                return True
        return self._publicar_fallback_antigo(
            documento, nome, competencia, indice, mensagem
        )

    def _sem_sessao_para_continuar(self, mensagem: str) -> bool:
        """True se a falha é de login e não há sessão para recuperar."""
        texto = str(mensagem or "").lower()
        if not any(indicio in texto for indicio in (
            "sessão", "sessao", "autentica", "token", "home",
            "sem contribuinte", "404",
        )):
            return False
        return not self.cookie and not _cdp_ativo()

    def _falha_rede_critica(self, mensagem: str) -> bool:
        """Para a rodada se o site estiver fora há vários documentos seguidos."""
        texto = str(mensagem).lower()
        indicadores = (
            "connection refused", "err_connection", "address unreachable",
            "err_address", "err_network", "timeout", "timed out",
            "não foi possível acessar", "nao foi possivel acessar",
            "net::",
        )
        if not any(indicio in texto for indicio in indicadores):
            self._falhas_rede = 0
            return False
        self._falhas_rede += 1
        if self._falhas_rede >= 3:
            self._parar = True
            self.aviso.emit(
                "Site inalcançável em 3 documentos seguidos — encerrando para "
                "não perder tempo. Verifique a conexão e o login no navegador; "
                "o progresso foi salvo e a consulta pode ser retomada."
            )
            return True
        return False

    def run(self) -> None:
        total = len(self.registros)
        ok = erros = processados = 0

        # Sem sessão autenticada o site redireciona para a home e TODAS as
        # consultas falham. Encerrar antes do loop evita rebaixar débitos
        # COBRÁVEL para INDETERMINADO e sobrescrever um relatório válido.
        if not self.cookie and not _cdp_ativo():
            aviso = (
                "Sem sessão válida: entre no sistema no navegador e use "
                "'Abrir Brave com depuração (9222)' ou 'Atualizar cookies' "
                "na aba Ambiente antes de consultar. Nenhum resultado foi alterado."
            )
            self.aviso.emit(aviso)
            self.finalizado.emit("Consulta cancelada: " + aviso)
            return

        for indice, registro in enumerate(self.registros, start=1):
            if self._parar:
                break
            # Retomada: pular os registros já processados numa execução anterior.
            if indice <= self.indice_inicial:
                continue

            # Índice real na planilha de entrada: a revisão trabalha com uma
            # lista filtrada, então "indice" não corresponde à linha original.
            indice_real = int(_campo(registro, "_indice_original") or indice)

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
            self.documento_iniciado.emit(str(documento).strip())
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

            # CPF/CNPJ com dígito verificador errado só gera erro-page no site
            # (13445, 13446, 404...). Não adianta repetir: marca como não
            # verificado e segue para o próximo, sem desperdiçar consultas.
            if not _documento_valido(documento, tipo):
                digitos = "".join(
                    caracter for caracter in documento if caracter.isdigit()
                )
                antigos = _debitos_relatorio_antigo().get(digitos, [])
                if antigos:
                    # CPF fictício do legado: os débitos vêm do relatório antigo.
                    novos = [
                        {
                            "Nome": nome, "CPF": documento,
                            "Tipo de Débito": debito["tipo_debito"],
                            "Competência": debito["competencia"],
                            "Valor (R$)": f"{debito['valor']:.2f}".replace(".", ","),
                            "Status": STATUS_PENDENTE,
                            "Erro": "",
                            "CIF": debito["cif"],
                            "ID Dívida": debito["id_divida"],
                            "Vencimento": debito["vencimento"],
                            "Situação (site)": str(
                                debito.get("situacao", "") or ""
                            ).strip(),
                            "_pendente_confirmacao": True,
                        }
                        for debito in antigos
                    ]
                else:
                    novos = [{
                        "Nome": nome, "CPF": documento,
                        "Tipo de Débito": "-",
                        "Competência": competencia_original or "-",
                        "Valor (R$)": "0,00",
                        "Status": "DOCUMENTO INVÁLIDO",
                        "Erro": (
                            "CPF/CNPJ inválido (dígito verificador) — "
                            "conferir a planilha."
                        ),
                        "Situação (site)": "",
                    }]
                self._publicar(documento, novos, indice_real)
                if self.arquivo_progresso is not None:
                    salvar_progresso(
                        self.arquivo_progresso, indice_real, total, self.resultados,
                        em_andamento=False, indice_atual=indice_real,
                        entrada=self.entrada, aba=self.aba,
                    )
                continue

            self.progresso.emit(
                int((indice - 1) * 100 / total), f"Consultando {indice}/{total}: {rotulo}"
            )
            # Avisa no arquivo de progresso que esta consulta COMEÇOU mas ainda
            # não terminou. Se o programa for fechado agora, a retomada refaz
            # este registro em vez de considerá-lo concluído.
            if self.arquivo_progresso is not None:
                salvar_progresso(
                    self.arquivo_progresso, indice_real - 1, total, self.resultados,
                    em_andamento=True, indice_atual=indice_real,
                    entrada=self.entrada, aba=self.aba,
                )

            try:
                sem_confirmacao: list[bool] = []
                if self.sem_navegador:
                    with _SESSAO_SITE_LOCK:
                        debitos = consultar_debito_http(
                            documento, nome, self.cookie, self.cookie_sessao,
                            self.timeout, tipo,
                            deve_parar=self.parar_solicitado,
                            situacoes=(
                                SITUACOES_VERIFICACAO_ESSENCIAIS
                                if self.incluir_pagos else None
                            ),
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
                            with _SESSAO_SITE_LOCK:
                                debitos = _executar_com_playwright(
                                    consultar_debito,
                                    documento, nome, self.cookie, self.timeout,
                                    tipo, self.cookie_sessao,
                                    deve_parar=self.parar_solicitado,
                                    sem_confirmacao=sem_confirmacao,
                                    incluir_pagos=self.incluir_pagos,
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
                            if (
                                "sessão expirada" in mensagem
                                or "sessao expirada" in mensagem
                            ):
                                resumo = (
                                    "sessão expirada — renovando cookies do "
                                    "navegador e refazendo a consulta"
                                )
                            else:
                                resumo = str(exc)
                            self.aviso.emit(
                                f"{documento}: {resumo} ({tentativa + 2}/3)."
                            )
                            self.progresso.emit(
                                max(0, int((indice - 1) * 100 / total)),
                                f"Refazendo {documento} com sessão nova "
                                f"({tentativa + 2}/3)",
                            )
                            self.msleep(1000)
                if not debitos:
                    # HTTP é experimental e, no Playwright, uma resposta sem
                    # confirmação do documento NÃO pode virar "sem débitos"
                    # confirmado. Nesses casos o item volta para o Revisar.
                    nao_confirmado = bool(sem_confirmacao) or self.sem_navegador
                    descricao = (
                        "Não foi possível verificar débitos" if nao_confirmado
                        else "Sem débitos verificados"
                    )
                    novos = [{
                        "Nome": nome, "CPF": documento,
                        "Tipo de Débito": descricao, "Competência": competencia_original or "-",
                        "Valor (R$)": "0,00",
                        "Status": (
                            STATUS_ERRO_CONSULTA if nao_confirmado else "VERIFICADO"
                        ),
                        "Erro": (
                            "Resultado sem confirmação suficiente — usar Revisar."
                            if nao_confirmado else ""
                        ),
                        "CIF": "-", "ID Dívida": "-", "Vencimento": "-",
                        "Situação (site)": (
                            "SEM REGISTROS" if not nao_confirmado else ""
                        ),
                    }]
                else:
                    novos = []
                    descartados = 0
                    for debito in debitos:
                        item = {
                            "Nome": nome, "CPF": documento,
                            "Tipo de Débito": _tipo_debito_com_migracao(
                                debito["tipo_debito"], registro
                            ),
                            "Competência": _competencia_resultado(
                                competencia_original, debito, self.modo_revisao
                            ),
                            "_competencia_site": debito["competencia"],
                            "Valor (R$)": f"{debito['valor']:.2f}".replace(".", ","),
                            "Status": "VERIFICADO", "Erro": "",
                            "CIF": debito.get("cif", ""),
                            "ID Dívida": debito.get("id_divida", ""),
                            "Vencimento": debito.get("vencimento", ""),
                            "Situação (site)": str(
                                debito.get("situacao", "") or ""
                            ).strip(),
                        }
                        if self.incluir_pagos:
                            # Consulta com TODAS as situações: classifica cada
                            # linha e mantém somente as dívidas em aberto.
                            # PAGO, PAGO MANUAL, canceladas etc. ficam fora da
                            # cobrança mesmo tendo aparecido no site.
                            pago, cobravel = classificar_situacao(
                                debito.get("situacao", ""),
                                float(debito.get("pago", 0) or 0),
                                float(debito.get("valor", 0) or 0),
                            )
                            item["PAGO?"], item["COBRÁVEL?"] = pago, cobravel
                            if cobravel != "SIM":
                                descartados += 1
                                continue
                        novos.append(item)
                    if descartados:
                        self.aviso.emit(
                            f"{documento}: {descartados} débito(s) pago(s) ou "
                            "encerrado(s) não entram na cobrança."
                        )
                self._publicar(documento, novos, indice_real)
                ok += 1
                self._falhas_rede = 0
            except ConsultaInterrompida:
                # O usuário clicou em Parar no meio da consulta: não publica
                # resultado para este item e encerra o loop imediatamente.
                break
            except (ConsultaNaoVerificada, ConsultaHttpNaoVerificada) as exc:
                if not self._publicar_falha(
                    documento, nome, competencia_original, indice_real,
                    str(exc), registro,
                ):
                    erros += 1
                    self._publicar(documento, [{
                        "Nome": nome, "CPF": documento,
                        "Tipo de Débito": "-",
                        "Competência": competencia_original or "-",
                        "Valor (R$)": "0,00",
                        "Status": STATUS_ERRO_CONSULTA,
                        "Erro": mensagem_amigavel_erro(exc),
                        "Situação (site)": "",
                    }], indice_real)
                if self._falha_rede_critica(str(exc)):
                    break
                if self._sem_sessao_para_continuar(str(exc)):
                    self.aviso.emit(
                        "Sessão expirada e não há login no navegador — "
                        "consulta encerrada para não gerar indeterminados. "
                        "Abra o Brave com depuração (9222) / atualize os "
                        "cookies e use 'Revisar' para refazer estes débitos."
                    )
                    break
            except Exception as exc:  # noqa: BLE001
                if not self._publicar_falha(
                    documento, nome, competencia_original, indice_real,
                    str(exc), registro,
                ):
                    erros += 1
                    self._publicar(documento, [{
                        "Nome": nome, "CPF": documento,
                        "Tipo de Débito": "-",
                        "Competência": competencia_original or "-",
                        "Valor (R$)": "0,00",
                        "Status": STATUS_ERRO_CONSULTA,
                        "Erro": mensagem_amigavel_erro(exc),
                        "Situação (site)": "",
                    }], indice_real)
                if self._falha_rede_critica(str(exc)):
                    break
                if self._sem_sessao_para_continuar(str(exc)):
                    self.aviso.emit(
                        "Sessão expirada e não há login no navegador — "
                        "consulta encerrada para não gerar indeterminados. "
                        "Abra o Brave com depuração (9222) / atualize os "
                        "cookies e use 'Revisar' para refazer estes débitos."
                    )
                    break

            processados = indice_real
            if self.arquivo_progresso is not None:
                salvar_progresso(
                    self.arquivo_progresso, indice_real, total, self.resultados,
                    em_andamento=False, indice_atual=indice_real,
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

        # Gravação final forçada: o salvamento comum é espaçado no tempo e
        # poderia deixar de registrar o último item concluído.
        if self.arquivo_progresso is not None:
            salvar_progresso(
                self.arquivo_progresso, processados, total, self.resultados,
                em_andamento=bool(self._parar), indice_atual=processados,
                entrada=self.entrada, aba=self.aba, forcado=True,
            )

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


# O Playwright sempre avisa isso no Linux Mint (ele só considera "oficial"
# o Ubuntu). É um aviso inofensivo: a build do Ubuntu 24.04 é a correta.
AVISO_PLAYWRIGHT_MINT = (
    "Chromium: Linux Mint usa a build compatível do Ubuntu 24.04 "
    "(aviso do Playwright — é normal e o navegador funciona)."
)


def _eh_aviso_fallback_playwright(texto: str) -> bool:
    return (
        "BEWARE: your OS is not officially supported by Playwright" in texto
    )


class ComandoWorker(QThread):
    """Executa comandos (pip/playwright) em segundo plano, sem travar a GUI."""

    saida = pyqtSignal(str)
    finalizado = pyqtSignal(bool, str)

    def __init__(self, comandos: list[list[str]], diretorio: str | None = None,
                 parent=None):
        super().__init__(parent)
        self.comandos = comandos
        self.diretorio = diretorio
        self._aviso_playwright_emitido = False

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
                    texto = linha.rstrip()
                    if _eh_aviso_fallback_playwright(texto):
                        # Mostra uma explicação única, em português, em vez do
                        # aviso em inglês que parece um erro.
                        if not self._aviso_playwright_emitido:
                            self._aviso_playwright_emitido = True
                            self.saida.emit(AVISO_PLAYWRIGHT_MINT)
                        continue
                    self.saida.emit(texto)
            codigo = processo.wait()
            if codigo != 0:
                self.finalizado.emit(
                    False,
                    f"Comando falhou (código {codigo}): {' '.join(comando)}",
                )
                return
        self.finalizado.emit(True, "Instalação concluída com sucesso.")


class BraveWorker(QThread):
    """Abre um navegador Chromium com depuração (CDP) para reusar a sessão."""

    saida = pyqtSignal(str)
    finalizado = pyqtSignal(bool, str)

    def __init__(self, fechar: bool = False, parent=None):
        super().__init__(parent)
        self.fechar = fechar

    @staticmethod
    def _caminho_brave() -> str | None:
        """Localiza qualquer navegador Chromium instalado."""
        apps = (
            "Brave Browser",
            "Google Chrome",
            "Microsoft Edge",
            "Vivaldi",
            "Chromium",
        )
        candidatos: list[str] = []
        for app in apps:
            candidatos.extend((
                str(Path.home() / f"Applications/{app}.app/Contents/MacOS/{app}"),
                f"/Applications/{app}.app/Contents/MacOS/{app}",
            ))
        for candidato in candidatos:
            if Path(candidato).exists():
                return candidato
        for nome in ("brave-browser", "brave", "google-chrome", "chrome",
                     "microsoft-edge", "vivaldi", "chromium", "chromium-browser"):
            encontrado = shutil.which(nome)
            if encontrado:
                return encontrado
        return None

    @staticmethod
    def _nome_app(binario: str) -> str:
        """Nome do app (para fechar via osascript) a partir do executável."""
        return Path(binario).name

    def run(self) -> None:
        if _cdp_ativo():
            self.finalizado.emit(
                True, "Navegador já está com a depuração ativa na porta 9222."
            )
            return

        binario = self._caminho_brave()
        if not binario:
            self.finalizado.emit(
                False,
                "Nenhum navegador Chromium encontrado (Brave, Chrome, Edge, "
                "Vivaldi ou Chromium). Instale um deles em /Applications.",
            )
            return
        nome_app = self._nome_app(binario)

        if self.fechar:
            self.saida.emit(f"Fechando {nome_app} para reabri-lo com depuração...")
            try:
                if sys.platform == "darwin":
                    subprocess.run(
                        ["osascript", "-e", f'tell application "{nome_app}" to quit'],
                        capture_output=True, text=True, timeout=15,
                    )
                else:
                    subprocess.run(
                        ["pkill", "-x", nome_app],
                        capture_output=True, timeout=10,
                    )
            except Exception as exc:  # noqa: BLE001
                self.saida.emit(f" Não foi possível fechar automaticamente: {exc}")
            self.msleep(1500)

        porta = CDP_URL.rsplit(":", 1)[-1]
        self.saida.emit(f"Abrindo {nome_app} com depuração na porta {porta}...")
        try:
            subprocess.Popen(
                [binario, f"--remote-debugging-port={porta}", HOME_URL],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
        except Exception as exc:  # noqa: BLE001
            self.finalizado.emit(False, f"Não foi possível abrir o navegador: {exc}")
            return

        # Aguarda até 30 s pelo endpoint de depuração.
        for _ in range(60):
            if _cdp_ativo():
                self.finalizado.emit(
                    True, f"{nome_app} aberto e conectado (porta {porta})."
                )
                return
            self.msleep(500)
        self.finalizado.emit(
            False,
            f"O {nome_app} abriu, mas a porta {porta} não respondeu. "
            f"Feche todas as janelas do {nome_app} e tente novamente.",
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
        QLineEdit, QComboBox, QPlainTextEdit, QDoubleSpinBox, QSpinBox {{
            background: {p["campo_fundo"]};
            color: {p["texto"]};
            border: 1px solid {p["borda"]};
            border-radius: 6px;
            padding: 6px 8px;
            selection-background-color: {p["selecao"]};
        }}
        QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus, QDoubleSpinBox:focus, QSpinBox:focus {{
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


class SalvarSaidasWorker(QThread):
    """Grava a OBSERVAÇÃO e a planilha Resultado fora da thread da interface.

    Gravar dezenas de milhares de linhas com openpyxl leva minutos; fazendo
    isso em _ao_finalizar (thread da GUI) o programa parecia travar ao parar.
    """

    concluido = pyqtSignal(str)

    def __init__(self, entrada: Path, resultados: list[dict],
                 destino_observacao: Path, destino_resultado: Path,
                 parent=None):
        super().__init__(parent)
        self.entrada = entrada
        self.resultados = list(resultados)
        self.destino_observacao = destino_observacao
        self.destino_resultado = destino_resultado

    def run(self) -> None:
        try:
            gravar_relatorio_com_observacao(
                self.entrada,
                mapa_observacoes(self.resultados),
                self.destino_observacao,
            )
            gravar_planilha(str(self.destino_resultado), self.resultados)
            self.concluido.emit(
                f"Relatorios salvos: {self.destino_observacao.name} e "
                f"{self.destino_resultado.name}"
            )
        except Exception as exc:  # noqa: BLE001
            self.concluido.emit(f"Erro ao salvar relatorios: {exc}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(
            "Consulta de Débitos — todos os débitos na OBSERVAÇÃO"
        )
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
        self._total_revisao = 0
        self._revisao_paralela_ativa = False
        self._documentos_em_revisao: set[str] = set()
        self.worker: ConsultaWorker | None = None
        self.worker_pagos: VerificarPagosWorker | None = None
        self.worker_cifs: VerificarCifsWorker | None = None
        self._fluxo_completo = False
        self._salvador_saidas: SalvarSaidasWorker | None = None
        self._salvamento_pendente = False
        # Destino dos relatórios automáticos: gerado uma vez por execução,
        # com data/hora no nome, para não sobrescrever relatórios anteriores.
        self._destino_auto: Path | None = None
        self._destino_auto_parcial: Path | None = None
        self.config_worker: QThread | None = None
        self._config = QSettings("ConsultaDebitos", "ConsultaDebitos")
        self.gerar_umalinha = _config_bool(
            self._config, "gerar_umalinha", True
        )

        self._montar_interface()
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage(
            "Pronto — planilha base carregada. Escolha a aba e execute."
        )

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
        # Pisca o botão Parar enquanto o aviso do ESC estiver ativo.
        self._esc_pisca_aceso = True
        self._timer_pisca_esc = QTimer(self)
        self._timer_pisca_esc.setInterval(350)
        self._timer_pisca_esc.timeout.connect(self._piscar_aviso_esc)

        # Renovação preventiva: relê o token do Brave a cada 2 minutos e
        # avisa se ele estiver perto de vencer.
        self._timer_token = QTimer(self)
        self._timer_token.setInterval(120000)
        self._timer_token.timeout.connect(self._renovar_token_periodicamente)
        self._timer_token.start()

        # Copia automaticamente os relatórios gerados para a pasta
        # compartilhada do VirtualBox (departamento d.a), por departamento.
        self._aviso_compartilhamento = False
        self._timer_compartilhamento = QTimer(self)
        self._timer_compartilhamento.setInterval(5000)
        self._timer_compartilhamento.timeout.connect(
            self._sincronizar_compartilhamento
        )
        self._timer_compartilhamento.start()
        QTimer.singleShot(1500, self._sincronizar_compartilhamento)

        self._log('Aplicação iniciada. Dica: pressione "ESC" 2x para parar as consultas.')

        # Reabre automaticamente a última planilha usada, para facilitar a
        # retomada de uma consulta interrompida.
        self.botao_reabrir.setEnabled(bool(self._config.value("ultima_planilha", "")))
        self.reabrir_ultima_planilha(silent=True)

    def _sincronizar_compartilhamento(self, manual: bool = False) -> None:
        """Copia os relatórios para a pasta 'departamento d.a' (VirtualBox)."""
        copias, erro = sincronizar_relatorios(_pasta_relatorios())
        if copias:
            mensagem = (
                f"Compartilhamento: {len(copias)} relatório(s) copiado(s) — "
                + ", ".join(copias)
            )
            self._log(mensagem)
            self.statusBar().showMessage(mensagem)
        if erro and not self._aviso_compartilhamento:
            self._aviso_compartilhamento = True
            self._log(f"Compartilhamento — aviso: {erro}")
        if manual and hasattr(self, "campo_ambiente"):
            if copias:
                texto = "Relatórios copiados:\n" + "\n".join(copias)
                if erro:
                    texto += f"\n\nAviso: {erro}"
            elif erro:
                texto = f"Não foi possível copiar: {erro}"
            else:
                texto = "Nenhum relatório novo para copiar."
            self.campo_ambiente.append(texto)
            QMessageBox.information(self, "Compartilhamento", texto)

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
            self._esc_pisca_aceso = True
            self._mostrar_aviso_esc(True, True)
            self._timer_esc.start()
            self._timer_pisca_esc.start()
            self.statusBar().showMessage(
                "Pressione ESC novamente para parar as consultas."
            )
            self._log("ESC pressionado — pressione novamente para parar as consultas.")
            return

        # Segundo aperto: interrompe e restaura o visual normal.
        self._resetar_aviso_esc()
        self.parar()

    def _mostrar_aviso_esc(self, ativo: bool, aceso: bool = True) -> None:
        """Realça (ou restaura) os controles de parada no aviso do ESC."""
        if ativo:
            if aceso:
                estilo_botao = (
                    "QPushButton { background: #B71C1C; color: #FFFFFF; "
                    "font-weight: 600; border: 1px solid #7F0000; "
                    "border-radius: 6px; padding: 4px 10px; }"
                )
            else:
                estilo_botao = (
                    "QPushButton { background: #FF8C00; color: #000000; "
                    "font-weight: 600; border: 1px solid #B36300; "
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

    def _piscar_aviso_esc(self) -> None:
        """Alterna a cor do botão Parar enquanto o aviso do ESC está ativo."""
        self._esc_pisca_aceso = not getattr(self, "_esc_pisca_aceso", True)
        self._mostrar_aviso_esc(True, self._esc_pisca_aceso)

    def _resetar_aviso_esc(self) -> None:
        """Cancela o aviso do ESC e devolve os controles ao visual normal."""
        self._esc_armado = False
        self._timer_esc.stop()
        self._timer_pisca_esc.stop()
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
        self._atualizar_painel_resultados()
        if not hasattr(self, "botao_revisar"):
            return
        rodando = (
            (self.worker is not None and self.worker.isRunning())
            or (self.worker_revisao is not None and self.worker_revisao.isRunning())
            or (self.worker_pagos is not None and self.worker_pagos.isRunning())
            or (self.worker_cifs is not None and self.worker_cifs.isRunning())
            or self._revisao_paralela_ativa
        )
        habilitado = not rodando and (
            bool(self.registros) or bool(self.resultados)
        )
        self.botao_revisar.setEnabled(habilitado)
        if hasattr(self, "botao_revisao_completa"):
            self.botao_revisao_completa.setEnabled(habilitado)
        self._atualizar_botao_pagos()

    def _atualizar_botao_pagos(self) -> None:
        """Habilita 'Verificar pagos' somente com planilha e sem consulta ativa."""
        botao = getattr(self, "botao_verificar_pagos", None)
        if botao is None:
            return
        rodando = (
            (self.worker is not None and self.worker.isRunning())
            or (self.worker_revisao is not None and self.worker_revisao.isRunning())
            or (self.worker_pagos is not None and self.worker_pagos.isRunning())
            or (self.worker_cifs is not None and self.worker_cifs.isRunning())
            or self._revisao_paralela_ativa
        )
        habilitado = not rodando and bool(self.caminho_entrada)
        botao.setEnabled(habilitado)
        botao_completo = getattr(self, "botao_consulta_completa", None)
        if botao_completo is not None:
            botao_completo.setEnabled(habilitado)
        botao_cifs = getattr(self, "botao_verificar_cifs", None)
        if botao_cifs is not None:
            botao_cifs.setEnabled(habilitado)

    def consulta_completa(self) -> None:
        """Revisão completa + classificação de pagos + validação de CIFs."""
        if self.worker is not None and self.worker.isRunning():
            return
        if self.worker_pagos is not None and self.worker_pagos.isRunning():
            return
        if self.worker_cifs is not None and self.worker_cifs.isRunning():
            return
        if not self._pergunta_sim_nao(
            "Consulta completa",
            "Vai executar em sequência:\n"
            "1) revisão completa de todos os documentos da aba,\n"
            "2) classificação de pagos,\n"
            "3) validação dos CIFs.\n\n"
            "Pode demorar. Deseja continuar?",
        ):
            return
        self._fluxo_completo = True
        self._log("Consulta completa iniciada (etapa 1/3: revisão completa).")
        self.revisao_completa()

    def verificar_pagos(self, auto: bool = False) -> None:
        """Cruza os débitos da OBSERVAÇÃO com o site para achar os cobráveis."""
        if self.worker_pagos is not None and self.worker_pagos.isRunning():
            return
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.information(
                self, "Verificar pagos",
                "Aguarde a consulta em andamento terminar antes de verificar.",
            )
            return

        # A OBSERVAÇÃO já preenchida fica no relatório de 1 linha; se ele não
        # existir, usa a planilha de entrada carregada.
        origem = (
            self._arquivo_destino_observacao()
            if self._usa_observacao()
            else (self.caminho_entrada or PLANILHA_BASE)
        )
        if not Path(origem).exists():
            origem = self.caminho_entrada or PLANILHA_BASE
        aba = (
            self.combo_aba.currentText()
            if self.combo_aba is not None and self.combo_aba.isEnabled()
            else ""
        )
        try:
            # Os resultados já obtidos (inclusive os débitos do snapshot
            # legado) entram no módulo novo para análise/classificação.
            registros = registros_para_analise(self.resultados)
            observacao = registros_com_observacao(origem, aba or None)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self, "Verificar pagos",
                f"Não foi possível ler os débitos para análise:\n{exc}",
            )
            return
        # Junta o que veio dos resultados com o que está na OBSERVAÇÃO, sem
        # duplicar tipo+competência do mesmo documento.
        por_documento = {r["Documento"]: r for r in registros}
        for registro_obs in observacao:
            documento = registro_obs["Documento"]
            if documento not in por_documento:
                por_documento[documento] = registro_obs
                continue
            destino = por_documento[documento]
            chaves = {
                (chave_tipo(d["tipo"]), chave_competencia(d["competencia"]))
                for d in destino["Debitos"]
            }
            for debito in registro_obs["Debitos"]:
                chave = (
                    chave_tipo(debito["tipo"]),
                    chave_competencia(debito["competencia"]),
                )
                if chave not in chaves:
                    chaves.add(chave)
                    destino["Debitos"].append(debito)
        registros = list(por_documento.values())
        if not registros:
            QMessageBox.information(
                self, "Verificar pagos",
                "Nenhum débito encontrado nos resultados nem na coluna "
                "OBSERVAÇÃO (débitos) desta aba. Execute a consulta primeiro.",
            )
            return

        total_debitos = sum(len(r.get("Debitos", [])) for r in registros)
        if not auto and not self._pergunta_sim_nao(
            "Verificar pagos",
            f"Vamos cruzar {total_debitos} débito(s) de {len(registros)} "
            f"documento(s) da aba '{aba or 'todas'}' com o site.\n\n"
            "O site será consultado com todos os filtros de situação. "
            "Deseja continuar?",
        ):
            return

        destino, detalhe = caminhos_saida(
            self.caminho_entrada, _pasta_relatorios(), PLANILHA_BASE.stem
        )
        self._carregar_cookie_brave(silent=True)
        self.barra_progresso.setValue(0)
        self.label_status.setText("Verificando pagos...")
        self.botao_executar.setEnabled(False)
        self.botao_parar.setEnabled(True)
        self.botao_parar_resultados.setEnabled(True)
        self.botao_escolher.setEnabled(False)
        self.botao_salvar.setEnabled(False)
        self.botao_revisar.setEnabled(False)
        if hasattr(self, "botao_revisao_completa"):
            self.botao_revisao_completa.setEnabled(False)
        self.botao_verificar_pagos.setEnabled(False)
        self.botao_verificar_cifs.setEnabled(False)

        self.worker_pagos = VerificarPagosWorker(
            registros,
            consultar_debito,
            self.campo_cookie.text(),
            self.cookie_sessao,
            destino=destino,
            detalhe=detalhe,
            origem=origem,
            pausa=self.campo_velocidade.value(),
            variar_pausa=self.campo_variar_pausa.isChecked(),
            controle=self.controle_consulta,
            parent=self,
        )
        self.worker_pagos.progresso.connect(self._ao_progresso)
        self.worker_pagos.aviso.connect(self._log)
        self.worker_pagos.documento_concluido.connect(self._ao_documento_pago)
        self.worker_pagos.finalizado.connect(self._ao_finalizar_pagos)
        self.worker_pagos.finished.connect(self._limpar_worker_pagos)
        self.worker_pagos.start()
        self._log(
            f"Verificação de cobrabilidade iniciada: {total_debitos} débito(s) "
            f"de {len(registros)} documento(s) da aba '{aba or 'todas'}'."
        )

    def _ao_documento_pago(self, informacao: dict) -> None:
        itens = informacao.get("Itens", [])
        cobraveis = sum(
            1 for item in itens if item.get("COBRÁVEL?") == "SIM"
        )
        self._log(
            f"{informacao.get('Documento', '')}: {len(itens)} débito(s) "
            f"cruzado(s), {cobraveis} cobrável(is)."
        )

    def _ao_finalizar_pagos(self, mensagem: str) -> None:
        self._resetar_aviso_esc()
        if self._fluxo_completo:
            self._fluxo_completo = False
            if "Concluído" in mensagem:
                self._log("Consulta completa (3/3): validando CIFs...")
                QTimer.singleShot(50, lambda: self.verificar_cifs(auto=True))
        self._log(mensagem)
        self.label_status.setText(
            "Concluído" if "Concluído" in mensagem else mensagem
        )
        self.botao_executar.setEnabled(True)
        self.botao_parar.setEnabled(False)
        self.botao_parar_resultados.setEnabled(False)
        self.botao_escolher.setEnabled(True)
        self.botao_salvar.setEnabled(bool(self.resultados))
        self.barra_progresso.setValue(100)
        self.statusBar().showMessage(mensagem)

    def _limpar_worker_pagos(self) -> None:
        self.worker_pagos = None
        self._atualizar_botao_revisar()

    def _arquivo_resultado_classico(self) -> Path:
        """Planilha com um débito por linha, onde ficam os CIFs."""
        stem = self.caminho_entrada.stem if self.caminho_entrada else PLANILHA_BASE.stem
        return _pasta_relatorios() / f"{stem}_resultado.xlsx"

    def verificar_cifs(self, auto: bool = False) -> None:
        """Consulta o Cadastro Imobiliário (CPF/CNPJ) e valida os CIFs."""
        if self.worker_cifs is not None and self.worker_cifs.isRunning():
            return
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.information(
                self, "Verificar CIFs",
                "Aguarde a consulta em andamento terminar antes de verificar.",
            )
            return

        try:
            documentos = carregar_cifs(
                self.resultados, self._arquivo_resultado_classico()
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self, "Verificar CIFs",
                f"Não foi possível ler os CIFs dos resultados:\n{exc}",
            )
            return
        if not documentos:
            QMessageBox.information(
                self, "Verificar CIFs",
                "Nenhum CIF foi encontrado nos débitos consultados. "
                "Execute a consulta primeiro.",
            )
            return

        total_cifs = sum(len(cifs) for cifs in documentos.values())
        if not auto and not self._pergunta_sim_nao(
            "Verificar CIFs",
            f"Vamos consultar o Cadastro Imobiliário do IPTU para "
            f"{len(documentos)} CPF(s)/CNPJ(s) e verificar {total_cifs} CIF(s).\n\n"
            "Um CIF só será considerado REAL se aparecer no cadastro com "
            "endereço. Deseja continuar?",
        ):
            return

        destino, detalhe = caminhos_saida_cifs(
            self.caminho_entrada, _pasta_relatorios(), PLANILHA_BASE.stem
        )
        self._carregar_cookie_brave(silent=True)
        self.barra_progresso.setValue(0)
        self.label_status.setText("Verificando CIFs...")
        self.botao_executar.setEnabled(False)
        self.botao_parar.setEnabled(True)
        self.botao_parar_resultados.setEnabled(True)
        self.botao_escolher.setEnabled(False)
        self.botao_salvar.setEnabled(False)
        self.botao_revisar.setEnabled(False)
        if hasattr(self, "botao_revisao_completa"):
            self.botao_revisao_completa.setEnabled(False)
        self.botao_verificar_pagos.setEnabled(False)
        self.botao_verificar_cifs.setEnabled(False)

        self.worker_cifs = VerificarCifsWorker(
            documentos,
            consultar_imoveis,
            self.campo_cookie.text(),
            self.cookie_sessao,
            destino=destino,
            detalhe=detalhe,
            origem=self._arquivo_resultado_classico(),
            pausa=self.campo_velocidade.value(),
            variar_pausa=self.campo_variar_pausa.isChecked(),
            controle=self.controle_consulta,
            parent=self,
        )
        self.worker_cifs.progresso.connect(self._ao_progresso)
        self.worker_cifs.aviso.connect(self._log)
        self.worker_cifs.finalizado.connect(self._ao_finalizar_cifs)
        self.worker_cifs.finished.connect(self._limpar_worker_cifs)
        self.worker_cifs.start()
        self._log(
            f"Verificação de CIFs iniciada: {total_cifs} CIF(s) de "
            f"{len(documentos)} documento(s)."
        )

    def _ao_finalizar_cifs(self, mensagem: str) -> None:
        self._resetar_aviso_esc()
        self._log(mensagem)
        self.label_status.setText(
            "Concluído" if "Concluído" in mensagem else mensagem
        )
        self.botao_executar.setEnabled(True)
        self.botao_parar.setEnabled(False)
        self.botao_parar_resultados.setEnabled(False)
        self.botao_escolher.setEnabled(True)
        self.botao_salvar.setEnabled(bool(self.resultados))
        self.barra_progresso.setValue(100)
        self.statusBar().showMessage(mensagem)

    def _limpar_worker_cifs(self) -> None:
        self.worker_cifs = None
        self._atualizar_botao_revisar()

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
        # Progressos antigos guardavam débitos do snapshot com texto de
        # "relatório antigo/sistema legado" na coluna Erro. Isso não é uma
        # falha de consulta: limpa o Erro e mantém o status de não confirmado
        # (a origem fica na flag _pendente_confirmacao, usada pelo módulo novo).
        for item in resultados:
            erro = str(item.get("Erro", "") or "").lower()
            item["Erro"] = mensagem_amigavel_erro(item.get("Erro", ""))
            status_antigo = str(item.get("Status", "") or "").strip().upper()
            if (
                "relatório antigo" in erro or "sistema legado" in erro
                or status_antigo in (
                    "SEM CONFIRMAÇÃO / PENDENTE DE ANÁLISE",
                    "LEGADO_NAO_CONFIRMADO",
                )
            ):
                item["Status"] = STATUS_PENDENTE
                item["Erro"] = ""
                item["_pendente_confirmacao"] = True
        # Repara índices de progressos antigos: cada CPF volta para a posição
        # real que ocupa na planilha de entrada.
        posicoes: dict[str, int] = {}
        for posicao, registro in enumerate(self.registros, start=1):
            documento = self._documento_registro(registro)
            if documento:
                posicoes.setdefault(documento, posicao)
        for item in resultados:
            documento = str(item.get("CPF", "")).strip()
            if documento in posicoes:
                item["_indice"] = posicoes[documento]
        self.resultados = _ordenar_resultados(resultados)
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

        # --- Aba Classificação e Verificação (primeira, no topo) ---
        aba_modulos = QWidget()
        modulos_layout = QVBoxLayout(aba_modulos)
        modulos_layout.addWidget(QLabel(
            "Classificação e verificação: cruze os débitos com o site e confira "
            "os CIFs. Selecione a aba desejada em Consulta antes de executar."
        ))
        legenda = QLabel(
            "<b>Legenda de alerta:</b> "
            "<span style='color:#1B5E20;'>■ VERDE — cobrável (confirmado)</span> &nbsp; "
            "<span style='color:#B26A00;'>■ LARANJA — indeterminado (analisar)</span> &nbsp; "
            "<span style='color:#9C6500;'>■ AMARELO — pendente de análise</span> &nbsp; "
            "<span style='color:#9C0006;'>■ VERMELHO — não cobrar (pago/cancelado/prescrito)</span>"
        )
        legenda.setWordWrap(True)
        modulos_layout.addWidget(legenda)

        acoes = QGroupBox("Ações")
        acoes_layout = QVBoxLayout(acoes)
        self.botao_consulta_completa = QPushButton("Consulta completa")
        self.botao_consulta_completa.setEnabled(False)
        self.botao_consulta_completa.setToolTip(
            "Executa em sequência: revisão completa (todas as situações), "
            "classificação de pagos e validação dos CIFs."
        )
        self.botao_consulta_completa.clicked.connect(self.consulta_completa)
        acoes_layout.addWidget(self.botao_consulta_completa)

        self.botao_verificar_pagos = QPushButton("Verificar pagos")
        self.botao_verificar_pagos.setEnabled(False)
        self.botao_verificar_pagos.setToolTip(
            "Cruza os débitos com o site e marca PAGO? e COBRÁVEL? "
            "(SIM / NÃO / NÃO CONFIRMADO) para cada dívida."
        )
        self.botao_verificar_pagos.clicked.connect(self.verificar_pagos)
        acoes_layout.addWidget(self.botao_verificar_pagos)

        self.botao_verificar_cifs = QPushButton("Verificar CIFs")
        self.botao_verificar_cifs.setEnabled(False)
        self.botao_verificar_cifs.setToolTip(
            "Consulta o Cadastro Imobiliário pelo CPF/CNPJ e marca CIF REAL? "
            "e ENDEREÇO (CIF) para cada imóvel."
        )
        self.botao_verificar_cifs.clicked.connect(self.verificar_cifs)
        acoes_layout.addWidget(self.botao_verificar_cifs)

        self.botao_umalinha = QPushButton()
        self.botao_umalinha.clicked.connect(self._alternar_gerar_umalinha)
        acoes_layout.addWidget(self.botao_umalinha)

        acoes.setMaximumWidth(560)
        modulos_layout.addWidget(acoes)

        instrucoes = QLabel(
            "Como usar:\n"
            "1. Na aba Consulta, escolha a planilha e a aba dos contribuintes.\n"
            "2. Clique em Consulta completa para revisar, classificar pagos e "
            "validar CIFs em sequência — ou use um botão individual.\n"
            "3. O resultado aparece na aba Resultados e os relatórios são "
            "gravados em relatorios/.\n\n"
            "Dica: mantenha o sistema aberto e logado no navegador com "
            "depuração (porta 9222) para maior estabilidade."
        )
        instrucoes.setWordWrap(True)
        modulos_layout.addWidget(instrucoes)
        modulos_layout.addStretch()

        # --- Aba Consulta ---
        # Com todos os controles (pausa, revisão paralela etc.), a aba pode
        # ficar mais alta que a janela: usa rolagem para nada ficar escondido.
        aba_consulta = QWidget()
        consulta_scroll = QScrollArea(aba_consulta)
        consulta_scroll.setWidgetResizable(True)
        consulta_conteudo = QWidget()
        consulta_layout = QVBoxLayout(consulta_conteudo)
        consulta_scroll.setWidget(consulta_conteudo)
        layout_aba_consulta = QVBoxLayout(aba_consulta)
        layout_aba_consulta.addWidget(consulta_scroll)

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
        self.botao_limpar_progresso = QPushButton("Limpar progresso")
        self.botao_limpar_progresso.setToolTip(
            "Apaga o progresso desta aba para começar do zero"
        )
        self.botao_limpar_progresso.clicked.connect(self.limpar_progresso)
        linha_arquivo.addWidget(self.campo_arquivo)
        linha_arquivo.addWidget(self.botao_escolher)
        linha_arquivo.addWidget(self.botao_reabrir)
        linha_arquivo.addWidget(self.botao_limpar_progresso)
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
        self.tabela_entrada.setMinimumHeight(200)
        self.tabela_entrada.setMaximumHeight(240)
        _configurar_tabela(
            self.tabela_entrada, ["Coluna"], zebra=True, scroll_sempre=True
        )
        entrada_layout.addWidget(self.tabela_entrada)
        consulta_layout.addWidget(grupo_entrada)

        grupo_sessao = QGroupBox("Sessão (leitura automática do navegador)")
        sessao_layout = QHBoxLayout(grupo_sessao)
        sessao_layout.addWidget(QLabel(f"Cookie ({COOKIE_NAME}):"))
        self.campo_cookie = QLineEdit()
        self.campo_cookie.setPlaceholderText("Cookie preenchido automaticamente do navegador")
        sessao_layout.addWidget(self.campo_cookie)
        sessao_layout.addWidget(QLabel("JSESSIONID:"))
        self.campo_sessao = QLineEdit()
        self.campo_sessao.setEchoMode(QLineEdit.EchoMode.Password)
        self.campo_sessao.setPlaceholderText(
            "Informe manualmente se a leitura automática falhar"
        )
        self.campo_sessao.setToolTip(
            "Sessão JSF do sistema. Necessária junto com o token JWT."
        )
        self.campo_sessao.textChanged.connect(self._ao_mudar_sessao)
        sessao_layout.addWidget(self.campo_sessao)
        self.campo_cookie.setEchoMode(QLineEdit.EchoMode.Password)
        self.campo_cookie.setToolTip("Token de sessão lido do navegador. Fica oculto por segurança.")
        self.botao_cookie_brave = QPushButton("Buscar no navegador")
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

        linha_inicio = QHBoxLayout()
        linha_inicio.addWidget(QLabel("Começar do registro:"))
        self.campo_registro_inicial = QSpinBox()
        self.campo_registro_inicial.setRange(1, 99999)
        self.campo_registro_inicial.setValue(1)
        self.campo_registro_inicial.setToolTip(
            "Para iniciar de um número específico. 1 = desde o início / "
            "retomada automática."
        )
        linha_inicio.addWidget(self.campo_registro_inicial)
        linha_inicio.addStretch()
        execucao_layout.addLayout(linha_inicio)

        self.label_dica_espaco = QLabel(
            'PRESSIONE "ESC" 2x PARA PARAR CONSULTA'
        )
        self.label_dica_espaco.setObjectName("dicaEspaco")
        execucao_layout.addWidget(self.label_dica_espaco)

        # O site é JSF/PrimeFaces: sem JavaScript não é possível nem ler o
        # formulário. A opção fica visível apenas como explicação.
        self.campo_sem_navegador = QCheckBox(
            "SEM navegador (indisponível — o site exige JavaScript)"
        )
        self.campo_sem_navegador.setChecked(False)
        self.campo_sem_navegador.setEnabled(False)
        self.campo_sem_navegador.setToolTip(
            "O sistema do ISS Digital usa JavaScript e não pode ser consultado "
            "por HTTP puro. A consulta sempre usa o navegador automático "
            "(Playwright) ou a sessão real do Chrome/Brave."
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
        self.campo_revisar_paralelo.toggled.connect(
            self._ao_alternar_revisao_paralela
        )

        # Controle de velocidade: pausa entre consultas (anti-detecção).
        linha_velocidade = QHBoxLayout()
        linha_velocidade.addWidget(QLabel("Pausa entre consultas:"))
        self.campo_velocidade = QDoubleSpinBox()
        self.campo_velocidade.setRange(0.0, 60.0)
        self.campo_velocidade.setSingleStep(0.5)
        self.campo_velocidade.setDecimals(1)
        self.campo_velocidade.setSuffix(" s")
        self.campo_velocidade.setMinimumWidth(100)
        self.campo_velocidade.setMinimumHeight(30)
        self.campo_velocidade.setAlignment(Qt.AlignmentFlag.AlignRight)
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
            _config_bool(self._config, "variar_pausa", False)
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
        self.label_revisao_paralela = QLabel("Revisão paralela: inativa")
        self.label_revisao_paralela.setObjectName("dicaEspaco")
        execucao_layout.addWidget(self.label_revisao_paralela)
        consulta_layout.addWidget(grupo_execucao)

        abas.addTab(aba_consulta, "Consulta")

        # --- Aba Resultados ---
        aba_resultados = QWidget()
        resultados_layout = QVBoxLayout(aba_resultados)

        painel = QGroupBox("Resumo por situação")
        painel_layout = QHBoxLayout(painel)
        self.labels_painel: dict[str, QLabel] = {}
        for chave, titulo in (
            ("TOTAL", "Total"),
            ("VERIFICADO", "VERIFICADO"),
            (STATUS_PENDENTE, "CLASSIFICADO — INDETERMINADO"),
            (STATUS_ERRO_CONSULTA, "ERRO DE CONSULTA"),
            ("DOCUMENTO INVÁLIDO", "DOCUMENTO INVÁLIDO"),
            ("NAO VERIFICADO", "NAO VERIFICADO"),
            ("ERRO", "ERRO"),
        ):
            rotulo = QLabel(f"{titulo}: 0")
            rotulo.setWordWrap(True)
            self.labels_painel[chave] = rotulo
            painel_layout.addWidget(rotulo)
        resultados_layout.addWidget(painel)

        self.tabela_resultados = QTableWidget()
        self.tabela_resultados.setMinimumHeight(220)
        _configurar_tabela(
            self.tabela_resultados,
            COLUNAS_SAIDA,
            larguras=[185, 150, 225, 105, 120, 270, 90, 100, 100, 165, None],
            scroll_sempre=True,
        )
        self.tabela_resultados.setItemDelegate(
            _DelegadoRevisao(self.tabela_resultados)
        )
        resultados_layout.addWidget(self.tabela_resultados)
        linha_resultados = QHBoxLayout()
        self.botao_revisar = QPushButton("Revisar")
        self.botao_revisar.setEnabled(False)
        if hasattr(self, "botao_revisao_completa"):
            self.botao_revisao_completa.setEnabled(False)
        self.botao_revisar.setToolTip(
            "Após parar: consulta somente o que ficou faltando. "
            "Depois de concluído: reconulta erros, NAO VERIFICADO, "
            "'Sem débitos' e linhas com CIF, ID Dívida ou Vencimento em branco"
        )
        self.botao_revisar.clicked.connect(self.revisar)
        linha_resultados.addWidget(self.botao_revisar)
        self.botao_revisao_completa = QPushButton("Revisão completa")
        self.botao_revisao_completa.setEnabled(False)
        self.botao_revisao_completa.setToolTip(
            "Reconsulta TODOS os CPFs/CNPJs da aba selecionada, inclusive os "
            "que já têm débito na OBSERVAÇÃO"
        )
        self.botao_revisao_completa.clicked.connect(self.revisao_completa)
        linha_resultados.addWidget(self.botao_revisao_completa)
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

        # 3ª etapa: classificação/verificação fica depois de Consulta e Resultados.
        abas.addTab(aba_modulos, "Classificação e Verificação")

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

        self.botao_abrir_brave = QPushButton("Abrir navegador com depuração (9222)")
        self.botao_abrir_brave.clicked.connect(self.abrir_brave)
        self.botao_abrir_sistema = QPushButton("Abrir sistema no navegador")
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

        self.botao_compartilhamento = QPushButton(
            "Copiar relatórios ao compartilhamento"
        )
        self.botao_compartilhamento.setToolTip(
            "Copia os relatórios para a raiz da pasta 'departamento d.a' "
            "do VirtualBox"
        )
        self.botao_compartilhamento.clicked.connect(
            lambda _checked=False: self._sincronizar_compartilhamento(manual=True)
        )
        self.botao_verificar = QPushButton("Verificar novamente")
        self.botao_verificar.clicked.connect(self._verificar_ambiente)
        linha_verificar = QHBoxLayout()
        linha_verificar.addWidget(self.botao_compartilhamento)
        linha_verificar.addWidget(self.botao_verificar)
        linha_verificar.addStretch()
        ambiente_layout.addLayout(linha_verificar)

        self.botoes_config = [
            self.botao_instalar_tudo, self.botao_instalar_deps,
            self.botao_instalar_navegador, self.botao_abrir_brave,
            self.botao_abrir_sistema, self.botao_atualizar_cookies,
            self.botao_compartilhamento,
        ]
        if sys.platform.startswith("linux"):
            self.botao_guest_additions = QPushButton(
                "Instalar Guest Additions (arrastar/soltar)"
            )
            self.botao_guest_additions.setToolTip(
                "Instala o VBoxClient — necessário para arrastar e soltar e "
                "para a área de transferência dentro da VM. Pede autorização "
                "(pkexec/sudo) e reinicia a sessão gráfica."
            )
            self.botao_guest_additions.clicked.connect(
                self.instalar_guest_additions
            )
            ambiente_layout.addWidget(self.botao_guest_additions)
            self.botoes_config.append(self.botao_guest_additions)
        abas.addTab(self.aba_ambiente, "Ambiente")
        abas.currentChanged.connect(self._ao_mudar_aba_principal)

        self._atualizar_botao_umalinha()

        # Tenta carregar automaticamente os cookies da sessão aberta no Brave.
        self.cookie_sessao = ""
        self._carregar_cookie_brave(silent=True)
        self._verificar_ambiente()

    # ---------- Ações ----------

    def _ao_mudar_sessao(self, texto: str) -> None:
        """Mantém a sessão usada nas consultas em sincronia com o campo."""
        self.cookie_sessao = str(texto or "").strip()

    def _carregar_cookie_brave(self, silent: bool = False) -> None:
        """Lê os cookies de sessão do Brave e preenche o campo de cookie."""
        try:
            cookies = buscar_cookies_no_brave()
        except Exception as exc:  # noqa: BLE001
            if not silent:
                QMessageBox.warning(self, "Cookie do navegador", str(exc))
            self._log(f"Não foi possível ler os cookies do navegador: {exc}")
            return
        self.campo_cookie.setText(cookies["__Secure-DA"])
        # O JSESSIONID é opcional: o site o cria ao abrir a home.
        sessao = cookies.get("JSESSIONID", "")
        self.campo_sessao.setText(sessao)
        self.cookie_sessao = sessao
        navegador = brave_cookie.ULTIMO_NAVEGADOR or "navegador Chromium"
        detalhe = "JWT + JSESSIONID" if sessao else "JWT (JSESSIONID será criado pelo site)"
        self._log(f"Cookies de sessão obtidos do {navegador} ({detalhe}).")
        aviso = _aviso_expiracao(cookies["__Secure-DA"])
        if aviso:
            self._log(f"ATENÇÃO: {aviso}")
            if not silent:
                self.statusBar().showMessage(aviso)
        if not silent and not aviso:
            self.statusBar().showMessage("Cookies do navegador atualizados.")

    def _renovar_token_periodicamente(self) -> None:
        """Relê o token do Brave periodicamente e avisa antes de expirar."""
        if self.worker is not None and self.worker.isRunning():
            return  # o worker já renova o token durante a consulta
        cookie_atual = self.campo_cookie.text()
        cookie, sessao = _renovar_token(cookie_atual, self.cookie_sessao)
        if cookie != cookie_atual or sessao != self.cookie_sessao:
            self.campo_cookie.setText(cookie)
            self.cookie_sessao = sessao
            self._log("Token de autenticação renovado automaticamente do navegador.")
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

    def _atualizar_botao_umalinha(self) -> None:
        botao = getattr(self, "botao_umalinha", None)
        if botao is None:
            return
        if self.gerar_umalinha:
            botao.setText("1 linha: ATIVADA")
            botao.setToolTip(
                "Clique para DESATIVAR: o resultado passa a ser gravado como "
                "uma linha por débito, sem mexer na coluna OBSERVAÇÃO."
            )
            botao.setStyleSheet(
                "QPushButton { background-color: #1B5E20; color: white; "
                "font-weight: bold; }"
            )
        else:
            botao.setText("1 linha: DESATIVADA")
            botao.setToolTip(
                "Clique para ATIVAR: todos os débitos voltam a ser gravados "
                "em uma única linha na coluna OBSERVAÇÃO."
            )
            botao.setStyleSheet(
                "QPushButton { background-color: #555555; color: white; "
                "font-weight: bold; }"
            )

    def _alternar_gerar_umalinha(self) -> None:
        self.gerar_umalinha = not self.gerar_umalinha
        self._config.setValue("gerar_umalinha", self.gerar_umalinha)
        self._atualizar_botao_umalinha()
        estado = "ativada" if self.gerar_umalinha else "desativada"
        self._log(f"Geração em 1 linha (OBSERVAÇÃO) {estado}.")
        self.statusBar().showMessage(f"Saída em 1 linha {estado}.")

    def _usa_observacao(self) -> bool:
        """Usa a OBSERVAÇÃO somente se o botão estiver ativado e a aba tiver a coluna."""
        return bool(
            self.gerar_umalinha
            and self.caminho_entrada
            and self.caminho_entrada.suffix.lower() in (".xlsx", ".xlsm")
            and tem_coluna_observacao(self.caminho_entrada)
        )

    @staticmethod
    def _pergunta_sim_nao(titulo: str, texto: str,
                          padrao_sim: bool = False,
                          confirmar_nao: bool = False) -> bool:
        """Caixa de confirmação com botões explicitamente em português."""
        caixa = QMessageBox()
        caixa.setWindowTitle(titulo)
        caixa.setText(texto)
        caixa.setIcon(QMessageBox.Icon.Question)
        botao_sim = caixa.addButton("Sim", QMessageBox.ButtonRole.YesRole)
        botao_nao = caixa.addButton("Não", QMessageBox.ButtonRole.NoRole)
        caixa.setDefaultButton(botao_sim if padrao_sim else botao_nao)
        caixa.exec()
        if caixa.clickedButton() is botao_sim:
            return True
        if caixa.clickedButton() is None:
            # ESC ou fechar a caixa: mantem a opcao segura (retomar).
            return padrao_sim
        if confirmar_nao:
            # Confirmação extra para não descartar o progresso por engano.
            confirma = QMessageBox()
            confirma.setWindowTitle("Tem certeza?")
            confirma.setText(
                "Você escolheu NÃO continuar.\n\n"
                "O progresso salvo será descartado e a consulta começará do "
                "zero. Tem certeza?"
            )
            confirma.setIcon(QMessageBox.Icon.Warning)
            botao_voltar = confirma.addButton(
                "Voltar", QMessageBox.ButtonRole.NoRole
            )
            botao_descartar = confirma.addButton(
                "Sim, descartar e começar do zero", QMessageBox.ButtonRole.YesRole
            )
            confirma.setDefaultButton(botao_voltar)
            confirma.exec()
            if confirma.clickedButton() is not botao_descartar:
                return True  # voltou: mantém a retomada
        return False

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
            larguras=[185, 150, 225, 105, 120, 270, 90, 100, 100, 165, None],
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

        import ctypes.util
        import shutil

        item(True, f"Python {sys.version.split()[0]}")
        for pacote in ("PyQt6", "playwright", "openpyxl", "cryptography"):
            item(importlib.util.find_spec(pacote) is not None, f"{pacote} instalado")

        item(
            ctypes.util.find_library("xcb-cursor") is not None,
            "libxcb-cursor0 (plugin xcb do Qt) instalada",
        )
        item(
            shutil.which("secret-tool") is not None,
            "secret-tool (keyring dos cookies) instalado",
        )
        if sys.platform.startswith("linux"):
            vboxclient = shutil.which("VBoxClient")
            item(
                vboxclient is not None,
                (
                    "VBoxClient instalado (arrastar/soltar e área de transferência)"
                    if vboxclient else
                    "VBoxClient ausente — use 'Instalar Guest Additions' "
                    "na aba Ambiente para habilitar arrastar/soltar"
                ),
            )
        navegadores = [
            nome for nome in ("brave-browser", "brave", "google-chrome",
                              "chromium", "microsoft-edge", "vivaldi")
            if shutil.which(nome) is not None
        ]
        item(bool(navegadores), f"Navegador Chromium: {', '.join(navegadores) or 'nenhum'}")

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

        item(
            _cdp_ativo(),
            f"Navegador Chromium com depuração na porta {CDP_URL.rsplit(':', 1)[-1]}",
        )

        try:
            cookies = buscar_cookies_no_brave()
            sessao = cookies.get("JSESSIONID", "")
            self.cookie_sessao = sessao
            detalhe = (
                "JWT + JSESSIONID encontrados" if sessao
                else "JWT encontrado (JSESSIONID criado pelo site)"
            )
            item(True, f"Cookies do navegador: {detalhe}")
        except Exception as exc:  # noqa: BLE001
            item(False, f"Cookies do navegador: {exc}")

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

    def instalar_guest_additions(self) -> None:
        """Instala o VBoxClient (drag-and-drop / área de transferência)."""
        diretorio = str(Path(__file__).resolve().parent)
        script = Path(diretorio) / "instalar_guest_additions.sh"
        if not script.exists():
            QMessageBox.warning(
                self, "Guest Additions",
                f"Script não encontrado: {script}",
            )
            return
        self._iniciar_config_worker(
            ComandoWorker([[str(script)]], diretorio=diretorio, parent=self)
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
        """Fecha (se necessário) e abre um Chromium com depuração remota."""
        if _cdp_ativo():
            QMessageBox.information(
                self, "Navegador",
                "O navegador já está com a depuração ativa na porta 9222.",
            )
            return
        if self._pergunta_sim_nao(
            "Abrir navegador com depuração",
            "Para usar a sessão real do navegador (Brave, Chrome, Edge, "
            "Vivaldi ou Chromium) é necessário fechá-lo e reabri-lo com a "
            "depuração ativa (porta 9222).\n\n"
            "Deseja fechar o navegador e abrir novamente?",
        ):
            self._iniciar_config_worker(BraveWorker(fechar=True, parent=self))

    def abrir_sistema(self) -> None:
        """Abre a página do sistema no navegador Chromium detectado."""
        try:
            if sys.platform == "darwin":
                binario = BraveWorker._caminho_brave()
                if binario:
                    subprocess.Popen([binario, HOME_URL])
                else:
                    subprocess.Popen(["open", HOME_URL])
            else:
                import webbrowser

                webbrowser.open(HOME_URL)
            self._log_ambiente(f"Sistema aberto no navegador: {HOME_URL}")
            if not _cdp_ativo():
                self._log_ambiente(
                    "Dica: use 'Abrir navegador com depuração (9222)' para que "
                    "o programa reaproveite a sessão logada."
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

    def limpar_progresso(self) -> None:
        """Apaga o progresso da aba atual (ação explícita do usuário)."""
        if not self.caminho_entrada:
            return
        aba = self.combo_aba.currentText() if self.combo_aba.isEnabled() else ""
        if not self._pergunta_sim_nao(
            "Limpar progresso",
            f"Apagar o progresso salvo da aba '{aba or 'padrão'}'?\n\n"
            "A próxima execução começará do zero.",
            padrao_sim=False,
        ):
            return
        arquivos = [
            _arquivo_progresso(self.caminho_entrada, aba),
            _arquivo_progresso_revisao(self.caminho_entrada, aba),
            _arquivo_progresso_revisao_completa(self.caminho_entrada, aba),
        ]
        try:
            for arquivo in arquivos:
                if arquivo.exists():
                    arquivo.unlink()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Limpar progresso", str(exc))
            return
        self.resultados = []
        self.tabela_resultados.setRowCount(0)
        self._atualizar_painel_resultados()
        self._log("Progresso limpo — a próxima execução começa do zero.")
        self.statusBar().showMessage("Progresso limpo.")

    def reabrir_ultima_planilha(self, silent: bool = False) -> None:
        """Abre a planilha base desta versão (ou a última usada, se não houver)."""
        if PLANILHA_BASE.exists():
            caminho = str(PLANILHA_BASE)
        else:
            caminho = str(self._config.value("ultima_planilha", "") or "")
        aba = str(self._config.value("ultima_aba", "") or "")
        if not caminho or not Path(caminho).exists():
            self.botao_reabrir.setEnabled(False)
            if not silent:
                QMessageBox.information(
                    self, "Reabrir planilha",
                    "A última planilha não existe mais. Escolha uma planilha.",
                )
            return
        self._abrir_planilha(caminho, aba)
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

    def _progresso_pela_planilha(self, arquivo_progresso: Path | None,
                                 aba: str) -> dict | None:
        """Reconstrói o progresso a partir da planilha Resultado já gerada.

        Se o contador de progresso foi perdido (ESC no diálogo, arquivo
        sobrescrito etc.), os débitos continuam salvos no XLSX. Daí dá para
        saber quais documentos já foram consultados e retomar do próximo.
        """
        if not self.caminho_entrada or arquivo_progresso is None:
            return None
        planilha = _pasta_relatorios() / f"{self.caminho_entrada.stem}_resultado.xlsx"
        if not planilha.exists():
            return None

        from openpyxl import load_workbook

        wb = load_workbook(planilha, read_only=True, data_only=True)
        try:
            ws = wb["Resultado"] if "Resultado" in wb.sheetnames else wb.active
            cabecalho = [
                "" if c is None else str(c).strip()
                for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
            ]
            if "CPF" not in cabecalho:
                return None
            col_cpf = cabecalho.index("CPF")
            consultados: set[str] = set()
            linhas: list[dict] = []
            for linha in ws.iter_rows(min_row=2, values_only=True):
                doc = re.sub(r"\D", "", str(linha[col_cpf] or ""))
                if not doc:
                    continue
                consultados.add(doc)
                item = {
                    cabecalho[i]: linha[i]
                    for i in range(min(len(cabecalho), len(linha)))
                    if cabecalho[i]
                }
                linhas.append(item)
        finally:
            wb.close()

        # Se a planilha classica foi sobrescrita com poucos resultados, usa
        # as observacoes ja gravadas no relatorio de 1 linha como fonte.
        if len(consultados) < 10:
            relatorio = (
                _pasta_relatorios()
                / f"{self.caminho_entrada.stem}_1_LINHA_TODOS_OS_DEBITOS.xlsx"
            )
            if relatorio.exists():
                wb2 = load_workbook(relatorio, read_only=True, data_only=True)
                try:
                    aba_ws = wb2[aba] if aba in wb2.sheetnames else None
                    if aba_ws is not None:
                        local = localizar_cabecalho(aba_ws)
                        if local is not None:
                            linha_cab, col_doc, col_obs = local
                            for linha in aba_ws.iter_rows(
                                min_row=linha_cab + 1, values_only=True
                            ):
                                doc = re.sub(
                                    r"\D", "",
                                    str(linha[col_doc - 1]
                                        if col_doc <= len(linha) else ""),
                                )
                                obs = str(
                                    linha[col_obs - 1] or ""
                                    if col_obs <= len(linha) else ""
                                ).strip()
                                if doc and obs:
                                    consultados.add(doc)
                finally:
                    wb2.close()

        if not consultados:
            return None

        indices: dict[str, int] = {}
        for indice, registro in enumerate(self.registros, start=1):
            chave_doc = next(
                (k for k in registro if "cpf" in k.lower() or "cnpj" in k.lower()),
                None,
            )
            valor = registro.get(chave_doc, "") if chave_doc else ""
            doc = re.sub(r"\D", "", str(valor or ""))
            if doc:
                indices[doc] = indice
        ultimo = max((indices[d] for d in consultados if d in indices), default=0)
        if ultimo <= 0:
            return None
        for item in linhas:
            doc = re.sub(r"\D", "", str(item.get("CPF", "")))
            item["_indice"] = indices.get(doc, 0)

        salvar_progresso(
            arquivo_progresso, ultimo, len(self.registros), linhas,
            em_andamento=False, indice_atual=ultimo,
            entrada=str(self.caminho_entrada), aba=aba, forcado=True,
        )
        self._log(
            f"Progresso reconstruído da planilha salva: {ultimo} registro(s) "
            "já consultado(s)."
        )
        return ler_progresso(arquivo_progresso)

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
        # Contador zerado/suspeito: reconstrói do que já está na planilha
        # Resultado salva, em vez de recomeçar do registro 1.
        if (
            (progresso is None or int(progresso.get("processados", 0)) < 10)
            and arquivo_progresso is not None
        ):
            reconstruido = self._progresso_pela_planilha(arquivo_progresso, aba)
            if reconstruido is not None:
                progresso = reconstruido
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

            # Retomagem AUTOMÁTICA: a caixa antiga podia ser fechada com ESC
            # e o programa interpretava isso como "descartar o progresso".
            indice_inicial = refazer - 1
            # Descarta os resultados do registro que será refeito.
            resultados_iniciais = [
                item for item in progresso["resultados"]
                if int(item.get("_indice") or 0) < refazer
            ]
            self._log(
                f"Retomando automaticamente no registro {refazer} de "
                f"{len(self.registros)}{detalhe}."
            )
        elif progresso and progresso["processados"] >= len(self.registros):
            salvar_progresso(
                arquivo_progresso, 0, len(self.registros), [],
                entrada=entrada, aba=aba,
            )

        inicio_manual = self._registro_inicial_manual()
        if inicio_manual > 1:
            indice_inicial = inicio_manual - 1
            resultados_iniciais = []
            self._log(
                f"Consulta começando no registro {inicio_manual} "
                "(definido manualmente)."
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
        if hasattr(self, "botao_revisao_completa"):
            self.botao_revisao_completa.setEnabled(False)

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
        self._atualizar_botao_pagos()
        self._avisar_se_http()
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
        self._fluxo_completo = False
        self._consulta_interrompida = True
        if self.worker_revisao is not None and self.worker_revisao.isRunning():
            self.worker_revisao.stop()
            self._log("Parada solicitada na revisão paralela.")
        if self.worker is not None:
            self.worker.stop()
            self.label_status.setText("Parando após o item atual...")
            self._log("Parada solicitada pelo usuário.")
        if self.worker_pagos is not None and self.worker_pagos.isRunning():
            self.worker_pagos.stop()
            self.label_status.setText("Parando após o documento atual...")
            self._log("Parada solicitada na verificação de pagos.")
        if self.worker_cifs is not None and self.worker_cifs.isRunning():
            self.worker_cifs.stop()
            self.label_status.setText("Parando após o documento atual...")
            self._log("Parada solicitada na verificação de CIFs.")

    def _ao_progresso(self, percentual: int, mensagem: str) -> None:
        self._atualizar_painel_resultados()
        self.barra_progresso.setValue(percentual)
        self.label_status.setText(mensagem)
        self.statusBar().showMessage(mensagem)

    def _ao_progresso_revisao(self, percentual: int, mensagem: str) -> None:
        """Mostra o andamento da revisão paralela sem mexer na barra principal."""
        if self._total_revisao:
            feitos = len({
                str(r.get("CPF", "")).strip()
                for r in self.resultados_revisao
                if str(r.get("CPF", "")).strip()
            })
            # Mantém o nome/CPF que veio na mensagem do worker.
            rotulo = ""
            if ":" in mensagem:
                rotulo = mensagem.split(":", 1)[1].strip()
            mensagem = f"Processado {feitos}/{self._total_revisao}"
            if rotulo:
                mensagem += f": {rotulo}"
        self.label_revisao_paralela.setText(f"Revisão paralela: {mensagem}")

    def _inserir_linha_tabela(self, item: dict) -> None:
        """Insere um item na tabela sem alterar a lista ``resultados``.

        Usado ao popular a tabela com resultados já existentes (retomada),
        evitando que eles sejam adicionados uma segunda vez à lista.
        """
        posicao = self.tabela_resultados.rowCount()
        if posicao >= MAX_LINHAS_TABELA:
            # Janela deslizante: descarta a linha mais antiga para que a tabela
            # continue mostrando o que esta sendo consultado agora, mantendo
            # o consumo de memoria limitado.
            self.tabela_resultados.removeRow(0)
            posicao = self.tabela_resultados.rowCount()
        self.tabela_resultados.insertRow(posicao)
        self.tabela_resultados.setRowHeight(posicao, 30)
        for coluna, nome in enumerate(COLUNAS_SAIDA):
            valor = str(item.get(nome, ""))
            if nome == "Valor (R$)":
                valor = f"R$ {_para_numero(valor):.2f}".replace(".", ",")
            self.tabela_resultados.setItem(
                posicao, coluna, QTableWidgetItem(valor)
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
        if status in ("DOCUMENTO INVÁLIDO", STATUS_PENDENTE):
            return p["texto_neutro"]
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

    def _registro_inicial_manual(self) -> int:
        """Número do registro escolhido manualmente (1 = automático)."""
        campo = getattr(self, "campo_registro_inicial", None)
        return int(campo.value()) if campo is not None else 1

    def _avisar_se_http(self) -> None:
        """Alerta quando o modo HTTP experimental está ativado."""
        campo = getattr(self, "campo_sem_navegador", None)
        if campo is None or not campo.isChecked():
            return
        aviso = (
            "AVISO: 'Tentar SEM navegador (HTTP direto)' está ATIVADO. "
            "O site exige JavaScript; nesse modo as consultas ficam "
            "NAO VERIFICADO. Desmarque essa opção na aba Consulta."
        )
        self._log(aviso)
        self.statusBar().showMessage(aviso)

    def _atualizar_painel_resultados(self) -> None:
        """Atualiza contadores por situação na aba Resultados."""
        labels = getattr(self, "labels_painel", None)
        if not labels:
            return
        from collections import Counter
        contagens = Counter(
            str(item.get("Status", "") or "").strip() or "(vazio)"
            for item in self.resultados
        )
        cores = {
            "TOTAL": "#333333",
            "VERIFICADO": "#1B5E20",
            STATUS_PENDENTE: "#9C6500",
            STATUS_ERRO_CONSULTA: "#9C0006",
            "DOCUMENTO INVÁLIDO": "#666666",
            "NAO VERIFICADO": "#B26A00",
            "ERRO": "#9C0006",
        }
        titulos = {STATUS_PENDENTE: "CLASSIFICADO — INDETERMINADO"}
        for chave, rotulo in labels.items():
            valor = (
                len(self.resultados) if chave == "TOTAL"
                else contagens.get(chave, 0)
            )
            titulo = titulos.get(chave, chave if chave != "TOTAL" else "Total")
            rotulo.setText(
                f"<b style='color:{cores.get(chave, "#333333")}'>"
                f"{titulo}: {valor}</b>"
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
        status = str(item.get("Status", "")).strip().upper()
        # Documento inválido e snapshot legado não voltam para a revisão:
        # refazer a consulta não adianta (o site não aceita um e o outro
        # depende de fonte externa, não da sessão atual).
        # Indeterminado com Erro é falha de consulta: merece nova tentativa.
        # Indeterminado sem Erro é dado migrado analisado, sem reconsulta útil.
        if status == STATUS_PENDENTE:
            return bool(str(item.get("Erro", "") or "").strip())
        if status == "DOCUMENTO INVÁLIDO":
            return False
        if status not in ("OK", "VERIFICADO"):
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
        return _pincel_revisao()

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
                    item.setData(_MARCA_REVISAO, True)
                    item.setBackground(pincel)
                    # Mantém a cor de texto do tema: preto some no tema escuro.

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
            for coluna in range(self.tabela_resultados.columnCount()):
                item = self.tabela_resultados.item(linha, coluna)
                if item is not None:
                    item.setData(_MARCA_REVISAO, False)
            for resultado in self.resultados:
                if str(resultado.get("CPF", "")).strip() == documento:
                    self._aplicar_estilo_linha_resultado(linha, resultado)
                    break
            else:
                # Linha sem resultado correspondente: ainda assim restaura a
                # cor de texto do tema para não ficar invisível.
                cor_texto = QBrush(QColor(PALETA["texto"]))
                for coluna in range(self.tabela_resultados.columnCount()):
                    item = self.tabela_resultados.item(linha, coluna)
                    if item is not None:
                        item.setForeground(cor_texto)
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

    def _ao_alternar_revisao_paralela(self, marcado: bool) -> None:
        """ Liga/desliga a revisão paralela direto pelo checkbox."""
        if marcado:
            if self.worker_revisao is not None and self.worker_revisao.isRunning():
                return
            if self.worker is not None and self.worker.isRunning():
                self._iniciar_revisao_paralela()
                self._log("Revisão paralela iniciada pelo botão.")
            else:
                self._log(
                    "Revisão paralela ativada — começará ao clicar em Executar."
                )
        elif self.worker_revisao is not None and self.worker_revisao.isRunning():
            self.worker_revisao.stop()
            self._log("Revisão paralela interrompida pelo botão.")

    def _iniciar_revisao_paralela(self, documentos: set[str] | None = None) -> None:
        """Revisa pendências em outro worker enquanto a consulta principal roda."""
        if self.worker_revisao is not None and self.worker_revisao.isRunning():
            return
        if documentos is None:
            documentos = self._documentos_pendentes()
        if not documentos:
            self.label_revisao_paralela.setText(
                "Revisão paralela: sem pendências para revisar"
            )
            self._log("Revisão paralela: não há pendências para revisar.")
            return

        # Retoma a revisão pelo progresso salvo, sem recomeçar do primeiro.
        self.resultados_revisao = []
        total_documentos = len(documentos)
        if self.caminho_entrada:
            aba = self.combo_aba.currentText() if self.combo_aba.isEnabled() else ""
            arquivo_rev = _arquivo_progresso_revisao(self.caminho_entrada, aba)
            progresso_rev = (
                ler_progresso(arquivo_rev) if arquivo_rev.exists() else None
            )
            if progresso_rev:
                ja_revisados = {
                    str(r.get("CPF", "")).strip()
                    for r in (progresso_rev.get("resultados") or [])
                    if str(r.get("CPF", "")).strip()
                }
                if ja_revisados:
                    self.resultados_revisao = list(
                        progresso_rev.get("resultados") or []
                    )
                    documentos = documentos - ja_revisados
                    self._log(
                        "Revisão paralela retomada: "
                        f"{len(ja_revisados)} documento(s) já revisado(s)."
                    )
        self._total_revisao = total_documentos
        if not documentos:
            self.label_revisao_paralela.setText(
                "Revisão paralela: tudo já revisado"
            )
            self._log("Revisão paralela: nada a fazer (progresso já completo).")
            return

        registros = [
            dict(r, _indice_original=i)
            for i, r in enumerate(self.registros, start=1)
            if self._documento_registro(r) in documentos
        ]
        if not registros:
            self.label_revisao_paralela.setText(
                "Revisão paralela: registros originais não encontrados"
            )
            self._log("Revisão paralela: registros originais não encontrados.")
            return

        # A faixa é aplicada só ao documento que estiver sendo consultado.
        self.worker_revisao = ConsultaWorker(
            registros,
            self.campo_cookie.text(),
            cookie_sessao=self.cookie_sessao,
            modo_revisao=True,
            incluir_pagos=True,
            sem_navegador=self.campo_sem_navegador.isChecked(),
            pausa=max(0.5, self.campo_velocidade.value()),
            variar_pausa=self.campo_variar_pausa.isChecked(),
            controle=self.controle_consulta,
            parent=self,
        )
        self.worker_revisao.documento_revisado.connect(
            self._ao_documento_revisado_paralelo
        )
        self.worker_revisao.documento_iniciado.connect(
            self._ao_documento_iniciado_revisao
        )
        self.worker_revisao.progresso.connect(self._ao_progresso_revisao)
        self.worker_revisao.finalizado.connect(self._ao_finalizar_revisao_paralela)
        self.worker_revisao.finished.connect(self._limpar_worker_revisao)
        self.worker_revisao.aviso.connect(self._log)
        self._revisao_paralela_ativa = True
        self.botao_revisar.setEnabled(False)
        if hasattr(self, "botao_revisao_completa"):
            self.botao_revisao_completa.setEnabled(False)
        self.label_revisao_paralela.setText(
            f"Revisão paralela: iniciando ({len(registros)} pendente(s))..."
        )
        self.statusBar().showMessage(
            f"Revisão paralela iniciada: {len(registros)} pendente(s)."
        )
        self.worker_revisao.start()
        self._atualizar_botao_pagos()
        self._log(f"Revisão paralela iniciada: {len(registros)} pendente(s).")

    def _ao_documento_iniciado_revisao(self, documento: str) -> None:
        """Destaca somente o documento que está sendo revisado agora."""
        self._desmarcar_em_revisao()
        documento = str(documento).strip()
        if documento:
            self._marcar_em_revisao({documento})

    def _ao_documento_revisado_paralelo(self, documento: str, novos: list[dict]) -> None:
        """Guarda a revisão paralela à parte e salva seu progresso."""
        self.resultados_revisao.extend(novos)
        self._salvar_progresso_revisao()

    def _salvar_progresso_revisao(self) -> None:
        """Persiste o andamento da revisão para retomar sem recomeçar."""
        if not self.caminho_entrada or not self._total_revisao:
            return
        aba = self.combo_aba.currentText() if self.combo_aba.isEnabled() else ""
        arquivo = _arquivo_progresso_revisao(self.caminho_entrada, aba)
        feitos = {
            str(r.get("CPF", "")).strip()
            for r in self.resultados_revisao
            if str(r.get("CPF", "")).strip()
        }
        salvar_progresso(
            arquivo, len(feitos), self._total_revisao, self.resultados_revisao,
            em_andamento=True, indice_atual=len(feitos),
            entrada=str(self.caminho_entrada), aba=aba,
        )

    def _ao_finalizar_revisao_paralela(self, mensagem: str) -> None:
        self._log(f"Revisão paralela: {mensagem}")
        self.label_revisao_paralela.setText(
            "Revisão paralela: interrompida" if "interrompido" in mensagem.lower()
            else "Revisão paralela: concluída"
        )
        self._revisao_paralela_ativa = False
        # Concluída: limpa o progresso da revisão. Interrompida: mantém para
        # retomar de onde parou na próxima execução.
        if self.caminho_entrada:
            aba = self.combo_aba.currentText() if self.combo_aba.isEnabled() else ""
            arquivo_rev = _arquivo_progresso_revisao(self.caminho_entrada, aba)
            if "interrompido" in mensagem.lower():
                self._salvar_progresso_revisao()
            elif arquivo_rev.exists():
                try:
                    arquivo_rev.unlink()
                except Exception:  # noqa: BLE001
                    pass
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
        # Reordena pela planilha e reconstrói a tabela na ordem correta.
        self.resultados = _ordenar_resultados(self.resultados)
        self.tabela_resultados.setRowCount(0)
        for item in self.resultados:
            self._inserir_linha_tabela(item)
        # Persiste a mesclagem: os documentos revisados ficam de fora da
        # consulta principal e, sem isso, sumiriam do progresso ao reabrir.
        if self.caminho_entrada:
            aba = (
                self.combo_aba.currentText()
                if self.combo_aba.isEnabled() else ""
            )
            arquivo = _arquivo_progresso(self.caminho_entrada, aba)
            progresso = ler_progresso(arquivo)
            if progresso is not None:
                salvar_progresso(
                    arquivo,
                    progresso["processados"],
                    progresso["total"],
                    self.resultados,
                    em_andamento=bool(progresso.get("em_andamento")),
                    indice_atual=int(progresso.get("indice_atual") or 0),
                    entrada=str(self.caminho_entrada),
                    aba=aba,
                )
        self._log(f"Revisão paralela mesclada: {total} linha(s) atualizada(s).")
        self.botao_salvar.setEnabled(bool(self.resultados))
        self._atualizar_botao_revisar()

    def _arquivo_destino_observacao(self) -> Path:
        """Arquivo único onde a OBSERVAÇÃO (débitos) é preenchida."""
        stem = self.caminho_entrada.stem if self.caminho_entrada else PLANILHA_BASE.stem
        return _pasta_relatorios() / f"{stem}_1_LINHA_TODOS_OS_DEBITOS.xlsx"

    def _preencher_observacoes_memoria(self) -> int:
        """Aplica ao arquivo destino os débitos já obtidos na sessão."""
        if not self.resultados:
            return 0
        mapa = mapa_observacoes(self.resultados)
        if not mapa:
            return 0
        gravar_relatorio_com_observacao(
            self.caminho_entrada or PLANILHA_BASE,
            mapa,
            self._arquivo_destino_observacao(),
        )
        return len(mapa)

    def _registros_observacao_vazia(self) -> list[dict]:
        """CPFs/CNPJs das linhas com OBSERVAÇÃO (débitos) ainda vazia."""
        if self._usa_observacao():
            destino = self._arquivo_destino_observacao()
            origem = destino if destino.exists() else (
                self.caminho_entrada or PLANILHA_BASE
            )
        else:
            origem = self.caminho_entrada or PLANILHA_BASE
        if not origem or not Path(origem).exists():
            return []
        combo = getattr(self, "combo_aba", None)
        aba = combo.currentText() if combo is not None else ""
        return registros_observacao_vazia(Path(origem), aba or None)

    def _progresso_revisao_completa(self) -> Path | None:
        if not self.caminho_entrada:
            return None
        aba = (
            self.combo_aba.currentText()
            if self.combo_aba is not None and self.combo_aba.isEnabled()
            else ""
        )
        return _arquivo_progresso_revisao_completa(self.caminho_entrada, aba)

    def _iniciar_revisao_observacoes(
            self, registros: list[dict], incluir_pagos: bool = True,
            arquivo_progresso: Path | None = None, indice_inicial: int = 0,
            resultados_iniciais: list[dict] | None = None) -> None:
        """Lança a consulta dos CPFs que ficaram com OBSERVAÇÃO vazia."""
        self._avisar_se_http()
        self._registros_revisao = list(registros)
        self._carregar_cookie_brave(silent=True)
        self.barra_progresso.setValue(0)
        self.label_status.setText("Revisando observações vazias...")
        self.botao_executar.setEnabled(False)
        self.botao_parar.setEnabled(True)
        self.botao_parar_resultados.setEnabled(True)
        self.botao_salvar.setEnabled(False)
        self.botao_revisar.setEnabled(False)
        if hasattr(self, "botao_revisao_completa"):
            self.botao_revisao_completa.setEnabled(False)
        self.botao_escolher.setEnabled(False)

        self.worker = ConsultaWorker(
            registros,
            self.campo_cookie.text(),
            cookie_sessao=self.cookie_sessao,
            indice_inicial=indice_inicial,
            arquivo_progresso=arquivo_progresso,
            resultados_iniciais=resultados_iniciais,
            modo_revisao=True,
            incluir_pagos=True,
            sem_navegador=self.campo_sem_navegador.isChecked(),
            pausa=self.campo_velocidade.value(),
            variar_pausa=self.campo_variar_pausa.isChecked(),
            controle=self.controle_consulta,
            entrada=str(self.caminho_entrada) if self.caminho_entrada else "",
            aba=(
                self.combo_aba.currentText()
                if self.combo_aba.isEnabled() else ""
            ),
            parent=self,
        )
        self.worker.progresso.connect(self._ao_progresso)
        self.worker.aviso.connect(self._log)
        self.worker.documento_revisado.connect(self._ao_documento_revisado)
        self.worker.documento_iniciado.connect(self._ao_documento_iniciado_revisao)
        self.worker.finalizado.connect(self._ao_finalizar)
        self.worker.finished.connect(self._limpar_worker)
        self.worker.start()
        self._atualizar_botao_pagos()
        self._log(
            f"Revisando {len(registros)} documento(s) com OBSERVAÇÃO vazia."
        )

    def revisao_completa(self) -> None:
        """Reconsulta TODOS os documentos da aba selecionada."""
        if self.worker is not None and self.worker.isRunning():
            return
        if self.worker_revisao is not None and self.worker_revisao.isRunning():
            return
        if not self.registros:
            QMessageBox.information(
                self, "Revisão completa", "Nenhuma planilha carregada."
            )
            return

        vistos: set[str] = set()
        documentos: list[dict] = []
        for registro in self.registros:
            documento = self._documento_registro(registro)
            if not documento or documento in vistos:
                continue
            vistos.add(documento)
            documentos.append(dict(registro))
        if not documentos:
            QMessageBox.information(
                self, "Revisão completa",
                "Nenhum CPF/CNPJ válido encontrado na aba selecionada.",
            )
            return

        self._consulta_interrompida = False
        self._revisao_concluida.clear()
        arquivo = self._progresso_revisao_completa()
        resultados_iniciais: list[dict] = []
        progresso = ler_progresso(arquivo) if arquivo and arquivo.exists() else None
        if progresso:
            resultados_iniciais = list(progresso.get("resultados") or [])
        # Só já foi conferido quem tem resultado confirmado. Usa o progresso
        # da revisão completa e, se ele não existir, os resultados já
        # carregados do progresso principal (execuções feitas antes desta
        # versão). Documentos que falharam voltam para a fila.
        base_resultados = resultados_iniciais or self.resultados
        concluidos = {
            re.sub(r"\D", "", str(item.get("CPF", "")))
            for item in base_resultados
            if str(item.get("Status", "")).strip().upper()
            in ("OK", "VERIFICADO")
        }
        antes = len(documentos)
        documentos = [
            registro for registro in documentos
            if re.sub(r"\D", "", self._documento_registro(registro))
            not in concluidos
        ]
        if progresso:
            self.resultados = _ordenar_resultados(resultados_iniciais)
            self.tabela_resultados.setRowCount(0)
            for item in self.resultados:
                self._inserir_linha_tabela(item)
            self._atualizar_painel_resultados()
        if len(documentos) != antes:
            self._log(
                f"Retomando revisão completa: {len(documentos)} documento(s) "
                f"pendente(s) de {antes}."
            )
        inicio_manual = self._registro_inicial_manual()
        if inicio_manual > 1:
            documentos = documentos[inicio_manual - 1:]
            self._log(
                f"Revisão completa começando no registro {inicio_manual} "
                "(definido manualmente)."
            )
        self._revisao_completa_ativa = True
        self._log(
            f"Revisão completa: {len(documentos)} documento(s) da aba "
            f"'{self.combo_aba.currentText() if self.combo_aba.isEnabled() else ''}'."
        )
        self._iniciar_revisao_observacoes(
            documentos, incluir_pagos=True, arquivo_progresso=arquivo,
            resultados_iniciais=resultados_iniciais,
        )

    def revisar(self) -> None:
        """Após uma parada, consulta só o que falta; senão, os itens pendentes."""
        if self.worker is not None and self.worker.isRunning():
            return
        if self.worker_revisao is not None and self.worker_revisao.isRunning():
            return

        # Versão "Lista 1 aba": primeiro aplica os débitos já obtidos; depois
        # consulta somente o que continua com OBSERVAÇÃO (débitos) vazia.
        if self._usa_observacao():
            self._consulta_interrompida = False
            self._revisao_concluida.clear()
            ja_preenchidos = self._preencher_observacoes_memoria()
            if ja_preenchidos:
                self._log(
                    f"OBSERVAÇÃO preenchida com {ja_preenchidos} CPF(s) "
                    "já consultado(s) nesta sessão."
                )
            pendentes = self._registros_observacao_vazia()
            if not pendentes:
                QMessageBox.information(
                    self, "Revisão",
                    "Não há observações vazias para revisar. "
                    "Os débitos já obtidos foram aplicados ao arquivo.",
                )
                return
            self._iniciar_revisao_observacoes(pendentes, incluir_pagos=True)
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
                dict(r, _indice_original=int(_campo(r, "_indice_original") or i))
                for i, r in enumerate(base, start=1)
                if self._documento_registro(r) not in concluidos
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
                dict(r, _indice_original=i)
                for i, r in enumerate(self.registros, start=1)
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
        # A faixa acompanha apenas o documento em consulta, via sinal do worker.

        self._carregar_cookie_brave(silent=True)
        self.barra_progresso.setValue(0)
        self.label_status.setText("Revisando...")
        self.botao_executar.setEnabled(False)
        self.botao_parar.setEnabled(True)
        self.botao_parar_resultados.setEnabled(True)
        self.botao_salvar.setEnabled(False)
        self.botao_revisar.setEnabled(False)
        if hasattr(self, "botao_revisao_completa"):
            self.botao_revisao_completa.setEnabled(False)
        self.botao_escolher.setEnabled(False)

        self.worker = ConsultaWorker(
            registros,
            self.campo_cookie.text(),
            cookie_sessao=self.cookie_sessao,
            modo_revisao=True,
            incluir_pagos=True,
            sem_navegador=self.campo_sem_navegador.isChecked(),
            pausa=self.campo_velocidade.value(),
            variar_pausa=self.campo_variar_pausa.isChecked(),
            controle=self.controle_consulta,
            parent=self,
        )
        self.worker.progresso.connect(self._ao_progresso)
        self.worker.aviso.connect(self._log)
        self.worker.documento_revisado.connect(self._ao_documento_revisado)
        self.worker.documento_iniciado.connect(self._ao_documento_iniciado_revisao)
        self.worker.finalizado.connect(self._ao_finalizar)
        self.worker.finished.connect(self._limpar_worker)
        self.worker.start()
        self._atualizar_botao_pagos()
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
        if getattr(self, "_revisao_completa_ativa", False):
            self._revisao_completa_ativa = False
            if "Concluído" in mensagem:
                arquivo = self._progresso_revisao_completa()
                if arquivo and arquivo.exists():
                    try:
                        arquivo.unlink()
                    except OSError:
                        pass
        if self._fluxo_completo and "Concluído" in mensagem:
            if self._usa_observacao():
                try:
                    self._preencher_observacoes_memoria()
                except Exception as exc:  # noqa: BLE001
                    self._log(f"Não foi possível atualizar a OBSERVAÇÃO: {exc}")
            self._log("Consulta completa (2/3): classificando pagos...")
            QTimer.singleShot(50, lambda: self.verificar_pagos(auto=True))
            return
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

        # Nesta versão, ao concluir (consulta ou revisão) a OBSERVAÇÃO do
        # arquivo destino é atualizada automaticamente com os débitos obtidos.
        if "Concluído" in mensagem and self._usa_observacao():
            try:
                self._agendar_salvamento_saidas()
            except Exception as exc:  # noqa: BLE001
                self._log(f"Não foi possível atualizar a OBSERVAÇÃO: {exc}")
        elif "Concluído" in mensagem and self.resultados:
            # Com a saída em 1 linha DESATIVADA, salva direto a planilha
            # clássica (uma linha por débito), sem tocar na OBSERVAÇÃO.
            try:
                self._salvar_automatico()
            except Exception as exc:  # noqa: BLE001
                self._log(f"Não foi possível salvar o resultado: {exc}")

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
        """Salva em arquivo NOVO com data/hora — nunca sobrescreve outro."""
        if not self.resultados:
            return
        relatorio_com_observacao = self._usa_observacao()
        if relatorio_com_observacao:
            destino = (
                _pasta_relatorios()
                / f"{self.caminho_entrada.stem}_1_LINHA_TODOS_OS_DEBITOS.xlsx"
            )
        else:
            atributo = "_destino_auto_parcial" if parcial else "_destino_auto"
            destino = getattr(self, atributo)
            if destino is None:
                stem = (
                    self.caminho_entrada.stem
                    if self.caminho_entrada else "RELATORIO_DEBITOS"
                )
                sufixo = "resultado_parcial" if parcial else "resultado"
                destino = _pasta_relatorios() / (
                    f"{_novo_nome_relatorio(stem, sufixo)}.xlsx"
                )
                setattr(self, atributo, destino)
        try:
            if relatorio_com_observacao:
                self._agendar_salvamento_saidas()
            else:
                gravar_planilha(str(destino), self.resultados)
        except Exception as exc:  # noqa: BLE001
            self._log(f"Não foi possível salvar o relatório parcial: {exc}")
            QMessageBox.warning(self, "Relatório parcial", f"Não foi possível salvar: {exc}")
            return
        if relatorio_com_observacao:
            mensagem = f"OBSERVAÇÃO atualizada automaticamente em: {destino}"
        else:
            mensagem = (
                f"Relatório {'parcial ' if parcial else ''}"
                f"salvo automaticamente em: {destino}"
            )
        self._log(mensagem)
        self.statusBar().showMessage(mensagem)

    def _limpar_worker(self) -> None:
        self.worker = None
        self._atualizar_botao_revisar()

    def _gravar_resultado_classico(self) -> Path:
        """Gera a planilha Resultado estilo REVISADO: um débito por linha."""
        stem = self.caminho_entrada.stem if self.caminho_entrada else PLANILHA_BASE.stem
        destino = _pasta_relatorios() / f"{stem}_resultado.xlsx"
        gravar_planilha(str(destino), self.resultados)
        return destino

    def _limpar_salvador_saidas(self) -> None:
        self._salvador_saidas = None
        if self._salvamento_pendente:
            self._salvamento_pendente = False
            self._agendar_salvamento_saidas()

    def _agendar_salvamento_saidas(self) -> None:
        """Dispara a gravacao dos dois relatorios em segundo plano."""
        if not self._usa_observacao():
            return
        if self._salvador_saidas is not None and self._salvador_saidas.isRunning():
            self._salvamento_pendente = True
            self._log("Ja ha um salvamento em andamento; sera refeito ao terminar.")
            return
        stem = self.caminho_entrada.stem
        observacao = (
            _pasta_relatorios() / f"{stem}_1_LINHA_TODOS_OS_DEBITOS.xlsx"
        )
        resultado = _pasta_relatorios() / f"{stem}_resultado.xlsx"
        self._salvador_saidas = SalvarSaidasWorker(
            self.caminho_entrada, self.resultados, observacao, resultado, self
        )
        self._salvador_saidas.concluido.connect(self._log)
        self._salvador_saidas.finished.connect(self._limpar_salvador_saidas)
        self._salvador_saidas.start()
        self._log("Salvando relatorios em segundo plano...")
        self.statusBar().showMessage("Salvando relatorios em segundo plano...")

    def salvar_resultado(self) -> None:
        if not self.resultados:
            return

        # Quando a planilha de entrada é o relatório contábil, o destino é
        # SEMPRE o mesmo arquivo: RELATORIO_CONTABILIDADE_2022_ANALISTAS_1_LINHA_TODOS_OS_DEBITOS.xlsx.
        # Ele é preenchido na coluna OBSERVAÇÃO (débitos); se já existir, é
        # atualizado (não é criado outro arquivo).
        relatorio_com_observacao = self._usa_observacao()

        if relatorio_com_observacao:
            destino = (
                _pasta_relatorios()
                / f"{self.caminho_entrada.stem}_1_LINHA_TODOS_OS_DEBITOS.xlsx"
            )
            caminho = None
        else:
            stem = (
                self.caminho_entrada.stem
                if self.caminho_entrada else "RELATORIO_DEBITOS"
            )
            nome = f"{_novo_nome_relatorio(stem)}.xlsx"
            caminho, _ = QFileDialog.getSaveFileName(
                self, "Salvar resultado",
                str(_pasta_relatorios() / nome),
                "Excel (*.xlsx);;CSV (*.csv)",
            )
            if not caminho:
                return
            destino = None

        try:
            if relatorio_com_observacao:
                self._agendar_salvamento_saidas()
            else:
                destino = gravar_planilha(caminho, self.resultados)
        except PermissionError:
            QMessageBox.critical(
                self, "Erro ao salvar",
                "Não foi possível gravar. Feche a planilha "
                "RELATORIO_CONTABILIDADE_2022_ANALISTAS_1_LINHA_TODOS_OS_DEBITOS.xlsx "
                "se ela estiver aberta.",
            )
            return
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Erro ao salvar", str(exc))
            return
        mensagem = (
            f"OBSERVAÇÃO preenchida em: {destino}"
            if relatorio_com_observacao
            else f"Resultado salvo em: {destino}"
        )
        self._log(mensagem)
        self.statusBar().showMessage(mensagem)

    def closeEvent(self, event) -> None:  # noqa: N802
        if self._salvador_saidas is not None and self._salvador_saidas.isRunning():
            self._salvador_saidas.wait(2000)
        if self.worker_revisao is not None and self.worker_revisao.isRunning():
            self.worker_revisao.stop()
            self.worker_revisao.wait(3000)
        if self.worker is not None and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(3000)
        if self.worker_pagos is not None and self.worker_pagos.isRunning():
            self.worker_pagos.stop()
            self.worker_pagos.wait(3000)
        if self.worker_cifs is not None and self.worker_cifs.isRunning():
            self.worker_cifs.stop()
            self.worker_cifs.wait(3000)
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
