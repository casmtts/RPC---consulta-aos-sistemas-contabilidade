"""Tentativa de consulta SEM navegador (HTTP direto).

O site é JSF/PrimeFaces e normalmente exige JavaScript, então este modo é
experimental. Ele reenvia o formulário com o ``javax.faces.ViewState`` obtido
no próprio servidor e lê as tabelas da resposta HTML.

Regra de segurança: se a resposta não puder ser confirmada, levanta
``ConsultaNaoVerificada`` — nunca devolve "sem débitos" por dúvida.
"""

from __future__ import annotations

import html as html_mod
import http.cookiejar
import re
import unicodedata
import urllib.parse
import urllib.request
from datetime import datetime
from typing import Dict, List

from config import BASE_URL, HOME_URL, COOKIE_NAME
from regras_debito import situacao_eh_debito

DOMINIO_JWT = ".desenvolvecidade.com.br"
DOMINIO_SESSAO = "santanadoaraguaia-pa.desenvolvecidade.com.br"
SITUACOES_ATIVAS = ("DEBITO",)


class ConsultaNaoVerificada(Exception):
    """A consulta HTTP não pôde ser confirmada (nunca vira 'sem débitos')."""


class ConsultaInterrompidaHttp(Exception):
    """O usuário pediu para parar no meio da consulta HTTP."""


def _atributos(tag: str) -> Dict[str, str]:
    return {
        chave.lower(): valor
        for chave, valor in re.findall(r'([\w:.\-]+)\s*=\s*"([^"]*)"', tag)
    }


def _texto_limpo(fragmento: str) -> str:
    sem_tags = re.sub(r"<[^>]+>", " ", fragmento)
    # html.unescape é muito lento (pode ser 5-10x mais lento)
    # O site ja devolve texto limpo em geral, entao podemos pular
    return sem_tags.strip()


def _texto_visivel(fragmento: str) -> str:
    """Texto da resposta sem scripts, estilos nem painéis ocultos.

    O JSF mantém Mensagens de consultas anteriores em painéis escondidos;
    ler esse texto gerava falso "sem débitos".
    """
    sem_ocultos = re.sub(
        r"<(script|style)\b[^>]*>.*?</\1>", " ", fragmento, flags=re.I | re.S
    )
    # Campos hidden são elementos vazios: remove só a tag, nunca o conteúdo
    # seguinte (senão apagaria texto visível legítimo).
    sem_ocultos = re.sub(
        r"<(?:input|meta|link|br|hr)\b[^>]*(?:hidden|display\s*:\s*none)[^>]*>",
        " ", sem_ocultos, flags=re.I,
    )
    # Contêineres ocultos têm conteúdo: removem tag + conteúdo.
    sem_ocultos = re.sub(
        r"<(div|span|p|table|tr|td|section|ul|li|form)\b[^>]*"
        r"(?:hidden|display\s*:\s*none)[^>]*>.*?</\1>",
        " ", sem_ocultos, flags=re.I | re.S,
    )
    return _texto_limpo(sem_ocultos)


def _valor_campo_resposta(pagina: str, sufixo: str) -> str:
    """Lê o valor atual do campo do documento na resposta do servidor."""
    for tag in re.findall(r"<input\b[^>]*>", pagina, re.I):
        atributos = _atributos(tag)
        if (
            atributos.get("id", "").endswith(sufixo)
            or atributos.get("name", "").endswith(sufixo)
        ):
            return str(atributos.get("value", ""))
    return ""


def _sem_acento(texto: str) -> str:
    return "".join(
        caractere for caractere in unicodedata.normalize("NFD", str(texto).lower())
        if unicodedata.category(caractere) != "Mn"
    )


def _para_numero(texto: str) -> float:
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


def _montar_opener(cookie_value: str, cookie_sessao: str):
    jar = http.cookiejar.CookieJar()
    for nome, valor, dominio, caminho in (
        (COOKIE_NAME, cookie_value, DOMINIO_JWT, "/"),
        ("JSESSIONID", cookie_sessao, DOMINIO_SESSAO, "/nfsd"),
    ):
        if not valor:
            continue
        jar.set_cookie(http.cookiejar.Cookie(
            version=0, name=nome, value=valor, port=None, port_specified=False,
            domain=dominio, domain_specified=True,
            domain_initial_dot=dominio.startswith("."),
            path=caminho, path_specified=True, secure=True, expires=None,
            discard=False, comment=None, comment_url=None, rest={},
        ))
    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(jar)
    )
    opener.addheaders = [(
        "User-Agent",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36",
    )]
    return opener


def _tipo_e_documento(documento: str, tipo: str) -> tuple[str, str]:
    digitos = re.sub(r"\D", "", documento or "")
    tipo_norm = re.sub(r"\s+", "", _sem_acento(tipo)).upper()

    # A quantidade de dígitos decide o tipo (CPF=11, CNPJ=14), não o nome da
    # coluna da planilha.
    if tipo_norm in ("IMOVEL", "CIF"):
        return "IMOVEL", "documentoContribuinteCif"
    if len(digitos) == 11:
        return "FISICA", "documentoContribuinteCpf"
    if len(digitos) == 14:
        return "JURIDICA", "documentoContribuinteCnpj"
    if not tipo_norm:
        raise ConsultaNaoVerificada(
            "Não foi possível inferir FÍSICA/JURÍDICA do documento informado."
        )
    mapa = {
        "FISICA": ("FISICA", "documentoContribuinteCpf"),
        "FISICO": ("FISICA", "documentoContribuinteCpf"),
        "CPF": ("FISICA", "documentoContribuinteCpf"),
        "JURIDICA": ("JURIDICA", "documentoContribuinteCnpj"),
        "JURIDICO": ("JURIDICA", "documentoContribuinteCnpj"),
        "CNPJ": ("JURIDICA", "documentoContribuinteCnpj"),
        "IMOVEL": ("IMOVEL", "documentoContribuinteCif"),
    }
    if tipo_norm not in mapa:
        raise ConsultaNaoVerificada(f"Tipo de pessoa desconhecido: {tipo!r}.")
    return mapa[tipo_norm]


def _extrair_tabelas(pagina_html: str, somente_debito: bool = True) -> List[dict]:
    """Extrai débitos das tabelas da resposta HTML."""
    debitos: List[dict] = []
    for tabela in re.findall(r"<table\b[^>]*>(.*?)</table>", pagina_html,
                             re.I | re.S):
        linhas = re.findall(r"<tr\b[^>]*>(.*?)</tr>", tabela, re.I | re.S)
        if len(linhas) < 2:
            continue
        def celulas(linha):
            return [
                _texto_limpo(c)
                for c in re.findall(r"<t[dh]\b[^>]*>(.*?)</t[dh]>", linha,
                                    re.I | re.S)
            ]
        cabecalhos = celulas(linhas[0])
        texto = _sem_acento(" ".join(cabecalhos))
        if not any(palavra in texto for palavra in
                   ("competencia", "divida", "situacao", "origem")):
            continue
        for linha in linhas[1:]:
            valores = celulas(linha)
            if not any(valores):
                continue
            item = {
                cabecalhos[i]: valor
                for i, valor in enumerate(valores) if i < len(cabecalhos)
            }
            def por_chaves(chaves):
                # Prioridade das opções, não das colunas da tabela.
                for opcao in chaves:
                    alvo = _sem_acento(opcao)
                    for chave, valor in item.items():
                        if _sem_acento(chave) == alvo:
                            return valor
                for opcao in chaves:
                    alvo = _sem_acento(opcao)
                    for chave, valor in item.items():
                        if alvo in _sem_acento(chave):
                            return valor
                return ""
            debitos.append({
                "tipo_debito": por_chaves(
                    ["origem", "tipo", "debito", "tributo"]),
                "competencia": por_chaves(["competencia", "ano", "exercicio"]),
                "situacao": por_chaves(["situacao", "status"]),
                "valor": _para_numero(por_chaves(
                    ["atualizado (r$)", "dívida (r$)", "divida (r$)",
                     "total (r$)", "valor", "lancado (r$)"])),
                "cif": por_chaves(["cif"]),
                "id_divida": por_chaves(["id dívida", "id divida"]),
                "vencimento": por_chaves(["vencimento"]),
                "pago": _para_numero(por_chaves(["pago (r$)"])),
            })
    # Mantém somente dívidas em aberto: situação DÉBITO, valor > 0 e nada pago.
    if somente_debito:
        debitos = [
            debito for debito in debitos
            if situacao_eh_debito(debito.get("situacao", ""))
            and debito["valor"] > 0
            and debito["pago"] <= 0
        ]
    return debitos


def consultar_debito_http(
    documento: str,
    nome: str = "",
    cookie_value: str = "",
    cookie_sessao: str = "",
    timeout: int = 30,
    tipo: str = "",
    deve_parar=None, situacoes=None,
) -> List[dict]:
    """Consulta um documento por HTTP puro (sem navegador)."""
    def verificar_parada() -> None:
        if deve_parar is not None and deve_parar():
            raise ConsultaInterrompidaHttp()

    verificar_parada()
    tipo_pessoa, sufixo_documento = _tipo_e_documento(documento, tipo)
    opener = _montar_opener(cookie_value, cookie_sessao)

    def buscar(url: str, dados: bytes | None = None):
        verificar_parada()
        requisicao = urllib.request.Request(url, data=dados)
        with opener.open(requisicao, timeout=timeout) as resposta:
            return resposta.geturl(), resposta.read().decode(
                "utf-8", errors="replace"
            )

    # Abre a home para estabelecer a sessão e depois o relatório.
    try:
        _, _ = buscar(HOME_URL)
        url_final, pagina = buscar(BASE_URL)
        verificar_parada()
    except ConsultaInterrompidaHttp:
        raise  # o botão Parar precisa interromper de verdade
    except Exception as exc:  # noqa: BLE001
        raise ConsultaNaoVerificada(f"Falha ao acessar o site: {exc}") from exc

    if "home.jsf" in url_final.lower() and "relatorio" not in url_final.lower():
        raise ConsultaNaoVerificada(
            "Site redirecionou para a home (sessão sem contribuinte)."
        )

    campos = [(_atributos(tag), tag) for tag in
              re.findall(r"<(?:input|button)\b[^>]*>", pagina, re.I)]
    formularios = [_atributos(tag) for tag in
                   re.findall(r"<form\b[^>]*>", pagina, re.I)]
    if not formularios:
        raise ConsultaNaoVerificada("Formulário não encontrado na resposta.")
    acao = formularios[0].get("action", BASE_URL)
    acao = urllib.parse.urljoin(BASE_URL, acao)

    def campo_por_id(sufixo: str) -> str:
        for atributos, _ in campos:
            if atributos.get("id", "").endswith(sufixo):
                return atributos.get("name", "")
        return ""

    def campo_por_nome(sufixo: str) -> str:
        for atributos, _ in campos:
            if atributos.get("name", "").endswith(sufixo):
                return atributos.get("name", "")
        return ""

    nome_documento = campo_por_id(sufixo_documento)
    nome_ano = campo_por_id("anoInicial_input") or campo_por_nome("anoInicial_input")
    nome_radio = campo_por_nome("tipoPessoaContribuinte")
    nome_botao = campo_por_id("botaoAcaoPesquisar")

    if not nome_documento or not nome_radio or not nome_botao:
        raise ConsultaNaoVerificada(
            "Campos do formulário não encontrados na resposta HTTP — "
            "o site provavelmente exige JavaScript."
        )

    dados: Dict[str, str] = {}
    for atributos, _ in campos:
        if atributos.get("type", "").lower() == "hidden" and atributos.get("name"):
            dados[atributos["name"]] = atributos.get("value", "")

    if "javax.faces.ViewState" not in dados:
        raise ConsultaNaoVerificada(
            "ViewState do JSF não encontrado — não é seguro enviar a consulta."
        )

    digitos = "".join(c for c in (documento or "") if c.isdigit())
    dados[nome_documento] = digitos
    dados[nome_radio] = tipo_pessoa
    if nome_ano:
        dados[nome_ano] = str(datetime.now().year - 9)
    for atributos, _ in campos:
        if (atributos.get("name", "").endswith("situacaoDebito")
                and atributos.get("value") in (situacoes or SITUACOES_ATIVAS)):
            dados[atributos["name"]] = "true"
    dados[nome_botao] = "Pesquisar"

    corpo = urllib.parse.urlencode(dados).encode("utf-8")
    try:
        url_resposta, pagina_resposta = buscar(acao, corpo)
        verificar_parada()
    except ConsultaInterrompidaHttp:
        raise  # o botão Parar precisa interromper de verdade
    except Exception as exc:  # noqa: BLE001
        raise ConsultaNaoVerificada(f"Falha ao enviar a consulta: {exc}") from exc

    if "home.jsf" in url_resposta.lower() and "relatorio" not in url_resposta.lower():
        raise ConsultaNaoVerificada(
            "A consulta foi redirecionada para a home — sessão inválida."
        )

    # Confirmação do documento e do contribuinte na resposta: sem isso, um
    # resultado de outra sessão/CPF poderia ser publicado como deste CPF.
    valor_documento = _valor_campo_resposta(pagina_resposta, sufixo_documento)
    if valor_documento and "".join(c for c in valor_documento if c.isdigit()) != digitos:
        raise ConsultaNaoVerificada(
            f"A resposta está com o documento {valor_documento!r}, mas foi "
            f"pedido {digitos} — resultado descartado."
        )
    visivel = _texto_visivel(pagina_resposta)
    if nome.strip():
        nome_norm = _sem_acento(nome)
        visivel_norm = _sem_acento(visivel)
        partes = [p for p in re.findall(r"[a-z]{4,}", nome_norm)][:3]
        if nome_norm not in visivel_norm and not (
            partes and all(parte in visivel_norm for parte in partes)
        ):
            raise ConsultaNaoVerificada(
                "O nome do contribuinte da resposta não confere com a "
                "planilha — resultado descartado."
            )

    texto = visivel.lower()
    if any(trecho in texto for trecho in (
        "nenhum registro", "nenhum debito", "nenhuma guia", "sem resultados",
        "nao foram encontrados", "não foram encontrados", "nenhum lancamento",
        "nenhum lançamento", "no records found",
    )):
        return []

    debitos_brutos = _extrair_tabelas(pagina_resposta, somente_debito=False)
    if not debitos_brutos:
        raise ConsultaNaoVerificada(
            "A resposta HTTP não trouxe tabela de resultados nem mensagem "
            "de 'nenhum registro' — não é seguro afirmar que não há débitos."
        )
    debitos = _extrair_tabelas(pagina_resposta)
    if not debitos:
        # Havia linhas, mas nenhuma era DÉBITO confirmado. Dívida ativa,
        # ajuizado, pago etc. não podem virar "sem débitos".
        raise ConsultaNaoVerificada(
            "A resposta trouxe linhas, mas nenhuma com situação DÉBITO "
            "confirmado — não é seguro afirmar que não há débitos."
        )
    return debitos
