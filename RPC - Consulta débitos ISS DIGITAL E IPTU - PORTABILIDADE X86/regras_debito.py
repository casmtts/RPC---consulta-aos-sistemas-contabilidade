"""Regra única do projeto: o que conta como débito cobrável confirmado.

Centralizar esta decisão evita que um módulo (consulta, OBSERVAÇÃO, fallback
legado, verificação de pagos) aceite como cobrável algo que outro módulo
rejeitou. Na dúvida, a resposta é **não** — nunca gerar cobrança indevida.
"""

from __future__ import annotations

import re
import unicodedata

def mensagem_amigavel_erro(texto: object) -> str:
    """Troca mensagens técnicas por texto simples para quem lê o relatório."""
    original = str(texto or "").strip()
    if not original:
        return ""
    baixo = original.lower()
    if (
        "javascript" in baixo or "campos do formulário" in baixo
        or "campos do formulario" in baixo
    ):
        return "Não foi possível consultar o sistema no momento — tente novamente mais tarde."
    if any(code in baixo for code in ("503", "502", "504", "500")):
        return "Sistema temporariamente indisponível — tente novamente mais tarde."
    if "sessão expirada" in baixo or "sessao expirada" in baixo or "404" in baixo:
        return "Sessão do sistema expirada — atualize os cookies e refaça a consulta."
    if any(x in baixo for x in ("timeout", "connection refused", "unreachable", "net::")):
        return "Site inalcançável no momento — tente novamente mais tarde."
    return original


# "COBRÁVEL" é gravado por formatar_parcial() quando o site confirma dívida
# viva, com valor devido positivo e sem pagamento/encerramento. Também é um
# débito confirmado: precisa continuar alimentando a OBSERVAÇÃO quando a
# planilha de resultados já sai classificada (com Status colorido e MOTIVO).
STATUS_CONFIRMADOS = ("OK", "VERIFICADO", "COBRÁVEL")
# Nome neutro exibido na interface para débitos que ainda precisam de
# confirmação. Troque AQUI quando o nome definitivo for informado.
STATUS_PENDENTE = "CLASSIFICADO — INDETERMINADO"
# Consulta não concluída (sessão expirada, site fora, resposta duvidosa).
# Nunca pode conviver com COBRÁVEL: sem confirmação integral, não há cobrança.
STATUS_ERRO_CONSULTA = "ERRO DE CONSULTA"
TIPOS_SEM_DEBITO = ("-", "SEM DEBITOS", "SEM DÉBITOS")


def sem_acento(texto: object) -> str:
    valor = unicodedata.normalize("NFKD", str(texto or ""))
    return "".join(
        caractere for caractere in valor
        if not unicodedata.combining(caractere)
    ).lower().strip()


def situacao_eh_debito(situacao: object) -> bool:
    """True somente para a situação exata DÉBITO (não dívida ativa/pago)."""
    return sem_acento(situacao) == "debito"


def eh_debito_cobravel(item: dict) -> bool:
    """True somente para um débito real, confirmado e sem ressalva.

    Rejeita explicitamente: status diferente de OK/VERIFICADO, qualquer
    texto de erro/ressalva (inclusive "relatório antigo"), tipo "Sem
    débitos" ou "-", competência vazia/"-" e débito sem valor positivo
    informado quando o campo existe.
    """
    status = str(item.get("Status", "") or "").strip().upper()
    if status not in STATUS_CONFIRMADOS:
        return False
    if str(item.get("Erro", "") or "").strip():
        return False

    tipo = str(item.get("Tipo de Débito", "") or "").strip()
    competencia = str(item.get("Competência", "") or "").strip()
    if not tipo or not competencia:
        return False
    tipo_norm = sem_acento(tipo)
    if tipo_norm in TIPOS_SEM_DEBITO or tipo_norm.startswith("sem debitos"):
        return False
    if re.fullmatch(r"-{1,}", tipo.strip()) or re.fullmatch(r"-{1,}", competencia.strip()):
        return False
    return True
