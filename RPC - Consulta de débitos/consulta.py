"""Ponto de integração com o site de débitos.

A consulta real (Playwright: login, busca, extração) entra aqui na próxima etapa.
Por enquanto, consultar() falha de forma explícita, para que a interface
registre cada item como ERRO em vez de inventar resultados.
"""


class ConsultaErro(Exception):
    """Falha esperada de consulta; a mensagem vai para a coluna 'erro'."""


def consultar(registro: dict[str, str]) -> list[dict[str, str]]:
    """Consulta os débitos de um registro da planilha.

    Retorna uma lista com um dicionário por débito, com as chaves:
    tipo_debito, ano, situacao. Lista vazia significa "sem débitos".
    """
    raise ConsultaErro("Consulta ao site ainda não implementada")
