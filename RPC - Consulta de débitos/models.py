"""
Modelos de dados para o sistema de consulta de débitos
"""
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, List
from enum import Enum


class StatusConsulta(Enum):
    PENDENTE = "pendente"
    PROCESSANDO = "processando"
    COMPLETO = "completo"
    ERRO = "erro"


@dataclass
class Debito:
    """Dados básicos do devedor"""
    cpf: str
    nome: str


@dataclass
class TipoDebito:
    """Informações sobre o tipo de débito"""
    tipo: str
    competencia: str
    valor: float


@dataclass
class ResultadoDebito:
    """Resultado completo de uma consulta"""
    cpf: str
    nome: str
    tipo_debito: str
    competencia: str
    valor: float
    consultado_em: datetime = field(default_factory=datetime.now)
    status: StatusConsulta = StatusConsulta.PENDENTE


@dataclass
class ErroConsulta:
    """Erro de consulta"""
    cpf: str
    nome: str
    erro: str
    consultado_em: datetime = field(default_factory=datetime.now)


@dataclass
class ConsultaBatch:
    """Consulta de lote"""
    id: int
    cpf: str
    nome: str
    status: StatusConsulta = StatusConsulta.PENDENTE
    resultado: Optional[ResultadoDebito] = None
    erro: Optional[str] = None
    consultado_em: datetime = field(default_factory=datetime.now)


# Estrutura de dados para planilha
PLANILLA_CABECALOS = [
    "Nome",
    "CPF",
    "Tipo de Débito",
    "Competência",
    "Valor (R$)",
    "Status"
]
