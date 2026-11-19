"""
Configurações do sistema de consulta de débitos
"""

# URL base do sistema
BASE_URL = "https://santanadoaraguaia-pa.desenvolvecidade.com.br/nfsd/pages/relatorio/relatorioLancamentoDebito.jsf"

# URL da home — o site redireciona para ela na PRIMEIRA requisição de uma
# sessão nova (contexto sem JSESSIONID). Visitar a home primeiro "aquece" a
# sessão autenticada e permite acessar o relatório diretamente depois.
HOME_URL = "https://santanadoaraguaia-pa.desenvolvecidade.com.br/nfsd/home.jsf"

# Endereço de depuração do Brave. Quando disponível, o programa usa a sessão
# real do navegador do usuário (a mesma em que o "Pesquisar" funciona), em vez
# de um navegador isolado com cookies copiados.
CDP_URL = "http://127.0.0.1:9222"

# Configurações de cookie de sessão (para manter login)
COOKIE_NAME = "__Secure-DA"  # Nome do cookie de sessão
COOKIE_VALUE = ""  # Cole o valor do cookie aqui (vazio até você preencher)

# Parâmetros de scraping
MAX_CONCURRENT_REQUESTS = 50
TIMEOUT_PER_REQUEST = 30
EXTRACTION_DELAY = 1.5  # segundos entre consultas (1,5s)

# Configurações de planilha
INPUT_FILE_EXTENSIONS = [".csv", ".xlsx", ".xls"]
OUTPUT_FILE_EXTENSION = ".xlsx"
ERROR_LOG_FILE = "errors.log"

# Configurações de lote
LOTES_POR_CONSULTA = 50

# Timeout do navegador
BROWSER_TIMEOUT = 30

# Elementos do DOM (serão atualizados após análise)
# Formulário
FORM_ID = None
CPF_INPUT_ID = None
NOME_INPUT_ID = None
BUSCAR_BUTTON_ID = None

# Tabelas de resultados
TABLE_SELECTOR = None
CPF_COL_INDEX = None
NOME_COL_INDEX = None
TIPO_COL_INDEX = None
COMPETENCIA_COL_INDEX = None
VALOR_COL_INDEX = None
