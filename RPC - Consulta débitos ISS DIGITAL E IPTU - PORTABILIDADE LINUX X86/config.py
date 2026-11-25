"""Configurações usadas pelo programa (main.py)."""

# URL do relatório de débitos e da home (a home "aquece" a sessão autenticada)
BASE_URL = "https://santanadoaraguaia-pa.desenvolvecidade.com.br/nfsd/pages/relatorio/relatorioLancamentoDebito.jsf"
HOME_URL = "https://santanadoaraguaia-pa.desenvolvecidade.com.br/nfsd/home.jsf"

# Depuração remota do Brave: permite reaproveitar a sessão já logada
CDP_URL = "http://127.0.0.1:9222"

# Cookie JWT lido do perfil do Brave (valor nunca é gravado em arquivo)
COOKIE_NAME = "__Secure-DA"

# Intervalo entre consultas, em segundos, para não sobrecarregar o site
EXTRACTION_DELAY = 1.5

# Cadastro Imobiliário do IPTU Digital: consulta por CPF/CNPJ que devolve os
# imóveis (CIF) com endereço — usada para verificar se o CIF é real.
IPTU_URL = "https://santanadoaraguaia-pa.desenvolvecidade.com.br/iptu/#/cadastroImobiliario"
