# RPC de Consulta de Débitos - Sistema de Automação

## 📋 Descrição

Sistema para consultar e extrair dados de débitos de múltiplos devedores do site do Santana de Araçuaia-PA e gerar planilhas organizadas.

## 🎯 Objetivo

Automatizar a consulta de mais de 3500 CPFs e extrair informações sobre débitos (tipo, competência, valor) para uma planilha consolidada.

## 📦 Recursos

- **Interface Gráfica Completa (PyQt6)**: GUI intuitiva para upload e gestão de planilhas
- **Scraper com Playwright**: Automação de navegador para consulta de dados
- **Processamento em Lotes**: Consulta paralela de múltiplos CPFs
- **Gestão de Erros**: Log separado para erros de consulta
- **Exportação de Resultados**: Planilha XLSX/CSV com todos os dados

## 🚀 Instalação

### 1. Instalar dependências Python

```bash
pip install -r requirements.txt
```

### 2. Instalar Playwright e browsers

```bash
playwright install
playwright install chromium
```

### 3. Ou usar o script de instalação

```bash
chmod +x install.sh
./install.sh
```

## 📁 Estrutura do Projeto

```
RPC - Consulta de débitos/
├── main.py                          # GUI principal com interface gráfica
├── analyze.py                       # Script para analisar estrutura da página
├── config.py                        # Configurações do sistema
├── models.py                        # Modelos de dados
├── database.py                      # Gerenciamento de planilhas
├── scraper/
│   └── scraper.py                  # Lógica de scraping com Playwright
├── install.sh                       # Script de instalação
├── requirements.txt                 # Dependências
├── errors.log                       # Log de erros (gerado automaticamente)
├── resultados_*.xlsx               # Planilhas de resultado (geradas)
└── README.md                        # Este arquivo
```

## 💻 Como Usar

### 1. Iniciar o Sistema

```bash
python main.py
```

### 2. Carregar Planilha de Entrada

- Clique em "Escolher planilha..."
- Selecione um arquivo CSV ou XLSX com a lista de CPFs e nomes
- O sistema mostrará uma prévia dos dados

### 3. Configurar Parâmetros (Opcional)

- Número de lotes por consulta (padrão: 50)
- Timeout por consulta (padrão: 30s)
- Número de tentativas de retry

### 4. Executar Consulta

- Clique em "Executar Consulta"
- O sistema processará os débitos em paralelo
- Barra de progresso mostrará o status em tempo real
- Log mostrará detalhes de cada consulta

### 5. Verificar Resultados

- Após conclusão, o sistema mostrará o resultado
- Planilha de saída será gerada automaticamente
- Erros serão registrados em `errors.log`

## 📊 Formato da Planilha de Entrada

A planilha de entrada deve conter as seguintes colunas:

| Coluna | Descrição | Exemplo |
|--------|-----------|---------|
| Nome | Nome completo do devedor | João Silva |
| CPF | CPF do devedor | 123.456.789-00 |
| Tipo Débito | Tipo de débito (opcional) | ICMS |
| Competência | Ano/competência (opcional) | 2024 |
| Valor (R$) | Valor do débito (opcional) | 150.00 |

## 📊 Formato da Planilha de Saída

A planilha de saída conterá as seguintes colunas:

| Coluna | Descrição | Exemplo |
|--------|-----------|---------|
| Nome | Nome completo do devedor | João Silva |
| CPF | CPF do devedor | 123.456.789-00 |
| Tipo de Débito | Tipo de débito encontrado | ICMS |
| Competência | Ano/competência | 2024 |
| Valor (R$) | Valor do débito | 150.00 |
| Status | Status da consulta | COMPLETO/ERRO |

## 🔧 Configurações Customizáveis

### Configurar no arquivo `config.py`

```python
# URL base do sistema
BASE_URL = "https://santanadoaraguaia-pa.desenvolvecidade.com.br/nfsd/pages/relatorio/relatorioLancamentoDebito.jsf"

# Configurações de scraping
MAX_CONCURRENT_REQUESTS = 50
TIMEOUT_PER_REQUEST = 30
RETRY_ATTEMPTS = 3

# Configurações de lote
LOTES_POR_CONSULTA = 50

# Timeout do navegador
BROWSER_TIMEOUT = 30
```

### Elementos do DOM

Após análise da página, os elementos do DOM podem ser configurados:

```python
# Elementos do formulário
FORM_ID = "formulario_id"
CPF_INPUT_ID = "cpf_input_id"
NOME_INPUT_ID = "nome_input_id"
BUSCAR_BUTTON_ID = "botao_pesquisar_id"

# Tabelas de resultados
TABLE_SELECTOR = "table(selector)"
CPF_COL_INDEX = 0
NOME_COL_INDEX = 1
TIPO_COL_INDEX = 2
COMPETENCIA_COL_INDEX = 3
VALOR_COL_INDEX = 4
```

## ⚠️ Considerações Importantes

1. **Autenticação**: O sistema funciona com acesso direto à página (sem login)
2. **Rate Limiting**: Configurar timeout e retry para evitar bloqueios
3. **Cookies/Sessão**: Playwright gerencia sessões automaticamente
4. **JavaScript**: Playwright executa JavaScript da página automaticamente
5. **Imagens**: O sistema não extrai imagens (apenas dados estruturados)

## 🛠️ Desenvolvimento

### Análise da Página Web

Para analisar a estrutura da página e identificar elementos do DOM:

```bash
python analyze.py
```

Este script mostrará:
- Elementos do formulário
- Campos de input (CPF, Nome)
- Botões (Pesquisar)
- Tabelas de resultados
- Cabeçalhos das tabelas
- Classes e IDs dos elementos

### Customização do Scraper

Após análise, atualize o scraper.py com os elementos corretos:

```python
def _preencher_formulario(self, page, cpf, nome):
    cpf_input = page.query_selector("input[name='cpf']")
    cpf_input.fill(cpf)
    
    nome_input = page.query_selector("input[name='nome']")
    nome_input.fill(nome)
```

## 📝 Log de Erros

Todos os erros de consulta são registrados em `errors.log` com:
- Data e hora
- CPF do devedor
- Nome do devedor
- Detalhes do erro

## 🔄 Versionamento

Versão: 1.0.0

Data de criação: 2024

## 📞 Contato

Desenvolvido para consulta automatizada de débitos
