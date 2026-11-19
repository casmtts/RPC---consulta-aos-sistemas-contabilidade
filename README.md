# RPC — Consulta aos Sistemas de Contabilidade

Conjunto de ferramentas para **consultar débitos em lote** no sistema NFS-e do
município de **Santana do Araguaia-PA**, a partir de uma planilha de
CPFs/CNPJs/CIFs, com interface gráfica, exportação dos resultados e retomada
de consultas interrompidas.

O nome "RPC" vem do fluxo de trabalho usado na rotina de contabilidade: rodar
uma **consulta automatizada ao sistema de débitos** (RPC = *Remote Procedure
Call*, no sentido de consulta ao sistema) e consolidar o resultado em planilha.

> Este repositório contém **somente o código-fonte**. Planilhas, CSVs,
> relatórios, credenciais e ambientes virtuais **não** são versionados
> (veja `.gitignore`).

---

## Sumário

- [Versões incluídas](#versões-incluídas)
- [O que o programa faz](#o-que-o-programa-faz)
- [Requisitos](#requisitos)
- [Instalação](#instalação)
- [Fluxo de uso](#fluxo-de-uso)
- [Planilha de entrada](#planilha-de-entrada)
- [Resultado e status](#resultado-e-status)
- [Parar, retomar e revisar](#parar-retomar-e-revisar)
- [Sessão, cookies e token de autenticação](#sessão-cookies-e-token-de-autenticação)
- [Modo com e sem navegador](#modo-com-e-sem-navegador)
- [Tema e atalhos](#tema-e-atalhos)
- [Estrutura dos arquivos](#estrutura-dos-arquivos)
- [Solução de problemas](#solução-de-problemas)
- [Privacidade e segurança](#privacidade-e-segurança)

---

## Versões incluídas

| Pasta | Para que serve |
|-------|----------------|
| `RPC - Consulta de débitos/` | Versão original (**contabilidade**). Base histórica do projeto, com GUI PyQt6, scraper, modelos e banco de planilhas. |
| `RPC - Consulta de débitos - REVISADO/` | **Versão recomendada** para Macs Apple Silicon. Corrigida e evoluída: retomada, Revisar inteligente, renovação de token, temas, retry automático. |
| `RPC - Consulta de débitos - PORTABILIDADE X86/` | Mesmo programa revisado, preparado para **Mac Intel (x86_64)**, com virtualenv e browsers do Playwright separados (evita erro `bad CPU type`). |

Use **REVISADO** em Apple Silicon e **PORTABILIDADE X86** em Intel. A pasta
base serve como referência do projeto original.

Cada versão traz também a pasta `cpf_cnpj/` com utilitários independentes para
consultar situação cadastral de CNPJ ( BrasilAPI / ReceitaWS ), em GUI e terminal.

---

## O que o programa faz

1. Lê uma planilha (`.xlsx`, `.xlsm` ou `.csv`) com CPFs, CNPJs ou CIFs.
2. Abre o sistema de débitos usando a **sessão já logada no Brave** (via CDP,
   porta `9222`) ou um Chromium invisível com os cookies do Brave.
3. Consulta cada devedor, aplicando o filtro de situação **somente DÉBITO**
   (dívida ativa, ajuizado, pago, cancelado e prescrito ficam de fora).
4. Grava os resultados em uma tabela agrupada por devedor (cores alternadas)
   e exporta em **XLSX ou CSV**.
5. Salva progresso parcial a cada registro, permitindo **retomar** e
   **revisar somente o que falta** após uma interrupção.

Saídas incluem `CIF`, `ID Dívida` e `Vencimento`, para identificar, por
exemplo, qual imóvel (CIF) está devendo IPTU.

---

## Requisitos

- **macOS** (o leitor de cookies do Brave usa Keychain + banco de cookies do Chromium).
- **Python 3.10+** (3.9 na versão x86).
- Dependências: `PyQt6`, `playwright`, `openpyxl`, `cryptography`
  (ver `requirements.txt` / `requirements-x86.txt`).
- **Brave** com o sistema aberto e logado (ou Chromium headless com os cookies).
- Chromium do Playwright (instalado pelo próprio programa).

---

## Instalação

### Versão REVISADO (Apple Silicon)

Duplo clique em `Consulta de Débitos.app` (ou em
`Abrir Consulta de Débitos.command`). O programa escolhe um Python com as
dependências e instala o que faltar.

Se o macOS bloquear: **botão direito → Abrir → Abrir**.

Pelo terminal:

```bash
cd "RPC - Consulta de débitos - REVISADO"
pip install -r requirements.txt
python -m playwright install chromium
python3 main.py
```

### Versão PORTABILIDADE X86 (Intel)

1. Duplo clique em `Instalar dependências x86.command`
   (ou `bash install_x86.sh`).
2. Duplo clique em `Consulta de Débitos (x86).app`.

O instalador cria `.venv-x86/` e instala browsers em
`~/Library/Caches/ms-playwright-x86`, separados da instalação arm64.

---

## Fluxo de uso

1. Aba **Ambiente** → **Instalar tudo** (primeira vez) e
   **Abrir Brave com depuração (9222)**.
2. No Brave, entre no sistema normalmente e **deixe a aba logada**.
3. Volte ao programa → aba **Ambiente** → **Atualizar cookies**.
4. Aba **Consulta** → **Escolher planilha...** → confira a aba → **Executar**.
5. Acompanhe a barra de progresso e a aba **Log**.
6. Aba **Resultados** → **Revisar** (se houver pendências) →
   **Salvar resultado...** (XLSX ou CSV).

Dicas:
- **Pausa entre consultas** ajustável; **Variar a pausa** aplica oscilação
  aleatória de ±30% para não criar padrão detectável.
- O filtro de situação considera **somente DÉBITO**.
- O tipo de pessoa é inferido pela quantidade de dígitos, mesmo que o
  cabeçalho da coluna esteja errado (um CPF em coluna "CNPJ" não vira JURÍDICA).

---

## Planilha de entrada

Aceita `.xlsx`, `.xlsm` e `.csv`. O cabeçalho é detectado automaticamente,
mesmo abaixo de títulos, e as abas são listadas para escolha.

| Coluna | Descrição | Obrigatória |
|--------|-----------|-------------|
| `Nome` | Nome do devedor | Não |
| `CPF` | CPF (11 dígitos) | Sim, ou CNPJ/CIF |
| `CNPJ` | CNPJ (14 dígitos) | Sim, ou CPF |
| `CIF` | CIF do imóvel | Sim, se for consulta de IMÓVEL |
| `Tipo de Pessoa` | `FÍSICA`, `JURÍDICA` ou `IMOVEL` | Não (inferido) |

---

## Resultado e status

Colunas exportadas: **Nome, CPF, Tipo de Débito, Competência, Valor (R$),
Status, CIF, ID Dívida, Vencimento e Erro**.

| Status | Significado |
|--------|-------------|
| `VERIFICADO` | Consulta concluída com sucesso (com débitos ou "Sem débitos"). |
| `NAO VERIFICADO` | O site não permitiu confirmar o resultado (erro de sessão, 404, página de erro, dúvida). **Nunca** é tratado como "sem débitos". |
| `ERRO` | Falha inesperada na consulta. |

**Competência:** quando o site oferece o filtro **"Pesquisar por: competência /
lançamento / vencimento"**, o programa escolhe **Competência** antes de
consultar, para que o próprio site devolva a competência original do débito
(ex.: `fev/22`) em vez da referência da cobrança atual (`1/2026`). Quando o
filtro é aplicado, o log registra: *"Filtro do site 'Pesquisar por: Competência'
aplicado"*. Se o filtro não existir na tela, o programa usa a `Competência
Débito` (ou `Competência Guia`) da planilha de entrada como fallback. Na
**revisão** de itens já consultados, a competência sempre vem da planilha — a
data em que o débito foi gerado (ex.: `2022`) —, corrigindo os resultados
antigos que mostravam a data da consulta em vez do período original.

Os relatórios — tanto o **salvo automaticamente** ao parar quanto o gerado
pelo botão **Salvar resultado...** — são gravados na pasta **`relatorios/`**,
criada automaticamente dentro da pasta do programa. A planilha de entrada e o
arquivo de progresso continuam no local original.

Descrições especiais em `Tipo de Débito`:

- `Sem débitos` — pesquisa executada e confirmada.
- `Sem débitos (NÃO CONFIRMADO)` — o site mostrou ausência de débitos, mas o
  documento não pôde ser 100% confirmado. **Sempre vai para o Revisar**,
  porque a planilha de entrada é uma lista de devedores.
- `(M)` no final (ex.: `IPTU (M)`) — débito **migrado**. O programa preserva a
  marcação que já vem do site e também acrescenta `(M)` quando a coluna
  `Origem` da planilha de entrada indica migração (ex.: `IPTU (M)`).

---

## Parar, retomar e revisar

- **Parar:** pressione **ESC duas vezes** (1º aperto muda a cor dos botões e
  do aviso; 2º aperto interrompe) ou clique em **Parar**. O item atual termina
  e o progresso é salvo.
- **Retomar:** ao abrir, a última planilha é reaberta automaticamente e os
  resultados salvos aparecem na aba Resultados. Clicando em **Executar**, o
  programa pergunta se deseja continuar. O último registro é sempre refeito.
- **Revisar:** disponível já na abertura (com planilha carregada). Após parar,
  consulta **somente o que falta** — registros sem resultado, itens
  `NÃO CONFIRMADO`, `NAO VERIFICADO` e erros. Débitos confirmados não são
  repetidos. Se a revisão também for interrompida, um novo clique continua de
  onde parou.
- **Revisão paralela:** marque **Revisar pendências em paralelo** na aba
  Consulta. Ao clicar em **Executar**, um segundo worker reconfulta ao mesmo
  tempo os itens já pendentes (erro, `NAO VERIFICADO`, sem débitos). Os
  resultados revisados ficam isolados e só entram na tabela quando as duas
  consultas terminam, evitando conflito entre elas. Os documentos pendentes
  são **excluídos da consulta principal** — o worker principal os ignora —,
  então cada rotina atua num conjunto distinto de CPFs e nunca consulta o
  mesmo documento ao mesmo tempo. As linhas em revisão aparecem com **fundo
  laranja de faixas diagonais** (aspecto de “em obra”) e voltam às cores
  normais quando a revisão termina. O andamento aparece no rótulo **Revisão
  paralela: ...** na aba Consulta — inclusive a mensagem *sem pendências para
  revisar*, que explica quando não há itens pendentes para a segunda consulta.
  O **ESC 2x** / botão **Parar** encerra as duas consultas.

---

## Sessão, cookies e token de autenticação

O programa **não armazena senhas nem cookies em arquivo**. Ele lê, somente em
memória, os dois cookies do sistema no perfil do Brave:

- `__Secure-DA` — JWT de autenticação;
- `JSESSIONID` — sessão JSF (perfil selecionado no servidor).

Renovação automática:

- Se falta **menos de 1 minuto** para o JWT expirar, o token é relido do Brave
  antes da próxima consulta — **sem interromper o trabalho**.
- Se uma consulta falha por **token/sessão expirado** (404, página de erro
  `13429`, "Sessão expirada", erro de autenticação), o programa pega o token
  novo do Brave e **refaz aquela consulta** (até 3 tentativas).
- Se todas as tentativas falharem, o item fica `NAO VERIFICADO` e entra no
  Revisar; o log orienta a entrar no sistema novamente no Brave.

> O programa não tem credenciais para fazer login sozinho. Manter a aba do
> sistema aberta e logada no Brave é o que permite a renovação do token.

---

## Modo com e sem navegador

- **Padrão (com navegador):** Playwright. Se o Brave estiver com CDP ativo
  (porta 9222), o programa usa uma aba dele; caso contrário, abre um Chromium
  invisível com os cookies do Brave. É o modo mais confiável — o site é
  JSF/PrimeFaces e exige JavaScript.
- **"Tentar SEM navegador (HTTP direto)":** experimental (`consulta_http.py`).
  Reenvia o formulário com `javax.faces.ViewState` e lê a tabela da resposta.
  Se o resultado não puder ser confirmado, o item fica `NAO VERIFICADO` —
  nunca um falso "sem débitos".

Erros conhecidos tratados com retry: **HTTP 404** e **página de erro
(código 13429)**, ambos geralmente causados por viewstate/sessão expirada.

---

## Tema e atalhos

- Botão no canto superior direito alterna **claro/escuro**; a preferência fica
  salva. Tema escuro: preto, detalhes laranja e letras brancas.
- **ESC 2x** — parar consultas (1º aperto avisa mudando a cor; 2º executa).
- **Ambiente → Atualizar cookies** — relê manualmente JWT e JSESSIONID do Brave.

---

## Estrutura dos arquivos

```
RPC - Consulta de débitos - REVISADO/
├── main.py              # GUI (PyQt6), planilhas, worker e instalação
├── config.py            # URLs, porta CDP, cookie e pausa entre consultas
├── brave_cookie.py      # leitura segura dos cookies do Brave + validade do JWT
├── consulta_http.py     # modo experimental sem navegador
├── requirements.txt     # dependências
├── GUIDE_DE_USO.txt     # guia rápido
├── RELATORIO_REVISAO.md # histórico de correções
├── Abrir Consulta de Débitos.command
└── cpf_cnpj/            # utilitários de situação cadastral de CNPJ
```

A versão X86 acrescenta `install_x86.sh`, `requirements-x86.txt` e o `.app`
próprio. A versão base contém ainda `models.py`, `database.py`, `scraper/` e
scripts de análise.

---

## Solução de problemas

| Sintoma | O que fazer |
|---------|-------------|
| `bad CPU type` | Você está na arquitetura errada: use REVISADO (Apple Silicon) ou X86 (Intel). |
| "Cookie não encontrado no Brave" | Abra o sistema no Brave, faça login e clique em **Atualizar cookies**. |
| `NAO VERIFICADO ... status 404` | Sessão/viewstate expirada: o programa já tenta 3x com token novo; se persistir, entre no sistema novamente e use **Revisar**. |
| `página de erro (código 13429/13414)` | Mesmo tratamento acima: token renovado e consulta refeita automaticamente. |
| `Botão 'Pesquisar' não encontrado` | O programa abre sessão nova e repete a consulta (até 3x). Se persistir, confira se a aba do sistema está logada no Brave. |
| `Page.goto: net::ERR_ABORTED` | Navegação abortada também dispara nova tentativa com sessão nova. |
| "Playwright Sync API inside the asyncio loop" | Corrigido: cada sessão roda em thread limpa. |
| macOS não abre o `.app` | Botão direito → Abrir → Abrir (app assinado localmente, sem certificado Apple). |
| Brave não inicia com depuração | Aba **Ambiente** → **Abrir Brave com depuração (9222)**. |

Logs: `~/Library/Logs/ConsultaDebitos.log` (REVISADO) e
`~/Library/Logs/ConsultaDebitos-x86.log` (X86).

---

## Privacidade e segurança

- Nenhuma cópia do banco de cookies é gravada em disco (abertura somente-leitura).
- A consulta SQL filtra por **nomes e host exatos** — nenhum outro cookie é lido.
- A chave do Keychain ("Brave Safe Storage") fica apenas em memória.
- Tokens e cookies nunca são enviados a terceiros.
- Planilhas, CSVs e relatórios com dados de devedores **não** entram neste
  repositório.
