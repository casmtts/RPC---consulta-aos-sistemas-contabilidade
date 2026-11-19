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
2. Abre o sistema de débitos usando a **sessão já logada no navegador (Brave, Chrome, Edge, Vivaldi ou Chromium)** (via CDP,
   porta `9222`) ou um Chromium invisível com os cookies do navegador.
3. Consulta cada devedor, aplicando o filtro de situação **somente DÉBITO**
   (dívida ativa, ajuizado, pago, cancelado e prescrito ficam de fora).
4. Grava os resultados em uma tabela agrupada por devedor (cores alternadas)
   e exporta em **XLSX ou CSV**.
5. Salva progresso parcial a cada registro, permitindo **retomar** e
   **revisar somente o que falta** após uma interrupção.
6. Para **CPFs fictícios do sistema legado** (`000.000.xxx-xx`, que o site
   rejeita), localiza os débitos no relatório antigo
   `relatorio-lancamentos-debito-divida-ativa-*.csv` (procurado na pasta do
   programa, na pasta pai ou no diretório atual) e preenche o resultado com
   esses débitos — sem consultar o site e sem marcar como inválido.

Saídas incluem `CIF`, `ID Dívida` e `Vencimento`, para identificar, por
exemplo, qual imóvel (CIF) está devendo IPTU.

---

## Requisitos

- **macOS** (o leitor de cookies usa Keychain + banco de cookies do Chromium).
- **Python 3.10+** (3.9 na versão x86).
- Dependências: `PyQt6`, `playwright`, `openpyxl`, `cryptography`
  (ver `requirements.txt` / `requirements-x86.txt`).
- **Brave, Chrome, Edge, Vivaldi ou Chromium** com o sistema aberto e logado — o programa detecta qual está logado.
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
   **Abrir navegador com depuração (9222)**.
2. No navegador, entre no sistema normalmente e **deixe a aba logada**.
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

As linhas são sempre ordenadas pela posição da guia na **planilha de entrada**
(índice `_indice`), inclusive as que vêm da revisão paralela. Assim, o
primeiro nome do relatório bate com o primeiro nome da planilha enviada para
análise — a ordem não muda quando a revisão mescla seus resultados. Os índices
de progressos antigos também são reparados na abertura: cada CPF volta para a
linha real que ocupa na planilha.

| Status | Significado |
|--------|-------------|
| `VERIFICADO` | Consulta concluída com sucesso (com débitos ou "Sem débitos"). |
| `NAO VERIFICADO` | O site não permitiu confirmar o resultado (erro de sessão, 404, página de erro, dúvida). **Nunca** é tratado como "sem débitos". |
| `ERRO` | Falha inesperada na consulta. |
| `DOCUMENTO INVÁLIDO` | CPF/CNPJ com dígito verificador errado. Não é consultado e **não entra no Revisar** — refazer não resolve; corrija a planilha. |

**Cores da tabela de resultados:**

| Cor | Onde | Significado |
|-----|------|-------------|
| **Vermelho escuro** | texto da coluna `Status` | `ERRO` — a consulta falhou; o item entra no Revisar. |
| **Laranja** | texto da coluna `Status` | `NAO VERIFICADO` — o site não confirmou o resultado; entra no Revisar. |
| **Neutro/cinza** | texto da coluna `Status` | `DOCUMENTO INVÁLIDO` — erro no CPF/CNPJ da planilha; não entra no Revisar. |
| **Verde** | texto da coluna `Status` | `VERIFICADO` — consulta concluída. |
| **Laranja/amarelo com faixas diagonais** | fundo da linha inteira | **Documento em revisão naquele instante** (faixa de “em obra”). Some quando a revisão passa para o próximo. Não é erro — é só destaque visual. |

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
  programa pergunta se deseja continuar. Se você clicar em **Não**, aparece
  uma confirmação **“Tem certeza?”** antes de descartar o progresso — o botão
  **Voltar** mantém a retomada. O último registro é sempre refeito.
- **Revisar:** disponível já na abertura (com planilha carregada). Após parar,
  consulta **somente o que falta** — registros sem resultado, itens
  `NÃO CONFIRMADO`, `NAO VERIFICADO` e erros. Débitos confirmados não são
  repetidos. Se a revisão também for interrompida, um novo clique continua de
  onde parou.
- **Revisão paralela:** marque **Revisar pendências em paralelo** na aba
  Consulta. A revisão começa **na hora em que você marca o checkbox**, mesmo
  com a consulta já rodando — não é preciso parar e reiniciar. Se marcar antes
  de **Executar**, ela inicia junto com a consulta. Um segundo worker reconfulta
  os itens já pendentes (erro, `NAO VERIFICADO`, sem débitos) enquanto a
  consulta principal varre o restante da planilha. Desmarcar o checkbox
  interrompe a revisão (o progresso é mantido).

  - Os dois workers compartilham o mesmo `JSESSIONID` do navegador. Se
    consultassem o site ao mesmo tempo, a sessão JSF trocaria o contribuinte
    selecionado e um veria o documento do outro (erro “o site está com o
    documento X, mas foi pedido Y”). Por isso o acesso ao site é
    **serializado por um lock**: as duas rotinas rodam na mesma execução,
    mas nunca pisam no site simultaneamente.
  - As linhas em revisão aparecem com **fundo laranja escuro com faixas
    diagonais** bem visíveis (aspecto de “em obra”), aplicadas **somente ao
    documento que está sendo revisado naquele instante**. Ao passar para o
    próximo, a faixa anterior é removida. Isso vale também para o **Revisar**
    manual.
  - O andamento aparece no rótulo **Revisão paralela: ...** na aba Consulta,
    inclusive a mensagem *sem pendências para revisar*, que explica quando
    não há itens pendentes para a segunda consulta. O **ESC 2x** / botão
    **Parar** encerra as duas consultas.
  - A revisão tem **progresso próprio** (`_revisao_progresso.json`): se for
    interrompida, retoma de onde parou na próxima execução, sem recomeçar do
    primeiro pendente. Ao concluir, o arquivo é removido. Ao mesclar, os
    resultados revisados são gravados também no progresso da consulta, para
    nenhuma linha desaparecer ao reabrir o programa.

---

## Sessão, cookies e token de autenticação

O programa **não armazena senhas nem cookies em arquivo**. Ele lê, somente em
memória, os dois cookies do sistema no perfil do navegador detectado:

- `__Secure-DA` — JWT de autenticação;
- `JSESSIONID` — sessão JSF (perfil selecionado no servidor).

Renovação automática:

- Se falta **menos de 1 minuto** para o JWT expirar, o token é relido do navegador
  antes da próxima consulta — **sem interromper o trabalho**.
- Se uma consulta falha por **token/sessão expirado** (404, página de erro
  `13429`, "Sessão expirada", erro de autenticação), o programa pega o token
  novo do navegador e **refaz aquela consulta** (até 3 tentativas).
- Se todas as tentativas falharem, o item fica `NAO VERIFICADO` e entra no
  Revisar; o log orienta a entrar no sistema novamente no navegador.

> O programa não tem credenciais para fazer login sozinho. Manter a aba do
> sistema aberta e logada no navegador é o que permite a renovação do token.

---

## Modo com e sem navegador

- **Padrão (com navegador):** Playwright. Se o navegador estiver com CDP ativo
  (porta 9222), o programa usa uma aba dele; caso contrário, abre um Chromium
  invisível com os cookies do navegador. É o modo mais confiável — o site é
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
- **ESC 2x** — parar consultas (1º aperto faz o botão **Parar** piscar em
  laranja/vermelho; 2º executa a parada).
- **Ambiente → Atualizar cookies** — relê manualmente JWT e JSESSIONID do navegador.

---

## Estrutura dos arquivos

```
RPC - Consulta de débitos - REVISADO/
├── main.py              # GUI (PyQt6), planilhas, worker e instalação
├── config.py            # URLs, porta CDP, cookie e pausa entre consultas
├── brave_cookie.py      # leitura segura dos cookies do navegador + validade do JWT
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
| "Cookie não encontrado" | Abra o sistema no navegador, faça login e clique em **Atualizar cookies**. |
| `NAO VERIFICADO ... status 404` | Sessão/viewstate expirada: o programa já tenta 3x com token novo; se persistir, entre no sistema novamente e use **Revisar**. |
| `página de erro (código 13429/13414)` | Mesmo tratamento acima: token renovado e consulta refeita automaticamente. |
| `Botão 'Pesquisar' não encontrado` | O programa abre sessão nova e repete a consulta (até 3x). Se persistir, confira se a aba do sistema está logada no navegador. |
| `O site está com o documento X, mas foi pedido Y` | O site re-renderizou o campo e restaurou um CPF anterior. O programa redigita o documento e confirma o valor antes de pesquisar; se ainda divergir, descarta o resultado (nunca aceita um débito de outro CPF). |
| `CPF/CNPJ inválido (dígito verificador)` | CPFs do tipo `000.000.xxx-xx` com dígitos verificadores errados são rejeitados pelo site (erros 13445/13446/404). O programa **não consulta** esses documentos no site: procura os débitos no relatório antigo e os lança no resultado. Se não houver débito lá, marca `DOCUMENTO INVÁLIDO` (fora do Revisar). |
| `Page.goto: net::ERR_ABORTED` | Navegação abortada também dispara nova tentativa com sessão nova. |
| "Playwright Sync API inside the asyncio loop" | Corrigido: cada sessão roda em thread limpa. |
| macOS não abre o `.app` | Botão direito → Abrir → Abrir (app assinado localmente, sem certificado Apple). |
| Navegador não inicia com depuração | Aba **Ambiente** → **Abrir navegador com depuração (9222)**. |

Logs: `~/Library/Logs/ConsultaDebitos.log` (REVISADO) e
`~/Library/Logs/ConsultaDebitos-x86.log` (X86).

---

## Privacidade e segurança

- Nenhuma cópia do banco de cookies é gravada em disco (abertura somente-leitura).
- A consulta SQL filtra por **nomes e host exatos** — nenhum outro cookie é lido.
- A chave do Keychain de cada navegador (ex.: "Brave Safe Storage", "Chrome Safe Storage") fica apenas em memória.
- Tokens e cookies nunca são enviados a terceiros.
- Planilhas, CSVs e relatórios com dados de devedores **não** entram neste
  repositório.
