# RPC — Consulta débitos ISS DIGITAL E IPTU

Programa para consultar débitos no **ISS Digital** de Santana do Araguaia-PA a
partir de uma planilha de CPFs/CNPJs, com interface gráfica e exportação do
resultado. Verifica também a situação de pagamento de cada dívida e a
validade dos CIFs no **Cadastro Imobiliário do IPTU Digital**.

## O que o programa faz

Ao abrir, o programa carrega automaticamente a planilha base desta pasta:
`RELATORIO_CONTABILIDADE_2022_ANALISTAS.xlsx` (ela já vem copiada aqui).

A consulta de débitos puxa **somente a situação DÉBITO**. A mudança em relação
às versões antigas está na **saída**.

O botão **1 linha: ATIVADA/DESATIVADA** (aba **Classificação e Verificação**) liga ou desliga a
gravação na coluna `OBSERVAÇÃO (débitos)`:

- **ATIVADA** (padrão): todos os débitos de cada pessoa vão em uma única linha
  da OBSERVAÇÃO, atualizando sempre `..._1_LINHA_TODOS_OS_DEBITOS.xlsx`;
- **DESATIVADA**: a OBSERVAÇÃO não é alterada e o resultado é salvo no formato
  clássico, **uma linha por débito**, em `..._resultado.xlsx`.

A preferência fica salva e é reaplicada na próxima abertura.

- Quando a planilha de entrada tem a coluna **`OBSERVAÇÃO (débitos)`** (como
  o `RELATORIO_CONTABILIDADE_2022_ANALISTAS.xlsx` desta pasta), o botão
  **Salvar resultado...** preenche **sempre o mesmo arquivo**:
  `relatorios/RELATORIO_CONTABILIDADE_2022_ANALISTAS_1_LINHA_TODOS_OS_DEBITOS.xlsx`.
  Ele não abre "Salvar como" e **não cria outro arquivo**: se o arquivo já
  existir, a coluna OBSERVAÇÃO dele é atualizada. As abas dos analistas são
  preservadas.
- Cada pessoa (CPF/CNPJ) recebe **todos os seus débitos em UMA ÚNICA LINHA**
  na célula da coluna OBSERVAÇÃO, no formato
  **`Tipo de Débito – Competência; Tipo de Débito – Competência; ...`**. Exemplo:

  ```
  IPTU (M) – fev/22; TAXA DE LICENÇA PARA LOCALIZAÇÃO E FUNCIONAMENTO (M) – fev/22; TAXA DE PRESTAÇÃO DE SERVIÇOS (M) – fev/22
  ```

- Parcelas com o mesmo tipo/competência aparecem uma única vez, pois a
  observação mostra somente esses dois campos. A coluna é alargada para a
  linha inteira ficar visível.
- Se a planilha de entrada **não** tiver a coluna OBSERVAÇÃO, o programa
  continua gerando a planilha "Resultado" de sempre (uma linha por débito),
  mas **sem CNPJs**: documentos de 14 dígitos são descartados dessa saída.

### Botão Consulta completa (tudo em uma etapa)

O botão **Consulta completa**, na aba **Classificação e Verificação**,
executa em sequência:

1. **Revisão completa** de todos os documentos da aba selecionada, consultando
   todas as situações relevantes (não só `DÉBITO`) e mantendo no resultado
   **somente os débitos em aberto**;
2. **Classificação de pagos** — cruza por ID Dívida/CIF/tipo+competência e
   gera `..._VERIFICACAO_PAGOS.xlsx` e `..._DETALHE_PAGOS.xlsx`, incluindo
   débitos pagos que o site mostrou e que não estavam na planilha;
3. **Validação dos CIFs** no Cadastro Imobiliário, gerando
   `..._VERIFICACAO_CIFS.xlsx` e `..._DETALHE_CIFS.xlsx`.

É a forma mais completa de "ver os débitos, classificar e validar" de uma vez.
Pode ser interrompida com **ESC 2x**. A revisão completa grava progresso
próprio: ao executar novamente, **retoma somente os documentos sem resultado
confirmado** (os que falharam por sessão/erro voltam para a fila) e mantém os
resultados já obtidos. Para começar do zero, use **Limpar progresso**.

### Botão Verificar pagos — quais dívidas o gestor ainda pode cobrar

O botão **Verificar pagos** (aba **Classificação e Verificação**, a 3ª etapa) cruza os débitos já listados
na `OBSERVAÇÃO (débitos)` com a situação **atual** informada pelo site.
Ele consulta cada documento com **todas as situações relevantes** (27 filtros
do ISS Digital, incluindo `PAGO`, `PAGO MANUAL`, `PAGO POR COMPENSAÇÃO`,
`CANCELADO`, `PRESCRITO`, etc.) para não perder nenhuma forma de pagamento, e
classifica cada débito por **ID Dívida** (prioridade), CIF + competência ou
tipo + competência. A saída tem duas colunas:

| Coluna | Valores |
|--------|---------|
| `PAGO?` | `PAGO`, `EM ABERTO`, `CANCELADO`, `PRESCRITO`, `NÃO CONFIRMADO`, `NÃO VERIFICADO` |
| `COBRÁVEL?` | `SIM` (dívida viva), `NÃO` (paga/cancelada/prescrita/remida), `NÃO CONFIRMADO` |
| `SITUAÇÃO (SITE)` | Situação informada pelo site para cada débito: `DÉBITO`, `AJUIZADO`, `PAGO MANUAL`, `CANCELADO`, etc. |
| `NÍVEL DE ALERTA` | Cor de alerta da cobrança (verde/laranja/amarelo/vermelho), explicada abaixo |

### Nível de alerta (cores)

| Cor | Significado |
|-----|-------------|
| **Verde — COBRÁVEL** | Confirmado no site e em aberto: pode cobrar |
| **Laranja — INDETERMINADO** | O sistema não conseguiu decidir (sem correspondência exata, conflito, ID não localizado) **ou** há inconsistência de datas típica da migração — competência 2022 com vencimento 2021, por exemplo. Não cobrar sem análise |
| **Amarelo — PENDENTE** | Pendente de análise/falha na consulta: não cobrar ainda |
| **Vermelho — NÃO COBRAR** | Pago, pago manual, cancelado, prescrito, remido, suspenso etc. — mesma cor para todos os encerramentos |

A legenda também aparece na aba **Classificação e Verificação**.

Os débitos da classificação — na OBSERVAÇÃO, nas colunas cruzadas e no
relatório detalhado — são sempre listados em **ano decrescente** (mais novo
primeiro), na mesma ordem usada pela planilha de 1 linha.

Os valores monetários são gravados como **texto já formatado** (`R$ 253,91`).
Isso evita que a formatação numérica do Excel mude com o idioma e exiba
`R$ 253,910` em vez de `R$ 253,91`.

O módulo novo analisa **todos** os débitos dos resultados, inclusive os que
 ainda aguardam confirmação (marcados internamente). Quando a consulta
 confirma que não há débitos, a descrição é **`Sem débitos verificados`** e o
 status é `VERIFICADO`; quando não há confirmação, a descrição é
 `Não foi possível verificar débitos` e o status é `NAO VERIFICADO` — nunca
 um falso "sem débitos".

Nunca é atribuída situação "por chute": sem correspondência exata de tipo e
competência, o débito fica `NÃO CONFIRMADO`.

**Regra de segurança contra cobrança indevida:** `COBRÁVEL? = SIM` só aparece
para situação explicitamente de dívida viva (débito, dívida ativa, ajuizado,
consolidado, parcelado). Todas as formas de pagamento/encerramento são
tratadas como `NÃO`: `PAGO`, `PAGO MANUAL`, `PAGO PARCELAMENTO`, `PAGO POR
COMPENSAÇÃO`, `PAGO POR DAÇÃO`, `PAGO RESTITUIDO`, `AJUIZADO PAGO`, `DÍVIDA
ATIVA PAGA`, `REMIDO`, `QUITADO`, `BAIXADO`, `PRESCRITO`, `CANCELADO`,
`SUBSTITUÍDO`, `ISENTO`, `IMUNE`, `NÃO TRIBUTÁVEL` e suspensões judicial/
administrativa. Um débito com valor devido zero e valor pago positivo também
é marcado como pago. Situações ambíguas (parcelamento revertido/descumprido,
boleto gerado, lançamento em AINF, em fiscalização, situação vazia) ficam
`NÃO CONFIRMADO` — nunca viram cobrança.

Foram levantadas as listas completas do sistema: **33 situações do ISS
Digital** (checkboxes `situaçãoDebito`, usados na consulta) e **27 situações
de pagamento do IPTU Digital**. Todas elas estão mapeadas, inclusive casos
traiçoeiros como `RETIDO_DENTRO_MUNICIPIO_DEBITO` — embora contenha a palavra
"débito", o tributo foi retido na fonte e **não** pode ser cobrado. O arquivo
`_DETALHE_PAGOS.xlsx` traz a aba **Situações (auditoria)** com as 60 situações
e como cada uma é classificada.

Arquivos gerados em `relatorios/` (o relatório original não é alterado):

- `<planilha>_VERIFICACAO_PAGOS.xlsx` — cópia do relatório com `PAGO?` e
  `COBRÁVEL?`, na mesma ordem dos débitos da observação;
- `<planilha>_DETALHE_PAGOS.xlsx` — uma linha por débito, com a situação do
  site, valores, CIF, ID Dívida e Vencimento.

A verificação respeita a pausa entre consultas e pode ser interrompida com
**ESC 2x**; os resultados parciais são salvos ao parar.

### Botão Verificar CIFs — o CIF é real?

O botão **Verificar CIFs** (aba **Classificação e Verificação**) é a etapa seguinte. Para cada
CPF/CNPJ com débitos, ele consulta o **Cadastro Imobiliário** do IPTU Digital
(`#/cadastroImobiliario`, digitando o CPF/CNPJ no campo próprio) e cruza os
CIFs obtidos com os CIFs das dívidas. Um CIF só é considerado real quando
aparece no cadastro **com endereço**.

| Coluna | Valores |
|--------|---------|
| `CIF REAL?` | `SIM` (endereço com número ou lote), `NÃO` (endereço vazio ou sem número e sem lote → CIF fictícia), `NÃO CONFIRMADO` (não apareceu), `NÃO VERIFICADO` (falha na consulta) |
| `ENDEREÇO (CIF)` | Endereço informado pelo cadastro |

Arquivos gerados em `relatorios/`:

- `<planilha>_VERIFICACAO_CIFS.xlsx` — cópia do resultado com `CIF REAL?` e
  `ENDEREÇO (CIF)`;
- `<planilha>_DETALHE_CIFS.xlsx` — um CIF por linha, com endereço, inscrição
  imobiliária e situação do cadastro/imóvel.

Também respeita a pausa e pode ser interrompida com **ESC 2x**.

O campo **Começar do registro** (aba Consulta) permite iniciar a consulta ou a revisão completa a partir de um número específico da planilha. Deixe em **1** para usar a retomada automática.

A aba **Consulta** não oferece mais o modo "HTTP direto": o site é
JSF/PrimeFaces e exige JavaScript, então essa opção ficou desativada com a
explicação na própria tela. A consulta sempre usa o navegador automático ou a
sessão real (porta 9222). O cookie JWT aparece oculto no campo de sessão.

A aba **Classificação e Verificação** agrupa as ações em um painel e traz
instruções de uso e a legenda das cores.

A aba **Resultados** começa com um painel **Resumo por situação**, com os
contadores atualizados automaticamente: Total, VERIFICADO, SEM CONFIRMAÇÃO /
PENDENTE DE ANÁLISE, DOCUMENTO INVÁLIDO, NAO VERIFICADO e ERRO.

## Proteção jurídica do analista

Toda saída (XLSX ou CSV) traz aviso de que:

- os dados vêm dos sistemas oficiais ISS Digital / IPTU Digital (Desenvolve
  Cidade — Prefeitura de Santana do Araguaia/PA) e todo cruzamento deriva
  dessa base na data da consulta;
- o arquivo é **auxiliar**, não é certidão, comprovante, lançamento nem ato
  administrativo, e não substitui a consulta ao sistema oficial;
- a decisão final de cobrar, suspender ou baixar é do **gestor público**,
  com base no sistema oficial e nos autos do processo;
- gerar o arquivo não é certificação nem declaração de quem o operou, e não
  há assinatura de servidor público — o rodapé registra data/hora e a conta
  que gerou, como rastro, não como assinatura;
- `COBRÁVEL` é classificação de apoio baseada em padrão cadastral (anos
  consecutivos de IPTU) e na situação consultada — **não é prova de que não
  houve pagamento**;
- a legenda explica verde/laranja/vermelho/amarelo (débito duplicado, que
  deve ser conferido antes de qualquer cobrança).

No CSV, por não haver aba de aviso, a mesma advertência acompanha **cada
linha** na coluna `AVISO`.

## Aviso nos relatórios

Todo arquivo gerado começa com a aba **AVISO**, contendo a origem dos dados
(sistemas oficiais ISS Digital / IPTU Digital mantidos por Desenvolve Cidade),
a explicação de que todo cruzamento/classificação deriva da base de dados do
site, a natureza auxiliar do relatório, a ressalva de atualidade, a
responsabilidade da decisão final pelo gestor e o uso interno. A aba é criada
na posição 0 e não altera os cabeçalhos nem as abas dos analistas.

## Auditoria contra cobrança indevida

Uma auditoria completa do projeto resultou nas seguintes garantias:

Toda situação que o programa não consegue classificar com segurança —
site fora, resposta duvidosa, situação desconhecida ou divergência de dados —
recebe **`CLASSIFICADO — INDETERMINADO`**. Se houver Erro, o item volta no
Revisar; se não houver (dado migrado já analisado), fica apenas como
indeterminado, sem cobrança automática.

- **`regras_debito.py`** é a regra única do que conta como débito cobrável.
  `"Sem débitos"`, `"-"`, status diferente de `OK/VERIFICADO`, erros,
  ressalvas e débitos ainda não confirmados nunca entram na OBSERVAÇÃO.
- **Arquivo de apoio mais recente:** é escolhido pela data de modificação
  (não por ordem alfabética), é ignorado se tiver mais de 30 dias, verifica
  também a situação da dívida ativa e seus débitos são gravados com o status
  `CLASSIFICADO — INDETERMINADO` — nunca como `VERIFICADO`. A coluna **Erro permanece
  limpa**. Progressos antigos são normalizados na abertura.
- **Consulta Playwright:** débitos só são publicados quando o documento foi
  confirmado na resposta; na dúvida, o item volta para o Revisar.
- **Modo HTTP:** confere CPF **e** nome na resposta, lê só texto visível e
  nunca transforma "linhas sem DÉBITO" em "sem débitos". Um "sem débitos" do
  HTTP fica `NAO VERIFICADO`.
- **Cookies:** o `JSESSIONID` do host exato tem prioridade e não é
  sobrescrito pelo cookie de domínio do mesmo nome.
- **Verificação de pagos:** pareamento exige mês e competência iguais; se
  mais de um débito tem o mesmo tipo/competência com situações diferentes,
  o item fica `NÃO CONFIRMADO`.
- **Tipos de débito:** foram levantados os **576 tipos** que o site mostra
  (`tipos_debito_site.txt`). A chave de comparação ignora acentos, `(M)`,
  artigos e a ordem dos termos, casando nomes descritivos
  (`TAXA DE ALVARÁ DE OBRAS RESIDENCIAL`) com códigos
  (`TAXA_ALVARA_OBRAS_RESIDENCIAL`). Mesmo assim, a prioridade continua sendo
  o **ID Dívida**.
- **Verificação de CIFs:** o resultado só é aplicado ao par exato
  `(documento, CIF)` — nunca reaproveitado de outro CPF.

### Botões Revisar e Revisão completa

- **Revisar**: aplica os débitos já obtidos e consulta somente os CPFs/CNPJs
  com `OBSERVAÇÃO (débitos)` vazia **da aba selecionada**.
- **Revisão completa**: reconulta **todos os documentos da aba selecionada**,
  inclusive os que já têm débito preenchido. Use quando quiser refazer toda
  a conferência da planilha (na aba `PF - Caliu`, são 2.022 CPFs distintos).


Nesta versão o **Revisar** foi adaptado para investigar as observações em
branco:

1. Primeiro aplica ao arquivo destino **todos os débitos já obtidos** nesta
   sessão (sem nova consulta).
2. Depois consulta, em modo invisível, **somente os documentos cuja coluna
   OBSERVAÇÃO ainda está vazia — e apenas da aba selecionada no programa**
   (se você escolheu uma aba de CPF, não aparece CNPJ de outra aba).
3. Ao concluir, o arquivo é atualizado automaticamente. Quem realmente não
   tem débito continua com a célula vazia (não é inventada observação).

> Isso pode demorar: na aba `PF - Caliu`, por exemplo, há milhares de CPFs sem
> observação, e boa parte deles não tem débito. Você pode parar com **ESC 2x**;
> o progresso é salvo e um novo **Revisar** retoma exatamente os que ficaram
> vazios na mesma aba. Para as outras abas, selecione cada uma e repita.

> **CPFs do sistema legado** (`000.000.xxx-xx`) têm dígito verificador inválido.
> No **Revisar** eles são enviados ao site mesmo assim: se o site rejeitar, a
> classificação anterior é mantida e o resumo informa quantas consultas
> forçadas falharam. Na consulta normal eles continuam sendo gravados como
> `DOCUMENTO INVÁLIDO`. Para recuperar os débitos sem depender do site, use um
> CSV de apoio `relatorio-lancamentos-debito-divida-ativa-*.csv` recente (até
> 30 dias) na pasta do programa. A correção definitiva é substituir esses
> códigos pelos CPFs reais na planilha.

### Gerar a cópia preenchida sem consultar o site

O arquivo `preencher_observacao.py` usa a base de débitos já existente
(`relatorios/Análise e rastreio 2022 - cópia_resultado_parcial.xlsx`) e
preenche o relatório sem precisar abrir o navegador:

```bash
python3 preencher_observacao.py
```

Saída (sempre a mesma): `relatorios/RELATORIO_CONTABILIDADE_2022_ANALISTAS_1_LINHA_TODOS_OS_DEBITOS.xlsx`.

> **Atenção:** a base de débitos disponível veio de uma consulta parcial da
> aba `PF - Caliu`. Por isso, as abas `PJ - Angela`, `PJ - Cid` e
> `PF - Gabriela Godoy` ficam sem observação até que essas abas também sejam
> consultadas no programa (basta selecionar cada aba na aba Consulta e rodar).

## Como abrir (sem terminal)

Dê **duplo clique** em:

```
Consulta de Débitos.app
```

Ele está dentro desta pasta e abre o programa diretamente. Você pode
arrastá-lo para o **Dock** ou para a pasta **Aplicativos** para ter o ícone
sempre à mão.

**Se o macOS disser que não pode abrir:** o aplicativo foi criado localmente
e não tem certificado pago da Apple. Clique com o **botão direito** no ícone
→ **Abrir** → **Abrir** na confirmação. Isso só precisa ser feito uma vez.
Outra opção: **Ajustes do Sistema → Privacidade e Segurança → Abrir assim mesmo**.

**Plano B (também sem digitar):** duplo clique em
`Abrir Consulta de Débitos.command`. Abre o programa sozinho; aparece uma
janela do Terminal, mas você não escreve nada.

O aplicativo já foi assinado localmente e teve o bloqueio de quarentena
removido. Mesmo assim, o macOS pode pedir a confirmação acima por ser um
app sem certificado da Apple.

O aplicativo escolhe automaticamente um Python com as dependências e, se
alguma biblioteca estiver faltando, instala tudo sozinho antes de abrir.
Mensagens de erro ficam em `~/Library/Logs/ConsultaDebitos.log`.

Se preferir abrir pelo terminal (equivalente ao duplo clique):

```bash
cd "RPC - Consulta débitos ISS DIGITAL E IPTU"
python3 main.py
```

No **canto superior direito** da janela há o botão **🌙 Escuro / ☀️ Claro**
para alternar entre os temas claro e escuro. A escolha fica salva e é aplicada
na próxima abertura. O tema escuro é **preto com detalhes laranja e letras
brancas**.

A janela tem quatro abas:

| Aba | O que faz |
|-----|-----------|
| **Consulta** | Escolhe a planilha de entrada, mostra a prévia, gerencia cookies e executa/para as consultas |
| **Resultados** | Tabela com os débitos agrupados por pessoa: todos os débitos do 1º nome em branco, do 2º em cinza-claro, alternando; botões **Revisar** e **Parar consultas** |
| **Log** | Histórico de todas as operações |
| **Ambiente** | Instala dependências/navegador, abre o Brave com depuração, atualiza cookies e verifica pré-requisitos — sem usar o terminal |

### Fluxo recomendado

1. Aba **Ambiente** → clique em **Instalar tudo** (só na primeira vez) e em
   **Abrir Brave com depuração (9222)**. O programa pergunta antes de fechar
   o Brave, abre o navegador com a porta 9222 e abre o sistema.
2. No Brave, entre no sistema normalmente e deixe a aba logada.
3. Volte ao programa, aba **Ambiente** → **Atualizar cookies**.
4. Aba **Consulta** → **Escolher planilha...** → **Executar**.
   - Marque/desmarque **Tentar SEM navegador (HTTP direto)** para alternar
     entre os dois modos (detalhes abaixo).
   - O filtro de situação considera **somente DÉBITO** (dívida ativa e
     ajuizado ficam de fora).
   - Ajuste **Pausa entre consultas** (em segundos) na aba Consulta. Aumente
     se o site ficar lento ou bloquear; marque **Variar a pausa** para aplicar
     uma oscilação aleatória de ±30% e não criar um padrão detectável.
   - Pressione **ESC 2x** para parar as consultas: o 1º aperto muda a cor
     dos botões **Parar** e do aviso; o 2º aperto interrompe de fato.
     O aviso **PRESSIONE "ESC" 2x PARA PARAR CONSULTA** aparece na tela.
   - Se interromper, o progresso é salvo e a retomada é oferecida na próxima execução (ver abaixo).
5. Aba **Resultados** → use **Revisar** para reconquistar apenas os itens com
   erro/não verificados/sem débitos, depois **Salvar resultado...** (XLSX ou CSV).

## Formato da planilha de entrada

Aceita `.xlsx`, `.xlsm` e `.csv`. O cabeçalho é detectado automaticamente,
mesmo abaixo de títulos, e abas são listadas para escolha.

Colunas usadas (nome ignorando acentos/maiúsculas):

| Coluna | Descrição | Obrigatória |
|--------|-----------|-------------|
| Nome | Nome do devedor | Não |
| CPF | CPF (11 dígitos) | Sim, ou CNPJ/CIF |
| CNPJ | CNPJ (14 dígitos) | Sim, ou CPF |
| Tipo de Pessoa | FÍSICA, JURÍDICA ou IMÓVEL | Não — **inferido pelos dígitos**, mesmo que o cabeçalho da coluna esteja errado (por exemplo, um CPF em uma coluna chamada "CNPJ") |

## Saída

Colunas: Nome, CPF, Tipo de Débito, Competência, Valor (R$), Status, **CIF**,
**ID Dívida**, **Vencimento**, Erro e **MOTIVO**. Assim dá para saber, por
exemplo, qual CIF (imóvel) está devendo IPTU.

Todo XLSX gerado (inclusive os arquivos com data/hora) já sai com o **Status
classificado e colorido** — `COBRÁVEL` (verde), `NÃO COBRAR` (vermelho),
`CLASSIFICADO — INDETERMINADO` (laranja), `DUPLICADO — CONFERIR` (amarelo) ou
`ERRO DE CONSULTA` — e com a coluna **MOTIVO** explicando cada classificação.

Internamente a consulta usa `OK`, `VERIFICADO`, `NAO VERIFICADO` e `ERRO`. O
programa só afirma "Sem débitos" quando a pesquisa foi realmente executada;
na menor dúvida ele marca `NAO VERIFICADO` para não gerar falso negativo.

## Retomada de consultas interrompidas

- O programa **lembra a última planilha** e a reabre automaticamente ao iniciar.
  O botão **Reabrir última** serve para abrir novamente se você tiver trocado de planilha.
- Ao pressionar **Executar** com progresso salvo, ele pergunta se deseja retomar.
- **O último registro é sempre consultado de novo** — inclusive quando a busca
  tinha começado e foi interrompida no meio. O progresso guarda um sinalizador
  `em_andamento` para identificar esse caso.
- Os resultados do registro refeito são descartados antes de reconstruir a
  tabela, evitando duplicatas.
- Ao parar com ESC 2x ou com o botão **Parar**, o item atual termina e o
  progresso é gravado; nada é perdido.

## Com ou sem navegador

O botão de marcação na aba **Consulta** alterna entre:

- **Desmarcado (padrão)** — usa Playwright. É o modo mais confiável porque o
  site (JSF/PrimeFaces) precisa de JavaScript. Se o Brave estiver com a porta
  9222 ativa, o programa usa uma aba dele; se não estiver, usa um Chromium
  **invisível** (headless) com os cookies do Brave.
- **Marcado** — tenta HTTP direto, sem abrir navegador (`consulta_http.py`).
  É experimental: reenvia o formulário com o `javax.faces.ViewState` e lê a
  tabela da resposta. Se o site exigir JavaScript, a resposta não pode ser
  confirmada e o item fica como `NAO VERIFICADO` — nunca "sem débitos".

## Configuração

`config.py` contém as URLs, a porta CDP, o nome do cookie e o intervalo entre
consultas (`EXTRACTION_DELAY`). Senhas/cookies não ficam salvos em arquivo:
são lidos do perfil do Brave apenas em memória.

## Arquivos

- `main.py` — programa principal (PyQt6): interface, planilhas, consulta e instalação
- `config.py` — URLs, porta de depuração, cookie e intervalo entre consultas
- `brave_cookie.py` — leitura segura dos cookies do Brave
- `consulta_http.py` — modo experimental sem navegador (HTTP direto)
- `requirements.txt` — dependências
- `GUIDE_DE_USO.txt` — guia rápido
- `RELATORIO_REVISAO.md` — histórico das correções aplicadas

A pasta `cpf_cnpj/` é independente deste programa e foi mantida.
Protótipos e scripts que não eram mais usados foram removidos.

## Requisitos

Python 3.10+, PyQt6, Playwright, openpyxl e cryptography (ver
`requirements.txt`). O navegador Chromium do Playwright é instalado pelo
botão **Instalar navegador** da aba Ambiente.
