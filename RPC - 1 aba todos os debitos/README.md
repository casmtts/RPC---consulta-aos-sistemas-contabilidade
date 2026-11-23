# Consulta de Débitos — Lista 1 aba: todos os débitos na OBSERVAÇÃO

Programa para consultar débitos no sistema NFS-e de Santana do Araguaia-PA a
partir de uma planilha de CPFs/CNPJs, com interface gráfica e exportação do
resultado.

## O que mudou nesta versão

Ao abrir, o programa carrega automaticamente a planilha base desta pasta:
`RELATORIO_CONTABILIDADE_2022_ANALISTAS.xlsx` (ela já vem copiada aqui).

A consulta continua **igual à versão REVISADA**. A mudança é na **saída**:

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
cd "RPC - Consulta de débitos - REVISADO"
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
**ID Dívida**, **Vencimento** e Erro. Assim dá para saber, por exemplo, qual
CIF (imóvel) está devendo IPTU.

Status possíveis: `OK`, `NAO VERIFICADO` e `ERRO`. O programa só afirma
"Sem débitos" quando a pesquisa foi realmente executada; na
menor dúvida ele marca `NAO VERIFICADO` para não gerar falso negativo.

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

Para **Mac Intel (x86_64)** ou para rodar x86 sob Rosetta, use a pasta
`RPC - Consulta de débitos - PORTABILIDADE X86`, que possui executável e
instalador próprios e um diretório de browsers do Playwright separado.

## Requisitos

Python 3.10+, PyQt6, Playwright, openpyxl e cryptography (ver
`requirements.txt`). O navegador Chromium do Playwright é instalado pelo
botão **Instalar navegador** da aba Ambiente.
