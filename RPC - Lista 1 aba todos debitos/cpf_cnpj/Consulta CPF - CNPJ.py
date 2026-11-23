#!/usr/bin/env python3
"""
Interface gráfica para consultar CNPJs na Receita Federal.

- Botão "Fazer upload da base" para escolher um .txt ou .csv com os CNPJs.
- A lista mostra a base inteira, com o status de cada CNPJ (Pendente, Consultando, Consultado, Erro).
- O painel de status mostra os totais e o andamento das consultas.
- "Gerar relatório" salva um .txt com o resumo e os CNPJs agrupados por situação cadastral.
- "Exportar CSV" salva a tabela completa.

Usa tkinter (já vem com o Python) e o módulo consulta_cnpj.py, que deve estar na mesma pasta.
"""

import csv
import os
import queue
import re
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from consulta_cnpj import CAMPOS_CSV, cnpj_valido, consultar, limpar

INTERVALO_PADRAO = 1.5  # segundos entre consultas (ReceitaWS aceita ~3/min no plano gratuito)

PENDENTE = "Pendente"
CONSULTANDO = "Consultando"
CONSULTADO = "Consultado"
ERRO = "Erro"

COLUNAS = [
    ("status", "Status", 100),
    ("cnpj", "CNPJ", 140),
    ("situacao", "Situação", 150),
    ("data_situacao", "Data situação", 110),
    ("razao_social", "Razão social", 300),
    ("nome_fantasia", "Nome fantasia", 200),
    ("municipio", "Município", 150),
    ("uf", "UF", 50),
    ("fonte", "Fonte", 90),
    ("erro", "Erro", 300),
]
CHAVES = [c[0] for c in COLUNAS]


def _textos_planilha(caminho: str) -> list[str]:
    """Devolve os textos das células de uma planilha .xlsx. Se houver coluna com 'cnpj' ou 'cpf', usa só ela."""
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("para abrir .xlsx instale o openpyxl:  pip3 install openpyxl  (ou salve como CSV)")
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    textos: list[str] = []
    for ws in wb.worksheets:
        linhas = list(ws.iter_rows(values_only=True))
        if not linhas:
            continue
        idx = None
        for i, nome in enumerate(linhas[0]):
            nome_l = str(nome or "").lower()
            if "cnpj" in nome_l or "cpf" in nome_l:
                idx = i
                break
        corpo = linhas[1:] if idx is not None else linhas
        for linha in corpo:
            celulas = [linha[idx]] if idx is not None and idx < len(linha) else ([] if idx is not None else list(linha))
            for v in celulas:
                if v is None or isinstance(v, datetime):  # datas não são documentos
                    continue
                if isinstance(v, float) and v.is_integer():
                    v = int(v)
                textos.append(str(v))
    return textos


def ler_base(caminho: str) -> list[str]:
    """Extrai CNPJs válidos e sem repetição de um .txt, .csv ou .xlsx."""
    if caminho.lower().endswith(".xlsx"):
        textos = _textos_planilha(caminho)
    else:
        with open(caminho, encoding="utf-8-sig") as f:
            textos = re.split(r"[;,\t|\s]+", f.read())

    encontrados: list[str] = []
    vistos: set[str] = set()
    for texto in textos:
        c = limpar(texto)
        if c.isdigit() and 12 <= len(c) <= 13:  # zeros à esquerda removidos pelo Excel
            c = c.zfill(14)
        if len(c) == 14 and cnpj_valido(c) and c not in vistos:
            vistos.add(c)
            encontrados.append(c)
    return encontrados


def item_vazio(cnpj: str) -> dict:
    """Uma linha da base, com todas as colunas vazias exceto o CNPJ e o status."""
    item = {chave: "" for chave in CHAVES}
    item["cnpj"] = cnpj
    item["status"] = PENDENTE
    return item


def resumo(itens: list[dict]) -> dict:
    """Totais usados no painel de status."""
    total = len(itens)
    consultados = [i for i in itens if i["status"] == CONSULTADO]
    ativas = sum(1 for i in consultados if i["situacao"].upper() == "ATIVA")
    return {
        "total": total,
        "pendentes": sum(1 for i in itens if i["status"] == PENDENTE),
        "consultando": sum(1 for i in itens if i["status"] == CONSULTANDO),
        "consultados": len(consultados),
        "ativas": ativas,
        "nao_ativas": len(consultados) - ativas,
        "erros": sum(1 for i in itens if i["status"] == ERRO),
    }


def tag_da_linha(item: dict) -> str:
    if item["status"] == ERRO:
        return "erro"
    if item["status"] == CONSULTANDO:
        return "consultando"
    if item["status"] == CONSULTADO:
        return "ativa" if item["situacao"].upper() == "ATIVA" else "inativa"
    return ""


def gerar_relatorio_texto(itens: list[dict], arquivo_base: str, agora: datetime) -> str:
    """Monta o relatório em texto: resumo, situação cadastral, detalhes por situação, erros e pendentes."""
    r = resumo(itens)
    linhas = [
        "RELATÓRIO DE CONSULTA DE SITUAÇÃO CADASTRAL (RECEITA FEDERAL) - CNPJ",
        "=" * 72,
        f"Gerado em: {agora:%d/%m/%Y %H:%M}",
        f"Base: {arquivo_base or '-'}",
        "",
        "RESUMO",
        f"  Total na base:            {r['total']}",
        f"  Consultados com sucesso:  {r['consultados']}",
        f"  Com erro:                 {r['erros']}",
        f"  Pendentes:                {r['pendentes']}",
        "",
        "SITUAÇÃO CADASTRAL",
    ]

    grupos: dict[str, list[dict]] = {}
    for item in itens:
        if item["status"] == CONSULTADO:
            grupos.setdefault(item["situacao"] or "(sem situação)", []).append(item)

    if not grupos:
        linhas.append("  Nenhuma consulta concluída.")
    for situacao in sorted(grupos):
        linhas.append(f"  {situacao}: {len(grupos[situacao])}")
    linhas.append("")

    for situacao in sorted(grupos):
        linhas.append(f"DETALHES - {situacao} ({len(grupos[situacao])})")
        for item in grupos[situacao]:
            linhas.append(
                f"  {item['cnpj']} | {item['data_situacao'] or '-'} | "
                f"{item['uf'] or '-'} | {item['razao_social'] or '-'}"
            )
        linhas.append("")

    erros = [i for i in itens if i["status"] == ERRO]
    if erros:
        linhas.append(f"ERROS ({len(erros)})")
        for item in erros:
            linhas.append(f"  {item['cnpj']} | {item['erro']}")
        linhas.append("")

    pendentes = [i for i in itens if i["status"] == PENDENTE]
    if pendentes:
        linhas.append(f"PENDENTES ({len(pendentes)})")
        for item in pendentes:
            linhas.append(f"  {item['cnpj']}")
        linhas.append("")

    return "\n".join(linhas) + "\n"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Consulta CNPJ - Receita Federal")
        self.geometry("1200x720")
        self.minsize(900, 520)

        self.itens: list[dict] = []
        self.arquivo_base = ""
        self.fila: queue.Queue = queue.Queue()
        self.parar = threading.Event()
        self.consultando = False

        self._montar_interface()
        self.atualizar_contadores()

    # ---------- interface ----------
    def _montar_interface(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Linha 0: botões de ação e opções
        barra = ttk.Frame(self, padding=(8, 8, 8, 0))
        barra.grid(row=0, column=0, sticky="ew")

        linha1 = ttk.Frame(barra)
        linha1.pack(fill="x")
        self.btn_upload = ttk.Button(linha1, text="Fazer upload da base (.txt/.csv/.xlsx)", command=self.upload_base)
        self.btn_upload.pack(side="left")
        self.btn_consultar = ttk.Button(linha1, text="Consultar", command=self.iniciar_consulta, state="disabled")
        self.btn_consultar.pack(side="left", padx=6)
        self.btn_parar = ttk.Button(linha1, text="Parar", command=self.parar_consulta, state="disabled")
        self.btn_parar.pack(side="left")
        self.btn_relatorio = ttk.Button(linha1, text="Gerar relatório...", command=self.gerar_relatorio, state="disabled")
        self.btn_relatorio.pack(side="left", padx=6)
        self.btn_exportar = ttk.Button(linha1, text="Exportar CSV...", command=self.exportar, state="disabled")
        self.btn_exportar.pack(side="left")

        linha2 = ttk.Frame(barra, padding=(0, 6, 0, 0))
        linha2.pack(fill="x")
        ttk.Label(linha2, text="Intervalo (s):").pack(side="left", padx=(0, 4))
        self.var_intervalo = tk.StringVar(value=str(INTERVALO_PADRAO))
        ttk.Entry(linha2, textvariable=self.var_intervalo, width=6).pack(side="left")
        ttk.Label(linha2, text="Filtrar:").pack(side="left", padx=(16, 4))
        self.var_filtro = tk.StringVar()
        self.var_filtro.trace_add("write", lambda *_: self.atualizar_lista())
        ttk.Entry(linha2, textvariable=self.var_filtro, width=24).pack(side="left")

        # Linha 1: painel de status
        painel = ttk.LabelFrame(self, text="Status das consultas", padding=8)
        painel.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
        self.var_contadores = tk.StringVar()
        ttk.Label(painel, textvariable=self.var_contadores).pack(anchor="w")
        self.progresso = ttk.Progressbar(painel, mode="determinate")
        self.progresso.pack(fill="x", pady=4)
        self.var_agora = tk.StringVar(value="")
        ttk.Label(painel, textvariable=self.var_agora).pack(anchor="w")

        # Linha 2: lista da base (ocupa o espaço que sobra)
        meio = ttk.Frame(self, padding=(8, 4, 8, 0))
        meio.grid(row=2, column=0, sticky="nsew")
        meio.grid_rowconfigure(0, weight=1)
        meio.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(meio, columns=CHAVES, show="headings", selectmode="extended", height=18)
        for chave, titulo, largura in COLUNAS:
            self.tree.heading(chave, text=titulo)
            self.tree.column(chave, width=largura, anchor="w", stretch=True)
        self.tree.tag_configure("ativa", background="#e6f4ea")
        self.tree.tag_configure("inativa", background="#fdecea")
        self.tree.tag_configure("erro", foreground="#9e9e9e")
        self.tree.tag_configure("consultando", background="#fff8c4")

        rolagem_y = ttk.Scrollbar(meio, orient="vertical", command=self.tree.yview)
        rolagem_x = ttk.Scrollbar(meio, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=rolagem_y.set, xscrollcommand=rolagem_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        rolagem_y.grid(row=0, column=1, sticky="ns")
        rolagem_x.grid(row=1, column=0, sticky="ew")

        # Linha 3: mensagem de status
        rodape = ttk.Frame(self, padding=8)
        rodape.grid(row=3, column=0, sticky="ew")
        self.var_msg = tk.StringVar(value="Clique em 'Fazer upload da base' para começar.")
        ttk.Label(rodape, textvariable=self.var_msg).pack(anchor="w")

    # ---------- ações ----------
    def upload_base(self):
        if self.consultando:
            return
        caminho = filedialog.askopenfilename(
            title="Escolha a base de CNPJs",
            filetypes=[("Base de CNPJs", "*.txt *.csv *.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            return
        try:
            cnpjs = ler_base(caminho)
        except Exception as e:
            messagebox.showerror("Erro ao ler a base", str(e))
            return
        if not cnpjs:
            messagebox.showwarning("Base vazia", "Nenhum CNPJ válido encontrado no arquivo.")
            return

        self.itens = [item_vazio(c) for c in cnpjs]
        self.arquivo_base = os.path.basename(caminho)
        self.var_agora.set("")
        self.atualizar_lista()
        self.atualizar_contadores()
        self.btn_consultar.configure(state="normal")
        self._estado_botoes()
        self.var_msg.set(f"{len(cnpjs)} CNPJs válidos carregados de {self.arquivo_base}.")

    def iniciar_consulta(self):
        if self.consultando or not self.itens:
            return
        try:
            intervalo = max(0.0, float(self.var_intervalo.get().replace(",", ".")))
        except ValueError:
            messagebox.showerror("Intervalo inválido", "Informe um número, por exemplo 1.5.")
            return

        # Consulta os que ainda não foram consultados (pendentes e os com erro, para tentar de novo).
        pares = [(i, it["cnpj"]) for i, it in enumerate(self.itens) if it["status"] != CONSULTADO]
        if not pares:
            self.var_msg.set("Todos os CNPJs desta base já foram consultados.")
            return

        self.parar.clear()
        self.consultando = True
        self._estado_botoes()
        self.var_msg.set(f"Consultando {len(pares)} CNPJs...")
        threading.Thread(target=self._trabalhador, args=(pares, intervalo), daemon=True).start()
        self.after(100, self._verificar_fila)

    def parar_consulta(self):
        if self.consultando:
            self.parar.set()
            self.var_agora.set("Parando após a consulta atual...")

    def _trabalhador(self, pares: list[tuple[int, str]], intervalo: float):
        try:
            for n, (idx, cnpj) in enumerate(pares):
                if self.parar.is_set():
                    break
                self.fila.put(("inicio", idx, cnpj))
                resultado = consultar(cnpj)
                self.fila.put(("resultado", idx, resultado))
                if n < len(pares) - 1 and intervalo:
                    if self.parar.wait(intervalo):  # acorda na hora se o usuário clicar em Parar
                        break
        finally:
            self.fila.put(("fim", None, None))

    def _verificar_fila(self):
        try:
            while True:
                tipo, idx, dados = self.fila.get_nowait()
                if tipo == "inicio":
                    self.itens[idx]["status"] = CONSULTANDO
                    self.var_agora.set(f"Consultando agora: {dados}")
                    self._atualizar_linha(idx)
                elif tipo == "resultado":
                    novo = item_vazio(dados["cnpj"])
                    novo.update(dados)
                    novo["status"] = ERRO if "erro" in dados else CONSULTADO
                    self.itens[idx] = novo
                    self._atualizar_linha(idx)
                elif tipo == "fim":
                    self._finalizar()
                    return
                self.atualizar_contadores()
        except queue.Empty:
            pass
        if self.consultando:
            self.after(100, self._verificar_fila)

    def _finalizar(self):
        self.consultando = False
        self.var_agora.set("")
        r = resumo(self.itens)
        if self.parar.is_set():
            self.var_msg.set("Consulta interrompida. Clique em Consultar para continuar de onde parou.")
        else:
            self.var_msg.set(f"Concluído: {r['consultados']} consultados, {r['erros']} com erro.")
        self.atualizar_contadores()
        self._estado_botoes()

    def _estado_botoes(self):
        tem_base = bool(self.itens)
        tem_resultado = any(i["status"] != PENDENTE for i in self.itens)
        livre = not self.consultando
        self.btn_upload.configure(state="normal" if livre else "disabled")
        self.btn_consultar.configure(state="normal" if (livre and tem_base) else "disabled")
        self.btn_parar.configure(state="disabled" if livre else "normal")
        self.btn_relatorio.configure(state="normal" if (livre and tem_resultado) else "disabled")
        self.btn_exportar.configure(state="normal" if (livre and tem_resultado) else "disabled")

    def gerar_relatorio(self):
        if self.consultando or not any(i["status"] != PENDENTE for i in self.itens):
            return
        caminho = filedialog.asksaveasfilename(
            title="Salvar relatório",
            defaultextension=".txt",
            initialfile="relatorio_cnpj.txt",
            filetypes=[("Texto", "*.txt")],
        )
        if not caminho:
            return
        texto = gerar_relatorio_texto(self.itens, self.arquivo_base, datetime.now())
        with open(caminho, "w", encoding="utf-8-sig") as f:
            f.write(texto)
        self.var_msg.set(f"Relatório salvo em {caminho}")

    def exportar(self):
        exportar = [i for i in self.itens if i["status"] != PENDENTE]
        if not exportar:
            return
        caminho = filedialog.asksaveasfilename(
            title="Salvar resultado", defaultextension=".csv", filetypes=[("CSV", "*.csv")]
        )
        if not caminho:
            return
        with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=CAMPOS_CSV, extrasaction="ignore")
            w.writeheader()
            w.writerows(exportar)
        self.var_msg.set(f"Resultado salvo em {caminho}")

    # ---------- painel e lista ----------
    def atualizar_contadores(self):
        if not self.itens:
            self.var_contadores.set("Nenhuma base carregada.")
            self.progresso.configure(maximum=1, value=0)
            return
        r = resumo(self.itens)
        self.var_contadores.set(
            f"Total: {r['total']}   |   Pendentes: {r['pendentes']}   |   "
            f"Consultados: {r['consultados']}   |   Ativas: {r['ativas']}   |   "
            f"Não ativas: {r['nao_ativas']}   |   Erros: {r['erros']}"
        )
        feitos = r["consultados"] + r["erros"]
        self.progresso.configure(maximum=max(1, r["total"]), value=feitos)

    def _valores(self, item: dict) -> list[str]:
        return [str(item.get(chave, "") or "") for chave in CHAVES]

    def _passa_filtro(self, item: dict, termo: str) -> bool:
        if not termo:
            return True
        return termo in " ".join(self._valores(item)).lower()

    def atualizar_lista(self):
        """Reconstrói a lista inteira (usado ao carregar a base e ao mudar o filtro)."""
        termo = self.var_filtro.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        for idx, item in enumerate(self.itens):
            if self._passa_filtro(item, termo):
                tag = tag_da_linha(item)
                self.tree.insert("", "end", iid=str(idx), values=self._valores(item),
                                 tags=(tag,) if tag else ())

    def _atualizar_linha(self, idx: int):
        """Atualiza só uma linha, sem reconstruir a lista (mantém a rolagem)."""
        iid = str(idx)
        item = self.itens[idx]
        termo = self.var_filtro.get().strip().lower()
        if not self._passa_filtro(item, termo):
            if self.tree.exists(iid):
                self.tree.delete(iid)
            return
        tag = tag_da_linha(item)
        if self.tree.exists(iid):
            self.tree.item(iid, values=self._valores(item), tags=(tag,) if tag else ())
        else:
            self.atualizar_lista()


if __name__ == "__main__":
    App().mainloop()
