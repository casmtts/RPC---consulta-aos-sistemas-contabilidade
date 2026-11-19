"""Consulta de débitos: escolha a planilha (.xlsx ou .csv), execute e gere o resultado.

Arquivo único: não precisa de outros arquivos .py na pasta.
"""

from __future__ import annotations

import csv
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

COLUNAS_RESULTADO = ["tipo_debito", "ano", "situacao", "status_consulta", "erro"]
MAX_LINHAS_PREVIA = 500
SEPARADOR_CSV_SAIDA = ";"  # Excel em português usa ponto e vírgula


# ============ Consulta ao site (próxima etapa: Playwright) ============

class ConsultaErro(Exception):
    """Falha esperada de consulta; a mensagem vai para a coluna 'erro'."""


def consultar(registro: dict) -> list:
    """Retorna uma lista de débitos: [{"tipo_debito":..., "ano":..., "situacao":...}].
    Lista vazia significa "sem débitos". Ainda não implementado.
    """
    raise ConsultaErro("Consulta ao site ainda não implementada")


# ============ Leitura e gravação de planilhas ============

def ler_planilha(caminho) -> tuple:
    p = Path(caminho)
    ext = p.suffix.lower()
    if ext == ".csv":
        return _ler_csv(p)
    if ext == ".xlsx":
        return _ler_xlsx(p)
    raise ValueError(f"Formato não suportado: {ext or 'sem extensão'}. Use .xlsx ou .csv")


def gravar_planilha(caminho_saida, cabecalho: list, linhas: list) -> Path:
    p = Path(caminho_saida)
    ext = p.suffix.lower()
    if ext == ".csv":
        _gravar_csv(p, cabecalho, linhas)
    elif ext == ".xlsx":
        _gravar_xlsx(p, cabecalho, linhas)
    else:
        raise ValueError(f"Formato de saída não suportado: {ext}")
    return p


def caminho_saida_padrao(caminho_entrada) -> Path:
    p = Path(caminho_entrada)
    return p.with_name(f"{p.stem}_resultado{p.suffix.lower()}")


def _ler_csv(p: Path) -> tuple:
    # utf-8-sig remove o BOM que o Excel costuma colocar no início do arquivo
    with p.open(newline="", encoding="utf-8-sig") as f:
        amostra = f.read(4096)
        f.seek(0)
        try:
            dialeto = csv.Sniffer().sniff(amostra, delimiters=",;\t")
        except csv.Error:
            dialeto = csv.excel
        leitor = csv.reader(f, dialect=dialeto)
        cabecalho = [c.strip() for c in next(leitor, [])]
        linhas = []
        for valores in leitor:
            if not any(v.strip() for v in valores):
                continue
            linhas.append({
                h: (valores[i].strip() if i < len(valores) else "")
                for i, h in enumerate(cabecalho) if h
            })
    return [h for h in cabecalho if h], linhas


def _gravar_csv(p: Path, cabecalho: list, linhas: list) -> None:
    with p.open("w", newline="", encoding="utf-8-sig") as f:
        escritor = csv.DictWriter(
            f, fieldnames=cabecalho, delimiter=SEPARADOR_CSV_SAIDA, extrasaction="ignore"
        )
        escritor.writeheader()
        escritor.writerows(linhas)


def _ler_xlsx(p: Path) -> tuple:
    from openpyxl import load_workbook  # só exigido se houver .xlsx

    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        ws = wb.active
        iterador = ws.iter_rows(values_only=True)
        cabecalho_raw = next(iterador, None) or ()
        cabecalho = ["" if c is None else str(c).strip() for c in cabecalho_raw]
        linhas = []
        for valores in iterador:
            if valores is None or all(v in (None, "") for v in valores):
                continue
            linhas.append({
                h: ("" if i >= len(valores) or valores[i] is None else str(valores[i]).strip())
                for i, h in enumerate(cabecalho) if h
            })
    finally:
        wb.close()
    return [h for h in cabecalho if h], linhas


def _gravar_xlsx(p: Path, cabecalho: list, linhas: list) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"
    ws.append(cabecalho)
    for linha in linhas:
        ws.append([linha.get(h, "") for h in cabecalho])
    wb.save(p)


# ============ Interface ============

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Consulta de débitos")
        self.geometry("960x640")
        self.minsize(720, 480)

        self.caminho_entrada: Path | None = None
        self.cabecalho: list = []
        self.linhas: list = []
        self.fila: queue.Queue = queue.Queue()
        self.parar_evento = threading.Event()
        self.thread: threading.Thread | None = None

        self._montar_tela()
        self.after(100, self._drenar_fila)
        # Ao abrir, já pede a planilha. Se cancelar, a janela continua aberta.
        self.after(300, self.escolher_planilha)

    def _montar_tela(self) -> None:
        topo = ttk.Frame(self, padding=10)
        topo.pack(fill="x")
        ttk.Button(topo, text="Escolher planilha...", command=self.escolher_planilha).pack(side="left")
        self.lbl_arquivo = ttk.Label(topo, text="Nenhuma planilha selecionada (.xlsx ou .csv)")
        self.lbl_arquivo.pack(side="left", padx=10)

        previa = ttk.LabelFrame(self, text="Prévia da planilha de entrada", padding=6)
        previa.pack(fill="both", expand=True, padx=10)
        self.tabela = ttk.Treeview(previa, show="headings")
        rolagem_y = ttk.Scrollbar(previa, orient="vertical", command=self.tabela.yview)
        rolagem_x = ttk.Scrollbar(previa, orient="horizontal", command=self.tabela.xview)
        self.tabela.configure(yscrollcommand=rolagem_y.set, xscrollcommand=rolagem_x.set)
        self.tabela.grid(row=0, column=0, sticky="nsew")
        rolagem_y.grid(row=0, column=1, sticky="ns")
        rolagem_x.grid(row=1, column=0, sticky="ew")
        previa.rowconfigure(0, weight=1)
        previa.columnconfigure(0, weight=1)

        acoes = ttk.Frame(self, padding=10)
        acoes.pack(fill="x")
        self.btn_executar = ttk.Button(acoes, text="Executar", command=self.executar, state="disabled")
        self.btn_executar.pack(side="left")
        self.btn_parar = ttk.Button(acoes, text="Parar", command=self.parar, state="disabled")
        self.btn_parar.pack(side="left", padx=6)
        self.progresso = ttk.Progressbar(acoes, mode="determinate", length=260)
        self.progresso.pack(side="left", padx=10)
        self.lbl_status = ttk.Label(acoes, text="Pronto")
        self.lbl_status.pack(side="left")

        logs = ttk.LabelFrame(self, text="Log", padding=6)
        logs.pack(fill="both", padx=10, pady=(0, 10))
        self.log = tk.Text(logs, height=10, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True)

    def _escrever_log(self, mensagem: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", mensagem + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def escolher_planilha(self) -> None:
        caminho = filedialog.askopenfilename(
            title="Escolha a planilha",
            filetypes=[("Planilhas", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv")],
        )
        if not caminho:
            return
        try:
            cabecalho, linhas = ler_planilha(caminho)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Erro ao ler planilha", str(exc))
            return
        if not linhas:
            messagebox.showwarning("Planilha vazia", "Nenhuma linha com dados foi encontrada.")
            return

        self.caminho_entrada = Path(caminho)
        self.cabecalho = cabecalho
        self.linhas = linhas
        self._preencher_previa()
        self.lbl_arquivo.configure(text=f"{self.caminho_entrada.name} — {len(linhas)} linha(s)")
        self.btn_executar.configure(state="normal")
        self._escrever_log(f"Planilha carregada: {self.caminho_entrada.name} ({len(linhas)} linha(s))")

    def _preencher_previa(self) -> None:
        self.tabela.delete(*self.tabela.get_children())
        self.tabela["columns"] = self.cabecalho
        for coluna in self.cabecalho:
            self.tabela.heading(coluna, text=coluna)
            self.tabela.column(coluna, width=max(90, min(260, len(coluna) * 10)), stretch=True)
        for linha in self.linhas[:MAX_LINHAS_PREVIA]:
            self.tabela.insert("", "end", values=[linha.get(c, "") for c in self.cabecalho])

    def executar(self) -> None:
        if self.thread and self.thread.is_alive():
            return
        if not self.linhas:
            return
        self.parar_evento.clear()
        self.btn_executar.configure(state="disabled")
        self.btn_parar.configure(state="normal")
        self.progresso.configure(value=0)
        self.lbl_status.configure(text="Executando...")
        self.thread = threading.Thread(
            target=self._processar,
            args=(list(self.linhas), list(self.cabecalho), self.caminho_entrada),
            daemon=True,
        )
        self.thread.start()

    def parar(self) -> None:
        self.parar_evento.set()
        self.lbl_status.configure(text="Parando após o item atual...")

    def _processar(self, linhas, cabecalho, caminho_entrada: Path) -> None:
        """Roda em outra thread. Não toca na interface; só põe mensagens na fila."""
        total = len(linhas)
        resultados: list = []
        ok = erros = 0

        for i, registro in enumerate(linhas, start=1):
            if self.parar_evento.is_set():
                self.fila.put(("log", "Execução interrompida pelo usuário."))
                break

            rotulo = next((v for v in registro.values() if v), f"linha {i}")
            try:
                debitos = consultar(registro)
                if not debitos:
                    resultados.append({**registro, "status_consulta": "OK", "erro": "sem débitos"})
                for debito in debitos:
                    resultados.append({**registro, **debito, "status_consulta": "OK", "erro": ""})
                ok += 1
                self.fila.put(("log", f"[{i}/{total}] {rotulo}: {len(debitos)} débito(s)"))
            except ConsultaErro as exc:
                resultados.append({**registro, "status_consulta": "ERRO", "erro": str(exc)})
                erros += 1
                self.fila.put(("log", f"[{i}/{total}] {rotulo}: ERRO — {exc}"))
            except Exception as exc:  # noqa: BLE001
                resultados.append({**registro, "status_consulta": "ERRO", "erro": repr(exc)})
                erros += 1
                self.fila.put(("log", f"[{i}/{total}] {rotulo}: ERRO inesperado — {exc!r}"))

            self.fila.put(("progresso", i * 100 / total))

        if not resultados:
            self.fila.put(("fim", "Nenhum resultado gerado."))
            return

        colunas_saida = list(dict.fromkeys(cabecalho + COLUNAS_RESULTADO))
        try:
            saida = gravar_planilha(caminho_saida_padrao(caminho_entrada), colunas_saida, resultados)
        except PermissionError:
            self.fila.put(("erro", "Não foi possível gravar o resultado. Feche a planilha de saída se ela estiver aberta."))
            return
        except Exception as exc:  # noqa: BLE001
            self.fila.put(("erro", f"Falha ao gravar o resultado: {exc}"))
            return

        self.fila.put(("fim", f"Concluído: {ok} OK, {erros} com erro. Resultado salvo em {saida.name}"))

    def _drenar_fila(self) -> None:
        try:
            while True:
                tipo, valor = self.fila.get_nowait()
                if tipo == "log":
                    self._escrever_log(valor)
                elif tipo == "progresso":
                    self.progresso.configure(value=valor)
                elif tipo == "fim":
                    self._finalizar(valor)
                elif tipo == "erro":
                    self._finalizar(valor)
                    messagebox.showerror("Erro", valor)
        except queue.Empty:
            pass
        self.after(100, self._drenar_fila)

    def _finalizar(self, mensagem: str) -> None:
        self._escrever_log(mensagem)
        self.lbl_status.configure(text="Pronto")
        self.btn_executar.configure(state="normal" if self.linhas else "disabled")
        self.btn_parar.configure(state="disabled")


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
