"""
Gerenciamento de planilhas e dados
"""

from pathlib import Path
from typing import List, Dict, Tuple
from datetime import datetime
import pandas as pd
from openpyxl import Workbook, load_workbook


class DebitorDatabase:
    """Gerenciamento de planilhas e dados de débitos"""
    
    def __init__(self, input_file: str, output_file: str, error_log: str):
        self.input_file = Path(input_file)
        self.output_file = Path(output_file)
        self.error_log = Path(error_log)
        self.input_data: Dict = {}
        
    def carregar_planilha_input(self) -> Dict:
        """Carregar planilha de entrada com lista de CPFs e nomes"""
        if not self.input_file.exists():
            raise FileNotFoundError(f"Arquivo de entrada não encontrado: {self.input_file}")
        
        # Tentar detectar formato e carregar
        ext = self.input_file.suffix.lower()
        
        if ext == ".csv":
            self.input_data = self._ler_csv()
        elif ext in [".xlsx", ".xls"]:
            self.input_data = self._ler_xlsx()
        else:
            raise ValueError(f"Formato não suportado: {ext}")
        
        print(f"Planilha carregada: {self.input_file.name} ({len(self.input_data)} registros)")
        return self.input_data
    
    def _ler_csv(self) -> Dict:
        """Ler arquivo CSV"""
        cabecalho = None
        linhas = {}
        
        with self.input_file.open("r", encoding="utf-8-sig") as f:
            import csv
            leitor = csv.DictReader(f)
            cabecalho = leitor.fieldnames or []
            
            for linha in leitor:
                # Converter lista de valores em dicionário
                item = {}
                for i, val in enumerate(linha):
                    if i < len(cabecalho):
                        item[cabecalho[i]] = val.strip()
                linhas.update(item)
        
        return linhas if cabecalho else {}
    
    def _ler_xlsx(self) -> Dict:
        """Ler arquivo XLSX"""
        wb = load_workbook(self.input_file, read_only=True, data_only=True)
        ws = wb.active
        
        cabecalho = None
        linhas = {}
        
        for row in ws.iter_rows(values_only=True):
            # Converter lista de valores em dicionário
            if cabecalho is None:
                cabecalho = [str(cell) if cell is not None else "" for cell in row]
                continue
            
            if all(cell in (None, "") for cell in row):
                continue
            
            item = {}
            for i, val in enumerate(row):
                if i < len(cabecalho):
                    item[cabecalho[i]] = str(val).strip() if val is not None else ""
            linhas.update(item)
        
        wb.close()
        return linhas if cabecalho else {}
    
    def salvar_planilha_output(self, dados: List[Dict]) -> Path:
        """Salvar planilha de resultados"""
        if not dados:
            raise ValueError("Nenhum dado para salvar")
        
        # Preparar cabecalhos (padronizado)
        cabecalho = ["Nome", "CPF", "Tipo de Débito", "Competência", "Valor (R$)", "Status"]
        
        # Converter dados para lista de dicionários
        lista_dados = []
        for d in dados:
            item = {
                "Nome": d.get("nome", ""),
                "CPF": d.get("cpf", ""),
                "Tipo de Débito": d.get("tipo_debito", ""),
                "Competência": d.get("competencia", ""),
                "Valor (R$)": str(d.get("valor", 0.0)).replace('.', ','),
                "Status": d.get("status", "").upper()
            }
            lista_dados.append(item)
        
        # Criar arquivo
        try:
            output_path = self.output_file
            ext = output_path.suffix.lower()
            
            if ext in [".csv", ""]:
                self._gravar_csv(output_path, cabecalho, lista_dados)
            elif ext in [".xlsx", ".xls"]:
                self._gravar_xlsx(output_path, cabecalho, lista_dados)
            else:
                raise ValueError(f"Formato de saída não suportado: {ext}")
            
            print(f"Planilha salva: {output_path.name}")
            return output_path
            
        except PermissionError:
            raise
        except Exception as e:
            raise Exception(f"Erro ao salvar: {e}")
    
    def _gravar_csv(self, path: Path, cabecalho: List, dados: List[Dict]):
        """Gravar planilha CSV"""
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            import csv
            escritor = csv.DictWriter(f, fieldnames=cabecalho, extrasaction="ignore")
            escritor.writeheader()
            escritor.writerows(dados)
    
    def _gravar_xlsx(self, path: Path, cabecalho: List, dados: List[Dict]):
        """Gravar planilha XLSX"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultado"
        ws.append(cabecalho)
        
        for d in dados:
            row = []
            for h in cabecalho:
                val = d.get(h, "")
                if h == "Valor (R$)":
                    # Converter para formato decimal
                    if str(val):
                        val = str(float(val).replace(",", "."))
                row.append(str(val) if val is not None else "")
            ws.append(row)
        
        wb.save(path)
    
    def salvar_log_erros(self, erro: str, cpf: str, nome: str):
        """Salvar erro no log"""
        try:
            with open(self.error_log, "a", encoding="utf-8") as f:
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                erro_msg = f"[{timestamp}] CPF: {cpf} | Nome: {nome} | Erro: {erro}"
                f.write(erro_msg + "\n")
            print(f"Erro registrado: {cpf} - {erro}")
        except Exception as e:
            print(f"Erro ao salvar log: {e}")
