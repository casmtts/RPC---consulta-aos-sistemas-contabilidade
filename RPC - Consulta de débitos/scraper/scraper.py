"""
Scraper com Playwright para consultar débitos
"""

import os
from pathlib import Path
from typing import List, Dict, Optional
from playwright.sync_api import sync_playwright, Page, Frame
import time
import requests
from urllib.parse import quote_plus


class DebitorScraper:
    """Scraper para consultar débitos via Playwright"""
    
    BASE_URL = "https://santanadoaraguaia-pa.desenvolvecidade.com.br/nfsd/pages/relatorio/relatorioLancamentoDebito.jsf"
    TIMEOUT = 30  # segundos
    RETRY_ATTEMPTS = 3
    EXTRACTION_TIMEOUT = 1.5  # segundos entre consultas
    
    def __init__(self, base_url: str = None, timeout: int = None, retries: int = None, cookie_value: str = None):
        self.base_url = base_url or self.BASE_URL
        self.timeout = timeout or self.TIMEOUT
        self.retries = retries or self.RETRY_ATTEMPTS
        self.cookie_value = cookie_value or ""
        self.driver = None
        self.current_page = None
        self._load_playwright()
        
    def _load_playwright(self):
        """Inicializar Playwright"""
        # Instalar browsers se necessário
        self._ensure_browsers()
        # Se tiver cookie, aplicar
        if self.cookie_value:
            self._apply_cookie()
    
    def _apply_cookie(self):
        """Aplicar cookie ao navegador"""
        try:
            # Adicionar cookie ao Chromium
            cookie = {
                'name': 'sessionid',  # Nome genérico (será substituído pelo cookie real)
                'value': self.cookie_value,
                'domain': '.desenvolvecidade.com.br',
                'path': '/',
                'secure': True,
                'sameSite': 'Lax'
            }
            self.driver.add_cookies([cookie])
            print("Cookie aplicado com sucesso")
        except Exception as e:
            print(f"Erro ao aplicar cookie: {e}")
        
    def _ensure_browsers(self):
        """Instalar browsers se necessário"""
        # Verificar se browsers estão instalados
        import subprocess
        try:
            subprocess.run(["playwright", "install"], check=True, timeout=60)
            print("Browsers Playwright instalados")
        except Exception as e:
            print(f"Erro ao instalar browsers: {e}")
            print("Execute: playwright install")
    
    def update_cookie(self, cookie_value: str):
        """Atualizar o cookie de sessão"""
        self.cookie_value = cookie_value
        self._apply_cookie()
        print(f"Cookie atualizado: {cookie_value[:50]}...")
    
    def consultar_debito(self, cpf: str, nome: str, timeout: int = None) -> List[Dict]:
        """
        Consultar débitos para um CPF e nome específico
        
        Args:
            cpf: CPF do devedor (formato: 000.000.000-00)
            nome: Nome do devedor
            timeout: Timeout personalizado (padrão: 30s)
            
        Returns:
            Lista de débitos encontrados
            
        Raises:
            Exception: Se a consulta falhar (sessão expirada, timeout, etc)
        """
        debitos = []
        timeout_custom = timeout or self.timeout
        
        try:
            # Verificar se cookie está válido
            if not self.cookie_value:
                print("⚠️  Nenhum cookie configurado. Acesso sem autenticação.")
                # Tentar acessar sem cookie (pode falhar se exigir login)
            else:
                print(f"🔑 Usando cookie de sessão: {self.cookie_value[:30]}...")
            
            # Abrir navegador
            self.driver.start()
            browser = self.driver.chromium.launch(headless=True)
            page = browser.new_page()
            
            print(f"Consultando CPF: {cpf}, Nome: {nome}")
            print(f"Timeout: {timeout_custom}s")
            
            # Navegar até a página
            print(f"Navegando para: {self.base_url}")
            
            # Preencher URL com CPF e Nome
            url_completo = self.base_url
            
            # Tenta acessar diretamente primeiro
            try:
                page.goto(url_completo, wait_until="domcontentloaded", timeout=timeout_custom * 2)
                print("Página acessada")
            except Exception as e:
                print(f"⚠️  Erro ao acessar: {e}")
                raise
            
            # Esperar por elementos carregados (1,5s)
            time.sleep(self.extraction_timeout)
            
            # Preencher formulário
            self._preencher_formulario(page, cpf, nome)
            
            # Clicar no botão pesquisar
            self._executar_pesquisa(page)
            
            # Esperar pelos resultados (1,5s)
            time.sleep(self.extraction_timeout)
            
            # Extrair dados da tabela
            print("Extraindo dados da tabela...")
            debitos = self._extrair_dados_tabela(page)
            
            # Limpar navegador
            browser.close()
            
            print(f"✅ Consulta concluída: {len(debitos)} débito(s) encontrado(s)")
            
        except Exception as e:
            error_msg = f"Erro na consulta: {e}"
            print(f"❌ {error_msg}")
            raise
            
        return debitos
    
    def _preencher_formulario(self, page: Page, cpf: str, nome: str):
        """Preencher o formulário com CPF e Nome"""
        try:
            print("Preenchendo formulário...")
            
            # Tenta várias formas de encontrar os campos
            # 1. Por nome
            cpf_input = page.query_selector("input[name='cpf']")
            if cpf_input:
                cpf_input.fill(cpf)
                print(f"  CPF preenchido (via name='cpf'): {cpf}")
            else:
                # 2. Por ID - tenta
                cpf_input = page.query_selector("input[id*='cpf']")
                if cpf_input:
                    cpf_input.fill(cpf)
                    print(f"  CPF preenchido (via id*='cpf'): {cpf}")
                else:
                    print(f"  ⚠️  Campo CPF não encontrado")
            
            # 3. Tenta por aria-label
            cpf_input = page.query_selector('[aria-label*="CPF"]')
            if cpf_input:
                cpf_input.fill(cpf)
                print(f"  CPF preenchido (via aria-label): {cpf}")
            
            # Preencher nome
            nome_input = page.query_selector("input[name='nome']")
            if nome_input:
                nome_input.fill(nome)
                print(f"  Nome preenchido (via name='nome'): {nome}")
            else:
                # Tenta por ID
                nome_input = page.query_selector("input[id*='nome']")
                if nome_input:
                    nome_input.fill(nome)
                    print(f"  Nome preenchido (via id*='nome'): {nome}")
                else:
                    print(f"  ⚠️  Campo Nome não encontrado")
            
            # 4. Tenta por aria-label
            nome_input = page.query_selector('[aria-label*="Nome"]')
            if nome_input:
                nome_input.fill(nome)
                print(f"  Nome preenchido (via aria-label): {nome}")
            
            print("  Formulário preenchido")
            
        except Exception as e:
            print(f"Erro ao preencher formulário: {e}")
            raise
    
    def _executar_pesquisa(self, page: Page):
        """Executar a busca preenchendo o formulário"""
        try:
            print("Executando pesquisa...")
            
            # Tenta várias formas de encontrar o botão pesquisar
            # 1. Por role/aria-label
            buscar_button = page.query_selector("button[role='button'][value='Pesquisar']")
            if buscar_button:
                buscar_button.click()
                print("  Botão 'Pesquisar' clicado (via role/value)")
            else:
                # 2. Por ID
                buscar_button = page.query_selector("button[id*='pesquisar']")
                if buscar_button:
                    buscar_button.click()
                    print("  Botão 'Pesquisar' clicado (via id*='pesquisar')")
                else:
                    # 3. Por aria-label
                    buscar_button = page.query_selector('[aria-label*="Pesquisar"]')
                    if buscar_button:
                        buscar_button.click()
                        print("  Botão 'Pesquisar' clicado (via aria-label)")
                    else:
                        # 4. Por texto
                        buscar_button = page.query_selector("button:has-text('Pesquisar')")
                        if buscar_button:
                            buscar_button.click()
                            print("  Botão 'Pesquisar' clicado (via texto)")
                        else:
                            print("  ⚠️  Botão 'Pesquisar' não encontrado")
            
            print("Pesquisa executada")
            
        except Exception as e:
            print(f"Erro ao executar pesquisa: {e}")
            raise
    
    def _extrair_dados_tabela(self, page: Page) -> List[Dict]:
        """
        Extrair dados da tabela de resultados
        
        Returns:
            Lista de débitos com estrutura:
            [
                {
                    'tipo_debito': str,
                    'competencia': str,
                    'situacao': str,
                    'valor': float
                },
                ...
            ]
        """
        debitos = []
        
        try:
            print("Extraindo dados da tabela...")
            
            # Encontrar todas as tabelas
            tables = page.query_selector_all("table")
            
            for tabela in tables:
                print(f"  Analisando tabela...")
                
                # Encontrar cabeçalho
                headers = tabela.query_selector_all("th")
                
                if not headers:
                    print("    Tabela sem cabeçalho")
                    continue
                
                # Extrair cabeçalhos
                cabecalho = []
                for th in headers:
                    text = self._get_text_element(th)
                    if text:
                        cabecalho.append(text.strip())
                
                print(f"    Cabeçalhos ({len(cabecalho)}): {cabecalho}")
                
                # Encontrar linhas de dados
                rows = tabela.query_selector_all("tr")
                
                for row in rows[1:]:  # Pular cabeçalho
                    dados = {}
                    
                    # Encontrar células
                    cells = row.query_selector_all("td, th")
                    
                    for cell in cells:
                        text = self._get_text_element(cell)
                        if text:
                            # Encontrar índice da coluna
                            col_index = cabecalho.index(text.strip()) if text.strip() in cabecalho else -1
                            if col_index >= 0:
                                dados[col_index] = text.strip()
                    
                    # Construir objeto de débito
                    if dados:
                        debito = self._transformar_dados_para_debito(dados, cabecalho)
                        debitos.append(debito)
            
            print(f"✅ Extraído {len(debitos)} débito(s)")
            
        except Exception as e:
            print(f"⚠️  Erro ao extrair dados: {e}")
            # Se falhar, tentar extrair de forma mais simples
            print("Tentando extrair de forma simplificada...")
            debitos = self._extrair_dados_simplificado(page)
        
        return debitos
    
    def _extrair_dados_simplificado(self, page: Page) -> List[Dict]:
        """
        Extrair dados da tabela de forma simplificada (fallback)
        """
        debitos = []
        
        try:
            print("Extraindo dados de forma simplificada...")
            
            # Tentar encontrar tabelas com classes específicas
            tables = page.query_selector_all("table.debito, table[type='debito'], table[title*='débito']")
            
            for tabela in tables:
                # Extrair cabeçalho
                headers = tabela.query_selector_all("th")
                cabecalho = []
                
                for th in headers:
                    text = th.inner_text().strip()
                    if text.lower() in ['tipo', 'tipo de débito', 'competência', 'ano', 'situacao', 'valor']:
                        cabecalho.append(text)
                
                if not cabecalho:
                    continue
                
                # Extrair linhas
                rows = tabela.query_selector_all("tr")
                
                for row in rows:
                    cells = row.query_selector_all("td")
                    
                    if len(cells) >= len(cabecalho):
                        item = {}
                        for i, cell in enumerate(cells[:len(cabecalho)]):
                            text = cell.inner_text().strip()
                            item[cabecalho[i]] = text
                        
                        # Transformar
                        debito = self._transformar_dados_para_debito(item, cabecalho)
                        debitos.append(debito)
            
            print(f"✅ Extraído {len(debitos)} débito(s) (método simplificado)")
            
        except Exception as e:
            print(f"⚠️  Erro no método simplificado: {e}")
        
        return debitos
    
    def _get_text_element(self, element) -> str:
        """Obter texto de um elemento"""
        try:
            # Try inner_text first
            text = element.inner_text()
            if text:
                return text
            
            # Try aria-label
            aria_label = element.get_attribute('aria-label')
            if aria_label:
                return aria_label
            
            # Try title
            title = element.get_attribute('title')
            if title:
                return title
            
            # Try text_content
            text_content = element.text_content()
            if text_content:
                return str(text_content).strip()
                
        except Exception:
            pass
        
        return ""
    
    def _transformar_dados_para_debito(self, dados: Dict, cabecalho: List[str]) -> Dict:
        """
        Transformar dados da tabela para estrutura de débito
        
        Args:
            dados: Dict com dados por índice de coluna
            cabecalho: Lista de cabeçalhos da tabela
            
        Returns:
            Dict com estrutura de débito
        """
        # Mapear cabecalhos para índices
        cabecalho_index = {col: idx for idx, col in enumerate(cabecalho)}
        
        # Encontrar valores relevantes
        tipo_debito = dados.get(cabecalho_index.get('tipo_debito', 0), '')
        competencia = dados.get(cabecalho_index.get('competencia', 0), '')
        situacao = dados.get(cabecalho_index.get('situacao', 0), '')
        valor_str = dados.get(cabecalho_index.get('valor', 0), '')
        
        # Transformar valor
        valor = 0.0
        if valor_str:
            # Normalizar e converter para float
            valor_str = str(valor_str).replace('R$', '').replace('r$', '')  # Remove moeda
            valor_str = ''.join(c for c in valor_str if c.isdigit() or c in ',.')  # Mantém apenas dígitos e ponto/vírgula
            try:
                valor = float(valor_str.replace(',', '.'))
            except ValueError:
                print(f"  ⚠️  Valor não numérico: {valor_str}")
                valor = 0.0
        else:
            print(f"  ⚠️  Valor não encontrado")
        
        return {
            'tipo_debito': tipo_debito,
            'competencia': competencia,
            'situacao': situacao,
            'valor': valor
        }
