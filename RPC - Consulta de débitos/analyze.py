"""
Script para analisar a estrutura do DOM da página de consulta de débitos
"""

from pathlib import Path
from typing import List, Dict, Optional
import sys

sys.path.insert(0, str(Path(__file__).parent))

from playwright.sync_api import sync_playwright
from scraper.scraper import DebitorScraper


def analisar_pagina():
    """
    Analisa a estrutura da página para identificar:
    - Elementos do formulário (campos CPF, Nome, botão pesquisar)
    - Tabelas de resultados (colunas e índices)
    """
    url = "https://santanadoaraguaia-pa.desenvolvecidade.com.br/nfsd/pages/relatorio/relatorioLancamentoDebito.jsf"
    
    print("=" * 80)
    print("ANÁLISE DA ESTRUTURA DO DOM")
    print("=" * 80)
    
    # Abrir navegador
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        
        # Acessar a página
        print(f"\nAcessando URL: {url}")
        page.goto(url, wait_until="networkidle", timeout=30000)
        
        # Identificar elementos do formulário
        print("\n" + "=" * 80)
        print("1. ELEMENTOS DO FORMULÁRIO")
        print("=" * 80)
        
        # Encontrar todos os inputs
        inputs = page.query_selector_all("input, select, textarea")
        print(f"\nTotal de inputs: {len(inputs)}")
        
        # Encontrar botões
        buttons = page.query_selector_all("button, [role='button']")
        print(f"\nTotal de botões: {len(buttons)}")
        
        # Identificar tabelas
        print("\n" + "=" * 80)
        print("2. TABELAS DE RESULTADOS")
        print("=" * 80)
        
        tables = page.query_selector_all("table")
        print(f"\nTotal de tabelas: {len(tables)}")
        
        for i, tabela in enumerate(tables, 1):
            print(f"\nTabela {i}:")
            print(f"  Tag: {tabela.tag_name()}")
            print(f"  Atributos: {dict(tabela.attributes)}")
            
            # Encontrar cabeçalho
            headers = tabela.query_selector_all("th")
            print(f"  Cabeçalhos ({len(headers)}):")
            
            for j, th in enumerate(headers, 1):
                text = th.inner_text() or th.get_attribute("aria-label", "") or th.get_attribute("title", "")
                print(f"    Coluna {j}: '{text}'")
            
            # Encontrar linhas
            rows = tabela.query_selector_all("tr")
            print(f"  Linhas totais: {len(rows)}")
        
        # Verificar elementos com classes específicas
        print("\n" + "=" * 80)
        print("3. BUSCA POR CLASSES ESPECÍFICAS")
        print("=" * 80)
        
        classes_to_find = ["debito", "tipo", "tipo_debito", "competencia", 
                         "valor", "ano", "situacao", "status"]
        
        for classe in classes_to_find:
            elements = page.query_selector_all(f"[class*='{classe}']")
            if elements:
                print(f"\nClasses '{classe}': {len(elements)} elementos encontrados")
        
        browser.close()
        
        print("\n" + "=" * 80)
        print("ANÁLISE COMPLETA! Verifique o console para ver detalhes.")
        print("=" * 80)


if __name__ == "__main__":
    try:
        analisar_pagina()
    except Exception as e:
        print(f"\nErro ao analisar página: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
