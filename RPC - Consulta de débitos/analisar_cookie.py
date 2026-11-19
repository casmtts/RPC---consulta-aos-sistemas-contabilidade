"""
Script para analisar a página e identificar elementos do DOM
Requer um cookie de sessão válido
"""

from pathlib import Path
from playwright.sync_api import sync_playwright
from scraper.scraper import DebitorScraper

# Cookie do usuário (substitua pelo cookie real)
COOKIE_VALUE = ""  # Cole seu cookie aqui

def analisar_pagina(cookie_value=""):
    """
    Analisa a estrutura da página para identificar:
    - Elementos do formulário (campos CPF, Nome, botão pesquisar)
    - Tabelas de resultados (colunas e índices)
    """
    url = "https://santanadoaraguaia-pa.desenvolvecidade.com.br/nfsd/pages/relatorio/relatorioLancamentoDebito.jsf"
    
    print("=" * 80)
    print("ANÁLISE DA ESTRUTURA DO DOM")
    print("=" * 80)
    
    # Criar scraper com cookie
    scraper = DebitorScraper(
        base_url=url,
        timeout=30,
        cookie_value=cookie_value
    )
    
    # Abrir navegador
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        
        # Acessar a página
        print(f"\nAcessando URL: {url}")
        page.goto(url, wait_until="networkidle", timeout=30000)
        
        # Preencher formulário
        print("\n" + "=" * 80)
        print("1. PREENCHENDO FORMULÁRIO COM TESTE")
        print("=" * 80)
        
        cpf_teste = "123.456.789-00"
        nome_teste = "Teste"
        
        print(f"\nTestando com CPF: {cpf_teste}, Nome: {nome_teste}")
        
        try:
            scraper._preencher_formulario(page, cpf_teste, nome_teste)
            print("\nFormulário preenchido com sucesso!")
        except Exception as e:
            print(f"\nErro ao preencher formulário: {e}")
            import traceback
            traceback.print_exc()
        
        # Clicar no botão pesquisar
        print("\n" + "=" * 80)
        print("2. CLICANDO NO BOTÃO PESQUISAR")
        print("=" * 80)
        
        try:
            scraper._executar_pesquisa(page)
            print("\nBotão pesquisar clicado com sucesso!")
        except Exception as e:
            print(f"\nErro ao clicar no botão: {e}")
            import traceback
            traceback.print_exc()
        
        # Esperar resultados
        print("\n" + "=" * 80)
        print("3. ESPERANDO RESULTADOS (2 segundos)")
        print("=" * 80)
        import time
        time.sleep(2)
        
        # Extrair dados da tabela
        print("\n" + "=" * 80)
        print("4. EXTRAINDO DADOS DA TABELA")
        print("=" * 80)
        
        try:
            debitos = scraper._extrair_dados_tabela(page)
            print(f"\nExtraindo {len(debitos)} débito(s)")
            
            if debitos:
                print("\nPrimeiro débito extraído:")
                for i, d in enumerate(debitos[:3], 1):  # Mostrar apenas os 3 primeiros
                    print(f"  {i}. Tipo: {d.get('tipo_debito', 'N/A')}")
                    print(f"     Competência: {d.get('competencia', 'N/A')}")
                    print(f"     Valor: {d.get('valor', 'N/A')}")
            else:
                print("\nNenhum débito encontrado!")
                print("Verificando se há tabelas na página...")
                tables = page.query_selector_all("table")
                print(f"Encontradas {len(tables)} tabelas")
                
                for i, tabela in enumerate(tables, 1):
                    print(f"\nTabela {i}:")
                    headers = tabela.query_selector_all("th, td")
                    print("  Colunas:")
                    for j, cell in enumerate(headers[:10], 1):  # Mostrar apenas as 10 primeiras
                        text = cell.inner_text() or cell.get_attribute("aria-label") or ""
                        print(f"    {j}. '{text}'")
            
        except Exception as e:
            print(f"\nErro ao extrair dados: {e}")
            import traceback
            traceback.print_exc()
        
        browser.close()
        
        print("\n" + "=" * 80)
        print("ANÁLISE COMPLETA! Verifique o console para ver detalhes.")
        print("=" * 80)

if __name__ == "__main__":
    try:
        if not analisar_pagina(COOKIE_VALUE):
            print("\nErro na análise!")
    except Exception as e:
        print(f"\nErro ao analisar página: {e}")
        import traceback
        traceback.print_exc()
