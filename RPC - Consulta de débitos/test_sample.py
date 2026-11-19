"""
Teste e exemplo de uso do sistema de consulta de débitos
"""

from pathlib import Path
from config import BASE_URL, LOTES_POR_CONSULTA, BROWSER_TIMEOUT
from models import StatusConsulta, PLANILLA_CABECALOS
from database import DebitorDatabase
from scraper.scraper import DebitorScraper


def test_scraper():
    """Teste básico do scraper"""
    print("=" * 80)
    print("TESTE DO SCRAPER")
    print("=" * 80)
    
    # Criar scraper
    scraper = DebitorScraper(base_url=BASE_URL, timeout=30)
    
    # Testar consulta (use CPF e nome válidos para teste)
    cpf_teste = "123.456.789-00"
    nome_teste = "João Silva"
    
    print(f"\nConsultando CPF: {cpf_teste}, Nome: {nome_teste}")
    
    try:
        debitos = scraper.consultar_debito(cpf_teste, nome_teste, timeout=30)
        
        print(f"\nResultados: {len(debitos)} débito(s) encontrado(s)")
        
        for i, debito in enumerate(debitos, 1):
            print(f"\n{i}. Tipo: {debito['tipo_debito']}")
            print(f"   Ano: {debito['ano']}")
            print(f"   Situacao: {debito['situacao']}")
            print(f"   Valor: R$ {debito['valor']:.2f}")
            
    except Exception as e:
        print(f"\nErro na consulta: {e}")
        print("\nDicas:")
        print("1. Verifique se a URL está correta")
        print("2. Verifique se o navegador pode acessar a página")
        print("3. Tente abrir manualmente a URL e verifique os campos")
        print("4. Analise a página com: python analyze.py")
    
    return debitos


def test_database():
    """Teste básico do database"""
    print("\n" + "=" * 80)
    print("TESTE DO DATABASE")
    print("=" * 80)
    
    # Criar instance do database
    db = DebitorDatabase("teste_input.csv", "teste_output.xlsx", "test_errors.log")
    
    # Criar arquivo de teste de entrada
    input_path = Path("teste_input.csv")
    
    # Escrever arquivo de teste
    with input_path.open("w", encoding="utf-8-sig") as f:
        f.write("Nome,CPF,Tipo Débito,Competência,Valor (R$)\n")
        f.write("João Silva,123.456.789-00,ICMS,2024,150.00\n")
        f.write("Maria Santos,987.654.321-00,ISS,2024,200.00\n")
        f.write("José Pereira,456.789.123-00,ICMS,2024,100.00\n")
    
    print(f"\nArquivo de teste criado: {input_path.name}")
    
    # Carregar
    try:
        db.carregar_planilha_input(str(input_path))
        print(f"\nRegistros carregados: {len(db.input_data)}")
        
        for cpf, dados in db.input_data.items():
            print(f"  CPF: {cpf}, Nome: {dados.get('nome', 'N/A')}")
            
    except Exception as e:
        print(f"\nErro ao carregar input: {e}")
    
    # Salvar resultado
    output_path = Path("teste_output.xlsx")
    
    try:
        # Criar dados de saída
        dados_saida = [
            {
                "Nome": "João Silva",
                "CPF": "123.456.789-00",
                "Tipo de Débito": "ICMS",
                "Competência": "2024",
                "Valor (R$)": "150.00",
                "Status": "COMPLETO"
            },
            {
                "Nome": "Maria Santos",
                "CPF": "987.654.321-00",
                "Tipo de Débito": "ISS",
                "Competência": "2024",
                "Valor (R$)": "200.00",
                "Status": "COMPLETO"
            }
        ]
        
        db.salvar_planilha_output(dados_saida, output_path)
        print(f"\nPlanilha salva: {output_path.name}")
        
    except Exception as e:
        print(f"\nErro ao salvar output: {e}")


if __name__ == "__main__":
    print("\n=== Testando Sistema de Consulta de Débitos ===\n")
    
    # Testar scraper
    debitos = test_scraper()
    
    # Testar database
    test_database()
    
    print("\n" + "=" * 80)
    print("TESTES CONCLUÍDOS")
    print("=" * 80)
    print("\nPróximos passos:")
    print("1. Realizar análise da página: python analyze.py")
    print("2. Identificar elementos do DOM")
    print("3. Atualizar scraper.py com os elementos corretos")
    print("4. Testar com dados reais")
    print("5. Rodar o sistema principal: python main.py")
