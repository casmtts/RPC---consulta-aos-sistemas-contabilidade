#!/usr/bin/env python3
"""
Script para verificar a estrutura do sistema
"""

import os
from pathlib import Path

def verificar_estrutura():
    """Verificar se todos os arquivos necessários existem"""
    base_dir = Path("/Users/casmtts/Documents/TRIBUTOS E TERRA/RPC - Consulta de débitos")
    
    print("=" * 80)
    print("VERIFICAÇÃO DA ESTRUTURA DO SISTEMA")
    print("=" * 80)
    
    # Lista de arquivos esperados
    arquivos_esperados = [
        "main.py",
        "config.py", 
        "models.py",
        "database.py",
        "scraper/scraper.py",
        "analyze.py",
        "test_sample.py",
        "test_sample.py",
        "test_sample.py",
        "install.sh",
        "requirements.txt",
        "README.md",
        "GUIDE_DE_USO.txt"
    ]
    
    print("\\n📁 Arquivos principais:")
    print("-" * 80)
    
    todos_existem = True
    
    for arquivo in arquivos_esperados:
        caminho = base_dir / arquivo
        existe = caminho.exists()
        status = "✅" if existe else "❌"
        print(f"{status} {arquivo}")
        
        if not existe:
            todos_existem = False
    
    print("\\n" + "=" * 80)
    
    if todos_existem:
        print("✅ Todos os arquivos principais estão presentes!")
        print("\\nPróximos passos:")
        print("1. Instalar dependências: pip install -r requirements.txt")
        print("2. Instalar Playwright: playwright install")
        print("3. Instalar navegador: playwright install chromium")
        print("4. Analisar página: python analyze.py")
        print("5. Testar: python test_sample.py")
        print("6. Rodar o sistema: python main.py")
    else:
        print("❌ Alguns arquivos estão ausentes!")
        print("\\nPara adicionar arquivos faltantes, use:")
        print("  touch /caminho/filho_faltante")
    
    print("\\n" + "=" * 80)
    
    # Verificar diretório scraper
    print("\\n📁 Verificando diretório scraper:")
    scraper_dir = base_dir / "scraper"
    if scraper_dir.exists():
        print(f"✅ Directório scraper existe")
        scraper_files = list(scraper_dir.glob("*"))
        print(f"   Arquivos no scraper: {len(scraper_files)}")
    else:
        print("❌ Directório scraper não existe")
    
    print("\\n" + "=" * 80)
    
    # Tamanho total dos arquivos
    print("\\n📊 Resumo:")
    print("-" * 80)
    
    total_tamanho = 0
    for arquivo in base_dir.glob("*"):
        if arquivo.is_file() and arquivo.suffix != ".pyc":
            total_tamanho += arquivo.stat().st_size
            print(f"📄 {arquivo.name}: {arquivo.stat().st_size / 1024:.2f} KB")
    
    print(f"\\nTotal: {total_tamanho / 1024 / 1024:.2f} MB")
    
    print("\\n" + "=" * 80)
    print("VERIFICAÇÃO COMPLETADA")
    print("=" * 80)

if __name__ == "__main__":
    verificar_estrutura()
