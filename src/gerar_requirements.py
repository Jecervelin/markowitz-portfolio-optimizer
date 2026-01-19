import os

# Lista exata das bibliotecas usadas no seu projeto Markowitz Pro
libs = [
    "numpy",
    "pandas",
    "yfinance",
    "matplotlib",
    "PyPortfolioOpt",
    "scikit-learn",  # Necessário para o Ledoit-Wolf Shrinkage
    "xlsxwriter",    # Gráficos no Excel
    "openpyxl",      # Leitura de Excel/Config
    "fpdf"           # Geração do PDF
]

file_name = "requirements.txt"

try:
    with open(file_name, "w") as f:
        f.write("# Dependencias do Projeto Markowitz Pro\n")
        for lib in libs:
            f.write(f"{lib}\n")
            
    print(f"Sucesso! Arquivo '{file_name}' gerado com as {len(libs)} bibliotecas essenciais.")
    print("Para instalar em outro PC, use: pip install -r requirements.txt")
    
except Exception as e:
    print(f"Erro ao criar arquivo: {e}")