from fpdf import FPDF
import os

class ModernPDF(FPDF):
    def header(self):
        # Cabeçalho com fundo Azul Escuro
        self.set_font('Arial', 'B', 16)
        self.set_fill_color(44, 62, 80) 
        self.set_text_color(255, 255, 255) 
        self.cell(0, 15, "Otimizador de Portfólio & Backtest (Markowitz Pro)", 0, 1, 'C', 1)
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128)
        self.cell(0, 10, f'Pagina {self.page_no()}', 0, 0, 'C')

    def chapter_title(self, label):
        self.set_font('Arial', 'B', 12)
        self.set_text_color(44, 62, 80)
        self.set_fill_color(230, 230, 230)
        self.cell(0, 8, f"  {label}", 0, 1, 'L', 1)
        self.ln(4)

    def chapter_body(self, text):
        self.set_font('Arial', '', 10)
        self.set_text_color(0)
        text = text.replace('“', '"').replace('”', '"').replace('’', "'")
        text = text.encode('latin-1', 'replace').decode('latin-1')
        self.multi_cell(0, 6, text)
        self.ln(2)

    def chapter_list(self, items):
        self.set_font('Arial', '', 10)
        self.set_text_color(0)
        for item in items:
            text = f"{chr(149)} {item}"
            text = text.replace('“', '"').replace('”', '"').replace('’', "'")
            text = text.encode('latin-1', 'replace').decode('latin-1')
            self.multi_cell(0, 6, text)
        self.ln(4)
    
    def code_block(self, lines):
        self.set_font('Courier', '', 9)
        self.set_fill_color(245, 245, 245)
        self.set_text_color(0, 0, 0)
        
        for line in lines:
            line = line.encode('latin-1', 'replace').decode('latin-1')
            self.cell(0, 5, f"  {line}", 0, 1, 'L', 1)
        self.ln(4)
        self.set_text_color(0)

# --- CONTEÚDO ---
pdf = ModernPDF()
pdf.add_page()

# 1. INTRO
pdf.chapter_body(
    "Ferramenta de engenharia financeira em Python que une a Teoria Moderna do Portfólio (Markowitz) "
    "com validação prática via Backtest. O sistema calcula a alocação ideal em dois cenários "
    "e compara a performance da IA contra uma carteira manual."
)
pdf.ln(2)

# 2. FUNCIONALIDADES
pdf.chapter_title("Funcionalidades Principais")
funcionalidades = [
    "Otimização de Média-Variância (PyPortfolioOpt).",
    "Matriz de Covariância Robusta (Ledoit-Wolf Shrinkage).",
    "Cenários Duplos: Sem Restrições (0-100%) vs. Com Restrições (Compliance).",
    "Automação Inteligente: Leitura automática de configurações.",
    "Relatórios Excel: Dashboards com gráficos nativos."
]
pdf.chapter_list(funcionalidades)

# 3. INSTALAÇÃO
pdf.chapter_title("Instalação e Dependências")
pdf.chapter_body("Instale as bibliotecas necessárias com o comando abaixo:")
comandos_install = [
    "pip install numpy pandas yfinance matplotlib",
    "pip install PyPortfolioOpt xlsxwriter openpyxl scikit-learn"
]
pdf.code_block(comandos_install)

# 4. COMO UTILIZAR (AGORA COM ARQUIVOS DE SAÍDA)
pdf.chapter_title("Fluxo de Trabalho e Arquivos Gerados")

# PASSO 1
pdf.set_font('Arial', 'B', 10)
pdf.cell(0, 6, "Passo 1: Configurar", 0, 1)
pdf.set_font('Arial', '', 10)
pdf.chapter_body("Edite 'data/raw/assets.csv' inserindo os tickers (ex: AAPL, WEGE3.SA).")
pdf.ln(2)

# PASSO 2
pdf.set_font('Arial', 'B', 10)
pdf.cell(0, 6, "Passo 2: Otimizar (Motor Matemático)", 0, 1)
pdf.set_font('Arial', '', 10)
pdf.chapter_body("Calcula a Fronteira Eficiente e gera na pasta 'data/processed':")
# Lista de Saídas
saidas_opt = [
    "carteira_recomendada.csv: Os pesos da carteira segura para o backtest.",
    "analise_portfolio_pro.xlsx: Relatório técnico contendo a aba 'Config' (metadados)."
]
pdf.chapter_list(saidas_opt)
pdf.code_block(["python markowitz_optimizer.py"])

# PASSO 3
pdf.set_font('Arial', 'B', 10)
pdf.cell(0, 6, "Passo 3: Validar (Backtest)", 0, 1)
pdf.set_font('Arial', '', 10)
pdf.chapter_body("Lê as configurações, realiza a simulação de 12 meses e gera:")
# Lista de Saídas
saidas_back = [
    "Relatorio_Final_Completo.xlsx: O Dashboard final com gráficos comparativos e tabelas."
]
pdf.chapter_list(saidas_back)
pdf.code_block(["python compare_strategies.py"])

# 5. RESULTADOS
pdf.chapter_title("Legenda do Gráfico Final")
resultados = [
    "Linha Cinza (Inicial): Sua carteira original passiva.",
    "Linha Azul (Sem Restrições): Potencial máximo teórico (Alocação 0-100%).",
    "Linha Verde (Com Restrições): Sugestão equilibrada (Respeita seus limites).",
    "Linha Laranja (Benchmark): Referência de mercado (Nasdaq-100)."
]
pdf.chapter_list(resultados)

# SALVAR
output_path = "README_Markowitz_V3_Final.pdf"
pdf.output(output_path)
print(f"PDF atualizado gerado em: {output_path}")