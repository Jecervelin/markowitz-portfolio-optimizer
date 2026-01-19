# --- 1. BIBLIOTECAS ---
import numpy as np
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import os
from datetime import datetime, timedelta

# Estilo visual limpo e profissional
plt.style.use('seaborn-v0_8-darkgrid')

# --- 2. CONFIGURAÇÃO GERAL ---
print("--- ANÁLISE FINAL: Inicial vs Markowitz vs Benchmark ---")

# Definição Inteligente dos Caminhos
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) 
if not os.path.exists(os.path.join(BASE_DIR, 'data')): BASE_DIR = os.getcwd()

RAW_DIR = os.path.join(BASE_DIR, 'data', 'raw')
PROC_DIR = os.path.join(BASE_DIR, 'data', 'processed')

FILE_MANUAL = os.path.join(RAW_DIR, 'assets.csv')
FILE_ROBO_CSV = os.path.join(PROC_DIR, 'carteira_recomendada.csv')
FILE_ROBO_XLSX = os.path.join(PROC_DIR, 'analise_portfolio_pro.xlsx')
FILE_SAIDA_XLSX = os.path.join(PROC_DIR, 'Relatorio_Final_Completo.xlsx')

CAPITAL = 10000.00
BENCHMARK_TICKER = 'QQQ'
BENCHMARK_NOME = 'Benchmark (NASDAQ)'
JANELA_DIAS = 365

# --- 3. DETECÇÃO AUTOMÁTICA DAS REGRAS ---
print("\n[1/7] Lendo configurações do Otimizador...")

# Valor padrão caso não encontre o arquivo
INFO_RESTRICOES = "Restrições Desconhecidas"

try:
    if os.path.exists(FILE_ROBO_XLSX):
        # Tenta ler a aba 'Config' onde salvamos MIN e MAX
        try:
            df_config = pd.read_excel(FILE_ROBO_XLSX, sheet_name='Config')
            
            # Busca os valores nas linhas correspondentes
            min_val = df_config.loc[df_config['Parametro'] == 'MIN_ALOCACAO', 'Valor'].values[0]
            max_val = df_config.loc[df_config['Parametro'] == 'MAX_ALOCACAO', 'Valor'].values[0]
            
            INFO_RESTRICOES = f"Min {min_val:.1%} | Max {max_val:.0%}"
            print(f" > Sucesso: Regras detectadas ({INFO_RESTRICOES})")
            
        except ValueError:
            print(" [AVISO] Aba 'Config' não encontrada no Excel. Usando nome genérico.")
            INFO_RESTRICOES = "Com Restrições"
    else:
        print(" [ERRO] Arquivo do otimizador não encontrado.")
except Exception as e:
    print(f" [ERRO] Falha ao ler config: {e}")
    INFO_RESTRICOES = "Com Restrições"

# Define os nomes finais para os Gráficos
NOME_LIVRE = "MARKOWITZ (SEM RESTRIÇÕES)"
NOME_RESTRITO = f"MARKOWITZ ({INFO_RESTRICOES})"


# --- 4. CARREGAMENTO DAS CARTEIRAS ---
print("\n[2/7] Carregando Portfólios...")

cart_manual = {}
cart_restrita = {}
cart_sem_restricao = {}

# A. CARTEIRA INICIAL (Manual)
try:
    df_orig = pd.read_csv(FILE_MANUAL, sep=None, engine='python')
    # Detecta colunas
    col_t = next((c for c in ['Ticker', 'Symbol', 'Código'] if c in df_orig.columns), None)
    col_p = next((c for c in ['Weight (%)', 'Weight', 'Peso'] if c in df_orig.columns), 'Weight (%)')
    
    if col_t:
        df_orig[col_t] = df_orig[col_t].astype(str).str.strip()
        cart_manual = dict(zip(df_orig[col_t], df_orig[col_p] / 100))
        print(f" > Inicial: {len(cart_manual)} ativos.")
except: pass

# B. RESTRITA (CSV - É a carteira oficial para backtest)
try:
    if os.path.exists(FILE_ROBO_CSV):
        df_opt = pd.read_csv(FILE_ROBO_CSV, header=None, index_col=0)
        cart_restrita = pd.to_numeric(df_opt.iloc[:, 0], errors='coerce').dropna().to_dict()
        cart_restrita = {str(k).strip(): v for k, v in cart_restrita.items() if v > 0.001}
        print(f" > Restrita: {len(cart_restrita)} ativos.")
except: pass

# C. LIVRE (EXCEL - Aba Comparativo)
try:
    if os.path.exists(FILE_ROBO_XLSX):
        df_comp = pd.read_excel(FILE_ROBO_XLSX, sheet_name='Comparativo Alocacao', index_col=0)
        col_livre = next((c for c in df_comp.columns if 'Livre' in c or 'Sem Limites' in c), None)
        if col_livre:
            cart_sem_restricao = df_comp[col_livre].dropna().to_dict()
            cart_sem_restricao = {str(k).strip(): v for k, v in cart_sem_restricao.items() if v > 0.001}
            print(f" > Livre:    {len(cart_sem_restricao)} ativos.")
            
        # Fallback: Se o CSV falhou, tenta pegar a Restrita do Excel também
        if not cart_restrita:
            col_rest = next((c for c in df_comp.columns if 'Restrito' in c), None)
            if col_rest:
                cart_restrita = df_comp[col_rest].dropna().to_dict()
                print(f" > Restrita (Recuperada do Excel): {len(cart_restrita)} ativos.")
except: pass

if not cart_restrita and not cart_sem_restricao:
    print("ERRO CRÍTICO: Nenhuma carteira otimizada encontrada. Rode o 'markowitz_optimizer_pro.py' primeiro.")
    exit()


# --- 5. DADOS DE MERCADO ---
print("\n[3/7] Baixando Cotações...")
todos_ativos = list(set(
    list(cart_manual.keys()) + list(cart_restrita.keys()) + list(cart_sem_restricao.keys()) + [BENCHMARK_TICKER]
))

end_date = datetime.today()
start_date = end_date - timedelta(days=JANELA_DIAS)

try:
    # auto_adjust=True para dividendos/splits
    dados_raw = yf.download(todos_ativos, start=start_date, end=end_date, progress=False, auto_adjust=True)
    
    # Tratamento MultiIndex (Correção Yfinance)
    if 'Close' in dados_raw.columns and isinstance(dados_raw.columns, pd.MultiIndex):
        dados = dados_raw['Close']
    else:
        dados = dados_raw
    dados.dropna(inplace=True)
    print(f" > Dados baixados: {dados.shape[0]} dias de pregão.")
except Exception as e:
    print(f"Erro download: {e}")
    exit()


# --- 6. SIMULAÇÃO ---
print("\n[4/7] Simulando Performance...")

def simular(pesos, df, capital):
    if not pesos: return pd.Series(0, index=df.index)
    
    # Filtra apenas ativos que conseguimos baixar
    ativos = {k: v for k, v in pesos.items() if k in df.columns}
    
    # Renormaliza para 100%
    soma = sum(ativos.values())
    if soma == 0: return pd.Series(0, index=df.index)
    ativos = {k: v / soma for k, v in ativos.items()}
    
    # Cálculo
    ret_acum = df[list(ativos.keys())] / df[list(ativos.keys())].iloc[0]
    saldo = pd.Series(0.0, index=df.index)
    for at, pe in ativos.items():
        saldo += ret_acum[at] * pe * capital
    return saldo

saldo_manual = simular(cart_manual, dados, CAPITAL)
saldo_restrita = simular(cart_restrita, dados, CAPITAL)
saldo_sem_restricao = simular(cart_sem_restricao, dados, CAPITAL)
saldo_bench = (dados[BENCHMARK_TICKER] / dados[BENCHMARK_TICKER].iloc[0]) * CAPITAL


# --- 7. MÉTRICAS E TABELA ---
print("\n[5/7] Calculando Indicadores...")

def get_metrics(serie):
    if serie.empty or serie.iloc[0] == 0: return 0, 0, 0
    ret = (serie.iloc[-1] / serie.iloc[0]) - 1
    vol = serie.pct_change().std() * (252**0.5)
    sha = ret / vol if vol > 0 else 0
    return ret, vol, sha

metrics = []
# Adiciona na lista apenas se a carteira existir
if not saldo_manual.empty and saldo_manual.iloc[-1] > 0:
    metrics.append(('CARTEIRA INICIAL', saldo_manual))

if not saldo_sem_restricao.empty and saldo_sem_restricao.iloc[-1] > 0:
    metrics.append((NOME_LIVRE, saldo_sem_restricao))

if not saldo_restrita.empty and saldo_restrita.iloc[-1] > 0:
    metrics.append((NOME_RESTRITO, saldo_restrita))

metrics.append((BENCHMARK_NOME, saldo_bench))

# Exibe Tabela no Console
resumo_dados = {'Estratégia': [], 'Retorno': [], 'Risco (Vol)': [], 'Sharpe': [], 'Saldo Final': []}

print("-" * 110)
print(f"{'ESTRATÉGIA':<40} | {'RETORNO':<10} | {'RISCO':<10} | {'SHARPE':<8} | {'SALDO':<12}")
print("-" * 110)

for nome, serie in metrics:
    r, v, s = get_metrics(serie)
    resumo_dados['Estratégia'].append(nome)
    resumo_dados['Retorno'].append(r)
    resumo_dados['Risco (Vol)'].append(v)
    resumo_dados['Sharpe'].append(s)
    resumo_dados['Saldo Final'].append(serie.iloc[-1])
    print(f"{nome:<40} | {r:<10.2%} | {v:<10.2%} | {s:<8.2f} | R$ {serie.iloc[-1]:,.2f}")
print("-" * 110)


# --- 8. EXPORTAÇÃO EXCEL ---
print("\n[6/7] Gerando Relatório Excel...")

df_resumo = pd.DataFrame(resumo_dados)
dict_hist = {nome: serie for nome, serie in metrics}
df_hist = pd.DataFrame(dict_hist)
df_hist.index = pd.to_datetime(df_hist.index).date

try:
    with pd.ExcelWriter(FILE_SAIDA_XLSX, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formatos
        fmt_pct = workbook.add_format({'num_format': '0.00%'})
        fmt_money = workbook.add_format({'num_format': 'R$ #,##0.00'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        fmt_center = workbook.add_format({'align': 'center'})
        
        # --- ABA 1: DASHBOARD ---
        sheet_dash = workbook.add_worksheet('Dashboard')
        df_resumo.to_excel(writer, sheet_name='Dashboard', startrow=1, startcol=1, index=False)
        
        # Ajuste de largura das colunas
        sheet_dash.set_column('B:B', 40) # Bem largo para caber "Markowitz (Min X% | Max Y%)"
        sheet_dash.set_column('C:D', 15, fmt_pct)
        sheet_dash.set_column('E:E', 12, fmt_center)
        sheet_dash.set_column('F:F', 18, fmt_money)
        
        # Aplica cabeçalho
        for idx, val in enumerate(df_resumo.columns):
            sheet_dash.write(1, idx+1, val, fmt_header)
            
        # Gráfico de Barras
        chart_bar = workbook.add_chart({'type': 'column'})
        chart_bar.add_series({
            'name': 'Retorno Total',
            'categories': ['Dashboard', 2, 1, 2 + len(df_resumo)-1, 1], # Nomes das estratégias
            'values':     ['Dashboard', 2, 2, 2 + len(df_resumo)-1, 2], # Valores de Retorno
            'data_labels': {'value': True, 'num_format': '0.0%'},
            'points': [
                {'fill': {'color': 'gray'}},    # Manual
                {'fill': {'color': '#00B0F0'}}, # Livre (Azul)
                {'fill': {'color': '#00B050'}}, # Restrito (Verde)
                {'fill': {'color': '#FFC000'}}  # Benchmark (Laranja)
            ]
        })
        chart_bar.set_title({'name': 'Rentabilidade Acumulada'})
        chart_bar.set_legend({'position': 'none'})
        sheet_dash.insert_chart('H2', chart_bar)
        
        # --- ABA 2: HISTÓRICO ---
        df_hist.to_excel(writer, sheet_name='Dados_Historicos')
        sheet_data = writer.sheets['Dados_Historicos']
        sheet_data.set_column('A:A', 12, workbook.add_format({'num_format': 'dd/mm/yyyy'}))
        sheet_data.set_column('B:Z', 20, fmt_money)
        
        # Gráfico de Linha
        chart_line = workbook.add_chart({'type': 'line'})
        max_row = len(df_hist) + 1
        
        # Mapa de cores consistente
        color_map = {
            'CARTEIRA INICIAL': 'gray',
            NOME_LIVRE: '#00B0F0',
            NOME_RESTRITO: '#00B050',
            BENCHMARK_NOME: '#FFC000'
        }
        
        for i, col_name in enumerate(df_hist.columns):
            color = color_map.get(col_name, 'black') # Preto se não achar cor
            chart_line.add_series({
                'name':       ['Dados_Historicos', 0, i+1],
                'categories': ['Dados_Historicos', 1, 0, max_row, 0],
                'values':     ['Dados_Historicos', 1, i+1, max_row, i+1],
                'line':       {'color': color, 'width': 2.25}
            })
            
        chart_line.set_title({'name': f'Evolução Patrimonial ({INFO_RESTRICOES})'})
        chart_line.set_size({'width': 900, 'height': 500})
        chart_line.set_y_axis({'name': 'Patrimônio (R$)', 'major_gridlines': {'visible': True}})
        sheet_dash.insert_chart('B10', chart_line)
        
    print(f" > Excel salvo: {FILE_SAIDA_XLSX}")

except Exception as e:
    print(f" [ERRO] Falha ao salvar Excel: {e}")
    print("Dica: Feche o arquivo Excel se ele estiver aberto.")


# --- 9. PLOTAGEM RÁPIDA (PREVIEW) ---
print("\n[7/7] Exibindo Gráfico...")
plt.figure(figsize=(12, 7))

# Plota na ordem lógica visual (Benchmark ao fundo, depois as linhas de destaque)
plt.plot(saldo_bench.index, saldo_bench, color="#FF4800", alpha=0.6, linestyle=':', label=BENCHMARK_NOME)

if not saldo_manual.empty and saldo_manual.iloc[-1] > 0:
    plt.plot(saldo_manual.index, saldo_manual, color='magenta', linestyle='--', label='Carteira Inicial')

if not saldo_sem_restricao.empty and saldo_sem_restricao.iloc[-1] > 0:
    plt.plot(saldo_sem_restricao.index, saldo_sem_restricao, color='#00B0F0', linewidth=1.5, alpha=0.8, label=NOME_LIVRE)

if not saldo_restrita.empty and saldo_restrita.iloc[-1] > 0:
    plt.plot(saldo_restrita.index, saldo_restrita, color='#00B050', linewidth=2.5, label=NOME_RESTRITO)

plt.title(f'Performance: Com vs Sem Restrições ({INFO_RESTRICOES})', fontsize=14)
plt.ylabel('Patrimônio (R$)')
plt.legend()
plt.tight_layout()
plt.show()