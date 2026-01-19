# --- 1. BIBLIOTECAS ---
import numpy as np
import pandas as pd
import yfinance as yf
import os
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from pypfopt import risk_models, expected_returns
from pypfopt.efficient_frontier import EfficientFrontier 
from pypfopt import plotting

# Configuração visual
plt.style.use('seaborn-v0_8-darkgrid')

# ==============================================================================
# CONFIGURAÇÃO DE RESTRIÇÕES (SEUS AJUSTES)
# ==============================================================================
MIN_ALOCACAO = 0.05  # X % (Obriga a ter pelo menos X % de cada ativo)
MAX_ALOCACAO = 0.30  # Y % (Nenhum ativo pode passar de Y % da carteira)
# ==============================================================================

# --- 2. DIRETÓRIOS E DADOS ---
print(f"--- Markowitz Pro: Otimização ({MIN_ALOCACAO:.0%} a {MAX_ALOCACAO:.0%}) ---")

base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) 
if not os.path.exists(os.path.join(base_dir, 'data')): base_dir = os.getcwd()

raw_dir = os.path.join(base_dir, 'data', 'raw')
processed_dir = os.path.join(base_dir, 'data', 'processed')
os.makedirs(processed_dir, exist_ok=True)

caminho_arquivo = os.path.join(raw_dir, 'assets.csv')

# Leitura e Tratamento
try:
    df_assets = pd.read_csv(caminho_arquivo, sep=None, engine='python')
    col_ativo = next((c for c in ['Ticker', 'Symbol', 'Código'] if c in df_assets.columns), None)
    if not col_ativo: raise ValueError("Coluna de Ticker não encontrada.")
    assets = df_assets[col_ativo].astype(str).str.strip().tolist()
    assets = list(dict.fromkeys(assets)) # Remove duplicatas
except Exception as e:
    print(f"Erro no CSV: {e}")
    exit()

# Download
anos_historico = 4
data_hoje = datetime.today()
data_inicio = data_hoje - timedelta(days=anos_historico * 365)
print(f"Ativos: {len(assets)}")

try:
    print("\n[1/4] Baixando Cotações...")
    dados = yf.download(assets, start=data_inicio, end=data_hoje, progress=False, auto_adjust=True)
    
    # Tratamento para MultiIndex (Correção para versões novas do yfinance/pandas)
    if 'Close' in dados.columns and isinstance(dados.columns, pd.MultiIndex):
        precos = dados['Close']
    else:
        precos = dados
        
    precos.dropna(axis=1, how='all', inplace=True)
    precos.dropna(inplace=True)
except Exception as e:
    print(f"Erro download: {e}")
    exit()

# --- 3. CÁLCULO DE RISCO E RETORNO ---
print("\n[2/4] Calculando Matrizes (Mu & Sigma)...")
mu = expected_returns.mean_historical_return(precos, frequency=252)
S = risk_models.CovarianceShrinkage(precos).ledoit_wolf()

# --- 4. OTIMIZAÇÃO DUPLA ---
print("\n[3/4] Otimizando Cenários...")
risk_free = 0.045

# CENÁRIO A: SEM RESTRIÇÃO (0% a 100%) - "Teórico Puro" (Estrela Azul)
try:
    ef_uncons = EfficientFrontier(mu, S, weight_bounds=(0, 1))
    w_uncons = ef_uncons.max_sharpe(risk_free_rate=risk_free)
    clean_uncons = ef_uncons.clean_weights()
    ret_un, vol_un, sha_un = ef_uncons.portfolio_performance(verbose=False, risk_free_rate=risk_free)
except Exception as e:
    print(f"Erro na otimização livre: {e}")
    clean_uncons = {}
    ret_un, vol_un, sha_un = 0, 0, 0

# CENÁRIO B: COM RESTRIÇÃO (MIN a MAX) - "Prático Seguro" (Estrela Dourada)
try:
    ef_cons = EfficientFrontier(mu, S, weight_bounds=(MIN_ALOCACAO, MAX_ALOCACAO))
    w_cons = ef_cons.max_sharpe(risk_free_rate=risk_free)
    clean_cons = ef_cons.clean_weights()
    ret_co, vol_co, sha_co = ef_cons.portfolio_performance(verbose=False, risk_free_rate=risk_free)
except Exception as e:
    print(f"\n[ERRO NA OTIMIZAÇÃO RESTRITA]: {e}")
    print("Dica: Verifique se Min * N_Ativos <= 100%. Se o Mínimo for muito alto, a conta não fecha.")
    clean_cons = {}
    ret_co, vol_co, sha_co = 0, 0, 0

# --- 5. VISUALIZAÇÃO DOS DADOS (TABELA COMPARATIVA) ---
col_name_restrito = f'Restrito ({MIN_ALOCACAO:.0%}-{MAX_ALOCACAO:.0%})'
col_name_livre = 'Livre (0%-100%)'

df_compare = pd.DataFrame({
    col_name_livre: pd.Series(clean_uncons),
    col_name_restrito: pd.Series(clean_cons)
})
df_compare.fillna(0, inplace=True)

# Filtra para mostrar apenas ativos com alocação
df_compare = df_compare.loc[(df_compare.abs() > 0.0001).any(axis=1)] 
df_compare = df_compare.sort_values(by=col_name_restrito, ascending=False)

print("\n" + "="*70)
print(f"{'COMPARAÇÃO DE ALOCAÇÃO':^70}")
print("="*70)
print(df_compare.applymap(lambda x: f"{x:.2%}"))
print("-" * 70)

# --- 6. EXPORTAÇÃO (COM METADADOS DE CONFIG) ---
print("\n[4/4] Salvando Arquivos...")

# A. Salva a carteira SEGURA (Constrained) para o backtest CSV
file_csv = os.path.join(processed_dir, "carteira_recomendada.csv")
if clean_cons:
    pd.Series(clean_cons)[pd.Series(clean_cons) > 0].to_csv(file_csv, header=False)
    print(f" > CSV (Carteira Restrita) salvo em: {file_csv}")

# B. Excel Completo (Com Aba de Configuração para o Analisador)
file_excel = os.path.join(processed_dir, "analise_portfolio_pro.xlsx")
try:
    with pd.ExcelWriter(file_excel, engine='openpyxl') as writer:
        # Aba 1: Comparativo
        df_compare.to_excel(writer, sheet_name='Comparativo Alocacao')
        
        # Aba 2: Preços
        precos.to_excel(writer, sheet_name='Precos Historicos')
        
        # Aba 3: Métricas
        pd.DataFrame({
            'Métrica': ['Retorno', 'Volatilidade', 'Sharpe'],
            'Livre': [ret_un, vol_un, sha_un],
            'Restrito': [ret_co, vol_co, sha_co]
        }).to_excel(writer, sheet_name='Metricas', index=False)
        
        # Aba 4: CONFIGURAÇÃO (O pulo do gato para o automator)
        pd.DataFrame({
            'Parametro': ['MIN_ALOCACAO', 'MAX_ALOCACAO'],
            'Valor': [MIN_ALOCACAO, MAX_ALOCACAO]
        }).to_excel(writer, sheet_name='Config', index=False)
        
    print(f" > Excel salvo com CONFIGURAÇÕES em: {file_excel}")
except Exception as e:
    print(f"Erro ao salvar Excel: {e}")

# --- 7. GRÁFICO DA FRONTEIRA ---
print(" > Gerando Gráfico...")

fig, ax = plt.subplots(figsize=(10, 6))

# Curva Teórica (Sempre 0 a 1)
ef_curve = EfficientFrontier(mu, S, weight_bounds=(0, 1))
plotting.plot_efficient_frontier(ef_curve, ax=ax, show_assets=False)

# Ativos Individuais
ax.scatter(np.sqrt(np.diag(S)), mu, s=30, color="black", label="Ativos Individuais", alpha=0.5)

# 1. ESTRELA AZUL (Sem Restrições)
ax.scatter(vol_un, ret_un, c='blue', s=300, marker='*', label=f'Carteira sem restrições (Sharpe: {sha_un:.2f})', zorder=10)

# 2. ESTRELA DOURADA (Com Restrições)
if ret_co > 0:
    ax.scatter(vol_co, ret_co, c='gold', s=300, marker='*', edgecolors='black', label=f'Carteira com restrições (Sharpe: {sha_co:.2f})', zorder=10)

ax.set_title(f'Fronteira Eficiente: Min {MIN_ALOCACAO:.0%} | Max {MAX_ALOCACAO:.0%}')
ax.set_xlabel('Volatilidade (Risco Anual)')
ax.set_ylabel('Retorno Esperado (Anual)')
ax.legend()

plt.tight_layout()
plt.show()