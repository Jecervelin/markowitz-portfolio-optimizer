import yfinance as yf

# O ticker que você quer testar
ticker_symbol = "XYZ"

print(f"--- Testando conexão com: {ticker_symbol} ---")

try:
    # 1. Tenta baixar o histórico recente (5 dias)
    # progress=False esconde a barra de carregamento para limpar a tela
    dados = yf.download(ticker_symbol, period="5d", progress=False)

    # 2. Verifica se o DataFrame voltou vazio
    if dados.empty:
        print(f"❌ ERRO: O ticker '{ticker_symbol}' não retornou dados.")
        print("Possíveis causas: Ticker incorreto, ativo deslistado ou bloqueio de IP.")
    else:
        print(f"✅ SUCESSO: Dados encontrados para '{ticker_symbol}'!")
        print("\nÚltimos valores baixados:")
        print(dados.tail())
        
        # Opcional: Tentar pegar informações extras (Setor, Nome completo)
        try:
            info = yf.Ticker(ticker_symbol).info
            nome = info.get('shortName') or info.get('longName')
            preco = info.get('currentPrice')
            print(f"\nNome: {nome}")
            print(f"Preço Atual (aprox): ${preco}")
        except:
            print("\n(Info de metadados não disponível, mas os preços históricos funcionam)")

except Exception as e:
    print(f"❌ ERRO CRÍTICO: {e}")