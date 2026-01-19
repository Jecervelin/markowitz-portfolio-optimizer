@echo off
setlocal
cd /d "%~dp0"

echo ========================================================
echo      SETUP DE AMBIENTE (VERSAO CORRIGIDA)
echo ========================================================
echo.

:: 1. Verificacao Python
echo [1/6] Verificando Python...
python --version
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado no PATH!
    pause
    exit /b
)

:: 2. Criar Ambiente Virtual (.venv)
:: OBS: Removi parenteses dos textos para evitar erro de sintaxe
if not exist ".venv" (
    echo [2/6] Criando ambiente virtual novo...
    python -m venv .venv
) else (
    echo [INFO] Pasta .venv ja existe. Pulando criacao.
)

:: 3. Atualizar PIP
echo [3/6] Atualizando o PIP...
.\.venv\Scripts\python -m pip install --upgrade pip

:: 4. Instalar Bibliotecas
echo [4/6] Instalando Bibliotecas Completas...
echo Aguarde...
.\.venv\Scripts\pip install pandas numpy matplotlib seaborn yfinance scipy statsmodels scikit-learn jupyterlab openpyxl xlsxwriter PyPortfolioOpt --no-cache-dir

:: 5. Criar Pastas
echo [5/6] Criando estrutura de pastas...
if not exist "src" mkdir src
if not exist "notebooks" mkdir notebooks
if not exist "data\raw" mkdir "data\raw"
if not exist "data\processed" mkdir "data\processed"
if not exist "models" mkdir models

if not exist "src\__init__.py" type nul > "src\__init__.py"

:: 6. Requirements
echo [6/6] Gerando requirements.txt...
.\.venv\Scripts\pip freeze > requirements.txt

echo.
echo ========================================================
echo      SUCESSO! AMBIENTE PRONTO.
echo ========================================================
pause