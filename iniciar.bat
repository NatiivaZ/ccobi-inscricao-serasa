@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   Pagamento e Inscrição SERASA
echo   Automação SIFAMA - ANTT
echo ========================================
echo.
echo Verificando Python...
python --version
if errorlevel 1 (
    py --version
    if errorlevel 1 (
        echo ERRO: Python não encontrado! Instale o Python primeiro.
        pause
        exit /b 1
    )
    set PY=py
) else (
    set PY=python
)
echo.
echo Instalando/Atualizando dependências...
%PY% -m pip install --upgrade pip -q
%PY% -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo ERRO: Falha ao instalar dependências!
    pause
    exit /b 1
)
echo.
echo ========================================
echo Iniciando sistema...
echo ========================================
echo.
%PY% automacao_sifama_integrada.py
if errorlevel 1 (
    echo.
    echo Ocorreu um erro. Veja a mensagem acima.
    pause
    exit /b 1
)
pause
