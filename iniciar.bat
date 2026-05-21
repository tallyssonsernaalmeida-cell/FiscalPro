@echo off
title FiscalPro - Servidor
color 0A
echo.
echo  ============================================
echo   FISCALPRO - Iniciando servidor...
echo  ============================================
echo.

cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERRO] Python nao encontrado! Instale em python.org
    pause
    exit
)

if not exist ".deps_ok" (
    echo  Instalando dependencias, aguarde...
    pip install -r requirements.txt --quiet
    echo. > .deps_ok
    echo  Dependencias instaladas!
    echo.
)

echo  Abrindo navegador...
timeout /t 2 /nobreak >nul
start http://127.0.0.1:5000

echo  Servidor rodando em: http://127.0.0.1:5000
echo  Admin: admin@fiscal.app / admin123
echo  Pressione CTRL+C para parar.
echo.

python app.py
pause
