@echo off
chcp 1252 >nul
title Instalar ATHOS Download - Agendador Windows

cd /d "C:\Users\IAF\Desktop\Script\3C"

echo ===================================
echo Instalando ATHOS Download
echo com Agendador do Windows
echo ===================================
echo.

REM 1. Criar EXE
echo [1/4] Criando executavel...
python -m PyInstaller --onefile --noconsole --name "ATHOS_Download" --clean --noconfirm athos_servico_final.py >nul 2>&1

if exist "dist\ATHOS_Download.exe" (
    echo OK: Executavel criado!
) else (
    echo ERRO: Falha ao criar EXE
    pause
    exit /b 1
)

REM 2. Criar pasta permanente
echo [2/4] Criando pasta permanente...
set "DESTINO=%USERPROFILE%\ATHOS_Download"
if not exist "%DESTINO%" mkdir "%DESTINO%"
copy /y "dist\ATHOS_Download.exe" "%DESTINO%\ATHOS_Download.exe" >nul
echo OK: %DESTINO%

REM 3. Remover atalho antigo da inicializacao (se existir)
echo [3/4] Removendo configuracoes antigas...
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
if exist "%STARTUP%\ATHOS_Download.lnk" del /q "%STARTUP%\ATHOS_Download.lnk"

REM 4. Criar tarefa no Agendador do Windows
echo [4/4] Criando tarefa agendada...

schtasks /create /tn "ATHOS Download" /tr "\"%DESTINO%\ATHOS_Download.exe\"" /sc daily /st 08:00 /f /rl highest

if %errorlevel% equ 0 (
    echo.
    echo ===================================
    echo INSTALACAO CONCLUIDA!
    echo ===================================
    echo.
    echo Configuracao:
    echo - Nome: ATHOS Download
    echo - Horario: 08:00 todo dia
    echo - Programa: %DESTINO%\ATHOS_Download.exe
    echo - Modo: Agendador do Windows
    echo - ACORDA PC: SIM (mesmo hibernado)
    echo.
    echo O Windows vai:
    echo 1. Acordar o PC as 08:00 (se hibernado)
    echo 2. Executar o download automaticamente
    echo 3. Abrir a pasta com os audios
    echo.
    echo Nao gera arquivos de log.
    echo.
) else (
    echo.
    echo ERRO ao criar tarefa. Tentando com PowerShell...
    
    powershell -Command "$action = New-ScheduledTaskAction -Execute '%DESTINO%\ATHOS_Download.exe'; $trigger = New-ScheduledTaskTrigger -Daily -At '08:00'; $settings = New-ScheduledTaskSettingsSet -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries; Register-ScheduledTask -TaskName 'ATHOS Download' -Action $action -Trigger $trigger -Settings $settings -Force"
    
    if %errorlevel% equ 0 (
        echo OK: Tarefa criada via PowerShell!
    ) else (
        echo ERRO: Nao foi possivel criar a tarefa.
    )
)

echo.
echo Para testar agora:
echo schtasks /run /tn "ATHOS Download"
echo.

set /p teste="Deseja testar agora? (S/N): "
if /i "%teste%"=="S" (
    echo Executando teste...
    schtasks /run /tn "ATHOS Download"
    echo Teste iniciado! Verifique a pasta.
)

pause