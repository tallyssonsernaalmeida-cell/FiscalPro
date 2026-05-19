@echo off
chcp 65001 >nul
title Instalador Robô Olos Mailing

echo ╔══════════════════════════════════════════╗
echo ║  INSTALADOR DO ROBÔ OLOS MAILING        ║
echo ╚══════════════════════════════════════════╝
echo.

cd /d "C:\Users\IAF\Desktop\Script\3C\SRC (CODIGO PRINCIPAL DO SISTEMA)"

echo 📦 Instalando dependências Python...
pip install -r requirements.txt

echo.
echo ⚙️  Configurando inicialização automática...

:: Criar tarefa agendada para iniciar com Windows
schtasks /create /tn "RoboOlosMailing" /tr "python \"C:\Users\IAF\Desktop\Script\3C\SRC (CODIGO PRINCIPAL DO SISTEMA)\automacao_olos_mailing.py\" --agendar" /sc onstart /ru SYSTEM /rl highest /f

echo.
echo ✅ Instalação concluída!
echo ⏰ O robô iniciará automaticamente com o Windows
echo 📋 Verifique o arquivo config.ini para configurar credenciais
echo.
pause