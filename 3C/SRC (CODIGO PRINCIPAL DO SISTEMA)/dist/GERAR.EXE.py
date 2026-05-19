"""
Gerador do Executável ATHOS Relatório - VERSÃO CORRIGIDA
Execute: python gerar_exe.py
"""

import os
import subprocess
import sys
import time

def instalar_pyinstaller():
    """Instala o PyInstaller corretamente"""
    print("📦 Instalando PyInstaller...")
    try:
        # Usa --user para evitar problemas de permissão
        subprocess.run([sys.executable, "-m", "pip", "install", "--user", "pyinstaller"], 
                       check=True, capture_output=True, text=True)
        print("✅ PyInstaller instalado com sucesso!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro na instalação: {e}")
        if e.stderr:
            print(f"Detalhes: {e.stderr}")
        return False

def verificar_pyinstaller():
    """Verifica se o PyInstaller está disponível"""
    try:
        # Tenta importar o PyInstaller
        subprocess.run([sys.executable, "-m", "PyInstaller", "--version"], 
                       capture_output=True, check=True)
        return True
    except:
        return False

def gerar_exe():
    print("=" * 60)
    print("GERADOR DO EXECUTÁVEL ATHOS")
    print("=" * 60)
    
    # Verifica se o script existe
    if not os.path.exists("athos_monitor.py"):
        print("\n❌ ERRO: Arquivo 'athos_monitor.py' não encontrado!")
        print("   Certifique-se que este script está na mesma pasta do athos_monitor.py")
        input("\nPressione Enter para sair...")
        return
    
    # Verifica/Instala PyInstaller
    if not verificar_pyinstaller():
        print("\n⚠️  PyInstaller não encontrado!")
        if not instalar_pyinstaller():
            print("\n❌ Falha ao instalar PyInstaller")
            print("\nTente instalar manualmente:")
            print("   pip install pyinstaller")
            input("\nPressione Enter para sair...")
            return
        print("\n🔄 Aguardando instalação...")
        time.sleep(2)
    
    # Remove arquivos antigos
    print("\n🗑️  Removendo arquivos antigos...")
    pastas_remover = ["build", "dist", "__pycache__"]
    for pasta in pastas_remover:
        if os.path.exists(pasta):
            import shutil
            try:
                shutil.rmtree(pasta)
                print(f"   ✓ Removido: {pasta}")
            except:
                print(f"   ⚠️ Não foi possível remover: {pasta}")
    
    # Remove .spec antigo
    spec_file = "ATHOS_Relatorio.spec"
    if os.path.exists(spec_file):
        try:
            os.remove(spec_file)
            print(f"   ✓ Removido: {spec_file}")
        except:
            pass
    
    # Comando para gerar o executável
    print("\n⚙️ Gerando executável (pode levar 2-3 minutos)...")
    print("   Aguarde...\n")
    
    # Usa comando mais simples e direto
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name", "ATHOS_Relatorio",
        "athos_monitor.py"
    ]
    
    try:
        # Executa e mostra a saída em tempo real
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, 
                                   text=True, bufsize=1, universal_newlines=True)
        
        for line in process.stdout:
            print(f"   {line.strip()}")
        
        process.wait()
        
        if process.returncode == 0:
            print("\n" + "=" * 60)
            print("✅ EXECUTÁVEL GERADO COM SUCESSO!")
            print("=" * 60)
            exe_path = os.path.abspath("dist/ATHOS_Relatorio.exe")
            print(f"📁 Localização: {exe_path}")
            
            if os.path.exists(exe_path):
                tamanho = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"📏 Tamanho: {tamanho:.2f} MB")
                print("\n💡 COMO USAR:")
                print("   1. Copie o arquivo 'ATHOS_Relatorio.exe' para a pasta desejada")
                print("   2. Execute com duplo clique (não vai abrir o CMD)")
                print("   3. Para criar atalho: Clique direito -> Enviar para -> Área de Trabalho")
                print("   4. Certifique-se que o arquivo 'downloads_realizados.json' está na mesma pasta")
            else:
                print("⚠️ Executável gerado mas não encontrado no local esperado")
        else:
            print("\n❌ ERRO AO GERAR EXECUTÁVEL!")
            print(f"Código de erro: {process.returncode}")
            
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        print("\nTente gerar manualmente com o comando:")
        print('   pyinstaller --onefile --windowed --name "ATHOS_Relatorio" athos_monitor.py')
    
    print("\n" + "=" * 60)
    input("Pressione Enter para sair...")

def gerar_sem_console():
    """Versão alternativa que mantém console para debug"""
    print("=" * 60)
    print("GERADOR COM CONSOLE (para debug)")
    print("=" * 60)
    
    if not os.path.exists("athos_monitor.py"):
        print("\n❌ ERRO: Arquivo 'athos_monitor.py' não encontrado!")
        input("Pressione Enter para sair...")
        return
    
    cmd = [
        "pyinstaller",
        "--onefile",
        "--name", "ATHOS_Relatorio_Debug",
        "athos_monitor.py"
    ]
    
    print("\n⚠️ Versão COM console (mostra erros se houver)")
    print("⚙️ Gerando...\n")
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ Gerado! Use ATHOS_Relatorio_Debug.exe para testar")
            print(f"📁 Local: {os.path.abspath('dist/ATHOS_Relatorio_Debug.exe')}")
        else:
            print("❌ Erro:")
            print(result.stderr)
    except Exception as e:
        print(f"❌ Erro: {e}")
    
    input("\nPressione Enter para sair...")

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("GERADOR DO EXECUTÁVEL ATHOS")
    print("=" * 60)
    print("\nOpções:")
    print("1 - Gerar executável SEM console (recomendado)")
    print("2 - Gerar executável COM console (para debug)")
    print("3 - Sair")
    
    opcao = input("\nEscolha uma opção (1/2/3): ").strip()
    
    if opcao == "1":
        gerar_exe()
    elif opcao == "2":
        gerar_sem_console()
    else:
        print("Saindo...")
