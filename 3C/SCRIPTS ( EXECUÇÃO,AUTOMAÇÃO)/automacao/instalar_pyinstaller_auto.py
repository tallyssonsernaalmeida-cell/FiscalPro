import subprocess
import sys
import os

def encontrar_python():
    """Encontra o caminho do Python no sistema"""
    caminhos = [
        r"C:\Users\IAF\AppData\Local\Microsoft\WindowsApps\python.exe",
        r"C:\Users\IAF\AppData\Local\Python\bin\python.exe",
        r"C:\Python312\python.exe",
        r"C:\Python311\python.exe",
        r"C:\Python310\python.exe",
    ]
    
    for caminho in caminhos:
        if os.path.exists(caminho):
            return caminho
    
    # Procura em outras pastas comuns
    pastas = [
        r"C:\Users\IAF\AppData\Local\Programs\Python",
        r"C:\Program Files\Python",
        r"C:\Program Files (x86)\Python",
    ]
    
    for pasta in pastas:
        if os.path.exists(pasta):
            for root, dirs, files in os.walk(pasta):
                for file in files:
                    if file == "python.exe":
                        return os.path.join(root, file)
    
    return None

def instalar_pyinstaller():
    python_path = encontrar_python()
    
    if not python_path:
        print("❌ Python não encontrado!")
        print("Por favor, instale o Python em: https://python.org")
        return False
    
    print(f"✅ Python encontrado: {python_path}")
    
    # Instala o pip se não existir
    print("\n📦 Atualizando pip...")
    subprocess.run([python_path, "-m", "ensurepip"], check=False)
    
    print("\n📦 Instalando PyInstaller...")
    result = subprocess.run([python_path, "-m", "pip", "install", "pyinstaller"], 
                           capture_output=True, text=True)
    
    if result.returncode == 0:
        print("✅ PyInstaller instalado com sucesso!")
        return True
    else:
        print(f"❌ Erro: {result.stderr}")
        return False

def gerar_executavel():
    python_path = encontrar_python()
    
    if not python_path:
        print("❌ Python não encontrado!")
        return
    
    print(f"\n⚙️ Gerando executável com: {python_path}")
    result = subprocess.run([
        python_path, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name", "ATHOS_Relatorio",
        "athos_monitor.py"
    ])
    
    if result.returncode == 0:
        print("\n✅ Executável gerado!")
        print("📁 Local: dist/ATHOS_Relatorio.exe")
    else:
        print("\n❌ Erro ao gerar executável")

if __name__ == "__main__":
    print("=" * 60)
    print("INSTALADOR E GERADOR AUTOMÁTICO")
    print("=" * 60)
    
    if instalar_pyinstaller():
        gerar_executavel()
    
    input("\nPressione Enter para sair...")