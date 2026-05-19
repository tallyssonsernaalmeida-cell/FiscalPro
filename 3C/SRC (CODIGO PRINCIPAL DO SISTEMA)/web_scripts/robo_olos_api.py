# -*- coding: utf-8 -*-
"""
ROBO OLOS MAILING v6.0 - UPLOAD VIA API DIRETA
Muito mais rapido e confiavel!
"""

import pymssql
import csv
import os
import sys
import time
import json
import requests
from datetime import datetime
from pathlib import Path
from colorama import init, Fore, Style

init(autoreset=True)

# ============================================
# CONFIGURACOES
# ============================================

SQL_SERVER = "DESKTOP-74TBS68"
SQL_DATABASE = "IAF_BD"
SQL_USERNAME = "ControlDesk"
SQL_PASSWORD = "Control@2026"

OLOS_URL = "https://iaf.oloschannel.com.br"
OLOS_LOGIN_URL = f"{OLOS_URL}/Olos/login.aspx"
OLOS_IMPORT_URL = f"{OLOS_URL}/OPS_Custom/ImportExportWeb/ImportFiles"
OLOS_STATUS_URL = f"{OLOS_URL}/ImportExport/ImportStatus.aspx"
OLOS_USERNAME = "Tallysson.Almeida"
OLOS_PASSWORD = "Iaf@2026"

PASTA_RAIZ = Path(r"C:\Users\IAF\Desktop\Script\3C\SRC (CODIGO PRINCIPAL DO SISTEMA)")
PASTA_ARQUIVOS = PASTA_RAIZ / "arquivos_olos"
PASTA_ARQUIVOS.mkdir(parents=True, exist_ok=True)

# Sessao global para manter cookies
sessao = requests.Session()

# ============================================
# CONSULTA SQL PADRAO
# ============================================

CONSULTA_PADRAO = """WITH Ultimo_CPC AS (
    SELECT processo, fone, DATEDIFF(DAY, MAX(data_atual), GETDATE()) AS dias_ultimo_cpc
    FROM BaseAcionamentos WITH (NOLOCK)
    WHERE cpc = 1 AND ferramenta IN ('DISCADOR', 'DISCADOR MANUAL')
    GROUP BY processo, fone
),
BasePrincipal AS (
    SELECT a.processo, a.devedor, a.fone, a.hot_number, a.uid_devedor, b.saldo,
        LEFT(REPLACE(REPLACE(REPLACE(COALESCE(d.faixa_acoes, ''), '.', ''), '>', ''), '-', ''), 10) AS faixa_acoes,
        c.dias_ultimo_cpc, a.end_uf
    FROM BaseDevedor AS a WITH (NOLOCK)
    JOIN base_dia AS b WITH (NOLOCK) ON a.id = b.idDevedor
    LEFT JOIN Ultimo_CPC AS c ON a.processo = c.processo AND a.fone = c.fone
    LEFT JOIN base_consolidada AS d ON a.processo = d.processo
    WHERE b.status_processo NOT IN ('DEVOLUCAO', 'COBRANCA SUSPENSA', 'JURIDICO', 'ACORDO', 'AJUIZADO', 'QUITADO', 'FALECIDO', 'TRATATIVA COGNITO')
      AND LTRIM(RTRIM(UPPER(a.status_telefone))) NOT LIKE '%INATIVO%'
      AND b.carteira IN ('RESCISAO', 'GARANTIA')
      AND hot_number BETWEEN 0 AND 9999999999
),
FonesNumerados AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY processo ORDER BY hot_number DESC) AS rn
    FROM BasePrincipal
)
SELECT CONCAT(processo, ',', devedor, ',',
    MAX(CASE WHEN rn = 1 THEN fone END), ',',
    MAX(CASE WHEN rn = 2 THEN fone END), ',',
    MAX(CASE WHEN rn = 3 THEN fone END), ',',
    MAX(CASE WHEN rn = 4 THEN fone END), ',',
    MAX(CASE WHEN rn = 5 THEN fone END), ',',
    uid_devedor, ',', MAX(saldo), ',', MAX(faixa_acoes), ',',
    MAX(dias_ultimo_cpc), ',', MAX(end_uf)) AS mailing
FROM FonesNumerados
WHERE rn <= 5
GROUP BY processo, devedor, uid_devedor
ORDER BY MAX(dias_ultimo_cpc) DESC"""

# ============================================
# FUNCOES DE EXIBICAO
# ============================================

def limpar_tela():
    os.system('cls' if os.name == 'nt' else 'clear')

def cabecalho():
    limpar_tela()
    print(Fore.CYAN + Style.BRIGHT + """
╔══════════════════════════════════════════╗
║     ROBO OLOS MAILING v6.0              ║
║     Upload via API Direta               ║
╚══════════════════════════════════════════╝
""" + Style.RESET_ALL)

def print_ok(msg):
    print(Fore.GREEN + "  [OK] " + msg + Style.RESET_ALL)

def print_erro(msg):
    print(Fore.RED + "  [ERRO] " + msg + Style.RESET_ALL)

def print_info(msg):
    print(Fore.CYAN + "  [...] " + msg + Style.RESET_ALL)

def print_titulo(msg):
    print(Fore.YELLOW + Style.BRIGHT + "\n" + "="*50)
    print(Fore.YELLOW + Style.BRIGHT + "  " + msg)
    print(Fore.YELLOW + Style.BRIGHT + "="*50 + Style.RESET_ALL)

def print_destaque(msg):
    print(Fore.MAGENTA + Style.BRIGHT + "\n  >>> " + msg + " <<<" + Style.RESET_ALL)

# ============================================
# CONSULTA SQL
# ============================================

def executar_consulta(query):
    """Executa a consulta SQL e retorna os dados"""
    print_titulo("EXECUTANDO CONSULTA SQL")
    
    try:
        print_info("Conectando ao SQL Server...")
        
        conn = pymssql.connect(
            server=SQL_SERVER,
            database=SQL_DATABASE,
            user=SQL_USERNAME,
            password=SQL_PASSWORD
        )
        
        print_ok("Conexao estabelecida!")
        
        cursor = conn.cursor()
        
        print_info("Executando consulta...")
        inicio = datetime.now()
        
        cursor.execute(query)
        resultados = cursor.fetchall()
        
        tempo = (datetime.now() - inicio).total_seconds()
        
        lista_mailing = []
        for row in resultados:
            linha = str(row[0]).strip() if row[0] else ""
            if linha:
                lista_mailing.append(linha)
        
        cursor.close()
        conn.close()
        
        print_ok(f"Tempo de execucao: {tempo:.2f} segundos")
        print_destaque(f"TOTAL DE REGISTROS ENCONTRADOS: {len(lista_mailing)}")
        
        return lista_mailing
        
    except Exception as e:
        print_erro(f"Erro na consulta: {e}")
        return None

# ============================================
# GERAR CSV
# ============================================

def gerar_csv(dados, nome_arquivo=None):
    """Gera arquivo CSV"""
    print_titulo("GERANDO ARQUIVO CSV")
    
    try:
        if nome_arquivo is None:
            timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
            nome_arquivo = f"GERAL_RESCISAO_{timestamp}"
        
        if not nome_arquivo.endswith('.csv'):
            nome_arquivo += '.csv'
        
        caminho = PASTA_ARQUIVOS / nome_arquivo
        
        print_info(f"Salvando como: {nome_arquivo}")
        
        with open(caminho, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=',', quoting=csv.QUOTE_MINIMAL)
            writer.writerow(['processo,fone'])
            for linha in dados:
                writer.writerow([linha])
        
        tamanho_kb = caminho.stat().st_size / 1024
        
        print_ok(f"Arquivo gerado com sucesso!")
        print_ok(f"Tamanho: {tamanho_kb:.1f} KB")
        print_ok(f"Linhas: {len(dados)}")
        print(Fore.GREEN + f"\n  SALVO EM: {caminho}" + Style.RESET_ALL)
        
        return caminho
        
    except Exception as e:
        print_erro(f"Erro ao gerar CSV: {e}")
        return None

# ============================================
# LOGIN VIA API
# ============================================

def login_api():
    """Faz login e obtem cookies/token"""
    print_titulo("LOGIN NO OLOS (API)")
    
    try:
        # Primeiro, acessar pagina de login para pegar cookies
        print_info("Obtendo cookies iniciais...")
        resposta = sessao.get(OLOS_LOGIN_URL)
        print_ok(f"Status: {resposta.status_code}")
        
        # Extrair token antiforgery se existir
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(resposta.text, 'html.parser')
        
        # Procurar __RequestVerificationToken
        token = None
        for input_tag in soup.find_all('input'):
            if input_tag.get('name') == '__RequestVerificationToken':
                token = input_tag.get('value')
                break
        
        # Dados de login
        dados_login = {
            'UserTxt': OLOS_USERNAME,
            'Password': OLOS_PASSWORD,
            'BtnOK': 'OK'
        }
        
        if token:
            dados_login['__RequestVerificationToken'] = token
        
        # Headers
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        print_info("Enviando credenciais...")
        resposta_login = sessao.post(OLOS_LOGIN_URL, data=dados_login, headers=headers, allow_redirects=True)
        
        print_ok(f"Status: {resposta_login.status_code}")
        print_ok(f"URL final: {resposta_login.url}")
        
        # Verificar se logou
        if "login" not in resposta_login.url.lower():
            print_ok("Login realizado com sucesso!")
            
            # Agora acessar o Painel de Customizacoes
            print_info("Acessando Painel de Customizacoes...")
            painel_url = f"{OLOS_URL}/OPS_Custom"
            resposta_painel = sessao.get(painel_url)
            print_ok(f"Painel status: {resposta_painel.status_code}")
            
            return True
        else:
            print_erro("Falha no login!")
            return False
            
    except Exception as e:
        print_erro(f"Erro no login: {e}")
        return False

# ============================================
# UPLOAD VIA API
# ============================================

def upload_api(caminho_arquivo):
    """Faz upload do arquivo via API"""
    print_titulo("UPLOAD VIA API")
    
    try:
        # Ler arquivo
        print_info(f"Lendo arquivo: {caminho_arquivo.name}")
        
        with open(caminho_arquivo, 'r', encoding='utf-8') as f:
            conteudo = f.read()
        
        # Preparar dados do formulario
        # Nota: Os campos exatos dependem da API do Olos
        
        dados_upload = {
            'file': (caminho_arquivo.name, conteudo, 'text/csv'),
            'campaign': 'Manual_QuintoCred',
            'type': 'import'
        }
        
        # Headers para upload
        headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'X-Requested-With': 'XMLHttpRequest',
            'Origin': OLOS_URL,
            'Referer': f"{OLOS_URL}/OPS_Custom/ImportExportWeb/ImportFiles"
        }
        
        print_info(f"Enviando para: {OLOS_IMPORT_URL}")
        
        resposta = sessao.post(
            OLOS_IMPORT_URL,
            files={'file': (caminho_arquivo.name, open(caminho_arquivo, 'rb'), 'text/csv')},
            headers=headers
        )
        
        print_ok(f"Status: {resposta.status_code}")
        
        if resposta.status_code == 200:
            print_ok("Upload enviado com sucesso!")
            
            try:
                resultado = resposta.json()
                print_ok(f"Resposta: {json.dumps(resultado, indent=2)}")
            except:
                print_info(f"Resposta: {resposta.text[:200]}")
            
            return True
        else:
            print_erro(f"Erro no upload: {resposta.text[:200]}")
            return False
            
    except Exception as e:
        print_erro(f"Erro no upload: {e}")
        return False

# ============================================
# VERIFICAR STATUS
# ============================================

def verificar_status_api():
    """Verifica status da importacao via API"""
    print_titulo("VERIFICANDO STATUS")
    
    try:
        print_info("Consultando status...")
        
        resposta = sessao.get(OLOS_STATUS_URL)
        
        if resposta.status_code == 200:
            if "Imported" in resposta.text:
                print_destaque("ARQUIVO IMPORTADO COM SUCESSO!")
                return True
            else:
                print_info("Status: Ainda processando ou nao encontrado")
                return False
        else:
            print_erro(f"Erro ao consultar status: {resposta.status_code}")
            return False
            
    except Exception as e:
        print_erro(f"Erro: {e}")
        return False

# ============================================
# MENU PRINCIPAL
# ============================================

def menu_principal():
    """Menu principal"""
    
    logado = False
    
    while True:
        cabecalho()
        
        status_login = Fore.GREEN + "LOGADO" + Style.RESET_ALL if logado else Fore.RED + "NAO LOGADO" + Style.RESET_ALL
        
        print(Fore.WHITE + f"""
  Status: {status_login}
  
  OPCOES:
  ┌────────────────────────────────────────┐
  │ """ + Fore.GREEN + "1" + Fore.WHITE + """ - COLAR NOVA CONSULTA SQL          │
  │ """ + Fore.CYAN + "2" + Fore.WHITE + """ - USAR CONSULTA PADRAO              │
  │ """ + Fore.YELLOW + "3" + Fore.WHITE + """ - FAZER LOGIN NO OLOS (API)       │
  │ """ + Fore.MAGENTA + "4" + Fore.WHITE + """ - UPLOAD DE ARQUIVO EXISTENTE    │
  │ """ + Fore.BLUE + "5" + Fore.WHITE + """ - VERIFICAR STATUS DA IMPORTACAO   │
  │ """ + Fore.RED + "0" + Fore.WHITE + """ - SAIR                              │
  └────────────────────────────────────────┘
        """ + Style.RESET_ALL)
        
        opcao = input("  Escolha: ").strip()
        
        if opcao == '1':
            menu_nova_consulta(logado)
        elif opcao == '2':
            processar_consulta(CONSULTA_PADRAO, logado)
        elif opcao == '3':
            logado = login_api()
            print(Fore.YELLOW + "\nPressione ENTER para continuar..." + Style.RESET_ALL)
            input()
        elif opcao == '4':
            menu_upload_existente()
        elif opcao == '5':
            verificar_status_api()
            print(Fore.YELLOW + "\nPressione ENTER para continuar..." + Style.RESET_ALL)
            input()
        elif opcao == '0':
            print(Fore.GREEN + "\n  Ate logo!" + Style.RESET_ALL)
            break
        else:
            print_erro("Opcao invalida!")
            time.sleep(1)

def menu_nova_consulta(logado):
    """Menu para colar nova consulta"""
    cabecalho()
    
    print_titulo("COLAR NOVA CONSULTA SQL")
    print(Fore.CYAN + """
  Cole sua consulta SQL e digite 'FIM' para terminar
  Ou digite 'PADRAO' para usar a consulta padrao
    """ + Style.RESET_ALL)
    
    linhas = []
    while True:
        linha = input()
        if linha.strip().upper() == 'FIM':
            break
        if linha.strip().upper() == 'PADRAO':
            processar_consulta(CONSULTA_PADRAO, logado)
            return
        linhas.append(linha)
    
    if linhas:
        query = '\n'.join(linhas)
        print(Fore.YELLOW + f"\n  Deseja executar? (S/N): " + Style.RESET_ALL, end='')
        if input().strip().upper() == 'S':
            processar_consulta(query, logado)

def processar_consulta(query, logado):
    """Processa consulta e oferece opcoes"""
    dados = executar_consulta(query)
    
    if not dados:
        return
    
    print(Fore.CYAN + "\n  Como deseja nomear o arquivo?" + Style.RESET_ALL)
    nome = input("  > ").strip()
    
    if not nome:
        timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
        nome = f"GERAL_RESCISAO_{timestamp}"
    
    caminho = gerar_csv(dados, nome)
    
    if caminho and logado:
        print(Fore.YELLOW + "\n  Fazer upload via API? (S/N): " + Style.RESET_ALL, end='')
        if input().strip().upper() == 'S':
            upload_api(caminho)
    
    print(Fore.YELLOW + "\nPressione ENTER para continuar..." + Style.RESET_ALL)
    input()

def menu_upload_existente():
    """Upload de arquivo ja gerado"""
    cabecalho()
    print_titulo("UPLOAD DE ARQUIVO EXISTENTE")
    
    arquivos = list(PASTA_ARQUIVOS.glob("*.csv"))
    
    if not arquivos:
        print_erro("Nenhum arquivo encontrado!")
        return
    
    for i, arq in enumerate(arquivos, 1):
        tam = arq.stat().st_size / 1024
        print(f"  {Fore.GREEN}{i}{Style.RESET_ALL} - {arq.name} ({tam:.1f} KB)")
    
    try:
        op = int(input("\n  Escolha: ").strip())
        if 1 <= op <= len(arquivos):
            upload_api(arquivos[op - 1])
    except:
        pass
    
    print(Fore.YELLOW + "\nPressione ENTER para continuar..." + Style.RESET_ALL)
    input()

# ============================================
# INICIAR
# ============================================

if __name__ == "__main__":
    import traceback
    try:
        menu_principal()
    except KeyboardInterrupt:
        print(Fore.GREEN + "\n\n  Ate logo!" + Style.RESET_ALL)
    except Exception as e:
        print_erro(f"Erro: {e}")
        traceback.print_exc()
        input()
