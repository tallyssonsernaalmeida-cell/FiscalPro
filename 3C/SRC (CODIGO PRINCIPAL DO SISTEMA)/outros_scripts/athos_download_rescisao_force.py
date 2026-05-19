"""
ATHOS AUTO DOWNLOAD - RESCISÃO (FORÇADO)
Baixa data específica: 04/05/2026
Campanhas: 91-360 e 360+
"""

import os
import re
import json
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
from bs4 import BeautifulSoup

# ================================================================
# CONFIGURAÇÕES
# ================================================================
LOGIN_URL = "https://iaf.oloschannel.com.br/Olos/Login.aspx"
SEARCH_URL = "https://iaf.oloschannel.com.br/RecordingRetrieve/RecordingRetrieveList.aspx"
DOWNLOAD_BASE = "https://iaf.oloschannel.com.br/RecordingRetrieve/PlayMp3.aspx"

# Credenciais
USUARIO = "Tallysson.Almeida"
SENHA = "Iaf@2026"

# Filtros
HORA_INI = "07:00:00"
HORA_FIM = "21:00:00"
DURACAO_MIN = 60

# DATA ESPECÍFICA PARA BAIXAR
DATA_ALVO = "04/05/2026"

# ================================================================
# CONFIGURAÇÕES DAS CAMPANHAS DE RESCISÃO
# ================================================================
CAMPANHAS_RESCISAO = {
    "91-360": {
        "infos": ["101081", "491120", "5121180"],
        "pasta_base": r"C:\Users\IAF\OneDrive\Ligações Monitoria\2026\Rescisão\91-360"
    },
    "360+": {
        "infos": ["6181360", "7361540", "8541720", "97211080"],
        "pasta_base": r"C:\Users\IAF\OneDrive\Ligações Monitoria\2026\Rescisão\360+"
    }
}

MESES_PASTA = {
    "01": "JANEIRO 01", "02": "FEVEREIRO 02", "03": "MARÇO 03",
    "04": "ABRIL 04", "05": "MAIO 05", "06": "JUNHO 06",
    "07": "JULHO 07", "08": "AGOSTO 08", "09": "SETEMBRO 09",
    "10": "OUTUBRO 10", "11": "NOVEMBRO 11", "12": "DEZEMBRO 12",
}

FERIADOS = {
    "01/01/2026": "Ano Novo", "03/04/2026": "Sexta-feira Santa",
    "21/04/2026": "Tiradentes", "01/05/2026": "Dia do Trabalho",
    "07/09/2026": "Independência", "12/10/2026": "Nossa Sra. Aparecida",
    "02/11/2026": "Finados", "25/12/2026": "Natal",
}

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ================================================================
# FUNÇÕES AUXILIARES
# ================================================================
def log(mensagem):
    print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] {mensagem}")

def formatar_pasta_dia(dia):
    """Formata o nome da pasta do dia:
    Dias 1-9: 2 dígitos (04, 05, 06)
    Dias 10-31: 3 dígitos (010, 011, 012)"""
    if 1 <= dia <= 9:
        return f"{dia:02d}"
    else:
        return f"{dia:03d}"

def fazer_login():
    """Faz login na plataforma OLOS"""
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'pt-BR,pt;q=0.9',
    })
    
    try:
        response = session.get(LOGIN_URL, timeout=30)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        viewstate = soup.find('input', {'id': '__VIEWSTATE'})
        viewstategenerator = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})
        eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})
        
        payload = {
            '__VIEWSTATE': viewstate['value'] if viewstate else '',
            '__VIEWSTATEGENERATOR': viewstategenerator['value'] if viewstategenerator else '',
            '__EVENTVALIDATION': eventvalidation['value'] if eventvalidation else '',
            'UserTxt': USUARIO,
            'Password': SENHA,
            'BtnOK': 'OK',
        }
        
        response = session.post(LOGIN_URL, data=payload, timeout=30)
        
        if 'RecordingRetrieveList' in response.url:
            log("✅ Login realizado com sucesso!")
            return session
        else:
            log("❌ Falha no login!")
            return None
    except Exception as e:
        log(f"❌ Erro no login: {e}")
        return None

def buscar_chamadas_por_info(session, info_adicional, data_str):
    """Busca chamadas filtrando por Info. Adicionais"""
    
    response = session.get(SEARCH_URL, timeout=30)
    soup = BeautifulSoup(response.text, 'html.parser')
    
    viewstate = soup.find('input', {'id': '__VIEWSTATE'})
    viewstategenerator = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})
    eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})
    
    payload = {
        '__VIEWSTATE': viewstate['value'] if viewstate else '',
        '__VIEWSTATEGENERATOR': viewstategenerator['value'] if viewstategenerator else '',
        '__EVENTVALIDATION': eventvalidation['value'] if eventvalidation else '',
        'ctl00$ToolkitScriptManager1': '',
        'ctl00$PageContent$TabContainer1$TabPanelDateTime$StartDate': data_str,
        'ctl00$PageContent$TabContainer1$TabPanelDateTime$EndDate': data_str,
        'ctl00$PageContent$TabContainer1$TabPanelDateTime$StartHour': HORA_INI,
        'ctl00$PageContent$TabContainer1$TabPanelDateTime$EndHour': HORA_FIM,
        'ctl00$PageContent$TabContainer1$TabPanelCalls$DurationStart': str(DURACAO_MIN),
        'ctl00$PageContent$TabContainer1$TabPanelCalls$DurationStartComparision': '>=',
        'ctl00$PageContent$TabContainer1$TabPanelMailing$AdditionalInformation': info_adicional,
        'ctl00$PageContent$rdTypeRecord': 'rdTypeRecordCall',
        'ctl00$PageContent$Button1X': 'Pesquisa',
        '__ASYNCPOST': 'true',
        'ctl00_PageContent_UpdatePanel1': 'ctl00_PageContent_UpdatePanel1',
    }
    
    response = session.post(SEARCH_URL, data=payload, timeout=30)
    
    chamadas = []
    soup = BeautifulSoup(response.text, 'html.parser')
    tabela = soup.find('table', {'id': 'ctl00_PageContent_ListRecording'})
    
    if tabela:
        linhas = tabela.find_all('tr', class_=re.compile(r'table-result'))
        
        for linha in linhas:
            celulas = linha.find_all('td')
            if len(celulas) >= 9:
                call_id = None
                file_name = None
                data_hora = None
                
                for i, celula in enumerate(celulas):
                    texto = celula.get_text(strip=True)
                    if i == 0:
                        file_name = texto
                    elif i == 7:
                        call_id = texto
                    elif i == 2:
                        data_hora = texto
                
                if call_id and file_name:
                    chamadas.append({
                        'CallIdMaster': call_id,
                        'FileName': file_name,
                        'DateTime': data_hora,
                        'InfoAdicional': info_adicional,
                    })
    
    return chamadas

def baixar_audio(session, chamada, pasta_destino):
    """Baixa um áudio específico para a pasta determinada"""
    
    call_id = chamada.get('CallIdMaster')
    file_name = chamada.get('FileName')
    
    if not call_id or not file_name:
        return "ERRO"
    
    os.makedirs(pasta_destino, exist_ok=True)
    
    info = chamada.get('InfoAdicional', '')
    filename = f"{info}_{call_id}_{file_name}"
    filepath = os.path.join(pasta_destino, re.sub(r'[\\/*?:"<>|]', "_", filename))
    
    if os.path.exists(filepath):
        return "EXISTE"
    
    url = f"{DOWNLOAD_BASE}?FileName={file_name}&CallIdMaster={call_id}"
    
    try:
        response = session.get(url, timeout=60, stream=True)
        if response.status_code == 200:
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            return "OK"
        else:
            return f"ERRO_{response.status_code}"
    except Exception as e:
        return f"ERRO_{str(e)[:50]}"

def processar_campanha(campanha_nome, config, data_ligacao, session):
    """Processa uma campanha inteira (todos os Info. Adicionais)"""
    
    data_str = data_ligacao.strftime("%d/%m/%Y")
    mes_nome = MESES_PASTA.get(data_ligacao.strftime("%m"), "OUTROS")
    dia_pasta = formatar_pasta_dia(data_ligacao.day)
    
    pasta_destino = os.path.join(config["pasta_base"], mes_nome, dia_pasta)
    
    log(f"\n{'='*60}")
    log(f"📁 Campanha: {campanha_nome}")
    log(f"📂 Destino: {pasta_destino}")
    log(f"{'='*60}")
    
    total_chamadas = 0
    total_ok = 0
    total_existe = 0
    total_erro = 0
    
    for info in config["infos"]:
        log(f"\n🔍 Buscando Info. Adicional: {info}")
        
        chamadas = buscar_chamadas_por_info(session, info, data_str)
        
        if not chamadas:
            log(f"   ⚠️ Nenhuma chamada encontrada")
            continue
        
        log(f"   ✅ {len(chamadas)} chamadas encontradas")
        total_chamadas += len(chamadas)
        
        for chamada in chamadas:
            status = baixar_audio(session, chamada, pasta_destino)
            if status == "OK":
                total_ok += 1
                log(f"   ⬇️ OK: {chamada['CallIdMaster']}")
            elif status == "EXISTE":
                total_existe += 1
            else:
                total_erro += 1
                log(f"   ❌ ERRO: {chamada['CallIdMaster']} - {status}")
    
    return {
        "campanha": campanha_nome,
        "pasta": pasta_destino,
        "chamadas": total_chamadas,
        "ok": total_ok,
        "existe": total_existe,
        "erro": total_erro
    }

def main():
    print("=" * 60)
    print(" ATHOS AUTO DOWNLOAD - RESCISÃO (FORÇADO)")
    print(f" Baixando data específica: {DATA_ALVO}")
    print("=" * 60)
    
    # Converte a data alvo
    data_ligacao = datetime.strptime(DATA_ALVO, "%d/%m/%Y")
    data_str = data_ligacao.strftime("%d/%m/%Y")
    mes_nome = MESES_PASTA.get(data_ligacao.strftime("%m"), "OUTROS")
    dia_pasta = formatar_pasta_dia(data_ligacao.day)
    
    log(f"🎙 Baixando ligações de: {data_str}")
    log(f"📁 Pasta: {mes_nome}\\{dia_pasta}")
    
    # Login
    session = fazer_login()
    if not session:
        log("❌ Falha na autenticação")
        return
    
    # Processar ambas as campanhas
    resultados = []
    
    for campanha_nome, config in CAMPANHAS_RESCISAO.items():
        resultado = processar_campanha(campanha_nome, config, data_ligacao, session)
        resultados.append(resultado)
    
    # Resumo final
    print("\n" + "=" * 60)
    print("📊 RESUMO FINAL")
    print("=" * 60)
    
    for r in resultados:
        print(f"\n📁 {r['campanha']}")
        print(f"   Pasta: {r['pasta']}")
        print(f"   Chamadas: {r['chamadas']}")
        print(f"   Baixados: {r['ok']} | Existentes: {r['existe']} | Erros: {r['erro']}")
    
    print("\n" + "=" * 60)
    log("✅ Processamento concluído!")
    print("=" * 60)

if __name__ == "__main__":
    main()
