# athos_download_maio.py - VERSÃO FINAL COM VALIDAÇÃO E LOG DE ERROS

import os
import re
import json
import time
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests

# ================================================================
# CONFIGURAÇÕES
# ================================================================
BASE_URL   = "https://athossoluoes.3c.plus/api/v1"
AUDIO_BASE = "https://app.3c.plus/api/v1/calls"
PASTA_BASE = r"C:\Users\IAF\OneDrive\Ligações Monitoria\2026\Assinatura\GERAL"

CAMPANHAS     = {"ASSINATURA": 253640}
QUALIFICACOES = [204212, 169571, 169563, 169562, 169561, 169560, 169559,
                 169558, 169557, 169556, 169540, 166376, 166375]

AGENT_IDS = "210227,216056,211372,211769,191932,211387,211733,216032,211743,211735,211749,211383,211746,216029,213029,216054,211369,211745,211368,216058,204819,190334,210950,216057,207698,206299,215984,211739,211370,212829,212827,211771,211742,211738,211728,211727,211371,208580,202490,206300,205338,183863,194776,190404,190409"

STATUS_FIXO  = 7
HORA_INI     = "07:00:00"
HORA_FIM     = "20:00:00"
MAX_WORKERS  = 5
TIMEOUT      = 120
RETRIES      = 15
PAUSE_RETRY  = 3

MESES_PASTA = {
    "01": "JANEIRO 01", "02": "FEVEREIRO 02", "03": "MARÇO 03",
    "04": "ABRIL 04",   "05": "MAIO 05",       "06": "JUNHO 06",
    "07": "JULHO 07",   "08": "AGOSTO 08",     "09": "SETEMBRO 09",
    "10": "OUTUBRO 10", "11": "NOVEMBRO 11",   "12": "DEZEMBRO 12",
}

FERIADOS = {
    "01/01/2026": "Ano Novo",             "03/04/2026": "Sexta-feira Santa",
    "21/04/2026": "Tiradentes",           "01/05/2026": "Dia do Trabalho",
    "07/09/2026": "Independência",        "12/10/2026": "Nossa Sra. Aparecida",
    "02/11/2026": "Finados",              "25/12/2026": "Natal",
}

# Caminho fixo do script — evita problema quando roda via SYSTEM
SCRIPT_DIR = r"C:\Users\IAF\Desktop\Script\3C\SRC (CODIGO PRINCIPAL DO SISTEMA)\outros_scripts"
CRED_FILE  = os.path.join(SCRIPT_DIR, ".athos_painel_creds.json")
TOKEN_FILE = os.path.join(SCRIPT_DIR, ".athos_painel_token.json")
LOG_FILE   = os.path.join(SCRIPT_DIR, "athos_download_log.txt")

# ================================================================
# UTILITÁRIOS
# ================================================================
def log(msg):
    linha = f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] {msg}"
    print(linha)
    # Salva no arquivo de log para consultar depois
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except:
        pass

def formatar_pasta_dia(dia):
    """
    Dias 1-9   → 2 dígitos: 04, 05, 06, 07, 08, 09
    Dias 10-31 → 3 dígitos: 010, 011, 012, 031
    """
    dia = int(dia)  # garante que é inteiro
    if dia < 10:
        return f"0{dia}"   # força string manual: 04, 05, 06, 07...
    else:
        return f"0{dia}"   # força string manual: 010, 011, 012...

def eh_dia_util(data):
    if data.weekday() >= 5:
        return False
    return data.strftime("%d/%m/%Y") not in FERIADOS

def dia_util_anterior(data):
    d = data - timedelta(days=1)
    while not eh_dia_util(d):
        d -= timedelta(days=1)
    return d

# ================================================================
# AUTENTICAÇÃO
# ================================================================
def fazer_login():
    if not os.path.exists(CRED_FILE):
        log(f"❌ Arquivo de credenciais não encontrado: {CRED_FILE}")
        modelo = {"usuario": "seu_usuario", "senha": "sua_senha"}
        with open(CRED_FILE, "w", encoding="utf-8") as f:
            json.dump(modelo, f, indent=2)
        log("   Edite o arquivo com suas credenciais e execute novamente")
        return None

    try:
        with open(CRED_FILE, "r", encoding="utf-8") as f:
            creds = json.load(f)
    except Exception as e:
        log(f"❌ Erro ao ler credenciais: {e}")
        return None

    usuario = creds.get("usuario", "")
    senha   = creds.get("senha", "")

    if not usuario or not senha or usuario == "seu_usuario":
        log("❌ Credenciais inválidas ou não configuradas!")
        return None

    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, "r", encoding="utf-8") as f:
                token = json.load(f).get("token")
            if token:
                r = requests.get(
                    f"{BASE_URL}/me",
                    headers={"Authorization": f"Bearer {token}"},
                    timeout=5
                )
                if r.status_code == 200:
                    log("✔ Token válido")
                    return token
        except:
            pass

    try:
        log("🔄 Fazendo login...")
        r = requests.post(
            f"{BASE_URL}/authenticate?user={usuario}",
            json={"user": usuario, "password": senha, "token_type": "jwt"},
            timeout=10
        )
        if r.status_code == 200:
            token = r.json()["data"]["api_token"]
            with open(TOKEN_FILE, "w", encoding="utf-8") as f:
                json.dump({"token": token}, f)
            log("✔ Login realizado!")
            return token
        else:
            log(f"❌ Falha no login - Status: {r.status_code}")
            return None
    except Exception as e:
        log(f"❌ Erro no login: {e}")
        return None

# ================================================================
# BUSCA DE CHAMADAS
# ================================================================
def buscar_chamadas(token, data_str):
    d          = datetime.strptime(data_str, "%d/%m/%Y")
    start_date = f"{d.strftime('%Y-%m-%d')} {HORA_INI}"
    end_date   = f"{d.strftime('%Y-%m-%d')} {HORA_FIM}"
    headers    = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    calls, ids_vistos, page = [], set(), 0

    while True:
        params = {
            "page":             page,
            "per_page":         100,
            "start_date":       start_date,
            "end_date":         end_date,
            "call_mode":        "all",
            "type":             "all",
            "campaigns[]":      CAMPANHAS["ASSINATURA"],
            "statuses":         STATUS_FIXO,
            "simple_paginate":  "true",
            "qualification":    ",".join(map(str, QUALIFICACOES)),
            "agent_ids":        AGENT_IDS,
            "minimum_duration": 60,
        }
        try:
            r = requests.get(f"{BASE_URL}/calls", headers=headers, params=params, timeout=30)
            if r.status_code != 200:
                log(f"⚠️ Página {page} retornou status {r.status_code} — parando busca")
                break

            data = r.json().get("data", [])
            if not data:
                break

            novos = 0
            for c in data:
                cid = c.get("id")
                if cid in ids_vistos:
                    continue
                if not str(c.get("recording") or "").startswith("http"):
                    continue
                if not str(c.get("qualification") or "").upper().startswith("CPC"):
                    continue
                ids_vistos.add(cid)
                calls.append(c)
                novos += 1

            log(f"   📄 Página {page}: +{novos} (total: {len(calls)})")
            page += 1
            if page > 50:
                break

        except Exception as e:
            log(f"⚠️ Erro na página {page}: {e} — aguardando 3s...")
            time.sleep(3)

    return calls

# ================================================================
# DOWNLOAD DE ÁUDIO
# ================================================================
def baixar_audio(chamada, token):
    call_id   = chamada.get("id")
    call_date = str(chamada.get("call_date", ""))[:10]
    number    = str(chamada.get("number", "unknown"))
    list_name = str(chamada.get("list", "unknown"))

    try:
        data_obj = datetime.strptime(call_date, "%Y-%m-%d")

        if data_obj.month != 5:
            return "PULAR"

        mes            = MESES_PASTA.get(data_obj.strftime("%m"), "OUTROS")
        dia_num        = int(data_obj.strftime("%d"))
        nome_pasta_dia = formatar_pasta_dia(dia_num)
        pasta          = os.path.join(PASTA_BASE, mes, nome_pasta_dia)

        filename = f"{number}_{data_obj.strftime('%d%m%Y')}_{list_name}_{call_id}.mp3"
        filepath = os.path.join(pasta, re.sub(r'[\\/*?:"<>|]', "_", filename))

        if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
            return "EXISTE"

        if os.path.exists(filepath):
            os.remove(filepath)

        if not os.path.exists(pasta):
            log(f"   ❌ Pasta não encontrada: {pasta}")
            return "ERRO"

        link = chamada.get("recording", "")
        if not link.startswith("http"):
            link = f"{AUDIO_BASE}/{call_id}/recording"

        headers = {"Authorization": f"Bearer {token}"}

        for tentativa in range(1, RETRIES + 1):
            try:
                r = requests.get(link, headers=headers, timeout=TIMEOUT, stream=True)

                if r.status_code == 200:
                    with open(filepath, "wb") as f:
                        for chunk in r.iter_content(8192):
                            if chunk:
                                f.write(chunk)

                    if os.path.getsize(filepath) > 1024:
                        log(f"   ✅ BAIXADO: {call_id}")
                        return "OK"
                    else:
                        log(f"   ⚠️ Arquivo muito pequeno ({call_id}), tentativa {tentativa}/{RETRIES}")
                        os.remove(filepath)

                elif r.status_code == 429:
                    log(f"   ⚠️ Rate limit (429) para {call_id}, tentativa {tentativa}/{RETRIES} — aguardando 10s")
                    time.sleep(10)
                    continue

                else:
                    log(f"   ⚠️ Status {r.status_code} para {call_id}, tentativa {tentativa}/{RETRIES}")

            except Exception as e:
                log(f"   ⚠️ Tentativa {tentativa}/{RETRIES} - Erro: {str(e)[:80]}")

            time.sleep(PAUSE_RETRY)

        log(f"   ❌ ERRO DEFINITIVO: {call_id} — esgotou {RETRIES} tentativas")
        return "ERRO"

    except Exception as e:
        log(f"   ❌ Erro crítico em {call_id}: {str(e)[:100]}")
        return "ERRO"

# ================================================================
# MAIN
# ================================================================
def main():
    log("=" * 50)
    log("ATHOS DOWNLOAD - ASSINATURA (MAIO)")
    log("=" * 50)

    hoje = datetime.now()

    if not eh_dia_util(hoje):
        log(f"{hoje.strftime('%d/%m/%Y')} - Não é dia útil")
        return

    data_ligacao = dia_util_anterior(hoje)

    if data_ligacao.month != 5:
        log(f"Ligação de {data_ligacao.strftime('%d/%m/%Y')} não é de MAIO — ignorando")
        return

    data_str   = data_ligacao.strftime("%d/%m/%Y")
    dia_num    = int(data_ligacao.strftime("%d"))
    nome_pasta = formatar_pasta_dia(dia_num)

    log(f"Data hoje: {hoje.strftime('%d/%m/%Y')}")
    log(f"Baixando ligações de: {data_str}")
    log(f"Pasta destino: MAIO 05\\{nome_pasta}")

    token = fazer_login()
    if not token:
        log("❌ Falha na autenticação")
        return

    calls = buscar_chamadas(token, data_str)
    if not calls:
        log(f"⚠️ Nenhuma chamada encontrada para {data_str}")
        return

    log(f"✅ {len(calls)} chamadas encontradas")

    resultados = {"OK": 0, "EXISTE": 0, "ERRO": 0, "PULAR": 0}
    erros_ids  = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(baixar_audio, c, token): c for c in calls}
        for i, future in enumerate(as_completed(futures), 1):
            chamada = futures[future]
            status  = future.result()
            resultados[status] += 1

            if status == "ERRO":
                erros_ids.append(chamada.get("id"))

            if i % 10 == 0 or i == len(calls):
                log(f"Progresso: {i}/{len(calls)}")

    log("=" * 50)
    log(f"✅ DOWNLOAD CONCLUÍDO!")
    log(f"   OK: {resultados['OK']} | Existia: {resultados['EXISTE']} | Erros: {resultados['ERRO']}")

    if erros_ids:
        log(f"⚠️ IDs com erro ({len(erros_ids)}): {', '.join(erros_ids)}")

    log(f"📁 Pasta: MAIO 05\\{nome_pasta}")
    log("=" * 50)


if __name__ == "__main__":
    main()
