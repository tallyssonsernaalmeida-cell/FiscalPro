import os
import pandas as pd
import requests
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
from colorama import init, Fore
from send2trash import send2trash
import time
from datetime import datetime

init(autoreset=True)

# ---------------- CONFIG ----------------
PASTA_BASE = r"C:\Users\IAF\OneDrive\Ligações Monitoria\2026\Assinatura\GERAL"

TIMEOUT = 30
RETRIES = 5
PAUSE_BETWEEN_RETRIES = 2
MAX_WORKERS = 15

DOWNLOADS_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
ERROR_LOG = os.path.join(PASTA_BASE, "erros.txt")

MESES_PASTA = {
    "01": "JANEIRO 01",
    "02": "FEVEREIRO 02",
    "03": "MARÇO 03",
    "04": "ABRIL 04",
    "05": "MAIO 05",
    "06": "JUNHO 06",
    "07": "JULHO 07",
    "08": "AGOSTO 08",
    "09": "SETEMBRO 09",
    "10": "OUTUBRO 10",
    "11": "NOVEMBRO 11",
    "12": "DEZEMBRO 12",
}

# ---------------- FUNÇÕES ----------------
def get_latest_csv(folder):
    while True:
        arquivos = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(".csv")
        ]

        if not arquivos:
            time.sleep(1)
            continue

        arquivo = max(arquivos, key=os.path.getmtime)

        size1 = os.path.getsize(arquivo)
        time.sleep(1)
        size2 = os.path.getsize(arquivo)

        if size1 == size2:
            return arquivo

def safe_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

def safe_foldername(name):
    return re.sub(r'[\\/*?:"<>| ]', "_", name)

def download_audio(row):
    link = str(row.get('recording', '')).strip()
    idx = row.get('idx', 0)

    if not link.startswith("http"):
        return "IGNORADO", None

    number = str(row.get('number', 'unknown'))
    call_date_str = str(row.get('call_date', 'unknown'))
    list_name = str(row.get('list_name', 'unknown'))

    try:
        data_obj = datetime.strptime(call_date_str.split()[0], "%d/%m/%Y")
        mes = data_obj.strftime("%m")
        dia = data_obj.strftime("%d").zfill(3)
        pasta_mes = MESES_PASTA.get(mes)
    except:
        return "ERRO", None

    pasta_destino = os.path.join(PASTA_BASE, pasta_mes, dia)

    if not os.path.exists(pasta_destino):
        return "ERRO", None

    filename = f"{number}_{safe_foldername(call_date_str)}_{list_name}_{idx}.mp3"
    filepath = os.path.join(pasta_destino, safe_filename(filename))

    if os.path.exists(filepath):
        return "JA_EXISTE", pasta_destino

    for _ in range(RETRIES):
        try:
            r = requests.get(link, timeout=TIMEOUT)
            if r.status_code == 200:
                with open(filepath, "wb") as f:
                    f.write(r.content)
                return "SUCESSO", pasta_destino
        except:
            pass
        time.sleep(PAUSE_BETWEEN_RETRIES)

    return "ERRO", None

# ---------------- INÍCIO ----------------
print(f"""
{Fore.CYAN}╔══════════════════════════════════════════════════════════╗
║        BAIXANDO OS AUDIOS DE ASSINATURA                  ║
╠══════════════════════════════════════════════════════════╣
║  CAMPANHA : ASSINATURA                                  ║
║  PERÍODO  : 07:00:00 ATÉ 20:00:00                        ║
║  OPERAÇÃO : DOWNLOAD DE ÁUDIOS                           ║
║  USUÁRIO  : TALLYSSON ALMEIDA                           ║
╚══════════════════════════════════════════════════════════╝
""")

print(f"{Fore.YELLOW}>> AGUARDANDO DADOS DO RELATÓRIO...\n")

# ---------------- CSV ----------------
CSV_FILE = get_latest_csv(DOWNLOADS_PATH)
print(f"{Fore.GREEN}>> ARQUIVO CARREGADO: {os.path.basename(CSV_FILE)}\n")

df = pd.read_csv(CSV_FILE, sep=";", encoding="utf-8", low_memory=False)
df.columns = df.columns.str.strip().str.lower()
df['idx'] = df.index

# ---------------- EXECUÇÃO ----------------
print(f"{Fore.CYAN}>> INICIANDO PROCESSAMENTO DE ÁUDIOS...\n")

results = {"SUCESSO":0, "ERRO":0, "JA_EXISTE":0, "IGNORADO":0}
ultima_pasta = None

with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
    futures = [executor.submit(download_audio, row) for _, row in df.iterrows()]

    with tqdm(
        total=len(futures),
        ncols=80,
        bar_format=f"{Fore.GREEN}{{l_bar}}{{bar}}{Fore.RESET}| {{n_fmt}}/{{total_fmt}}"
    ) as pbar:

        for f in as_completed(futures):
            status, pasta = f.result()
            results[status] += 1

            if status == "SUCESSO":
                ultima_pasta = pasta

            pbar.update(1)

# ---------------- FINAL ----------------
print(f"""
{Fore.CYAN}╔══════════════════════════════════════════════════════════╗
║                PROCESSAMENTO FINALIZADO                  ║
╠══════════════════════════════════════════════════════════╣
║  STATUS : CONCLUÍDO COM SUCESSO                          ║
║  MÓDULO : DOWNLOAD DE ÁUDIOS                             ║
╚══════════════════════════════════════════════════════════╝
""")

print(f"{Fore.WHITE}RESUMO OPERACIONAL\n")
print(f"{Fore.GREEN}  ✔ Downloads realizados : {results['SUCESSO']}")
print(f"{Fore.BLUE}  ℹ Arquivos existentes  : {results['JA_EXISTE']}")
print(f"{Fore.CYAN}  ⚠ Registros ignorados : {results['IGNORADO']}")
print(f"{Fore.RED}  ✖ Falhas no processo  : {results['ERRO']}")

# ---------------- LIXEIRA ----------------
try:
    send2trash(CSV_FILE)
    print(f"\n{Fore.YELLOW}>> EXCEL.CSV MOVIDO PARA LIXEIRA")
except Exception as e:
    print(f"{Fore.RED}Erro ao mover para lixeira: {e}")

# ---------------- ABRIR PASTA ----------------
if ultima_pasta and os.path.exists(ultima_pasta):
    print(f"{Fore.GREEN}>> AGUARDE A PASTA DESTINO PARA CONFERENCIA...\n")
    os.startfile(ultima_pasta)
