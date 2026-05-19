import os
import time
import glob

from colorama import Fore, init
init(autoreset=True)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime

# =========================
# CONFIGURAÇÃO DE PASTA
# =========================
PASTA_BASE = r"C:\Users\IAF\OneDrive\Ligações Monitoria\2026\Assinatura\GERAL"

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

# =========================
# CONECTAR AO CHROME
# =========================
options = Options()
options.debugger_address = "localhost:9222"
driver = webdriver.Chrome(options=options)

driver.execute_cdp_cmd(
    "Page.setDownloadBehavior",
    {"behavior": "allow", "downloadPath": PASTA_BASE}
)

# =========================
# CONTROLE
# =========================
processados = set()
total_baixados = 0

# =========================
# FUNÇÕES
# =========================
def barra(atual, total, baixados):
    tamanho = 20
    progresso = int((atual / total) * tamanho) if total else 0
    azul = Fore.CYAN + "█" * progresso
    cinza = Fore.WHITE + "░" * (tamanho - progresso)
    print(f"\r[{azul}{cinza}{Fore.RESET}] {atual}/{total} | Baixados: {Fore.CYAN}{baixados}{Fore.RESET}", end="")

def formatar_dia(dia_num):
    if dia_num < 10:
        return f"{dia_num:02d}"   # 1 → "01", 9 → "09"
    else:
        return f"0{dia_num}"      # 10 → "010", 13 → "013", 14 → "014"

def mover_ultimo_download(telefone, pasta_destino, timestamp_antes, timeout=20):
    """
    Aguarda um novo arquivo .mp3 aparecer na PASTA_BASE após o timestamp_antes,
    garantindo que é o arquivo recém-baixado e não um anterior.
    """
    inicio = time.time()
    while time.time() - inicio < timeout:
        # Aguarda terminar qualquer download em andamento
        temporarios = glob.glob(os.path.join(PASTA_BASE, "*.crdownload"))
        if temporarios:
            time.sleep(1)
            continue

        arquivos = glob.glob(os.path.join(PASTA_BASE, "*.mp3"))

        # Filtra apenas arquivos criados/modificados APÓS o início do download
        novos = [f for f in arquivos if os.path.getmtime(f) >= timestamp_antes]

        if novos:
            # Pega o mais recente entre os novos
            novo_arquivo = max(novos, key=os.path.getmtime)
            novo_nome = os.path.join(pasta_destino, f"{telefone}_{int(time.time())}.mp3")
            try:
                os.rename(novo_arquivo, novo_nome)
                return True
            except:
                return False

        time.sleep(1)
    return False

def scroll_tabela():
    for _ in range(10):
        driver.execute_script("window.scrollBy(0, 1000);")
        time.sleep(0.5)

def esperar_botao_pronto(expandida, tentativas=12, intervalo=1):
    for _ in range(tentativas):
        try:
            anchors = expandida.find_elements(By.XPATH, ".//a")
            for a in anchors:
                href = a.get_attribute("href")
                if href and "blob" in href and a.is_displayed() and a.is_enabled():
                    return a
        except:
            pass
        time.sleep(intervalo)
    return None

# =========================
# PROCESSAR
# =========================
def processar_pagina():
    global total_baixados

    scroll_tabela()
    time.sleep(2)

    linhas = driver.find_elements(By.XPATH, "//table/tbody/tr")
    total = len(linhas)
    atual = 0

    for linha in linhas:
        atual += 1
        try:
            texto = linha.text
            if not texto or texto in processados:
                barra(atual, total, total_baixados)
                continue

            processados.add(texto)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", linha)
            time.sleep(0.5)

            # telefone
            try:
                telefone = linha.find_element(By.XPATH, ".//strong").text
                telefone = telefone.replace("(", "").replace(")", "").replace("-", "").replace(" ", "")
            except:
                telefone = f"linha_{atual}"

            # DATA
            try:
                celulas = linha.find_elements(By.XPATH, ".//td")
                data_texto = None
                for td in celulas:
                    txt = td.text.strip()
                    if "/" in txt and "às" in txt:
                        data_texto = txt
                        break

                if not data_texto:
                    barra(atual, total, total_baixados)
                    continue

                data_obj = datetime.strptime(data_texto.split(" às ")[0].strip(), "%d/%m/%Y")
                mes = data_obj.strftime("%m")
                dia = formatar_dia(data_obj.day)
                nome_mes = MESES_PASTA.get(mes, "OUTROS")
                pasta_destino = os.path.join(PASTA_BASE, nome_mes, dia)

                if not os.path.exists(pasta_destino):
                    print(f"\n  [!] Pasta não encontrada: {pasta_destino}")
                    barra(atual, total, total_baixados)
                    continue

            except Exception as e:
                print(f"\n  [!] Erro ao ler data: {e}")
                pasta_destino = PASTA_BASE

            # abrir linha
            try:
                celula = linha.find_element(By.XPATH, ".//td[last()]")
                ActionChains(driver).move_to_element(celula).click().perform()
                time.sleep(2)
            except:
                barra(atual, total, total_baixados)
                continue

            # expandida
            expandida = None
            for _ in range(5):
                try:
                    expandida = linha.find_element(By.XPATH, "following-sibling::tr[1]")
                    if expandida.is_displayed():
                        break
                except:
                    time.sleep(1)

            if not expandida:
                barra(atual, total, total_baixados)
                continue

            # botão download
            botao = esperar_botao_pronto(expandida)
            if not botao:
                barra(atual, total, total_baixados)
                continue

            # ✅ Marca o timestamp ANTES de iniciar o download
            timestamp_antes = time.time()

            driver.execute_script("arguments[0].click();", botao)

            # ✅ Passa o timestamp para garantir que pega o arquivo correto
            if mover_ultimo_download(telefone, pasta_destino, timestamp_antes):
                total_baixados += 1

            barra(atual, total, total_baixados)

        except:
            barra(atual, total, total_baixados)
            continue

# =========================
# EXECUÇÃO
# =========================
input(">>> Deixe o relatório pronto e pressione ENTER <<<\n")

while True:
    processar_pagina()

    try:
        print("\nIndo para próxima página...")
        antes = driver.find_element(By.XPATH, "//table/tbody/tr[1]").text
        proximo = driver.find_element(By.XPATH, "//button[@aria-label='Próximo' or contains(., 'Próximo')]")
        driver.execute_script("arguments[0].click();", proximo)
        WebDriverWait(driver, 10).until(
            lambda d: d.find_element(By.XPATH, "//table/tbody/tr[1]").text != antes
        )
        time.sleep(2)
    except:
        print("\nFim das páginas")
        break

print("\n")
print(Fore.CYAN + "=" * 40)
print(Fore.GREEN + "✔  DOWNLOADS CONCLUÍDOS COM SUCESSO")
print(Fore.CYAN + f"   TOTAL DE ÁUDIOS BAIXADOS: {total_baixados}")
print(Fore.CYAN + "=" * 40)
