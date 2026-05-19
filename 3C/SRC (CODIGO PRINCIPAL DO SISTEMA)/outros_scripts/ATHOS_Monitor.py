"""
ATHOS - RELATÓRIO DE LIGAÇÕES DO 3C PLUS - CAMPANHA: ASSINATURA
Painel de Histórico de Downloads - COM LOGO DA EMPRESA
"""

import os
import json
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk  # Precisa instalar: pip install Pillow

# ================================================================
# CORES
# ================================================================
ATHOS_AZUL  = "#1B2B8A"
ATHOS_BG    = "#0d0f1a"
ATHOS_BG2   = "#13172a"
ATHOS_BG3   = "#1a1f35"
VERDE       = "#22c55e"
VERMELHO    = "#f87171"
AMARELO     = "#facc15"
CINZA       = "#A0A8B8"
BRANCO      = "#ffffff"
AZUL_CLARO  = "#6ea4ff"
LARANJA     = "#ff9f4a"

# ================================================================
# CONFIG
# ================================================================
PASTA_BASE = r"C:\Users\IAF\OneDrive\Ligações Monitoria\2026\Assinatura\GERAL"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(SCRIPT_DIR, "downloads_realizados.json")

MESES_PASTA = {
    "01": "JANEIRO 01",  "02": "FEVEREIRO 02", "03": "MARÇO 03",
    "04": "ABRIL 04",    "05": "MAIO 05",       "06": "JUNHO 06",
    "07": "JULHO 07",    "08": "AGOSTO 08",     "09": "SETEMBRO 09",
    "10": "OUTUBRO 10",  "11": "NOVEMBRO 11",   "12": "DEZEMBRO 12",
}

FERIADOS_2026 = {
    "01/01/2026": "Ano Novo", "03/04/2026": "Sexta-feira Santa",
    "21/04/2026": "Tiradentes", "01/05/2026": "Dia do Trabalho",
    "07/09/2026": "Independência", "12/10/2026": "Nossa Sra. Aparecida",
    "02/11/2026": "Finados", "25/12/2026": "Natal",
}

# ================================================================
# FUNÇÕES DE DATA
# ================================================================
def eh_dia_util(data):
    if data.weekday() >= 5:
        return False
    if data.strftime("%d/%m/%Y") in FERIADOS_2026:
        return False
    return True

def dia_util_anterior(data):
    d = data - timedelta(days=1)
    while not eh_dia_util(d):
        d -= timedelta(days=1)
    return d

def proximo_dia_util(data):
    d = data + timedelta(days=1)
    while not eh_dia_util(d):
        d += timedelta(days=1)
    return d

def nome_dia(data):
    dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
    return dias[data.weekday()]

def formatar_pasta_dia(dia):
    if 1 <= dia <= 9:
        return f"{dia:02d}"
    else:
        return f"{dia:03d}"

# ================================================================
# HISTÓRICO DO DOWNLOADER
# ================================================================
def carregar_historico_downloader():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {"baixados": []}

def salvar_historico_downloader(historico):
    try:
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(historico, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar: {e}")

def ja_foi_baixado_pelo_downloader(data):
    data_str = data.strftime("%d/%m/%Y")
    historico = carregar_historico_downloader()
    return data_str in historico.get("baixados", [])

def sincronizar_com_pastas():
    encontrados = []
    
    for mes_cod, mes_nome in MESES_PASTA.items():
        pasta_mes = os.path.join(PASTA_BASE, mes_nome)
        if not os.path.exists(pasta_mes):
            continue
            
        for dia_pasta in os.listdir(pasta_mes):
            pasta_dia = os.path.join(pasta_mes, dia_pasta)
            if not os.path.isdir(pasta_dia):
                continue
                
            mp3s = [f for f in os.listdir(pasta_dia) if f.endswith('.mp3')]
            if mp3s:
                try:
                    dia_num = int(dia_pasta)
                    data = datetime(2026, int(mes_cod), dia_num)
                    data_str = data.strftime("%d/%m/%Y")
                    encontrados.append(data_str)
                except:
                    pass
    
    historico = {"baixados": sorted(list(set(encontrados)))}
    salvar_historico_downloader(historico)
    return encontrados

# ================================================================
# LÓGICA DO PRÓXIMO DOWNLOAD
# ================================================================
def encontrar_proximo_download_real(hoje):
    datas_baixadas = carregar_historico_downloader().get("baixados", [])
    
    def ja_foi_baixada(data):
        return data.strftime("%d/%m/%Y") in datas_baixadas
    
    hora_atual = hoje.hour
    minuto_atual = hoje.minute
    
    if eh_dia_util(hoje):
        if hora_atual < 8 or (hora_atual == 8 and minuto_atual == 0):
            data_teste = hoje
        else:
            data_teste = proximo_dia_util(hoje)
    else:
        data_teste = proximo_dia_util(hoje)
    
    for _ in range(30):
        if not eh_dia_util(data_teste):
            data_teste = proximo_dia_util(data_teste)
            continue
        
        data_ligacao = dia_util_anterior(data_teste)
        
        if not ja_foi_baixada(data_ligacao):
            return data_teste, data_ligacao
        
        data_teste = proximo_dia_util(data_teste)
    
    data_teste = proximo_dia_util(hoje)
    return data_teste, dia_util_anterior(data_teste)

# ================================================================
# FUNÇÕES DO MONITOR
# ================================================================
def pasta_do_dia(data):
    mes = data.strftime("%m")
    dia_num = int(data.strftime("%d"))
    dia = formatar_pasta_dia(dia_num)
    nome_mes = MESES_PASTA.get(mes, "")
    if not nome_mes:
        return None, None
    pasta = os.path.join(PASTA_BASE, nome_mes, dia)
    pasta_display = f"{nome_mes}\\{dia}"
    return pasta, pasta_display

def contar_mp3(pasta):
    if not pasta or not os.path.exists(pasta):
        return 0
    return len(list(Path(pasta).glob("*.mp3")))

def data_modificacao_pasta(pasta):
    if not pasta or not os.path.exists(pasta):
        return None
    mp3s = list(Path(pasta).glob("*.mp3"))
    if not mp3s:
        return None
    mais_recente = max(mp3s, key=lambda f: f.stat().st_mtime)
    return datetime.fromtimestamp(mais_recente.stat().st_mtime)

def obter_historico(dias=60):
    historico = []
    hoje = datetime.now().date()

    for i in range(dias):
        data = datetime.now() - timedelta(days=i)
        data_dt = data.date()
        data_str = data.strftime("%d/%m/%Y")
        dia_semana = nome_dia(data)

        pasta_path, pasta_display = pasta_do_dia(data)

        if not eh_dia_util(data):
            if data.weekday() == 5:
                tipo = "SÁBADO"
            elif data.weekday() == 6:
                tipo = "DOMINGO"
            else:
                nome = FERIADOS_2026.get(data.strftime("%d/%m/%Y"), "")
                tipo = f"FERIADO — {nome}" if nome else "FERIADO"

            historico.append({
                "data": data_str, "dia": dia_semana, "qtd": 0,
                "status": tipo, "download_em": "---",
                "pasta": pasta_display or "---",
                "pasta_path": pasta_path or "",
                "util": False,
            })
            continue

        qtd = contar_mp3(pasta_path)
        dt_baixado = data_modificacao_pasta(pasta_path)
        
        baixado_pelo_downloader = ja_foi_baixado_pelo_downloader(data)

        if qtd > 0 or baixado_pelo_downloader:
            status = "BAIXADO"
            download_em = dt_baixado.strftime("%d/%m/%Y") if dt_baixado else "desconhecido"
        elif data_dt < hoje:
            status = "SEM CPC"
            download_em = proximo_dia_util(data).strftime("%d/%m/%Y")
        elif data_dt == hoje:
            status = "PENDENTE"
            download_em = proximo_dia_util(data).strftime("%d/%m/%Y")
        else:
            status = "FUTURO"
            download_em = proximo_dia_util(data).strftime("%d/%m/%Y")

        historico.append({
            "data": data_str,
            "dia": dia_semana,
            "qtd": qtd if qtd > 0 else ("PENDENTE" if status == "PENDENTE" else 0),
            "status": status,
            "download_em": download_em,
            "pasta": pasta_display or "---",
            "pasta_path": pasta_path or "",
            "util": True,
        })

    return historico

# ================================================================
# UI PRINCIPAL COM LOGO
# ================================================================
class PainelHistorico:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ATHOS - Relatório de Ligações do 3C Plus - Campanha: ASSINATURA")
        self.root.geometry("1200x750")
        self.root.configure(bg=ATHOS_BG)
        self.root.resizable(True, True)

        # ============================================================
        # CARREGAR LOGO DA EMPRESA
        # ============================================================
        self.logo_image = None
        self.logo_photo = None
        self.carregar_logo()
        
        self._build_ui()
        self._centralizar()
        self.atualizar()

    def carregar_logo(self):
        """Tenta carregar a logo da empresa de diferentes locais"""
        possiveis_caminhos = [
            os.path.join(SCRIPT_DIR, "logo.png"),
            os.path.join(SCRIPT_DIR, "logo.jpg"),
            os.path.join(SCRIPT_DIR, "logo.jpeg"),
            os.path.join(SCRIPT_DIR, "ATHOS_Logo.png"),
            os.path.join(SCRIPT_DIR, "logo_empresa.png"),
            r"C:\Users\IAF\Desktop\logo.png",
            r"C:\Users\IAF\Pictures\logo.png",
        ]
        
        for caminho in possiveis_caminhos:
            if os.path.exists(caminho):
                try:
                    # Abre e redimensiona a imagem
                    img = Image.open(caminho)
                    img = img.resize((60, 60), Image.Resampling.LANCZOS)
                    self.logo_photo = ImageTk.PhotoImage(img)
                    print(f"✅ Logo carregada de: {caminho}")
                    return
                except Exception as e:
                    print(f"Erro ao carregar logo de {caminho}: {e}")
        
        print("⚠️ Logo não encontrada. Usando apenas texto.")
        # Cria um placeholder com texto
        self.logo_photo = None

    def _centralizar(self):
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"1200x750+{(sw-1200)//2}+{(sh-750)//2}")

    def _build_ui(self):
        # HEADER COM LOGO
        header = tk.Frame(self.root, bg=ATHOS_BG2, height=100)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        # Frame para logo e título lado a lado
        top_frame = tk.Frame(header, bg=ATHOS_BG2)
        top_frame.pack(pady=(10,2))
        
        # Logo (se existir)
        if self.logo_photo:
            logo_label = tk.Label(top_frame, image=self.logo_photo, bg=ATHOS_BG2)
            logo_label.pack(side="left", padx=(10,15))
        
        # Título
        tk.Label(top_frame, text="ATHOS - RELATÓRIO DE LIGAÇÕES DO 3C PLUS",
                 font=("Consolas", 18, "bold"), fg=BRANCO, bg=ATHOS_BG2).pack(side="left")
        
        # Subtítulo
        tk.Label(header, text="Campanha: ASSINATURA",
                 font=("Consolas", 11), fg=AZUL_CLARO, bg=ATHOS_BG2).pack()
        tk.Label(header, text="Com lógica do Downloader Automático | Pula datas já baixadas",
                 font=("Consolas", 9), fg=CINZA, bg=ATHOS_BG2).pack(pady=(2,5))

        # BANNER PRÓXIMO DOWNLOAD
        banner = tk.Frame(self.root, bg=ATHOS_BG3,
                          highlightbackground=ATHOS_AZUL, highlightthickness=2)
        banner.pack(fill="x", padx=15, pady=(10,0))
        self.lbl_banner = tk.Label(banner, text="Carregando...",
                                   font=("Consolas", 12, "bold"),
                                   fg=LARANJA, bg=ATHOS_BG3, justify="center")
        self.lbl_banner.pack(pady=10)

        # CARDS
        cards = tk.Frame(self.root, bg=ATHOS_BG)
        cards.pack(fill="x", padx=15, pady=8)
        self.c_pend = self._card(cards, "DIAS PENDENTES", "0", VERMELHO)
        self.c_prox = self._card(cards, "PRÓXIMO DOWNLOAD", "---", AZUL_CLARO)
        self.c_prox_lig = self._card(cards, "LIGAÇÃO DE", "---", VERDE)

        # TABELA
        frame_tab = tk.Frame(self.root, bg=ATHOS_BG3)
        frame_tab.pack(fill="both", expand=True, padx=15, pady=(0,4))

        scroll_y = tk.Scrollbar(frame_tab, orient="vertical")
        scroll_y.pack(side="right", fill="y")
        scroll_x = tk.Scrollbar(frame_tab, orient="horizontal")
        scroll_x.pack(side="bottom", fill="x")

        cols = ("data", "dia", "qtd", "status", "download_em", "pasta_caminho")

        self.tree = ttk.Treeview(frame_tab, columns=cols, show="headings",
                                 height=20,
                                 yscrollcommand=scroll_y.set,
                                 xscrollcommand=scroll_x.set)

        self.tree.heading("data", text="DATA LIGAÇÃO")
        self.tree.heading("dia", text="DIA")
        self.tree.heading("qtd", text="ÁUDIOS")
        self.tree.heading("status", text="STATUS")
        self.tree.heading("download_em", text="DOWNLOAD EM")
        self.tree.heading("pasta_caminho", text="PASTA DESTINO")

        self.tree.column("data", width=120, anchor="center")
        self.tree.column("dia", width=90, anchor="center")
        self.tree.column("qtd", width=80, anchor="center")
        self.tree.column("status", width=220, anchor="w")
        self.tree.column("download_em", width=130, anchor="center")
        self.tree.column("pasta_caminho", width=550, anchor="w")

        estilo = ttk.Style()
        estilo.theme_use("default")
        estilo.configure("Treeview",
                         background=ATHOS_BG2, foreground=BRANCO,
                         fieldbackground=ATHOS_BG2,
                         font=("Consolas", 10), rowheight=26)
        estilo.configure("Treeview.Heading",
                         background=ATHOS_BG3, foreground=CINZA,
                         font=("Consolas", 9, "bold"))
        estilo.map("Treeview", background=[("selected", ATHOS_AZUL)])

        self.tree.tag_configure("baixado", foreground=VERDE)
        self.tree.tag_configure("pendente", foreground=AMARELO)
        self.tree.tag_configure("sem_cpc", foreground=CINZA)
        self.tree.tag_configure("sabado", foreground=AMARELO)
        self.tree.tag_configure("domingo", foreground=AMARELO)
        self.tree.tag_configure("feriado", foreground=VERMELHO)
        self.tree.tag_configure("proximo", foreground=LARANJA, font=("Consolas", 10, "bold"))

        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>", self._abrir_pasta_selecionada)

        # RODAPÉ
        rodape = tk.Frame(self.root, bg=ATHOS_BG)
        rodape.pack(fill="x", padx=15, pady=6)

        tk.Button(rodape, text="📂 ABRIR PASTA BASE", bg=ATHOS_BG3, fg=BRANCO,
                  font=("Consolas", 10, "bold"),
                  command=lambda: os.startfile(PASTA_BASE),
                  padx=12, pady=5, relief="flat", cursor="hand2").pack(side="left", padx=3)

        tk.Button(rodape, text="🔄 ATUALIZAR", bg=ATHOS_AZUL, fg=BRANCO,
                  font=("Consolas", 10, "bold"),
                  command=self.atualizar,
                  padx=12, pady=5, relief="flat", cursor="hand2").pack(side="left", padx=3)

        tk.Button(rodape, text="💾 EXPORTAR TXT", bg=ATHOS_BG3, fg=BRANCO,
                  font=("Consolas", 10),
                  command=self.exportar,
                  padx=12, pady=5, relief="flat", cursor="hand2").pack(side="left", padx=3)

        tk.Button(rodape, text="🔄 SINCRONIZAR", bg=ATHOS_BG3, fg=BRANCO,
                  font=("Consolas", 10),
                  command=self.sincronizar,
                  padx=12, pady=5, relief="flat", cursor="hand2").pack(side="left", padx=3)

        tk.Label(rodape, text="ATHOS SOLUÇÕES EM RECUPERAÇÃO DE CRÉDITO",
                 font=("Consolas", 8), fg=AZUL_CLARO, bg=ATHOS_BG).pack(side="left", padx=15)

        self.lbl_hora = tk.Label(rodape, text="",
                                 font=("Consolas", 9), fg=CINZA, bg=ATHOS_BG)
        self.lbl_hora.pack(side="right")

    def _card(self, parent, titulo, valor, cor):
        card = tk.Frame(parent, bg=ATHOS_BG2,
                        highlightbackground="#2a3060", highlightthickness=1,
                        padx=10, pady=8)
        card.pack(side="left", expand=True, fill="x", padx=3)
        tk.Label(card, text=titulo, font=("Consolas", 8),
                 fg="#555e7a", bg=ATHOS_BG2).pack()
        lbl = tk.Label(card, text=valor,
                       font=("Consolas", 18, "bold"), fg=cor, bg=ATHOS_BG2)
        lbl.pack()
        return lbl

    def sincronizar(self):
        encontrados = sincronizar_com_pastas()
        messagebox.showinfo("Sincronizado", 
                           f"Sincronizado com sucesso!\n{len(encontrados)} dias encontrados nas pastas.")
        self.atualizar()

    def atualizar(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        historico = obter_historico(60)
        hoje = datetime.now()
        dias_pend = 0

        data_download, data_ligacao = encontrar_proximo_download_real(hoje)
        
        self.lbl_banner.config(
            text=f"🎯 PRÓXIMO DOWNLOAD REAL: {data_download.strftime('%d/%m/%Y')} ({nome_dia(data_download)}) às 08:00\n"
                 f"📞 Ligação de: {data_ligacao.strftime('%d/%m/%Y')} ({nome_dia(data_ligacao)})"
        )
        self.c_prox.config(text=data_download.strftime("%d/%m/%Y"))
        self.c_prox_lig.config(text=data_ligacao.strftime("%d/%m/%Y"))

        for item in historico:
            status = item["status"]
            qtd = item["qtd"]
            data_item = datetime.strptime(item["data"], "%d/%m/%Y")

            if status == "BAIXADO":
                tag = "baixado"
            elif status == "PENDENTE":
                tag = "pendente"
                dias_pend += 1
            elif status == "SEM CPC":
                tag = "sem_cpc"
            elif status == "SÁBADO":
                tag = "sabado"
            elif status == "DOMINGO":
                tag = "domingo"
            elif status.startswith("FERIADO"):
                tag = "feriado"
            else:
                tag = "sem_cpc"

            if data_item.date() == data_ligacao.date():
                status = "⬇️ PRÓXIMO DOWNLOAD ⬇️"
                tag = "proximo"

            pasta_path = item.get("pasta_path", "")
            caminho_completo = pasta_path if pasta_path else "---"

            self.tree.insert("", "end", values=(
                item["data"], item["dia"], str(qtd),
                status, item["download_em"], caminho_completo,
            ), tags=(tag,))

        self.c_pend.config(text=str(dias_pend))
        self.lbl_hora.config(text=f"Atualizado: {hoje.strftime('%H:%M:%S')}")

        self.root.after(30000, self.atualizar)

    def _abrir_pasta_selecionada(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        valores = self.tree.item(sel[0], "values")
        pasta = valores[5] if len(valores) > 5 else ""
        if pasta and pasta != "---" and os.path.exists(pasta):
            os.startfile(pasta)
        else:
            messagebox.showinfo("Pasta", f"Pasta não encontrada:\n{pasta}")

    def exportar(self):
        nome = f"Relatorio_ATHOS_{datetime.now().strftime('%d-%m-%Y_%H%M')}.txt"
        caminho = os.path.join(os.path.expanduser("~"), "Desktop", nome)

        with open(caminho, "w", encoding="utf-8") as f:
            f.write("ATHOS - RELATÓRIO DE LIGAÇÕES DO 3C PLUS\n")
            f.write("Campanha: ASSINATURA\n")
            f.write("=" * 80 + "\n\n")
            f.write(self.lbl_banner.cget("text") + "\n\n")
            f.write(f"{'DATA':<14} {'DIA':<10} {'ÁUDIOS':<8} {'STATUS':<20} "
                    f"{'DOWNLOAD EM':<14} PASTA\n")
            f.write("-" * 80 + "\n")
            for item in self.tree.get_children():
                v = self.tree.item(item, "values")
                f.write(f"{v[0]:<14} {v[1]:<10} {v[2]:<8} {v[3]:<20} {v[4]:<14} {v[5]}\n")

        os.startfile(caminho)


# ================================================================
# MAIN
# ================================================================
if __name__ == "__main__":
    # Instalar Pillow se não tiver
    try:
        from PIL import Image, ImageTk
    except ImportError:
        import subprocess
        print("Instalando Pillow...")
        subprocess.run(["pip", "install", "Pillow"])
        from PIL import Image, ImageTk
    
    if not os.path.exists(HISTORY_FILE):
        historico_inicial = {
            "baixados": [
                "09/04/2026", "10/04/2026", "13/04/2026", "14/04/2026", "15/04/2026",
                "16/04/2026", "17/04/2026", "20/04/2026", "22/04/2026", "23/04/2026",
                "24/04/2026", "27/04/2026", "28/04/2026", "29/04/2026", "30/04/2026"
            ]
        }
        salvar_historico_downloader(historico_inicial)
    
    app = PainelHistorico()
    app.root.mainloop()
