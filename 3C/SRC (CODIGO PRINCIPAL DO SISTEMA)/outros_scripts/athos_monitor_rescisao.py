"""
ATHOS - RELATÓRIO DE LIGAÇÕES - OLOS
Campanha: Manual_QuintoCred (Rescisão)
Monitoramento de downloads por faixa (Info. Adicionais)
"""

import os
import json
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

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
# CONFIGURAÇÕES
# ================================================================
CAMPANHAS_RESCISAO = {
    "91-360": {
        "nome": "91-360",
        "faixas": ["101081", "491120", "5121180"],
        "pasta_base": r"C:\Users\IAF\OneDrive\Ligações Monitoria\2026\Rescisão\91-360"
    },
    "360+": {
        "nome": "360+",
        "faixas": ["6181360", "7361540", "8541720", "97211080"],
        "pasta_base": r"C:\Users\IAF\OneDrive\Ligações Monitoria\2026\Rescisão\360+"
    }
}

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

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

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
    """Formata o nome da pasta do dia:
    Dias 1-9: 2 dígitos (04, 05, 06)
    Dias 10-31: 3 dígitos (010, 011, 012)"""
    if 1 <= dia <= 9:
        return f"{dia:02d}"
    else:
        return f"{dia:03d}"

# ================================================================
# FUNÇÕES DO MONITOR
# ================================================================
def contar_mp3(pasta):
    if not pasta or not os.path.exists(pasta):
        return 0
    return len(list(Path(pasta).glob("*.mp3")))

def obter_audios_por_faixa(data):
    resultado = {}
    
    for campanha, config in CAMPANHAS_RESCISAO.items():
        for faixa in config["faixas"]:
            mes_nome = MESES_PASTA.get(data.strftime("%m"), "OUTROS")
            dia_pasta = formatar_pasta_dia(data.day)
            pasta = os.path.join(config["pasta_base"], mes_nome, dia_pasta)
            
            qtd = contar_mp3(pasta)
            resultado[faixa] = {
                "quantidade": qtd,
                "campanha": campanha,
                "pasta": pasta
            }
    
    return resultado

def obter_historico(dias=60):
    historico = []
    hoje = datetime.now().date()

    for i in range(dias):
        data = datetime.now() - timedelta(days=i)
        data_dt = data.date()
        data_str = data.strftime("%d/%m/%Y")
        dia_semana = nome_dia(data)

        if not eh_dia_util(data):
            if data.weekday() == 5:
                tipo = "SÁBADO"
            elif data.weekday() == 6:
                tipo = "DOMINGO"
            else:
                nome = FERIADOS_2026.get(data.strftime("%d/%m/%Y"), "")
                tipo = f"FERIADO — {nome}" if nome else "FERIADO"

            historico.append({
                "data": data_str, "dia": dia_semana,
                "status": tipo, "download_em": "---",
                "faixas": {}
            })
            continue

        faixas = obter_audios_por_faixa(data)
        total_audios = sum(f["quantidade"] for f in faixas.values())
        
        if total_audios > 0:
            status = "BAIXADO"
            download_em = data_str
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
            "status": status,
            "download_em": download_em,
            "faixas": faixas
        })

    return historico

def encontrar_proximo_download_real(hoje):
    if eh_dia_util(hoje):
        if hoje.hour < 8 or (hoje.hour == 8 and hoje.minute == 0):
            return hoje, dia_util_anterior(hoje)
        else:
            prox = proximo_dia_util(hoje)
            return prox, dia_util_anterior(prox)
    else:
        prox = proximo_dia_util(hoje)
        return prox, dia_util_anterior(prox)

# ================================================================
# UI PRINCIPAL
# ================================================================
class PainelRescisao:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ATHOS - Relatório de Ligações - OLOS - Campanha: Manual_QuintoCred (Rescisão)")
        self.root.geometry("1400x750")
        self.root.configure(bg=ATHOS_BG)
        self.root.resizable(True, True)

        self._build_ui()
        self._centralizar()
        self.atualizar()

    def _centralizar(self):
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"1400x750+{(sw-1400)//2}+{(sh-750)//2}")

    def _build_ui(self):
        header = tk.Frame(self.root, bg=ATHOS_BG2, height=90)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        tk.Label(header, text="ATHOS - RELATÓRIO DE LIGAÇÕES - OLOS",
                 font=("Consolas", 18, "bold"), fg=BRANCO, bg=ATHOS_BG2).pack(pady=(12,2))
        tk.Label(header, text="Campanha: Manual_QuintoCred (Rescisão)",
                 font=("Consolas", 11), fg=AZUL_CLARO, bg=ATHOS_BG2).pack()
        tk.Label(header, text="Monitoramento de downloads por faixa (Info. Adicionais)",
                 font=("Consolas", 9), fg=CINZA, bg=ATHOS_BG2).pack(pady=(2,5))

        banner = tk.Frame(self.root, bg=ATHOS_BG3,
                          highlightbackground=ATHOS_AZUL, highlightthickness=2)
        banner.pack(fill="x", padx=15, pady=(10,0))
        self.lbl_banner = tk.Label(banner, text="Carregando...",
                                   font=("Consolas", 12, "bold"),
                                   fg=LARANJA, bg=ATHOS_BG3, justify="center")
        self.lbl_banner.pack(pady=10)

        cards = tk.Frame(self.root, bg=ATHOS_BG)
        cards.pack(fill="x", padx=15, pady=8)
        self.c_prox = self._card(cards, "PRÓXIMO DOWNLOAD", "---", AZUL_CLARO)
        self.c_prox_lig = self._card(cards, "LIGAÇÃO DE", "---", LARANJA)

        frame_tab = tk.Frame(self.root, bg=ATHOS_BG3)
        frame_tab.pack(fill="both", expand=True, padx=15, pady=(0,4))

        scroll_y = tk.Scrollbar(frame_tab, orient="vertical")
        scroll_y.pack(side="right", fill="y")
        scroll_x = tk.Scrollbar(frame_tab, orient="horizontal")
        scroll_x.pack(side="bottom", fill="x")

        cols = ["data", "dia", "status", "download_em"]
        todas_faixas = ["101081", "491120", "5121180", "6181360", "7361540", "8541720", "97211080"]

        for faixa in todas_faixas:
            cols.append(f"faixa_{faixa}")

        self.tree = ttk.Treeview(frame_tab, columns=cols, show="headings",
                                 height=18,
                                 yscrollcommand=scroll_y.set,
                                 xscrollcommand=scroll_x.set)

        self.tree.heading("data", text="DATA LIGAÇÃO")
        self.tree.heading("dia", text="DIA")
        self.tree.heading("status", text="STATUS")
        self.tree.heading("download_em", text="DOWNLOAD EM")
        
        for faixa in todas_faixas:
            self.tree.heading(f"faixa_{faixa}", text=f"{faixa}")

        self.tree.column("data", width=100, anchor="center")
        self.tree.column("dia", width=80, anchor="center")
        self.tree.column("status", width=180, anchor="w")
        self.tree.column("download_em", width=100, anchor="center")
        
        for faixa in todas_faixas:
            self.tree.column(f"faixa_{faixa}", width=70, anchor="center")

        estilo = ttk.Style()
        estilo.theme_use("default")
        estilo.configure("Treeview",
                         background=ATHOS_BG2, foreground=BRANCO,
                         fieldbackground=ATHOS_BG2,
                         font=("Consolas", 9), rowheight=24)
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
        self.tree.tag_configure("proximo", foreground=LARANJA, font=("Consolas", 9, "bold"))

        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        self.tree.pack(fill="both", expand=True)

        rodape = tk.Frame(self.root, bg=ATHOS_BG)
        rodape.pack(fill="x", padx=15, pady=6)

        tk.Button(rodape, text="📂 ABRIR PASTA 91-360", bg=ATHOS_BG3, fg=BRANCO,
                  font=("Consolas", 10, "bold"),
                  command=lambda: self._abrir_pasta(CAMPANHAS_RESCISAO["91-360"]["pasta_base"]),
                  padx=12, pady=5, relief="flat", cursor="hand2").pack(side="left", padx=3)

        tk.Button(rodape, text="📂 ABRIR PASTA 360+", bg=ATHOS_BG3, fg=BRANCO,
                  font=("Consolas", 10, "bold"),
                  command=lambda: self._abrir_pasta(CAMPANHAS_RESCISAO["360+"]["pasta_base"]),
                  padx=12, pady=5, relief="flat", cursor="hand2").pack(side="left", padx=3)

        tk.Button(rodape, text="🔄 ATUALIZAR", bg=ATHOS_AZUL, fg=BRANCO,
                  font=("Consolas", 10, "bold"),
                  command=self.atualizar,
                  padx=12, pady=5, relief="flat", cursor="hand2").pack(side="left", padx=3)

        tk.Button(rodape, text="💾 EXPORTAR TXT", bg=ATHOS_BG3, fg=BRANCO,
                  font=("Consolas", 10),
                  command=self.exportar,
                  padx=12, pady=5, relief="flat", cursor="hand2").pack(side="left", padx=3)

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

    def _abrir_pasta(self, pasta):
        if os.path.exists(pasta):
            os.startfile(pasta)
        else:
            messagebox.showinfo("Pasta", f"Pasta não encontrada:\n{pasta}")

    def atualizar(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        historico = obter_historico(60)
        hoje = datetime.now()
        ordem_faixas = ["101081", "491120", "5121180", "6181360", "7361540", "8541720", "97211080"]

        data_download, data_ligacao = encontrar_proximo_download_real(hoje)
        
        self.lbl_banner.config(
            text=f"🎯 PRÓXIMO DOWNLOAD REAL: {data_download.strftime('%d/%m/%Y')} ({nome_dia(data_download)}) às 08:00\n"
                 f"📞 Ligação de: {data_ligacao.strftime('%d/%m/%Y')} ({nome_dia(data_ligacao)})"
        )
        self.c_prox.config(text=data_download.strftime("%d/%m/%Y"))
        self.c_prox_lig.config(text=data_ligacao.strftime("%d/%m/%Y"))

        for item in historico:
            status = item["status"]
            faixas = item.get("faixas", {})

            valores = [item["data"], item["dia"], status, item["download_em"]]
            
            for faixa in ordem_faixas:
                qtd = faixas.get(faixa, {}).get("quantidade", 0)
                valores.append(qtd)

            if status == "BAIXADO":
                tag_linha = "baixado"
            elif status == "PENDENTE":
                tag_linha = "pendente"
            elif status == "SEM CPC":
                tag_linha = "sem_cpc"
            elif status == "SÁBADO":
                tag_linha = "sabado"
            elif status == "DOMINGO":
                tag_linha = "domingo"
            elif status.startswith("FERIADO"):
                tag_linha = "feriado"
            else:
                tag_linha = "sem_cpc"

            data_item = datetime.strptime(item["data"], "%d/%m/%Y")
            if data_item.date() == data_ligacao.date():
                valores[2] = "⬇️ PRÓXIMO ⬇️"
                tag_linha = "proximo"

            self.tree.insert("", "end", values=valores, tags=(tag_linha,))

        self.lbl_hora.config(text=f"Atualizado: {hoje.strftime('%H:%M:%S')}")
        self.root.after(30000, self.atualizar)

    def exportar(self):
        nome = f"Relatorio_Rescisao_{datetime.now().strftime('%d-%m-%Y_%H%M')}.txt"
        caminho = os.path.join(os.path.expanduser("~"), "Desktop", nome)

        with open(caminho, "w", encoding="utf-8") as f:
            f.write("ATHOS - RELATÓRIO DE LIGAÇÕES - OLOS\n")
            f.write("Campanha: Manual_QuintoCred (Rescisão)\n")
            f.write("=" * 100 + "\n\n")
            f.write(self.lbl_banner.cget("text") + "\n\n")
            
            f.write(f"{'DATA':<12} {'DIA':<10} {'STATUS':<22} {'DOWNLOAD':<12} ")
            for faixa in ["101081", "491120", "5121180", "6181360", "7361540", "8541720", "97211080"]:
                f.write(f"{faixa:<8} ")
            f.write("\n")
            f.write("-" * 140 + "\n")
            
            for item in self.tree.get_children():
                v = self.tree.item(item, "values")
                f.write(f"{v[0]:<12} {v[1]:<10} {str(v[2])[:20]:<22} {v[3]:<12} ")
                for i in range(4, len(v)):
                    f.write(f"{str(v[i]):<8} ")
                f.write("\n")

        os.startfile(caminho)


# ================================================================
# MAIN
# ================================================================
if __name__ == "__main__":
    app = PainelRescisao()
    app.root.mainloop()
