"""
Verificador de Downloads ATHOS
Mostra status de cada dia: baixado ou pendente
"""

import os
from datetime import datetime, timedelta
from pathlib import Path

PASTA_BASE = r"C:\Users\IAF\OneDrive\Ligações Monitoria\2026\Assinatura\GERAL"

MESES = {
    "01": "JANEIRO 01", "02": "FEVEREIRO 02", "03": "MARÇO 03",
    "04": "ABRIL 04", "05": "MAIO 05", "06": "JUNHO 06",
    "07": "JULHO 07", "08": "AGOSTO 08", "09": "SETEMBRO 09",
    "10": "OUTUBRO 10", "11": "NOVEMBRO 11", "12": "DEZEMBRO 12",
}

def verificar_downloads():
    print("=" * 70)
    print("VERIFICAÇÃO DE DOWNLOADS ATHOS")
    print("=" * 70)
    print()
    
    if not os.path.exists(PASTA_BASE):
        print(f"❌ Pasta base não encontrada: {PASTA_BASE}")
        return
    
    total_geral = 0
    dias_com_audio = 0
    dias_sem_audio = 0
    
    # Verificar últimos 30 dias
    print("ÚLTIMOS 30 DIAS:")
    print("-" * 70)
    print(f"{'DATA':<15} {'DIA':<8} {'ÁUDIOS':<10} {'STATUS':<20} {'PASTA'}")
    print("-" * 70)
    
    for dias_atras in range(30):
        data = datetime.now() - timedelta(days=dias_atras)
        mes = data.strftime("%m")
        dia = data.strftime("%d").zfill(3)
        nome_mes = MESES.get(mes, "")
        data_fmt = data.strftime("%d/%m/%Y")
        
        if nome_mes:
            pasta = os.path.join(PASTA_BASE, nome_mes, dia)
            
            if os.path.exists(pasta):
                audios = list(Path(pasta).glob("*.mp3"))
                qtd = len(audios)
                total_geral += qtd
                
                if qtd > 0:
                    status = "✅ BAIXADO"
                    dias_com_audio += 1
                else:
                    status = "📂 PASTA VAZIA"
                    dias_sem_audio += 1
                
                print(f"{data_fmt:<15} {dia:<8} {qtd:<10} {status:<20} {nome_mes}\\{dia}")
            else:
                # Verificar se é dia útil (seg-sex)
                if data.weekday() < 5:  # 0=seg, 6=dom
                    status = "❌ NÃO BAIXADO"
                else:
                    status = "⚪ FIM DE SEMANA"
                
                print(f"{data_fmt:<15} {dia:<8} {'0':<10} {status:<20} ---")
                dias_sem_audio += 1
    
    print("-" * 70)
    print()
    
    # Resumo
    print("=" * 70)
    print("RESUMO")
    print("=" * 70)
    print(f"Total de áudios: {total_geral}")
    print(f"Dias com áudios: {dias_com_audio}")
    print(f"Dias sem áudios: {dias_sem_audio}")
    print()
    
    # Próximos dias que precisam ser baixados
    print("DIAS QUE PRECISAM DE DOWNLOAD:")
    print("-" * 50)
    
    hoje = datetime.now()
    for dias_atras in range(1, 8):  # Últimos 7 dias
        data = hoje - timedelta(days=dias_atras)
        mes = data.strftime("%m")
        dia = data.strftime("%d").zfill(3)
        nome_mes = MESES.get(mes, "")
        
        if nome_mes and data.weekday() < 5:  # Dias úteis
            pasta = os.path.join(PASTA_BASE, nome_mes, dia)
            
            if not os.path.exists(pasta) or len(list(Path(pasta).glob("*.mp3"))) == 0:
                print(f"  ⚠ {data.strftime('%d/%m/%Y')} - {nome_mes}\\{dia} - PENDENTE")
    
    print()
    print("=" * 70)
    
    # Próxima execução
    print("PRÓXIMA EXECUÇÃO AGENDADA:")
    print(f"  ⏰ Amanhã às 08:00")
    print(f"  📅 Baixará áudios do dia: {hoje.strftime('%d/%m/%Y')}")
    print("=" * 70)

if __name__ == "__main__":
    verificar_downloads()
    input("\nPressione Enter para sair...")