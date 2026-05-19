"""
Mostra campo 'list' de cada registro CPC do dia 27/04
Rode: python debug_api.py
"""
import json, requests, os

TOKEN_FILE = os.path.join(os.path.expanduser("~"), ".athos_painel_token.json")
BASE_URL   = "https://athossoluoes.3c.plus/api/v1"

def load_token():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE) as f:
            return json.load(f).get("token")
    return None

token = load_token()
if not token:
    token = input("Cole seu Bearer token: ").strip()

headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

todas   = []
ids_vis = set()
for page in range(0, 20):
    url = (
        f"{BASE_URL}/calls?page={page}"
        f"&start_date=2026-04-27+07:00:00"
        f"&end_date=2026-04-27+20:00:00"
        f"&call_mode=all&statuses=7&type=all"
        f"&campaigns[]=253640"
        f"&qualification=204212,169571,169563,169562,169560,169561,169558,169559,169540,169557,169556,166376,166375"
        f"&simple_paginate=true&order_by_desc=call_date&include=campaign_rel"
    )
    r    = requests.get(url, headers=headers, timeout=30)
    data = r.json().get("data", [])
    if not data:
        break
    for c in data:
        cid = c.get("id","")
        if cid not in ids_vis:
            ids_vis.add(cid)
            if str(c.get("recording") or "").startswith("http"):
                if str(c.get("qualification") or "").upper().startswith("CPC"):
                    todas.append(c)

print(f"Total CPC únicos: {len(todas)}")
print()

# Mostra campo list de cada um
from collections import Counter
listas = Counter()
for i, c in enumerate(todas):
    lista = str(c.get("list",""))
    qual  = c.get("qualification","")
    listas[lista] += 1
    print(f"[{i+1:02d}] list='{lista}' | {qual}")

print()
print("Distribuição de listas:")
for l, n in listas.most_common():
    print(f"  {n}x '{l}'")

input("\nPressione ENTER para fechar...")
