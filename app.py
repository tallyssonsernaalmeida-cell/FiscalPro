from flask import (Flask, render_template, request, redirect,
                   url_for, session, send_file, flash)
import pandas as pd
import os, json, io
from datetime import datetime, timedelta
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import ofxparse
from openpyxl import Workbook
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = 'fiscalpro_secret_2025_xK9#mL'
app.jinja_env.filters['enumerate'] = enumerate

# Admin: sessão longa mas SEMPRE exige login (sem remember-me)
# Cookie zerado = tem que logar de novo — comportamento padrão do Flask
app.permanent_session_lifetime = timedelta(hours=12)

USERS_FILE   = 'users.json'
SUPPORT_FILE = 'support.json'
LOGS_FILE    = 'logs.json'
UPLOAD_FOLDER= 'uploads'

for d in [UPLOAD_FOLDER, 'user_data']:
    os.makedirs(d, exist_ok=True)

ALLOWED = {'ofx','xlsx','xls','xlsm','csv','pdf'}

# ─────────────────────────────────────────────────────────────
# LOGS
# ─────────────────────────────────────────────────────────────
def add_log(acao, usuario='sistema', detalhe=''):
    logs = []
    if os.path.exists(LOGS_FILE):
        try:
            with open(LOGS_FILE, 'r', encoding='utf-8') as f:
                logs = json.load(f)
        except:
            logs = []
    logs.insert(0, {
        'data':    datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
        'acao':    acao,
        'usuario': usuario,
        'detalhe': detalhe
    })
    logs = logs[:500]  # mantém últimos 500
    with open(LOGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(logs, f, ensure_ascii=False, indent=2)

def get_logs():
    if not os.path.exists(LOGS_FILE): return []
    try:
        with open(LOGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return []

# ─────────────────────────────────────────────────────────────
# USUÁRIOS
# ─────────────────────────────────────────────────────────────
def load_users():
    if not os.path.exists(USERS_FILE):
        default = {"admin@fiscal.app": {
            "password": generate_password_hash("admin123"),
            "nome": "Administrador", "role": "admin",
            "files": [], "whatsapp": "",
            "created_at": datetime.now().isoformat(),
            "ativo": True
        }}
        save_users(default)
        return default
    with open(USERS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_users(u):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(u, f, ensure_ascii=False, indent=4)

def load_support():
    if not os.path.exists(SUPPORT_FILE): return []
    try:
        with open(SUPPORT_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return []

def save_support(m):
    with open(SUPPORT_FILE, 'w', encoding='utf-8') as f:
        json.dump(m, f, ensure_ascii=False, indent=4)

def get_ctx():
    """Contexto base para todos os templates."""
    users = load_users()
    user  = users.get(session.get('user'), {})
    return {
        'nome':     user.get('nome', ''),
        'email':    session.get('user', ''),
        'is_admin': user.get('role') == 'admin'
    }

def require_login():
    if 'user' not in session:
        return redirect(url_for('login'))

def require_admin():
    if 'user' not in session:
        return redirect(url_for('login'))
    users = load_users()
    if users.get(session['user'], {}).get('role') != 'admin':
        return redirect(url_for('dashboard'))

# ─────────────────────────────────────────────────────────────
# CATEGORIZAÇÃO
# ─────────────────────────────────────────────────────────────
def categorizar(desc):
    d = str(desc).lower()
    if any(x in d for x in ['salário','salario','holerite','folha']):            return 'Salário/Renda'
    if any(x in d for x in ['freelance','comissão','comissao','honorário']):     return 'Renda Extra'
    if any(x in d for x in ['pix recebid','ted recebid','transferência recebida','depósito','deposito']): return 'Transferência Recebida'
    if any(x in d for x in ['pix enviad','ted enviad','transferência enviada']): return 'Transferência Enviada'
    if any(x in d for x in ['aluguel','condomínio','condominio','iptu','financiamento']): return 'Moradia'
    if any(x in d for x in ['supermercado','mercado','hortifruti','açougue','padaria','restaurante','ifood','lanchonete']): return 'Alimentação'
    if any(x in d for x in ['uber','99 ','cabify','ônibus','metrô','gasolina','combustível','posto','estacionamento']): return 'Transporte'
    if any(x in d for x in ['energia','água ','agua ','internet','telefone','celular','claro','vivo','tim ','net ','gás']): return 'Contas & Serviços'
    if any(x in d for x in ['farmácia','farmacia','drogaria','médico','medico','hospital','plano de saúde','unimed']): return 'Saúde'
    if any(x in d for x in ['netflix','spotify','amazon','disney','hbo','cinema','teatro','ingresso']): return 'Lazer'
    if any(x in d for x in ['curso','faculdade','escola','educação','livro','livraria']): return 'Educação'
    if any(x in d for x in ['roupa','calçado','moda','loja','zara','renner','riachuelo']): return 'Vestuário'
    return 'Outros'

# ─────────────────────────────────────────────────────────────
# PROCESSAMENTO DE ARQUIVOS
# ─────────────────────────────────────────────────────────────
def process_ofx(fp):
    try:
        with open(fp, 'r', encoding='utf-8', errors='ignore') as f:
            ofx = ofxparse.OfxParser.parse(f)
        out = []
        for acc in ofx.accounts:
            for t in acc.statement.transactions:
                v = float(t.amount)
                d = (t.memo or '')[:120]
                out.append({'data': str(t.date)[:10], 'descricao': d, 'valor': v,
                            'tipo': 'Crédito' if v > 0 else 'Débito',
                            'categoria': categorizar(d), 'pago': False})
        return out
    except Exception as e:
        print(f'OFX: {e}'); return []

def process_excel(fp):
    try:
        ext = fp.rsplit('.', 1)[-1].lower()
        if ext == 'csv':
            df = pd.read_csv(fp, sep=None, engine='python', encoding='utf-8', errors='ignore')
        elif ext in ['xlsx','xlsm']:
            df = pd.read_excel(fp, engine='openpyxl')
        else:
            df = pd.read_excel(fp, engine='xlrd')

        df.columns = [str(c).strip().lower() for c in df.columns]
        col_d = next((c for c in df.columns if any(k in c for k in ['data','date','dt'])), None)
        col_v = next((c for c in df.columns if any(k in c for k in ['valor','value','amount','montante'])), None)
        col_t = next((c for c in df.columns if any(k in c for k in ['descri','memo','histor','lançamento','lancamento','complement'])), None)

        out = []
        for _, row in df.iterrows():
            v = 0.0
            if col_v:
                try:
                    v = float(str(row[col_v]).replace('R$','').replace(' ','').replace('.','').replace(',','.'))
                except: v = 0.0
            desc = str(row[col_t])[:120] if col_t else 'Sem descrição'
            data = str(row[col_d])[:10]  if col_d  else 'N/A'
            out.append({'data': data, 'descricao': desc, 'valor': v,
                        'tipo': 'Crédito' if v > 0 else 'Débito',
                        'categoria': categorizar(desc), 'pago': False})
        return out
    except Exception as e:
        print(f'Excel/CSV: {e}'); return []

# ─────────────────────────────────────────────────────────────
# DASHBOARD DATA
# ─────────────────────────────────────────────────────────────
def get_dash(email):
    users = load_users()
    files = users.get(email, {}).get('files', [])
    txs   = []
    for fi in files:
        if not os.path.exists(fi['path']): continue
        ext = fi['path'].rsplit('.', 1)[-1].lower()
        txs.extend(process_ofx(fi['path']) if ext == 'ofx' else process_excel(fi['path']))

    rec  = sum(t['valor'] for t in txs if t['valor'] > 0)
    desp = sum(abs(t['valor']) for t in txs if t['valor'] < 0)
    saldo= rec - desp

    cat_d, cat_r = {}, {}
    for t in txs:
        cat = t['categoria']
        if t['valor'] < 0: cat_d[cat] = cat_d.get(cat, 0) + abs(t['valor'])
        else:              cat_r[cat] = cat_r.get(cat, 0) + t['valor']

    return {
        'transactions':    txs,
        'total_receitas':  round(rec, 2),
        'total_despesas':  round(desp, 2),
        'saldo':           round(saldo, 2),
        'status':          'Lucro ✅' if saldo > 0 else ('Prejuízo ⚠️' if saldo < 0 else 'Equilibrado ⚖️'),
        'qtd_transacoes':  len(txs),
        'categorias':      cat_d,
        'categorias_rec':  cat_r,
        'top_categorias':  sorted(cat_d.items(), key=lambda x: x[1], reverse=True)[:5],
        'percentual_lucro':round((saldo/rec)*100, 1) if rec > 0 else 0
    }

# ─────────────────────────────────────────────────────────────
# EXCEL COMPLETO
# ─────────────────────────────────────────────────────────────
def _brd():
    s = Side(style='thin', color='D0D0D0')
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(cor): return PatternFill(start_color=cor, end_color=cor, fill_type='solid')

def _fnt(cor='1E293B', bold=False, size=10):
    return Font(name='Calibri', color=cor, bold=bold, size=size)

def build_excel(data):
    G, GL, R, RL = '22a060','E8F8EF','C0392B','FDEDEC'
    DARK, GRAY   = '1E293B','F4F6F8'
    wb  = Workbook()
    brd = _brd()

    rec_list  = [t for t in data['transactions'] if t['valor'] > 0]
    desp_list = [t for t in data['transactions'] if t['valor'] < 0]
    tot_r = data['total_receitas']
    tot_d = data['total_despesas']
    saldo = data['saldo']
    pct   = data['percentual_lucro']
    cats  = sorted(data['categorias'].items(), key=lambda x: x[1], reverse=True)
    cats_r= sorted(data['categorias_rec'].items(), key=lambda x: x[1], reverse=True)

    # ══════════════════════════════════════
    # ABA 1 — PAINEL RESUMO
    # ══════════════════════════════════════
    ws = wb.active
    ws.title = '📊 Painel Resumo'
    ws.sheet_view.showGridLines = False

    # Título
    ws.merge_cells('A1:I1')
    ws['A1'].value = '🌿  FISCALPRO — RELATÓRIO FINANCEIRO COMPLETO'
    ws['A1'].font  = Font(name='Calibri', bold=True, size=18, color='FFFFFF')
    ws['A1'].fill  = _fill(G)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 42

    ws.merge_cells('A2:I2')
    ws['A2'].value = f'Gerado em {datetime.now().strftime("%d/%m/%Y às %H:%M")}   •   {data["qtd_transacoes"]} transações'
    ws['A2'].font  = _fnt('6B7280', size=10)
    ws['A2'].fill  = _fill(GRAY)
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    # KPIs linha 4-6
    kpis = [
        ('B', 'RECEITAS TOTAIS',  f'R$ {tot_r:,.2f}',  G,      GL),
        ('D', 'DESPESAS TOTAIS',  f'R$ {tot_d:,.2f}',  R,      RL),
        ('F', 'SALDO LÍQUIDO',    f'R$ {saldo:,.2f}',  G if saldo>=0 else R, GL if saldo>=0 else RL),
        ('H', 'MARGEM DE LUCRO',  f'{pct}%',           '1D4ED8','EFF6FF'),
    ]
    ws.row_dimensions[4].height = 16
    ws.row_dimensions[5].height = 38
    ws.row_dimensions[6].height = 14
    for col, label, val, clr, bg in kpis:
        for rr in [4,5,6]:
            ws[f'{col}{rr}'].fill = _fill(bg)
        ws[f'{col}4'].value = label
        ws[f'{col}4'].font  = Font(name='Calibri', bold=True, size=9, color='6B7280')
        ws[f'{col}4'].alignment = Alignment(horizontal='center', vertical='center')
        ws[f'{col}5'].value = val
        ws[f'{col}5'].font  = Font(name='Calibri', bold=True, size=16, color=clr)
        ws[f'{col}5'].alignment = Alignment(horizontal='center', vertical='center')

    for col in 'ABCDEFGHI':
        ws.column_dimensions[col].width = 18

    # Tabela despesas por cat
    r = 9
    ws.merge_cells(f'A{r}:D{r}')
    ws[f'A{r}'].value = 'DESPESAS POR CATEGORIA'
    ws[f'A{r}'].font  = Font(name='Calibri', bold=True, size=13, color=G)
    ws.row_dimensions[r].height = 22
    r += 1

    for i, h in enumerate(['Categoria','Total (R$)','% do Total','Qtd'], 1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(G)
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = brd
    ws.row_dimensions[r].height = 18
    cat_start = r + 1; r += 1

    for i, (cat, val) in enumerate(cats):
        pct_c = round(val/tot_d*100,1) if tot_d > 0 else 0
        qtd   = sum(1 for t in desp_list if t['categoria']==cat)
        bg    = GL if i%2==0 else 'FFFFFF'
        for ci, v in enumerate([cat, round(val,2), f'{pct_c}%', qtd], 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.fill = _fill(bg); c.border = brd
            c.font = _fnt(R, True) if ci==2 else _fnt()
            c.alignment = Alignment(horizontal='left' if ci==1 else 'center', vertical='center')
        r += 1
    cat_end = r - 1

    # Tabela receitas por cat
    r += 1
    ws.merge_cells(f'A{r}:D{r}')
    ws[f'A{r}'].value = 'RECEITAS POR CATEGORIA'
    ws[f'A{r}'].font  = Font(name='Calibri', bold=True, size=13, color=G)
    ws.row_dimensions[r].height = 22
    r += 1

    for i, h in enumerate(['Categoria','Total (R$)','% do Total','Qtd'], 1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(G)
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = brd
    r += 1

    for i, (cat, val) in enumerate(cats_r):
        pct_c = round(val/tot_r*100,1) if tot_r > 0 else 0
        qtd   = sum(1 for t in rec_list if t['categoria']==cat)
        bg    = GL if i%2==0 else 'FFFFFF'
        for ci, v in enumerate([cat, round(val,2), f'{pct_c}%', qtd], 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.fill = _fill(bg); c.border = brd
            c.font = _fnt(G, True) if ci==2 else _fnt()
            c.alignment = Alignment(horizontal='left' if ci==1 else 'center', vertical='center')
        r += 1

    # Gráfico pizza — despesas
    if cats:
        pie = PieChart()
        pie.title  = 'Despesas por Categoria'
        pie.style  = 10; pie.width = 14; pie.height = 12
        pie.add_data(Reference(ws, min_col=2, min_row=cat_start-1, max_row=cat_end))
        pie.set_categories(Reference(ws, min_col=1, min_row=cat_start, max_row=cat_end))
        ws.add_chart(pie, 'F9')

    # Gráfico barras — rec vs desp
    ws['H2'].value = 'Receitas';  ws['H3'].value = tot_r
    ws['I2'].value = 'Despesas';  ws['I3'].value = tot_d
    bar = BarChart()
    bar.type = 'col'; bar.title = 'Receitas vs Despesas'
    bar.style = 10; bar.width = 12; bar.height = 10
    bar.add_data(Reference(ws, min_col=8, max_col=9, min_row=2, max_row=3), titles_from_data=True)
    ws.add_chart(bar, 'F24')

    # ══════════════════════════════════════
    # ABA 2 — RECEITAS
    # ══════════════════════════════════════
    w2 = wb.create_sheet('💰 Receitas')
    w2.sheet_view.showGridLines = False

    w2.merge_cells('A1:F1')
    w2['A1'].value = '💰  RECEITAS DETALHADAS'
    w2['A1'].font  = Font(name='Calibri', bold=True, size=15, color='FFFFFF')
    w2['A1'].fill  = _fill(G)
    w2['A1'].alignment = Alignment(horizontal='center', vertical='center')
    w2.row_dimensions[1].height = 32

    w2.merge_cells('A2:F2')
    w2['A2'].value = f'Total: R$ {tot_r:,.2f}   |   {len(rec_list)} lançamentos'
    w2['A2'].font  = Font(name='Calibri', bold=True, size=11, color=G)
    w2['A2'].fill  = _fill(GL)
    w2['A2'].alignment = Alignment(horizontal='center', vertical='center')

    for i, h in enumerate(['Data','Descrição','Categoria','Valor (R$)','% da Receita','Status'], 1):
        c = w2.cell(row=3, column=i, value=h)
        c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(G)
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = brd

    for i, t in enumerate(sorted(rec_list, key=lambda x: x['data'], reverse=True)):
        rr  = i + 4
        pct_t = round(t['valor']/tot_r*100,1) if tot_r > 0 else 0
        bg  = GL if i%2==0 else 'FFFFFF'
        for ci, v in enumerate([t['data'], t['descricao'], t['categoria'],
                                  round(t['valor'],2), f'{pct_t}%',
                                  '✅ Pago' if t.get('pago') else '⚠️ Pendente'], 1):
            c = w2.cell(row=rr, column=ci, value=v)
            c.fill = _fill(bg); c.border = brd
            c.font = _fnt(G, True) if ci==4 else _fnt()
            c.alignment = Alignment(horizontal='left' if ci==2 else 'center', vertical='center')

    lr = len(rec_list) + 4
    w2.merge_cells(f'A{lr}:C{lr}')
    w2[f'A{lr}'].value = 'TOTAL RECEITAS'
    w2[f'A{lr}'].font  = _fnt('FFFFFF', True, 11); w2[f'A{lr}'].fill = _fill(G)
    w2[f'A{lr}'].alignment = Alignment(horizontal='center', vertical='center')
    w2[f'D{lr}'].value = round(tot_r, 2)
    w2[f'D{lr}'].font  = _fnt('FFFFFF', True, 11); w2[f'D{lr}'].fill = _fill(G)
    w2[f'D{lr}'].alignment = Alignment(horizontal='center', vertical='center')

    for col, w in zip('ABCDEF', [13,50,22,14,14,14]):
        w2.column_dimensions[col].width = w

    # ══════════════════════════════════════
    # ABA 3 — DESPESAS
    # ══════════════════════════════════════
    w3 = wb.create_sheet('📤 Despesas')
    w3.sheet_view.showGridLines = False

    w3.merge_cells('A1:F1')
    w3['A1'].value = '📤  DESPESAS DETALHADAS'
    w3['A1'].font  = Font(name='Calibri', bold=True, size=15, color='FFFFFF')
    w3['A1'].fill  = _fill(R)
    w3['A1'].alignment = Alignment(horizontal='center', vertical='center')
    w3.row_dimensions[1].height = 32

    w3.merge_cells('A2:F2')
    w3['A2'].value = f'Total: R$ {tot_d:,.2f}   |   {len(desp_list)} lançamentos'
    w3['A2'].font  = Font(name='Calibri', bold=True, size=11, color=R)
    w3['A2'].fill  = _fill(RL)
    w3['A2'].alignment = Alignment(horizontal='center', vertical='center')

    for i, h in enumerate(['Data','Descrição','Categoria','Valor (R$)','% da Despesa','Status'], 1):
        c = w3.cell(row=3, column=i, value=h)
        c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(R)
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = brd

    for i, t in enumerate(sorted(desp_list, key=lambda x: x['data'], reverse=True)):
        rr  = i + 4
        pct_t = round(abs(t['valor'])/tot_d*100,1) if tot_d > 0 else 0
        bg  = RL if i%2==0 else 'FFFFFF'
        for ci, v in enumerate([t['data'], t['descricao'], t['categoria'],
                                  round(abs(t['valor']),2), f'{pct_t}%',
                                  '✅ Pago' if t.get('pago') else '⚠️ Pendente'], 1):
            c = w3.cell(row=rr, column=ci, value=v)
            c.fill = _fill(bg); c.border = brd
            c.font = _fnt(R, True) if ci==4 else _fnt()
            c.alignment = Alignment(horizontal='left' if ci==2 else 'center', vertical='center')

    # Gráfico pizza despesas
    if cats:
        pie2 = PieChart()
        pie2.title = 'Distribuição de Despesas'; pie2.style = 10
        pie2.width = 16; pie2.height = 12
        pie2.add_data(Reference(ws, min_col=2, min_row=cat_start-1, max_row=cat_end))
        pie2.set_categories(Reference(ws, min_col=1, min_row=cat_start, max_row=cat_end))
        w3.add_chart(pie2, 'H4')

    # Gráfico barras despesas
    bar2 = BarChart()
    bar2.type = 'bar'; bar2.title = 'Despesas por Categoria'
    bar2.style = 10; bar2.width = 16; bar2.height = 12
    bar2.add_data(Reference(ws, min_col=2, min_row=cat_start-1, max_row=cat_end), titles_from_data=True)
    bar2.set_categories(Reference(ws, min_col=1, min_row=cat_start, max_row=cat_end))
    w3.add_chart(bar2, 'H20')

    lr3 = len(desp_list) + 4
    w3.merge_cells(f'A{lr3}:C{lr3}')
    w3[f'A{lr3}'].value = 'TOTAL DESPESAS'
    w3[f'A{lr3}'].font  = _fnt('FFFFFF', True, 11); w3[f'A{lr3}'].fill = _fill(R)
    w3[f'A{lr3}'].alignment = Alignment(horizontal='center', vertical='center')
    w3[f'D{lr3}'].value = round(tot_d, 2)
    w3[f'D{lr3}'].font  = _fnt('FFFFFF', True, 11); w3[f'D{lr3}'].fill = _fill(R)
    w3[f'D{lr3}'].alignment = Alignment(horizontal='center', vertical='center')

    for col, w in zip('ABCDEF', [13,50,22,14,14,14]):
        w3.column_dimensions[col].width = w

    # ══════════════════════════════════════
    # ABA 4 — ANÁLISE %
    # ══════════════════════════════════════
    w4 = wb.create_sheet('📈 Análise')
    w4.sheet_view.showGridLines = False

    w4.merge_cells('A1:F1')
    w4['A1'].value = '📈  ANÁLISE PERCENTUAL E COMPARATIVO'
    w4['A1'].font  = Font(name='Calibri', bold=True, size=15, color='FFFFFF')
    w4['A1'].fill  = _fill(G)
    w4['A1'].alignment = Alignment(horizontal='center', vertical='center')
    w4.row_dimensions[1].height = 32

    kpi_items = [
        ('Total Receitas',     f'R$ {tot_r:,.2f}',  G),
        ('Total Despesas',     f'R$ {tot_d:,.2f}',  R),
        ('Saldo Líquido',      f'R$ {saldo:,.2f}',  G if saldo>=0 else R),
        ('Margem de Lucro',    f'{pct}%',            '1D4ED8'),
        ('Qtd Transações',     str(data['qtd_transacoes']), DARK),
        ('Maior Despesa',      f'R$ {max((abs(t["valor"]) for t in desp_list), default=0):,.2f}', R),
        ('Maior Receita',      f'R$ {max((t["valor"] for t in rec_list), default=0):,.2f}', G),
        ('Ticket Médio Desp',  f'R$ {(tot_d/len(desp_list)):,.2f}' if desp_list else 'R$ 0,00', R),
        ('Ticket Médio Rec',   f'R$ {(tot_r/len(rec_list)):,.2f}'  if rec_list  else 'R$ 0,00', G),
    ]
    for i, (label, val, clr) in enumerate(kpi_items):
        rr = i + 3
        c1 = w4.cell(row=rr, column=1, value=label)
        c1.font = _fnt(DARK, True); c1.fill = _fill(GRAY); c1.border = brd
        c1.alignment = Alignment(horizontal='left', vertical='center')
        c2 = w4.cell(row=rr, column=2, value=val)
        c2.font = Font(name='Calibri', bold=True, size=11, color=clr)
        c2.fill = _fill(GRAY); c2.border = brd
        c2.alignment = Alignment(horizontal='center', vertical='center')

    # Tabela comparativa
    r4 = 14
    w4.merge_cells(f'A{r4}:F{r4}')
    w4[f'A{r4}'].value = 'COMPARATIVO RECEITA vs DESPESA POR CATEGORIA'
    w4[f'A{r4}'].font  = Font(name='Calibri', bold=True, size=13, color=G)
    w4.row_dimensions[r4].height = 22
    r4 += 1

    for i, h in enumerate(['Categoria','Desp (R$)','% Desp','Rec (R$)','% Rec','Saldo Cat'], 1):
        c = w4.cell(row=r4, column=i, value=h)
        c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(G)
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = brd
    r4 += 1

    all_cats = sorted(set(list(data['categorias'].keys()) + list(data['categorias_rec'].keys())))
    comp_start = r4
    for i, cat in enumerate(all_cats):
        dv = data['categorias'].get(cat, 0)
        rv = data['categorias_rec'].get(cat, 0)
        dp = round(dv/tot_d*100,1) if tot_d > 0 else 0
        rp = round(rv/tot_r*100,1) if tot_r > 0 else 0
        sc = round(rv - dv, 2)
        bg = GL if i%2==0 else 'FFFFFF'
        for ci, v in enumerate([cat, round(dv,2), f'{dp}%', round(rv,2), f'{rp}%', sc], 1):
            c = w4.cell(row=r4, column=ci, value=v)
            c.fill = _fill(bg); c.border = brd
            c.font = _fnt(R, True) if ci==2 else (_fnt(G, True) if ci in [4,6] else _fnt())
            c.alignment = Alignment(horizontal='left' if ci==1 else 'center', vertical='center')
        r4 += 1

    for col, w in zip('ABCDEF', [26,15,10,15,10,14]):
        w4.column_dimensions[col].width = w

    # Gráfico comparativo por categoria
    bar3 = BarChart()
    bar3.type = 'col'; bar3.title = 'Receita vs Despesa por Categoria'
    bar3.style = 10; bar3.width = 20; bar3.height = 14
    bar3.add_data(Reference(w4, min_col=2, max_col=4, min_row=14, max_row=r4-1), titles_from_data=True)
    bar3.set_categories(Reference(w4, min_col=1, min_row=15, max_row=r4-1))
    w4.add_chart(bar3, 'H3')

    # ══════════════════════════════════════
    # ABA 5 — TODAS AS TRANSAÇÕES
    # ══════════════════════════════════════
    w5 = wb.create_sheet('📋 Transações')
    w5.sheet_view.showGridLines = False

    w5.merge_cells('A1:G1')
    w5['A1'].value = '📋  TODAS AS TRANSAÇÕES'
    w5['A1'].font  = Font(name='Calibri', bold=True, size=15, color='FFFFFF')
    w5['A1'].fill  = _fill(G)
    w5['A1'].alignment = Alignment(horizontal='center', vertical='center')
    w5.row_dimensions[1].height = 32

    for i, h in enumerate(['Data','Descrição','Categoria','Tipo','Valor (R$)','% no Fluxo','Status'], 1):
        c = w5.cell(row=2, column=i, value=h)
        c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(G)
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = brd

    for i, t in enumerate(sorted(data['transactions'], key=lambda x: x['data'], reverse=True)):
        rr   = i + 3
        base = tot_r if t['valor'] > 0 else tot_d
        pct_t= round(abs(t['valor'])/base*100,2) if base > 0 else 0
        bg   = GL if t['valor'] > 0 else RL
        for ci, v in enumerate([t['data'], t['descricao'], t['categoria'],
                                  t['tipo'], round(abs(t['valor']),2),
                                  f'{pct_t}%', '✅' if t.get('pago') else '⚠️'], 1):
            c = w5.cell(row=rr, column=ci, value=v)
            c.fill = _fill(bg); c.border = brd
            c.font = _fnt(G if t['valor']>0 else R, True) if ci==5 else _fnt()
            c.alignment = Alignment(horizontal='left' if ci==2 else 'center', vertical='center')

    for col, w in zip('ABCDEFG', [13,50,22,10,14,12,10]):
        w5.column_dimensions[col].width = w

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out

# ─────────────────────────────────────────────────────────────
# ROTAS — AUTH
# ─────────────────────────────────────────────────────────────
@app.route('/')
def index():
    if 'user' not in session: return redirect(url_for('login'))
    users = load_users()
    role  = users.get(session['user'], {}).get('role')
    return redirect(url_for('admin_dashboard') if role == 'admin' else url_for('dashboard'))

@app.route('/login', methods=['GET','POST'])
def login():
    error = ''; prefill = ''
    if request.method == 'POST':
        email = request.form.get('email','').strip().lower()
        senha = request.form.get('senha','')
        users = load_users()
        if email in users and check_password_hash(users[email]['password'], senha):
            if not users[email].get('ativo', True):
                error = 'Conta desativada. Contate o administrador.'
            else:
                session.permanent = False   # sessão expira ao fechar o browser
                session['user']   = email
                add_log('Login', email, f'IP: {request.remote_addr}')
                role = users[email].get('role')
                return redirect(url_for('admin_dashboard') if role == 'admin' else url_for('dashboard'))
        else:
            error = 'Email ou senha inválidos.'
            prefill = email
    return render_template('login.html', error=error, prefill_email=prefill)

@app.route('/cadastro', methods=['GET','POST'])
def cadastro():
    error = ''; success = ''
    if request.method == 'POST':
        nome  = request.form.get('nome','').strip()
        email = request.form.get('email','').strip().lower()
        pwd   = request.form.get('password','')
        cpwd  = request.form.get('confirm_password','')
        if not all([nome, email, pwd]):  error = 'Preencha todos os campos.'
        elif pwd != cpwd:                error = 'Senhas não conferem.'
        elif len(pwd) < 6:               error = 'Mínimo 6 caracteres.'
        else:
            users = load_users()
            if email in users:
                error = 'E-mail já cadastrado.'
            else:
                users[email] = {'password': generate_password_hash(pwd), 'nome': nome,
                                'role': 'user', 'files': [], 'whatsapp': '',
                                'created_at': datetime.now().isoformat(), 'ativo': True}
                save_users(users)
                add_log('Novo cadastro', email)
                success = 'Conta criada! Faça login.'
    return render_template('cadastro.html', error=error, success=success)

@app.route('/logout')
def logout():
    email = session.get('user','')
    users = load_users()
    if email in users:
        for f in users[email].get('files', []):
            if os.path.exists(f.get('path','')): os.remove(f['path'])
        users[email]['files'] = []
        save_users(users)
    add_log('Logout', email)
    session.clear()
    return redirect(url_for('login'))

# ─────────────────────────────────────────────────────────────
# ROTAS — ADMIN
# ─────────────────────────────────────────────────────────────
@app.route('/admin')
def admin_dashboard():
    redir = require_admin()
    if redir: return redir
    users = load_users()
    ctx   = get_ctx()
    lista = []
    for em, d in users.items():
        lista.append({'email': em, 'nome': d['nome'], 'role': d['role'],
                      'files_count': len(d.get('files',[])),
                      'whatsapp': d.get('whatsapp',''),
                      'created_at': d.get('created_at','')[:10],
                      'ativo': d.get('ativo', True)})
    return render_template('admin_dashboard.html', **ctx, active='admin',
                           usuarios=lista,
                           total_usuarios=len(users),
                           total_ativos=sum(1 for d in users.values() if d.get('ativo', True)),
                           support_messages=load_support(),
                           logs=get_logs()[:50])

@app.route('/admin/criar-usuario', methods=['POST'])
def admin_criar_usuario():
    redir = require_admin()
    if redir: return redir
    nome  = request.form.get('nome','').strip()
    email = request.form.get('email','').strip().lower()
    pwd   = request.form.get('password','')
    role  = request.form.get('role','user')
    users = load_users()
    if email in users:
        flash('E-mail já cadastrado!', 'danger')
    elif not all([nome, email, pwd]):
        flash('Preencha todos os campos!', 'danger')
    else:
        users[email] = {'password': generate_password_hash(pwd), 'nome': nome,
                        'role': role, 'files': [], 'whatsapp': '',
                        'created_at': datetime.now().isoformat(), 'ativo': True}
        save_users(users)
        add_log('Admin criou usuário', session['user'], f'Criado: {email} ({role})')
        flash(f'Usuário {nome} criado com sucesso!', 'success')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/editar-usuario', methods=['POST'])
def admin_editar_usuario():
    redir = require_admin()
    if redir: return redir
    email_orig = request.form.get('email_orig','').strip().lower()
    novo_nome  = request.form.get('nome','').strip()
    novo_email = request.form.get('email','').strip().lower()
    novo_role  = request.form.get('role','user')
    novo_pwd   = request.form.get('password','').strip()
    users      = load_users()

    if email_orig not in users:
        flash('Usuário não encontrado.', 'danger')
        return redirect(url_for('admin_dashboard'))

    # Não deixa remover o próprio admin
    if email_orig == session['user'] and novo_role != 'admin':
        flash('Você não pode rebaixar sua própria conta.', 'danger')
        return redirect(url_for('admin_dashboard'))

    u = users.pop(email_orig)
    u['nome'] = novo_nome
    u['role'] = novo_role
    if novo_pwd and len(novo_pwd) >= 6:
        u['password'] = generate_password_hash(novo_pwd)
    if novo_email != email_orig and novo_email in users:
        flash('Novo e-mail já em uso.', 'danger')
        users[email_orig] = u
    else:
        users[novo_email] = u
        if email_orig == session['user']:
            session['user'] = novo_email
    save_users(users)
    add_log('Admin editou usuário', session['user'], f'{email_orig} → {novo_email}')
    flash('Usuário atualizado!', 'success')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/remover-usuario', methods=['POST'])
def remover_usuario():
    redir = require_admin()
    if redir: return redir
    email = request.form.get('email','').strip().lower()
    users = load_users()
    if email == session['user']:
        flash('Você não pode remover sua própria conta!', 'danger')
        return redirect(url_for('admin_dashboard'))
    if email in users:
        for f in users[email].get('files',[]):
            if os.path.exists(f.get('path','')): os.remove(f['path'])
        del users[email]
        save_users(users)
        add_log('Admin removeu usuário', session['user'], email)
        flash(f'Usuário {email} removido.', 'success')
    else:
        flash('Usuário não encontrado.', 'danger')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/toggle-ativo', methods=['POST'])
def toggle_ativo():
    redir = require_admin()
    if redir: return redir
    email = request.form.get('email','').strip().lower()
    users = load_users()
    if email in users and email != session['user']:
        users[email]['ativo'] = not users[email].get('ativo', True)
        save_users(users)
        estado = 'ativado' if users[email]['ativo'] else 'desativado'
        add_log(f'Admin {estado} usuário', session['user'], email)
        flash(f'Usuário {email} {estado}.', 'success')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/responder-suporte', methods=['POST'])
def responder_suporte():
    redir = require_admin()
    if redir: return redir
    idx     = int(request.form.get('index',-1))
    resposta= request.form.get('resposta','').strip()
    msgs    = load_support()
    if 0 <= idx < len(msgs):
        msgs[idx]['lido']    = True
        msgs[idx]['resposta']= resposta
        msgs[idx]['respondido_em'] = datetime.now().isoformat()
        save_support(msgs)
        flash('Resposta enviada!', 'success')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/configuracoes', methods=['GET','POST'])
def admin_configuracoes():
    redir = require_admin()
    if redir: return redir
    users = load_users(); ctx = get_ctx()
    error = ''; success = ''
    if request.method == 'POST':
        acao = request.form.get('acao','')
        if acao == 'alterar_email':
            novo = request.form.get('novo_email','').strip().lower()
            if not novo: error = 'E-mail vazio.'
            elif novo in users and novo != session['user']: error = 'E-mail já em uso.'
            else:
                users[novo] = users.pop(session['user'])
                save_users(users); session['user'] = novo; success = 'E-mail alterado!'
        elif acao == 'alterar_senha':
            sa = request.form.get('senha_atual','')
            ns = request.form.get('nova_senha','')
            cs = request.form.get('confirmar_senha','')
            if not check_password_hash(users[session['user']]['password'], sa): error = 'Senha atual incorreta.'
            elif ns != cs:   error = 'Senhas não conferem.'
            elif len(ns) < 6:error = 'Mínimo 6 caracteres.'
            else:
                users[session['user']]['password'] = generate_password_hash(ns)
                save_users(users); success = 'Senha alterada!'
    return render_template('configuracoes.html', **ctx, active='admin_config',
                           admin_email=session['user'], error=error, success=success)

# ─────────────────────────────────────────────────────────────
# ROTAS — USUÁRIO
# ─────────────────────────────────────────────────────────────
@app.route('/dashboard')
def dashboard():
    redir = require_login()
    if redir: return redir
    users    = load_users()
    user     = users.get(session['user'])
    ctx      = get_ctx()
    dash     = get_dash(session['user'])
    files    = user.get('files', [])
    has_data = len(dash['transactions']) > 0
    return render_template('dashboard.html', **ctx, active='dashboard',
                           dash=dash, files=files, has_data=has_data)

@app.route('/conciliacao')
def conciliacao():
    redir = require_login()
    if redir: return redir
    ctx  = get_ctx()
    dash = get_dash(session['user'])
    rec  = [t for t in dash['transactions'] if t['valor'] > 0]
    desp = [t for t in dash['transactions'] if t['valor'] < 0]
    return render_template('conciliacao.html', **ctx, active='conciliacao',
                           creditos=rec, debitos=desp,
                           total_creditos=sum(t['valor'] for t in rec),
                           total_debitos=sum(abs(t['valor']) for t in desp))

@app.route('/upload', methods=['POST'])
def upload_file():
    redir = require_login()
    if redir: return redir
    if 'file' not in request.files: return redirect(url_for('dashboard'))
    file = request.files['file']
    if file.filename == '': return redirect(url_for('dashboard'))
    if file and '.' in file.filename and file.filename.rsplit('.',1)[1].lower() in ALLOWED:
        fname = secure_filename(f"{session['user']}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
        fpath = os.path.join(UPLOAD_FOLDER, fname)
        file.save(fpath)
        users = load_users()
        users[session['user']].setdefault('files', []).append({
            'name': file.filename, 'path': fpath,
            'type': file.filename.rsplit('.',1)[1].lower(),
            'date': datetime.now().isoformat()
        })
        save_users(users)
        add_log('Upload de arquivo', session['user'], file.filename)
        flash(f'Arquivo "{file.filename}" importado com sucesso!', 'success')
    return redirect(url_for('dashboard'))

@app.route('/remover-arquivo/<path:filename>')
def remover_arquivo(filename):
    redir = require_login()
    if redir: return redir
    users = load_users()
    nova  = [f for f in users[session['user']].get('files',[]) if f['name'] != filename]
    for f in users[session['user']].get('files',[]):
        if f['name'] == filename and os.path.exists(f['path']):
            os.remove(f['path'])
    users[session['user']]['files'] = nova
    save_users(users)
    return redirect(url_for('dashboard'))

@app.route('/configuracoes', methods=['GET','POST'])
def user_configuracoes():
    redir = require_login()
    if redir: return redir
    users = load_users(); user = users.get(session['user'])
    ctx   = get_ctx(); error = ''; success = ''
    if request.method == 'POST':
        users[session['user']]['whatsapp'] = request.form.get('whatsapp','').strip()
        save_users(users); success = 'WhatsApp atualizado!'
    return render_template('user_config.html', **ctx, active='config',
                           whatsapp=user.get('whatsapp',''), error=error, success=success)

@app.route('/suporte', methods=['GET','POST'])
def suporte():
    redir = require_login()
    if redir: return redir
    ctx = get_ctx(); success = ''
    if request.method == 'POST':
        msg = request.form.get('mensagem','').strip()
        if msg:
            msgs = load_support()
            msgs.append({'de': session['user'], 'nome': ctx['nome'],
                         'mensagem': msg, 'data': datetime.now().isoformat(), 'lido': False})
            save_support(msgs); success = 'Mensagem enviada!'
    return render_template('suporte.html', **ctx, active='suporte', success=success)

# ─────────────────────────────────────────────────────────────
# ROTAS — DOWNLOADS
# ─────────────────────────────────────────────────────────────
@app.route('/download/excel')
def download_excel():
    redir = require_login()
    if redir: return redir
    dash = get_dash(session['user'])
    add_log('Download Excel', session['user'])
    return send_file(build_excel(dash),
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,
                     download_name=f'FiscalPro_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')

@app.route('/download/conciliacao')
def download_conciliacao():
    redir = require_login()
    if redir: return redir
    dash = get_dash(session['user'])
    G    = '22a060'; GL = 'E8F8EF'; R = 'C0392B'; RL = 'FDEDEC'
    wb   = Workbook(); brd = _brd()

    def make_sheet(ws, titulo, lista, clr, bg):
        ws.sheet_view.showGridLines = False
        ws.merge_cells('A1:F1')
        ws['A1'].value = titulo
        ws['A1'].font  = Font(name='Calibri', bold=True, size=14, color='FFFFFF')
        ws['A1'].fill  = _fill(clr)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 30
        for i, h in enumerate(['Data','Descrição','Categoria','Valor (R$)','% do Total','Status'], 1):
            c = ws.cell(row=2, column=i, value=h)
            c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(clr)
            c.alignment = Alignment(horizontal='center', vertical='center'); c.border = brd
        total = sum(abs(t['valor']) for t in lista)
        for j, t in enumerate(lista):
            rr = j + 3
            pct= round(abs(t['valor'])/total*100,1) if total > 0 else 0
            fill_alt = bg if j%2==0 else 'FFFFFF'
            for ci, v in enumerate([t['data'], t['descricao'][:70], t['categoria'],
                                      round(abs(t['valor']),2), f'{pct}%',
                                      '✅ Pago' if t.get('pago') else '⚠️ Pendente'], 1):
                c = ws.cell(row=rr, column=ci, value=v)
                c.fill = _fill(fill_alt); c.border = brd
                c.alignment = Alignment(horizontal='left' if ci==2 else 'center', vertical='center')
        for col, w in zip('ABCDEF', [13,50,22,14,12,14]):
            ws.column_dimensions[col].width = w

    ws1 = wb.active; ws1.title = 'Receitas'
    make_sheet(ws1, 'CONCILIAÇÃO — RECEITAS', [t for t in dash['transactions'] if t['valor']>0], G, GL)
    ws2 = wb.create_sheet('Despesas')
    make_sheet(ws2, 'CONCILIAÇÃO — DESPESAS', [t for t in dash['transactions'] if t['valor']<0], R, RL)

    out = io.BytesIO(); wb.save(out); out.seek(0)
    return send_file(out,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,
                     download_name=f'Conciliacao_{datetime.now().strftime("%Y%m%d")}.xlsx')

# ─────────────────────────────────────────────────────────────
if __name__ == '__main__':
    users = load_users()
    print('\n' + '='*50)
    print('  🌿 FiscalPro iniciando...')
    print('='*50)
    print('  URL:   http://127.0.0.1:5000')
    print('  Admin: admin@fiscal.app')
    print('  Senha: admin123')
    print('='*50 + '\n')
    app.run(debug=True, host='127.0.0.1', port=5000)
