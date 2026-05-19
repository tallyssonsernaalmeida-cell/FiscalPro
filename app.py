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
app.permanent_session_lifetime = timedelta(hours=12)

USERS_FILE    = 'users.json'
SUPPORT_FILE  = 'support.json'
LOGS_FILE     = 'logs.json'
UPLOAD_FOLDER = 'uploads'
USER_DATA     = 'user_data'
CONCILIACAO_FOLDER = 'conciliacao'

for d in [UPLOAD_FOLDER, USER_DATA, CONCILIACAO_FOLDER]:
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
    logs = logs[:500]
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
    return None

def require_admin():
    redir = require_login()
    if redir: return redir
    users = load_users()
    if users.get(session['user'], {}).get('role') != 'admin':
        flash('Acesso restrito a administradores.', 'warning')
        return redirect(url_for('dashboard'))
    return None

# ─────────────────────────────────────────────────────────────
# CATEGORIZAÇÃO
# ─────────────────────────────────────────────────────────────
def categorizar(desc):
    d = str(desc).lower()
    if any(x in d for x in ['salário','salario','holerite','folha']):             return 'Salário/Renda'
    if any(x in d for x in ['freelance','comissão','comissao','honorário']):      return 'Renda Extra'
    if any(x in d for x in ['pix recebid','ted recebid','transferência recebida',
                              'depósito','deposito','estorno']):                   return 'Transferência Recebida'
    if any(x in d for x in ['pix enviad','ted enviad','transferência enviada']):  return 'Transferência Enviada'
    if any(x in d for x in ['aluguel','condomínio','condominio','iptu',
                              'financiamento','prestação','prestacao']):           return 'Moradia'
    if any(x in d for x in ['supermercado','mercado','hortifruti','açougue',
                              'padaria','restaurante','ifood','lanchonete',
                              'alimentação','alimentacao','comida']):              return 'Alimentação'
    if any(x in d for x in ['uber','99 ','cabify','ônibus','onibus','metrô','metro',
                              'gasolina','combustível','combustivel','posto',
                              'estacionamento','pedágio','pedagio']):              return 'Transporte'
    if any(x in d for x in ['energia','água','agua','internet','telefone',
                              'celular','claro','vivo','tim','net','gás','gas']):  return 'Contas & Serviços'
    if any(x in d for x in ['farmácia','farmacia','drogaria','médico','medico',
                              'hospital','plano de saúde','unimed','exame',
                              'consulta','remédio','remedio']):                    return 'Saúde'
    if any(x in d for x in ['netflix','spotify','amazon','disney','hbo',
                              'cinema','teatro','ingresso','show']):               return 'Lazer'
    if any(x in d for x in ['curso','faculdade','escola','educação','educacao',
                              'livro','livraria','mensalidade']):                  return 'Educação'
    if any(x in d for x in ['roupa','calçado','calcado','moda','loja','zara',
                              'renner','riachuelo','shopping']):                   return 'Vestuário'
    if any(x in d for x in ['imposto','taxa','multa','darf','irpf','irpj']):      return 'Impostos & Taxas'
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
                out.append({
                    'data': str(t.date)[:10],
                    'descricao': d,
                    'valor': v,
                    'tipo': 'Crédito' if v > 0 else 'Débito',
                    'categoria': categorizar(d),
                    'pago': False
                })
        return out
    except Exception as e:
        print(f'ERRO OFX: {e}')
        return []

def process_excel(fp):
    """Processa arquivo Excel/CSV e retorna lista de transações."""
    try:
        ext = fp.rsplit('.', 1)[-1].lower()
        print(f'Processando arquivo: {fp} (ext: {ext})')
        
        if ext == 'csv':
            df = pd.read_csv(fp, sep=None, engine='python', encoding='utf-8', errors='ignore')
        elif ext in ['xlsx','xlsm']:
            df = pd.read_excel(fp, engine='openpyxl')
        else:
            df = pd.read_excel(fp, engine='xlrd')

        print(f'Colunas encontradas: {list(df.columns)}')
        print(f'Total de linhas: {len(df)}')
        
        df.columns = [str(c).strip().lower() for c in df.columns]
        col_d = next((c for c in df.columns if any(k in c for k in ['data','date','dt'])), None)
        col_v = next((c for c in df.columns if any(k in c for k in ['valor','value','amount','montante'])), None)
        col_t = next((c for c in df.columns if any(k in c for k in ['descri','memo','histor','lançamento','lancamento','complement','nome'])), None)

        print(f'Coluna data: {col_d}, Coluna valor: {col_v}, Coluna desc: {col_t}')

        out = []
        for idx, row in df.iterrows():
            v = 0.0
            if col_v:
                try:
                    raw = str(row[col_v]).replace('R$','').replace(' ','').strip()
                    if raw == '' or raw == 'nan':
                        v = 0.0
                    else:
                        if ',' in raw and '.' in raw:
                            if raw.rfind(',') > raw.rfind('.'):
                                raw = raw.replace('.','').replace(',','.')
                            else:
                                raw = raw.replace(',','')
                        elif ',' in raw:
                            raw = raw.replace(',','.')
                        v = float(raw)
                except Exception as e:
                    print(f'Erro ao converter valor na linha {idx}: {row[col_v]} -> {e}')
                    v = 0.0
            
            desc = str(row[col_t])[:120] if col_t and str(row[col_t]) != 'nan' else 'Sem descrição'
            data = str(row[col_d])[:10] if col_d and str(row[col_d]) != 'nan' else 'N/A'
            
            out.append({
                'data': data,
                'descricao': desc,
                'valor': v,
                'tipo': 'Crédito' if v > 0 else 'Débito',
                'categoria': categorizar(desc),
                'pago': False
            })
        
        print(f'Total de transações processadas: {len(out)}')
        return out
    except Exception as e:
        print(f'ERRO Excel/CSV: {e}')
        import traceback
        traceback.print_exc()
        return []

# ─────────────────────────────────────────────────────────────
# DASHBOARD DATA
# ─────────────────────────────────────────────────────────────
def get_dash(email):
    users = load_users()
    files = users.get(email, {}).get('files', [])
    txs   = []
    
    print(f'\n=== get_dash para {email} ===')
    print(f'Arquivos encontrados: {len(files)}')
    
    for fi in files:
        path = fi.get('path','')
        print(f'  Arquivo: {fi.get("name","")} -> {path}')
        if not os.path.exists(path):
            print(f'    ARQUIVO NÃO ENCONTRADO: {path}')
            continue
        ext = path.rsplit('.', 1)[-1].lower()
        if ext == 'ofx':
            resultado = process_ofx(path)
        else:
            resultado = process_excel(path)
        print(f'    Transações extraídas: {len(resultado)}')
        txs.extend(resultado)

    rec  = sum(t['valor'] for t in txs if t['valor'] > 0)
    desp = sum(abs(t['valor']) for t in txs if t['valor'] < 0)
    saldo = round(rec - desp, 2)

    cat_d, cat_r = {}, {}
    for t in txs:
        cat = t['categoria']
        if t['valor'] < 0:
            cat_d[cat] = cat_d.get(cat, 0) + abs(t['valor'])
        else:
            cat_r[cat] = cat_r.get(cat, 0) + t['valor']

    por_mes = {}
    for t in txs:
        mes = t['data'][:7] if len(t['data']) >= 7 else t['data']
        if mes not in por_mes:
            por_mes[mes] = {'creditos': 0, 'debitos': 0}
        if t['valor'] > 0:
            por_mes[mes]['creditos'] += t['valor']
        else:
            por_mes[mes]['debitos'] += abs(t['valor'])

    return {
        'transacoes':      txs,
        'transactions':    txs,
        'total_receitas':  round(rec, 2),
        'total_creditos':  round(rec, 2),
        'total_despesas':  round(desp, 2),
        'total_debitos':   round(desp, 2),
        'saldo':           saldo,
        'status':          'Lucro' if saldo > 0 else ('Prejuízo' if saldo < 0 else 'Equilibrado'),
        'total_transacoes': len(txs),
        'qtd_transacoes':  len(txs),
        'por_categoria':   [{'categoria': k, 'valor': round(v,2)} for k,v in sorted(cat_d.items(), key=lambda x:x[1], reverse=True)],
        'categorias':      cat_d,
        'categorias_rec':  cat_r,
        'top_categorias':  sorted(cat_d.items(), key=lambda x: x[1], reverse=True)[:5],
        'por_mes':         [{'mes': k, 'creditos': round(v['creditos'],2), 'debitos': round(v['debitos'],2)} for k,v in sorted(por_mes.items())],
        'percentual_lucro': round((saldo/rec)*100, 1) if rec > 0 else 0
    }

# ─────────────────────────────────────────────────────────────
# EXCEL COMPLETO
# ─────────────────────────────────────────────────────────────
def _brd():
    s = Side(style='thin', color='D0D0D0')
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(cor):
    return PatternFill(start_color=cor, end_color=cor, fill_type='solid')

def _fnt(cor='1E293B', bold=False, size=10):
    return Font(name='Calibri', color=cor, bold=bold, size=size)

def build_excel(data):
    G, GL, R, RL = '22a060','E8F8EF','C0392B','FDEDEC'
    wb  = Workbook()
    brd = _brd()

    rec_list  = [t for t in data['transactions'] if t['valor'] > 0]
    desp_list = [t for t in data['transactions'] if t['valor'] < 0]
    tot_r = data['total_receitas']
    tot_d = data['total_despesas']
    saldo = data['saldo']
    pct   = data['percentual_lucro']
    cats  = sorted(data['categorias'].items(), key=lambda x: x[1], reverse=True)

    # ABA 1 - PAINEL RESUMO
    ws = wb.active
    ws.title = 'Painel Resumo'
    ws.sheet_view.showGridLines = False

    ws.merge_cells('A1:I1')
    ws['A1'].value = 'FISCALPRO - RELATORIO FINANCEIRO COMPLETO'
    ws['A1'].font  = Font(name='Calibri', bold=True, size=18, color='FFFFFF')
    ws['A1'].fill  = _fill(G)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 42

    ws.merge_cells('A2:I2')
    ws['A2'].value = f'Gerado em {datetime.now().strftime("%d/%m/%Y as %H:%M")}   -   {data["qtd_transacoes"]} transacoes'
    ws['A2'].font  = _fnt('6B7280', size=10)
    ws['A2'].fill  = _fill('F4F6F8')
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    kpis = [
        ('B', 'RECEITAS TOTAIS',  f'R$ {tot_r:,.2f}',  G,      GL),
        ('D', 'DESPESAS TOTAIS',  f'R$ {tot_d:,.2f}',  R,      RL),
        ('F', 'SALDO LIQUIDO',    f'R$ {saldo:,.2f}',  G if saldo>=0 else R, GL if saldo>=0 else RL),
        ('H', 'MARGEM DE LUCRO',  f'{pct}%',           '1D4ED8','EFF6FF'),
    ]
    for col, label, val, clr, bg in kpis:
        ws.merge_cells(f'{col}4:{col}6')
        ws[f'{col}4'].value = f'{label}\n{val}'
        ws[f'{col}4'].font  = Font(name='Calibri', bold=True, size=12, color=clr)
        ws[f'{col}4'].fill  = _fill(bg)
        ws[f'{col}4'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[4].height = 50

    for col in 'ABCDEFGHI':
        ws.column_dimensions[col].width = 18

    r = 7
    ws.merge_cells(f'A{r}:D{r}')
    ws[f'A{r}'].value = 'DESPESAS POR CATEGORIA'
    ws[f'A{r}'].font  = Font(name='Calibri', bold=True, size=13, color=G)
    r += 1
    for i, h in enumerate(['Categoria','Total (R$)','% do Total','Qtd'], 1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(G)
        c.alignment = Alignment(horizontal='center', vertical='center'); c.border = brd
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

    if cats:
        pie = PieChart()
        pie.title = 'Despesas por Categoria'; pie.style = 10
        pie.width = 14; pie.height = 12
        pie.add_data(Reference(ws, min_col=2, min_row=cat_start-1, max_row=cat_end))
        pie.set_categories(Reference(ws, min_col=1, min_row=cat_start, max_row=cat_end))
        ws.add_chart(pie, 'F9')

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out

# ─────────────────────────────────────────────────────────────
# ROTAS — AUTH
# ─────────────────────────────────────────────────────────────
@app.route('/')
def index():
    if 'user' not in session:
        return redirect(url_for('login'))
    users = load_users()
    role  = users.get(session['user'], {}).get('role')
    if role == 'admin':
        return redirect(url_for('admin_dashboard'))
    return redirect(url_for('dashboard'))

@app.route('/login', methods=['GET','POST'])
def login():
    error = ''
    prefill = ''
    if request.method == 'POST':
        email = request.form.get('email','').strip().lower()
        senha = request.form.get('senha','')
        users = load_users()
        if email in users and check_password_hash(users[email]['password'], senha):
            if not users[email].get('ativo', True):
                error = 'Conta desativada. Contate o administrador.'
            else:
                session['user'] = email
                add_log('Login', email, f'IP: {request.remote_addr}')
                flash(f'Bem-vindo(a), {users[email]["nome"]}!', 'success')
                if users[email].get('role') == 'admin':
                    return redirect(url_for('admin_dashboard'))
                return redirect(url_for('dashboard'))
        else:
            error = 'Email ou senha inválidos.'
            prefill = email
    return render_template('login.html', error=error, prefill_email=prefill)

@app.route('/cadastro', methods=['GET','POST'])
def cadastro():
    error = ''
    if request.method == 'POST':
        nome  = request.form.get('nome','').strip()
        email = request.form.get('email','').strip().lower()
        pwd   = request.form.get('password','')
        cpwd  = request.form.get('confirm_password','')
        if not all([nome, email, pwd]):
            error = 'Preencha todos os campos obrigatórios.'
        elif pwd != cpwd:
            error = 'Senhas não conferem.'
        elif len(pwd) < 6:
            error = 'A senha deve ter no mínimo 6 caracteres.'
        else:
            users = load_users()
            if email in users:
                error = 'E-mail já cadastrado.'
            else:
                users[email] = {
                    'password': generate_password_hash(pwd),
                    'nome': nome,
                    'role': 'user',
                    'files': [],
                    'whatsapp': request.form.get('whatsapp','').strip(),
                    'created_at': datetime.now().isoformat(),
                    'ativo': True
                }
                save_users(users)
                add_log('Novo cadastro', email)
                # LOGIN AUTOMÁTICO
                session['user'] = email
                flash(f'Bem-vindo(a), {nome}! Sua conta foi criada com sucesso.', 'success')
                return redirect(url_for('dashboard'))
    return render_template('cadastro.html', error=error)

@app.route('/logout')
def logout():
    add_log('Logout', session.get('user',''))
    session.clear()
    flash('Você saiu da conta.', 'info')
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
        lista.append({
            'email': em,
            'nome': d['nome'],
            'role': d['role'],
            'files_count': len(d.get('files',[])),
            'whatsapp': d.get('whatsapp',''),
            'created_at': d.get('created_at','')[:10],
            'ativo': d.get('ativo', True)
        })
    return render_template('admin_dashboard.html',
                           **ctx,
                           active='admin',
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
        users[email] = {
            'password': generate_password_hash(pwd),
            'nome': nome,
            'role': role,
            'files': [],
            'whatsapp': '',
            'created_at': datetime.now().isoformat(),
            'ativo': True
        }
        save_users(users)
        add_log('Admin criou usuário', session['user'], f'Criado: {email} ({role})')
        flash(f'Usuário {nome} criado com sucesso!', 'success')
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

@app.route('/admin/remover-usuario', methods=['POST'])
def remover_usuario():
    redir = require_admin()
    if redir: return redir
    email = request.form.get('email','').strip().lower()
    users = load_users()
    if email == session['user']:
        flash('Você não pode remover sua própria conta!', 'danger')
    elif email in users:
        for f in users[email].get('files',[]):
            if os.path.exists(f.get('path','')): os.remove(f['path'])
        del users[email]
        save_users(users)
        add_log('Admin removeu usuário', session['user'], email)
        flash(f'Usuário {email} removido.', 'success')
    else:
        flash('Usuário não encontrado.', 'danger')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/configuracoes', methods=['GET','POST'])
def configuracoes():
    redir = require_admin()
    if redir: return redir
    users = load_users()
    ctx   = get_ctx()
    error = ''
    success = ''
    if request.method == 'POST':
        acao = request.form.get('acao','')
        if acao == 'alterar_email':
            novo = request.form.get('novo_email','').strip().lower()
            if not novo:
                error = 'E-mail vazio.'
            elif novo in users and novo != session['user']:
                error = 'E-mail já em uso.'
            else:
                users[novo] = users.pop(session['user'])
                save_users(users)
                session['user'] = novo
                success = 'E-mail alterado com sucesso!'
        elif acao == 'alterar_senha':
            sa = request.form.get('senha_atual','')
            ns = request.form.get('nova_senha','')
            cs = request.form.get('confirmar_senha','')
            if not check_password_hash(users[session['user']]['password'], sa):
                error = 'Senha atual incorreta.'
            elif ns != cs:
                error = 'Senhas não conferem.'
            elif len(ns) < 6:
                error = 'Mínimo 6 caracteres.'
            else:
                users[session['user']]['password'] = generate_password_hash(ns)
                save_users(users)
                success = 'Senha alterada com sucesso!'
    return render_template('Configurações.html',
                           **ctx,
                           active='admin_config',
                           admin_email=session['user'],
                           error=error,
                           success=success,
                           config={},
                           sys_python='Python 3',
                           sys_platform=os.name)

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
    files    = user.get('files', []) if user else []
    has_data = len(dash.get('transactions', [])) > 0

    # Paginação
    pagina = request.args.get('pagina', 1, type=int)
    por_pagina = 20
    transacoes = list(dash.get('transactions', []))
    total_paginas = max(1, (len(transacoes) + por_pagina - 1) // por_pagina)
    inicio = (pagina - 1) * por_pagina
    dash['transactions'] = transacoes[inicio:inicio + por_pagina]
    dash['pagina'] = pagina
    dash['total_paginas'] = total_paginas

    return render_template('dashboard.html',
                           **ctx,
                           active='dashboard',
                           dash=dash,
                           files=files,
                           has_data=has_data)

@app.route('/conciliacao', methods=['GET','POST'])
def conciliacao():
    redir = require_login()
    if redir: return redir
    ctx  = get_ctx()
    
    # Processar uploads de conciliação
    if request.method == 'POST':
        if 'arquivo_extrato' in request.files:
            file = request.files['arquivo_extrato']
            if file.filename != '':
                fname = secure_filename(f"extrato_{session['user']}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
                fpath = os.path.join(CONCILIACAO_FOLDER, fname)
                file.save(fpath)
                flash(f'Extrato "{file.filename}" carregado!', 'success')
        
        if 'arquivo_notas' in request.files:
            file = request.files['arquivo_notas']
            if file.filename != '':
                fname = secure_filename(f"notas_{session['user']}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
                fpath = os.path.join(CONCILIACAO_FOLDER, fname)
                file.save(fpath)
                flash(f'Notas "{file.filename}" carregadas!', 'success')
    
    # Buscar dados
    dash = get_dash(session['user'])
    rec  = [t for t in dash['transactions'] if t['valor'] > 0]
    desp = [t for t in dash['transactions'] if t['valor'] < 0]

    mes_selecionado = request.args.get('mes', '')
    if mes_selecionado:
        rec  = [t for t in rec if t['data'].startswith(mes_selecionado)]
        desp = [t for t in desp if t['data'].startswith(mes_selecionado)]

    meses = sorted(set(t['data'][:7] for t in dash['transactions'] if len(t['data'])>=7))
    
    # Conciliação automática (mesmo valor = conciliado)
    conciliados = 0
    for r in rec:
        for d in desp:
            if abs(r['valor']) == abs(d['valor']):
                conciliados += 1
                break

    total_itens = len(rec) + len(desp)
    
    # Recomendações
    recomendacoes = []
    if dash['saldo'] < 0:
        recomendacoes.append('⚠️ Seu saldo está negativo. Reveja suas despesas.')
    if dash['total_despesas'] > dash['total_receitas']:
        recomendacoes.append('📉 Suas despesas superam as receitas. Crie um orçamento.')
    if dash['percentual_lucro'] < 10:
        recomendacoes.append('💡 Margem de lucro abaixo de 10%. Considere reduzir custos.')
    if not recomendacoes:
        recomendacoes.append('✅ Suas finanças estão saudáveis!')

    return render_template('conciliacao.html',
                           **ctx,
                           active='conciliacao',
                           extrato=rec,
                           notas=desp,
                           creditos=rec,
                           debitos=desp,
                           total_creditos=sum(t['valor'] for t in rec),
                           total_debitos=sum(abs(t['valor']) for t in desp),
                           conciliados=conciliados,
                           total_itens=total_itens,
                           meses_disponiveis=meses,
                           mes_selecionado=mes_selecionado,
                           recomendacoes=recomendacoes)

@app.route('/upload', methods=['POST'])
def upload_file():
    redir = require_login()
    if redir: return redir
    if 'file' not in request.files and 'arquivo' not in request.files:
        flash('Nenhum arquivo selecionado.', 'warning')
        return redirect(url_for('dashboard'))

    file = request.files.get('file') or request.files.get('arquivo')
    if file.filename == '':
        flash('Nome de arquivo vazio.', 'warning')
        return redirect(url_for('dashboard'))

    if '.' in file.filename and file.filename.rsplit('.',1)[1].lower() in ALLOWED:
        fname = secure_filename(f"{session['user']}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
        fpath = os.path.join(UPLOAD_FOLDER, fname)
        file.save(fpath)
        users = load_users()
        if session['user'] not in users:
            users[session['user']] = {'files': []}
        users[session['user']].setdefault('files', []).append({
            'name': file.filename,
            'path': fpath,
            'type': file.filename.rsplit('.',1)[1].lower(),
            'date': datetime.now().isoformat()
        })
        save_users(users)
        add_log('Upload de arquivo', session['user'], file.filename)
        flash(f'Arquivo "{file.filename}" importado com sucesso!', 'success')
    else:
        flash(f'Formato não permitido. Use: {", ".join(ALLOWED)}', 'danger')
    return redirect(url_for('dashboard'))

@app.route('/remover-arquivo/<path:filename>')
def remover_arquivo(filename):
    redir = require_login()
    if redir: return redir
    users = load_users()
    nova  = []
    for f in users.get(session['user'], {}).get('files', []):
        if f['name'] == filename:
            if os.path.exists(f.get('path','')):
                os.remove(f['path'])
        else:
            nova.append(f)
    if session['user'] in users:
        users[session['user']]['files'] = nova
    save_users(users)
    flash(f'Arquivo "{filename}" removido.', 'info')
    return redirect(url_for('dashboard'))

@app.route('/configuracoes_usuario', methods=['GET','POST'])
def configuracoes_usuario():
    redir = require_login()
    if redir: return redir
    users = load_users()
    user  = users.get(session['user'], {})
    ctx   = get_ctx()
    error = ''
    success = ''
    if request.method == 'POST':
        acao = request.form.get('acao','')
        if acao == 'alterar_senha':
            sa = request.form.get('senha_atual','')
            ns = request.form.get('nova_senha','')
            cs = request.form.get('confirmar_senha','')
            if not check_password_hash(users[session['user']]['password'], sa):
                error = 'Senha atual incorreta.'
            elif ns != cs:
                error = 'Senhas não conferem.'
            elif len(ns) < 6:
                error = 'Mínimo 6 caracteres.'
            else:
                users[session['user']]['password'] = generate_password_hash(ns)
                save_users(users)
                success = 'Senha alterada com sucesso!'
        elif acao == 'preferencias':
            users[session['user']]['notificacoes'] = request.form.get('notificacoes') == 'on'
            save_users(users)
            success = 'Preferências salvas!'
        else:
            users[session['user']]['nome'] = request.form.get('nome', user.get('nome','')).strip()
            users[session['user']]['whatsapp'] = request.form.get('whatsapp','').strip()
            save_users(users)
            success = 'Dados atualizados com sucesso!'
    return render_template('user_configuracoes.html',
                           **ctx,
                           active='config',
                           usuario=users.get(session['user'], {}),
                           error=error,
                           success=success)

@app.route('/suporte', methods=['GET','POST'])
def suporte():
    redir = require_login()
    if redir: return redir
    ctx = get_ctx()
    if request.method == 'POST':
        msg = request.form.get('mensagem','').strip()
        if msg:
            msgs = load_support()
            msgs.append({
                'de': session['user'],
                'nome': ctx['nome'],
                'mensagem': msg,
                'data': datetime.now().isoformat(),
                'lido': False
            })
            save_support(msgs)
            flash('Mensagem enviada com sucesso!', 'success')
        else:
            flash('Escreva uma mensagem.', 'warning')
    return render_template('suporte.html', **ctx, active='suporte')

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

@app.route('/download/csv')
def download_csv():
    redir = require_login()
    if redir: return redir
    dash = get_dash(session['user'])
    df = pd.DataFrame(dash['transactions'])
    out = io.StringIO()
    df.to_csv(out, index=False)
    out.seek(0)
    add_log('Download CSV', session['user'])
    return send_file(io.BytesIO(out.getvalue().encode('utf-8')),
                     mimetype='text/csv',
                     as_attachment=True,
                     download_name=f'FiscalPro_{datetime.now().strftime("%Y%m%d_%H%M")}.csv')

@app.route('/download/conciliacao')
def download_conciliacao():
    redir = require_login()
    if redir: return redir
    dash = get_dash(session['user'])
    wb = Workbook()
    G, GL, R, RL = '22a060','E8F8EF','C0392B','FDEDEC'
    brd = _brd()

    # ABA RECEITAS
    ws1 = wb.active
    ws1.title = 'Receitas'
    ws1.sheet_view.showGridLines = False
    ws1['A1'].value = 'CONCILIACAO - RECEITAS'
    ws1['A1'].font = Font(name='Calibri', bold=True, size=14, color='FFFFFF')
    ws1['A1'].fill = _fill(G)
    for i, h in enumerate(['Data','Descricao','Categoria','Valor (R$)'], 1):
        c = ws1.cell(row=2, column=i, value=h)
        c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(G); c.border = brd
    recs = [t for t in dash['transactions'] if t['valor'] > 0]
    for j, t in enumerate(recs):
        rr = j + 3
        for ci, v in enumerate([t['data'], t['descricao'][:60], t['categoria'], round(t['valor'],2)], 1):
            c = ws1.cell(row=rr, column=ci, value=v)
            c.fill = _fill(GL if j%2==0 else 'FFFFFF'); c.border = brd

    # ABA DESPESAS
    ws2 = wb.create_sheet('Despesas')
    ws2.sheet_view.showGridLines = False
    ws2['A1'].value = 'CONCILIACAO - DESPESAS'
    ws2['A1'].font = Font(name='Calibri', bold=True, size=14, color='FFFFFF')
    ws2['A1'].fill = _fill(R)
    for i, h in enumerate(['Data','Descricao','Categoria','Valor (R$)'], 1):
        c = ws2.cell(row=2, column=i, value=h)
        c.font = _fnt('FFFFFF', True, 11); c.fill = _fill(R); c.border = brd
    desps = [t for t in dash['transactions'] if t['valor'] < 0]
    for j, t in enumerate(desps):
        rr = j + 3
        for ci, v in enumerate([t['data'], t['descricao'][:60], t['categoria'], round(abs(t['valor']),2)], 1):
            c = ws2.cell(row=rr, column=ci, value=v)
            c.fill = _fill(RL if j%2==0 else 'FFFFFF'); c.border = brd

    # ABA RESUMO
    ws3 = wb.create_sheet('Resumo')
    ws3.sheet_view.showGridLines = False
    ws3['A1'].value = 'RESUMO DA CONCILIACAO'
    ws3['A1'].font = Font(name='Calibri', bold=True, size=14, color='FFFFFF')
    ws3['A1'].fill = _fill('1D4ED8')
    dados_resumo = [
        ('Total de Receitas', f'R$ {dash["total_receitas"]:,.2f}'),
        ('Total de Despesas', f'R$ {dash["total_despesas"]:,.2f}'),
        ('Saldo', f'R$ {dash["saldo"]:,.2f}'),
        ('Total de Transacoes', str(dash['qtd_transacoes'])),
        ('Margem de Lucro', f'{dash["percentual_lucro"]}%'),
        ('Status', dash['status']),
    ]
    for i, (label, val) in enumerate(dados_resumo):
        ws3.cell(row=i+3, column=1, value=label).font = Font(bold=True)
        ws3.cell(row=i+3, column=2, value=val)

    for ws in [ws1, ws2]:
        for col, w in zip('ABCD', [13,50,22,14]):
            ws.column_dimensions[col].width = w

    out = io.BytesIO(); wb.save(out); out.seek(0)
    return send_file(out,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,
                     download_name=f'Conciliacao_{datetime.now().strftime("%Y%m%d")}.xlsx')

# ─────────────────────────────────────────────────────────────
if __name__ == '__main__':
    print('\n' + '='*50)
    print('  FiscalPro iniciando...')
    print('='*50)
    print('  URL:   http://127.0.0.1:5000')
    print('  Admin: admin@fiscal.app')
    print('  Senha: admin123')
    print('='*50 + '\n')
    app.run(debug=True, host='127.0.0.1', port=5000)
