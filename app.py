"""
Controle de Produtividade Banho — Pintura Eletrostática
Fluxo: Preparação (1 ou 2 operadores) -> Preencher OPs -> Fila do banho -> Banho -> Concluído

Recursos:
- Grade de 19 cestos com cadeado
- Pausar/retomar o tempo de preparação (café, ginástica) — desconta do total
- Parar o tempo e só depois preencher os dados
- Múltiplas OPs por cesto (botão Adicionar OP)
- 1 ou 2 operadores por cesto (definido ao iniciar)
- Lista mestra do SAP (Excel) com importação otimizada p/ grande volume
- Dashboards (admin e público) + export Excel pré-banho e banho

Banco: PostgreSQL (Railway). Local sem DATABASE_URL -> SQLite.
"""
import os
import io
import csv
import json
import threading
from datetime import datetime, timedelta
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, jsonify, send_file
)
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import create_engine, Column, Integer, String, DateTime, Float, Text, inspect, text
from sqlalchemy.orm import declarative_base, sessionmaker, scoped_session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'troque-esta-chave-em-producao')

DATABASE_URL = os.environ.get('DATABASE_URL', '')
if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)
if not DATABASE_URL:
    DATABASE_URL = 'sqlite:///dados_local.db'

engine = create_engine(DATABASE_URL, pool_pre_ping=True, pool_recycle=280)
Session = scoped_session(sessionmaker(bind=engine))
Base = declarative_base()

TOTAL_CESTOS = 19

# Senha de acesso ao painel público (gerência). Pode trocar por variável de ambiente.
SENHA_PAINEL = os.environ.get('SENHA_PAINEL', 'Decio2026@')

PROCESSOS = [
    "AÇO SEM OXIDAÇÃO", "AÇO COM OXIDAÇÃO", "ALUMÍNIO",
    "MINIMIZADO SEM OXIDAÇÃO", "MINIMIZADO COM OXIDAÇÃO", "INOX",
]

ST_PREPARANDO = 'PREPARANDO'   # cronômetro rodando
ST_PREENCHER = 'PREENCHER'     # tempo parado, aguardando dados
ST_FILA_BANHO = 'FILA_BANHO'   # aguardando banho
ST_EM_BANHO = 'EM_BANHO'       # em banho
ST_CONCLUIDO = 'CONCLUIDO'
ESTADOS_ATIVOS = (ST_PREPARANDO, ST_PREENCHER, ST_FILA_BANHO, ST_EM_BANHO)


# ─────────────────────────────────────────────────────────────────────────────
# Modelos
# ─────────────────────────────────────────────────────────────────────────────
class Usuario(Base):
    __tablename__ = 'usuarios'
    id = Column(Integer, primary_key=True)
    login = Column(String(50), unique=True, nullable=False)
    nome = Column(String(120), nullable=False)
    senha_hash = Column(String(255), nullable=False)
    perfil = Column(String(20), nullable=False)  # admin, prep, banho

    def to_dict(self):
        return {'id': self.id, 'login': self.login, 'nome': self.nome, 'perfil': self.perfil}


class ItemMestre(Base):
    __tablename__ = 'itens_mestre'
    id = Column(Integer, primary_key=True)
    ordem = Column(String(60), unique=True, nullable=False, index=True)  # OP
    material = Column(String(60), default='')                            # código
    texto_breve = Column(String(255), default='')                        # descrição
    quantidade = Column(Integer, default=0)

    def to_dict(self):
        return {'ordem': self.ordem, 'material': self.material,
                'texto_breve': self.texto_breve, 'quantidade': self.quantidade}


class Card(Base):
    __tablename__ = 'cards'
    id = Column(Integer, primary_key=True)
    estado = Column(String(20), nullable=False, index=True)

    numero_cesto = Column(Integer, nullable=False)
    processo = Column(String(60), default='')
    tipo = Column(String(20), default='Normal')

    # 1ª OP nos campos diretos (compatível com Excel/histórico); lista completa em itens_json
    ordem = Column(String(60), default='')
    material = Column(String(60), default='')
    texto_breve = Column(String(255), default='')
    quantidade = Column(Integer, default=0)
    itens_json = Column(Text, default='')
    observacao = Column(Text, default='')

    operador_prep = Column(String(120), default='')    # operador 1
    operador_prep2 = Column(String(120), default='')   # operador 2 (opcional)
    n_operadores = Column(Integer, default=1)
    operador_banho = Column(String(120), default='')

    prep_inicio = Column(DateTime)
    prep_fim = Column(DateTime)
    prep_minutos = Column(Float, default=0)

    pausado = Column(Integer, default=0)
    pausa_inicio = Column(DateTime)
    pausa_acumulada_seg = Column(Integer, default=0)

    banho_inicio = Column(DateTime)
    banho_fim = Column(DateTime)
    banho_minutos = Column(Float, default=0)

    criado_em = Column(DateTime, default=datetime.utcnow)

    def to_dict(self):
        def fmt(dt):
            return (dt - timedelta(hours=3)).strftime('%d/%m/%Y %H:%M:%S') if dt else ''

        def iso(dt):
            return dt.isoformat() + 'Z' if dt else ''
        try:
            itens = json.loads(self.itens_json) if self.itens_json else []
        except (ValueError, TypeError):
            itens = []
        if not itens and self.ordem:
            itens = [{'ordem': self.ordem, 'material': self.material,
                      'texto_breve': self.texto_breve, 'quantidade': self.quantidade}]
        qtd_total = sum(int(i.get('quantidade') or 0) for i in itens) if itens else (self.quantidade or 0)
        # peso e área totais = soma do unitário (por código SAP) × quantidade de cada item
        peso_total = 0.0
        area_total = 0.0
        for it in itens:
            a_unit, p_unit = _area_peso_do_codigo(it.get('material', ''))
            q = int(it.get('quantidade') or 0)
            area_total += a_unit * q
            peso_total += p_unit * q
        return {
            'id': self.id, 'estado': self.estado,
            'numero_cesto': self.numero_cesto,
            'processo': self.processo, 'tipo': self.tipo,
            'ordem': self.ordem, 'material': self.material,
            'texto_breve': self.texto_breve, 'quantidade': self.quantidade,
            'itens': itens, 'qtd_total': qtd_total, 'n_itens': len(itens),
            'peso_total': round(peso_total, 2), 'area_total': round(area_total, 3),
            'observacao': self.observacao or '',
            'operador_prep': self.operador_prep, 'operador_prep2': self.operador_prep2 or '',
            'n_operadores': self.n_operadores or 1,
            'operador_banho': self.operador_banho,
            'prep_inicio': fmt(self.prep_inicio), 'prep_fim': fmt(self.prep_fim),
            'prep_minutos': round(self.prep_minutos or 0, 1),
            'banho_inicio': fmt(self.banho_inicio), 'banho_fim': fmt(self.banho_fim),
            'banho_minutos': round(self.banho_minutos or 0, 1),
            'prep_inicio_iso': iso(self.prep_inicio),
            'prep_fim_iso': iso(self.prep_fim),
            'banho_inicio_iso': iso(self.banho_inicio),
            'pausado': bool(self.pausado),
            'pausa_inicio_iso': iso(self.pausa_inicio),
            'pausa_acumulada_seg': self.pausa_acumulada_seg or 0,
            'data_ref': (self.banho_fim - timedelta(hours=3)).strftime('%Y-%m-%d') if self.banho_fim else '',
        }


# ─────────────────────────────────────────────────────────────────────────────
# Init + migração + seed
# ─────────────────────────────────────────────────────────────────────────────
def _migrar_colunas():
    """Cria colunas novas em tabelas que já existem (deploy sobre banco antigo)."""
    insp = inspect(engine)
    if 'cards' not in insp.get_table_names():
        return
    existentes = {c['name'] for c in insp.get_columns('cards')}
    novas = {
        'itens_json': 'TEXT', 'operador_prep2': "VARCHAR(120) DEFAULT ''",
        'n_operadores': 'INTEGER DEFAULT 1', 'pausado': 'INTEGER DEFAULT 0',
        'pausa_inicio': 'TIMESTAMP NULL', 'pausa_acumulada_seg': 'INTEGER DEFAULT 0',
    }
    with engine.begin() as conn:
        for col, tipo in novas.items():
            if col not in existentes:
                try:
                    conn.execute(text(f'ALTER TABLE cards ADD COLUMN {col} {tipo}'))
                except Exception:
                    pass


def init_db():
    Base.metadata.create_all(engine)
    _migrar_colunas()
    db = Session()
    try:
        if db.query(Usuario).count() == 0:
            seed = [('admin', 'Administrador', 'admin123', 'admin'),
                    ('banho', 'Operador de Banho', 'banho123', 'banho')]
            for i in range(1, 7):
                seed.append((f'op{i}', f'Operador {i}', 'op1234', 'prep'))
            for login, nome, senha, perfil in seed:
                db.add(Usuario(login=login, nome=nome,
                               senha_hash=generate_password_hash(senha), perfil=perfil))
            db.commit()
    finally:
        db.close()


def login_required(*perfis):
    def deco(f):
        @wraps(f)
        def wrapper(*a, **kw):
            if 'usuario' not in session:
                return redirect(url_for('login'))
            if perfis and session.get('perfil') not in perfis and session.get('perfil') != 'admin':
                return redirect(url_for('login'))
            return f(*a, **kw)
        return wrapper
    return deco


def _norm_ordem(v):
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


# ─────────────────────────────────────────────────────────────────────────────
# Lista mestra do SAP e Área/Peso — carregados de arquivos do GitHub p/ memória
# (sem banco: troque o .xlsx no GitHub e dê redeploy para atualizar)
# ─────────────────────────────────────────────────────────────────────────────
_lista_lock = threading.Lock()
_lista_por_ordem = {}
_lista_por_material = {}
_lista_status = {'carregada': False, 'total': 0, 'erro': None}

LISTA_MESTRA_ARQUIVOS = ['lista_mestra.xlsx', 'lista_mestra.csv', 'lista_mestra.txt',
                         'exemplo_lista_mestra_sap.txt']
AREA_PESO_ARQUIVOS = ['area_peso.xlsx', 'area_peso.csv', 'area_peso.txt']

_areapeso_por_sap = {}
_areapeso_status = {'carregada': False, 'total': 0, 'erro': None}


def _norm_str(v):
    if v is None:
        return ''
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def _achar_arquivo(lista_nomes):
    base = os.path.dirname(os.path.abspath(__file__))
    for nome in lista_nomes:
        caminho = os.path.join(base, nome)
        if os.path.isfile(caminho):
            return caminho
    return None


def _achar_colunas(linhas):
    """Detecta as colunas da lista mestra pelo nome no cabeçalho."""
    def norm(s):
        return str(s).strip().lower() if s is not None else ''
    for i, row in enumerate(linhas[:10]):
        if not row:
            continue
        nomes = [norm(c) for c in row]
        idx = {}
        for j, nome in enumerate(nomes):
            if nome == 'ordem' and 'ordem' not in idx:
                idx['ordem'] = j
            elif nome == 'material' and 'material' not in idx:
                idx['material'] = j
            elif 'texto breve' in nome and 'texto' not in idx:
                idx['texto'] = j
            elif ('quantidade da ordem' in nome or nome == 'quantidade total'
                  or nome == 'quantidade') and 'qtd' not in idx:
                idx['qtd'] = j
        if 'ordem' in idx and 'material' in idx:
            return i, idx
    return None


def _parsear_linhas_mestre(linhas):
    achado = _achar_colunas(linhas)
    if achado:
        cab_idx, col = achado
        i_ordem = col.get('ordem', 0)
        i_mat = col.get('material', 1)
        i_texto = col.get('texto')
        i_qtd = col.get('qtd')
        inicio = cab_idx + 1
    else:
        i_ordem, i_mat, i_texto, i_qtd = 0, 2, 3, 4
        inicio = 0

    def val(row, idx):
        if idx is None or idx >= len(row) or row[idx] is None:
            return ''
        return str(row[idx]).strip()

    por_ordem, por_material = {}, {}
    for row in linhas[inicio:]:
        if not row or all(c is None or str(c).strip() == '' for c in row):
            continue
        ordem = _norm_ordem(row[i_ordem]) if i_ordem < len(row) and row[i_ordem] is not None else ''
        if not ordem or not ordem.replace('.', '').isdigit():
            continue
        material = val(row, i_mat)
        texto = val(row, i_texto)
        q = val(row, i_qtd)
        try:
            qtd = int(float(q)) if q else 0
        except (ValueError, TypeError):
            qtd = 0
        item = {'ordem': ordem, 'material': material, 'texto_breve': texto, 'quantidade': qtd}
        por_ordem[ordem] = item
        if material and material not in por_material:
            por_material[material] = item
    return por_ordem, por_material


def _ler_planilha(caminho):
    """Lê xlsx/csv/txt e devolve lista de linhas (cada linha = lista de células)."""
    linhas = []
    nome = caminho.lower()
    if nome.endswith('.csv') or nome.endswith('.txt'):
        with open(caminho, encoding='utf-8-sig', errors='replace') as f:
            raw = f.read()
        sep = '\t' if raw.count('\t') > raw.count(';') and raw.count('\t') > raw.count(',') \
            else (';' if raw.count(';') > raw.count(',') else ',')
        linhas = list(csv.reader(io.StringIO(raw), delimiter=sep))
    else:
        wb = load_workbook(caminho, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            linhas.append(list(row))
    return linhas


def carregar_lista_mestre():
    global _lista_por_ordem, _lista_por_material, _lista_status
    caminho = _achar_arquivo(LISTA_MESTRA_ARQUIVOS)
    if not caminho:
        with _lista_lock:
            _lista_status = {'carregada': False, 'total': 0,
                             'erro': 'lista_mestra.xlsx não encontrada na raiz do projeto.'}
        print('[lista_mestra] AVISO: nenhum arquivo encontrado.')
        return
    try:
        linhas = _ler_planilha(caminho)
        por_ordem, por_material = _parsear_linhas_mestre(linhas)
        with _lista_lock:
            _lista_por_ordem = por_ordem
            _lista_por_material = por_material
            _lista_status = {'carregada': True, 'total': len(por_ordem), 'erro': None}
        print(f'[lista_mestra] Carregada: {len(por_ordem)} ordens de "{os.path.basename(caminho)}".')
    except Exception as e:
        with _lista_lock:
            _lista_status = {'carregada': False, 'total': 0, 'erro': str(e)}
        print(f'[lista_mestra] ERRO: {e}')


def carregar_area_peso():
    """Carrega area_peso (Código SAP -> área m² e peso kg unitários)."""
    global _areapeso_por_sap, _areapeso_status
    caminho = _achar_arquivo(AREA_PESO_ARQUIVOS)
    if not caminho:
        _areapeso_status = {'carregada': False, 'total': 0, 'erro': 'area_peso não encontrado.'}
        print('[area_peso] AVISO: nenhum arquivo encontrado.')
        return
    try:
        linhas = _ler_planilha(caminho)

        def norm(s):
            return str(s).strip().lower() if s is not None else ''
        i_sap = i_area = i_peso = None
        inicio = 0
        for i, row in enumerate(linhas[:10]):
            if not row:
                continue
            for j, c in enumerate(row):
                n = norm(c)
                if 'codigo sap' in n or 'código sap' in n:
                    i_sap = j
                elif 'area' in n or 'área' in n:
                    i_area = j
                elif 'peso' in n:
                    i_peso = j
            if i_sap is not None:
                inicio = i + 1
                break
        if i_sap is None:
            i_sap, i_area, i_peso, inicio = 3, 1, 2, 1

        mapa = {}
        for row in linhas[inicio:]:
            if not row or all(c is None or str(c).strip() == '' for c in row):
                continue
            sap = _norm_str(row[i_sap]) if i_sap is not None and i_sap < len(row) else ''
            if not sap:
                continue

            def num(idx):
                try:
                    return float(row[idx]) if idx is not None and idx < len(row) and row[idx] not in (None, '') else 0.0
                except (ValueError, TypeError):
                    return 0.0
            mapa[sap] = {'area_m2': num(i_area) / 1_000_000.0, 'peso_kg': num(i_peso)}
        _areapeso_por_sap = mapa
        _areapeso_status = {'carregada': True, 'total': len(mapa), 'erro': None}
        print(f'[area_peso] Carregada: {len(mapa)} códigos de "{os.path.basename(caminho)}".')
    except Exception as e:
        _areapeso_status = {'carregada': False, 'total': 0, 'erro': str(e)}
        print(f'[area_peso] ERRO: {e}')


def _area_peso_do_codigo(material):
    """Retorna (area_m2, peso_kg) unitários para um código SAP, ou (0,0)."""
    d = _areapeso_por_sap.get(_norm_str(material))
    if d:
        return d['area_m2'], d['peso_kg']
    return 0.0, 0.0





@app.route('/login', methods=['GET', 'POST'])
def login():
    erro = None
    if request.method == 'POST':
        login_u = request.form.get('usuario', '').strip()
        senha = request.form.get('senha', '')
        db = Session()
        try:
            u = db.query(Usuario).filter_by(login=login_u).first()
            if u and check_password_hash(u.senha_hash, senha):
                session['usuario'] = u.login
                session['nome'] = u.nome
                session['perfil'] = u.perfil
                destino = {'admin': 'dashboard', 'banho': 'tela_banho',
                           'prep': 'tela_prep'}.get(u.perfil, 'login')
                return redirect(url_for(destino))
            erro = 'Usuário ou senha incorretos.'
        finally:
            db.close()
    return render_template('login.html', erro=erro)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


def _lista_operadores_prep():
    db = Session()
    try:
        return [u.nome for u in db.query(Usuario)
                .filter(Usuario.perfil.in_(('prep', 'banho', 'admin')))
                .order_by(Usuario.nome).all()]
    finally:
        db.close()


@app.route('/preparacao')
@login_required('prep', 'banho')
def tela_prep():
    return render_template('prep.html', nome=session.get('nome'),
                           perfil=session.get('perfil'), processos=PROCESSOS,
                           operadores=_lista_operadores_prep())


@app.route('/banho')
@login_required('banho')
def tela_banho():
    return render_template('banho.html', nome=session.get('nome'), perfil=session.get('perfil'))


@app.route('/dashboard')
@login_required('admin')
def dashboard():
    return render_template('dashboard.html', nome=session.get('nome'), processos=PROCESSOS)


@app.route('/painel', methods=['GET', 'POST'])
def painel_publico():
    # Painel público protegido por senha simples (compartilhada com a gerência)
    if request.method == 'POST':
        if request.form.get('senha', '') == SENHA_PAINEL:
            session['painel_ok'] = True
            return redirect(url_for('painel_publico'))
        return render_template('painel_senha.html', erro='Senha incorreta.')
    if not session.get('painel_ok'):
        return render_template('painel_senha.html', erro=None)
    return render_template('painel.html', processos=PROCESSOS)


@app.route('/admin/usuarios', methods=['GET', 'POST'])
@login_required('admin')
def admin_usuarios():
    db = Session()
    msg = None
    try:
        if request.method == 'POST':
            acao = request.form.get('acao')
            if acao == 'adicionar':
                nu = request.form.get('novo_usuario', '').strip()
                nn = request.form.get('novo_nome', '').strip()
                ns = request.form.get('nova_senha', '')
                npf = request.form.get('novo_perfil', 'prep')
                if nu and nn and ns and not db.query(Usuario).filter_by(login=nu).first():
                    db.add(Usuario(login=nu, nome=nn,
                                   senha_hash=generate_password_hash(ns), perfil=npf))
                    db.commit()
                    msg = f'Usuário {nn} adicionado.'
                elif db.query(Usuario).filter_by(login=nu).first():
                    msg = 'Esse login já existe.'
            elif acao == 'remover':
                u = db.query(Usuario).filter_by(login=request.form.get('usuario_remover')).first()
                if u and u.login != 'admin':
                    db.delete(u)
                    db.commit()
                    msg = 'Usuário removido.'
            elif acao == 'senha':
                u = db.query(Usuario).filter_by(login=request.form.get('usuario_senha')).first()
                nova = request.form.get('senha_nova', '')
                if u and nova:
                    u.senha_hash = generate_password_hash(nova)
                    db.commit()
                    msg = f'Senha de {u.nome} atualizada.'
        usuarios = [u.to_dict() for u in db.query(Usuario).order_by(Usuario.id).all()]
        return render_template('usuarios.html', usuarios=usuarios, nome=session.get('nome'), msg=msg)
    finally:
        db.close()


@app.route('/admin/mestre', methods=['GET', 'POST'])
@login_required('admin')
def admin_mestre():
    msg = None
    # Botão "Recarregar da memória": relê os arquivos do projeto
    if request.method == 'POST' and request.form.get('acao') == 'recarregar':
        carregar_lista_mestre()
        carregar_area_peso()
        msg = 'Lista mestra e área/peso recarregadas dos arquivos.'

    with _lista_lock:
        status = dict(_lista_status)
        amostra = list(_lista_por_ordem.values())[:25]
        total = status.get('total', 0)
    arquivo_lista = os.path.basename(_achar_arquivo(LISTA_MESTRA_ARQUIVOS) or '') or 'não encontrado'
    arquivo_ap = os.path.basename(_achar_arquivo(AREA_PESO_ARQUIVOS) or '') or 'não encontrado'
    return render_template('mestre.html', nome=session.get('nome'),
                           msg=msg, total_itens=total, amostra=amostra,
                           status=status, arquivo_info=arquivo_lista,
                           areapeso_status=_areapeso_status, areapeso_arquivo=arquivo_ap)


def _admin_mestre_antigo_desativado():
    db = Session()
    msg = None
    try:
        # Garante que a tabela exista (deploy sobre banco antigo pode não tê-la)
        try:
            Base.metadata.create_all(engine)
        except Exception:
            pass

        if request.method == 'POST':
            f = request.files.get('arquivo')
            if f and f.filename:
                nome = f.filename.lower()
                try:
                    linhas = []
                    if nome.endswith('.csv') or nome.endswith('.txt'):
                        raw = f.stream.read().decode('utf-8-sig', errors='replace')
                        sep = '\t' if raw.count('\t') > raw.count(';') and raw.count('\t') > raw.count(',') \
                            else (';' if raw.count(';') > raw.count(',') else ',')
                        linhas = list(csv.reader(io.StringIO(raw), delimiter=sep))
                    else:
                        wb = load_workbook(f, read_only=True, data_only=True)
                        ws = wb.active
                        for row in ws.iter_rows(values_only=True):
                            linhas.append(list(row))
                    novos, atual = importar_mestre(db, linhas)
                    msg = f'Importado: {novos} novas ordens, {atual} atualizadas.'
                except Exception as e:
                    db.rollback()
                    msg = f'Erro ao importar: {e}'
        try:
            total = db.query(ItemMestre).count()
            amostra = [i.to_dict() for i in db.query(ItemMestre).limit(25).all()]
        except Exception as e:
            db.rollback()
            total, amostra = 0, []
            if not msg:
                msg = f'Banco em preparação ({e}). Tente importar o relatório.'
        return render_template('mestre.html', nome=session.get('nome'),
                               msg=msg, total_itens=total, amostra=amostra)
    finally:
        db.close()


def _serial_val(v):
    if isinstance(v, datetime):
        return v.isoformat() + 'Z'
    return v


@app.route('/api/admin/db_status')
@login_required('admin')
def api_db_status():
    """Mostra se o banco é PostgreSQL (seguro) ou SQLite (temporário)."""
    tipo = 'postgresql' if DATABASE_URL.startswith('postgresql') else 'sqlite'
    db = Session()
    try:
        n_cards = db.query(Card).count()
        n_users = db.query(Usuario).count()
        n_mestre = db.query(ItemMestre).count()
    except Exception:
        n_cards = n_users = n_mestre = -1
    finally:
        db.close()
    return jsonify({'tipo': tipo, 'seguro': tipo == 'postgresql',
                    'cards': n_cards, 'usuarios': n_users, 'mestre': n_mestre})


@app.route('/api/admin/backup')
@login_required('admin')
def api_admin_backup():
    """Baixa backup completo (cards, usuários e lista mestra) em JSON."""
    db = Session()
    try:
        def full(obj, model):
            return {c.name: _serial_val(getattr(obj, c.name)) for c in model.__table__.columns}
        dados = {
            'versao': 1,
            'gerado_em': datetime.utcnow().isoformat() + 'Z',
            'cards': [full(c, Card) for c in db.query(Card).all()],
            'usuarios': [full(u, Usuario) for u in db.query(Usuario).all()],
            'itens_mestre': [full(i, ItemMestre) for i in db.query(ItemMestre).all()],
        }
        buf = io.BytesIO(json.dumps(dados, ensure_ascii=False, indent=2).encode('utf-8'))
        buf.seek(0)
        stamp = datetime.now().strftime('%Y%m%d_%H%M')
        return send_file(buf, as_attachment=True,
                         download_name=f'backup_banho_{stamp}.json',
                         mimetype='application/json')
    finally:
        db.close()


@app.route('/api/admin/restaurar', methods=['POST'])
@login_required('admin')
def api_admin_restaurar():
    """Restaura backup JSON. Só ADICIONA o que não existe — nunca apaga."""
    f = request.files.get('arquivo')
    if not f or not f.filename:
        return jsonify({'sucesso': False, 'erro': 'Envie o arquivo de backup.'}), 400
    try:
        dados = json.loads(f.stream.read().decode('utf-8-sig', errors='replace'))
    except Exception as e:
        return jsonify({'sucesso': False, 'erro': f'Arquivo inválido: {e}'}), 400
    db = Session()
    rc = ru = rm = 0
    try:
        def set_cols(obj, model, src):
            for col in model.__table__.columns:
                if col.name in src:
                    val = src[col.name]
                    if isinstance(col.type, DateTime) and val:
                        try:
                            val = datetime.fromisoformat(str(val).replace('Z', ''))
                        except (ValueError, TypeError):
                            val = None
                    setattr(obj, col.name, val)

        logins = {u.login for u in db.query(Usuario).all()}
        for u in dados.get('usuarios', []):
            if u.get('login') and u['login'] not in logins:
                novo = Usuario(login=u['login'], nome=u.get('nome', ''),
                               senha_hash=u.get('senha_hash', ''), perfil=u.get('perfil', 'prep'))
                db.add(novo); ru += 1
        ids = {c.id for c in db.query(Card.id).all()}
        for cd in dados.get('cards', []):
            if cd.get('id') and cd['id'] not in ids:
                novo = Card(); set_cols(novo, Card, cd); db.add(novo); rc += 1
        ordens = {i.ordem for i in db.query(ItemMestre).all()}
        for it in dados.get('itens_mestre', []):
            if it.get('ordem') and it['ordem'] not in ordens:
                novo = ItemMestre(); set_cols(novo, ItemMestre, it); db.add(novo); rm += 1
        db.commit()
        return jsonify({'sucesso': True, 'cards': rc, 'usuarios': ru, 'mestre': rm})
    except Exception as e:
        db.rollback()
        return jsonify({'sucesso': False, 'erro': str(e)}), 500
    finally:
        db.close()


@app.route('/api/admin/testar_op/<path:ordem>')
@login_required('admin')
def api_admin_testar_op(ordem):
    db = Session()
    try:
        item = db.query(ItemMestre).filter_by(ordem=_norm_ordem(ordem)).first()
        if item:
            return jsonify({'encontrado': True, **item.to_dict()})
        return jsonify({'encontrado': False, 'ordem': _norm_ordem(ordem)})
    finally:
        db.close()


def importar_mestre(db, linhas):
    """Otimizado p/ grande volume (17 mil+): 1 consulta inicial + inserção em lote."""
    existentes = {o.ordem: o for o in db.query(ItemMestre).all()}
    novos = atual = 0
    novos_objs = []
    vistos = set()
    for row in linhas:
        if not row or all(c is None or str(c).strip() == '' for c in row):
            continue
        c0 = str(row[0]).strip().lower()
        if 'ordem' in c0 or c0 in ('order',):
            continue
        ordem = _norm_ordem(row[0])
        if not ordem or not ordem.replace('.', '').isdigit():
            continue
        material = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ''
        texto = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ''
        try:
            qtd = int(float(row[4])) if len(row) > 4 and row[4] not in (None, '') else 0
        except (ValueError, TypeError):
            qtd = 0
        if ordem in existentes:
            ex = existentes[ordem]
            ex.material, ex.texto_breve, ex.quantidade = material, texto, qtd
            atual += 1
        elif ordem not in vistos:
            novos_objs.append(ItemMestre(ordem=ordem, material=material,
                                         texto_breve=texto, quantidade=qtd))
            vistos.add(ordem)
            novos += 1
        if len(novos_objs) >= 1000:
            db.bulk_save_objects(novos_objs)
            db.commit()
            novos_objs = []
    if novos_objs:
        db.bulk_save_objects(novos_objs)
    db.commit()
    return novos, atual


# ─────────────────────────────────────────────────────────────────────────────
# APIs — grade e fluxo
# ─────────────────────────────────────────────────────────────────────────────
@app.route('/api/cestos')
@login_required('prep', 'banho')
def api_cestos():
    db = Session()
    try:
        ativos = db.query(Card).filter(Card.estado.in_(ESTADOS_ATIVOS)).all()
        mapa = {c.numero_cesto: c for c in ativos}
        grade = [{'numero': n, 'ocupado': n in mapa,
                  'card': mapa[n].to_dict() if n in mapa else None}
                 for n in range(1, TOTAL_CESTOS + 1)]
        return jsonify(grade)
    finally:
        db.close()


@app.route('/api/buscar_ordem/<path:ordem>')
@login_required('prep', 'banho')
def api_buscar_ordem(ordem):
    o = _norm_ordem(ordem)
    with _lista_lock:
        item = _lista_por_ordem.get(o)
    if item:
        return jsonify({'encontrado': True, **item})
    return jsonify({'encontrado': False, 'ordem': o})


@app.route('/api/buscar_codigo/<path:codigo>')
@login_required('prep', 'banho')
def api_buscar_codigo(codigo):
    """Busca a descrição de um item pelo código (Material), quando não há OP."""
    cod = _norm_str(codigo)
    with _lista_lock:
        item = _lista_por_material.get(cod)
    if item:
        return jsonify({'encontrado': True, 'material': cod,
                        'texto_breve': item.get('texto_breve', '')})
    return jsonify({'encontrado': False, 'material': cod})


@app.route('/api/prep/iniciar', methods=['POST'])
@login_required('prep', 'banho')
def api_prep_iniciar():
    d = request.json or {}
    try:
        numero = int(d.get('numero_cesto'))
    except (ValueError, TypeError):
        return jsonify({'sucesso': False, 'erro': 'Cesto inválido.'}), 400
    if not (1 <= numero <= TOTAL_CESTOS):
        return jsonify({'sucesso': False, 'erro': 'Cesto fora do intervalo.'}), 400
    db = Session()
    try:
        if db.query(Card).filter(Card.numero_cesto == numero,
                                 Card.estado.in_(ESTADOS_ATIVOS)).first():
            return jsonify({'sucesso': False, 'erro': f'Cesto {numero} já está em uso.'}), 400
        try:
            n_op = 2 if int(d.get('n_operadores', 1)) == 2 else 1
        except (ValueError, TypeError):
            n_op = 1
        op2 = (d.get('operador_prep2') or '').strip() if n_op == 2 else ''
        card = Card(estado=ST_PREPARANDO, numero_cesto=numero,
                    operador_prep=session.get('nome', ''),
                    operador_prep2=op2, n_operadores=n_op,
                    prep_inicio=datetime.utcnow())
        db.add(card)
        db.commit()
        return jsonify({'sucesso': True, 'id': card.id})
    finally:
        db.close()


@app.route('/api/prep/pausar', methods=['POST'])
@login_required('prep', 'banho')
def api_prep_pausar():
    d = request.json or {}
    db = Session()
    try:
        card = db.query(Card).get(int(d.get('id', 0)))
        if not card or card.estado != ST_PREPARANDO:
            return jsonify({'sucesso': False, 'erro': 'Cesto não está em preparação.'}), 404
        agora = datetime.utcnow()
        if card.pausado:
            if card.pausa_inicio:
                card.pausa_acumulada_seg = (card.pausa_acumulada_seg or 0) + \
                    int((agora - card.pausa_inicio).total_seconds())
            card.pausado = 0
            card.pausa_inicio = None
        else:
            card.pausado = 1
            card.pausa_inicio = agora
        db.commit()
        return jsonify({'sucesso': True, 'pausado': bool(card.pausado)})
    finally:
        db.close()


@app.route('/api/prep/parar', methods=['POST'])
@login_required('prep', 'banho')
def api_prep_parar():
    d = request.json or {}
    db = Session()
    try:
        card = db.query(Card).get(int(d.get('id', 0)))
        if not card or card.estado != ST_PREPARANDO:
            return jsonify({'sucesso': False, 'erro': 'Cesto não está em preparação.'}), 404
        agora = datetime.utcnow()
        if card.pausado and card.pausa_inicio:
            card.pausa_acumulada_seg = (card.pausa_acumulada_seg or 0) + \
                int((agora - card.pausa_inicio).total_seconds())
            card.pausado = 0
            card.pausa_inicio = None
        card.prep_fim = agora
        bruto = (card.prep_fim - card.prep_inicio).total_seconds()
        card.prep_minutos = round(max(0, bruto - (card.pausa_acumulada_seg or 0)) / 60, 1)
        card.estado = ST_PREENCHER
        db.commit()
        return jsonify({'sucesso': True, 'prep_minutos': card.prep_minutos})
    finally:
        db.close()


def _salvar_itens(card, d):
    """Recebe lista de itens (OPs) e grava em itens_json + campos diretos (1ª OP)."""
    itens = d.get('itens')
    if itens is None:
        # compatibilidade: item único vindo de campos soltos
        itens = [{'ordem': d.get('ordem', ''), 'material': d.get('material', ''),
                  'texto_breve': d.get('texto_breve', ''), 'quantidade': d.get('quantidade', 0)}]
    norm = []
    for it in itens:
        ordem = _norm_ordem(it.get('ordem', '')) if it.get('ordem') else ''
        if not ordem and not it.get('material'):
            continue
        try:
            q = int(it.get('quantidade') or 0)
        except (ValueError, TypeError):
            q = 0
        norm.append({'ordem': ordem, 'material': (it.get('material') or '').strip(),
                     'texto_breve': (it.get('texto_breve') or '').strip(), 'quantidade': q})
    card.itens_json = json.dumps(norm, ensure_ascii=False)
    if norm:
        card.ordem = norm[0]['ordem']
        card.material = norm[0]['material']
        card.texto_breve = norm[0]['texto_breve']
        card.quantidade = sum(i['quantidade'] for i in norm)  # qtd total do cesto


def _aplicar_meta(card, d):
    for campo in ('processo', 'tipo', 'observacao'):
        if campo in d:
            setattr(card, campo, (d.get(campo) or '').strip())


@app.route('/api/prep/finalizar', methods=['POST'])
@login_required('prep', 'banho')
def api_prep_finalizar():
    d = request.json or {}
    db = Session()
    try:
        card = db.query(Card).get(int(d.get('id', 0)))
        if not card or card.estado not in (ST_PREENCHER, ST_PREPARANDO):
            return jsonify({'sucesso': False, 'erro': 'Card não encontrado.'}), 404
        if card.estado == ST_PREPARANDO:
            agora = datetime.utcnow()
            if card.pausado and card.pausa_inicio:
                card.pausa_acumulada_seg = (card.pausa_acumulada_seg or 0) + \
                    int((agora - card.pausa_inicio).total_seconds())
                card.pausado = 0
                card.pausa_inicio = None
            card.prep_fim = agora
            bruto = (card.prep_fim - card.prep_inicio).total_seconds()
            card.prep_minutos = round(max(0, bruto - (card.pausa_acumulada_seg or 0)) / 60, 1)
        _aplicar_meta(card, d)
        _salvar_itens(card, d)
        card.estado = ST_FILA_BANHO
        db.commit()
        return jsonify({'sucesso': True})
    finally:
        db.close()


@app.route('/api/card/editar', methods=['POST'])
@login_required('prep', 'banho')
def api_card_editar():
    d = request.json or {}
    db = Session()
    try:
        card = db.query(Card).get(int(d.get('id', 0)))
        if not card:
            return jsonify({'sucesso': False, 'erro': 'Card não encontrado.'}), 404
        _aplicar_meta(card, d)
        if 'itens' in d:
            _salvar_itens(card, d)
        if d.get('prep_minutos') not in (None, ''):
            try:
                card.prep_minutos = round(float(d.get('prep_minutos')), 1)
            except (ValueError, TypeError):
                pass
        # editar operadores
        if 'n_operadores' in d:
            try:
                card.n_operadores = 2 if int(d.get('n_operadores')) == 2 else 1
            except (ValueError, TypeError):
                pass
            card.operador_prep2 = (d.get('operador_prep2') or '').strip() if card.n_operadores == 2 else ''
        db.commit()
        return jsonify({'sucesso': True})
    finally:
        db.close()


@app.route('/api/banho/fila')
@login_required('banho')
def api_banho_fila():
    db = Session()
    try:
        fila = db.query(Card).filter_by(estado=ST_FILA_BANHO).order_by(Card.prep_fim).all()
        emb = db.query(Card).filter_by(estado=ST_EM_BANHO).order_by(Card.banho_inicio).all()
        return jsonify({'fila': [c.to_dict() for c in fila],
                        'em_banho': [c.to_dict() for c in emb]})
    finally:
        db.close()


@app.route('/api/banho/iniciar', methods=['POST'])
@login_required('banho')
def api_banho_iniciar():
    d = request.json or {}
    db = Session()
    try:
        card = db.query(Card).get(int(d.get('id', 0)))
        if not card or card.estado != ST_FILA_BANHO:
            return jsonify({'sucesso': False, 'erro': 'Card não está na fila.'}), 404
        card.banho_inicio = datetime.utcnow()
        card.operador_banho = session.get('nome', '')
        card.estado = ST_EM_BANHO
        db.commit()
        return jsonify({'sucesso': True})
    finally:
        db.close()


@app.route('/api/banho/finalizar', methods=['POST'])
@login_required('banho')
def api_banho_finalizar():
    d = request.json or {}
    db = Session()
    try:
        card = db.query(Card).get(int(d.get('id', 0)))
        if not card or card.estado != ST_EM_BANHO:
            return jsonify({'sucesso': False, 'erro': 'Card não está em banho.'}), 404
        card.banho_fim = datetime.utcnow()
        card.banho_minutos = round((card.banho_fim - card.banho_inicio).total_seconds() / 60, 1)
        card.estado = ST_CONCLUIDO
        db.commit()
        return jsonify({'sucesso': True, 'banho_minutos': card.banho_minutos})
    finally:
        db.close()


@app.route('/api/agora')
def api_agora():
    """Hora do servidor (UTC) p/ sincronizar cronômetros e começar do 0:00."""
    return jsonify({'agora_iso': datetime.utcnow().isoformat() + 'Z'})


@app.route('/api/cesto/mudar_numero', methods=['POST'])
@login_required('prep', 'banho')
def api_cesto_mudar_numero():
    """Corrige o número de um cesto cadastrado errado."""
    d = request.json or {}
    db = Session()
    try:
        card = db.query(Card).get(int(d.get('id', 0)))
        if not card:
            return jsonify({'sucesso': False, 'erro': 'Cesto não encontrado.'}), 404
        try:
            novo = int(d.get('numero_cesto'))
        except (ValueError, TypeError):
            return jsonify({'sucesso': False, 'erro': 'Número inválido.'}), 400
        if not (1 <= novo <= TOTAL_CESTOS):
            return jsonify({'sucesso': False, 'erro': f'O número deve ser entre 1 e {TOTAL_CESTOS}.'}), 400
        ocupado = db.query(Card).filter(Card.numero_cesto == novo,
                                        Card.estado.in_(ESTADOS_ATIVOS),
                                        Card.id != card.id).first()
        if ocupado:
            return jsonify({'sucesso': False, 'erro': f'O cesto {novo} já está em uso.'}), 400
        card.numero_cesto = novo
        db.commit()
        return jsonify({'sucesso': True})
    finally:
        db.close()


@app.route('/api/cesto/cancelar', methods=['POST'])
@login_required('prep', 'banho')
def api_cesto_cancelar():
    """Cancela (remove) um cesto cadastrado errado. Não cancela concluído."""
    d = request.json or {}
    db = Session()
    try:
        card = db.query(Card).get(int(d.get('id', 0)))
        if not card:
            return jsonify({'sucesso': False, 'erro': 'Cesto não encontrado.'}), 404
        if card.estado == ST_CONCLUIDO:
            return jsonify({'sucesso': False, 'erro': 'Não é possível cancelar um cesto concluído.'}), 400
        db.delete(card)
        db.commit()
        return jsonify({'sucesso': True})
    finally:
        db.close()


# ─────────────────────────────────────────────────────────────────────────────
# Dados dos dashboards
# ─────────────────────────────────────────────────────────────────────────────
def _coletar_dados(de=None, ate=None):
    db = Session()
    try:
        cards = db.query(Card).filter_by(estado=ST_CONCLUIDO).all()

        def dentro(c):
            if not c.banho_fim:
                return False
            dia = (c.banho_fim - timedelta(hours=3)).date()
            if de and dia < de:
                return False
            if ate and dia > ate:
                return False
            return True
        cards = [c for c in cards if dentro(c)]
        ativos = db.query(Card).filter(Card.estado.in_(ESTADOS_ATIVOS)).order_by(Card.id.desc()).all()
        normais = sum(1 for c in cards if c.tipo == 'Normal')
        retrab = sum(1 for c in cards if c.tipo == 'Retrabalho')
        tp = [c.prep_minutos for c in cards if c.prep_minutos]
        tb = [c.banho_minutos for c in cards if c.banho_minutos]
        por_proc, por_dia = {}, {}
        peso_total_geral = 0.0
        area_total_geral = 0.0
        pecas_total_geral = 0
        for c in cards:
            p = c.processo or 'Sem processo'
            por_proc[p] = por_proc.get(p, 0) + 1
            dia = (c.banho_fim - timedelta(hours=3)).strftime('%d/%m')
            por_dia[dia] = por_dia.get(dia, 0) + 1
            dd = c.to_dict()
            peso_total_geral += dd['peso_total']
            area_total_geral += dd['area_total']
            pecas_total_geral += dd['qtd_total']
        return {
            'total': len(cards), 'normais': normais, 'retrabalhos': retrab,
            'em_andamento': len(ativos),
            'media_prep': round(sum(tp) / len(tp), 1) if tp else 0,
            'media_banho': round(sum(tb) / len(tb), 1) if tb else 0,
            'por_processo': por_proc, 'por_dia': por_dia,
            'peso_total_geral': round(peso_total_geral, 1),
            'area_total_geral': round(area_total_geral, 2),
            'pecas_total_geral': pecas_total_geral,
            'ativos': [c.to_dict() for c in ativos],
            'registros': [c.to_dict() for c in sorted(cards, key=lambda x: x.id, reverse=True)[:200]],
        }
    finally:
        db.close()


def _parse_datas():
    def pd(s):
        try:
            return datetime.strptime(s, '%Y-%m-%d').date()
        except (ValueError, TypeError):
            return None
    return pd(request.args.get('de')), pd(request.args.get('ate'))


@app.route('/api/dashboard/dados')
@login_required('admin')
def api_dashboard_dados():
    de, ate = _parse_datas()
    return jsonify(_coletar_dados(de, ate))


@app.route('/api/painel/dados')
def api_painel_dados():
    de, ate = _parse_datas()
    return jsonify(_coletar_dados(de, ate))


# ─────────────────────────────────────────────────────────────────────────────
# Export Excel (uma linha por OP, para detalhar cestos com várias OPs)
# ─────────────────────────────────────────────────────────────────────────────
def _estilo_cabecalho(ws, headers):
    fill = PatternFill("solid", fgColor="0F3D5C")
    font = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style='thin', color='D0D7DE')
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(bottom=thin)
    ws.row_dimensions[1].height = 28


def _gerar_excel(tipo):
    db = Session()
    try:
        cards = db.query(Card).filter_by(estado=ST_CONCLUIDO).order_by(Card.id).all()
        wb = Workbook()
        ws = wb.active
        if tipo == 'prebanho':
            ws.title = 'Pre-Banho'
            headers = ['ID', 'Cesto', 'OP (Ordem)', 'Código', 'Texto breve', 'Qtd',
                       'Área (m²)', 'Peso (kg)',
                       'Processo', 'Tipo', 'Operador 1', 'Operador 2', 'Nº oper.',
                       'Início', 'Fim', 'Tempo prep (min)', 'Observação']
            larg = [6, 7, 14, 14, 30, 7, 12, 12, 22, 12, 16, 16, 9, 19, 19, 13, 28]
        else:
            ws.title = 'Banho'
            headers = ['ID', 'Cesto', 'OP (Ordem)', 'Código', 'Texto breve', 'Qtd',
                       'Área (m²)', 'Peso (kg)',
                       'Processo', 'Tipo', 'Operador banho',
                       'Início banho', 'Fim banho', 'Tempo banho (min)']
            larg = [6, 7, 14, 14, 30, 7, 12, 12, 22, 12, 18, 19, 19, 14]
        _estilo_cabecalho(ws, headers)
        for c in cards:
            dd = c.to_dict()
            itens = dd['itens'] or [{'ordem': dd['ordem'], 'material': dd['material'],
                                     'texto_breve': dd['texto_breve'], 'quantidade': dd['quantidade']}]
            for it in itens:  # uma linha por OP
                a_unit, p_unit = _area_peso_do_codigo(it.get('material', ''))
                q_it = int(it.get('quantidade') or 0)
                area_it = round(a_unit * q_it, 3)
                peso_it = round(p_unit * q_it, 2)
                if tipo == 'prebanho':
                    ws.append([dd['id'], dd['numero_cesto'], it['ordem'], it['material'],
                               it['texto_breve'], it['quantidade'], area_it, peso_it,
                               dd['processo'], dd['tipo'],
                               dd['operador_prep'], dd['operador_prep2'], dd['n_operadores'],
                               dd['prep_inicio'], dd['prep_fim'], dd['prep_minutos'], dd['observacao']])
                else:
                    ws.append([dd['id'], dd['numero_cesto'], it['ordem'], it['material'],
                               it['texto_breve'], it['quantidade'], area_it, peso_it,
                               dd['processo'], dd['tipo'],
                               dd['operador_banho'], dd['banho_inicio'], dd['banho_fim'], dd['banho_minutos']])
        for i, w in enumerate(larg, 1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
        ws.freeze_panes = 'A2'
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf
    finally:
        db.close()


@app.route('/api/download/prebanho')
@login_required('admin')
def download_prebanho():
    buf = _gerar_excel('prebanho')
    stamp = datetime.now().strftime('%Y%m%d_%H%M')
    return send_file(buf, as_attachment=True, download_name=f'prebanho_{stamp}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/download/banho')
@login_required('admin')
def download_banho():
    buf = _gerar_excel('banho')
    stamp = datetime.now().strftime('%Y%m%d_%H%M')
    return send_file(buf, as_attachment=True, download_name=f'banho_{stamp}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.teardown_appcontext
def remove_session(exc=None):
    Session.remove()


init_db()
carregar_lista_mestre()
carregar_area_peso()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
