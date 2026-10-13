import os
import glob
import time
import uuid
import asyncio
import json
import unicodedata
import logging
import warnings
import threading
import urllib.parse
from datetime import datetime, timedelta

from flask import Flask, render_template, jsonify, request, send_from_directory
from werkzeug.middleware.proxy_fix import ProxyFix
import pandas as pd
import edge_tts
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

warnings.simplefilter(action='ignore', category=UserWarning)
warnings.filterwarnings("ignore", module="openpyxl")
pd.set_option('future.no_silent_downcasting', True)
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
METAS_REPAROS_PATH = os.path.join(BASE_DIR, 'metas_reparos.json')
metas_reparos_lock = threading.Lock()
app = Flask(__name__, 
            template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=os.path.join(BASE_DIR, 'static'))

app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

@app.after_request
def add_no_cache_header(response):
    if request.path in ('/sw.js', '/static/manifest.webmanifest'):
        response.headers['Cache-Control'] = 'no-cache'
    elif request.path.startswith('/static/'):
        response.headers['Cache-Control'] = 'public, max-age=86400'
    else:
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    return response

# --- ESTADO GLOBAL DO SERVIÇO ---
class AppState:
    def __init__(self):
        for f in glob.glob("temp_voice_*.mp3"):
            try: os.remove(f)
            except: pass

        self.driver = None
        self.df_hierarquia = None
        self.df_producao_raw = None
        self.df_reparos_raw = None
        self.mapa_controladores = {}
        
        self.status_msg = "Aguardando inicialização..."
        self.bloqueio_processamento = False
        self.ultima_atualizacao = None
        self.ultima_fala = {}
        
        self.audio_pendente = None
        self.config = {
            "rotacao_ativa": True,
            "tempo_tela": 30,
            "auto_refresh": True,
            "intervalo_refresh_min": 5,
            "voz_ativa": True,
            "intervalo_voz_min": 15
        }

state = AppState()

def carregar_metas_reparos():
    """Carrega metas persistidas por regional/coordenador sem derrubar o painel se o JSON estiver inválido."""
    if not os.path.exists(METAS_REPAROS_PATH):
        return {}
    try:
        with metas_reparos_lock, open(METAS_REPAROS_PATH, 'r', encoding='utf-8') as arquivo:
            dados = json.load(arquivo)
        return dados if isinstance(dados, dict) else {}
    except Exception as e:
        logging.error(f"Erro ao carregar metas de reparos: {e}")
        return {}

def salvar_metas_reparos(dados):
    """Grava de forma atômica para evitar um arquivo parcial durante atualizações concorrentes."""
    temporario = f"{METAS_REPAROS_PATH}.tmp"
    with metas_reparos_lock:
        with open(temporario, 'w', encoding='utf-8') as arquivo:
            json.dump(dados, arquivo, ensure_ascii=False, indent=2, sort_keys=True)
        os.replace(temporario, METAS_REPAROS_PATH)

def normalizar_chave_meta(valor):
    return str(valor or '').strip().upper()

def aplicar_filtros_hierarquia_disponiveis(df, gerentes=None, supervisores=None):
    """Aplica apenas filtros hierárquicos realmente presentes na fonte atual.

    Relatórios de tarefas abertas podem não trazer recurso/hierarquia. Nessa
    situação, filtrar por supervisor produziria um falso painel zerado.
    """
    resultado = df
    ignorados = []
    pares = [
        ('coordenação', 'Temp_Gerente', gerentes or []),
        ('supervisão', 'Temp_Sup', supervisores or [])
    ]
    for nome, coluna, selecionados in pares:
        if not selecionados:
            continue
        disponiveis = set()
        if coluna in resultado.columns:
            disponiveis = set(resultado[coluna].fillna('').astype(str).str.strip()) - {''}
        valores_validos = [valor for valor in selecionados if valor in disponiveis]
        if valores_validos:
            resultado = resultado[resultado[coluna].isin(valores_validos)]
        else:
            ignorados.append(nome)
    return resultado, ignorados

# --- CARREGAMENTO LOCAL EXCEL ---
def carregar_controladores_local():
    try:
        for arquivo in glob.glob("*.xlsx"):
            if "~$" in arquivo: continue
            try:
                df_scan = pd.read_excel(arquivo, nrows=10)
                cols = [str(c).strip().upper() for c in df_scan.columns]
                if 'SEGMENTO' in cols and 'NOME' in cols and 'MATRICULA' in cols:
                    df = pd.read_excel(arquivo)
                    df.columns = [str(c).strip().upper() for c in df.columns]
                    df['CONTROLADOR'] = df['NOME'].astype(str).str.strip().str.upper()
                    df['SEGMENTO'] = df['SEGMENTO'].astype(str).str.strip().str.upper().str.split(' - ').str[0]
                    mapa = dict(zip(df['SEGMENTO'], df['CONTROLADOR']))
                    return mapa, f"Amarrado: {len(mapa)} controladores."
            except Exception:
                pass
        return {}, "Base de Controladores não localizada."
    except Exception as e:
        return {}, f"Erro Controladores: {e}"

# --- DRIVER SELENIUM ---
def init_driver():
    if state.driver: return state.driver
    d_dir = os.path.abspath("./temp_salesforce")
    os.makedirs(d_dir, exist_ok=True)
    
    opts = Options()
    opts.add_experimental_option("prefs", {
        "download.default_directory": d_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    })
    opts.add_argument("--log-level=3")
    opts.add_argument("--silent")
    opts.add_argument("--disable-logging")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    opts.add_argument("--start-maximized")

    try:
        service = Service("./chromedriver.exe", log_output=os.devnull)
        driver = webdriver.Chrome(service=service, options=opts)
        state.driver = driver
        state.status_msg = "🟢 Driver Ativo"
        return driver
    except Exception as e:
        state.status_msg = f"❌ Driver: {e}"
        return None

def baixar_arquivo_bg(report_id):
    if not state.driver: return None
    d_dir = os.path.abspath("./temp_salesforce")
    for f in glob.glob(os.path.join(d_dir, "*")):
        try: os.remove(f)
        except: pass

    antigos = set(glob.glob(os.path.join(d_dir, "*.xls*")))
    try:
        state.driver.get(f"https://oimoveltrialorg2021.lightning.force.com/lightning/r/Report/{report_id}/view")
        time.sleep(4)
        state.driver.get(f"https://oimoveltrialorg2021.lightning.force.com/servlet/PrintableViewDownloadServlet?isStandardRequest=true&reportId={report_id}")
        
        for _ in range(60):
            time.sleep(1)
            atuais = set(glob.glob(os.path.join(d_dir, "*.xls*")))
            novos = [f for f in (atuais - antigos) if not f.endswith('.crdownload') and not f.endswith('.tmp')]
            if novos:
                alvo = novos[0]
                t1 = os.path.getsize(alvo)
                time.sleep(1.5)
                t2 = os.path.getsize(alvo)
                if t1 > 0 and t1 == t2:
                    return alvo
        return None
    except Exception as e:
        logging.error(f"Erro ao baixar relatório {report_id}: {e}")
        return None

# --- PROCESSAMENTO PANDAS ---
def processar_hierarquia(caminho):
    try:
        try: df_raw = pd.read_excel(caminho, header=None)
        except:
            dfs = pd.read_html(caminho)
            if not dfs: return False
            df_raw = dfs[0]

        h_idx = -1
        for i, row in df_raw.head(30).iterrows():
            txt = " ".join([str(v).lower() for v in row.values])
            if "supervisor" in txt and ("gerente" in txt or "coordenador" in txt):
                h_idx = i
                break
        if h_idx == -1: return False

        df = df_raw.iloc[h_idx+1:].copy()
        df.columns = df_raw.iloc[h_idx]
        df = df.ffill()

        ren = {}
        for c in df.columns:
            cs = str(c).strip()
            if any(k in cs for k in ['Cont.', 'Contato', 'Telef', 'Cel']): continue
            if 'Recurso' in cs or 'Nome' in cs: ren[c] = 'Nome'
            elif 'Supervisor' in cs: ren[c] = 'Supervisor'
            elif 'Gerente' in cs or 'Coordenador' in cs: ren[c] = 'Gerente'

        df = df.rename(columns=ren)
        df = df.loc[:, ~df.columns.duplicated()]
        if 'Nome' not in df.columns or 'Supervisor' not in df.columns: return False

        df = df.drop_duplicates(subset=['Nome'])
        df = df[~df['Nome'].astype(str).str.contains('Total|Subtotal', case=False, na=False)]
        if 'Gerente' not in df.columns: df['Gerente'] = "NÃO INFORMADO"

        df['Gerente'] = df['Gerente'].astype(str).str.strip().str.upper()
        df['Supervisor'] = df['Supervisor'].astype(str).str.strip().str.upper()
        df['Nome'] = df['Nome'].astype(str).str.replace('\xa0', ' ').str.strip().str.upper()

        substituicoes = {'JOSE ROBERTO GARCIA': 'ROGERIO ALVES DE MAGALHAES'}
        df['Gerente'] = df['Gerente'].replace(substituicoes)
        df['Supervisor'] = df['Supervisor'].replace(substituicoes)
        df['Nome'] = df['Nome'].replace(substituicoes)

        # BLINDAGEM DE DUPLICIDADE: Atribui supervisor ao gerente onde tem mais M.O
        ranking_ger = df.groupby(['Supervisor', 'Gerente'])['Nome'].nunique().reset_index(name='QTD')
        ranking_ger = ranking_ger.sort_values(by=['Supervisor', 'QTD'], ascending=[True, False])
        
        mapa_ger_dominante = dict(zip(
            ranking_ger.drop_duplicates(subset=['Supervisor'])['Supervisor'], 
            ranking_ger.drop_duplicates(subset=['Supervisor'])['Gerente']
        ))
        df['Gerente'] = df['Supervisor'].map(mapa_ger_dominante).fillna(df['Gerente'])

        col_uf = next((c for c in df.columns if any(k in str(c).strip().upper() for k in ['UF', 'ESTADO', 'ESTADO/PROVÍNCIA', 'ESTADO_UF'])), None)
        if col_uf:
            def mapear_reg(v):
                v = str(v).strip().upper()
                return 'SUL' if v == 'PR' else ('NORTE' if v in ['AM', 'AP', 'PA', 'RR'] else 'OUTROS')
            df['REGIAO_HIERARQUIA'] = df[col_uf].apply(mapear_reg)
        else:
            df['REGIAO_HIERARQUIA'] = 'OUTROS'

        df = df[~df['Nome'].str.contains('CANCELAMENTO|REAGENDAMENTO|BUFFER', case=False, na=False)]
        state.df_hierarquia = df[['Nome', 'Supervisor', 'Gerente', 'REGIAO_HIERARQUIA']]
        return True
    except Exception as e:
        logging.error(f"Erro ao processar hierarquia: {e}")
        return False

def processar_producao(caminho):
    try:
        try: df_raw = pd.read_excel(caminho)
        except:
            dfs = pd.read_html(caminho)
            if not dfs: return False
            df_raw = dfs[0]

        h_idx = -1
        for i, r in df_raw.iterrows():
            if 'Número SA' in str(r.values):
                h_idx = i
                break
        if h_idx == -1: return False

        df = df_raw.iloc[h_idx:].copy()
        df.columns = df.iloc[0]
        df = df.iloc[1:].reset_index(drop=True)

        col_status = next((c for c in df.columns if ('Estado' in str(c) or 'Status' in str(c)) and 'UF' not in str(c).upper() and 'PROV' not in str(c).upper()), 'Estado')
        col_cia = 'ID Companhia' if 'ID Companhia' in df.columns else 'Companhia'
        col_macro = 'Macro Atividade'
        col_nome = next((c for c in df.columns if 'Recurso' in str(c) and 'Nome' in str(c)), None)
        if not col_nome: return False

        df = df[df[col_macro].astype(str).str.contains('FTTH', case=False, na=False)].copy()
        
        if col_cia in df.columns:
            df['Tenant'] = df[col_cia].apply(lambda x: 'NIO' if str(x).upper() in ['OI', 'NIO'] else ('TIM' if str(x).upper() == 'TIM' else 'DEMAIS'))
        else:
            df['Tenant'] = 'DEMAIS'

        df['Nome_Clean'] = df[col_nome].astype(str).str.replace('\xa0', ' ').str.strip().str.upper()
        df['Nome_Clean'] = df['Nome_Clean'].replace({'JOSE ROBERTO GARCIA': 'ROGERIO ALVES DE MAGALHAES'})
        df = df[~df['Nome_Clean'].str.contains('BUFFER|CANCELAMENTO|REAGENDAMENTO|HIGIENIZACAO', case=False, na=False)]

        if state.df_hierarquia is not None:
            g_map = dict(zip(state.df_hierarquia['Nome'], state.df_hierarquia['Gerente']))
            s_map = dict(zip(state.df_hierarquia['Nome'], state.df_hierarquia['Supervisor']))
            r_map = dict(zip(state.df_hierarquia['Nome'], state.df_hierarquia['REGIAO_HIERARQUIA']))
            df['Temp_Gerente'] = df['Nome_Clean'].map(g_map).fillna('')
            df['Temp_Sup'] = df['Nome_Clean'].map(s_map).fillna('')
            df['REGIAO_HIER'] = df['Nome_Clean'].map(r_map)
        else:
            df['Temp_Gerente'] = ''; df['Temp_Sup'] = ''; df['REGIAO_HIER'] = None

        col_setor = next((c for c in df.columns if 'Território' in str(c)), None)
        if col_setor:
            df['TEMP_SETOR'] = df[col_setor].astype(str).str.strip().str.upper()
            df['REGIAO_SETOR'] = df['TEMP_SETOR'].str[:2].map({'PR': 'SUL', 'AM': 'NORTE', 'AP': 'NORTE', 'PA': 'NORTE', 'RR': 'NORTE'})
        else:
            df['REGIAO_SETOR'] = None

        df['REGIAO'] = df['REGIAO_HIER'].combine_first(df['REGIAO_SETOR']).fillna('OUTROS')
        df = df[df['REGIAO'].isin(['SUL', 'NORTE'])]

        col_ini_plan = next((c for c in df.columns if str(c).strip().upper() in ['INÍCIO AGENDAMENTO', 'INICIO AGENDAMENTO'] or ('INÍCIO' in str(c).upper() and 'AGENDAD' in str(c).upper())), None)
        col_fim_plan = next((c for c in df.columns if str(c).strip().upper() in ['FIM AGENDAMENTO'] or ('FIM' in str(c).upper() and 'AGENDAD' in str(c).upper())), None)
        col_ini_real = next((c for c in df.columns if str(c).strip().upper() in ['INÍCIO EXECUÇÃO', 'INICIO EXECUÇÃO'] or ('INÍCIO' in str(c).upper() and 'EXEC' in str(c).upper()) or 'INÍCIO REAL' in str(c).upper()), None)
        col_motivo = next((c for c in df.columns if str(c).strip().upper() in ['DESCRIÇÃO', 'DESCRICAO', 'MOTIVO', 'MOTIVO DA PENDÊNCIA', 'SUBSTATUS', 'RAZÃO', 'OBSERVAÇÃO'] or 'PENDÊNCIA' in str(c).upper()), None)
        col_endereco = next((c for c in df.columns if any(k in str(c).upper() for k in ['ENDEREÇO', 'ENDERECO', 'RUA', 'LOGRADOURO'])), None)

        df['COL_MOTIVO_NOME'] = col_motivo
        df['COL_STATUS_NOME'] = col_status
        df['COL_MACRO_NOME'] = col_macro
        df['COL_NOME_SERV'] = col_nome
        df['COL_INI_REAL'] = col_ini_real
        df['COL_ENDERECO_NOME'] = col_endereco
        df['COL_SETOR_NOME'] = col_setor

        agora = pd.Timestamp(datetime.now())
        hoje_date = agora.date()

        def calcular_sla(row):
            if not col_ini_plan or not col_fim_plan: return 'Sem SLA'
            try:
                ini_plan = pd.to_datetime(row[col_ini_plan], dayfirst=True, errors='coerce')
                fim_plan = pd.to_datetime(row[col_fim_plan], dayfirst=True, errors='coerce')
                if pd.isna(ini_plan) or pd.isna(fim_plan): return 'Sem Data'
                if ini_plan.date() != hoje_date: return 'Outras Datas'
                
                status = str(row[col_status]).strip()
                status_ativos = ["Concluído com sucesso", "Concluído sem sucesso", "Em execução", "Fechado em WFM"]
                has_real = col_ini_real and pd.notna(row[col_ini_real])
                ini_real = pd.to_datetime(row[col_ini_real], dayfirst=True, errors='coerce') if has_real else None

                if status in status_ativos:
                    if ini_real is not None:
                        return 'No Prazo' if ini_plan <= ini_real <= fim_plan else 'Fora do Prazo'
                    return 'Fora do Prazo' if agora > fim_plan else 'Aguardando Início'
                else:
                    if agora < ini_plan: return 'Aguardando Início'
                    return 'No Prazo' if agora <= fim_plan else 'Fora do Prazo'
            except Exception:
                return 'Erro'

        df['SLA_STATUS'] = df.apply(calcular_sla, axis=1)
        state.df_producao_raw = df.copy()
        return True
    except Exception as e:
        logging.error(f"Erro ao processar produção: {e}")
        return False

def processar_reparos_icom(caminho):
    try:
        try: df_raw = pd.read_excel(caminho)
        except:
            dfs = pd.read_html(caminho)
            if not dfs: return False
            df_raw = dfs[0]

        h_idx = -1
        for i, r in df_raw.head(20).iterrows():
            linha_str = str(r.values).upper()
            if 'TERRITÓRIO DE SERVIÇO: NOME' in linha_str or 'NÚMERO SA' in linha_str:
                h_idx = i
                break
        
        df = df_raw.iloc[h_idx:].copy() if h_idx != -1 else df_raw
        if h_idx != -1:
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)

        df.columns = [str(c).strip() for c in df.columns]

        col_pronto = next((c for c in df.columns if 'PRONTO EXECU' in str(c).upper()), None)
        if col_pronto:
            df = df[df[col_pronto].astype(str).str.strip().str.upper().isin(['VERDADEIRO', 'TRUE', '1', '1.0'])].copy()

        col_setor = next((c for c in df.columns if 'Território' in str(c)), None)
        if not col_setor: return False
        df['SETOR'] = df[col_setor].fillna('SEM SETOR').astype(str).str.strip().str.upper()

        col_nome_serv = next((c for c in df.columns if 'Recurso' in str(c) and 'Nome' in str(c)), None)
        if col_nome_serv and state.df_hierarquia is not None:
            df['Nome_Clean'] = df[col_nome_serv].astype(str).str.replace('\xa0', ' ').str.strip().str.upper()
            df['Nome_Clean'] = df['Nome_Clean'].replace({'JOSE ROBERTO GARCIA': 'ROGERIO ALVES DE MAGALHAES'})
            df = df[~df['Nome_Clean'].str.contains('BUFFER|CANCELAMENTO|REAGENDAMENTO|HIGIENIZACAO', case=False, na=False)]
            
            g_map = dict(zip(state.df_hierarquia['Nome'], state.df_hierarquia['Gerente']))
            s_map = dict(zip(state.df_hierarquia['Nome'], state.df_hierarquia['Supervisor']))
            r_map = dict(zip(state.df_hierarquia['Nome'], state.df_hierarquia['REGIAO_HIERARQUIA']))
            df['Temp_Gerente'] = df['Nome_Clean'].map(g_map).fillna('')
            df['Temp_Sup'] = df['Nome_Clean'].map(s_map).fillna('')
            df['REGIAO_HIER'] = df['Nome_Clean'].map(r_map)
        else:
            df['Temp_Gerente'] = ''; df['Temp_Sup'] = ''; df['REGIAO_HIER'] = None

        df['UF_SIGLA'] = df['SETOR'].str[:2]
        df['REGIAO_SETOR'] = df['UF_SIGLA'].map({'PR': 'SUL', 'AM': 'NORTE', 'AP': 'NORTE', 'PA': 'NORTE', 'RR': 'NORTE'})
        df['REGIAO'] = df['REGIAO_SETOR'].combine_first(df['REGIAO_HIER']).fillna('OUTROS')
        df = df[df['REGIAO'].isin(['SUL', 'NORTE'])]

        col_cia = 'ID Companhia' if 'ID Companhia' in df.columns else ('Companhia' if 'Companhia' in df.columns else None)
        if col_cia:
            df['Tenant'] = df[col_cia].apply(lambda x: 'NIO' if str(x).upper() in ['OI', 'NIO'] else ('TIM' if str(x).upper() == 'TIM' else 'DEMAIS'))
        else:
            df['Tenant'] = 'DEMAIS'

        hoje = datetime.now().date()
        amanha = hoje + timedelta(days=1)
        col_inicio = 'Início Execução' if 'Início Execução' in df.columns else None
        col_abertura = next((c for c in df.columns if 'Abertura' in str(c)), None)

        def classificar_quando(row):
            val_inicio = row.get(col_inicio) if col_inicio else None
            if pd.isna(val_inicio): return 'Futuro'
            try:
                val_date = pd.to_datetime(val_inicio, dayfirst=True).date()
                if val_date <= hoje: return 'Hoje'
                return 'Amanhã' if val_date == amanha else 'Futuro'
            except Exception:
                return 'Futuro'

        df['QUANDO_CAT'] = df.apply(classificar_quando, axis=1) if col_inicio or col_abertura else 'Futuro'

        agora = datetime.now()
        df['FAIXA_AGING'] = 'N/A'
        
        if col_abertura:
            df['DATA_ABERTURA'] = pd.to_datetime(df[col_abertura], dayfirst=True, errors='coerce')
            df['HORA_ABERTURA'] = df['DATA_ABERTURA'].dt.strftime('%H:00').fillna('N/A')
            
            col_macro = 'Macro Atividade' if 'Macro Atividade' in df.columns else None
            if col_macro:
                mask_rep = df[col_macro].astype(str).str.contains('REP', case=False, na=False)
                mask_abertura = df['DATA_ABERTURA'].notna()
                mask_valida = mask_rep & mask_abertura

                if mask_valida.any():
                    horas = (agora - df.loc[mask_valida, 'DATA_ABERTURA']).dt.total_seconds() / 3600
                    
                    df.loc[mask_valida & (horas <= 24), 'FAIXA_AGING'] = 'Até 24h'
                    df.loc[mask_valida & (horas > 24) & (horas <= 48), 'FAIXA_AGING'] = '24 a 48h'
                    df.loc[mask_valida & (horas > 48) & (horas <= 72), 'FAIXA_AGING'] = '48 a 72h'
                    df.loc[mask_valida & (horas > 72), 'FAIXA_AGING'] = '> 72h'
        else:
            df['DATA_ABERTURA'] = pd.NaT
            df['HORA_ABERTURA'] = 'N/A'

        state.df_reparos_raw = df.copy()
        return True
    except Exception as e:
        logging.error(f"Erro ao processar reparos: {e}")
        return False

# --- COMPILADORES ANALÍTICOS ---
def compilar_cockpit_reparos(df_reg):
    hoje_qty = len(df_reg[df_reg['QUANDO_CAT'] == 'Hoje'])
    amanha_qty = len(df_reg[df_reg['QUANDO_CAT'] == 'Amanhã'])
    futuro_qty = len(df_reg[df_reg['QUANDO_CAT'] == 'Futuro'])
    tot_d0 = hoje_qty + amanha_qty + futuro_qty
    pct_d0 = (hoje_qty / tot_d0 * 100) if tot_d0 > 0 else 0
    tot_d1 = amanha_qty + futuro_qty
    pct_d1 = (amanha_qty / tot_d1 * 100) if tot_d1 > 0 else 0
    pct_futuro = (futuro_qty / tot_d0 * 100) if tot_d0 > 0 else 0
    return {
        'hoje': hoje_qty, 'amanha': amanha_qty, 'futuro': futuro_qty,
        'tot_d0': tot_d0, 'pct_d0': round(pct_d0, 1),
        'tot_d1': tot_d1, 'pct_d1': round(pct_d1, 1),
        'pct_futuro': round(pct_futuro, 1)
    }

def compilar_dados_producao(df_regiao, df_hierarquia):
    if df_regiao.empty:
        return {'tabela': [], 'top10': [], 'kpis': {'NIO': 'INST: 0 | REP: 0 | %AG: 0.0%', 'TIM': 'INST: 0 | REP: 0 | %AG: 0.0%', 'DEMAIS': 'INST: 0 | REP: 0 | %AG: 0.0%', 'TOTAL': 'INST: 0 | REP: 0 | GERAL: 0 | %AG: 0.0%'}, 'meta': 0, 'realizado': 0}

    col_status = df_regiao['COL_STATUS_NOME'].iloc[0] if 'COL_STATUS_NOME' in df_regiao.columns else 'Estado'
    col_macro = df_regiao['COL_MACRO_NOME'].iloc[0] if 'COL_MACRO_NOME' in df_regiao.columns else 'Macro Atividade'
    col_nome_serv = df_regiao['COL_NOME_SERV'].iloc[0] if 'COL_NOME_SERV' in df_regiao.columns else 'Recurso de serviço: Nome'

    df_suc = df_regiao[df_regiao[col_status].astype(str).str.contains('Concluído com sucesso', case=False, na=False)].copy()
    
    def calc_mac(d_suc, d_all, is_total=False):
        i = d_suc[col_macro].astype(str).str.contains('INST|MUD', case=False, na=False, regex=True).sum() if col_macro in d_suc.columns else 0
        r = d_suc[col_macro].astype(str).str.contains('REP', case=False, na=False).sum() if col_macro in d_suc.columns else 0
        npz = (d_all['SLA_STATUS'] == 'No Prazo').sum()
        fpz = (d_all['SLA_STATUS'] == 'Fora do Prazo').sum()
        tot_sla = npz + fpz
        sla = (npz / tot_sla * 100) if tot_sla > 0 else 0
        if is_total: return f"INST: {i} | REP: {r} | GERAL: {i+r} | %AG: {sla:.1f}%"
        return f"INST: {i} | REP: {r} | %AG: {sla:.1f}%"

    kpis = {
        'NIO': calc_mac(df_suc[df_suc['Tenant']=='NIO'], df_regiao[df_regiao['Tenant']=='NIO']),
        'TIM': calc_mac(df_suc[df_suc['Tenant']=='TIM'], df_regiao[df_regiao['Tenant']=='TIM']),
        'DEMAIS': calc_mac(df_suc[df_suc['Tenant']=='DEMAIS'], df_regiao[df_regiao['Tenant']=='DEMAIS']),
        'TOTAL': calc_mac(df_suc, df_regiao, is_total=True)
    }

    # Usa a hierarquia já resolvida no próprio conjunto filtrado. Refazer o
    # vínculo pelo nome do técnico poderia trazer o coordenador de nível master.
    hierarquia_resolvida = all(
        coluna in df_regiao.columns
        for coluna in ['Nome_Clean', 'Temp_Gerente', 'Temp_Sup']
    )
    colunas_grupo = ['Nome_Clean', 'Temp_Gerente', 'Temp_Sup'] if hierarquia_resolvida else [col_nome_serv]

    tec = df_regiao.groupby(colunas_grupo, dropna=False).agg(
        SUC=(col_status, lambda x: x.astype(str).str.contains('Concluído com sucesso', case=False, na=False).sum()),
        SEM_SUC=(col_status, lambda x: x.astype(str).str.contains('Concluído sem sucesso', case=False, na=False).sum()),
        EXEC=(col_status, lambda x: x.astype(str).str.contains('Em execução', case=False, na=False).sum()),
        DESL=(col_status, lambda x: x.astype(str).str.contains('Em deslocamento', case=False, na=False).sum()),
        CAIXA=(col_status, lambda x: x.isin(['Em deslocamento','Atribuído','Recebido','Em execução','Fechado em WFM']).sum()),
        NO_PRAZO=('SLA_STATUS', lambda x: (x == 'No Prazo').sum()),
        FORA_PRAZO=('SLA_STATUS', lambda x: (x == 'Fora do Prazo').sum())
    ).reset_index()

    if hierarquia_resolvida:
        merged = tec.rename(columns={
            'Nome_Clean': 'Nome',
            'Temp_Gerente': 'Gerente',
            'Temp_Sup': 'Supervisor'
        })
        merged['Nome'] = merged['Nome'].astype(str).str.strip().str.upper()
        merged['Gerente'] = merged['Gerente'].astype(str).str.strip().str.upper()
        merged['Supervisor'] = merged['Supervisor'].astype(str).str.strip().str.upper()
        merged = merged[(merged['Gerente'] != '') & (merged['Supervisor'] != '')]
    else:
        tec = tec.rename(columns={col_nome_serv: 'Nome'})
        tec['Nome'] = tec['Nome'].astype(str).str.strip().str.upper()
        merged = pd.merge(tec, df_hierarquia, on='Nome', how='inner')
    merged['EFICACIA_VAL'] = (merged['SUC']/(merged['SUC']+merged['SEM_SUC'])).fillna(0)
    top10_df = merged[merged['SUC']>0].sort_values(['SUC','EFICACIA_VAL'], ascending=False).head(10)
    top10 = [{'Nome': r['Nome'], 'SUC': int(r['SUC'])} for _, r in top10_df.iterrows()]

    if merged.empty:
        return {'tabela': [], 'top10': [], 'kpis': kpis, 'meta': 0, 'realizado': 0}

    res = merged.groupby(['Gerente','Supervisor']).agg(
        MO_CAMPO=('Nome','nunique'), ZERADOS=('SUC', lambda x: (x==0).sum()), SUC=('SUC','sum'), SEM_SUC=('SEM_SUC','sum'),
        EXEC=('EXEC', 'sum'), DESL=('DESL', 'sum'), CAIXA=('CAIXA','sum'), NO_PRAZO=('NO_PRAZO','sum'), FORA_PRAZO=('FORA_PRAZO','sum')
    ).reset_index()

    res['Gerente'] = res['Gerente'].astype(str).str.strip().str.upper()
    res['Supervisor'] = res['Supervisor'].astype(str).str.strip().str.upper()

    # GARANTIA DE TIPAGEM NUMÉRICA: Evita descarte nas somas
    cols_numericas = ['MO_CAMPO', 'ZERADOS', 'SUC', 'SEM_SUC', 'EXEC', 'DESL', 'CAIXA', 'NO_PRAZO', 'FORA_PRAZO']
    for c in cols_numericas:
        res[c] = pd.to_numeric(res[c], errors='coerce').fillna(0)
    
    def projecao(r):
        tot = r['SUC'] + r['SEM_SUC']
        ef = (r['SUC'] / tot) if tot > 0 else 0.7
        return int(r['SUC'] + (r['CAIXA'] * ef))

    res['PROJECAO'] = res.apply(projecao, axis=1)
    res['ATV_TEC'] = (res['SUC'] / res['MO_CAMPO']).round(2)
    res['EFICACIA'] = (res['SUC'] / (res['SUC'] + res['SEM_SUC']) * 100).fillna(0).round(1)

    rows = []
    grps = []
    for g, gd in res.groupby('Gerente'):
        s_mo = gd['MO_CAMPO'].sum()
        s_suc = gd['SUC'].sum()
        atv = s_suc / s_mo if s_mo > 0 else 0
        grps.append((g, atv, gd))
    
    grps.sort(key=lambda x: x[1], reverse=True)
    for g, atv, gd in grps:
        s_mo = int(gd['MO_CAMPO'].sum())
        s_zer = int(gd['ZERADOS'].sum())
        s_suc = int(gd['SUC'].sum())
        s_sem_suc = int(gd['SEM_SUC'].sum())
        s_exec = int(gd['EXEC'].sum())
        s_desl = int(gd['DESL'].sum())
        s_caixa = int(gd['CAIXA'].sum())
        s_proj = int(gd['PROJECAO'].sum())
        s_np = gd['NO_PRAZO'].sum()
        s_fp = gd['FORA_PRAZO'].sum()
        
        ef = (s_suc / (s_suc + s_sem_suc) * 100) if (s_suc + s_sem_suc) > 0 else 0
        sla_val = (s_np / (s_np + s_fp) * 100) if (s_np + s_fp) > 0 else 0
        
        rows.append({
            'COORDENADOR': g, 'SUPERVISOR': '', 'CONTROLADOR': state.mapa_controladores.get(g, ''),
            'MO_ATIVA': s_mo, 'ZERADOS': s_zer, 'COM_SUCESSO': s_suc,
            'SEM_SUCESSO': s_sem_suc, 'EM_EXEC': s_exec, 'EM_DESL': s_desl,
            'CAIXA_GERAL': s_caixa, 'PROJECAO': s_proj,
            'ATIVIDADE_POR_TECNICO': round(atv, 2), 'SLA_PCT': round(sla_val, 1), 'EFICACIA': round(ef, 1),
            'TIPO': 'COORD'
        })
        for _, r in gd.sort_values('ATV_TEC', ascending=False).iterrows():
            sla_sup = (r['NO_PRAZO'] / (r['NO_PRAZO'] + r['FORA_PRAZO']) * 100) if (r['NO_PRAZO'] + r['FORA_PRAZO']) > 0 else 0
            rows.append({
                'COORDENADOR': '', 'SUPERVISOR': r['Supervisor'], 'CONTROLADOR': state.mapa_controladores.get(r['Supervisor'], ''),
                'MO_ATIVA': int(r['MO_CAMPO']), 'ZERADOS': int(r['ZERADOS']), 'COM_SUCESSO': int(r['SUC']),
                'SEM_SUCESSO': int(r['SEM_SUC']), 'EM_EXEC': int(r['EXEC']), 'EM_DESL': int(r['DESL']),
                'CAIXA_GERAL': int(r['CAIXA']), 'PROJECAO': int(r['PROJECAO']),
                'ATIVIDADE_POR_TECNICO': round(r['ATV_TEC'], 2), 'SLA_PCT': round(sla_sup, 1), 'EFICACIA': round(r['EFICACIA'], 1),
                'TIPO': 'SUP'
            })

    # CÁLCULO DIRETO E SEGURO DO TOTAL GERAL
    tot_suc = float(res['SUC'].sum())
    tot_sem_suc = float(res['SEM_SUC'].sum())
    tot_mo = float(res['MO_CAMPO'].sum())
    tot_np = float(res['NO_PRAZO'].sum())
    tot_fp = float(res['FORA_PRAZO'].sum())

    atv_tot = (tot_suc / tot_mo) if tot_mo > 0 else 0
    ef_tot = (tot_suc / (tot_suc + tot_sem_suc) * 100) if (tot_suc + tot_sem_suc) > 0 else 0
    sla_tot = (tot_np / (tot_np + tot_fp) * 100) if (tot_np + tot_fp) > 0 else 0

    rows.append({
        'COORDENADOR': 'TOTAL GERAL', 'SUPERVISOR': '', 'CONTROLADOR': '',
        'MO_ATIVA': int(tot_mo), 'ZERADOS': int(res['ZERADOS'].sum()), 'COM_SUCESSO': int(tot_suc),
        'SEM_SUCESSO': int(tot_sem_suc), 'EM_EXEC': int(res['EXEC'].sum()), 'EM_DESL': int(res['DESL'].sum()),
        'CAIXA_GERAL': int(res['CAIXA'].sum()), 'PROJECAO': int(res['PROJECAO'].sum()),
        'ATIVIDADE_POR_TECNICO': round(atv_tot, 2), 'SLA_PCT': round(sla_tot, 1), 'EFICACIA': round(ef_tot, 1),
        'TIPO': 'TOTAL'
    })

    return {
        'tabela': rows,
        'top10': top10,
        'kpis': kpis,
        'meta': int(tot_mo * 5),
        'realizado': int(tot_suc),
        'mo_total': int(tot_mo)
    }

def compilar_performance_reparos(df, regiao):
    """Replica a visão gerencial da planilha usando a base viva de produção de reparos."""
    vazio = {'tabela': [], 'motivos_sem_sucesso': [], 'motivos_com_sucesso': [], 'metas': {}}
    if df is None or df.empty:
        return vazio

    col_status = df['COL_STATUS_NOME'].iloc[0] if 'COL_STATUS_NOME' in df.columns else None
    col_macro = df['COL_MACRO_NOME'].iloc[0] if 'COL_MACRO_NOME' in df.columns else None
    col_motivo = df['COL_MOTIVO_NOME'].iloc[0] if 'COL_MOTIVO_NOME' in df.columns else None
    if not col_status or col_status not in df.columns:
        return vazio

    if col_macro and col_macro in df.columns:
        df = df[df[col_macro].astype(str).str.contains('REP', case=False, na=False)].copy()
    if df.empty:
        return vazio

    status_norm = df[col_status].astype(str).map(
        lambda valor: ''.join(c for c in unicodedata.normalize('NFKD', valor) if not unicodedata.combining(c)).strip().upper()
    )
    df['_SUCESSO'] = status_norm.str.contains('CONCLU', na=False) & status_norm.str.contains('COM SUCESSO', na=False)
    df['_SEM_SUCESSO'] = status_norm.str.contains('CONCLU', na=False) & status_norm.str.contains('SEM SUCESSO', na=False)
    df['_CAIXA'] = status_norm.str.startswith(('EM DESLOC', 'ATRIBU', 'RECEBIDO', 'EM EXECU', 'FECHADO EM WFM'), na=False)

    metas_todas = carregar_metas_reparos()
    metas_regiao = metas_todas.get(normalizar_chave_meta(regiao), {})
    linhas = []

    for gerente, grupo in df.groupby('Temp_Gerente', dropna=False):
        gerente = normalizar_chave_meta(gerente) or 'NÃO INFORMADO'
        sucesso = int(grupo['_SUCESSO'].sum())
        sem_sucesso = int(grupo['_SEM_SUCESSO'].sum())
        caixa = int(grupo['_CAIXA'].sum())
        finalizados = sucesso + sem_sucesso
        eficacia = (sucesso / finalizados * 100) if finalizados else 0

        if 'Nome_Clean' in grupo.columns:
            por_tecnico = grupo.groupby('Nome_Clean')['_SUCESSO'].sum()
            zerados = int((por_tecnico == 0).sum())
        else:
            zerados = 0

        projetado = int(sucesso + caixa * (eficacia / 100) + 0.5)
        meta = max(0, int(metas_regiao.get(gerente, 0) or 0))
        linhas.append({
            'COORDENADOR': gerente, 'SUCESSO': sucesso, 'SEM_SUCESSO': sem_sucesso,
            'CAIXA': caixa, 'ZERADOS': zerados, 'EFICACIA': round(eficacia, 1),
            'PROJETADO': projetado, 'DIF_META': projetado - meta, 'META': meta,
            'GAP_META': sucesso - meta,
            'PCT_META': round((sucesso / meta * 100) if meta else 0, 1), 'TIPO': 'COORD'
        })

    linhas.sort(key=lambda item: (item['PCT_META'], item['PROJETADO']), reverse=True)
    if linhas:
        total_sucesso = sum(item['SUCESSO'] for item in linhas)
        total_sem = sum(item['SEM_SUCESSO'] for item in linhas)
        total_meta = sum(item['META'] for item in linhas)
        total_final = total_sucesso + total_sem
        linhas.append({
            'COORDENADOR': 'TOTAL GERAL', 'SUCESSO': total_sucesso, 'SEM_SUCESSO': total_sem,
            'CAIXA': sum(item['CAIXA'] for item in linhas), 'ZERADOS': sum(item['ZERADOS'] for item in linhas),
            'EFICACIA': round((total_sucesso / total_final * 100) if total_final else 0, 1),
            'PROJETADO': sum(item['PROJETADO'] for item in linhas),
            'DIF_META': sum(item['DIF_META'] for item in linhas), 'META': total_meta,
            'GAP_META': total_sucesso - total_meta,
            'PCT_META': round((total_sucesso / total_meta * 100) if total_meta else 0, 1), 'TIPO': 'TOTAL'
        })

    def ranking_motivos(mask, limite=10):
        if not col_motivo or col_motivo not in df.columns:
            return []
        serie = df.loc[mask, col_motivo].fillna('NÃO INFORMADO').astype(str).str.strip()
        serie = serie[serie != '']
        contagem = serie.value_counts().head(limite)
        total = int(contagem.sum())
        return [
            {'MOTIVO': motivo, 'QUANTIDADE': int(qtd), 'PCT': round((qtd / total * 100) if total else 0, 1)}
            for motivo, qtd in contagem.items()
        ]

    return {
        'tabela': linhas,
        'motivos_sem_sucesso': ranking_motivos(df['_SEM_SUCESSO']),
        'motivos_com_sucesso': ranking_motivos(df['_SUCESSO']),
        'metas': metas_regiao
    }

# --- SÍNTESE DE ÁUDIO ---
async def _async_gerar_audio(txt, filename):
    c = edge_tts.Communicate(txt, "pt-BR-AntonioNeural", rate="+20%")
    await c.save(filename)

def processar_audio(txt):
    for f in glob.glob("temp_voice_*.mp3"):
        try: os.remove(f)
        except: pass
    fname = f"temp_voice_{uuid.uuid4().hex[:8]}.mp3"
    asyncio.run(_async_gerar_audio(txt, fname))
    return fname

def gerar_mensagem_voz(dados_prod, regiao):
    try:
        tabela = dados_prod.get('tabela', [])
        if not tabela: return "Atualização concluída."
        
        coords = [r for r in tabela if r['TIPO'] == 'COORD']
        sups = [r for r in tabela if r['TIPO'] == 'SUP']
        total = next((r for r in tabela if r['TIPO'] == 'TOTAL'), None)
        
        tot_suc = total['COM_SUCESSO'] if total else 0
        tot_zer = total['ZERADOS'] if total else 0
        
        melhor_c = max(coords, key=lambda x: x['ATIVIDADE_POR_TECNICO'])['COORDENADOR'].title() if coords else "N/A"
        pior_c = min(coords, key=lambda x: x['ATIVIDADE_POR_TECNICO'])['COORDENADOR'].title() if coords else "N/A"
        melhor_s = max(sups, key=lambda x: x['ATIVIDADE_POR_TECNICO'])['SUPERVISOR'].title() if sups else "N/A"
        pior_s = min(sups, key=lambda x: x['ATIVIDADE_POR_TECNICO'])['SUPERVISOR'].title() if sups else "N/A"
        
        return (f"Atualização concluída para a regional {regiao}. Total de {tot_suc} atividades com sucessos e {tot_zer} técnicos zerados. "
                f"Destaque positivo para a coordenação de {melhor_c} e supervisor {melhor_s}. "
                f"Atenção necessária para a coordenação de {pior_c} e supervisor {pior_s}.")
    except Exception:
        return "Atualização concluída."

# --- TAREFAS EM SEGUNDO PLANO ---
def loop_atualizacao_background():
    while True:
        try:
            cfg = state.config
            if cfg["auto_refresh"] and state.driver and state.df_hierarquia is not None and not state.bloqueio_processamento:
                deve_rodar = False
                if state.ultima_atualizacao is None:
                    deve_rodar = True
                elif (datetime.now() - state.ultima_atualizacao).total_seconds() > (cfg["intervalo_refresh_min"] * 60):
                    deve_rodar = True

                if deve_rodar:
                    state.bloqueio_processamento = True
                    state.status_msg = "🔄 Robô extraindo dados no Salesforce..."
                    f_prod = baixar_arquivo_bg("00ON400000MWfkwMAD")
                    if f_prod: processar_producao(f_prod)

                    f_rep = baixar_arquivo_bg("00ON400000Kmw9RMAR")
                    if f_rep: processar_reparos_icom(f_rep)

                    state.ultima_atualizacao = datetime.now()
                    state.status_msg = f"✅ Atualizado: {state.ultima_atualizacao.strftime('%H:%M:%S')}"
                    state.bloqueio_processamento = False

                    if cfg["voz_ativa"]:
                        for reg in ['NORTE', 'SUL']:
                            ult = state.ultima_fala.get(reg)
                            if not ult or (datetime.now() - ult).total_seconds() > (cfg["intervalo_voz_min"] * 60):
                                if state.df_producao_raw is not None:
                                    try:
                                        df_r = state.df_producao_raw[state.df_producao_raw['REGIAO'] == reg]
                                        d_analise = compilar_dados_producao(df_r, state.df_hierarquia)
                                        msg = gerar_mensagem_voz(d_analise, reg)
                                        state.audio_pendente = processar_audio(msg)
                                        state.ultima_fala[reg] = datetime.now()
                                    except Exception as e:
                                        state.audio_pendente = None
                                        state.ultima_fala[reg] = datetime.now()
                                        detalhe = "serviço de voz recusou a conexão (HTTP 403)" if "403" in str(e) else str(e)
                                        logging.warning(f"Falha ao gerar áudio para a regional {reg}: {detalhe}. Nova tentativa após o intervalo configurado.")
                                    break
        except Exception as e:
            state.status_msg = f"❌ Erro Update: {e}"
            state.bloqueio_processamento = False
        time.sleep(10)

threading.Thread(target=loop_atualizacao_background, daemon=True).start()

# --- ROTAS FLASK ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/sw.js')
def service_worker():
    response = send_from_directory(app.static_folder, 'sw.js')
    response.headers['Service-Worker-Allowed'] = '/'
    return response

@app.route('/audio/<filename>')
def get_audio(filename):
    return send_from_directory('.', filename)

@app.route('/api/iniciar_driver', methods=['POST'])
def rota_iniciar():
    d = init_driver()
    if d:
        d.get("https://oimoveltrialorg2021.lightning.force.com/")
        if not state.mapa_controladores:
            m, _ = carregar_controladores_local()
            state.mapa_controladores = m
        return jsonify({"sucesso": True, "msg": "Driver iniciado com sucesso."})
    return jsonify({"sucesso": False, "msg": "Falha ao inicializar chromedriver."})

@app.route('/api/carregar_hierarquia', methods=['POST'])
def rota_hierarquia():
    if not state.driver: return jsonify({"sucesso": False, "msg": "Driver inativo."})
    f = baixar_arquivo_bg("00ON400000KCybBMAT")
    if f and processar_hierarquia(f):
        return jsonify({"sucesso": True, "total": len(state.df_hierarquia)})
    return jsonify({"sucesso": False, "msg": "Falha ao processar arquivo da hierarquia."})

@app.route('/api/filtros', methods=['GET'])
def rota_filtros():
    g_set, s_set, set_set, m_set, t_set = set(), set(), set(), set(), set()
    
    if state.df_hierarquia is not None and not state.df_hierarquia.empty:
        g_set.update(state.df_hierarquia['Gerente'].dropna().unique())
        s_set.update(state.df_hierarquia['Supervisor'].dropna().unique())

    if state.df_producao_raw is not None and not state.df_producao_raw.empty:
        df_p = state.df_producao_raw
        if 'Temp_Gerente' in df_p.columns: g_set.update(df_p['Temp_Gerente'].dropna().unique())
        if 'Temp_Sup' in df_p.columns: s_set.update(df_p['Temp_Sup'].dropna().unique())
        if 'COL_SETOR_NOME' in df_p.columns and df_p['COL_SETOR_NOME'].iloc[0]:
            col_s = df_p['COL_SETOR_NOME'].iloc[0]
            if col_s in df_p.columns: set_set.update(df_p[col_s].dropna().unique())
        if 'Tenant' in df_p.columns: t_set.update(df_p['Tenant'].dropna().unique())

    if state.df_reparos_raw is not None and not state.df_reparos_raw.empty:
        df_r = state.df_reparos_raw
        if 'Temp_Gerente' in df_r.columns: g_set.update(df_r['Temp_Gerente'].dropna().unique())
        if 'Temp_Sup' in df_r.columns: s_set.update(df_r['Temp_Sup'].dropna().unique())
        if 'SETOR' in df_r.columns: set_set.update(df_r['SETOR'].dropna().unique())
        if 'Macro Atividade' in df_r.columns: m_set.update(df_r['Macro Atividade'].dropna().unique())
        if 'Tenant' in df_r.columns: t_set.update(df_r['Tenant'].dropna().unique())

    if not m_set: m_set = {'INST-FTTH', 'REP-FTTH', 'MUD-FTTH'}
    if not t_set: t_set = {'NIO', 'TIM', 'DEMAIS'}

    return jsonify({
        "gerentes": sorted(list(g_set - {'', 'NÃO INFORMADO'})),
        "supervisores": sorted(list(s_set - {''})),
        "setores": sorted([str(s).strip() for s in set_set if str(s).strip() != '']),
        "macros": sorted(list(m_set)),
        "tenants": sorted(list(t_set))
    })

@app.route('/api/metas_reparos', methods=['POST'])
def rota_metas_reparos():
    body = request.get_json(silent=True) or {}
    regiao = normalizar_chave_meta(body.get('regiao'))
    metas_recebidas = body.get('metas', {})
    if regiao not in {'NORTE', 'SUL'} or not isinstance(metas_recebidas, dict):
        return jsonify({'sucesso': False, 'msg': 'Regional ou metas inválidas.'}), 400

    metas_limpas = {}
    try:
        for coordenador, valor in metas_recebidas.items():
            chave = normalizar_chave_meta(coordenador)
            if not chave or chave == 'TOTAL GERAL':
                continue
            numero = int(valor)
            if numero < 0 or numero > 100000:
                raise ValueError('Meta fora do intervalo permitido')
            metas_limpas[chave] = numero
    except (TypeError, ValueError):
        return jsonify({'sucesso': False, 'msg': 'As metas devem ser números inteiros entre 0 e 100000.'}), 400

    todas = carregar_metas_reparos()
    todas[regiao] = metas_limpas
    try:
        salvar_metas_reparos(todas)
    except Exception as e:
        logging.error(f"Erro ao salvar metas de reparos: {e}")
        return jsonify({'sucesso': False, 'msg': 'Não foi possível persistir as metas.'}), 500
    return jsonify({'sucesso': True, 'msg': f'Metas de {regiao.title()} salvas.', 'metas': metas_limpas})

@app.route('/api/dados', methods=['POST'])
def rota_dados():
    body = request.get_json() or {}
    tela = body.get('tela', 'PRODUCAO')
    regiao = body.get('regiao', 'NORTE')
    filtros = body.get('filtros', {})

    f_gerente = filtros.get('gerente', [])
    f_sup = filtros.get('supervisor', [])
    f_setor = filtros.get('setor', [])
    f_macro = filtros.get('macro', [])
    f_tenant = filtros.get('tenant', [])

    resposta = {
        "status_msg": state.status_msg,
        "ultima_atualizacao": state.ultima_atualizacao.strftime('%H:%M:%S') if state.ultima_atualizacao else 'N/A',
        "bloqueio_processamento": state.bloqueio_processamento,
        "audio_pendente": state.audio_pendente,
        "payload": {}
    }
    state.audio_pendente = None

    if tela == 'PRODUCAO':
        if state.df_producao_raw is not None:
            df = state.df_producao_raw[state.df_producao_raw['REGIAO'] == regiao].copy()
            if f_gerente: df = df[df['Temp_Gerente'].isin(f_gerente)]
            if f_sup: df = df[df['Temp_Sup'].isin(f_sup)]
            if f_setor and not df.empty:
                col_s = df['COL_SETOR_NOME'].iloc[0]
                if col_s: df = df[df[col_s].isin(f_setor)]
            if f_tenant and 'Tenant' in df.columns: df = df[df['Tenant'].isin(f_tenant)]
            payload_producao = compilar_dados_producao(df, state.df_hierarquia)
            if f_sup:
                payload_producao['tabela'] = [
                    linha for linha in payload_producao.get('tabela', [])
                    if linha.get('TIPO') == 'SUP'
                ]
            resposta['payload'] = payload_producao

    elif tela == 'PERFORMANCE_REPAROS':
        if state.df_producao_raw is not None:
            df = state.df_producao_raw[state.df_producao_raw['REGIAO'] == regiao].copy()
            if f_gerente: df = df[df['Temp_Gerente'].isin(f_gerente)]
            if f_sup: df = df[df['Temp_Sup'].isin(f_sup)]
            if f_setor and not df.empty:
                col_s = df['COL_SETOR_NOME'].iloc[0]
                if col_s: df = df[df[col_s].isin(f_setor)]
            if f_tenant and 'Tenant' in df.columns: df = df[df['Tenant'].isin(f_tenant)]
            resposta['payload'] = compilar_performance_reparos(df, regiao)

    elif tela == 'PENDENCIAS':
        if state.df_producao_raw is not None:
            df = state.df_producao_raw.copy()
            col_status = df['COL_STATUS_NOME'].iloc[0] if 'COL_STATUS_NOME' in df.columns else None
            col_motivo = df['COL_MOTIVO_NOME'].iloc[0] if 'COL_MOTIVO_NOME' in df.columns else None
            col_setor = df['COL_SETOR_NOME'].iloc[0] if 'COL_SETOR_NOME' in df.columns else None

            if col_status and col_motivo:
                df = df[(df['REGIAO'] == regiao) & (df[col_status].astype(str).str.contains('Concluído sem sucesso', case=False, na=False))]
                if f_gerente: df = df[df['Temp_Gerente'].isin(f_gerente)]
                if f_sup: df = df[df['Temp_Sup'].isin(f_sup)]
                if f_setor and col_setor: df = df[df[col_setor].isin(f_setor)]
                if f_tenant and 'Tenant' in df.columns: df = df[df['Tenant'].isin(f_tenant)]
                
                if not df.empty:
                    cont = df[col_motivo].value_counts().reset_index()
                    cont.columns = ['MOTIVO', 'QUANTIDADE']
                    resposta['payload'] = {
                        'lista': cont.to_dict(orient='records'),
                        'max': int(cont['QUANTIDADE'].max())
                    }

    elif tela == 'EXECUCAO':
        if state.df_producao_raw is not None:
            df = state.df_producao_raw.copy()
            col_ini = df['COL_INI_REAL'].iloc[0] if 'COL_INI_REAL' in df.columns else None
            col_status = df['COL_STATUS_NOME'].iloc[0] if 'COL_STATUS_NOME' in df.columns else None
            col_end = df['COL_ENDERECO_NOME'].iloc[0] if 'COL_ENDERECO_NOME' in df.columns else None
            col_setor = df['COL_SETOR_NOME'].iloc[0] if 'COL_SETOR_NOME' in df.columns else None

            if col_ini and col_status:
                df = df[(df['REGIAO'] == regiao) & (df[col_status].astype(str).str.strip().str.upper() == 'EM EXECUÇÃO')].dropna(subset=[col_ini]).copy()
                df['ENDERECO'] = df[col_end].fillna('Não Informado') if col_end and col_end in df.columns else 'Não Informado'
                if f_gerente: df = df[df['Temp_Gerente'].isin(f_gerente)]
                if f_sup: df = df[df['Temp_Sup'].isin(f_sup)]
                if f_setor and col_setor: df = df[df[col_setor].isin(f_setor)]
                if f_tenant and 'Tenant' in df.columns: df = df[df['Tenant'].isin(f_tenant)]

                if not df.empty:
                    agora_dt = datetime.now()
                    df['MINUTOS'] = (agora_dt - pd.to_datetime(df[col_ini], dayfirst=True, errors='coerce')).dt.total_seconds() / 60
                    df = df[df['MINUTOS'] > 80].sort_values(by='MINUTOS', ascending=False)
                    
                    itens = []
                    for _, r in df.iterrows():
                        m = r['MINUTOS']
                        itens.append({
                            'SUPERVISOR': r['Temp_Sup'],
                            'COLABORADOR': r['Nome_Clean'],
                            'ENDERECO': r['ENDERECO'],
                            'TEMPO': f"{int(m // 60):02d}:{int(m % 60):02d}:00",
                            'MAPA': f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(str(r['ENDERECO']))}" if r['ENDERECO'] != 'Não Informado' else None
                        })
                    resposta['payload'] = {'itens': itens}

    elif tela == 'ZERADOS':
        if state.df_producao_raw is not None:
            df = state.df_producao_raw.copy()
            col_status = df['COL_STATUS_NOME'].iloc[0] if 'COL_STATUS_NOME' in df.columns else None
            col_setor = df['COL_SETOR_NOME'].iloc[0] if 'COL_SETOR_NOME' in df.columns else None
            if col_status:
                df = df[df['REGIAO'] == regiao].copy()
                if f_gerente: df = df[df['Temp_Gerente'].isin(f_gerente)]
                if f_sup: df = df[df['Temp_Sup'].isin(f_sup)]
                if f_setor and col_setor: df = df[df[col_setor].isin(f_setor)]
                if f_tenant and 'Tenant' in df.columns: df = df[df['Tenant'].isin(f_tenant)]

                if not df.empty:
                    df_tec = df.groupby(['Temp_Gerente', 'Temp_Sup', 'Nome_Clean']).agg(
                        SUC=(col_status, lambda x: x.astype(str).str.contains('Concluído com sucesso', case=False, na=False).sum())
                    ).reset_index()
                    mo_sup = df_tec.groupby(['Temp_Gerente', 'Temp_Sup'])['Nome_Clean'].nunique().reset_index(name='MO_ATIVA')
                    df_zer = df_tec[df_tec['SUC'] == 0].copy()
                    qtd_zer = df_zer.groupby(['Temp_Gerente', 'Temp_Sup'])['Nome_Clean'].nunique().reset_index(name='QTD_ZERADOS')

                    stats = pd.merge(mo_sup, qtd_zer, on=['Temp_Gerente', 'Temp_Sup'], how='left').fillna({'QTD_ZERADOS': 0})
                    stats['PCT_ZERADOS'] = (stats['QTD_ZERADOS'] / stats['MO_ATIVA'] * 100).round(1)

                    tot_mo = int(stats['MO_ATIVA'].sum())
                    tot_zer = int(stats['QTD_ZERADOS'].sum())
                    pct_geral = (tot_zer / tot_mo * 100) if tot_mo > 0 else 0

                    detalhes = pd.merge(df_zer[['Temp_Gerente', 'Temp_Sup', 'Nome_Clean']], stats, on=['Temp_Gerente', 'Temp_Sup'], how='left')
                    detalhes = detalhes.sort_values(by=['PCT_ZERADOS', 'Temp_Sup', 'Nome_Clean'], ascending=[False, True, True])

                    resposta['payload'] = {
                        'tot_mo': tot_mo,
                        'tot_zer': tot_zer,
                        'pct_geral': round(pct_geral, 1),
                        'lista': detalhes.rename(columns={'Temp_Gerente': 'GERENTE', 'Temp_Sup': 'SUPERVISOR', 'Nome_Clean': 'TECNICO'}).to_dict(orient='records')
                    }

    elif tela == 'REPAROS':
        geral_vazio = {'hoje': 0, 'amanha': 0, 'futuro': 0, 'tot_d0': 0, 'pct_d0': 0.0, 'tot_d1': 0, 'pct_d1': 0.0, 'pct_futuro': 0.0}
        inst_vazio = {'nio': {'hoje': 0, 'pct': 0.0}, 'tim': {'hoje': 0, 'pct': 0.0}, 'demais': {'hoje': 0, 'pct': 0.0}}
        rep_vazio = {'nio': {'hoje': 0, 'pct': 0.0}, 'tim': {'hoje': 0, 'pct': 0.0}, 'demais': {'hoje': 0, 'pct': 0.0}}

        if state.df_reparos_raw is not None and not state.df_reparos_raw.empty:
            df = state.df_reparos_raw[state.df_reparos_raw['REGIAO'] == regiao].copy()
            df, filtros_hierarquia_ignorados = aplicar_filtros_hierarquia_disponiveis(df, f_gerente, f_sup)
            if f_setor: df = df[df['SETOR'].isin(f_setor)]
            if f_tenant and 'Tenant' in df.columns: df = df[df['Tenant'].isin(f_tenant)]

            if not df.empty:
                rep_geral = compilar_cockpit_reparos(df)
                df_inst = df[df['Macro Atividade'].astype(str).str.contains('INST', case=False, na=False)]
                df_rep_macro = df[df['Macro Atividade'].astype(str).str.contains('REP', case=False, na=False)]

                def format_atrib(sub_df):
                    if sub_df.empty: return {"hoje": 0, "pct": 0.0}
                    h = len(sub_df[sub_df['QUANDO_CAT'] == 'Hoje'])
                    tot = len(sub_df)
                    return {"hoje": h, "pct": round((h / tot * 100) if tot > 0 else 0, 1)}

                cards_macro = {}
                macros_alvo = f_macro if f_macro else ['INST-FTTH', 'REP-FTTH']
                for m_name in macros_alvo:
                    m_df = df[df['Macro Atividade'].astype(str).str.contains(m_name, case=False, na=False)]
                    if not m_df.empty:
                        cards_macro[m_name] = compilar_cockpit_reparos(m_df)

                hoje_exato = datetime.now().date()
                df_hora = df_rep_macro[(df_rep_macro['DATA_ABERTURA'].notna()) & (df_rep_macro['DATA_ABERTURA'].dt.date == hoje_exato)]
                vol_hora = []
                if not df_hora.empty and 'HORA_ABERTURA' in df_hora.columns:
                    vh = df_hora[df_hora['HORA_ABERTURA'] != 'N/A'].groupby('HORA_ABERTURA').size().reset_index(name='VOL')
                    vol_hora = vh.sort_values(by='HORA_ABERTURA').to_dict(orient='records')

                resposta['payload'] = {
                    'geral': rep_geral,
                    'aging_geral': df_rep_macro[df_rep_macro['FAIXA_AGING'] != 'N/A']['FAIXA_AGING'].value_counts().to_dict(),
                    'aging_nio': df_rep_macro[(df_rep_macro['Tenant'] == 'NIO') & (df_rep_macro['FAIXA_AGING'] != 'N/A')]['FAIXA_AGING'].value_counts().to_dict(),
                    'inst': {
                        'nio': format_atrib(df_inst[df_inst['Tenant'] == 'NIO']),
                        'tim': format_atrib(df_inst[df_inst['Tenant'] == 'TIM']),
                        'demais': format_atrib(df_inst[df_inst['Tenant'] == 'DEMAIS'])
                    },
                    'rep': {
                        'nio': format_atrib(df_rep_macro[df_rep_macro['Tenant'] == 'NIO']),
                        'tim': format_atrib(df_rep_macro[df_rep_macro['Tenant'] == 'TIM']),
                        'demais': format_atrib(df_rep_macro[df_rep_macro['Tenant'] == 'DEMAIS'])
                    },
                    'macros': cards_macro,
                    'vol_hora': vol_hora,
                    'filtros_hierarquia_ignorados': filtros_hierarquia_ignorados
                }
            else:
                resposta['payload'] = {
                    'geral': geral_vazio, 'aging_geral': {}, 'aging_nio': {},
                    'inst': inst_vazio, 'rep': rep_vazio,
                    'macros': {}, 'vol_hora': [],
                    'filtros_hierarquia_ignorados': filtros_hierarquia_ignorados
                }
        else:
            resposta['payload'] = {
                'geral': geral_vazio, 'aging_geral': {}, 'aging_nio': {},
                'inst': inst_vazio, 'rep': rep_vazio,
                'macros': {}, 'vol_hora': []
            }

    elif tela == 'REPAROS_SETOR':
        if state.df_reparos_raw is not None:
            df = state.df_reparos_raw[state.df_reparos_raw['REGIAO'] == regiao].copy()
            df, filtros_hierarquia_ignorados = aplicar_filtros_hierarquia_disponiveis(df, f_gerente, f_sup)
            if f_setor: df = df[df['SETOR'].isin(f_setor)]
            if f_macro: df = df[df['Macro Atividade'].isin(f_macro)]
            if f_tenant and 'Tenant' in df.columns: df = df[df['Tenant'].isin(f_tenant)]

            if not df.empty:
                rep_geral = compilar_cockpit_reparos(df)
                linhas = []
                for s_nome, df_s in df.groupby('SETOR'):
                    d_s = compilar_cockpit_reparos(df_s)
                    linhas.append({
                        'SETOR': s_nome,
                        'HOJE': d_s['hoje'],
                        'AMANHA': d_s['amanha'],
                        'FUTURO': d_s['futuro'],
                        'TOTAL': d_s['tot_d0'],
                        'PCT_D0': d_s['pct_d0'],
                        'TOTAL_FUTURO': d_s['tot_d1'],
                        'PCT_FUTURA': d_s['pct_futuro']
                    })
                linhas.sort(key=lambda x: x['TOTAL'], reverse=True)
                resposta['payload'] = {
                    'geral': rep_geral,
                    'setores': linhas,
                    'filtros_hierarquia_ignorados': filtros_hierarquia_ignorados
                }

    return jsonify(resposta)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
