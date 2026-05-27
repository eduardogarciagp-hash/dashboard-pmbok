# =============================================================================
# DASHBOARD EXECUTIVO DE PORTFÓLIO — PMBOK 8ª EDIÇÃO  v3.0
# Leitura direta de XML do MS Project + Section 2 Governança por Projeto
# =============================================================================
# INSTALAÇÃO:  pip install streamlit pandas plotly openpyxl requests
# EXECUÇÃO:    streamlit run dashboard_pmo_v3.py
# =============================================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from datetime import date, datetime, timedelta
import io, json, re, requests, xml.etree.ElementTree as ET

# ──────────────────────────────────────────────────────────────────────────────
# 0. PÁGINA
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard PMO Executivo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────────────
# 1. CSS EXECUTIVO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
*, .stApp { font-family: 'Inter', sans-serif; }
.stApp { background: #EEF5EE; }
/* Caixas de texto brancas na Section C */
.stTextArea textarea {
    background-color: #FFFFFF !important;
    border: 1px solid #CBD5E1 !important;
    border-radius: 6px !important;
}
/* Barra do expander (nome do projeto) branca */
.stExpander {
    background-color: #FFFFFF !important;
    border: 1px solid #CBD5E1 !important;
    border-radius: 8px !important;
}
.stExpander > details > summary {
    background-color: #FFFFFF !important;
    border-radius: 8px !important;
}
.section-title {
    color:#1B2A4A; font-size:14px; font-weight:700;
    border-left:4px solid #2563EB; padding-left:10px;
    margin:20px 0 10px 0; text-transform:uppercase; letter-spacing:.06em;
}
.kpi-card {
    background:#fff; border-radius:10px; padding:16px 20px 14px 20px;
    border-left:5px solid #2563EB;
    box-shadow:0 1px 6px rgba(0,0,0,.08);
}
.kpi-label { color:#9AA5BE; font-size:11px; font-weight:600;
             text-transform:uppercase; letter-spacing:.07em; }
.kpi-value { color:#4A5568 !important; font-size:28px; font-weight:700; margin:4px 0 2px 0; }
.kpi-sub   { color:#6B7A99; font-size:11px; }
.badge-green  { background:#D4EDDA; color:#1A6B3A; padding:2px 8px;
                border-radius:12px; font-size:11px; font-weight:600; }
.badge-yellow { background:#FFF3CD; color:#856404; padding:2px 8px;
                border-radius:12px; font-size:11px; font-weight:600; }
.badge-red    { background:#F8D7DA; color:#842029; padding:2px 8px;
                border-radius:12px; font-size:11px; font-weight:600; }
.proj-header-bdf  { background:#1B2A4A; color:#fff; padding:8px 14px;
                    border-radius:6px; font-weight:700; font-size:13px; margin:14px 0 6px 0; }
.proj-header-ce   { background:#1B4332; color:#fff; padding:8px 14px;
                    border-radius:6px; font-weight:700; font-size:13px; margin:14px 0 6px 0; }
.proj-header-ea   { background:#5C2D19; color:#fff; padding:8px 14px;
                    border-radius:6px; font-weight:700; font-size:13px; margin:14px 0 6px 0; }
.crit-row-red     { background:#FFF0F0; border-left:4px solid #DC3545;
                    border-radius:6px; padding:10px 14px; margin:4px 0; }
.crit-row-yellow  { background:#FFFAED; border-left:4px solid #FFC107;
                    border-radius:6px; padding:10px 14px; margin:4px 0; }
.crit-label { font-size:10px; font-weight:600; color:#6B7A99;
              text-transform:uppercase; letter-spacing:.05em; }
.crit-value { font-size:12px; color:#1B2A4A; margin-top:2px; }
hr.section-sep { border:none; border-top:2px solid #E2E8F0; margin:20px 0; }
.stDataFrame { border-radius:8px; overflow:hidden; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 2. PARSER XML MS PROJECT
# ──────────────────────────────────────────────────────────────────────────────
NS = "http://schemas.microsoft.com/project"

def _tag(t):
    return f"{{{NS}}}{t}"

def _cfn(cf, k):
    v = cf.get(k, '')
    try:
        return float(v.replace(',', '.')) if v else None
    except:
        return None

def parse_xml(file_obj, project_name: str) -> list[dict]:
    """Parse MS Project XML and return list of task dicts."""
    try:
        tree = ET.parse(file_obj)
    except ET.ParseError:
        return []
    root = tree.getroot()

    ext_map = {}
    for ea in root.findall(f".//{_tag('ExtendedAttribute')}"):
        fid = ea.findtext(_tag('FieldID'))
        fn  = ea.findtext(_tag('FieldName'))
        if fid and fn:
            ext_map[fid] = fn

    tasks = []
    for task in root.findall(f".//{_tag('Task')}"):
        def g(t):
            el = task.find(_tag(t))
            return el.text if el is not None else None

        uid = g('UID')
        if uid == '0':
            continue
        name = g('Name')
        if not name:
            continue

        level      = int(g('OutlineLevel') or 99)
        pct        = float(g('PercentComplete') or 0)
        start      = (g('Start')          or '')[:10]
        finish     = (g('Finish')         or '')[:10]
        b_finish   = (g('BaselineFinish') or '')[:10]
        is_mile    = g('Milestone') == '1'
        is_sum     = g('Summary')   == '1'

        cf = {}
        for ev in task.findall(_tag('ExtendedAttribute')):
            fid = ev.findtext(_tag('FieldID'))
            val = ev.findtext(_tag('Value'))
            fn  = ext_map.get(fid, fid)
            cf[fn] = val

        pv     = _cfn(cf, 'Número4') or 0
        ev_val = _cfn(cf, 'Número6') or 0

        # ── IDP — regra por formato de arquivo ───────────────────────────────
        # Cockpit Engenharia: Texto3 contém o IDP calculado pela fórmula
        #   do MS Project (fórmula customizada da empresa). Número6 = 0 neste
        #   template, por isso Texto3 é a fonte correta.
        # BDF / Esteira Analytics: usam Número6 (EV dias) e Número4 (PV dias).
        #   IDP = Número6 / Número4.
        # Detecção automática: se Texto3 tem valor numérico válido E Número6=0
        #   → usa Texto3 (Cockpit). Caso contrário → Número6/Número4.

        idp_direct = None
        _t3 = (cf.get('Texto3', '') or '').replace(',', '.').strip()
        try:
            _t3_val = float(_t3)
            if _t3_val > 0:
                idp_direct = _t3_val
        except:
            pass

        if idp_direct is not None and ev_val == 0:
            # Cockpit: usa fórmula MS Project via Texto3
            spi = round(idp_direct, 4)
        elif pv > 0 and ev_val > 0:
            # BDF / Esteira: IDP = EV / PV (dias)
            spi = round(ev_val / pv, 4)
        else:
            spi = None

        status = (cf.get('Texto2', '') or '').strip()
        resp   = (cf.get('Texto4', '') or cf.get('Texto5', '')).strip()
        desvio = _cfn(cf, 'Número2')

        tasks.append({
            'projeto': project_name, 'uid': uid, 'nome': name,
            'nivel': level, 'pct': pct, 'pv': pv, 'ev': ev_val, 'spi': spi,
            'status': status, 'resp': resp,
            'inicio': start, 'termino': finish, 'baseline_termino': b_finish,
            'is_milestone': is_mile, 'is_summary': is_sum,
            'desvio_dias': desvio,
        })
    return tasks


def build_df(tasks: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(tasks)
    if df.empty:
        return df
    df['pct_ratio'] = df['pct'] / 100
    df['spi_num']   = pd.to_numeric(df['spi'], errors='coerce')
    for col in ['inicio', 'termino', 'baseline_termino']:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    return df


# ──────────────────────────────────────────────────────────────────────────────
# 3. INSIGHTS IA (Claude API)
# ──────────────────────────────────────────────────────────────────────────────
# Logo embutida em base64
_LOGO_B64 = "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"

ANTHROPIC_API = "https://api.anthropic.com/v1/messages"

@st.cache_data(show_spinner=False, ttl=3600)
def gerar_governanca_ia(proj_name: str, tarefas_json: str) -> list:
    prompt = (
        f"Voce e um especialista senior em PMO certificado pelo PMI, com dominio em PMBOK 8a Edicao.\n\n"
        f"Analise os dados do projeto '{proj_name}' exportados do MS Project.\n"
        f"Identifique os 2 a 4 pontos criticos mais relevantes para apresentacao a diretores executivos.\n\n"
        f"Para cada ponto critico gere:\n"
        f"- titulo: titulo executivo conciso (max 80 chars) com o indicador principal\n"
        f"- impacto: impacto no negocio em 3-5 frases com consequencias concretas para a empresa\n"
        f"- causa: causa raiz tecnica identificada nos dados com nomes de tarefas, SPI e datas\n"
        f"- plano: plano de acao com etapas numeradas, responsaveis genericos e prazos realistas\n\n"
        f"Dados:\n{tarefas_json}\n\n"
        f"Responda SOMENTE em JSON valido sem markdown:\n"
        f'[{{"titulo":"...","impacto":"...","causa":"...","plano":"..."}}]'
    )
    try:
        resp = requests.post(
            ANTHROPIC_API,
            headers={"Content-Type": "application/json"},
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 3000,
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=60,
        )
        data = resp.json()
        text = data["content"][0]["text"]
        text = re.sub(r'```json|```', '', text).strip()
        return json.loads(text)
    except Exception:
        return []

@st.cache_data(show_spinner=False, ttl=3600)
def gerar_insights_ia(proj_name: str, criticos_json: str) -> dict:
    """Chama Claude para gerar insights de governança para as tarefas críticas."""
    prompt = f"""Você é um especialista em PMO e PMBOK 8ª Edição. 
Analise os dados abaixo do projeto '{proj_name}' e para cada tarefa crítica (SPI < 0,95) gere:
1. Impacto no Negócio (1-2 frases objetivas)
2. Causa Raiz (hipótese técnica baseada nos dados)
3. Plano de Ação (ações concretas com responsáveis genéricos e prazo)

Dados: {criticos_json}

Responda SOMENTE em JSON válido, sem markdown, no formato:
[
  {{
    "nome": "nome da tarefa",
    "nivel": 2,
    "spi": 0.50,
    "impacto": "...",
    "causa_raiz": "...",
    "plano_acao": "..."
  }}
]"""

    try:
        resp = requests.post(
            ANTHROPIC_API,
            headers={"Content-Type": "application/json"},
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 1500,
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=30,
        )
        data = resp.json()
        text = data["content"][0]["text"]
        text = re.sub(r'```json|```', '', text).strip()
        return {"status": "ok", "items": json.loads(text)}
    except Exception as e:
        return {"status": "error", "msg": str(e), "items": []}


# ──────────────────────────────────────────────────────────────────────────────
# 4. HELPERS UI
# ──────────────────────────────────────────────────────────────────────────────
def badge(spi):
    if spi is None:
        return '<span class="badge-yellow">N/A</span>'
    if spi >= 0.99:
        return f'<span class="badge-green">IDP {spi:.2f} ▲</span>'
    if spi >= 0.95:
        return f'<span class="badge-yellow">IDP {spi:.2f} !</span>'
    return f'<span class="badge-red">IDP {spi:.2f} ▼</span>'

def idp_face(spi):
    if spi is None:
        return "", "#9AA5BE", "N/A"
    if spi >= 0.99:
        return "", "#059669", "Em dia"
    if spi >= 0.95:
        return "", "#D97706", "Em alerta"
    return "", "#DC2626", "Em atraso"

def kpi_card(label, valor, sub="", cor="#2563EB"):
    st.markdown(f"""
    <div class="kpi-card" style="border-left-color:{cor}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{valor}</div>
        <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

PROJ_CSS = {
    "Business Data Fabric": "proj-header-bdf",
    "Cockpit Engenharia":   "proj-header-ce",
    "Esteira Analytics":    "proj-header-ea",
}

PROJETOS_FIXOS = [
    "Business Data Fabric",
    "Cockpit Engenharia",
    "Esteira Analytics",
    "Demonstrativo Financeiro",
    "Automacao Order to Cash",
    "IA Copilot",
]
# Lista dinâmica (fixos + extras adicionados pelo usuário)
# Resolvida após inicialização do session_state (ver abaixo)


# ──────────────────────────────────────────────────────────────────────────────
# 5. SIDEBAR — UPLOAD & FILTROS
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('## 📊 Dashboard PMO')
    st.markdown('---')

    # ── Importar XML ──────────────────────────────────────────────────────────
    st.markdown('### 📂 Importar Projetos (XML)')
    st.caption('Exporte pelo MS Project: Arquivo → Salvar Como → XML do Project')
    uploaded = st.file_uploader(
        'Selecione um ou mais arquivos .xml',
        type=['xml'],
        accept_multiple_files=True,
    )
    st.markdown('---')

    # ── Configurações ─────────────────────────────────────────────────────────
    st.markdown('### ⚙️ Configurações')
    data_ref    = st.date_input('📅 Data de Referência', value=date.today())
    filtro_proj = st.multiselect('🔍 Filtrar por Projeto', options=[], key='filtro_proj')
    spi_limiar  = st.slider('⚠️ Limiar IDP (crítico)', 0.70, 1.00, 0.95, 0.01)
    usar_ia     = st.toggle('🤖 Insights com IA (Claude API)', value=False)
    st.markdown('---')

    # ── Projetos do Portfólio ────────────────────────────────────────────────────
    st.markdown('### ✏️ Projetos do Portfolio')
    st.caption('Edite IDP e datas de cada projeto.')
    if 'idp_override' not in st.session_state:
        st.session_state.idp_override = {}
    if 'proj_manual' not in st.session_state:
        st.session_state.proj_manual = {}
    if 'projetos_extras' not in st.session_state:
        st.session_state.projetos_extras = []

    # Adicionar novo projeto
    with st.expander("➕ Adicionar projeto", expanded=False):
        novo_nome = st.text_input("Nome do projeto", placeholder="Ex: Projeto X", key="novo_proj_nome")
        if st.button("Adicionar", key="btn_add_proj", use_container_width=True):
            nome = novo_nome.strip()
            todos = PROJETOS_FIXOS + st.session_state.projetos_extras
            if nome and nome not in todos:
                st.session_state.projetos_extras.append(nome)
                st.session_state.proj_manual[nome] = {'idp': 0.0, 'pct': 0, 'inicio': '', 'termino': ''}
                st.rerun()
            elif not nome:
                st.warning("Digite um nome.")
            else:
                st.warning("Projeto já existe.")

    # Remover projetos extras
    if st.session_state.projetos_extras:
        with st.expander("🗑️ Remover projeto", expanded=False):
            proj_rem = st.selectbox("Selecione", st.session_state.projetos_extras, key="sel_rem_proj")
            if st.button("Remover", key="btn_rem_proj", use_container_width=True):
                st.session_state.projetos_extras.remove(proj_rem)
                st.session_state.proj_manual.pop(proj_rem, None)
                st.session_state.idp_override.pop(proj_rem, None)
                st.rerun()

    _sb_idp_placeholder = st.empty()
    st.markdown('---')

    # ── Exportar ──────────────────────────────────────────────────────────────
    st.markdown('### 📥 Exportar')
    _sb_export_placeholder = st.empty()


# ──────────────────────────────────────────────────────────────────────────────
# 6. CARREGAR DADOS
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_projects(files_meta: list) -> pd.DataFrame:
    all_tasks = []
    for name, content in files_meta:
        proj_name = re.sub(r'_\d{4}_\d{2}_\d{2}\.xml$', '', name).replace('_', ' ').strip()
        tasks = parse_xml(io.BytesIO(content), proj_name)
        all_tasks.extend(tasks)
    return build_df(all_tasks)


if uploaded:
    files_meta = [(f.name, f.read()) for f in uploaded]
    df_xml = load_projects(files_meta)
else:
    df_xml = pd.DataFrame()

if 'proj_manual' not in st.session_state:
    st.session_state.proj_manual = {}

def _ensure_proj(proj):
    if not df_xml.empty:
        rows = df_xml[(df_xml['projeto']==proj) & (df_xml['nivel']==1)]
        if not rows.empty:
            return rows
    m = st.session_state.proj_manual.get(proj, {})
    idp_v = float(m['idp']) if m.get('idp') else None
    return pd.DataFrame([{
        'projeto': proj, 'uid': '0', 'nome': proj,
        'nivel': 1, 'pct': float(m.get('pct', 0)),
        'pv': 0.0, 'ev': 0.0, 'spi': idp_v,
        'status': '', 'resp': '',
        'inicio':   pd.Timestamp(m['inicio'])  if m.get('inicio')  else pd.NaT,
        'termino':  pd.Timestamp(m['termino']) if m.get('termino') else pd.NaT,
        'baseline_termino': pd.NaT,
        'is_milestone': False, 'is_summary': False, 'desvio_dias': None,
    }])

# Lista dinâmica: fixos + extras adicionados pelo usuário
PROJETOS_PORTFOLIO = PROJETOS_FIXOS + st.session_state.get('projetos_extras', [])
projetos_disp = PROJETOS_PORTFOLIO

dfs = []
for proj in PROJETOS_PORTFOLIO:
    dfs.append(_ensure_proj(proj))
    if not df_xml.empty:
        filhas = df_xml[(df_xml['projeto']==proj) & (df_xml['nivel']>1)]
        if not filhas.empty:
            dfs.append(filhas)

df_view = build_df(pd.concat(dfs, ignore_index=True)) if dfs else pd.DataFrame()
# Lista dinâmica: fixos + extras


# ──────────────────────────────────────────────────────────────────────────────
# 7. CABEÇALHO
# ──────────────────────────────────────────────────────────────────────────────
# ── Modo Apresentação ────────────────────────────────────────────────────────
if 'modo_apresentacao' not in st.session_state:
    st.session_state.modo_apresentacao = False

# Cabeçalho com logo via components (suporta base64 longo)
_data_ref_str   = data_ref.strftime('%d/%m/%Y')
_n_proj         = len(PROJETOS_FIXOS)
st.components.v1.html(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
  body{{margin:0;padding:0;background:transparent;font-family:'Inter',sans-serif;}}
  .hdr{{display:flex;align-items:center;justify-content:space-between;
        padding:4px 0 6px 0;}}
  .hdr h1{{color:#1B2A4A;font-size:30px;font-weight:700;margin:0;line-height:1.2;}}
  .hdr p{{color:#9AA5BE;font-size:13px;margin:4px 0 0 0;}}
  .hdr img{{height:104px;object-fit:contain;opacity:0.92;}}
</style>
<div class='hdr'>
  <div>
    <h1>Dashboard Executivo - Digital</h1>
    <p>PMBOK 8ª Ed. &middot; Referência: {_data_ref_str} &middot; {_n_proj} projetos no portfólio</p>
  </div>
  <img src='data:image/png;base64,{_LOGO_B64}' alt='Logo'>
</div>
""", height=120, scrolling=False)

_apresentando = st.session_state.modo_apresentacao

# JS: colapsa sidebar no modo apresentação, expande no modo edição
# JS robusto: colapsa/expande sidebar testando múltiplos seletores
_collapsed = 'true' if _apresentando else 'false'
st.components.v1.html(f"""
<script>
(function(){{
  var want_collapsed = {_collapsed};
  var attempts = 0;

  function findCollapseBtn() {{
    var doc = window.parent.document;
    // Seletores em ordem de prioridade (Streamlit Cloud + local)
    var selectors = [
      '[data-testid="stSidebarCollapseButton"] button',
      '[data-testid="stSidebarCollapseButton"]',
      'section[data-testid="stSidebar"] button:first-child',
      '[data-testid="stSidebar"] button',
    ];
    for (var i = 0; i < selectors.length; i++) {{
      var el = doc.querySelector(selectors[i]);
      if (el) return el;
    }}
    return null;
  }}

  function findExpandBtn() {{
    var doc = window.parent.document;
    var selectors = [
      '[data-testid="stSidebarCollapsedControl"] button',
      '[data-testid="stSidebarCollapsedControl"]',
      '[aria-label="Open sidebar"]',
      '[aria-label="Abrir barra lateral"]',
    ];
    for (var i = 0; i < selectors.length; i++) {{
      var el = doc.querySelector(selectors[i]);
      if (el) return el;
    }}
    return null;
  }}

  function isSidebarOpen() {{
    var doc = window.parent.document;
    var sidebar = doc.querySelector('[data-testid="stSidebar"]');
    if (!sidebar) return true;
    // Collapsed control visível = sidebar fechada
    var collapsedCtrl = doc.querySelector('[data-testid="stSidebarCollapsedControl"]');
    if (collapsedCtrl) return false;
    // aria-expanded = false = fechada
    if (sidebar.getAttribute('aria-expanded') === 'false') return false;
    // Verifica largura
    var rect = sidebar.getBoundingClientRect();
    if (rect.width < 50) return false;
    return true;
  }}

  var timer = setInterval(function() {{
    attempts++;
    if (attempts > 40) {{ clearInterval(timer); return; }}

    var open = isSidebarOpen();

    if (want_collapsed && open) {{
      var btn = findCollapseBtn();
      if (btn) {{ btn.click(); clearInterval(timer); }}
    }} else if (!want_collapsed && !open) {{
      var btn2 = findExpandBtn();
      if (btn2) {{ btn2.click(); clearInterval(timer); }}
    }} else {{
      // Estado já correto
      clearInterval(timer);
    }}
  }}, 200);
}})();
</script>
""", height=0)

st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)

if df_view.empty:
    st.info("📂 Importe arquivos XML do MS Project pela barra lateral para visualizar o dashboard.")
    st.stop()


# ──────────────────────────────────────────────────────────────────────────────
# 8. SECTION KPIs
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">A. Governança Estratégica — KPIs do Portfólio</div>',
            unsafe_allow_html=True)

df_root = df_view[
    (df_view['nivel'] == 1) &
    (~df_view['nome'].isin(['MS Project_Teste_Formulas_2', 'NOME DO PROJETO']))
]
n_projetos_calc  = df_view['projeto'].nunique()
pct_media_calc   = df_root['pct'].mean() if not df_root.empty else 0
crits_count_calc = df_view[(df_view['spi_num'] < spi_limiar) & df_view['spi_num'].notna()]['projeto'].nunique()
marcos_tot_calc  = df_view[df_view['is_milestone']].shape[0]

# IDP por projeto — sempre 6 projetos fixos
_l1_rows = df_view[
    (df_view['nivel'] == 1) &
    (~df_view['nome'].isin(['MS Project_Teste_Formulas_2', 'NOME DO PROJETO']))
]

idp_por_projeto = {}
for proj in PROJETOS_PORTFOLIO:
    grp = _l1_rows[_l1_rows['projeto'] == proj]
    if grp.empty:
        # Sem dados: usa override manual
        idp_por_projeto[proj] = float(st.session_state.idp_override.get(proj, 0)) or None
        continue
    total_pv = grp['pv'].sum()
    total_ev = grp['ev'].sum()
    if total_ev > 0 and total_pv > 0:
        idp_por_projeto[proj] = round(total_ev / total_pv, 4)
    else:
        valid = grp[grp['spi_num'].notna() & (grp['pv'] > 0)]
        if not valid.empty:
            idp_por_projeto[proj] = round(
                (valid['spi_num'] * valid['pv']).sum() / valid['pv'].sum(), 4)
        elif grp['spi_num'].notna().any():
            idp_por_projeto[proj] = round(grp['spi_num'].dropna().mean(), 4)
        else:
            idp_por_projeto[proj] = float(st.session_state.idp_override.get(proj, 0)) or None

_vals = [v for v in idp_por_projeto.values() if v]
spi_medio_calc = round(sum(_vals) / len(_vals), 4) if _vals else None

# ── IDP por projeto: inputs no sidebar, lidos aqui ──────────────────────────
if 'idp_override' not in st.session_state:
    st.session_state.idp_override = {}

# Renderiza inputs no sidebar — todos os 6 projetos sempre visíveis
idp_por_projeto_final = {}
with _sb_idp_placeholder.container():
    if not st.session_state.get('modo_apresentacao', False):
        for proj in PROJETOS_PORTFOLIO:
            # IDP calculado do XML ou do override manual
            idp_calc = idp_por_projeto.get(proj)
            st.markdown(f"**{proj[:30]}**")
            col_idp, col_pct = st.columns(2)
            with col_idp:
                idp_editado = st.number_input(
                    "IDP", min_value=0.0, max_value=2.0,
                    value=float(st.session_state.idp_override.get(proj, idp_calc or 0.0)),
                    step=0.01, format='%.2f',
                    key=f'idp_input_{proj}'
                )
            with col_pct:
                pct_editado = st.number_input(
                    "% Concl.", min_value=0, max_value=100,
                    value=int(st.session_state.proj_manual.get(proj, {}).get('pct', 0)),
                    step=1, key=f'pct_input_{proj}'
                )
            # Datas (apenas para projetos sem XML)
            has_xml = not df_xml.empty and proj in df_xml['projeto'].values
            if not has_xml:
                col_ini, col_fim = st.columns(2)
                with col_ini:
                    ini_val = st.session_state.proj_manual.get(proj, {}).get('inicio', '')
                    ini_str = st.text_input("Início", value=ini_val,
                        placeholder="AAAA-MM-DD", key=f'ini_{proj}')
                with col_fim:
                    fim_val = st.session_state.proj_manual.get(proj, {}).get('termino', '')
                    fim_str = st.text_input("Término", value=fim_val,
                        placeholder="AAAA-MM-DD", key=f'fim_{proj}')
                st.session_state.proj_manual[proj] = {
                    'idp': idp_editado, 'pct': pct_editado,
                    'inicio': ini_str, 'termino': fim_str,
                }
            else:
                st.session_state.proj_manual[proj] = {
                    'idp': idp_editado, 'pct': pct_editado,
                }
            st.session_state.idp_override[proj] = idp_editado
            idp_por_projeto_final[proj] = idp_editado
            st.markdown("<hr style='margin:6px 0;border-color:#eee'>", unsafe_allow_html=True)
    else:
        for proj in PROJETOS_PORTFOLIO:
            idp_calc = idp_por_projeto.get(proj)
            idp_por_projeto_final[proj] = float(st.session_state.idp_override.get(proj, idp_calc or 0.0))


# KPIs recalculados a partir dos IDPs editados
_idp_vals = [v for v in idp_por_projeto_final.values() if v]
spi_medio = round(sum(_idp_vals) / len(_idp_vals), 4) if _idp_vals else None
n_projetos = len(idp_por_projeto_final)
pct_media  = pct_media_calc

# Upstream / Downstream — editáveis no sidebar
if 'kpi_upstream'   not in st.session_state: st.session_state.kpi_upstream   = 0
if 'kpi_downstream' not in st.session_state: st.session_state.kpi_downstream = 0

# KPI cards (primeira linha)
c1, c2, c3, c4, c5 = st.columns(5)
with c1: kpi_card("PROJETOS ATIVOS", str(n_projetos), "monitorados", "#2563EB")
with c2: kpi_card("CONCLUSÃO MÉDIA", f"{pct_media:.1f}%", "do portfólio", "#059669")
with c3:
    spi_str = f"{spi_medio:.2f}" if spi_medio else "N/A"
    cor3 = "#DC2626" if (spi_medio and spi_medio < 0.95) else ("#D97706" if (spi_medio and spi_medio < 0.99) else "#059669")
    kpi_card("IDP PORTFÓLIO", spi_str, "índice de desempenho", cor3)
with c4:
    if not _apresentando:
        st.markdown("<div class='kpi-card' style='border-left-color:#0EA5E9'>", unsafe_allow_html=True)
        st.markdown("<div class='kpi-label'>PROJETOS EM UPSTREAM</div>", unsafe_allow_html=True)
        st.session_state.kpi_upstream = st.number_input(
            "upstream", min_value=0, max_value=50,
            value=st.session_state.kpi_upstream,
            step=1, label_visibility="collapsed", key="inp_upstream"
        )
        st.markdown("<div class='kpi-sub'>em iniciação / planejamento</div></div>", unsafe_allow_html=True)
    else:
        kpi_card("PROJETOS EM UPSTREAM", str(st.session_state.kpi_upstream),
                 "em iniciação / planejamento", "#0EA5E9")
with c5:
    if not _apresentando:
        st.markdown("<div class='kpi-card' style='border-left-color:#7C3AED'>", unsafe_allow_html=True)
        st.markdown("<div class='kpi-label'>PROJETOS EM DOWNSTREAM</div>", unsafe_allow_html=True)
        st.session_state.kpi_downstream = st.number_input(
            "downstream", min_value=0, max_value=50,
            value=st.session_state.kpi_downstream,
            step=1, label_visibility="collapsed", key="inp_downstream"
        )
        st.markdown("<div class='kpi-sub'>em execução / encerramento</div></div>", unsafe_allow_html=True)
    else:
        kpi_card("PROJETOS EM DOWNSTREAM", str(st.session_state.kpi_downstream),
                 "em execução / encerramento", "#7C3AED")

# Carinhas por projeto (segunda linha) — 6 projetos, caixas de tamanho uniforme
cols_face = st.columns(len(idp_por_projeto_final))
for i, (proj, idp_val) in enumerate(idp_por_projeto_final.items()):
    face, cor_face, label_face = idp_face(idp_val)
    idp_txt = f"{idp_val:.2f}" if idp_val else "N/A"
    with cols_face[i]:
        st.markdown(f"""
<div style='background:#fff;border-radius:10px;
            padding:10px 12px 10px 12px;
            border-left:5px solid {cor_face};
            box-shadow:0 1px 6px rgba(0,0,0,.08);
            margin-top:12px;
            min-height:90px;
            display:flex;flex-direction:column;justify-content:space-between;'>
  <div style='font-size:10px;font-weight:700;color:#4A5568;
              text-transform:uppercase;letter-spacing:.05em;
              line-height:1.4;word-break:break-word;'>{proj}</div>
  <div>
    <div style='font-size:24px;font-weight:700;color:{cor_face};
                line-height:1.1;margin-top:6px;'>IDP {idp_txt}</div>
    <div style='font-size:11px;color:{cor_face};font-weight:600;
                margin-top:2px;'>{label_face}</div>
  </div>
</div>""", unsafe_allow_html=True)
st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 9. SECTION 1 — ROADMAP: UMA LINHA POR PROJETO COM MARCOS E MODAL
# ──────────────────────────────────────────────────────────────────────────────
import streamlit.components.v1 as components

st.markdown('<div class="section-title">B. Roadmap Executivo & Marcos de Valor</div>',
            unsafe_allow_html=True)

# ── Coleta dados por projeto ──────────────────────────────────────────────────
SKIP_NAMES = ['MS Project_Teste_Formulas_2', 'NOME DO PROJETO']

def _ts(dt):
    import pandas as _pd
    if _pd.isna(dt): return None
    return int(dt.normalize().value // 1_000_000)

import json as _json

proj_data = []
for proj in PROJETOS_PORTFOLIO:
    df_p = df_view[df_view['projeto'] == proj] if not df_view.empty else pd.DataFrame()

    # Linha raiz L1 do projeto para datas início/fim
    l1 = pd.DataFrame()
    if not df_p.empty:
        l1 = df_p[
            (df_p['nivel'] == 1) &
            (~df_p['nome'].isin(SKIP_NAMES)) &
            df_p['inicio'].notna() &
            df_p['termino'].notna()
        ]
        if l1.empty:
            l1 = df_p[df_p['inicio'].notna() & df_p['termino'].notna()]

    # Se sem datas: usa dados manuais do sidebar ou placeholder
    if l1.empty:
        m = st.session_state.proj_manual.get(proj, {})
        ini_s = m.get('inicio', '')
        fim_s = m.get('termino', '')
        try:
            ini_ts = pd.Timestamp(ini_s) if ini_s else pd.Timestamp('2026-01-01')
            fim_ts = pd.Timestamp(fim_s) if fim_s else pd.Timestamp('2026-12-31')
        except:
            ini_ts = pd.Timestamp('2026-01-01')
            fim_ts = pd.Timestamp('2026-12-31')
        pct_v = float(m.get('pct', 0))
        l1 = pd.DataFrame([{
            'projeto': proj, 'nome': proj, 'nivel': 1,
            'pct': pct_v, 'spi_num': None,
            'inicio': ini_ts, 'termino': fim_ts,
            'baseline_termino': pd.NaT,
            'is_milestone': False, 'is_summary': False,
        }])

    r0 = l1.iloc[0]
    idp_val = idp_por_projeto_final.get(proj)

    # Cor neutra — barra cinza, progresso por %
    bar_cor   = "#CBD5E1"
    fase_nome = ""

    # ── Marcos curados por análise PMO PMBOK 8ª Edição ──────────────────────────
    # Marcos definidos com critério técnico: eventos significativos que representam
    # conclusão de entrega, decisão crítica ou transição de fase.
    # Curados manualmente após análise do XML de cada projeto.

    MARCOS_CURADOS = {
        "Business Data Fabric": [
            {"nome": "Kick-off Projeto",               "data": "2026-03-25", "pct": 100},
            {"nome": "Assinatura Contrato MS Fabric",  "data": "2026-06-09", "pct": 0},
            {"nome": "Assinatura Contrato Implantação","data": "2026-06-09", "pct": 0},
            {"nome": "Ativação SKU / Go-Live Fabric",  "data": "2026-06-23", "pct": 0},
            {"nome": "Publicação Política Governança", "data": "2026-04-30", "pct": 0},
            {"nome": "Demanda Concluída – SAC GO-LIVE","data": "2026-07-02", "pct": 0},
            {"nome": "Encerramento do Projeto",        "data": "2026-12-11", "pct": 0},
        ],
        "Cockpit Engenharia": [
            {"nome": "Aprovação TAP",                  "data": "2026-05-22", "pct": 64},
            {"nome": "Kick-off Fase 1",                "data": "2026-06-05", "pct": 100},
            {"nome": "Assinatura Contrato Fase 2",     "data": "2026-08-14", "pct": 0},
            {"nome": "GO-LIVE",                        "data": "2026-10-29", "pct": 0},
        ],
        "Esteira Analytics": [
            {"nome": "Definição Entregas Q1",          "data": "2026-01-02", "pct": 100},
            {"nome": "Assinatura Contrato Estratégia", "data": "2026-03-06", "pct": 100},
            {"nome": "Go Live Estratégia de Dados",    "data": "2026-04-30", "pct": 100},
            {"nome": "Go Live CEO Digital Boardroom",  "data": "2026-04-30", "pct": 100},
            {"nome": "Conclusão Entregas Q1",          "data": "2026-05-27", "pct": 87},
            {"nome": "Conclusão Entregas Q2",          "data": "2026-06-30", "pct": 0},
            {"nome": "Conclusão Entregas Q3",          "data": "2026-09-30", "pct": 0},
            {"nome": "Conclusão Entregas Q4",          "data": "2026-12-31", "pct": 0},
        ],
    }

    # Busca marcos curados para este projeto; fallback para L2 se projeto não mapeado
    marcos_def = MARCOS_CURADOS.get(proj)
    if marcos_def:
        marcos = []
        for md in marcos_def:
            dt = pd.Timestamp(md["data"])
            pct_m = md["pct"]
            concluido = pct_m >= 100
            # Atrasado: % < 100 e data já passou
            atrasado = (not concluido) and (dt.date() < date.today())
            marcos.append({
                "nome":      md["nome"],
                "termino":   int(dt.normalize().value // 1_000_000),
                "baseline":  None,
                "pct":       float(pct_m),
                "status":    "",
                "resp":      "",
                "nivel":     2,
                "concluido": concluido,
                "atrasado":  atrasado,
            })
        df_miles = pd.DataFrame()  # não usado abaixo quando marcos já preenchido
    else:
        marcos = []
        df_miles = df_p[
            (df_p['nivel'] == 2) &
            df_p['termino'].notna() &
            (~df_p['nome'].isin(SKIP_NAMES))
        ].head(8)

    # Se marcos já foram preenchidos pelos curados, não sobrescrever
    if not marcos:
        for _, m in df_miles.iterrows():
            concluido = m['pct'] >= 100
            atrasado  = (
                pd.notna(m['baseline_termino']) and
                pd.notna(m['termino']) and
                m['termino'] > m['baseline_termino']
            )
            marcos.append({
                "nome":      m['nome'][:60],
                "termino":   _ts(m['termino']),
                "baseline":  _ts(m['baseline_termino']),
                "pct":       float(m['pct']),
                "status":    m['status'][:30] if m['status'] else "",
                "resp":      m['resp'][:30] if m['resp'] else "",
                "nivel":     int(m['nivel']),
                "concluido": concluido,
                "atrasado":  atrasado,
            })

    # Subfases L2 para resumo no modal da barra
    subfases = []
    for _, s in df_p[(df_p['nivel']==2) & (~df_p['nome'].isin(SKIP_NAMES))].iterrows():
        idp_s = round(s['spi_num'],2) if pd.notna(s['spi_num']) else None
        subfases.append({
            "nome": s['nome'][:50],
            "pct":  float(s['pct']),
            "idp":  idp_s,
            "termino": _ts(s['termino']),
        })

    # Sem segmentos de fase — barra simples com % de conclusão
    _segs = []

    proj_data.append({
        "projeto":  proj,
        "inicio":   _ts(r0['inicio']),
        "termino":  _ts(r0['termino']),
        "baseline": _ts(r0['baseline_termino']),
        "pct":      float(r0['pct']),
        "idp":      idp_val,
        "cor":      bar_cor,
        "fase":     fase_nome,
        "segs":     _segs,
        "marcos":   marcos,
        "subfases": subfases,
    })

if proj_data:
    hoje_ts       = int(pd.Timestamp(date.today()).normalize().value // 1_000_000)
    proj_json     = _json.dumps(proj_data, ensure_ascii=False)
    row_h         = 90
    total_h       = len(proj_data) * row_h + 130
    _modo_apres_js = 'true' if _apresentando else 'false'

    html_roadmap = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'Inter','Segoe UI',sans-serif;}}
body{{background:#0F1623;color:#E2E8F0;overflow:visible;}}
#wrap{{position:relative;width:100%;padding:0 16px 16px 16px;overflow:visible;}}

/* HEADER */
#header-row{{position:relative;height:32px;margin-bottom:2px;}}
.month-cell{{
  position:absolute;height:100%;
  display:flex;align-items:center;justify-content:center;
  font-size:11px;font-weight:700;color:#475569;
  text-transform:uppercase;letter-spacing:.07em;
  border-right:1px solid #1E2D42;
}}

/* BODY */
#body{{position:relative;width:100%;overflow:visible;}}
.gridline{{position:absolute;top:0;bottom:0;width:1px;background:rgba(100,116,139,.15);pointer-events:none;}}
.today-line{{position:absolute;top:0;bottom:0;width:2px;background:rgba(99,179,237,.6);pointer-events:none;z-index:5;}}
.today-lbl{{
  position:absolute;top:4px;
  font-size:11px;color:#63B3ED;font-weight:700;
  background:#0F1623;padding:1px 5px;border-radius:3px;white-space:nowrap;
  transform:translateX(-50%);
}}

/* PROJECT ROW */
.proj-row{{
  position:absolute;width:100%;
  border-bottom:1px solid #1A2535;
  display:flex;align-items:center;
}}
.proj-row:hover .proj-label{{color:#E2E8F0;}}

/* LABEL */
.proj-label{{
  flex-shrink:0;
  display:flex;flex-direction:column;justify-content:center;
  padding:0 10px 0 4px;
  z-index:3;
}}
.proj-label .pname{{font-size:13px;font-weight:700;color:#CBD5E1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.proj-label .pidp{{font-size:12px;font-weight:600;margin-top:2px;}}

/* TIMELINE AREA */
.tl-area{{position:relative;flex:1;height:100%;}}

/* BOTÃO NOVO MARCO */
.add-marco-btn{{
  flex-shrink:0;width:28px;height:28px;
  display:flex;align-items:center;justify-content:center;
  background:rgba(59,130,246,.12);
  border:1px dashed rgba(59,130,246,.35);
  border-radius:7px;cursor:pointer;
  font-size:15px;color:#3B82F6;
  margin:auto 8px auto 6px;
  transition:background .15s,border-color .15s,transform .15s;
  user-select:none;flex-shrink:0;
}}
.add-marco-btn:hover{{
  background:rgba(59,130,246,.25);
  border-color:#3B82F6;
  transform:scale(1.12);
}}

/* BAR */
.bar{{
  position:absolute;border-radius:5px;cursor:pointer;overflow:hidden;
  transition:filter .15s;display:flex;align-items:center;
}}
.bar:hover{{filter:brightness(1.3);z-index:10;}}
.bar-progress{{
  position:absolute;left:0;top:0;bottom:0;
  background:rgba(255,255,255,.2);border-radius:5px 0 0 5px;pointer-events:none;
}}
.bar-txt{{
  position:relative;z-index:1;
  font-size:11px;font-weight:700;color:rgba(255,255,255,.95);
  white-space:nowrap;padding:0 8px;letter-spacing:.02em;
}}

/* MARCO WRAPPER */
.marco-wrap{{
  position:absolute;
  display:flex;flex-direction:column;align-items:center;
  cursor:pointer;z-index:15;
  transform:translateX(-50%);
  pointer-events:all;
}}
/* MARCO DIAMANTE */
.marco{{
  width:13px;height:13px;
  transform:rotate(45deg);
  transition:transform .15s,filter .15s;
  border-radius:2px;flex-shrink:0;
}}
.marco-wrap:hover .marco{{transform:rotate(45deg) scale(1.5);filter:brightness(1.5);}}
.marco.done{{background:#22C55E;box-shadow:0 0 8px #22C55E88;}}
.marco.ok  {{background:#3B82F6;box-shadow:0 0 6px #3B82F666;}}
.marco.late{{
  background:#EF4444;box-shadow:0 0 12px #EF444499;
  animation:pulse 1.4s ease-in-out infinite;
}}
@keyframes pulse{{0%,100%{{box-shadow:0 0 8px #EF444499;}}50%{{box-shadow:0 0 20px #EF4444CC;}}}}
/* MARCO GO-LIVE — estrela amarela */
.marco-golive{{
  width:18px;height:18px;
  background:none;box-shadow:none;transform:none;
  border-radius:0;
  display:flex;align-items:center;justify-content:center;
  font-size:15px;line-height:1;cursor:pointer;
  transition:transform .15s,filter .15s;
  margin-top:-2px;
}}
.marco-wrap:hover .marco-golive{{transform:scale(1.4);filter:brightness(1.3);}}
/* CONECTOR vertical diamante → label */
.marco-stem{{
  width:1px;height:14px;
  background:rgba(148,163,184,.35);
  flex-shrink:0;
}}
/* MARCO LABEL — horizontal, bem abaixo da barra */
.marco-lbl{{
  margin-top:0;
  font-size:10px;font-weight:600;color:#FFFFFF;
  white-space:nowrap;
  writing-mode:horizontal-tb;
  transform:none;
  max-width:88px;
  overflow:hidden;text-overflow:ellipsis;
  line-height:1.3;letter-spacing:.01em;
  text-align:center;
  padding:2px 4px;
  border-radius:3px;
  background:rgba(15,22,35,.7);
  border:1px solid rgba(100,116,139,.18);
}}
.marco-wrap:hover .marco-lbl{{
  color:#E2E8F0;
  border-color:rgba(148,163,184,.5);
  background:rgba(30,41,59,.95);
}}

/* HOVER TOOLTIP DO MARCO — position fixed, segue o mouse */
#marco-tip{{
  display:none;
  position:fixed;
  background:#1E293B;
  border:1px solid #334155;
  border-radius:10px;
  padding:14px 18px;
  min-width:210px;
  box-shadow:0 16px 48px rgba(0,0,0,.8),0 0 0 1px rgba(99,179,237,.08);
  font-size:13px;line-height:1.8;color:#E2E8F0;
  pointer-events:none;
  z-index:9999;
  white-space:normal;
  transition:opacity .1s;
}}
#marco-tip.show{{display:block;}}
.mt-title{{font-size:14px;font-weight:700;color:#E2E8F0;margin-bottom:8px;
           border-bottom:1px solid #2D3F55;padding-bottom:6px;
           display:flex;align-items:center;gap:6px;}}
.mt-diamond{{width:9px;height:9px;transform:rotate(45deg);
             border-radius:1px;flex-shrink:0;display:inline-block;}}
.mt-row{{display:flex;justify-content:space-between;gap:20px;margin-bottom:3px;}}
.mt-label{{color:#475569;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;}}
.mt-val{{color:#CBD5E1;font-weight:600;text-align:right;}}
.mt-badge{{display:inline-block;padding:2px 10px;border-radius:10px;
           font-size:12px;font-weight:700;margin-top:8px;}}

/* MODAL */
#modal-overlay{{
  display:none;position:fixed;inset:0;z-index:1000;
  background:rgba(0,0,0,.65);backdrop-filter:blur(3px);
  align-items:center;justify-content:center;
}}
#modal-overlay.open{{display:flex;}}
#modal{{
  background:#1E293B;border:1px solid #334155;
  border-radius:14px;padding:24px 28px;
  width:min(520px,92vw);max-height:80vh;overflow-y:auto;
  box-shadow:0 24px 64px rgba(0,0,0,.7);
  animation:fadeIn .18s ease;
}}
@keyframes fadeIn{{from{{opacity:0;transform:translateY(-12px);}}to{{opacity:1;transform:translateY(0);}}}}
#modal-close{{
  float:right;cursor:pointer;font-size:18px;color:#64748B;
  background:none;border:none;padding:0;line-height:1;
}}
#modal-close:hover{{color:#E2E8F0;}}
#modal h2{{font-size:15px;font-weight:700;color:#E2E8F0;margin-bottom:4px;}}
#modal .modal-sub{{font-size:11px;color:#64748B;margin-bottom:14px;}}
.modal-sep{{border:none;border-top:1px solid #2D3F55;margin:12px 0;}}
.modal-kpi-row{{display:flex;gap:16px;margin-bottom:12px;flex-wrap:wrap;}}
.modal-kpi{{
  background:#0F1E2E;border-radius:8px;padding:10px 14px;flex:1;min-width:110px;
}}
.modal-kpi .mk-label{{font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.07em;}}
.modal-kpi .mk-val  {{font-size:18px;font-weight:700;color:#E2E8F0;margin-top:2px;}}
.modal-kpi .mk-sub  {{font-size:10px;color:#64748B;}}
.modal-section-title{{font-size:10px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px;}}
.marco-item{{
  display:flex;align-items:center;gap:10px;
  padding:7px 10px;border-radius:6px;margin-bottom:4px;
  background:#0F1E2E;
}}
.marco-dot{{width:10px;height:10px;border-radius:2px;transform:rotate(45deg);flex-shrink:0;}}
.marco-info{{flex:1;}}
.marco-info .mi-nome{{font-size:11px;font-weight:600;color:#CBD5E1;}}
.marco-info .mi-meta{{font-size:10px;color:#64748B;margin-top:1px;}}
.subfase-item{{
  display:flex;align-items:center;justify-content:space-between;
  padding:5px 8px;border-radius:5px;margin-bottom:3px;
  background:#0F1E2E;font-size:10px;
}}
.sf-nome{{color:#94A3B8;flex:1;}}
.sf-pct{{font-weight:700;color:#E2E8F0;margin:0 8px;}}
.sf-idp{{font-weight:700;}}

/* LEGENDA */
#legend{{display:flex;gap:18px;flex-wrap:wrap;padding:10px 0 0 0;}}
.leg{{display:flex;align-items:center;gap:5px;font-size:11px;color:#475569;}}
.leg-d{{width:9px;height:9px;transform:rotate(45deg);border-radius:1px;flex-shrink:0;}}
.leg-b{{width:20px;height:7px;border-radius:2px;flex-shrink:0;}}

/* CONTEXT MENU */
#ctx-menu{{
  display:none;position:fixed;z-index:99999;
  background:#1E293B;border:1px solid #334155;
  border-radius:10px;padding:6px;
  box-shadow:0 12px 40px rgba(0,0,0,.8);
  min-width:160px;
  animation:fadeIn .12s ease;
}}
#ctx-menu.show{{display:block;}}
.ctx-item{{
  display:flex;align-items:center;gap:10px;
  padding:9px 14px;border-radius:7px;
  font-size:12px;font-weight:600;color:#CBD5E1;
  cursor:pointer;transition:background .1s;
  user-select:none;
}}
.ctx-item:hover{{background:#0F1E2E;}}
.ctx-item.danger{{color:#F87171;}}
.ctx-item.danger:hover{{background:rgba(239,68,68,.1);}}
.ctx-sep{{height:1px;background:#2D3F55;margin:4px 0;}}

/* EDIT MODAL */
#edit-overlay{{
  display:none;position:fixed;inset:0;z-index:10000;
  background:rgba(0,0,0,.7);backdrop-filter:blur(4px);
  align-items:center;justify-content:center;
}}
#edit-overlay.open{{display:flex;}}
#edit-modal{{
  background:#1E293B;border:1px solid #334155;
  border-radius:14px;padding:28px 32px;
  width:min(440px,92vw);
  box-shadow:0 24px 64px rgba(0,0,0,.8);
  animation:fadeIn .18s ease;
}}
#edit-modal h3{{
  font-size:15px;font-weight:700;color:#E2E8F0;
  margin-bottom:20px;display:flex;align-items:center;gap:8px;
}}
.edit-field{{margin-bottom:16px;}}
.edit-label{{
  font-size:10px;font-weight:700;color:#475569;
  text-transform:uppercase;letter-spacing:.07em;
  margin-bottom:6px;display:block;
}}
.edit-input{{
  width:100%;background:#0F1E2E;
  border:1px solid #334155;border-radius:8px;
  padding:9px 12px;font-size:13px;color:#E2E8F0;
  outline:none;transition:border .15s;
  font-family:inherit;
}}
.edit-input:focus{{border-color:#3B82F6;box-shadow:0 0 0 3px rgba(59,130,246,.15);}}
.edit-actions{{display:flex;gap:10px;margin-top:20px;justify-content:flex-end;}}
.btn-save{{
  background:#3B82F6;color:#fff;border:none;
  padding:9px 22px;border-radius:8px;
  font-size:12px;font-weight:700;cursor:pointer;
  transition:background .15s;
}}
.btn-save:hover{{background:#2563EB;}}
.btn-cancel{{
  background:transparent;color:#64748B;
  border:1px solid #334155;
  padding:9px 18px;border-radius:8px;
  font-size:12px;font-weight:600;cursor:pointer;
  transition:all .15s;
}}
.btn-cancel:hover{{color:#CBD5E1;border-color:#475569;}}
</style>
</head>
<body>
<div id="wrap">
  <div id="header-row"></div>
  <div id="body"></div>
  <div id="legend">
    <div class="leg"><div class="leg-d" style="background:#22C55E"></div>Marco Concluído</div>
    <div class="leg"><div class="leg-d" style="background:#3B82F6"></div>Marco no Prazo</div>
    <div class="leg"><div class="leg-d" style="background:#EF4444"></div>Marco Atrasado</div>
    <div class="leg"><span style="font-size:13px;line-height:1;">⭐</span>&nbsp;Go-Live</div>
    <div class="leg" style="color:#63B3ED;border-left:2px solid #63B3ED;padding-left:5px;margin-left:12px;">Linha Hoje</div>
  </div>
</div>

<div id="marco-tip"></div>

<!-- CONTEXT MENU -->
<div id="ctx-menu">
  <div class="ctx-item" id="ctx-edit">✏️ Editar Marco</div>
  <div class="ctx-sep"></div>
  <div class="ctx-item danger" id="ctx-del">🗑️ Excluir Marco</div>
</div>

<!-- EDIT MODAL -->
<div id="edit-overlay">
  <div id="edit-modal">
    <h3><span id="edit-diamond" style="width:12px;height:12px;transform:rotate(45deg);border-radius:2px;display:inline-block;background:#3B82F6;flex-shrink:0;"></span> Editar Marco</h3>
    <div class="edit-field">
      <label class="edit-label">Nome do Marco</label>
      <input class="edit-input" id="edit-nome" type="text" placeholder="Nome do marco...">
    </div>
    <div style="display:flex;gap:14px;">
      <div class="edit-field" style="flex:1;">
        <label class="edit-label">Data de Término</label>
        <input class="edit-input" id="edit-data" type="date">
      </div>
      <div class="edit-field" style="flex:1;">
        <label class="edit-label">% Concluído</label>
        <input class="edit-input" id="edit-pct" type="number" min="0" max="100" step="1" placeholder="0">
      </div>
    </div>
    <div class="edit-actions">
      <button class="btn-cancel" onclick="closeEditModal()">Cancelar</button>
      <button class="btn-save" onclick="saveEditModal()">Salvar</button>
    </div>
  </div>
</div>

<!-- MODAL -->
<div id="modal-overlay">
  <div id="modal">
    <button id="modal-close" onclick="closeModal()">✕</button>
    <div id="modal-content"></div>
  </div>
</div>

<script>
const PROJECTS       = {proj_json};
const marcoTip       = document.getElementById('marco-tip');
const HOJE_MS        = {hoje_ts};
const MODO_APRESENTAR = {_modo_apres_js};
const ROW_H    = {row_h};
const LABEL_W  = 200;

// ── Min/Max ──────────────────────────────────────────────────────────────────
let minMs = Infinity, maxMs = -Infinity;
PROJECTS.forEach(p => {{
  if(p.inicio  != null) {{ minMs=Math.min(minMs,p.inicio);  maxMs=Math.max(maxMs,p.inicio);  }}
  if(p.termino != null) {{ minMs=Math.min(minMs,p.termino); maxMs=Math.max(maxMs,p.termino); }}
  p.marcos.forEach(m => {{ if(m.termino!=null) {{ minMs=Math.min(minMs,m.termino); maxMs=Math.max(maxMs,m.termino); }} }});
}});
const SPAN = maxMs - minMs;
const MIN  = minMs - SPAN*0.01;
const MAX  = maxMs + SPAN*0.03;

function d2p(ms) {{ return ((ms-MIN)/(MAX-MIN))*100; }}

// ── Meses ────────────────────────────────────────────────────────────────────
const MNAMES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
function genMonths() {{
  const months=[]; const d=new Date(MIN); d.setUTCDate(1);
  while(d.getTime()<=MAX+86400000*31) {{
    const s=d.getTime(); const nd=new Date(d); nd.setUTCMonth(nd.getUTCMonth()+1);
    months.push({{start:s,end:nd.getTime(),d:new Date(s)}});
    d.setUTCMonth(d.getUTCMonth()+1); if(d.getUTCFullYear()>2030)break;
  }}
  return months;
}}
const MONTHS=genMonths();

// ── Header ───────────────────────────────────────────────────────────────────
const hdr=document.getElementById('header-row');
MONTHS.forEach(m=>{{
  const x0=d2p(m.start), x1=d2p(m.end);
  if(x1<0||x0>100) return;
  const l=Math.max(0,x0), r=Math.min(100,x1), w=r-l;
  if(w<0.5) return;
  const el=document.createElement('div'); el.className='month-cell';
  el.style.left=`calc(${{LABEL_W}}px + (${{l/100}} * (100% - ${{LABEL_W}}px)))`;
  el.style.width=`calc(${{w/100}} * (100% - ${{LABEL_W}}px))`;
  el.textContent=MNAMES[m.d.getUTCMonth()]+'/'+String(m.d.getUTCFullYear()).slice(2);
  hdr.appendChild(el);
}});

// ── Body ─────────────────────────────────────────────────────────────────────
const body=document.getElementById('body');
body.style.height=(ROW_H*PROJECTS.length)+'px';
body.style.position='relative';

MONTHS.forEach(m=>{{
  const xp=d2p(m.start); if(xp<0||xp>100) return;
  const gl=document.createElement('div'); gl.className='gridline';
  gl.style.left=`calc(${{LABEL_W}}px + (${{xp/100}} * (100% - ${{LABEL_W}}px)))`;
  body.appendChild(gl);
}});

// Linha Hoje
if(HOJE_MS>=MIN&&HOJE_MS<=MAX){{
  const xp=d2p(HOJE_MS);
  const tl=document.createElement('div'); tl.className='today-line';
  const left=`calc(${{LABEL_W}}px + (${{xp/100}} * (100% - ${{LABEL_W}}px)))`;
  tl.style.left=left;
  const lbl=document.createElement('div'); lbl.className='today-lbl';
  lbl.style.left=left;
  const hd=new Date(HOJE_MS);
  lbl.textContent='Hoje '+hd.getUTCDate()+'\/'+(hd.getUTCMonth()+1);
  body.appendChild(tl); body.appendChild(lbl);
}}

// ── Helpers ──────────────────────────────────────────────────────────────────
function isGolive(nome){{
  // Detecta qualquer variação contendo as letras g-o-l-i-v-e em sequência
  var n=(nome||'').toLowerCase().replace(/[^a-z]/g,'');
  return n.includes('golive');
}}

function fmtDate(ms){{
  if(ms==null) return '—';
  const d=new Date(ms);
  return d.getUTCDate().toString().padStart(2,'0')+'/'+(d.getUTCMonth()+1).toString().padStart(2,'0')+'/'+d.getUTCFullYear();
}}
function idpColor(v){{
  if(v==null) return '#64748B';
  if(v>=0.99) return '#22C55E';
  if(v>=0.95) return '#D97706';
  return '#EF4444';
}}
function idpFace(v){{
  if(v==null) return '';
  if(v>=0.99) return '';
  if(v>=0.95) return '';
  return '';
}}

// ── Rows ─────────────────────────────────────────────────────────────────────
PROJECTS.forEach((p,i)=>{{
  const row=document.createElement('div');
  row.className='proj-row';
  row.style.height=ROW_H+'px';
  row.style.top=(i*ROW_H)+'px';
  row.style.position='absolute';
  row.style.width='100%';
  row.style.background=i%2===0?'rgba(255,255,255,.018)':'transparent';

  // Label
  const lbl=document.createElement('div');
  lbl.className='proj-label';
  lbl.style.width=LABEL_W+'px';
  const idpC=idpColor(p.idp);
  const idpTxt=p.idp!=null?p.idp.toFixed(2):'N/A';
  lbl.innerHTML=`
    <div class="pname" title="${{p.projeto}}">${{p.projeto.length>22?p.projeto.slice(0,22)+'…':p.projeto}}</div>
    <div class="pidp" style="color:${{idpC}}">IDP ${{idpTxt}}</div>`;
  row.appendChild(lbl);

  // Timeline area
  const tla=document.createElement('div');
  tla.className='tl-area';

  // BAR simples: fundo branco/cinza claro, preenchimento cinza por % conclusão
  if(p.inicio!=null&&p.termino!=null){{
    let barEndMs = p.termino;
    if(p.marcos&&p.marcos.length>0)
      p.marcos.forEach(m=>{{ if(m.termino!=null&&m.termino>barEndMs) barEndMs=m.termino; }});

    const x0all=d2p(p.inicio), x1all=d2p(barEndMs);
    const lpAll=Math.max(0,x0all), rpAll=Math.min(100,x1all);

    const barWrap=document.createElement('div');
    barWrap.style.cssText=`position:absolute;top:18px;height:28px;border-radius:6px;
      overflow:hidden;cursor:pointer;
      background:#F1F5F9;border:1px solid #CBD5E1;
      left:${{lpAll}}%;width:${{rpAll-lpAll}}%;`;
    barWrap.addEventListener('click',()=>openModal(p));
    barWrap.title='Clique para detalhes';

    // Preenchimento cinza proporcional ao % concluído
    const fill=document.createElement('div');
    fill.style.cssText=`position:absolute;top:0;left:0;bottom:0;
      width:${{p.pct}}%;
      background:#94A3B8;
      border-radius:5px 0 0 5px;`;
    barWrap.appendChild(fill);

    // Texto % no centro
    const bt=document.createElement('div');
    bt.style.cssText=`position:absolute;top:0;left:0;right:0;bottom:0;
      display:flex;align-items:center;padding:0 10px;
      font-size:12px;font-weight:700;color:#475569;font-family:'Inter','Segoe UI',sans-serif;`;
    bt.textContent=p.pct.toFixed(0)+'% concluído';
    barWrap.appendChild(bt);

    p._barEl = barWrap;
    p._barLp = lpAll;
    tla.appendChild(barWrap);
  }}


  // MARCOS sobre a barra com label abaixo
  p.marcos.forEach(m=>{{
    if(m.termino==null) return;
    const xp=d2p(m.termino);
    if(xp<0||xp>100) return;

    const wrap=document.createElement('div');
    wrap.className='marco-wrap';
    wrap.style.left=`${{xp}}%`;
    // Diamante centrado na barra (barra: top=16px h=32px → centro=32px)
    wrap.style.top='26px';

    // Diamante
    const mk=document.createElement('div');
    if(isGolive(m.nome)){{
      mk.className='marco-golive';
      mk.textContent='⭐';
    }} else {{
      mk.className='marco '+(m.concluido?'done':m.atrasado?'late':'ok');
    }}
    wrap.appendChild(mk);

    // Conector vertical
    const stem=document.createElement('div');
    stem.className='marco-stem';
    wrap.appendChild(stem);

    // Label horizontal abaixo
    const lbl=document.createElement('div');
    lbl.className='marco-lbl';
    lbl.textContent=m.nome;
    wrap.appendChild(lbl);

    // Tooltip rico — position:fixed, segue mouse, nunca fica atrás
    const mc2=m.concluido?'#22C55E':m.atrasado?'#EF4444':'#3B82F6';
    const ms2=m.concluido?'✅ Concluído':m.atrasado?'🔴 Atrasado':'🔵 No prazo';
    const badgeBg=m.concluido?'rgba(34,197,94,.18)':m.atrasado?'rgba(239,68,68,.18)':'rgba(59,130,246,.18)';
    const tipHTML=`
      <div class="mt-title">
        <span class="mt-diamond" style="background:${{mc2}}"></span>
        ${{m.nome}}
      </div>
      <div class="mt-row">
        <span class="mt-label">Data</span>
        <span class="mt-val">${{fmtDate(m.termino)}}</span>
      </div>
      <div class="mt-row">
        <span class="mt-label">% Concluído</span>
        <span class="mt-val">${{m.pct.toFixed(0)}}%</span>
      </div>
      ${{m.baseline?`<div class="mt-row">
        <span class="mt-label">Baseline</span>
        <span class="mt-val">${{fmtDate(m.baseline)}}</span>
      </div>`:''}}
      <div style="margin-top:8px;">
        <span class="mt-badge" style="background:${{badgeBg}};color:${{mc2}}">${{ms2}}</span>
      </div>`;

    wrap.addEventListener('mouseenter', ()=>{{
      marcoTip.innerHTML=tipHTML;
      marcoTip.classList.add('show');
    }});
    wrap.addEventListener('mousemove', e=>{{
      // Posição relativa ao viewport, offset para não cobrir o cursor
      let tx=e.clientX+14, ty=e.clientY-10;
      // Evitar saída pela direita
      if(tx+220>window.innerWidth) tx=e.clientX-230;
      // Evitar saída pelo topo
      if(ty<8) ty=e.clientY+20;
      marcoTip.style.left=tx+'px';
      marcoTip.style.top=ty+'px';
    }});
    wrap.addEventListener('mouseleave', ()=>{{
      marcoTip.classList.remove('show');
    }});

    // Left click → modal detalhes
    wrap.addEventListener('click', e=>{{ e.stopPropagation(); openModalMarco(m, p); }});

    // Right click → context menu
    wrap.addEventListener('contextmenu', e=>{{
      e.preventDefault(); e.stopPropagation();
      marcoTip.classList.remove('show');
      ctxTargetMarco = m;
      ctxTargetProj  = p;
      ctxTargetWrap  = wrap;
      const cx=e.clientX, cy=e.clientY;
      ctxMenu.style.left=(cx+2)+'px';
      ctxMenu.style.top=(cy+2)+'px';
      ctxMenu.classList.add('show');
    }});

    tla.appendChild(wrap);
  }});

  row.appendChild(tla);

  // ── Botão ➕ novo marco ────────────────────────────────────────────────────
  const addBtn=document.createElement('div');
  addBtn.className='add-marco-btn';
  addBtn.textContent='➕';
  addBtn.title='Adicionar novo marco para '+p.projeto;
  if(MODO_APRESENTAR) addBtn.style.display='none';
  addBtn.addEventListener('click', e=>{{
    e.stopPropagation();
    marcoTip.classList.remove('show');
    // Abre modal de edição em modo "Novo Marco"
    ctxTargetMarco = null;   // null = modo criação
    ctxTargetProj  = p;
    ctxTargetWrap  = null;
    document.getElementById('edit-diamond').style.background='#3B82F6';
    document.getElementById('edit-nome').value='';
    // Sugere data de término do projeto como padrão
    const dt=p.termino?new Date(p.termino):new Date();
    document.getElementById('edit-data').value=
      dt.getUTCFullYear()+'-'+
      String(dt.getUTCMonth()+1).padStart(2,'0')+'-'+
      String(dt.getUTCDate()).padStart(2,'0');
    document.getElementById('edit-pct').value='0';
    document.getElementById('edit-modal').querySelector('h3').innerHTML=
      '<span id="edit-diamond" style="width:12px;height:12px;transform:rotate(45deg);border-radius:2px;display:inline-block;background:#3B82F6;flex-shrink:0;"></span> Novo Marco — '+p.projeto;
    editOverlay.classList.add('open');
    // Store tla so we can append the new wrap after save
    editOverlay._tla = tla;
  }});
  row.appendChild(addBtn);

  body.appendChild(row);
}});

// ── Modal ────────────────────────────────────────────────────────────────────
const overlay=document.getElementById('modal-overlay');
const modalContent=document.getElementById('modal-content');

function closeModal(){{ overlay.classList.remove('open'); }}
overlay.addEventListener('click', e=>{{ if(e.target===overlay) closeModal(); }});
document.addEventListener('keydown', e=>{{ if(e.key==='Escape') closeModal(); }});

function openModal(p){{
  const idpC=idpColor(p.idp);
  const idpTxt=p.idp!=null?p.idp.toFixed(2):'N/A';
  const idpFc=idpFace(p.idp);

  let subfHTML='';
  if(p.subfases&&p.subfases.length>0){{
    subfHTML='<div class="modal-section-title">Subfases</div>';
    p.subfases.forEach(s=>{{
      const sc=idpColor(s.idp);
      const si=s.idp!=null?s.idp.toFixed(2):'—';
      subfHTML+=`<div class="subfase-item">
        <span class="sf-nome">${{s.nome}}</span>
        <span class="sf-pct">${{s.pct.toFixed(0)}}%</span>
        <span class="sf-idp" style="color:${{sc}}">IDP ${{si}}</span>
        <span style="font-size:9px;color:#475569;margin-left:6px">${{fmtDate(s.termino)}}</span>
      </div>`;
    }});
  }}

  let mkHTML='';
  if(p.marcos&&p.marcos.length>0){{
    mkHTML='<hr class="modal-sep"><div class="modal-section-title">Marcos</div>';
    p.marcos.forEach(m=>{{
      const mc=m.concluido?'#22C55E':m.atrasado?'#EF4444':'#3B82F6';
      const ms2=m.concluido?'✅ Concluído':m.atrasado?'🔴 Atrasado':'🔵 No prazo';
      mkHTML+=`<div class="marco-item">
        <div class="marco-dot" style="background:${{mc}}"></div>
        <div class="marco-info">
          <div class="mi-nome">${{m.nome}}</div>
          <div class="mi-meta">${{ms2}} · ${{fmtDate(m.termino)}}${{m.baseline?` · Baseline: ${{fmtDate(m.baseline)}}`:''}}</div>
        </div>
      </div>`;
    }});
  }}

  modalContent.innerHTML=`
    <h2>${{p.projeto}}</h2>
    <div class="modal-sub">${{fmtDate(p.inicio)}} → ${{fmtDate(p.termino)}}</div>
    <div class="modal-kpi-row">
      <div class="modal-kpi">
        <div class="mk-label">Conclusão</div>
        <div class="mk-val">${{p.pct.toFixed(0)}}%</div>
      </div>
      <div class="modal-kpi">
        <div class="mk-label">IDP</div>
        <div class="mk-val" style="color:${{idpC}}">${{idpFc}} ${{idpTxt}}</div>
        <div class="mk-sub" style="color:${{idpC}}">${{p.idp>=0.99?'Em dia':p.idp>=0.95?'Em alerta':p.idp?'Em atraso':'N/A'}}</div>
      </div>
      <div class="modal-kpi">
        <div class="mk-label">Baseline</div>
        <div class="mk-val" style="font-size:13px">${{fmtDate(p.baseline)}}</div>
      </div>
    </div>
    ${{subfHTML}}
    ${{mkHTML}}`;
  overlay.classList.add('open');
}}

// ── Context menu & Edit logic ────────────────────────────────────────────────
const ctxMenu      = document.getElementById('ctx-menu');
const editOverlay  = document.getElementById('edit-overlay');
let ctxTargetMarco = null;
let ctxTargetProj  = null;
let ctxTargetWrap  = null;

// Hide context menu on any click elsewhere
document.addEventListener('click', ()=>ctxMenu.classList.remove('show'));
document.addEventListener('keydown', e=>{{
  if(e.key==='Escape') {{ ctxMenu.classList.remove('show'); closeEditModal(); }}
}});

// ── Editar ───────────────────────────────────────────────────────────────────
document.getElementById('ctx-edit').addEventListener('click', ()=>{{
  if(!ctxTargetMarco) return;
  ctxMenu.classList.remove('show');
  const mc2=ctxTargetMarco.concluido?'#22C55E':ctxTargetMarco.atrasado?'#EF4444':'#3B82F6';
  document.getElementById('edit-diamond').style.background=mc2;
  document.getElementById('edit-nome').value=ctxTargetMarco.nome;
  // Format timestamp to yyyy-mm-dd
  const dt=ctxTargetMarco.termino?new Date(ctxTargetMarco.termino):new Date();
  document.getElementById('edit-data').value=
    dt.getUTCFullYear()+'-'+
    String(dt.getUTCMonth()+1).padStart(2,'0')+'-'+
    String(dt.getUTCDate()).padStart(2,'0');
  document.getElementById('edit-pct').value=Math.round(ctxTargetMarco.pct);
  editOverlay.classList.add('open');
}});

function closeEditModal(){{ editOverlay.classList.remove('open'); }}
editOverlay.addEventListener('click', e=>{{ if(e.target===editOverlay) closeEditModal(); }});

function saveEditModal(){{
  const novoNome = document.getElementById('edit-nome').value.trim();
  const novaData = document.getElementById('edit-data').value;
  const novoPct  = parseInt(document.getElementById('edit-pct').value)||0;
  if(!novoNome) {{ closeEditModal(); return; }}

  let termMs = null;
  if(novaData){{
    const [y,mo,d]=novaData.split('-').map(Number);
    termMs=Date.UTC(y,mo-1,d);
  }}
  const concluido=novoPct>=100;
  const atrasado=(!concluido)&&termMs&&(termMs<Date.now());

  if(ctxTargetMarco && ctxTargetWrap){{
    // ── MODO EDIÇÃO ──────────────────────────────────────────────────────────
    ctxTargetMarco.nome=novoNome;
    ctxTargetMarco.pct=novoPct;
    if(termMs) ctxTargetMarco.termino=termMs;
    ctxTargetMarco.concluido=concluido;
    ctxTargetMarco.atrasado=!!atrasado;

    const newCls='marco '+(concluido?'done':atrasado?'late':'ok');
    const diamond=ctxTargetWrap.querySelector('.marco');
    if(diamond) diamond.className=newCls;
    const lbl=ctxTargetWrap.querySelector('.marco-lbl');
    if(lbl) lbl.textContent=novoNome;
    const xp=d2p(ctxTargetMarco.termino);
    if(xp>=0&&xp<=100) ctxTargetWrap.style.left=xp+'%';

  }} else if(ctxTargetProj){{
    // ── MODO CRIAÇÃO ─────────────────────────────────────────────────────────
    const novoMarco={{
      nome:novoNome, termino:termMs, baseline:null,
      pct:novoPct, status:'', resp:'', nivel:2,
      concluido:concluido, atrasado:!!atrasado
    }};
    ctxTargetProj.marcos.push(novoMarco);

    // Atualiza a barra do projeto se o novo marco ultrapassar o término atual
    if(termMs!=null && ctxTargetProj._barEl){{
      let newBarEnd = ctxTargetProj.termino;
      ctxTargetProj.marcos.forEach(m=>{{
        if(m.termino!=null && m.termino > newBarEnd) newBarEnd = m.termino;
      }});
      const newRp = Math.min(100, d2p(newBarEnd));
      const newWp = newRp - ctxTargetProj._barLp;
      if(newWp > 0.05) ctxTargetProj._barEl.style.width = newWp + '%';
    }}

    const tla=editOverlay._tla;
    if(tla&&termMs!=null){{
      const xp=d2p(termMs);
      if(xp>=0&&xp<=100){{
        const mc2=concluido?'#22C55E':atrasado?'#EF4444':'#3B82F6';
        const ms2=concluido?'✅ Concluído':atrasado?'🔴 Atrasado':'🔵 No prazo';
        const badgeBg=concluido?'rgba(34,197,94,.18)':atrasado?'rgba(239,68,68,.18)':'rgba(59,130,246,.18)';

        const wrap=document.createElement('div');
        wrap.className='marco-wrap';
        wrap.style.left=xp+'%'; wrap.style.top='26px';

        const mk=document.createElement('div');
        if(isGolive(novoNome)){{
          mk.className='marco-golive';
          mk.textContent='⭐';
        }} else {{
          mk.className='marco '+(concluido?'done':atrasado?'late':'ok');
        }}
        wrap.appendChild(mk);

        const stem=document.createElement('div'); stem.className='marco-stem';
        wrap.appendChild(stem);

        const lblEl=document.createElement('div'); lblEl.className='marco-lbl';
        lblEl.textContent=novoNome; wrap.appendChild(lblEl);

        const tipHTML=`<div class="mt-title"><span class="mt-diamond" style="background:${{mc2}}"></span>${{novoNome}}</div>
          <div class="mt-row"><span class="mt-label">Data</span><span class="mt-val">${{fmtDate(termMs)}}</span></div>
          <div class="mt-row"><span class="mt-label">% Concluído</span><span class="mt-val">${{novoPct}}%</span></div>
          <div style="margin-top:8px;"><span class="mt-badge" style="background:${{badgeBg}};color:${{mc2}}">${{ms2}}</span></div>`;

        wrap.addEventListener('mouseenter',()=>{{ marcoTip.innerHTML=tipHTML; marcoTip.classList.add('show'); }});
        wrap.addEventListener('mousemove',e=>{{
          let tx=e.clientX+14,ty=e.clientY-10;
          if(tx+220>window.innerWidth) tx=e.clientX-230;
          if(ty<8) ty=e.clientY+20;
          marcoTip.style.left=tx+'px'; marcoTip.style.top=ty+'px';
        }});
        wrap.addEventListener('mouseleave',()=>marcoTip.classList.remove('show'));
        wrap.addEventListener('click',e=>{{ e.stopPropagation(); openModalMarco(novoMarco, ctxTargetProj); }});
        wrap.addEventListener('contextmenu',e=>{{
          e.preventDefault(); e.stopPropagation();
          marcoTip.classList.remove('show');
          ctxTargetMarco=novoMarco; ctxTargetProj=ctxTargetProj;
          ctxTargetWrap=wrap;
          ctxMenu.style.left=(e.clientX+2)+'px'; ctxMenu.style.top=(e.clientY+2)+'px';
          ctxMenu.classList.add('show');
        }});
        tla.appendChild(wrap);
      }}
    }}
  }}
  closeEditModal();
}}

// ── Excluir ──────────────────────────────────────────────────────────────────
document.getElementById('ctx-del').addEventListener('click', ()=>{{
  if(!ctxTargetMarco||!ctxTargetProj||!ctxTargetWrap) return;
  ctxMenu.classList.remove('show');
  const idx=ctxTargetProj.marcos.indexOf(ctxTargetMarco);
  if(idx>-1) ctxTargetProj.marcos.splice(idx,1);
  ctxTargetWrap.remove();
  ctxTargetMarco=null; ctxTargetWrap=null;
}});

function openModalMarco(m, p){{
  const mc=m.concluido?'#22C55E':m.atrasado?'#EF4444':'#3B82F6';
  const ms2=m.concluido?'✅ Concluído':m.atrasado?'🔴 Atrasado':'🔵 No prazo';
  modalContent.innerHTML=`
    <h2>♦ ${{m.nome}}</h2>
    <div class="modal-sub">${{p.projeto}}</div>
    <div class="modal-kpi-row">
      <div class="modal-kpi">
        <div class="mk-label">Status</div>
        <div class="mk-val" style="font-size:14px;color:${{mc}}">${{ms2}}</div>
      </div>
      <div class="modal-kpi">
        <div class="mk-label">Data</div>
        <div class="mk-val" style="font-size:14px">${{fmtDate(m.termino)}}</div>
      </div>
      <div class="modal-kpi">
        <div class="mk-label">Baseline</div>
        <div class="mk-val" style="font-size:14px">${{fmtDate(m.baseline)}}</div>
      </div>
    </div>
    <hr class="modal-sep">
    <div class="modal-kpi-row">
      <div class="modal-kpi">
        <div class="mk-label">% Concluído</div>
        <div class="mk-val">${{m.pct.toFixed(0)}}%</div>
      </div>
      ${{m.resp?`<div class="modal-kpi"><div class="mk-label">Responsável</div><div class="mk-val" style="font-size:13px">${{m.resp}}</div></div>`:'' }}
      ${{m.status?`<div class="modal-kpi"><div class="mk-label">Status Texto</div><div class="mk-val" style="font-size:12px">${{m.status}}</div></div>`:'' }}
    </div>`;
  overlay.classList.add('open');
}}
</script>
</body></html>"""

    components.html(html_roadmap, height=total_h, scrolling=False)

# ──────────────────────────────────────────────────────────────────────────────
# ──────────────────────────────────────────────────────────────────────────────
# 11. SECTION 2 — GOVERNANÇA DE INCERTEZAS (linhas editáveis + adicionar)
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="section-title">C. Governança de Incertezas: Pontos Críticos e Planos de Ação</div>',
    unsafe_allow_html=True,
)

# ── Persistência: dict proj_key -> lista de linhas [{titulo, impacto, causa, plano}]
if 'gov_data' not in st.session_state:
    st.session_state.gov_data = {}
if 'gov_del' not in st.session_state:
    st.session_state.gov_del = {}
# Versão dos defaults — incrementar para forçar recarga do conteúdo PMO
_GOV_VERSION = "v3"
if st.session_state.get('gov_version') != _GOV_VERSION:
    st.session_state.gov_data = {}
    st.session_state.gov_version = _GOV_VERSION

# ── Análise PMO PMBOK 8ª Ed. — Pontos críticos pré-preenchidos por especialista ─
# Baseado na análise técnica dos XMLs em 21/05/2026.
# Campos editáveis — o usuário pode alterar a qualquer momento.
GOV_DEFAULTS = {
    "Business_Data_Fabric": [
        {
            "titulo": "Implantação MS Fabric — IDP 0,62 com DOWNStream em colapso (SPI=0,30)",
            "impacto": (
                "O projeto Business Data Fabric apresenta IDP consolidado de 0,62 — realiza apenas 62% do trabalho previsto para a data. "
                "O gargalo mais crítico está no DOWNStream da Implantação com SPI=0,30: apenas 29% do trabalho foi executado no período em que 88% deveria estar concluído. "
                "Sem recuperação imediata, o Go-Live de Dezembro/2026 está em risco direto, incluindo a migração dos dashboards Power BI para MS Fabric, o POV de Faturamento e o cutover do SAC. "
                "O impacto financeiro se estende ao ROI previsto para 2027 com a plataforma unificada de dados."
            ),
            "causa": (
                "1. Revisão da Minuta pelo Jurídico (SPI=0,36 e 0,40): dois contratos críticos — MS Fabric e Implantação — com minutas travadas no jurídico, bloqueando assinaturas e início formal das atividades downstream. "
                "2. VALIDAÇÃO SSA com SPI=0,37: ambiente SAC PRD/QAS sem equalização concluída, impedindo o cutover planejado para Junho. "
                "3. Recursos DataEX e HVAR com alocação abaixo do necessário para o escopo contratado. "
                "4. Treinamento Técnico de Data Analysts com SPI=0,27: capacitação de domínios praticamente não iniciada, gerando risco de adoção pós-implantação."
            ),
            "plano": (
                "1. [Imediato] PMO convoca reunião de crise com DataEX/HVAR até 27/Mai para nivelamento de capacidade — foco nos épicos 3 a 6 do DOWNStream. "
                "2. [Até 02/Jun] William/Supply consolidar sessão jurídica única para fechamento simultâneo das duas minutas (Fabric + Implantação). "
                "3. [Até 15/Jun] Rodrigo Evangelista entrega plano de sprint de equalização PRD/QAS com critérios de aceite definidos e aprovados pelo sponsor. "
                "4. [Até 30/Jun] Iniciar trilha de capacitação de Data Analysts em blocos de 2 semanas — sem essa entrega o CoE não opera. "
                "5. [Gate 30/Jun] Revisão executiva: se SPI não atingir 0,75, acionar plano de contingência de escopo reduzido para Go-Live."
            ),
        },
        {
            "titulo": "SAC Test Tenant — Go-Live 02/Jul em risco com SPI=0,62",
            "impacto": (
                "O SAC Test Tenant (Rodrigo Evangelista) tem Go-Live planejado para 02/Jul/2026 com SPI=0,62. "
                "A equalização dos ambientes PRD/QAS está incompleta, a VALIDAÇÃO SSA atingiu apenas 37% do planejado, e o processo formal de Gestão de Mudança (Change) não foi iniciado. "
                "Um Go-Live sem homologação adequada representa risco operacional direto no ambiente produtivo — instabilidade de dados SAC em produção pode paralisar operações críticas do negócio."
            ),
            "causa": (
                "1. Correção de ambientes SolvePlan/Rodrigo Evangelista com SPI=0,60 — atividades de equalização atrasadas em relação ao baseline. "
                "2. Transporte de objetos entre ambientes QAS→PRD ainda pendente, bloqueando a validação final. "
                "3. Processo de Change não iniciado: sem aprovação formal da mudança, o Go-Live não pode ser executado dentro do processo regulatório interno. "
                "4. Congelamento do ambiente PRD previsto para 24/Jun sem pré-requisitos concluídos."
            ),
            "plano": (
                "1. [25/Mai–20/Jun] Sprint dedicado de equalização PRD/QAS liderado por Rodrigo — entrega incremental semanal reportada ao PMO. "
                "2. [Até 28/Mai] Abrir processo de Change imediatamente — responsável: Gestor de Mudança designado. "
                "3. [Até 15/Jun] Definir critérios formais de Go/No-Go com sponsor e área de negócio. "
                "4. [24/Jun] Congelamento do ambiente PRD — data não negociável. Qualquer pendência aberta nessa data ativa o plano de contingência. "
                "5. [Contingência] Se Go-Live não puder ocorrer em 02/Jul, definir janela alternativa de Jul/2026 com mínimo impacto operacional."
            ),
        },
        {
            "titulo": "Governança de Dados — Letramento e capacitação com SPI=0,40",
            "impacto": (
                "O pilar de Governança do BDF apresenta SPI=0,83 no nível L1, mas o DOWNStream de Letramento e Capacitação tem SPI=0,40 — realizou apenas 39% do planejado. "
                "Data Owners, Data Custodians e Data Citizens dos Domínios Indústria e Agropecuária não estão certificados. "
                "Sem esses guardiões habilitados, o modelo de Governança de Dados 1.0 entregue pela Delaware não será operado corretamente, tornando o investimento ineficaz e o catálogo de dados inutilizável."
            ),
            "causa": (
                "1. Treinamento Técnico de Data Analysts com SPI=0,27 e Data Transforms praticamente não iniciado. "
                "2. Publicação da Política de Governança de Dados (Abr/2026) zerada — sem política publicada não há base normativa para o programa. "
                "3. Domínios Indústria (55%) e Agropecuária (52%) com catálogos de dados e monitoramento incompletos. "
                "4. Processo de Quality Gate definido mas sem aplicação nos domínios ativos."
            ),
            "plano": (
                "1. [Até 30/Mai] Delaware publica Política de Governança de Dados 1.0 — pré-requisito para todos os treinamentos subsequentes. "
                "2. [Jun–Jul] Trilha intensiva: Data Owners e Custodians (2 semanas) → Data Citizens (2 semanas) → Data Analysts (4 semanas). "
                "3. [Até 30/Jun] Concluir catálogos de dados dos Domínios Indústria e Agropecuária com Quality Gate aplicado. "
                "4. [Jul/2026] Cerimônia formal de ativação do CoE com presença do sponsor — marco de encerramento da fase de capacitação."
            ),
        },
    ],
    "Cockpit_Engenharia": [
        {
            "titulo": "Aprovação TAP por Marcel — Gargalo crítico no caminho do projeto (SPI=0,64)",
            "impacto": (
                "O Cockpit Engenharia opera com IDP=0,98 (fórmula MS Project), indicando desempenho geral satisfatório. "
                "Porém, a aprovação da TAP por Marcel com SPI=0,64 é o único ponto de falha que pode comprometer toda a FASE 2. "
                "Sem TAP formalmente aprovada: o contrato com IntechPRO não pode ser assinado, o Kick-off da FASE 2 (ajustes SolvePlan) não pode ser formalizado, e o Go-Live de Out/2026 fica em risco de postergação. "
                "O projeto está essencialmente com uma dependência crítica de uma única pessoa em uma única atividade."
            ),
            "causa": (
                "1. Marcel com disponibilidade insuficiente no período — múltiplas demandas concorrentes fora do projeto. "
                "2. Ciclo de aprovação mais longo que o planejado: o processo de aprovação da TAP (SPI=0,85 na fase de solicitação) não foi priorizado adequadamente. "
                "3. Ausência de escalada formal: o PMO não ativou o caminho de escalada executiva dentro do prazo estabelecido. "
                "4. Dependência sequencial rígida: aprovação TAP → assinatura contrato → Kick-off FASE 2, sem paralelismo possível."
            ),
            "plano": (
                "1. [Esta semana] Dumont agenda sessão de 30 minutos com Marcel — leva resumo executivo de 1 página com os riscos de atraso mapeados. "
                "2. [Até 30/Mai] Se Marcel não aprovar, PMO escala formalmente ao patrocinador executivo com registro documentado. "
                "3. [Imediato] Paralelizar o que for possível: iniciar análise dos dashs SolvePlan (FASE 2 UPStream) independente da TAP. "
                "4. [Até 05/Jun] Revisão do impacto no cronograma da FASE 2: mapear quais atividades podem iniciar sem TAP e quais ficam bloqueadas. "
                "5. [Contingência] Se TAP não aprovada até 10/Jun, revisar data do Go-Live com sponsor."
            ),
        },
        {
            "titulo": "Testes FASE 1 — Integrados não iniciados, risco de qualidade no Go-Live",
            "impacto": (
                "O DOWNStream FASE 1 está em 90% de conclusão, mas os Testes Integrados estão em 0% — ainda não iniciados. "
                "Os Testes Unitários atingiram 95%, mas com cobertura possivelmente abaixo do critério de aceite da EF aprovada por Amanda (Key-user). "
                "Ir ao Go-Live sem testes integrados completos representa risco real de defeitos em produção nas interfaces entre módulos do Cockpit, podendo gerar retrabalho pós-implantação e impacto na operação de engenharia."
            ),
            "causa": (
                "1. Recursos técnicos divididos entre finalização dos testes unitários e início das análises da FASE 2 (SolvePlan). "
                "2. Prazo apertado: Testes Unitários previsto para terminar 26/Mai, Integrados para 29/Mai — janela de apenas 3 dias. "
                "3. Critério de aceite dos testes não formalmente validado com Amanda antes da execução."
            ),
            "plano": (
                "1. [Até 29/Mai] Concluir 100% dos Testes Unitários — sem desvio de recursos para FASE 2. "
                "2. [27–29/Mai] Validação formal dos resultados dos testes com Amanda (Key-user) — assinatura do termo de aceite. "
                "3. [Até 02/Jun] Iniciar e concluir Testes Integrados — prioridade máxima na semana. "
                "4. [Até 05/Jun] Reunião de Go/No-Go para liberação do ambiente de homologação com PMO e Key-user. "
                "5. Apenas após aceite dos testes integrados iniciar a FASE 2."
            ),
        },
    ],
    "Esteira_Analytics": [
        {
            "titulo": "CEO Digital Boardroom — Etapa 3 com SPI=0,05, Go-Live Nov/2026 em risco severo",
            "impacto": (
                "A Etapa 3 do CEO Digital Boardroom — entrega dos 4 painéis executivos — apresenta SPI=0,05 com apenas 4% concluído de 4,86% planejado. "
                "Esse é o pior índice de todo o portfólio. Com Go-Live previsto para 30/Nov/2026, o projeto tem menos de 7 meses para entregar 4 painéis executivos complexos, incluindo indicadores críticos da C-Suite. "
                "A não entrega compromete diretamente o pilar de 'Visibilidade Executiva Digital' do programa de transformação de dados da empresa, gerando risco reputacional do PMO perante a diretoria."
            ),
            "causa": (
                "1. Escopo da Etapa 3 ainda indefinido: ausência de owner designado para os 4 painéis e KPIs da C-Suite não mapeados. "
                "2. Aprovação da proposta Bridge (consultoria) pendente — sem contrato, não há recursos técnicos alocados para a Etapa 3. "
                "3. Transição formal entre Etapa 2 (Estratégia de Dados, 100% concluída) e Etapa 3 não executada — gap de governança entre fases. "
                "4. Dados-fonte dos painéis executivos não estruturados: CEO Digital Boardroom depende de dados que ainda não estão disponíveis na camada analítica."
            ),
            "plano": (
                "1. [Até 28/Mai] Reunião de alinhamento com Bridge para definir MVP dos 4 painéis — escopo mínimo viável para Go-Live Nov/2026. "
                "2. [Até 30/Mai] Nomear Product Owner interno do CEO Digital Boardroom com dedicação mínima de 40%. "
                "3. [Até 05/Jun] Mapear os dados-fonte de cada painel e validar disponibilidade na camada analítica — sem dados não há painel. "
                "4. [Jun/2026] Definir Go-Live parcial como meta: 2 painéis até Ago/2026, 2 painéis restantes até Nov/2026. "
                "5. [Gate Jun/2026] Se aprovação Bridge não concluída, escalar ao CIO com análise de impacto formal."
            ),
        },
        {
            "titulo": "Entregas Q2 — SPI=0,54 com escopo 'A Definir' e ausência de owner",
            "impacto": (
                "As Entregas Q2 da Esteira Analytics (Abr–Jun/2026) têm SPI=0,54 com o único entregável listado como 'A Definir'. "
                "Com o encerramento do Q2 previsto para 30/Jun/2026, existe risco concreto de o quarter ser encerrado sem nenhuma entrega formal registrada. "
                "Isso representa perda de valor entregue no período e acúmulo de pressão sobre Q3 e Q4, que já têm seus próprios escopos indefinidos."
            ),
            "causa": (
                "1. UPStream Esteira com SPI=0,50: o processo de definição de entregas Q2 não foi concluído dentro do prazo planejado. "
                "2. Ausência de owner designado para as entregas Q2 — nenhum responsável formal por escopo, prazo e qualidade. "
                "3. Backlog de demandas das áreas não foi priorizado para o período Q2, gerando vazio de planejamento."
            ),
            "plano": (
                "1. [Até 27/Mai] Convocar sessão de escopo Q2 com todas as áreas demandantes — output: lista priorizada de no mínimo 3 entregas. "
                "2. [Até 30/Mai] Nomear owner por entrega Q2 com comprometimento formal de prazo. "
                "3. [Paralelo] Iniciar definição de entregas Q3 já nessa sessão — não repetir o mesmo erro no próximo quarter. "
                "4. [Semanal] PMO monitora progresso Q2 com alerta vermelho ao sponsor se 30/Jun chegar sem entregas formalmente aceitas."
            ),
        },
        {
            "titulo": "Estratégia de Dados e Digital Boardroom — IDP 0,51 com DOWNStream crítico",
            "impacto": (
                "O sub-projeto Estratégia de Dados e Digital Boardroom tem IDP=0,51 — realiza metade do trabalho planejado. "
                "O DOWNStream deste sub-projeto, que engloba o CEO Digital Boardroom, está com SPI=0,50 e Fim=Nov/2026. "
                "Se o ritmo não for corrigido, o encerramento do programa de Estratégia de Dados ficará incompleto, comprometendo a entrega do roadmap digital prometido para 2026."
            ),
            "causa": (
                "1. CEO Digital Boardroom com SPI=0,37 puxando o índice do sub-projeto para baixo. "
                "2. UPStream concluído (100%) mas sem transição estruturada para o DOWNStream — gap de handoff entre fases. "
                "3. Capacidade técnica da consultoria Bridge não contratada formalmente para a fase de execução."
            ),
            "plano": (
                "1. Tratar o CEO Digital Boardroom como projeto independente com PMO dedicado a partir de Jun/2026. "
                "2. Formalizar o handoff entre UPStream (concluído) e DOWNStream com ata de reunião e aceite do patrocinador. "
                "3. Contrato Bridge: definir prazo limite de assinatura (05/Jun) com cláusula de penalidade por atraso no início. "
                "4. Revisão mensal do IDP do sub-projeto com o sponsor — alerta se IDP não atingir 0,70 até Jul/2026."
            ),
        },
    ],
}

projetos_gov = PROJETOS_PORTFOLIO

for proj in projetos_gov:
    k = proj.replace(" ", "_").replace("/", "_")

    # Inicializa com análise PMO se ainda não foi editado pelo usuário
    if k not in st.session_state.gov_data:
        if k in GOV_DEFAULTS:
            st.session_state.gov_data[k] = GOV_DEFAULTS[k]
        elif usar_ia:
            df_proj = df_view[df_view['projeto'] == proj].copy()
            _t = df_proj[(df_proj['spi_num'].notna())&(df_proj['nivel']<=4)][
                ['nome','nivel','pct','spi_num','inicio','termino','status']].copy()
            _t['inicio']  = _t['inicio'].dt.strftime('%Y-%m-%d').fillna('')
            _t['termino'] = _t['termino'].dt.strftime('%Y-%m-%d').fillna('')
            _t = _t.rename(columns={'spi_num':'spi'})
            with st.spinner(f"Analisando {proj} com IA..."):
                _res = gerar_governanca_ia(proj, json.dumps(_t.to_dict('records'), ensure_ascii=False))
            st.session_state.gov_data[k] = _res if _res else [{"titulo":"","impacto":"","causa":"","plano":""}]
            if _res: st.rerun()
        else:
            st.session_state.gov_data[k] = [{"titulo":"","impacto":"","causa":"","plano":""}]

    # Badge de alerta pelo IDP
    idp_val = idp_por_projeto_final.get(proj)
    alerta  = "🔴" if (idp_val and idp_val < 0.95) else ("⚠️" if (idp_val and idp_val < 0.99) else "✅")
    idp_txt = f"IDP {idp_val:.2f}" if idp_val else "IDP N/A"
    n_linhas = len(st.session_state.gov_data[k])
    label_exp = f"{alerta} {proj} · {idp_txt} · {n_linhas} ponto(s)"

    with st.expander(label_exp, expanded=False):

        linhas = st.session_state.gov_data[k]
        modo_edicao = st.session_state.get(f"edit_mode_{k}", False) and not _apresentando

        # ── MODO LEITURA (padrão executivo) ──────────────────────────────────
        if not modo_edicao:
            for idx, linha in enumerate(linhas):
                titulo  = linha.get("titulo", "") or f"Ponto Crítico {idx+1}"
                impacto = linha.get("impacto", "")
                causa   = linha.get("causa", "")
                plano   = linha.get("plano", "")
                # Só exibe linhas com conteúdo
                if not any([impacto, plano]): continue
                if idx > 0:
                    st.markdown("<hr style='border:1px solid #F0F2F6;margin:14px 0'>", unsafe_allow_html=True)
                st.markdown(f"<strong style='font-size:14px;color:#1B2A4A'>{titulo}</strong>", unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"<div style='font-size:10px;font-weight:700;color:#9AA5BE;text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px'>📌 Impacto no Negócio</div><div style='font-size:13px;color:#1B2A4A;line-height:1.6'>{impacto or '—'}</div>", unsafe_allow_html=True)
                with c2:
                    st.markdown(f"<div style='font-size:10px;font-weight:700;color:#9AA5BE;text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px'>✅ Plano de Ação</div><div style='font-size:13px;color:#1B2A4A;line-height:1.6'>{plano or '—'}</div>", unsafe_allow_html=True)

            st.markdown("")
            if not _apresentando:
                _bc = st.columns([1, 1]) if usar_ia else [st.container()]
                with _bc[0]:
                    if st.button("✏️ Editar", key=f"btn_edit_{k}", use_container_width=True):
                        st.session_state[f"edit_mode_{k}"] = True
                        st.rerun()
                if usar_ia:
                    with _bc[1]:
                        if st.button("🤖 Regenerar IA", key=f"btn_ia_{k}", use_container_width=True):
                            df_proj = df_view[df_view['projeto'] == proj].copy()
                            _t = df_proj[(df_proj['spi_num'].notna())&(df_proj['nivel']<=4)][
                                ['nome','nivel','pct','spi_num','inicio','termino','status']].copy()
                            _t['inicio']  = _t['inicio'].dt.strftime('%Y-%m-%d').fillna('')
                            _t['termino'] = _t['termino'].dt.strftime('%Y-%m-%d').fillna('')
                            _t = _t.rename(columns={'spi_num':'spi'})
                            with st.spinner(f"Analisando {proj}..."):
                                _res = gerar_governanca_ia.__wrapped__(proj, json.dumps(_t.to_dict('records'), ensure_ascii=False))
                            if _res:
                                st.session_state.gov_data[k] = _res
                                st.rerun()

        # ── MODO EDIÇÃO ───────────────────────────────────────────────────────
        else:
            to_delete = []

            def _auto_h(txt, min_h=100, chars_per_line=52, line_h=20, pad=40):
                if not txt: return min_h
                lines = sum(max(1, (len(p)//chars_per_line)+1) for p in txt.split('\n'))
                return max(min_h, lines * line_h + pad)

            for idx, linha in enumerate(linhas):
                if idx > 0:
                    st.markdown("<hr style='border:1px dashed #E2E8F0;margin:10px 0'>", unsafe_allow_html=True)

                col_titulo, col_del = st.columns([9, 1])
                with col_titulo:
                    linhas[idx]["titulo"] = st.text_input(
                        f"Ponto crítico {idx+1}",
                        value=linhas[idx].get("titulo", ""),
                        placeholder=f"Nome do ponto crítico {idx+1}...",
                        key=f"titulo_{k}_{idx}",
                        label_visibility="collapsed",
                    )
                with col_del:
                    if st.button("🗑️", key=f"del_{k}_{idx}", help="Remover"):
                        to_delete.append(idx)

                _hmax = max(
                    _auto_h(linhas[idx].get("impacto","")),
                    _auto_h(linhas[idx].get("plano","")),
                )
                c1, c2 = st.columns(2)
                with c1:
                    linhas[idx]["impacto"] = st.text_area("📌 Impacto no Negócio",
                        value=linhas[idx].get("impacto",""), height=_hmax,
                        placeholder="Descreva o impacto no negócio...",
                        key=f"impacto_{k}_{idx}")
                with c2:
                    linhas[idx]["plano"] = st.text_area("✅ Plano de Ação",
                        value=linhas[idx].get("plano",""), height=_hmax,
                        placeholder="Ações, responsáveis e prazo...",
                        key=f"plano_{k}_{idx}")

            for idx in sorted(to_delete, reverse=True):
                if len(linhas) > 1:
                    linhas.pop(idx)
            st.session_state.gov_data[k] = linhas

            col_add, col_save = st.columns([1, 1])
            with col_add:
                if st.button("➕ Adicionar ponto crítico", key=f"add_{k}"):
                    st.session_state.gov_data[k].append({"titulo":"","impacto":"","causa":"","plano":""})
                    # Não usar st.rerun() aqui — preserva o estado do roadmap JS
                    st.session_state[f"edit_mode_{k}"] = True
            with col_save:
                if st.button("✅ Concluir edição", key=f"btn_save_{k}", type="primary"):
                    st.session_state[f"edit_mode_{k}"] = False

# 12. EXPORTAR RELATÓRIO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)

def gerar_pdf_executivo(df_view, idp_por_projeto, gov_data, data_ref):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        HRFlowable, KeepTogether
    )
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

    buf = io.BytesIO()
    W, H = A4
    mg = 1.8 * cm

    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=mg, rightMargin=mg, topMargin=mg, bottomMargin=mg,
        title="Dashboard Executivo - Digital")

    NAVY  = rl_colors.HexColor("#1B2A4A")
    BLUE  = rl_colors.HexColor("#2563EB")
    SLATE = rl_colors.HexColor("#64748B")
    LGRAY = rl_colors.HexColor("#F1F5F9")
    MGRAY = rl_colors.HexColor("#CBD5E1")
    GREEN = rl_colors.HexColor("#059669")
    AMBER = rl_colors.HexColor("#D97706")
    RED   = rl_colors.HexColor("#DC2626")
    WHITE = rl_colors.white

    def cor_idp(v):
        if v is None: return SLATE
        if v >= 0.99: return GREEN
        if v >= 0.95: return AMBER
        return RED

    def label_idp(v):
        if v is None: return "N/A"
        if v >= 0.99: return "Em dia"
        if v >= 0.95: return "Em alerta"
        return "Em atraso"

    def sty(name, **kw):
        base = {"fontName":"Helvetica","fontSize":9,"textColor":NAVY,"leading":13,"alignment":TA_LEFT}
        base.update(kw)
        return ParagraphStyle(name, **base)

    S_TITLE  = sty("title",  fontName="Helvetica-Bold", fontSize=18, textColor=NAVY, spaceAfter=2)
    S_SUB    = sty("sub",    fontSize=8, textColor=SLATE)
    S_SEC    = sty("sec",    fontName="Helvetica-Bold", fontSize=10, textColor=BLUE, leading=16)
    S_LABEL  = sty("label",  fontName="Helvetica-Bold", fontSize=7, textColor=SLATE, leading=10)
    S_VAL    = sty("val",    fontName="Helvetica-Bold", fontSize=20, textColor=NAVY, leading=22)
    S_SMALL  = sty("small",  fontSize=8, textColor=SLATE, leading=11)
    S_BODY   = sty("body",   fontSize=8.5, textColor=NAVY, leading=13)
    S_TITULO = sty("titulo", fontName="Helvetica-Bold", fontSize=9, textColor=NAVY, leading=13)

    story = []

    # ── Cabeçalho ─────────────────────────────────────────────────────────────
    story.append(Paragraph("Dashboard Executivo - Digital", S_TITLE))
    story.append(Paragraph(
        f"PMBOK 8a Edicao  |  Referencia: {data_ref.strftime('%d/%m/%Y')}  |  "
        f"Gerado em: {date.today().strftime('%d/%m/%Y')}",
        S_SUB))
    story.append(HRFlowable(width="100%", thickness=2, color=BLUE, spaceAfter=14))

    # ── A. KPIs ───────────────────────────────────────────────────────────────
    story.append(Paragraph("A. Governanca Estrategica - KPIs do Portfolio", S_SEC))
    story.append(Spacer(1, 0.2*cm))

    df_root   = df_view[(df_view['nivel']==1) &
                        (~df_view['nome'].isin(['MS Project_Teste_Formulas_2','NOME DO PROJETO']))]
    n_proj    = df_view['projeto'].nunique()
    pct_media = df_root['pct'].mean() if not df_root.empty else 0
    idp_vals  = [v for v in idp_por_projeto.values() if v]
    spi_med   = round(sum(idp_vals)/len(idp_vals), 2) if idp_vals else None
    crits_n   = sum(1 for v in idp_por_projeto.values() if v and v < 0.95)
    marcos_n  = df_view[df_view['is_milestone']].shape[0]

    def kpi_cell(label, val, sub=""):
        return [Paragraph(label, S_LABEL), Paragraph(str(val), S_VAL), Paragraph(sub, S_SMALL)]

    cw5 = (W - 2*mg) / 5
    t_kpi = Table([[
        kpi_cell("PROJETOS ATIVOS", n_proj, "monitorados"),
        kpi_cell("CONCLUSAO MEDIA", f"{pct_media:.1f}%", "do portfolio"),
        kpi_cell("IDP PORTFOLIO", f"{spi_med:.2f}" if spi_med else "N/A", "indice de desempenho"),
        kpi_cell("PROJETOS CRITICOS", crits_n, "IDP < 0,95"),
        kpi_cell("MARCOS", marcos_n, "identificados"),
    ]], colWidths=[cw5]*5)
    t_kpi.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), LGRAY),
        ('BOX',           (0,0),(-1,-1), 0.5, MGRAY),
        ('INNERGRID',     (0,0),(-1,-1), 0.3, MGRAY),
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
        ('TOPPADDING',    (0,0),(-1,-1), 8),
        ('BOTTOMPADDING', (0,0),(-1,-1), 8),
        ('LEFTPADDING',   (0,0),(-1,-1), 10),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 0.25*cm))

    # IDP por projeto
    if idp_por_projeto:
        cw_idp = (W - 2*mg) / len(idp_por_projeto)
        idp_row = []
        for proj, v in idp_por_projeto.items():
            c = cor_idp(v)
            idp_row.append([
                Paragraph(proj[:28], S_LABEL),
                Paragraph(f"IDP {v:.2f}" if v else "N/A",
                          ParagraphStyle("iv", fontName="Helvetica-Bold",
                                         fontSize=16, textColor=c, leading=18)),
                Paragraph(label_idp(v),
                          ParagraphStyle("il", fontSize=8, textColor=c, leading=10)),
            ])
        t_idp = Table([idp_row], colWidths=[cw_idp]*len(idp_por_projeto))
        t_idp.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,-1), WHITE),
            ('BOX',           (0,0),(-1,-1), 0.5, MGRAY),
            ('INNERGRID',     (0,0),(-1,-1), 0.3, MGRAY),
            ('TOPPADDING',    (0,0),(-1,-1), 8),
            ('BOTTOMPADDING', (0,0),(-1,-1), 8),
            ('LEFTPADDING',   (0,0),(-1,-1), 10),
            ('VALIGN',        (0,0),(-1,-1), 'TOP'),
        ]))
        for i, (proj, v) in enumerate(idp_por_projeto.items()):
            t_idp.setStyle(TableStyle([('LINEBEFORE',(i,0),(i,0), 4, cor_idp(v))]))
        story.append(t_idp)

    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=MGRAY, spaceAfter=10))

    # ── B. Roadmap ────────────────────────────────────────────────────────────
    story.append(Paragraph("B. Roadmap Executivo & Marcos de Valor", S_SEC))
    story.append(Spacer(1, 0.2*cm))

    MARCOS_CURADOS_PDF = {
        "Business Data Fabric": [
            {"nome": "Kick-off Projeto",               "data": "2026-03-25", "pct": 100},
            {"nome": "Assinatura Contrato MS Fabric",  "data": "2026-06-09", "pct": 0},
            {"nome": "Assinatura Contrato Implantacao","data": "2026-06-09", "pct": 0},
            {"nome": "Ativacao SKU / Go-Live Fabric",  "data": "2026-06-23", "pct": 0},
            {"nome": "Publicacao Politica Governanca", "data": "2026-04-30", "pct": 0},
            {"nome": "SAC GO-LIVE",                   "data": "2026-07-02", "pct": 0},
            {"nome": "Encerramento",                   "data": "2026-12-11", "pct": 0},
        ],
        "Cockpit Engenharia": [
            {"nome": "Aprovacao TAP",                  "data": "2026-05-22", "pct": 64},
            {"nome": "Kick-off Fase 1",                "data": "2026-06-05", "pct": 100},
            {"nome": "Assinatura Contrato Fase 2",     "data": "2026-08-14", "pct": 0},
            {"nome": "GO-LIVE",                        "data": "2026-10-29", "pct": 0},
        ],
        "Esteira Analytics": [
            {"nome": "Definicao Entregas Q1",          "data": "2026-01-02", "pct": 100},
            {"nome": "Assinatura Contrato Estrategia", "data": "2026-03-06", "pct": 100},
            {"nome": "Go Live Estrategia de Dados",    "data": "2026-04-30", "pct": 100},
            {"nome": "Conclusao Entregas Q1",          "data": "2026-05-27", "pct": 87},
            {"nome": "Conclusao Entregas Q2",          "data": "2026-06-30", "pct": 0},
            {"nome": "Conclusao Entregas Q3",          "data": "2026-09-30", "pct": 0},
            {"nome": "Conclusao Entregas Q4",          "data": "2026-12-31", "pct": 0},
        ],
    }

    from datetime import datetime as _dtpdf

    # Calcula min/max de todas as datas para escala
    all_dates = []
    df_l1 = df_view[(df_view['nivel']==1) &
                    (~df_view['nome'].isin(['MS Project_Teste_Formulas_2','NOME DO PROJETO']))]
    for _, r in df_l1.iterrows():
        if pd.notna(r['inicio']):  all_dates.append(r['inicio'].date())
        if pd.notna(r['termino']): all_dates.append(r['termino'].date())
    for ms_list in MARCOS_CURADOS_PDF.values():
        for m in ms_list:
            all_dates.append(_dtpdf.strptime(m['data'], '%Y-%m-%d').date())

    if not all_dates:
        story.append(Paragraph("Dados de roadmap nao disponiveis.", sty("na")))
    else:
        from datetime import timedelta
        d_min = min(all_dates)
        d_max = max(all_dates)
        span  = (d_max - d_min).days or 1

        BAR_W    = W - 2*mg - 3.5*cm   # largura util da barra
        LABEL_W  = 3.4*cm
        ROW_H    = 0.75*cm
        HOJE     = date.today()

        def d2x(d):
            return LABEL_W + ((d - d_min).days / span) * BAR_W

        from reportlab.platypus import Flowable

        class RoadmapFlowable(Flowable):
            def __init__(self, df_l1, marcos_dict, d_min, d_max, span,
                         bar_w, label_w, row_h, hoje, cor_idp_fn, idp_proj, projetos):
                super().__init__()
                self.df_l1       = df_l1
                self.marcos_dict = marcos_dict
                self.d_min       = d_min
                self.d_max       = d_max
                self.span        = span
                self.bar_w_base  = bar_w
                self.label_w     = label_w
                self.row_h       = row_h
                self.hoje        = hoje
                self.cor_idp_fn  = cor_idp_fn
                self.idp_proj    = idp_proj
                self.projetos    = projetos
                self.HEADER_H    = 0.6*cm
                self.total_h     = self.HEADER_H + len(projetos) * row_h

            def wrap(self, avW, avH):
                self.avW = avW
                return (avW, self.total_h)

            def draw(self):
                c         = self.canv
                label_w   = self.label_w
                bar_w     = self.avW - label_w
                HEADER_H  = self.HEADER_H
                total_h   = self.total_h
                d_min     = self.d_min
                d_max     = self.d_max
                span      = self.span
                row_h     = self.row_h
                hoje      = self.hoje

                NAVY    = rl_colors.HexColor("#1B2A4A")
                MGRAY   = rl_colors.HexColor("#CBD5E1")
                LGRAY   = rl_colors.HexColor("#F1F5F9")
                SLATE   = rl_colors.HexColor("#64748B")
                GREEN   = rl_colors.HexColor("#059669")
                AMBER   = rl_colors.HexColor("#D97706")
                RED     = rl_colors.HexColor("#DC2626")
                BLUEC   = rl_colors.HexColor("#3B82F6")
                TODAY_C = rl_colors.HexColor("#63B3ED")

                def d2x(d):
                    return label_w + ((d - d_min).days / span) * bar_w

                # ── Meses header ─────────────────────────────────────────────
                from datetime import date as _dt_date
                cur   = _dt_date(d_min.year, d_min.month, 1)
                y_hdr = total_h - HEADER_H
                while cur <= d_max:
                    next_m = _dt_date(cur.year + (cur.month // 12), (cur.month % 12) + 1, 1)
                    x0 = max(d2x(cur),    label_w)
                    x1 = min(d2x(next_m), label_w + bar_w)
                    w  = x1 - x0
                    if w > 2:
                        c.setFillColor(LGRAY)
                        c.rect(x0, y_hdr, w, HEADER_H, fill=1, stroke=0)
                        c.setFillColor(SLATE)
                        c.setFont("Helvetica-Bold", 5.5)
                        c.drawCentredString(x0 + w/2, y_hdr + HEADER_H*0.3, cur.strftime('%b/%y'))
                    xg = d2x(cur)
                    if label_w <= xg <= label_w + bar_w:
                        c.setStrokeColor(MGRAY); c.setLineWidth(0.3)
                        c.line(xg, 0, xg, total_h)
                    cur = next_m
                c.setStrokeColor(MGRAY); c.setLineWidth(0.4)
                c.rect(label_w, y_hdr, bar_w, HEADER_H, fill=0, stroke=1)

                # ── Linha de hoje ─────────────────────────────────────────────
                if d_min <= hoje <= d_max:
                    xh = d2x(hoje)
                    c.setStrokeColor(TODAY_C); c.setLineWidth(1.2); c.setDash(3,3)
                    c.line(xh, 0, xh, y_hdr); c.setDash()
                    c.setFillColor(TODAY_C); c.setFont("Helvetica-Bold", 5)
                    c.drawCentredString(xh, y_hdr - 0.15*cm, "HOJE")

                # ── Rows ──────────────────────────────────────────────────────
                for i, proj in enumerate(self.projetos):
                    y_row = total_h - HEADER_H - (i + 1) * row_h
                    bg    = LGRAY if i % 2 == 0 else rl_colors.white
                    c.setFillColor(bg)
                    c.rect(0, y_row, label_w + bar_w, row_h, fill=1, stroke=0)
                    c.setStrokeColor(MGRAY); c.setLineWidth(0.3)
                    c.line(0, y_row, label_w + bar_w, y_row)

                    v = self.idp_proj.get(proj)
                    cor_v = self.cor_idp_fn(v)
                    c.setFillColor(NAVY); c.setFont("Helvetica-Bold", 7)
                    short = proj[:20]+"..." if len(proj)>20 else proj
                    c.drawString(3, y_row + row_h*0.55, short)
                    c.setFont("Helvetica", 6); c.setFillColor(cor_v)
                    c.drawString(3, y_row + row_h*0.2, f"IDP {v:.2f}" if v else "N/A")

                    row = self.df_l1[self.df_l1['projeto']==proj]
                    if not row.empty:
                        r0 = row.iloc[0]
                        if pd.notna(r0['inicio']) and pd.notna(r0['termino']):
                            ini_d = r0['inicio'].date()
                            fim_d = r0['termino'].date()
                            for m in self.marcos_dict.get(proj,[]):
                                md2 = _dtpdf.strptime(m['data'],'%Y-%m-%d').date()
                                if md2 > fim_d: fim_d = md2
                            xs = max(d2x(ini_d), label_w)
                            xe = min(d2x(fim_d), label_w + bar_w)
                            bh = row_h * 0.38
                            yb = y_row + (row_h - bh)/2
                            c.setFillColor(rl_colors.HexColor("#F1F5F9"))
                            c.setStrokeColor(MGRAY); c.setLineWidth(0.4)
                            c.roundRect(xs, yb, xe-xs, bh, 2, fill=1, stroke=1)
                            fw = (xe-xs)*float(r0['pct'])/100
                            if fw > 1:
                                c.setFillColor(rl_colors.HexColor("#94A3B8"))
                                c.roundRect(xs, yb, fw, bh, 2, fill=1, stroke=0)
                            c.setFillColor(SLATE); c.setFont("Helvetica-Bold", 5.5)
                            c.drawCentredString((xs+xe)/2, yb+bh*0.25, f"{r0['pct']:.0f}% concluido")

                    for m in self.marcos_dict.get(proj,[]):
                        md = _dtpdf.strptime(m['data'],'%Y-%m-%d').date()
                        if not (d_min <= md <= d_max): continue
                        xm = d2x(md)
                        ym_c = y_row + row_h*0.72
                        nome_lower = m['nome'].lower().replace('-','').replace(' ','')
                        if 'golive' in nome_lower:
                            c.setFillColor(AMBER); c.setFont("Helvetica-Bold", 9)
                            c.drawCentredString(xm, ym_c - 3, "*")
                        else:
                            mk_c = GREEN if m['pct']>=100 else (RED if md<hoje else BLUEC)
                            sz = 3.5
                            c.setFillColor(mk_c)
                            p = c.beginPath()
                            p.moveTo(xm, ym_c+sz); p.lineTo(xm+sz, ym_c)
                            p.lineTo(xm, ym_c-sz); p.lineTo(xm-sz, ym_c); p.close()
                            c.drawPath(p, fill=1, stroke=0)
                        c.setFillColor(SLATE); c.setFont("Helvetica", 4)
                        lbl = m['nome'][:12]+"..." if len(m['nome'])>12 else m['nome']
                        c.drawCentredString(xm, y_row + 1.5, lbl)

                c.setStrokeColor(MGRAY); c.setLineWidth(0.5)
                c.rect(label_w, 0, bar_w, total_h, fill=0, stroke=1)

        projetos = sorted(df_l1['projeto'].unique().tolist())

        roadmap = RoadmapFlowable(
            df_l1       = df_l1,
            marcos_dict = MARCOS_CURADOS_PDF,
            d_min       = d_min, d_max=d_max, span=span,
            bar_w       = BAR_W, label_w=LABEL_W, row_h=ROW_H,
            hoje        = HOJE,
            cor_idp_fn  = cor_idp,
            idp_proj    = idp_por_projeto,
            projetos    = projetos,
        )
        story.append(roadmap)

    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=MGRAY, spaceAfter=10))

    # ── C. Governança ─────────────────────────────────────────────────────────
    story.append(Paragraph(
        "C. Governanca de Incertezas: Pontos Criticos e Planos de Acao", S_SEC))
    story.append(Spacer(1, 0.2*cm))

    PROJ_COLORS = {
        "Business Data Fabric": rl_colors.HexColor("#1B2A4A"),
        "Cockpit Engenharia":   rl_colors.HexColor("#1B4332"),
        "Esteira Analytics":    rl_colors.HexColor("#5C2D19"),
    }

    for proj in sorted(df_view['projeto'].unique().tolist()):
        # Tenta múltiplos formatos de chave
        k1 = proj.replace(" ","_").replace("/","_")
        k2 = proj
        k3 = proj.replace(" ","").replace("/","")
        linhas_raw = gov_data.get(k1) or gov_data.get(k2) or gov_data.get(k3) or []
        linhas = [l for l in linhas_raw
                  if any([l.get("impacto"), l.get("causa"), l.get("plano")])]
        if not linhas:
            continue

        v          = idp_por_projeto.get(proj)
        proj_color = PROJ_COLORS.get(proj, NAVY)
        idp_s      = f"IDP {v:.2f}" if v else "N/A"

        # Header do projeto
        t_hdr = Table([[
            Paragraph(proj, ParagraphStyle(
                "ph", fontName="Helvetica-Bold", fontSize=11,
                textColor=WHITE, leading=14)),
            Paragraph(f"{idp_s} | {label_idp(v)}", ParagraphStyle(
                "pi", fontName="Helvetica-Bold", fontSize=9,
                textColor=WHITE, leading=14, alignment=TA_RIGHT)),
        ]], colWidths=[(W-2*mg)*0.65, (W-2*mg)*0.35])
        t_hdr.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,-1), proj_color),
            ('TOPPADDING',    (0,0),(-1,-1), 8),
            ('BOTTOMPADDING', (0,0),(-1,-1), 8),
            ('LEFTPADDING',   (0,0),(0,0),  12),
            ('RIGHTPADDING',  (1,0),(1,0),  12),
            ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
        ]))
        story.append(KeepTogether([t_hdr]))

        # Pontos críticos
        cw3 = (W - 2*mg) / 3
        for idx, linha in enumerate(linhas):
            titulo  = linha.get("titulo", f"Ponto {idx+1}")
            impacto = linha.get("impacto", "")
            causa   = linha.get("causa",   "")
            plano   = linha.get("plano",   "")
            row_bg  = LGRAY if idx % 2 == 0 else WHITE

            t_titulo = Table(
                [[Paragraph(titulo, S_TITULO)]],
                colWidths=[W - 2*mg])
            t_titulo.setStyle(TableStyle([
                ('BACKGROUND',  (0,0),(-1,-1), row_bg),
                ('TOPPADDING',  (0,0),(-1,-1), 8),
                ('LEFTPADDING', (0,0),(-1,-1), 8),
                ('LINEBEFORE',  (0,0),(-1,-1), 3, cor_idp(v)),
            ]))

            t_cols = Table([[
                [Paragraph("IMPACTO NO NEGOCIO", S_LABEL), Paragraph(impacto, S_BODY)],
                [Paragraph("CAUSA RAIZ",          S_LABEL), Paragraph(causa,   S_BODY)],
                [Paragraph("PLANO DE ACAO",        S_LABEL), Paragraph(plano,   S_BODY)],
            ]], colWidths=[cw3, cw3, cw3])
            t_cols.setStyle(TableStyle([
                ('BACKGROUND',    (0,0),(-1,-1), row_bg),
                ('VALIGN',        (0,0),(-1,-1), 'TOP'),
                ('TOPPADDING',    (0,0),(-1,-1), 8),
                ('BOTTOMPADDING', (0,0),(-1,-1), 10),
                ('LEFTPADDING',   (0,0),(-1,-1), 8),
                ('RIGHTPADDING',  (0,0),(-1,-1), 8),
                ('INNERGRID',     (0,0),(-1,-1), 0.3, MGRAY),
                ('BOX',           (0,0),(-1,-1), 0.3, MGRAY),
            ]))
            story.append(KeepTogether([t_titulo, t_cols]))

        story.append(Spacer(1, 0.4*cm))

    # ── Rodapé ────────────────────────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=0.5, color=MGRAY, spaceBefore=10))
    story.append(Paragraph(
        f"Dashboard Executivo - Digital  |  PMBOK 8a Edicao  |  PMI  |  "
        f"{date.today().strftime('%d/%m/%Y')}",
        ParagraphStyle("foot", fontSize=7, textColor=SLATE, alignment=TA_CENTER)
    ))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ── Export PDF in sidebar ─────────────────────────────────────────────────────
with _sb_export_placeholder.container():
    if st.button('📄 Gerar PDF', use_container_width=True):
        with st.spinner('Gerando PDF executivo...'):
            pdf_bytes = gerar_pdf_executivo(
                df_view,
                idp_por_projeto_final,
                st.session_state.gov_data,
                data_ref,
            )
        st.download_button(
            '⬇️ Baixar PDF',
            data=pdf_bytes,
            file_name=f'dashboard_executivo_{date.today().isoformat()}.pdf',
            mime='application/pdf',
        )


# ── Botão Apresentar / Modo Edição no rodapé ─────────────────────────────────
st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)
_col_foot_l, _col_foot_r = st.columns([6, 1])
with _col_foot_l:
    st.markdown(
        f"<p style='color:#C0C8D8;font-size:10px;margin-top:8px'>"
        f"Dashboard Executivo · PMBOK® 8ª Edição · PMI · Gerado em {date.today().strftime('%d/%m/%Y')}"
        f"</p>", unsafe_allow_html=True,
    )
with _col_foot_r:
    if st.session_state.modo_apresentacao:
        if st.button("⚙️ Modo Edição", use_container_width=True, key="btn_apres_footer"):
            st.session_state.modo_apresentacao = False
            st.rerun()
    else:
        if st.button("🎯 Apresentar", use_container_width=True, type="primary", key="btn_apres_footer"):
            st.session_state.modo_apresentacao = True
            for _k in list(st.session_state.keys()):
                if _k.startswith("edit_mode_"):
                    st.session_state[_k] = False
            st.rerun()
