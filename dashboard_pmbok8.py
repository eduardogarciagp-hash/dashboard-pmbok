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

dfs = []
for proj in PROJETOS_FIXOS:
    dfs.append(_ensure_proj(proj))
    if not df_xml.empty:
        filhas = df_xml[(df_xml['projeto']==proj) & (df_xml['nivel']>1)]
        if not filhas.empty:
            dfs.append(filhas)

df_view = build_df(pd.concat(dfs, ignore_index=True)) if dfs else pd.DataFrame()
projetos_disp = PROJETOS_FIXOS


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
  .hdr h1{{color:#1B2A4A;font-size:22px;font-weight:700;margin:0;line-height:1.2;}}
  .hdr p{{color:#9AA5BE;font-size:11px;margin:3px 0 0 0;}}
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
for proj in PROJETOS_FIXOS:
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
        for proj in PROJETOS_FIXOS:
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
        for proj in PROJETOS_FIXOS:
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
for proj in PROJETOS_FIXOS:
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
  font-size:9px;font-weight:700;color:#475569;
  text-transform:uppercase;letter-spacing:.07em;
  border-right:1px solid #1E2D42;
}}

/* BODY */
#body{{position:relative;width:100%;overflow:visible;}}
.gridline{{position:absolute;top:0;bottom:0;width:1px;background:rgba(100,116,139,.15);pointer-events:none;}}
.today-line{{position:absolute;top:0;bottom:0;width:2px;background:rgba(99,179,237,.6);pointer-events:none;z-index:5;}}
.today-lbl{{
  position:absolute;top:4px;
  font-size:9px;color:#63B3ED;font-weight:700;
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
.proj-label .pname{{font-size:11px;font-weight:700;color:#CBD5E1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.proj-label .pidp{{font-size:10px;font-weight:600;margin-top:2px;}}

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
  font-size:9px;font-weight:700;color:rgba(255,255,255,.95);
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
  font-size:8px;font-weight:600;color:#64748B;
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
  font-size:11px;line-height:1.7;color:#E2E8F0;
  pointer-events:none;
  z-index:9999;
  white-space:normal;
  transition:opacity .1s;
}}
#marco-tip.show{{display:block;}}
.mt-title{{font-size:12px;font-weight:700;color:#E2E8F0;margin-bottom:8px;
           border-bottom:1px solid #2D3F55;padding-bottom:6px;
           display:flex;align-items:center;gap:6px;}}
.mt-diamond{{width:9px;height:9px;transform:rotate(45deg);
             border-radius:1px;flex-shrink:0;display:inline-block;}}
.mt-row{{display:flex;justify-content:space-between;gap:20px;margin-bottom:3px;}}
.mt-label{{color:#475569;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;}}
.mt-val{{color:#CBD5E1;font-weight:600;text-align:right;}}
.mt-badge{{display:inline-block;padding:2px 10px;border-radius:10px;
           font-size:10px;font-weight:700;margin-top:8px;}}

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
.modal-kpi .mk-label{{font-size:9px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.07em;}}
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
.leg{{display:flex;align-items:center;gap:5px;font-size:9px;color:#475569;}}
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
      font-size:10px;font-weight:700;color:#475569;font-family:'Inter','Segoe UI',sans-serif;`;
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

projetos_gov = PROJETOS_FIXOS

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
                st.markdown(f"**{titulo}**", unsafe_allow_html=False)
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
                    st.rerun()
            with col_save:
                if st.button("✅ Concluir edição", key=f"btn_save_{k}", type="primary"):
                    st.session_state[f"edit_mode_{k}"] = False
                    st.rerun()

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
