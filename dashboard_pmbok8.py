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
.stApp { background: #F4F6FB; }
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
.kpi-value { color:#1B2A4A; font-size:28px; font-weight:700; margin:4px 0 2px 0; }
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

        pv      = _cfn(cf, 'Número4') or 0
        ev_val  = _cfn(cf, 'Número6') or 0

        # IDP = %_concluido / %_planejado da linha L1 do projeto
        # %_concluido = Texto1 (BDF/Esteira) ou Número1 (Cockpit)
        # %_planejado = Número7 (BDF/Esteira)
        # IDP direto  = Texto3 (Cockpit, já calculado pelo MS Project)
        def _cfpct(k):
            v = (cf.get(k, '') or '').replace(',', '.').replace('%', '').strip()
            try: return float(v) if v else None
            except: return None

        pct_real = _cfpct('Texto1') or _cfpct('Número1')
        pct_plan = _cfpct('Número7')
        idp_direct_raw = (cf.get('Texto3', '') or '').replace(',', '.').strip()
        try:    idp_direct = float(idp_direct_raw)
        except: idp_direct = None

        if pct_plan and pct_plan > 0 and pct_real is not None:
            spi = round(pct_real / pct_plan, 4)
        elif idp_direct is not None:
            spi = round(idp_direct, 4)
        elif pv > 0 and ev_val > 0:
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
ANTHROPIC_API = "https://api.anthropic.com/v1/messages"

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
        return "⚪", "#9AA5BE", "N/A"
    if spi >= 0.99:
        return "😊", "#059669", "Em dia"
    if spi >= 0.95:
        return "😐", "#D97706", "Em alerta"
    return "😟", "#DC2626", "Em atraso"

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

    # ── Edição de IDPs por projeto ────────────────────────────────────────────
    st.markdown('### ✏️ Editar IDPs por Projeto')
    st.caption('Altere os valores abaixo para recalcular os KPIs.')
    if 'idp_override' not in st.session_state:
        st.session_state.idp_override = {}
    _sb_idp_override = {}
    # Placeholder — preenchido após carregar dados (ver abaixo)
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
    df_full = load_projects(files_meta)
else:
    df_full = pd.DataFrame()

# Atualiza opções de filtro
if not df_full.empty:
    projetos_disp = sorted(df_full['projeto'].unique().tolist())
    # Update filtro options (hack: rerun needed; use session)
    if not st.session_state.get('filtro_proj'):
        filtro_proj = projetos_disp
    df_view = df_full[df_full['projeto'].isin(filtro_proj)] if filtro_proj else df_full
else:
    df_view = df_full
    projetos_disp = []


# ──────────────────────────────────────────────────────────────────────────────
# 7. CABEÇALHO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(
    f"<h1 style='color:#1B2A4A;font-size:24px;font-weight:700;margin-bottom:2px'>"
    f"Dashboard Executivo - Digital</h1>"
    f"<p style='color:#9AA5BE;font-size:12px'>PMBOK 8ª Ed. · Referência: "
    f"{data_ref.strftime('%d/%m/%Y')} · {len(projetos_disp)} projeto(s) carregado(s)</p>",
    unsafe_allow_html=True,
)
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

# IDP por projeto: pega a primeira linha L1 de cada projeto
_l1_rows = df_view[
    (df_view['nivel'] == 1) &
    (~df_view['nome'].isin(['MS Project_Teste_Formulas_2', 'NOME DO PROJETO']))
].drop_duplicates(subset='projeto', keep='first')

idp_por_projeto = {}
for _, r in _l1_rows.iterrows():
    idp_por_projeto[r['projeto']] = round(r['spi_num'], 4) if pd.notna(r['spi_num']) else None

_vals = [v for v in idp_por_projeto.values() if v is not None]
spi_medio_calc = round(sum(_vals) / len(_vals), 4) if _vals else None

# ── IDP por projeto: inputs no sidebar, lidos aqui ──────────────────────────
if 'idp_override' not in st.session_state:
    st.session_state.idp_override = {}

# Renderiza inputs no sidebar via placeholder
idp_por_projeto_final = {}
with _sb_idp_placeholder.container():
    for proj, idp_calc in idp_por_projeto.items():
        idp_editado = st.number_input(
            proj[:28],
            value=float(st.session_state.idp_override.get(proj, idp_calc or 0.0)),
            step=0.01, format='%.2f',
            key=f'idp_input_{proj}'
        )
        st.session_state.idp_override[proj] = idp_editado
        idp_por_projeto_final[proj] = idp_editado


# KPIs recalculados a partir dos IDPs editados
_idp_vals   = [v for v in idp_por_projeto_final.values() if v]
spi_medio   = round(sum(_idp_vals) / len(_idp_vals), 4) if _idp_vals else None
crits_count = sum(1 for v in idp_por_projeto_final.values() if v and v < spi_limiar)
n_projetos  = len(idp_por_projeto_final)
pct_media   = pct_media_calc
marcos_tot  = marcos_tot_calc

# KPI cards (primeira linha)
c1, c2, c3, c4, c5 = st.columns(5)
with c1: kpi_card("PROJETOS ATIVOS", str(n_projetos), "monitorados", "#2563EB")
with c2: kpi_card("CONCLUSÃO MÉDIA", f"{pct_media:.1f}%", "do portfólio", "#059669")
with c3:
    spi_str = f"{spi_medio:.2f}" if spi_medio else "N/A"
    cor3 = "#DC2626" if (spi_medio and spi_medio < 0.95) else ("#D97706" if (spi_medio and spi_medio < 0.99) else "#059669")
    kpi_card("IDP PORTFÓLIO", spi_str, "índice de desempenho", cor3)
with c4: kpi_card("PROJETOS CRÍTICOS", str(crits_count), f"com IDP < {spi_limiar}", "#DC2626")
with c5: kpi_card("MARCOS NO PORTFÓLIO", str(marcos_tot), "identificados", "#7C3AED")

# Carinhas por projeto (segunda linha)
cols_face = st.columns(len(idp_por_projeto_final))
for i, (proj, idp_val) in enumerate(idp_por_projeto_final.items()):
    face, cor_face, label_face = idp_face(idp_val)
    idp_txt = f"{idp_val:.2f}" if idp_val else "N/A"
    with cols_face[i]:
        st.markdown(f"""
<div style='background:#fff;border-radius:10px;padding:10px 16px 8px 16px;
            border-left:5px solid {cor_face};box-shadow:0 1px 6px rgba(0,0,0,.08);
            text-align:center;'>
  <div style='font-size:28px;line-height:1.1;'>{face}</div>
  <div style='font-size:18px;font-weight:700;color:{cor_face};margin:2px 0;'>IDP {idp_txt}</div>
  <div style='font-size:11px;color:{cor_face};font-weight:600;'>{label_face}</div>
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
for proj in sorted(df_view['projeto'].unique()):
    df_p = df_view[df_view['projeto'] == proj]

    # Linha raiz L1 do projeto para datas início/fim
    l1 = df_p[
        (df_p['nivel'] == 1) &
        (~df_p['nome'].isin(SKIP_NAMES)) &
        df_p['inicio'].notna() &
        df_p['termino'].notna()
    ]
    if l1.empty:
        l1 = df_p[df_p['inicio'].notna() & df_p['termino'].notna()]
    if l1.empty: continue

    r0      = l1.iloc[0]
    idp_val = idp_por_projeto_final.get(proj)

    # Cor da barra por etapa PMBOK detectada pelo % de conclusão
    pct_proj = float(r0['pct'])
    if pct_proj == 0:
        bar_cor = "#6366F1"   # Iniciação
    elif pct_proj < 20:
        bar_cor = "#3B82F6"   # Planejamento
    elif pct_proj < 85:
        bar_cor = "#F59E0B"   # Execução e Controle
    else:
        bar_cor = "#22C55E"   # Encerramento

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

    proj_data.append({
        "projeto":  proj,
        "inicio":   _ts(r0['inicio']),
        "termino":  _ts(r0['termino']),
        "baseline": _ts(r0['baseline_termino']),
        "pct":      float(r0['pct']),
        "idp":      idp_val,
        "cor":      bar_cor,
        "marcos":   marcos,
        "subfases": subfases,
    })

if proj_data:
    hoje_ts     = int(pd.Timestamp(date.today()).normalize().value // 1_000_000)
    proj_json   = _json.dumps(proj_data, ensure_ascii=False)
    row_h       = 90
    total_h     = len(proj_data) * row_h + 130

    html_roadmap = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'Inter','Segoe UI',sans-serif;}}
body{{background:#0F1623;color:#E2E8F0;overflow-x:hidden;}}
#wrap{{position:relative;width:100%;padding:0 16px 16px 16px;}}

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
#body{{position:relative;width:100%;}}
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
.marco-wrap:hover .marco{{transform:rotate(45deg) scale(1.5);filter:brightness(1.4);}}
.marco.done{{background:#22C55E;box-shadow:0 0 8px #22C55E88;}}
.marco.ok  {{background:#3B82F6;box-shadow:0 0 6px #3B82F666;}}
.marco.late{{
  background:#EF4444;box-shadow:0 0 12px #EF444499;
  animation:pulse 1.4s ease-in-out infinite;
}}
@keyframes pulse{{0%,100%{{box-shadow:0 0 8px #EF444499;}}50%{{box-shadow:0 0 20px #EF4444CC;}}}}
/* MARCO LABEL — horizontal, abaixo da barra */
.marco-lbl{{
  margin-top:5px;
  font-size:7.5px;font-weight:600;color:#94A3B8;
  white-space:nowrap;
  writing-mode:horizontal-tb;
  transform:none;
  max-width:90px;
  overflow:hidden;text-overflow:ellipsis;
  line-height:1.3;letter-spacing:.01em;
  text-align:center;
}}
.marco-wrap:hover .marco-lbl{{color:#E2E8F0;}}

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
</style>
</head>
<body>
<div id="wrap">
  <div id="header-row"></div>
  <div id="body"></div>
  <div id="legend">
    <div class="leg"><div class="leg-b" style="background:#6366F1"></div>Iniciação (0%)</div>
    <div class="leg"><div class="leg-b" style="background:#3B82F6"></div>Planejamento (&lt;20%)</div>
    <div class="leg"><div class="leg-b" style="background:#F59E0B"></div>Execução e Controle (20–84%)</div>
    <div class="leg"><div class="leg-b" style="background:#22C55E"></div>Encerramento (≥85%)</div>
    <div class="leg" style="margin-left:12px;"><div class="leg-d" style="background:#22C55E"></div>Marco Concluído</div>
    <div class="leg"><div class="leg-d" style="background:#3B82F6"></div>Marco no Prazo</div>
    <div class="leg"><div class="leg-d" style="background:#EF4444"></div>Marco Atrasado</div>
    <div class="leg" style="color:#63B3ED;border-left:2px solid #63B3ED;padding-left:5px;margin-left:12px;">Linha Hoje</div>
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
const PROJECTS = {proj_json};
const HOJE_MS  = {hoje_ts};
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
  if(v==null) return '⚪';
  if(v>=0.99) return '😊';
  if(v>=0.95) return '😐';
  return '😟';
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
    <div class="pidp" style="color:${{idpC}}">${{idpFace(p.idp)}} IDP ${{idpTxt}}</div>`;
  row.appendChild(lbl);

  // Timeline area
  const tla=document.createElement('div');
  tla.className='tl-area';

  // BAR
  if(p.inicio!=null&&p.termino!=null){{
    const x0=d2p(p.inicio), x1=d2p(p.termino);
    const lp=Math.max(0,x0), rp=Math.min(100,x1), wp=rp-lp;
    if(wp>0.05){{
      const bar=document.createElement('div');
      bar.className='bar';
      bar.style.left=`${{lp}}%`; bar.style.width=`${{wp}}%`;
      bar.style.height='32px'; bar.style.top='16px';
      bar.style.background=p.cor; bar.style.opacity='0.85';

      const prog=document.createElement('div');
      prog.className='bar-progress'; prog.style.width=p.pct+'%';
      bar.appendChild(prog);

      const bt=document.createElement('div');
      bt.className='bar-txt';
      bt.textContent=p.pct.toFixed(0)+'% concluído';
      bar.appendChild(bt);

      bar.addEventListener('click', ()=>openModal(p));
      bar.title='Clique para ver detalhes';
      tla.appendChild(bar);
    }}
  }}

  // MARCOS sobre a barra com label abaixo
  p.marcos.forEach(m=>{{
    if(m.termino==null) return;
    const xp=d2p(m.termino);
    if(xp<0||xp>100) return;

    const wrap=document.createElement('div');
    wrap.className='marco-wrap';
    wrap.style.left=`${{xp}}%`;
    // Posiciona o diamante no centro vertical da barra (top=16px, h=32px → centro=32px)
    // O wrap começa no topo do diamante; label fica naturalmente abaixo
    wrap.style.top='24px';

    const mk=document.createElement('div');
    mk.className='marco '+(m.concluido?'done':m.atrasado?'late':'ok');
    wrap.appendChild(mk);

    const lbl=document.createElement('div');
    lbl.className='marco-lbl';
    // Full name — CSS truncates with ellipsis at max-width:90px
    lbl.textContent=m.nome;
    lbl.title=m.nome;
    wrap.appendChild(lbl);

    wrap.addEventListener('click', e=>{{ e.stopPropagation(); openModalMarco(m, p); }});
    wrap.title=m.nome;
    tla.appendChild(wrap);
  }});

  row.appendChild(tla);
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
    st.session_state.gov_del = {}   # controle de exclusão pendente

projetos_gov = sorted(df_view['projeto'].unique().tolist())

for proj in projetos_gov:
    k = proj.replace(" ", "_").replace("/", "_")

    # Inicializa com 1 linha vazia se não existir
    if k not in st.session_state.gov_data:
        st.session_state.gov_data[k] = [{"titulo": "", "impacto": "", "causa": "", "plano": ""}]

    # Badge de alerta pelo IDP
    idp_val = idp_por_projeto_final.get(proj)
    alerta  = "🔴" if (idp_val and idp_val < 0.95) else ("⚠️" if (idp_val and idp_val < 0.99) else "✅")
    idp_txt = f"IDP {idp_val:.2f}" if idp_val else "IDP N/A"
    n_linhas = len(st.session_state.gov_data[k])
    label_exp = f"{alerta} {proj} · {idp_txt} · {n_linhas} ponto(s)"

    with st.expander(label_exp, expanded=False):

        linhas = st.session_state.gov_data[k]
        to_delete = []

        for idx, linha in enumerate(linhas):
            # Separador visual entre linhas
            if idx > 0:
                st.markdown("<hr style='border:1px dashed #E2E8F0;margin:10px 0'>", unsafe_allow_html=True)

            # Título da linha + botão excluir
            col_titulo, col_del = st.columns([9, 1])
            with col_titulo:
                linhas[idx]["titulo"] = st.text_input(
                    f"Ponto crítico {idx + 1}",
                    value=linhas[idx].get("titulo", ""),
                    placeholder=f"Nome do ponto crítico {idx + 1}...",
                    key=f"titulo_{k}_{idx}",
                    label_visibility="collapsed",
                )
            with col_del:
                if st.button("🗑️", key=f"del_{k}_{idx}", help="Remover esta linha"):
                    to_delete.append(idx)

            # 3 campos de conteúdo
            c1, c2, c3 = st.columns(3)
            with c1:
                linhas[idx]["impacto"] = st.text_area(
                    "📌 Impacto no Negócio",
                    value=linhas[idx].get("impacto", ""),
                    height=120,
                    placeholder="Descreva o impacto no negócio...",
                    key=f"impacto_{k}_{idx}",
                )
            with c2:
                linhas[idx]["causa"] = st.text_area(
                    "🔍 Causa Raiz (Hipótese)",
                    value=linhas[idx].get("causa", ""),
                    height=120,
                    placeholder="Descreva a causa raiz identificada...",
                    key=f"causa_{k}_{idx}",
                )
            with c3:
                linhas[idx]["plano"] = st.text_area(
                    "✅ Plano de Ação",
                    value=linhas[idx].get("plano", ""),
                    height=120,
                    placeholder="Ações, responsáveis e prazo...",
                    key=f"plano_{k}_{idx}",
                )

        # Aplica exclusões (de trás para frente para não deslocar índices)
        for idx in sorted(to_delete, reverse=True):
            if len(linhas) > 1:   # mantém pelo menos 1 linha
                linhas.pop(idx)
        st.session_state.gov_data[k] = linhas

        # Botão adicionar nova linha
        st.markdown("")
        if st.button(f"➕ Adicionar ponto crítico", key=f"add_{k}"):
            st.session_state.gov_data[k].append(
                {"titulo": "", "impacto": "", "causa": "", "plano": ""}
            )
            st.rerun()

# 12. EXPORTAR RELATÓRIO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)

def gerar_excel(df_base, df_crits):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Aba 1 — Tarefas completas
        df_base.to_excel(writer, sheet_name='Tarefas', index=False)
        # Aba 2 — Críticos
        df_crits.to_excel(writer, sheet_name='Pontos_Criticos', index=False)
        # Aba 3 — EVM Resumo
        df_evm_exp = df_base[df_base['nivel'] <= 2].copy()
        df_evm_exp.to_excel(writer, sheet_name='EVM_Resumo', index=False)
    return output.getvalue()

# ── Export in sidebar ───────────────────────────────────────────────────────
with _sb_export_placeholder.container():
    if st.button('📥 Gerar Excel Executivo', use_container_width=True):
        excel_bytes = gerar_excel(df_view, df_criticos)
        st.download_button(
            '⬇️ Baixar .xlsx',
            data=excel_bytes,
            file_name=f'relatorio_pmo_{date.today().isoformat()}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )


st.markdown(
    f"<p style='text-align:center;color:#C0C8D8;font-size:10px;margin-top:30px'>"
    f"Dashboard Executivo · PMBOK® 8ª Edição · PMI · Gerado em {date.today().strftime('%d/%m/%Y')}"
    f"</p>", unsafe_allow_html=True,
)
