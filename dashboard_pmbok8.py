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
    st.markdown("## 📊 Dashboard PMO")
    st.markdown("---")
    st.markdown("### 📂 Importar Projetos (XML)")
    st.caption("Exporte pelo MS Project: Arquivo → Salvar Como → XML do Project")
    uploaded = st.file_uploader(
        "Selecione um ou mais arquivos .xml",
        type=["xml"],
        accept_multiple_files=True,
    )
    st.markdown("---")
    data_ref = st.date_input("📅 Data de Referência", value=date.today())
    st.markdown("---")
    filtro_proj = st.multiselect("🔍 Filtrar por Projeto", options=[], key="filtro_proj")
    spi_limiar  = st.slider("⚠️ Limiar SPI (crítico)", 0.70, 1.00, 0.95, 0.01)
    st.markdown("---")
    usar_ia = st.toggle("🤖 Gerar insights com IA (Claude API)", value=False)


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

# ── Modo edição ───────────────────────────────────────────────────────────────
if 'kpi_edit' not in st.session_state:
    st.session_state.kpi_edit = False
if 'kpi_override' not in st.session_state:
    st.session_state.kpi_override = {}

col_title, col_btn = st.columns([8, 1])
with col_btn:
    if st.button("✏️ Editar" if not st.session_state.kpi_edit else "✅ Salvar", use_container_width=True):
        st.session_state.kpi_edit = not st.session_state.kpi_edit

if st.session_state.kpi_edit:
    st.caption("Edite os valores abaixo e clique em ✅ Salvar para aplicar.")
    ec1, ec2, ec3, ec4, ec5 = st.columns(5)
    ov = st.session_state.kpi_override
    with ec1:
        ov['n_projetos']  = st.number_input("Projetos Ativos",    value=int(ov.get('n_projetos',  n_projetos_calc)),  step=1, min_value=0)
    with ec2:
        ov['pct_media']   = st.number_input("Conclusão Média (%)", value=float(ov.get('pct_media',  round(pct_media_calc,1))), step=0.1, format="%.1f")
    with ec3:
        ov['spi_medio']   = st.number_input("IDP Portfólio",       value=float(ov.get('spi_medio',  spi_medio_calc or 0.0)), step=0.01, format="%.2f")
    with ec4:
        ov['crits_count'] = st.number_input("Projetos Críticos",   value=int(ov.get('crits_count', crits_count_calc)), step=1, min_value=0)
    with ec5:
        ov['marcos_tot']  = st.number_input("Marcos no Portfólio", value=int(ov.get('marcos_tot',  marcos_tot_calc)),  step=1, min_value=0)

    st.caption("IDP por projeto:")
    idp_cols = st.columns(max(1, len(idp_por_projeto)))
    for i, (proj, idp_calc) in enumerate(idp_por_projeto.items()):
        with idp_cols[i]:
            ov[f'idp_{proj}'] = st.number_input(
                proj[:25], value=float(ov.get(f'idp_{proj}', idp_calc or 0.0)),
                step=0.01, format="%.2f", key=f"idp_edit_{proj}"
            )

# ── Lê valores (override ou calculado) ───────────────────────────────────────
ov = st.session_state.kpi_override
n_projetos  = int(ov.get('n_projetos',  n_projetos_calc))
pct_media   = float(ov.get('pct_media',  pct_media_calc))
spi_medio   = float(ov.get('spi_medio',  spi_medio_calc or 0.0)) or None
crits_count = int(ov.get('crits_count', crits_count_calc))
marcos_tot  = int(ov.get('marcos_tot',  marcos_tot_calc))

idp_por_projeto_final = {
    proj: float(ov.get(f'idp_{proj}', idp_por_projeto.get(proj) or 0.0))
    for proj in idp_por_projeto
}

# ── Renderiza KPI cards ───────────────────────────────────────────────────────
c1, c2, c3, c4, c5 = st.columns(5)
with c1: kpi_card("PROJETOS ATIVOS", str(n_projetos), "monitorados", "#2563EB")
with c2: kpi_card("CONCLUSÃO MÉDIA", f"{pct_media:.1f}%", "do portfólio", "#059669")
with c3:
    spi_str = f"{spi_medio:.2f}" if spi_medio else "N/A"
    cor3 = "#DC2626" if (spi_medio and spi_medio < 0.95) else ("#D97706" if (spi_medio and spi_medio < 0.99) else "#059669")
    kpi_card("IDP PORTFÓLIO", spi_str, "índice de desempenho", cor3)
with c4: kpi_card("PROJETOS CRÍTICOS", str(crits_count), f"com IDP < {spi_limiar}", "#DC2626")
with c5: kpi_card("MARCOS NO PORTFÓLIO", str(marcos_tot), "identificados", "#7C3AED")

# ── Linha de IDP por projeto com carinhas ────────────────────────────────────
if idp_por_projeto_final:
    cols_idp = st.columns(len(idp_por_projeto_final))
    for i, (proj, idp_val) in enumerate(idp_por_projeto_final.items()):
        face, cor_face, label_face = idp_face(idp_val)
        idp_txt = f"{idp_val:.2f}" if idp_val is not None else "N/A"
        with cols_idp[i]:
            st.markdown(f"""
<div style="background:#fff;border-radius:10px;padding:12px 16px;
            border-left:5px solid {cor_face};box-shadow:0 1px 6px rgba(0,0,0,.08);
            text-align:center;margin-top:8px;">
  <div style="font-size:10px;font-weight:600;color:#9AA5BE;
              text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px;">
    {proj}
  </div>
  <div style="font-size:28px;line-height:1.1;">{face}</div>
  <div style="font-size:20px;font-weight:700;color:{cor_face};margin:2px 0;">
    IDP {idp_txt}
  </div>
  <div style="font-size:11px;color:{cor_face};font-weight:600;">{label_face}</div>
</div>""", unsafe_allow_html=True)

st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 9. SECTION 1 — ROADMAP PIXEL-PERFECT (HTML/JS interpolação linear)
# ──────────────────────────────────────────────────────────────────────────────
import streamlit.components.v1 as components

st.markdown('<div class="section-title">Section 1 — Roadmap Executivo & Marcos de Valor</div>',
            unsafe_allow_html=True)

df_gantt = df_view[
    (df_view['nivel'].isin([1, 2])) &
    (df_view['inicio'].notna()) &
    (df_view['termino'].notna()) &
    (~df_view['nome'].isin(['MS Project_Teste_Formulas_2', 'NOME DO PROJETO']))
].copy()

df_milestones = df_view[
    df_view['is_milestone'] &
    df_view['termino'].notna() &
    (~df_view['nome'].isin(['MS Project_Teste_Formulas_2', 'NOME DO PROJETO']))
].copy()

if not df_gantt.empty:
    import json as _json

    # ── Serializa dados para JS (UTC timestamps em ms) ────────────────────────
    def _ts(dt):
        """Pandas Timestamp → UTC epoch ms (sem deslocamento de fuso)."""
        import pandas as _pd
        if _pd.isna(dt):
            return None
        return int(dt.normalize().value // 1_000_000)

    rows_js = []
    for _, r in df_gantt.iterrows():
        idp_v = r['spi_num']
        if idp_v is None or (hasattr(idp_v, '__float__') and __import__('math').isnan(float(idp_v))):
            bar_color = "#4A6FA5"
        elif idp_v < 0.95:
            bar_color = "#E05252"
        elif idp_v < 0.99:
            bar_color = "#D97706"
        else:
            bar_color = "#22C55E"

        rows_js.append({
            "projeto": r['projeto'],
            "nome":    r['nome'][:55],
            "nivel":   int(r['nivel']),
            "pct":     float(r['pct']),
            "idp":     round(float(idp_v), 2) if (idp_v is not None and not __import__('math').isnan(float(idp_v or 0))) else None,
            "inicio":  _ts(r['inicio']),
            "termino": _ts(r['termino']),
            "baseline":_ts(r['baseline_termino']),
            "cor":     bar_color,
        })

    miles_js = []
    for _, m in df_milestones.iterrows():
        concluido = m['pct'] >= 100
        atrasado  = (
            pd.notna(m['baseline_termino']) and
            pd.notna(m['termino']) and
            m['termino'] > m['baseline_termino']
        )
        miles_js.append({
            "projeto":  m['projeto'],
            "nome":     m['nome'][:55],
            "termino":  _ts(m['termino']),
            "baseline": _ts(m['baseline_termino']),
            "concluido": concluido,
            "atrasado":  atrasado,
            "pct":       float(m['pct']),
        })

    hoje_ts = int(pd.Timestamp(date.today()).normalize().value // 1_000_000)

    data_json   = _json.dumps(rows_js,  ensure_ascii=False)
    miles_json  = _json.dumps(miles_js, ensure_ascii=False)

    row_h   = 36   # altura por linha px
    n_rows  = len(rows_js)
    body_h  = max(200, n_rows * row_h)
    total_h = body_h + 110  # header + legenda

    html_roadmap = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  *{{box-sizing:border-box;margin:0;padding:0;font-family:'Inter','Segoe UI',sans-serif;}}
  body{{background:#0F1623;color:#E2E8F0;overflow-x:hidden;}}
  #wrap{{position:relative;width:100%;padding:0 16px 16px 16px;}}

  /* HEADER MESES */
  #header-row{{display:flex;position:relative;height:32px;margin-bottom:0;}}
  .month-cell{{
    position:absolute;height:100%;
    display:flex;align-items:center;justify-content:center;
    font-size:10px;font-weight:600;color:#64748B;
    text-transform:uppercase;letter-spacing:.06em;
    border-right:1px solid #1E2D42;
  }}

  /* GRID BODY */
  #body{{position:relative;width:100%;}}
  .gridline{{
    position:absolute;top:0;bottom:0;width:1px;
    background:rgba(100,116,139,.18);pointer-events:none;
  }}
  .today-line{{
    position:absolute;top:0;bottom:0;width:2px;
    background:rgba(99,179,237,.55);pointer-events:none;z-index:10;
  }}
  .today-label{{
    position:absolute;top:-22px;
    transform:translateX(-50%);
    font-size:9px;color:#63B3ED;font-weight:700;
    background:#0F1623;padding:1px 4px;border-radius:3px;
    white-space:nowrap;
  }}

  /* ROWS */
  .gantt-row{{
    position:relative;width:100%;
    border-bottom:1px solid #1A2535;
    display:flex;align-items:center;
  }}
  .row-label{{
    position:absolute;left:0;top:0;bottom:0;
    display:flex;align-items:center;
    font-size:10px;color:#94A3B8;white-space:nowrap;
    overflow:hidden;text-overflow:ellipsis;
    pointer-events:none;
    z-index:5;padding-left:4px;
    max-width:180px;
  }}
  .row-label.nivel1{{color:#CBD5E1;font-weight:700;font-size:11px;}}

  /* BARS */
  .bar{{
    position:absolute;border-radius:4px;cursor:pointer;
    transition:filter .15s,opacity .15s;
    display:flex;align-items:center;overflow:hidden;
  }}
  .bar:hover{{filter:brightness(1.25);z-index:20;}}
  .bar-text{{
    font-size:9px;font-weight:600;color:rgba(255,255,255,.9);
    white-space:nowrap;overflow:hidden;padding:0 6px;
  }}
  /* progress fill inside bar */
  .bar-progress{{
    position:absolute;left:0;top:0;bottom:0;
    background:rgba(255,255,255,.18);border-radius:4px 0 0 4px;
    pointer-events:none;
  }}

  /* MILESTONES */
  .milestone{{
    position:absolute;width:14px;height:14px;
    transform:rotate(45deg) translate(-50%,-50%);
    cursor:pointer;z-index:15;transition:transform .15s,filter .15s;
  }}
  .milestone:hover{{transform:rotate(45deg) translate(-50%,-50%) scale(1.5);filter:brightness(1.4);}}
  .milestone.done{{background:#22C55E;box-shadow:0 0 8px #22C55E88;}}
  .milestone.ok{{background:#3B82F6;box-shadow:0 0 6px #3B82F666;}}
  .milestone.late{{
    background:#EF4444;
    box-shadow:0 0 10px #EF444488;
    animation:pulse-red 1.4s ease-in-out infinite;
  }}
  @keyframes pulse-red{{
    0%,100%{{box-shadow:0 0 8px #EF444499;}}
    50%{{box-shadow:0 0 18px #EF4444DD,0 0 30px #EF444455;}}
  }}

  /* TOOLTIP */
  #tooltip{{
    position:fixed;pointer-events:none;z-index:9999;
    background:#1E293B;border:1px solid #334155;
    border-radius:8px;padding:10px 14px;min-width:200px;
    box-shadow:0 8px 32px rgba(0,0,0,.5);
    font-size:11px;line-height:1.6;color:#E2E8F0;
    opacity:0;transition:opacity .12s;
  }}
  #tooltip.show{{opacity:1;}}
  #tooltip b{{color:#93C5FD;}}
  #tooltip .tt-sep{{border-top:1px solid #334155;margin:6px 0;}}

  /* LEGENDA */
  #legend{{
    display:flex;gap:20px;align-items:center;
    padding:10px 0 0 0;flex-wrap:wrap;
  }}
  .leg-item{{display:flex;align-items:center;gap:6px;font-size:10px;color:#64748B;}}
  .leg-diamond{{width:10px;height:10px;transform:rotate(45deg);flex-shrink:0;}}
  .leg-bar{{width:24px;height:8px;border-radius:2px;flex-shrink:0;}}
</style>
</head>
<body>
<div id="wrap">
  <div id="header-row"></div>
  <div id="body"></div>
  <div id="legend">
    <div class="leg-item"><div class="leg-diamond" style="background:#22C55E"></div>Marco Concluído</div>
    <div class="leg-item"><div class="leg-diamond" style="background:#3B82F6"></div>Marco no Prazo</div>
    <div class="leg-item"><div class="leg-diamond" style="background:#EF4444"></div>Marco Atrasado</div>
    <div class="leg-item"><div class="leg-bar" style="background:#22C55E"></div>IDP ≥ 0,99</div>
    <div class="leg-item"><div class="leg-bar" style="background:#D97706"></div>IDP ≥ 0,95</div>
    <div class="leg-item"><div class="leg-bar" style="background:#E05252"></div>IDP &lt; 0,95</div>
    <div class="leg-item"><div class="leg-bar" style="background:#4A6FA5"></div>IDP N/A</div>
  </div>
</div>
<div id="tooltip"></div>

<script>
const ROWS     = {data_json};
const MILES    = {miles_json};
const HOJE_MS  = {hoje_ts};
const ROW_H    = {row_h};
const LABEL_W  = 190;  // px reserved for row labels on left

// ── 1. CALCULAR MinDate / MaxDate ──────────────────────────────────────────
let minMs = Infinity, maxMs = -Infinity;
ROWS.forEach(r => {{
  if(r.inicio  != null) {{ minMs = Math.min(minMs, r.inicio);  maxMs = Math.max(maxMs, r.inicio);  }}
  if(r.termino != null) {{ minMs = Math.min(minMs, r.termino); maxMs = Math.max(maxMs, r.termino); }}
}});
MILES.forEach(m => {{
  if(m.termino != null) {{ minMs = Math.min(minMs, m.termino); maxMs = Math.max(maxMs, m.termino); }}
}});
// Extend slightly so bars don't clip at edges
const SPAN = maxMs - minMs;
const MIN  = minMs - SPAN * 0.01;
const MAX  = maxMs + SPAN * 0.03;

// ── 2. FUNÇÃO DE INTERPOLAÇÃO LINEAR ──────────────────────────────────────
function dateToPercent(ms) {{
  return ((ms - MIN) / (MAX - MIN)) * 100;
}}

// ── 3. GERAR MESES DO CABEÇALHO ───────────────────────────────────────────
function generateMonths(minMs, maxMs) {{
  const months = [];
  const d = new Date(minMs);
  d.setUTCDate(1);
  while(d.getTime() <= maxMs + 86400000 * 31) {{
    const start = d.getTime();
    const nextD = new Date(d);
    nextD.setUTCMonth(nextD.getUTCMonth() + 1);
    const end = nextD.getTime();
    months.push({{ start, end,
      label: d.toLocaleDateString('pt-BR', {{month:'short', year:'2-digit', timeZone:'UTC'}})
    }});
    d.setUTCMonth(d.getUTCMonth() + 1);
    if(d.getUTCFullYear() > 2030) break;
  }}
  return months;
}}

const MONTHS = generateMonths(MIN, MAX);
const MNAMES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

// ── 4. RENDERIZAR CABEÇALHO ────────────────────────────────────────────────
const headerRow = document.getElementById('header-row');
// Available width for timeline = total - LABEL_W (will use percent inside relative)
MONTHS.forEach(m => {{
  const x0 = dateToPercent(m.start);
  const x1 = dateToPercent(m.end);
  if(x1 < 0 || x0 > 100) return;
  const left  = Math.max(0, x0);
  const right = Math.min(100, x1);
  const w = right - left;
  if(w < 0.5) return;

  // Adjust for label width: shift right by LABEL_W px equivalent
  const el = document.createElement('div');
  el.className = 'month-cell';
  el.style.left  = `calc(${{LABEL_W}}px + (${{left / 100}} * (100% - ${{LABEL_W}}px)))`;
  el.style.width = `calc(${{w / 100}} * (100% - ${{LABEL_W}}px))`;
  const dt = new Date(m.start);
  el.textContent = MNAMES[dt.getUTCMonth()] + '/' + String(dt.getUTCFullYear()).slice(2);
  headerRow.appendChild(el);
}});

// ── 5. RENDERIZAR BODY ────────────────────────────────────────────────────
const body = document.getElementById('body');
body.style.height = (ROW_H * ROWS.length) + 'px';
body.style.position = 'relative';

// Gridlines for each month start
MONTHS.forEach(m => {{
  const xp = dateToPercent(m.start);
  if(xp < 0 || xp > 100) return;
  const gl = document.createElement('div');
  gl.className = 'gridline';
  gl.style.left = `calc(${{LABEL_W}}px + (${{xp / 100}} * (100% - ${{LABEL_W}}px)))`;
  body.appendChild(gl);
}});

// Linha de hoje
if(HOJE_MS >= MIN && HOJE_MS <= MAX) {{
  const xp = dateToPercent(HOJE_MS);
  const tl = document.createElement('div');
  tl.className = 'today-line';
  tl.style.left = `calc(${{LABEL_W}}px + (${{xp / 100}} * (100% - ${{LABEL_W}}px)))`;
  const lbl = document.createElement('div');
  lbl.className = 'today-label';
  lbl.style.left = tl.style.left;
  lbl.style.position = 'absolute';
  lbl.style.top = '-22px';
  const hd = new Date(HOJE_MS);
  lbl.textContent = 'Hoje ' + hd.getUTCDate() + '/' + (hd.getUTCMonth()+1);
  body.appendChild(tl);
  body.appendChild(lbl);
}}

// ── 6. TOOLTIP ────────────────────────────────────────────────────────────
const tooltip = document.getElementById('tooltip');
function showTip(html, e) {{
  tooltip.innerHTML = html;
  tooltip.classList.add('show');
  moveTip(e);
}}
function moveTip(e) {{
  let x = e.clientX + 14, y = e.clientY - 10;
  if(x + 220 > window.innerWidth) x = e.clientX - 230;
  tooltip.style.left = x + 'px';
  tooltip.style.top  = y + 'px';
}}
function hideTip() {{ tooltip.classList.remove('show'); }}
document.addEventListener('mousemove', e => {{ if(tooltip.classList.contains('show')) moveTip(e); }});

function fmtDate(ms) {{
  if(ms == null) return '—';
  const d = new Date(ms);
  return d.getUTCDate().toString().padStart(2,'0') + '/'
       + (d.getUTCMonth()+1).toString().padStart(2,'0') + '/'
       + d.getUTCFullYear();
}}
function idpLabel(v) {{
  if(v == null) return 'N/A';
  if(v >= 0.99) return '<span style="color:#22C55E">' + v.toFixed(2) + ' 😊 Em dia</span>';
  if(v >= 0.95) return '<span style="color:#D97706">' + v.toFixed(2) + ' 😐 Em alerta</span>';
  return '<span style="color:#EF4444">' + v.toFixed(2) + ' 😟 Em atraso</span>';
}}

// ── 7. RENDERIZAR LINHAS + BARRAS ─────────────────────────────────────────
ROWS.forEach((r, i) => {{
  const rowEl = document.createElement('div');
  rowEl.className = 'gantt-row';
  rowEl.style.height = ROW_H + 'px';
  rowEl.style.top = (i * ROW_H) + 'px';
  rowEl.style.position = 'absolute';
  rowEl.style.width = '100%';
  // alternating row bg
  rowEl.style.background = i % 2 === 0 ? 'rgba(255,255,255,.018)' : 'transparent';

  // Label
  const lbl = document.createElement('div');
  lbl.className = 'row-label' + (r.nivel === 1 ? ' nivel1' : '');
  lbl.style.width = LABEL_W + 'px';
  lbl.title = '[' + r.projeto + '] ' + r.nome;
  // Show short version
  const short = r.nome.length > 22 ? r.nome.slice(0,22)+'…' : r.nome;
  lbl.textContent = short;
  rowEl.appendChild(lbl);

  if(r.inicio != null && r.termino != null) {{
    const x0p = dateToPercent(r.inicio);
    const x1p = dateToPercent(r.termino);
    const leftPct  = Math.max(0, x0p);
    const rightPct = Math.min(100, x1p);
    const wPct = rightPct - leftPct;
    if(wPct > 0.05) {{
      const bar = document.createElement('div');
      bar.className = 'bar';
      bar.style.left   = `calc(${{LABEL_W}}px + (${{leftPct/100}} * (100% - ${{LABEL_W}}px)))`;
      bar.style.width  = `calc(${{wPct/100}} * (100% - ${{LABEL_W}}px))`;
      bar.style.height = (r.nivel === 1 ? ROW_H * 0.42 : ROW_H * 0.55) + 'px';
      bar.style.top    = ((ROW_H - (r.nivel === 1 ? ROW_H*0.42 : ROW_H*0.55)) / 2) + 'px';
      bar.style.background = r.cor;
      bar.style.opacity = r.nivel === 1 ? '0.55' : '0.80';

      // Progress fill
      const prog = document.createElement('div');
      prog.className = 'bar-progress';
      prog.style.width = r.pct + '%';
      bar.appendChild(prog);

      // Bar text
      const bt = document.createElement('div');
      bt.className = 'bar-text';
      const idpTxt = r.idp != null ? ' · IDP ' + r.idp.toFixed(2) : '';
      bt.textContent = r.pct.toFixed(0) + '%' + idpTxt;
      bar.appendChild(bt);

      // Tooltip
      bar.addEventListener('mouseenter', e => {{
        showTip(`<b>${{r.nome}}</b><div class="tt-sep"></div>
          <b>Projeto:</b> ${{r.projeto}}<br>
          <b>Início:</b> ${{fmtDate(r.inicio)}}<br>
          <b>Término:</b> ${{fmtDate(r.termino)}}<br>
          <b>Baseline:</b> ${{fmtDate(r.baseline)}}<br>
          <b>Conclusão:</b> ${{r.pct.toFixed(0)}}%<br>
          <b>IDP:</b> ${{idpLabel(r.idp)}}`, e);
      }});
      bar.addEventListener('mouseleave', hideTip);
      rowEl.appendChild(bar);
    }}
  }}
  body.appendChild(rowEl);
}});

// ── 8. RENDERIZAR MARCOS ──────────────────────────────────────────────────
MILES.forEach(m => {{
  if(m.termino == null) return;
  const xp = dateToPercent(m.termino);
  if(xp < 0 || xp > 100) return;

  // Find matching row index
  const rowIdx = ROWS.findIndex(r =>
    r.projeto === m.projeto &&
    Math.abs((r.inicio || 0) - (m.termino || 0)) < 86400000 * 400
  );
  const yCenter = rowIdx >= 0
    ? rowIdx * ROW_H + ROW_H / 2
    : ROWS.length * ROW_H / 2;

  const ms = document.createElement('div');
  ms.className = 'milestone' + (m.concluido ? ' done' : m.atrasado ? ' late' : ' ok');
  ms.style.left = `calc(${{LABEL_W}}px + (${{xp / 100}} * (100% - ${{LABEL_W}}px)))`;
  ms.style.top  = yCenter + 'px';
  ms.title = m.nome;

  ms.addEventListener('mouseenter', e => {{
    const status = m.concluido ? '✅ Concluído' : m.atrasado ? '🔴 Atrasado' : '🔵 No prazo';
    showTip(`<b>♦ ${{m.nome}}</b><div class="tt-sep"></div>
      <b>Projeto:</b> ${{m.projeto}}<br>
      <b>Data:</b> ${{fmtDate(m.termino)}}<br>
      <b>Baseline:</b> ${{fmtDate(m.baseline)}}<br>
      <b>Status:</b> ${{status}}`, e);
  }});
  ms.addEventListener('mouseleave', hideTip);
  body.appendChild(ms);
}});
</script>
</body>
</html>
"""

    components.html(html_roadmap, height=total_h, scrolling=False)

st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 10. SECTION EVM — MATRIZ POR PROJETO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Section EVM — Análise de Valor Agregado por Projeto</div>',
            unsafe_allow_html=True)

df_evm = df_view[
    (df_view['nivel'] <= 2) &
    (~df_view['nome'].isin(['MS Project_Teste_Formulas_2', 'NOME DO PROJETO']))
].copy()

if not df_evm.empty:
    tbl_data = []
    for _, row in df_evm.iterrows():
        spi_v = row['spi_num']
        tbl_data.append({
            'Projeto':   row['projeto'],
            'Fase/Resumo': row['nome'][:55],
            'Nív': int(row['nivel']),
            '% Concl.': f"{row['pct']:.0f}%",
            'PV (dias)': f"{row['pv']:.0f}",
            'EV (dias)': f"{row['ev']:.0f}",
            'SPI': f"{spi_v:.2f}" if pd.notna(spi_v) else "N/A",
            'Alerta': ('🔴 CRÍTICO' if (pd.notna(spi_v) and spi_v < 0.80)
                       else ('⚠️ ATENÇÃO' if (pd.notna(spi_v) and spi_v < spi_limiar)
                             else '✅ OK')),
            'Término': row['termino'].strftime('%d/%m/%Y') if pd.notna(row['termino']) else '-',
        })
    st.dataframe(
        pd.DataFrame(tbl_data),
        use_container_width=True, hide_index=True,
        column_config={
            'SPI': st.column_config.TextColumn('SPI'),
            'Alerta': st.column_config.TextColumn('Alerta'),
        }
    )

st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 11. SECTION 2 — GOVERNANÇA DE INCERTEZAS (separada por projeto)
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="section-title">Section 2 — Governança de Incertezas: '
    'Pontos Críticos e Planos de Ação</div>',
    unsafe_allow_html=True,
)
st.caption(f"Exibindo tarefas/fases com SPI < {spi_limiar} · Separado por projeto")

df_criticos = df_view[
    (df_view['spi_num'] < spi_limiar) &
    df_view['spi_num'].notna()
].copy()

if df_criticos.empty:
    st.success(f"✅ Nenhuma tarefa com SPI abaixo de {spi_limiar} no portfólio atual.")
else:
    projetos_com_crits = df_criticos['projeto'].unique().tolist()

    for proj in projetos_com_crits:
        df_proj_crit = df_criticos[df_criticos['projeto'] == proj].copy()
        df_proj_crit = df_proj_crit.sort_values(['nivel', 'spi_num'])
        n_crits = len(df_proj_crit['nome'].unique())
        spi_min = df_proj_crit['spi_num'].min()
        alerta_proj = "🔴" if spi_min < 0.80 else "⚠️"

        label_expander = (
            f"{alerta_proj} {proj} "
            f"— {n_crits} ponto(s) crítico(s) "
            f"· SPI mín: {spi_min:.2f}"
        )

        with st.expander(label_expander, expanded=False):

            # ── IA ou estático ──────────────────────────────────────────
            items_ia = []
            if usar_ia:
                with st.spinner(f"🤖 Gerando insights para {proj}..."):
                    crit_payload = df_proj_crit[
                        ['nome', 'nivel', 'spi_num', 'pct', 'status', 'resp',
                         'termino', 'baseline_termino']
                    ].rename(columns={'spi_num': 'spi', 'termino': 'fim_atual',
                                       'baseline_termino': 'baseline'}).to_dict('records')
                    result = gerar_insights_ia(proj, json.dumps(crit_payload, default=str))
                    if result['status'] == 'ok':
                        items_ia = result['items']

            ia_map = {it.get('nome', '').strip(): it for it in items_ia}

            # ── Renderiza cards ─────────────────────────────────────────
            seen = set()
            for _, row in df_proj_crit.iterrows():
                key = row['nome'].strip()
                if key in seen:
                    continue
                seen.add(key)

                spi_v     = row['spi_num']
                row_class = "crit-row-red" if spi_v < 0.80 else "crit-row-yellow"
                spi_badge = badge(spi_v)

                ia_item = ia_map.get(key, {})
                impacto  = ia_item.get('impacto',  '— Análise IA não disponível; revise manualmente.')
                causa    = ia_item.get('causa_raiz','— Identificar gargalo com o responsável.')
                plano    = ia_item.get('plano_acao','— Agendar revisão semanal com owner.')

                fim_str  = row['termino'].strftime('%d/%m/%Y') if pd.notna(row['termino']) else '-'
                base_str = row['baseline_termino'].strftime('%d/%m/%Y') if pd.notna(row['baseline_termino']) else '-'

                st.markdown(f"""
<div class="{row_class}">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
    <span style="font-size:13px;font-weight:600;color:#1B2A4A;">
      {'&nbsp;' * (row['nivel']-1) * 4}L{row['nivel']} — {key[:65]}
    </span>
    <span>{spi_badge}&nbsp;&nbsp;
      <span style="font-size:11px;color:#6B7A99">{row['pct']:.0f}% concluído &nbsp;|&nbsp; Fim: {fim_str} &nbsp;|&nbsp; Baseline: {base_str}</span>
    </span>
  </div>
  <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;">
    <div>
      <div class="crit-label">📌 Impacto no Negócio</div>
      <div class="crit-value">{impacto}</div>
    </div>
    <div>
      <div class="crit-label">🔍 Causa Raiz (Hipótese)</div>
      <div class="crit-value">{causa}</div>
    </div>
    <div>
      <div class="crit-label">✅ Plano de Ação</div>
      <div class="crit-value">{plano}</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

            st.markdown("")


# ──────────────────────────────────────────────────────────────────────────────
# 12. EXPORTAR RELATÓRIO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)
st.markdown('<div class="section-title">Exportar Relatório Executivo</div>', unsafe_allow_html=True)

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

col_btn, _ = st.columns([1, 4])
with col_btn:
    if st.button("📥 Gerar Excel Executivo"):
        excel_bytes = gerar_excel(df_view, df_criticos)
        st.download_button(
            "⬇️ Baixar .xlsx",
            data=excel_bytes,
            file_name=f"relatorio_pmo_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.markdown(
    f"<p style='text-align:center;color:#C0C8D8;font-size:10px;margin-top:30px'>"
    f"Dashboard Executivo · PMBOK® 8ª Edição · PMI · Gerado em {date.today().strftime('%d/%m/%Y')}"
    f"</p>", unsafe_allow_html=True,
)
