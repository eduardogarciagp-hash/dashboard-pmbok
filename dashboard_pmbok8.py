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

        # SPI: prefer Texto3 (Cockpit), fallback to Número6/Número4
        spi_direct = None
        raw = (cf.get('Texto3', '') or '').replace(',', '.').strip()
        try:
            spi_direct = float(raw)
        except:
            pass
        spi = spi_direct if spi_direct is not None else (
            round(ev_val / pv, 3) if pv > 0 else None
        )

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
    if spi >= 0.95:
        return f'<span class="badge-green">SPI {spi:.2f} ▲</span>'
    if spi >= 0.80:
        return f'<span class="badge-yellow">SPI {spi:.2f} !</span>'
    return f'<span class="badge-red">SPI {spi:.2f} ▼</span>'

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
    f"Dashboard de Governança e Valor do Portfólio</h1>"
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
n_projetos  = df_view['projeto'].nunique()
pct_media   = df_root['pct'].mean() if not df_root.empty else 0
spi_medio   = df_root['spi_num'].dropna().mean() if not df_root.empty else None
crits_count = df_view[(df_view['spi_num'] < spi_limiar) & df_view['spi_num'].notna()]['projeto'].nunique()
marcos_tot  = df_view[df_view['is_milestone']].shape[0]

c1, c2, c3, c4, c5 = st.columns(5)
with c1: kpi_card("PROJETOS ATIVOS", str(n_projetos), "monitorados", "#2563EB")
with c2: kpi_card("CONCLUSÃO MÉDIA", f"{pct_media:.1f}%", "do portfólio", "#059669")
with c3:
    spi_str = f"{spi_medio:.3f}" if spi_medio else "N/A"
    cor3 = "#DC2626" if (spi_medio and spi_medio < 0.90) else ("#F59E0B" if (spi_medio and spi_medio < 0.95) else "#059669")
    kpi_card("SPI PORTFÓLIO", spi_str, "índice de desempenho", cor3)
with c4: kpi_card("PROJETOS CRÍTICOS", str(crits_count), f"com SPI < {spi_limiar}", "#DC2626")
with c5: kpi_card("MARCOS NO PORTFÓLIO", str(marcos_tot), "identificados", "#7C3AED")

st.markdown("<hr class='section-sep'>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 9. SECTION 1 — ROADMAP (GANTT)
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Section 1 — Roadmap Executivo & Marcos de Valor</div>',
            unsafe_allow_html=True)

df_gantt = df_view[
    (df_view['nivel'].isin([1, 2])) &
    (df_view['inicio'].notna()) &
    (df_view['termino'].notna()) &
    (~df_view['nome'].isin(['MS Project_Teste_Formulas_2', 'NOME DO PROJETO']))
].copy()

if not df_gantt.empty:
    df_gantt['label'] = df_gantt.apply(
        lambda r: f"[{r['projeto']}] {r['nome'][:40]}", axis=1)
    df_gantt['cor'] = df_gantt['spi_num'].apply(
        lambda s: '#DC2626' if (s and s < 0.80) else
                  ('#F59E0B' if (s and s < 0.95) else '#2563EB'))
    df_gantt['spi_txt'] = df_gantt['spi_num'].apply(
        lambda s: f"SPI {s:.2f}" if pd.notna(s) else "")

    fig_gantt = go.Figure()
    for _, row in df_gantt.iterrows():
        fig_gantt.add_trace(go.Bar(
            x=[(row['termino'] - row['inicio']).days],
            y=[row['label']],
            base=[row['inicio'].strftime('%Y-%m-%d')],
            orientation='h',
            marker_color=row['cor'],
            marker_opacity=0.80 if row['nivel'] == 2 else 0.50,
            marker_line_width=0,
            width=0.6 if row['nivel'] == 2 else 0.3,
            text=f"{row['pct']:.0f}% {row['spi_txt']}",
            textposition='inside',
            insidetextanchor='middle',
            hovertemplate=(
                f"<b>{row['nome']}</b><br>"
                f"Projeto: {row['projeto']}<br>"
                f"Início: {row['inicio'].strftime('%d/%m/%Y')}<br>"
                f"Término: {row['termino'].strftime('%d/%m/%Y')}<br>"
                f"Conclusão: {row['pct']:.0f}%<br>"
                f"SPI: {row['spi_txt']}<extra></extra>"
            ),
            showlegend=False,
        ))

    # Marcos
    df_marcos = df_view[df_view['is_milestone'] & df_view['termino'].notna()]
    for _, m in df_marcos.iterrows():
        label = f"[{m['projeto']}] {m['nome'][:40]}"
        fig_gantt.add_trace(go.Scatter(
            x=[m['termino']], y=[label],
            mode='markers+text',
            marker=dict(symbol='diamond', size=10, color='#7C3AED'),
            text=['♦'], textposition='middle right',
            hovertemplate=f"<b>Marco:</b> {m['nome']}<extra></extra>",
            showlegend=False,
        ))

    fig_gantt.add_vline(
        x=date.today().strftime('%Y-%m-%d'),
        line_dash="dash", line_color="#6B7A99", line_width=1.5,
        annotation_text=f"Hoje {date.today().strftime('%d/%m')}",
        annotation_position="top right",
    )
    fig_gantt.update_layout(
        barmode='overlay', height=max(280, len(df_gantt) * 32 + 60),
        plot_bgcolor='#FAFBFE', paper_bgcolor='#FAFBFE',
        margin=dict(l=10, r=10, t=10, b=30),
        xaxis=dict(type='date', tickformat='%b/%y', gridcolor='#E8ECF4',
                   showgrid=True, tickfont=dict(size=10)),
        yaxis=dict(autorange='reversed', tickfont=dict(size=10), gridcolor='#E8ECF4'),
        font=dict(family='Inter'),
    )
    st.plotly_chart(fig_gantt, use_container_width=True)

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
        css_cls = PROJ_CSS.get(proj, "proj-header-bdf")
        st.markdown(f'<div class="{css_cls}">📌 {proj}</div>', unsafe_allow_html=True)

        df_proj_crit = df_criticos[df_criticos['projeto'] == proj].copy()
        df_proj_crit = df_proj_crit.sort_values(['nivel', 'spi_num'])

        # ── IA ou estático ──────────────────────────────────────────────────
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

        # ── Renderiza cards ─────────────────────────────────────────────────
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
