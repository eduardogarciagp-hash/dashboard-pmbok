# =============================================================================
# DASHBOARD EXECUTIVO DE PORTFÓLIO — PMBOK 8ª EDIÇÃO
# Foco: Entrega de Valor | EVM | Roadmap | Gestão de Riscos
# Framework: Streamlit
# =============================================================================
# INSTALAÇÃO:
#   pip install streamlit pandas plotly openpyxl
# EXECUÇÃO:
#   streamlit run dashboard_pmbok8.py
# =============================================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from datetime import date, timedelta
import io

# ──────────────────────────────────────────────────────────────────────────────
# 0. CONFIGURAÇÃO DA PÁGINA
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Executivo | Portfólio",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stApp { background-color: #F4F6F9; }
    section[data-testid="stSidebar"] { background-color: #1B2A4A; }
    section[data-testid="stSidebar"] * { color: #FFFFFF !important; }
    .kpi-card {
        background: #FFFFFF; border-radius: 10px;
        padding: 20px 24px; border-left: 4px solid #1B2A4A;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .kpi-label { font-size: 12px; color: #6B7A99; font-weight: 600;
                 text-transform: uppercase; letter-spacing: .5px; }
    .kpi-value { font-size: 32px; font-weight: 700; color: #1B2A4A; line-height: 1.2; }
    .kpi-sub   { font-size: 12px; color: #9AA5BE; margin-top: 4px; }
    .badge-green  { background:#E6F4EA; color:#1E7E34; padding:3px 10px;
                    border-radius:12px; font-size:12px; font-weight:600; }
    .badge-yellow { background:#FFF8E1; color:#B8860B; padding:3px 10px;
                    border-radius:12px; font-size:12px; font-weight:600; }
    .badge-red    { background:#FDECEA; color:#C62828; padding:3px 10px;
                    border-radius:12px; font-size:12px; font-weight:600; }
    .section-title {
        font-size: 13px; font-weight: 700; color: #6B7A99;
        text-transform: uppercase; letter-spacing: 1px;
        margin: 28px 0 12px 0; border-bottom: 1px solid #E2E8F0; padding-bottom: 6px;
    }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 1. FUNÇÕES EVM — sem uso de .apply() para evitar KeyError
# ──────────────────────────────────────────────────────────────────────────────
def calcular_evm(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula SPI, SV e CPI usando operações vetoriais do numpy.
    Evita o erro KeyError que ocorre com pandas.apply() em certas versões.
    """
    for col in ["EV", "PV", "AC"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # SPI = EV / PV  (0 quando PV = 0)
    df["SPI"] = np.where(df["PV"] > 0, (df["EV"] / df["PV"]).round(3), np.nan)
    # SV  = EV - PV
    df["SV"]  = df["EV"] - df["PV"]
    # CPI = EV / AC  (0 quando AC = 0)
    df["CPI"] = np.where(df["AC"] > 0, (df["EV"] / df["AC"]).round(3), np.nan)

    def status_spi(s):
        if pd.isna(s):   return "N/A"
        if s >= 1.0:     return "🟢 No Prazo"
        if s >= 0.90:    return "🟡 Atenção"
        return "🔴 Crítico"

    df["Status_SPI"] = df["SPI"].apply(status_spi)
    return df


# ──────────────────────────────────────────────────────────────────────────────
# 2. MOCK DATA
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data
def carregar_dados_mock() -> pd.DataFrame:
    hoje = date.today()
    dados = [
        {"Portfolio":"Transformação Digital","Projeto":"ERP Cloud Migration",
         "ID":1,"Nome da Tarefa":"ERP Cloud Migration","Tipo":"Fase",
         "Inicio":hoje-timedelta(120),"Termino":hoje+timedelta(60),
         "Termino_Baseline":hoje+timedelta(50),"Pct_Concluida":0.62,
         "AC":850000,"PV":780000,"EV":700000,"Marco":False,
         "Recursos":"Ana Costa","Responsavel":"Ana Costa","Causa_Raiz":"","Plano_Acao":""},

        {"Portfolio":"Transformação Digital","Projeto":"ERP Cloud Migration",
         "ID":2,"Nome da Tarefa":"Migração de Dados","Tipo":"Tarefa",
         "Inicio":hoje-timedelta(60),"Termino":hoje+timedelta(10),
         "Termino_Baseline":hoje-timedelta(5),"Pct_Concluida":0.45,
         "AC":210000,"PV":190000,"EV":160000,"Marco":False,
         "Recursos":"João Lima","Responsavel":"João Lima","Causa_Raiz":"","Plano_Acao":""},

        {"Portfolio":"Transformação Digital","Projeto":"ERP Cloud Migration",
         "ID":3,"Nome da Tarefa":"Go-Live ERP","Tipo":"Marco",
         "Inicio":hoje+timedelta(60),"Termino":hoje+timedelta(60),
         "Termino_Baseline":hoje+timedelta(50),"Pct_Concluida":0.0,
         "AC":0,"PV":0,"EV":0,"Marco":True,
         "Recursos":"Ana Costa","Responsavel":"Ana Costa","Causa_Raiz":"","Plano_Acao":""},

        {"Portfolio":"Transformação Digital","Projeto":"Data Analytics Platform",
         "ID":4,"Nome da Tarefa":"Data Analytics Platform","Tipo":"Fase",
         "Inicio":hoje-timedelta(90),"Termino":hoje+timedelta(90),
         "Termino_Baseline":hoje+timedelta(90),"Pct_Concluida":0.50,
         "AC":420000,"PV":400000,"EV":410000,"Marco":False,
         "Recursos":"Maria Souza","Responsavel":"Maria Souza","Causa_Raiz":"","Plano_Acao":""},

        {"Portfolio":"Transformação Digital","Projeto":"Data Analytics Platform",
         "ID":5,"Nome da Tarefa":"Lançamento MVP Analytics","Tipo":"Marco",
         "Inicio":hoje+timedelta(30),"Termino":hoje+timedelta(30),
         "Termino_Baseline":hoje+timedelta(30),"Pct_Concluida":0.0,
         "AC":0,"PV":0,"EV":0,"Marco":True,
         "Recursos":"Maria Souza","Responsavel":"Maria Souza","Causa_Raiz":"","Plano_Acao":""},

        {"Portfolio":"Expansão de Mercado","Projeto":"Abertura Filial Sul",
         "ID":6,"Nome da Tarefa":"Abertura Filial Sul","Tipo":"Fase",
         "Inicio":hoje-timedelta(150),"Termino":hoje+timedelta(30),
         "Termino_Baseline":hoje-timedelta(10),"Pct_Concluida":0.80,
         "AC":1100000,"PV":960000,"EV":820000,"Marco":False,
         "Recursos":"Carlos Melo","Responsavel":"Carlos Melo","Causa_Raiz":"","Plano_Acao":""},

        {"Portfolio":"Expansão de Mercado","Projeto":"Abertura Filial Sul",
         "ID":7,"Nome da Tarefa":"Inauguração Filial Sul","Tipo":"Marco",
         "Inicio":hoje-timedelta(10),"Termino":hoje-timedelta(10),
         "Termino_Baseline":hoje-timedelta(30),"Pct_Concluida":0.0,
         "AC":0,"PV":0,"EV":0,"Marco":True,
         "Recursos":"Carlos Melo","Responsavel":"Carlos Melo","Causa_Raiz":"","Plano_Acao":""},

        {"Portfolio":"Expansão de Mercado","Projeto":"Expansão E-commerce",
         "ID":8,"Nome da Tarefa":"Expansão E-commerce","Tipo":"Fase",
         "Inicio":hoje-timedelta(45),"Termino":hoje+timedelta(120),
         "Termino_Baseline":hoje+timedelta(120),"Pct_Concluida":0.30,
         "AC":155000,"PV":160000,"EV":158000,"Marco":False,
         "Recursos":"Lucia Ferreira","Responsavel":"Lucia Ferreira","Causa_Raiz":"","Plano_Acao":""},

        {"Portfolio":"Expansão de Mercado","Projeto":"Expansão E-commerce",
         "ID":9,"Nome da Tarefa":"Lançamento Loja Online v2","Tipo":"Marco",
         "Inicio":hoje+timedelta(120),"Termino":hoje+timedelta(120),
         "Termino_Baseline":hoje+timedelta(120),"Pct_Concluida":0.0,
         "AC":0,"PV":0,"EV":0,"Marco":True,
         "Recursos":"Lucia Ferreira","Responsavel":"Lucia Ferreira","Causa_Raiz":"","Plano_Acao":""},
    ]
    df = pd.DataFrame(dados)
    df = calcular_evm(df)
    df["Marco_Atrasado"] = (df["Marco"] == True) & (df["Termino"] > df["Termino_Baseline"])
    df["Inicio_str"]           = pd.to_datetime(df["Inicio"]).dt.strftime("%Y-%m-%d")
    df["Termino_str"]          = pd.to_datetime(df["Termino"]).dt.strftime("%Y-%m-%d")
    df["Termino_Baseline_str"] = pd.to_datetime(df["Termino_Baseline"]).dt.strftime("%Y-%m-%d")
    return df


# ──────────────────────────────────────────────────────────────────────────────
# 3. CARREGAMENTO — Excel real ou mock
# ──────────────────────────────────────────────────────────────────────────────
def carregar_dados(arquivo_excel=None) -> pd.DataFrame:
    if arquivo_excel is None:
        return carregar_dados_mock()

    df_raw = pd.read_excel(arquivo_excel)

    # Normaliza nomes de colunas (remove espaços extras e padroniza)
    df_raw.columns = df_raw.columns.str.strip()

    # Mapeamento flexível de colunas do MS Project → nomes internos
    col_map = {
        "Início": "Inicio", "Inicio": "Inicio",
        "Término": "Termino", "Termino": "Termino",
        "% concluída": "Pct_Concluida", "% Concluída": "Pct_Concluida",
        "Custo Real (CR)": "AC", "AC": "AC", "Custo Real": "AC",
        "COTA": "PV", "PV": "PV",
        "COTR": "EV", "EV": "EV",
        "Nomes dos Recursos": "Recursos",
        "Nome da Tarefa": "Nome da Tarefa",
        "Marco": "Marco",
    }
    df_raw.rename(columns={k: v for k, v in col_map.items() if k in df_raw.columns}, inplace=True)

    # Garante colunas opcionais
    defaults = {
        "Portfolio": "", "Projeto": "", "Tipo": "Tarefa",
        "Termino_Baseline": df_raw.get("Termino", pd.Series(dtype="object")),
        "Responsavel": "", "Causa_Raiz": "", "Plano_Acao": "",
        "Pct_Concluida": 0.0,
    }
    for col, val in defaults.items():
        if col not in df_raw.columns:
            df_raw[col] = val

    # Converte % para decimal se vier como inteiro (ex: 75 → 0.75)
    df_raw["Pct_Concluida"] = pd.to_numeric(df_raw["Pct_Concluida"], errors="coerce").fillna(0)
    if df_raw["Pct_Concluida"].max() > 1:
        df_raw["Pct_Concluida"] = df_raw["Pct_Concluida"] / 100

    # Normaliza coluna Marco
    df_raw["Marco"] = df_raw["Marco"].astype(str).str.lower().isin(["sim", "true", "1", "yes"])

    # Calcula EVM com função vetorial segura
    df_raw = calcular_evm(df_raw)

    df_raw["Marco_Atrasado"] = (
        df_raw["Marco"] &
        (pd.to_datetime(df_raw["Termino"], errors="coerce") >
         pd.to_datetime(df_raw["Termino_Baseline"], errors="coerce"))
    )

    for campo, col in [("Inicio_str","Inicio"),("Termino_str","Termino"),
                       ("Termino_Baseline_str","Termino_Baseline")]:
        df_raw[campo] = pd.to_datetime(df_raw[col], errors="coerce").dt.strftime("%Y-%m-%d")

    return df_raw


# ──────────────────────────────────────────────────────────────────────────────
# 4. COMPONENTES UI
# ──────────────────────────────────────────────────────────────────────────────
def render_kpi_card(label, valor, sub="", cor="#1B2A4A"):
    st.markdown(f"""
    <div class="kpi-card" style="border-left-color:{cor}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{valor}</div>
        <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

def badge_spi(spi):
    if spi is None or (isinstance(spi, float) and np.isnan(spi)):
        return '<span class="badge-yellow">N/A</span>'
    if spi >= 1.0:
        return f'<span class="badge-green">SPI {spi:.2f} ▲</span>'
    elif spi >= 0.90:
        return f'<span class="badge-yellow">SPI {spi:.2f} !</span>'
    return f'<span class="badge-red">SPI {spi:.2f} ▼</span>'


# ──────────────────────────────────────────────────────────────────────────────
# 5. SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Dashboard PMO")
    st.markdown("---")
    st.markdown("### 📂 Fonte de Dados")
    arquivo = st.file_uploader(
        "Importe o Excel do MS Project",
        type=["xlsx", "xls"],
        help="Exporte pelo MS Project: Arquivo > Salvar Como > Excel Workbook",
    )
    st.markdown("---")
    st.markdown("### 🔍 Filtros")
    data_relatorio = st.date_input("Data de Referência", value=date.today())


# ──────────────────────────────────────────────────────────────────────────────
# 6. CARREGA DADOS E FILTROS
# ──────────────────────────────────────────────────────────────────────────────
df_full = carregar_dados(arquivo)

portfolios = ["Todos"] + sorted(df_full["Portfolio"].dropna().unique().tolist())
portfolio_sel = st.sidebar.selectbox("Portfólio", portfolios)

df_port = df_full if portfolio_sel == "Todos" else df_full[df_full["Portfolio"] == portfolio_sel]
projetos = ["Todos"] + sorted(df_port["Projeto"].dropna().unique().tolist())
projeto_sel = st.sidebar.selectbox("Projeto", projetos)

df = df_port if projeto_sel == "Todos" else df_port[df_port["Projeto"] == projeto_sel]
df_fases  = df[df["Tipo"].isin(["Fase", "Projeto"])]
df_marcos = df[df["Marco"] == True]


# ──────────────────────────────────────────────────────────────────────────────
# 7. CABEÇALHO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(
    f"<h1 style='color:#1B2A4A;font-size:26px;font-weight:700;margin-bottom:2px'>"
    f"Dashboard Executivo de Portfólio</h1>"
    f"<p style='color:#9AA5BE;font-size:13px'>PMBOK 8ª Ed. · Referência: "
    f"{data_relatorio.strftime('%d/%m/%Y')} · Portfólio: <b>{portfolio_sel}</b> · "
    f"Projeto: <b>{projeto_sel}</b></p>",
    unsafe_allow_html=True,
)
st.markdown("---")


# ──────────────────────────────────────────────────────────────────────────────
# 8. SEÇÃO A — KPIs GERAIS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">A. Governança Estratégica — Visão Holística</div>',
            unsafe_allow_html=True)

n_projetos  = df_fases["Projeto"].nunique()
pct_media   = df_fases["Pct_Concluida"].mean() * 100 if not df_fases.empty else 0
spi_medio   = df_fases["SPI"].dropna().mean() if not df_fases.empty else 0
marcos_crit = int(df_marcos["Marco_Atrasado"].sum()) if not df_marcos.empty else 0
n_criticos  = int((df_fases["SPI"].dropna() < 0.90).sum())

cor_spi = "#C62828" if spi_medio < 0.90 else "#B8860B" if spi_medio < 1.0 else "#1E7E34"

c1, c2, c3, c4, c5 = st.columns(5)
with c1: render_kpi_card("Projetos Monitorados", str(n_projetos), "no portfólio")
with c2: render_kpi_card("Conclusão Média", f"{pct_media:.1f}%", "avanço físico")
with c3: render_kpi_card("SPI Médio", f"{spi_medio:.2f}", "Índice de Desempenho de Prazo", cor_spi)
with c4: render_kpi_card("Marcos em Atraso", str(marcos_crit), "além da baseline",
                          "#C62828" if marcos_crit > 0 else "#1E7E34")
with c5: render_kpi_card("Projetos Críticos", str(n_criticos), "SPI < 0.90",
                          "#C62828" if n_criticos > 0 else "#1E7E34")


# ──────────────────────────────────────────────────────────────────────────────
# 9. SEÇÃO B — EVM
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">B. Indicadores EVM — Desempenho de Prazo e Custo</div>',
            unsafe_allow_html=True)

if df_fases.empty:
    st.info("Nenhum dado de fase disponível.")
else:
    df_evm = df_fases.groupby("Projeto", as_index=False).agg(
        PV=("PV","sum"), EV=("EV","sum"), AC=("AC","sum"),
        Pct=("Pct_Concluida","mean"),
    )
    df_evm["SPI"] = np.where(df_evm["PV"]>0, (df_evm["EV"]/df_evm["PV"]).round(3), np.nan)
    df_evm["CPI"] = np.where(df_evm["AC"]>0, (df_evm["EV"]/df_evm["AC"]).round(3), np.nan)
    df_evm["SV"]  = df_evm["EV"] - df_evm["PV"]

    cores = ["#C62828" if s<0.90 else "#B8860B" if s<1.0 else "#1E7E34"
             for s in df_evm["SPI"].fillna(0)]

    fig_bubble = go.Figure()
    fig_bubble.add_trace(go.Scatter(
        x=df_evm["SPI"], y=df_evm["CPI"],
        mode="markers+text",
        marker=dict(size=df_evm["EV"].clip(lower=1)/12000, color=cores,
                    opacity=0.85, line=dict(width=1, color="white")),
        text=df_evm["Projeto"], textposition="top center",
        textfont=dict(size=11, color="#1B2A4A"),
        hovertemplate="<b>%{text}</b><br>SPI: %{x:.2f}<br>CPI: %{y:.2f}<extra></extra>",
    ))
    fig_bubble.add_hline(y=1.0, line_dash="dot", line_color="#888", line_width=1)
    fig_bubble.add_vline(x=1.0, line_dash="dot", line_color="#888", line_width=1)
    fig_bubble.add_vline(x=0.90, line_dash="dash", line_color="#C62828", line_width=1,
                         annotation_text="Limite Crítico", annotation_position="top right",
                         annotation_font=dict(size=10, color="#C62828"))
    fig_bubble.update_layout(
        title=dict(text="Quadrante de Saúde — SPI × CPI", font=dict(size=13,color="#1B2A4A")),
        xaxis_title="SPI (Prazo)", yaxis_title="CPI (Custo)",
        plot_bgcolor="white", paper_bgcolor="white",
        height=360, margin=dict(l=40,r=20,t=50,b=40),
        xaxis=dict(gridcolor="#F0F0F0", zeroline=False),
        yaxis=dict(gridcolor="#F0F0F0", zeroline=False),
    )

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        st.plotly_chart(fig_bubble, use_container_width=True)
    with col_b2:
        st.markdown("**Resumo EVM por Projeto**")
        df_show = df_evm[["Projeto","PV","EV","AC","SPI","CPI","SV"]].copy()
        df_show.columns = ["Projeto","PV (R$)","EV (R$)","AC (R$)","SPI","CPI","SV (R$)"]
        for c in ["PV (R$)","EV (R$)","AC (R$)","SV (R$)"]:
            df_show[c] = df_show[c].apply(lambda x: f"R$ {x:,.0f}".replace(",","."))

        def hl(row):
            s = float(row["SPI"]) if not pd.isna(row["SPI"]) else 1.0
            if s < 0.90: return ["background-color:#FDECEA"]*len(row)
            if s < 1.0:  return ["background-color:#FFF8E1"]*len(row)
            return ["background-color:#E6F4EA"]*len(row)

        st.dataframe(df_show.style.apply(hl, axis=1), use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────────────────────────
# 10. SEÇÃO C — GANTT / ROADMAP
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">C. Roadmap Executivo — Cronograma e Marcos</div>',
            unsafe_allow_html=True)

df_gantt = df[df["Tipo"] == "Fase"].copy()

if df_gantt.empty:
    st.info("Nenhuma fase encontrada para o filtro selecionado.")
else:
    fig_gantt = go.Figure()
    hoje_str  = date.today().strftime("%Y-%m-%d")

    for _, row in df_gantt.iterrows():
        spi_val = row["SPI"] if not pd.isna(row["SPI"]) else 1.0
        cor = "#C62828" if spi_val < 0.90 else "#B8860B" if spi_val < 1.0 else "#1B2A4A"
        dur = (pd.to_datetime(row["Termino_str"]) - pd.to_datetime(row["Inicio_str"])).days
        fig_gantt.add_trace(go.Bar(
            x=[dur], y=[f"{row['Projeto']} — {row['Nome da Tarefa']}"],
            base=[row["Inicio_str"]], orientation="h",
            marker=dict(color=cor, opacity=0.85),
            text=f"{row['Pct_Concluida']*100:.0f}%",
            textposition="inside", insidetextanchor="middle",
            hovertemplate=(f"<b>{row['Nome da Tarefa']}</b><br>"
                           f"Início: {row['Inicio_str']}<br>Término: {row['Termino_str']}<br>"
                           f"SPI: {f'{spi_val:.2f}'}<br>"
                           f"Avanço: {row['Pct_Concluida']*100:.0f}%<extra></extra>"),
            showlegend=False,
        ))

    for _, row in df_marcos.iterrows():
        y_ref = df_gantt[df_gantt["Projeto"] == row["Projeto"]]["Nome da Tarefa"].values
        y_lbl = f"{row['Projeto']} — {y_ref[0]}" if len(y_ref) > 0 else row["Projeto"]
        cor_m = "#C62828" if row["Marco_Atrasado"] else "#1E7E34"
        fig_gantt.add_trace(go.Scatter(
            x=[row["Termino_str"]], y=[y_lbl],
            mode="markers+text",
            marker=dict(symbol="diamond", size=14, color=cor_m,
                        line=dict(width=1.5, color="white")),
            text=[f"◆ {row['Nome da Tarefa']}"], textposition="top center",
            textfont=dict(size=10, color=cor_m),
            hovertemplate=(f"<b>Marco: {row['Nome da Tarefa']}</b><br>"
                           f"Baseline: {row['Termino_Baseline_str']}<br>"
                           f"Atual: {row['Termino_str']}<br>"
                           f"{'⚠️ Atrasado' if row['Marco_Atrasado'] else '✅ No prazo'}<extra></extra>"),
            showlegend=False,
        ))

    # Linha "Hoje" como shape manual (compatível com eixo de datas no Plotly recente)
    fig_gantt.add_shape(
        type="line", xref="x", yref="paper",
        x0=hoje_str, x1=hoje_str, y0=0, y1=1,
        line=dict(color="#4A90D9", width=2, dash="dash"),
    )
    fig_gantt.add_annotation(
        x=hoje_str, y=1, yref="paper", text="Hoje",
        showarrow=False, xanchor="left",
        font=dict(size=11, color="#4A90D9"),
    )
    fig_gantt.update_layout(
        barmode="overlay", plot_bgcolor="white", paper_bgcolor="white",
        height=max(300, len(df_gantt)*60+80),
        margin=dict(l=10,r=20,t=20,b=40),
        xaxis=dict(type="date", gridcolor="#F0F0F0", tickformat="%b/%y"),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig_gantt, use_container_width=True)
    st.markdown("<small style='color:#9AA5BE'>🔵 No prazo &nbsp;|&nbsp; 🟡 Atenção &nbsp;|&nbsp; "
                "🔴 Crítico &nbsp;|&nbsp; ◆ Verde = Marco OK &nbsp;|&nbsp; ◆ Vermelho = Marco atrasado</small>",
                unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 11. SEÇÃO D — PONTOS CRÍTICOS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">D. Matriz de Pontos Críticos — Planos de Ação</div>',
            unsafe_allow_html=True)

df_crit_f = df_fases[df_fases["SPI"].notna() & (df_fases["SPI"] < 0.90)].copy()
df_crit_m = df_marcos[df_marcos["Marco_Atrasado"] == True].copy()
df_criticos = pd.concat([df_crit_f, df_crit_m]).drop_duplicates(subset=["Projeto"])

if df_criticos.empty:
    st.success("✅ Nenhum desvio crítico identificado no portfólio selecionado.")
else:
    st.warning(f"⚠️ **{len(df_criticos)} projeto(s) com desvio crítico.** "
               "Preencha os campos antes da reunião de diretoria.")

    if "acoes" not in st.session_state:
        st.session_state["acoes"] = {}

    for _, row in df_criticos.iterrows():
        pk = row["Projeto"]
        if pk not in st.session_state["acoes"]:
            st.session_state["acoes"][pk] = {
                "impacto":"", "causa":"", "plano":"",
                "resp": row.get("Responsavel",""),
            }

        spi_lbl = f"{row['SPI']:.2f}" if not pd.isna(row["SPI"]) else "N/A"
        with st.expander(f"🔴 {pk} — SPI: {spi_lbl}", expanded=True):
            ca, cb = st.columns(2)
            with ca:
                st.markdown(f"**Status:** {badge_spi(row['SPI'])}", unsafe_allow_html=True)
                st.markdown(f"**Avanço Físico:** {row['Pct_Concluida']*100:.0f}%")
                st.markdown(f"**SV:** R$ {row['SV']:,.0f}".replace(",","."))
            with cb:
                st.markdown(f"**Término Previsto:** {row['Termino_str']}")
                st.markdown(f"**Baseline:** {row['Termino_Baseline_str']}")
                desvio = (pd.to_datetime(row["Termino_str"]) -
                          pd.to_datetime(row["Termino_Baseline_str"])).days
                st.markdown(f"**Desvio:** {'🔴 +'+str(desvio)+' dias' if desvio>0 else '🟢 '+str(desvio)+' dias'}")

            st.markdown("---")
            c1,c2,c3,c4 = st.columns(4)
            with c1:
                st.session_state["acoes"][pk]["impacto"] = st.text_area(
                    "Impacto no Negócio", value=st.session_state["acoes"][pk]["impacto"],
                    height=90, placeholder="Ex: Atraso impacta receita de R$ 500k", key=f"imp_{pk}")
            with c2:
                st.session_state["acoes"][pk]["causa"] = st.text_area(
                    "Causa Raiz", value=st.session_state["acoes"][pk]["causa"],
                    height=90, placeholder="Ex: Volume de dados 3x maior que estimado", key=f"cau_{pk}")
            with c3:
                st.session_state["acoes"][pk]["plano"] = st.text_area(
                    "Plano de Ação", value=st.session_state["acoes"][pk]["plano"],
                    height=90, placeholder="Ex: +2 DBAs por 3 semanas", key=f"pla_{pk}")
            with c4:
                st.session_state["acoes"][pk]["resp"] = st.text_input(
                    "Responsável", value=st.session_state["acoes"][pk]["resp"], key=f"res_{pk}")


# ──────────────────────────────────────────────────────────────────────────────
# 12. EXPORTAÇÃO EXCEL
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Exportar Relatório Executivo</div>', unsafe_allow_html=True)

def gerar_excel(df_base, acoes):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_exp = df_base.groupby("Projeto", as_index=False).agg(
            Portfolio=("Portfolio","first"),
            PV=("PV","sum"), EV=("EV","sum"), AC=("AC","sum"),
            Pct=("Pct_Concluida","mean"),
        )
        df_exp["SPI"] = np.where(df_exp["PV"]>0,(df_exp["EV"]/df_exp["PV"]).round(3),np.nan)
        df_exp["CPI"] = np.where(df_exp["AC"]>0,(df_exp["EV"]/df_exp["AC"]).round(3),np.nan)
        df_exp.to_excel(writer, sheet_name="EVM_Consolidado", index=False)

        rows = [{"Projeto":p,"Impacto":d["impacto"],"Causa Raiz":d["causa"],
                 "Plano de Ação":d["plano"],"Responsável":d["resp"]}
                for p,d in acoes.items()]
        if rows:
            pd.DataFrame(rows).to_excel(writer, sheet_name="Planos_de_Acao", index=False)

        df_m = df[df["Marco"]==True][["Portfolio","Projeto","Nome da Tarefa",
                                       "Termino_str","Termino_Baseline_str","Marco_Atrasado"]].copy()
        df_m.columns=["Portfólio","Projeto","Marco","Término","Baseline","Atrasado"]
        df_m.to_excel(writer, sheet_name="Marcos", index=False)
    return output.getvalue()

col_e1, _ = st.columns([1,4])
with col_e1:
    if st.button("📥 Gerar Excel Executivo", use_container_width=True):
        excel_bytes = gerar_excel(df, st.session_state.get("acoes", {}))
        st.download_button(
            "⬇️ Baixar Relatório (.xlsx)", data=excel_bytes,
            file_name=f"relatorio_executivo_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.markdown("---")
st.markdown(
    f"<p style='text-align:center;color:#C0C8D8;font-size:11px'>"
    f"Dashboard Executivo · PMBOK® 8ª Edição · PMI · {date.today().strftime('%d/%m/%Y')}</p>",
    unsafe_allow_html=True,
)
