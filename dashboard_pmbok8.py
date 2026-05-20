# =============================================================================
# DASHBOARD EXECUTIVO DE PORTFÓLIO — PMBOK 8ª EDIÇÃO
# Foco: Entrega de Valor | EVM | Roadmap | Gestão de Riscos
# Framework: Streamlit
# Autor: Gerado via Claude (Anthropic) — Especialista PMP + Ciência de Dados
# =============================================================================
# INSTALAÇÃO:
#   pip install streamlit pandas plotly openpyxl
# EXECUÇÃO:
#   streamlit run dashboard_pmbok8.py
# =============================================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import date, timedelta
import numpy as np

# ──────────────────────────────────────────────────────────────────────────────
# 0. CONFIGURAÇÃO GLOBAL DA PÁGINA
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Executivo | Portfólio",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Paleta corporativa via CSS inline ────────────────────────────────────────
st.markdown("""
<style>
    /* Fundo geral e sidebar */
    .stApp { background-color: #F4F6F9; }
    section[data-testid="stSidebar"] { background-color: #1B2A4A; }
    section[data-testid="stSidebar"] * { color: #FFFFFF !important; }

    /* Cards de KPI */
    .kpi-card {
        background: #FFFFFF;
        border-radius: 10px;
        padding: 20px 24px;
        border-left: 4px solid #1B2A4A;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .kpi-label { font-size: 12px; color: #6B7A99; font-weight: 600;
                 text-transform: uppercase; letter-spacing: .5px; }
    .kpi-value { font-size: 32px; font-weight: 700; color: #1B2A4A; line-height: 1.2; }
    .kpi-sub   { font-size: 12px; color: #9AA5BE; margin-top: 4px; }

    /* Badges de status */
    .badge-green  { background:#E6F4EA; color:#1E7E34; padding:3px 10px;
                    border-radius:12px; font-size:12px; font-weight:600; }
    .badge-yellow { background:#FFF8E1; color:#B8860B; padding:3px 10px;
                    border-radius:12px; font-size:12px; font-weight:600; }
    .badge-red    { background:#FDECEA; color:#C62828; padding:3px 10px;
                    border-radius:12px; font-size:12px; font-weight:600; }

    /* Títulos de seção */
    .section-title {
        font-size: 13px; font-weight: 700; color: #6B7A99;
        text-transform: uppercase; letter-spacing: 1px;
        margin: 28px 0 12px 0; border-bottom: 1px solid #E2E8F0; padding-bottom: 6px;
    }

    /* Tabela de pontos críticos */
    .stDataFrame { border-radius: 10px !important; }
    div[data-testid="stExpander"] { background: #FFFFFF; border-radius: 10px; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 1. MOCK DATA — Estrutura fiel ao export padrão do MS Project
#    Substitua esta seção por: df = pd.read_excel("seu_arquivo.xlsx")
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data
def carregar_dados_mock() -> pd.DataFrame:
    """
    Simula o DataFrame exportado do MS Project.
    Colunas espelham os campos nativos do MS Project + EVM calculado.
    """
    hoje = date.today()

    dados = [
        # ── PORTFÓLIO A — Transformação Digital ──────────────────────────────
        {
            "Portfolio": "Transformação Digital",
            "Projeto": "ERP Cloud Migration",
            "ID": 1, "Nome da Tarefa": "ERP Cloud Migration",
            "Tipo": "Fase",  # Fase | Tarefa | Marco
            "Inicio": hoje - timedelta(days=120),
            "Termino": hoje + timedelta(days=60),
            "Termino_Baseline": hoje + timedelta(days=50),
            "Pct_Concluida": 0.62,
            "AC": 850_000,   # Custo Real
            "PV": 780_000,   # COTA — Custo Orçado Agendado
            "EV": 700_000,   # COTR — Custo Orçado Realizado
            "Marco": False,
            "Recursos": "Ana Costa, João Lima",
            "Responsavel": "Ana Costa",
            "Causa_Raiz": "",
            "Plano_Acao": "",
        },
        {
            "Portfolio": "Transformação Digital",
            "Projeto": "ERP Cloud Migration",
            "ID": 2, "Nome da Tarefa": "Migração de Dados",
            "Tipo": "Tarefa",
            "Inicio": hoje - timedelta(days=60),
            "Termino": hoje + timedelta(days=10),
            "Termino_Baseline": hoje - timedelta(days=5),
            "Pct_Concluida": 0.45,
            "AC": 210_000, "PV": 190_000, "EV": 160_000,
            "Marco": False,
            "Recursos": "João Lima",
            "Responsavel": "João Lima",
            "Causa_Raiz": "",
            "Plano_Acao": "",
        },
        {
            "Portfolio": "Transformação Digital",
            "Projeto": "ERP Cloud Migration",
            "ID": 3, "Nome da Tarefa": "Go-Live ERP",
            "Tipo": "Marco",
            "Inicio": hoje + timedelta(days=60),
            "Termino": hoje + timedelta(days=60),
            "Termino_Baseline": hoje + timedelta(days=50),
            "Pct_Concluida": 0.0,
            "AC": 0, "PV": 0, "EV": 0,
            "Marco": True,
            "Recursos": "Ana Costa",
            "Responsavel": "Ana Costa",
            "Causa_Raiz": "",
            "Plano_Acao": "",
        },
        # ── PORTFÓLIO A — Data Analytics ─────────────────────────────────────
        {
            "Portfolio": "Transformação Digital",
            "Projeto": "Data Analytics Platform",
            "ID": 4, "Nome da Tarefa": "Data Analytics Platform",
            "Tipo": "Fase",
            "Inicio": hoje - timedelta(days=90),
            "Termino": hoje + timedelta(days=90),
            "Termino_Baseline": hoje + timedelta(days=90),
            "Pct_Concluida": 0.50,
            "AC": 420_000, "PV": 400_000, "EV": 410_000,
            "Marco": False,
            "Recursos": "Maria Souza",
            "Responsavel": "Maria Souza",
            "Causa_Raiz": "",
            "Plano_Acao": "",
        },
        {
            "Portfolio": "Transformação Digital",
            "Projeto": "Data Analytics Platform",
            "ID": 5, "Nome da Tarefa": "Lançamento MVP Analytics",
            "Tipo": "Marco",
            "Inicio": hoje + timedelta(days=30),
            "Termino": hoje + timedelta(days=30),
            "Termino_Baseline": hoje + timedelta(days=30),
            "Pct_Concluida": 0.0,
            "AC": 0, "PV": 0, "EV": 0,
            "Marco": True,
            "Recursos": "Maria Souza",
            "Responsavel": "Maria Souza",
            "Causa_Raiz": "",
            "Plano_Acao": "",
        },
        # ── PORTFÓLIO B — Expansão de Mercado ────────────────────────────────
        {
            "Portfolio": "Expansão de Mercado",
            "Projeto": "Abertura Filial Sul",
            "ID": 6, "Nome da Tarefa": "Abertura Filial Sul",
            "Tipo": "Fase",
            "Inicio": hoje - timedelta(days=150),
            "Termino": hoje + timedelta(days=30),
            "Termino_Baseline": hoje - timedelta(days=10),
            "Pct_Concluida": 0.80,
            "AC": 1_100_000, "PV": 960_000, "EV": 820_000,
            "Marco": False,
            "Recursos": "Carlos Melo",
            "Responsavel": "Carlos Melo",
            "Causa_Raiz": "",
            "Plano_Acao": "",
        },
        {
            "Portfolio": "Expansão de Mercado",
            "Projeto": "Abertura Filial Sul",
            "ID": 7, "Nome da Tarefa": "Inauguração Filial Sul",
            "Tipo": "Marco",
            "Inicio": hoje - timedelta(days=10),
            "Termino": hoje - timedelta(days=10),
            "Termino_Baseline": hoje - timedelta(days=30),
            "Pct_Concluida": 0.0,
            "AC": 0, "PV": 0, "EV": 0,
            "Marco": True,
            "Recursos": "Carlos Melo",
            "Responsavel": "Carlos Melo",
            "Causa_Raiz": "",
            "Plano_Acao": "",
        },
        {
            "Portfolio": "Expansão de Mercado",
            "Projeto": "Expansão E-commerce",
            "ID": 8, "Nome da Tarefa": "Expansão E-commerce",
            "Tipo": "Fase",
            "Inicio": hoje - timedelta(days=45),
            "Termino": hoje + timedelta(days=120),
            "Termino_Baseline": hoje + timedelta(days=120),
            "Pct_Concluida": 0.30,
            "AC": 155_000, "PV": 160_000, "EV": 158_000,
            "Marco": False,
            "Recursos": "Lucia Ferreira",
            "Responsavel": "Lucia Ferreira",
            "Causa_Raiz": "",
            "Plano_Acao": "",
        },
        {
            "Portfolio": "Expansão de Mercado",
            "Projeto": "Expansão E-commerce",
            "ID": 9, "Nome da Tarefa": "Lançamento Loja Online v2",
            "Tipo": "Marco",
            "Inicio": hoje + timedelta(days=120),
            "Termino": hoje + timedelta(days=120),
            "Termino_Baseline": hoje + timedelta(days=120),
            "Pct_Concluida": 0.0,
            "AC": 0, "PV": 0, "EV": 0,
            "Marco": True,
            "Recursos": "Lucia Ferreira",
            "Responsavel": "Lucia Ferreira",
            "Causa_Raiz": "",
            "Plano_Acao": "",
        },
    ]

    df = pd.DataFrame(dados)

    # ── Cálculos EVM derivados ────────────────────────────────────────────────
    # SPI (Índice de Desempenho de Prazo) = EV / PV
    df["SPI"] = df.apply(
        lambda r: round(r["EV"] / r["PV"], 3) if r["PV"] > 0 else None, axis=1
    )
    # SV (Variância de Prazo) = EV - PV
    df["SV"] = df["EV"] - df["PV"]

    # CPI (Índice de Desempenho de Custo) = EV / AC
    df["CPI"] = df.apply(
        lambda r: round(r["EV"] / r["AC"], 3) if r["AC"] > 0 else None, axis=1
    )

    # Status semafórico baseado no SPI
    def status_spi(spi):
        if spi is None:
            return "N/A"
        if spi >= 1.0:
            return "🟢 No Prazo"
        elif spi >= 0.90:
            return "🟡 Atenção"
        else:
            return "🔴 Crítico"

    df["Status_SPI"] = df["SPI"].apply(status_spi)

    # Marco atrasado: Termino real > Termino_Baseline E Pct_Concluida < 1.0
    df["Marco_Atrasado"] = (
        (df["Marco"] == True) &
        (df["Termino"] > df["Termino_Baseline"])
    )

    # Converte datas para string (compatível com Plotly)
    df["Inicio_str"]            = pd.to_datetime(df["Inicio"]).dt.strftime("%Y-%m-%d")
    df["Termino_str"]           = pd.to_datetime(df["Termino"]).dt.strftime("%Y-%m-%d")
    df["Termino_Baseline_str"]  = pd.to_datetime(df["Termino_Baseline"]).dt.strftime("%Y-%m-%d")

    return df


# ──────────────────────────────────────────────────────────────────────────────
# 2. CARREGAMENTO DE DADOS — Real ou Mock
# ──────────────────────────────────────────────────────────────────────────────
def carregar_dados(arquivo_excel=None) -> pd.DataFrame:
    """
    Se um arquivo Excel for enviado (via st.file_uploader),
    lê e padroniza as colunas. Caso contrário, retorna mock data.

    Mapeamento esperado de colunas do MS Project:
        'ID', 'Nome da Tarefa', 'Modo da Tarefa', 'Duração',
        'Início', 'Término', 'Precedentes', 'Nomes dos Recursos',
        '% concluída', 'Custo Real (CR)', 'COTA', 'COTR', 'Marco'
    """
    if arquivo_excel is not None:
        df_raw = pd.read_excel(arquivo_excel)

        # ── Normalização de colunas (adapte ao seu export) ──────────────────
        col_map = {
            "Início": "Inicio",
            "Término": "Termino",
            "% concluída": "Pct_Concluida",
            "Custo Real (CR)": "AC",
            "COTA": "PV",
            "COTR": "EV",
            "Nomes dos Recursos": "Recursos",
            "Nome da Tarefa": "Nome da Tarefa",
            "Marco": "Marco",
        }
        df_raw.rename(columns=col_map, inplace=True)

        # Campos adicionais que não vêm do MS Project nativamente
        for col in ["Portfolio", "Projeto", "Tipo", "Termino_Baseline",
                    "Responsavel", "Causa_Raiz", "Plano_Acao"]:
            if col not in df_raw.columns:
                df_raw[col] = ""

        # Recalcula EVM
        df_raw["SPI"] = df_raw.apply(
            lambda r: round(r["EV"] / r["PV"], 3) if r["PV"] > 0 else None, axis=1
        )
        df_raw["SV"]  = df_raw["EV"] - df_raw["PV"]
        df_raw["CPI"] = df_raw.apply(
            lambda r: round(r["EV"] / r["AC"], 3) if r["AC"] > 0 else None, axis=1
        )

        def status_spi(spi):
            if spi is None: return "N/A"
            if spi >= 1.0:  return "🟢 No Prazo"
            if spi >= 0.90: return "🟡 Atenção"
            return "🔴 Crítico"

        df_raw["Status_SPI"]     = df_raw["SPI"].apply(status_spi)
        df_raw["Marco_Atrasado"] = (
            (df_raw["Marco"].astype(str).str.lower().isin(["sim", "true", "1"])) &
            (pd.to_datetime(df_raw["Termino"]) > pd.to_datetime(df_raw["Termino_Baseline"]))
        )
        df_raw["Inicio_str"]           = pd.to_datetime(df_raw["Inicio"]).dt.strftime("%Y-%m-%d")
        df_raw["Termino_str"]          = pd.to_datetime(df_raw["Termino"]).dt.strftime("%Y-%m-%d")
        df_raw["Termino_Baseline_str"] = pd.to_datetime(df_raw["Termino_Baseline"]).dt.strftime("%Y-%m-%d")
        return df_raw

    return carregar_dados_mock()


# ──────────────────────────────────────────────────────────────────────────────
# 3. COMPONENTES DE UI
# ──────────────────────────────────────────────────────────────────────────────
def render_kpi_card(label: str, valor: str, sub: str = "", cor_borda: str = "#1B2A4A"):
    """Renderiza um card KPI com estilo corporativo."""
    st.markdown(f"""
    <div class="kpi-card" style="border-left-color:{cor_borda}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{valor}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


def badge_spi(spi):
    """Retorna HTML de badge colorido conforme SPI."""
    if spi is None:
        return '<span class="badge-yellow">N/A</span>'
    if spi >= 1.0:
        return f'<span class="badge-green">SPI {spi:.2f} ▲</span>'
    elif spi >= 0.90:
        return f'<span class="badge-yellow">SPI {spi:.2f} !</span>'
    else:
        return f'<span class="badge-red">SPI {spi:.2f} ▼</span>'


# ──────────────────────────────────────────────────────────────────────────────
# 4. SIDEBAR — Filtros e Upload
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image(
        "https://via.placeholder.com/180x40/1B2A4A/FFFFFF?text=PORTFÓLIO+PMO",
        use_container_width=True,
    )
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
# 5. CARREGA E FILTRA DADOS
# ──────────────────────────────────────────────────────────────────────────────
df_full = carregar_dados(arquivo)

portfolios_disponiveis = ["Todos"] + sorted(df_full["Portfolio"].unique().tolist())
portfolio_sel = st.sidebar.selectbox("Portfólio", portfolios_disponiveis)

projetos_filtrados = df_full if portfolio_sel == "Todos" else df_full[df_full["Portfolio"] == portfolio_sel]
projetos_disponiveis = ["Todos"] + sorted(projetos_filtrados["Projeto"].unique().tolist())
projeto_sel = st.sidebar.selectbox("Projeto", projetos_disponiveis)

df = projetos_filtrados if projeto_sel == "Todos" else projetos_filtrados[projetos_filtrados["Projeto"] == projeto_sel]

# Apenas linhas de Fase/Projeto (exclui tarefas granulares dos KPIs gerais)
df_fases = df[df["Tipo"].isin(["Fase", "Projeto"])]
df_marcos = df[df["Marco"] == True]


# ──────────────────────────────────────────────────────────────────────────────
# 6. CABEÇALHO
# ──────────────────────────────────────────────────────────────────────────────
col_logo, col_titulo = st.columns([1, 5])
with col_titulo:
    st.markdown(
        f"<h1 style='color:#1B2A4A;font-size:26px;font-weight:700;margin-bottom:2px'>"
        f"Dashboard Executivo de Portfólio</h1>"
        f"<p style='color:#9AA5BE;font-size:13px'>PMBOK 8ª Ed. — Entrega de Valor | "
        f"Referência: {data_relatorio.strftime('%d/%m/%Y')} | "
        f"Portfólio: <b>{portfolio_sel}</b> | Projeto: <b>{projeto_sel}</b></p>",
        unsafe_allow_html=True,
    )

st.markdown("---")


# ──────────────────────────────────────────────────────────────────────────────
# 7. SEÇÃO A — KPIs GERAIS DO PORTFÓLIO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">A. Governança Estratégica — Visão Holística do Portfólio</div>',
            unsafe_allow_html=True)

n_projetos   = df_fases["Projeto"].nunique()
pct_media    = df_fases["Pct_Concluida"].mean() * 100 if not df_fases.empty else 0
spi_medio    = df_fases["SPI"].dropna().mean() if not df_fases.empty else 0
marcos_crit  = df_marcos["Marco_Atrasado"].sum() if not df_marcos.empty else 0
n_criticos   = (df_fases["SPI"].dropna() < 0.90).sum()

# Saúde geral: média ponderada de SPI vs meta 1.0
saude_pct = min(spi_medio * 100, 100) if spi_medio else 0
if saude_pct >= 95:
    cor_saude, label_saude = "#1E7E34", "SAUDÁVEL"
elif saude_pct >= 90:
    cor_saude, label_saude = "#B8860B", "ATENÇÃO"
else:
    cor_saude, label_saude = "#C62828", "CRÍTICO"

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    render_kpi_card("Projetos Monitorados", str(n_projetos), "no portfólio selecionado")
with c2:
    render_kpi_card("Conclusão Média", f"{pct_media:.1f}%", "% médio de avanço físico")
with c3:
    render_kpi_card("SPI Médio do Portfólio", f"{spi_medio:.2f}",
                    "Índice de Desempenho de Prazo",
                    cor_borda="#C62828" if spi_medio < 0.90 else "#B8860B" if spi_medio < 1.0 else "#1E7E34")
with c4:
    render_kpi_card("Marcos em Atraso", str(int(marcos_crit)),
                    "marcos além da baseline",
                    cor_borda="#C62828" if marcos_crit > 0 else "#1E7E34")
with c5:
    render_kpi_card("Projetos Críticos (SPI<0.9)", str(int(n_criticos)),
                    "requerem ação imediata",
                    cor_borda="#C62828" if n_criticos > 0 else "#1E7E34")


# ──────────────────────────────────────────────────────────────────────────────
# 8. SEÇÃO B — EVM: INDICADORES DE VALOR AGREGADO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">B. Indicadores EVM — Desempenho de Prazo e Custo (PMBOK 8)</div>',
            unsafe_allow_html=True)

if df_fases.empty:
    st.info("Nenhum dado de fase disponível para o filtro selecionado.")
else:
    # ── Tabela de EVM por projeto ─────────────────────────────────────────────
    df_evm = df_fases.groupby("Projeto", as_index=False).agg(
        PV=("PV", "sum"), EV=("EV", "sum"), AC=("AC", "sum"),
        Pct_Concluida=("Pct_Concluida", "mean"),
    )
    df_evm["SPI"] = (df_evm["EV"] / df_evm["PV"]).round(3)
    df_evm["SV"]  = df_evm["EV"] - df_evm["PV"]
    df_evm["CPI"] = (df_evm["EV"] / df_evm["AC"]).round(3)
    # EAC (Estimativa no Término) = BAC / CPI  →  BAC aproximado via PV total
    df_evm["BAC"] = df_evm["PV"] / df_evm["Pct_Concluida"].clip(lower=0.01)
    df_evm["EAC"] = (df_evm["BAC"] / df_evm["CPI"]).round(0)

    # Gráfico de bolhas SPI x CPI (quadrante de saúde)
    fig_bubble = go.Figure()
    cores_bubble = [
        "#C62828" if s < 0.90 else "#B8860B" if s < 1.0 else "#1E7E34"
        for s in df_evm["SPI"]
    ]
    fig_bubble.add_trace(go.Scatter(
        x=df_evm["SPI"], y=df_evm["CPI"],
        mode="markers+text",
        marker=dict(size=df_evm["EV"]/12000, color=cores_bubble, opacity=0.85,
                    line=dict(width=1, color="white")),
        text=df_evm["Projeto"],
        textposition="top center",
        textfont=dict(size=11, color="#1B2A4A"),
        hovertemplate=(
            "<b>%{text}</b><br>SPI: %{x:.2f}<br>CPI: %{y:.2f}<extra></extra>"
        ),
    ))
    # Linhas de referência (meta = 1.0)
    fig_bubble.add_hline(y=1.0, line_dash="dot", line_color="#888", line_width=1)
    fig_bubble.add_vline(x=1.0, line_dash="dot", line_color="#888", line_width=1)
    fig_bubble.add_vline(x=0.90, line_dash="dash", line_color="#C62828",
                         line_width=1, annotation_text="Limite Crítico SPI",
                         annotation_position="top right",
                         annotation_font=dict(size=10, color="#C62828"))
    fig_bubble.update_layout(
        title=dict(text="Quadrante de Saúde — SPI × CPI", font=dict(size=13, color="#1B2A4A")),
        xaxis_title="SPI (Prazo)", yaxis_title="CPI (Custo)",
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(color="#1B2A4A"), height=360,
        margin=dict(l=40, r=20, t=50, b=40),
        xaxis=dict(gridcolor="#F0F0F0", zeroline=False),
        yaxis=dict(gridcolor="#F0F0F0", zeroline=False),
    )

    col_bubble, col_evm_table = st.columns([1, 1])
    with col_bubble:
        st.plotly_chart(fig_bubble, use_container_width=True)

    with col_evm_table:
        st.markdown("**Resumo EVM por Projeto**")
        df_exibir = df_evm[["Projeto", "PV", "EV", "AC", "SPI", "CPI", "SV"]].copy()
        df_exibir.columns = ["Projeto", "PV (R$)", "EV (R$)", "AC (R$)", "SPI", "CPI", "SV (R$)"]
        for col_m in ["PV (R$)", "EV (R$)", "AC (R$)", "SV (R$)"]:
            df_exibir[col_m] = df_exibir[col_m].apply(lambda x: f"R$ {x:,.0f}".replace(",", "."))

        def highlight_spi(row):
            spi = float(row["SPI"])
            if spi < 0.90:
                return ["background-color:#FDECEA"] * len(row)
            elif spi < 1.0:
                return ["background-color:#FFF8E1"] * len(row)
            return ["background-color:#E6F4EA"] * len(row)

        st.dataframe(
            df_exibir.style.apply(highlight_spi, axis=1),
            use_container_width=True, hide_index=True,
        )


# ──────────────────────────────────────────────────────────────────────────────
# 9. SEÇÃO C — ROADMAP EXECUTIVO (GANTT DE ALTO NÍVEL + MARCOS)
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">C. Roadmap Executivo — Cronograma e Marcos</div>',
            unsafe_allow_html=True)

df_gantt = df[df["Tipo"].isin(["Fase"])].copy()

if df_gantt.empty:
    st.info("Nenhuma fase encontrada para o filtro selecionado.")
else:
    fig_gantt = go.Figure()
    hoje_str = date.today().strftime("%Y-%m-%d")

    # ── Barras de Gantt por fase ──────────────────────────────────────────────
    for _, row in df_gantt.iterrows():
        cor_barra = (
            "#C62828" if (row["SPI"] is not None and row["SPI"] < 0.90)
            else "#B8860B" if (row["SPI"] is not None and row["SPI"] < 1.0)
            else "#1B2A4A"
        )
        fig_gantt.add_trace(go.Bar(
            x=[(pd.to_datetime(row["Termino_str"]) - pd.to_datetime(row["Inicio_str"])).days],
            y=[f"{row['Projeto']} — {row['Nome da Tarefa']}"],
            base=[row["Inicio_str"]],
            orientation="h",
            marker=dict(color=cor_barra, opacity=0.85),
            text=f"{row['Pct_Concluida']*100:.0f}%",
            textposition="inside",
            insidetextanchor="middle",
            hovertemplate=(
                f"<b>{row['Nome da Tarefa']}</b><br>"
                f"Início: {row['Inicio_str']}<br>"
                f"Término: {row['Termino_str']}<br>"
                f"SPI: {row['SPI']}<br>"
                f"Avanço: {row['Pct_Concluida']*100:.0f}%<extra></extra>"
            ),
            showlegend=False,
        ))

    # ── Marcos sobrepostos ────────────────────────────────────────────────────
    for _, row in df_marcos.iterrows():
        label_proj = f"{row['Projeto']} — {row['Nome da Tarefa']}"
        cor_marco  = "#C62828" if row["Marco_Atrasado"] else "#1E7E34"

        # Pega o y-index da fase correspondente (aproximação)
        y_ref = df_gantt[df_gantt["Projeto"] == row["Projeto"]]["Nome da Tarefa"].values
        y_label = f"{row['Projeto']} — {y_ref[0]}" if len(y_ref) > 0 else label_proj

        fig_gantt.add_trace(go.Scatter(
            x=[row["Termino_str"]],
            y=[y_label],
            mode="markers+text",
            marker=dict(symbol="diamond", size=14, color=cor_marco,
                        line=dict(width=1.5, color="white")),
            text=[f"◆ {row['Nome da Tarefa']}"],
            textposition="top center",
            textfont=dict(size=10, color=cor_marco),
            hovertemplate=(
                f"<b>Marco: {row['Nome da Tarefa']}</b><br>"
                f"Previsto: {row['Termino_Baseline_str']}<br>"
                f"Atual: {row['Termino_str']}<br>"
                f"Status: {'⚠️ Atrasado' if row['Marco_Atrasado'] else '✅ No prazo'}<extra></extra>"
            ),
            showlegend=False,
        ))

    # Linha de hoje
    fig_gantt.add_vline(
        x=hoje_str, line_color="#4A90D9", line_width=2, line_dash="dash",
        annotation_text="Hoje", annotation_position="top",
        annotation_font=dict(size=11, color="#4A90D9"),
    )

    fig_gantt.update_layout(
        barmode="overlay",
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(color="#1B2A4A", size=12),
        height=max(300, len(df_gantt) * 60 + 80),
        margin=dict(l=10, r=20, t=20, b=40),
        xaxis=dict(
            type="date", gridcolor="#F0F0F0",
            tickformat="%b/%y", tickfont=dict(size=11),
        ),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
    )

    st.plotly_chart(fig_gantt, use_container_width=True)

    # Legenda manual
    st.markdown(
        "<small style='color:#9AA5BE'>"
        "🔵 No prazo &nbsp;|&nbsp; 🟡 Atenção (SPI 0.9–1.0) &nbsp;|&nbsp; 🔴 Crítico (SPI < 0.9) &nbsp;|&nbsp;"
        "◆ Verde = Marco no prazo &nbsp;|&nbsp; ◆ Vermelho = Marco atrasado"
        "</small>",
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────────────────────────────────────
# 10. SEÇÃO D — MATRIZ DE PONTOS CRÍTICOS E PLANOS DE AÇÃO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">D. Matriz de Pontos Críticos — Planos de Ação (Domínio de Riscos)</div>',
            unsafe_allow_html=True)

# Filtra críticos: SPI < 0.90 OU marco atrasado
df_criticos_fases  = df_fases[df_fases["SPI"].notna() & (df_fases["SPI"] < 0.90)].copy()
df_criticos_marcos = df_marcos[df_marcos["Marco_Atrasado"] == True].copy()

# Une e deduplica por projeto
df_criticos = pd.concat([df_criticos_fases, df_criticos_marcos]).drop_duplicates(subset=["Projeto"])

if df_criticos.empty:
    st.success("✅ Nenhum desvio crítico identificado no portfólio selecionado.")
else:
    st.warning(
        f"⚠️ **{len(df_criticos)} projeto(s) com desvio crítico** identificado(s). "
        "Preencha os campos de causa e plano de ação antes da reunião de diretoria."
    )

    # Inicializa session_state para persistir edições dos campos textuais
    if "acoes" not in st.session_state:
        st.session_state["acoes"] = {}

    for _, row in df_criticos.iterrows():
        projeto_key = row["Projeto"]
        if projeto_key not in st.session_state["acoes"]:
            st.session_state["acoes"][projeto_key] = {
                "impacto": "",
                "causa":   "",
                "plano":   "",
                "resp":    row.get("Responsavel", ""),
            }

        with st.expander(f"🔴 {projeto_key} — SPI: {row['SPI']:.2f if row['SPI'] else 'N/A'}", expanded=True):
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown(f"**Status EVM:** {badge_spi(row['SPI'])}", unsafe_allow_html=True)
                st.markdown(f"**Avanço Físico:** {row['Pct_Concluida']*100:.0f}%")
                st.markdown(f"**SV (Variância de Prazo):** R$ {row['SV']:,.0f}".replace(",", "."))

            with col_b:
                st.markdown(f"**Término Previsto:** {row['Termino_str']}")
                st.markdown(f"**Baseline Término:** {row['Termino_Baseline_str']}")
                desvio_dias = (
                    pd.to_datetime(row["Termino_str"]) - pd.to_datetime(row["Termino_Baseline_str"])
                ).days
                if desvio_dias > 0:
                    st.markdown(f"**Desvio:** 🔴 +{desvio_dias} dias em atraso")
                else:
                    st.markdown(f"**Desvio:** 🟢 {desvio_dias} dias adiantado")

            st.markdown("---")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.session_state["acoes"][projeto_key]["impacto"] = st.text_area(
                    "Impacto no Negócio",
                    value=st.session_state["acoes"][projeto_key]["impacto"],
                    height=90,
                    placeholder="Ex: Atraso na entrega ao cliente X impacta receita de R$ 500k",
                    key=f"impacto_{projeto_key}",
                )
            with c2:
                st.session_state["acoes"][projeto_key]["causa"] = st.text_area(
                    "Causa Raiz do Desvio",
                    value=st.session_state["acoes"][projeto_key]["causa"],
                    height=90,
                    placeholder="Ex: Migração de dados subestimada — volume 3x maior",
                    key=f"causa_{projeto_key}",
                )
            with c3:
                st.session_state["acoes"][projeto_key]["plano"] = st.text_area(
                    "Plano de Ação Proposto",
                    value=st.session_state["acoes"][projeto_key]["plano"],
                    height=90,
                    placeholder="Ex: Alocação de 2 DBAs adicionais por 3 semanas + hora extra",
                    key=f"plano_{projeto_key}",
                )
            with c4:
                st.session_state["acoes"][projeto_key]["resp"] = st.text_input(
                    "Responsável",
                    value=st.session_state["acoes"][projeto_key]["resp"],
                    key=f"resp_{projeto_key}",
                )


# ──────────────────────────────────────────────────────────────────────────────
# 11. EXPORTAR RELATÓRIO — Botão para gerar Excel executivo
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Exportar Relatório Executivo</div>', unsafe_allow_html=True)

import io

def gerar_excel_executivo(df_base, acoes: dict) -> bytes:
    """Gera um Excel consolidado com EVM + Planos de Ação para a diretoria."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Aba 1: EVM Consolidado
        df_exp = df_base.groupby("Projeto", as_index=False).agg(
            Portfolio=("Portfolio", "first"),
            PV=("PV", "sum"), EV=("EV", "sum"), AC=("AC", "sum"),
            Pct_Concluida=("Pct_Concluida", "mean"),
            SPI=("SPI", "mean"), CPI=("CPI", "mean"),
        ).round(3)
        df_exp.to_excel(writer, sheet_name="EVM_Consolidado", index=False)

        # Aba 2: Pontos Críticos + Planos de Ação
        rows_acoes = []
        for proj, dados in acoes.items():
            rows_acoes.append({
                "Projeto": proj,
                "Impacto no Negócio": dados["impacto"],
                "Causa Raiz": dados["causa"],
                "Plano de Ação": dados["plano"],
                "Responsável": dados["resp"],
            })
        if rows_acoes:
            pd.DataFrame(rows_acoes).to_excel(writer, sheet_name="Planos_de_Acao", index=False)

        # Aba 3: Marcos
        df_marcos_exp = df[df["Marco"] == True][
            ["Portfolio", "Projeto", "Nome da Tarefa",
             "Termino_str", "Termino_Baseline_str", "Marco_Atrasado"]
        ].copy()
        df_marcos_exp.columns = ["Portfólio", "Projeto", "Marco",
                                  "Término Atual", "Baseline", "Atrasado"]
        df_marcos_exp.to_excel(writer, sheet_name="Marcos", index=False)

    return output.getvalue()


col_exp1, col_exp2 = st.columns([1, 4])
with col_exp1:
    if st.button("📥 Exportar Excel Executivo", use_container_width=True):
        acoes = st.session_state.get("acoes", {})
        excel_bytes = gerar_excel_executivo(df, acoes)
        st.download_button(
            label="⬇️ Baixar Relatório (.xlsx)",
            data=excel_bytes,
            file_name=f"relatorio_executivo_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:#C0C8D8;font-size:11px'>"
    "Dashboard Executivo · Baseado no PMBOK® 8ª Edição · PMI · "
    f"Gerado em {date.today().strftime('%d/%m/%Y')}"
    "</p>",
    unsafe_allow_html=True,
)
