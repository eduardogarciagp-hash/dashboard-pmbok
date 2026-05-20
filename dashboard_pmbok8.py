# =============================================================================
# DASHBOARD EXECUTIVO DE PORTFÓLIO — PMBOK 8ª EDIÇÃO
# Layout: Arquitetura Executiva para Diretoria
# =============================================================================
# INSTALAÇÃO:  pip install streamlit pandas plotly openpyxl
# EXECUÇÃO:    streamlit run dashboard_pmbok8.py
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
    page_title="Governança de Portfólio | Diretoria",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    .stApp { background-color: #F0F2F6; }
    .block-container { padding: 1.5rem 2rem 2rem 2rem !important; }
    .kpi-wrap {
        background: #FFFFFF; border-radius: 12px;
        padding: 22px 24px 18px 24px; border-top: 4px solid #0D1B2A;
        box-shadow: 0 1px 6px rgba(0,0,0,0.07); height: 100%;
    }
    .kpi-label { font-size: 11px; font-weight: 700; color: #8A9BB5;
                 text-transform: uppercase; letter-spacing: .8px; }
    .kpi-value { font-size: 42px; font-weight: 800; color: #0D1B2A;
                 line-height: 1.1; margin: 6px 0 4px 0; }
    .kpi-sub   { font-size: 12px; color: #B0BAD0; }
    .section-box {
        background: #FFFFFF; border-radius: 12px;
        padding: 22px 26px; box-shadow: 0 1px 6px rgba(0,0,0,0.07);
        margin-bottom: 18px;
    }
    .section-header {
        font-size: 11px; font-weight: 700; color: #8A9BB5;
        text-transform: uppercase; letter-spacing: 1px;
        border-bottom: 1px solid #E8ECF3;
        padding-bottom: 10px; margin-bottom: 16px;
    }
    .badge { display:inline-block; padding:3px 12px; border-radius:20px;
             font-size:11px; font-weight:700; }
    .badge-green  { background:#E6F4EA; color:#1E7E34; }
    .badge-yellow { background:#FFF8E1; color:#B8860B; }
    .badge-red    { background:#FDECEA; color:#C62828; }
    section[data-testid="stSidebar"] { background-color: #0D1B2A; }
    section[data-testid="stSidebar"] * { color: #FFFFFF !important; }
    .modebar { display: none !important; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 1. FUNÇÕES EVM — vetorial
# ──────────────────────────────────────────────────────────────────────────────
def calcular_evm(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["EV", "PV", "AC"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df["SPI"] = np.where(df["PV"] > 0, (df["EV"] / df["PV"]).round(3), np.nan)
    df["SV"]  = df["EV"] - df["PV"]
    df["CPI"] = np.where(df["AC"] > 0, (df["EV"] / df["AC"]).round(3), np.nan)
    def status(s):
        if pd.isna(s): return "N/A"
        if s >= 1.0:   return "🟢 No Prazo"
        if s >= 0.90:  return "🟡 Atenção"
        return "🔴 Crítico"
    df["Status_SPI"] = df["SPI"].apply(status)
    return df


# ──────────────────────────────────────────────────────────────────────────────
# 2. MOCK DATA
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data
def carregar_dados_mock() -> pd.DataFrame:
    hoje = date.today()
    dados = [
        {"Portfolio":"Portfólio Geral","Projeto":"ERP Cloud Migration","ID":1,
         "Nome da Tarefa":"ERP — Fase 1: Planejamento","Tipo":"Fase","Nivel":1,
         "Inicio":hoje-timedelta(120),"Termino":hoje-timedelta(30),
         "Termino_Baseline":hoje-timedelta(30),"Pct_Concluida":1.0,
         "AC":420000,"PV":400000,"EV":400000,"Marco":False,
         "Recursos":"Ana Costa","Responsavel":"Ana Costa","Causa_Raiz":"","Plano_Acao":""},
        {"Portfolio":"Portfólio Geral","Projeto":"ERP Cloud Migration","ID":2,
         "Nome da Tarefa":"ERP — Fase 2: Implementação","Tipo":"Fase","Nivel":1,
         "Inicio":hoje-timedelta(30),"Termino":hoje+timedelta(60),
         "Termino_Baseline":hoje+timedelta(45),"Pct_Concluida":0.38,
         "AC":430000,"PV":380000,"EV":300000,"Marco":False,
         "Recursos":"Ana Costa","Responsavel":"Ana Costa","Causa_Raiz":"","Plano_Acao":""},
        {"Portfolio":"Portfólio Geral","Projeto":"ERP Cloud Migration","ID":3,
         "Nome da Tarefa":"Kickoff ERP","Tipo":"Marco","Nivel":2,
         "Inicio":hoje-timedelta(120),"Termino":hoje-timedelta(120),
         "Termino_Baseline":hoje-timedelta(120),"Pct_Concluida":1.0,
         "AC":0,"PV":0,"EV":0,"Marco":True,
         "Recursos":"Ana Costa","Responsavel":"Ana Costa","Causa_Raiz":"","Plano_Acao":""},
        {"Portfolio":"Portfólio Geral","Projeto":"ERP Cloud Migration","ID":4,
         "Nome da Tarefa":"Go-Live ERP","Tipo":"Marco","Nivel":2,
         "Inicio":hoje+timedelta(60),"Termino":hoje+timedelta(60),
         "Termino_Baseline":hoje+timedelta(45),"Pct_Concluida":0.0,
         "AC":0,"PV":0,"EV":0,"Marco":True,
         "Recursos":"Ana Costa","Responsavel":"Ana Costa","Causa_Raiz":"","Plano_Acao":""},
        {"Portfolio":"Portfólio Geral","Projeto":"Data Analytics","ID":5,
         "Nome da Tarefa":"Analytics — Fase 1: Discovery","Tipo":"Fase","Nivel":1,
         "Inicio":hoje-timedelta(90),"Termino":hoje+timedelta(20),
         "Termino_Baseline":hoje+timedelta(20),"Pct_Concluida":0.75,
         "AC":200000,"PV":195000,"EV":200000,"Marco":False,
         "Recursos":"Maria Souza","Responsavel":"Maria Souza","Causa_Raiz":"","Plano_Acao":""},
        {"Portfolio":"Portfólio Geral","Projeto":"Data Analytics","ID":6,
         "Nome da Tarefa":"Analytics — Fase 2: Entrega","Tipo":"Fase","Nivel":1,
         "Inicio":hoje+timedelta(20),"Termino":hoje+timedelta(100),
         "Termino_Baseline":hoje+timedelta(100),"Pct_Concluida":0.0,
         "AC":0,"PV":0,"EV":0,"Marco":False,
         "Recursos":"Maria Souza","Responsavel":"Maria Souza","Causa_Raiz":"","Plano_Acao":""},
        {"Portfolio":"Portfólio Geral","Projeto":"Data Analytics","ID":7,
         "Nome da Tarefa":"MVP Analytics","Tipo":"Marco","Nivel":2,
         "Inicio":hoje+timedelta(100),"Termino":hoje+timedelta(100),
         "Termino_Baseline":hoje+timedelta(100),"Pct_Concluida":0.0,
         "AC":0,"PV":0,"EV":0,"Marco":True,
         "Recursos":"Maria Souza","Responsavel":"Maria Souza","Causa_Raiz":"","Plano_Acao":""},
        {"Portfolio":"Portfólio Geral","Projeto":"Filial Sul","ID":8,
         "Nome da Tarefa":"Filial Sul — Fase 1: Obras","Tipo":"Fase","Nivel":1,
         "Inicio":hoje-timedelta(150),"Termino":hoje-timedelta(20),
         "Termino_Baseline":hoje-timedelta(40),"Pct_Concluida":1.0,
         "AC":600000,"PV":500000,"EV":480000,"Marco":False,
         "Recursos":"Carlos Melo","Responsavel":"Carlos Melo","Causa_Raiz":"","Plano_Acao":""},
        {"Portfolio":"Portfólio Geral","Projeto":"Filial Sul","ID":9,
         "Nome da Tarefa":"Filial Sul — Fase 2: Operação","Tipo":"Fase","Nivel":1,
         "Inicio":hoje-timedelta(20),"Termino":hoje+timedelta(40),
         "Termino_Baseline":hoje+timedelta(10),"Pct_Concluida":0.30,
         "AC":500000,"PV":400000,"EV":310000,"Marco":False,
         "Recursos":"Carlos Melo","Responsavel":"Carlos Melo","Causa_Raiz":"","Plano_Acao":""},
        {"Portfolio":"Portfólio Geral","Projeto":"Filial Sul","ID":10,
         "Nome da Tarefa":"Inauguração Filial","Tipo":"Marco","Nivel":2,
         "Inicio":hoje+timedelta(40),"Termino":hoje+timedelta(40),
         "Termino_Baseline":hoje+timedelta(10),"Pct_Concluida":0.0,
         "AC":0,"PV":0,"EV":0,"Marco":True,
         "Recursos":"Carlos Melo","Responsavel":"Carlos Melo","Causa_Raiz":"","Plano_Acao":""},
    ]
    df = pd.DataFrame(dados)
    df = calcular_evm(df)
    df["Marco_Atrasado"] = df["Marco"] & (df["Termino"] > df["Termino_Baseline"])
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
    df_raw.columns = df_raw.columns.str.strip()

    with st.expander("🔍 Diagnóstico: colunas detectadas no Excel", expanded=False):
        st.code("\n".join([f"{i+1:02d}. {c}" for i, c in enumerate(df_raw.columns)]))
        st.dataframe(df_raw.head(3), use_container_width=True)

    col_map = {
        "Início":"Inicio","Start":"Inicio",
        "Término":"Termino","Finish":"Termino",
        "Nome":"Nome da Tarefa","Name":"Nome da Tarefa",
        "Task Name":"Nome da Tarefa","Nome da Tarefa":"Nome da Tarefa",
        "Nível da estrutura de tópicos":"Nivel","Outline Level":"Nivel",
        "Duração":"Duracao","Duration":"Duracao",
        "Id":"ID",
        "Nomes dos Recursos":"Recursos","Resource Names":"Recursos",
        "% concluída":"Pct_Concluida","% Concluída":"Pct_Concluida",
        "% Complete":"Pct_Concluida","% Completo":"Pct_Concluida",
        "Custo Real (CR)":"AC","Custo Real":"AC","Actual Cost":"AC","ACWP":"AC","AC":"AC",
        "COTA":"PV","Baseline Cost":"PV","BCWS":"PV","PV":"PV",
        "COTR":"EV","Earned Value":"EV","BCWP":"EV","EV":"EV",
        "Marco":"Marco","Milestone":"Marco",
        "Projeto":"Projeto","Project":"Projeto",
        "Portfólio":"Portfolio","Portfolio":"Portfolio",
    }
    df_raw.rename(columns={k:v for k,v in col_map.items() if k in df_raw.columns}, inplace=True)

    nome_arquivo = getattr(arquivo_excel, "name", "Projeto")
    nome_projeto = nome_arquivo.replace(".xlsx","").replace(".xls","")

    if "Portfolio"     not in df_raw.columns: df_raw["Portfolio"]     = "Portfólio Geral"
    if "Projeto"       not in df_raw.columns: df_raw["Projeto"]       = nome_projeto
    if "Responsavel"   not in df_raw.columns: df_raw["Responsavel"]   = ""
    if "Causa_Raiz"    not in df_raw.columns: df_raw["Causa_Raiz"]    = ""
    if "Plano_Acao"    not in df_raw.columns: df_raw["Plano_Acao"]    = ""
    if "Pct_Concluida" not in df_raw.columns: df_raw["Pct_Concluida"] = 0.0
    if "AC"            not in df_raw.columns: df_raw["AC"]            = 0.0
    if "PV"            not in df_raw.columns: df_raw["PV"]            = 0.0
    if "EV"            not in df_raw.columns: df_raw["EV"]            = 0.0
    if "Recursos"      not in df_raw.columns: df_raw["Recursos"]      = ""

    df_raw["Inicio"]  = pd.to_datetime(df_raw.get("Inicio"),  errors="coerce")
    df_raw["Termino"] = pd.to_datetime(df_raw.get("Termino"), errors="coerce")
    df_raw = df_raw.dropna(subset=["Inicio","Termino"])

    if "Termino_Baseline" not in df_raw.columns:
        df_raw["Termino_Baseline"] = df_raw["Termino"]

    df_raw["Pct_Concluida"] = pd.to_numeric(df_raw["Pct_Concluida"], errors="coerce").fillna(0)
    if df_raw["Pct_Concluida"].max() > 1:
        df_raw["Pct_Concluida"] = df_raw["Pct_Concluida"] / 100

    if "Marco" not in df_raw.columns:
        if "Duracao" in df_raw.columns:
            df_raw["Marco"] = df_raw["Duracao"].astype(str).str.strip().isin(
                ["0","0 dias","0d","0 days"])
        else:
            df_raw["Marco"] = False
    else:
        df_raw["Marco"] = df_raw["Marco"].astype(str).str.lower().isin(
            ["sim","true","1","yes"])

    if "Nivel" in df_raw.columns:
        df_raw["Nivel"] = pd.to_numeric(df_raw["Nivel"], errors="coerce").fillna(99)
        def definir_tipo(row):
            if row["Marco"]:      return "Marco"
            if row["Nivel"] <= 1: return "Fase"
            return "Tarefa"
        df_raw["Tipo"] = df_raw.apply(definir_tipo, axis=1)
    else:
        df_raw["Tipo"] = df_raw["Marco"].apply(lambda m: "Marco" if m else "Fase")

    df_raw = calcular_evm(df_raw)
    df_raw["Marco_Atrasado"] = (
        df_raw["Marco"] &
        (pd.to_datetime(df_raw["Termino"], errors="coerce") >
         pd.to_datetime(df_raw["Termino_Baseline"], errors="coerce"))
    )
    df_raw["Inicio_str"]           = df_raw["Inicio"].dt.strftime("%Y-%m-%d")
    df_raw["Termino_str"]          = df_raw["Termino"].dt.strftime("%Y-%m-%d")
    df_raw["Termino_Baseline_str"] = pd.to_datetime(
        df_raw["Termino_Baseline"], errors="coerce").dt.strftime("%Y-%m-%d")
    return df_raw


# ──────────────────────────────────────────────────────────────────────────────
# 4. SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    st.markdown("---")
    arquivo  = st.file_uploader("📂 Importar Excel MS Project", type=["xlsx","xls"])
    data_ref = st.date_input("📅 Data de Referência", value=date.today())
    logo_url = st.text_input("🖼️ URL do Logo (opcional)", value="")


# ──────────────────────────────────────────────────────────────────────────────
# 5. CARREGA DADOS
# ──────────────────────────────────────────────────────────────────────────────
df_full        = carregar_dados(arquivo)
projetos_lista = ["Todos"] + sorted(df_full["Projeto"].dropna().unique().tolist())


# ──────────────────────────────────────────────────────────────────────────────
# 6. CABEÇALHO + FILTROS
# ──────────────────────────────────────────────────────────────────────────────
col_logo, col_title, col_filtros = st.columns([1, 4, 3])

with col_logo:
    if logo_url:
        st.image(logo_url, width=120)
    else:
        st.markdown(
            "<div style='width:72px;height:48px;background:#1B3A6B;border-radius:8px;"
            "display:flex;align-items:center;justify-content:center;"
            "color:white;font-weight:800;font-size:13px'>PMO</div>",
            unsafe_allow_html=True)

with col_title:
    st.markdown(
        "<div style='padding-top:4px'>"
        "<span style='font-size:19px;font-weight:800;color:#0D1B2A'>"
        "DASHBOARD DE GOVERNANÇA E VALOR DO PORTFÓLIO</span><br>"
        f"<span style='font-size:12px;color:#8A9BB5'>PMBOK 8ª Ed. · "
        f"Referência: {data_ref.strftime('%d/%m/%Y')}</span></div>",
        unsafe_allow_html=True)

with col_filtros:
    visao_sel   = st.radio("Visão",
                           ["Portfólio Global","Diretoria de Tecnologia","Unidade de Negócio"],
                           horizontal=True, label_visibility="collapsed")
    projeto_sel = st.selectbox("Projeto", projetos_lista, label_visibility="collapsed")

st.markdown("<hr style='margin:10px 0 14px 0;border-color:#E0E4EC'>", unsafe_allow_html=True)

# Aplica filtro
df        = df_full if projeto_sel == "Todos" else df_full[df_full["Projeto"] == projeto_sel]
df_fases  = df[df["Tipo"] == "Fase"]
df_marcos = df[df["Marco"] == True]


# ──────────────────────────────────────────────────────────────────────────────
# 7. BARRA DE SAÚDE HOLÍSTICA
# ──────────────────────────────────────────────────────────────────────────────
spi_saude = df_fases["SPI"].dropna().mean() if not df_fases.empty else None
pct_saude = min(float(spi_saude), 1.0) * 100 if spi_saude is not None else 0
n_crit_s  = int((df_fases["SPI"].dropna() < 0.90).sum())
n_total_s = int(df_fases["SPI"].dropna().count())

if spi_saude is None or n_total_s == 0:
    cor_saude   = "#8A9BB5"; label_saude = "SEM DADOS EVM"
elif pct_saude >= 95:
    cor_saude   = "#1E7E34"; label_saude = "ESTRATÉGIA SAUDÁVEL"
elif pct_saude >= 90:
    cor_saude   = "#B8860B"; label_saude = "ATENÇÃO — DESVIOS MODERADOS"
else:
    cor_saude   = "#C62828"; label_saude = "CRÍTICO — INTERVENÇÃO NECESSÁRIA"

if spi_saude and n_total_s > 0:
    spi_str   = f"{spi_saude:.2f}"
    barra_html = f"""
    <div style='background:#FFFFFF;border-radius:10px;padding:14px 22px;
                box-shadow:0 1px 6px rgba(0,0,0,0.07);margin-bottom:16px'>
      <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>
        <span style='font-size:11px;font-weight:700;color:#8A9BB5;
                     text-transform:uppercase;letter-spacing:.8px'>
          SAÚDE HOLÍSTICA DA ESTRATÉGIA — {visao_sel.upper()}
        </span>
        <span style='font-size:12px;font-weight:700;color:{cor_saude}'>{label_saude}</span>
      </div>
      <div style='background:#F0F2F6;border-radius:20px;height:10px;width:100%;overflow:hidden'>
        <div style='background:{cor_saude};height:10px;width:{pct_saude:.1f}%;border-radius:20px'></div>
      </div>
      <div style='display:flex;justify-content:space-between;margin-top:6px'>
        <span style='font-size:11px;color:#8A9BB5'>
          SPI Médio: <b style='color:{cor_saude}'>{spi_str}</b>
          &nbsp;·&nbsp; {n_crit_s} de {n_total_s} projetos em zona crítica
        </span>
        <span style='font-size:11px;color:#8A9BB5'>{pct_saude:.1f}% da meta de entrega</span>
      </div>
    </div>"""
else:
    barra_html = f"""
    <div style='background:#FFFFFF;border-radius:10px;padding:14px 22px;
                box-shadow:0 1px 6px rgba(0,0,0,0.07);margin-bottom:16px'>
      <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>
        <span style='font-size:11px;font-weight:700;color:#8A9BB5;
                     text-transform:uppercase;letter-spacing:.8px'>
          SAÚDE HOLÍSTICA DA ESTRATÉGIA — {visao_sel.upper()}
        </span>
        <span style='font-size:12px;font-weight:700;color:#8A9BB5'>SEM DADOS EVM</span>
      </div>
      <div style='background:#F0F2F6;border-radius:20px;height:10px;width:100%;overflow:hidden'>
        <div style='background:#D0D6E0;height:10px;width:100%;border-radius:20px'></div>
      </div>
      <div style='margin-top:6px'>
        <span style='font-size:11px;color:#8A9BB5'>
          Importe um Excel com colunas EVM (COTA/PV, COTR/EV, Custo Real/AC) para habilitar.
        </span>
      </div>
    </div>"""

st.markdown(barra_html, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 8. CARDS KPI
# ──────────────────────────────────────────────────────────────────────────────
n_projetos = df_fases["Projeto"].nunique()
pct_media  = df_fases["Pct_Concluida"].mean() * 100 if not df_fases.empty else 0
spi_medio  = df_fases["SPI"].dropna().mean() if not df_fases.empty else None
marcos_at  = int(df_marcos["Marco_Atrasado"].sum()) if not df_marcos.empty else 0

def cor_spi(s):
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return "#8A9BB5", "N/A", "Sem dados EVM"
    if s >= 1.0:  return "#1E7E34", f"{s:.2f} ✅", "No prazo"
    if s >= 0.90: return "#B8860B", f"{s:.2f} ⚠️", "Atenção"
    return "#C62828", f"{s:.2f} 🔴", "Crítico"

cor1, val_spi, label_spi = cor_spi(spi_medio)
cor_m = "#C62828" if marcos_at > 0 else "#1E7E34"

c1, c2, c3, c4 = st.columns(4)
for col_widget, label, valor, sub, cor_top in [
    (c1, "PROJETOS ATIVOS",  str(n_projetos),      "no portfólio",        "#0D1B2A"),
    (c2, "CONCLUSÃO MÉDIA",  f"{pct_media:.1f}%",  "avanço físico médio", "#0D1B2A"),
    (c3, "SPI DO PORTFÓLIO", val_spi,              label_spi,             cor1),
    (c4, "MARCOS EM ATRASO", str(marcos_at),       "além da baseline",    cor_m),
]:
    with col_widget:
        st.markdown(f"""
        <div class="kpi-wrap" style="border-top-color:{cor_top}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value" style="color:{cor_top}">{valor}</div>
            <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 9. SECTION 1 — ROADMAP EXECUTIVO & MARCOS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-header">📅 SECTION 1 — ROADMAP EXECUTIVO & MARCOS DE VALOR</div>',
            unsafe_allow_html=True)

df_gantt   = df[df["Tipo"] == "Fase"].copy()
df_marcos2 = df[df["Marco"] == True].copy()

if df_gantt.empty:
    st.info("Nenhuma fase detectada. Verifique o arquivo ou use os dados de exemplo.")
else:
    projetos_gantt = df_gantt["Projeto"].unique().tolist()
    paleta = ["#1B3A6B","#2E6DA4","#1A7F5A","#7B3FA0","#B85C00","#8B0000","#2C6E7A","#5A4A00"]
    cores_proj = {p: paleta[i % len(paleta)] for i, p in enumerate(projetos_gantt)}

    fig = go.Figure()
    hoje_str = date.today().strftime("%Y-%m-%d")
    hoje_dt  = date.today()

    # ── Barras de fase ────────────────────────────────────────────────────────
    for _, row in df_gantt.iterrows():
        dur     = max((pd.to_datetime(row["Termino_str"]) -
                       pd.to_datetime(row["Inicio_str"])).days, 1)
        spi_val = row["SPI"] if not pd.isna(row.get("SPI", np.nan)) else None

        if spi_val is None:
            cor_barra = cores_proj.get(row["Projeto"], "#1B3A6B")
        elif spi_val < 0.90:
            cor_barra = "#C62828"
        elif spi_val < 1.0:
            cor_barra = "#B8860B"
        else:
            cor_barra = cores_proj.get(row["Projeto"], "#1B3A6B")

        label_pct = f"{row['Pct_Concluida']*100:.0f}%" if row["Pct_Concluida"] > 0 else ""
        spi_hover = f"<br>SPI: {spi_val:.2f}" if spi_val else ""

        fig.add_trace(go.Bar(
            x=[dur], y=[row["Projeto"]],
            base=[row["Inicio_str"]],
            orientation="h",
            marker=dict(color=cor_barra, opacity=0.88, line=dict(width=0)),
            text=label_pct,
            textposition="inside", insidetextanchor="middle",
            textfont=dict(size=11, color="white", family="Arial"),
            hovertemplate=(
                f"<b>{row['Nome da Tarefa']}</b><br>"
                f"Início: {row['Inicio_str']} → Término: {row['Termino_str']}<br>"
                f"Avanço: {row['Pct_Concluida']*100:.0f}%"
                + spi_hover + "<extra></extra>"
            ),
            showlegend=False,
        ))

    # ── Marcos — com status calculado dentro do loop (bug corrigido) ──────────
    for _, row in df_marcos2.iterrows():
        termino_dt = pd.to_datetime(row["Termino_str"]).date()
        concluido  = row["Pct_Concluida"] >= 1.0
        atrasado   = bool(row["Marco_Atrasado"])   # ← variável definida aqui
        y_proj     = row["Projeto"]                 # ← variável definida aqui

        if concluido:
            cor_m2 = "#1E7E34"; simbolo = "star";   label_icone = "📅"
        elif atrasado:
            cor_m2 = "#C62828"; simbolo = "circle"; label_icone = "🔴"
        elif termino_dt > hoje_dt:
            cor_m2 = "#8A9BB5"; simbolo = "diamond"; label_icone = "🔶"
        else:
            cor_m2 = "#1E7E34"; simbolo = "diamond"; label_icone = "♦"

        status_txt = (
            "📅 CONCLUÍDO"  if concluido
            else "🔴 ATRASADO" if atrasado
            else "🔶 AGENDADO" if termino_dt > hoje_dt
            else "♦ NO PRAZO"
        )

        fig.add_trace(go.Scatter(
            x=[row["Termino_str"]],
            y=[y_proj],
            mode="markers+text",
            marker=dict(symbol=simbolo, size=16, color=cor_m2,
                        line=dict(width=2, color="white")),
            text=[f"  {label_icone} {row['Nome da Tarefa']}"],
            textposition="middle right",
            textfont=dict(size=10, color=cor_m2),
            hovertemplate=(
                f"<b>{row['Nome da Tarefa']}</b><br>"
                f"Baseline: {row['Termino_Baseline_str']}<br>"
                f"Atual: {row['Termino_str']}<br>"
                f"Status: {status_txt}<extra></extra>"
            ),
            showlegend=False,
        ))

    # Linha "Hoje"
    fig.add_shape(
        type="line", xref="x", yref="paper",
        x0=hoje_str, x1=hoje_str, y0=0, y1=1,
        line=dict(color="#4A90D9", width=2, dash="dash"),
    )
    fig.add_annotation(
        x=hoje_str, y=1.02, yref="paper",
        text="Hoje", showarrow=False, xanchor="center",
        font=dict(size=11, color="#4A90D9"),
    )

    fig.update_layout(
        barmode="stack",
        plot_bgcolor="white", paper_bgcolor="white",
        height=max(280, len(projetos_gantt) * 72 + 80),
        margin=dict(l=10, r=180, t=30, b=40),
        xaxis=dict(type="date", gridcolor="#F0F2F6",
                   tickformat="%b/%y", tickfont=dict(size=11, color="#6B7A99"),
                   showline=False),
        yaxis=dict(autorange="reversed", showgrid=False,
                   tickfont=dict(size=12, color="#0D1B2A")),
        hoverlabel=dict(bgcolor="white", font_size=12),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("""
    <div style='display:flex;gap:20px;flex-wrap:wrap;font-size:11px;
                color:#6B7A99;margin-top:-8px;padding:8px 4px'>
      <span>📅 <b style='color:#1E7E34'>Concluído</b></span>
      <span>♦ <b style='color:#1E7E34'>No Prazo</b></span>
      <span>🔴 <b style='color:#C62828'>Atrasado</b></span>
      <span>🔶 <b style='color:#8A9BB5'>Agendado / Futuro</b></span>
      <span><span style='color:#C62828'>■</span> Fase crítica (SPI&lt;0.90)</span>
      <span><span style='color:#B8860B'>■</span> Fase em atenção</span>
      <span><b style='color:#4A90D9'>— —</b> Hoje</span>
    </div>""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 10. SECTION 2 — EVM
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-header">📊 SECTION 2 — ANÁLISE DE VALOR AGREGADO (EVM) POR PROJETO</div>',
            unsafe_allow_html=True)

tem_evm = df_fases[["PV","EV","AC"]].sum().sum() > 0

if not tem_evm:
    st.info("ℹ️ Dados de EVM não encontrados. Exporte do MS Project incluindo "
            "COTA (PV), COTR (EV) e Custo Real (AC) para habilitar esta seção.")
else:
    df_evm = df_fases.groupby("Projeto", as_index=False).agg(
        PV=("PV","sum"), EV=("EV","sum"), AC=("AC","sum"),
        Pct=("Pct_Concluida","mean"),
    )
    df_evm["SPI"] = np.where(df_evm["PV"]>0, (df_evm["EV"]/df_evm["PV"]).round(3), np.nan)
    df_evm["CPI"] = np.where(df_evm["AC"]>0, (df_evm["EV"]/df_evm["AC"]).round(3), np.nan)
    df_evm["SV"]  = df_evm["EV"] - df_evm["PV"]
    df_evm["CV"]  = df_evm["EV"] - df_evm["AC"]

    cores_b = [
        "#C62828" if (not np.isnan(s) and s < 0.90)
        else "#B8860B" if (not np.isnan(s) and s < 1.0)
        else "#1B3A6B"
        for s in df_evm["SPI"].fillna(1)
    ]
    ev_max = df_evm["EV"].max()
    tam_bolha = (df_evm["EV"].clip(lower=1) / max(ev_max, 1) * 40 + 15)

    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(
        x=df_evm["SPI"], y=df_evm["CPI"],
        mode="markers+text",
        marker=dict(size=tam_bolha, color=cores_b, opacity=0.80,
                    line=dict(width=2, color="white")),
        text=df_evm["Projeto"], textposition="top center",
        textfont=dict(size=11, color="#0D1B2A"),
        hovertemplate="<b>%{text}</b><br>SPI: %{x:.2f}<br>CPI: %{y:.2f}<extra></extra>",
    ))
    fig2.add_hline(y=1.0, line_dash="dot", line_color="#CCCCCC", line_width=1)
    fig2.add_vline(x=1.0, line_dash="dot", line_color="#CCCCCC", line_width=1)
    fig2.add_vline(x=0.90, line_dash="dash", line_color="#C62828", line_width=1)
    cpi_max = df_evm["CPI"].dropna().max() if not df_evm["CPI"].isna().all() else 1.1
    fig2.add_annotation(x=0.90, y=cpi_max + 0.05, text="Limite crítico",
                        showarrow=False, font=dict(size=10, color="#C62828"), xanchor="left")
    fig2.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        height=320, margin=dict(l=40,r=20,t=20,b=40),
        xaxis=dict(title="SPI — Prazo", gridcolor="#F0F2F6", zeroline=False),
        yaxis=dict(title="CPI — Custo", gridcolor="#F0F2F6", zeroline=False),
    )

    col_g, col_t = st.columns([1,1])
    with col_g:
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})
    with col_t:
        st.markdown("**Indicadores por Projeto**")
        df_show = df_evm[["Projeto","SPI","CPI","SV","CV","Pct"]].copy()
        df_show.columns = ["Projeto","SPI","CPI","SV (R$)","CV (R$)","Avanço"]
        df_show["SV (R$)"] = df_show["SV (R$)"].apply(lambda x: f"R$ {x:,.0f}".replace(",","."))
        df_show["CV (R$)"] = df_show["CV (R$)"].apply(lambda x: f"R$ {x:,.0f}".replace(",","."))
        df_show["Avanço"]  = df_show["Avanço"].apply(lambda x: f"{x*100:.0f}%")

        def hl(row):
            try:
                s = float(row["SPI"])
                if s < 0.90: return ["background:#FDECEA"] * len(row)
                if s < 1.0:  return ["background:#FFF8E1"] * len(row)
            except: pass
            return ["background:#E6F4EA"] * len(row)

        st.dataframe(df_show.style.apply(hl, axis=1),
                     use_container_width=True, hide_index=True)

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 11. SECTION 3 — GOVERNANÇA DE INCERTEZAS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-header">⚠️ SECTION 3 — GOVERNANÇA DE INCERTEZAS: PONTOS CRÍTICOS E PLANOS DE AÇÃO</div>',
            unsafe_allow_html=True)

df_crit = pd.concat([
    df_fases[df_fases["SPI"].notna() & (df_fases["SPI"] < 0.90)],
    df_marcos[df_marcos["Marco_Atrasado"] == True],
]).drop_duplicates(subset=["Projeto"])

if df_crit.empty:
    st.success("✅ Nenhum desvio crítico identificado. Portfólio sob controle.")
else:
    st.warning(f"⚠️ {len(df_crit)} projeto(s) requerem atenção da diretoria. "
               "Preencha os campos abaixo antes da reunião.")

    if "acoes" not in st.session_state:
        st.session_state["acoes"] = {}

    for _, row in df_crit.iterrows():
        pk  = row["Projeto"]
        spi = row.get("SPI", None)
        spi_str = f"{spi:.2f}" if (spi is not None and not np.isnan(float(spi))) else "N/A"

        if pk not in st.session_state["acoes"]:
            st.session_state["acoes"][pk] = {
                "impacto":"", "causa":"", "plano":"",
                "resp": str(row.get("Responsavel","")),
            }

        desvio_dias = int((
            pd.to_datetime(row["Termino_str"]) -
            pd.to_datetime(row["Termino_Baseline_str"])
        ).days)

        with st.expander(
            f"🔴 {pk}  |  SPI: {spi_str}  |  Desvio: +{desvio_dias} dias",
            expanded=True
        ):
            m1, m2, m3 = st.columns(3)
            with m1:
                st.markdown(f"**Avanço Físico:** {row['Pct_Concluida']*100:.0f}%")
                st.markdown(f"**Término Previsto:** {row['Termino_str']}")
            with m2:
                st.markdown(f"**Baseline:** {row['Termino_Baseline_str']}")
                sv_val = row.get("SV", 0) or 0
                st.markdown(f"**SV (Variância Prazo):** R$ {float(sv_val):,.0f}".replace(",","."))
            with m3:
                eh_critico = (spi is not None and not np.isnan(float(spi)) and float(spi) < 0.90)
                badge = ('<span class="badge badge-red">CRÍTICO</span>' if eh_critico
                         else '<span class="badge badge-yellow">ATENÇÃO</span>')
                st.markdown(badge, unsafe_allow_html=True)
                if row.get("Marco_Atrasado"):
                    st.markdown('<span class="badge badge-red">MARCO ATRASADO</span>',
                                unsafe_allow_html=True)

            st.markdown("---")
            fa, fb, fc, fd = st.columns(4)
            with fa:
                st.session_state["acoes"][pk]["impacto"] = st.text_area(
                    "🏢 Impacto no Negócio",
                    value=st.session_state["acoes"][pk]["impacto"], height=100,
                    placeholder="Ex: Atraso impacta receita de R$ 500k no Q3",
                    key=f"imp_{pk}")
            with fb:
                st.session_state["acoes"][pk]["causa"] = st.text_area(
                    "🔍 Causa Raiz",
                    value=st.session_state["acoes"][pk]["causa"], height=100,
                    placeholder="Ex: Volume de dados 3x maior que estimado",
                    key=f"cau_{pk}")
            with fc:
                st.session_state["acoes"][pk]["plano"] = st.text_area(
                    "🎯 Plano de Ação",
                    value=st.session_state["acoes"][pk]["plano"], height=100,
                    placeholder="Ex: +2 consultores por 3 semanas + revisão de escopo",
                    key=f"pla_{pk}")
            with fd:
                st.session_state["acoes"][pk]["resp"] = st.text_input(
                    "👤 Responsável",
                    value=st.session_state["acoes"][pk]["resp"],
                    key=f"res_{pk}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 12. EXPORTAÇÃO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-header">📥 EXPORTAR RELATÓRIO EXECUTIVO</div>',
            unsafe_allow_html=True)

def gerar_excel(df_base, acoes):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        g = df_base.groupby("Projeto", as_index=False).agg(
            Portfolio=("Portfolio","first"),
            PV=("PV","sum"), EV=("EV","sum"), AC=("AC","sum"),
            Pct=("Pct_Concluida","mean"),
        )
        g["SPI"] = np.where(g["PV"]>0,(g["EV"]/g["PV"]).round(3),np.nan)
        g["CPI"] = np.where(g["AC"]>0,(g["EV"]/g["AC"]).round(3),np.nan)
        g.to_excel(writer, sheet_name="EVM_Consolidado", index=False)

        rows_ac = [{"Projeto":p,"Impacto":d["impacto"],"Causa Raiz":d["causa"],
                    "Plano de Ação":d["plano"],"Responsável":d["resp"]}
                   for p,d in acoes.items()]
        if rows_ac:
            pd.DataFrame(rows_ac).to_excel(writer, sheet_name="Planos_de_Acao", index=False)

        if not df_marcos.empty:
            dm = df_marcos[["Portfolio","Projeto","Nome da Tarefa",
                             "Termino_str","Termino_Baseline_str","Marco_Atrasado"]].copy()
            dm.columns = ["Portfólio","Projeto","Marco","Término","Baseline","Atrasado"]
            dm.to_excel(writer, sheet_name="Marcos", index=False)

    return output.getvalue()

col_e1, _ = st.columns([1,4])
with col_e1:
    if st.button("📥 Gerar Excel Executivo", use_container_width=True):
        excel_bytes = gerar_excel(df, st.session_state.get("acoes",{}))
        st.download_button(
            "⬇️ Baixar Relatório (.xlsx)", data=excel_bytes,
            file_name=f"governanca_portfolio_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.markdown('</div>', unsafe_allow_html=True)
st.markdown(
    f"<p style='text-align:center;color:#C0C8D8;font-size:11px;margin-top:20px'>"
    f"Dashboard Executivo · PMBOK® 8ª Edição · PMI · {date.today().strftime('%d/%m/%Y')}</p>",
    unsafe_allow_html=True,
)
