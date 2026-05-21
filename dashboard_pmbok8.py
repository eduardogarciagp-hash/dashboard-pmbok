# =============================================================================
# DASHBOARD EXECUTIVO DE PORTFÓLIO — PMBOK 8ª EDIÇÃO
# Layout: Arquitetura Executiva para Diretoria
# Multi-Excel + Preenchimento Automático via IA (Claude API)
# =============================================================================
# INSTALAÇÃO:  pip install streamlit pandas plotly openpyxl requests
# EXECUÇÃO:    streamlit run dashboard_pmbok8.py
# =============================================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from datetime import date, timedelta
import io
import json
import requests

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
# 1. FUNÇÕES EVM
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
# 2. MAPA DE COLUNAS MS PROJECT → INTERNO
# ──────────────────────────────────────────────────────────────────────────────
MAPA_COLUNAS = {
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

# ──────────────────────────────────────────────────────────────────────────────
# 3. PARSER DE DATAS PT-BR
#    Converte "22 Dezembro 2025 08:00" → Timestamp
#    Também lida com formato serial do Excel e dd/mm/aaaa
# ──────────────────────────────────────────────────────────────────────────────
MESES_PT = {
    "janeiro":"01","fevereiro":"02","março":"03","marco":"03",
    "abril":"04","maio":"05","junho":"06","julho":"07",
    "agosto":"08","setembro":"09","outubro":"10",
    "novembro":"11","dezembro":"12",
}

def parse_data(valor) -> pd.Timestamp:
    if pd.isna(valor) or str(valor).strip() in ("","nan","None","NaT"):
        return pd.NaT
    # Já é datetime
    if isinstance(valor, (pd.Timestamp, pd.datetime if hasattr(pd,"datetime") else type(None))):
        return pd.Timestamp(valor)
    # Serial numérico do Excel (ex: 46078)
    try:
        n = float(valor)
        if 30000 < n < 80000:
            return pd.Timestamp("1899-12-30") + timedelta(days=int(n))
    except (ValueError, TypeError):
        pass
    # String — tenta PT-BR por extenso
    s = str(valor).strip().lower()
    for nome, num in MESES_PT.items():
        s = s.replace(nome, num)
    partes = s.split()
    if len(partes) >= 3:
        s_data = " ".join(partes[:3])
        try:
            return pd.Timestamp(s_data, dayfirst=True)
        except Exception:
            pass
    # Fallback genérico
    return pd.to_datetime(valor, dayfirst=True, errors="coerce")

def parse_serie(serie: pd.Series) -> pd.Series:
    return serie.apply(parse_data)


# ──────────────────────────────────────────────────────────────────────────────
# 4. NORMALIZAÇÃO DE UM ARQUIVO EXCEL
# ──────────────────────────────────────────────────────────────────────────────
def normalizar_arquivo(arquivo) -> pd.DataFrame:
    conteudo = arquivo.read()
    try:
        # Lê sem forçar dtype — preserva datetime do Excel automaticamente
        df = pd.read_excel(io.BytesIO(conteudo))
    except Exception as e:
        st.error(f"❌ Erro ao ler '{arquivo.name}': {e}")
        return pd.DataFrame()

    df.columns = df.columns.str.strip()
    df.dropna(how="all", inplace=True)

    # Amostra bruta ANTES do mapeamento
    amostra = {c: str(df[c].dropna().iloc[0]) if df[c].notna().any() else "vazio"
               for c in df.columns}

    with st.expander(f"🔍 '{arquivo.name}' — colunas detectadas", expanded=False):
        st.code("\n".join([f"{i+1:02d}. {c}  →  {amostra[c]}"
                           for i, c in enumerate(df.columns)]))

    # Mapeamento
    df.rename(columns={k: v for k, v in MAPA_COLUNAS.items() if k in df.columns},
              inplace=True)

    # Parse de datas — aceita datetime nativo, serial ou string PT-BR
    for col in ["Inicio", "Termino"]:
        if col in df.columns:
            df[col] = parse_serie(df[col])
        else:
            df[col] = pd.NaT

    linhas_ok = df["Inicio"].notna() & df["Termino"].notna()

    if linhas_ok.sum() == 0:
        st.error(
            f"❌ **'{arquivo.name}'**: datas não reconhecidas.\n\n"
            "**Como corrigir:**\n"
            "1. Abra o arquivo no Excel\n"
            "2. Selecione as colunas **Início** e **Término**\n"
            "3. Copie tudo para um **Excel novo** (Colar Especial → Valores)\n"
            "4. Formate as colunas de data como **dd/mm/aaaa**\n"
            "5. Salve e reimporte"
        )
        return pd.DataFrame()

    df = df[linhas_ok].copy()

    # Nome do projeto = nome do arquivo se coluna ausente
    nome_arquivo = arquivo.name.replace(".xlsx","").replace(".xls","").strip()
    if "Projeto" not in df.columns or df["Projeto"].isna().all():
        df["Projeto"] = nome_arquivo
    else:
        df["Projeto"] = df["Projeto"].fillna(nome_arquivo)

    if "Portfolio" not in df.columns:
        df["Portfolio"] = "Portfólio Geral"

    # Defaults
    for col, val in {"Responsavel":"","Causa_Raiz":"","Plano_Acao":"",
                     "Pct_Concluida":0.0,"AC":0.0,"PV":0.0,"EV":0.0,"Recursos":""}.items():
        if col not in df.columns:
            df[col] = val

    if "Termino_Baseline" not in df.columns:
        df["Termino_Baseline"] = df["Termino"]

    # % Concluída
    df["Pct_Concluida"] = pd.to_numeric(df["Pct_Concluida"], errors="coerce").fillna(0)
    if df["Pct_Concluida"].max() > 1:
        df["Pct_Concluida"] = df["Pct_Concluida"] / 100

    # EVM numérico
    for col in ["AC","PV","EV"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Marco: duração zero
    if "Marco" not in df.columns:
        if "Duracao" in df.columns:
            df["Marco"] = df["Duracao"].astype(str).str.strip().isin(
                ["0","0 dias","0d","0 days","0?"])
        else:
            df["Marco"] = False
    else:
        df["Marco"] = df["Marco"].astype(str).str.lower().isin(["sim","true","1","yes"])

    # Tipo por Nível
    if "Nivel" in df.columns:
        df["Nivel"] = pd.to_numeric(df["Nivel"], errors="coerce").fillna(99)
        def tipo(row):
            if row["Marco"]:      return "Marco"
            if row["Nivel"] <= 1: return "Fase"
            return "Tarefa"
        df["Tipo"] = df.apply(tipo, axis=1)
    else:
        df["Tipo"] = df["Marco"].apply(lambda m: "Marco" if m else "Fase")

    df = calcular_evm(df)

    df["Marco_Atrasado"] = (
        df["Marco"] &
        (df["Termino"] > pd.to_datetime(df["Termino_Baseline"], errors="coerce"))
    )
    df["Inicio_str"]           = df["Inicio"].dt.strftime("%Y-%m-%d")
    df["Termino_str"]          = df["Termino"].dt.strftime("%Y-%m-%d")
    df["Termino_Baseline_str"] = pd.to_datetime(
        df["Termino_Baseline"], errors="coerce").dt.strftime("%Y-%m-%d")

    return df


# ──────────────────────────────────────────────────────────────────────────────
# 5. CARREGAMENTO MULTI-ARQUIVO
# ──────────────────────────────────────────────────────────────────────────────
def carregar_multi_excel(arquivos: list) -> pd.DataFrame:
    frames = [normalizar_arquivo(a) for a in arquivos]
    frames = [f for f in frames if not f.empty]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


# ──────────────────────────────────────────────────────────────────────────────
# 6. MOCK DATA
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
# 7. IA — Análise automática via Claude API
# ──────────────────────────────────────────────────────────────────────────────
def gerar_analise_ia(projeto: str, dados: dict) -> dict:
    prompt = f"""Você é especialista sênior em Gestão de Projetos (PMP).
Analise os dados abaixo e gere uma análise executiva objetiva.

PROJETO: {projeto}
SPI: {dados.get('spi','N/A')}
Avanço Físico: {dados.get('pct',0)*100:.0f}%
Desvio: +{dados.get('desvio_dias',0)} dias
Término Previsto: {dados.get('termino','N/A')}
Baseline: {dados.get('baseline','N/A')}
Tarefas críticas: {dados.get('tarefas_criticas','N/A')}

Responda SOMENTE com JSON válido, sem markdown:
{{"impacto":"frase objetiva sobre impacto no negócio (max 200 chars)",
  "causa":"frase objetiva sobre causa raiz (max 200 chars)",
  "plano":"frase objetiva com plano de ação (max 200 chars)"}}"""
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"Content-Type":"application/json"},
            json={"model":"claude-sonnet-4-20250514","max_tokens":400,
                  "messages":[{"role":"user","content":prompt}]},
            timeout=20,
        )
        if resp.status_code == 200:
            texto = resp.json()["content"][0]["text"].strip()
            texto = texto.replace("```json","").replace("```","").strip()
            return json.loads(texto)
    except Exception:
        pass
    return {"impacto":"⚠️ Preencha manualmente.",
            "causa":  "⚠️ Preencha manualmente.",
            "plano":  "⚠️ Preencha manualmente."}


# ──────────────────────────────────────────────────────────────────────────────
# 8. SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    st.markdown("---")
    arquivos = st.file_uploader(
        "📂 Importar Excel(s) do MS Project",
        type=["xlsx","xls"],
        accept_multiple_files=True,
        help="Selecione um ou mais arquivos Excel exportados do MS Project.",
    )
    if arquivos:
        st.markdown(f"**{len(arquivos)} arquivo(s):**")
        for a in arquivos:
            st.markdown(f"• {a.name}")
    st.markdown("---")
    data_ref = st.date_input("📅 Data de Referência", value=date.today())
    logo_url = st.text_input("🖼️ URL do Logo (opcional)", value="")
    st.markdown("---")
    usar_ia  = st.toggle("🤖 Preencher Section 2 com IA", value=False)


# ──────────────────────────────────────────────────────────────────────────────
# 9. CARREGA DADOS
# ──────────────────────────────────────────────────────────────────────────────
if arquivos:
    df_full = carregar_multi_excel(arquivos)
    if df_full.empty:
        st.warning("⚠️ Nenhum dado válido carregado. Verifique as mensagens acima.")
        st.stop()
else:
    df_full = carregar_dados_mock()

projetos_lista = ["Todos"] + sorted(df_full["Projeto"].dropna().unique().tolist())


# ──────────────────────────────────────────────────────────────────────────────
# 10. CABEÇALHO + FILTROS
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
    fonte = f"{len(arquivos)} arquivo(s)" if arquivos else "dados de exemplo"
    st.markdown(
        "<div style='padding-top:4px'>"
        "<span style='font-size:19px;font-weight:800;color:#0D1B2A'>"
        "DASHBOARD DE GOVERNANÇA E VALOR DO PORTFÓLIO</span><br>"
        f"<span style='font-size:12px;color:#8A9BB5'>PMBOK 8ª Ed. · "
        f"Referência: {data_ref.strftime('%d/%m/%Y')} · {fonte}</span></div>",
        unsafe_allow_html=True)

with col_filtros:
    visao_sel   = st.radio("Visão",
                           ["Portfólio Global","Diretoria de Tecnologia","Unidade de Negócio"],
                           horizontal=True, label_visibility="collapsed")
    projeto_sel = st.selectbox("Projeto", projetos_lista, label_visibility="collapsed")

st.markdown("<hr style='margin:10px 0 14px 0;border-color:#E0E4EC'>", unsafe_allow_html=True)

df        = df_full if projeto_sel == "Todos" else df_full[df_full["Projeto"] == projeto_sel]
df_fases  = df[df["Tipo"] == "Fase"]
df_marcos = df[df["Marco"] == True]


# ──────────────────────────────────────────────────────────────────────────────
# 11. BARRA DE SAÚDE HOLÍSTICA
# ──────────────────────────────────────────────────────────────────────────────
spi_saude = df_fases["SPI"].dropna().mean() if not df_fases.empty else None
pct_saude = min(float(spi_saude), 1.0) * 100 if spi_saude else 0
n_crit_s  = int((df_fases["SPI"].dropna() < 0.90).sum())
n_total_s = int(df_fases["SPI"].dropna().count())

if spi_saude and n_total_s > 0:
    if pct_saude >= 95:   cor_s, lbl_s = "#1E7E34", "ESTRATÉGIA SAUDÁVEL"
    elif pct_saude >= 90: cor_s, lbl_s = "#B8860B", "ATENÇÃO — DESVIOS MODERADOS"
    else:                 cor_s, lbl_s = "#C62828", "CRÍTICO — INTERVENÇÃO NECESSÁRIA"
    st.markdown(f"""
    <div style='background:#FFF;border-radius:10px;padding:14px 22px;
                box-shadow:0 1px 6px rgba(0,0,0,0.07);margin-bottom:16px'>
      <div style='display:flex;justify-content:space-between;margin-bottom:8px'>
        <span style='font-size:11px;font-weight:700;color:#8A9BB5;text-transform:uppercase;letter-spacing:.8px'>
          SAÚDE HOLÍSTICA — {visao_sel.upper()}</span>
        <span style='font-size:12px;font-weight:700;color:{cor_s}'>{lbl_s}</span>
      </div>
      <div style='background:#F0F2F6;border-radius:20px;height:10px;overflow:hidden'>
        <div style='background:{cor_s};height:10px;width:{pct_saude:.1f}%;border-radius:20px'></div>
      </div>
      <div style='display:flex;justify-content:space-between;margin-top:6px'>
        <span style='font-size:11px;color:#8A9BB5'>SPI Médio:
          <b style='color:{cor_s}'>{spi_saude:.2f}</b> · {n_crit_s} de {n_total_s} em zona crítica
        </span>
        <span style='font-size:11px;color:#8A9BB5'>{pct_saude:.1f}% da meta</span>
      </div>
    </div>""", unsafe_allow_html=True)
else:
    st.markdown(f"""
    <div style='background:#FFF;border-radius:10px;padding:14px 22px;
                box-shadow:0 1px 6px rgba(0,0,0,0.07);margin-bottom:16px'>
      <div style='display:flex;justify-content:space-between;margin-bottom:8px'>
        <span style='font-size:11px;font-weight:700;color:#8A9BB5;text-transform:uppercase;letter-spacing:.8px'>
          SAÚDE HOLÍSTICA — {visao_sel.upper()}</span>
        <span style='font-size:12px;font-weight:700;color:#8A9BB5'>SEM DADOS EVM</span>
      </div>
      <div style='background:#F0F2F6;border-radius:20px;height:10px;overflow:hidden'>
        <div style='background:#D0D6E0;height:10px;width:100%;border-radius:20px'></div>
      </div>
      <div style='margin-top:6px'>
        <span style='font-size:11px;color:#8A9BB5'>
          Importe Excel com colunas COTA/PV, COTR/EV e Custo Real/AC para habilitar.</span>
      </div>
    </div>""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 12. CARDS KPI
# ──────────────────────────────────────────────────────────────────────────────
n_proj    = df_fases["Projeto"].nunique()
pct_med   = df_fases["Pct_Concluida"].mean() * 100 if not df_fases.empty else 0
spi_med   = df_fases["SPI"].dropna().mean() if not df_fases.empty else None
marcos_at = int(df_marcos["Marco_Atrasado"].sum()) if not df_marcos.empty else 0

def cor_spi(s):
    if s is None or (isinstance(s,float) and np.isnan(s)):
        return "#8A9BB5","N/A","Sem dados EVM"
    if s >= 1.0:  return "#1E7E34",f"{s:.2f} ✅","No prazo"
    if s >= 0.90: return "#B8860B",f"{s:.2f} ⚠️","Atenção"
    return "#C62828",f"{s:.2f} 🔴","Crítico"

c1,c2,c3,c4 = st.columns(4)
cor1,val_spi,lbl_spi = cor_spi(spi_med)
for cw,lbl,val,sub,cor in [
    (c1,"PROJETOS ATIVOS",str(n_proj),"no portfólio","#0D1B2A"),
    (c2,"CONCLUSÃO MÉDIA",f"{pct_med:.1f}%","avanço físico médio","#0D1B2A"),
    (c3,"SPI DO PORTFÓLIO",val_spi,lbl_spi,cor1),
    (c4,"MARCOS EM ATRASO",str(marcos_at),"além da baseline",
     "#C62828" if marcos_at>0 else "#1E7E34"),
]:
    with cw:
        st.markdown(f"""
        <div class="kpi-wrap" style="border-top-color:{cor}">
            <div class="kpi-label">{lbl}</div>
            <div class="kpi-value" style="color:{cor}">{val}</div>
            <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 13. SECTION 1 — ROADMAP EXECUTIVO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-header">📅 SECTION 1 — ROADMAP EXECUTIVO & MARCOS DE VALOR</div>',
            unsafe_allow_html=True)

df_gantt   = df[df["Tipo"] == "Fase"].copy()
df_marcos2 = df[df["Marco"] == True].copy()

if df_gantt.empty:
    st.info("Nenhuma fase detectada. Verifique o arquivo ou use os dados de exemplo.")
else:
    projs  = df_gantt["Projeto"].unique().tolist()
    paleta = ["#1B3A6B","#2E6DA4","#1A7F5A","#7B3FA0","#B85C00","#8B0000","#2C6E7A","#5A4A00"]
    cores_proj = {p: paleta[i % len(paleta)] for i,p in enumerate(projs)}
    fig    = go.Figure()
    hoje_str = date.today().strftime("%Y-%m-%d")
    hoje_dt  = date.today()

    for _, row in df_gantt.iterrows():
        dur     = max((pd.to_datetime(row["Termino_str"]) -
                       pd.to_datetime(row["Inicio_str"])).days, 1)
        spi_val = row["SPI"] if not pd.isna(row.get("SPI",np.nan)) else None
        if spi_val is None:   cor_b = cores_proj.get(row["Projeto"],"#1B3A6B")
        elif spi_val < 0.90:  cor_b = "#C62828"
        elif spi_val < 1.0:   cor_b = "#B8860B"
        else:                 cor_b = cores_proj.get(row["Projeto"],"#1B3A6B")

        pct_txt = f"{row['Pct_Concluida']*100:.0f}%" if row["Pct_Concluida"] > 0 else ""
        nome    = row["Nome da Tarefa"]
        txt     = f"{nome}  {pct_txt}".strip() if pct_txt else nome

        fig.add_trace(go.Bar(
            x=[dur], y=[row["Projeto"]], base=[row["Inicio_str"]],
            orientation="h",
            marker=dict(color=cor_b, opacity=0.88, line=dict(width=0)),
            text=txt, textposition="inside", insidetextanchor="middle",
            textfont=dict(size=11,color="white",family="Arial"),
            hovertemplate=(
                f"<b>{nome}</b><br>"
                f"{row['Inicio_str']} → {row['Termino_str']}<br>"
                f"Avanço: {row['Pct_Concluida']*100:.0f}%"
                + (f"<br>SPI: {spi_val:.2f}" if spi_val else "")
                + "<extra></extra>"),
            showlegend=False,
        ))

    for _, row in df_marcos2.iterrows():
        tdt       = pd.to_datetime(row["Termino_str"]).date()
        concluido = row["Pct_Concluida"] >= 1.0
        atrasado  = bool(row["Marco_Atrasado"])
        if concluido:         cor_m,simb,icone = "#1E7E34","star","📅"
        elif atrasado:        cor_m,simb,icone = "#C62828","circle","🔴"
        elif tdt > hoje_dt:   cor_m,simb,icone = "#8A9BB5","diamond","🔶"
        else:                 cor_m,simb,icone = "#1E7E34","diamond","♦"

        st_txt  = "📅 CONCLUÍDO" if concluido else "🔴 ATRASADO" if atrasado else "🔶 AGENDADO" if tdt>hoje_dt else "♦ NO PRAZO"
        nm_c    = row["Nome da Tarefa"][:22]+"…" if len(row["Nome da Tarefa"])>22 else row["Nome da Tarefa"]
        rotulo  = f"  {icone} {nm_c} ({pd.to_datetime(row['Termino_str']).strftime('%d/%b')})"

        fig.add_trace(go.Scatter(
            x=[row["Termino_str"]], y=[row["Projeto"]],
            mode="markers+text",
            marker=dict(symbol=simb, size=16, color=cor_m, line=dict(width=2,color="white")),
            text=[rotulo], textposition="middle right",
            textfont=dict(size=10,color=cor_m),
            hovertemplate=f"<b>{row['Nome da Tarefa']}</b><br>Baseline: {row['Termino_Baseline_str']}<br>Atual: {row['Termino_str']}<br>{st_txt}<extra></extra>",
            showlegend=False,
        ))

    # Grade trimestral
    dts  = pd.to_datetime(df_gantt["Inicio_str"].tolist()+df_gantt["Termino_str"].tolist())
    dmin = dts.min()-timedelta(days=15)
    dmax = dts.max()+timedelta(days=30)
    for ano in range(dmin.year, dmax.year+2):
        for mes in [1,4,7,10]:
            qt = pd.Timestamp(year=ano,month=mes,day=1)
            if dmin <= qt <= dmax:
                fig.add_shape(type="line",xref="x",yref="paper",
                              x0=qt.strftime("%Y-%m-%d"),x1=qt.strftime("%Y-%m-%d"),
                              y0=0,y1=1,line=dict(color="#E0E4EC",width=1,dash="dot"))
                fig.add_annotation(x=qt.strftime("%Y-%m-%d"),y=1.04,yref="paper",
                                   text=f"Q{(qt.month-1)//3+1} {qt.year}",
                                   showarrow=False,xanchor="center",
                                   font=dict(size=10,color="#8A9BB5"))

    fig.add_shape(type="line",xref="x",yref="paper",
                  x0=hoje_str,x1=hoje_str,y0=0,y1=1,
                  line=dict(color="#4A90D9",width=2,dash="dash"))
    fig.add_annotation(x=hoje_str,y=1.02,yref="paper",text="Hoje",
                       showarrow=False,xanchor="center",font=dict(size=11,color="#4A90D9"))

    fig.update_layout(
        barmode="stack", plot_bgcolor="white", paper_bgcolor="white",
        height=max(300, len(projs)*80+80),
        margin=dict(l=10,r=200,t=45,b=40),
        xaxis=dict(type="date",range=[dmin.strftime("%Y-%m-%d"),dmax.strftime("%Y-%m-%d")],
                   gridcolor="#F8F9FB",tickformat="%d/%b",dtick="M1",
                   tickfont=dict(size=10,color="#B0BAD0"),showline=False,zeroline=False),
        yaxis=dict(autorange="reversed",showgrid=False,
                   tickfont=dict(size=12,color="#0D1B2A",family="Arial")),
        hoverlabel=dict(bgcolor="white",font_size=12),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    st.markdown("""
    <div style='display:flex;gap:20px;flex-wrap:wrap;font-size:11px;color:#6B7A99;padding:8px 4px'>
      <span>📅 <b style='color:#1E7E34'>Concluído</b></span>
      <span>♦ <b style='color:#1E7E34'>No Prazo</b></span>
      <span>🔴 <b style='color:#C62828'>Atrasado</b></span>
      <span>🔶 <b style='color:#8A9BB5'>Futuro</b></span>
      <span><span style='color:#C62828'>■</span> Fase crítica SPI&lt;0.90</span>
      <span><span style='color:#B8860B'>■</span> Atenção</span>
      <span><b style='color:#4A90D9'>— —</b> Hoje</span>
    </div>""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 14. SECTION 2 — GOVERNANÇA DE INCERTEZAS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown(
    '<div class="section-header">⚠️ SECTION 2 — GOVERNANÇA DE INCERTEZAS: PONTOS CRÍTICOS E PLANOS DE AÇÃO'
    '<span style="font-weight:400;font-size:10px;color:#B0BAD0;margin-left:12px">'
    '💡 SPI &lt; 0.95 ou marcos atrasados</span></div>',
    unsafe_allow_html=True)

df_crit = pd.concat([
    df_fases[df_fases["SPI"].notna() & (df_fases["SPI"] < 0.95)],
    df_marcos[df_marcos["Marco_Atrasado"] == True],
]).drop_duplicates(subset=["Projeto"])

if df_crit.empty:
    st.success("✅ Nenhum alerta ativo. Portfólio dentro do limiar de desempenho.")
else:
    st.markdown(f"<p style='font-size:12px;color:#6B7A99;margin-bottom:16px'>"
                f"<b>{len(df_crit)}</b> projeto(s) com alertas ativos.</p>",
                unsafe_allow_html=True)

    if "acoes" not in st.session_state:
        st.session_state["acoes"] = {}

    # Botão IA global
    if usar_ia:
        if st.button("🤖 Analisar todos com IA", use_container_width=False):
            with st.spinner("Analisando com IA..."):
                for _, row in df_crit.iterrows():
                    pk = row["Projeto"]
                    tf = df_fases[df_fases["Projeto"]==pk].nsmallest(3,"SPI")
                    ts = "; ".join(tf["Nome da Tarefa"].tolist()) if not tf.empty else "N/A"
                    spi_v = row.get("SPI",None)
                    dados = {
                        "spi": f"{spi_v:.2f}" if spi_v and not np.isnan(float(spi_v)) else "N/A",
                        "pct": float(row["Pct_Concluida"]),
                        "desvio_dias": int((pd.to_datetime(row["Termino_str"])-
                                           pd.to_datetime(row["Termino_Baseline_str"])).days),
                        "termino": row["Termino_str"],
                        "baseline": row["Termino_Baseline_str"],
                        "sv": float(row.get("SV",0) or 0),
                        "tarefas_criticas": ts,
                    }
                    an = gerar_analise_ia(pk, dados)
                    if pk not in st.session_state["acoes"]:
                        st.session_state["acoes"][pk] = {"impacto":"","causa":"","plano":"","resp":"","prazo":""}
                    st.session_state["acoes"][pk].update(
                        {"impacto":an.get("impacto",""),"causa":an.get("causa",""),"plano":an.get("plano","")})
            st.success("✅ Análise concluída! Revise antes da reunião.")

    for _, row in df_crit.iterrows():
        pk    = row["Projeto"]
        spi   = row.get("SPI",None)
        spi_f = float(spi) if spi is not None and not np.isnan(float(spi)) else None
        spi_s = f"{spi_f:.2f}" if spi_f else "N/A"
        crit  = spi_f is not None and spi_f < 0.90
        cor_n = "#C62828" if crit else "#B8860B"
        bg_n  = "#FDECEA" if crit else "#FFF8E1"
        lbl_n = "ALERTA"  if crit else "ATENÇÃO"

        if pk not in st.session_state["acoes"]:
            st.session_state["acoes"][pk] = {
                "impacto":"","causa":"","plano":"",
                "resp":str(row.get("Responsavel","")),"prazo":""}

        desvio = int((pd.to_datetime(row["Termino_str"])-
                      pd.to_datetime(row["Termino_Baseline_str"])).days)

        st.markdown(f"""
        <div style='background:{bg_n};border-left:4px solid {cor_n};
                    border-radius:0 10px 10px 0;padding:14px 20px;margin-bottom:6px'>
          <div style='display:flex;justify-content:space-between'>
            <span style='font-size:13px;font-weight:700;color:#0D1B2A'>
              [ <span style='color:{cor_n}'>{lbl_n}</span> ] &nbsp; PROJETO: {pk}
            </span>
            <span style='font-size:11px;color:#8A9BB5'>
              SPI: <b style='color:{cor_n}'>{spi_s}</b> &nbsp;·&nbsp;
              Desvio: <b>+{desvio} dias</b> &nbsp;·&nbsp;
              Avanço: <b>{row['Pct_Concluida']*100:.0f}%</b>
            </span>
          </div>
        </div>""", unsafe_allow_html=True)

        with st.expander("▸ Preencher / editar plano de ação", expanded=True):
            if usar_ia:
                if st.button(f"🤖 Analisar '{pk}' com IA", key=f"ia_{pk}"):
                    with st.spinner(f"Analisando {pk}..."):
                        tf = df_fases[df_fases["Projeto"]==pk].nsmallest(3,"SPI")
                        ts = "; ".join(tf["Nome da Tarefa"].tolist()) if not tf.empty else "N/A"
                        an = gerar_analise_ia(pk,{"spi":spi_s,"pct":float(row["Pct_Concluida"]),
                            "desvio_dias":desvio,"termino":row["Termino_str"],
                            "baseline":row["Termino_Baseline_str"],
                            "sv":float(row.get("SV",0) or 0),"tarefas_criticas":ts})
                        st.session_state["acoes"][pk].update(
                            {"impacto":an.get("impacto",""),"causa":an.get("causa",""),"plano":an.get("plano","")})
                    st.rerun()

            fa, fb = st.columns(2)
            with fa:
                st.session_state["acoes"][pk]["impacto"] = st.text_area(
                    "├── 🏢 Impacto no Negócio",
                    value=st.session_state["acoes"][pk]["impacto"], height=90,
                    placeholder="Ex: Risco de paralisação no processamento de relatórios.",
                    key=f"imp_{pk}")
                st.session_state["acoes"][pk]["causa"] = st.text_area(
                    "├── 🔍 Causa Raiz",
                    value=st.session_state["acoes"][pk]["causa"], height=90,
                    placeholder="Ex: Lentidão na validação de acessos de segurança.",
                    key=f"cau_{pk}")
            with fb:
                st.session_state["acoes"][pk]["plano"] = st.text_area(
                    "└── 🎯 Plano de Ação",
                    value=st.session_state["acoes"][pk]["plano"], height=90,
                    placeholder="Ex: Força-tarefa Cyber Security + fornecedor.",
                    key=f"pla_{pk}")
                r1, r2 = st.columns(2)
                with r1:
                    st.session_state["acoes"][pk]["resp"] = st.text_input(
                        "└── 👤 Responsável",
                        value=st.session_state["acoes"][pk]["resp"],
                        placeholder="Ex: Rafael Mechis", key=f"res_{pk}")
                with r2:
                    st.session_state["acoes"][pk]["prazo"] = st.text_input(
                        "└── 📅 Prazo",
                        value=st.session_state["acoes"][pk]["prazo"],
                        placeholder="Ex: 22/Mai", key=f"prz_{pk}")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 15. EXPORTAÇÃO
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-header">📥 EXPORTAR RELATÓRIO EXECUTIVO</div>',
            unsafe_allow_html=True)

def gerar_excel_export(df_base, acoes):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        g = df_base.groupby("Projeto", as_index=False).agg(
            Portfolio=("Portfolio","first"),
            PV=("PV","sum"),EV=("EV","sum"),AC=("AC","sum"),Pct=("Pct_Concluida","mean"))
        g["SPI"] = np.where(g["PV"]>0,(g["EV"]/g["PV"]).round(3),np.nan)
        g["CPI"] = np.where(g["AC"]>0,(g["EV"]/g["AC"]).round(3),np.nan)
        g.to_excel(w, sheet_name="EVM_Consolidado", index=False)

        rows = [{"Projeto":p,"Impacto":d["impacto"],"Causa Raiz":d["causa"],
                 "Plano":d["plano"],"Responsável":d["resp"],"Prazo":d.get("prazo","")}
                for p,d in acoes.items()]
        if rows:
            pd.DataFrame(rows).to_excel(w, sheet_name="Planos_de_Acao", index=False)

        dm = df_base[df_base["Marco"]==True]
        if not dm.empty:
            dm2 = dm[["Portfolio","Projeto","Nome da Tarefa",
                       "Termino_str","Termino_Baseline_str","Marco_Atrasado"]].copy()
            dm2.columns = ["Portfólio","Projeto","Marco","Término","Baseline","Atrasado"]
            dm2.to_excel(w, sheet_name="Marcos", index=False)
    return out.getvalue()

col_e, _ = st.columns([1,4])
with col_e:
    if st.button("📥 Gerar Excel Executivo", use_container_width=True):
        xls = gerar_excel_export(df, st.session_state.get("acoes",{}))
        st.download_button(
            "⬇️ Baixar Relatório (.xlsx)", data=xls,
            file_name=f"governanca_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown(
    f"<p style='text-align:center;color:#C0C8D8;font-size:11px;margin-top:20px'>"
    f"Dashboard Executivo · PMBOK® 8ª Edição · PMI · {date.today().strftime('%d/%m/%Y')}</p>",
    unsafe_allow_html=True)
