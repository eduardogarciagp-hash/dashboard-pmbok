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
    .ia-badge {
        display:inline-block; background:#EEF2FF; color:#3730A3;
        font-size:10px; font-weight:700; padding:2px 8px;
        border-radius:10px; margin-left:8px; vertical-align:middle;
    }
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
# 2. PIPELINE DE NORMALIZAÇÃO — 1 arquivo Excel → DataFrame padronizado
# ──────────────────────────────────────────────────────────────────────────────
MAPA_COLUNAS = {
    # Datas
    "Início":"Inicio","Start":"Inicio",
    "Término":"Termino","Finish":"Termino",
    # Tarefa
    "Nome":"Nome da Tarefa","Name":"Nome da Tarefa",
    "Task Name":"Nome da Tarefa","Nome da Tarefa":"Nome da Tarefa",
    # Estrutura
    "Nível da estrutura de tópicos":"Nivel","Outline Level":"Nivel",
    # Duração
    "Duração":"Duracao","Duration":"Duracao",
    # ID
    "Id":"ID",
    # Recursos
    "Nomes dos Recursos":"Recursos","Resource Names":"Recursos",
    # % avanço
    "% concluída":"Pct_Concluida","% Concluída":"Pct_Concluida",
    "% Complete":"Pct_Concluida","% Completo":"Pct_Concluida",
    # EVM
    "Custo Real (CR)":"AC","Custo Real":"AC","Actual Cost":"AC","ACWP":"AC","AC":"AC",
    "COTA":"PV","Baseline Cost":"PV","BCWS":"PV","PV":"PV",
    "COTR":"EV","Earned Value":"EV","BCWP":"EV","EV":"EV",
    # Marco
    "Marco":"Marco","Milestone":"Marco",
    # Projeto / Portfolio
    "Projeto":"Projeto","Project":"Projeto",
    "Portfólio":"Portfolio","Portfolio":"Portfolio",
}

def normalizar_arquivo(arquivo) -> pd.DataFrame:
    """
    Lê um único arquivo Excel e retorna um DataFrame normalizado
    com todas as colunas internas padronizadas.
    O nome do projeto é derivado do nome do arquivo se não houver coluna Projeto.
    """
    try:
        # Lê sem conversão automática de datas para inspecionar valores brutos
        df = pd.read_excel(arquivo, dtype=str)
    except Exception as e:
        st.warning(f"⚠️ Erro ao ler '{arquivo.name}': {e}")
        return pd.DataFrame()

    df.columns = df.columns.str.strip()

    # Diagnóstico colapsado
    with st.expander(f"🔍 Diagnóstico — '{arquivo.name}'", expanded=False):
        st.code("\n".join([
            f"{i+1:02d}. {c}  →  ex: {str(df[c].dropna().iloc[0]) if df[c].notna().any() else 'vazio'}"
            for i, c in enumerate(df.columns)
        ]))
        st.dataframe(df.head(3), use_container_width=True)

    # Renomeia pelo mapa
    df.rename(columns={k: v for k, v in MAPA_COLUNAS.items() if k in df.columns}, inplace=True)

    # ── Tenta parsear datas em TODAS as colunas candidatas ───────────────────
    # O MS Project às vezes exporta datas como serial numérico do Excel (ex: 46000)
    # ou como string "qui 01/01/2026". Tentamos todos os formatos.
    def parse_data_coluna(serie: pd.Series) -> pd.Series:
        # Tenta direto
        result = pd.to_datetime(serie, errors="coerce", dayfirst=True)
        # Se falhou muito, tenta como serial numérico Excel (dias desde 1900-01-01)
        if result.notna().sum() < len(serie) * 0.3:
            try:
                numerico = pd.to_numeric(serie, errors="coerce")
                serial_ok = numerico.notna() & (numerico > 1000) & (numerico < 100000)
                if serial_ok.sum() > len(serie) * 0.3:
                    from datetime import datetime
                    result = numerico.apply(
                        lambda x: (datetime(1899, 12, 30) + timedelta(days=int(x)))
                        if pd.notna(x) and 1000 < x < 100000 else pd.NaT
                    )
            except Exception:
                pass
        return result

    # Remove linhas completamente vazias
    df.dropna(how="all", inplace=True)

    # Nome do projeto = coluna Projeto (se existir) ou nome do arquivo
    nome_arquivo = arquivo.name.replace(".xlsx","").replace(".xls","").strip()
    if "Projeto" not in df.columns or df["Projeto"].isna().all():
        df["Projeto"] = nome_arquivo
    else:
        # Preenche células vazias com o nome do arquivo
        df["Projeto"] = df["Projeto"].fillna(nome_arquivo)

    if "Portfolio" not in df.columns:
        df["Portfolio"] = "Portfólio Geral"

    # Defaults para colunas ausentes
    defaults = {
        "Responsavel":"","Causa_Raiz":"","Plano_Acao":"",
        "Pct_Concluida":0.0,"AC":0.0,"PV":0.0,"EV":0.0,"Recursos":"",
    }
    for col, val in defaults.items():
        if col not in df.columns:
            df[col] = val

    # ── Datas — busca inteligente se colunas não foram mapeadas ─────────────
    def detectar_coluna_data(df_: pd.DataFrame, col_interna: str,
                              candidatos_extras: list) -> pd.Series:
        # 1. Nome interno já mapeado — tenta parsear mesmo se parece vazio
        if col_interna in df_.columns:
            result = parse_data_coluna(df_[col_interna])
            if result.notna().sum() > len(df_) * 0.3:
                return result

        # 2. Candidatos extras
        for c in candidatos_extras:
            if c in df_.columns:
                result = parse_data_coluna(df_[c])
                if result.notna().sum() > len(df_) * 0.3:
                    return result

        # 3. Varredura automática em todas as colunas
        melhor_serie  = pd.Series([pd.NaT] * len(df_))
        melhor_count  = 0
        for c in df_.columns:
            if c in ["ID","Duracao","Nivel","Pct_Concluida","AC","PV","EV"]:
                continue  # pula colunas sabidamente não-data
            result = parse_data_coluna(df_[c])
            cnt = result.notna().sum()
            if cnt > melhor_count:
                melhor_count = cnt
                melhor_serie = result

        return melhor_serie if melhor_count > len(df_) * 0.3 else pd.Series([pd.NaT] * len(df_))

    df["Inicio"]  = detectar_coluna_data(df, "Inicio",
                        ["Início","Start","Data Início","Data de Início",
                         "data_inicio","inicio","DT_INICIO"])
    df["Termino"] = detectar_coluna_data(df, "Termino",
                        ["Término","Finish","Data Término","Data de Término",
                         "data_termino","termino","DT_TERMINO"])

    linhas_validas = df["Inicio"].notna() & df["Termino"].notna()
    if linhas_validas.sum() == 0:
        # Mostra diagnóstico detalhado para ajudar o usuário
        with st.expander(f"❌ Diagnóstico de erro — '{arquivo.name}'", expanded=True):
            st.error("Não foi possível detectar colunas de data. "
                     "Veja as colunas encontradas abaixo e informe ao suporte.")
            st.code("\n".join([f"{i+1:02d}. {c}  →  ex: {str(df[c].dropna().iloc[0]) if df[c].notna().any() else 'vazio'}"
                               for i, c in enumerate(df.columns)]))
        return pd.DataFrame()

    df = df[linhas_validas].copy()

    if "Termino_Baseline" not in df.columns:
        df["Termino_Baseline"] = df["Termino"]

    # % Concluída — normaliza para 0..1
    df["Pct_Concluida"] = pd.to_numeric(df["Pct_Concluida"], errors="coerce").fillna(0)
    if df["Pct_Concluida"].max() > 1:
        df["Pct_Concluida"] = df["Pct_Concluida"] / 100

    # Marco: duração zero ou coluna explícita
    if "Marco" not in df.columns:
        if "Duracao" in df.columns:
            df["Marco"] = df["Duracao"].astype(str).str.strip().isin(
                ["0","0 dias","0d","0 days"])
        else:
            df["Marco"] = False
    else:
        df["Marco"] = df["Marco"].astype(str).str.lower().isin(
            ["sim","true","1","yes"])

    # Tipo por Nível de estrutura
    if "Nivel" in df.columns:
        df["Nivel"] = pd.to_numeric(df["Nivel"], errors="coerce").fillna(99)
        def definir_tipo(row):
            if row["Marco"]:      return "Marco"
            if row["Nivel"] <= 1: return "Fase"
            return "Tarefa"
        df["Tipo"] = df.apply(definir_tipo, axis=1)
    else:
        df["Tipo"] = df["Marco"].apply(lambda m: "Marco" if m else "Fase")

    # EVM
    df = calcular_evm(df)

    # Marco atrasado
    df["Marco_Atrasado"] = (
        df["Marco"] &
        (pd.to_datetime(df["Termino"], errors="coerce") >
         pd.to_datetime(df["Termino_Baseline"], errors="coerce"))
    )

    # Strings de data para Plotly
    df["Inicio_str"]           = df["Inicio"].dt.strftime("%Y-%m-%d")
    df["Termino_str"]          = df["Termino"].dt.strftime("%Y-%m-%d")
    df["Termino_Baseline_str"] = pd.to_datetime(
        df["Termino_Baseline"], errors="coerce").dt.strftime("%Y-%m-%d")

    return df


# ──────────────────────────────────────────────────────────────────────────────
# 3. MOCK DATA
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
# 4. CARREGAMENTO MULTI-ARQUIVO
# ──────────────────────────────────────────────────────────────────────────────
def carregar_multi_excel(arquivos: list) -> pd.DataFrame:
    """
    Recebe lista de arquivos uploadados.
    Normaliza cada um e concatena num único DataFrame.
    """
    frames = []
    for arq in arquivos:
        df_norm = normalizar_arquivo(arq)
        if not df_norm.empty:
            frames.append(df_norm)

    if not frames:
        return pd.DataFrame()

    df_total = pd.concat(frames, ignore_index=True)

    # Garante unicidade de projetos duplicados entre arquivos distintos
    # (ex: dois arquivos com coluna Projeto = "Alpha" → distingue pelo nome do arquivo)
    return df_total


# ──────────────────────────────────────────────────────────────────────────────
# 5. IA — Análise automática de projetos críticos via Claude API
# ──────────────────────────────────────────────────────────────────────────────
def gerar_analise_ia(projeto: str, dados_projeto: dict) -> dict:
    """
    Envia os dados do projeto crítico para a API do Claude e retorna
    os campos de Impacto, Causa Raiz e Plano de Ação preenchidos automaticamente.
    """
    prompt = f"""Você é um especialista sênior em Gestão de Projetos (PMP) e deve analisar 
os dados abaixo de um projeto com desvio crítico de prazo e gerar uma análise executiva objetiva.

DADOS DO PROJETO:
- Nome: {projeto}
- SPI (Índice de Desempenho de Prazo): {dados_projeto.get('spi', 'N/A')}
- Avanço Físico: {dados_projeto.get('pct', 0)*100:.0f}%
- Desvio em dias: +{dados_projeto.get('desvio_dias', 0)} dias
- Término Previsto: {dados_projeto.get('termino', 'N/A')}
- Baseline de Término: {dados_projeto.get('baseline', 'N/A')}
- Tarefas com maior atraso: {dados_projeto.get('tarefas_criticas', 'N/A')}
- Variância de Prazo (SV): R$ {dados_projeto.get('sv', 0):,.0f}

Responda SOMENTE com um JSON válido no formato abaixo, sem texto adicional, sem markdown:
{{
  "impacto": "Uma frase objetiva sobre o impacto no negócio (máx 200 chars)",
  "causa": "Uma frase objetiva sobre a provável causa raiz (máx 200 chars)",
  "plano": "Uma frase objetiva com o plano de ação recomendado (máx 200 chars)"
}}"""

    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"Content-Type": "application/json"},
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 500,
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=20,
        )
        if resp.status_code == 200:
            texto = resp.json()["content"][0]["text"].strip()
            # Remove possíveis blocos markdown ```json ... ```
            texto = texto.replace("```json","").replace("```","").strip()
            return json.loads(texto)
    except Exception:
        pass

    # Fallback se a API não responder
    return {
        "impacto": "⚠️ Análise automática indisponível. Preencha manualmente.",
        "causa":   "⚠️ Análise automática indisponível. Preencha manualmente.",
        "plano":   "⚠️ Análise automática indisponível. Preencha manualmente.",
    }


# ──────────────────────────────────────────────────────────────────────────────
# 6. SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    st.markdown("---")

    # Upload múltiplo
    arquivos = st.file_uploader(
        "📂 Importar Excel(s) do MS Project",
        type=["xlsx","xls"],
        accept_multiple_files=True,   # ← múltiplos arquivos
        help="Selecione um ou mais arquivos Excel exportados do MS Project.",
    )

    if arquivos:
        st.markdown(f"**{len(arquivos)} arquivo(s) carregado(s):**")
        for arq in arquivos:
            st.markdown(f"• {arq.name}")

    st.markdown("---")
    data_ref = st.date_input("📅 Data de Referência", value=date.today())
    logo_url = st.text_input("🖼️ URL do Logo (opcional)", value="")
    st.markdown("---")
    st.markdown("**🤖 Análise por IA**")
    usar_ia = st.toggle("Preencher Section 3 com IA", value=False,
                        help="Usa a API do Claude para gerar automaticamente "
                             "Impacto, Causa Raiz e Plano de Ação dos projetos críticos.")


# ──────────────────────────────────────────────────────────────────────────────
# 7. CARREGA DADOS
# ──────────────────────────────────────────────────────────────────────────────
if arquivos:
    df_full = carregar_multi_excel(arquivos)
    if df_full.empty:
        st.error("Nenhum dado válido encontrado nos arquivos. Verifique o formato.")
        st.stop()

    # Diagnóstico colapsado
    with st.expander(f"🔍 Diagnóstico: {len(arquivos)} arquivo(s) carregado(s)", expanded=False):
        for arq in arquivos:
            st.markdown(f"**{arq.name}**")
            try:
                df_diag = pd.read_excel(arq)
                df_diag.columns = df_diag.columns.str.strip()
                st.code("\n".join([f"{i+1:02d}. {c}" for i, c in enumerate(df_diag.columns)]))
                st.dataframe(df_diag.head(2), use_container_width=True)
            except Exception as e:
                st.warning(f"Erro: {e}")
else:
    df_full = carregar_dados_mock()

projetos_lista = ["Todos"] + sorted(df_full["Projeto"].dropna().unique().tolist())


# ──────────────────────────────────────────────────────────────────────────────
# 8. CABEÇALHO + FILTROS
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
        f"Referência: {data_ref.strftime('%d/%m/%Y')} · "
        f"{len(arquivos) if arquivos else 'Mock'} fonte(s) de dados</span></div>",
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
# 9. BARRA DE SAÚDE HOLÍSTICA
# ──────────────────────────────────────────────────────────────────────────────
spi_saude = df_fases["SPI"].dropna().mean() if not df_fases.empty else None
pct_saude = min(float(spi_saude), 1.0) * 100 if spi_saude is not None else 0
n_crit_s  = int((df_fases["SPI"].dropna() < 0.90).sum())
n_total_s = int(df_fases["SPI"].dropna().count())

if spi_saude is None or n_total_s == 0:
    cor_saude = "#8A9BB5"; label_saude = "SEM DADOS EVM"
elif pct_saude >= 95:
    cor_saude = "#1E7E34"; label_saude = "ESTRATÉGIA SAUDÁVEL"
elif pct_saude >= 90:
    cor_saude = "#B8860B"; label_saude = "ATENÇÃO — DESVIOS MODERADOS"
else:
    cor_saude = "#C62828"; label_saude = "CRÍTICO — INTERVENÇÃO NECESSÁRIA"

if spi_saude and n_total_s > 0:
    st.markdown(f"""
    <div style='background:#FFFFFF;border-radius:10px;padding:14px 22px;
                box-shadow:0 1px 6px rgba(0,0,0,0.07);margin-bottom:16px'>
      <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>
        <span style='font-size:11px;font-weight:700;color:#8A9BB5;text-transform:uppercase;letter-spacing:.8px'>
          SAÚDE HOLÍSTICA DA ESTRATÉGIA — {visao_sel.upper()}
        </span>
        <span style='font-size:12px;font-weight:700;color:{cor_saude}'>{label_saude}</span>
      </div>
      <div style='background:#F0F2F6;border-radius:20px;height:10px;width:100%;overflow:hidden'>
        <div style='background:{cor_saude};height:10px;width:{pct_saude:.1f}%;border-radius:20px'></div>
      </div>
      <div style='display:flex;justify-content:space-between;margin-top:6px'>
        <span style='font-size:11px;color:#8A9BB5'>
          SPI Médio: <b style='color:{cor_saude}'>{spi_saude:.2f}</b>
          &nbsp;·&nbsp; {n_crit_s} de {n_total_s} projetos em zona crítica
        </span>
        <span style='font-size:11px;color:#8A9BB5'>{pct_saude:.1f}% da meta de entrega</span>
      </div>
    </div>""", unsafe_allow_html=True)
else:
    st.markdown(f"""
    <div style='background:#FFFFFF;border-radius:10px;padding:14px 22px;
                box-shadow:0 1px 6px rgba(0,0,0,0.07);margin-bottom:16px'>
      <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>
        <span style='font-size:11px;font-weight:700;color:#8A9BB5;text-transform:uppercase;letter-spacing:.8px'>
          SAÚDE HOLÍSTICA DA ESTRATÉGIA — {visao_sel.upper()}
        </span>
        <span style='font-size:12px;font-weight:700;color:#8A9BB5'>SEM DADOS EVM</span>
      </div>
      <div style='background:#F0F2F6;border-radius:20px;height:10px;width:100%;overflow:hidden'>
        <div style='background:#D0D6E0;height:10px;width:100%;border-radius:20px'></div>
      </div>
      <div style='margin-top:6px'>
        <span style='font-size:11px;color:#8A9BB5'>
          Importe Excel com colunas EVM (COTA/PV, COTR/EV, Custo Real/AC) para habilitar.
        </span>
      </div>
    </div>""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 10. CARDS KPI
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
# 11. SECTION 1 — ROADMAP EXECUTIVO
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

        label_pct   = f"{row['Pct_Concluida']*100:.0f}%" if row["Pct_Concluida"] > 0 else ""
        nome_fase   = row["Nome da Tarefa"]
        texto_barra = f"{nome_fase}  {label_pct}".strip() if label_pct else nome_fase
        spi_hover   = f"<br>SPI: {spi_val:.2f}" if spi_val else ""

        fig.add_trace(go.Bar(
            x=[dur], y=[row["Projeto"]],
            base=[row["Inicio_str"]],
            orientation="h",
            marker=dict(color=cor_barra, opacity=0.88, line=dict(width=0)),
            text=texto_barra,
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

    for _, row in df_marcos2.iterrows():
        termino_dt = pd.to_datetime(row["Termino_str"]).date()
        concluido  = row["Pct_Concluida"] >= 1.0
        atrasado   = bool(row["Marco_Atrasado"])
        y_proj     = row["Projeto"]

        if concluido:
            cor_m2 = "#1E7E34"; simbolo = "star";    label_icone = "📅"
        elif atrasado:
            cor_m2 = "#C62828"; simbolo = "circle";  label_icone = "🔴"
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
        data_fmt   = pd.to_datetime(row["Termino_str"]).strftime("%d/%b")
        nome_curto = row["Nome da Tarefa"][:22] + "…" if len(row["Nome da Tarefa"]) > 22 else row["Nome da Tarefa"]
        rotulo     = f"  {label_icone} {nome_curto} ({data_fmt})"

        fig.add_trace(go.Scatter(
            x=[row["Termino_str"]], y=[y_proj],
            mode="markers+text",
            marker=dict(symbol=simbolo, size=16, color=cor_m2,
                        line=dict(width=2, color="white")),
            text=[rotulo],
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

    # Grade trimestral
    todas_datas = pd.to_datetime(
        df_gantt["Inicio_str"].tolist() + df_gantt["Termino_str"].tolist())
    data_min = todas_datas.min() - timedelta(days=15)
    data_max = todas_datas.max() + timedelta(days=30)
    trimestres = []
    for ano in range(data_min.year, data_max.year + 2):
        for mes in [1, 4, 7, 10]:
            dt = pd.Timestamp(year=ano, month=mes, day=1)
            if data_min <= dt <= data_max:
                trimestres.append(dt)

    for qt in trimestres:
        fig.add_shape(type="line", xref="x", yref="paper",
                      x0=qt.strftime("%Y-%m-%d"), x1=qt.strftime("%Y-%m-%d"),
                      y0=0, y1=1, line=dict(color="#E0E4EC", width=1, dash="dot"))
        fig.add_annotation(
            x=qt.strftime("%Y-%m-%d"), y=1.04, yref="paper",
            text=f"Q{(qt.month-1)//3+1} {qt.year}",
            showarrow=False, xanchor="center",
            font=dict(size=10, color="#8A9BB5", family="Arial"))

    fig.add_shape(type="line", xref="x", yref="paper",
                  x0=hoje_str, x1=hoje_str, y0=0, y1=1,
                  line=dict(color="#4A90D9", width=2, dash="dash"))
    fig.add_annotation(x=hoje_str, y=1.02, yref="paper", text="Hoje",
                       showarrow=False, xanchor="center",
                       font=dict(size=11, color="#4A90D9"))

    fig.update_layout(
        barmode="stack",
        plot_bgcolor="white", paper_bgcolor="white",
        height=max(300, len(projetos_gantt) * 80 + 80),
        margin=dict(l=10, r=200, t=45, b=40),
        xaxis=dict(
            type="date",
            range=[data_min.strftime("%Y-%m-%d"), data_max.strftime("%Y-%m-%d")],
            gridcolor="#F8F9FB", tickformat="%d/%b", dtick="M1",
            tickfont=dict(size=10, color="#B0BAD0"),
            showline=False, zeroline=False,
        ),
        yaxis=dict(autorange="reversed", showgrid=False,
                   tickfont=dict(size=12, color="#0D1B2A", family="Arial")),
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
# 12. SECTION 2 — GOVERNANÇA DE INCERTEZAS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown(
    '<div class="section-header">⚠️ SECTION 2 — GOVERNANÇA DE INCERTEZAS: PONTOS CRÍTICOS E PLANOS DE AÇÃO'
    '<span style="font-weight:400;font-size:10px;color:#B0BAD0;margin-left:12px">'
    '💡 Exibe apenas projetos com SPI &lt; 0.95 ou marcos críticos atrasados</span></div>',
    unsafe_allow_html=True,
)

df_crit = pd.concat([
    df_fases[df_fases["SPI"].notna() & (df_fases["SPI"] < 0.95)],
    df_marcos[df_marcos["Marco_Atrasado"] == True],
]).drop_duplicates(subset=["Projeto"])

if df_crit.empty:
    st.success("✅ Nenhum alerta ativo. Todos os projetos operam dentro do limiar de desempenho.")
else:
    st.markdown(
        f"<p style='font-size:12px;color:#6B7A99;margin-bottom:16px'>"
        f"<b>{len(df_crit)}</b> projeto(s) com alertas ativos nesta data de referência.</p>",
        unsafe_allow_html=True,
    )

    if "acoes" not in st.session_state:
        st.session_state["acoes"] = {}

    # Botão de análise IA global
    if usar_ia:
        if st.button("🤖 Gerar análise automática com IA para todos os projetos críticos",
                     use_container_width=False):
            with st.spinner("Analisando projetos com IA..."):
                for _, row in df_crit.iterrows():
                    pk = row["Projeto"]
                    spi_v = row.get("SPI", None)

                    # Tarefas mais críticas do projeto
                    tarefas_df = df_fases[df_fases["Projeto"] == pk].nsmallest(3, "SPI")
                    tarefas_str = "; ".join(tarefas_df["Nome da Tarefa"].tolist()) if not tarefas_df.empty else "N/A"

                    dados = {
                        "spi":          f"{spi_v:.2f}" if spi_v and not np.isnan(float(spi_v)) else "N/A",
                        "pct":          float(row["Pct_Concluida"]),
                        "desvio_dias":  int((pd.to_datetime(row["Termino_str"]) -
                                            pd.to_datetime(row["Termino_Baseline_str"])).days),
                        "termino":      row["Termino_str"],
                        "baseline":     row["Termino_Baseline_str"],
                        "sv":           float(row.get("SV", 0) or 0),
                        "tarefas_criticas": tarefas_str,
                    }
                    analise = gerar_analise_ia(pk, dados)

                    if pk not in st.session_state["acoes"]:
                        st.session_state["acoes"][pk] = {
                            "impacto":"","causa":"","plano":"",
                            "resp": str(row.get("Responsavel","")), "prazo":"",
                        }
                    st.session_state["acoes"][pk]["impacto"] = analise.get("impacto","")
                    st.session_state["acoes"][pk]["causa"]   = analise.get("causa","")
                    st.session_state["acoes"][pk]["plano"]   = analise.get("plano","")
            st.success("✅ Análise IA concluída! Revise os campos abaixo antes da reunião.")

    # Renderiza cards por projeto crítico
    for _, row in df_crit.iterrows():
        pk  = row["Projeto"]
        spi = row.get("SPI", None)
        spi_f   = float(spi) if (spi is not None and not np.isnan(float(spi))) else None
        spi_str = f"{spi_f:.2f}" if spi_f is not None else "N/A"

        eh_critico = spi_f is not None and spi_f < 0.90
        nivel_label = "ALERTA" if eh_critico else "ATENÇÃO"
        cor_nivel   = "#C62828" if eh_critico else "#B8860B"
        bg_nivel    = "#FDECEA" if eh_critico else "#FFF8E1"
        borda_nivel = "#C62828" if eh_critico else "#B8860B"

        if pk not in st.session_state["acoes"]:
            st.session_state["acoes"][pk] = {
                "impacto":"","causa":"","plano":"",
                "resp": str(row.get("Responsavel","")), "prazo":"",
            }

        desvio_dias = int((
            pd.to_datetime(row["Termino_str"]) -
            pd.to_datetime(row["Termino_Baseline_str"])
        ).days)

        st.markdown(f"""
        <div style='background:{bg_nivel};border-left:4px solid {borda_nivel};
                    border-radius:0 10px 10px 0;padding:14px 20px;margin-bottom:6px'>
          <div style='display:flex;justify-content:space-between;align-items:center'>
            <span style='font-size:13px;font-weight:700;color:#0D1B2A'>
              [ <span style='color:{cor_nivel}'>{nivel_label}</span> ] &nbsp; PROJETO: {pk}
            </span>
            <span style='font-size:11px;color:#8A9BB5'>
              SPI: <b style='color:{cor_nivel}'>{spi_str}</b>
              &nbsp;·&nbsp; Desvio: <b>+{desvio_dias} dias</b>
              &nbsp;·&nbsp; Avanço: <b>{row["Pct_Concluida"]*100:.0f}%</b>
            </span>
          </div>
        </div>""", unsafe_allow_html=True)

        with st.expander("▸ Preencher / editar plano de ação", expanded=True):
            # Botão IA individual por projeto
            if usar_ia:
                if st.button(f"🤖 Analisar '{pk}' com IA", key=f"ia_{pk}"):
                    with st.spinner(f"Analisando {pk}..."):
                        tarefas_df  = df_fases[df_fases["Projeto"] == pk].nsmallest(3,"SPI")
                        tarefas_str = "; ".join(tarefas_df["Nome da Tarefa"].tolist()) if not tarefas_df.empty else "N/A"
                        dados = {
                            "spi": spi_str, "pct": float(row["Pct_Concluida"]),
                            "desvio_dias": desvio_dias,
                            "termino": row["Termino_str"], "baseline": row["Termino_Baseline_str"],
                            "sv": float(row.get("SV",0) or 0), "tarefas_criticas": tarefas_str,
                        }
                        analise = gerar_analise_ia(pk, dados)
                        st.session_state["acoes"][pk]["impacto"] = analise.get("impacto","")
                        st.session_state["acoes"][pk]["causa"]   = analise.get("causa","")
                        st.session_state["acoes"][pk]["plano"]   = analise.get("plano","")
                    st.rerun()

            fa, fb = st.columns(2)
            with fa:
                st.session_state["acoes"][pk]["impacto"] = st.text_area(
                    "├── 🏢 Impacto no Negócio",
                    value=st.session_state["acoes"][pk]["impacto"], height=90,
                    placeholder="Ex: Risco de paralisação no processamento de relatórios.",
                    key=f"imp_{pk}")
                st.session_state["acoes"][pk]["causa"] = st.text_area(
                    "├── 🔍 Causa Raiz do Desvio",
                    value=st.session_state["acoes"][pk]["causa"], height=90,
                    placeholder="Ex: Lentidão na validação de acessos de segurança.",
                    key=f"cau_{pk}")
            with fb:
                st.session_state["acoes"][pk]["plano"] = st.text_area(
                    "└── 🎯 Plano de Ação Proposto",
                    value=st.session_state["acoes"][pk]["plano"], height=90,
                    placeholder="Ex: Força-tarefa entre Cyber Security e fornecedor.",
                    key=f"pla_{pk}")
                cr1, cr2 = st.columns(2)
                with cr1:
                    st.session_state["acoes"][pk]["resp"] = st.text_input(
                        "    └── 👤 Responsável",
                        value=st.session_state["acoes"][pk]["resp"],
                        placeholder="Ex: Rafael Mechis", key=f"res_{pk}")
                with cr2:
                    st.session_state["acoes"][pk]["prazo"] = st.text_input(
                        "    └── 📅 Prazo",
                        value=st.session_state["acoes"][pk]["prazo"],
                        placeholder="Ex: 22/Mai", key=f"prz_{pk}")

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 13. EXPORTAÇÃO
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

        rows_ac = [
            {"Projeto":p,"Impacto":d["impacto"],"Causa Raiz":d["causa"],
             "Plano de Ação":d["plano"],"Responsável":d["resp"],"Prazo":d.get("prazo","")}
            for p,d in acoes.items()
        ]
        if rows_ac:
            pd.DataFrame(rows_ac).to_excel(writer, sheet_name="Planos_de_Acao", index=False)

        df_m = df_base[df_base["Marco"]==True]
        if not df_m.empty:
            dm = df_m[["Portfolio","Projeto","Nome da Tarefa",
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
