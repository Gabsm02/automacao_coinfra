"""
Dashboard para visualizar o histórico gravado na tabela `historico_tas_infra`
do MariaDB (a mesma tabela alimentada pelo Coinfra.py).

Como rodar:
    pip install streamlit pandas sqlalchemy pymysql plotly python-dotenv
    python -m streamlit run dash.py

O dashboard lê as configurações de conexão do mesmo arquivo .env usado
pelo Coinfra.py (DB_HOST, DB_PORT, DB_USER, DB_SENHA, DB_NOME, DB_TABELA).
"""

import os
import sys
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from dotenv import load_dotenv
from sqlalchemy import create_engine

# ==========================================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================================

st.set_page_config(
    page_title="Dashboard Coinfra",
    page_icon="📊",
    layout="wide",
)

load_dotenv()


def obter_env(nome, obrigatorio=True, padrao=None):
    valor = os.getenv(nome, padrao)
    if obrigatorio and not valor:
        st.error(
            f"❌ A variável '{nome}' não foi encontrada no arquivo .env. "
            f"Confira se ele existe na mesma pasta deste dashboard."
        )
        st.stop()
    return valor


DB_HOST = obter_env("DB_HOST")
DB_PORT = int(obter_env("DB_PORT", padrao="3306"))
DB_USER = obter_env("DB_USER")
DB_SENHA = obter_env("DB_SENHA")
DB_NOME = obter_env("DB_NOME")
DB_TABELA = obter_env("DB_TABELA")


# ==========================================================
# CONEXÃO E CARGA DE DADOS
# ==========================================================

@st.cache_resource
def obter_engine():
    string_conexao = f"mysql+pymysql://{DB_USER}:{DB_SENHA}@{DB_HOST}:{DB_PORT}/{DB_NOME}"
    return create_engine(string_conexao)


@st.cache_data(ttl=300)  # os dados ficam em cache por 5 minutos
def carregar_dados():
    engine = obter_engine()
    query = f"SELECT * FROM `{DB_TABELA}`"
    df = pd.read_sql(query, con=engine)

    if "data_execucao" in df.columns:
        df["data_execucao"] = pd.to_datetime(df["data_execucao"], errors="coerce")

    return df


try:
    df = carregar_dados()
except Exception as erro:
    st.error(f"❌ Não foi possível conectar ao MariaDB: {erro}")
    st.stop()

if df.empty:
    st.warning("A tabela ainda não tem nenhum registro.")
    st.stop()


# ==========================================================
# BARRA LATERAL - FILTROS
# ==========================================================

st.sidebar.header("🔍 Filtros")

# Filtro de período (data_execucao)
if "data_execucao" in df.columns and df["data_execucao"].notna().any():
    data_min = df["data_execucao"].min().date()
    data_max = df["data_execucao"].max().date()
    periodo = st.sidebar.date_input(
        "Período de execução",
        value=(data_min, data_max),
        min_value=data_min,
        max_value=data_max,
    )
else:
    periodo = None

# Filtro de UF
if "UF" in df.columns:
    ufs_disponiveis = sorted(df["UF"].dropna().unique().tolist())
    ufs_selecionadas = st.sidebar.multiselect("UF", ufs_disponiveis, default=ufs_disponiveis)
else:
    ufs_selecionadas = None

# Filtro de Responsável
if "Responsável" in df.columns:
    responsaveis_disponiveis = sorted(df["Responsável"].dropna().unique().tolist())
    responsaveis_selecionados = st.sidebar.multiselect(
        "Responsável", responsaveis_disponiveis, default=responsaveis_disponiveis
    )
else:
    responsaveis_selecionados = None

# Filtro de Município
if "Município" in df.columns:
    municipios_disponiveis = sorted(df["Município"].dropna().unique().tolist())
    municipios_selecionados = st.sidebar.multiselect("Município", municipios_disponiveis)
else:
    municipios_selecionados = None


# ==========================================================
# APLICA OS FILTROS
# ==========================================================

df_filtrado = df.copy()

if periodo and len(periodo) == 2:
    inicio, fim = periodo
    df_filtrado = df_filtrado[
        (df_filtrado["data_execucao"].dt.date >= inicio)
        & (df_filtrado["data_execucao"].dt.date <= fim)
    ]

if ufs_selecionadas:
    df_filtrado = df_filtrado[df_filtrado["UF"].isin(ufs_selecionadas)]

if responsaveis_selecionados:
    df_filtrado = df_filtrado[df_filtrado["Responsável"].isin(responsaveis_selecionados)]

if municipios_selecionados:
    df_filtrado = df_filtrado[df_filtrado["Município"].isin(municipios_selecionados)]


# ==========================================================
# CABEÇALHO E MÉTRICAS
# ==========================================================

st.title("📊 Dashboard Coinfra - Histórico TAS Infra")
st.caption(f"Última atualização dos dados: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Total de registros (filtrados)", len(df_filtrado))
if "Município" in df_filtrado.columns:
    col2.metric("Municípios distintos", df_filtrado["Município"].nunique())
if "DDD" in df_filtrado.columns:
    col3.metric("DDDs distintos", df_filtrado["DDD"].nunique())
if "data_execucao" in df_filtrado.columns and df_filtrado["data_execucao"].notna().any():
    col4.metric("Execuções distintas", df_filtrado["data_execucao"].dt.date.nunique())

st.divider()


# ==========================================================
# GRÁFICOS
# ==========================================================

col_esq, col_dir = st.columns(2)

with col_esq:
    if "DDD" in df_filtrado.columns:
        st.subheader("Registros por DDD")
        contagem_ddd = df_filtrado["DDD"].value_counts().reset_index()
        contagem_ddd.columns = ["DDD", "Quantidade"]
        fig_ddd = px.bar(contagem_ddd, x="DDD", y="Quantidade")
        st.plotly_chart(fig_ddd, use_container_width=True)

with col_dir:
    if "Município" in df_filtrado.columns:
        st.subheader("Top 15 municípios")
        contagem_municipio = (
            df_filtrado["Município"].value_counts().head(15).reset_index()
        )
        contagem_municipio.columns = ["Município", "Quantidade"]
        fig_municipio = px.bar(
            contagem_municipio, x="Quantidade", y="Município", orientation="h"
        )
        fig_municipio.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_municipio, use_container_width=True)

if "Responsável" in df_filtrado.columns:
    st.subheader("Registros por Responsável")
    contagem_resp = df_filtrado["Responsável"].value_counts().reset_index()
    contagem_resp.columns = ["Responsável", "Quantidade"]
    fig_resp = px.pie(contagem_resp, names="Responsável", values="Quantidade")
    st.plotly_chart(fig_resp, use_container_width=True)

if "data_execucao" in df_filtrado.columns and df_filtrado["data_execucao"].notna().any():
    st.subheader("Evolução de registros por execução")
    evolucao = (
        df_filtrado.groupby(df_filtrado["data_execucao"].dt.date)
        .size()
        .reset_index(name="Quantidade")
    )
    evolucao.columns = ["Data", "Quantidade"]
    fig_evolucao = px.line(evolucao, x="Data", y="Quantidade", markers=True)
    st.plotly_chart(fig_evolucao, use_container_width=True)


# ==========================================================
# TABELA DE DADOS + DOWNLOAD
# ==========================================================

st.divider()
st.subheader("📋 Dados detalhados")
st.dataframe(df_filtrado, use_container_width=True)

csv = df_filtrado.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️ Baixar dados filtrados (CSV)",
    data=csv,
    file_name="historico_tas_infra_filtrado.csv",
    mime="text/csv",
)