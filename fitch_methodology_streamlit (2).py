# fitch_comparator_app.py
import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px
from pyxlsb import open_workbook

st.set_page_config(
    page_title="Fitch Sovereign Data Comparator",
    layout="wide"
)

# =========================================================
# Utils
# =========================================================

def read_xlsb_to_df(file):
    """
    Lê a PRIMEIRA aba do XLSB e retorna DataFrame cru
    """
    data = []
    with open_workbook(file) as wb:
        sheet = wb.get_sheet(1)
        for row in sheet.rows():
            data.append([cell.v for cell in row])
    return pd.DataFrame(data)


def find_data_start(df):
    """
    Encontra a linha onde começam os dados (países / medianas)
    """
    for i in range(len(df)):
        val = str(df.iloc[i, 0]).lower()
        if re.match(r"[a-z]{3,}sov", val) or "eur" in val:
            return i
    return None


def normalize_long(df_raw):
    """
    Converte a planilha gigante em formato LONGO
    """
    start = find_data_start(df_raw)
    if start is None:
        raise ValueError("Não foi possível localizar o início dos dados")

    header_1 = df_raw.iloc[start - 3].fillna("")
    header_2 = df_raw.iloc[start - 2].fillna("")
    header_3 = df_raw.iloc[start - 1].fillna("")

    df = df_raw.iloc[start:].reset_index(drop=True)

    df.columns = [
        f"{header_1[i]} | {header_2[i]} | {header_3[i]}".strip()
        for i in range(len(df.columns))
    ]

    # Metadados básicos (sempre nas primeiras colunas)
    meta_cols = df.columns[:10]

    value_cols = df.columns[10:]

    long_df = df.melt(
        id_vars=meta_cols,
        value_vars=value_cols,
        var_name="raw_indicator",
        value_name="value"
    )

    # Extrai ano
    long_df["year"] = (
        long_df["raw_indicator"]
        .str.extract(r"(19\d{2}|20\d{2})")[0]
    )

    # Limpa nome do indicador
    long_df["indicator"] = (
        long_df["raw_indicator"]
        .str.replace(r"\|.*", "", regex=True)
        .str.strip()
    )

    # País / grupo
    long_df["entity"] = df.iloc[:, 0].values.repeat(len(value_cols))

    long_df["value"] = pd.to_numeric(long_df["value"], errors="coerce")

    return long_df.dropna(subset=["year"])


# =========================================================
# UI
# =========================================================

st.title("Fitch Global Sovereign Data Comparator")

uploaded = st.file_uploader(
    "📤 Envie o arquivo XLSB da Fitch",
    type=["xlsb"]
)

if not uploaded:
    st.info("Envie o arquivo para iniciar")
    st.stop()

with st.spinner("Lendo arquivo XLSB..."):
    raw_df = read_xlsb_to_df(uploaded)

with st.spinner("Normalizando dados..."):
    df_long = normalize_long(raw_df)

tab_met, tab_dash, tab_data = st.tabs(
    ["📘 Metodologia", "📊 Dashboard", "📋 Dados"]
)

# =========================================================
# Metodologia
# =========================================================
with tab_met:
    st.markdown("""
### Fitch Global Sovereign Data Comparator

Este dashboard utiliza a base pública do **Fitch Sovereign Comparator**,
organizada em milhares de indicadores macroeconômicos, fiscais, externos
e institucionais.

**Tratamento aplicado:**
- Leitura direta do XLSB original
- Ignora cabeçalhos multinível
- Converte dados para formato longo
- Nenhuma interpolação ou ajuste estatístico

**Observação importante:**
Este material é **exclusivamente analítico** e não substitui
avaliações oficiais da Fitch Ratings.
""")

# =========================================================
# Dashboard
# =========================================================
with tab_dash:
    st.subheader("Dashboard interativo")

    col1, col2, col3 = st.columns(3)

    with col1:
        entity = st.selectbox(
            "País / Grupo",
            sorted(df_long["entity"].unique())
        )

    with col2:
        indicator = st.selectbox(
            "Indicador",
            sorted(df_long["indicator"].unique())
        )

    with col3:
        years = st.multiselect(
            "Anos",
            sorted(df_long["year"].unique()),
            default=sorted(df_long["year"].unique())[-5:]
        )

    plot_df = df_long[
        (df_long["entity"] == entity) &
        (df_long["indicator"] == indicator) &
        (df_long["year"].isin(years))
    ].sort_values("year")

    fig = px.line(
        plot_df,
        x="year",
        y="value",
        markers=True,
        title=f"{indicator} — {entity}"
    )

    st.plotly_chart(fig, use_container_width=True)

# =========================================================
# Dados
# =========================================================
with tab_data:
    st.subheader("Tabela de dados")

    view = st.radio(
        "Formato",
        ["Longo", "Pivot"],
        horizontal=True
    )

    if view == "Longo":
        display = df_long
    else:
        display = df_long.pivot_table(
            index=["entity", "indicator"],
            columns="year",
            values="value",
            aggfunc="first"
        ).reset_index()

    st.dataframe(display, use_container_width=True)

    csv = display.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Download CSV",
        csv,
        file_name="fitch_comparator.csv",
        mime="text/csv"
    )
