# ============================================================
# Fitch Sovereign Comparator – Streamlit App
# ============================================================

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

# ============================================================
# 1. Leitura do XLSB
# ============================================================

def read_xlsb(file):
    rows = []
    with open_workbook(file) as wb:
        sheet = wb.get_sheet(1)
        for row in sheet.rows():
            rows.append([cell.v for cell in row])
    return pd.DataFrame(rows)


def find_data_start(df):
    """
    Encontra onde começam países/medianas
    """
    for i in range(len(df)):
        v = str(df.iloc[i, 0]).lower()
        if re.match(r"[a-z]{3,}sov", v) or "eur" in v:
            return i
    raise ValueError("Não foi possível localizar início dos dados")


def normalize_long(df_raw):
    start = find_data_start(df_raw)

    h1 = df_raw.iloc[start - 3].fillna("")
    h2 = df_raw.iloc[start - 2].fillna("")
    h3 = df_raw.iloc[start - 1].fillna("")

    df = df_raw.iloc[start:].reset_index(drop=True)

    df.columns = [
        f"{h1[i]} | {h2[i]} | {h3[i]}".strip()
        for i in range(len(df.columns))
    ]

    meta_cols = df.columns[:10]
    value_cols = df.columns[10:]

    long = df.melt(
        id_vars=meta_cols,
        value_vars=value_cols,
        var_name="raw_indicator",
        value_name="value"
    )

    long["year"] = long["raw_indicator"].str.extract(r"(19\d{2}|20\d{2})")
    long["indicator"] = long["raw_indicator"].str.replace(r"\|.*", "", regex=True).str.strip()
    long["entity"] = df.iloc[:, 0].values.repeat(len(value_cols))

    long["value"] = pd.to_numeric(long["value"], errors="coerce")

    return long.dropna(subset=["year"])

# ============================================================
# 2. UI – Upload
# ============================================================

st.title("Fitch Global Sovereign Data Comparator")

uploaded = st.file_uploader(
    "📤 Envie o arquivo XLSB da Fitch",
    type=["xlsb"]
)

if not uploaded:
    st.info("Envie o arquivo para iniciar")
    st.stop()

with st.spinner("Lendo e processando XLSB..."):
    raw_df = read_xlsb(uploaded)
    df_long = normalize_long(raw_df)

# ============================================================
# 3. Abas do App
# ============================================================

tab_met, tab_dash, tab_data = st.tabs(
    ["📘 Metodologia", "📊 Dashboard", "📋 Dados"]
)

# ============================================================
# Aba 1 – Metodologia
# ============================================================

with tab_met:
    st.markdown("""
### Fitch Global Sovereign Data Comparator

Este aplicativo lê diretamente a **base pública do Fitch Sovereign Comparator**
no formato XLSB.

**Tratamento aplicado:**
- Leitura direta do XLSB (sem Excel intermediário)
- Cabeçalhos multinível ignorados
- Conversão para formato longo
- Nenhuma interpolação ou ajuste estatístico

**Uso:**
- Ferramenta analítica
- Não substitui avaliações oficiais da Fitch Ratings
""")

# ============================================================
# Aba 2 – Dashboard
# ============================================================

with tab_dash:
    st.subheader("Dashboard interativo")

    c1, c2, c3 = st.columns(3)

    with c1:
        entity = st.selectbox(
            "País / Grupo",
            sorted(df_long["entity"].unique())
        )

    with c2:
        indicator = st.selectbox(
            "Indicador",
            sorted(df_long["indicator"].unique())
        )

    with c3:
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

# ============================================================
# Aba 3 – Dados
# ============================================================

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
