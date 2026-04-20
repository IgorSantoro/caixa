import streamlit as st
import pandas as pd
import os
import sys
from io import BytesIO

st.set_page_config(layout="wide")

# =============================
# ESTADO
# =============================
if "empresa" not in st.session_state:
    st.session_state.empresa = None

# =============================
# FORMATAÇÃO
# =============================
def formato_br(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# =============================
# EXPORTAÇÃO EXCEL
# =============================
def exportar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")
    return output.getvalue()

# =============================
# CARREGAR BASE
# =============================
@st.cache_data
def carregar_dados():
    arquivo = "base.xls"

    if not os.path.exists(arquivo):
        st.error("Arquivo não encontrado")
        st.stop()

    return pd.read_excel(arquivo)

df = carregar_dados()

# =============================
# REMOVER CANCELADAS
# =============================
df = df[
    df["dsc_situacao"]
    .astype(str)
    .str.upper()
    .str.strip() != "CANCELADA"
]

# =============================
# TRATAMENTO
# =============================
colunas = [
    "cod_pagamento","dat_pagamento","cf_empresa",
    "vlr_principal","vlr_multa",
    "vlr_juro_encargo","vlr_outra_entidade","vlr_total"
]

for col in colunas:
    if col not in df.columns:
        df[col] = 0

df["cod_pagamento"] = df["cod_pagamento"].astype(str).str.strip()
df = df[df["cod_pagamento"].isin(["2362","2484","2089","2372"])]

valores = colunas[3:]

for col in valores:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

df["dat_pagamento"] = pd.to_datetime(df["dat_pagamento"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["dat_pagamento"])

df["Ano"] = df["dat_pagamento"].dt.year
df["Trimestre"] = df["dat_pagamento"].dt.quarter

df["Imposto"] = df["cod_pagamento"].map({
    "2362":"IRPJ","2089":"IRPJ",
    "2484":"CSLL","2372":"CSLL"
})

if df.empty:
    st.error("Sem dados após tratamento")
    st.stop()

# =============================
# FILTROS (MELHORADOS)
# =============================
st.sidebar.header("Filtros")

ano = st.sidebar.selectbox("Ano", sorted(df["Ano"].unique()))
tri = st.sidebar.selectbox("Trimestre", [1,2,3,4])

empresas_sel = st.sidebar.multiselect(
    "Empresa",
    options=sorted(df["cf_empresa"].unique()),
    default=None
)

imposto_sel = st.sidebar.multiselect(
    "Imposto",
    options=["IRPJ", "CSLL"],
    default=["IRPJ", "CSLL"]
)

df_filtrado = df[
    (df["Ano"] == ano) &
    (df["Trimestre"] == tri) &
    (df["Imposto"].isin(imposto_sel))
]

if empresas_sel:
    df_filtrado = df_filtrado[df_filtrado["cf_empresa"].isin(empresas_sel)]

# =============================
# FUNÇÃO TABELA
# =============================
def mostrar_tabela(df):
    if sys.version_info >= (3, 13):
        st.table(df)
    else:
        st.dataframe(df)

# =============================
# ABAS
# =============================
abas = st.tabs(["📊 Dashboard", "📑 Analítico", "📈 Consolidados"])

# =============================
# DASHBOARD
# =============================
with abas[0]:
    st.markdown("## 📊 Visão Corporativa")

    empresas = (
        df_filtrado.groupby("cf_empresa")["vlr_total"]
        .sum()
        .sort_values(ascending=False)
        .index
    )

    for emp in empresas:
        df_emp = df_filtrado[df_filtrado["cf_empresa"]==emp]

        total = df_emp["vlr_total"].sum()
        irpj = df_emp[df_emp["Imposto"]=="IRPJ"]["vlr_total"].sum()
        csll = df_emp[df_emp["Imposto"]=="CSLL"]["vlr_total"].sum()

        st.markdown(f"""
        <div style="border:1px solid #ddd; padding:10px; margin-bottom:10px;">
            <b>{emp}</b><br>
            Total: R$ {formato_br(total)}<br>
            IRPJ: R$ {formato_br(irpj)}<br>
            CSLL: R$ {formato_br(csll)}
        </div>
        """, unsafe_allow_html=True)

        if st.button(f"Selecionar {emp}"):
            st.session_state.empresa = emp

# =============================
# ANALÍTICO + EXPORTAÇÃO
# =============================
with abas[1]:
    emp = st.session_state.empresa

    if not emp:
        st.info("Selecione uma empresa no Dashboard")
    else:
        df_emp = df_filtrado[df_filtrado["cf_empresa"]==emp]

        st.markdown(f"## {emp}")

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Principal", f"R$ {formato_br(df_emp['vlr_principal'].sum())}")
        c2.metric("Multa", f"R$ {formato_br(df_emp['vlr_multa'].sum())}")
        c3.metric("Juros", f"R$ {formato_br(df_emp['vlr_juro_encargo'].sum())}")
        c4.metric("Outros", f"R$ {formato_br(df_emp['vlr_outra_entidade'].sum())}")
        c5.metric("Total", f"R$ {formato_br(df_emp['vlr_total'].sum())}")

        st.markdown("### Detalhamento")

        mostrar_tabela(df_emp)

        # 🔽 EXPORTAR
        st.download_button(
            "📥 Exportar Analítico para Excel",
            exportar_excel(df_emp),
            file_name=f"analitico_{emp}.xlsx"
        )

# =============================
# CONSOLIDADO + EXPORTAÇÃO
# =============================
with abas[2]:
    st.markdown("## Consolidação")

    df_consol = df_filtrado.copy()

    consolidado = (
        df_consol.groupby("cf_empresa")[valores]
        .sum()
        .reset_index()
    )

    mostrar_tabela(consolidado)

    # 🔽 EXPORTAR
    st.download_button(
        "📥 Exportar Consolidado para Excel",
        exportar_excel(consolidado),
        file_name="consolidado.xlsx"
    )
