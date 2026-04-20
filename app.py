import streamlit as st
import pandas as pd
import os
import sys
import io
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch

# =============================
# CONFIG & ESTILO CORPORATIVO
# =============================
st.set_page_config(
    layout="wide",
    page_title="Painel Tributário | IRPJ & CSLL",
    page_icon="📊",
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

    /* ── Reset & Base ── */
    html, body, [class*="css"] {
        font-family: 'DM Sans', sans-serif;
    }

    .main { background: #0D1117; }
    .block-container { padding: 2rem 2.5rem 3rem; }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: #161B22;
        border-right: 1px solid #21262D;
    }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] .stMarkdown {
        color: #C9D1D9 !important;
    }
    [data-testid="stSidebar"] .stSelectbox > div > div {
        background: #21262D;
        border: 1px solid #30363D;
        color: #C9D1D9;
        border-radius: 8px;
    }

    /* ── Header ── */
    .corp-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 1.25rem 2rem;
        background: linear-gradient(135deg, #161B22 0%, #1C2128 100%);
        border: 1px solid #21262D;
        border-radius: 14px;
        margin-bottom: 1.75rem;
        box-shadow: 0 4px 24px rgba(0,0,0,0.4);
    }
    .corp-header-left h1 {
        font-size: 1.45rem;
        font-weight: 700;
        color: #E6EDF3;
        margin: 0 0 2px;
        letter-spacing: -0.3px;
    }
    .corp-header-left p {
        font-size: 0.82rem;
        color: #8B949E;
        margin: 0;
        font-weight: 400;
    }
    .corp-badge {
        background: linear-gradient(135deg, #1F6FEB 0%, #388BFD 100%);
        color: #fff;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 5px 14px;
        border-radius: 20px;
        letter-spacing: 0.5px;
    }

    /* ── Section title ── */
    .section-title {
        font-size: 1.05rem;
        font-weight: 600;
        color: #E6EDF3;
        margin: 1.5rem 0 0.85rem;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid #21262D;
        letter-spacing: -0.2px;
    }

    /* ── KPI Cards ── */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
        gap: 14px;
        margin-bottom: 1.5rem;
    }
    .kpi-card {
        background: #161B22;
        border: 1px solid #21262D;
        border-radius: 12px;
        padding: 1.1rem 1.25rem;
        transition: border-color 0.2s, transform 0.2s;
    }
    .kpi-card:hover { border-color: #388BFD; transform: translateY(-2px); }
    .kpi-label {
        font-size: 0.72rem;
        font-weight: 600;
        color: #8B949E;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 1.2rem;
        font-weight: 700;
        color: #E6EDF3;
        font-family: 'DM Mono', monospace;
        letter-spacing: -0.5px;
    }
    .kpi-value.green { color: #3FB950; }
    .kpi-value.blue  { color: #388BFD; }
    .kpi-value.amber { color: #D29922; }
    .kpi-value.red   { color: #F85149; }

    /* ── Company Cards ── */
    .company-card {
        background: #161B22;
        border: 1px solid #21262D;
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        transition: border-color 0.2s, box-shadow 0.2s;
    }
    .company-card:hover {
        border-color: #388BFD;
        box-shadow: 0 0 0 1px #388BFD22;
    }
    .company-name {
        font-size: 0.95rem;
        font-weight: 600;
        color: #E6EDF3;
        margin-bottom: 4px;
    }
    .company-sub {
        font-size: 0.78rem;
        color: #8B949E;
    }
    .company-values {
        display: flex;
        gap: 2rem;
        text-align: right;
    }
    .company-val-block .val-label {
        font-size: 0.68rem;
        color: #8B949E;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .company-val-block .val-num {
        font-size: 0.92rem;
        font-weight: 700;
        color: #E6EDF3;
        font-family: 'DM Mono', monospace;
    }
    .company-val-block .val-num.total { color: #388BFD; }

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] {
        background: #161B22;
        border-radius: 10px;
        padding: 4px;
        border: 1px solid #21262D;
        gap: 2px;
    }
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        color: #8B949E;
        border-radius: 7px;
        font-weight: 500;
        font-size: 0.88rem;
        padding: 0.45rem 1.1rem;
        border: none;
    }
    .stTabs [aria-selected="true"] {
        background: #21262D !important;
        color: #E6EDF3 !important;
    }

    /* ── Dataframe / Table ── */
    .stDataFrame { border-radius: 10px; overflow: hidden; }
    [data-testid="stDataFrameResizable"] {
        border: 1px solid #21262D;
        border-radius: 10px;
    }

    /* ── Search box ── */
    .stTextInput > div > input {
        background: #161B22 !important;
        border: 1px solid #30363D !important;
        color: #E6EDF3 !important;
        border-radius: 8px !important;
        font-family: 'DM Sans', sans-serif;
    }
    .stTextInput > div > input::placeholder { color: #484F58 !important; }
    .stTextInput label { color: #8B949E !important; font-size: 0.82rem !important; }

    /* ── Buttons ── */
    .stButton > button {
        background: #21262D;
        color: #C9D1D9;
        border: 1px solid #30363D;
        border-radius: 8px;
        font-family: 'DM Sans', sans-serif;
        font-weight: 500;
        font-size: 0.84rem;
        padding: 0.4rem 1rem;
        transition: all 0.15s;
    }
    .stButton > button:hover {
        background: #388BFD;
        border-color: #388BFD;
        color: #fff;
    }

    /* ── Download button ── */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1F6FEB 0%, #388BFD 100%) !important;
        color: #fff !important;
        border: none !important;
        border-radius: 8px !important;
        font-family: 'DM Sans', sans-serif;
        font-weight: 600;
        font-size: 0.84rem;
        padding: 0.45rem 1.2rem;
        letter-spacing: 0.2px;
    }
    .stDownloadButton > button:hover {
        opacity: 0.88;
        transform: translateY(-1px);
    }

    /* ── Info/success boxes ── */
    .stAlert { border-radius: 10px; }

    /* ── Metric override ── */
    [data-testid="stMetric"] {
        background: #161B22;
        border: 1px solid #21262D;
        border-radius: 12px;
        padding: 1rem 1.2rem;
    }
    [data-testid="stMetricLabel"] { color: #8B949E !important; font-size: 0.75rem !important; }
    [data-testid="stMetricValue"] { color: #E6EDF3 !important; font-family: 'DM Mono', monospace !important; }

    /* ── Sidebar logo area ── */
    .sidebar-logo {
        text-align: center;
        padding: 1.5rem 1rem 1rem;
        border-bottom: 1px solid #21262D;
        margin-bottom: 1.25rem;
    }
    .sidebar-logo .logo-icon {
        font-size: 2rem;
        margin-bottom: 6px;
    }
    .sidebar-logo .logo-text {
        font-size: 0.78rem;
        font-weight: 600;
        color: #8B949E;
        text-transform: uppercase;
        letter-spacing: 1.5px;
    }
    .filter-divider {
        border: none;
        border-top: 1px solid #21262D;
        margin: 1rem 0;
    }
    .tag-irpj {
        background: #1F3A5F; color: #79C0FF;
        padding: 2px 8px; border-radius: 4px;
        font-size: 0.72rem; font-weight: 600;
    }
    .tag-csll {
        background: #3A1F5F; color: #D2A8FF;
        padding: 2px 8px; border-radius: 4px;
        font-size: 0.72rem; font-weight: 600;
    }
    .empty-state {
        text-align: center;
        padding: 3rem 1rem;
        color: #484F58;
    }
    .empty-state .icon { font-size: 2.5rem; margin-bottom: 10px; }
    .empty-state p { font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

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
# CARREGAR BASE
# =============================
@st.cache_data
def carregar_dados():
    arquivo = "base.xls"
    if not os.path.exists(arquivo):
        st.error("Arquivo base.xls não encontrado.")
        st.stop()
    return pd.read_excel(arquivo)

df = carregar_dados()

# ── Remover canceladas ──
df = df[
    df["dsc_situacao"].astype(str).str.upper().str.strip() != "CANCELADA"
]

# ── Tratamento ──
colunas = [
    "cod_pagamento", "dat_pagamento", "cf_empresa",
    "vlr_principal", "vlr_multa",
    "vlr_juro_encargo", "vlr_outra_entidade", "vlr_total"
]
for col in colunas:
    if col not in df.columns:
        df[col] = 0

df["cod_pagamento"] = df["cod_pagamento"].astype(str).str.strip()
df = df[df["cod_pagamento"].isin(["2362", "2484", "2089", "2372"])]

valores = colunas[3:]
for col in valores:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

df["dat_pagamento"] = pd.to_datetime(df["dat_pagamento"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["dat_pagamento"])

df["Ano"] = df["dat_pagamento"].dt.year
df["Trimestre"] = df["dat_pagamento"].dt.quarter

df["Imposto"] = df["cod_pagamento"].map({
    "2362": "IRPJ", "2089": "IRPJ",
    "2484": "CSLL", "2372": "CSLL"
})

if df.empty:
    st.error("Sem dados após tratamento.")
    st.stop()

# =============================
# EXPORTAR EXCEL
# =============================
def exportar_excel(dataframe: pd.DataFrame, nome_aba: str = "Dados") -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=nome_aba)
        workbook  = writer.book
        worksheet = writer.sheets[nome_aba]

        # Formatos
        fmt_header = workbook.add_format({
            "bold": True, "bg_color": "#1F6FEB", "font_color": "#FFFFFF",
            "border": 1, "border_color": "#FFFFFF", "align": "center",
            "valign": "vcenter", "font_name": "Calibri", "font_size": 11,
        })
        fmt_money = workbook.add_format({
            "num_format": 'R$ #,##0.00', "font_name": "Calibri",
            "font_size": 10, "align": "right",
        })
        fmt_default = workbook.add_format({
            "font_name": "Calibri", "font_size": 10,
        })
        fmt_alt = workbook.add_format({
            "font_name": "Calibri", "font_size": 10,
            "bg_color": "#F0F4FF",
        })

        money_cols = {c for c in dataframe.columns if "vlr_" in c.lower()}

        for col_idx, col_name in enumerate(dataframe.columns):
            worksheet.write(0, col_idx, col_name, fmt_header)
            col_width = max(len(str(col_name)) + 4,
                            dataframe[col_name].astype(str).str.len().max() + 2
                            if not dataframe.empty else 12)
            worksheet.set_column(col_idx, col_idx, min(col_width, 30))

        for row_idx, row in dataframe.iterrows():
            base_fmt = fmt_alt if (row_idx % 2 == 0) else fmt_default
            for col_idx, col_name in enumerate(dataframe.columns):
                val = row[col_name]
                if col_name in money_cols:
                    try:
                        worksheet.write_number(row_idx + 1, col_idx, float(val), fmt_money)
                    except Exception:
                        worksheet.write(row_idx + 1, col_idx, val, base_fmt)
                else:
                    worksheet.write(row_idx + 1, col_idx, val, base_fmt)

        worksheet.freeze_panes(1, 0)
    return buffer.getvalue()

# =============================
# SIDEBAR
# =============================
with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
        <div class="logo-icon">🏦</div>
        <div class="logo-text">Painel Tributário</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("**📅 Período**")
    ano = st.selectbox("Ano", sorted(df["Ano"].unique()), label_visibility="collapsed")
    tri = st.selectbox("Trimestre", [1, 2, 3, 4],
                       format_func=lambda x: f"Q{x} — {['Jan-Mar','Abr-Jun','Jul-Set','Out-Dez'][x-1]}",
                       label_visibility="collapsed")

    st.markdown("<hr class='filter-divider'>", unsafe_allow_html=True)

    st.markdown("**🏢 Empresa**")
    empresas_disp = ["Todas"] + sorted(df["cf_empresa"].dropna().unique().tolist())
    empresa_filtro = st.selectbox("Empresa", empresas_disp, label_visibility="collapsed")

    st.markdown("<hr class='filter-divider'>", unsafe_allow_html=True)
    st.markdown("**🧾 Imposto**")
    imp_filtro = st.multiselect("Imposto", ["IRPJ", "CSLL"],
                                default=["IRPJ", "CSLL"], label_visibility="collapsed")

# =============================
# FILTRO PRINCIPAL
# =============================
df_tri = df[(df["Ano"] == ano) & (df["Trimestre"] == tri)]
if empresa_filtro != "Todas":
    df_tri = df_tri[df_tri["cf_empresa"] == empresa_filtro]
if imp_filtro:
    df_tri = df_tri[df_tri["Imposto"].isin(imp_filtro)]

# =============================
# HEADER
# =============================
st.markdown(f"""
<div class="corp-header">
    <div class="corp-header-left">
        <h1>📊 Painel IRPJ & CSLL</h1>
        <p>Análise de pagamentos tributários · {ano} · Q{tri} · {['Jan–Mar','Abr–Jun','Jul–Set','Out–Dez'][tri-1]}</p>
    </div>
    <div class="corp-badge">CORPORATIVO</div>
</div>
""", unsafe_allow_html=True)

# =============================
# KPIs GLOBAIS
# =============================
total_geral  = df_tri["vlr_total"].sum()
total_irpj   = df_tri[df_tri["Imposto"] == "IRPJ"]["vlr_total"].sum()
total_csll   = df_tri[df_tri["Imposto"] == "CSLL"]["vlr_total"].sum()
total_multa  = df_tri["vlr_multa"].sum()
total_juros  = df_tri["vlr_juro_encargo"].sum()
n_empresas   = df_tri["cf_empresa"].nunique()

st.markdown("""
<div class="kpi-grid">
""", unsafe_allow_html=True)

kpi_cols = st.columns(6)
kpi_cols[0].metric("Total Geral",    f"R$ {formato_br(total_geral)}")
kpi_cols[1].metric("IRPJ",           f"R$ {formato_br(total_irpj)}")
kpi_cols[2].metric("CSLL",           f"R$ {formato_br(total_csll)}")
kpi_cols[3].metric("Multas",         f"R$ {formato_br(total_multa)}")
kpi_cols[4].metric("Juros/Encargos", f"R$ {formato_br(total_juros)}")
kpi_cols[5].metric("Empresas",       str(n_empresas))

st.markdown("</div>", unsafe_allow_html=True)

# =============================
# ABAS
# =============================
abas = st.tabs(["📊 Dashboard", "📑 Analítico", "📈 Consolidados"])

# ── aba 0: DASHBOARD ──────────────────────────────────
with abas[0]:
    st.markdown('<div class="section-title">Visão por Empresa</div>', unsafe_allow_html=True)

    # Pesquisa
    busca = st.text_input("🔍  Buscar empresa...", placeholder="Digite o nome ou código da empresa")

    empresas_rank = (
        df_tri.groupby("cf_empresa")["vlr_total"]
        .sum()
        .sort_values(ascending=False)
        .index.tolist()
    )

    if busca:
        empresas_rank = [e for e in empresas_rank if busca.upper() in str(e).upper()]

    if not empresas_rank:
        st.markdown("""
        <div class="empty-state">
            <div class="icon">🔍</div>
            <p>Nenhuma empresa encontrada para "<b>{}</b>"</p>
        </div>
        """.format(busca), unsafe_allow_html=True)
    else:
        for emp in empresas_rank:
            df_emp = df_tri[df_tri["cf_empresa"] == emp]
            total  = df_emp["vlr_total"].sum()
            irpj   = df_emp[df_emp["Imposto"] == "IRPJ"]["vlr_total"].sum()
            csll   = df_emp[df_emp["Imposto"] == "CSLL"]["vlr_total"].sum()
            multa  = df_emp["vlr_multa"].sum()

            tag_irpj = f'<span class="tag-irpj">IRPJ</span>' if irpj > 0 else ""
            tag_csll = f'<span class="tag-csll">CSLL</span>' if csll > 0 else ""

            st.markdown(f"""
            <div class="company-card">
                <div>
                    <div class="company-name">{emp} &nbsp; {tag_irpj} {tag_csll}</div>
                    <div class="company-sub">Multas: R$ {formato_br(multa)}</div>
                </div>
                <div class="company-values">
                    <div class="company-val-block">
                        <div class="val-label">IRPJ</div>
                        <div class="val-num">R$ {formato_br(irpj)}</div>
                    </div>
                    <div class="company-val-block">
                        <div class="val-label">CSLL</div>
                        <div class="val-num">R$ {formato_br(csll)}</div>
                    </div>
                    <div class="company-val-block">
                        <div class="val-label">Total</div>
                        <div class="val-num total">R$ {formato_br(total)}</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if st.button(f"Ver detalhes → {emp}", key=f"btn_{emp}"):
                st.session_state.empresa = emp

    # Exportar dashboard
    st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)
    df_export_dash = df_tri.groupby("cf_empresa")[valores + ["Imposto"]].agg(
        {**{v: "sum" for v in valores}, "Imposto": lambda x: ", ".join(sorted(x.unique()))}
    ).reset_index()
    st.download_button(
        label="⬇ Exportar Dashboard (.xlsx)",
        data=exportar_excel(df_export_dash, "Dashboard"),
        file_name=f"dashboard_{ano}_Q{tri}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ── aba 1: ANALÍTICO ──────────────────────────────────
with abas[1]:
    emp = st.session_state.empresa

    if not emp:
        st.markdown("""
        <div class="empty-state">
            <div class="icon">🏢</div>
            <p>Selecione uma empresa no Dashboard para ver o detalhamento.</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        df_emp = df_tri[df_tri["cf_empresa"] == emp].copy()

        st.markdown(f'<div class="section-title">🏢 {emp}</div>', unsafe_allow_html=True)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Principal", f"R$ {formato_br(df_emp['vlr_principal'].sum())}")
        c2.metric("Multa",     f"R$ {formato_br(df_emp['vlr_multa'].sum())}")
        c3.metric("Juros",     f"R$ {formato_br(df_emp['vlr_juro_encargo'].sum())}")
        c4.metric("Outros",    f"R$ {formato_br(df_emp['vlr_outra_entidade'].sum())}")
        c5.metric("Total",     f"R$ {formato_br(df_emp['vlr_total'].sum())}")

        st.markdown('<div class="section-title">Detalhamento de Pagamentos</div>', unsafe_allow_html=True)

        # Busca no analítico
        busca_analitico = st.text_input(
            "🔍 Filtrar registros...",
            placeholder="Buscar por código, data, imposto...",
            key="busca_analitico"
        )

        cols_exibir = ["dat_pagamento", "cod_pagamento", "Imposto"] + valores
        df_det = df_emp[cols_exibir].copy()
        df_det["dat_pagamento"] = df_det["dat_pagamento"].dt.strftime("%d/%m/%Y")

        if busca_analitico:
            mask = df_det.apply(
                lambda row: busca_analitico.upper() in " ".join(row.astype(str).str.upper()), axis=1
            )
            df_det = df_det[mask]

        if df_det.empty:
            st.info("Nenhum registro encontrado para o filtro informado.")
        else:
            # Exibe sem formatar moeda para manter ordenação
            st.dataframe(
                df_det.style.format({v: "R$ {:,.2f}" for v in valores}),
                use_container_width=True,
                hide_index=True,
            )

        # Exportar analítico
        col_exp1, col_exp2 = st.columns([1, 3])
        with col_exp1:
            st.download_button(
                label="⬇ Exportar Analítico (.xlsx)",
                data=exportar_excel(df_emp[cols_exibir].assign(
                    dat_pagamento=df_emp["dat_pagamento"].dt.strftime("%d/%m/%Y")
                ), "Analítico"),
                file_name=f"analitico_{emp}_{ano}_Q{tri}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ── aba 2: CONSOLIDADOS ───────────────────────────────
with abas[2]:
    st.markdown('<div class="section-title">Consolidação Geral</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Principal", f"R$ {formato_br(df_tri['vlr_principal'].sum())}")
    c2.metric("Multa",     f"R$ {formato_br(df_tri['vlr_multa'].sum())}")
    c3.metric("Juros",     f"R$ {formato_br(df_tri['vlr_juro_encargo'].sum())}")
    c4.metric("Total",     f"R$ {formato_br(df_tri['vlr_total'].sum())}")

    st.markdown('<div class="section-title">Por Empresa</div>', unsafe_allow_html=True)

    # Busca no consolidado
    busca_consol = st.text_input("🔍 Filtrar empresas...", placeholder="Nome ou código", key="busca_consol")

    consolidado = (
        df_tri.groupby("cf_empresa")[valores]
        .sum()
        .reset_index()
        .sort_values("vlr_total", ascending=False)
    )

    if busca_consol:
        consolidado = consolidado[
            consolidado["cf_empresa"].astype(str).str.upper().str.contains(busca_consol.upper())
        ]

    if consolidado.empty:
        st.info("Nenhuma empresa encontrada.")
    else:
        st.dataframe(
            consolidado.style.format({v: "R$ {:,.2f}" for v in valores}),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown('<div class="section-title">Por Imposto</div>', unsafe_allow_html=True)
    por_imposto = (
        df_tri.groupby("Imposto")[valores]
        .sum()
        .reset_index()
    )
    st.dataframe(
        por_imposto.style.format({v: "R$ {:,.2f}" for v in valores}),
        use_container_width=True,
        hide_index=True,
    )

    # Exportar consolidado
    st.markdown("")
    col_e1, col_e2, col_e3 = st.columns([1, 1, 2])
    with col_e1:
        st.download_button(
            label="⬇ Exportar por Empresa (.xlsx)",
            data=exportar_excel(consolidado, "Consolidado por Empresa"),
            file_name=f"consolidado_empresas_{ano}_Q{tri}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_e2:
        st.download_button(
            label="⬇ Exportar por Imposto (.xlsx)",
            data=exportar_excel(por_imposto, "Consolidado por Imposto"),
            file_name=f"consolidado_imposto_{ano}_Q{tri}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
