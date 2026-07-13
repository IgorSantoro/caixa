import streamlit as st
import pandas as pd
import os
import io
import html
import re

# =============================
# CONFIG
# =============================
st.set_page_config(
    layout="wide",
    page_title="Painel Tributário",
    page_icon="📊",
    initial_sidebar_state="expanded",
)

# =============================
# CONSOLIDADOS — mapeamento grupo → empresas
# =============================
# Preencha cada lista com o nome EXATO das empresas (conforme aparecem na
# coluna cf_empresa do base.xls). Pode colar quantas quiser em cada lista.
# Ex.: "ESA": ["EMPRESA A LTDA", "EMPRESA B LTDA"]
CONSOLIDADOS = {
    "ESA": [],
    "Geração": [
        "126 - PARQUE EOLICO SOBRADINHO",
        "137 - ENERGISA GERAÇÃO USINA MAURÍCIO S.A.",
        "161 - ENERGISA GERAÇÃO CENTRAL SOLAR COREMAS S.A.",
        "197 - ENERGISA GERACAO CENTRAL EOLICA BOA ESPERANCA S.A.",
        "198 - ENERGISA GERACAO CENTRAL EOLICA MANDACARU S.A.",
        "196 - ENERGISA GERACAO CENTRAL EOLICA ALECRIM S.A.",
        "199 - ENERGISA GERACAO CENTRAL EOLICA UMBUZEIRO-MUQUIM S.A.",
        "231 - ENERGISA GERACAO CENTRAL SOLAR RIO DO PEIXE I S.A.",
        "232 - ENERGISA GERACAO CENTRAL SOLAR RIO DO PEIXE II S.A.",
    ],
    "Soluções": [
        "015 - ENERGISA SOLUÇÕES S.A.",
        "160 - ENERGISA SOLUCOES CONSTRUCOES E SERV EM LINHAS E REDES S.A.",
    ],
    "Alsol": [
        "230 - ALSOL ENERGIAS RENOVAVEIS S/A",
        "233 - LARALSOL EMPREENDIMENTOS ENERGETICOS LTDA",
        "246 - URB ENERGIA LIMPA LTDA",
        "254 - REENERGISA GERACAO FOTOVOLTAICA I LTDA",
        "255 - REENERGISA GERAÇÃO FOTOVOLTAICA II S.A.",
        "256 - RENESOLAR ENGENHARIA ELETRICA LTDA",
        "257 - FLOWSOLAR ENGENHARIA ELETRICA LTDA",
        "258 - CARBONSOLAR ENGENHARIA ELETRICA LTDA",
        "281 - REENERGISA GERACAO FOTOVOLTAICA IV S.A.",
        "282 - REENERGISA GERACAO FOTOVOLTAICA VI S.A.",
        "283 - REENERGISA GERACAO FOTOVOLTAICA III S.A.",
        "284 - SPE UFV VISION VII LAGOA FORMOSA 2,4MW S.A.",
        "285 - REENERGISA GERACAO FOTOVOLTAICA VIII S.A.",
        "286 - REENERGISA GERACAO FOTOVOLTAICA V S.A.",
        "301 - ANGULO45 PARTICIPAÇÕES S.A.",
        "302 - ANGULO45 EMPREENDIMENTOS S.A.",
    ],
    "Nova Denerge": [
        "176 - DENERGE DESENVOLVIMENTO ENERGÉTICO SA",
        "260 - NOVA GEMINI TRANSMISSAO DE ENERGIA S.A. (NOVA DENERGE)",
    ],
    "EBIOGÁS": [
        "266 - AGRIC ADUBOS E GESTÃO DE RESÍDUOS INDUSTRIAIS E COMERCIAIS LTDA",
        "267 - ENERGISA BIOGAS S.A.",
        "308 - LUREAN S.A.",
    ],
    "EDISGAS": [
        "268 - COMPANHIA DE GAS DO ESPIRITO SANTO - ES GAS",
        "269 - ENERGISA DISTRIBUIÇÃO DE GAS SA",
        "303 - INFRA GAS E ENERGIA S/A",
        "304 - NORGAS S.A.",
    ],
    "Sobradinho": [
        "126 - PARQUE EOLICO SOBRADINHO",
        "236 - ENERGISA GERACAO CENTRAL EOLICA MARAVILHA I S.A.",
        "237 - ENERGISA GERACAO CENTRAL EOLICA MARAVILHA II S.A.",
        "238 - ENERGISA GERACAO CENTRAL EOLICA MARAVILHA III S.A.",
        "239 - ENERGISA GERACAO CENTRAL EOLICA MARAVILHA IV S.A.",
        "240 - ENERGISA GERACAO CENTRAL EOLICA MARAVILHA V S.A.",
    ],
    "EPNE": [
        "027 - ENERGISA PARAIBA DISTRIBUIDORA DE ENERGIA SA",
        "274 - ENERGISA PARTICIPACOES NORDESTE S.A.",
    ],
    "ETE": [
        "216 - ENERGISA PARA TRANSMISSORA DE ENERGIA I S/A",
        "217 - ENERGISA GOIAS TRANSMISSORA DE ENERGIA I S.A.",
        "224 - ENERGISA PARA TRANSMISSORA DE ENERGIA II S.A.",
        "225 - ENERGISA TRANSMISSÃO DE ENERGIA S.A.",
        "228 - ENERGISA TOCANTINS TRANSMISSORA DE ENERGIA S.A.",
        "241 - ENERGISA AMAZONAS TRANSMISSORA DE ENERGIA S.A.",
        "242 - ENERGISA TOCANTINS TRANSMISSORA DE ENERGIA II S.A.",
        "243 - ENERGISA AMAPA TRANSMISSORA DE ENERGIA S.A.",
        "247 - ENERGISA PARANAÍTA TRANSMISSORA DE ENERGIA",
        "261 - ENERGISA AMAZONAS TRANSMISSORA DE ENERGIA II S.A.",
        "262 - ENERGISA TRANSMISSORA DE ENERGIA IX S.A.",
        "263 - ENERGISA TRANSMISSORA DE ENERGIA VII S.A.",
        "271 - ENERGISA MARANHÃO TRANSMISSORA DE ENERGIA I S/A",
        "272 - ENERGISA TRANSMISSORA DE ENERGIA V S.A.",
        "273 - ENERGISA TRANSMISSORA DE ENERGIA VIII S.A.",
    ],
    "GEMINI": [
        "248 - GEMINI ENERGY S.A.",
        "249 - LINHAS DE MACAPA TRANSMISSORA DE ENERGIA S.A.",
        "250 - LINHAS DE XINGU TRANSMISSORA DE ENERGIA S.A.",
        "251 - PLENA OPERACAO E MANUTENCAO DE TRANSMISSORAS DE ENERGIA LTDA",
        "252 - LINHAS DE TAUBATE TRANSMISSORA DE ENERGIA SA",
        "253 - LINHAS DE ITACAIUNAS TRANSMISSORA DE ENERGIA LTDA.",
    ],
    "REDE": [
        "168 - REDE ENERGIA PARTICIPAÇÕES S.A",
        "170 - MULTI ENERGISA SERVICOS S.A",
        "178 - QMRA PARTICIPAÇÕES S/A",
        "182 - REDE POWER DO BRASIL S.A.",
        "184 - COMPANHIA TECNICA DE COMERCIALIZACAO DE ENERGIA",
        "190 - ENERGISA TOCANTINS - DISTRIBUIDORA DE ENERGIA S/A",
        "191 - ENERGISA MATO GROSSO - DISTRIBUIDORA DE ENERGIA S.A.",
        "193 - ENERGISA MATO GROSSO DO SUL - DISTRIBUIDORA DE ENERGIA S.A.",
    ],
}

# ESA consolida TODAS as empresas de TODOS os outros consolidados.
# Montada automaticamente a partir das listas acima — se você adicionar,
# remover ou editar uma empresa em qualquer grupo, a ESA acompanha sozinha.
CONSOLIDADOS["ESA"] = sorted({
    empresa
    for grupo, lista in CONSOLIDADOS.items()
    for empresa in lista
})

# =============================
# CSS
# =============================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

*, html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif !important;
}

.stApp { background: #060810; }
.block-container { padding: 1.75rem 2rem 3rem !important; max-width: 100% !important; }

/* ─── sidebar ─── */
[data-testid="stSidebar"] {
    background: #0C0F1A !important;
    border-right: 1px solid #1A2035 !important;
}
[data-testid="stSidebar"] * { color: #8892A4 !important; }
[data-testid="stSidebar"] strong { color: #C2CCDF !important; }
[data-testid="stSidebar"] .stSelectbox > div > div {
    background: #111527 !important;
    border: 1px solid #1E2640 !important;
    color: #C2CCDF !important;
    border-radius: 6px !important;
}

/* ─── top bar ─── */
.top-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 1.1rem 1.75rem;
    background: linear-gradient(135deg, #0C0F1A 0%, #101525 100%);
    border: 1px solid #1A2035;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 32px rgba(0,0,0,0.5);
}
.top-bar-left { display: flex; align-items: center; gap: 14px; }
.top-bar-icon {
    width: 42px; height: 42px;
    background: linear-gradient(135deg, #1A3A6A, #1F6FEB);
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 1.2rem;
}
.top-bar-title { font-size: 1.1rem; font-weight: 700; color: #E2E8F4; letter-spacing: -0.3px; }
.top-bar-sub   { font-size: 0.78rem; color: #3A4A60; margin-top: 2px; }
.top-bar-right { display: flex; align-items: center; gap: 10px; }
.top-bar-period {
    font-size: 0.78rem; font-weight: 600;
    color: #3A8FF5; background: #0D2040;
    border: 1px solid #1A3A6A;
    padding: 6px 16px; border-radius: 20px; letter-spacing: 0.3px;
}
.top-bar-live {
    display: flex; align-items: center; gap: 6px;
    font-size: 0.72rem; color: #34C77B; font-weight: 600;
}
.dot-live {
    width: 7px; height: 7px;
    background: #34C77B; border-radius: 50%;
    animation: pulse-dot 2s infinite;
}
@keyframes pulse-dot {
    0%, 100% { opacity: 1; transform: scale(1); }
    50%       { opacity: 0.4; transform: scale(0.7); }
}

/* ─── KPI row ─── */
.kpi-row {
    display: grid;
    grid-template-columns: repeat(6, 1fr);
    gap: 10px;
    margin-bottom: 1.75rem;
}
.kpi {
    background: #0C0F1A;
    border: 1px solid #1A2035;
    border-radius: 12px;
    padding: 1.1rem 1.2rem 1rem;
    transition: border-color 0.18s, transform 0.18s, box-shadow 0.18s;
    position: relative;
    overflow: hidden;
}
.kpi::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
}
.kpi.blue::before   { background: linear-gradient(90deg,#3A8FF5,#60AFFF); }
.kpi.green::before  { background: linear-gradient(90deg,#34C77B,#5DE09B); }
.kpi.amber::before  { background: linear-gradient(90deg,#F0A429,#FFBF57); }
.kpi.red::before    { background: linear-gradient(90deg,#F05A5A,#FF8080); }
.kpi.purple::before { background: linear-gradient(90deg,#9B6EF3,#BF9FFF); }
.kpi.teal::before   { background: linear-gradient(90deg,#3ABFBF,#60DFDF); }
.kpi:hover {
    border-color: #2A3555;
    transform: translateY(-2px);
    box-shadow: 0 8px 32px rgba(0,0,0,0.4);
}
.kpi-icon  { font-size: 1.1rem; margin-bottom: 8px; opacity: 0.85; }
.kpi-label {
    font-size: 0.63rem; font-weight: 700; color: #3A4A60;
    text-transform: uppercase; letter-spacing: 1.1px; margin-bottom: 8px;
}
.kpi-val {
    font-size: 1.02rem; font-weight: 700; color: #E2E8F4;
    font-family: 'IBM Plex Mono', monospace; letter-spacing: -0.5px;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.kpi-val.blue   { color: #3A8FF5; }
.kpi-val.green  { color: #34C77B; }
.kpi-val.amber  { color: #F0A429; }
.kpi-val.purple { color: #9B6EF3; }
.kpi-val.teal   { color: #3ABFBF; }
.kpi-val.red    { color: #F05A5A; }
.kpi-sub { font-size: 0.65rem; color: #2A3A50; margin-top: 5px; }

/* ─── section heading ─── */
.sec-head {
    font-size: 0.63rem; font-weight: 700; color: #3A4A60;
    text-transform: uppercase; letter-spacing: 1.3px;
    margin: 1.5rem 0 0.85rem;
    display: flex; align-items: center; gap: 10px;
}
.sec-head::after { content: ""; flex: 1; height: 1px; background: #111827; }

/* ─── text inputs ─── */
.stTextInput > div > input {
    background: #0C0F1A !important;
    border: 1px solid #1E2640 !important;
    color: #C2CCDF !important;
    border-radius: 9px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 0.88rem !important;
    padding: 0.58rem 1rem !important;
    transition: border-color 0.15s, box-shadow 0.15s !important;
}
.stTextInput > div > input:focus {
    border-color: #3A8FF5 !important;
    box-shadow: 0 0 0 3px rgba(58,143,245,0.1) !important;
}
.stTextInput > div > input::placeholder { color: #2A3A50 !important; }
.stTextInput label {
    color: #3A4A60 !important; font-size: 0.65rem !important;
    font-weight: 700 !important; text-transform: uppercase !important;
    letter-spacing: 0.9px !important;
}

/* ─── company table ─── */
.ctable { width: 100%; border-collapse: collapse; font-size: 0.875rem; }
.ctable thead tr { background: #080B15; border-bottom: 1px solid #1A2035; }
.ctable thead th {
    padding: 11px 16px; text-align: left;
    font-size: 0.62rem; font-weight: 700; color: #3A4A60;
    text-transform: uppercase; letter-spacing: 1.1px;
}
.ctable thead th.num { text-align: right; }
.ctable tbody tr { border-bottom: 1px solid #0D1020; transition: background 0.1s; }
.ctable tbody tr:hover { background: #0F1428; }
.ctable td { padding: 13px 16px; color: #9AAEC8; vertical-align: middle; }
.ctable td.emp-name { color: #C2CCDF; font-weight: 600; font-size: 0.88rem; }
.ctable td.num {
    text-align: right;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem; color: #6A7A95;
}
.ctable td.total-num {
    text-align: right;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.88rem; font-weight: 700; color: #3A8FF5;
}
.pill {
    display: inline-block;
    padding: 2px 7px; border-radius: 4px;
    font-size: 0.6rem; font-weight: 700;
    margin-left: 5px; letter-spacing: 0.4px; vertical-align: middle;
}
.pill-irpj { background: #0D2040; color: #3A8FF5; border: 1px solid #1A3A6A; }
.pill-csll  { background: #1A0D3A; color: #9B6EF3; border: 1px solid #35196A; }
.bar-wrap { height: 3px; background: #111827; border-radius: 2px; margin-top: 8px; width: 100%; }
.bar-fill { height: 3px; border-radius: 2px; }
.share-badge {
    display: inline-block;
    font-size: 0.68rem; font-weight: 700;
    padding: 2px 8px; border-radius: 4px;
    font-family: 'IBM Plex Mono', monospace;
}
.share-high { background: #0D2A1A; color: #34C77B; }
.share-mid  { background: #2A1A0D; color: #F0A429; }
.share-low  { background: #1A0D0D; color: #F05A5A; }

/* ─── result count ─── */
.result-count {
    font-size: 0.68rem; color: #3A4A60; font-weight: 600;
    padding: 3px 10px; border: 1px solid #1A2035;
    border-radius: 20px; background: #0C0F1A;
}

/* ─── tabs ─── */
.stTabs [data-baseweb="tab-list"] {
    background: #0C0F1A !important;
    border: 1px solid #1A2035 !important;
    border-radius: 8px !important;
    padding: 3px !important; gap: 2px !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important; color: #3A4A60 !important;
    border-radius: 6px !important; font-weight: 500 !important;
    font-size: 0.84rem !important; padding: 0.42rem 1.1rem !important;
    border: none !important;
}
.stTabs [aria-selected="true"] {
    background: #151C30 !important; color: #C2CCDF !important;
}

/* ─── dataframe ─── */
[data-testid="stDataFrameResizable"] {
    border: 1px solid #1A2035 !important;
    border-radius: 10px !important; overflow: hidden !important;
}

/* ─── metrics ─── */
[data-testid="stMetric"] {
    background: #0C0F1A !important; border: 1px solid #1A2035 !important;
    border-radius: 10px !important; padding: 1rem 1.2rem !important;
}
[data-testid="stMetricLabel"] {
    color: #3A4A60 !important; font-size: 0.68rem !important;
    text-transform: uppercase !important; letter-spacing: 0.9px !important;
}
[data-testid="stMetricValue"] {
    color: #E2E8F4 !important; font-family: 'IBM Plex Mono', monospace !important;
}

/* ─── download button ─── */
.stDownloadButton > button {
    background: #0D2040 !important; color: #3A8FF5 !important;
    border: 1px solid #1A3A6A !important; border-radius: 7px !important;
    font-weight: 600 !important; font-size: 0.82rem !important;
    padding: 0.45rem 1.1rem !important; transition: all 0.15s !important;
}
.stDownloadButton > button:hover {
    background: #3A8FF5 !important; color: #fff !important;
    border-color: #3A8FF5 !important; transform: translateY(-1px) !important;
}

/* ─── regular button ─── */
.stButton > button {
    background: #111527 !important; color: #7A90B0 !important;
    border: 1px solid #1E2640 !important; border-radius: 7px !important;
    font-size: 0.82rem !important; transition: all 0.15s !important;
}
.stButton > button:hover {
    border-color: #3A8FF5 !important; color: #3A8FF5 !important;
}

/* ─── info box ─── */
.info-box {
    background: #0C0F1A; border: 1px dashed #1A2035; border-radius: 12px;
    padding: 3rem 1.5rem; text-align: center; color: #2E3A50; font-size: 0.88rem;
}
.info-box .ico { font-size: 2.2rem; margin-bottom: 12px; opacity: 0.6; }

/* ─── summary cards ─── */
.summary-row {
    display: grid; grid-template-columns: repeat(3, 1fr);
    gap: 12px; margin-bottom: 1.4rem;
}
.sum-card {
    background: #0C0F1A; border: 1px solid #1A2035;
    border-radius: 12px; padding: 1.1rem 1.4rem;
    display: flex; align-items: center; gap: 14px;
}
.sum-card-icon {
    width: 38px; height: 38px; border-radius: 9px;
    display: flex; align-items: center; justify-content: center;
    font-size: 1rem; flex-shrink: 0;
}
.sum-card-icon.blue   { background: #0D2040; }
.sum-card-icon.green  { background: #0D2A1A; }
.sum-card-icon.purple { background: #1A0D3A; }
.sum-card-label { font-size: 0.62rem; font-weight: 700; color: #3A4A60; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px; }
.sum-card-val   { font-size: 1rem; font-weight: 700; font-family: 'IBM Plex Mono', monospace; }
.sum-card-val.blue   { color: #3A8FF5; }
.sum-card-val.green  { color: #34C77B; }
.sum-card-val.purple { color: #9B6EF3; }

/* selectbox */
.stSelectbox > div > div {
    background: #0C0F1A !important; border: 1px solid #1E2640 !important;
    color: #C2CCDF !important; border-radius: 7px !important;
}
.stMultiSelect > div > div {
    background: #0C0F1A !important; border: 1px solid #1E2640 !important;
    color: #C2CCDF !important; border-radius: 7px !important;
}
[data-baseweb="tag"] { background: #111E38 !important; border: 1px solid #1A3A6A !important; }

/* sidebar helpers */
.sdiv { border: none; border-top: 1px solid #1A2035; margin: 1rem 0; }
.slabel {
    font-size: 0.63rem !important; font-weight: 700 !important;
    color: #2E3A50 !important; text-transform: uppercase !important;
    letter-spacing: 1.2px !important; margin-bottom: 4px !important;
    display: block !important;
}
</style>
""", unsafe_allow_html=True)

# =============================
# ESTADO
# =============================
if "empresa" not in st.session_state:
    st.session_state.empresa = None

# =============================
# UTILS
# =============================
def fmt(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def esc(texto):
    """Escapa texto vindo do usuário ou da planilha antes de injetar em HTML."""
    return html.escape(str(texto))


def extrair_codigo(texto):
    """Extrai o código numérico no início do texto (ex.: '126 - PARQUE EOLICO...' -> 126).
    Usado para casar empresas com os consolidados pelo CÓDIGO, que é estável,
    em vez do nome, que pode variar de grafia entre o relatório e a lista informada."""
    if texto is None:
        return None
    m = re.match(r"\s*0*(\d+)", str(texto))
    return int(m.group(1)) if m else None


def exportar_excel(dataframe: pd.DataFrame, nome_aba: str = "Dados") -> bytes:
    """Exporta com openpyxl — sem dependência de xlsxwriter."""
    buf  = io.BytesIO()
    aba  = nome_aba[:31]
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=aba)
        ws = writer.sheets[aba]
        try:
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            hdr_fill   = PatternFill("solid", fgColor="1F3A6A")
            hdr_font   = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
            alt_fill   = PatternFill("solid", fgColor="EFF3FB")
            row_border = Border(bottom=Side(style="thin", color="DDE3F0"))
            money_cols = {c for c in dataframe.columns if "vlr_" in str(c).lower()}

            # Cabeçalho
            for ci, cell in enumerate(ws[1], 1):
                cell.fill      = hdr_fill
                cell.font      = hdr_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[1].height = 26
                col_name = dataframe.columns[ci - 1]
                w = max(
                    len(str(col_name)) + 4,
                    dataframe[col_name].astype(str).str.len().max() + 2
                    if not dataframe.empty else 12
                )
                ws.column_dimensions[get_column_letter(ci)].width = min(w, 30)

            # Linhas de dados
            for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
                fill = alt_fill if ri % 2 == 0 else None
                for ci, cell in enumerate(row, 1):
                    col_name = dataframe.columns[ci - 1]
                    if fill:
                        cell.fill = fill
                    cell.border = row_border
                    cell.font   = Font(name="Calibri", size=10)
                    if col_name in money_cols:
                        try:
                            cell.value         = float(cell.value) if cell.value is not None else 0.0
                            cell.number_format = 'R$ #,##0.00'
                            cell.alignment     = Alignment(horizontal="right")
                        except Exception:
                            pass
            ws.freeze_panes = "A2"
        except Exception as e:
            # Não interrompe a exportação, mas avisa que o Excel saiu sem estilo
            st.warning(f"Excel exportado sem formatação avançada ({e}).", icon="⚠️")
    return buf.getvalue()


# =============================
# DADOS  —  UM ARQUIVO POR TRIMESTRE
# =============================
# base1 -> 1º trimestre, base2 -> 2º, base3 -> 3º, base4 -> 4º.
#
# IMPORTANTE: aqui o TRIMESTRE é definido pelo ARQUIVO, e não pela data de
# pagamento. No IRPJ/CSLL trimestral o recolhimento normalmente cai no
# trimestre seguinte ao da apuração (a apuração do 1º tri é paga em abril,
# que já é o 2º tri). Se o trimestre saísse da data, o pagamento apareceria
# no período errado. Por isso quem manda é o arquivo.
#
# Basta subir para o mesmo diretório do app: base1.xls, base2.xls, etc.
# (também aceita .xlsx). Suba só os trimestres que já existirem — o painel
# se ajusta sozinho aos trimestres presentes.
DIR_BASE = os.path.dirname(os.path.abspath(__file__))

ARQUIVOS_TRIMESTRE = {
    1: ["base1.xls", "base1.xlsx"],
    2: ["base2.xls", "base2.xlsx"],
    3: ["base3.xls", "base3.xlsx"],
    4: ["base4.xls", "base4.xlsx"],
}


@st.cache_data
def carregar_dados(caminho: str, mtime: float):
    """mtime entra como parte da chave de cache: se o arquivo mudar no disco,
    o Streamlit percebe e recarrega automaticamente."""
    return pd.read_excel(caminho)


def _localizar(candidatos):
    """Retorna o primeiro caminho existente dentre os nomes candidatos."""
    for nome in candidatos:
        caminho = os.path.join(DIR_BASE, nome)
        if os.path.exists(caminho):
            return caminho
    return None


frames = []
tris_no_sistema = []   # trimestres que têm arquivo carregado (1..4)

for _tri, _cands in ARQUIVOS_TRIMESTRE.items():
    _caminho = _localizar(_cands)
    if _caminho:
        _parcial = carregar_dados(_caminho, os.path.getmtime(_caminho)).copy()
        _parcial["__tri_arquivo"] = _tri
        frames.append(_parcial)
        tris_no_sistema.append(_tri)

MODO_POR_ARQUIVO = len(frames) > 0

if MODO_POR_ARQUIVO:
    df = pd.concat(frames, ignore_index=True)
else:
    # Compatibilidade: se nenhum base1..base4 existir, usa o base.xls antigo
    # e, nesse caso, o trimestre volta a ser derivado da data de pagamento.
    _legado = _localizar(["base.xls", "base.xlsx"])
    if not _legado:
        st.error(
            "Nenhum arquivo de base encontrado. Suba ao menos um de "
            "base1.xls … base4.xls (ou o base.xls antigo) na mesma pasta do app."
        )
        st.stop()
    df = carregar_dados(_legado, os.path.getmtime(_legado)).copy()
    df["__tri_arquivo"] = pd.NA

df = df[df["dsc_situacao"].astype(str).str.upper().str.strip() != "CANCELADA"]

colunas = [
    "cod_pagamento", "dat_pagamento", "cf_empresa",
    "vlr_principal", "vlr_multa", "vlr_juro_encargo", "vlr_outra_entidade", "vlr_total"
]
for col in colunas:
    if col not in df.columns:
        df[col] = 0

# Normaliza cod_pagamento removendo ".0" que aparece quando a coluna vem como float
df["cod_pagamento"] = (
    df["cod_pagamento"].astype(str).str.strip()
    .str.replace(r"\.0$", "", regex=True)
)

MAPA_IMPOSTO = {"2362": "IRPJ", "2089": "IRPJ", "2484": "CSLL", "2372": "CSLL"}
df = df[df["cod_pagamento"].isin(MAPA_IMPOSTO)]

# Código numérico da empresa (mais confiável que o nome para casar com os consolidados,
# já que o nome pode variar de grafia entre o relatório e a lista informada)
df["cod_empresa"] = df["cf_empresa"].apply(extrair_codigo)

valores = colunas[3:]
for col in valores:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

n_antes = len(df)
df["dat_pagamento"] = pd.to_datetime(df["dat_pagamento"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["dat_pagamento"])
n_descartados = n_antes - len(df)

df["Ano"] = df["dat_pagamento"].dt.year

# Trimestre: no modo por arquivo, quem manda é o arquivo (base1->Q1, ...).
# No modo compatibilidade (base.xls antigo), deriva da data como antes.
if MODO_POR_ARQUIVO:
    df["Trimestre"] = df["__tri_arquivo"].astype("Int64")
else:
    df["Trimestre"] = df["dat_pagamento"].dt.quarter

df["Imposto"] = df["cod_pagamento"].map(MAPA_IMPOSTO)

if df.empty:
    st.error("Sem dados após tratamento.")
    st.stop()

# Trimestres realmente presentes (após todos os filtros)
tris_disponiveis = sorted(int(t) for t in df["Trimestre"].dropna().unique())

# =============================
# SIDEBAR — só período
# =============================
with st.sidebar:
    st.markdown("""
    <div style="padding:1.5rem 0.5rem 1.1rem;border-bottom:1px solid #1A2035;margin-bottom:1.25rem;">
        <div style="font-size:1.5rem;margin-bottom:8px;">📊</div>
        <div style="font-size:0.92rem;font-weight:700;color:#C2CCDF;">Painel Tributário</div>
        <div style="font-size:0.68rem;color:#2E3A50;margin-top:3px;letter-spacing:0.5px;">IRPJ &amp; CSLL</div>
    </div>
    """, unsafe_allow_html=True)

    if n_descartados > 0:
        st.warning(f"{n_descartados} registro(s) com data inválida foram ignorados.", icon="⚠️")

    st.markdown('<span class="slabel">Ano de Referência</span>', unsafe_allow_html=True)
    ano = st.selectbox("Ano", sorted(df["Ano"].unique()), label_visibility="collapsed")

    st.markdown("<hr class='sdiv'>", unsafe_allow_html=True)
    st.markdown('<span class="slabel">Trimestre</span>', unsafe_allow_html=True)
    tri = st.selectbox(
        "Trimestre", tris_disponiveis,
        format_func=lambda x: f"Q{x} — {['Jan–Mar','Abr–Jun','Jul–Set','Out–Dez'][x-1]}",
        label_visibility="collapsed"
    )

    st.markdown("<hr class='sdiv'>", unsafe_allow_html=True)

    # mini‑resumo na sidebar
    df_side = df[(df["Ano"] == ano) & (df["Trimestre"] == tri)]
    n_side  = df_side["cf_empresa"].nunique()
    t_side  = df_side["vlr_total"].sum()
    st.markdown(f"""
    <div style="padding:0.5rem 0 0.75rem;">
        <div class="slabel" style="margin-bottom:10px;">Resumo Rápido</div>
        <div style="display:flex;flex-direction:column;gap:8px;">
            <div style="background:#111527;border:1px solid #1E2640;border-radius:8px;padding:0.7rem 0.9rem;">
                <div style="font-size:0.58rem;color:#2E3A50;text-transform:uppercase;letter-spacing:1px;margin-bottom:3px;">Empresas</div>
                <div style="font-size:1.1rem;font-weight:700;color:#3ABFBF;font-family:'IBM Plex Mono',monospace;">{n_side}</div>
            </div>
            <div style="background:#111527;border:1px solid #1E2640;border-radius:8px;padding:0.7rem 0.9rem;">
                <div style="font-size:0.58rem;color:#2E3A50;text-transform:uppercase;letter-spacing:1px;margin-bottom:3px;">Total do Período</div>
                <div style="font-size:0.85rem;font-weight:700;color:#3A8FF5;font-family:'IBM Plex Mono',monospace;">R$ {fmt(t_side)}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # indicador de quais arquivos/trimestres foram carregados — ajuda a
    # conferir se o upload no GitHub funcionou.
    if MODO_POR_ARQUIVO:
        carregados_txt = ", ".join(f"Q{t}" for t in sorted(tris_no_sistema))
        faltando_txt   = ", ".join(f"Q{t}" for t in (1, 2, 3, 4) if t not in tris_no_sistema)
        st.markdown(f"""
        <div style="padding:0.25rem 0 0.5rem;">
            <div class="slabel" style="margin-bottom:6px;">Arquivos no Sistema</div>
            <div style="font-size:0.62rem;color:#34C77B;">Carregados: {carregados_txt or '—'}</div>
            <div style="font-size:0.62rem;color:#3A4A60;margin-top:2px;">Ausentes: {faltando_txt or 'nenhum'}</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="padding:0.25rem 0 0.5rem;">
            <div class="slabel" style="margin-bottom:6px;">Fonte</div>
            <div style="font-size:0.62rem;color:#F0A429;">base.xls (modo compatibilidade — trimestre pela data)</div>
        </div>
        """, unsafe_allow_html=True)

# =============================
# FILTRO POR PERÍODO
# =============================
df_tri = df[(df["Ano"] == ano) & (df["Trimestre"] == tri)].copy()

# Base do ACUMULADO: todos os trimestres do ano selecionado (ignora o Trimestre).
df_ano = df[df["Ano"] == ano].copy()

# Se a empresa selecionada no Dashboard não existir mais neste período, limpa a seleção
if st.session_state.empresa and st.session_state.empresa not in df_tri["cf_empresa"].unique():
    st.session_state.empresa = None

# =============================
# TOP BAR
# =============================
tl = ['Janeiro–Março', 'Abril–Junho', 'Julho–Setembro', 'Outubro–Dezembro'][tri - 1]
st.markdown(f"""
<div class="top-bar">
    <div class="top-bar-left">
        <div class="top-bar-icon">📊</div>
        <div>
            <div class="top-bar-title">Painel de Pagamentos Tributários</div>
            <div class="top-bar-sub">IRPJ — Imposto de Renda Pessoa Jurídica &nbsp;|&nbsp; CSLL — Contribuição Social sobre o Lucro Líquido</div>
        </div>
    </div>
    <div class="top-bar-right">
        <div class="top-bar-live"><div class="dot-live"></div>Ativo</div>
        <div class="top-bar-period">{ano} &nbsp;·&nbsp; Q{tri} &nbsp;·&nbsp; {tl}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =============================
# KPIs
# =============================
t_geral = df_tri["vlr_total"].sum()
t_irpj  = df_tri[df_tri["Imposto"] == "IRPJ"]["vlr_total"].sum()
t_csll  = df_tri[df_tri["Imposto"] == "CSLL"]["vlr_total"].sum()
t_multa = df_tri["vlr_multa"].sum()
t_juros = df_tri["vlr_juro_encargo"].sum()
n_emp   = df_tri["cf_empresa"].nunique()

pct_irpj  = (t_irpj  / t_geral * 100) if t_geral > 0 else 0
pct_csll  = (t_csll  / t_geral * 100) if t_geral > 0 else 0
pct_multa = (t_multa / t_geral * 100) if t_geral > 0 else 0

st.markdown(f"""
<div class="kpi-row">
    <div class="kpi blue">
        <div class="kpi-icon">💰</div>
        <div class="kpi-label">Total Geral</div>
        <div class="kpi-val blue">R$ {fmt(t_geral)}</div>
        <div class="kpi-sub">100% do período</div>
    </div>
    <div class="kpi green">
        <div class="kpi-icon">🟩</div>
        <div class="kpi-label">IRPJ</div>
        <div class="kpi-val green">R$ {fmt(t_irpj)}</div>
        <div class="kpi-sub">{pct_irpj:.1f}% do total</div>
    </div>
    <div class="kpi purple">
        <div class="kpi-icon">🟪</div>
        <div class="kpi-label">CSLL</div>
        <div class="kpi-val purple">R$ {fmt(t_csll)}</div>
        <div class="kpi-sub">{pct_csll:.1f}% do total</div>
    </div>
    <div class="kpi amber">
        <div class="kpi-icon">⚠️</div>
        <div class="kpi-label">Multas</div>
        <div class="kpi-val amber">R$ {fmt(t_multa)}</div>
        <div class="kpi-sub">{pct_multa:.1f}% do total</div>
    </div>
    <div class="kpi red">
        <div class="kpi-icon">📈</div>
        <div class="kpi-label">Juros & Encargos</div>
        <div class="kpi-val red">R$ {fmt(t_juros)}</div>
        <div class="kpi-sub">acréscimos legais</div>
    </div>
    <div class="kpi teal">
        <div class="kpi-icon">🏢</div>
        <div class="kpi-label">Empresas</div>
        <div class="kpi-val teal">{n_emp}</div>
        <div class="kpi-sub">contribuintes ativos</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =============================
# ABAS
# =============================
abas = st.tabs(["  📊 Dashboard  ", "  📑 Analítico  ", "  📈 Consolidado  "])

# ─── DASHBOARD ────────────────────────────────────────
with abas[0]:

    # Filtros inline: busca + tipo de imposto + exportar
    col_s, col_imp, col_exp = st.columns([3, 1.4, 1])
    with col_s:
        busca = st.text_input(
            "Pesquisar empresa",
            placeholder="🔍  Nome, código ou parte do nome...",
            key="busca_dash"
        )
    with col_imp:
        imp_sel = st.multiselect(
            "Tipo de imposto",
            ["IRPJ", "CSLL"],
            default=["IRPJ", "CSLL"],
            key="imp_dash"
        )

    # Dataset filtrado pelo imposto selecionado
    df_dash = df_tri[df_tri["Imposto"].isin(imp_sel)] if imp_sel else df_tri.copy()

    # Ranking por total
    rank = (
        df_dash.groupby("cf_empresa")["vlr_total"]
        .sum().sort_values(ascending=False)
    )

    # Aplica busca por texto
    if busca.strip():
        rank = rank[rank.index.astype(str).str.upper().str.contains(busca.strip().upper())]

    max_val  = rank.max() if not rank.empty else 1
    n_result = len(rank)

    with col_exp:
        df_dash_exp = (
            df_dash.groupby("cf_empresa")[valores]
            .sum().reset_index().sort_values("vlr_total", ascending=False)
        )
        st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
        st.download_button(
            "⬇  Exportar (.xlsx)",
            data=exportar_excel(df_dash_exp, "Dashboard"),
            file_name=f"dashboard_{ano}_Q{tri}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Cabeçalho da lista com contagem
    hc, rc = st.columns([6, 1])
    with hc:
        st.markdown('<div class="sec-head">Empresas</div>', unsafe_allow_html=True)
    with rc:
        st.markdown(
            f"<div style='padding-top:1.6rem;text-align:right'>"
            f"<span class='result-count'>{n_result} resultado{'s' if n_result != 1 else ''}</span>"
            f"</div>",
            unsafe_allow_html=True
        )

    if rank.empty:
        msg = f'para <strong>"{esc(busca)}"</strong>' if busca.strip() else "neste período"
        st.markdown(f"""
        <div class="info-box">
            <div class="ico">🔍</div>
            <p>Nenhuma empresa encontrada {msg}.</p>
        </div>""", unsafe_allow_html=True)
    else:
        # Pré-agrega uma única vez em vez de refiltrar o dataframe a cada empresa do loop
        agg_imp = df_dash.groupby(["cf_empresa", "Imposto"])["vlr_total"].sum().unstack(fill_value=0)
        agg_ext = df_dash.groupby("cf_empresa")[["vlr_multa", "vlr_juro_encargo"]].sum()

        rows_html = ""
        for emp, total in rank.items():
            irpj  = agg_imp.loc[emp, "IRPJ"] if emp in agg_imp.index and "IRPJ" in agg_imp.columns else 0
            csll  = agg_imp.loc[emp, "CSLL"] if emp in agg_imp.index and "CSLL" in agg_imp.columns else 0
            multa = agg_ext.loc[emp, "vlr_multa"] if emp in agg_ext.index else 0
            juros = agg_ext.loc[emp, "vlr_juro_encargo"] if emp in agg_ext.index else 0
            pct   = (total / max_val * 100) if max_val > 0 else 0
            share = (total / t_geral * 100) if t_geral > 0 else 0

            share_class = "share-high" if share >= 20 else ("share-mid" if share >= 8 else "share-low")
            bar_color   = "#3A8FF5" if irpj >= csll else "#9B6EF3"

            pill_i = '<span class="pill pill-irpj">IRPJ</span>' if irpj > 0 else ""
            pill_c = '<span class="pill pill-csll">CSLL</span>' if csll > 0 else ""

            rows_html += f"""
            <tr>
                <td class="emp-name">
                    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:4px;">
                        <span>{esc(emp)}{pill_i}{pill_c}</span>
                        <span class="share-badge {share_class}">{share:.1f}%</span>
                    </div>
                    <div class="bar-wrap">
                        <div class="bar-fill" style="width:{pct:.1f}%;background:linear-gradient(90deg,{bar_color},{bar_color}66);"></div>
                    </div>
                </td>
                <td class="num">R$ {fmt(irpj)}</td>
                <td class="num">R$ {fmt(csll)}</td>
                <td class="num">R$ {fmt(multa)}</td>
                <td class="num">R$ {fmt(juros)}</td>
                <td class="total-num">R$ {fmt(total)}</td>
            </tr>"""

        st.markdown(f"""
        <div style="background:#0A0D18;border:1px solid #1A2035;border-radius:12px;overflow:hidden;">
        <table class="ctable">
            <thead><tr>
                <th>Empresa &nbsp;<span style="opacity:.35;font-weight:400;font-size:0.6rem;">% participação</span></th>
                <th class="num">IRPJ</th>
                <th class="num">CSLL</th>
                <th class="num">Multas</th>
                <th class="num">Juros</th>
                <th class="num">Total</th>
            </tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
        </div>""", unsafe_allow_html=True)

        st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
        emp_sel = st.selectbox(
            "Abrir empresa no Analítico →",
            ["— selecione uma empresa —"] + list(rank.index),
            key="sel_emp_dash"
        )
        if emp_sel != "— selecione uma empresa —":
            st.session_state.empresa = emp_sel
            st.success(f"✓  **{emp_sel}** carregada — acesse a aba **📑 Analítico**")

# ─── ANALÍTICO ────────────────────────────────────────
with abas[1]:
    emp = st.session_state.empresa

    if not emp:
        st.markdown("""
        <div class="info-box">
            <div class="ico">🏢</div>
            <p>Selecione uma empresa no Dashboard para ver o detalhamento completo.</p>
        </div>""", unsafe_allow_html=True)
    else:
        df_emp = df_tri[df_tri["cf_empresa"] == emp].copy()
        st.markdown(f'<div class="sec-head">{esc(emp)}</div>', unsafe_allow_html=True)

        total_emp = df_emp["vlr_total"].sum()
        irpj_emp  = df_emp[df_emp["Imposto"] == "IRPJ"]["vlr_total"].sum()
        csll_emp  = df_emp[df_emp["Imposto"] == "CSLL"]["vlr_total"].sum()

        st.markdown(f"""
        <div class="summary-row">
            <div class="sum-card">
                <div class="sum-card-icon blue">💰</div>
                <div>
                    <div class="sum-card-label">Total da Empresa</div>
                    <div class="sum-card-val blue">R$ {fmt(total_emp)}</div>
                </div>
            </div>
            <div class="sum-card">
                <div class="sum-card-icon green">🟩</div>
                <div>
                    <div class="sum-card-label">IRPJ</div>
                    <div class="sum-card-val green">R$ {fmt(irpj_emp)}</div>
                </div>
            </div>
            <div class="sum-card">
                <div class="sum-card-icon purple">🟪</div>
                <div>
                    <div class="sum-card-label">CSLL</div>
                    <div class="sum-card-val purple">R$ {fmt(csll_emp)}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Principal", f"R$ {fmt(df_emp['vlr_principal'].sum())}")
        c2.metric("Multa",     f"R$ {fmt(df_emp['vlr_multa'].sum())}")
        c3.metric("Juros",     f"R$ {fmt(df_emp['vlr_juro_encargo'].sum())}")
        c4.metric("Outros",    f"R$ {fmt(df_emp['vlr_outra_entidade'].sum())}")
        c5.metric("Registros", str(len(df_emp)))

        st.markdown('<div class="sec-head">Registros de Pagamento</div>', unsafe_allow_html=True)

        col_s, col_dl = st.columns([4, 1])
        with col_s:
            busca_a = st.text_input(
                "Filtrar registros",
                placeholder="🔍  Código, data, tipo de imposto...",
                key="busca_analitico"
            )

        cols_exib = ["dat_pagamento", "cod_pagamento", "Imposto"] + valores
        df_det = df_emp[cols_exib].copy()
        df_det["dat_pagamento"] = df_det["dat_pagamento"].dt.strftime("%d/%m/%Y")

        if busca_a.strip():
            mask = df_det.apply(
                lambda row: busca_a.strip().upper() in " ".join(row.astype(str).str.upper()), axis=1
            )
            df_det = df_det[mask]

        with col_dl:
            st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
            export_det = df_emp[cols_exib].copy()
            export_det["dat_pagamento"] = export_det["dat_pagamento"].dt.strftime("%d/%m/%Y")
            st.download_button(
                "⬇  Exportar (.xlsx)",
                data=exportar_excel(export_det, "Analítico"),
                file_name=f"analitico_{emp}_{ano}_Q{tri}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_analitico"
            )

        if df_det.empty:
            st.markdown(
                '<div class="info-box"><p>Nenhum registro encontrado para o filtro informado.</p></div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                f"<div style='font-size:0.68rem;color:#3A4A60;margin-bottom:6px;'>"
                f"{len(df_det)} registro(s) encontrado(s)</div>",
                unsafe_allow_html=True
            )
            st.dataframe(
                df_det.style.format({v: "R$ {:,.2f}" for v in valores}),
                use_container_width=True,
                hide_index=True,
            )

# ─── CONSOLIDADO ──────────────────────────────────────
def render_bloco_consolidado(df_base: pd.DataFrame, chave_export: str, sufixo_arquivo: str = None):
    """Renderiza o resumo padrão (KPIs + por empresa + por imposto) para um
    subconjunto qualquer do dataframe. Reutilizado na Visão Geral, no Acumulado
    e em cada consolidado por grupo.

    sufixo_arquivo controla o nome dos arquivos exportados; se None, usa o
    período selecionado (ano_Qtri). No Acumulado passamos 'ano_acumulado'."""

    if sufixo_arquivo is None:
        sufixo_arquivo = f"{ano}_Q{tri}"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Principal", f"R$ {fmt(df_base['vlr_principal'].sum())}")
    c2.metric("Multa",     f"R$ {fmt(df_base['vlr_multa'].sum())}")
    c3.metric("Juros",     f"R$ {fmt(df_base['vlr_juro_encargo'].sum())}")
    c4.metric("Total",     f"R$ {fmt(df_base['vlr_total'].sum())}")

    # ── Por Empresa ──
    st.markdown('<div class="sec-head">Por Empresa</div>', unsafe_allow_html=True)
    col_s2, col_dl2 = st.columns([4, 1])
    with col_s2:
        busca_c = st.text_input(
            "Filtrar empresas",
            placeholder="🔍  Nome ou código...",
            key=f"busca_{chave_export}"
        )

    consol = (
        df_base.groupby("cf_empresa")[valores]
        .sum().reset_index().sort_values("vlr_total", ascending=False)
    )
    if busca_c.strip():
        consol = consol[
            consol["cf_empresa"].astype(str).str.upper().str.contains(busca_c.strip().upper())
        ]

    with col_dl2:
        st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
        st.download_button(
            "⬇  Exportar (.xlsx)",
            data=exportar_excel(consol, "Por Empresa"),
            file_name=f"{chave_export}_empresa_{sufixo_arquivo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{chave_export}_emp"
        )

    if consol.empty:
        st.markdown(
            '<div class="info-box"><p>Sem dados para o filtro atual.</p></div>',
            unsafe_allow_html=True
        )
    else:
        st.dataframe(
            consol.style.format({v: "R$ {:,.2f}" for v in valores}),
            use_container_width=True, hide_index=True,
        )

    # ── Por Imposto ──
    st.markdown('<div class="sec-head">Por Imposto</div>', unsafe_allow_html=True)
    por_imp = df_base.groupby("Imposto")[valores].sum().reset_index()

    col_ti, col_dl3 = st.columns([4, 1])
    with col_dl3:
        st.download_button(
            "⬇  Exportar (.xlsx)",
            data=exportar_excel(por_imp, "Por Imposto"),
            file_name=f"{chave_export}_imposto_{sufixo_arquivo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{chave_export}_imp"
        )
    st.dataframe(
        por_imp.style.format({v: "R$ {:,.2f}" for v in valores}),
        use_container_width=True, hide_index=True,
    )


with abas[2]:
    nomes_subabas = (
        ["  🗂 Visão Geral  ", "  Σ Acumulado  "]
        + [f"  {nome}  " for nome in CONSOLIDADOS.keys()]
    )
    subabas = st.tabs(nomes_subabas)

    # ── Visão Geral — trimestre selecionado, empresas dos consolidados ──
    # Soma apenas as empresas que pertencem a algum consolidado (união de
    # todos os grupos, pelo código numérico), não a base inteira.
    with subabas[0]:
        st.markdown('<div class="sec-head">Resumo do Período — Todas as Empresas dos Consolidados</div>', unsafe_allow_html=True)

        codigos_todos = sorted({
            c
            for lista in CONSOLIDADOS.values()
            for c in (extrair_codigo(e) for e in lista)
            if c is not None
        })
        df_visao_geral = df_tri[df_tri["cod_empresa"].isin(codigos_todos)]

        if df_visao_geral.empty:
            st.markdown(
                '<div class="info-box"><p>Nenhuma empresa consolidada teve pagamentos neste período.</p></div>',
                unsafe_allow_html=True
            )
        else:
            render_bloco_consolidado(df_visao_geral, "consolidado_geral")

    # ── ACUMULADO — soma de TODOS os trimestres presentes no ano ──
    with subabas[1]:
        rotulos_tri = {1: "Q1", 2: "Q2", 3: "Q3", 4: "Q4"}
        tris_ano = sorted(int(t) for t in df_ano["Trimestre"].dropna().unique())
        lista_tri_txt = ", ".join(rotulos_tri.get(t, f"Q{t}") for t in tris_ano) or "—"

        st.markdown(
            f'<div class="sec-head">Acumulado {ano} — soma dos trimestres no sistema '
            f'({lista_tri_txt})</div>',
            unsafe_allow_html=True
        )
        st.caption(
            "Esta aba independe do trimestre selecionado na barra lateral: "
            "soma todos os trimestres já carregados do ano de referência."
        )

        # Universo consolidado = união de todos os códigos de todos os grupos
        codigos_todos = sorted({
            c
            for lista in CONSOLIDADOS.values()
            for c in (extrair_codigo(e) for e in lista)
            if c is not None
        })
        df_acum = df_ano[df_ano["cod_empresa"].isin(codigos_todos)]

        if df_acum.empty:
            st.markdown(
                '<div class="info-box"><p>Nenhuma empresa consolidada teve pagamentos nos trimestres carregados.</p></div>',
                unsafe_allow_html=True
            )
        else:
            # ── Acumulado por consolidado (grupo) ──
            st.markdown('<div class="sec-head">Acumulado por Consolidado</div>', unsafe_allow_html=True)

            linhas_grupo = []
            for nome_grupo, lista_empresas in CONSOLIDADOS.items():
                if nome_grupo == "ESA":
                    continue  # ESA = união de todos; representada no total geral abaixo
                cods = {c for c in (extrair_codigo(e) for e in lista_empresas) if c is not None}
                sub = df_acum[df_acum["cod_empresa"].isin(cods)]
                if sub.empty:
                    continue
                linhas_grupo.append({
                    "Consolidado":      nome_grupo,
                    "IRPJ":             sub.loc[sub["Imposto"] == "IRPJ", "vlr_total"].sum(),
                    "CSLL":             sub.loc[sub["Imposto"] == "CSLL", "vlr_total"].sum(),
                    "vlr_multa":        sub["vlr_multa"].sum(),
                    "vlr_juro_encargo": sub["vlr_juro_encargo"].sum(),
                    "vlr_total":        sub["vlr_total"].sum(),
                })

            if linhas_grupo:
                df_grupos = pd.DataFrame(linhas_grupo).sort_values("vlr_total", ascending=False)
                col_gi, col_gdl = st.columns([4, 1])
                with col_gdl:
                    st.download_button(
                        "⬇  Exportar (.xlsx)",
                        data=exportar_excel(df_grupos, "Acumulado grupos"),
                        file_name=f"acumulado_grupos_{ano}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_acum_grupos"
                    )
                st.dataframe(
                    df_grupos.style.format({
                        c: "R$ {:,.2f}" for c in
                        ["IRPJ", "CSLL", "vlr_multa", "vlr_juro_encargo", "vlr_total"]
                    }),
                    use_container_width=True, hide_index=True,
                )
                st.caption(
                    "Algumas empresas pertencem a mais de um consolidado (ex.: código 126, "
                    "em Geração e Sobradinho), então a soma dos grupos pode não bater com o "
                    "Total Consolidado abaixo, que conta cada empresa uma única vez."
                )

            # ── Total consolidado acumulado (todas as empresas, sem dupla contagem) ──
            st.markdown('<div class="sec-head">Total Consolidado Acumulado</div>', unsafe_allow_html=True)
            render_bloco_consolidado(df_acum, "acumulado_total", sufixo_arquivo=f"{ano}_acumulado")

    # ── Um sub-tab por grupo (trimestre selecionado), conforme CONSOLIDADOS ──
    for i, (nome_grupo, lista_empresas) in enumerate(CONSOLIDADOS.items(), start=2):
        with subabas[i]:
            st.markdown(f'<div class="sec-head">Consolidado {esc(nome_grupo)}</div>', unsafe_allow_html=True)

            if not lista_empresas:
                st.markdown(f"""
                <div class="info-box">
                    <div class="ico">🏗️</div>
                    <p>Nenhuma empresa configurada ainda para o consolidado <strong>{esc(nome_grupo)}</strong>.<br>
                    Assim que a lista de empresas for definida no dicionário <code>CONSOLIDADOS</code>, esta aba será preenchida automaticamente.</p>
                </div>""", unsafe_allow_html=True)
                continue

            # Casa empresas pelo CÓDIGO numérico (ex.: 126), não pelo nome — o nome pode
            # vir com grafia levemente diferente no relatório, mas o código é estável.
            codigos_grupo = sorted({c for c in (extrair_codigo(e) for e in lista_empresas) if c is not None})
            sem_codigo = [e for e in lista_empresas if extrair_codigo(e) is None]

            df_grupo = df_tri[df_tri["cod_empresa"].isin(codigos_grupo)]

            codigos_encontrados = set(df_grupo["cod_empresa"].dropna().unique())
            codigos_faltantes   = [c for c in codigos_grupo if c not in codigos_encontrados]

            if codigos_faltantes:
                nomes_faltantes = [
                    e for e in lista_empresas if extrair_codigo(e) in codigos_faltantes
                ]
                st.info(
                    f"{len(codigos_encontrados)} de {len(codigos_grupo)} empresa(s) do grupo têm pagamentos neste período. "
                    f"Sem movimento (código): {', '.join(nomes_faltantes[:10])}"
                    + ("..." if len(nomes_faltantes) > 10 else ""),
                    icon="ℹ️"
                )
            if sem_codigo:
                st.warning(
                    f"Não foi possível identificar o código destas empresas na lista do consolidado "
                    f"(verifique se começam com o número): {', '.join(sem_codigo)}",
                    icon="⚠️"
                )

            if df_grupo.empty:
                st.markdown(
                    '<div class="info-box"><p>Nenhum pagamento encontrado para este grupo no período selecionado.</p></div>',
                    unsafe_allow_html=True
                )
            else:
                render_bloco_consolidado(df_grupo, f"consolidado_{nome_grupo.lower().replace(' ', '_')}")
