import base64
import html
from pathlib import Path
from io import BytesIO
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

# =====================================================
# CONFIGURAÇÃO GERAL - EIrox DRE ONLINE
# =====================================================
st.set_page_config(
    page_title="Eirox DRE Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "DRE_Consolidado_Moderno.xlsx"
DATA_FILE_ROOT = BASE_DIR / "DRE_Consolidado_Moderno.xlsx"
LOGO_FILE = BASE_DIR / "assets" / "logo_eirox.png"

MESES_ORDEM = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}

# =====================================================
# CSS - DARK ENTERPRISE EIrox
# =====================================================
st.markdown(
    """
<style>
    .stApp { background: radial-gradient(circle at top left, #0B1B2D 0%, #07111F 42%, #050A12 100%); }
    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #07111F 0%, #0E1A2B 100%); border-right: 1px solid rgba(0,175,255,.25); }
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; max-width: 98%; }
    .eirox-hero {
        border: 1px solid rgba(0,175,255,.25);
        background: linear-gradient(135deg, rgba(5,12,24,.96), rgba(12,31,52,.94));
        border-radius: 22px;
        padding: 22px 26px;
        margin-bottom: 20px;
        box-shadow: 0 18px 45px rgba(0,0,0,.35);
    }
    .eirox-title { font-size: 34px; font-weight: 800; color: #F8FBFF; margin: 0; letter-spacing: .2px; }
    .eirox-subtitle { color: #9FB4CC; margin-top: 6px; font-size: 15px; }
    .kpi-card {
        background: linear-gradient(145deg, rgba(14,26,43,.98), rgba(8,17,31,.98));
        border: 1px solid rgba(0,175,255,.20);
        border-radius: 18px;
        padding: 18px 18px;
        box-shadow: 0 10px 28px rgba(0,0,0,.28);
        min-height: 112px;
    }
    .kpi-title { color:#91A4BA; font-size: 13px; font-weight: 600; margin-bottom: 8px; text-transform: uppercase; letter-spacing:.4px; }
    .kpi-value { color:#FFFFFF; font-size: 26px; font-weight: 800; line-height: 1.15; }
    .kpi-help { color:#00AFFF; font-size: 12px; margin-top: 8px; }
    .section-title { font-size: 24px; font-weight: 800; margin: 18px 0 10px 0; color:#FFFFFF; }
    .section-note { color:#9FB4CC; margin-bottom: 14px; }
    div[data-testid="stDataFrame"] { border-radius: 16px; overflow: hidden; border: 1px solid rgba(0,175,255,.18); }
    .small-muted { color:#9FB4CC; font-size: 13px; }
</style>
""",
    unsafe_allow_html=True,
)

# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================
def image_to_base64(path: Path) -> str:
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def fmt_moeda(v):
    try:
        v = float(v)
    except Exception:
        return "R$ 0,00"
    s = f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s


def fmt_pct(v):
    try:
        v = float(v)
    except Exception:
        return "0,00%"
    return f"{v*100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def normalizar_mes(m):
    return str(m).strip().lower()


def mes_key(m):
    m = normalizar_mes(m)
    if m == "sem mês" or m == "sem mes":
        return (99, 99)
    try:
        abrev, ano = m.split("/")
        return (int("20" + ano), MESES_ORDEM.get(abrev[:3], 98))
    except Exception:
        return (98, 98)


def ordenar_meses(meses):
    return sorted([m for m in meses if pd.notna(m)], key=mes_key)


def mes_para_data(m):
    """Converte rótulos como jan/26 em uma data no primeiro dia do mês."""
    m_norm = normalizar_mes(m)
    try:
        abrev, ano = m_norm.split('/')
        mes = MESES_ORDEM.get(abrev[:3])
        if mes is None:
            return None
        ano_int = int(ano)
        if ano_int < 100:
            ano_int += 2000
        return date(ano_int, mes, 1)
    except Exception:
        return None


def filtrar_meses_fechados(meses):
    """Mantém somente meses encerrados, excluindo automaticamente o mês atual e meses futuros."""
    hoje = date.today()
    inicio_mes_atual = date(hoje.year, hoje.month, 1)
    fechados = []
    for mes in meses:
        dt = mes_para_data(mes)
        if dt is not None and dt < inicio_mes_atual:
            fechados.append(mes)
    return ordenar_meses(fechados)


def safe_read_excel(file_or_path):
    return pd.ExcelFile(file_or_path)


@st.cache_data(show_spinner=False)
def carregar_workbook_from_bytes(content: bytes):
    bio = BytesIO(content)
    xl = safe_read_excel(bio)
    sheets = {}
    for s in xl.sheet_names:
        try:
            sheets[s] = pd.read_excel(BytesIO(content), sheet_name=s)
        except Exception:
            pass
    return sheets


@st.cache_data(show_spinner=False)
def carregar_workbook_from_path(path_str: str):
    xl = safe_read_excel(path_str)
    sheets = {}
    for s in xl.sheet_names:
        try:
            sheets[s] = pd.read_excel(path_str, sheet_name=s)
        except Exception:
            pass
    return sheets


def escolher_arquivo():
    """Carrega a base automaticamente do banco/arquivo padrão do projeto.
    Não exibe upload na tela.
    Prioridade:
    1) data/DRE_Consolidado_Moderno.xlsx
    2) DRE_Consolidado_Moderno.xlsx na raiz do projeto
    """
    if DATA_FILE.exists():
        return carregar_workbook_from_path(str(DATA_FILE)), str(DATA_FILE)
    if DATA_FILE_ROOT.exists():
        return carregar_workbook_from_path(str(DATA_FILE_ROOT)), str(DATA_FILE_ROOT)
    return {}, "Nenhum arquivo encontrado"


def obter_df(sheets, nome, default_cols=None):
    if nome in sheets:
        return sheets[nome].copy()
    if default_cols:
        return pd.DataFrame(columns=default_cols)
    return pd.DataFrame()


def ultimo_mes_com_receita(dados):
    if dados.empty or "Mês" not in dados.columns or "Linha DRE" not in dados.columns:
        return None
    rec = dados[dados["Linha DRE"].astype(str).str.contains("RECEITA OPERACIONAL BRUTA", case=False, na=False)]
    rec = rec[rec["Valor"].fillna(0) != 0]
    rec = rec[~rec["Mês"].astype(str).str.lower().isin(["sem mês", "sem mes"])]
    if rec.empty:
        meses = ordenar_meses(dados["Mês"].dropna().unique())
        meses = [m for m in meses if str(m).lower() not in ["sem mês", "sem mes"]]
        return meses[-1] if meses else None
    meses = ordenar_meses(rec["Mês"].unique())
    return meses[-1] if meses else None


def valor_linha_mes(dados, linha_contains, mes):
    if dados.empty or mes is None:
        return 0.0
    f = dados["Linha DRE"].astype(str).str.contains(linha_contains, case=False, regex=False, na=False) & (dados["Mês"].astype(str) == str(mes))
    if not f.any():
        return 0.0
    return float(pd.to_numeric(dados.loc[f, "Valor"], errors="coerce").fillna(0).sum())


def percentual_linha_mes(dados, linha_contains, mes):
    if dados.empty or mes is None:
        return 0.0
    f = dados["Linha DRE"].astype(str).str.contains(linha_contains, case=False, regex=False, na=False) & (dados["Mês"].astype(str) == str(mes))
    if not f.any() or "% Receita" not in dados.columns:
        return 0.0
    return float(pd.to_numeric(dados.loc[f, "% Receita"], errors="coerce").fillna(0).mean())


def montar_dre_para_tabela(dre_estruturada, meses_selecionados):
    if dre_estruturada.empty:
        return pd.DataFrame()
    base_cols = [c for c in ["Ordem", "Seção", "Linha DRE", "Nível", "Tipo"] if c in dre_estruturada.columns]
    cols = base_cols[:]
    for m in meses_selecionados:
        for sufixo in ["Valor", "%"]:
            c = f"{m} {sufixo}"
            if c in dre_estruturada.columns:
                cols.append(c)
    df = dre_estruturada[cols].copy()
    for c in df.columns:
        if str(c).endswith(" Valor"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).apply(fmt_moeda)
        elif str(c).endswith(" %"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).apply(fmt_pct)
    # esconde colunas técnicas na exibição
    display = df.drop(columns=[c for c in ["Ordem", "Seção", "Nível", "Tipo"] if c in df.columns])
    return display



def classe_linha_dre(linha):
    texto = str(linha).upper()
    # Subtotais principais em amarelo executivo
    if "(=)" in texto or "LUCRO" in texto or "EBITDA" in texto or "LAIR" in texto:
        return "subtotal"
    # Blocos principais em azul Eirox claro
    if texto.startswith(("1.", "2.", "4.", "6.", "8.", "10.")):
        return "bloco"
    # Agrupadores internos
    if texto in [
        "DESPESAS COM PESSOAL",
        "DESPESAS ADMINISTRATIVAS E OCUPAÇÃO",
        "DESPESAS COM VENDAS E MARKETING",
    ]:
        return "grupo"
    return "normal"


def render_tabela_dre_html(tabela: pd.DataFrame, height: int = 720) -> str:
    """Renderiza a DRE em HTML puro para evitar erro de DOM/removeChild do Streamlit."""
    if tabela.empty:
        return ""

    cols = [str(c) for c in tabela.columns]
    html_rows = []
    for _, row in tabela.iterrows():
        linha = str(row.get("Linha DRE", ""))
        cls = classe_linha_dre(linha)
        cells = []
        for i, col in enumerate(cols):
            val = "" if pd.isna(row.get(col, "")) else str(row.get(col, ""))
            val = html.escape(val)
            extra = " first-col" if i == 0 else ""
            cells.append(f"<td class='{extra}'>{val}</td>")
        html_rows.append(f"<tr class='{cls}'>" + "".join(cells) + "</tr>")

    ths = []
    for i, col in enumerate(cols):
        extra = " first-col" if i == 0 else ""
        ths.append(f"<th class='{extra}'>{html.escape(col)}</th>")

    return f"""
<style>
.dre-table-wrap {{
    width: 100%;
    max-height: {height}px;
    overflow: auto;
    border: 1px solid rgba(0,175,255,.25);
    border-radius: 18px;
    background: #0E1A2B;
    box-shadow: 0 18px 45px rgba(0,0,0,.30);
}}
table.dre-table {{
    width: max-content;
    min-width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    font-size: 15px;
    color: #F8FBFF;
}}
.dre-table th {{
    position: sticky;
    top: 0;
    z-index: 4;
    background: #171C27;
    color: #B9C3D0;
    text-align: left;
    font-weight: 700;
    padding: 12px 12px;
    border-bottom: 1px solid rgba(255,255,255,.12);
    border-right: 1px solid rgba(255,255,255,.10);
    white-space: nowrap;
}}
.dre-table td {{
    padding: 11px 12px;
    border-bottom: 1px solid rgba(255,255,255,.08);
    border-right: 1px solid rgba(255,255,255,.08);
    white-space: nowrap;
    font-weight: 500;
}}
.dre-table .first-col {{
    position: sticky;
    left: 0;
    z-index: 3;
    min-width: 360px;
    max-width: 520px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}
.dre-table th.first-col {{ z-index: 6; }}
.dre-table tr.normal td {{ background: #0E1A2B; color: #F8FBFF; }}
.dre-table tr.normal td.first-col {{ background: #0E1A2B; }}
.dre-table tr.bloco td {{ background: #D9EAFF; color: #06101D; font-weight: 900; }}
.dre-table tr.bloco td.first-col {{ background: #D9EAFF; }}
.dre-table tr.subtotal td {{ background: #FFD11A; color: #06101D; font-weight: 950; }}
.dre-table tr.subtotal td.first-col {{ background: #FFD11A; }}
.dre-table tr.grupo td {{ background: #263243; color: #FFFFFF; font-weight: 900; font-style: italic; }}
.dre-table tr.grupo td.first-col {{ background: #263243; }}
.dre-table tr:hover td {{ filter: brightness(1.08); }}
</style>
<div class="dre-table-wrap">
<table class="dre-table">
<thead><tr>{''.join(ths)}</tr></thead>
<tbody>{''.join(html_rows)}</tbody>
</table>
</div>
"""

def row_style(row):
    linha = str(row.get("Linha DRE", "")).upper()
    # Subtotais principais em amarelo executivo
    if "(=)" in linha or "LUCRO" in linha or "EBITDA" in linha or "LAIR" in linha:
        return ["background-color: #FFD11A; color: #06101D; font-weight: 900;" for _ in row]
    # Blocos principais numerados em azul claro
    if linha.startswith(("1.", "2.", "4.", "6.", "8.", "10.")):
        return ["background-color: #D9EAFF; color: #06101D; font-weight: 900;" for _ in row]
    # Agrupadores de despesas em cinza premium
    if linha in [
        "DESPESAS COM PESSOAL",
        "DESPESAS ADMINISTRATIVAS E OCUPAÇÃO",
        "DESPESAS COM VENDAS E MARKETING",
    ]:
        return ["background-color: #2B3444; color: #FFFFFF; font-weight: 900; font-style: italic;" for _ in row]
    return ["background-color: #0E1A2B; color: #F8FBFF;" for _ in row]


def plot_line(df, title, y_title="Valor"):
    if df.empty:
        st.info("Sem dados para exibir.")
        return
    fig = px.line(df, x="Mês", y="Valor", color="Indicador", markers=True, title=title)
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color="#F8FBFF",
        title_font_size=20,
        yaxis_title=y_title,
        xaxis_title="Mês",
        legend_title_text="",
        margin=dict(l=20, r=20, t=55, b=20),
    )
    st.plotly_chart(fig, use_container_width=True)


# =====================================================
# CARREGAMENTO
# =====================================================
sheets, fonte = escolher_arquivo()

if not sheets:
    st.error("Nenhum arquivo de dados foi encontrado. Coloque DRE_Consolidado_Moderno.xlsx na pasta data/ ou na raiz do projeto.")
    st.stop()

dre = obter_df(sheets, "DRE")
dre_estruturada = obter_df(sheets, "DRE_ESTRUTURADA")
dados = obter_df(sheets, "DADOS_DRE")
contas = obter_df(sheets, "CONTAS_CLASSIFICADAS")
nao_class = obter_df(sheets, "NAO_CLASSIFICADOS")
vendas = obter_df(sheets, "VENDAS_BASE")
estoque = obter_df(sheets, "ESTOQUE_BASE")
resumo_loja = obter_df(sheets, "RESUMO_LOJA")
checks = obter_df(sheets, "CHECKS")

# Normalizações mínimas
if not dados.empty:
    dados["Valor"] = pd.to_numeric(dados.get("Valor", 0), errors="coerce").fillna(0)
    if "% Receita" in dados.columns:
        dados["% Receita"] = pd.to_numeric(dados["% Receita"], errors="coerce").fillna(0)
if not dre_estruturada.empty:
    if "Ordem" in dre_estruturada.columns:
        dre_estruturada = dre_estruturada.sort_values("Ordem")

meses = []
if not dre_estruturada.empty:
    meses = sorted(set(str(c).replace(" Valor", "").replace(" %", "") for c in dre_estruturada.columns if str(c).endswith(" Valor")), key=mes_key)
    meses = [m for m in meses if m.lower() not in ["sem mês", "sem mes"]]
if not meses and not dados.empty and "Mês" in dados.columns:
    meses = ordenar_meses(dados["Mês"].unique())
    meses = [m for m in meses if str(m).lower() not in ["sem mês", "sem mes"]]

# =====================================================
# SIDEBAR
# =====================================================
logo_b64 = image_to_base64(LOGO_FILE)
if logo_b64:
    st.sidebar.markdown(f"<div style='text-align:center'><img src='data:image/png;base64,{logo_b64}' style='max-width:230px; width:100%; margin-bottom:10px;'></div>", unsafe_allow_html=True)
else:
    st.sidebar.markdown("# EIROX")

st.sidebar.markdown("### Navegação")
pagina = st.sidebar.radio(
    "",
    [
        "📊 Painel Executivo",
        "📈 DRE Gerencial",
        "🏪 Resultado por Loja",
        "📦 Estoque",
        "⚠️ Auditoria DRE",
        "📋 Dados Base",
    ],
    label_visibility="collapsed",
)

st.sidebar.markdown("---")
st.sidebar.caption(f"Fonte: {fonte}")

meses_fechados = filtrar_meses_fechados(meses)

if meses_fechados:
    meses_sel = st.sidebar.multiselect(
        "Meses",
        meses_fechados,
        default=meses_fechados,
        help="Filtro automático: exibe somente meses fechados, sem considerar o mês atual."
    )
elif meses:
    meses_sel = st.sidebar.multiselect(
        "Meses",
        meses,
        default=meses,
        help="Não foi possível identificar meses fechados; exibindo todos os meses disponíveis."
    )
else:
    meses_sel = []

mes_atual = ultimo_mes_com_receita(dados)

# =====================================================
# HEADER
# =====================================================
st.markdown("<div class='eirox-hero'>", unsafe_allow_html=True)
col_logo, col_title = st.columns([1, 4])
with col_logo:
    if LOGO_FILE.exists():
        st.image(str(LOGO_FILE), use_container_width=True)
with col_title:
    st.markdown("<h1 class='eirox-title'>DRE Empresa Online</h1>", unsafe_allow_html=True)
    st.markdown("<div class='eirox-subtitle'>Dashboard financeiro gerencial no padrão Eirox Pricing Online • DRE moderno com meses em colunas, auditoria e indicadores executivos.</div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# =====================================================
# PÁGINAS
# =====================================================
if pagina == "📊 Painel Executivo":
    st.markdown("<div class='section-title'>Painel Executivo</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='section-note'>Mês de referência: <b>{mes_atual or 'não identificado'}</b>. Não classificados ficam fora dos cálculos principais e aparecem na Auditoria DRE.</div>", unsafe_allow_html=True)

    receita = valor_linha_mes(dados, "RECEITA OPERACIONAL BRUTA", mes_atual)
    lucro_bruto = valor_linha_mes(dados, "LUCRO BRUTO", mes_atual)
    ebitda = valor_linha_mes(dados, "RESULTADO ANTES DO RESULTADO FINANCEIRO", mes_atual)
    lucro_liquido = valor_linha_mes(dados, "LUCRO LÍQUIDO", mes_atual)
    margem_ebitda = percentual_linha_mes(dados, "RESULTADO ANTES DO RESULTADO FINANCEIRO", mes_atual)
    margem_liquida = percentual_linha_mes(dados, "LUCRO LÍQUIDO", mes_atual)

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    kpis = [
        (c1, "Receita Bruta", fmt_moeda(receita), "Base de vendas"),
        (c2, "Lucro Bruto", fmt_moeda(lucro_bruto), "Após CMV"),
        (c3, "EBITDA", fmt_moeda(ebitda), fmt_pct(margem_ebitda)),
        (c4, "Lucro Líquido", fmt_moeda(lucro_liquido), fmt_pct(margem_liquida)),
        (c5, "Margem EBITDA", fmt_pct(margem_ebitda), "Sobre receita"),
        (c6, "Margem Líquida", fmt_pct(margem_liquida), "Resultado final"),
    ]
    for col, title, value, help_txt in kpis:
        with col:
            st.markdown(f"<div class='kpi-card'><div class='kpi-title'>{title}</div><div class='kpi-value'>{value}</div><div class='kpi-help'>{help_txt}</div></div>", unsafe_allow_html=True)

    st.markdown("<div class='section-title'>Evolução dos principais indicadores</div>", unsafe_allow_html=True)
    if not dados.empty and meses_sel:
        linhas_ind = [
            ("Receita", "RECEITA OPERACIONAL BRUTA"),
            ("Lucro Bruto", "LUCRO BRUTO"),
            ("EBITDA", "RESULTADO ANTES DO RESULTADO FINANCEIRO"),
            ("Lucro Líquido", "LUCRO LÍQUIDO"),
        ]
        rows = []
        for indicador, chave in linhas_ind:
            for m in meses_sel:
                rows.append({"Indicador": indicador, "Mês": m, "Valor": valor_linha_mes(dados, chave, m)})
        evol = pd.DataFrame(rows)
        plot_line(evol, "Receita, Lucro Bruto, EBITDA e Lucro Líquido")

elif pagina == "📈 DRE Gerencial":
    st.markdown("<div class='section-title'>DRE Gerencial</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-note'>Modelo mantido no formato da aba DRE original: meses em colunas, valores e percentuais lado a lado. Todas as linhas estruturais foram destacadas.</div>", unsafe_allow_html=True)

    tabela = montar_dre_para_tabela(dre_estruturada, meses_sel)
    if tabela.empty:
        st.warning("A aba DRE_ESTRUTURADA não foi encontrada ou está vazia.")
    else:
        # Tabela HTML estável: evita o erro de navegador/Streamlit "removeChild"
        st.markdown(render_tabela_dre_html(tabela, height=720), unsafe_allow_html=True)

elif pagina == "🏪 Resultado por Loja":
    st.markdown("<div class='section-title'>Resultado por Loja</div>", unsafe_allow_html=True)
    if resumo_loja.empty:
        st.info("A aba RESUMO_LOJA não foi encontrada.")
    else:
        df = resumo_loja.copy()
        if "Mês" in df.columns and meses_sel:
            df = df[df["Mês"].astype(str).isin(meses_sel)]
        for c in ["Receita", "Despesas", "Resultado Caixa Simplificado"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        st.dataframe(df, use_container_width=True, height=350)
        if {"Loja", "Resultado Caixa Simplificado"}.issubset(df.columns):
            g = df.groupby("Loja", as_index=False)["Resultado Caixa Simplificado"].sum()
            fig = px.bar(g, x="Loja", y="Resultado Caixa Simplificado", title="Resultado Caixa Simplificado por Loja")
            fig.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#F8FBFF")
            st.plotly_chart(fig, use_container_width=True)

elif pagina == "📦 Estoque":
    st.markdown("<div class='section-title'>Estoque x Receita</div>", unsafe_allow_html=True)
    if estoque.empty:
        st.info("A aba ESTOQUE_BASE não foi encontrada.")
    else:
        df = estoque.copy()
        if "Mês" in df.columns and meses_sel:
            df = df[df["Mês"].astype(str).isin(meses_sel)]
        if "Valor Estoque" in df.columns:
            df["Valor Estoque"] = pd.to_numeric(df["Valor Estoque"], errors="coerce").fillna(0)
        st.dataframe(df, use_container_width=True, height=350)
        if {"Mês", "Valor Estoque"}.issubset(df.columns):
            g = df.groupby("Mês", as_index=False)["Valor Estoque"].sum()
            g["Mês"] = pd.Categorical(g["Mês"], categories=meses, ordered=True)
            g = g.sort_values("Mês")
            fig = px.line(g, x="Mês", y="Valor Estoque", markers=True, title="Evolução do Estoque a Custo")
            fig.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#F8FBFF")
            st.plotly_chart(fig, use_container_width=True)

elif pagina == "⚠️ Auditoria DRE":
    st.markdown("<div class='section-title'>Auditoria DRE</div>", unsafe_allow_html=True)
    valor_nao = 0
    qtde_nao = 0
    if not nao_class.empty:
        if "Valor" in nao_class.columns:
            valor_nao = pd.to_numeric(nao_class["Valor"], errors="coerce").fillna(0).sum()
        if "Qtde" in nao_class.columns:
            qtde_nao = pd.to_numeric(nao_class["Qtde"], errors="coerce").fillna(0).sum()
        else:
            qtde_nao = len(nao_class)
    receita_total = dados[dados["Linha DRE"].astype(str).str.contains("RECEITA OPERACIONAL BRUTA", case=False, na=False)]["Valor"].sum() if not dados.empty else 0
    impacto = valor_nao / receita_total if receita_total else 0
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Não classificados</div><div class='kpi-value'>{int(qtde_nao)}</div><div class='kpi-help'>Registros/agrupamentos</div></div>", unsafe_allow_html=True)
    with c2:
        st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Valor não classificado</div><div class='kpi-value'>{fmt_moeda(valor_nao)}</div><div class='kpi-help'>Fora do DRE principal</div></div>", unsafe_allow_html=True)
    with c3:
        st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Impacto estimado</div><div class='kpi-value'>{fmt_pct(impacto)}</div><div class='kpi-help'>Sobre receita total</div></div>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>Itens para classificação</div>", unsafe_allow_html=True)
    st.dataframe(nao_class, use_container_width=True, height=520)

else:
    st.markdown("<div class='section-title'>Dados Base</div>", unsafe_allow_html=True)
    aba = st.selectbox("Selecione a base", list(sheets.keys()))
    st.dataframe(sheets[aba], use_container_width=True, height=650)
