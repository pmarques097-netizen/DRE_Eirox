# -*- coding: utf-8 -*-
"""
Eirox DRE Empresa Online - versão profissional
- Lê automaticamente o arquivo data/DRE_Consolidado_Moderno.xlsx ou DRE_Consolidado_Moderno.xlsx
- Mantém o formato da aba DRE: meses em colunas e Valor/% lado a lado
- Filtra por padrão somente meses fechados do ano atual
- Destaques visuais para todos os tópicos do DRE e resultados estratégicos
- Formatação em moeda brasileira e percentual brasileiro
"""

from __future__ import annotations

import base64
import re
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

# =========================
# CONFIGURAÇÃO GERAL
# =========================
st.set_page_config(
    page_title="Eirox DRE Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

PASTA_APP = Path(__file__).resolve().parent
CAMINHOS_BASE = [
    PASTA_APP / "data" / "DRE_Consolidado_Moderno.xlsx",
    PASTA_APP / "DRE_Consolidado_Moderno.xlsx",
    Path.cwd() / "data" / "DRE_Consolidado_Moderno.xlsx",
    Path.cwd() / "DRE_Consolidado_Moderno.xlsx",
    Path.home() / "Desktop" / "Dre" / "data" / "DRE_Consolidado_Moderno.xlsx",
    Path.home() / "Desktop" / "Dre" / "DRE_Consolidado_Moderno.xlsx",
]
LOGO_PATHS = [
    PASTA_APP / "assets" / "logo_eirox.png",
    PASTA_APP / "logo_eirox.png",
    Path.cwd() / "assets" / "logo_eirox.png",
    Path.cwd() / "logo_eirox.png",
]

MESES_PT = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"
}
MESES_INV = {v: k for k, v in MESES_PT.items()}

# =========================
# CSS PROFISSIONAL
# =========================
def inject_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
        .stApp { background: radial-gradient(circle at top, #07192d 0, #020914 42%, #020713 100%); color: #f8fafc; }
        section[data-testid="stSidebar"] { background: #020811; border-right: 1px solid #123c64; }
        .main .block-container { padding-top: 2rem; max-width: 1480px; }
        .hero { text-align:center; padding: 10px 0 18px 0; }
        .hero img { max-width: 145px; margin-bottom: 8px; }
        .hero h1 { font-size: 3.2rem; line-height: 1.05; margin: 0; font-weight: 800; color: #f8fafc; }
        .hero p { font-size: 1.05rem; color: #52b8ff; margin-top: 12px; }
        .kpi-grid { display: grid; grid-template-columns: repeat(4, minmax(180px, 1fr)); gap: 14px; margin: 18px 0 24px 0; }
        .kpi-card { background: linear-gradient(180deg, #0f2035 0%, #081426 100%); border: 1px solid #1d4f79; border-radius: 18px; padding: 18px; box-shadow: 0 10px 28px rgba(0,0,0,.28); }
        .kpi-title { color: #9fb3c8; font-size: .84rem; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; }
        .kpi-value { font-size: 1.55rem; font-weight: 800; color: #ffffff; margin-top: 8px; }
        .kpi-sub { color: #52b8ff; font-size: .85rem; margin-top: 4px; }
        .section-title { font-weight: 800; font-size: 1.55rem; margin: 28px 0 10px; color: #fff; }
        .alert-ok { background:#0c3f2a; color:#5cff9d; padding:13px 16px; border-radius:12px; border:1px solid #238a55; font-weight:700; }
        .alert-err { background:#3a1b26; color:#ff6673; padding:13px 16px; border-radius:12px; border:1px solid #773344; font-weight:700; }
        .path-box { background:#09243a; color:#9bd3ff; padding:12px; border-radius:12px; font-size:.82rem; word-break:break-all; }
        .dre-wrap { border:1px solid #164b78; border-radius: 16px; overflow:auto; max-height: 760px; background:#06111f; box-shadow:0 10px 30px rgba(0,0,0,.25); }
        table.dre-table { border-collapse: collapse; width: 100%; min-width: 1120px; font-size: 14px; }
        table.dre-table th { position: sticky; top: 0; z-index: 3; background:#171b26; color:#dbe8f7; padding:12px 10px; border:1px solid #253850; text-align:right; white-space:nowrap; }
        table.dre-table th:first-child { left:0; z-index:4; text-align:left; min-width:410px; }
        table.dre-table td { padding:11px 10px; border:1px solid #20324a; text-align:right; white-space:nowrap; color:#f8fafc; }
        table.dre-table td:first-child { position: sticky; left:0; z-index:2; text-align:left; min-width:410px; background:inherit; }
        tr.linha-normal { background:#071426; }
        tr.linha-bloco { background:#cfe1f3; color:#001021; font-weight:800; }
        tr.linha-bloco td { color:#001021; font-weight:800; }
        tr.linha-resultado { background:#ffd21f; color:#001021; font-weight:900; }
        tr.linha-resultado td { color:#001021; font-weight:900; }
        tr.linha-grupo { background:#202938; color:#f7fbff; font-weight:800; font-style:italic; }
        tr.linha-grupo td { color:#f7fbff; font-weight:800; }
        tr.linha-estrategica { background:#ffca0a; color:#001021; font-weight:900; }
        tr.linha-estrategica td { color:#001021; font-weight:900; }
        .small-note { color:#9aa7b5; font-size:.86rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )

# =========================
# HELPERS DE FORMATAÇÃO
# =========================
def localizar_arquivo(candidatos: List[Path]) -> Optional[Path]:
    for p in candidatos:
        if p.exists():
            return p
    return None


def img_to_base64(path: Path) -> str:
    try:
        return base64.b64encode(path.read_bytes()).decode("utf-8")
    except Exception:
        return ""


def norm(txt) -> str:
    if pd.isna(txt):
        return ""
    s = str(txt).strip().upper()
    repl = {"Á":"A","À":"A","Â":"A","Ã":"A","É":"E","Ê":"E","Í":"I","Ó":"O","Ô":"O","Õ":"O","Ú":"U","Ç":"C"}
    for a,b in repl.items():
        s = s.replace(a,b)
    s = re.sub(r"\s+", " ", s)
    return s


def br_money(v) -> str:
    try:
        if pd.isna(v) or v == "":
            return "R$ 0,00"
        v = float(v)
        txt = f"R$ {v:,.2f}"
        return txt.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def br_percent(v) -> str:
    try:
        if pd.isna(v) or v == "":
            return "0,00%"
        v = float(v)
        # se veio como 41,69 em vez de 0,4169, mantém como 41,69%
        if abs(v) <= 1.5:
            v *= 100
        return f"{v:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00%"


def parse_num(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float, np.number)):
        return float(v)
    s = str(v).strip()
    if s in ["", "-", "nan", "None"]:
        return 0.0
    s = s.replace("R$", "").replace("%", "").replace(" ", "")
    # padrão brasileiro
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def mes_key(mes: str) -> Tuple[int, int]:
    m = re.search(r"([a-zç]{3})/(\d{2})", str(mes).lower())
    if not m:
        return (9999, 99)
    mm = MESES_INV.get(m.group(1), 99)
    yy = 2000 + int(m.group(2))
    return (yy, mm)


def meses_fechados_ano_atual(meses: List[str]) -> List[str]:
    hoje = datetime.today()
    ano_atual = hoje.year
    mes_atual = hoje.month
    fechados = []
    for m in meses:
        ano, mes = mes_key(m)
        if ano == ano_atual and mes < mes_atual:
            fechados.append(m)
    return fechados or meses

# =========================
# LEITURA DO EXCEL DRE
# =========================
def escolher_aba(path: Path) -> str:
    xls = pd.ExcelFile(path)
    nomes = xls.sheet_names
    for preferida in ["DRE", "DRE_ESTRUTURADA", "DRE GERENCIAL", "DRE_ESTRUTURADA_MODERNA"]:
        if preferida in nomes:
            return preferida
    # procura aba que contenha Linha DRE
    for aba in nomes:
        try:
            tmp = pd.read_excel(path, sheet_name=aba, header=None, nrows=15)
            if tmp.astype(str).apply(lambda col: col.str.contains("Linha DRE", case=False, na=False)).any().any():
                return aba
        except Exception:
            pass
    return nomes[0]


@st.cache_data(show_spinner=False)
def carregar_dre(path_str: str) -> Tuple[pd.DataFrame, List[str], str]:
    path = Path(path_str)
    aba = escolher_aba(path)

    raw = pd.read_excel(path, sheet_name=aba, header=None, dtype=object)
    raw = raw.dropna(how="all").reset_index(drop=True)

    # Caso 1: arquivo já com cabeçalho plano: Linha DRE, jan/26 Valor, jan/26 %, ...
    first_rows_text = raw.head(8).astype(str).to_string()
    header_row = None
    line_col = None

    for i in range(min(20, len(raw))):
        for j in range(raw.shape[1]):
            if "LINHA DRE" in norm(raw.iat[i, j]):
                header_row = i
                line_col = j
                break
        if header_row is not None:
            break

    if header_row is None:
        # tenta ler como dataframe comum
        df_try = pd.read_excel(path, sheet_name=aba)
        df_try.columns = [str(c).strip() for c in df_try.columns]
        linha_col = None
        for c in df_try.columns:
            if "LINHA DRE" in norm(c) or norm(c) in ["DRE", "CONTA", "LINHA"]:
                linha_col = c
                break
        if linha_col is None:
            raise ValueError(f"Não encontrei a coluna Linha DRE na aba {aba}.")
        dre = df_try.rename(columns={linha_col: "Linha DRE"}).copy()
        meses = sorted({c.replace(" Valor", "").replace(" %", "") for c in dre.columns if re.search(r"[a-zç]{3}/\d{2}", str(c).lower())}, key=mes_key)
        return dre, meses, aba

    # Se linha anterior tem meses e a linha atual tem Valor/%
    mes_row = header_row - 1 if header_row > 0 else header_row

    col_map: Dict[int, str] = {line_col: "Linha DRE"}
    current_month = ""

    for col in range(line_col + 1, raw.shape[1]):
        maybe_month = raw.iat[mes_row, col] if mes_row >= 0 else ""
        if pd.notna(maybe_month) and re.search(r"[a-zç]{3}/\d{2}", str(maybe_month).lower()):
            current_month = re.search(r"[a-zç]{3}/\d{2}", str(maybe_month).lower()).group(0)

        sub = norm(raw.iat[header_row, col])
        if sub in ["VALOR", "VALORES"]:
            col_map[col] = f"{current_month} Valor"
        elif sub in ["%", "PERCENTUAL"]:
            col_map[col] = f"{current_month} %"
        else:
            # Caso cabeçalho já esteja em uma célula só: jan/26 Valor
            text = str(raw.iat[header_row, col])
            m = re.search(r"([a-zç]{3}/\d{2}).*(valor|%)", text.lower())
            if m:
                col_map[col] = f"{m.group(1)} {'%' if '%' in text else 'Valor'}"

    data = raw.iloc[header_row + 1:].copy()
    data = data[[c for c in col_map.keys() if c < data.shape[1]]]
    data.columns = [col_map[c] for c in col_map.keys() if c < raw.shape[1]]
    data = data.dropna(how="all")
    data["Linha DRE"] = data["Linha DRE"].astype(str).str.strip()
    data = data[(data["Linha DRE"] != "") & (data["Linha DRE"].str.lower() != "nan")]

    # remove linhas de título/rodapé que não são DRE
    data = data[~data["Linha DRE"].str.contains("Mês Referência|Realizado|Atualizado automaticamente", case=False, na=False)]

    meses = []
    for c in data.columns:
        m = re.match(r"([a-zç]{3}/\d{2})\s+(Valor|%)", str(c).lower())
        if m and m.group(1) not in meses:
            meses.append(m.group(1))
    meses = sorted(meses, key=mes_key)

    # força numérico nas colunas
    for c in data.columns:
        if c != "Linha DRE":
            data[c] = data[c].apply(parse_num)

    return data.reset_index(drop=True), meses, aba

# =========================
# CLASSIFICAÇÃO VISUAL
# =========================
def classe_linha(linha: str) -> str:
    l = norm(linha)
    if any(x in l for x in [
        "RECEITA OPERACIONAL LIQUIDA",
        "LUCRO BRUTO",
        "RESULTADO ANTES DO RESULTADO FINANCEIRO",
        "EBITDA",
        "LAJIDA",
        "RESULTADO ANTES DOS TRIBUTOS",
        "LAIR",
        "LUCRO LIQUIDO",
        "MARGEM DE CONTRIBUICAO TOTAL",
        "MARGEM DE CONTRIBUICAO PERCENTUAL",
        "PONTO DE EQUILIBRIO",
        "POSICAO FINAL",
    ]):
        return "linha-resultado"
    if re.match(r"^\d+\.", l) or any(x in l for x in [
        "RECEITA OPERACIONAL BRUTA", "DEDUCOES DA RECEITA BRUTA", "CUSTOS DAS VENDAS",
        "DESPESAS OPERACIONAIS", "RESULTADO FINANCEIRO", "TRIBUTOS SOBRE O LUCRO",
        "CONCILIACAO DE FLUXO", "ANALISE DE PONTO DE EQUILIBRIO", "CALCULOS INTERMEDIARIOS"
    ]):
        return "linha-bloco"
    if any(x in l for x in [
        "DESPESAS COM PESSOAL", "DESPESAS ADMINISTRATIVAS", "DESPESAS COM VENDAS",
        "MARKETING", "INDICADORES", "RECEITAS", "ESTOQUE A CUSTO", "CONTAS A PAGAR FORNECEDOR"
    ]):
        return "linha-grupo"
    return "linha-normal"


def formatar_valor_por_linha(linha: str, coluna: str, valor) -> str:
    l = norm(linha)
    if coluna.endswith(" %"):
        return br_percent(valor)
    if "PERCENTUAL" in l and coluna.endswith(" Valor"):
        return br_percent(valor)
    return br_money(valor)


def montar_tabela_html(df: pd.DataFrame, meses: List[str]) -> str:
    cols = ["Linha DRE"]
    for m in meses:
        if f"{m} Valor" in df.columns:
            cols.append(f"{m} Valor")
        if f"{m} %" in df.columns:
            cols.append(f"{m} %")
    show = df[cols].copy()

    html = '<div class="dre-wrap"><table class="dre-table"><thead><tr>'
    for c in show.columns:
        html += f"<th>{c}</th>"
    html += "</tr></thead><tbody>"
    for _, row in show.iterrows():
        linha = row["Linha DRE"]
        cls = classe_linha(linha)
        html += f'<tr class="{cls}">'
        for c in show.columns:
            if c == "Linha DRE":
                txt = str(row[c])
            else:
                txt = formatar_valor_por_linha(linha, c, row[c])
            html += f"<td>{txt}</td>"
        html += "</tr>"
    html += "</tbody></table></div>"
    return html

# =========================
# INDICADORES
# =========================
def buscar_linha(df: pd.DataFrame, termos: List[str]) -> Optional[pd.Series]:
    for termo in termos:
        mask = df["Linha DRE"].apply(lambda x: termo in norm(x))
        if mask.any():
            return df.loc[mask].iloc[0]
    return None


def valor_linha(df: pd.DataFrame, termos: List[str], mes: str, tipo="Valor") -> float:
    row = buscar_linha(df, termos)
    col = f"{mes} {tipo}"
    if row is None or col not in df.columns:
        return 0.0
    return parse_num(row[col])


def card(title: str, value: str, sub: str = ""):
    return f"""
    <div class="kpi-card">
      <div class="kpi-title">{title}</div>
      <div class="kpi-value">{value}</div>
      <div class="kpi-sub">{sub}</div>
    </div>
    """


def painel_kpis(df: pd.DataFrame, mes: str):
    receita = valor_linha(df, ["RECEITA OPERACIONAL BRUTA"], mes)
    rec_liq = valor_linha(df, ["RECEITA OPERACIONAL LIQUIDA"], mes)
    lucro_bruto = valor_linha(df, ["LUCRO BRUTO"], mes)
    ebitda = valor_linha(df, ["RESULTADO ANTES DO RESULTADO FINANCEIRO", "EBITDA", "LAJIDA"], mes)
    lucro_liq = valor_linha(df, ["LUCRO LIQUIDO"], mes)
    cmv_pct = valor_linha(df, ["CUSTOS DAS VENDAS", "MERCADORIAS (CMP)", "CUSTO MEDIO DE VENDA"], mes, "%")

    margem_ebitda = ebitda / receita if receita else 0
    margem_liq = lucro_liq / receita if receita else 0

    html = '<div class="kpi-grid">'
    html += card("Receita Bruta", br_money(receita), mes)
    html += card("Receita Líquida", br_money(rec_liq), mes)
    html += card("Lucro Bruto", br_money(lucro_bruto), br_percent(lucro_bruto/receita if receita else 0))
    html += card("EBITDA", br_money(ebitda), br_percent(margem_ebitda))
    html += card("Lucro Líquido", br_money(lucro_liq), br_percent(margem_liq))
    html += card("CMV %", br_percent(cmv_pct), "sobre receita")
    html += card("Margem EBITDA", br_percent(margem_ebitda), "resultado operacional")
    html += card("Margem Líquida", br_percent(margem_liq), "resultado final")
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def grafico_evolucao(df: pd.DataFrame, meses: List[str]):
    linhas = {
        "Receita Bruta": ["RECEITA OPERACIONAL BRUTA"],
        "Lucro Bruto": ["LUCRO BRUTO"],
        "EBITDA": ["RESULTADO ANTES DO RESULTADO FINANCEIRO", "EBITDA"],
        "Lucro Líquido": ["LUCRO LIQUIDO"],
    }
    fig = go.Figure()
    for nome, termos in linhas.items():
        vals = [valor_linha(df, termos, m) for m in meses]
        fig.add_trace(go.Scatter(x=meses, y=vals, mode="lines+markers", name=nome))
    fig.update_layout(
        height=390,
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=35, b=20),
        legend=dict(orientation="h", y=1.12),
        yaxis_tickprefix="R$ ",
    )
    st.plotly_chart(fig, use_container_width=True)


def resultados_estrategicos(df: pd.DataFrame, mes: str):
    receita = valor_linha(df, ["RECEITA TOTAL", "RECEITA OPERACIONAL BRUTA"], mes)
    margem_contrib = valor_linha(df, ["MARGEM DE CONTRIBUICAO TOTAL"], mes)
    margem_pct = valor_linha(df, ["MARGEM DE CONTRIBUICAO PERCENTUAL"], mes, "Valor")
    ponto_eq = valor_linha(df, ["PONTO DE EQUILIBRIO"], mes)
    cmv = valor_linha(df, ["CUSTO MEDIO DE VENDA", "CUSTOS DAS VENDAS"], mes)
    folga = receita - ponto_eq if ponto_eq else 0

    html = '<div class="kpi-grid">'
    html += card("Margem de Contribuição", br_money(margem_contrib), br_percent(margem_pct))
    html += card("Ponto de Equilíbrio", br_money(ponto_eq), "receita mínima")
    html += card("Folga Operacional", br_money(folga), "receita - ponto equilíbrio")
    html += card("CMV", br_money(cmv), br_percent(cmv/receita if receita else 0))
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

# =========================
# APP
# =========================
def main():
    inject_css()

    if st.sidebar.button("🔄 Atualizar Base", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📁 Base monitorada")

    arquivo_base = localizar_arquivo(CAMINHOS_BASE)
    logo = localizar_arquivo(LOGO_PATHS)

    if arquivo_base:
        st.sidebar.markdown('<div class="alert-ok">DRE_Consolidado_Moderno.xlsx</div>', unsafe_allow_html=True)
        st.sidebar.markdown(f'<div class="path-box">{arquivo_base}</div>', unsafe_allow_html=True)
    else:
        st.sidebar.markdown('<div class="alert-err">Base não encontrada</div>', unsafe_allow_html=True)
        st.sidebar.write("Coloque o arquivo em `data/DRE_Consolidado_Moderno.xlsx` ou na raiz do projeto.")

    logo_html = ""
    if logo:
        b64 = img_to_base64(logo)
        logo_html = f'<img src="data:image/png;base64,{b64}" />'

    st.markdown(
        f"""
        <div class="hero">
            {logo_html}
            <h1>DRE Empresa Online</h1>
            <p>Dashboard financeiro gerencial • DRE no formato aprovado • Meses fechados do ano atual</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not arquivo_base:
        st.error("Nenhuma base foi encontrada. Verifique o arquivo DRE_Consolidado_Moderno.xlsx.")
        return

    try:
        df, meses, aba = carregar_dre(str(arquivo_base))
    except Exception as e:
        st.error(f"Erro ao ler a base: {e}")
        return

    if not meses:
        st.error(f"Nenhum mês foi identificado na aba {aba}. Verifique se a coluna Linha DRE e os meses estão no formato jan/26 Valor | jan/26 %. ")
        return

    meses_padrao = meses_fechados_ano_atual(meses)
    selected_months = st.sidebar.multiselect(
        "Meses",
        options=meses,
        default=meses_padrao,
    )
    if not selected_months:
        selected_months = meses_padrao

    ultimo = selected_months[-1]
    st.sidebar.markdown("---")
    st.sidebar.caption(f"Aba lida: {aba}")
    st.sidebar.caption(f"Meses encontrados: {', '.join(meses)}")

    tab1, tab2, tab3 = st.tabs(["📊 Painel Executivo", "📋 DRE Gerencial", "🎯 Resultados Estratégicos"])

    with tab1:
        st.markdown('<div class="section-title">Painel Executivo</div>', unsafe_allow_html=True)
        painel_kpis(df, ultimo)
        st.markdown('<div class="section-title">Evolução Mensal</div>', unsafe_allow_html=True)
        grafico_evolucao(df, selected_months)

    with tab2:
        st.markdown('<div class="section-title">DRE Gerencial</div>', unsafe_allow_html=True)
        st.markdown('<div class="small-note">Modelo mantido no formato aprovado: meses em colunas, valores e percentuais lado a lado, com todos os blocos estruturais destacados.</div>', unsafe_allow_html=True)
        st.markdown(montar_tabela_html(df, selected_months), unsafe_allow_html=True)

    with tab3:
        st.markdown('<div class="section-title">Resultados Estratégicos</div>', unsafe_allow_html=True)
        resultados_estrategicos(df, ultimo)
        linhas_resultado = df[df["Linha DRE"].apply(lambda x: classe_linha(x) in ["linha-resultado", "linha-bloco", "linha-grupo"])]
        st.markdown(montar_tabela_html(linhas_resultado, selected_months), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
