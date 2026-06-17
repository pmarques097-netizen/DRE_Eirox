# -*- coding: utf-8 -*-
"""
Eirox DRE Online - Versão Final Profissional
Lê diretamente a aba DADOS_DRE do DRE_Consolidado_Moderno.xlsx.
Mantém o modelo aprovado do DRE, com meses em colunas, destaques por tipo,
moeda brasileira e indicadores executivos.
"""

from pathlib import Path
from datetime import datetime
import base64
import re
import pandas as pd
import streamlit as st

try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_OK = True
except Exception:
    PLOTLY_OK = False

# =========================================================
# CONFIGURAÇÃO
# =========================================================
st.set_page_config(
    page_title="Eirox DRE Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

MESES_PT = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}
MESES_PT_REV = {v: k for k, v in MESES_PT.items()}

# =========================================================
# FUNÇÕES UTILITÁRIAS
# =========================================================

def moeda_br(valor) -> str:
    """Formata número como moeda brasileira."""
    try:
        v = float(valor)
    except Exception:
        v = 0.0
    s = f"R$ {v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct_br(valor) -> str:
    """Formata percentual. Aceita 0.42 ou 42."""
    try:
        v = float(valor)
    except Exception:
        v = 0.0
    # Na base, normalmente vem como 0.42 = 42%.
    if abs(v) <= 2:
        v = v * 100
    s = f"{v:,.2f}%"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def numero(valor) -> float:
    try:
        if pd.isna(valor):
            return 0.0
        return float(valor)
    except Exception:
        return 0.0


def mes_key(mes: str):
    mes = str(mes).strip().lower()
    m = re.match(r"^([a-zç]{3})/(\d{2})$", mes)
    if not m:
        return (9999, 99)
    nome, ano = m.groups()
    return (2000 + int(ano), MESES_PT.get(nome[:3], 99))


def eh_mes_valido(mes: str) -> bool:
    return bool(re.match(r"^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/\d{2}$", str(mes).strip().lower()))


def meses_fechados_ano_atual(meses):
    """Retorna somente meses do ano atual anteriores ao mês atual."""
    hoje = datetime.today()
    ano_atual = hoje.year
    mes_atual = hoje.month
    fechados = []
    for m in meses:
        ano, mes_num = mes_key(m)
        if ano == ano_atual and mes_num < mes_atual:
            fechados.append(m)
    return fechados


def procurar_arquivo_base() -> Path | None:
    """Procura o DRE_Consolidado_Moderno.xlsx em locais compatíveis local/online."""
    base = Path(__file__).resolve().parent
    candidatos = [
        base / "data" / "DRE_Consolidado_Moderno.xlsx",
        base / "DRE_Consolidado_Moderno.xlsx",
        Path.cwd() / "data" / "DRE_Consolidado_Moderno.xlsx",
        Path.cwd() / "DRE_Consolidado_Moderno.xlsx",
        Path.home() / "Desktop" / "Dre" / "data" / "DRE_Consolidado_Moderno.xlsx",
        Path.home() / "Desktop" / "Dre" / "DRE_Consolidado_Moderno.xlsx",
        Path(r"C:\Users\Comercial\Desktop\Dre\data\DRE_Consolidado_Moderno.xlsx"),
        Path(r"C:\Users\Comercial\Desktop\Dre\DRE_Consolidado_Moderno.xlsx"),
    ]
    for c in candidatos:
        if c.exists():
            return c
    return None


def logo_base64() -> str | None:
    base = Path(__file__).resolve().parent
    candidatos = [
        base / "assets" / "logo_eirox.png",
        base / "logo_eirox.png",
        Path.cwd() / "assets" / "logo_eirox.png",
        Path.cwd() / "logo_eirox.png",
    ]
    for p in candidatos:
        if p.exists():
            try:
                return base64.b64encode(p.read_bytes()).decode("utf-8")
            except Exception:
                pass
    return None


@st.cache_data(show_spinner=False, ttl=0)
def carregar_dados(caminho: str):
    xl = pd.ExcelFile(caminho)
    abas = xl.sheet_names

    if "DADOS_DRE" not in abas:
        raise ValueError(f"A aba DADOS_DRE não foi encontrada. Abas existentes: {abas}")

    dados = pd.read_excel(caminho, sheet_name="DADOS_DRE")
    dados.columns = [str(c).strip() for c in dados.columns]

    obrig = ["Ordem", "Seção", "Linha DRE", "Nível", "Tipo", "Mês", "Valor", "% Receita"]
    faltantes = [c for c in obrig if c not in dados.columns]
    if faltantes:
        raise ValueError(f"Colunas ausentes na aba DADOS_DRE: {faltantes}")

    dados = dados.copy()
    dados["Mês"] = dados["Mês"].astype(str).str.strip().str.lower()
    dados["Linha DRE"] = dados["Linha DRE"].astype(str).str.strip()
    dados["Tipo"] = dados["Tipo"].astype(str).str.strip()
    dados["Ordem"] = pd.to_numeric(dados["Ordem"], errors="coerce").fillna(999999)
    dados["Nível"] = pd.to_numeric(dados["Nível"], errors="coerce").fillna(0).astype(int)
    dados["Valor"] = pd.to_numeric(dados["Valor"], errors="coerce").fillna(0.0)
    dados["% Receita"] = pd.to_numeric(dados["% Receita"], errors="coerce").fillna(0.0)

    # Remove meses técnicos que não devem aparecer no painel principal.
    dados = dados[dados["Mês"].apply(eh_mes_valido)].copy()

    resumo_loja = pd.DataFrame()
    if "RESUMO_LOJA" in abas:
        resumo_loja = pd.read_excel(caminho, sheet_name="RESUMO_LOJA")
        resumo_loja.columns = [str(c).strip() for c in resumo_loja.columns]
        if "Mês" in resumo_loja.columns:
            resumo_loja["Mês"] = resumo_loja["Mês"].astype(str).str.strip().str.lower()
            resumo_loja = resumo_loja[resumo_loja["Mês"].apply(eh_mes_valido)].copy()

    nao_class = pd.DataFrame()
    if "NAO_CLASSIFICADOS" in abas:
        nao_class = pd.read_excel(caminho, sheet_name="NAO_CLASSIFICADOS")
        nao_class.columns = [str(c).strip() for c in nao_class.columns]
        if "Mês" in nao_class.columns:
            nao_class["Mês"] = nao_class["Mês"].astype(str).str.strip().str.lower()

    checks = pd.DataFrame()
    if "CHECKS" in abas:
        checks = pd.read_excel(caminho, sheet_name="CHECKS")
        checks.columns = [str(c).strip() for c in checks.columns]

    return dados, resumo_loja, nao_class, checks


def valor_linha(dados, linha_busca, mes, contains=True):
    base = dados[dados["Mês"].eq(mes)]
    if contains:
        base = base[base["Linha DRE"].str.upper().str.contains(linha_busca.upper(), na=False, regex=False)]
    else:
        base = base[base["Linha DRE"].str.upper().eq(linha_busca.upper())]
    if base.empty:
        return 0.0
    return numero(base.sort_values("Ordem").iloc[0]["Valor"])


def pct_linha(dados, linha_busca, mes, contains=True):
    base = dados[dados["Mês"].eq(mes)]
    if contains:
        base = base[base["Linha DRE"].str.upper().str.contains(linha_busca.upper(), na=False, regex=False)]
    else:
        base = base[base["Linha DRE"].str.upper().eq(linha_busca.upper())]
    if base.empty:
        return 0.0
    return numero(base.sort_values("Ordem").iloc[0]["% Receita"])


def construir_tabela_html(dados, meses):
    """Cria tabela HTML estável com destaques, sem usar st.dataframe Styler."""
    base = dados[dados["Mês"].isin(meses)].copy()
    if base.empty:
        return "<div class='alerta'>Nenhum dado encontrado para os meses selecionados.</div>"

    ordem_linhas = (
        base[["Ordem", "Seção", "Linha DRE", "Nível", "Tipo"]]
        .drop_duplicates()
        .sort_values("Ordem")
    )

    def classe_linha(linha, tipo, secao):
        linha_up = str(linha).upper()
        tipo = str(tipo).lower()
        secao = str(secao).upper()

        resultados_estrategicos = [
            "RECEITA TOTAL",
            "MARGEM DE CONTRIBUIÇÃO TOTAL",
            "MARGEM DE CONTRIBUIÇÃO PERCENTUAL",
            "PONTO DE EQUILÍBRIO",
        ]
        blocos_estrategicos = [
            "DESPESAS FIXAS",
            "DESPESAS VARIÁVEIS",
            "CUSTOS E DESPESAS FIXAS TOTAIS",
            "CUSTO MÉDIO DE VENDA",
            "TOTAL DE CUSTOS VARIÁVEIS",
        ]

        if any(x in linha_up for x in resultados_estrategicos):
            return "linha-amarela"
        if any(x in linha_up for x in blocos_estrategicos):
            return "linha-azul"
        if "resultado_amarelo" in tipo:
            return "linha-amarela"
        if "subtotal_azul" in tipo:
            return "linha-azul"
        if "grupo_italico" in tipo:
            return "linha-grupo"
        return "linha-detalhe"

    html = []
    html.append("<div class='dre-scroll'><table class='dre-table'>")
    html.append("<thead><tr><th class='col-linha'>Linha DRE</th>")
    for mes in meses:
        html.append(f"<th>{mes} Valor</th><th>{mes} %</th>")
    html.append("</tr></thead><tbody>")

    for _, row in ordem_linhas.iterrows():
        linha = row["Linha DRE"]
        tipo = row["Tipo"]
        secao = row["Seção"]
        nivel = int(row["Nível"])
        classe = classe_linha(linha, tipo, secao)
        padding = 12 + (nivel * 18)
        html.append(f"<tr class='{classe}'>")
        html.append(f"<td class='col-linha' style='padding-left:{padding}px'>{linha}</td>")

        for mes in meses:
            filtro = base[(base["Linha DRE"].eq(linha)) & (base["Mês"].eq(mes))]
            if filtro.empty:
                val = 0.0
                pct = 0.0
            else:
                val = numero(filtro.iloc[0]["Valor"])
                pct = numero(filtro.iloc[0]["% Receita"])

            if "MARGEM DE CONTRIBUIÇÃO PERCENTUAL" in str(linha).upper():
                val_txt = pct_br(val)
                pct_txt = "-"
            else:
                val_txt = moeda_br(val)
                pct_txt = pct_br(pct)

            html.append(f"<td class='num'>{val_txt}</td><td class='num pct'>{pct_txt}</td>")
        html.append("</tr>")

    html.append("</tbody></table></div>")
    return "".join(html)


def gerar_serie(dados, linha_busca, meses):
    return pd.DataFrame({
        "Mês": meses,
        "Valor": [valor_linha(dados, linha_busca, m) for m in meses],
    })

# =========================================================
# CSS PROFISSIONAL
# =========================================================
st.markdown(
    """
<style>
    :root {
        --bg:#07111f;
        --panel:#0b1728;
        --card:#101c2f;
        --line:#22324a;
        --blue:#0aa7ff;
        --blue-soft:#d8ecff;
        --yellow:#ffd21f;
        --text:#f8fbff;
        --muted:#aeb8c8;
        --green:#21e17b;
        --red:#ff5c6c;
    }
    .stApp {background: radial-gradient(circle at top, #0b1c30 0%, #07111f 35%, #040b14 100%); color:var(--text);}
    [data-testid="stSidebar"] {background:#050c15; border-right:1px solid #17324c;}
    .block-container {padding-top: 2.2rem; padding-bottom: 2rem; max-width: 1550px;}
    h1, h2, h3 {color:var(--text)!important; font-weight:800!important;}
    .hero {text-align:center; padding: 20px 0 14px 0;}
    .hero img {max-width: 145px; margin-bottom: 12px;}
    .hero h1 {font-size: 54px; line-height:1.05; margin: 4px 0 8px 0; letter-spacing:-1px;}
    .hero p {color:#48baff; font-size:20px; margin:0;}
    .card-grid {display:grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap:14px; margin:20px 0 10px 0;}
    .kpi-card {background:linear-gradient(180deg, rgba(21,36,58,0.98), rgba(10,20,35,0.98)); border:1px solid #1e3b5b; border-radius:18px; padding:18px 18px; box-shadow: 0 12px 30px rgba(0,0,0,0.18);}
    .kpi-title {color:var(--muted); font-size:14px; font-weight:700; text-transform:uppercase; letter-spacing:.6px;}
    .kpi-value {color:#fff; font-size:28px; font-weight:900; margin-top:8px;}
    .kpi-sub {color:#66c7ff; font-size:13px; margin-top:4px;}
    .section-title {font-size:28px; font-weight:900; margin:28px 0 8px 0;}
    .section-sub {color:var(--muted); margin-bottom:18px;}
    .dre-scroll {overflow:auto; border:1px solid #173a5a; border-radius:16px; max-height: 680px; background:#081426; box-shadow: inset 0 0 0 1px rgba(255,255,255,0.02);}
    table.dre-table {border-collapse:collapse; width:max-content; min-width:100%; font-size:16px;}
    .dre-table th {position:sticky; top:0; z-index:3; background:#171c27; color:#dce8f7; padding:14px 12px; border-bottom:2px solid #26384f; border-right:1px solid #26384f; text-align:right; white-space:nowrap; font-weight:800;}
    .dre-table th.col-linha {left:0; z-index:4; text-align:left; min-width:470px;}
    .dre-table td {padding:13px 12px; border-bottom:1px solid #20304a; border-right:1px solid #20304a; white-space:nowrap;}
    .dre-table td.col-linha {position:sticky; left:0; z-index:2; min-width:470px; text-align:left;}
    .num {text-align:right; font-weight:700;}
    .pct {min-width:85px;}
    .linha-detalhe td {background:#0b1728; color:#ffffff;}
    .linha-detalhe td.col-linha {background:#0b1728; color:#ffffff; font-weight:500;}
    .linha-azul td {background:#d7e9fb; color:#00142a; font-weight:900;}
    .linha-azul td.col-linha {background:#d7e9fb; color:#00142a; font-weight:900;}
    .linha-amarela td {background:#ffd21f; color:#06101e; font-weight:1000;}
    .linha-amarela td.col-linha {background:#ffd21f; color:#06101e; font-weight:1000;}
    .linha-grupo td {background:#eef2f7; color:#0d1827; font-weight:900; font-style:italic;}
    .linha-grupo td.col-linha {background:#eef2f7; color:#0d1827; font-weight:900; font-style:italic;}
    .alerta {background:#3b1d2a; color:#ff6976; padding:20px 24px; border-radius:12px; border:1px solid #713146; font-weight:800;}
    .ok-box {background:#083e25; border:1px solid #127a45; color:#4fff93; padding:14px 16px; border-radius:12px; font-weight:800;}
    .info-box {background:#0d2b45; border:1px solid #1e5d8c; color:#d8ecff; padding:14px 16px; border-radius:12px; font-weight:700;}
    @media (max-width: 1100px){.card-grid{grid-template-columns: repeat(2, minmax(0, 1fr));}.hero h1{font-size:40px;}}
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown("### ⚙️ Atualização")
    st.write("As informações são lidas do arquivo consolidado do DRE. Ao trocar as bases, atualize o consolidado e clique em Atualizar Base.")
    if st.button("🔄 Atualizar Base", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.divider()
    st.markdown("### 📁 Base monitorada")
    caminho_base = procurar_arquivo_base()
    if caminho_base:
        st.markdown("<div class='ok-box'>DRE_Consolidado_Moderno.xlsx</div>", unsafe_allow_html=True)
        st.caption(str(caminho_base))
    else:
        st.markdown("<div class='alerta'>Arquivo não encontrado.</div>", unsafe_allow_html=True)
        st.caption("Coloque em data/DRE_Consolidado_Moderno.xlsx ou na raiz do projeto.")

# =========================================================
# HEADER
# =========================================================
logo64 = logo_base64()
if logo64:
    logo_html = f"<img src='data:image/png;base64,{logo64}' />"
else:
    logo_html = "<div style='color:#32b7ff;font-weight:900;font-size:22px'>EIROX</div>"

st.markdown(
    f"""
<div class='hero'>
    {logo_html}
    <h1>DRE Empresa Online</h1>
    <p>Dashboard financeiro gerencial • DRE no formato aprovado • Meses fechados do ano atual</p>
</div>
""",
    unsafe_allow_html=True,
)

# =========================================================
# CARGA
# =========================================================
if not caminho_base:
    st.stop()

try:
    dados, resumo_loja, nao_class, checks = carregar_dados(str(caminho_base))
except Exception as e:
    st.markdown(f"<div class='alerta'>Erro ao ler a base: {e}</div>", unsafe_allow_html=True)
    st.stop()

if dados.empty:
    st.markdown("<div class='alerta'>A aba DADOS_DRE está vazia.</div>", unsafe_allow_html=True)
    st.stop()

# Meses disponíveis
all_months = sorted(dados["Mês"].dropna().unique().tolist(), key=mes_key)
closed_default = meses_fechados_ano_atual(all_months)
if not closed_default:
    # Fallback: usa todos os meses válidos com receita maior que zero, exceto o último se for mês atual.
    closed_default = all_months

with st.sidebar:
    st.divider()
    st.markdown("### 📅 Filtros")
    selected_months = st.multiselect(
        "Meses",
        options=all_months,
        default=closed_default,
        help="Por padrão são selecionados somente os meses fechados do ano atual.",
    )

if not selected_months:
    st.markdown("<div class='alerta'>Selecione pelo menos um mês.</div>", unsafe_allow_html=True)
    st.stop()
selected_months = sorted(selected_months, key=mes_key)
ultimo_mes = selected_months[-1]

# =========================================================
# KPIs EXECUTIVOS
# =========================================================
receita_bruta = valor_linha(dados, "1. RECEITA OPERACIONAL BRUTA", ultimo_mes)
receita_liq = valor_linha(dados, "3. (=) RECEITA OPERACIONAL LÍQUIDA", ultimo_mes)
lucro_bruto = valor_linha(dados, "5. (=) LUCRO BRUTO", ultimo_mes)
ebitda = valor_linha(dados, "EBITDA", ultimo_mes)
lair = valor_linha(dados, "RESULTADO ANTES DOS TRIBUTOS", ultimo_mes)
lucro_liq = valor_linha(dados, "11. (=) LUCRO LÍQUIDO", ultimo_mes)
cmv = valor_linha(dados, "Mercadorias (CMP)", ultimo_mes)
desp_op = valor_linha(dados, "6. (-) DESPESAS OPERACIONAIS", ultimo_mes)

margem_bruta = lucro_bruto / receita_bruta if receita_bruta else 0
margem_ebitda = ebitda / receita_bruta if receita_bruta else 0
margem_liq = lucro_liq / receita_bruta if receita_bruta else 0
cmv_pct = cmv / receita_bruta if receita_bruta else 0
desp_pct = desp_op / receita_bruta if receita_bruta else 0

st.markdown(f"<div class='section-title'>Painel Executivo — {ultimo_mes}</div>", unsafe_allow_html=True)
st.markdown(
    f"""
<div class='card-grid'>
    <div class='kpi-card'><div class='kpi-title'>Receita Bruta</div><div class='kpi-value'>{moeda_br(receita_bruta)}</div><div class='kpi-sub'>Base de faturamento</div></div>
    <div class='kpi-card'><div class='kpi-title'>Receita Líquida</div><div class='kpi-value'>{moeda_br(receita_liq)}</div><div class='kpi-sub'>{pct_br(receita_liq/receita_bruta if receita_bruta else 0)} da receita</div></div>
    <div class='kpi-card'><div class='kpi-title'>Lucro Bruto</div><div class='kpi-value'>{moeda_br(lucro_bruto)}</div><div class='kpi-sub'>Margem {pct_br(margem_bruta)}</div></div>
    <div class='kpi-card'><div class='kpi-title'>EBITDA</div><div class='kpi-value'>{moeda_br(ebitda)}</div><div class='kpi-sub'>Margem {pct_br(margem_ebitda)}</div></div>
    <div class='kpi-card'><div class='kpi-title'>Lucro Líquido</div><div class='kpi-value'>{moeda_br(lucro_liq)}</div><div class='kpi-sub'>Margem {pct_br(margem_liq)}</div></div>
    <div class='kpi-card'><div class='kpi-title'>CMV</div><div class='kpi-value'>{pct_br(cmv_pct)}</div><div class='kpi-sub'>{moeda_br(cmv)}</div></div>
    <div class='kpi-card'><div class='kpi-title'>Despesas Operacionais</div><div class='kpi-value'>{pct_br(desp_pct)}</div><div class='kpi-sub'>{moeda_br(desp_op)}</div></div>
    <div class='kpi-card'><div class='kpi-title'>LAIR</div><div class='kpi-value'>{moeda_br(lair)}</div><div class='kpi-sub'>Resultado antes dos tributos</div></div>
</div>
""",
    unsafe_allow_html=True,
)

# =========================================================
# GRÁFICOS
# =========================================================
if PLOTLY_OK:
    st.markdown("<div class='section-title'>Evolução Mensal</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    evol = pd.concat([
        gerar_serie(dados, "1. RECEITA OPERACIONAL BRUTA", selected_months).assign(Indicador="Receita Bruta"),
        gerar_serie(dados, "5. (=) LUCRO BRUTO", selected_months).assign(Indicador="Lucro Bruto"),
        gerar_serie(dados, "EBITDA", selected_months).assign(Indicador="EBITDA"),
        gerar_serie(dados, "11. (=) LUCRO LÍQUIDO", selected_months).assign(Indicador="Lucro Líquido"),
    ], ignore_index=True)
    with c1:
        fig = px.line(evol, x="Mês", y="Valor", color="Indicador", markers=True, title="Receita, Lucro Bruto, EBITDA e Lucro Líquido")
        fig.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=390)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        margens = pd.DataFrame({
            "Mês": selected_months,
            "Margem Bruta": [pct_linha(dados, "5. (=) LUCRO BRUTO", m) * 100 for m in selected_months],
            "Margem EBITDA": [pct_linha(dados, "EBITDA", m) * 100 for m in selected_months],
            "Margem Líquida": [pct_linha(dados, "11. (=) LUCRO LÍQUIDO", m) * 100 for m in selected_months],
        }).melt(id_vars="Mês", var_name="Indicador", value_name="Percentual")
        fig2 = px.line(margens, x="Mês", y="Percentual", color="Indicador", markers=True, title="Margens (%)")
        fig2.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=390, yaxis_ticksuffix="%")
        st.plotly_chart(fig2, use_container_width=True)

# =========================================================
# DRE GERENCIAL
# =========================================================
st.markdown("<div class='section-title'>DRE Gerencial</div>", unsafe_allow_html=True)
st.markdown("<div class='section-sub'>Modelo mantido no formato aprovado: meses em colunas, valores e percentuais lado a lado, com blocos estruturais destacados.</div>", unsafe_allow_html=True)
st.markdown(construir_tabela_html(dados, selected_months), unsafe_allow_html=True)

# =========================================================
# RESULTADOS ESTRATÉGICOS
# =========================================================
st.markdown("<div class='section-title'>Resultados Estratégicos</div>", unsafe_allow_html=True)
mc_total = valor_linha(dados, "Margem de Contribuição Total", ultimo_mes)
mc_pct = valor_linha(dados, "Margem de Contribuição Percentual", ultimo_mes)
pe = valor_linha(dados, "Ponto de Equilíbrio", ultimo_mes)
folga = receita_bruta - pe if pe else 0
st.markdown(
    f"""
<div class='card-grid'>
    <div class='kpi-card'><div class='kpi-title'>Margem de Contribuição</div><div class='kpi-value'>{moeda_br(mc_total)}</div><div class='kpi-sub'>{pct_br(mc_pct)}</div></div>
    <div class='kpi-card'><div class='kpi-title'>Ponto de Equilíbrio</div><div class='kpi-value'>{moeda_br(pe)}</div><div class='kpi-sub'>Receita mínima estimada</div></div>
    <div class='kpi-card'><div class='kpi-title'>Folga Operacional</div><div class='kpi-value'>{moeda_br(folga)}</div><div class='kpi-sub'>Receita - Ponto de Equilíbrio</div></div>
    <div class='kpi-card'><div class='kpi-title'>Alavancagem Operacional</div><div class='kpi-value'>{(receita_bruta / pe if pe else 0):.2f}x</div><div class='kpi-sub'>Receita / Ponto de Equilíbrio</div></div>
</div>
""",
    unsafe_allow_html=True,
)

# =========================================================
# RESULTADO POR LOJA
# =========================================================
if not resumo_loja.empty and {"Loja", "Mês", "Receita", "Despesas"}.issubset(set(resumo_loja.columns)):
    st.markdown("<div class='section-title'>Resultado por Loja</div>", unsafe_allow_html=True)
    base_loja = resumo_loja[resumo_loja["Mês"].isin(selected_months)].copy()
    for col in ["Receita", "Despesas", "Resultado Caixa Simplificado"]:
        if col in base_loja.columns:
            base_loja[col] = pd.to_numeric(base_loja[col], errors="coerce").fillna(0)
    resumo = base_loja.groupby("Loja", as_index=False).agg({
        "Receita": "sum",
        "Despesas": "sum",
        "Resultado Caixa Simplificado": "sum" if "Resultado Caixa Simplificado" in base_loja.columns else "sum",
    })
    resumo["Margem Caixa"] = resumo["Resultado Caixa Simplificado"] / resumo["Receita"].replace(0, pd.NA)
    resumo["Receita"] = resumo["Receita"].apply(moeda_br)
    resumo["Despesas"] = resumo["Despesas"].apply(moeda_br)
    resumo["Resultado Caixa Simplificado"] = resumo["Resultado Caixa Simplificado"].apply(moeda_br)
    resumo["Margem Caixa"] = resumo["Margem Caixa"].fillna(0).apply(pct_br)
    st.dataframe(resumo, use_container_width=True, hide_index=True)

# =========================================================
# AUDITORIA
# =========================================================
with st.expander("⚠️ Auditoria DRE — Não Classificados", expanded=False):
    if nao_class.empty:
        st.success("Nenhuma conta não classificada encontrada na base.")
    else:
        base_nc = nao_class.copy()
        if "Mês" in base_nc.columns:
            base_nc = base_nc[base_nc["Mês"].isin(selected_months)]
        valor_nc = pd.to_numeric(base_nc.get("Valor", 0), errors="coerce").fillna(0).sum() if not base_nc.empty else 0
        qtde_nc = pd.to_numeric(base_nc.get("Qtde", 0), errors="coerce").fillna(0).sum() if not base_nc.empty else 0
        st.markdown(
            f"<div class='info-box'>Valor não classificado no filtro: <b>{moeda_br(valor_nc)}</b> • Quantidade: <b>{int(qtde_nc)}</b></div>",
            unsafe_allow_html=True,
        )
        st.dataframe(base_nc, use_container_width=True, hide_index=True)

with st.expander("✅ Checks da Base", expanded=False):
    if checks.empty:
        st.info("A aba CHECKS não foi encontrada.")
    else:
        st.dataframe(checks, use_container_width=True, hide_index=True)
