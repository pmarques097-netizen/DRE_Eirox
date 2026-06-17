# -*- coding: utf-8 -*-
"""
Eirox DRE Online - versão com busca completa dos meses
- Mantém o modelo DRE aprovado: meses em colunas, Valor e % lado a lado
- Lê a base consolidada automaticamente de data/DRE_Consolidado_Moderno.xlsx ou da raiz do projeto
- Filtra por padrão somente os meses fechados do ano atual
- Não limita mais aos últimos 3 meses
- Formata valores em moeda brasileira
"""

from pathlib import Path
from datetime import date
import re
import base64
import pandas as pd
import streamlit as st

# =========================
# CONFIGURAÇÃO DA PÁGINA
# =========================
st.set_page_config(
    page_title="Eirox DRE Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# ESTILO DARK ENTERPRISE
# =========================
st.markdown(
    """
    <style>
    :root {
        --bg: #07111f;
        --panel: #0c192b;
        --panel2: #101826;
        --line: #1c334e;
        --text: #f6f9ff;
        --muted: #aab4c0;
        --blue: #00a7ff;
        --blue2: #0d6efd;
        --yellow: #ffd21a;
        --cyan: #4cc9ff;
    }
    .stApp { background: var(--bg); color: var(--text); }
    [data-testid="stSidebar"] { background: #050b13; border-right: 1px solid #14304c; }
    h1, h2, h3 { color: var(--text); font-weight: 800; }
    .subtitle { color: #79c7ff; font-size: 1.05rem; text-align:center; margin-top:-10px; }
    .kpi-grid { display:grid; grid-template-columns: repeat(4, minmax(180px, 1fr)); gap:14px; margin: 18px 0 24px 0; }
    .kpi-card { background: linear-gradient(145deg,#0c192b,#091525); border:1px solid #173b60; border-radius:18px; padding:18px; box-shadow:0 10px 28px rgba(0,0,0,.25); }
    .kpi-label { color:#9fb0c0; font-size:.88rem; font-weight:700; text-transform:uppercase; letter-spacing:.04em; }
    .kpi-value { color:#ffffff; font-size:1.65rem; font-weight:900; margin-top:8px; }
    .kpi-sub { color:#69c7ff; font-size:.85rem; margin-top:5px; }
    .eirox-alert { background:#3b202b; border:1px solid #67364a; color:#ff7070; padding:18px 22px; border-radius:12px; font-weight:700; }
    .dre-wrap { overflow:auto; max-height: 72vh; border:1px solid #173b60; border-radius:16px; background:#0b1628; }
    table.dre-table { border-collapse: separate; border-spacing:0; width:100%; min-width:1200px; font-size:15px; }
    .dre-table th { position:sticky; top:0; z-index:2; background:#171b25; color:#c8d1dd; text-align:right; border-bottom:1px solid #25364b; padding:13px 12px; white-space:nowrap; }
    .dre-table th:first-child { text-align:left; position:sticky; left:0; z-index:3; min-width:420px; }
    .dre-table td { padding:12px 12px; border-bottom:1px solid #1b2b40; border-right:1px solid #1b2b40; color:#ffffff; text-align:right; white-space:nowrap; }
    .dre-table td:first-child { text-align:left; position:sticky; left:0; z-index:1; background:inherit; min-width:420px; }
    .row-base td { background:#cfe0f2 !important; color:#00152a !important; font-weight:900; }
    .row-result td { background:#ffd21a !important; color:#00152a !important; font-weight:950; }
    .row-group td { background:#182d46 !important; color:#ffffff !important; font-weight:900; }
    .row-result td { background:#ffd21a !important; color:#00152a !important; font-weight:950; box-shadow: inset 0 0 0 1px rgba(0,0,0,.08); }
    .row-base td { background:#cfe0f2 !important; color:#00152a !important; font-weight:900; }
    .row-normal td { background:#081629; }
    .row-normal:hover td { background:#10233a; }
    .logo-box { text-align:center; margin-top: 8px; }
    .logo-box img { max-width: 160px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================
# FUNÇÕES AUXILIARES
# =========================
MESES_PT = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}
MESES_ORDEM = {v: k for k, v in MESES_PT.items()}


def br_money(v):
    try:
        v = float(v)
    except Exception:
        v = 0.0
    s = f"R$ {v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def br_pct(v):
    try:
        v = float(v)
    except Exception:
        v = 0.0
    if abs(v) > 1.5:  # caso venha como 38 em vez de 0.38
        v = v / 100
    return f"{v*100:.2f}%".replace(".", ",")


def mes_para_tuple(mes):
    m = re.match(r"^([a-zç]{3})/(\d{2})$", str(mes).strip().lower())
    if not m:
        return (9999, 99)
    nome, ano2 = m.groups()
    return (2000 + int(ano2), MESES_PT.get(nome, 99))


def ordenar_meses(meses):
    return sorted([m for m in meses if isinstance(m, str)], key=mes_para_tuple)


def meses_fechados_ano_atual(meses):
    hoje = date.today()
    ano_atual = hoje.year
    mes_atual = hoje.month
    fechados = []
    for m in ordenar_meses(meses):
        ano, mes_num = mes_para_tuple(m)
        if ano == ano_atual and mes_num < mes_atual:
            fechados.append(m)
    # fallback para ambientes onde a data do servidor não bata com a base
    if not fechados:
        meses_validos = [m for m in ordenar_meses(meses) if mes_para_tuple(m)[0] != 9999]
        if len(meses_validos) > 1:
            fechados = meses_validos[:-1]
        else:
            fechados = meses_validos
    return fechados


def localizar_base():
    cwd = Path.cwd()
    candidatos_base = [
        cwd,
        cwd.parent,
        Path(__file__).resolve().parent,
        Path(__file__).resolve().parent.parent,
        Path.home() / "Desktop" / "Dre",
    ]
    nomes = [
        Path("data") / "DRE_Consolidado_Moderno.xlsx",
        Path("DRE_Consolidado_Moderno.xlsx"),
        Path("data") / "DRE_Consolidado_Robusto.xlsx",
        Path("DRE_Consolidado_Robusto.xlsx"),
    ]
    for base in candidatos_base:
        for nome in nomes:
            arq = base / nome
            if arq.exists():
                return arq
    return None


def localizar_logo():
    cwd = Path.cwd()
    candidatos = [
        cwd / "assets" / "logo_eirox.png",
        cwd / "assets" / "logo eirox(3).png",
        Path(__file__).resolve().parent / "assets" / "logo_eirox.png",
        Path(__file__).resolve().parent / "assets" / "logo eirox(3).png",
        Path.home() / "Desktop" / "Dre" / "assets" / "logo_eirox.png",
    ]
    for p in candidatos:
        if p.exists():
            return p
    return None


@st.cache_data(show_spinner=False)
def carregar_excel(path_str):
    path = Path(path_str)
    xl = pd.ExcelFile(path)
    sheets = {}
    for sh in xl.sheet_names:
        try:
            sheets[sh] = pd.read_excel(path, sheet_name=sh)
        except Exception:
            pass
    return sheets


def montar_dre_formatado(df, meses_sel):
    # Esperado: aba DRE_ESTRUTURADA com Linha DRE, Tipo, Nível, Ordem e colunas <mes> Valor/%
    df = df.copy()
    if "Ordem" in df.columns:
        df = df.sort_values("Ordem")
    cols = ["Linha DRE"]
    for m in meses_sel:
        if f"{m} Valor" in df.columns:
            cols.append(f"{m} Valor")
        if f"{m} %" in df.columns:
            cols.append(f"{m} %")
    cols = [c for c in cols if c in df.columns]
    out = df[cols].copy()

    # Formata todos os valores em padrão brasileiro.
    # Exceção: linhas percentuais estratégicas precisam aparecer como %, não como R$.
    linhas_percentuais = [
        "MARGEM DE CONTRIBUIÇÃO PERCENTUAL",
        "MARGEM DE CONTRIBUICAO PERCENTUAL",
    ]

    for idx, row in out.iterrows():
        linha_txt = str(row.get("Linha DRE", "")).upper()
        linha_percentual = any(x in linha_txt for x in linhas_percentuais)

        for c in out.columns:
            if c == "Linha DRE":
                continue
            if c.endswith(" Valor"):
                out.at[idx, c] = br_pct(row[c]) if linha_percentual else br_money(row[c])
            elif c.endswith(" %"):
                out.at[idx, c] = br_pct(row[c])
    return out


def row_class(linha):
    t = str(linha).upper()

    # Resultados principais e indicadores finais: destaque amarelo executivo
    resultados = [
        "(=)",
        "LUCRO BRUTO",
        "EBITDA",
        "LAJIDA",
        "LAIR",
        "LUCRO LÍQUIDO",
        "LUCRO LIQUIDO",
        "RECEITA OPERACIONAL LÍQUIDA",
        "RECEITA OPERACIONAL LIQUIDA",
        "RECEITA TOTAL",
        "MARGEM DE CONTRIBUIÇÃO TOTAL",
        "MARGEM DE CONTRIBUICAO TOTAL",
        "MARGEM DE CONTRIBUIÇÃO PERCENTUAL",
        "MARGEM DE CONTRIBUICAO PERCENTUAL",
        "PONTO DE EQUILÍBRIO",
        "PONTO DE EQUILIBRIO",
        "POSIÇÃO FINAL",
        "POSICAO FINAL",
    ]
    if any(x in t for x in resultados):
        return "row-result"

    # Blocos estruturais numerados: azul claro Eirox
    if re.match(r"^\s*\d+\.", t):
        return "row-base"

    # Agrupadores e linhas de análise: cinza premium
    grupos = [
        "DESPESAS COM PESSOAL",
        "DESPESAS ADMINISTRATIVAS",
        "DESPESAS COM VENDAS",
        "CONCILIAÇÃO",
        "CONCILIACAO",
        "ANÁLISE DE PONTO DE EQUILÍBRIO",
        "ANALISE DE PONTO DE EQUILIBRIO",
        "CÁLCULOS INTERMEDIÁRIOS",
        "CALCULOS INTERMEDIARIOS",
        "DESPESAS FIXAS",
        "DESPESAS VARIÁVEIS",
        "DESPESAS VARIAVEIS",
        "CUSTOS E DESPESAS FIXAS",
        "CUSTO MÉDIO DE VENDA",
        "CUSTO MEDIO DE VENDA",
        "TOTAL DE CUSTOS VARIÁVEIS",
        "TOTAL DE CUSTOS VARIAVEIS",
    ]
    if any(x in t for x in grupos):
        return "row-group"

    return "row-normal"


def dataframe_para_html(df):
    html = ['<div class="dre-wrap"><table class="dre-table"><thead><tr>']
    for c in df.columns:
        html.append(f"<th>{c}</th>")
    html.append("</tr></thead><tbody>")
    for _, r in df.iterrows():
        cls = row_class(r.get("Linha DRE", ""))
        html.append(f'<tr class="{cls}">')
        for c in df.columns:
            html.append(f"<td>{r.get(c, '')}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "".join(html)


def pegar_valor_linha(dre, linha_contains, mes):
    if dre is None or dre.empty:
        return 0
    filtro = dre["Linha DRE"].astype(str).str.upper().str.contains(linha_contains.upper(), regex=False, na=False)
    col = f"{mes} Valor"
    if not filtro.any() or col not in dre.columns:
        return 0
    return float(pd.to_numeric(dre.loc[filtro, col], errors="coerce").fillna(0).iloc[0])


def pegar_pct_linha(dre, linha_contains, mes):
    if dre is None or dre.empty:
        return 0
    filtro = dre["Linha DRE"].astype(str).str.upper().str.contains(linha_contains.upper(), regex=False, na=False)
    col = f"{mes} %"
    if not filtro.any() or col not in dre.columns:
        return 0
    return float(pd.to_numeric(dre.loc[filtro, col], errors="coerce").fillna(0).iloc[0])

# =========================
# SIDEBAR
# =========================
with st.sidebar:
    st.markdown("### ⚙️ Atualização")
    st.caption("As informações são lidas do arquivo consolidado gerado pelo banco do DRE. Ao trocar as bases, atualize o consolidado e clique em Atualizar Base.")
    if st.button("🔄 Atualizar Base", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    st.markdown("### 📁 Base monitorada")
    base_path = localizar_base()
    if base_path:
        st.success(base_path.name)
        st.caption(str(base_path))
    else:
        st.error("DRE_Consolidado_Moderno.xlsx não encontrado")

# =========================
# CABEÇALHO
# =========================
logo = localizar_logo()
if logo:
    with open(logo, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    st.markdown(f'<div class="logo-box"><img src="data:image/png;base64,{b64}" /></div>', unsafe_allow_html=True)

st.markdown("<h1 style='text-align:center;font-size:3.5rem;margin-top:0;'>DRE Empresa Online</h1>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Dashboard financeiro gerencial no padrão Eirox Pricing Online • Meses fechados do ano atual • Busca completa das bases</div>", unsafe_allow_html=True)

# =========================
# CARREGAMENTO
# =========================
if not base_path:
    st.markdown("<div class='eirox-alert'>Base não encontrada. Coloque DRE_Consolidado_Moderno.xlsx na raiz do projeto ou dentro da pasta data.</div>", unsafe_allow_html=True)
    st.stop()

sheets = carregar_excel(str(base_path))
dre = sheets.get("DRE_ESTRUTURADA")
dados = sheets.get("DADOS_DRE")
nao_class = sheets.get("NAO_CLASSIFICADOS", pd.DataFrame())
resumo_loja = sheets.get("RESUMO_LOJA", pd.DataFrame())

if dre is None or dre.empty:
    st.markdown("<div class='eirox-alert'>A aba DRE_ESTRUTURADA não foi encontrada no consolidado.</div>", unsafe_allow_html=True)
    st.stop()

# identifica todos os meses presentes nas colunas do DRE, não apenas últimos 3
all_months = []
for c in dre.columns:
    m = re.match(r"^([a-zç]{3}/\d{2}) Valor$", str(c).strip().lower())
    if m:
        all_months.append(m.group(1))
all_months = ordenar_meses(list(dict.fromkeys(all_months)))

# remove colunas técnicas sem mês
all_months_validos = [m for m in all_months if m.lower() != "sem mês"]
default_months = meses_fechados_ano_atual(all_months_validos)

if not all_months_validos:
    st.markdown("<div class='eirox-alert'>Nenhum mês foi identificado na aba DRE_ESTRUTURADA.</div>", unsafe_allow_html=True)
    st.stop()

with st.sidebar:
    st.divider()
    st.markdown("### 📅 Meses")
    st.caption("Padrão: meses fechados do ano atual. Agora o sistema busca todos os meses existentes no arquivo.")
    selected_months = st.multiselect(
        "Selecione os meses",
        options=all_months_validos,
        default=default_months,
        key="meses_dre",
    )
    if not selected_months:
        selected_months = default_months or all_months_validos

selected_months = ordenar_meses(selected_months)
ultimo = selected_months[-1]

# =========================
# KPIS
# =========================
receita = pegar_valor_linha(dre, "1. RECEITA OPERACIONAL BRUTA", ultimo)
receita_liquida = pegar_valor_linha(dre, "3. (=) RECEITA OPERACIONAL LÍQUIDA", ultimo)
lucro_bruto = pegar_valor_linha(dre, "5. (=) LUCRO BRUTO", ultimo)
ebitda = pegar_valor_linha(dre, "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO", ultimo)
lucro_liquido = pegar_valor_linha(dre, "11. (=) LUCRO LÍQUIDO", ultimo)
mg_ebitda = pegar_pct_linha(dre, "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO", ultimo)
mg_liquida = pegar_pct_linha(dre, "11. (=) LUCRO LÍQUIDO", ultimo)

st.markdown(f"<p style='text-align:center;color:#aab4c0;'>Último mês selecionado: <b>{ultimo}</b> • Meses encontrados na base: <b>{', '.join(all_months_validos)}</b></p>", unsafe_allow_html=True)

st.markdown(
    f"""
    <div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Receita Bruta</div><div class="kpi-value">{br_money(receita)}</div><div class="kpi-sub">{ultimo}</div></div>
        <div class="kpi-card"><div class="kpi-label">Receita Líquida</div><div class="kpi-value">{br_money(receita_liquida)}</div><div class="kpi-sub">{br_pct(receita_liquida/receita if receita else 0)}</div></div>
        <div class="kpi-card"><div class="kpi-label">Lucro Bruto</div><div class="kpi-value">{br_money(lucro_bruto)}</div><div class="kpi-sub">Margem {br_pct(lucro_bruto/receita if receita else 0)}</div></div>
        <div class="kpi-card"><div class="kpi-label">EBITDA</div><div class="kpi-value">{br_money(ebitda)}</div><div class="kpi-sub">Margem {br_pct(mg_ebitda)}</div></div>
        <div class="kpi-card"><div class="kpi-label">Lucro Líquido</div><div class="kpi-value">{br_money(lucro_liquido)}</div><div class="kpi-sub">Margem {br_pct(mg_liquida)}</div></div>
        <div class="kpi-card"><div class="kpi-label">Meses no filtro</div><div class="kpi-value">{len(selected_months)}</div><div class="kpi-sub">{', '.join(selected_months)}</div></div>
        <div class="kpi-card"><div class="kpi-label">Não classificados</div><div class="kpi-value">{0 if nao_class.empty else len(nao_class)}</div><div class="kpi-sub">Fora do DRE principal</div></div>
        <div class="kpi-card"><div class="kpi-label">Base</div><div class="kpi-value">OK</div><div class="kpi-sub">{base_path.name}</div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================
# ABAS
# =========================
aba1, aba2, aba3, aba4 = st.tabs(["📊 Painel Executivo", "📋 DRE Gerencial", "🏪 Resultado por Loja", "⚠️ Auditoria DRE"])

with aba1:
    st.subheader("Evolução mensal")
    evol = []
    for m in selected_months:
        evol.append({
            "Mês": m,
            "Receita Bruta": pegar_valor_linha(dre, "1. RECEITA OPERACIONAL BRUTA", m),
            "Lucro Bruto": pegar_valor_linha(dre, "5. (=) LUCRO BRUTO", m),
            "EBITDA": pegar_valor_linha(dre, "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO", m),
            "Lucro Líquido": pegar_valor_linha(dre, "11. (=) LUCRO LÍQUIDO", m),
        })
    evol_df = pd.DataFrame(evol).set_index("Mês")
    st.line_chart(evol_df)

with aba2:
    st.subheader("DRE Gerencial")
    st.caption("Modelo mantido no formato aprovado: meses em colunas, valores e percentuais lado a lado, com todos os blocos estruturais destacados.")
    dre_fmt = montar_dre_formatado(dre, selected_months)
    st.markdown(dataframe_para_html(dre_fmt), unsafe_allow_html=True)

with aba3:
    st.subheader("Resultado por Loja")
    if resumo_loja is not None and not resumo_loja.empty:
        loja_df = resumo_loja[resumo_loja["Mês"].isin(selected_months)].copy() if "Mês" in resumo_loja.columns else resumo_loja.copy()
        for c in ["Receita", "Despesas", "Resultado Caixa Simplificado"]:
            if c in loja_df.columns:
                loja_df[c] = pd.to_numeric(loja_df[c], errors="coerce").fillna(0)
        st.dataframe(loja_df, use_container_width=True, hide_index=True)
    else:
        st.info("Resumo por loja não encontrado no consolidado.")

with aba4:
    st.subheader("Auditoria de não classificados")
    st.caption("Esses itens ficam fora dos cálculos principais até serem classificados no plano de contas.")
    if nao_class is not None and not nao_class.empty:
        st.dataframe(nao_class, use_container_width=True, hide_index=True)
    else:
        st.success("Nenhum item não classificado encontrado.")
