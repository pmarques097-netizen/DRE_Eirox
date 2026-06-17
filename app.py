# -*- coding: utf-8 -*-
"""
Eirox DRE Online Premium
VERSÃO CONGELADA - DRE EIROX ENTERPRISE PREMIUM v1.0
Versão baseada na base congelada DRE_Consolidado_Moderno.xlsx.
Fonte principal: aba DADOS_DRE.

Como executar:
    streamlit run app.py
"""

from __future__ import annotations

import base64
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# =========================================================
# CONFIGURAÇÃO
# =========================================================
st.set_page_config(
    page_title="Eirox DRE Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = Path(__file__).resolve().parent
POSSIVEIS_BASES = [
    APP_DIR / "data" / "DRE_Consolidado_Moderno.xlsx",
    APP_DIR / "DRE_Consolidado_Moderno.xlsx",
    Path.cwd() / "data" / "DRE_Consolidado_Moderno.xlsx",
    Path.cwd() / "DRE_Consolidado_Moderno.xlsx",
    Path.home() / "Desktop" / "Dre" / "data" / "DRE_Consolidado_Moderno.xlsx",
    Path.home() / "Desktop" / "Dre" / "DRE_Consolidado_Moderno.xlsx",
]

POSSIVEIS_LOGOS = [
    APP_DIR / "assets" / "logo_eirox.png",
    APP_DIR / "assets" / "logo eirox(3).png",
    APP_DIR / "assets" / "logo eirox.png",
    APP_DIR / "logo_eirox.png",
    APP_DIR / "logo eirox(3).png",
    Path.cwd() / "assets" / "logo_eirox.png",
    Path.cwd() / "assets" / "logo eirox(3).png",
    Path.cwd() / "logo_eirox.png",
]

MESES_ORDEM = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


# =========================================================
# CONTROLE DE ACESSO
# =========================================================
USUARIOS = {
    "paulomarques": {"senha": "031730", "perfil": "admin", "nome": "Paulo Marques"},
    "admin": {"senha": "031730", "perfil": "admin", "nome": "Administrador"},
    "ubiratan": {"senha": "031730", "perfil": "visualizacao", "nome": "Ubiratan"},
    "vanderlei": {"senha": "031730", "perfil": "visualizacao", "nome": "Vanderlei"},
}


# =========================================================
# CSS PREMIUM
# =========================================================
st.markdown(
    """
<style>
:root{
    --bg:#07111f;
    --panel:#0b1728;
    --panel2:#101c2f;
    --line:#20334d;
    --text:#f5f7fb;
    --muted:#9aa7b8;
    --blue:#00a8ff;
    --blue2:#0068ff;
    --yellow:#ffd21f;
    --green:#32e875;
    --red:#ff5470;
}
html, body, [data-testid="stAppViewContainer"]{background: radial-gradient(circle at top, #0c2038 0%, #06101d 35%, #030914 100%) !important; color:var(--text)!important;}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#050b14 0%,#07111f 100%)!important; border-right:1px solid #103a5c;}
[data-testid="stSidebar"] *{color:#f5f7fb!important;}
.block-container{padding-top:1.2rem; padding-bottom:2rem; max-width: 1600px;}
.sidebar-logo{display:flex; justify-content:center; align-items:center; margin:12px 0 8px 0;}
.sidebar-title{text-align:center; font-size:1.05rem; font-weight:800; margin-bottom:2px; color:#ffffff;}
.sidebar-subtitle{text-align:center; color:#6cc7ff!important; font-size:.82rem; margin-bottom:18px;}
.sidebar-section{font-size:.78rem; font-weight:800; letter-spacing:.12em; color:#7ebeff!important; margin:18px 0 8px 0; text-transform:uppercase;}
div[data-baseweb="select"] > div{background:#0d1b2e!important; border-color:#1f4d77!important; border-radius:14px!important;}
.stMultiSelect [data-baseweb="tag"]{background:#00a8ff!important; color:white!important; border-radius:10px!important; font-weight:700!important;}
.hero{padding:24px 26px; border-radius:26px; background:linear-gradient(135deg,rgba(0,168,255,.18),rgba(8,18,32,.94)); border:1px solid rgba(0,168,255,.28); box-shadow: 0 12px 35px rgba(0,0,0,.30); margin-bottom:18px;}
.hero h1{font-size:3.15rem; line-height:1; margin:0; font-weight:900; letter-spacing:-.04em; color:#fff;}
.hero p{font-size:1.05rem; margin:.65rem 0 0 0; color:#8bd4ff;}
.small-meta{font-size:.86rem; color:#aab7c8; margin-top:10px;}
.kpi-card{background:linear-gradient(180deg,#10223a,#091728); border:1px solid rgba(0,168,255,.28); border-radius:22px; padding:20px 18px; box-shadow: 0 8px 28px rgba(0,0,0,.22); min-height:126px;}
.kpi-card .label{font-size:.82rem; text-transform:uppercase; letter-spacing:.09em; color:#9fb5cc; font-weight:800; margin-bottom:8px;}
.kpi-card .value{font-size:1.55rem; color:#fff; font-weight:900; white-space:nowrap;}
.kpi-card .delta{font-size:.86rem; margin-top:8px; color:#9fb5cc; font-weight:700;}
.delta-pos{color:#32e875!important;} .delta-neg{color:#ff5470!important;} .delta-neutral{color:#9fb5cc!important;}
.section-title{font-size:1.55rem; font-weight:900; color:#fff; margin:26px 0 10px 0;}
.section-caption{color:#9aa7b8; margin:-4px 0 16px 0;}
.eirox-table-wrap{overflow:auto; border:1px solid rgba(0,168,255,.25); border-radius:20px; max-height:720px; box-shadow:0 12px 35px rgba(0,0,0,.28); background:#091526;}
table.eirox-table{border-collapse:collapse; width:max-content; min-width:100%; font-size:.92rem;}
table.eirox-table th{position:sticky; top:0; z-index:2; background:#141b27; color:#c9d3df; padding:13px 14px; border-bottom:1px solid #2a3d55; border-right:1px solid #2a3d55; text-align:right; white-space:nowrap;}
table.eirox-table th:first-child{left:0; z-index:3; text-align:left; min-width:440px;}
table.eirox-table td{padding:12px 14px; border-bottom:1px solid #20334d; border-right:1px solid #20334d; text-align:right; white-space:nowrap; color:#fff; font-weight:650;}
table.eirox-table td:first-child{position:sticky; left:0; z-index:1; background:#0b1728; text-align:left; min-width:440px; max-width:520px; white-space:normal; font-weight:700;}
table.eirox-table tr.detalhe td{background:#0b1728;}
table.eirox-table tr.detalhe td:first-child{background:#0b1728; font-weight:600;}
table.eirox-table tr.subtotal_azul td{background:#cfe3f8!important; color:#00101e!important; font-weight:950;}
table.eirox-table tr.subtotal_azul td:first-child{background:#cfe3f8!important; color:#00101e!important;}
table.eirox-table tr.resultado_amarelo td{background:#ffd21f!important; color:#030914!important; font-weight:950;}
table.eirox-table tr.resultado_amarelo td:first-child{background:#ffd21f!important; color:#030914!important;}
table.eirox-table tr.agrupador td{background:#1b2534!important; color:#fff!important; font-weight:900; font-style:italic;}
table.eirox-table tr.agrupador td:first-child{background:#1b2534!important;}
/* Reforço visual final dos tópicos do DRE */
table.eirox-table tr.subtotal_azul td{box-shadow: inset 5px 0 0 #00a8ff;}
table.eirox-table tr.resultado_amarelo td{box-shadow: inset 5px 0 0 #ff9f1c; font-size:.96rem;}
table.eirox-table tr.agrupador td{box-shadow: inset 5px 0 0 #64748b;}
.audit-box{background:rgba(255,84,112,.12); border:1px solid rgba(255,84,112,.35); padding:16px 18px; border-radius:18px; color:#ffb5c1; font-weight:800;}
.ok-box{background:rgba(50,232,117,.12); border:1px solid rgba(50,232,117,.35); padding:16px 18px; border-radius:18px; color:#baf7cb; font-weight:800;}
.footer{color:#6d7d90; font-size:.78rem; text-align:center; margin-top:30px; padding-top:18px; border-top:1px solid rgba(255,255,255,.08);}
[data-testid="stMetric"]{background:linear-gradient(180deg,#10223a,#091728); padding:16px; border-radius:18px; border:1px solid rgba(0,168,255,.22);}

/* Menu lateral fixo em radio, sem dropdown */
div[role="radiogroup"] label{
    background:linear-gradient(90deg,rgba(0,168,255,.12),rgba(255,255,255,.03));
    border:1px solid rgba(0,168,255,.20);
    border-radius:14px;
    padding:10px 12px!important;
    margin:7px 0!important;
    transition:all .18s ease-in-out;
}
div[role="radiogroup"] label:hover{
    border-color:rgba(0,168,255,.55);
    background:linear-gradient(90deg,rgba(0,168,255,.25),rgba(255,255,255,.05));
}
div[role="radiogroup"] label[data-baseweb="radio"]{width:100%;}


.login-card{max-width:460px; margin:8vh auto 0 auto; padding:34px 32px; border-radius:26px; background:linear-gradient(180deg,#10223a,#07111f); border:1px solid rgba(0,168,255,.32); box-shadow:0 18px 50px rgba(0,0,0,.35);}
.login-title{text-align:center; font-size:2rem; font-weight:950; color:#fff; margin:12px 0 6px 0;}
.login-sub{text-align:center; color:#8bd4ff; font-weight:700; margin-bottom:20px;}
.user-chip{background:rgba(0,168,255,.12); border:1px solid rgba(0,168,255,.28); border-radius:14px; padding:10px 12px; margin:10px 0; color:#d9f2ff; font-weight:800; text-align:center;}
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# FUNÇÕES
# =========================================================
def encontrar_arquivo(paths: list[Path]) -> Path | None:
    for p in paths:
        if p.exists():
            return p
    return None


def img_base64(path: Path) -> str | None:
    if not path or not path.exists():
        return None
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def mes_sort_key(mes: str) -> tuple[int, int]:
    s = str(mes).lower().strip()
    if s in ("sem mês", "sem mes", "nan", "none", ""):
        return (9999, 99)
    m = re.match(r"([a-zç]{3})/(\d{2})", s)
    if not m:
        return (9999, 98)
    nome, ano2 = m.group(1), int(m.group(2))
    return (2000 + ano2, MESES_ORDEM.get(nome[:3], 99))


def meses_fechados_ano_atual(meses: list[str]) -> list[str]:
    hoje = datetime.today()
    ano_atual = hoje.year
    mes_atual = hoje.month
    filtrados = []
    for mes in meses:
        ano, mes_num = mes_sort_key(mes)
        if ano == ano_atual and mes_num < mes_atual:
            filtrados.append(mes)
    # fallback para base histórica quando ano atual do servidor for diferente do ano da base
    if not filtrados:
        anos = sorted({mes_sort_key(m)[0] for m in meses if mes_sort_key(m)[0] < 9999})
        if anos:
            ano_base = anos[-1]
            # se a base for 2026 e o servidor também estiver em 2026, exclui mês atual; caso contrário mostra todos os meses úteis
            for mes in meses:
                ano, mes_num = mes_sort_key(mes)
                if ano == ano_base and mes_num < 99:
                    if ano_base == ano_atual:
                        if mes_num < mes_atual:
                            filtrados.append(mes)
                    else:
                        filtrados.append(mes)
    return filtrados


def brl(v) -> str:
    try:
        v = float(v)
    except Exception:
        v = 0.0
    s = f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s


def pct(v) -> str:
    try:
        v = float(v)
    except Exception:
        v = 0.0
    # aceita 0.42 como 42%, aceita 42 como 42%
    if abs(v) <= 1.5:
        v = v * 100
    return f"{v:.2f}%".replace(".", ",")


def parse_float(v) -> float:
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace("%", "").replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


@st.cache_data(show_spinner=False)
def carregar_base(caminho: str):
    path = Path(caminho)
    sheets = pd.ExcelFile(path).sheet_names
    if "DADOS_DRE" not in sheets:
        raise ValueError(f"A aba DADOS_DRE não foi encontrada. Abas disponíveis: {', '.join(sheets)}")

    dados = pd.read_excel(path, sheet_name="DADOS_DRE")
    dados.columns = [str(c).strip() for c in dados.columns]
    obrigatorias = ["Ordem", "Seção", "Linha DRE", "Nível", "Tipo", "Mês", "Valor", "% Receita"]
    faltantes = [c for c in obrigatorias if c not in dados.columns]
    if faltantes:
        raise ValueError(f"Colunas ausentes em DADOS_DRE: {faltantes}")

    dados["Valor"] = dados["Valor"].apply(parse_float)
    dados["% Receita"] = dados["% Receita"].apply(parse_float)
    dados["Mês"] = dados["Mês"].astype(str).str.strip()
    dados["Linha DRE"] = dados["Linha DRE"].astype(str).str.strip()
    dados["Tipo"] = dados["Tipo"].fillna("detalhe").astype(str).str.strip()
    dados["Ordem"] = pd.to_numeric(dados["Ordem"], errors="coerce").fillna(999999).astype(int)
    dados["Nível"] = pd.to_numeric(dados["Nível"], errors="coerce").fillna(0).astype(int)

    resumo_loja = pd.DataFrame()
    if "RESUMO_LOJA" in sheets:
        resumo_loja = pd.read_excel(path, sheet_name="RESUMO_LOJA")
        resumo_loja.columns = [str(c).strip() for c in resumo_loja.columns]
        for col in ["Receita", "Despesas", "Resultado Caixa Simplificado"]:
            if col in resumo_loja.columns:
                resumo_loja[col] = resumo_loja[col].apply(parse_float)
        if "Mês" in resumo_loja.columns:
            resumo_loja["Mês"] = resumo_loja["Mês"].astype(str).str.strip()

    nao_classificados = pd.DataFrame()
    if "NAO_CLASSIFICADOS" in sheets:
        nao_classificados = pd.read_excel(path, sheet_name="NAO_CLASSIFICADOS")
        nao_classificados.columns = [str(c).strip() for c in nao_classificados.columns]
        if "Valor" in nao_classificados.columns:
            nao_classificados["Valor"] = nao_classificados["Valor"].apply(parse_float)
        if "Mês" in nao_classificados.columns:
            nao_classificados["Mês"] = nao_classificados["Mês"].astype(str).str.strip()

    checks = pd.DataFrame()
    if "CHECKS" in sheets:
        checks = pd.read_excel(path, sheet_name="CHECKS")
        checks.columns = [str(c).strip() for c in checks.columns]

    return dados, resumo_loja, nao_classificados, checks


def valor_linha(df: pd.DataFrame, linha_contains: str, mes: str) -> float:
    filtro = df["Linha DRE"].str.upper().str.contains(linha_contains.upper(), regex=False, na=False) & (df["Mês"] == mes)
    if not filtro.any():
        return 0.0
    return float(df.loc[filtro, "Valor"].iloc[0])


def pct_linha(df: pd.DataFrame, linha_contains: str, mes: str) -> float:
    filtro = df["Linha DRE"].str.upper().str.contains(linha_contains.upper(), regex=False, na=False) & (df["Mês"] == mes)
    if not filtro.any():
        return 0.0
    return float(df.loc[filtro, "% Receita"].iloc[0])


def delta_mes(df: pd.DataFrame, linha_contains: str, meses: list[str]) -> tuple[float, str]:
    if len(meses) < 2:
        return 0.0, "Sem comparação"
    atual, anterior = meses[-1], meses[-2]
    v_atual = valor_linha(df, linha_contains, atual)
    v_ant = valor_linha(df, linha_contains, anterior)
    if v_ant == 0:
        return 0.0, f"{atual} vs {anterior}"
    return (v_atual - v_ant) / abs(v_ant), f"{atual} vs {anterior}"


def make_kpi(label: str, value: str, delta_value: float | None = None, delta_label: str = ""):
    if delta_value is None:
        delta_html = ""
    else:
        cls = "delta-pos" if delta_value >= 0 else "delta-neg"
        sinal = "+" if delta_value >= 0 else ""
        delta_html = f'<div class="delta {cls}">{sinal}{pct(delta_value)} • {delta_label}</div>'
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="label">{label}</div>
            <div class="value">{value}</div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )



def classificar_destaque_linha(linha: str, tipo_original: str = "") -> str:
    """Classifica visualmente a linha do DRE sem alterar cálculo.
    Mantém o Tipo vindo da base, mas reforça os tópicos principais.
    """
    txt = str(linha).upper().strip()
    tipo_original = str(tipo_original).strip()

    resultados_amarelos = [
        "3. (=) RECEITA OPERACIONAL LÍQUIDA",
        "3. (=) RECEITA OPERACIONAL LIQUIDA",
        "5. (=) LUCRO BRUTO",
        "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO",
        "EBITDA",
        "LAJIDA",
        "9. (=) RESULTADO ANTES DOS TRIBUTOS",
        "LAIR",
        "11. (=) LUCRO LÍQUIDO",
        "11. (=) LUCRO LIQUIDO",
        "POSIÇÃO FINAL",
        "POSICAO FINAL",
        "MARGEM DE CONTRIBUIÇÃO TOTAL",
        "MARGEM DE CONTRIBUICAO TOTAL",
        "MARGEM DE CONTRIBUIÇÃO PERCENTUAL",
        "MARGEM DE CONTRIBUICAO PERCENTUAL",
        "PONTO DE EQUILÍBRIO EM VALOR MONETÁRIO",
        "PONTO DE EQUILIBRIO EM VALOR MONETARIO",
    ]
    subtotais_azuis = [
        "1. RECEITA OPERACIONAL BRUTA",
        "2. (-) DEDUÇÕES DA RECEITA BRUTA",
        "2. (-) DEDUCOES DA RECEITA BRUTA",
        "4. (-) CUSTOS DAS VENDAS",
        "6. (-) DESPESAS OPERACIONAIS",
        "8. (+/-) RESULTADO FINANCEIRO LÍQUIDO",
        "8. (+/-) RESULTADO FINANCEIRO LIQUIDO",
        "10. (-) TRIBUTOS SOBRE O LUCRO",
        "--- CONCILIAÇÃO DE FLUXO DE CAIXA",
        "--- CONCILIACAO DE FLUXO DE CAIXA",
        "RECEITAS",
        "ESTOQUE A CUSTO",
        "CMV",
        "RECEITA TOTAL",
        "DESPESAS FIXAS",
        "DESPESAS VARIÁVEIS",
        "DESPESAS VARIAVEIS",
        "CUSTOS E DESPESAS FIXAS TOTAIS",
        "CUSTO MÉDIO DE VENDA",
        "CUSTO MEDIO DE VENDA",
        "TOTAL DE CUSTOS VARIÁVEIS",
        "TOTAL DE CUSTOS VARIAVEIS",
    ]
    agrupadores = [
        "DESPESAS COM PESSOAL",
        "DESPESAS ADMINISTRATIVAS E OCUPAÇÃO",
        "DESPESAS ADMINISTRATIVAS E OCUPACAO",
        "DESPESAS COM VENDAS E MARKETING",
        "ANÁLISE DE PONTO DE EQUILÍBRIO",
        "ANALISE DE PONTO DE EQUILIBRIO",
        "CÁLCULOS INTERMEDIÁRIOS",
        "CALCULOS INTERMEDIARIOS",
    ]

    if any(p in txt for p in resultados_amarelos):
        return "resultado_amarelo"
    if any(p in txt for p in subtotais_azuis):
        return "subtotal_azul"
    if any(p in txt for p in agrupadores):
        return "agrupador"
    if tipo_original in ["subtotal_azul", "resultado_amarelo", "agrupador"]:
        return tipo_original
    return "detalhe"

def dre_pivot_html(dados: pd.DataFrame, meses: list[str], secao: str | None = None) -> str:
    base = dados.copy()
    if secao:
        base = base[base["Seção"].astype(str).str.upper() == secao.upper()]
    base = base[base["Mês"].isin(meses)]
    ordem = base[["Ordem", "Linha DRE", "Nível", "Tipo"]].drop_duplicates().sort_values("Ordem")

    linhas_html = []
    header = "<tr><th>Linha DRE</th>" + "".join([f"<th>{m} Valor</th><th>{m} %</th>" for m in meses]) + "</tr>"
    for _, row in ordem.iterrows():
        linha = row["Linha DRE"]
        tipo = classificar_destaque_linha(linha, row.get("Tipo", "detalhe"))
        nivel = int(row.get("Nível", 0))
        indent = "&nbsp;" * (nivel * 4)
        tds = [f"<td>{indent}{linha}</td>"]
        for mes in meses:
            rec = base[(base["Ordem"] == row["Ordem"]) & (base["Mês"] == mes)]
            if rec.empty:
                val, prc = 0.0, 0.0
            else:
                val = rec["Valor"].iloc[0]
                prc = rec["% Receita"].iloc[0]
            # Corrige linha de margem percentual para não aparecer como moeda
            if "MARGEM DE CONTRIBUIÇÃO PERCENTUAL" in str(linha).upper():
                tds.append(f"<td>{pct(val)}</td><td>{pct(prc)}</td>")
            else:
                tds.append(f"<td>{brl(val)}</td><td>{pct(prc)}</td>")
        linhas_html.append(f"<tr class='{tipo}'>" + "".join(tds) + "</tr>")
    return f"<div class='eirox-table-wrap'><table class='eirox-table'>{header}{''.join(linhas_html)}</table></div>"


def serie_linha(dados: pd.DataFrame, linha_contains: str, meses: list[str]) -> pd.DataFrame:
    rows = []
    for m in meses:
        rows.append({"Mês": m, "Valor": valor_linha(dados, linha_contains, m)})
    return pd.DataFrame(rows)


# =========================================================
# AUTENTICAÇÃO
# =========================================================
def tela_login() -> bool:
    """Retorna True quando o usuário está autenticado."""
    if st.session_state.get("autenticado", False):
        return True

    logo_path_login = encontrar_arquivo(POSSIVEIS_LOGOS)
    logo_html = ""
    if logo_path_login:
        b64 = img_base64(logo_path_login)
        logo_html = f"<div style='text-align:center;'><img src='data:image/png;base64,{b64}' width='190'></div>"

    st.markdown(
        f"""
        <div class='login-card'>
            {logo_html}
            <div class='login-title'>Eirox DRE Online</div>
            <div class='login-sub'>Acesso ao Painel Financeiro Enterprise</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.form("form_login", clear_on_submit=False):
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar")

    if entrar:
        user = USUARIOS.get(str(usuario).strip().lower())
        if user and senha == user["senha"]:
            st.session_state["autenticado"] = True
            st.session_state["usuario"] = str(usuario).strip().lower()
            st.session_state["nome_usuario"] = user["nome"]
            st.session_state["perfil"] = user["perfil"]
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")

    st.stop()


def botao_logout_sidebar():
    nome = st.session_state.get("nome_usuario", "Usuário")
    perfil = st.session_state.get("perfil", "")
    st.sidebar.markdown("<div class='sidebar-section'>Acesso</div>", unsafe_allow_html=True)
    st.sidebar.caption(f"Usuário: {nome}")
    st.sidebar.caption(f"Perfil: {perfil}")
    if st.sidebar.button("Sair", use_container_width=True):
        for k in ["autenticado", "usuario", "nome_usuario", "perfil"]:
            st.session_state.pop(k, None)
        st.rerun()


tela_login()

# =========================================================
# CARREGAMENTO
# =========================================================
base_path = encontrar_arquivo(POSSIVEIS_BASES)
logo_path = encontrar_arquivo(POSSIVEIS_LOGOS)

# Sidebar premium
if logo_path:
    b64 = img_base64(logo_path)
    st.sidebar.markdown(f"<div class='sidebar-logo'><img src='data:image/png;base64,{b64}' width='210'></div>", unsafe_allow_html=True)
else:
    st.sidebar.markdown("<div class='sidebar-title'>EIROX</div>", unsafe_allow_html=True)

st.sidebar.markdown("<div class='sidebar-title'>DRE Financeiro Online</div>", unsafe_allow_html=True)
st.sidebar.markdown("<div class='sidebar-subtitle'>Enterprise Premium</div>", unsafe_allow_html=True)
botao_logout_sidebar()

st.sidebar.markdown("<div class='sidebar-section'>Navegação</div>", unsafe_allow_html=True)
pagina = st.sidebar.radio(
    "",
    [
        "📊 Painel CEO",
        "📑 DRE Gerencial",
        "📈 Evolução Mensal",
        "🏪 Resultado por Loja",
        "⚖️ Resultados Estratégicos",
        "⚠️ Auditoria DRE",
    ],
    label_visibility="collapsed",
)

if not base_path:
    st.markdown("<div class='hero'><h1>DRE Empresa Online</h1><p>Base não encontrada.</p></div>", unsafe_allow_html=True)
    st.error("Não encontrei o arquivo DRE_Consolidado_Moderno.xlsx. Coloque em data/DRE_Consolidado_Moderno.xlsx ou na raiz da pasta do app.")
    st.stop()

try:
    dados, resumo_loja, nao_classificados, checks = carregar_base(str(base_path))
except Exception as e:
    st.markdown("<div class='hero'><h1>DRE Empresa Online</h1><p>Erro ao carregar base.</p></div>", unsafe_allow_html=True)
    st.error(f"Erro ao ler a base: {e}")
    st.stop()

all_months = sorted([m for m in dados["Mês"].dropna().unique().tolist() if str(m).lower() not in ["sem mês", "sem mes", "nan", "none", ""]], key=mes_sort_key)
default_months = meses_fechados_ano_atual(all_months)
if not default_months:
    default_months = all_months

st.sidebar.markdown("<div class='sidebar-section'>Filtros</div>", unsafe_allow_html=True)
meses_sel = st.sidebar.multiselect("Meses", all_months, default=default_months)
meses_sel = sorted(meses_sel, key=mes_sort_key)
if not meses_sel:
    st.warning("Selecione pelo menos um mês.")
    st.stop()

st.sidebar.markdown("<div class='footer'>EIROX FINANCIAL ANALYTICS<br>Versão Enterprise Premium</div>", unsafe_allow_html=True)

ultimo_mes = meses_sel[-1]
penultimo_mes = meses_sel[-2] if len(meses_sel) >= 2 else None

# Header
logo_top = ""
if logo_path:
    b64 = img_base64(logo_path)
    logo_top = f"<img src='data:image/png;base64,{b64}' width='118' style='margin-bottom:12px;'>"

st.markdown(
    f"""
    <div class="hero" style="text-align:center;">
        {logo_top}
        <h1>DRE Empresa Online</h1>
        <p>Dashboard financeiro gerencial • DRE no formato aprovado • Indicadores executivos premium</p>
        <div class="small-meta">Base: {base_path.name} • Período filtrado: {meses_sel[0]} a {meses_sel[-1]} • Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# KPIs fixos
delta_label = ""
cols = st.columns(5)
with cols[0]:
    d, dl = delta_mes(dados, "RECEITA OPERACIONAL BRUTA", meses_sel)
    make_kpi("Receita Bruta", brl(valor_linha(dados, "RECEITA OPERACIONAL BRUTA", ultimo_mes)), d if len(meses_sel) > 1 else None, dl)
with cols[1]:
    d, dl = delta_mes(dados, "RECEITA OPERACIONAL LÍQUIDA", meses_sel)
    make_kpi("Receita Líquida", brl(valor_linha(dados, "RECEITA OPERACIONAL LÍQUIDA", ultimo_mes)), d if len(meses_sel) > 1 else None, dl)
with cols[2]:
    d, dl = delta_mes(dados, "LUCRO BRUTO", meses_sel)
    make_kpi("Lucro Bruto", brl(valor_linha(dados, "LUCRO BRUTO", ultimo_mes)), d if len(meses_sel) > 1 else None, dl)
with cols[3]:
    d, dl = delta_mes(dados, "EBITDA", meses_sel)
    make_kpi("EBITDA", brl(valor_linha(dados, "EBITDA", ultimo_mes)), d if len(meses_sel) > 1 else None, dl)
with cols[4]:
    d, dl = delta_mes(dados, "LUCRO LÍQUIDO", meses_sel)
    make_kpi("Lucro Líquido", brl(valor_linha(dados, "LUCRO LÍQUIDO", ultimo_mes)), d if len(meses_sel) > 1 else None, dl)

st.markdown("<br>", unsafe_allow_html=True)

# =========================================================
# PÁGINAS
# =========================================================
if pagina == "📊 Painel CEO":
    st.markdown("<div class='section-title'>📊 Painel CEO</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-caption'>Visão executiva do último mês filtrado com comparação mensal.</div>", unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        make_kpi("Margem Bruta", pct(pct_linha(dados, "LUCRO BRUTO", ultimo_mes)))
    with c2:
        make_kpi("Margem EBITDA", pct(pct_linha(dados, "EBITDA", ultimo_mes)))
    with c3:
        make_kpi("Margem Líquida", pct(pct_linha(dados, "LUCRO LÍQUIDO", ultimo_mes)))

    st.markdown("<div class='section-title'>Evolução dos principais resultados</div>", unsafe_allow_html=True)
    evol = []
    for nome, busca in [("Receita Bruta", "RECEITA OPERACIONAL BRUTA"), ("Lucro Bruto", "LUCRO BRUTO"), ("EBITDA", "EBITDA"), ("Lucro Líquido", "LUCRO LÍQUIDO")]:
        tmp = serie_linha(dados, busca, meses_sel)
        tmp["Indicador"] = nome
        evol.append(tmp)
    evol_df = pd.concat(evol, ignore_index=True)
    fig = px.line(evol_df, x="Mês", y="Valor", color="Indicador", markers=True, template="plotly_dark")
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=430, margin=dict(l=20,r=20,t=30,b=20))
    st.plotly_chart(fig, use_container_width=True)

elif pagina == "📑 DRE Gerencial":
    st.markdown("<div class='section-title'>📑 DRE Gerencial</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-caption'>Modelo completo no formato aprovado: DRE, conciliação de fluxo de caixa, indicadores e ponto de equilíbrio ao final.</div>", unsafe_allow_html=True)
    st.markdown(dre_pivot_html(dados, meses_sel, secao=None), unsafe_allow_html=True)

elif pagina == "📈 Evolução Mensal":
    st.markdown("<div class='section-title'>📈 Evolução Mensal</div>", unsafe_allow_html=True)
    indicadores = {
        "Receita Bruta": "RECEITA OPERACIONAL BRUTA",
        "Receita Líquida": "RECEITA OPERACIONAL LÍQUIDA",
        "Lucro Bruto": "LUCRO BRUTO",
        "EBITDA": "EBITDA",
        "Lucro Líquido": "LUCRO LÍQUIDO",
        "CMV": "CUSTOS DAS VENDAS",
    }
    escolha = st.multiselect("Indicadores", list(indicadores.keys()), default=["Receita Bruta", "EBITDA", "Lucro Líquido"])
    evol = []
    for nome in escolha:
        tmp = serie_linha(dados, indicadores[nome], meses_sel)
        tmp["Indicador"] = nome
        evol.append(tmp)
    if evol:
        evol_df = pd.concat(evol, ignore_index=True)
        fig = px.bar(evol_df, x="Mês", y="Valor", color="Indicador", barmode="group", template="plotly_dark", text_auto=False)
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=480)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Selecione pelo menos um indicador.")

elif pagina == "🏪 Resultado por Loja":
    st.markdown("<div class='section-title'>🏪 Resultado por Loja</div>", unsafe_allow_html=True)
    if resumo_loja.empty:
        st.info("A aba RESUMO_LOJA não foi encontrada na base.")
    else:
        loja_df = resumo_loja[resumo_loja["Mês"].isin(meses_sel)].copy()
        if loja_df.empty:
            st.info("Sem dados por loja para os meses selecionados.")
        else:
            resumo = loja_df.groupby("Loja", dropna=False)[["Receita", "Despesas", "Resultado Caixa Simplificado"]].sum().reset_index()
            resumo = resumo.sort_values("Resultado Caixa Simplificado", ascending=False)
            fig = px.bar(resumo, x="Loja", y="Resultado Caixa Simplificado", template="plotly_dark", text_auto=".2s")
            fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=420)
            st.plotly_chart(fig, use_container_width=True)
            show = resumo.copy()
            for col in ["Receita", "Despesas", "Resultado Caixa Simplificado"]:
                show[col] = show[col].apply(brl)
            st.dataframe(show, use_container_width=True, hide_index=True)

elif pagina == "⚖️ Resultados Estratégicos":
    st.markdown("<div class='section-title'>⚖️ Resultados Estratégicos</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-caption'>Ponto de equilíbrio, margem de contribuição e custos variáveis em formato executivo.</div>", unsafe_allow_html=True)
    sec = dados[dados["Seção"].astype(str).str.upper().str.contains("PONTO", na=False)]
    if sec.empty:
        sec = dados[dados["Linha DRE"].str.upper().str.contains("MARGEM DE CONTRIBUIÇÃO|PONTO DE EQUILÍBRIO|RECEITA TOTAL|DESPESAS FIXAS|DESPESAS VARIÁVEIS|CUSTO MÉDIO", na=False, regex=True)]
    # Cards principais do último mês
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        make_kpi("Margem de Contribuição", brl(valor_linha(dados, "MARGEM DE CONTRIBUIÇÃO TOTAL", ultimo_mes)))
    with c2:
        make_kpi("Margem Contribuição %", pct(valor_linha(dados, "MARGEM DE CONTRIBUIÇÃO PERCENTUAL", ultimo_mes)))
    with c3:
        make_kpi("Ponto de Equilíbrio", brl(valor_linha(dados, "PONTO DE EQUILÍBRIO", ultimo_mes)))
    with c4:
        folga = valor_linha(dados, "RECEITA TOTAL", ultimo_mes) - valor_linha(dados, "PONTO DE EQUILÍBRIO", ultimo_mes)
        make_kpi("Folga Operacional", brl(folga))
    if sec.empty:
        st.info("Não encontrei a seção de ponto de equilíbrio na base.")
    else:
        st.markdown(dre_pivot_html(sec, meses_sel), unsafe_allow_html=True)

elif pagina == "⚠️ Auditoria DRE":
    st.markdown("<div class='section-title'>⚠️ Auditoria DRE</div>", unsafe_allow_html=True)
    if nao_classificados.empty:
        st.markdown("<div class='ok-box'>Nenhum item não classificado encontrado.</div>", unsafe_allow_html=True)
    else:
        aud = nao_classificados.copy()
        if "Mês" in aud.columns:
            aud = aud[aud["Mês"].isin(meses_sel) | aud["Mês"].astype(str).str.lower().isin(["sem mês", "sem mes"])]
        valor_nc = aud["Valor"].sum() if "Valor" in aud.columns else 0
        qtde_nc = aud["Qtde"].sum() if "Qtde" in aud.columns else len(aud)
        c1, c2 = st.columns(2)
        with c1:
            make_kpi("Valor não classificado", brl(valor_nc))
        with c2:
            make_kpi("Qtde não classificada", f"{int(qtde_nc):,}".replace(",", "."))
        show = aud.copy()
        if "Valor" in show.columns:
            show["Valor"] = show["Valor"].apply(brl)
        st.dataframe(show, use_container_width=True, hide_index=True)

st.markdown("<div class='footer'>EIROX FINANCIAL ANALYTICS • DRE Online Premium • Base congelada evolutiva</div>", unsafe_allow_html=True)
