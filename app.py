# -*- coding: utf-8 -*-
"""
Eirox DRE Online - Premium CEO
Base congelada: lê diretamente a aba DADOS_DRE do DRE_Consolidado_Moderno.xlsx.
Melhoria 1 aplicada: Página CEO com cards executivos e comparação Mês Atual vs Mês Anterior.
"""
from pathlib import Path
from datetime import datetime
import base64
import re
import pandas as pd
import streamlit as st

try:
    import plotly.express as px
    PLOTLY_OK = True
except Exception:
    PLOTLY_OK = False

st.set_page_config(page_title="Eirox DRE Online", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

MESES_PT = {"jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,"jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12}

# =========================================================
# UTILITÁRIOS
# =========================================================
def moeda_br(valor) -> str:
    try:
        v = float(valor)
    except Exception:
        v = 0.0
    s = f"R$ {v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def pct_br(valor) -> str:
    try:
        v = float(valor)
    except Exception:
        v = 0.0
    if abs(v) <= 2:
        v *= 100
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
    hoje = datetime.today()
    ano_atual, mes_atual = hoje.year, hoje.month
    fechados = []
    for m in meses:
        ano, mes_num = mes_key(m)
        if ano == ano_atual and mes_num < mes_atual:
            fechados.append(m)
    return fechados

def procurar_arquivo_base() -> Path | None:
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

def delta_abs(valor_atual, valor_anterior):
    return valor_atual - valor_anterior

def delta_pct(valor_atual, valor_anterior):
    if not valor_anterior:
        return 0.0
    return (valor_atual - valor_anterior) / abs(valor_anterior)

def cor_delta(valor):
    return "delta-pos" if valor >= 0 else "delta-neg"

def texto_delta(valor_atual, valor_anterior, is_pct=False):
    d = delta_abs(valor_atual, valor_anterior)
    dp = delta_pct(valor_atual, valor_anterior)
    sinal = "+" if d >= 0 else ""
    if is_pct:
        return f"{sinal}{pct_br(d)} | {sinal}{pct_br(dp)} vs mês anterior"
    return f"{sinal}{moeda_br(d)} | {sinal}{pct_br(dp)} vs mês anterior"

def gerar_serie(dados, linha_busca, meses):
    return pd.DataFrame({"Mês": meses, "Valor": [valor_linha(dados, linha_busca, m) for m in meses]})

def construir_tabela_html(dados, meses):
    base = dados[dados["Mês"].isin(meses)].copy()
    if base.empty:
        return "<div class='alerta'>Nenhum dado encontrado para os meses selecionados.</div>"
    ordem_linhas = base[["Ordem", "Seção", "Linha DRE", "Nível", "Tipo"]].drop_duplicates().sort_values("Ordem")

    def classe_linha(linha, tipo):
        linha_up = str(linha).upper()
        tipo = str(tipo).lower()
        resultados = ["RECEITA TOTAL", "MARGEM DE CONTRIBUIÇÃO TOTAL", "MARGEM DE CONTRIBUIÇÃO PERCENTUAL", "PONTO DE EQUILÍBRIO"]
        blocos = ["DESPESAS FIXAS", "DESPESAS VARIÁVEIS", "CUSTOS E DESPESAS FIXAS TOTAIS", "CUSTO MÉDIO DE VENDA", "TOTAL DE CUSTOS VARIÁVEIS"]
        if any(x in linha_up for x in resultados) or "resultado_amarelo" in tipo:
            return "linha-amarela"
        if any(x in linha_up for x in blocos) or "subtotal_azul" in tipo:
            return "linha-azul"
        if "grupo_italico" in tipo:
            return "linha-grupo"
        return "linha-detalhe"

    html = ["<div class='dre-scroll'><table class='dre-table'>", "<thead><tr><th class='col-linha'>Linha DRE</th>"]
    for mes in meses:
        html.append(f"<th>{mes} Valor</th><th>{mes} %</th>")
    html.append("</tr></thead><tbody>")

    for _, row in ordem_linhas.iterrows():
        linha = row["Linha DRE"]
        classe = classe_linha(linha, row["Tipo"])
        padding = 12 + (int(row["Nível"]) * 18)
        html.append(f"<tr class='{classe}'><td class='col-linha' style='padding-left:{padding}px'>{linha}</td>")
        for mes in meses:
            filtro = base[(base["Linha DRE"].eq(linha)) & (base["Mês"].eq(mes))]
            if filtro.empty:
                val, pct = 0.0, 0.0
            else:
                val, pct = numero(filtro.iloc[0]["Valor"]), numero(filtro.iloc[0]["% Receita"])
            if "MARGEM DE CONTRIBUIÇÃO PERCENTUAL" in str(linha).upper():
                val_txt, pct_txt = pct_br(val), "-"
            else:
                val_txt, pct_txt = moeda_br(val), pct_br(pct)
            html.append(f"<td class='num'>{val_txt}</td><td class='num pct'>{pct_txt}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "".join(html)

# =========================================================
# CSS PREMIUM
# =========================================================
st.markdown("""
<style>
:root{--bg:#050b14;--panel:#07111f;--blue:#00aaff;--cyan:#5bd7ff;--yellow:#ffd21f;--text:#f8fbff;--muted:#a8b3c4;--green:#24e782;--red:#ff5c6c;}
.stApp{background:radial-gradient(circle at 20% 0%,rgba(0,170,255,.16) 0%,transparent 28%),linear-gradient(145deg,#050b14 0%,#071527 45%,#03070d 100%);color:var(--text)}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#040912 0%,#07111f 100%);border-right:1px solid rgba(0,170,255,.35);box-shadow:8px 0 30px rgba(0,0,0,.28)}
[data-testid="stSidebar"] *{color:#f8fbff}.block-container{padding-top:1.4rem;padding-bottom:2.5rem;max-width:1580px}h1,h2,h3{color:var(--text)!important;font-weight:900!important}
.hero-premium{position:relative;overflow:hidden;text-align:center;margin:0 auto 22px auto;padding:34px 30px 30px;border:1px solid rgba(0,170,255,.28);border-radius:26px;background:linear-gradient(135deg,rgba(0,170,255,.12),rgba(255,255,255,.02)),linear-gradient(180deg,rgba(12,30,52,.95),rgba(4,10,18,.98));box-shadow:0 24px 70px rgba(0,0,0,.35),inset 0 1px 0 rgba(255,255,255,.05)}
.hero-premium img{max-width:170px;margin-bottom:14px}.hero-premium h1{font-size:54px;line-height:1.02;margin:4px 0 12px}.hero-premium p{color:#64cfff;font-size:19px;margin:0}.premium-badge{display:inline-flex;padding:8px 14px;border:1px solid rgba(0,170,255,.4);border-radius:999px;background:rgba(0,170,255,.08);color:#bfeeff;font-weight:800;font-size:12px;letter-spacing:.5px;text-transform:uppercase;margin-bottom:12px}
.card-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin:20px 0 10px}.card-grid-5{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:16px;margin:20px 0 10px}.card-grid-3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px;margin:20px 0 10px}
.kpi-card{position:relative;overflow:hidden;background:linear-gradient(135deg,rgba(0,170,255,.10),rgba(255,255,255,.02)),linear-gradient(180deg,rgba(18,35,58,.98),rgba(8,18,32,.98));border:1px solid rgba(91,215,255,.20);border-radius:22px;padding:20px;box-shadow:0 18px 48px rgba(0,0,0,.28),inset 0 1px 0 rgba(255,255,255,.06)}
.kpi-card:after{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,var(--blue),var(--cyan),transparent)}.kpi-title{color:#9eb4ce;font-size:13px;font-weight:900;text-transform:uppercase;letter-spacing:.7px}.kpi-value{color:#fff;font-size:28px;font-weight:1000;margin-top:9px;letter-spacing:-.5px}.kpi-sub{color:#62cfff;font-size:13px;margin-top:6px;font-weight:700}.delta-pos{color:#24e782!important}.delta-neg{color:#ff5c6c!important}.ceo-card{min-height:142px}.ceo-card .kpi-value{font-size:26px}.ceo-highlight{border-color:rgba(255,210,31,.55);box-shadow:0 18px 52px rgba(255,210,31,.10)}.ceo-highlight:after{background:linear-gradient(90deg,#ffd21f,#f7b500,transparent)}
.section-title{font-size:30px;font-weight:1000;margin:34px 0 8px;letter-spacing:-.5px;display:flex;align-items:center;gap:10px}.section-title:before{content:"";width:7px;height:30px;border-radius:999px;background:linear-gradient(180deg,var(--blue),var(--cyan));box-shadow:0 0 18px rgba(0,170,255,.7)}.section-sub{color:var(--muted);margin-bottom:18px;font-size:15px}
.dre-scroll{overflow:auto;border:1px solid rgba(0,170,255,.32);border-radius:22px;max-height:720px;background:rgba(5,13,25,.88);box-shadow:0 22px 70px rgba(0,0,0,.38),inset 0 1px 0 rgba(255,255,255,.04)}.dre-table{border-collapse:separate;border-spacing:0;width:max-content;min-width:100%;font-size:15px}.dre-table th{position:sticky;top:0;z-index:5;background:linear-gradient(180deg,#172132 0%,#101827 100%);color:#d9e4f2;padding:15px 14px;border-bottom:1px solid #233b58;border-right:1px solid #23324a;text-align:right;font-weight:900}.dre-table th.col-linha{left:0;z-index:7;text-align:left;min-width:440px}.dre-table td{padding:14px;border-bottom:1px solid #1d304a;border-right:1px solid #1d304a;white-space:nowrap;background:#081426;color:#f8fbff;font-weight:650}.dre-table td.num{text-align:right;font-variant-numeric:tabular-nums}.dre-table td.pct{font-weight:800}.dre-table td.col-linha{position:sticky;left:0;z-index:3;background:#081426;text-align:left;min-width:440px;font-weight:750}.dre-table tr:hover td{filter:brightness(1.08)}
.linha-azul td{background:linear-gradient(90deg,#d9ecff 0%,#cfe6fb 100%);color:#00142a;font-weight:1000}.linha-azul td.col-linha{background:#d9ecff;color:#00142a;font-weight:1000}.linha-amarela td{background:linear-gradient(90deg,#ffd21f 0%,#ffc400 100%);color:#06101e;font-weight:1000}.linha-amarela td.col-linha{background:#ffd21f;color:#06101e;font-weight:1000}.linha-grupo td{background:linear-gradient(90deg,#eef3f8 0%,#e4edf7 100%);color:#0d1827;font-weight:1000;font-style:italic}.linha-grupo td.col-linha{background:#eef3f8;color:#0d1827;font-weight:1000;font-style:italic}
.alerta{background:#3b1d2a;color:#ff6976;padding:20px 24px;border-radius:16px;border:1px solid #713146;font-weight:900}.ok-box{background:#083e25;border:1px solid #127a45;color:#4fff93;padding:14px 16px;border-radius:14px;font-weight:900}.info-box{background:#0d2b45;border:1px solid #1e5d8c;color:#d8ecff;padding:14px 16px;border-radius:14px;font-weight:800}
.stButton>button{border-radius:14px;border:1px solid rgba(91,215,255,.35);background:linear-gradient(135deg,#10233b,#0c1b2e);color:#fff;font-weight:900;height:48px}.stButton>button:hover{border-color:#5bd7ff;box-shadow:0 0 18px rgba(0,170,255,.25)}
@media(max-width:1200px){.card-grid,.card-grid-5{grid-template-columns:repeat(2,minmax(0,1fr))}.card-grid-3{grid-template-columns:1fr}.hero-premium h1{font-size:40px}.dre-table th.col-linha,.dre-table td.col-linha{min-width:330px}}
</style>
""", unsafe_allow_html=True)

# =========================================================
# SIDEBAR E CARGA
# =========================================================
with st.sidebar:
    st.markdown("### ⚙️ Atualização")
    if st.button("🔄 Atualizar Base", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    st.divider()
    st.markdown("### 📁 Base monitorada")
    caminho_base = procurar_arquivo_base()
    if caminho_base:
        st.markdown("<div class='ok-box'>DRE_Consolidado_Moderno.xlsx</div>", unsafe_allow_html=True)
        st.caption(str(caminho_base))
    else:
        st.markdown("<div class='alerta'>Arquivo não encontrado.</div>", unsafe_allow_html=True)

logo64 = logo_base64()
logo_html = f"<img src='data:image/png;base64,{logo64}' />" if logo64 else "<div style='color:#32b7ff;font-weight:900;font-size:22px'>EIROX</div>"
st.markdown(f"""
<div class='hero-premium'>
    {logo_html}
    <div class='premium-badge'>Premium Financial Intelligence</div>
    <h1>DRE Empresa Online</h1>
    <p>Dashboard financeiro executivo • Página CEO • DRE gerencial • Visual Premium Eirox</p>
</div>
""", unsafe_allow_html=True)

if not caminho_base:
    st.stop()
try:
    dados, resumo_loja, nao_class, checks = carregar_dados(str(caminho_base))
except Exception as e:
    st.markdown(f"<div class='alerta'>Erro ao ler a base: {e}</div>", unsafe_allow_html=True); st.stop()

all_months = sorted(dados["Mês"].dropna().unique().tolist(), key=mes_key)
closed_default = meses_fechados_ano_atual(all_months) or all_months

with st.sidebar:
    st.divider()
    st.markdown("### 🧭 Navegação")
    pagina = st.radio("Tela", ["👔 Painel CEO", "📊 DRE Gerencial", "📈 Evolução Mensal", "🏪 Resultado por Loja", "⚠️ Auditoria DRE"], label_visibility="collapsed")
    st.divider()
    st.markdown("### 📅 Filtros")
    selected_months = st.multiselect("Meses", options=all_months, default=closed_default, help="Por padrão são selecionados somente os meses fechados do ano atual.")

if not selected_months:
    st.markdown("<div class='alerta'>Selecione pelo menos um mês.</div>", unsafe_allow_html=True); st.stop()
selected_months = sorted(selected_months, key=mes_key)
ultimo_mes = selected_months[-1]
mes_anterior = selected_months[-2] if len(selected_months) >= 2 else None

st.markdown(f"<div class='info-box'><b>Base ativa:</b> DADOS_DRE • <b>Meses:</b> {', '.join(selected_months)} • <b>Último mês:</b> {ultimo_mes}</div>", unsafe_allow_html=True)

# Valores base
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

if mes_anterior:
    ant = {
        "receita_bruta": valor_linha(dados, "1. RECEITA OPERACIONAL BRUTA", mes_anterior),
        "receita_liq": valor_linha(dados, "3. (=) RECEITA OPERACIONAL LÍQUIDA", mes_anterior),
        "lucro_bruto": valor_linha(dados, "5. (=) LUCRO BRUTO", mes_anterior),
        "ebitda": valor_linha(dados, "EBITDA", mes_anterior),
        "lucro_liq": valor_linha(dados, "11. (=) LUCRO LÍQUIDO", mes_anterior),
    }
    ant["margem_bruta"] = ant["lucro_bruto"] / ant["receita_bruta"] if ant["receita_bruta"] else 0
    ant["margem_ebitda"] = ant["ebitda"] / ant["receita_bruta"] if ant["receita_bruta"] else 0
    ant["margem_liq"] = ant["lucro_liq"] / ant["receita_bruta"] if ant["receita_bruta"] else 0
else:
    ant = {"receita_bruta":0,"receita_liq":0,"lucro_bruto":0,"ebitda":0,"lucro_liq":0,"margem_bruta":0,"margem_ebitda":0,"margem_liq":0}

# =========================================================
# PÁGINAS
# =========================================================
if pagina == "👔 Painel CEO":
    st.markdown(f"<div class='section-title'>Painel CEO — {ultimo_mes}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='section-sub'>Indicadores executivos com comparação automática contra {mes_anterior or 'mês anterior'}.</div>", unsafe_allow_html=True)
    st.markdown(f"""
    <div class='card-grid-5'>
        <div class='kpi-card ceo-card ceo-highlight'><div class='kpi-title'>Receita Bruta</div><div class='kpi-value'>{moeda_br(receita_bruta)}</div><div class='kpi-sub {cor_delta(delta_abs(receita_bruta, ant['receita_bruta']))}'>{texto_delta(receita_bruta, ant['receita_bruta'])}</div></div>
        <div class='kpi-card ceo-card'><div class='kpi-title'>Receita Líquida</div><div class='kpi-value'>{moeda_br(receita_liq)}</div><div class='kpi-sub {cor_delta(delta_abs(receita_liq, ant['receita_liq']))}'>{texto_delta(receita_liq, ant['receita_liq'])}</div></div>
        <div class='kpi-card ceo-card ceo-highlight'><div class='kpi-title'>Lucro Bruto</div><div class='kpi-value'>{moeda_br(lucro_bruto)}</div><div class='kpi-sub {cor_delta(delta_abs(lucro_bruto, ant['lucro_bruto']))}'>{texto_delta(lucro_bruto, ant['lucro_bruto'])}</div></div>
        <div class='kpi-card ceo-card ceo-highlight'><div class='kpi-title'>EBITDA</div><div class='kpi-value'>{moeda_br(ebitda)}</div><div class='kpi-sub {cor_delta(delta_abs(ebitda, ant['ebitda']))}'>{texto_delta(ebitda, ant['ebitda'])}</div></div>
        <div class='kpi-card ceo-card ceo-highlight'><div class='kpi-title'>Lucro Líquido</div><div class='kpi-value'>{moeda_br(lucro_liq)}</div><div class='kpi-sub {cor_delta(delta_abs(lucro_liq, ant['lucro_liq']))}'>{texto_delta(lucro_liq, ant['lucro_liq'])}</div></div>
    </div>
    <div class='card-grid-3'>
        <div class='kpi-card'><div class='kpi-title'>Margem Bruta</div><div class='kpi-value'>{pct_br(margem_bruta)}</div><div class='kpi-sub {cor_delta(delta_abs(margem_bruta, ant['margem_bruta']))}'>{texto_delta(margem_bruta, ant['margem_bruta'], is_pct=True)}</div></div>
        <div class='kpi-card'><div class='kpi-title'>Margem EBITDA</div><div class='kpi-value'>{pct_br(margem_ebitda)}</div><div class='kpi-sub {cor_delta(delta_abs(margem_ebitda, ant['margem_ebitda']))}'>{texto_delta(margem_ebitda, ant['margem_ebitda'], is_pct=True)}</div></div>
        <div class='kpi-card'><div class='kpi-title'>Margem Líquida</div><div class='kpi-value'>{pct_br(margem_liq)}</div><div class='kpi-sub {cor_delta(delta_abs(margem_liq, ant['margem_liq']))}'>{texto_delta(margem_liq, ant['margem_liq'], is_pct=True)}</div></div>
    </div>
    <div class='card-grid'>
        <div class='kpi-card'><div class='kpi-title'>CMV</div><div class='kpi-value'>{pct_br(cmv_pct)}</div><div class='kpi-sub'>{moeda_br(cmv)}</div></div>
        <div class='kpi-card'><div class='kpi-title'>Despesas Operacionais</div><div class='kpi-value'>{pct_br(desp_pct)}</div><div class='kpi-sub'>{moeda_br(desp_op)}</div></div>
        <div class='kpi-card'><div class='kpi-title'>LAIR</div><div class='kpi-value'>{moeda_br(lair)}</div><div class='kpi-sub'>Resultado antes dos tributos</div></div>
        <div class='kpi-card'><div class='kpi-title'>Mês analisado</div><div class='kpi-value'>{ultimo_mes}</div><div class='kpi-sub'>Comparado com {mes_anterior or '-'}</div></div>
    </div>
    """, unsafe_allow_html=True)

    if PLOTLY_OK:
        st.markdown("<div class='section-title'>Tendência Executiva</div>", unsafe_allow_html=True)
        evol = pd.concat([
            gerar_serie(dados, "1. RECEITA OPERACIONAL BRUTA", selected_months).assign(Indicador="Receita Bruta"),
            gerar_serie(dados, "EBITDA", selected_months).assign(Indicador="EBITDA"),
            gerar_serie(dados, "11. (=) LUCRO LÍQUIDO", selected_months).assign(Indicador="Lucro Líquido"),
        ], ignore_index=True)
        fig = px.line(evol, x="Mês", y="Valor", color="Indicador", markers=True, title="Receita, EBITDA e Lucro Líquido")
        fig.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=430)
        st.plotly_chart(fig, use_container_width=True)

elif pagina == "📊 DRE Gerencial":
    st.markdown("<div class='section-title'>DRE Gerencial</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-sub'>Modelo mantido no formato aprovado: meses em colunas, valores e percentuais lado a lado, com blocos estruturais destacados.</div>", unsafe_allow_html=True)
    st.markdown(construir_tabela_html(dados, selected_months), unsafe_allow_html=True)

elif pagina == "📈 Evolução Mensal":
    st.markdown("<div class='section-title'>Evolução Mensal</div>", unsafe_allow_html=True)
    if PLOTLY_OK:
        evol = pd.concat([
            gerar_serie(dados, "1. RECEITA OPERACIONAL BRUTA", selected_months).assign(Indicador="Receita Bruta"),
            gerar_serie(dados, "5. (=) LUCRO BRUTO", selected_months).assign(Indicador="Lucro Bruto"),
            gerar_serie(dados, "EBITDA", selected_months).assign(Indicador="EBITDA"),
            gerar_serie(dados, "11. (=) LUCRO LÍQUIDO", selected_months).assign(Indicador="Lucro Líquido"),
        ], ignore_index=True)
        fig = px.line(evol, x="Mês", y="Valor", color="Indicador", markers=True, title="Indicadores em R$")
        fig.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=430)
        st.plotly_chart(fig, use_container_width=True)

        margens = pd.DataFrame({
            "Mês": selected_months,
            "Margem Bruta": [pct_linha(dados, "5. (=) LUCRO BRUTO", m) * 100 for m in selected_months],
            "Margem EBITDA": [pct_linha(dados, "EBITDA", m) * 100 for m in selected_months],
            "Margem Líquida": [pct_linha(dados, "11. (=) LUCRO LÍQUIDO", m) * 100 for m in selected_months],
        }).melt(id_vars="Mês", var_name="Indicador", value_name="Percentual")
        fig2 = px.line(margens, x="Mês", y="Percentual", color="Indicador", markers=True, title="Margens (%)")
        fig2.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=430, yaxis_ticksuffix="%")
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Plotly não está instalado. Instale com: pip install plotly")

elif pagina == "🏪 Resultado por Loja":
    st.markdown("<div class='section-title'>Resultado por Loja</div>", unsafe_allow_html=True)
    if not resumo_loja.empty and {"Loja", "Mês", "Receita", "Despesas"}.issubset(set(resumo_loja.columns)):
        base_loja = resumo_loja[resumo_loja["Mês"].isin(selected_months)].copy()
        for col in ["Receita", "Despesas", "Resultado Caixa Simplificado"]:
            if col in base_loja.columns:
                base_loja[col] = pd.to_numeric(base_loja[col], errors="coerce").fillna(0)
        agg_cols = {"Receita":"sum", "Despesas":"sum"}
        if "Resultado Caixa Simplificado" in base_loja.columns:
            agg_cols["Resultado Caixa Simplificado"] = "sum"
        resumo = base_loja.groupby("Loja", as_index=False).agg(agg_cols)
        if "Resultado Caixa Simplificado" not in resumo.columns:
            resumo["Resultado Caixa Simplificado"] = resumo["Receita"] - resumo["Despesas"]
        resumo["Margem Caixa"] = resumo["Resultado Caixa Simplificado"] / resumo["Receita"].replace(0, pd.NA)
        resumo_fmt = resumo.copy()
        for col in ["Receita", "Despesas", "Resultado Caixa Simplificado"]:
            resumo_fmt[col] = resumo_fmt[col].apply(moeda_br)
        resumo_fmt["Margem Caixa"] = resumo_fmt["Margem Caixa"].fillna(0).apply(pct_br)
        st.dataframe(resumo_fmt, use_container_width=True, hide_index=True)
        if PLOTLY_OK and not resumo.empty:
            fig = px.bar(resumo, x="Loja", y="Resultado Caixa Simplificado", title="Ranking de Resultado por Loja")
            fig.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=420)
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("A aba RESUMO_LOJA não foi encontrada ou não possui as colunas necessárias.")

elif pagina == "⚠️ Auditoria DRE":
    st.markdown("<div class='section-title'>Auditoria DRE — Não Classificados</div>", unsafe_allow_html=True)
    if nao_class.empty:
        st.success("Nenhuma conta não classificada encontrada na base.")
    else:
        base_nc = nao_class.copy()
        if "Mês" in base_nc.columns:
            base_nc = base_nc[base_nc["Mês"].isin(selected_months)]
        valor_nc = pd.to_numeric(base_nc.get("Valor", 0), errors="coerce").fillna(0).sum() if not base_nc.empty else 0
        qtde_nc = pd.to_numeric(base_nc.get("Qtde", 0), errors="coerce").fillna(0).sum() if not base_nc.empty else 0
        st.markdown(f"<div class='info-box'>Valor não classificado no filtro: <b>{moeda_br(valor_nc)}</b> • Quantidade: <b>{int(qtde_nc)}</b></div>", unsafe_allow_html=True)
        st.dataframe(base_nc, use_container_width=True, hide_index=True)
    with st.expander("✅ Checks da Base", expanded=False):
        if checks.empty:
            st.info("A aba CHECKS não foi encontrada.")
        else:
            st.dataframe(checks, use_container_width=True, hide_index=True)
