# -*- coding: utf-8 -*-
"""
Eirox DRE Online - Atualização automática pelas pastas locais
Execute na pasta C:\\Users\\Comercial\\Desktop\\Dre com:
    streamlit run app.py
"""

from __future__ import annotations

import base64
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st

# =========================
# CONFIGURAÇÃO GERAL
# =========================
st.set_page_config(
    page_title="Eirox DRE Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

PASTAS_OBRIGATORIAS = [
    "CONTAS A PAGAR - DRE",
    "PLANO DE CONTAS - LEGENDA",
    "POSIÇÃO DE ESTOQUE",
    "VENDA POR PAGAMENTO",
]

MESES_PT = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
}
MESES_NOME = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
}

# =========================
# ESTILO
# =========================
st.markdown(
    """
    <style>
    .stApp { background: #07111f; color: #f8fafc; }
    [data-testid="stSidebar"] { background: #050b14; border-right: 1px solid #12375b; }
    .main .block-container { padding-top: 1.4rem; max-width: 1500px; }
    h1, h2, h3 { color: #f8fafc !important; font-weight: 800 !important; }
    .subtitulo { color:#93c5fd; font-size:18px; text-align:center; margin-top:-10px; }
    .card {
        background: linear-gradient(180deg,#101b2c,#0a1526);
        border:1px solid #1f3b5c; border-radius:18px; padding:18px;
        box-shadow: 0 10px 28px rgba(0,0,0,.22);
    }
    .kpi-title { color:#94a3b8; font-size:14px; font-weight:700; text-transform:uppercase; }
    .kpi-value { color:#ffffff; font-size:27px; font-weight:900; margin-top:5px; }
    .kpi-sub { color:#38bdf8; font-size:13px; margin-top:4px; }
    .ok { color:#22c55e; font-weight:800; }
    .bad { color:#fb7185; font-weight:800; }
    .dre-wrap { overflow-x:auto; border:1px solid #173554; border-radius:16px; background:#081426; }
    table.dre { border-collapse:collapse; width:100%; min-width:1300px; font-size:15px; }
    table.dre th { position:sticky; top:0; background:#171b24; color:#cbd5e1; padding:12px 10px; border:1px solid #263244; text-align:left; white-space:nowrap; font-weight:700; }
    table.dre td { padding:11px 10px; border:1px solid #203149; white-space:nowrap; color:#f8fafc; }
    table.dre td.num { text-align:right; font-variant-numeric: tabular-nums; }
    table.dre tr.item { background:#0b1729; }
    table.dre tr.bloco { background:#cfe0f3; }
    table.dre tr.bloco td { color:#07111f; font-weight:900; }
    table.dre tr.resultado { background:#ffd21a; }
    table.dre tr.resultado td { color:#05111f; font-weight:950; }
    table.dre tr.grupo { background:#263241; }
    table.dre tr.grupo td { color:#ffffff; font-weight:900; font-style:italic; }
    table.dre tr:hover { filter: brightness(1.08); }
    .footer-note { color:#94a3b8; font-size:13px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================
# FUNÇÕES BASE
# =========================
def encontrar_pasta_base() -> Path:
    candidatos = [
        Path.cwd(),
        Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd(),
        Path(__file__).resolve().parent.parent if "__file__" in globals() else Path.cwd().parent,
        Path.home() / "Desktop" / "Dre",
        Path.home() / "Área de Trabalho" / "Dre",
        Path.home() / "Area de Trabalho" / "Dre",
        Path(r"C:\Users\Comercial\Desktop\Dre"),
    ]
    for p in candidatos:
        if all((p / pasta).exists() for pasta in PASTAS_OBRIGATORIAS):
            return p
    return Path.cwd()

PASTA_BASE = encontrar_pasta_base()


def normalize(txt) -> str:
    if pd.isna(txt):
        return ""
    s = str(txt).strip().upper()
    mapa = str.maketrans("ÁÀÂÃÉÊÍÓÔÕÚÜÇ", "AAAAEEIOOOUUC")
    s = s.translate(mapa)
    s = re.sub(r"\s+", " ", s)
    return s


def fmt_brl(v) -> str:
    try:
        v = float(v)
    except Exception:
        v = 0.0
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v) -> str:
    try:
        v = float(v)
    except Exception:
        v = 0.0
    return f"{v*100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def limpar_valor(x) -> float:
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace("R$", "").replace("%", "")
    if s in ["-", "", "nan", "None"]:
        return 0.0
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def ler_excel_pasta(pasta: Path) -> pd.DataFrame:
    if not pasta.exists():
        return pd.DataFrame()
    arquivos = [a for a in list(pasta.rglob("*.xlsx")) + list(pasta.rglob("*.xls")) if not a.name.startswith("~$")]
    bases = []
    for arq in arquivos:
        try:
            df = pd.read_excel(arq)
            df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
            df["Arquivo Origem"] = arq.name
            df["Caminho Origem"] = str(arq)
            bases.append(df)
        except Exception as e:
            st.sidebar.warning(f"Falha ao ler {arq.name}: {e}")
    return pd.concat(bases, ignore_index=True) if bases else pd.DataFrame()


def localizar_coluna(df: pd.DataFrame, opcoes: List[str]) -> str | None:
    if df.empty:
        return None
    norm_cols = {normalize(c): c for c in df.columns}
    for op in opcoes:
        n = normalize(op)
        if n in norm_cols:
            return norm_cols[n]
    for c in df.columns:
        nc = normalize(c)
        if any(normalize(op) in nc for op in opcoes):
            return c
    return None


def mes_from_filename(nome: str, ano_padrao: int = 2026) -> str | None:
    n = normalize(nome)
    for mes_nome, mes_num in MESES_NOME.items():
        if mes_nome in n:
            return f"{MESES_PT[mes_num]}/{str(ano_padrao)[-2:]}"
    m = re.search(r"(20\d{2})[-_ ]?(0[1-9]|1[0-2])", n)
    if m:
        ano = int(m.group(1)); mes = int(m.group(2))
        return f"{MESES_PT[mes]}/{str(ano)[-2:]}"
    return None


def detectar_mes(df: pd.DataFrame, arquivo_col="Arquivo Origem") -> pd.Series:
    # tenta colunas explícitas de mês/data
    col_mes = localizar_coluna(df, ["Ano-mês", "Ano Mes", "Mes", "Mês", "Mês Referência", "Mes Referencia", "Competência", "Competencia"])
    if col_mes:
        def conv(v):
            if pd.isna(v): return None
            s = str(v).strip()
            ns = normalize(s)
            # jan/26
            m = re.search(r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[A-Z]*[/-]?(\d{2,4})", ns)
            if m:
                mes_txt = m.group(1)[:3].lower()
                ano = m.group(2)[-2:]
                if mes_txt == "jun": mes_txt = "jun"
                return f"{mes_txt}/{ano}"
            dt = pd.to_datetime(v, errors="coerce", dayfirst=True)
            if pd.notna(dt):
                return f"{MESES_PT[int(dt.month)]}/{str(int(dt.year))[-2:]}"
            return None
        serie = df[col_mes].apply(conv)
    else:
        serie = pd.Series([None] * len(df), index=df.index)

    # fallback pelo nome do arquivo
    if arquivo_col in df.columns:
        arquivo_mes = df[arquivo_col].apply(lambda x: mes_from_filename(str(x)))
        serie = serie.fillna(arquivo_mes)
    return serie


def mes_ordem(mes: str) -> int:
    if not mes or "/" not in str(mes):
        return 999999
    mtxt, ano = str(mes).split("/")[:2]
    rev = {v: k for k, v in MESES_PT.items()}
    return 2000 + int(ano) * 100 + rev.get(mtxt.lower()[:3], 99)


def meses_fechados(meses: List[str]) -> List[str]:
    hoje = datetime.today()
    atual = f"{MESES_PT[hoje.month]}/{str(hoje.year)[-2:]}"
    return [m for m in sorted(set(meses), key=mes_ordem) if m != atual]

# =========================
# CLASSIFICAÇÃO
# =========================
MAPEAMENTO = {
    "ICMS": ["ICMS"],
    "DAS": ["DAS"],
    "ISSQN": ["ISSQN"],
    "Mercadorias (CMP)": ["MERCADORIAS (CMP)", "MERCADORIA", "CMP", "CMV"],
    "Alimentação": ["ALIMENTACAO"],
    "Pagamento de Férias": ["PAGAMENTO DE FERIAS", "FERIAS"],
    "FGTS": ["FGTS"],
    "Salários Fixos + Horas Extras": ["SALARIOS FIXOS", "HORAS EXTRAS", "SALARIO"],
    "Comissões e premiações": ["COMISSOES", "COMISSAO", "PREMIACOES", "PREMIACAO"],
    "INSS": ["INSS"],
    "IRRF Folha": ["IRRF FOLHA", "IRRF"],
    "Cursos, Treinamentos, Viagens": ["CURSOS", "TREINAMENTOS", "VIAGENS"],
    "PPRA<PCMSO e Exames": ["PPRA", "PCMSO", "EXAMES"],
    "Rescisão (guia)": ["RESCISAO"],
    "Provisão 13 e Férias": ["PROVISAO 13", "PROVISAO FERIAS"],
    "Uniformes": ["UNIFORMES"],
    "Vale Transporte": ["VALE TRANSPORTE"],
    "Pro-Labore": ["PRO-LABORE", "PRO LABORE"],
    "Outras Despesas com Pessoal": ["OUTRAS DESPESAS COM PESSOAL"],
    "Aluguel&IPTU": ["ALUGUEL", "IPTU"],
    "Seguro Imóvel": ["SEGURO IMOVEL"],
    "Agua/Luz/Fone/Net": ["AGUA", "LUZ", "FONE", "NET", "ENERGIA", "INTERNET"],
    "Combustível": ["COMBUSTIVEL"],
    "Manutenção em Geral": ["MANUTENCAO EM GERAL"],
    "Manutenção Veículos/Motos": ["MANUTENCAO VEICULOS", "MANUTENCAO MOTOS"],
    "Mat. de Escritório/Informática": ["ESCRITORIO", "INFORMATICA"],
    "Mat. De Limpeza": ["LIMPEZA"],
    "Viagens": ["VIAGENS"],
    "Contábil (Terceiros)": ["CONTABIL"],
    "Sistemas (Terceiros)": ["SISTEMAS"],
    "Juridico (Terceiros)": ["JURIDICO"],
    "Assessorias/Consultorias/Treinamentos": ["ASSESSORIA", "CONSULTORIA", "TREINAMENTO"],
    "Taxas, Licenças e Contrib.": ["TAXAS", "LICENCAS", "CONTRIB"],
    "Outros (Terceiros)": ["OUTROS TERCEIROS"],
    "Outros (Despesas)": ["OUTROS DESPESAS", "OUTRAS DESPESAS"],
    "Quebra de caixa": ["QUEBRA DE CAIXA"],
    "Marketing/Publicidade": ["MARKETING", "PUBLICIDADE"],
    "Sistema Fidelidade": ["SISTEMA FIDELIDADE", "FIDELIDADE"],
    "Frete": ["FRETE"],
    "Associação de Classe (Royalties)": ["ASSOCIACAO DE CLASSE", "ROYALTIES"],
    "Juros Boletos": ["JUROS BOLETOS", "JUROS"],
    "Tarifas Bancarias": ["TARIFAS BANCARIAS", "TARIFA"],
    "Taxas de Cartão (MDR)": ["TAXAS DE CARTAO", "MDR", "CARTAO"],
    "IRPJ/CSLL": ["IRPJ", "CSLL"],
}


def classificar_plano(plano: str) -> str:
    t = normalize(plano)
    if not t or t in ["NAN", "NONE"]:
        return "NAO_CLASSIFICADO"
    for linha, palavras in MAPEAMENTO.items():
        if any(normalize(p) in t for p in palavras):
            return linha
    return "NAO_CLASSIFICADO"

# =========================
# PROCESSAMENTO DINÂMICO
# =========================
@st.cache_data(show_spinner=False, ttl=0)
def processar_dre(timestamp: float) -> Dict[str, pd.DataFrame]:
    base = encontrar_pasta_base()
    contas = ler_excel_pasta(base / "CONTAS A PAGAR - DRE")
    plano_legenda = ler_excel_pasta(base / "PLANO DE CONTAS - LEGENDA")
    estoque = ler_excel_pasta(base / "POSIÇÃO DE ESTOQUE")
    vendas = ler_excel_pasta(base / "VENDA POR PAGAMENTO")

    # VENDAS
    if not vendas.empty:
        col_total = localizar_coluna(vendas, ["Total", "Valor", "Valor Total", "Receita", "Venda"])
        vendas["Mês"] = detectar_mes(vendas)
        vendas["Valor Receita"] = vendas[col_total].apply(limpar_valor) if col_total else 0.0
    else:
        vendas = pd.DataFrame(columns=["Mês", "Valor Receita"])

    # CONTAS
    if not contas.empty:
        col_valor = localizar_coluna(contas, ["Valor Documento", "Valor", "Total", "Valor Total"])
        col_plano = localizar_coluna(contas, ["Plano de Contas", "Plano Contas", "Conta", "Classificação"])
        col_loja = localizar_coluna(contas, ["Unidade", "Loja", "Apelido Un. Neg.", "Un. Neg."])
        contas["Mês"] = detectar_mes(contas)
        contas["Valor Documento"] = contas[col_valor].apply(limpar_valor) if col_valor else 0.0
        contas["Plano de Contas"] = contas[col_plano].astype(str) if col_plano else ""
        contas["Loja"] = contas[col_loja].astype(str) if col_loja else "Todas"
        contas["Linha Classificada"] = contas["Plano de Contas"].apply(classificar_plano)
        contas_class = contas[contas["Linha Classificada"] != "NAO_CLASSIFICADO"].copy()
        nao_class = contas[contas["Linha Classificada"] == "NAO_CLASSIFICADO"].copy()
    else:
        contas = pd.DataFrame(columns=["Mês", "Valor Documento", "Plano de Contas", "Loja", "Linha Classificada"])
        contas_class = contas.copy()
        nao_class = contas.copy()

    # ESTOQUE
    if not estoque.empty:
        estoque["Mês"] = detectar_mes(estoque)
        col_est = localizar_coluna(estoque, ["Estoque X Custo Médio", "Estoque X Custo Medio", "Valor Estoque", "Estoque a Custo"])
        estoque["Valor Estoque"] = estoque[col_est].apply(limpar_valor) if col_est else 0.0
    else:
        estoque = pd.DataFrame(columns=["Mês", "Valor Estoque"])

    meses = meses_fechados(list(vendas["Mês"].dropna().unique()) + list(contas["Mês"].dropna().unique()) + list(estoque["Mês"].dropna().unique()))

    def receita(m): return vendas.loc[vendas["Mês"] == m, "Valor Receita"].sum()
    def soma_linha(m, linha): return contas_class.loc[(contas_class["Mês"] == m) & (contas_class["Linha Classificada"] == linha), "Valor Documento"].sum()
    def soma_lista(m, linhas): return sum(soma_linha(m, l) for l in linhas)

    pessoal = ["Alimentação", "Pagamento de Férias", "FGTS", "Salários Fixos + Horas Extras", "Comissões e premiações", "INSS", "IRRF Folha", "Cursos, Treinamentos, Viagens", "PPRA<PCMSO e Exames", "Rescisão (guia)", "Provisão 13 e Férias", "Uniformes", "Vale Transporte", "Pro-Labore", "Outras Despesas com Pessoal"]
    admin = ["Aluguel&IPTU", "Seguro Imóvel", "Agua/Luz/Fone/Net", "Combustível", "Manutenção em Geral", "Manutenção Veículos/Motos", "Mat. de Escritório/Informática", "Mat. De Limpeza", "Viagens", "Contábil (Terceiros)", "Sistemas (Terceiros)", "Juridico (Terceiros)", "Assessorias/Consultorias/Treinamentos", "Taxas, Licenças e Contrib.", "Outros (Terceiros)", "Outros (Despesas)", "Quebra de caixa"]
    vendas_mkt = ["Marketing/Publicidade", "Sistema Fidelidade", "Frete", "Associação de Classe (Royalties)"]
    financeiro = ["Juros Boletos", "Tarifas Bancarias", "Taxas de Cartão (MDR)"]

    linhas_modelo = [
        ("1. RECEITA OPERACIONAL BRUTA", "bloco", lambda m: receita(m)),
        ("Receita de Vendas de Mercadorias", "item", lambda m: receita(m)),
        ("2. (-) DEDUÇÕES DA RECEITA BRUTA", "bloco", lambda m: soma_lista(m, ["ICMS", "DAS", "ISSQN"])),
        ("ICMS", "item", lambda m: soma_linha(m, "ICMS")),
        ("DAS", "item", lambda m: soma_linha(m, "DAS")),
        ("ISSQN", "item", lambda m: soma_linha(m, "ISSQN")),
        ("3. (=) RECEITA OPERACIONAL LÍQUIDA", "resultado", lambda m: receita(m) - soma_lista(m, ["ICMS", "DAS", "ISSQN"])),
        ("4. (-) CUSTOS DAS VENDAS", "bloco", lambda m: soma_linha(m, "Mercadorias (CMP)")),
        ("Mercadorias (CMP)", "item", lambda m: soma_linha(m, "Mercadorias (CMP)")),
        ("5. (=) LUCRO BRUTO", "resultado", lambda m: (receita(m) - soma_lista(m, ["ICMS", "DAS", "ISSQN"])) - soma_linha(m, "Mercadorias (CMP)")),
        ("6. (-) DESPESAS OPERACIONAIS", "bloco", lambda m: soma_lista(m, pessoal + admin + vendas_mkt)),
        ("Despesas com Pessoal", "grupo", lambda m: soma_lista(m, pessoal)),
    ]
    linhas_modelo += [(x, "item", lambda m, x=x: soma_linha(m, x)) for x in pessoal]
    linhas_modelo += [("Despesas Administrativas e Ocupação", "grupo", lambda m: soma_lista(m, admin))]
    linhas_modelo += [(x, "item", lambda m, x=x: soma_linha(m, x)) for x in admin]
    linhas_modelo += [("Despesas com Vendas e Marketing", "grupo", lambda m: soma_lista(m, vendas_mkt))]
    linhas_modelo += [(x, "item", lambda m, x=x: soma_linha(m, x)) for x in vendas_mkt]
    linhas_modelo += [
        ("7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO (EBITDA/LAJIDA)", "resultado", lambda m: ((receita(m) - soma_lista(m, ["ICMS", "DAS", "ISSQN"])) - soma_linha(m, "Mercadorias (CMP)")) - soma_lista(m, pessoal + admin + vendas_mkt)),
        ("8. (+/-) RESULTADO FINANCEIRO LÍQUIDO", "bloco", lambda m: soma_lista(m, financeiro)),
    ]
    linhas_modelo += [(x, "item", lambda m, x=x: soma_linha(m, x)) for x in financeiro]
    linhas_modelo += [
        ("9. (=) RESULTADO ANTES DOS TRIBUTOS (LAIR)", "resultado", lambda m: (((receita(m) - soma_lista(m, ["ICMS", "DAS", "ISSQN"])) - soma_linha(m, "Mercadorias (CMP)")) - soma_lista(m, pessoal + admin + vendas_mkt)) - soma_lista(m, financeiro)),
        ("10. (-) TRIBUTOS SOBRE O LUCRO", "bloco", lambda m: soma_linha(m, "IRPJ/CSLL")),
        ("IRPJ/CSLL", "item", lambda m: soma_linha(m, "IRPJ/CSLL")),
        ("11. (=) LUCRO LÍQUIDO DO EXERCÍCIO", "resultado", lambda m: ((((receita(m) - soma_lista(m, ["ICMS", "DAS", "ISSQN"])) - soma_linha(m, "Mercadorias (CMP)")) - soma_lista(m, pessoal + admin + vendas_mkt)) - soma_lista(m, financeiro)) - soma_linha(m, "IRPJ/CSLL")),
    ]

    rows = []
    for linha, estilo, func in linhas_modelo:
        row = {"Linha DRE": linha, "Estilo": estilo}
        for m in meses:
            val = float(func(m))
            rec = receita(m)
            row[f"{m} Valor"] = val
            row[f"{m} %"] = (val / rec) if rec else 0.0
        rows.append(row)
    dre = pd.DataFrame(rows)

    # Série executiva
    long_rows = []
    for m in meses:
        rec = receita(m)
        ded = soma_lista(m, ["ICMS", "DAS", "ISSQN"])
        cmv = soma_linha(m, "Mercadorias (CMP)")
        lucro_bruto = rec - ded - cmv
        despesas = soma_lista(m, pessoal + admin + vendas_mkt)
        ebitda = lucro_bruto - despesas
        fin = soma_lista(m, financeiro)
        ir = soma_linha(m, "IRPJ/CSLL")
        lucro_liq = ebitda - fin - ir
        long_rows.extend([
            {"Mês": m, "Indicador": "Receita Bruta", "Valor": rec},
            {"Mês": m, "Indicador": "Lucro Bruto", "Valor": lucro_bruto},
            {"Mês": m, "Indicador": "EBITDA", "Valor": ebitda},
            {"Mês": m, "Indicador": "Lucro Líquido", "Valor": lucro_liq},
        ])
    indicadores = pd.DataFrame(long_rows)

    return {
        "dre": dre,
        "indicadores": indicadores,
        "contas": contas,
        "contas_classificadas": contas_class,
        "nao_classificados": nao_class,
        "vendas": vendas,
        "estoque": estoque,
        "plano": plano_legenda,
        "meses": pd.DataFrame({"Mês": meses}),
    }

# =========================
# HTML TABLE
# =========================
def render_dre_table(df: pd.DataFrame, meses: List[str]) -> str:
    cols = ["Linha DRE"]
    for m in meses:
        cols += [f"{m} Valor", f"{m} %"]
    html = '<div class="dre-wrap"><table class="dre"><thead><tr>'
    for c in cols:
        html += f"<th>{c}</th>"
    html += "</tr></thead><tbody>"
    for _, r in df.iterrows():
        estilo = r.get("Estilo", "item")
        html += f'<tr class="{estilo}">'
        html += f"<td>{r['Linha DRE']}</td>"
        for m in meses:
            html += f"<td class='num'>{fmt_brl(r.get(f'{m} Valor', 0))}</td>"
            html += f"<td class='num'>{fmt_pct(r.get(f'{m} %', 0))}</td>"
        html += "</tr>"
    html += "</tbody></table></div>"
    return html

# =========================
# APP
# =========================
def logo_html() -> str:
    for p in [PASTA_BASE / "assets" / "logo_eirox.png", PASTA_BASE / "assets" / "logo eirox(3).png", PASTA_BASE / "logo_eirox.png"]:
        if p.exists():
            data = base64.b64encode(p.read_bytes()).decode()
            return f'<img src="data:image/png;base64,{data}" style="height:44px; object-fit:contain;">'
    return "<b style='color:#38bdf8;font-size:28px'>EIROX</b>"

with st.sidebar:
    st.markdown("### ⚙️ Atualização")
    st.caption("As bases são lidas direto das pastas. Ao substituir arquivos, clique em Atualizar Base.")
    if st.button("🔄 Atualizar Base", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    st.markdown("### 📁 Pastas monitoradas")
    st.caption(str(PASTA_BASE))
    for pasta in PASTAS_OBRIGATORIAS:
        ok = (PASTA_BASE / pasta).exists()
        st.markdown(("✅" if ok else "❌") + f" {pasta}")

resultado = processar_dre(datetime.now().timestamp())
meses = resultado["meses"]["Mês"].tolist()

st.markdown(
    f"""
    <div style="display:flex; align-items:center; justify-content:center; gap:28px; margin-top:10px;">
        {logo_html()}
        <div style="font-size:52px; font-weight:950; color:#f8fafc;">DRE Empresa Online</div>
    </div>
    <div class="subtitulo">Dashboard financeiro gerencial no padrão Eirox Pricing Online • Atualização automática pelas pastas de base</div>
    """,
    unsafe_allow_html=True,
)

if not meses:
    st.error("Nenhum mês foi identificado nas bases. Verifique se as pastas de dados estão na mesma pasta do app ou em C:\\Users\\Comercial\\Desktop\\Dre.")
    st.stop()

# Filtro sempre inicia com meses fechados
meses_sel = st.sidebar.multiselect("Meses", options=meses, default=meses)
if not meses_sel:
    st.warning("Selecione pelo menos um mês.")
    st.stop()

# KPIs último mês selecionado
ultimo = sorted(meses_sel, key=mes_ordem)[-1]
dre = resultado["dre"]

def valor_linha(nome: str, mes: str) -> float:
    s = dre.loc[dre["Linha DRE"].eq(nome), f"{mes} Valor"]
    return float(s.iloc[0]) if len(s) else 0.0

def pct_linha(nome: str, mes: str) -> float:
    s = dre.loc[dre["Linha DRE"].eq(nome), f"{mes} %"]
    return float(s.iloc[0]) if len(s) else 0.0

st.markdown("## 📊 Painel Executivo")
kp1, kp2, kp3, kp4 = st.columns(4)
for col, titulo, linha in [
    (kp1, "Receita Bruta", "1. RECEITA OPERACIONAL BRUTA"),
    (kp2, "Lucro Bruto", "5. (=) LUCRO BRUTO"),
    (kp3, "EBITDA", "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO (EBITDA/LAJIDA)"),
    (kp4, "Lucro Líquido", "11. (=) LUCRO LÍQUIDO DO EXERCÍCIO"),
]:
    with col:
        st.markdown(f"<div class='card'><div class='kpi-title'>{titulo} • {ultimo}</div><div class='kpi-value'>{fmt_brl(valor_linha(linha, ultimo))}</div><div class='kpi-sub'>{fmt_pct(pct_linha(linha, ultimo))} da receita</div></div>", unsafe_allow_html=True)

st.markdown("## DRE Gerencial")
st.caption("Modelo mantido no formato da aba DRE original: meses em colunas, valores e percentuais lado a lado. Não classificados ficam fora dos cálculos principais.")
dre_filtrado_cols = ["Linha DRE", "Estilo"] + [c for m in meses_sel for c in [f"{m} Valor", f"{m} %"]]
st.markdown(render_dre_table(dre[dre_filtrado_cols], meses_sel), unsafe_allow_html=True)

st.markdown("## 📈 Evolução Mensal")
ind = resultado["indicadores"]
ind = ind[ind["Mês"].isin(meses_sel)].copy()
if not ind.empty:
    fig = px.line(ind, x="Mês", y="Valor", color="Indicador", markers=True)
    fig.update_layout(template="plotly_dark", paper_bgcolor="#07111f", plot_bgcolor="#07111f", font_color="#e5e7eb")
    st.plotly_chart(fig, use_container_width=True)

with st.expander("⚠️ Auditoria DRE - Não classificados", expanded=False):
    nc = resultado["nao_classificados"]
    valor_nc = nc["Valor Documento"].sum() if "Valor Documento" in nc.columns else 0
    st.markdown(f"**Registros não classificados:** {len(nc)}  |  **Valor:** {fmt_brl(valor_nc)}")
    if not nc.empty:
        mostrar = [c for c in ["Mês", "Loja", "Plano de Contas", "Valor Documento", "Arquivo Origem"] if c in nc.columns]
        st.dataframe(nc[mostrar], use_container_width=True, hide_index=True)

st.markdown("<div class='footer-note'>Eirox DRE Online • Base atualizada diretamente das pastas monitoradas.</div>", unsafe_allow_html=True)
