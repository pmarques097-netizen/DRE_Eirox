# -*- coding: utf-8 -*-
"""
Eirox DRE Online - Atualização Automática
- Mantém o modelo do DRE Gerencial com meses em colunas
- Lê automaticamente as bases nas pastas do projeto
- Ignora não classificados nos cálculos principais
- Mostra somente meses fechados por padrão
- Sem upload manual de arquivo

Estrutura esperada:
Dre/
├── app.py
├── assets/logo_eirox.png
├── CONTAS A PAGAR - DRE/
├── PLANO DE CONTAS - LEGENDA/
├── POSIÇÃO DE ESTOQUE/
└── VENDA POR PAGAMENTO/
"""

from __future__ import annotations

from pathlib import Path
from datetime import datetime
import re
import base64
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px

# =========================================================
# CONFIGURAÇÃO GERAL
# =========================================================
st.set_page_config(
    page_title="Eirox DRE Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = Path(__file__).resolve().parent
PASTA_BASE = APP_DIR

# Se o app estiver em uma subpasta, tenta localizar a pasta mãe do projeto
PASTAS_OBRIGATORIAS = [
    "CONTAS A PAGAR - DRE",
    "PLANO DE CONTAS - LEGENDA",
    "POSIÇÃO DE ESTOQUE",
    "VENDA POR PAGAMENTO",
]
if not all((PASTA_BASE / p).exists() for p in PASTAS_OBRIGATORIAS):
    if all((PASTA_BASE.parent / p).exists() for p in PASTAS_OBRIGATORIAS):
        PASTA_BASE = PASTA_BASE.parent

PASTA_CONTAS = PASTA_BASE / "CONTAS A PAGAR - DRE"
PASTA_PLANO = PASTA_BASE / "PLANO DE CONTAS - LEGENDA"
PASTA_ESTOQUE = PASTA_BASE / "POSIÇÃO DE ESTOQUE"
PASTA_VENDAS = PASTA_BASE / "VENDA POR PAGAMENTO"
PASTA_ASSETS = PASTA_BASE / "assets"
LOGO_PATHS = [
    PASTA_ASSETS / "logo_eirox.png",
    PASTA_ASSETS / "logo eirox.png",
    PASTA_BASE / "logo_eirox.png",
    PASTA_BASE / "logo eirox.png",
]

MESES_PT = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
}
MESES_NOME = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3, "ABRIL": 4,
    "MAIO": 5, "JUNHO": 6, "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9,
    "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}

# =========================================================
# CSS / TEMA
# =========================================================
st.markdown(
    """
<style>
:root {
  --bg: #07111f;
  --panel: #0d1a2e;
  --panel2: #111827;
  --border: #1f3a5f;
  --text: #f8fafc;
  --muted: #9ca3af;
  --blue: #00a6ff;
  --blue2: #0e78ff;
  --yellow: #ffd11a;
  --green: #22c55e;
  --red: #ef4444;
  --ice: #d8e9fb;
}
.stApp { background: radial-gradient(circle at top left, #0b2442 0, #07111f 35%, #050b13 100%); color: var(--text); }
[data-testid="stSidebar"] { background: #050b13; border-right: 1px solid #12385f; }
.block-container { padding-top: 1.2rem; max-width: 1500px; }
.eirox-header { display:flex; align-items:center; gap:28px; padding: 14px 4px 24px 4px; }
.eirox-logo { width: 170px; max-height: 74px; object-fit: contain; }
.eirox-title { font-size: 42px; font-weight: 900; line-height: 1.05; margin:0; color:#ffffff; letter-spacing:.3px; }
.eirox-subtitle { margin-top:8px; font-size: 17px; color:#9ecbff; }
.kpi-card { background: linear-gradient(145deg, rgba(13,26,46,.98), rgba(6,16,30,.96)); border:1px solid rgba(0,166,255,.24); border-radius:18px; padding:18px 18px; box-shadow: 0 10px 30px rgba(0,0,0,.34); min-height:122px; }
.kpi-label { color:#9ca3af; font-size:13px; font-weight:700; text-transform:uppercase; letter-spacing:.5px; }
.kpi-value { color:#ffffff; font-size:28px; font-weight:900; margin-top:8px; white-space:nowrap; }
.kpi-sub { color:#9ecbff; font-size:13px; margin-top:6px; }
.section-title { font-size:27px; font-weight:900; margin:28px 0 12px 0; color:#ffffff; }
.info-box { background: rgba(13,26,46,.75); border:1px solid rgba(0,166,255,.20); border-radius:14px; padding:14px 16px; color:#b8c7dc; }
.dre-wrap { overflow:auto; border:1px solid #143557; border-radius:16px; max-height: 720px; box-shadow: 0 12px 34px rgba(0,0,0,.28); }
table.dre-table { border-collapse: collapse; width: 100%; min-width: 1180px; font-size: 15px; background:#0b1628; }
.dre-table th { position: sticky; top:0; z-index:5; background:#151a24; color:#b9c4d0; padding:13px 10px; text-align:right; font-weight:700; border-bottom:2px solid #25364d; border-right:1px solid #26374d; white-space:nowrap; }
.dre-table th:first-child { left:0; z-index:6; text-align:left; min-width:360px; }
.dre-table td { padding:11px 10px; text-align:right; border-bottom:1px solid #1b2a3f; border-right:1px solid #243247; color:#f8fafc; white-space:nowrap; font-weight:600; }
.dre-table td:first-child { position:sticky; left:0; z-index:3; text-align:left; min-width:360px; background:#0b1628; }
.dre-table tr.item td { background:#0b1628; font-weight:500; }
.dre-table tr.item td:first-child { background:#0b1628; }
.dre-table tr.bloco td { background:#d8e9fb; color:#03101f; font-weight:900; }
.dre-table tr.bloco td:first-child { background:#d8e9fb; color:#03101f; }
.dre-table tr.resultado td { background:#ffd11a; color:#020617; font-weight:950; }
.dre-table tr.resultado td:first-child { background:#ffd11a; color:#020617; }
.dre-table tr.grupo td { background:#202736; color:#f8fafc; font-weight:900; font-style:italic; }
.dre-table tr.grupo td:first-child { background:#202736; color:#f8fafc; }
.dre-table tr.fluxo td { background:#17304d; color:#f8fafc; font-weight:900; }
.dre-table tr.fluxo td:first-child { background:#17304d; color:#f8fafc; }
.dre-table tr:hover td { filter:brightness(1.08); }
.small-muted { color:#9ca3af; font-size:13px; }
hr { border-color:#12385f; }
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================
def normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    txt = str(valor).strip().upper()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("ASCII")
    txt = re.sub(r"\s+", " ", txt)
    return txt


def dinheiro(v) -> str:
    try:
        v = float(v)
    except Exception:
        v = 0.0
    s = f"R$ {v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v) -> str:
    try:
        v = float(v)
    except Exception:
        v = 0.0
    return f"{v*100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def limpar_valor_serie(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype=float)
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    txt = s.astype(str).str.strip()
    txt = txt.str.replace("R$", "", regex=False)
    txt = txt.str.replace("%", "", regex=False)
    txt = txt.str.replace(" ", "", regex=False)
    # Trata formato BR: 1.234,56
    txt = txt.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    txt = txt.replace({"-": "0", "": "0", "nan": "0", "None": "0"})
    return pd.to_numeric(txt, errors="coerce").fillna(0.0)


def procurar_coluna(df: pd.DataFrame, opcoes: list[str]) -> str | None:
    if df.empty:
        return None
    mapa = {normalizar_texto(c): c for c in df.columns}
    for op in opcoes:
        opn = normalizar_texto(op)
        if opn in mapa:
            return mapa[opn]
    # procura por contém
    for op in opcoes:
        opn = normalizar_texto(op)
        for cn, c in mapa.items():
            if opn in cn or cn in opn:
                return c
    return None


def listar_arquivos_excel(pasta: Path) -> list[Path]:
    if not pasta.exists():
        return []
    arquivos = []
    for ext in ("*.xlsx", "*.xls"):
        arquivos.extend(pasta.rglob(ext))
    return [a for a in arquivos if not a.name.startswith("~$")]


def ler_excel_pasta(pasta: Path, tipo: str) -> pd.DataFrame:
    bases = []
    for arq in listar_arquivos_excel(pasta):
        try:
            df = pd.read_excel(arq)
            df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
            df["Arquivo Origem"] = arq.name
            df["Pasta Origem"] = tipo
            bases.append(df)
        except Exception as e:
            st.warning(f"Não foi possível ler {arq.name}: {e}")
    if not bases:
        return pd.DataFrame()
    return pd.concat(bases, ignore_index=True)


def mes_label(ano: int, mes: int) -> str:
    return f"{MESES_PT.get(mes, str(mes).zfill(2))}/{str(ano)[-2:]}"


def mes_key(label: str) -> tuple[int, int]:
    txt = normalizar_texto(label)
    m = re.search(r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[/\- ]?(\d{2,4})", txt)
    if m:
        mes = MESES_NOME.get(m.group(1), 1)
        ano = int(m.group(2))
        if ano < 100:
            ano += 2000
        return ano, mes
    return (9999, 12)


def extrair_mes_de_valor(valor) -> str | None:
    if pd.isna(valor):
        return None
    txt = str(valor).strip()
    norm = normalizar_texto(txt)

    # jan/26, fevereiro/2026 etc.
    m = re.search(r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ|JANEIRO|FEVEREIRO|MARCO|MARÇO|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)[/_\- ]?(\d{2,4})", norm)
    if m:
        mes = MESES_NOME.get(m.group(1))
        ano = int(m.group(2))
        if ano < 100:
            ano += 2000
        if mes:
            return mes_label(ano, mes)

    dt = pd.to_datetime(txt, errors="coerce", dayfirst=True)
    if pd.notna(dt):
        return mes_label(int(dt.year), int(dt.month))
    return None


def extrair_mes_de_arquivo(nome: str) -> str | None:
    norm = normalizar_texto(nome)
    ano_match = re.search(r"(20\d{2})", norm)
    ano = int(ano_match.group(1)) if ano_match else datetime.today().year
    for nome_mes, mes in MESES_NOME.items():
        if nome_mes in norm:
            return mes_label(ano, mes)
    return None


def meses_fechados(meses: list[str]) -> list[str]:
    hoje = datetime.today()
    resultado = []
    for m in meses:
        ano, mes = mes_key(m)
        if (ano, mes) < (hoje.year, hoje.month):
            resultado.append(m)
    return resultado


def img_to_base64(path: Path) -> str | None:
    if not path.exists():
        return None
    try:
        return base64.b64encode(path.read_bytes()).decode("utf-8")
    except Exception:
        return None

# =========================================================
# CLASSIFICAÇÃO DO DRE
# =========================================================
MAPEAMENTO = {
    "ICMS": ["ICMS"],
    "DAS": ["DAS"],
    "ISSQN": ["ISSQN"],
    "Mercadorias (CMP)": ["MERCADORIAS (CMP)", "MERCADORIA", "CMP", "CMV", "FORNECEDOR MERCADORIA", "COMPRA MERCADORIA"],
    "Alimentação": ["ALIMENTACAO", "ALIMENTACAO FUNCIONARIO"],
    "Pagamento de Férias": ["PAGAMENTO DE FERIAS", "FERIAS"],
    "FGTS": ["FGTS"],
    "Salários Fixos + Horas Extras": ["SALARIOS FIXOS", "HORAS EXTRAS", "SALARIO", "FOLHA"],
    "Comissões e Premiações": ["COMISSOES", "COMISSAO", "PREMIACOES", "PREMIACAO"],
    "INSS": ["INSS"],
    "IRRF Folha": ["IRRF FOLHA", "IRRF"],
    "Cursos, Treinamentos, Viagens": ["CURSOS", "TREINAMENTOS", "TREINAMENTO", "VIAGENS"],
    "PPRA<PCMSO e Exames": ["PPRA", "PCMSO", "EXAMES"],
    "Rescisão (guia)": ["RESCISAO"],
    "Provisão 13 e Férias": ["PROVISAO 13", "PROVISAO FERIAS", "13 SALARIO"],
    "Uniformes": ["UNIFORME"],
    "Vale Transporte": ["VALE TRANSPORTE", "VT"],
    "Pro-Labore": ["PRO-LABORE", "PRO LABORE"],
    "Outras Despesas com Pessoal": ["OUTRAS DESPESAS COM PESSOAL"],
    "Aluguel&IPTU": ["ALUGUEL", "IPTU"],
    "Seguro Imóvel": ["SEGURO IMOVEL", "SEGURO"],
    "Agua/Luz/Fone/Net": ["AGUA", "LUZ", "FONE", "NET", "ENERGIA", "INTERNET", "TELEFONE"],
    "Combustível": ["COMBUSTIVEL"],
    "Manutenção em Geral": ["MANUTENCAO EM GERAL", "MANUTENCAO"],
    "Manutenção Veículos/Motos": ["MANUTENCAO VEICULOS", "MANUTENCAO MOTOS", "VEICULOS", "MOTOS"],
    "Mat. de Escritório/Informática": ["ESCRITORIO", "INFORMATICA", "MATERIAL DE ESCRITORIO"],
    "Mat. De Limpeza": ["LIMPEZA"],
    "Viagens": ["VIAGENS"],
    "Contábil (Terceiros)": ["CONTABIL", "CONTABILIDADE"],
    "Sistemas (Terceiros)": ["SISTEMAS", "SISTEMA"],
    "Juridico (Terceiros)": ["JURIDICO", "ADVOGADO"],
    "Assessorias/Consultorias/Treinamentos": ["ASSESSORIA", "CONSULTORIA", "CONSULTORIAS"],
    "Taxas, Licenças e Contrib.": ["TAXAS", "LICENCAS", "CONTRIB", "ALVARA"],
    "Outros (Terceiros)": ["OUTROS TERCEIROS", "TERCEIROS"],
    "Outros (Despesas)": ["OUTROS DESPESAS", "OUTRAS DESPESAS", "RECARGA DE CELULAR", "DESPESAS COM SEGURANCA", "SEGURANCA"],
    "Quebra de caixa": ["QUEBRA DE CAIXA"],
    "Marketing/Publicidade": ["MARKETING", "PUBLICIDADE", "PROPAGANDA"],
    "Sistema Fidelidade": ["SISTEMA FIDELIDADE", "FIDELIDADE"],
    "Frete": ["FRETE"],
    "Associação de Classe (Royalties)": ["ASSOCIACAO DE CLASSE", "ROYALTIES"],
    "Juros Boletos": ["JUROS BOLETOS", "JUROS"],
    "Tarifas Bancarias": ["TARIFAS BANCARIAS", "TARIFA BANCARIA", "TARIFA"],
    "Taxas de Cartão (MDR)": ["TAXAS DE CARTAO", "MDR", "CARTAO", "CARTAO DE CREDITO"],
    "IRPJ/CSLL": ["IRPJ", "CSLL"],
    "Investimentos - Reformas": ["INVESTIMENTOS - REFORMAS", "REFORMAS", "REFORMA"],
    "Investimentos - Equipamentos": ["INVESTIMENTOS - EQUIPAMENTOS", "EQUIPAMENTOS"],
    "Investimentos - Expansão": ["INVESTIMENTOS - EXPANSAO", "EXPANSAO"],
    "Pagamento de Empréstimos (Principal)": ["EMPRESTIMOS", "EMPRESTIMO", "PRINCIPAL"],
    "Retiradas (Distribuição de Lucros)": ["RETIRADAS", "DISTRIBUICAO DE LUCROS", "DISTRIBUICAO"],
    "Parcelamento de Impostos": ["PARCELAMENTO DE IMPOSTOS", "PARCELAMENTO"],
    "Divergência Saídas": ["DIVERGENCIA SAIDAS", "DIVERGENCIA"],
}

DRE_LAYOUT = [
    ("1. RECEITA OPERACIONAL BRUTA", "bloco"),
    ("Receita de Vendas de Mercadorias", "item"),
    ("2. (-) DEDUÇÕES DA RECEITA BRUTA", "bloco"),
    ("ICMS", "item"), ("DAS", "item"), ("ISSQN", "item"),
    ("3. (=) RECEITA OPERACIONAL LÍQUIDA", "resultado"),
    ("4. (-) CUSTOS DAS VENDAS", "bloco"),
    ("Mercadorias (CMP)", "item"),
    ("5. (=) LUCRO BRUTO", "resultado"),
    ("6. (-) DESPESAS OPERACIONAIS", "bloco"),
    ("Despesas com Pessoal", "grupo"),
    ("Alimentação", "item"), ("Pagamento de Férias", "item"), ("FGTS", "item"),
    ("Salários Fixos + Horas Extras", "item"), ("Comissões e Premiações", "item"),
    ("INSS", "item"), ("IRRF Folha", "item"), ("Cursos, Treinamentos, Viagens", "item"),
    ("PPRA<PCMSO e Exames", "item"), ("Rescisão (guia)", "item"),
    ("Provisão 13 e Férias", "item"), ("Uniformes", "item"), ("Vale Transporte", "item"),
    ("Pro-Labore", "item"), ("Outras Despesas com Pessoal", "item"),
    ("Despesas Administrativas e Ocupação", "grupo"),
    ("Aluguel&IPTU", "item"), ("Seguro Imóvel", "item"), ("Agua/Luz/Fone/Net", "item"),
    ("Combustível", "item"), ("Manutenção em Geral", "item"), ("Manutenção Veículos/Motos", "item"),
    ("Mat. de Escritório/Informática", "item"), ("Mat. De Limpeza", "item"), ("Viagens", "item"),
    ("Contábil (Terceiros)", "item"), ("Sistemas (Terceiros)", "item"), ("Juridico (Terceiros)", "item"),
    ("Assessorias/Consultorias/Treinamentos", "item"), ("Taxas, Licenças e Contrib.", "item"),
    ("Outros (Terceiros)", "item"), ("Outros (Despesas)", "item"), ("Quebra de caixa", "item"),
    ("Despesas com Vendas e Marketing", "grupo"),
    ("Marketing/Publicidade", "item"), ("Sistema Fidelidade", "item"), ("Frete", "item"),
    ("Associação de Classe (Royalties)", "item"),
    ("7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO (EBITDA/LAJIDA)", "resultado"),
    ("8. (+/-) RESULTADO FINANCEIRO LÍQUIDO", "bloco"),
    ("Juros Boletos", "item"), ("Tarifas Bancarias", "item"), ("Taxas de Cartão (MDR)", "item"),
    ("9. (=) RESULTADO ANTES DOS TRIBUTOS (LAIR)", "resultado"),
    ("10. (-) TRIBUTOS SOBRE O LUCRO", "bloco"),
    ("IRPJ/CSLL", "item"),
    ("11. (=) LUCRO LÍQUIDO DO EXERCÍCIO", "resultado"),
    ("--- CONCILIAÇÃO DE FLUXO DE CAIXA (SAÍDAS NÃO-DRE) ---", "fluxo"),
    ("Investimentos - Reformas", "item"), ("Investimentos - Equipamentos", "item"),
    ("Investimentos - Expansão", "item"), ("Pagamento de Empréstimos (Principal)", "item"),
    ("Retiradas (Distribuição de Lucros)", "item"), ("Parcelamento de Impostos", "item"),
    ("Divergência Saídas", "item"), ("Posição Final", "resultado"),
]

PESSOAL = ["Alimentação", "Pagamento de Férias", "FGTS", "Salários Fixos + Horas Extras", "Comissões e Premiações", "INSS", "IRRF Folha", "Cursos, Treinamentos, Viagens", "PPRA<PCMSO e Exames", "Rescisão (guia)", "Provisão 13 e Férias", "Uniformes", "Vale Transporte", "Pro-Labore", "Outras Despesas com Pessoal"]
ADMIN = ["Aluguel&IPTU", "Seguro Imóvel", "Agua/Luz/Fone/Net", "Combustível", "Manutenção em Geral", "Manutenção Veículos/Motos", "Mat. de Escritório/Informática", "Mat. De Limpeza", "Viagens", "Contábil (Terceiros)", "Sistemas (Terceiros)", "Juridico (Terceiros)", "Assessorias/Consultorias/Treinamentos", "Taxas, Licenças e Contrib.", "Outros (Terceiros)", "Outros (Despesas)", "Quebra de caixa"]
MKT = ["Marketing/Publicidade", "Sistema Fidelidade", "Frete", "Associação de Classe (Royalties)"]
FIN = ["Juros Boletos", "Tarifas Bancarias", "Taxas de Cartão (MDR)"]
FLUXO = ["Investimentos - Reformas", "Investimentos - Equipamentos", "Investimentos - Expansão", "Pagamento de Empréstimos (Principal)", "Retiradas (Distribuição de Lucros)", "Parcelamento de Impostos", "Divergência Saídas"]


def classificar_conta(texto: str) -> str:
    t = normalizar_texto(texto)
    if not t or t in {"NAN", "NONE"}:
        return "NAO_CLASSIFICADO"
    for destino, palavras in MAPEAMENTO.items():
        for palavra in palavras:
            if normalizar_texto(palavra) in t:
                return destino
    return "NAO_CLASSIFICADO"

# =========================================================
# PROCESSAMENTO AUTOMÁTICO DAS PASTAS
# =========================================================
@st.cache_data(ttl=0, show_spinner=False)
def processar_bases(assinatura: str) -> dict:
    contas = ler_excel_pasta(PASTA_CONTAS, "CONTAS")
    plano = ler_excel_pasta(PASTA_PLANO, "PLANO")
    estoque = ler_excel_pasta(PASTA_ESTOQUE, "ESTOQUE")
    vendas = ler_excel_pasta(PASTA_VENDAS, "VENDAS")

    # ---------------- VENDAS / RECEITA ----------------
    vendas_proc = pd.DataFrame()
    receita_mes = {}
    if not vendas.empty:
        col_valor = procurar_coluna(vendas, ["Total", "Valor", "Valor Total", "Receita", "Vlr Total"])
        col_mes = procurar_coluna(vendas, ["Ano-mês", "Ano Mes", "Mês", "Mes", "Competência", "Competencia"])
        col_data = procurar_coluna(vendas, ["Data", "Data Venda", "Data Emissão", "Data Emissao"])
        col_loja = procurar_coluna(vendas, ["Unidade", "Loja", "Filial", "Apelido Un. Neg."])
        vendas_proc = vendas.copy()
        vendas_proc["Valor"] = limpar_valor_serie(vendas_proc[col_valor]) if col_valor else 0.0
        if col_mes:
            vendas_proc["Mês"] = vendas_proc[col_mes].apply(extrair_mes_de_valor)
        elif col_data:
            vendas_proc["Mês"] = vendas_proc[col_data].apply(extrair_mes_de_valor)
        else:
            vendas_proc["Mês"] = vendas_proc["Arquivo Origem"].apply(extrair_mes_de_arquivo)
        vendas_proc["Loja"] = vendas_proc[col_loja].astype(str) if col_loja else "Todas"
        vendas_proc = vendas_proc.dropna(subset=["Mês"])
        receita_mes = vendas_proc.groupby("Mês")["Valor"].sum().to_dict()

    # ---------------- CONTAS A PAGAR ----------------
    contas_proc = pd.DataFrame()
    nao_classificados = pd.DataFrame()
    valores_item_mes = {}
    valor_nao_classificado = 0.0
    qtd_nao_classificado = 0

    if not contas.empty:
        col_valor = procurar_coluna(contas, ["Valor Documento", "Valor", "Valor Total", "Total", "Vlr"])
        col_plano = procurar_coluna(contas, ["Plano de Contas", "Plano Contas", "Conta", "Classificação", "Classificacao"])
        col_loja = procurar_coluna(contas, ["Unidade", "Loja", "Filial", "Apelido Un. Neg."])
        col_data = procurar_coluna(contas, ["Data Pagamento", "Data Emissão", "Data Emissao", "Data Lançamento", "Data Lancamento", "Data"])

        contas_proc = contas.copy()
        contas_proc["Valor"] = limpar_valor_serie(contas_proc[col_valor]) if col_valor else 0.0
        contas_proc["Plano de Contas"] = contas_proc[col_plano].astype(str) if col_plano else ""
        contas_proc["Loja"] = contas_proc[col_loja].astype(str) if col_loja else "Todas"
        if col_data:
            contas_proc["Mês"] = contas_proc[col_data].apply(extrair_mes_de_valor)
        else:
            contas_proc["Mês"] = contas_proc["Arquivo Origem"].apply(extrair_mes_de_arquivo)

        # Tenta enriquecer com plano de contas legenda, se existir
        if not plano.empty and col_plano:
            col_plano_leg = procurar_coluna(plano, ["Plano de Contas", "Plano Contas", "Plano de contas", "Conta"])
            possiveis_destino = ["Destino na DRE Estruturada", "Destino DRE", "DRE", "Subgrupo", "Grupo"]
            col_dest = procurar_coluna(plano, possiveis_destino)
            if col_plano_leg and col_dest:
                tmp_plano = plano[[col_plano_leg, col_dest]].copy()
                tmp_plano["_k"] = tmp_plano[col_plano_leg].apply(normalizar_texto)
                tmp_plano = tmp_plano.drop_duplicates("_k")
                contas_proc["_k"] = contas_proc["Plano de Contas"].apply(normalizar_texto)
                contas_proc = contas_proc.merge(tmp_plano[["_k", col_dest]], on="_k", how="left")
                contas_proc["Texto Classificação"] = contas_proc["Plano de Contas"].astype(str) + " " + contas_proc[col_dest].astype(str)
            else:
                contas_proc["Texto Classificação"] = contas_proc["Plano de Contas"].astype(str)
        else:
            contas_proc["Texto Classificação"] = contas_proc["Plano de Contas"].astype(str)

        contas_proc["Item DRE"] = contas_proc["Texto Classificação"].apply(classificar_conta)
        contas_proc = contas_proc.dropna(subset=["Mês"])

        nao_classificados = contas_proc[contas_proc["Item DRE"] == "NAO_CLASSIFICADO"].copy()
        contas_classificadas = contas_proc[contas_proc["Item DRE"] != "NAO_CLASSIFICADO"].copy()
        valor_nao_classificado = float(nao_classificados["Valor"].sum()) if not nao_classificados.empty else 0.0
        qtd_nao_classificado = int(len(nao_classificados))
        valores_item_mes = contas_classificadas.groupby(["Item DRE", "Mês"])["Valor"].sum().to_dict()

    # ---------------- ESTOQUE ----------------
    estoque_mes = {}
    estoque_proc = pd.DataFrame()
    if not estoque.empty:
        estoque_proc = estoque.copy()
        col_valor = procurar_coluna(estoque_proc, ["Estoque X Custo Médio", "Estoque X Custo Medio", "Valor Estoque", "Estoque a Custo", "Total"])
        col_mes = procurar_coluna(estoque_proc, ["Mês", "Mes", "Ano-mês", "Competência", "Competencia"])
        col_data = procurar_coluna(estoque_proc, ["Data", "Data Estoque", "Referência", "Referencia"])
        if col_valor:
            estoque_proc["Valor"] = limpar_valor_serie(estoque_proc[col_valor])
            if col_mes:
                estoque_proc["Mês"] = estoque_proc[col_mes].apply(extrair_mes_de_valor)
            elif col_data:
                estoque_proc["Mês"] = estoque_proc[col_data].apply(extrair_mes_de_valor)
            else:
                estoque_proc["Mês"] = estoque_proc["Arquivo Origem"].apply(extrair_mes_de_arquivo)
            estoque_proc = estoque_proc.dropna(subset=["Mês"])
            estoque_mes = estoque_proc.groupby("Mês")["Valor"].sum().to_dict()

    todos_meses = sorted(set(receita_mes) | {m for _, m in valores_item_mes.keys()} | set(estoque_mes), key=mes_key)

    linhas = []
    for mes in todos_meses:
        rec = float(receita_mes.get(mes, 0.0))
        get = lambda item: float(valores_item_mes.get((item, mes), 0.0))
        ded = get("ICMS") + get("DAS") + get("ISSQN")
        rec_liq = rec - ded
        cmv = get("Mercadorias (CMP)")
        lucro_bruto = rec_liq - cmv
        desp_pessoal = sum(get(i) for i in PESSOAL)
        desp_admin = sum(get(i) for i in ADMIN)
        desp_mkt = sum(get(i) for i in MKT)
        desp_op = desp_pessoal + desp_admin + desp_mkt
        ebitda = lucro_bruto - desp_op
        res_fin = sum(get(i) for i in FIN)
        lair = ebitda - res_fin
        trib_lucro = get("IRPJ/CSLL")
        lucro_liq = lair - trib_lucro
        saidas_nao_dre = sum(get(i) for i in FLUXO)
        pos_final = lucro_liq - saidas_nao_dre

        valores_linha = {
            "1. RECEITA OPERACIONAL BRUTA": rec,
            "Receita de Vendas de Mercadorias": rec,
            "2. (-) DEDUÇÕES DA RECEITA BRUTA": ded,
            "ICMS": get("ICMS"), "DAS": get("DAS"), "ISSQN": get("ISSQN"),
            "3. (=) RECEITA OPERACIONAL LÍQUIDA": rec_liq,
            "4. (-) CUSTOS DAS VENDAS": cmv,
            "Mercadorias (CMP)": cmv,
            "5. (=) LUCRO BRUTO": lucro_bruto,
            "6. (-) DESPESAS OPERACIONAIS": desp_op,
            "Despesas com Pessoal": desp_pessoal,
            "Despesas Administrativas e Ocupação": desp_admin,
            "Despesas com Vendas e Marketing": desp_mkt,
            "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO (EBITDA/LAJIDA)": ebitda,
            "8. (+/-) RESULTADO FINANCEIRO LÍQUIDO": res_fin,
            "9. (=) RESULTADO ANTES DOS TRIBUTOS (LAIR)": lair,
            "10. (-) TRIBUTOS SOBRE O LUCRO": trib_lucro,
            "IRPJ/CSLL": trib_lucro,
            "11. (=) LUCRO LÍQUIDO DO EXERCÍCIO": lucro_liq,
            "--- CONCILIAÇÃO DE FLUXO DE CAIXA (SAÍDAS NÃO-DRE) ---": saidas_nao_dre,
            "Posição Final": pos_final,
        }
        # Itens individuais
        for item in list(MAPEAMENTO.keys()):
            valores_linha.setdefault(item, get(item))

        for ordem, (linha, estilo) in enumerate(DRE_LAYOUT):
            val = float(valores_linha.get(linha, 0.0))
            linhas.append({
                "Ordem": ordem,
                "Linha DRE": linha,
                "Estilo": estilo,
                "Mês": mes,
                "Valor": val,
                "% Receita": (val / rec) if rec else 0.0,
            })

    dre_long = pd.DataFrame(linhas)
    if not dre_long.empty:
        dre_long = dre_long.sort_values(["Ordem", "Mês"])

    return {
        "dre_long": dre_long,
        "contas": contas_proc,
        "vendas": vendas_proc,
        "estoque": estoque_proc,
        "nao_classificados": nao_classificados,
        "meses": todos_meses,
        "meses_fechados": meses_fechados(todos_meses),
        "valor_nao_classificado": valor_nao_classificado,
        "qtd_nao_classificado": qtd_nao_classificado,
        "estoque_mes": estoque_mes,
        "assinatura": assinatura,
    }


def assinatura_arquivos() -> str:
    partes = []
    for pasta in [PASTA_CONTAS, PASTA_PLANO, PASTA_ESTOQUE, PASTA_VENDAS]:
        for arq in listar_arquivos_excel(pasta):
            try:
                stt = arq.stat()
                partes.append(f"{arq.name}|{int(stt.st_mtime)}|{stt.st_size}")
            except Exception:
                partes.append(arq.name)
    return "::".join(sorted(partes))


def pivot_dre(dre_long: pd.DataFrame, meses: list[str]) -> pd.DataFrame:
    if dre_long.empty:
        return pd.DataFrame()
    base = dre_long[dre_long["Mês"].isin(meses)].copy()
    linhas = []
    for ordem, (linha, estilo) in enumerate(DRE_LAYOUT):
        row = {"Linha DRE": linha, "Estilo": estilo, "Ordem": ordem}
        for mes in meses:
            filtro = (base["Linha DRE"] == linha) & (base["Mês"] == mes)
            if filtro.any():
                val = float(base.loc[filtro, "Valor"].sum())
                perc = float(base.loc[filtro, "% Receita"].mean())
            else:
                val = 0.0
                perc = 0.0
            row[f"{mes} Valor"] = val
            row[f"{mes} %"] = perc
        linhas.append(row)
    return pd.DataFrame(linhas)


def render_dre_html(df: pd.DataFrame, meses: list[str]) -> str:
    if df.empty:
        return "<div class='info-box'>Nenhuma informação encontrada para os filtros selecionados.</div>"
    cols = ["Linha DRE"]
    for m in meses:
        cols += [f"{m} Valor", f"{m} %"]

    html = ["<div class='dre-wrap'><table class='dre-table'>"]
    html.append("<thead><tr>")
    for c in cols:
        html.append(f"<th>{c}</th>")
    html.append("</tr></thead><tbody>")

    for _, r in df.iterrows():
        estilo = str(r.get("Estilo", "item"))
        html.append(f"<tr class='{estilo}'>")
        html.append(f"<td>{r['Linha DRE']}</td>")
        for m in meses:
            html.append(f"<td>{dinheiro(r.get(f'{m} Valor', 0))}</td>")
            html.append(f"<td>{pct(r.get(f'{m} %', 0))}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "".join(html)

# =========================================================
# INTERFACE
# =========================================================
def header():
    logo_b64 = None
    for p in LOGO_PATHS:
        logo_b64 = img_to_base64(p)
        if logo_b64:
            break
    if logo_b64:
        logo_html = f"<img class='eirox-logo' src='data:image/png;base64,{logo_b64}' />"
    else:
        logo_html = "<div style='font-weight:900;color:#00a6ff;font-size:30px'>EIROX</div>"
    st.markdown(
        f"""
<div class="eirox-header">
  {logo_html}
  <div>
    <h1 class="eirox-title">DRE Empresa Online</h1>
    <div class="eirox-subtitle">Dashboard financeiro gerencial no padrão Eirox Pricing Online • Atualização automática pelas pastas de base</div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


def kpi(label: str, value: str, sub: str = ""):
    st.markdown(
        f"""
<div class="kpi-card">
  <div class="kpi-label">{label}</div>
  <div class="kpi-value">{value}</div>
  <div class="kpi-sub">{sub}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def obter_valor_linha(dre_long: pd.DataFrame, linha: str, meses: list[str]) -> float:
    if dre_long.empty:
        return 0.0
    f = (dre_long["Linha DRE"] == linha) & (dre_long["Mês"].isin(meses))
    return float(dre_long.loc[f, "Valor"].sum())


def main():
    header()

    with st.sidebar:
        st.markdown("### ⚙️ Atualização")
        st.caption("As bases são lidas direto das pastas. Ao substituir arquivos, clique em Atualizar Base.")
        if st.button("🔄 Atualizar Base", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        st.markdown("---")
        st.markdown("### 📂 Pastas monitoradas")
        st.caption(str(PASTA_BASE))
        for p in PASTAS_OBRIGATORIAS:
            ok = "✅" if (PASTA_BASE / p).exists() else "❌"
            st.write(f"{ok} {p}")

    assinatura = assinatura_arquivos()
    with st.spinner("Atualizando informações a partir das bases..."):
        dados = processar_bases(assinatura)

    dre_long = dados["dre_long"]
    meses = dados["meses"]
    meses_padrao = dados["meses_fechados"] or meses

    if not meses:
        st.error("Nenhum mês foi identificado nas bases. Verifique as pastas de dados.")
        st.stop()

    with st.sidebar:
        st.markdown("---")
        st.markdown("### 🗓️ Filtros")
        meses_sel = st.multiselect(
            "Meses fechados",
            options=meses,
            default=meses_padrao,
            help="Por padrão o mês atual fica fora da seleção."
        )
        if not meses_sel:
            meses_sel = meses_padrao

    meses_sel = sorted(meses_sel, key=mes_key)
    df_dre = pivot_dre(dre_long, meses_sel)

    receita = obter_valor_linha(dre_long, "1. RECEITA OPERACIONAL BRUTA", meses_sel)
    receita_liq = obter_valor_linha(dre_long, "3. (=) RECEITA OPERACIONAL LÍQUIDA", meses_sel)
    lucro_bruto = obter_valor_linha(dre_long, "5. (=) LUCRO BRUTO", meses_sel)
    ebitda = obter_valor_linha(dre_long, "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO (EBITDA/LAJIDA)", meses_sel)
    lucro_liq = obter_valor_linha(dre_long, "11. (=) LUCRO LÍQUIDO DO EXERCÍCIO", meses_sel)
    margem_ebitda = ebitda / receita if receita else 0
    margem_liq = lucro_liq / receita if receita else 0

    abas = st.tabs(["📊 Painel Executivo", "📋 DRE Gerencial", "📈 Evolução Mensal", "⚠️ Auditoria DRE", "📁 Bases"])

    with abas[0]:
        st.markdown("<div class='section-title'>Painel Executivo</div>", unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi("Receita Bruta", dinheiro(receita), "Meses fechados selecionados")
        with c2: kpi("Receita Líquida", dinheiro(receita_liq), pct(receita_liq / receita if receita else 0))
        with c3: kpi("Lucro Bruto", dinheiro(lucro_bruto), pct(lucro_bruto / receita if receita else 0))
        with c4: kpi("EBITDA", dinheiro(ebitda), pct(margem_ebitda))
        c5, c6, c7, c8 = st.columns(4)
        with c5: kpi("Lucro Líquido", dinheiro(lucro_liq), pct(margem_liq))
        with c6: kpi("Não Classificados", str(dados["qtd_nao_classificado"]), dinheiro(dados["valor_nao_classificado"]))
        with c7: kpi("Meses Fechados", str(len(meses_padrao)), ", ".join(meses_padrao[-3:]))
        with c8: kpi("Última Atualização", datetime.now().strftime("%d/%m/%Y %H:%M"), "Leitura direta das pastas")

    with abas[1]:
        st.markdown("<div class='section-title'>DRE Gerencial</div>", unsafe_allow_html=True)
        st.markdown("<div class='small-muted'>Modelo mantido no formato aprovado: meses em colunas, valores e percentuais lado a lado, com todos os blocos estruturais destacados.</div>", unsafe_allow_html=True)
        st.markdown(render_dre_html(df_dre, meses_sel), unsafe_allow_html=True)

    with abas[2]:
        st.markdown("<div class='section-title'>Evolução Mensal</div>", unsafe_allow_html=True)
        if not dre_long.empty:
            linhas_chave = [
                "1. RECEITA OPERACIONAL BRUTA",
                "5. (=) LUCRO BRUTO",
                "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO (EBITDA/LAJIDA)",
                "11. (=) LUCRO LÍQUIDO DO EXERCÍCIO",
            ]
            evo = dre_long[(dre_long["Linha DRE"].isin(linhas_chave)) & (dre_long["Mês"].isin(meses_sel))].copy()
            evo["Mês Ordem"] = evo["Mês"].apply(lambda x: mes_key(x)[0] * 100 + mes_key(x)[1])
            evo = evo.sort_values("Mês Ordem")
            fig = px.line(evo, x="Mês", y="Valor", color="Linha DRE", markers=True, template="plotly_dark")
            fig.update_layout(height=470, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)

    with abas[3]:
        st.markdown("<div class='section-title'>Auditoria DRE</div>", unsafe_allow_html=True)
        st.info("Os não classificados estão fora dos cálculos principais e aparecem aqui apenas para tratamento da classificação.")
        nc = dados["nao_classificados"]
        if nc.empty:
            st.success("Nenhuma conta não classificada encontrada.")
        else:
            cols = [c for c in ["Mês", "Loja", "Plano de Contas", "Valor", "Arquivo Origem"] if c in nc.columns]
            st.dataframe(nc[cols], use_container_width=True, hide_index=True)

    with abas[4]:
        st.markdown("<div class='section-title'>Bases Monitoradas</div>", unsafe_allow_html=True)
        st.write("Toda troca, substituição ou inclusão de arquivos nas pastas abaixo será considerada ao clicar em **Atualizar Base** ou recarregar o app.")
        arquivos = []
        for pasta in [PASTA_CONTAS, PASTA_PLANO, PASTA_ESTOQUE, PASTA_VENDAS]:
            for arq in listar_arquivos_excel(pasta):
                stt = arq.stat()
                arquivos.append({"Pasta": pasta.name, "Arquivo": arq.name, "Modificado em": datetime.fromtimestamp(stt.st_mtime).strftime("%d/%m/%Y %H:%M"), "Tamanho KB": round(stt.st_size/1024, 1)})
        st.dataframe(pd.DataFrame(arquivos), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()
