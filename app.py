import base64
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

try:
    import plotly.graph_objects as go
except Exception:
    go = None

st.set_page_config(page_title="Eirox DRE Online", page_icon="📊", layout="wide")

REQUIRED_FOLDERS = [
    "CONTAS A PAGAR - DRE",
    "PLANO DE CONTAS - LEGENDA",
    "POSIÇÃO DE ESTOQUE",
    "VENDA POR PAGAMENTO",
]

MONTH_MAP = {
    "jan": 1, "janeiro": 1,
    "fev": 2, "fevereiro": 2,
    "mar": 3, "março": 3, "marco": 3,
    "abr": 4, "abril": 4,
    "mai": 5, "maio": 5,
    "jun": 6, "junho": 6,
    "jul": 7, "julho": 7,
    "ago": 8, "agosto": 8,
    "set": 9, "setembro": 9,
    "out": 10, "outubro": 10,
    "nov": 11, "novembro": 11,
    "dez": 12, "dezembro": 12,
}
REV_MONTH = {1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun", 7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"}

PESSOAL = ["Alimentação", "Pagamento de Férias", "FGTS", "Salários Fixos + Horas Extras", "Comissões e Premiações", "INSS", "IRRF Folha", "Cursos, Treinamentos, Viagens", "PPRA<PCMSO e Exames", "Rescisão (guia)", "Provisão 13 e Férias", "Uniformes", "Vale Transporte", "Pro-Labore", "Outras Despesas com Pessoal"]
ADMIN = ["Aluguel&IPTU", "Seguro Imóvel", "Agua/Luz/Fone/Net", "Combustível", "Manutenção em Geral", "Manutenção Veículos/Motos", "Mat. de Escritório/Informática", "Mat. De Limpeza", "Viagens", "Contábil (Terceiros)", "Sistemas (Terceiros)", "Juridico (Terceiros)", "Assessorias/Consultorias/Treinamentos", "Taxas, Licenças e Contrib.", "Outros (Terceiros)", "Outros (Despesas)", "Quebra de caixa"]
VENDAS_MKT = ["Marketing/Publicidade", "Sistema Fidelidade", "Frete", "Associação de Classe (Royalties)"]
DEDUCOES = ["ICMS", "DAS", "ISSQN"]
FINANCEIRO = ["Juros Boletos", "Tarifas Bancarias", "Taxas de Cartão (MDR)"]
TRIBUTOS = ["IRPJ/CSLL"]

STRUCTURE = [
    ("1. RECEITA OPERACIONAL BRUTA", "total", "blue"),
    ("Receita de Vendas de Mercadorias", "detail", ""),
    ("2. (-) DEDUÇÕES DA RECEITA BRUTA", "total", "blue"),
    ("ICMS", "detail", ""), ("DAS", "detail", ""), ("ISSQN", "detail", ""),
    ("3. (=) RECEITA OPERACIONAL LÍQUIDA", "result", "yellow"),
    ("4. (-) CUSTOS DAS VENDAS", "total", "blue"),
    ("Mercadorias (CMP)", "detail", ""),
    ("5. (=) LUCRO BRUTO", "result", "yellow"),
    ("6. (-) DESPESAS OPERACIONAIS", "total", "blue"),
    ("Despesas com Pessoal", "group", "gray"),
    *[(x, "detail", "") for x in PESSOAL],
    ("Despesas Administrativas e Ocupação", "group", "gray"),
    *[(x, "detail", "") for x in ADMIN],
    ("Despesas com Vendas e Marketing", "group", "gray"),
    *[(x, "detail", "") for x in VENDAS_MKT],
    ("7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO (EBITDA/LAJIDA)", "result", "yellow"),
    ("8. (+/-) RESULTADO FINANCEIRO LÍQUIDO", "total", "blue"),
    *[(x, "detail", "") for x in FINANCEIRO],
    ("9. (=) RESULTADO ANTES DOS TRIBUTOS (LAIR)", "result", "yellow"),
    ("10. (-) TRIBUTOS SOBRE O LUCRO", "total", "blue"),
    *[(x, "detail", "") for x in TRIBUTOS],
    ("11. (=) LUCRO LÍQUIDO DO EXERCÍCIO", "result", "yellow"),
]
STYLE_MAP = {line: style for line, _kind, style in STRUCTURE}
KIND_MAP = {line: kind for line, kind, _style in STRUCTURE}

# -----------------------------
# Formatação e utilitários
# -----------------------------
def fmt_brl(v):
    try:
        v = float(v)
    except Exception:
        v = 0.0
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def fmt_pct(v):
    try:
        v = float(v)
    except Exception:
        v = 0.0
    if abs(v) <= 1:
        v *= 100
    return f"{v:.2f}%".replace(".", ",")


def norm_txt(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().lower()
    tr = str.maketrans("áàâãéêíóôõúç", "aaaaeeiooouc")
    return s.translate(tr)


def money_to_float(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace("R$", "").replace("%", "")
    s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^0-9\.-]", "", s)
    try:
        return float(s)
    except Exception:
        return 0.0


def month_key(m):
    s = norm_txt(m)
    if s in ["sem mes", "sem_mes", "nan", "none", ""]:
        return (9999, 99)
    match = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-zç]*[/-]?(\d{2,4})?", s)
    if match:
        mes = MONTH_MAP.get(match.group(1), 99)
        yy = match.group(2)
        ano = 2000 + int(yy) if yy and len(yy) == 2 else int(yy) if yy else 9999
        return (ano, mes)
    dt = pd.to_datetime(m, errors="coerce", dayfirst=True)
    if pd.notna(dt):
        return (int(dt.year), int(dt.month))
    return (9999, 99)


def month_label_from_value(v):
    if pd.isna(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return f"{REV_MONTH[int(v.month)]}/{str(int(v.year))[-2:]}"
    dt = pd.to_datetime(v, errors="coerce", dayfirst=True)
    if pd.notna(dt):
        return f"{REV_MONTH[int(dt.month)]}/{str(int(dt.year))[-2:]}"
    ano, mes = month_key(v)
    if mes != 99 and ano != 9999:
        return f"{REV_MONTH[mes]}/{str(ano)[-2:]}"
    return None


def is_current_month_label(m):
    ano, mes = month_key(m)
    hoje = datetime.today()
    return ano == hoje.year and mes == hoje.month


def closed_months(months):
    fechados = [m for m in months if not is_current_month_label(m)]
    return fechados if fechados else months


def find_project_base():
    here = Path(__file__).resolve().parent
    candidates = []
    for c in [Path.cwd(), here, here.parent, Path.home() / "Desktop" / "Dre", Path("C:/Users/Comercial/Desktop/Dre")]:
        if c not in candidates:
            candidates.append(c)
    for c in candidates:
        try:
            if all((c / f).exists() for f in REQUIRED_FOLDERS):
                return c, "pastas"
        except Exception:
            pass
    for c in candidates:
        try:
            if (c / "data" / "DRE_Consolidado_Moderno.xlsx").exists() or (c / "DRE_Consolidado_Moderno.xlsx").exists():
                return c, "excel"
        except Exception:
            pass
    return here, "nenhum"

BASE_DIR, BASE_MODE = find_project_base()


def get_logo():
    possible = [
        BASE_DIR / "assets" / "logo_eirox.png",
        BASE_DIR / "assets" / "logo eirox.png",
        BASE_DIR / "logo_eirox.png",
        BASE_DIR / "logo eirox.png",
        Path(__file__).resolve().parent / "assets" / "logo_eirox.png",
    ]
    for p in possible:
        if p.exists():
            try:
                return base64.b64encode(p.read_bytes()).decode()
            except Exception:
                return None
    return None


def read_all_excels(folder):
    files = list(folder.rglob("*.xlsx")) + list(folder.rglob("*.xls"))
    files = [f for f in files if not f.name.startswith("~$")]
    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f)
            df["ARQUIVO_ORIGEM"] = f.name
            df["PASTA_ORIGEM"] = folder.name
            dfs.append(df)
        except Exception:
            pass
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def find_col(df, names):
    if df.empty:
        return None
    cols_norm = {norm_txt(c): c for c in df.columns}
    for n in names:
        nn = norm_txt(n)
        if nn in cols_norm:
            return cols_norm[nn]
    for c in df.columns:
        cn = norm_txt(c)
        for n in names:
            if norm_txt(n) in cn:
                return c
    return None

# -----------------------------
# Leitura do consolidado ou das pastas
# -----------------------------
def parse_dre_sheet_from_excel(path):
    raw = pd.read_excel(path, sheet_name="DRE", header=None)
    header_row = None
    for i in range(min(25, len(raw))):
        if raw.iloc[i].astype(str).str.contains("Linha DRE", case=False, na=False).any():
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()
    month_row = header_row - 1
    months = []
    for col in range(1, raw.shape[1], 2):
        label = month_label_from_value(raw.iat[month_row, col] if month_row >= 0 else None)
        if label:
            months.append((col, label))
    rows = []
    for r in range(header_row + 1, len(raw)):
        linha = raw.iat[r, 0]
        if pd.isna(linha) or str(linha).strip() == "":
            continue
        linha = str(linha).strip()
        row = {"Linha DRE": linha, "_kind": KIND_MAP.get(linha, "detail"), "_style": STYLE_MAP.get(linha, "")}
        for col, m in months:
            row[f"{m} Valor"] = money_to_float(raw.iat[r, col] if col < raw.shape[1] else 0)
            row[f"{m} %"] = money_to_float(raw.iat[r, col + 1] if col + 1 < raw.shape[1] else 0)
        rows.append(row)
    return pd.DataFrame(rows)


def load_from_consolidated():
    possible = [BASE_DIR / "data" / "DRE_Consolidado_Moderno.xlsx", BASE_DIR / "DRE_Consolidado_Moderno.xlsx"]
    path = next((p for p in possible if p.exists()), None)
    if not path:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    xls = pd.ExcelFile(path)
    dre = parse_dre_sheet_from_excel(path) if "DRE" in xls.sheet_names else pd.DataFrame()
    nao = pd.read_excel(path, sheet_name="NAO_CLASSIFICADOS") if "NAO_CLASSIFICADOS" in xls.sheet_names else pd.DataFrame()
    resumo_loja = pd.read_excel(path, sheet_name="RESUMO_LOJA") if "RESUMO_LOJA" in xls.sheet_names else pd.DataFrame()
    estoque = pd.read_excel(path, sheet_name="ESTOQUE_BASE") if "ESTOQUE_BASE" in xls.sheet_names else pd.DataFrame()
    return dre, nao, resumo_loja, estoque


def classify_by_text(text):
    t = norm_txt(text)
    if "icms" in t: return "ICMS"
    if t == "das" or " das" in t: return "DAS"
    if "issqn" in t: return "ISSQN"
    if any(k in t for k in ["mercadoria", "cmp", "cmv", "fornecedor"]): return "Mercadorias (CMP)"
    if "fgts" in t: return "FGTS"
    if any(k in t for k in ["salario", "hora extra"]): return "Salários Fixos + Horas Extras"
    if any(k in t for k in ["comissao", "premiacao"]): return "Comissões e Premiações"
    if "inss" in t: return "INSS"
    if "irrf" in t: return "IRRF Folha"
    if any(k in t for k in ["curso", "treinamento", "viagem"]): return "Cursos, Treinamentos, Viagens"
    if "rescisao" in t: return "Rescisão (guia)"
    if any(k in t for k in ["pro labore", "pro-labore"]): return "Pro-Labore"
    if "alimentacao" in t: return "Alimentação"
    if any(k in t for k in ["aluguel", "iptu"]): return "Aluguel&IPTU"
    if "seguro imovel" in t: return "Seguro Imóvel"
    if any(k in t for k in ["agua", "luz", "fone", "internet", "energia", "net"]): return "Agua/Luz/Fone/Net"
    if "combustivel" in t: return "Combustível"
    if "manutencao" in t and any(k in t for k in ["veiculo", "moto"]): return "Manutenção Veículos/Motos"
    if "manutencao" in t: return "Manutenção em Geral"
    if any(k in t for k in ["escritorio", "informatica"]): return "Mat. de Escritório/Informática"
    if "limpeza" in t: return "Mat. De Limpeza"
    if "contabil" in t: return "Contábil (Terceiros)"
    if "sistema" in t and "fidelidade" not in t: return "Sistemas (Terceiros)"
    if "juridico" in t: return "Juridico (Terceiros)"
    if any(k in t for k in ["assessoria", "consultoria"]): return "Assessorias/Consultorias/Treinamentos"
    if any(k in t for k in ["taxa", "licenca", "contrib"]): return "Taxas, Licenças e Contrib."
    if any(k in t for k in ["marketing", "publicidade"]): return "Marketing/Publicidade"
    if "fidelidade" in t: return "Sistema Fidelidade"
    if "frete" in t: return "Frete"
    if "juros" in t: return "Juros Boletos"
    if "tarifa" in t: return "Tarifas Bancarias"
    if any(k in t for k in ["cartao", "mdr"]): return "Taxas de Cartão (MDR)"
    if any(k in t for k in ["irpj", "csll"]): return "IRPJ/CSLL"
    if "outros terceiros" in t: return "Outros (Terceiros)"
    if any(k in t for k in ["outros", "outras"]): return "Outros (Despesas)"
    return "NAO_CLASSIFICADO"


def infer_month_from_file(filename):
    s = norm_txt(filename)
    for k in MONTH_MAP:
        if k in s:
            return f"{k[:3]}/26"
    return None


def load_from_folders():
    venda = read_all_excels(BASE_DIR / "VENDA POR PAGAMENTO")
    contas = read_all_excels(BASE_DIR / "CONTAS A PAGAR - DRE")
    estoque = read_all_excels(BASE_DIR / "POSIÇÃO DE ESTOQUE")
    receita_by_month = {}
    if not venda.empty:
        col_total = find_col(venda, ["Total", "Valor", "Valor Total", "Valor Venda"])
        col_mes = find_col(venda, ["Ano-mês", "Ano Mes", "Mês", "Mes", "Data"])
        if col_total:
            venda["_VALOR"] = venda[col_total].apply(money_to_float)
            venda["_MES"] = venda[col_mes].apply(month_label_from_value) if col_mes else venda["ARQUIVO_ORIGEM"].apply(infer_month_from_file)
            venda["_MES"] = venda["_MES"].fillna("sem_mes")
            receita_by_month = venda.groupby("_MES")["_VALOR"].sum().to_dict()
    nao = pd.DataFrame()
    resumo_loja = pd.DataFrame()
    class_month = {}
    if not contas.empty:
        c_valor = find_col(contas, ["Valor Documento", "Valor", "Valor Total", "Total"])
        c_plano = find_col(contas, ["Plano de Contas", "Plano Contas", "Conta", "Descrição", "Historico"])
        c_mes = find_col(contas, ["Data Emissão", "Data Pagamento", "Mês", "Mes", "Data"])
        c_loja = find_col(contas, ["Unidade", "Loja", "Apelido Un. Neg.", "Un. Neg."])
        contas["_VALOR"] = contas[c_valor].apply(money_to_float) if c_valor else 0
        contas["_TEXTO"] = contas[c_plano].astype(str) if c_plano else ""
        contas["_CLASS"] = contas["_TEXTO"].apply(classify_by_text)
        contas["_MES"] = contas[c_mes].apply(month_label_from_value) if c_mes else contas["ARQUIVO_ORIGEM"].apply(infer_month_from_file)
        contas["_MES"] = contas["_MES"].fillna("sem_mes")
        contas["_LOJA"] = contas[c_loja].astype(str) if c_loja else "Não identificado"
        nao = contas[contas["_CLASS"] == "NAO_CLASSIFICADO"].copy()
        class_month = contas[contas["_CLASS"] != "NAO_CLASSIFICADO"].groupby(["_CLASS", "_MES"])["_VALOR"].sum().to_dict()
        resumo_loja = contas[contas["_CLASS"] != "NAO_CLASSIFICADO"].groupby(["_LOJA", "_MES"])["_VALOR"].sum().reset_index().rename(columns={"_LOJA": "Loja", "_MES": "Mês", "_VALOR": "Despesas Classificadas"})
    months = sorted(set(receita_by_month.keys()) | {m for _, m in class_month.keys()}, key=month_key)
    months = [m for m in months if m and m != "sem_mes"]
    if not months:
        return pd.DataFrame(), nao, resumo_loja, estoque
    rows = []
    for line, kind, style in STRUCTURE:
        row = {"Linha DRE": line, "_kind": kind, "_style": style}
        for m in months:
            receita = receita_by_month.get(m, 0)
            def val(name): return class_month.get((name, m), 0)
            ded = sum(val(x) for x in DEDUCOES)
            cmv = val("Mercadorias (CMP)")
            despesas_op = sum(val(x) for x in PESSOAL + ADMIN + VENDAS_MKT)
            financeiro = sum(val(x) for x in FINANCEIRO)
            tributos = sum(val(x) for x in TRIBUTOS)
            if line in ["1. RECEITA OPERACIONAL BRUTA", "Receita de Vendas de Mercadorias"]:
                v = receita
            elif line == "2. (-) DEDUÇÕES DA RECEITA BRUTA":
                v = ded
            elif line == "3. (=) RECEITA OPERACIONAL LÍQUIDA":
                v = receita - ded
            elif line == "4. (-) CUSTOS DAS VENDAS":
                v = cmv
            elif line == "5. (=) LUCRO BRUTO":
                v = receita - ded - cmv
            elif line == "6. (-) DESPESAS OPERACIONAIS":
                v = despesas_op
            elif line == "Despesas com Pessoal":
                v = sum(val(x) for x in PESSOAL)
            elif line == "Despesas Administrativas e Ocupação":
                v = sum(val(x) for x in ADMIN)
            elif line == "Despesas com Vendas e Marketing":
                v = sum(val(x) for x in VENDAS_MKT)
            elif line == "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO (EBITDA/LAJIDA)":
                v = receita - ded - cmv - despesas_op
            elif line == "8. (+/-) RESULTADO FINANCEIRO LÍQUIDO":
                v = financeiro
            elif line == "9. (=) RESULTADO ANTES DOS TRIBUTOS (LAIR)":
                v = receita - ded - cmv - despesas_op - financeiro
            elif line == "10. (-) TRIBUTOS SOBRE O LUCRO":
                v = tributos
            elif line == "11. (=) LUCRO LÍQUIDO DO EXERCÍCIO":
                v = receita - ded - cmv - despesas_op - financeiro - tributos
            else:
                v = val(line)
            row[f"{m} Valor"] = float(v)
            row[f"{m} %"] = float(v / receita) if receita else 0.0
        rows.append(row)
    return pd.DataFrame(rows), nao, resumo_loja, estoque


def get_raw_value(dre, line, month):
    col = f"{month} Valor"
    if col not in dre.columns:
        return 0.0
    s = dre.loc[dre["Linha DRE"].astype(str).str.contains(line, case=False, regex=False), col]
    return float(s.iloc[0]) if not s.empty else 0.0


def calc_series(dre, line, months):
    return [get_raw_value(dre, line, m) for m in months]


def make_metrics(dre, month):
    rec = get_raw_value(dre, "1. RECEITA OPERACIONAL BRUTA", month)
    rec_liq = get_raw_value(dre, "3. (=) RECEITA OPERACIONAL LÍQUIDA", month)
    cmv = get_raw_value(dre, "4. (-) CUSTOS DAS VENDAS", month)
    lb = get_raw_value(dre, "5. (=) LUCRO BRUTO", month)
    desp = get_raw_value(dre, "6. (-) DESPESAS OPERACIONAIS", month)
    ebitda = get_raw_value(dre, "7. (=) RESULTADO", month)
    lair = get_raw_value(dre, "9. (=) RESULTADO", month)
    lucro = get_raw_value(dre, "11. (=) LUCRO LÍQUIDO", month)
    ded = get_raw_value(dre, "2. (-) DEDUÇÕES", month)
    return {
        "Receita Bruta": rec,
        "Receita Líquida": rec_liq,
        "Deduções %": ded / rec if rec else 0,
        "CMV %": cmv / rec if rec else 0,
        "Lucro Bruto": lb,
        "Margem Bruta %": lb / rec if rec else 0,
        "Despesas Operacionais": desp,
        "Despesas Op. %": desp / rec if rec else 0,
        "EBITDA": ebitda,
        "Margem EBITDA %": ebitda / rec if rec else 0,
        "LAIR": lair,
        "Lucro Líquido": lucro,
        "Margem Líquida %": lucro / rec if rec else 0,
    }


def delta_pct_atual_vs_anterior(dre, line, months):
    if len(months) < 2:
        return ""
    atual = get_raw_value(dre, line, months[-1])
    ant = get_raw_value(dre, line, months[-2])
    if ant == 0:
        return ""
    return fmt_pct((atual - ant) / abs(ant))


def html_dre_table(df, months):
    rows_html = []
    header = "<tr><th class='col-linha'>Linha DRE</th>" + "".join([f"<th>{m} Valor</th><th>{m} %</th>" for m in months]) + "</tr>"
    for _, r in df.iterrows():
        style = r.get("_style", "")
        cls = "row-detail"
        if style == "yellow": cls = "row-yellow"
        elif style == "blue": cls = "row-blue"
        elif style == "gray": cls = "row-gray"
        line = str(r.get("Linha DRE", ""))
        cells = [f"<td class='col-linha'>{line}</td>"]
        for m in months:
            v = float(r.get(f"{m} Valor", 0) or 0)
            p = float(r.get(f"{m} %", 0) or 0)
            cells.append(f"<td>{fmt_brl(v)}</td><td>{fmt_pct(p)}</td>")
        rows_html.append(f"<tr class='{cls}'>" + "".join(cells) + "</tr>")
    return f"""
    <div class='dre-scroll'>
      <table class='dre-table'>
        <thead>{header}</thead>
        <tbody>{''.join(rows_html)}</tbody>
      </table>
    </div>
    """


def line_chart(title, months, series_dict):
    if go is None:
        chart_df = pd.DataFrame({"Mês": months, **series_dict}).set_index("Mês")
        st.line_chart(chart_df)
        return
    fig = go.Figure()
    for name, vals in series_dict.items():
        fig.add_trace(go.Scatter(x=months, y=vals, mode="lines+markers", name=name))
    fig.update_layout(
        title=title,
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=50, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=380,
    )
    fig.update_yaxes(tickprefix="R$ ", separatethousands=True)
    st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# CSS Enterprise
# -----------------------------
st.markdown("""
<style>
.stApp {background: #06111F; color: #F8FAFC;}
[data-testid="stSidebar"] {background: #050B13; border-right: 1px solid #17385A;}
h1, h2, h3 {color: #FFFFFF;}
.block-container {padding-top: 1.6rem; max-width: 1550px;}
.hero {text-align:center; margin: 6px 0 20px 0;}
.hero img {height: 54px; margin-bottom: 4px;}
.hero h1 {font-size: 58px; margin: 0; font-weight: 950; letter-spacing: -1px;}
.hero p {color:#7DC8FF; font-size:20px; margin-top:12px;}
.card {background: linear-gradient(135deg, #0B1628, #101E34);border: 1px solid #1E3A5F;border-radius: 18px;padding: 20px;box-shadow: 0 12px 30px rgba(0,0,0,.30); min-height: 126px;}
.metric-label {color:#9FB3C8;font-size:13px;font-weight:800;text-transform:uppercase;letter-spacing:.4px;}
.metric-value {color:#FFFFFF;font-size:28px;font-weight:950;margin-top:6px;white-space:nowrap;}
.metric-sub {color:#00AEEF;font-size:13px;margin-top:6px;font-weight:800;}
.section-title {font-size:26px;font-weight:950;margin:18px 0 10px 0;color:#fff;}
.dre-scroll {overflow:auto; border:1px solid #17385A; border-radius:14px; max-height:680px;}
.dre-table {border-collapse:collapse; width:max-content; min-width:100%; font-size:15px;}
.dre-table th {position: sticky; top:0; background:#171B24; color:#B8C3D0; z-index:2; padding:12px 10px; border:1px solid #263447; text-align:right; font-weight:800;}
.dre-table td {padding:11px 10px; border:1px solid #263447; text-align:right; white-space:nowrap;}
.dre-table .col-linha {position: sticky; left:0; z-index:1; text-align:left; min-width:360px; max-width:460px;}
.dre-table th.col-linha {z-index:3;}
.row-detail td {background:#0B1628; color:#F5F7FA;}
.row-detail .col-linha {background:#0B1628;}
.row-blue td {background:#D6E8FB; color:#06111F; font-weight:950;}
.row-blue .col-linha {background:#D6E8FB;}
.row-yellow td {background:#FFD11A; color:#06111F; font-weight:950;}
.row-yellow .col-linha {background:#FFD11A;}
.row-gray td {background:#2A3342; color:#FFFFFF; font-weight:900; font-style:italic;}
.row-gray .col-linha {background:#2A3342;}
.audit-box {background:#3B1D29; border:1px solid #6E3145; border-radius:14px; padding:16px; color:#FFB5C3;}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Renderização
# -----------------------------
logo = get_logo()
if logo:
    st.markdown(f"""
    <div class='hero'>
        <img src='data:image/png;base64,{logo}'>
        <h1>DRE Empresa Online</h1>
        <p>Dashboard financeiro gerencial no padrão Eirox Pricing Online • DRE executivo com indicadores, evolução mensal e auditoria</p>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("<div class='hero'><h1>DRE Empresa Online</h1><p>Dashboard financeiro gerencial no padrão Eirox Pricing Online</p></div>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Atualização")
    st.caption("As bases são lidas automaticamente. Ao substituir arquivos, clique em Atualizar Base.")
    if st.button("🔄 Atualizar Base", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.divider()
    st.markdown("### 📁 Base encontrada")
    st.caption(str(BASE_DIR))
    for f in REQUIRED_FOLDERS:
        st.markdown(("✅ " if (BASE_DIR / f).exists() else "⚪ ") + f)
    if (BASE_DIR / "data" / "DRE_Consolidado_Moderno.xlsx").exists() or (BASE_DIR / "DRE_Consolidado_Moderno.xlsx").exists():
        st.markdown("✅ DRE_Consolidado_Moderno.xlsx")
    st.caption(f"Modo de leitura: {BASE_MODE}")

try:
    if BASE_MODE == "pastas":
        dre, nao, resumo_loja, estoque = load_from_folders()
    elif BASE_MODE == "excel":
        dre, nao, resumo_loja, estoque = load_from_consolidated()
    else:
        dre, nao, resumo_loja, estoque = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    if dre.empty:
        st.error("Nenhuma base foi encontrada. Para rodar online, coloque DRE_Consolidado_Moderno.xlsx dentro da pasta data do repositório. Para rodar local, mantenha as 4 pastas de base junto do app.py.")
        st.stop()

    all_months = sorted({str(c).replace(" Valor", "") for c in dre.columns if str(c).endswith(" Valor")}, key=month_key)
    all_months = [m for m in all_months if month_key(m) != (9999, 99)]
    if not all_months:
        st.error("A base foi encontrada, mas não há colunas de mês no padrão 'jan/26 Valor'. Verifique a aba DRE do arquivo consolidado.")
        st.stop()

    default_months = closed_months(all_months)
    selected_months = st.sidebar.multiselect("Meses", all_months, default=default_months)
    if not selected_months:
        selected_months = default_months
    ultimo = selected_months[-1]
    anterior = selected_months[-2] if len(selected_months) >= 2 else None

    m = make_metrics(dre, ultimo)
    m_ant = make_metrics(dre, anterior) if anterior else None
    nao_valor = 0.0
    if not nao.empty:
        val_col = find_col(nao, ["Valor Documento", "Valor", "Valor Total", "_VALOR"])
        if val_col:
            nao_valor = nao[val_col].apply(money_to_float).sum()

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Painel Executivo", "📋 DRE Gerencial", "📈 Evolução Mensal", "🏪 Lojas/Operação", "⚠️ Auditoria DRE"])

    with tab1:
        st.markdown("<div class='section-title'>Painel Executivo</div>", unsafe_allow_html=True)
        cols = st.columns(4)
        cards = [
            ("Receita Bruta", fmt_brl(m["Receita Bruta"]), f"Mês referência: {ultimo}"),
            ("Receita Líquida", fmt_brl(m["Receita Líquida"]), f"{fmt_pct(m['Receita Líquida']/m['Receita Bruta'] if m['Receita Bruta'] else 0)} da receita"),
            ("Lucro Bruto", fmt_brl(m["Lucro Bruto"]), f"Margem {fmt_pct(m['Margem Bruta %'])}"),
            ("EBITDA", fmt_brl(m["EBITDA"]), f"Margem {fmt_pct(m['Margem EBITDA %'])}"),
        ]
        for col, (label, val, sub) in zip(cols, cards):
            with col:
                st.markdown(f"<div class='card'><div class='metric-label'>{label}</div><div class='metric-value'>{val}</div><div class='metric-sub'>{sub}</div></div>", unsafe_allow_html=True)
        cols = st.columns(4)
        cards2 = [
            ("Lucro Líquido", fmt_brl(m["Lucro Líquido"]), f"Margem {fmt_pct(m['Margem Líquida %'])}"),
            ("CMV", fmt_pct(m["CMV %"]), "Custo sobre receita"),
            ("Despesas Operacionais", fmt_pct(m["Despesas Op. %"]), fmt_brl(m["Despesas Operacionais"])),
            ("Não Classificados", f"{len(nao):,}".replace(",", "."), fmt_brl(nao_valor)),
        ]
        for col, (label, val, sub) in zip(cols, cards2):
            with col:
                st.markdown(f"<div class='card'><div class='metric-label'>{label}</div><div class='metric-value'>{val}</div><div class='metric-sub'>{sub}</div></div>", unsafe_allow_html=True)

        st.markdown("<div class='section-title'>Indicadores Gerenciais</div>", unsafe_allow_html=True)
        ind_df = pd.DataFrame({
            "Indicador": ["Margem Bruta", "Margem EBITDA", "Margem Líquida", "CMV sobre Receita", "Despesas Op. sobre Receita", "Deduções sobre Receita"],
            "Atual": [fmt_pct(m["Margem Bruta %"]), fmt_pct(m["Margem EBITDA %"]), fmt_pct(m["Margem Líquida %"]), fmt_pct(m["CMV %"]), fmt_pct(m["Despesas Op. %"]), fmt_pct(m["Deduções %"])],
            "Mês": [ultimo] * 6,
        })
        st.dataframe(ind_df, use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("<div class='section-title'>DRE Gerencial</div>", unsafe_allow_html=True)
        st.caption("Modelo mantido no formato aprovado: meses em colunas, valores e percentuais lado a lado, com todos os blocos estruturais destacados.")
        keep = ["Linha DRE", "_kind", "_style"]
        for mes in selected_months:
            keep += [f"{mes} Valor", f"{mes} %"]
        dre_view = dre[[c for c in keep if c in dre.columns]].copy()
        st.markdown(html_dre_table(dre_view, selected_months), unsafe_allow_html=True)

    with tab3:
        st.markdown("<div class='section-title'>Evolução Mensal</div>", unsafe_allow_html=True)
        line_chart("Receita, Lucro Bruto, EBITDA e Lucro Líquido", selected_months, {
            "Receita Bruta": calc_series(dre, "1. RECEITA OPERACIONAL BRUTA", selected_months),
            "Lucro Bruto": calc_series(dre, "5. (=) LUCRO BRUTO", selected_months),
            "EBITDA": calc_series(dre, "7. (=) RESULTADO", selected_months),
            "Lucro Líquido": calc_series(dre, "11. (=) LUCRO LÍQUIDO", selected_months),
        })
        evo = []
        for mes in selected_months:
            mm = make_metrics(dre, mes)
            evo.append({"Mês": mes, "Receita": fmt_brl(mm["Receita Bruta"]), "Margem Bruta": fmt_pct(mm["Margem Bruta %"]), "Margem EBITDA": fmt_pct(mm["Margem EBITDA %"]), "Margem Líquida": fmt_pct(mm["Margem Líquida %"]), "CMV %": fmt_pct(mm["CMV %"]), "Despesas Op. %": fmt_pct(mm["Despesas Op. %"])})
        st.dataframe(pd.DataFrame(evo), use_container_width=True, hide_index=True)

    with tab4:
        st.markdown("<div class='section-title'>Lojas e Operação</div>", unsafe_allow_html=True)
        if not resumo_loja.empty:
            st.dataframe(resumo_loja, use_container_width=True, height=420)
        else:
            st.info("Resumo por loja ainda não disponível nessa base consolidada. Quando a base de contas trouxer Unidade/Loja, esta tela será alimentada automaticamente.")
        if not estoque.empty:
            st.markdown("<div class='section-title'>Base de Estoque</div>", unsafe_allow_html=True)
            st.dataframe(estoque, use_container_width=True, height=320)

    with tab5:
        st.markdown("<div class='section-title'>Auditoria DRE</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='audit-box'><b>Não classificados fora dos cálculos principais:</b> {len(nao)} linhas • Valor estimado: {fmt_brl(nao_valor)}</div>", unsafe_allow_html=True)
        if not nao.empty:
            st.dataframe(nao, use_container_width=True, height=520)
        else:
            st.success("Nenhuma conta não classificada encontrada.")

except Exception as e:
    st.exception(e)
