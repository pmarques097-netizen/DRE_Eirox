import base64
from pathlib import Path
from datetime import datetime
import re

import pandas as pd
import streamlit as st

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
REV_MONTH = {1:"jan",2:"fev",3:"mar",4:"abr",5:"mai",6:"jun",7:"jul",8:"ago",9:"set",10:"out",11:"nov",12:"dez"}

PESSOAL = ["Alimentação","Pagamento de Férias","FGTS","Salários Fixos + Horas Extras","Comissões e Premiações","INSS","IRRF Folha","Cursos, Treinamentos, Viagens","PPRA<PCMSO e Exames","Rescisão (guia)","Provisão 13 e Férias","Uniformes","Vale Transporte","Pro-Labore","Outras Despesas com Pessoal"]
ADMIN = ["Aluguel&IPTU","Seguro Imóvel","Agua/Luz/Fone/Net","Combustível","Manutenção em Geral","Manutenção Veículos/Motos","Mat. de Escritório/Informática","Mat. De Limpeza","Viagens","Contábil (Terceiros)","Sistemas (Terceiros)","Juridico (Terceiros)","Assessorias/Consultorias/Treinamentos","Taxas, Licenças e Contrib.","Outros (Terceiros)","Outros (Despesas)","Quebra de caixa"]
VENDAS_MKT = ["Marketing/Publicidade","Sistema Fidelidade","Frete","Associação de Classe (Royalties)"]
DEDUCOES = ["ICMS","DAS","ISSQN"]
FINANCEIRO = ["Juros Boletos","Tarifas Bancarias","Taxas de Cartão (MDR)"]
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


def fmt_brl(v):
    try:
        v = float(v)
    except Exception:
        return "R$ 0,00"
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def fmt_pct(v):
    try:
        v = float(v)
    except Exception:
        return "0,00%"
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
    for c in [Path.cwd(), here, here.parent, Path.home() / "Desktop" / "Dre", Path(r"C:\Users\Comercial\Desktop\Dre")]:
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
    for p in [BASE_DIR / "assets" / "logo_eirox.png", BASE_DIR / "assets" / "logo eirox.png", BASE_DIR / "logo_eirox.png", BASE_DIR / "logo eirox.png"]:
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


def parse_dre_sheet_from_excel(path):
    raw = pd.read_excel(path, sheet_name="DRE", header=None)
    header_row = None
    for i in range(min(15, len(raw))):
        if raw.iloc[i].astype(str).str.contains("Linha DRE", case=False, na=False).any():
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()
    month_row = header_row - 1
    months = []
    for col in range(1, raw.shape[1], 2):
        m = raw.iat[month_row, col] if month_row >= 0 else None
        label = month_label_from_value(m)
        if label:
            months.append((col, label))
    rows = []
    style_map = {line: style for line, _kind, style in STRUCTURE}
    kind_map = {line: kind for line, kind, _style in STRUCTURE}
    for r in range(header_row + 1, len(raw)):
        linha = raw.iat[r, 0]
        if pd.isna(linha) or str(linha).strip() == "":
            continue
        linha = str(linha).strip()
        row = {"Linha DRE": linha, "_kind": kind_map.get(linha, "detail"), "_style": style_map.get(linha, "")}
        for col, m in months:
            val = raw.iat[r, col] if col < raw.shape[1] else 0
            pct = raw.iat[r, col+1] if col+1 < raw.shape[1] else 0
            row[f"{m} Valor"] = fmt_brl(money_to_float(val))
            row[f"{m} %"] = fmt_pct(pct)
        rows.append(row)
    return pd.DataFrame(rows)


def load_from_consolidated():
    possible = [BASE_DIR / "data" / "DRE_Consolidado_Moderno.xlsx", BASE_DIR / "DRE_Consolidado_Moderno.xlsx"]
    path = next((p for p in possible if p.exists()), None)
    if not path:
        return pd.DataFrame(), pd.DataFrame()
    xls = pd.ExcelFile(path)
    if "DRE" in xls.sheet_names:
        dre = parse_dre_sheet_from_excel(path)
    else:
        dre = pd.read_excel(path, sheet_name=xls.sheet_names[0])
    nao = pd.read_excel(path, sheet_name="NAO_CLASSIFICADOS") if "NAO_CLASSIFICADOS" in xls.sheet_names else pd.DataFrame()
    return dre, nao


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
    receita_by_month = {}
    if not venda.empty:
        col_total = find_col(venda, ["Total", "Valor", "Valor Total", "Valor Venda"])
        col_mes = find_col(venda, ["Ano-mês", "Ano Mes", "Mês", "Mes", "Data"])
        if col_total:
            venda["_VALOR"] = venda[col_total].apply(money_to_float)
            if col_mes:
                venda["_MES"] = venda[col_mes].apply(month_label_from_value)
            else:
                venda["_MES"] = venda["ARQUIVO_ORIGEM"].apply(infer_month_from_file)
            venda["_MES"] = venda["_MES"].fillna("sem_mes")
            receita_by_month = venda.groupby("_MES")["_VALOR"].sum().to_dict()
    nao = pd.DataFrame()
    class_month = {}
    if not contas.empty:
        c_valor = find_col(contas, ["Valor Documento", "Valor", "Valor Total", "Total"])
        c_plano = find_col(contas, ["Plano de Contas", "Plano Contas", "Conta", "Descrição", "Historico"])
        c_mes = find_col(contas, ["Data Emissão", "Data Pagamento", "Mês", "Mes", "Data"])
        contas["_VALOR"] = contas[c_valor].apply(money_to_float) if c_valor else 0
        contas["_TEXTO"] = contas[c_plano].astype(str) if c_plano else ""
        contas["_CLASS"] = contas["_TEXTO"].apply(classify_by_text)
        if c_mes:
            contas["_MES"] = contas[c_mes].apply(month_label_from_value)
        else:
            contas["_MES"] = contas["ARQUIVO_ORIGEM"].apply(infer_month_from_file)
        contas["_MES"] = contas["_MES"].fillna("sem_mes")
        nao = contas[contas["_CLASS"] == "NAO_CLASSIFICADO"].copy()
        class_month = contas[contas["_CLASS"] != "NAO_CLASSIFICADO"].groupby(["_CLASS","_MES"])["_VALOR"].sum().to_dict()
    months = sorted(set(receita_by_month.keys()) | {m for _,m in class_month.keys()}, key=month_key)
    months = [m for m in months if m and m != "sem_mes"]
    if not months:
        return pd.DataFrame(), nao
    rows = []
    for line, kind, style in STRUCTURE:
        row = {"Linha DRE": line, "_kind": kind, "_style": style}
        for m in months:
            receita = receita_by_month.get(m, 0)
            def val(name): return class_month.get((name,m), 0)
            if line in ["1. RECEITA OPERACIONAL BRUTA", "Receita de Vendas de Mercadorias"]:
                v = receita
            elif line == "2. (-) DEDUÇÕES DA RECEITA BRUTA":
                v = sum(val(x) for x in DEDUCOES)
            elif line == "3. (=) RECEITA OPERACIONAL LÍQUIDA":
                v = receita - sum(val(x) for x in DEDUCOES)
            elif line == "4. (-) CUSTOS DAS VENDAS":
                v = val("Mercadorias (CMP)")
            elif line == "5. (=) LUCRO BRUTO":
                v = receita - sum(val(x) for x in DEDUCOES) - val("Mercadorias (CMP)")
            elif line == "6. (-) DESPESAS OPERACIONAIS":
                v = sum(val(x) for x in PESSOAL + ADMIN + VENDAS_MKT)
            elif line == "Despesas com Pessoal":
                v = sum(val(x) for x in PESSOAL)
            elif line == "Despesas Administrativas e Ocupação":
                v = sum(val(x) for x in ADMIN)
            elif line == "Despesas com Vendas e Marketing":
                v = sum(val(x) for x in VENDAS_MKT)
            elif line == "7. (=) RESULTADO ANTES DO RESULTADO FINANCEIRO (EBITDA/LAJIDA)":
                v = receita - sum(val(x) for x in DEDUCOES) - val("Mercadorias (CMP)") - sum(val(x) for x in PESSOAL + ADMIN + VENDAS_MKT)
            elif line == "8. (+/-) RESULTADO FINANCEIRO LÍQUIDO":
                v = sum(val(x) for x in FINANCEIRO)
            elif line == "9. (=) RESULTADO ANTES DOS TRIBUTOS (LAIR)":
                v = receita - sum(val(x) for x in DEDUCOES) - val("Mercadorias (CMP)") - sum(val(x) for x in PESSOAL + ADMIN + VENDAS_MKT) - sum(val(x) for x in FINANCEIRO)
            elif line == "10. (-) TRIBUTOS SOBRE O LUCRO":
                v = sum(val(x) for x in TRIBUTOS)
            elif line == "11. (=) LUCRO LÍQUIDO DO EXERCÍCIO":
                v = receita - sum(val(x) for x in DEDUCOES) - val("Mercadorias (CMP)") - sum(val(x) for x in PESSOAL + ADMIN + VENDAS_MKT) - sum(val(x) for x in FINANCEIRO) - sum(val(x) for x in TRIBUTOS)
            else:
                v = val(line)
            row[f"{m} Valor"] = fmt_brl(v)
            row[f"{m} %"] = fmt_pct(v / receita if receita else 0)
        rows.append(row)
    return pd.DataFrame(rows), nao


def get_raw_value(dre, line, month):
    col = f"{month} Valor"
    if col not in dre.columns:
        return 0.0
    s = dre.loc[dre["Linha DRE"].astype(str).str.contains(line, case=False, regex=False), col]
    return money_to_float(s.iloc[0]) if not s.empty else 0.0


def style_dre(df):
    source = df.copy()
    show = source.drop(columns=[c for c in ["_kind","_style"] if c in source.columns])
    def row_style(row):
        style = source.loc[row.name, "_style"] if "_style" in source.columns else ""
        if style == "yellow":
            return ["background-color:#FFD11A;color:#06111F;font-weight:900;" for _ in row]
        if style == "blue":
            return ["background-color:#D6E8FB;color:#06111F;font-weight:900;" for _ in row]
        if style == "gray":
            return ["background-color:#2A3342;color:#FFFFFF;font-weight:800;font-style:italic;" for _ in row]
        return ["background-color:#0B1628;color:#F5F7FA;" for _ in row]
    return show.style.apply(row_style, axis=1)


st.markdown("""
<style>
.stApp {background: #06111F; color: #F8FAFC;}
[data-testid="stSidebar"] {background: #050B13;}
h1, h2, h3 {color: #FFFFFF;}
.card {background: linear-gradient(135deg, #0B1628, #101E34);border: 1px solid #1E3A5F;border-radius: 18px;padding: 22px;box-shadow: 0 12px 30px rgba(0,0,0,.30);}
.metric-label {color:#9FB3C8;font-size:14px;font-weight:700;}
.metric-value {color:#FFFFFF;font-size:30px;font-weight:900;margin-top:6px;}
.metric-sub {color:#00AEEF;font-size:13px;margin-top:6px;}
</style>
""", unsafe_allow_html=True)

logo = get_logo()
if logo:
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:22px;justify-content:center;margin-top:14px;">
        <img src="data:image/png;base64,{logo}" style="height:42px;">
        <h1 style="font-size:58px;margin:0;">DRE Empresa Online</h1>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("<h1 style='text-align:center;font-size:58px;'>DRE Empresa Online</h1>", unsafe_allow_html=True)

st.markdown("<p style='text-align:center;color:#7DC8FF;font-size:20px;'>Dashboard financeiro gerencial no padrão Eirox Pricing Online • Atualização automática pelas bases disponíveis</p>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Atualização")
    st.caption("Local: lê direto das pastas. Online: lê as pastas do repositório ou o arquivo data/DRE_Consolidado_Moderno.xlsx.")
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
    st.divider()
    st.caption(f"Modo de leitura: {BASE_MODE}")

try:
    if BASE_MODE == "pastas":
        dre, nao = load_from_folders()
    elif BASE_MODE == "excel":
        dre, nao = load_from_consolidated()
    else:
        dre, nao = pd.DataFrame(), pd.DataFrame()

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

    keep = ["Linha DRE", "_kind", "_style"]
    for m in selected_months:
        keep += [f"{m} Valor", f"{m} %"]
    dre_view = dre[[c for c in keep if c in dre.columns]].copy()

    ultimo = selected_months[-1]
    rec = get_raw_value(dre, "1. RECEITA OPERACIONAL BRUTA", ultimo)
    lb = get_raw_value(dre, "5. (=) LUCRO BRUTO", ultimo)
    ebitda = get_raw_value(dre, "7. (=) RESULTADO", ultimo)
    lucro = get_raw_value(dre, "11. (=) LUCRO LÍQUIDO", ultimo)

    c1,c2,c3,c4 = st.columns(4)
    for col, label, val, sub in [(c1, "Receita Bruta", rec, ultimo),(c2, "Lucro Bruto", lb, fmt_pct(lb/rec if rec else 0)),(c3, "EBITDA", ebitda, fmt_pct(ebitda/rec if rec else 0)),(c4, "Lucro Líquido", lucro, fmt_pct(lucro/rec if rec else 0))]:
        with col:
            st.markdown(f"<div class='card'><div class='metric-label'>{label}</div><div class='metric-value'>{fmt_brl(val)}</div><div class='metric-sub'>{sub}</div></div>", unsafe_allow_html=True)

    st.markdown("## DRE Gerencial")
    st.caption("Modelo mantido no formato aprovado: meses em colunas, valores e percentuais lado a lado, com todos os blocos estruturais destacados.")
    st.dataframe(style_dre(dre_view), use_container_width=True, height=620)

    with st.expander("⚠️ Auditoria de não classificados", expanded=False):
        st.metric("Linhas não classificadas", len(nao))
        if not nao.empty:
            st.dataframe(nao, use_container_width=True, height=360)

except Exception as e:
    st.exception(e)
