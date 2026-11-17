"""
Consolidador de ficheiros (CSV / XLSX / XLS) -> um único Excel ordenado por data.

Funciona com quaisquer colunas:
  - deteta automaticamente separador, decimal, codificação e linha de cabeçalho
  - deteta a coluna de data e ordena da mais antiga para a mais recente
  - junta colunas diferentes (união); células sem dados ficam vazias
  - se os ficheiros forem OMIE IndexHourValues, acrescenta resumos de preços

Local:  streamlit run app.py
"""
import csv
import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st

MAX_LINHAS_FOLHA = 1_048_000
NOMES_DATA = re.compile(r"(^|[^a-z])(date|data|fecha|dia|day|datetime|timestamp)([^a-z]|$)", re.I)


# =========================================================================== leitura
def decodificar(raw: bytes) -> str:
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


def detetar_separador(txt: str) -> str:
    amostra = "\n".join(txt.splitlines()[:50])
    try:
        return csv.Sniffer().sniff(amostra, delimiters=";,\t|").delimiter
    except csv.Error:
        cont = {s: amostra.count(s) for s in (";", ",", "\t", "|")}
        return max(cont, key=cont.get)


def ler_bruto(nome: str, raw: bytes, todas_folhas: bool):
    """Devolve lista de (folha, DataFrame sem cabeçalho, info)."""
    ext = nome.lower().rsplit(".", 1)[-1]
    if ext in ("csv", "txt"):
        txt = decodificar(raw)
        sep = detetar_separador(txt)
        df = pd.read_csv(io.StringIO(txt), sep=sep, header=None, dtype=str,
                         keep_default_na=False, engine="python", on_bad_lines="skip")
        return [("", df.replace("", None), f"sep='{sep}'")]
    engine = "xlrd" if ext == "xls" else None
    folhas = pd.read_excel(io.BytesIO(raw), sheet_name=None if todas_folhas else 0,
                           header=None, dtype=object, engine=engine)
    if isinstance(folhas, pd.DataFrame):
        folhas = {"": folhas}
    if ext in ("xlsx", "xlsm"):
        preencher_unidas(raw, folhas)
    return [(f, d, "") for f, d in folhas.items() if not d.dropna(how="all").empty]


def preencher_unidas(raw: bytes, folhas: dict):
    """Repete o valor das células unidas (merge) nas primeiras linhas, para cabeçalhos de vários níveis."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=False, data_only=True)
    except Exception:
        return
    nomes = list(folhas)
    for idx, nome in enumerate(nomes):
        ws = wb[nome] if nome in wb.sheetnames else wb.worksheets[idx]
        df = folhas[nome]
        for r in ws.merged_cells.ranges:
            if r.min_row > 60:
                continue
            v = df.iat[r.min_row - 1, r.min_col - 1] if r.min_row - 1 < len(df) \
                and r.min_col - 1 < df.shape[1] else None
            for i in range(r.min_row - 1, min(r.max_row, len(df))):
                for j in range(r.min_col - 1, min(r.max_col, df.shape[1])):
                    df.iat[i, j] = v


def e_texto(v) -> bool:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    if isinstance(v, (int, float, datetime, pd.Timestamp)):
        return False
    s = str(v).strip()
    return bool(s) and not re.fullmatch(r"[-+]?[\d.,\s%]+", s) and not parece_data(s)


def parece_data(s) -> bool:
    if isinstance(s, (datetime, pd.Timestamp)):
        return True
    s = str(s).strip()
    return bool(re.fullmatch(r"\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}([ T].*)?", s))


def detetar_cabecalho(df: pd.DataFrame, manual: int | None):
    """Índice da linha de cabeçalho e nº de linhas de cabeçalho a juntar."""
    if manual is not None:
        return max(manual - 1, 0), 1
    n = min(len(df), 60)
    largura = df.iloc[:n].notna().sum(axis=1)
    larg_ref = largura.max()
    # primeira linha "de dados": larga e com pelo menos metade das células não-texto
    def e_dados(i):
        linha = df.iloc[i].dropna()
        if len(linha) < max(2, 0.3 * larg_ref):
            return False
        return sum(not e_texto(v) for v in linha) >= 0.5 * len(linha)

    inicio = None
    for i in range(n):
        seguintes = [j for j in range(i, min(i + 3, len(df)))]
        if all(e_dados(j) for j in seguintes):
            inicio = i
            break
    if inicio is None or inicio == 0:
        return 0, 1
    # linhas de cabeçalho contíguas imediatamente acima (só texto)
    k, j = 0, inicio - 1
    while j >= 0 and k < 3:
        linha = df.iloc[j].dropna()
        if len(linha) == 0 or sum(e_texto(v) for v in linha) < 0.8 * len(linha):
            break
        k += 1
        j -= 1
    return inicio - 1, k


def construir_tabela(df: pd.DataFrame, manual: int | None):
    hdr, k = detetar_cabecalho(df, manual)
    if k == 0:
        cols = [f"Col_{i + 1}" for i in range(df.shape[1])]
        corpo = df.iloc[hdr + 1:]
    else:
        topo = df.iloc[hdr - k + 1: hdr + 1].copy()
        cols = []
        for c in range(df.shape[1]):
            partes = [str(v).strip() for v in topo.iloc[:, c] if v is not None and not pd.isna(v)
                      and str(v).strip()]
            # evita repetir o mesmo texto entre níveis
            limpo = []
            for p in partes:
                if p not in limpo:
                    limpo.append(p)
            cols.append(" | ".join(limpo) if limpo else f"Col_{c + 1}")
        corpo = df.iloc[hdr + 1:]
    cols = nomes_unicos(cols)
    corpo = corpo.copy()
    corpo.columns = cols
    corpo = corpo.dropna(how="all")
    corpo = corpo.loc[:, ~(corpo.isna().all() & pd.Series([c.startswith("Col_") for c in cols],
                                                           index=cols))]
    return corpo.reset_index(drop=True), hdr + 1


def nomes_unicos(cols):
    vistos, out = {}, []
    for c in cols:
        c = re.sub(r"\s+", " ", str(c)).strip()
        if c in vistos:
            vistos[c] += 1
            out.append(f"{c}_{vistos[c]}")
        else:
            vistos[c] = 0
            out.append(c)
    return out


# =========================================================================== tipos
def para_numero(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s
    t = s.astype(str).str.strip().str.replace(" ", "").str.replace(" ", "")
    ptbr = t.str.contains(r",\d+$", regex=True) & ~t.str.contains(r"\.\d+$", regex=True)
    t = t.where(~ptbr, t.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    t = t.str.replace(r",(?=\d{3}(\D|$))", "", regex=True)  # separador de milhares inglês
    return pd.to_numeric(t, errors="coerce")


def converter_tipos(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_numeric_dtype(s) or pd.api.types.is_datetime64_any_dtype(s):
            continue
        nn = s.dropna()
        if nn.empty:
            continue
        num = para_numero(nn)
        if num.notna().mean() >= 0.8:
            df[c] = para_numero(s)  # 'n.a.' e afins ficam vazios
    return df


def parse_datas(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    if pd.api.types.is_numeric_dtype(s):
        # números de série Excel plausíveis (1990-2100)
        ok = s.between(32874, 73051)
        return pd.to_datetime(s.where(ok), unit="D", origin="1899-12-30", errors="coerce")
    t = s.astype(str).str.strip()
    iso = t.str.match(r"^\d{4}-\d{1,2}-\d{1,2}").mean() > 0.5
    if iso:
        r = pd.to_datetime(t.str[:19], errors="coerce", format="mixed")
    else:
        r = pd.to_datetime(t, errors="coerce", dayfirst=True, format="mixed")
    return r


def detetar_coluna_data(df: pd.DataFrame):
    melhor, melhor_score = None, 0.0
    for c in df.columns:
        s = df[c].dropna()
        if s.empty:
            continue
        if pd.api.types.is_numeric_dtype(s) and not NOMES_DATA.search(str(c)):
            continue
        amostra = s.head(500)
        taxa = parse_datas(amostra).notna().mean()
        score = taxa + (0.3 if NOMES_DATA.search(str(c)) else 0)
        if taxa >= 0.8 and score > melhor_score:
            melhor, melhor_score = c, score
    return melhor


# =========================================================================== OMIE extra
COLS_OMIE = {"Date", "Area", "Period", "Period Start", "Value", "Volume"}


def resumos_omie(df: pd.DataFrame) -> dict:
    d = df[["Date", "Area", "Period", "Period Start", "Value", "Volume"]].copy()
    d["Date"] = parse_datas(d["Date"])
    d = d.dropna(subset=["Date", "Area", "Period"])
    d["Hora"] = pd.to_numeric(d["Period Start"].astype(str).str[:2], errors="coerce")
    h = d.groupby(["Date", "Area", "Hora"], as_index=False).agg(
        Preco=("Value", "mean"), Volume=("Volume", "sum"))

    def stats(g):
        p = g["Preco"].sort_values()
        sol = g[(g["Hora"] >= 10) & (g["Hora"] < 17)]
        vol = g["Volume"].sum()
        return pd.Series({
            "Base media (EUR/MWh)": g["Preco"].mean(),
            "Ponderado volume (EUR/MWh)": (g["Preco"] * g["Volume"]).sum() / vol if vol else None,
            "Solar 10h-17h (EUR/MWh)": sol["Preco"].mean() if len(sol) else None,
            "Minimo (EUR/MWh)": p.min(), "Maximo (EUR/MWh)": p.max(),
            "Spread max-min": p.max() - p.min(),
            "Spread 2h (top2-bottom2)": p.tail(2).mean() - p.head(2).mean(),
            "Spread 4h (top4-bottom4)": p.tail(4).mean() - p.head(4).mean(),
            "Horas preco<=0": int((g["Preco"] <= 0).sum()),
            "Volume total (MWh)": vol,
        })

    rd = h.groupby(["Date", "Area"]).apply(stats, include_groups=False).reset_index()
    num = [c for c in rd.columns if c not in ("Date", "Area")]
    rd["Mes"] = rd["Date"].dt.to_period("M").astype(str)
    g = rd.groupby(["Mes", "Area"])
    rm = g[num].mean().reset_index()
    rm.insert(2, "Dias", g.size().values)
    rm["Horas preco<=0"] = g["Horas preco<=0"].sum().values
    rm["Volume total (MWh)"] = g["Volume total (MWh)"].sum().values
    out = {"Resumo_Mensal": rm.round({c: 2 for c in num}),
           "Resumo_Diario": rd.drop(columns="Mes").round({c: 2 for c in num})}
    for a in sorted(d["Area"].astype(str).unique()):
        pv = d[d["Area"] == a].pivot_table(index="Date", columns="Period", values="Value",
                                           aggfunc="last")
        pv.columns = [f"P{int(c):02d}" for c in pv.columns]
        out[f"Preco_{a}"[:31]] = pv.reset_index()
    return out


# =========================================================================== excel
def gerar_excel(folhas: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm") as w:
        for nome, df in folhas.items():
            if len(df) > MAX_LINHAS_FOLHA:
                for i in range(0, len(df), MAX_LINHAS_FOLHA):
                    df.iloc[i:i + MAX_LINHAS_FOLHA].to_excel(
                        w, sheet_name=f"{nome}_{i // MAX_LINHAS_FOLHA + 1}"[:31], index=False)
            else:
                df.to_excel(w, sheet_name=nome[:31], index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for col in ws.iter_cols(max_row=min(ws.max_row, 200)):
                larg = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(9, larg + 2), 40)
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    if isinstance(c.value, datetime) and c.value.hour == 0 and c.value.minute == 0:
                        c.number_format = "yyyy-mm-dd"
    return buf.getvalue()


# =========================================================================== pipeline
def processar(ficheiros, todas_folhas=False, remover_dup=True, juntar_maiusc=True,
              incluir_origem=True, hdr_manual=None, prog=None):
    """ficheiros: lista de (nome, bytes). Devolve (df, folhas, log, n_bruto, tem_data)."""
    partes, log = [], []
    for i, f in enumerate(sorted(ficheiros, key=lambda x: x[0])):
        try:
            for folha, bruto, info in ler_bruto(f[0], f[1], todas_folhas):
                tab, linha_hdr = construir_tabela(bruto, hdr_manual)
                tab = converter_tipos(tab)
                cdata = detetar_coluna_data(tab)
                if cdata:
                    tab[cdata] = parse_datas(tab[cdata])
                    tab = tab.rename(columns={cdata: "__DATA__"}).assign(__NOMEDATA__=cdata)
                if incluir_origem:
                    tab["Ficheiro_origem"] = f[0]
                    if folha:
                        tab["Folha_origem"] = folha
                tab["__ORD__"] = range(len(tab))
                tab["__FICH__"] = i
                partes.append(tab)
                dmin = tab["__DATA__"].min() if cdata else None
                dmax = tab["__DATA__"].max() if cdata else None
                log.append({"Ficheiro": f[0], "Folha": folha, "Linhas": len(tab),
                            "Linha cabeçalho": linha_hdr, "Coluna de data": cdata or "(não detetada)",
                            "Data mín": dmin, "Data máx": dmax, "Info": info, "Erro": ""})
        except Exception as e:
            log.append({"Ficheiro": f[0], "Folha": "", "Linhas": 0, "Linha cabeçalho": None,
                        "Coluna de data": "", "Data mín": None, "Data máx": None, "Info": "",
                        "Erro": str(e)})
        if prog: prog((i + 1) / len(ficheiros))

    if not partes:
        return None, None, log, 0, False

    # nome da coluna de data final: o mais frequente
    nomes_data = [p["__NOMEDATA__"].iloc[0] for p in partes if "__NOMEDATA__" in p and len(p)]
    nome_data = max(set(nomes_data), key=nomes_data.count) if nomes_data else None

    if juntar_maiusc:
        canon = {}
        for p in partes:
            for c in p.columns:
                canon.setdefault(c.lower(), c)
        partes = [p.rename(columns={c: canon[c.lower()] for c in p.columns}) for p in partes]

    # mapa de colunas por ficheiro
    todas_cols = []
    for p in partes:
        for c in p.columns:
            if not c.startswith("__") and c not in todas_cols:
                todas_cols.append(c)
    mapa = pd.DataFrame([{**{"Ficheiro": l["Ficheiro"], "Folha": l["Folha"]},
                          **{c: ("✓" if c in p.columns else "") for c in todas_cols}}
                         for l, p in zip([l for l in log if not l["Erro"]], partes)])

    df = pd.concat(partes, ignore_index=True, sort=False)
    df = df.drop(columns=["__NOMEDATA__"], errors="ignore")
    n_bruto = len(df)
    tem_data = "__DATA__" in df.columns
    if remover_dup:
        chave = [c for c in df.columns if (c == "__DATA__" or not c.startswith("__"))
                 and c not in ("Ficheiro_origem", "Folha_origem")]
        df = df.drop_duplicates(subset=chave, keep="last")
    ordem = (["__DATA__"] if tem_data else []) + ["__FICH__", "__ORD__"]
    df = df.sort_values(ordem, kind="mergesort", na_position="last")
    if tem_data:
        alvo = nome_data if nome_data not in df.columns else f"{nome_data} (data)"
        df = df.rename(columns={"__DATA__": alvo})
        cols = [alvo] + [c for c in df.columns if c != alvo]
        df = df[cols]
    df = df.drop(columns=["__FICH__", "__ORD__"]).reset_index(drop=True)

    folhas = {"Consolidado": df}
    if COLS_OMIE.issubset(df.columns):
        try:
            folhas.update(resumos_omie(df))
        except Exception as e:
            log.append({"Ficheiro": "(resumos OMIE)", "Erro": str(e)})
    folhas["Ficheiros"] = pd.DataFrame(log)
    folhas["Colunas_por_ficheiro"] = mapa


    return df, folhas, log, n_bruto, tem_data


# =========================================================================== UI
CSS = """
<style>
:root{
  --glass-bg: rgba(255,255,255,.55);
  --glass-strong: rgba(255,255,255,.72);
  --glass-edge: rgba(255,255,255,.75);
  --glass-line: rgba(255,255,255,.35);
  --ink: #1d1d1f; --ink-2: #6e6e73;
  --accent: #0a84ff; --accent-2: #5e5ce6; --ok: #30d158; --warn: #ff9f0a; --err: #ff453a;
  --shadow: 0 10px 40px rgba(31,38,135,.12), 0 2px 6px rgba(0,0,0,.05);
  --radius: 22px;
}
@media (prefers-color-scheme: dark){
  :root{ --glass-bg: rgba(40,40,48,.50); --glass-strong: rgba(44,44,52,.70);
         --glass-edge: rgba(255,255,255,.22); --glass-line: rgba(255,255,255,.10);
         --ink:#f5f5f7; --ink-2:#a1a1a6; --shadow: 0 10px 40px rgba(0,0,0,.45); }
}
html, body, [class*="css"], .stApp { font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text",
  "Segoe UI Variable", "Segoe UI", system-ui, sans-serif; color: var(--ink); }

/* Fundo com cor para o vidro ter algo para refratar */
[data-testid="stAppViewContainer"]{
  background:
    radial-gradient(900px 600px at 8% 0%, rgba(10,132,255,.28), transparent 60%),
    radial-gradient(800px 600px at 95% 10%, rgba(94,92,230,.25), transparent 60%),
    radial-gradient(900px 700px at 60% 105%, rgba(48,209,88,.20), transparent 60%),
    linear-gradient(180deg, #f5f7fb 0%, #eef1f7 100%);
  background-attachment: fixed;
}
@media (prefers-color-scheme: dark){
  [data-testid="stAppViewContainer"]{
    background:
      radial-gradient(900px 600px at 8% 0%, rgba(10,132,255,.30), transparent 60%),
      radial-gradient(800px 600px at 95% 10%, rgba(94,92,230,.28), transparent 60%),
      radial-gradient(900px 700px at 60% 105%, rgba(48,209,88,.16), transparent 60%),
      linear-gradient(180deg, #0b0c10 0%, #121318 100%);
  }
}
[data-testid="stHeader"]{ background: transparent; }
.block-container{ max-width: 1120px; padding-top: 2.2rem; }

/* Material de vidro: camada clara para elementos de conteúdo */
.glass, [class*="st-key-glass"]{
  padding: 22px 24px !important;
  background: var(--glass-bg) !important;
  -webkit-backdrop-filter: blur(24px) saturate(180%); backdrop-filter: blur(24px) saturate(180%);
  border: 1px solid var(--glass-line) !important;
  border-top-color: var(--glass-edge) !important;          /* luz a bater na aresta */
  border-radius: var(--radius) !important;
  box-shadow: var(--shadow), inset 0 1px 0 rgba(255,255,255,.45);
  animation: materialize .45s cubic-bezier(.2,.8,.2,1) both;
}
@keyframes materialize{ from{ opacity:0; transform: scale(.985); filter: blur(6px);} to{ opacity:1; transform:none; filter:none;} }

/* Barra lateral: material mais pesado (estrutura) */
[data-testid="stSidebar"]{
  background: var(--glass-strong) !important;
  -webkit-backdrop-filter: blur(30px) saturate(160%); backdrop-filter: blur(30px) saturate(160%);
  border-right: 1px solid var(--glass-line);
}
[data-testid="stSidebar"] h2{ font-size: 1.05rem; letter-spacing: -.01em; }

/* Hero */
.hero{ padding: 28px 32px; margin-bottom: 18px; display:flex; align-items:center; gap:22px; }
.hero .icon{ width:64px; height:64px; border-radius:18px; flex:none; display:grid; place-items:center;
  background: linear-gradient(145deg, var(--accent), var(--accent-2)); color:#fff; font-size:30px;
  box-shadow: 0 8px 24px rgba(10,132,255,.35), inset 0 1px 0 rgba(255,255,255,.5); }
.hero h1{ margin:0; font-size: clamp(1.8rem, 3.4vw, 2.6rem); line-height:1.05; letter-spacing:-.025em; font-weight:700; }
.hero p{ margin:.35rem 0 0; color: var(--ink-2); font-size: 1rem; letter-spacing: -.005em; }

/* Passos */
.steps{ display:flex; gap:10px; flex-wrap:wrap; margin: 4px 0 18px; }
.step{ display:flex; align-items:center; gap:8px; padding:8px 14px; border-radius: 999px; font-size:.86rem;
  color: var(--ink-2); background: var(--glass-bg); border:1px solid var(--glass-line);
  -webkit-backdrop-filter: blur(16px); backdrop-filter: blur(16px); transition: all .3s ease; }
.step b{ width:22px; height:22px; border-radius:50%; display:grid; place-items:center; font-size:.75rem;
  background: rgba(120,120,128,.18); color: var(--ink-2); }
.step.on{ color: var(--ink); border-color: rgba(10,132,255,.35); }
.step.on b{ background: var(--accent); color:#fff; }
.step.done b{ background: var(--ok); color:#fff; }

.sec-title{ font-weight:650; font-size:1.08rem; letter-spacing:-.012em; margin: 0 0 .2rem; }
.sec-sub{ color: var(--ink-2); font-size:.88rem; margin: 0 0 .8rem; }

/* Zona de upload */
[data-testid="stFileUploaderDropzone"]{
  background: rgba(255,255,255,.28) !important; border: 1.5px dashed rgba(10,132,255,.45) !important;
  border-radius: 18px !important; padding: 26px !important; transition: background .2s ease, transform .2s ease;
}
[data-testid="stFileUploaderDropzone"]:hover{ background: rgba(10,132,255,.08) !important; }
@media (prefers-color-scheme: dark){ [data-testid="stFileUploaderDropzone"]{ background: rgba(255,255,255,.04) !important; } }

/* Botões em cápsula, feedback ao pressionar */
.stButton > button, [data-testid="stDownloadButton"] > button{
  border-radius: 999px !important; padding: .62rem 1.5rem !important; font-weight: 600 !important;
  letter-spacing: -.005em; border: 1px solid var(--glass-line) !important;
  background: var(--glass-strong) !important; color: var(--ink) !important;
  -webkit-backdrop-filter: blur(20px); backdrop-filter: blur(20px);
  box-shadow: 0 4px 14px rgba(0,0,0,.06), inset 0 1px 0 rgba(255,255,255,.55);
  transition: transform .1s ease-out, box-shadow .2s ease, filter .2s ease;
}
.stButton > button[kind="primary"], [data-testid="stDownloadButton"] > button[kind="primary"]{
  background: linear-gradient(180deg, #2b95ff 0%, var(--accent) 100%) !important; color:#fff !important;
  border-color: rgba(255,255,255,.25) !important;
  box-shadow: 0 8px 22px rgba(10,132,255,.35), inset 0 1px 0 rgba(255,255,255,.45);
}
.stButton > button:hover, [data-testid="stDownloadButton"] > button:hover{ filter: brightness(1.04); }
.stButton > button:active, [data-testid="stDownloadButton"] > button:active{ transform: scale(.97); }

/* Métricas como chips de vidro */
[data-testid="stMetric"]{
  background: var(--glass-bg); border:1px solid var(--glass-line); border-top-color: var(--glass-edge);
  -webkit-backdrop-filter: blur(20px) saturate(170%); backdrop-filter: blur(20px) saturate(170%);
  border-radius: 18px; padding: 14px 18px; box-shadow: 0 4px 16px rgba(0,0,0,.05);
}
[data-testid="stMetricLabel"] p{ color: var(--ink-2) !important; font-size:.8rem !important; letter-spacing:.01em; }
[data-testid="stMetricValue"]{ font-size: 1.75rem !important; letter-spacing: -.02em; font-weight: 650; }

/* Alertas, expanders, tabelas */
[data-testid="stAlert"]{ border-radius: 16px !important; -webkit-backdrop-filter: blur(16px); backdrop-filter: blur(16px); }
[data-testid="stExpander"] details{ border-radius: 16px !important; background: var(--glass-bg);
  border:1px solid var(--glass-line) !important; -webkit-backdrop-filter: blur(16px); backdrop-filter: blur(16px); }
[data-testid="stDataFrame"]{ border-radius: 14px; overflow: hidden; }
.stProgress > div > div > div > div{ background: linear-gradient(90deg, var(--accent), var(--accent-2)); }

.foot{ text-align:center; color: var(--ink-2); font-size:.78rem; margin-top: 28px; }

/* Acessibilidade */
@media (prefers-reduced-motion: reduce){
  *{ animation: none !important; transition: opacity .15s ease !important; }
  .stButton > button:active, [data-testid="stDownloadButton"] > button:active{ transform:none; }
}
@media (prefers-reduced-transparency: reduce){
  .glass, [data-testid="stSidebar"], [data-testid="stMetric"], .step,
  [class*="st-key-glass"]{
    -webkit-backdrop-filter:none !important; backdrop-filter:none !important; background: #fff !important; }
}
@media (prefers-contrast: more){
  .glass, [data-testid="stMetric"],
  [class*="st-key-glass"]{
    background: #fff !important; border: 1.5px solid #1d1d1f !important; }
}
</style>
"""

st.set_page_config(page_title="Consolidador", page_icon="📊", layout="wide")
st.markdown(CSS, unsafe_allow_html=True)

st.markdown("""
<div class="glass hero">
  <div class="icon">⧉</div>
  <div>
    <h1>Consolidador de ficheiros</h1>
    <p>CSV, XLSX ou XLS com quaisquer colunas, num único Excel ordenado por data.
       Processado em memória, nada fica guardado.</p>
  </div>
</div>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## Opções")
    todas_folhas = st.toggle("Ler todas as folhas dos Excel", value=False,
                             help="Desligado: lê só a primeira folha de cada ficheiro.")
    remover_dup = st.toggle("Remover linhas duplicadas", value=True,
                            help="Linhas iguais em todas as colunas de dados (ignora a origem).")
    juntar_maiusc = st.toggle("Juntar colunas ignorando maiúsculas", value=True)
    incluir_origem = st.toggle("Colunas de origem (ficheiro/folha)", value=True)
    hdr_auto = st.toggle("Detetar cabeçalho automaticamente", value=True)
    hdr_manual = None if hdr_auto else st.number_input("Linha do cabeçalho (1 = primeira)",
                                                       min_value=1, value=1, step=1)
    st.divider()
    nome_saida = st.text_input("Nome do Excel",
                               value=f"consolidado_{datetime.now():%Y%m%d_%H%M}.xlsx")

res = st.session_state.get("res")
passo = 3 if res else 2 if st.session_state.get("tem_fich") else 1


def chip(n, txt):
    cls = "done" if n < passo else "on" if n == passo else ""
    return f'<div class="step {cls}"><b>{"✓" if n < passo else n}</b>{txt}</div>'


st.markdown('<div class="steps">' + chip(1, "Carregar ficheiros") + chip(2, "Consolidar")
            + chip(3, "Descarregar Excel") + "</div>", unsafe_allow_html=True)

with st.container(border=True, key="glass_fich"):
    st.markdown('<p class="sec-title">Ficheiros</p><p class="sec-sub">Arrasta vários ficheiros de '
                'uma vez. Podem ter colunas e formatos diferentes.</p>', unsafe_allow_html=True)
    ficheiros = st.file_uploader("Ficheiros", accept_multiple_files=True, label_visibility="collapsed",
                                 type=["csv", "txt", "xlsx", "xlsm", "xls"])
    st.session_state["tem_fich"] = bool(ficheiros)
    assinatura = tuple((f.name, f.size) for f in ficheiros) if ficheiros else ()
    if res and res["assinatura"] != assinatura:
        st.session_state.pop("res", None)
        res = None
    col_b, col_i = st.columns([1, 3], vertical_alignment="center")
    correr = col_b.button("Consolidar", type="primary", disabled=not ficheiros,
                          use_container_width=True)
    if ficheiros:
        tot = sum(f.size for f in ficheiros) / 1024
        col_i.caption(f"{len(ficheiros)} ficheiro(s) · {tot:,.0f} KB".replace(",", " "))

if correr and ficheiros:
    barra = st.progress(0.0, text="A processar…")
    df, folhas, log, n_bruto, tem_data = processar(
        [(f.name, f.getvalue()) for f in ficheiros], todas_folhas, remover_dup, juntar_maiusc,
        incluir_origem, hdr_manual, prog=lambda v: barra.progress(v, text="A processar…"))
    barra.empty()
    if df is None:
        st.error("Nenhum ficheiro válido.")
        st.dataframe(pd.DataFrame(log), use_container_width=True)
        st.stop()
    st.session_state["res"] = {
        "assinatura": assinatura, "df": df, "log": log, "n_bruto": n_bruto, "tem_data": tem_data,
        "mapa": folhas["Colunas_por_ficheiro"], "xls": gerar_excel(folhas),
        "folhas": [k for k in folhas],
    }
    st.rerun()

if res:
    df, log = res["df"], res["log"]
    erros = [l for l in log if l.get("Erro")]
    with st.container(border=True, key="glass_res"):
        st.markdown('<p class="sec-title">Resultado</p>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ficheiros válidos", f"{len(log) - len(erros)}/{len(log)}")
        c2.metric("Linhas", f"{len(df):,}".replace(",", " "))
        c3.metric("Duplicados removidos", f"{res['n_bruto'] - len(df):,}".replace(",", " "))
        c4.metric("Colunas", len(df.columns))
        if res["tem_data"]:
            dcol = df.columns[0]
            st.success(f"Ordenado por **{dcol}**, de {df[dcol].min():%d/%m/%Y} "
                       f"a {df[dcol].max():%d/%m/%Y}.")
        else:
            st.warning("Sem coluna de data detetada. Mantida a ordem por ficheiro e linha.")
        if erros:
            st.error("Ficheiros com erro: " + "; ".join(f"{e['Ficheiro']}: {e['Erro']}" for e in erros))
        nome = nome_saida if nome_saida.lower().endswith(".xlsx") else nome_saida + ".xlsx"
        d1, d2 = st.columns([1, 3], vertical_alignment="center")
        d1.download_button("Descarregar Excel", data=res["xls"], file_name=nome, type="primary",
                           use_container_width=True, on_click="ignore",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.caption("Folhas: " + " · ".join(res["folhas"]))

    with st.container(border=True, key="glass_prev"):
        st.markdown('<p class="sec-title">Pré-visualização</p><p class="sec-sub">Primeiras 200 linhas.'
                    '</p>', unsafe_allow_html=True)
        st.dataframe(df.head(200), use_container_width=True, height=380)
        with st.expander("Registo de ficheiros"):
            st.dataframe(pd.DataFrame(log), use_container_width=True)
        with st.expander("Colunas por ficheiro"):
            st.dataframe(res["mapa"], use_container_width=True)

st.markdown('<p class="foot">Os ficheiros são processados em memória e descartados no fim da sessão.</p>',
            unsafe_allow_html=True)
