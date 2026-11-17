"""
Consolidação de Notas de Liquidação Mensal (NL_*.xlsx)

Lê as notas geradas pelo representante de mercado e calcula os valores
a partir dos dados brutos (Produção A:C e constantes do Resumo).
Não depende de valores em cache das fórmulas, por isso funciona com
ficheiros gerados por openpyxl que nunca foram abertos no Excel.
"""

import io
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Notas de Liquidação", layout="wide")


# ---------------------------------------------------------------------------
# Leitura de uma nota
# ---------------------------------------------------------------------------
def _num(v, default=0.0):
    """Converte para float; fórmulas ou vazios devolvem o valor por defeito."""
    if v is None or (isinstance(v, str) and v.startswith("=")):
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _data(v):
    if isinstance(v, datetime):
        return v
    return datetime.strptime(str(v).strip(), "%d-%m-%Y")


def ler_nota(ficheiro) -> dict:
    conteudo = ficheiro.getvalue()

    wb = openpyxl.load_workbook(io.BytesIO(conteudo), data_only=False, read_only=True)
    r = wb["Resumo"]
    cliente = r["C4"].value
    cpe = r["C5"].value
    inicio = _data(r["C8"].value)
    fim = _data(r["C9"].value)
    vol_fixo = _num(r["C14"].value, 0.0)          # fração remunerada a preço fixo
    vol_omie = 1.0 - vol_fixo                      # C15 = 1 - C14
    preco_fixo = _num(r["C21"].value, 0.0)         # €/MWh
    fee = _num(r["C22"].value, 0.0)                # €/MWh
    wb.close()

    df = pd.read_excel(io.BytesIO(conteudo), sheet_name="Produção", usecols="A:C")
    df.columns = ["data", "mwh", "preco"]
    df = df.dropna(subset=["data"])
    df["mwh"] = pd.to_numeric(df["mwh"], errors="coerce").fillna(0.0)
    df["preco"] = pd.to_numeric(df["preco"], errors="coerce").fillna(0.0)

    energia = df["mwh"].sum()
    valor_omie = (df["mwh"] * df["preco"] * vol_omie).sum()          # coluna D
    valor_fixo = (df["mwh"] * vol_fixo * preco_fixo).sum()           # coluna E
    custo_fee = (df["mwh"] * vol_omie * fee).sum()                   # coluna F
    total = valor_omie + valor_fixo - custo_fee

    # Controlo de qualidade: n.º de quartos de hora esperado no período
    dias = (fim - inicio).days + 1
    esperado = dias * 96
    intervalos = len(df)

    return {
        "Ficheiro": ficheiro.name,
        "Cliente": cliente,
        "CPE": cpe,
        "Mês": inicio.strftime("%Y-%m"),
        "Início": inicio.date(),
        "Fim": fim.date(),
        "Energia (MWh)": energia,
        "Vol. OMIE (%)": vol_omie * 100,
        "Valor OMIE (€)": valor_omie,
        "Valor Fixo (€)": valor_fixo,
        "Fee gestão (€)": custo_fee,
        "Total s/ IVA (€)": total,
        "Preço capturado (€/MWh)": valor_omie / (energia * vol_omie) if energia * vol_omie else None,
        "Preço médio OMIE (€/MWh)": df["preco"].mean(),
        "Intervalos": intervalos,
        "Intervalos esperados": esperado,
    }


# ---------------------------------------------------------------------------
# Excel de saída
# ---------------------------------------------------------------------------
def gerar_excel(detalhe: pd.DataFrame, por_cpe: pd.DataFrame, por_mes: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        detalhe.to_excel(xw, sheet_name="Detalhe", index=False)
        por_cpe.to_excel(xw, sheet_name="Por CPE", index=False)
        por_mes.to_excel(xw, sheet_name="Por mês", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                largura = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(largura + 2, 10), 45)
                for c in col[1:]:
                    if isinstance(c.value, float):
                        c.number_format = "#,##0.00"
            ws.freeze_panes = "A2"
    return buf.getvalue()


def agregar(df: pd.DataFrame, chave: list) -> pd.DataFrame:
    g = df.groupby(chave, as_index=False)[
        ["Energia (MWh)", "Valor OMIE (€)", "Valor Fixo (€)", "Fee gestão (€)", "Total s/ IVA (€)"]
    ].sum()
    g["Preço líquido (€/MWh)"] = g["Total s/ IVA (€)"] / g["Energia (MWh)"].where(g["Energia (MWh)"] != 0)
    return g


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------
st.title("Notas de Liquidação Mensal")
st.caption("Carregue os ficheiros NL_*.xlsx. Os valores são recalculados a partir dos dados de produção e preço.")

ficheiros = st.file_uploader(
    "Notas de liquidação (.xlsx)", type="xlsx", accept_multiple_files=True
)

if not ficheiros:
    st.stop()

linhas, erros = [], []
progresso = st.progress(0.0, text="A processar...")
for i, f in enumerate(ficheiros, start=1):
    try:
        linhas.append(ler_nota(f))
    except Exception as e:  # um ficheiro com erro não bloqueia os restantes
        erros.append(f"{f.name}: {type(e).__name__}: {e}")
    progresso.progress(i / len(ficheiros), text=f"A processar {i}/{len(ficheiros)}")
progresso.empty()

if erros:
    with st.expander(f"{len(erros)} ficheiro(s) com erro", expanded=True):
        for e in erros:
            st.error(e)

if not linhas:
    st.warning("Nenhum ficheiro foi processado.")
    st.stop()

detalhe = pd.DataFrame(linhas).sort_values(["CPE", "Mês"]).reset_index(drop=True)
por_cpe = agregar(detalhe, ["Cliente", "CPE"])
por_mes = agregar(detalhe, ["Mês"])

# Indicadores
c1, c2, c3, c4 = st.columns(4)
c1.metric("Ficheiros processados", f"{len(detalhe)} / {len(ficheiros)}")
c2.metric("Energia total", f"{detalhe['Energia (MWh)'].sum():,.1f} MWh".replace(",", " "))
c3.metric("Total s/ IVA", f"{detalhe['Total s/ IVA (€)'].sum():,.0f} €".replace(",", " "))
energia_total = detalhe["Energia (MWh)"].sum()
c4.metric(
    "Preço líquido médio",
    f"{detalhe['Total s/ IVA (€)'].sum() / energia_total:,.2f} €/MWh" if energia_total else "n.d.",
)

# Download logo no topo, fora de qualquer st.button para não desaparecer no rerun
st.download_button(
    "Descarregar resumo (Excel)",
    data=gerar_excel(detalhe, por_cpe, por_mes),
    file_name=f"resumo_notas_liquidacao_{datetime.now():%Y%m%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# Avisos de qualidade dos dados
incompletos = detalhe[detalhe["Intervalos"] != detalhe["Intervalos esperados"]]
if not incompletos.empty:
    st.warning(
        f"{len(incompletos)} nota(s) com número de quartos de hora diferente do esperado. "
        "Verificar com o representante."
    )

aba1, aba2, aba3 = st.tabs(["Por CPE", "Por mês", "Detalhe"])
fmt = {c: "{:,.2f}" for c in detalhe.select_dtypes("number").columns}
with aba1:
    st.dataframe(por_cpe.style.format({c: "{:,.2f}" for c in por_cpe.select_dtypes("number").columns}),
                 use_container_width=True, hide_index=True)
with aba2:
    st.dataframe(por_mes.style.format({c: "{:,.2f}" for c in por_mes.select_dtypes("number").columns}),
                 use_container_width=True, hide_index=True)
with aba3:
    st.dataframe(detalhe.style.format(fmt), use_container_width=True, hide_index=True)
