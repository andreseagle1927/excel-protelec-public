from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from src.reporting.pdf_report import build_section_report_pdf


ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data" / "clean"
BRAND = {
    "bg": "#f4f7f6",
    "surface": "#ffffff",
    "ink": "#0f1720",
    "muted": "#5f6c76",
    "accent": "#0f766e",
    "accent_soft": "#d8f0ee",
    "line": "#dce3e0",
}
CHART_HEIGHT = 340


@st.cache_data
def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    mov = pd.read_parquet(DATA / "fact_movimientos.parquet")
    sal = pd.read_parquet(DATA / "fact_saldos_tercero.parquet")
    cuentas = pd.read_parquet(DATA / "dim_cuentas.parquet")
    terceros_path = DATA / "dim_terceros.parquet"
    terceros = pd.read_parquet(terceros_path) if terceros_path.exists() else pd.DataFrame(columns=["tercero_id", "tercero_nombre"])
    quality = pd.read_csv(DATA / "quality_summary.csv")
    rejected = pd.read_csv(DATA / "rejected_rows.csv")

    if not mov.empty:
        mov["fecha"] = pd.to_datetime(mov["fecha"], errors="coerce")
        mov["mes"] = mov["fecha"].dt.to_period("M").astype(str)
        mov["neto"] = mov["debito"].fillna(0) - mov["credito"].fillna(0)

    if not sal.empty:
        sal["fecha_corte"] = pd.to_datetime(sal["fecha_corte"], errors="coerce")
        sal["delta_saldo"] = sal["saldo_actual"].fillna(0) - sal["saldo_anterior"].fillna(0)

    return mov, sal, cuentas, terceros, quality, rejected


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def fmt_cop(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return "$ 0,00"
    s = f"{float(value):,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {s}"


def period_label(df: pd.DataFrame) -> str:
    if df.empty or "fecha" not in df.columns or df["fecha"].dropna().empty:
        return "Sin periodo"
    start = df["fecha"].min()
    end = df["fecha"].max()
    if pd.isna(start) or pd.isna(end):
        return "Sin periodo"
    return f"{start:%d/%m/%Y} - {end:%d/%m/%Y}"


def compact_label(value: str | None, max_len: int = 42) -> str:
    text = str(value or "").strip()
    if len(text) <= max_len:
        return text
    return text[: max_len - 1].rstrip() + "…"


def apply_chart_theme(fig: go.Figure, *, height: int = CHART_HEIGHT) -> go.Figure:
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=dict(l=12, r=12, t=28, b=12),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font=dict(family="Source Sans 3, sans-serif", color=BRAND["ink"], size=13),
        hoverlabel=dict(bgcolor="#ffffff", font_size=12, font_family="Source Sans 3, sans-serif"),
        autosize=True,
    )
    fig.update_xaxes(showgrid=True, gridcolor="#edf3f1", zeroline=False, automargin=True)
    fig.update_yaxes(showgrid=True, gridcolor="#edf3f1", zeroline=False, automargin=True)
    return fig


def render_chart(target, fig: go.Figure) -> None:
    target.plotly_chart(
        fig,
        use_container_width=True,
        config={"responsive": True, "displaylogo": False},
    )


def calc_quality_score(quality_df: pd.DataFrame, rejected_df: pd.DataFrame) -> float:
    if quality_df.empty:
        return 0.0
    total_rows = quality_df["filas"].sum()
    total_rejected = quality_df["rechazadas"].sum()
    total_duplicates = quality_df["duplicados"].sum()
    total_nulls = quality_df["nulos_criticos"].sum()

    if total_rows <= 0:
        return 0.0

    reject_ratio = total_rejected / total_rows
    dup_ratio = total_duplicates / total_rows
    null_ratio = total_nulls / total_rows
    reject_reason_penalty = min(len(rejected_df["motivo"].dropna().unique()) * 0.5, 3.0) if not rejected_df.empty else 0

    score = 100 - (reject_ratio * 55 + dup_ratio * 30 + null_ratio * 15) * 100 - reject_reason_penalty
    return max(0.0, min(100.0, score))


def apply_filters(
    mov: pd.DataFrame,
    sal: pd.DataFrame,
    years: Iterable[int],
    cuentas: Iterable[str],
    terceros: Iterable[str],
    start_date: pd.Timestamp | None,
    end_date: pd.Timestamp | None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    mov_f = mov.copy()
    sal_f = sal.copy()

    years = list(years)
    cuentas = set(cuentas)
    terceros = set(terceros)

    if years:
        mov_f = mov_f[mov_f["anio"].isin(years)]
        sal_f = sal_f[sal_f["anio"].isin(years)]

    if cuentas:
        mov_f = mov_f[mov_f["cuenta"].astype(str).isin(cuentas)]
        sal_f = sal_f[sal_f["cuenta"].astype(str).isin(cuentas)]

    if terceros:
        mov_f = mov_f[mov_f["tercero_id"].astype(str).isin(terceros)]
        sal_f = sal_f[sal_f["tercero_id"].astype(str).isin(terceros)]

    if start_date is not None:
        mov_f = mov_f[mov_f["fecha"] >= pd.to_datetime(start_date)]
    if end_date is not None:
        mov_f = mov_f[mov_f["fecha"] <= pd.to_datetime(end_date)]

    return mov_f, sal_f


def render_pdf_button(pdf_data: bytes, file_name: str, key_suffix: str) -> None:
    st.download_button(
        "Descargar PDF ejecutivo",
        data=pdf_data,
        file_name=file_name,
        mime="application/pdf",
        use_container_width=True,
        key=f"pdf_exec_{key_suffix}",
    )


st.set_page_config(page_title="Dirección Financiera Materiales", page_icon="📊", layout="wide")

st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@500;700;800&family=Source+Sans+3:wght@400;600;700&display=swap');
html, body, [class*="css"] {{ font-family: 'Source Sans 3', sans-serif; background: {BRAND['bg']}; color: {BRAND['ink']}; }}
h1, h2, h3 {{ font-family: 'Manrope', sans-serif; letter-spacing: -0.02em; color: {BRAND['ink']}; }}
.hero {{ background: linear-gradient(130deg, #f8fefd 0%, #eef5f3 58%, #e6f1ee 100%); border: 1px solid {BRAND['line']}; border-radius: 18px; padding: 16px 18px; margin-bottom: 14px; }}
.meta {{ color: {BRAND['muted']}; font-size: 0.94rem; }}
[data-testid="stVerticalBlock"] > [style*="flex-direction: column"] > [data-testid="stVerticalBlockBorderWrapper"] {{ border-radius: 14px; }}
[data-testid="stMetric"] {{ background: {BRAND['surface']}; border: 1px solid {BRAND['line']}; border-radius: 14px; padding: 10px; }}
[data-testid="stMetricLabel"] p {{ color: {BRAND['muted']}; }}
[data-testid="stSidebar"] {{ background: #fbfdfd; border-right: 1px solid {BRAND['line']}; }}
[data-testid="stSidebar"] [data-baseweb="select"] > div,
[data-testid="stSidebar"] [data-baseweb="input"] > div {{ border-color: {BRAND['line']}; border-radius: 10px; }}
[data-testid="stSidebar"] [role="radiogroup"] > label {{ background: #f4f8f7; border: 1px solid {BRAND['line']}; border-radius: 10px; padding: 8px 10px; margin-bottom: 6px; }}
.stTabs [data-baseweb="tab-list"] {{ gap: 8px; }}
.stTabs [data-baseweb="tab"] {{ background: #ecf4f2; border-radius: 10px; padding: 6px 12px; }}
.stTabs [aria-selected="true"] {{ background: #d9ece8; }}
.kpi-grid {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; margin-bottom: 10px; }}
.kpi-card {{ background: {BRAND['surface']}; border: 1px solid {BRAND['line']}; border-radius: 12px; padding: 10px 12px; }}
.kpi-title {{ color: {BRAND['muted']}; font-size: 0.82rem; margin-bottom: 3px; }}
.kpi-value {{ color: {BRAND['ink']}; font-size: 1.15rem; font-weight: 700; line-height: 1.2; }}
.section-card {{ background: {BRAND['surface']}; border: 1px solid {BRAND['line']}; border-radius: 14px; padding: 10px 12px; margin-bottom: 10px; }}
@media (max-width: 1024px) {{
  .kpi-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
}}
@media (max-width: 1200px) {{
  [data-testid="stHorizontalBlock"] {{ flex-wrap: wrap; gap: 0.8rem; }}
  [data-testid="column"] {{ min-width: 100% !important; flex: 1 1 100% !important; }}
}}
@media (max-width: 768px) {{
  .hero {{ padding: 12px; border-radius: 14px; }}
  .meta {{ font-size: 0.82rem; line-height: 1.35; }}
  .kpi-grid {{ grid-template-columns: 1fr; }}
  .stTabs [data-baseweb="tab-list"] {{ flex-wrap: wrap; }}
  .stTabs [data-baseweb="tab"] {{ flex: 1 1 48%; justify-content: center; }}
  [data-testid="stDataFrame"] {{ max-height: 320px; overflow: auto; }}
  .js-plotly-plot,
  .js-plotly-plot .plot-container,
  .js-plotly-plot .svg-container {{ min-height: 280px !important; height: 280px !important; }}
}}
</style>
""",
    unsafe_allow_html=True,
)

if not (DATA / "fact_movimientos.parquet").exists():
    st.warning("No hay datos procesados. Ejecuta `python3 scripts/process_data.py`.")
    st.stop()

mov, sal, cuentas, terceros_dim, quality, rejected = load_data()

last_refresh = max((DATA / "fact_movimientos.parquet").stat().st_mtime, (DATA / "fact_saldos_tercero.parquet").stat().st_mtime)
last_refresh_ts = pd.to_datetime(last_refresh, unit="s")

st.markdown("<div class='hero'>", unsafe_allow_html=True)
st.title("📊 Dirección de Materiales · Protelec")
st.markdown("<div class='meta'>🏦 Plataforma de datos de contabilidad para analítica financiera y control ejecutivo.</div>", unsafe_allow_html=True)
quality_score = calc_quality_score(quality, rejected)
st.markdown(
    f"<div class='meta'>🗓️ Periodo disponible: <b>{period_label(mov)}</b> · 🔄 Última actualización: <b>{last_refresh_ts:%d/%m/%Y %H:%M}</b> · ✅ Calidad de datos: <b>{quality_score:.1f}/100</b></div>",
    unsafe_allow_html=True,
)
st.markdown("</div>", unsafe_allow_html=True)

st.sidebar.header("🧭 Control ejecutivo")
anios = sorted(set(mov["anio"].dropna().astype(int).tolist() + sal["anio"].dropna().astype(int).tolist()))
default_years = [max(anios)] if anios else []
anio_sel = st.sidebar.multiselect("Año fiscal", options=anios, default=default_years)

preset = st.sidebar.radio("Ventana de tiempo", ["Año completo", "YTD", "Último mes", "Personalizado"], index=0)
date_min = mov["fecha"].min() if not mov.empty else None
date_max = mov["fecha"].max() if not mov.empty else None
start_date, end_date = None, None
if date_min is not None and date_max is not None and not pd.isna(date_min) and not pd.isna(date_max):
    if preset == "YTD":
        start_date = pd.Timestamp(year=date_max.year, month=1, day=1)
        end_date = date_max
    elif preset == "Último mes":
        start_date = date_max - pd.Timedelta(days=30)
        end_date = date_max
    elif preset == "Personalizado":
        span = st.sidebar.date_input("Rango", value=(date_min.date(), date_max.date()))
        if isinstance(span, tuple) and len(span) == 2:
            start_date, end_date = pd.to_datetime(span[0]), pd.to_datetime(span[1])

cuentas_opts = sorted(mov["cuenta"].dropna().astype(str).unique().tolist())
cuenta_sel = st.sidebar.multiselect("Cuenta", options=cuentas_opts)

terceros_opts = sorted(mov["tercero_id"].dropna().astype(str).unique().tolist())
tercero_sel = st.sidebar.multiselect("Tercero", options=terceros_opts)

mov_f, sal_f = apply_filters(mov, sal, anio_sel, cuenta_sel, tercero_sel, start_date, end_date)

if start_date is not None and end_date is not None:
    periodo_pdf = f"{pd.to_datetime(start_date):%d/%m/%Y} - {pd.to_datetime(end_date):%d/%m/%Y}"
else:
    periodo_pdf = period_label(mov_f)
filtros_pdf = {
    "periodo": periodo_pdf,
    "anios": ", ".join(str(x) for x in sorted(anio_sel)) if anio_sel else "Todos",
}
pdf_name_all_tabs = f"reporte_materiales_{datetime.now():%Y%m%d_%H%M}.pdf"

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["Dirección", "Movimientos", "Saldos", "Gobierno de datos", "ETL y Calidad", "Diccionario"]
)

with tab1:
    render_pdf_button(
        build_section_report_pdf("direccion", mov_f, sal_f, cuentas, quality, rejected, filtros_pdf),
        pdf_name_all_tabs,
        "tab1",
    )

    total_debito = float(mov_f["debito"].sum()) if not mov_f.empty else 0.0
    total_credito = float(mov_f["credito"].sum()) if not mov_f.empty else 0.0
    total_saldo = float(sal_f["saldo_actual"].sum()) if not sal_f.empty else 0.0
    total_ops = int(mov_f["no_oper"].nunique()) if not mov_f.empty else 0
    cuentas_activas = int(mov_f["cuenta"].nunique()) if not mov_f.empty else 0
    terceros_mov = int(mov_f["tercero_id"].nunique()) if not mov_f.empty else 0
    terceros_saldo = int(sal_f["tercero_id"].nunique()) if not sal_f.empty else 0
    terceros_catalogo = int(terceros_dim["tercero_id"].nunique()) if not terceros_dim.empty else 0
    cuentas_catalogo = int(cuentas["cuenta"].nunique()) if not cuentas.empty else 0

    st.markdown("**Vista ejecutiva integrada: movimientos, saldos y cuentas**")

    st.markdown(
        """
        <div class="kpi-grid">
          <div class="kpi-card"><div class="kpi-title">Débito total</div><div class="kpi-value">{}</div></div>
          <div class="kpi-card"><div class="kpi-title">Crédito total</div><div class="kpi-value">{}</div></div>
          <div class="kpi-card"><div class="kpi-title">Saldo actual terceros</div><div class="kpi-value">{}</div></div>
          <div class="kpi-card"><div class="kpi-title">Operaciones únicas</div><div class="kpi-value">{}</div></div>
          <div class="kpi-card"><div class="kpi-title">Cuentas activas (mov)</div><div class="kpi-value">{}</div></div>
          <div class="kpi-card"><div class="kpi-title">Terceros en movimientos</div><div class="kpi-value">{}</div></div>
          <div class="kpi-card"><div class="kpi-title">Terceros en saldos</div><div class="kpi-value">{}</div></div>
          <div class="kpi-card"><div class="kpi-title">Terceros catálogo</div><div class="kpi-value">{}</div></div>
          <div class="kpi-card"><div class="kpi-title">Cuentas en catálogo</div><div class="kpi-value">{}</div></div>
        </div>
        """.format(
            fmt_cop(total_debito),
            fmt_cop(total_credito),
            fmt_cop(total_saldo),
            f"{total_ops:,}",
            f"{cuentas_activas:,}",
            f"{terceros_mov:,}",
            f"{terceros_saldo:,}",
            f"{terceros_catalogo:,}",
            f"{cuentas_catalogo:,}",
        ),
        unsafe_allow_html=True,
    )

    st.subheader("Datos clave")
    top_cuenta_debito = "n/d"
    top_tercero_saldo = "n/d"
    prom_debito = float(mov_f["debito"].mean()) if not mov_f.empty else 0.0
    neto_total = total_debito - total_credito

    if not mov_f.empty:
        tc = mov_f.groupby(["cuenta", "nombre_cuenta"], dropna=False)["debito"].sum().sort_values(ascending=False).head(1)
        if not tc.empty:
            idx = tc.index[0]
            top_cuenta_debito = f"{idx[0]} - {idx[1]}"
    if not sal_f.empty:
        tt = sal_f.groupby(["tercero_id", "tercero_nombre"], dropna=False)["saldo_actual"].sum().sort_values(ascending=False).head(1)
        if not tt.empty:
            idx = tt.index[0]
            top_tercero_saldo = f"{idx[0]} - {idx[1]}"

    map_pct = 0.0
    if not mov_f.empty:
        mov_ids = set(mov_f["tercero_id"].dropna().astype(str).str.strip())
        dim_ids = set(terceros_dim["tercero_id"].dropna().astype(str).str.strip()) if not terceros_dim.empty else set()
        map_pct = (len(mov_ids & dim_ids) / len(mov_ids) * 100) if mov_ids else 0.0

    resumen_clave = pd.DataFrame(
        [
            ["Cuenta líder por débito", top_cuenta_debito],
            ["Tercero líder por saldo", top_tercero_saldo],
            ["Neto movimientos", fmt_cop(neto_total)],
            ["Débito promedio por línea", fmt_cop(prom_debito)],
            ["Cobertura mapeo NIT -> Razón social", f"{map_pct:.1f}%"],
            ["Cobertura cuentas usadas/catálogo", f"{cuentas_activas}/{cuentas_catalogo}"],
        ],
        columns=["Indicador", "Valor"],
    )
    st.dataframe(resumen_clave, use_container_width=True, height=220)

    if mov_f.empty and sal_f.empty:
        st.info("Sin datos para filtros actuales.")
    else:
        c_left, c_right = st.columns(2)
        with c_left:
            st.subheader("Top cuentas por débito")
            if mov_f.empty:
                st.info("Sin movimientos para filtros actuales.")
            else:
                top_cuentas = (
                    mov_f.groupby(["cuenta", "nombre_cuenta"], dropna=False)["debito"]
                    .sum()
                    .reset_index()
                    .sort_values("debito", ascending=False)
                    .head(12)
                )
                top_cuentas["label_full"] = top_cuentas["cuenta"].astype(str) + " - " + top_cuentas["nombre_cuenta"].fillna("")
                top_cuentas["label"] = top_cuentas["label_full"].apply(compact_label)
                top_cuentas = top_cuentas.sort_values("debito", ascending=True)
                fig_cuentas = px.bar(
                    top_cuentas,
                    x="debito",
                    y="label",
                    orientation="h",
                    color="debito",
                )
                fig_cuentas.update_traces(
                    customdata=top_cuentas[["label_full"]].to_numpy(),
                    hovertemplate="<b>%{customdata[0]}</b><br>Débito: %{x:,.2f}<extra></extra>",
                )
                fig_cuentas.update_layout(coloraxis_showscale=False, xaxis_title="Débito", yaxis_title="Cuenta")
                fig_cuentas.update_xaxes(tickprefix="$ ")
                apply_chart_theme(fig_cuentas)
                render_chart(st, fig_cuentas)

        with c_right:
            st.subheader("Top terceros por saldo")
            if sal_f.empty:
                st.info("Sin saldos para filtros actuales.")
            else:
                top_terceros = (
                    sal_f.groupby(["tercero_id", "tercero_nombre"], dropna=False)["saldo_actual"]
                    .sum()
                    .reset_index()
                    .sort_values("saldo_actual", ascending=False)
                    .head(12)
                )
                top_terceros["label_full"] = top_terceros["tercero_id"].astype(str) + " - " + top_terceros["tercero_nombre"].fillna("")
                top_terceros["label"] = top_terceros["label_full"].apply(compact_label)
                top_terceros = top_terceros.sort_values("saldo_actual", ascending=True)
                fig_terceros = px.bar(
                    top_terceros,
                    x="saldo_actual",
                    y="label",
                    orientation="h",
                    color="saldo_actual",
                )
                fig_terceros.update_traces(
                    customdata=top_terceros[["label_full"]].to_numpy(),
                    hovertemplate="<b>%{customdata[0]}</b><br>Saldo actual: %{x:,.2f}<extra></extra>",
                )
                fig_terceros.update_layout(coloraxis_showscale=False, xaxis_title="Saldo actual", yaxis_title="Tercero")
                fig_terceros.update_xaxes(tickprefix="$ ")
                apply_chart_theme(fig_terceros)
                render_chart(st, fig_terceros)

with tab2:
    render_pdf_button(
        build_section_report_pdf("movimientos", mov_f, sal_f, cuentas, quality, rejected, filtros_pdf),
        pdf_name_all_tabs,
        "tab2",
    )

    st.subheader("Drilldown de movimientos")
    if mov_f.empty:
        st.info("Sin datos para filtros actuales.")
    else:
        col_a, col_b, col_c = st.columns(3)
        dd_cuenta = col_a.selectbox("1) Cuenta", ["Todas"] + sorted(mov_f["cuenta"].astype(str).unique().tolist()))
        mov_step = mov_f if dd_cuenta == "Todas" else mov_f[mov_f["cuenta"].astype(str) == dd_cuenta]

        dd_tercero = col_b.selectbox("2) Tercero", ["Todos"] + sorted(mov_step["tercero_id"].astype(str).unique().tolist()))
        mov_step = mov_step if dd_tercero == "Todos" else mov_step[mov_step["tercero_id"].astype(str) == dd_tercero]

        dd_oper = col_c.selectbox("3) Operación", ["Todas"] + sorted(mov_step["no_oper"].astype(str).unique().tolist()))
        mov_step = mov_step if dd_oper == "Todas" else mov_step[mov_step["no_oper"].astype(str) == dd_oper]

        search = st.text_input("Buscar por detalle o documento", value="").strip().lower()
        if search:
            mask = (
                mov_step["detalle"].fillna("").str.lower().str.contains(search)
                | mov_step["no_doc"].fillna("").astype(str).str.lower().str.contains(search)
            )
            mov_step = mov_step[mask]

        cols_show = [
            "anio",
            "fecha",
            "no_oper",
            "tipo_doc",
            "no_doc",
            "cuenta",
            "nombre_cuenta",
            "tercero_id",
            "tercero_nombre",
            "detalle",
            "debito",
            "credito",
            "neto",
        ]
        view = mov_step[cols_show].copy()
        view["tercero_nombre"] = view["tercero_nombre"].fillna("").astype(str).str.strip()
        view.loc[view["tercero_nombre"] == "", "tercero_nombre"] = "Sin nombre"
        view = view.sort_values(["fecha", "no_oper"]).reset_index(drop=True)
        st.dataframe(view, use_container_width=True, height=360)
        st.download_button("Descargar movimientos filtrados (CSV)", to_csv_bytes(view), file_name="movimientos_filtrados.csv", mime="text/csv")

        st.subheader("Top 8 repeticiones")
        t1, t2, t3 = st.columns(3)

        top_nombre_cuenta = mov_step["nombre_cuenta"].fillna("Sin nombre").astype(str).value_counts().head(8).reset_index()
        top_nombre_cuenta.columns = ["nombre_cuenta", "conteo"]
        t1.markdown("**Nombre cuenta más repetido (TB-MOV-01)**")
        t1.dataframe(top_nombre_cuenta, use_container_width=True, height=240)
        if not top_nombre_cuenta.empty:
            pie_nombre = px.pie(
                top_nombre_cuenta,
                names="nombre_cuenta",
                values="conteo",
                title="CH-MOV-01 · Distribución nombre_cuenta",
                hole=0.35,
            )
            pie_nombre.update_traces(
                textinfo="none",
                hovertemplate="%{label}<br>Registros: %{value}<br>%{percent}<extra></extra>",
            )
            pie_nombre.update_layout(
                showlegend=True,
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, title_text="Categoria"),
            )
            apply_chart_theme(pie_nombre)
            render_chart(t1, pie_nombre)

        top_tercero = mov_step["tercero_nombre"].fillna("Sin nombre").astype(str).str.strip().value_counts().head(8).reset_index()
        top_tercero.columns = ["tercero_nombre", "conteo"]
        t2.markdown("**Tercero nombre más repetido (TB-MOV-02)**")
        t2.dataframe(top_tercero, use_container_width=True, height=240)
        if not top_tercero.empty:
            pie_tercero = px.pie(
                top_tercero,
                names="tercero_nombre",
                values="conteo",
                title="CH-MOV-02 · Distribución tercero_nombre",
                hole=0.35,
            )
            pie_tercero.update_traces(
                textinfo="none",
                hovertemplate="%{label}<br>Registros: %{value}<br>%{percent}<extra></extra>",
            )
            pie_tercero.update_layout(
                showlegend=True,
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, title_text="Categoria"),
            )
            apply_chart_theme(pie_tercero)
            render_chart(t2, pie_tercero)

        top_tipo_doc = mov_step["tipo_doc"].fillna("Sin tipo").astype(str).str.strip().value_counts().head(8).reset_index()
        top_tipo_doc.columns = ["tipo_doc", "conteo"]
        t3.markdown("**Tipo doc más repetido (TB-MOV-03)**")
        t3.dataframe(top_tipo_doc, use_container_width=True, height=240)
        if not top_tipo_doc.empty:
            pie_tipo_doc = px.pie(
                top_tipo_doc,
                names="tipo_doc",
                values="conteo",
                title="CH-MOV-03 · Distribución tipo_doc",
                hole=0.35,
            )
            pie_tipo_doc.update_traces(
                textinfo="none",
                hovertemplate="%{label}<br>Registros: %{value}<br>%{percent}<extra></extra>",
            )
            pie_tipo_doc.update_layout(
                showlegend=True,
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, title_text="Categoria"),
            )
            apply_chart_theme(pie_tipo_doc)
            render_chart(t3, pie_tipo_doc)

with tab3:
    render_pdf_button(
        build_section_report_pdf("saldos", mov_f, sal_f, cuentas, quality, rejected, filtros_pdf),
        pdf_name_all_tabs,
        "tab3",
    )

    st.subheader("Análisis de saldos")
    if sal_f.empty:
        st.info("Sin datos para filtros actuales.")
    else:
        sal_cols = [
            "anio",
            "fecha_corte",
            "cuenta",
            "nombre_cuenta",
            "tercero_id",
            "tercero_nombre",
            "saldo_anterior",
            "debitos",
            "creditos",
            "saldo_actual",
            "delta_saldo",
        ]
        view_sal = sal_f[sal_cols].sort_values(["anio", "cuenta", "saldo_actual"], ascending=[True, True, False]).reset_index(drop=True)
        st.dataframe(view_sal, use_container_width=True, height=360)
        st.download_button("Descargar saldos filtrados (CSV)", to_csv_bytes(view_sal), file_name="saldos_filtrados.csv", mime="text/csv")

        st.subheader("Top 8 repeticiones")
        s1, s3 = st.columns(2)

        top_cuenta_sal = sal_f["nombre_cuenta"].fillna("Sin nombre").astype(str).str.strip().value_counts().head(8).reset_index()
        top_cuenta_sal.columns = ["nombre_cuenta", "conteo"]
        s1.markdown("**Cuenta más repetida (TB-SAL-01)**")
        s1.dataframe(top_cuenta_sal, use_container_width=True, height=240)
        if not top_cuenta_sal.empty:
            pie_cuenta_sal = px.pie(
                top_cuenta_sal,
                names="nombre_cuenta",
                values="conteo",
                title="CH-SAL-01 · Distribución nombre_cuenta",
                hole=0.35,
            )
            pie_cuenta_sal.update_traces(
                textinfo="none",
                hovertemplate="%{label}<br>Registros: %{value}<br>%{percent}<extra></extra>",
            )
            pie_cuenta_sal.update_layout(
                showlegend=True,
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, title_text="Categoria"),
            )
            apply_chart_theme(pie_cuenta_sal)
            render_chart(s1, pie_cuenta_sal)

        top_tercero_nom_sal = sal_f["tercero_nombre"].fillna("Sin nombre").astype(str).value_counts().head(8).reset_index()
        top_tercero_nom_sal.columns = ["tercero_nombre", "conteo"]
        s3.markdown("**Tercero nombre más repetido (TB-SAL-03)**")
        s3.dataframe(top_tercero_nom_sal, use_container_width=True, height=240)
        if not top_tercero_nom_sal.empty:
            pie_tercero_nom_sal = px.pie(
                top_tercero_nom_sal,
                names="tercero_nombre",
                values="conteo",
                title="CH-SAL-03 · Distribución tercero_nombre",
                hole=0.35,
            )
            pie_tercero_nom_sal.update_traces(
                textinfo="none",
                hovertemplate="%{label}<br>Registros: %{value}<br>%{percent}<extra></extra>",
            )
            pie_tercero_nom_sal.update_layout(
                showlegend=True,
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, title_text="Categoria"),
            )
            apply_chart_theme(pie_tercero_nom_sal)
            render_chart(s3, pie_tercero_nom_sal)

with tab4:
    render_pdf_button(
        build_section_report_pdf("gobierno", mov_f, sal_f, cuentas, quality, rejected, filtros_pdf),
        pdf_name_all_tabs,
        "tab4",
    )

    st.subheader("Score de calidad y reconciliación")

    g_col, q_col = st.columns([1, 2])
    with g_col:
        fig = go.Figure(
            go.Indicator(
                mode="gauge+number",
                value=quality_score,
                title={"text": "Calidad"},
                gauge={
                    "axis": {"range": [0, 100]},
                    "bar": {"color": "#0f766e"},
                    "steps": [
                        {"range": [0, 70], "color": "#f7d5d5"},
                        {"range": [70, 85], "color": "#f7eccf"},
                        {"range": [85, 100], "color": "#d8f0ee"},
                    ],
                    "threshold": {"line": {"color": "#0f1720", "width": 3}, "thickness": 0.75, "value": 85},
                },
            )
        )
        apply_chart_theme(fig)
        render_chart(st, fig)

    with q_col:
        st.dataframe(quality, use_container_width=True, height=250)

    st.subheader("Reconciliación por año")
    rec_mov = mov_f.groupby("anio", dropna=False)["neto"].sum().reset_index().rename(columns={"neto": "neto_movimientos"})
    rec_sal = sal_f.groupby("anio", dropna=False)["delta_saldo"].sum().reset_index().rename(columns={"delta_saldo": "delta_saldos"})
    rec = rec_mov.merge(rec_sal, on="anio", how="outer").fillna(0)
    rec["diferencia"] = rec["neto_movimientos"] - rec["delta_saldos"]
    st.dataframe(rec, use_container_width=True)

    st.subheader("Motivos de rechazo")
    if rejected.empty:
        st.success("No hay filas rechazadas en el lote actual.")
    else:
        rej = rejected.groupby("motivo", as_index=False).size().sort_values("size", ascending=False)
        rej["motivo_short"] = rej["motivo"].apply(compact_label)
        fig = px.bar(
            rej.sort_values("size", ascending=True),
            x="size",
            y="motivo_short",
            orientation="h",
            color="size",
        )
        fig.update_traces(
            customdata=rej.sort_values("size", ascending=True)[["motivo"]].to_numpy(),
            hovertemplate="<b>%{customdata[0]}</b><br>Filas: %{x}<extra></extra>",
        )
        fig.update_layout(coloraxis_showscale=False, xaxis_title="Filas", yaxis_title="Motivo")
        apply_chart_theme(fig)
        render_chart(st, fig)
        st.dataframe(rejected, use_container_width=True, height=300)

with tab5:
    render_pdf_button(
        build_section_report_pdf("etl", mov_f, sal_f, cuentas, quality, rejected, filtros_pdf),
        pdf_name_all_tabs,
        "tab5",
    )

    st.subheader("Archivos originales (fuente)")
    source_files = pd.DataFrame(
        [
            [
                "📗",
                "CUENTA MATERIALES.xlsx",
                "Catalogo de cuentas de materiales (codigo, nombre, nivel).",
                "Base para dimension de cuentas y mapeo semantico.",
            ],
            [
                "📗",
                "Movimiento contable Materiales 2025.xlsx",
                "Reporte contable de movimientos 2025 por operacion, cuenta y tercero.",
                "Se usa para construir fact_movimientos del periodo 2025.",
            ],
            [
                "📗",
                "Movimiento contable Materiales 2026.xlsx",
                "Reporte contable de movimientos 2026 con estructura equivalente a 2025.",
                "Se integra con 2025 para analisis comparativo y tendencias.",
            ],
            [
                "📗",
                "Saldos de terceros cuenta Materiales 25.xlsx",
                "Saldos por tercero y cuenta para corte 2025 (saldo anterior/debitos/creditos/saldo actual).",
                "Se usa para construir fact_saldos_tercero de 2025.",
            ],
            [
                "📗",
                "Saldos de terceros cuenta Materiales 26.xlsx",
                "Saldos por tercero y cuenta para corte 2026 con mismo layout de reporte.",
                "Se integra para comparativo interanual y reconciliacion.",
            ],
        ],
        columns=["Archivo", "Nombre", "Contenido original", "Uso en ETL"],
    )
    st.dataframe(source_files, use_container_width=True, height=260)

    st.subheader("Cómo funciona el ETL")
    st.markdown(
        "\n".join(
            [
                "**Extracción:** se leen 5 archivos Excel contables en formato de reporte.",
                "**Transformación:** se eliminan encabezados visuales, se unen filas partidas, se tipifican fechas/valores y se normalizan IDs.",
                "**Carga:** se construyen tablas limpias para análisis: `dim_cuentas`, `dim_terceros`, `fact_movimientos`, `fact_saldos_tercero`.",
            ]
        )
    )

    st.subheader("Por qué la calidad de origen es baja")
    bad_quality = pd.DataFrame(
        [
            ["Estructura tipo reporte", "Alta", "Cabeceras, subtotales y textos mezclados con datos."],
            ["Filas partidas", "Alta", "Nombres y descripciones continúan en la fila siguiente."],
            ["Tipos mezclados", "Media", "Misma columna contiene texto, números y moneda formateada."],
            ["IDs no consistentes", "Media", "Algunas filas de tercero no son ID numérico válido."],
            ["Duplicidad de líneas", "Media", "Existen líneas repetidas en movimientos."],
        ],
        columns=["Problema", "Severidad", "Impacto"],
    )
    st.dataframe(bad_quality, use_container_width=True, height=220)

    st.subheader("Qué mejora con el dataset limpio")
    b1, b2, b3 = st.columns(3)
    b1.metric("Filas procesadas", f"{int(quality['filas'].sum()) if not quality.empty else 0:,}")
    b2.metric("Filas rechazadas", f"{len(rejected):,}")
    b3.metric("Score calidad", f"{quality_score:.1f}/100")

    st.markdown(
        "\n".join(
            [
                "- KPIs consistentes entre periodos (2025 vs 2026).",
                "- Filtros y drilldown confiables por cuenta, tercero y operación.",
                "- Reconciliación trazable entre movimientos y saldos.",
                "- Base reutilizable para automatización y auditoría continua.",
            ]
        )
    )

    st.subheader("Trazabilidad de datos (lineage)")
    lineage = pd.DataFrame(
        [
            [
                "Movimiento contable Materiales 2025/2026.xlsx",
                "Parseo de bloques operativos + unión de detalle multilinea + tipificación",
                "fact_movimientos",
            ],
            [
                "Saldos de terceros cuenta Materiales 25/26.xlsx",
                "Detección de secciones Cuenta + consolidación de nombre + cálculo delta + dim de terceros",
                "fact_saldos_tercero, dim_terceros",
            ],
            [
                "CUENTA MATERIALES.xlsx",
                "Normalización catálogo y deduplicación de cuentas",
                "dim_cuentas",
            ],
        ],
        columns=["Fuente", "Transformación aplicada", "Salida limpia"],
    )
    st.dataframe(lineage, use_container_width=True, height=210)

with tab6:
    render_pdf_button(
        build_section_report_pdf("diccionario", mov_f, sal_f, cuentas, quality, rejected, filtros_pdf),
        pdf_name_all_tabs,
        "tab6",
    )

    st.subheader("Diccionario de datos")
    dictionary = pd.DataFrame(
        [
            ["fact_movimientos", "anio", "int", "Año contable del movimiento"],
            ["fact_movimientos", "fecha", "date", "Fecha del documento/operación"],
            ["fact_movimientos", "no_oper", "string", "Identificador de operación en reporte"],
            ["fact_movimientos", "cuenta", "string", "Código de cuenta contable"],
            ["fact_movimientos", "tercero_id", "string", "Identificación de tercero"],
            ["fact_movimientos", "debito", "float", "Valor débito línea"],
            ["fact_movimientos", "credito", "float", "Valor crédito línea"],
            ["fact_saldos_tercero", "saldo_anterior", "float", "Saldo inicial por tercero y cuenta"],
            ["fact_saldos_tercero", "debitos", "float", "Débitos acumulados del periodo"],
            ["fact_saldos_tercero", "creditos", "float", "Créditos acumulados del periodo"],
            ["fact_saldos_tercero", "saldo_actual", "float", "Saldo final por tercero y cuenta"],
            ["dim_cuentas", "nombre_cuenta", "string", "Descripción textual de la cuenta"],
        ],
        columns=["tabla", "campo", "tipo", "definición"],
    )
    st.dataframe(dictionary, use_container_width=True, height=380)

    st.caption("Regla de oro: IDs contables y terceros se conservan como texto para evitar pérdida de formato.")
