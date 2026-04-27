from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Dict

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


def _fmt_cop(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return "$ 0,00"
    s = f"{float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {s}"


def _safe_text(value: object) -> str:
    if value is None or pd.isna(value):
        return "-"
    return str(value)


def _make_table(data: list[list[str]], col_widths: list[float]) -> Table:
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#dce3e0")),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#f8fdfc")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5faf9")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


def build_executive_report_pdf(
    mov: pd.DataFrame,
    sal: pd.DataFrame,
    quality: pd.DataFrame,
    rejected: pd.DataFrame,
    filters: Dict[str, str],
) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title="Reporte Ejecutivo Materiales",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=20,
        textColor=colors.HexColor("#0f1720"),
        spaceAfter=8,
    )
    subtitle = ParagraphStyle(
        "Subtitle",
        parent=styles["Normal"],
        fontSize=10,
        textColor=colors.HexColor("#5f6c76"),
    )
    h2 = ParagraphStyle(
        "Heading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=13,
        textColor=colors.HexColor("#0f1720"),
        spaceBefore=10,
        spaceAfter=6,
    )
    normal = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#0f1720"),
    )

    total_debito = float(mov["debito"].sum()) if not mov.empty else 0.0
    total_credito = float(mov["credito"].sum()) if not mov.empty else 0.0
    saldo_neto = total_debito - total_credito
    n_oper = int(mov["no_oper"].nunique()) if not mov.empty else 0
    n_ter = int(mov["tercero_id"].nunique()) if not mov.empty else 0
    saldo_final = float(sal["saldo_actual"].sum()) if not sal.empty else 0.0

    content = [
        Paragraph("Reporte Ejecutivo de Materiales", title_style),
        Paragraph(
            f"Protelec · Emitido: {datetime.now():%d/%m/%Y %H:%M} · Filtros: {filters.get('periodo', '-')}, Año: {filters.get('anios', '-')}",
            subtitle,
        ),
        Spacer(1, 8),
    ]

    kpi_rows = [
        ["Indicador", "Valor"],
        ["Debito total", _fmt_cop(total_debito)],
        ["Credito total", _fmt_cop(total_credito)],
        ["Saldo neto", _fmt_cop(saldo_neto)],
        ["Saldo final terceros", _fmt_cop(saldo_final)],
        ["Operaciones", f"{n_oper:,}"],
        ["Terceros activos", f"{n_ter:,}"],
    ]
    content.append(_make_table(kpi_rows, [8 * cm, 8.5 * cm]))

    content.append(Paragraph("Hallazgos", h2))
    top_cuenta = "n/d"
    if not mov.empty:
        tc = mov.groupby(["cuenta", "nombre_cuenta"], dropna=False)["debito"].sum().sort_values(ascending=False).head(1)
        if not tc.empty:
            idx = tc.index[0]
            top_cuenta = f"{_safe_text(idx[0])} - {_safe_text(idx[1])}"
    top_ter = "n/d"
    if not sal.empty:
        tt = sal.groupby(["tercero_id", "tercero_nombre"], dropna=False)["saldo_actual"].sum().sort_values(ascending=False).head(1)
        if not tt.empty:
            idx = tt.index[0]
            top_ter = f"{_safe_text(idx[0])} - {_safe_text(idx[1])}"

    content.append(Paragraph(f"• Cuenta dominante por debito: <b>{top_cuenta}</b>", normal))
    content.append(Paragraph(f"• Tercero con mayor saldo: <b>{top_ter}</b>", normal))
    content.append(
        Paragraph(
            f"• Calidad de datos: <b>{quality['filas'].sum() if not quality.empty else 0}</b> filas procesadas, <b>{len(rejected) if rejected is not None else 0}</b> filas rechazadas.",
            normal,
        )
    )

    content.append(Paragraph("Top cuentas por debito", h2))
    top_cuentas = (
        mov.groupby(["cuenta", "nombre_cuenta"], dropna=False)["debito"].sum().reset_index().sort_values("debito", ascending=False).head(10)
        if not mov.empty
        else pd.DataFrame(columns=["cuenta", "nombre_cuenta", "debito"])
    )
    t_rows = [["Cuenta", "Nombre", "Debito"]]
    for _, r in top_cuentas.iterrows():
        t_rows.append([_safe_text(r.get("cuenta")), _safe_text(r.get("nombre_cuenta")), _fmt_cop(r.get("debito"))])
    if len(t_rows) == 1:
        t_rows.append(["-", "-", "-"])
    content.append(_make_table(t_rows, [3 * cm, 9 * cm, 4.5 * cm]))

    content.append(Paragraph("Top terceros por saldo", h2))
    top_terceros = (
        sal.groupby(["tercero_id", "tercero_nombre"], dropna=False)["saldo_actual"]
        .sum()
        .reset_index()
        .sort_values("saldo_actual", ascending=False)
        .head(10)
        if not sal.empty
        else pd.DataFrame(columns=["tercero_id", "tercero_nombre", "saldo_actual"])
    )
    s_rows = [["Tercero", "Nombre", "Saldo actual"]]
    for _, r in top_terceros.iterrows():
        s_rows.append([_safe_text(r.get("tercero_id")), _safe_text(r.get("tercero_nombre")), _fmt_cop(r.get("saldo_actual"))])
    if len(s_rows) == 1:
        s_rows.append(["-", "-", "-"])
    content.append(_make_table(s_rows, [3 * cm, 9 * cm, 4.5 * cm]))

    content.append(Paragraph("Calidad y control", h2))
    q_rows = [["Dataset", "Filas", "Rechazadas", "Duplicados", "Nulos criticos"]]
    if quality.empty:
        q_rows.append(["-", "0", "0", "0", "0"])
    else:
        for _, r in quality.iterrows():
            q_rows.append(
                [
                    _safe_text(r.get("dataset")),
                    f"{int(r.get('filas', 0)):,}",
                    f"{int(r.get('rechazadas', 0)):,}",
                    f"{int(r.get('duplicados', 0)):,}",
                    f"{int(r.get('nulos_criticos', 0)):,}",
                ]
            )
    content.append(_make_table(q_rows, [5 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm, 3 * cm]))

    content.append(Paragraph("Metodologia ETL y calidad de datos", h2))
    content.append(
        Paragraph(
            "El proceso ETL consolida archivos Excel en formato reporte hacia un modelo analitico estable. "
            "Primero se extraen las hojas fuente. Luego se transforman para remover cabeceras visuales, "
            "unir filas partidas, convertir tipos (fecha/moneda) y estandarizar identificadores. "
            "Finalmente se cargan tablas limpias para analisis y conciliacion.",
            normal,
        )
    )

    etl_rows = [
        ["Etapa", "Descripcion", "Salida"],
        ["Extraccion", "Lectura de 5 reportes contables de materiales", "Registros crudos"],
        [
            "Transformacion",
            "Limpieza de encabezados, consolidacion de multilinea, tipificacion y normalizacion",
            "Registros consistentes",
        ],
        ["Carga", "Construccion de tablas fact y dimension para BI", "fact_movimientos, fact_saldos_tercero, dim_cuentas"],
    ]
    content.append(_make_table(etl_rows, [2.8 * cm, 8.6 * cm, 4.1 * cm]))

    content.append(Paragraph("Por que la calidad de origen es baja", h2))
    reject_count = len(rejected) if rejected is not None else 0
    dup_count = int(quality["duplicados"].sum()) if not quality.empty else 0
    quality_issues = [
        ["Problema", "Evidencia", "Impacto"],
        [
            "Formato de reporte no tabular",
            "Cabeceras, subtotales y texto mezclados en las mismas hojas",
            "Dificulta agregaciones directas",
        ],
        ["Filas partidas (multilinea)", "Nombres y detalles continúan en filas siguientes", "Riesgo de registros incompletos"],
        ["Tipos mezclados por columna", "Moneda formateada, numeros y texto en la misma columna", "Errores de conversion"],
        ["IDs de tercero inconsistentes", f"{reject_count} filas rechazadas por validacion", "Baja trazabilidad"],
        ["Duplicados", f"{dup_count} registros duplicados detectados", "Distorsion de KPI"],
    ]
    content.append(_make_table(quality_issues, [4.1 * cm, 6.8 * cm, 4.6 * cm]))

    content.append(Paragraph("Por que se crea un dataset limpio", h2))
    content.append(Paragraph("• Permite comparabilidad entre periodos y cuentas sin sesgo de formato.", normal))
    content.append(Paragraph("• Habilita filtros confiables por cuenta, tercero y operacion en el tablero ejecutivo.", normal))
    content.append(Paragraph("• Soporta conciliacion: movimientos vs variacion de saldos por anio.", normal))
    content.append(Paragraph("• Reduce riesgo operativo en decisiones de compras y control financiero.", normal))

    doc.build(content)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def build_section_report_pdf(
    section: str,
    mov: pd.DataFrame,
    sal: pd.DataFrame,
    cuentas: pd.DataFrame,
    quality: pd.DataFrame,
    rejected: pd.DataFrame,
    filters: Dict[str, str],
) -> bytes:
    section = (section or "direccion").lower()
    if section == "direccion":
        return build_executive_report_pdf(mov, sal, quality, rejected, filters)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title=f"Reporte {section}",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("S-Title", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=18, textColor=colors.HexColor("#0f1720"), spaceAfter=8)
    subtitle = ParagraphStyle("S-Sub", parent=styles["Normal"], fontSize=10, textColor=colors.HexColor("#5f6c76"))
    h2 = ParagraphStyle("S-H2", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, textColor=colors.HexColor("#0f1720"), spaceBefore=10, spaceAfter=6)
    normal = ParagraphStyle("S-B", parent=styles["Normal"], fontSize=10, leading=14, textColor=colors.HexColor("#0f1720"))

    content = [
        Paragraph(f"Reporte - {section.capitalize()}", title_style),
        Paragraph(
            f"Protelec · Emitido: {datetime.now():%d/%m/%Y %H:%M} · Filtros: {filters.get('periodo', '-')}, Año: {filters.get('anios', '-')}",
            subtitle,
        ),
        Spacer(1, 8),
    ]

    if section == "movimientos":
        neto = float(mov["debito"].sum() - mov["credito"].sum()) if not mov.empty else 0.0
        kpi = [
            ["Indicador", "Valor"],
            ["Lineas", f"{len(mov):,}"],
            ["Operaciones unicas", f"{int(mov['no_oper'].nunique()) if not mov.empty else 0:,}"],
            ["Terceros unicos", f"{int(mov['tercero_id'].nunique()) if not mov.empty else 0:,}"],
            ["Debito", _fmt_cop(float(mov["debito"].sum()) if not mov.empty else 0.0)],
            ["Credito", _fmt_cop(float(mov["credito"].sum()) if not mov.empty else 0.0)],
            ["Neto", _fmt_cop(neto)],
        ]
        content.append(_make_table(kpi, [8 * cm, 8.5 * cm]))

        top_tipo = mov["tipo_doc"].fillna("Sin tipo").astype(str).value_counts().head(8).reset_index() if not mov.empty else pd.DataFrame(columns=["tipo_doc", "conteo"])
        if not top_tipo.empty:
            top_tipo.columns = ["tipo_doc", "conteo"]
        rows = [["Tipo doc", "Conteo"]] + [[_safe_text(r.get("tipo_doc")), f"{int(r.get('conteo', 0)):,}"] for _, r in top_tipo.iterrows()]
        if len(rows) == 1:
            rows.append(["-", "0"])
        content.append(Paragraph("Top 8 tipo_doc", h2))
        content.append(_make_table(rows, [11 * cm, 5.5 * cm]))

        if "tercero_nombre" in mov.columns:
            top_ter = mov["tercero_nombre"].fillna("Sin nombre").astype(str).str.strip().replace("", "Sin nombre").value_counts().head(8).reset_index()
            top_ter.columns = ["tercero_nombre", "conteo"]
            rows_t = [["Tercero nombre", "Conteo"]] + [[_safe_text(r.get("tercero_nombre")), f"{int(r.get('conteo', 0)):,}"] for _, r in top_ter.iterrows()]
            content.append(Paragraph("Top 8 tercero_nombre", h2))
            content.append(_make_table(rows_t, [11 * cm, 5.5 * cm]))

    elif section == "saldos":
        kpi = [
            ["Indicador", "Valor"],
            ["Registros", f"{len(sal):,}"],
            ["Cuentas unicas", f"{int(sal['cuenta'].nunique()) if not sal.empty else 0:,}"],
            ["Terceros unicos", f"{int(sal['tercero_id'].nunique()) if not sal.empty else 0:,}"],
            ["Saldo anterior", _fmt_cop(float(sal["saldo_anterior"].sum()) if not sal.empty else 0.0)],
            ["Debitos", _fmt_cop(float(sal["debitos"].sum()) if not sal.empty else 0.0)],
            ["Creditos", _fmt_cop(float(sal["creditos"].sum()) if not sal.empty else 0.0)],
            ["Saldo actual", _fmt_cop(float(sal["saldo_actual"].sum()) if not sal.empty else 0.0)],
        ]
        content.append(_make_table(kpi, [8 * cm, 8.5 * cm]))

        top_c = sal["nombre_cuenta"].fillna("Sin nombre").astype(str).str.strip().replace("", "Sin nombre").value_counts().head(8).reset_index() if not sal.empty else pd.DataFrame(columns=["nombre_cuenta", "conteo"])
        if not top_c.empty:
            top_c.columns = ["nombre_cuenta", "conteo"]
        rows_c = [["Nombre cuenta", "Conteo"]] + [[_safe_text(r.get("nombre_cuenta")), f"{int(r.get('conteo', 0)):,}"] for _, r in top_c.iterrows()]
        if len(rows_c) == 1:
            rows_c.append(["-", "0"])
        content.append(Paragraph("Top 8 nombre_cuenta", h2))
        content.append(_make_table(rows_c, [11 * cm, 5.5 * cm]))

        top_tn = sal["tercero_nombre"].fillna("Sin nombre").astype(str).str.strip().replace("", "Sin nombre").value_counts().head(8).reset_index() if not sal.empty else pd.DataFrame(columns=["tercero_nombre", "conteo"])
        if not top_tn.empty:
            top_tn.columns = ["tercero_nombre", "conteo"]
        rows_tn = [["Tercero nombre", "Conteo"]] + [[_safe_text(r.get("tercero_nombre")), f"{int(r.get('conteo', 0)):,}"] for _, r in top_tn.iterrows()]
        if len(rows_tn) == 1:
            rows_tn.append(["-", "0"])
        content.append(Paragraph("Top 8 tercero_nombre", h2))
        content.append(_make_table(rows_tn, [11 * cm, 5.5 * cm]))

    elif section == "gobierno":
        mov_g = mov.copy()
        if not mov_g.empty and "neto" not in mov_g.columns and {"debito", "credito"}.issubset(mov_g.columns):
            mov_g["neto"] = mov_g["debito"].fillna(0) - mov_g["credito"].fillna(0)
        sal_g = sal.copy()
        if not sal_g.empty and "delta_saldo" not in sal_g.columns and {"saldo_actual", "saldo_anterior"}.issubset(sal_g.columns):
            sal_g["delta_saldo"] = sal_g["saldo_actual"].fillna(0) - sal_g["saldo_anterior"].fillna(0)

        content.append(Paragraph("Calidad de datos", h2))
        q_rows = [["Dataset", "Filas", "Rechazadas", "Duplicados", "Nulos criticos"]]
        if not quality.empty:
            for _, r in quality.iterrows():
                q_rows.append([
                    _safe_text(r.get("dataset")),
                    f"{int(r.get('filas', 0)):,}",
                    f"{int(r.get('rechazadas', 0)):,}",
                    f"{int(r.get('duplicados', 0)):,}",
                    f"{int(r.get('nulos_criticos', 0)):,}",
                ])
        if len(q_rows) == 1:
            q_rows.append(["-", "0", "0", "0", "0"])
        content.append(_make_table(q_rows, [5 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm, 3 * cm]))

        rec_mov = mov_g.groupby("anio", dropna=False)["neto"].sum().reset_index().rename(columns={"neto": "neto_movimientos"}) if not mov_g.empty and "neto" in mov_g.columns else pd.DataFrame(columns=["anio", "neto_movimientos"])
        rec_sal = sal_g.groupby("anio", dropna=False)["delta_saldo"].sum().reset_index().rename(columns={"delta_saldo": "delta_saldos"}) if not sal_g.empty and "delta_saldo" in sal_g.columns else pd.DataFrame(columns=["anio", "delta_saldos"])
        rec = rec_mov.merge(rec_sal, on="anio", how="outer").fillna(0)
        if not rec.empty:
            rec["diferencia"] = rec["neto_movimientos"] - rec["delta_saldos"]
        content.append(Paragraph("Reconciliacion por anio", h2))
        rr = [["Anio", "Neto movimientos", "Delta saldos", "Diferencia"]]
        for _, r in rec.iterrows():
            rr.append([
                _safe_text(r.get("anio")),
                _fmt_cop(r.get("neto_movimientos")),
                _fmt_cop(r.get("delta_saldos")),
                _fmt_cop(r.get("diferencia")),
            ])
        if len(rr) == 1:
            rr.append(["-", "$ 0,00", "$ 0,00", "$ 0,00"])
        content.append(_make_table(rr, [3 * cm, 4.5 * cm, 4.5 * cm, 4.5 * cm]))

        content.append(Paragraph(f"Filas rechazadas: {len(rejected) if rejected is not None else 0}", normal))

    elif section == "etl":
        content.append(Paragraph("Metodologia ETL", h2))
        content.append(Paragraph("Extraccion de reportes Excel, transformacion con limpieza/normalizacion y carga en tablas analiticas para BI.", normal))
        etl_rows = [
            ["Etapa", "Descripcion", "Salida"],
            ["Extraccion", "5 archivos Excel contables", "Registros crudos"],
            ["Transformacion", "Limpieza de encabezados, tipificacion y consolidacion multilinea", "Registros consistentes"],
            ["Carga", "Modelo en fact/dim", "fact_movimientos, fact_saldos_tercero, dim_cuentas, dim_terceros"],
        ]
        content.append(_make_table(etl_rows, [2.8 * cm, 8.6 * cm, 4.1 * cm]))

    elif section == "diccionario":
        content.append(Paragraph("Diccionario de datos", h2))
        dict_rows = [
            ["Tabla", "Campo", "Descripcion"],
            ["fact_movimientos", "tercero_id", "NIT del tercero"],
            ["fact_movimientos", "tercero_nombre", "Razon social mapeada desde dim_terceros"],
            ["fact_saldos_tercero", "saldo_actual", "Saldo final por tercero y cuenta"],
            ["dim_cuentas", "nombre_cuenta", "Nombre de la cuenta contable"],
            ["dim_terceros", "tercero_nombre", "Razon social por NIT"],
        ]
        content.append(_make_table(dict_rows, [4.2 * cm, 3.5 * cm, 8.0 * cm]))

    else:
        content.append(Paragraph("Seccion no reconocida.", normal))

    doc.build(content)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes
