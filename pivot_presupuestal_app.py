"""
Constructor de Tabla Dinámica Presupuestal — MAP / SICOP
==========================================================

"""

from __future__ import annotations

import io
from datetime import date, datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Paleta SADER
# ---------------------------------------------------------------------------
ROJO_SADER = "9F2241"       # burgundy — encabezados principales
CAFE_SADER = "BC955C"       # café — subtítulos / acentos
CREMA_SADER = "F5EFE6"      # cream — fondo de filas de totales
GRIS_BORDE = "BFBFBF"

THIN = Side(style="thin", color=GRIS_BORDE)
BORDE = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FUENTES = {
    "MAP": "Módulo de Adecuaciones Presupuestarias (MAP)",
    "SICOP": "Sistema de Contabilidad y Presupuesto (SICOP)",
}

# ---------------------------------------------------------------------------
# Carga de datos
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def cargar_datos(archivo) -> pd.DataFrame:
    nombre = archivo.name.lower()
    if nombre.endswith(".csv"):
        df = pd.read_csv(archivo, encoding="utf-8", low_memory=False)
    elif nombre.endswith((".xlsx", ".xls")):
        df = pd.read_excel(archivo)
    else:
        raise ValueError("Formato no soportado. Sube un .csv o .xlsx")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def columnas_numericas(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]


def columnas_categoricas(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if c not in columnas_numericas(df)]


# ---------------------------------------------------------------------------
# Construcción de la tabla dinámica
# ---------------------------------------------------------------------------
def construir_pivote(
    df: pd.DataFrame,
    filas: list[str],
    columnas: list[str],
    valores: list[str],
    filtros: dict[str, list],
) -> pd.DataFrame:
    dff = df.copy()
    for col, seleccion in filtros.items():
        if seleccion:
            dff = dff[dff[col].isin(seleccion)]

    if not filas or not valores:
        return pd.DataFrame()

    if columnas:
        pivote = pd.pivot_table(
            dff,
            index=filas,
            columns=columnas,
            values=valores,
            aggfunc="sum",
            fill_value=0,
        )
        pivote.columns = [
            " | ".join(str(x) for x in c) if isinstance(c, tuple) else str(c)
            for c in pivote.columns
        ]
        pivote = pivote.reset_index()
    else:
        pivote = dff.groupby(filas, as_index=False)[valores].sum()

    # fila de Total general al principio
    fila_total = {c: "" for c in pivote.columns}
    if filas:
        fila_total[filas[0]] = "Total general"
    for c in pivote.columns:
        if c not in filas and pd.api.types.is_numeric_dtype(pivote[c]):
            fila_total[c] = pivote[c].sum()
    total_df = pd.DataFrame([fila_total])
    pivote = pd.concat([total_df, pivote], ignore_index=True)
    return pivote


# ---------------------------------------------------------------------------
# Exportación a Excel con formato "Estado del Ejercicio"
# ---------------------------------------------------------------------------
def exportar_excel_oref(
    pivote: pd.DataFrame,
    fuente: str,
    titulo: str,
    subtitulo: str,
    filas: list[str],
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"

    n_cols = len(pivote.columns)
    ultima_col_letra = get_column_letter(max(n_cols, 1))

    # --- Bloque de encabezado institucional -------------------------------
    ws.merge_cells(f"A1:{ultima_col_letra}1")
    ws["A1"] = "SADER — Secretaría de Agricultura y Desarrollo Rural"
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=11)
    ws["A1"].fill = PatternFill("solid", fgColor=ROJO_SADER)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells(f"A2:{ultima_col_letra}2")
    ws["A2"] = f"Fuente de datos: {FUENTES.get(fuente, fuente)}"
    ws["A2"].font = Font(italic=True, color="FFFFFF", size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=CAFE_SADER)
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.merge_cells(f"A4:{ultima_col_letra}4")
    ws["A4"] = titulo
    ws["A4"].font = Font(bold=True, size=13, color=ROJO_SADER)
    ws["A4"].alignment = Alignment(horizontal="center")

    ws.merge_cells(f"A5:{ultima_col_letra}5")
    ws["A5"] = subtitulo
    ws["A5"].font = Font(italic=True, size=10)
    ws["A5"].alignment = Alignment(horizontal="center")

    fila_encabezado = 7
    # --- Encabezados de columnas -------------------------------------------
    for j, col in enumerate(pivote.columns, start=1):
        celda = ws.cell(row=fila_encabezado, column=j, value=str(col))
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=ROJO_SADER)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = BORDE

    # --- Filas de datos ------------------------------------------------
    n_filas_agrupadoras = len(filas) if filas else 1
    for i, (_, fila) in enumerate(pivote.iterrows()):
        r = fila_encabezado + 1 + i
        es_total = str(fila.iloc[0]) == "Total general"
        for j, col in enumerate(pivote.columns, start=1):
            valor = fila[col]
            celda = ws.cell(row=r, column=j, value=valor)
            celda.border = BORDE
            if isinstance(valor, (int, float)) and col not in (filas or []):
                celda.number_format = "#,##0.00"
            if es_total:
                celda.font = Font(bold=True)
                celda.fill = PatternFill("solid", fgColor=CREMA_SADER)

    # --- Ajustes finales -----------------------------------------------
    ws.freeze_panes = ws.cell(row=fila_encabezado + 2, column=n_filas_agrupadoras + 1)
    for j, col in enumerate(pivote.columns, start=1):
        ancho = max(12, min(38, len(str(col)) + 4))
        ws.column_dimensions[get_column_letter(j)].width = ancho

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Interfaz Streamlit
# ---------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Tabla dinámica MAP / SICOP", layout="wide")
    st.title(" Constructor de reportes — MAP / SICOP")
    st.caption(
        "Integra reportes MAP y SICOP en un solo lugar, arma tu propia tabla "
        "dinámica y descárgala con el formato del Estado del Ejercicio."
    )

    with st.sidebar:
        st.header("1. Fuente de datos")
        fuente = st.radio("¿Qué reporte vas a trabajar?", ["MAP", "SICOP"], horizontal=True)
        archivo = st.file_uploader(f"Sube el archivo crudo de {fuente} (.csv o .xlsx)", type=["csv", "xlsx", "xls"])

    if not archivo:
        st.info("Sube un archivo en la barra lateral para empezar a armar tu reporte.")
        return

    df = cargar_datos(archivo)
    st.success(f"Archivo cargado: {archivo.name} — {len(df):,} filas, {len(df.columns)} columnas")

    cat_cols = columnas_categoricas(df)
    num_cols = columnas_numericas(df)

    st.subheader("2. Arma tu tabla dinámica")
    c1, c2, c3 = st.columns(3)
    with c1:
        filas = st.multiselect("Filas (agrupar por)", cat_cols, default=cat_cols[:2] if len(cat_cols) >= 2 else cat_cols)
    with c2:
        columnas_pivote = st.multiselect("Columnas (opcional, para pivotear)", [c for c in cat_cols if c not in filas])
    with c3:
        valores = st.multiselect("Valores (se suman)", num_cols, default=num_cols[:4] if len(num_cols) >= 4 else num_cols)

    with st.expander("Filtros (opcional)"):
        filtros = {}
        cols_filtro = st.multiselect("¿Qué columnas quieres filtrar?", cat_cols)
        for col in cols_filtro:
            opciones = sorted(df[col].dropna().unique().tolist(), key=str)
            filtros[col] = st.multiselect(f"Valores de {col}", opciones)

    pivote = construir_pivote(df, filas, columnas_pivote, valores, filtros)

    st.subheader("3. Vista previa")
    if pivote.empty:
        st.warning("Elige al menos una columna en Filas y una en Valores para ver la tabla.")
        return
    st.dataframe(pivote, use_container_width=True, height=420)

    st.subheader("4. Descargar")
    hoy = date.today().strftime("%d de %B de %Y")
    titulo_default = f"Estado del Ejercicio al {hoy}"
    titulo = st.text_input("Título del reporte", value=titulo_default)
    subtitulo = st.text_input("Subtítulo del reporte", value=f"Reporte {fuente} — datos seleccionados")

    excel_bytes = exportar_excel_oref(pivote, fuente, titulo, subtitulo, filas)
    nombre_archivo = f"Reporte_{fuente}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        " Descargar Excel con formato Estado del Ejercicio",
        data=excel_bytes,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
