"""
Constructor de Tabla Dinámica Presupuestal — MAP / SICOP
==========================================================
App Streamlit que integra los reportes crudos de MAP y SICOP, deja armar
cualquier reporte tipo tabla dinámica (agregar/quitar filas, columnas y
valores, filtrar por cualquier campo de la base — incluida Unidad
Responsable, Partida y Programa por nombre) y descarga el resultado en
Excel con el formato institucional del "Estado del Ejercicio".

Cómo correrla:
    pip install -r requirements.txt
    streamlit run pivot_presupuestal_app.py

Qué hace por ti automáticamente al cargar un archivo crudo:
    - Detecta la codificación (MAP = utf-8, SICOP = latin-1) sola.
    - Construye la Partida completa en SICOP (Capítulo+Concepto+Genérica+
      Específica) igual que en tus reportes actuales.
    - Junta los catálogos de catalogs/unidades.csv, catalogs/partidas.csv
      y catalogs/programas.csv para mostrar nombres, no solo códigos, en
      los filtros y en las filas del reporte. Si un código no está en el
      catálogo, muestra el código tal cual — puedes ir agregando filas a
      esos CSV para completar la cobertura.
    - Calcula, para cada familia de importes (Original, Modificado,
      Comprometido, Ejercido, Reservas, etc.), el total "Anual" y el
      "Al periodo" (acumulado de enero al mes que elijas), igual que en
      el Estado del Ejercicio.
    - Calcula el Importe Disponible como Modificado − Ejercido − Comprometido
      (fórmula verificada contra tu archivo de ejemplo, cuadra al centavo).
"""

from __future__ import annotations

import io
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Paleta y estilos — tomados directamente de formato_estado_del_ejercicio.xlsx
# ---------------------------------------------------------------------------
BURDEOS = "621333"     # encabezados de columnas
VERDE = "1E5B4F"       # encabezados de grupo ("Anual" / "Al periodo")
GRIS_TOTAL = "D9D9D9"  # fila de Total general
FORMATO_MONEDA = '_-"$"* #,##0.00_-;\\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'

THIN_GRIS = Side(style="thin", color="808080")
BORDE = Border(left=THIN_GRIS, right=THIN_GRIS, top=THIN_GRIS, bottom=THIN_GRIS)
THIN_BLANCO = Side(style="thin", color="FFFFFF")
BORDE_ENCABEZADO = Border(left=THIN_BLANCO, right=THIN_BLANCO, top=THIN_BLANCO, bottom=THIN_BLANCO)

MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
NOMBRES_MES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
               "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

CATALOGS_DIR = Path(__file__).parent / "catalogs"

# ---------------------------------------------------------------------------
# Definición de familias de importes por fuente
# (columna total ya existente en el crudo, o None si hay que sumarla)
# ---------------------------------------------------------------------------
def _cols_sicop(prefijo_2letras: list[str]) -> list[str]:
    return prefijo_2letras


FAMILIAS_SICOP = {
    "ORIGINAL": (["OREN", "ORFE", "ORMR", "ORAB", "ORMY", "ORJN", "ORJL", "ORAG", "ORSE", "OROC", "ORNO", "ORDI"], "ORIGINAL"),
    "AMPLIACION": (["AMEN", "AMFE", "AMMR", "AMAB", "AMMY", "AMJN", "AMJL", "AMAG", "AMSE", "AMOC", "AMNO", "AMDI"], "AMPLIACION"),
    "REDUCCION": (["REEN", "REFE", "REMR", "REAB", "REMY", "REJN", "REJL", "REAG", "RESE", "REOC", "RENO", "REDI"], "REDUCCION"),
    "RESERVAS": (["RESERVA_ENE", "RESERVA_FEB", "RESERVA_MZO", "RESERVA_ABR", "RESERVA_MAY", "RESERVA_JUN",
                  "RESERVA_JUL", "RESERVA_AGO", "RESERVA_SEP", "RESERVA_OCT", "RESERVA_NOV", "RESERVA_DIC"], "RESERVAS"),
    "MODIFICADO_AUTORIZADO": (["MOEN", "MOFE", "MOMR", "MOAB", "MOMY", "MOJN", "MOJL", "MOAG", "MOSE", "MOOC", "MONO", "MODI"], "MODIFICADO_AUTORIZADO"),
    "COMPROMETIDO": (["COEN", "COFE", "COMR", "COAB", "COMY", "COJN", "COJL", "COAG", "COSE", "COOC", "CONO", "CODI"], "COMPROMETIDO"),
    "EJERCIDO": (["EJEN", "EJFE", "EJMR", "EJAB", "EJMY", "EJJN", "EJJL", "EJAG", "EJSE", "EJOC", "EJNO", "EJDI"], "EJERCIDO"),
    "DEVENGADO": (["DVEN", "DVFE", "DVMR", "DVAB", "DVMY", "DVJN", "DVJL", "DVAG", "DVSE", "DVOC", "DVNO", "DVDI"], "DEVENGADO"),
}
CLAVES_SICOP = {"modificado": "MODIFICADO_AUTORIZADO", "ejercido": "EJERCIDO", "comprometido": "COMPROMETIDO"}

FAMILIAS_MAP = {
    f"{pref}": ([f"{pref}_{m}" for m in MESES], None)
    for pref in ["ORI", "AMP", "RED", "MOD", "CONG", "DESCONG", "EJE"]
}
CLAVES_MAP = {"modificado": "MOD", "ejercido": "EJE", "comprometido": None}

NOMBRES_FAMILIA = {
    "ORIGINAL": "Original", "ORI": "Original",
    "AMPLIACION": "Ampliación", "AMP": "Ampliación",
    "REDUCCION": "Reducción", "RED": "Reducción",
    "RESERVAS": "Reservas",
    "MODIFICADO_AUTORIZADO": "Modificado", "MOD": "Modificado",
    "COMPROMETIDO": "Comprometido",
    "EJERCIDO": "Ejercido", "EJE": "Ejercido",
    "DEVENGADO": "Devengado",
    "CONG": "Congelado", "DESCONG": "Descongelado",
}

FUENTES = {
    "MAP": "Módulo de Adecuaciones Presupuestarias (MAP)",
    "SICOP": "Sistema de Contabilidad y Presupuesto (SICOP)",
}


# ---------------------------------------------------------------------------
# Catálogos (código -> nombre) — se cargan de /catalogs, editables por ti
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def cargar_catalogo(nombre_archivo: str, col_codigo: str, col_nombre: str) -> dict[str, str]:
    ruta = CATALOGS_DIR / nombre_archivo
    if not ruta.exists():
        return {}
    try:
        cat = pd.read_csv(ruta, dtype=str)
        cat[col_codigo] = cat[col_codigo].str.strip()
        return dict(zip(cat[col_codigo], cat[col_nombre]))
    except Exception:
        return {}


def etiqueta_con_nombre(codigo, catalogo: dict[str, str]) -> str:
    if pd.isna(codigo):
        return "(sin dato)"
    clave = str(codigo).strip()
    clave_num = clave.split(".")[0] if clave.replace(".", "", 1).isdigit() else clave
    nombre = catalogo.get(clave) or catalogo.get(clave_num)
    return f"{clave} — {nombre}" if nombre else clave


# ---------------------------------------------------------------------------
# Carga y preparación de datos crudos
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Leyendo archivo...")
def cargar_crudo(archivo_bytes: bytes, nombre_archivo: str, fuente: str) -> pd.DataFrame:
    buffer = io.BytesIO(archivo_bytes)
    if nombre_archivo.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(buffer)
    else:
        encodings = ["utf-8", "latin-1", "cp1252"]
        df = None
        for enc in encodings:
            try:
                buffer.seek(0)
                df = pd.read_csv(buffer, encoding=enc, low_memory=False)
                break
            except UnicodeDecodeError:
                continue
        if df is None:
            raise ValueError("No se pudo leer el archivo con ninguna codificación común (utf-8/latin-1/cp1252).")
    df.columns = [str(c).strip() for c in df.columns]

    if fuente == "SICOP":
        for c in ["CAPITULO", "CONCEPTO", "PARTIDA_GENERICA", "PARTIDA_ESPECIFICA"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
        if {"CAPITULO", "CONCEPTO", "PARTIDA_GENERICA", "PARTIDA_ESPECIFICA"}.issubset(df.columns):
            df["Partida"] = (df["CAPITULO"] * 10000 + df["CONCEPTO"] * 1000
                              + df["PARTIDA_GENERICA"] * 100 + df["PARTIDA_ESPECIFICA"])
        if "ID_UNIDAD" in df.columns:
            df["Unidad Responsable"] = df["ID_UNIDAD"].astype(str).str.strip()
        if "PROGRAMA_PRESUPUESTARIO" in df.columns:
            df["Programa"] = df["PROGRAMA_PRESUPUESTARIO"].astype(str).str.strip()
    elif fuente == "MAP":
        if "PARTIDA" in df.columns:
            df["Partida"] = pd.to_numeric(df["PARTIDA"], errors="coerce")
        if "UNIDAD" in df.columns:
            df["Unidad Responsable"] = df["UNIDAD"].astype(str).str.strip()
        if "PROGRAMA" in df.columns:
            df["Programa"] = df["PROGRAMA"].astype(str).str.strip()

    return df


def enriquecer_con_catalogos(df: pd.DataFrame) -> pd.DataFrame:
    cat_ur = cargar_catalogo("unidades.csv", "codigo_ur", "nombre_ur")
    cat_partidas = cargar_catalogo("partidas.csv", "partida", "nombre_partida")
    cat_programas = cargar_catalogo("programas.csv", "programa", "nombre_programa")

    df = df.copy()
    if "Unidad Responsable" in df.columns:
        df["Unidad Responsable (nombre)"] = df["Unidad Responsable"].apply(lambda x: etiqueta_con_nombre(x, cat_ur))
    if "Partida" in df.columns:
        df["Partida (nombre)"] = df["Partida"].apply(lambda x: etiqueta_con_nombre(x, cat_partidas))
    if "Programa" in df.columns:
        df["Programa (nombre)"] = df["Programa"].apply(lambda x: etiqueta_con_nombre(x, cat_programas))
    return df


def agregar_periodos_y_disponible(df: pd.DataFrame, fuente: str, mes_corte_idx: int) -> tuple[pd.DataFrame, list[str]]:
    """Agrega columnas '<Familia> (Anual)' y '<Familia> (Al <mes>)' por cada
    familia de importes disponible, más 'Disponible (Anual)'/'Disponible (Al <mes>)'.
    Regresa el df enriquecido y la lista de columnas de valor calculadas (en orden lógico)."""
    familias = FAMILIAS_SICOP if fuente == "SICOP" else FAMILIAS_MAP
    claves = CLAVES_SICOP if fuente == "SICOP" else CLAVES_MAP
    mes_label = NOMBRES_MES[mes_corte_idx].capitalize()

    df = df.copy()
    columnas_valor = []
    for fam, (cols_mensuales, col_total_existente) in familias.items():
        presentes = [c for c in cols_mensuales if c in df.columns]
        if not presentes:
            continue
        nombre_bonito = NOMBRES_FAMILIA.get(fam, fam)
        col_anual = f"{nombre_bonito} (Anual)"
        if col_total_existente and col_total_existente in df.columns:
            df[col_anual] = pd.to_numeric(df[col_total_existente], errors="coerce").fillna(0)
        else:
            df[col_anual] = df[presentes].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)

        cols_al_periodo = [c for c in cols_mensuales[: mes_corte_idx + 1] if c in df.columns]
        col_periodo = f"{nombre_bonito} (Al {mes_label})"
        df[col_periodo] = df[cols_al_periodo].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if cols_al_periodo else 0.0

        columnas_valor += [col_anual, col_periodo]

    # Disponible = Modificado - Ejercido - Comprometido (verificado contra el
    # archivo de ejemplo: cuadra al centavo con las columnas "Importe Disponible").
    nombre_mod = NOMBRES_FAMILIA.get(claves["modificado"])
    nombre_eje = NOMBRES_FAMILIA.get(claves["ejercido"])
    nombre_com = NOMBRES_FAMILIA.get(claves["comprometido"]) if claves["comprometido"] else None

    col_mod_anual, col_mod_periodo = f"{nombre_mod} (Anual)", f"{nombre_mod} (Al {mes_label})"
    col_eje_anual = f"{nombre_eje} (Anual)"
    if col_mod_anual in df.columns and col_eje_anual in df.columns:
        com_anual = df[f"{nombre_com} (Anual)"] if nombre_com and f"{nombre_com} (Anual)" in df.columns else 0
        com_periodo = df[f"{nombre_com} (Al {mes_label})"] if nombre_com and f"{nombre_com} (Al {mes_label})" in df.columns else 0
        df["Disponible (Anual)"] = df[col_mod_anual] - df[col_eje_anual] - com_anual
        df["Disponible (Al " + mes_label + ")"] = df[col_mod_periodo] - df[col_eje_anual] - com_periodo
        columnas_valor += ["Disponible (Anual)", f"Disponible (Al {mes_label})"]

    return df, columnas_valor


def aplicar_depuracion_sicop(df: pd.DataFrame) -> pd.DataFrame:
    """Reglas que corrigieron discrepancias del Estado del Ejercicio en SICOP:
    excluir capítulo 1000 (servicios personales), capítulo 6000 (inversión) y
    la partida 39801. Revisa que apliquen a lo que quieres reportar antes de
    activarlas — no siempre corresponden a todos los reportes."""
    dff = df.copy()
    if "CAPITULO" in dff.columns:
        dff = dff[~dff["CAPITULO"].isin([1, 6])]
    if "Partida" in dff.columns:
        dff = dff[dff["Partida"] != 39801]
    return dff


def columnas_categoricas(df: pd.DataFrame, columnas_valor: list[str]) -> list[str]:
    return [c for c in df.columns if c not in columnas_valor and not pd.api.types.is_float_dtype(df[c])] or \
        [c for c in df.columns if c not in columnas_valor]


# ---------------------------------------------------------------------------
# Construcción de la tabla dinámica
# ---------------------------------------------------------------------------
def construir_pivote(df, filas, columnas, valores, filtros) -> pd.DataFrame:
    dff = df.copy()
    for col, seleccion in filtros.items():
        if seleccion:
            dff = dff[dff[col].isin(seleccion)]

    if not filas or not valores:
        return pd.DataFrame()

    if columnas:
        pivote = pd.pivot_table(dff, index=filas, columns=columnas, values=valores, aggfunc="sum", fill_value=0)
        pivote.columns = [" | ".join(str(x) for x in c) if isinstance(c, tuple) else str(c) for c in pivote.columns]
        pivote = pivote.reset_index()
    else:
        pivote = dff.groupby(filas, as_index=False)[valores].sum()

    fila_total = {c: "" for c in pivote.columns}
    if filas:
        fila_total[filas[0]] = "Total general"
    for c in pivote.columns:
        if c not in filas and pd.api.types.is_numeric_dtype(pivote[c]):
            fila_total[c] = pivote[c].sum()
    pivote = pd.concat([pd.DataFrame([fila_total]), pivote], ignore_index=True)
    return pivote


# ---------------------------------------------------------------------------
# Exportación a Excel con el formato del Estado del Ejercicio
# ---------------------------------------------------------------------------
def exportar_excel_oref(pivote: pd.DataFrame, fuente: str, linea1: str, linea2: str,
                         titulo: str, subtitulo: str, filas: list[str]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"
    ws.sheet_view.showGridLines = False

    n_cols = max(len(pivote.columns), 1)
    ultima_col = get_column_letter(n_cols)

    c = ws.cell(row=1, column=n_cols, value=linea1)
    c.font = Font(name="Arial", size=12, bold=True)
    c.alignment = Alignment(horizontal="right")

    c = ws.cell(row=2, column=n_cols, value=linea2)
    c.font = Font(name="Arial", size=11, bold=True)
    c.alignment = Alignment(horizontal="right")

    ws.merge_cells(f"A5:{ultima_col}5")
    c = ws["A5"]; c.value = titulo
    c.font = Font(name="Calibri", size=14, bold=True)
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[5].height = 18.75

    ws.merge_cells(f"A6:{ultima_col}6")
    c = ws["A6"]; c.value = subtitulo
    c.font = Font(name="Calibri", size=12, italic=True)
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[6].height = 15.75

    fila_encabezado = 9
    ws.row_dimensions[fila_encabezado].height = 30
    for j, col in enumerate(pivote.columns, start=1):
        celda = ws.cell(row=fila_encabezado, column=j, value=str(col))
        celda.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=BURDEOS)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = BORDE_ENCABEZADO

    n_filas_agrupadoras = max(len(filas), 1)
    for i, (_, fila) in enumerate(pivote.iterrows()):
        r = fila_encabezado + 1 + i
        es_total = str(fila.iloc[0]) == "Total general"
        for j, col in enumerate(pivote.columns, start=1):
            valor = fila[col]
            celda = ws.cell(row=r, column=j, value=valor)
            celda.border = BORDE
            celda.alignment = Alignment(vertical="center", horizontal="center" if col in (filas or []) else None)
            if isinstance(valor, (int, float)) and col not in (filas or []):
                celda.number_format = FORMATO_MONEDA
            if es_total:
                celda.font = Font(name="Arial", size=11, bold=True)
                celda.fill = PatternFill("solid", fgColor=GRIS_TOTAL)
            else:
                celda.font = Font(name="Calibri", size=11)

    ws.freeze_panes = ws.cell(row=fila_encabezado + 2, column=n_filas_agrupadoras + 1)
    ws.print_title_rows = f"1:{fila_encabezado}"
    for j, col in enumerate(pivote.columns, start=1):
        ancho = max(13, min(40, len(str(col)) + 6))
        ws.column_dimensions[get_column_letter(j)].width = ancho

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def fecha_desde_nombre_archivo(nombre: str) -> str | None:
    m = re.search(r"(\d{1,2})[-_]([A-ZÁÉÍÓÚa-záéíóú]+)[-_](\d{4})", nombre)
    if not m:
        return None
    dia, mes_txt, anio = m.groups()
    mes_txt = mes_txt.lower()
    meses_largos = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                    "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    mes_encontrado = next((m for m in meses_largos if m.startswith(mes_txt[:3])), None)
    if not mes_encontrado:
        return None
    return f"{int(dia)} de {mes_encontrado} de {anio}"


# ---------------------------------------------------------------------------
# Interfaz Streamlit
# ---------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Tabla dinámica MAP / SICOP", layout="wide")
    st.title(" Constructor de reportes — MAP / SICOP")
    st.caption(
        "Integra reportes MAP y SICOP en un solo lugar, arma cualquier reporte "
        "tipo tabla dinámica y descárgalo con el formato del Estado del Ejercicio."
    )

    with st.sidebar:
        st.header("1. Fuente de datos")
        fuente = st.radio("¿Qué vas a cargar?", ["MAP", "SICOP"], horizontal=True)
        archivo = st.file_uploader(f"Archivo crudo de {fuente} (.csv o .xlsx)", type=["csv", "xlsx", "xls"])
        st.divider()
        st.header("2. Periodo")
        hoy = date.today()
        mes_corte_idx = st.selectbox(
            "Corte de 'Al periodo' (acumulado enero → este mes)",
            options=list(range(12)),
            format_func=lambda i: NOMBRES_MES[i].capitalize(),
            index=min(hoy.month - 1, 11),
        )
        depurar_sicop = False
        if fuente == "SICOP":
            depurar_sicop = st.checkbox(
                "Excluir capítulo 1000, 6000 y partida 39801",
                value=False,
                help="Reglas que corrigieron antes discrepancias del Estado del Ejercicio. Revisa si aplican a tu reporte.",
            )

    if not archivo:
        st.info("Sube un archivo en la barra lateral para empezar a armar tu reporte.")
        return

    try:
        df = cargar_crudo(archivo.getvalue(), archivo.name, fuente)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        return

    df = enriquecer_con_catalogos(df)
    df, columnas_familia = agregar_periodos_y_disponible(df, fuente, mes_corte_idx)
    if fuente == "SICOP" and depurar_sicop:
        df = aplicar_depuracion_sicop(df)

    st.success(f"Archivo cargado: {archivo.name} — {len(df):,} filas")

    columnas_no_valor = [c for c in df.columns if c not in columnas_familia]
    columnas_num_extra = [c for c in columnas_no_valor if pd.api.types.is_numeric_dtype(df[c])]
    columnas_cat = [c for c in columnas_no_valor if c not in columnas_num_extra]

    st.subheader("2. Arma tu tabla dinámica")
    campos_sugeridos = [c for c in ["Unidad Responsable (nombre)", "Partida (nombre)", "Programa (nombre)"] if c in columnas_cat]
    c1, c2, c3 = st.columns(3)
    with c1:
        filas = st.multiselect("Filas (agrupar por)", columnas_cat, default=campos_sugeridos[:2] or columnas_cat[:1])
    with c2:
        columnas_pivote = st.multiselect("Columnas (opcional, para pivotear)", [c for c in columnas_cat if c not in filas])
    with c3:
        mostrar_mensuales = st.checkbox("Incluir columnas mensuales individuales en 'Valores'", value=False)
        opciones_valor = columnas_familia + (columnas_num_extra if mostrar_mensuales else [])
        default_valores = [c for c in columnas_familia if "(Anual)" in c][:4]
        valores = st.multiselect("Valores (se suman)", opciones_valor, default=default_valores)

    with st.expander("Filtros — cualquier columna de la base"):
        filtros = {}
        cols_filtro_default = [c for c in campos_sugeridos]
        cols_filtro = st.multiselect("¿Qué columnas quieres filtrar?", columnas_cat, default=cols_filtro_default)
        for col in cols_filtro:
            opciones = sorted(df[col].dropna().astype(str).unique().tolist())
            filtros[col] = st.multiselect(f"Valores de «{col}»", opciones, key=f"filtro_{col}")

    dff = df.copy()
    for col, seleccion in filtros.items():
        if seleccion:
            dff = dff[dff[col].astype(str).isin(seleccion)]

    pivote = construir_pivote(dff, filas, columnas_pivote, valores, {})

    st.subheader("3. Vista previa")
    if pivote.empty:
        st.warning("Elige al menos una columna en Filas y una en Valores para ver la tabla.")
        return
    st.dataframe(pivote, use_container_width=True, height=420)

    st.subheader("4. Descargar")
    fecha_archivo = fecha_desde_nombre_archivo(archivo.name)
    titulo_default = f"Estado del Ejercicio al {fecha_archivo}" if fecha_archivo else f"Estado del Ejercicio al {hoy.day} de {NOMBRES_MES[hoy.month-1]} de {hoy.year}"
    colA, colB = st.columns(2)
    with colA:
        linea1 = st.text_input("Encabezado — línea 1", value="Unidad de Administración y Finanzas")
        titulo = st.text_input("Título del reporte", value=titulo_default)
    with colB:
        linea2 = st.text_input("Encabezado — línea 2", value="Dirección General de Programación, Presupuesto y Finanzas")
        subtitulo = st.text_input("Subtítulo del reporte", value=f"Reporte {fuente} — datos seleccionados")

    excel_bytes = exportar_excel_oref(pivote, fuente, linea1, linea2, titulo, subtitulo, filas)
    nombre_archivo = f"Reporte_{fuente}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        " Descargar Excel con formato Estado del Ejercicio",
        data=excel_bytes,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
