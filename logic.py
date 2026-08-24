"""
Lógica de negocio (parsing, limpieza, cálculo de riesgo) para la
Tabla de Carga Semanal — Tienda Propia (Volcom / Rusty)
=========================================================

Primera versión del script de Streamlit que reemplaza el notebook de Colab
ejecutado celda a celda. Implementa el entregable fijo de los lunes descrito
en `Proceso_Semanal_Acordado_2026-08-17.md` y las reglas confirmadas por
Francisco en `Prototipo_Automatizacion_Reglas_Confirmadas_2026-08-24.md`.

Entrada esperada: archivo "4 Semanas - TP (Rip Curl & Prosurf) [DD.MM.AAAA].xlsx"
(hoja "Resumen 4 Semanas - Pro + Rip").

IMPORTANTE — esta es una ESTRUCTURA BASE, no un script ya probado contra el
archivo real dentro de esta sesión (no había un .xlsx adjunto para probar).
La lectura de las columnas 0-137 usa POSICIÓN (no nombre), porque el
Diccionario de Datos ya confirmó ese orden exacto contra datos reales. La
detección de las 28 tiendas (columnas 138-277) es HEURÍSTICA (busca patrones
"NNNN-NOMBRE" en las filas de encabezado) porque el diccionario no precisa en
qué fila exacta queda el nombre de cada tienda. La primera vez que se corra
contra un archivo real, revisar el panel "Diagnóstico de estructura" antes de
confiar en la tabla de salida, y ajustar `FILA_ENCABEZADO_FIJA` /
`COL_INICIO_TIENDAS_FIJA` manualmente si la detección automática falla.

No inventa datos: todo lo que no se pueda derivar directamente del archivo
(p.ej. una "Prioridad" numérica no definida por Francisco) se marca
explícitamente como criterio operativo propuesto, no como regla de negocio
confirmada.
"""

import io
import re

import numpy as np
import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# 1. REGLAS DE NEGOCIO CONFIRMADAS (ver docs del proyecto — no modificar sin
#    que Francisco confirme el cambio explícitamente)
# ---------------------------------------------------------------------------

MARCAS_POR_DEFECTO = ["Volcom", "Rusty"]

# Regla 4.2 — exclusión permanente, producto operativo (empaque/regalo).
FAMILIAS_EXCLUIR = ["Bolsa de Papel"]

# Regla 4.3 — cobertura semanal = Stock actual tienda / Venta Sem -1.
# Riesgo si Venta Sem -1 > 0 Y Cobertura < 2 semanas (atado al lead time de
# reposición real; aplica igual a todas las familias por ahora).
UMBRAL_COBERTURA_RIESGO = 2.0

# Regla 4.4 — tras la fusión societaria Prosurf -> Maui, la columna correcta
# para bodega disponible es el TOTAL agregado de todos los CD, no solo
# "01 - CD Tiendas Propias - Prosurf" (que ahora da 0 en el 100% de las filas).
COL_BODEGA_DISPONIBLE = "Stock Disponible CD - Pack"

# Reglas de tiendas confirmadas (Diccionario de Datos M300, sección "Reglas
# de tiendas") — prefijos de código de tienda.
PREFIJO_TIENDAS_EXCLUIR = ("50",)   # 50xx: excluir siempre del análisis
PREFIJO_TIENDAS_RIPCURL = ("30",)   # 30xx: casi sin stock V/R, se excluyen del análisis V/R
CODIGOS_MULTIMARCA = ("2022", "4006")  # multimarca — se excluyen del análisis estándar de TP

# Regla 4.1 — vigencia de temporada, SOLO aplica a Tienda Propia (no a Outlet).
def _temporada_vigente(temporada: str, anio_producto) -> bool:
    if pd.isna(temporada):
        return False
    t = str(temporada).strip().upper()
    try:
        anio = int(anio_producto)
    except (TypeError, ValueError):
        anio = None
    if t == "TODA TEMPORADA":
        return True
    if t == "INVIERNO":
        return anio == 2026
    if t == "VERANO":
        return anio in (2026, 2027)
    # COLEGIAL y cualquier otra combinación quedan fuera (excluidas del análisis TP).
    return False


# ---------------------------------------------------------------------------
# 2. ESTRUCTURA DE COLUMNAS 0-137 (fija, confirmada en el Diccionario de
#    Datos M300 Parte B, sección 3). Se usa por POSICIÓN, no por nombre de
#    encabezado, para no depender de bugs como espacios en blanco al inicio
#    del nombre de columna.
# ---------------------------------------------------------------------------

def _bloque_cd(prefijo_sufijo: str = "") -> list[str]:
    """Genera los 44 nombres de un bloque de 11 tipos de CD x (Total + 3 compañías)."""
    tipos_cd = [
        "01 - CD Tiendas Propias",
        "02 - CD Outlet",
        "03 - CD Liquidadora",
        "04 - CD Mayoreo Multitiendas",
        "05 - CD Mayoreo Norte/Sur",
        "06 - CD Exportación",
        "07 - CD E-Commerce",
        "08 - CD Reserva Stock",
        "09 - CD Traspaso",
        "10 - CD Consignación",
        "11 - CD Wholesale",
    ]
    nombres = []
    for tipo in tipos_cd:
        for compania in ("Total", "Maui", "Rip Curl", "Prosurf"):
            nombres.append(f"{tipo} - {compania}{prefijo_sufijo}")
    return nombres


COLUMNAS_FIJAS_0_137 = (
    [
        "Id Estilo", "Desc Estilo", "Color", "Id Pack", "Desc Pack", "Marca",
        "Linea", "Departamento", "Familia", "Sub_Familia", "Temporada",
        "Año Producto", "Definicion", "Genero", "Unid x Inner",
        "Inner x Master", "Unid x Master", "Producto", "Desc Color",
        "Temporada Local", "Status Precio",
    ]
    + [
        "Venta $ Neta Semana -1", "Costo $ Semana -1", "Venta Un Semana -4U",
        "Venta Un Semana -3U", "Venta Un Semana -2U", "Venta Un Semana -1U",
        "Venta Un 4 Semanas", "Venta Un Acum 1-Ene",
    ]
    + ["Stock Disponible CD - Pack"] + _bloque_cd() + ["_sep_cd_pack"]
    + ["Stock Disponible CD - Sku"] + _bloque_cd(" (Sku)") + ["_sep_cd_sku"]
    + [
        "Mercadería en Tránsito (Marítimo)", "Pend Envío", "Pend Recibido",
        "Stock Disponible Tiendas - MOM", "Stock Disponible CD - Ordenamiento",
        "Comentarios", "N Tiendas",
    ]
    + [
        "Parámetro", "Semanas Stock Tienda", "Sem Stock sin navidad",
        "Stock Total", "Semanas Stock Cadena", "Precio Normal (Bruto)",
        "Precio Promedio Ultima Semana", "Ultimo Costo", "DCTO Actual",
        "MG Actual",
    ]
)

METRICAS_POR_TIENDA = ["Vta # 4 Sem", "Vta # Sem -1", "Stock # MOM", "Pend. Envío", "Pend. Recib"]

assert len(COLUMNAS_FIJAS_0_137) == 138, f"Se esperaban 138 columnas fijas, hay {len(COLUMNAS_FIJAS_0_137)}"


# ---------------------------------------------------------------------------
# 3. LECTURA Y DETECCIÓN DE ESTRUCTURA
# ---------------------------------------------------------------------------

PATRON_TIENDA = re.compile(r"^\d{3,4}\s*-\s*.+")


@st.cache_data(show_spinner=False)
def leer_hojas_disponibles(archivo_bytes: bytes) -> list[str]:
    xls = pd.ExcelFile(io.BytesIO(archivo_bytes))
    return xls.sheet_names


@st.cache_data(show_spinner=False)
def detectar_fila_encabezado(archivo_bytes: bytes, hoja: str, max_filas: int = 10) -> int:
    """Busca la fila que contiene 'Marca' y 'Familia' como valores de celda —
    esas son las columnas dimensión que sabemos que existen siempre."""
    preview = pd.read_excel(io.BytesIO(archivo_bytes), sheet_name=hoja, header=None, nrows=max_filas)
    for i in range(len(preview)):
        valores = preview.iloc[i].astype(str).str.strip()
        if "Marca" in valores.values and "Familia" in valores.values:
            return i
    raise ValueError(
        "No se pudo detectar automáticamente la fila de encabezados en las "
        f"primeras {max_filas} filas (se buscó una fila con 'Marca' y "
        "'Familia'). Revisar el archivo manualmente e indicar la fila con "
        "el control 'Forzar fila de encabezado' en la barra lateral."
    )


def detectar_nombres_tienda(raw_encabezados: pd.DataFrame, col_inicio: int, n_cols: int) -> list[str]:
    """Escanea TODAS las filas de encabezado (0 hasta la fila de encabezado
    inclusive) buscando el patrón 'NNNN-NOMBRE' dentro de cada bloque de 5
    columnas y propaga el valor encontrado a las 5 columnas del bloque
    (por si la celda de Excel estaba combinada y pandas solo la deja en una)."""
    nombres = [None] * n_cols
    for _, fila in raw_encabezados.iterrows():
        for j in range(n_cols):
            col_idx = col_inicio + j
            if col_idx >= raw_encabezados.shape[1]:
                continue
            val = fila.iloc[col_idx]
            if pd.notna(val) and PATRON_TIENDA.match(str(val).strip()):
                nombres[j] = str(val).strip()

    for bloque_inicio in range(0, n_cols, 5):
        bloque_fin = min(bloque_inicio + 5, n_cols)
        nombre_bloque = next((nombres[j] for j in range(bloque_inicio, bloque_fin) if nombres[j]), None)
        for j in range(bloque_inicio, bloque_fin):
            if nombres[j] is None:
                nombres[j] = nombre_bloque or f"Tienda desconocida (bloque col {col_inicio + bloque_inicio})"
    return nombres


@st.cache_data(show_spinner=False)
def cargar_y_estructurar(archivo_bytes: bytes, hoja: str, fila_encabezado: int) -> tuple[pd.DataFrame, dict]:
    """Lee el archivo crudo y devuelve un DataFrame en formato ANCHO (una fila
    por Estilo+Color/Pack) con nombres de columna canónicos, más un dict de
    diagnóstico para mostrar en pantalla."""

    crudo = pd.read_excel(io.BytesIO(archivo_bytes), sheet_name=hoja, header=None)
    # eliminar columnas 100% vacías al final (el diccionario menciona 14 de sobra)
    crudo = crudo.dropna(axis=1, how="all")

    n_col_total = crudo.shape[1]
    n_col_tiendas = n_col_total - 138
    n_col_tiendas = max(n_col_tiendas - (n_col_tiendas % 5), 0)  # múltiplo de 5

    encabezados_para_deteccion = crudo.iloc[: fila_encabezado + 1]
    nombres_tienda_por_col = detectar_nombres_tienda(encabezados_para_deteccion, 138, n_col_tiendas)

    datos = crudo.iloc[fila_encabezado + 1 :].reset_index(drop=True)

    columnas_fijas = COLUMNAS_FIJAS_0_137[: min(138, n_col_total)]
    nombres_columnas = list(columnas_fijas)
    for j in range(n_col_tiendas):
        tienda = nombres_tienda_por_col[j]
        metrica = METRICAS_POR_TIENDA[j % 5]
        nombres_columnas.append(f"{tienda} || {metrica}")

    datos = datos.iloc[:, : len(nombres_columnas)]
    nombres_columnas, nombres_duplicados = _deduplicar_nombres(nombres_columnas)
    datos.columns = nombres_columnas

    # quitar columnas separadoras vacías generadas por el diccionario (_sep_*)
    datos = datos.drop(columns=[c for c in datos.columns if c.startswith("_sep_")], errors="ignore")

    tiendas_detectadas = sorted(set(nombres_tienda_por_col))
    diagnostico = {
        "fila_encabezado_usada": fila_encabezado,
        "columnas_totales_crudo": n_col_total,
        "columnas_esperadas_dictamen": 278,
        "n_tiendas_detectadas": len(tiendas_detectadas),
        "tiendas_detectadas": tiendas_detectadas,
        "filas_datos": len(datos),
        "nombres_duplicados": nombres_duplicados,
    }
    return datos, diagnostico


def _deduplicar_nombres(nombres: list[str]) -> tuple[list[str], list[str]]:
    """Garantiza nombres de columna únicos (requisito de pandas/Arrow para
    poder mostrar/exportar el DataFrame). Si el mismo nombre de tienda quedó
    asignado a más de un bloque de 5 columnas —lo cual señala que la
    detección automática de tiendas encontró el mismo texto en dos bloques
    distintos y probablemente necesita revisión manual—, se le agrega un
    sufijo [dup2], [dup3], etc. y se reporta en el diagnóstico para que quede
    visible, en vez de fallar silenciosamente o mezclar datos de dos bloques
    bajo el mismo nombre."""
    conteo: dict[str, int] = {}
    resultado = []
    duplicados: list[str] = []
    for nombre in nombres:
        conteo[nombre] = conteo.get(nombre, 0) + 1
        if conteo[nombre] == 1:
            resultado.append(nombre)
        else:
            resultado.append(f"{nombre} [dup{conteo[nombre]}]")
            if nombre not in duplicados:
                duplicados.append(nombre)
    return resultado, sorted(duplicados)


# ---------------------------------------------------------------------------
# 4. TRANSFORMACIÓN A FORMATO LARGO (SKU/Pack x Tienda) + LIMPIEZA
# ---------------------------------------------------------------------------

COLUMNAS_DIMENSION = [
    "Id Estilo", "Desc Estilo", "Color", "Id Pack", "Desc Pack", "Marca",
    "Linea", "Departamento", "Familia", "Sub_Familia", "Temporada",
    "Año Producto", "Definicion", "Genero", "Unid x Inner", "Inner x Master",
    "Unid x Master", COL_BODEGA_DISPONIBLE, "Semanas Stock Tienda",
]


def construir_formato_largo(datos_ancho: pd.DataFrame) -> pd.DataFrame:
    columnas_tienda = [c for c in datos_ancho.columns if " || " in c]
    tiendas = sorted({c.split(" || ")[0] for c in columnas_tienda})

    columnas_dim_presentes = [c for c in COLUMNAS_DIMENSION if c in datos_ancho.columns]
    piezas = []
    for tienda in tiendas:
        cols_tienda = {m: f"{tienda} || {m}" for m in METRICAS_POR_TIENDA if f"{tienda} || {m}" in datos_ancho.columns}
        if len(cols_tienda) < 2:
            continue
        sub = datos_ancho[columnas_dim_presentes + list(cols_tienda.values())].copy()
        sub = sub.rename(columns={v: k for k, v in cols_tienda.items()})
        sub["Tienda"] = tienda
        piezas.append(sub)

    largo = pd.concat(piezas, ignore_index=True) if piezas else pd.DataFrame()
    return largo


def limpiar_y_filtrar(
    largo: pd.DataFrame,
    marcas: list[str],
    excluir_familias: list[str],
    incluir_ripcurl: bool = False,
    incluir_multimarca: bool = False,
) -> pd.DataFrame:
    df = largo.copy()

    # Regla 4.5 — negativos en stock se tratan como 0.
    for col in ["Stock # MOM", COL_BODEGA_DISPONIBLE]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").clip(lower=0)
    for col in ["Vta # 4 Sem", "Vta # Sem -1"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df[df["Marca"].isin(marcas)]
    df = df[~df["Familia"].astype(str).str.strip().str.lower().isin([f.lower() for f in excluir_familias])]

    codigo_tienda = df["Tienda"].astype(str).str.extract(r"^(\d{3,4})")[0]
    prefijo2 = codigo_tienda.str[:2]

    excluir_50xx = prefijo2.isin(PREFIJO_TIENDAS_EXCLUIR)
    excluir_ripcurl = (~incluir_ripcurl) & prefijo2.isin(PREFIJO_TIENDAS_RIPCURL)
    excluir_multimarca = (~incluir_multimarca) & codigo_tienda.isin(CODIGOS_MULTIMARCA)

    df = df[~(excluir_50xx | excluir_ripcurl | excluir_multimarca)]

    # Regla 4.1 — vigencia de temporada (solo Tienda Propia).
    df["_vigente"] = df.apply(lambda r: _temporada_vigente(r.get("Temporada"), r.get("Año Producto")), axis=1)
    df = df[df["_vigente"]].drop(columns="_vigente")

    return df.reset_index(drop=True)


def calcular_tabla_riesgo(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d["Cobertura Sem"] = np.where(
        d["Vta # Sem -1"] > 0, d["Stock # MOM"] / d["Vta # Sem -1"].replace(0, np.nan), np.nan
    )
    en_riesgo = (d["Vta # Sem -1"] > 0) & (d["Cobertura Sem"] < UMBRAL_COBERTURA_RIESGO)
    riesgo = d[en_riesgo].copy()

    riesgo["Bodega Disponible"] = riesgo.get(COL_BODEGA_DISPONIBLE, 0).fillna(0) > 0
    riesgo["Motivo"] = f"Cobertura < {UMBRAL_COBERTURA_RIESGO:.0f} semanas con venta activa (Sem -1 > 0)"
    riesgo["Recomendación de carga"] = np.where(
        riesgo["Bodega Disponible"],
        "Evaluar carga de Master(s)/Pack(s) completo(s) — hay stock en bodega/CD",
        "Sin stock en bodega/CD — escalar a Compras/Logística",
    )

    # "Prioridad" NO es una regla de negocio confirmada por Francisco todavía
    # (queda como pendiente #3 del documento del 24-08). Este es un criterio
    # operativo de orden propuesto por defecto: primero lo accionable de
    # inmediato (con bodega), y dentro de cada grupo, lo de menor cobertura
    # primero (más urgente). Ajustar si Francisco define otro criterio.
    riesgo["Prioridad"] = np.where(riesgo["Bodega Disponible"], "Alta — accionable de inmediato", "Media — requiere escalar")
    riesgo = riesgo.sort_values(
        by=["Bodega Disponible", "Cobertura Sem"], ascending=[False, True]
    ).reset_index(drop=True)

    columnas_salida = [
        "Marca", "Familia", "Sub_Familia", "Tienda", "Desc Pack",
        "Vta # Sem -1", "Stock # MOM", "Bodega Disponible", "Cobertura Sem",
        "Motivo", "Recomendación de carga", "Prioridad",
        "Semanas Stock Tienda",
    ]
    columnas_salida = [c for c in columnas_salida if c in riesgo.columns]
    riesgo = riesgo[columnas_salida].rename(
        columns={"Vta # Sem -1": "Venta Sem -1", "Stock # MOM": "Stock Tienda"}
    )
    return riesgo


