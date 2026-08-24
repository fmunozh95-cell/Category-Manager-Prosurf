"""
Tabla de Carga Semanal — Tienda Propia (Volcom / Rusty)
=========================================================

Interfaz Streamlit. Toda la lógica de negocio (parsing del Excel, limpieza,
reglas de vigencia/riesgo/bodega) vive en `logic.py` para poder testearla
sin levantar la app (ver `test_sintetico.py`). Este archivo solo arma la UI.

Deploy en Streamlit Community Cloud: subir esta carpeta a un repo de GitHub
(`app.py`, `logic.py`, `requirements.txt`) y apuntar el deploy a `app.py`.
"""

import pandas as pd
import streamlit as st

from logic import (
    MARCAS_POR_DEFECTO,
    FAMILIAS_EXCLUIR,
    UMBRAL_COBERTURA_RIESGO,
    leer_hojas_disponibles,
    detectar_fila_encabezado,
    cargar_y_estructurar,
    construir_formato_largo,
    limpiar_y_filtrar,
    calcular_tabla_riesgo,
)

st.set_page_config(page_title="Tabla de Carga Semanal — Prosurf", layout="wide", page_icon="📦")

st.title("📦 Tabla de Carga Semanal — Tienda Propia (Volcom / Rusty)")
st.caption(
    "Entregable fijo de los lunes (ítem 11 del proyecto + reglas confirmadas el 24-08-2026). "
    "Reemplaza la ejecución celda a celda en Colab."
)

with st.expander("ℹ️ Qué hace este script y qué asume", expanded=False):
    st.markdown(
        f"""
- Lee el archivo **`4 Semanas - TP (Rip Curl & Prosurf) [DD.MM.AAAA].xlsx`**, hoja
  *"Resumen 4 Semanas - Pro + Rip"*.
- Detecta automáticamente la fila de encabezado (busca la fila con `Marca` y `Familia`).
- Las columnas 0-137 se leen **por posición**, según el orden confirmado en el
  Diccionario de Datos M300 (Parte B). Las columnas 138 en adelante (28 tiendas
  × 5 métricas) se arman detectando el patrón `NNNN-NOMBRE TIENDA` en las filas
  de encabezado.
- Aplica en orden: filtro de marca → exclusión de {", ".join(FAMILIAS_EXCLUIR)} →
  exclusión de tiendas 50xx / Rip Curl 30xx / multimarca → filtro de vigencia
  de temporada (regla 4.1, solo Tienda Propia) → cálculo de cobertura y riesgo
  de ruptura (cobertura < {UMBRAL_COBERTURA_RIESGO:.0f} semanas con venta activa,
  regla 4.3) → clasificación por disponibilidad de bodega (regla 4.4).
- **Primera vez con un archivo real:** revisa el panel "Diagnóstico de
  estructura" antes de confiar en los resultados — si el archivo cambió de
  formato, la detección automática puede fallar y hay que ajustarla.
- La columna **"Prioridad"** de la tabla final es un criterio operativo de
  orden que propuse yo (accionable primero, luego por cobertura ascendente),
  no una regla que Francisco haya confirmado — queda pendiente definirla con
  él (ver pendiente #3 del documento del 24-08).
        """
    )

with st.sidebar:
    st.header("1. Archivo de entrada")
    archivo = st.file_uploader(
        "4 Semanas - TP (Rip Curl & Prosurf) [DD.MM.AAAA].xlsx", type=["xlsx"]
    )

    st.header("2. Overrides manuales (opcional)")
    forzar_fila = st.number_input(
        "Forzar fila de encabezado (1 = primera fila del Excel). Dejar en 0 para autodetectar.",
        min_value=0, value=0, step=1,
    )

if archivo is None:
    st.info("Sube el archivo Excel en la barra lateral para comenzar.")
    st.stop()

archivo_bytes = archivo.getvalue()

try:
    hojas = leer_hojas_disponibles(archivo_bytes)
except Exception as e:
    st.error(f"No se pudo abrir el archivo como Excel: {e}")
    st.stop()

hoja_default = next((h for h in hojas if "resumen 4 semanas" in h.lower()), hojas[0])
hoja_sel = st.selectbox("Hoja a usar", options=hojas, index=hojas.index(hoja_default))

try:
    fila_encabezado = (forzar_fila - 1) if forzar_fila > 0 else detectar_fila_encabezado(archivo_bytes, hoja_sel)
except Exception as e:
    st.error(str(e))
    st.stop()

try:
    datos_ancho, diagnostico = cargar_y_estructurar(archivo_bytes, hoja_sel, fila_encabezado)
except Exception as e:
    st.error(f"Error al estructurar el archivo: {e}")
    st.stop()

# Lista de marcas armada con los valores REALES detectados en el archivo (no
# texto libre), para evitar diferencias de tipeo/mayúsculas y, de paso, para
# que si la columna "Marca" viene mal alineada se note de inmediato acá
# (aparecerían valores raros en vez de nombres de marca).
if "Marca" in datos_ancho.columns:
    marcas_disponibles = sorted(datos_ancho["Marca"].dropna().astype(str).str.strip().unique())
else:
    marcas_disponibles = []

marcas_default = [
    m for m in marcas_disponibles
    if m.strip().upper() in {d.upper() for d in MARCAS_POR_DEFECTO}
] or marcas_disponibles

with st.sidebar:
    st.header("3. Alcance del análisis")
    if marcas_disponibles:
        marcas_sel = st.multiselect(
            "Marcas a incluir (detectadas en el archivo)",
            options=marcas_disponibles,
            default=marcas_default,
        )
    else:
        st.error("No se encontró la columna 'Marca' en el archivo leído.")
        marcas_sel = []

    incluir_ripcurl = st.checkbox("Incluir tiendas Rip Curl (30xx)", value=False)
    incluir_multimarca = st.checkbox("Incluir tiendas multimarca (2022 / 4006)", value=False)

with st.expander("🔍 Diagnóstico de estructura detectada", expanded=False):
    col1, col2, col3 = st.columns(3)
    col1.metric("Fila de encabezado usada", diagnostico["fila_encabezado_usada"] + 1)
    col2.metric("Columnas leídas", diagnostico["columnas_totales_crudo"])
    col3.metric("Tiendas detectadas", diagnostico["n_tiendas_detectadas"])
    if diagnostico["columnas_totales_crudo"] != diagnostico["columnas_esperadas_dictamen"]:
        st.warning(
            f"El Diccionario de Datos documentó 278 columnas efectivas; este archivo tiene "
            f"{diagnostico['columnas_totales_crudo']}. La estructura pudo haber cambiado — "
            "revisar antes de confiar en el resultado."
        )
    if diagnostico["nombres_duplicados"]:
        st.error(
            "Se detectaron nombres repetidos al identificar bloques de tienda (probablemente "
            "el mismo código/nombre de tienda apareció en más de un bloque de 5 columnas). "
            "Para no mezclar datos de dos tiendas bajo el mismo nombre, esos bloques quedaron "
            "renombrados con un sufijo `[dup2]`, `[dup3]`, etc. y **se excluyen del análisis** "
            "hasta que se revisen manualmente. Nombres que chocaron:\n\n"
            + "\n".join(f"- {n}" for n in diagnostico["nombres_duplicados"])
        )
    st.write("Tiendas detectadas:")
    st.dataframe(pd.DataFrame({"Tienda": diagnostico["tiendas_detectadas"]}), use_container_width=True, hide_index=True)

    st.write(
        "**Valores reales encontrados en columnas dimensión clave** — esto confirma si las "
        "columnas fijas (0-137) calzaron con el contenido esperado. Si 'Marca' no muestra "
        "nombres de marca reconocibles acá, el problema es de alineación de columnas."
    )
    cols_preview = st.columns(3)
    for i, (col, valores) in enumerate(diagnostico["valores_columnas_clave"].items()):
        with cols_preview[i % 3]:
            st.markdown(f"**{col}**")
            if valores is None:
                st.caption("⚠️ Esta columna no existe con este nombre en el archivo leído.")
            else:
                st.dataframe(
                    pd.DataFrame(valores, columns=["Valor", "Filas"]),
                    use_container_width=True, hide_index=True, height=200,
                )

    st.write("Vista previa (primeras filas, formato ancho):")
    st.dataframe(datos_ancho.head(5), use_container_width=True)

largo = construir_formato_largo(datos_ancho)
if largo.empty:
    st.error(
        "No se pudo construir el formato largo (SKU x Tienda) — probablemente no se "
        "detectó ninguna tienda. Revisa el diagnóstico de estructura arriba."
    )
    st.stop()

df_filtrado, embudo = limpiar_y_filtrar(
    largo, marcas=marcas_sel, excluir_familias=FAMILIAS_EXCLUIR,
    incluir_ripcurl=incluir_ripcurl, incluir_multimarca=incluir_multimarca,
)

with st.expander("🔍 Embudo de filtrado (dónde se pierden filas)", expanded=df_filtrado.empty):
    st.dataframe(
        pd.DataFrame({"Etapa": list(embudo.keys()), "Filas": list(embudo.values())}),
        use_container_width=True, hide_index=True,
    )
    if df_filtrado.empty:
        st.error(
            "No quedó ninguna fila después de los filtros. Revisa la etapa donde las filas "
            "llegan a 0 — lo más común es que el valor real de 'Marca' no calce con las marcas "
            "seleccionadas arriba (revisa mayúsculas/espacios), o que la vigencia de temporada "
            "esté descartando todo porque 'Año Producto' no viene en el formato esperado."
        )

if df_filtrado.empty:
    st.stop()

tabla_riesgo = calcular_tabla_riesgo(df_filtrado)

st.header("Resultado")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Filas SKU × Tienda (vigentes)", f"{len(df_filtrado):,}".replace(",", "."))
c2.metric("Combinaciones en riesgo", f"{len(tabla_riesgo):,}".replace(",", "."))
c3.metric("Con bodega disponible", int(tabla_riesgo["Bodega Disponible"].sum()) if len(tabla_riesgo) else 0)
c4.metric("Sin bodega disponible", int((~tabla_riesgo["Bodega Disponible"]).sum()) if len(tabla_riesgo) else 0)

st.subheader("Tabla de carga — combinaciones SKU/Pack + Tienda en riesgo")
if tabla_riesgo.empty:
    st.success("No se detectaron combinaciones en riesgo con los filtros actuales.")
else:
    familias_disponibles = sorted(tabla_riesgo["Familia"].dropna().unique())
    familias_filtro = st.multiselect("Filtrar por Familia", options=familias_disponibles, default=[])
    tabla_mostrar = tabla_riesgo[tabla_riesgo["Familia"].isin(familias_filtro)] if familias_filtro else tabla_riesgo
    st.dataframe(tabla_mostrar, use_container_width=True, hide_index=True)

    csv = tabla_riesgo.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Descargar tabla de carga (CSV)",
        data=csv,
        file_name="tabla_carga_semanal.csv",
        mime="text/csv",
    )

    st.subheader("Familias más afectadas")
    st.bar_chart(tabla_riesgo["Familia"].value_counts())

    st.subheader("Tiendas más afectadas")
    st.bar_chart(tabla_riesgo["Tienda"].value_counts().head(15))

st.divider()
st.caption(
    "Recuerda: esta tabla es Nivel C (analizar + detectar + explicar + recomendar). "
    "La decisión final de carga la toma Francisco; las combinaciones 'sin bodega disponible' "
    "requieren escalar a Compras/Logística antes de poder actuar."
)
