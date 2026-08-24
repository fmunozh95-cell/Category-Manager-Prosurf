"""
Prueba de humo del pipeline con un archivo sintético que imita la estructura
documentada (138 columnas fijas + N tiendas x 5 métricas), para validar que
la lógica de parsing/limpieza/riesgo funciona de punta a punta ANTES de
tener el archivo real. No reemplaza probar con datos reales.
"""
import os
import sys
import numpy as np
import pandas as pd

sys.path.insert(0, ".")
from logic import (
    COLUMNAS_FIJAS_0_137, METRICAS_POR_TIENDA, COL_BODEGA_DISPONIBLE,
    detectar_fila_encabezado, cargar_y_estructurar, construir_formato_largo,
    limpiar_y_filtrar, calcular_tabla_riesgo, MARCAS_POR_DEFECTO, FAMILIAS_EXCLUIR,
)

TIENDAS = [
    "2201-INSIDE BUENAVENTURA",
    "2206-VOLCOM RIP CURL ALTO LAS CONDES",
    "2208-VOLCOM RIP CURL COSTANERA",
    "3002-RIP CURL FLORIDA C",   # Rip Curl -> debe excluirse por defecto
    "4006-MAUI VOLRIP LOS DOMINICOS",  # multimarca -> debe excluirse por defecto
    "5001-TIENDA FUERA DE ALCANCE",     # 50xx -> debe excluirse siempre
]

rng = np.random.default_rng(42)


def fila_producto(id_estilo, familia, temporada, anio, marca="Volcom"):
    fija = {c: None for c in COLUMNAS_FIJAS_0_137}
    fija["Id Estilo"] = id_estilo
    fija["Desc Estilo"] = f"Estilo {id_estilo}"
    fija["Color"] = "Black"
    fija["Id Pack"] = id_estilo + 1
    fija["Desc Pack"] = f"Pack {familia} {id_estilo}"
    fija["Marca"] = marca
    fija["Linea"] = "Vestuario"
    fija["Departamento"] = "Vestuario Hombre"
    fija["Familia"] = familia
    fija["Sub_Familia"] = "MC"
    fija["Temporada"] = temporada
    fija["Año Producto"] = anio
    fija["Definicion"] = "ADULTO"
    fija["Genero"] = "MAN"
    fija["Unid x Inner"] = 3
    fija["Inner x Master"] = 2
    fija["Unid x Master"] = 6
    fija[COL_BODEGA_DISPONIBLE] = rng.integers(0, 5)
    fija["Semanas Stock Tienda"] = rng.integers(1, 30)
    for c in COLUMNAS_FIJAS_0_137:
        if fija[c] is None:
            fija[c] = 0
    return fija


productos = [
    fila_producto(1000, "Poleron", "INVIERNO", 2026),   # vigente
    fila_producto(1001, "Polera", "VERANO", 2026),      # vigente
    fila_producto(1002, "Casaca", "VERANO", 2020),      # NO vigente (año viejo)
    fila_producto(1003, "Bolsa de Papel", "TODA TEMPORADA", 2026),  # excluir por familia
    fila_producto(1004, "Parka", "COLEGIAL", 2026),      # excluir (colegial)
    fila_producto(1005, "Polera", "INVIERNO", 2026, marca="Rip Curl"),  # excluir por marca
]

filas_datos = []
for prod in productos:
    fila = dict(prod)
    for tienda in TIENDAS:
        vta_sem1 = int(rng.integers(0, 10))
        stock = int(rng.integers(0, 15))
        fila[f"{tienda} || Vta # 4 Sem"] = vta_sem1 * 3
        fila[f"{tienda} || Vta # Sem -1"] = vta_sem1
        fila[f"{tienda} || Stock # MOM"] = stock
        fila[f"{tienda} || Pend. Envío"] = 0
        fila[f"{tienda} || Pend. Recib"] = 0
    filas_datos.append(fila)

# Forzar un caso de riesgo evidente: venta > 0 y stock bajo => cobertura < 2
filas_datos[0]["2208-VOLCOM RIP CURL COSTANERA || Vta # Sem -1"] = 8
filas_datos[0]["2208-VOLCOM RIP CURL COSTANERA || Stock # MOM"] = 3  # cobertura 0.375 -> riesgo
filas_datos[0][COL_BODEGA_DISPONIBLE] = 10  # con bodega disponible

nombres_tienda_columnas = []
for tienda in TIENDAS:
    for m in METRICAS_POR_TIENDA:
        nombres_tienda_columnas.append(f"{tienda} || {m}")

todas_columnas = list(COLUMNAS_FIJAS_0_137) + nombres_tienda_columnas
df_final = pd.DataFrame(filas_datos)[todas_columnas]

# --- Armar el excel crudo como lo entrega el sistema real: filas vacías/meta
# arriba, encabezado con nombre de tienda + metrica en distintas filas ---
fila_vacia = [None] * len(todas_columnas)

fila_nombre_tienda = [None] * len(COLUMNAS_FIJAS_0_137)
for tienda in TIENDAS:
    fila_nombre_tienda += [tienda, None, None, None, None]

fila_totales = [None] * len(todas_columnas)

fila_encabezado_final = list(COLUMNAS_FIJAS_0_137)
for _ in TIENDAS:
    fila_encabezado_final += METRICAS_POR_TIENDA

crudo = pd.DataFrame(
    [fila_vacia, fila_nombre_tienda, fila_totales, fila_encabezado_final],
    columns=range(len(todas_columnas)),
)
datos_sin_header = df_final.copy()
datos_sin_header.columns = range(len(todas_columnas))
crudo = pd.concat([crudo, datos_sin_header], ignore_index=True)

archivo_test = "test_4semtp_sintetico.xlsx"
with pd.ExcelWriter(archivo_test, engine="openpyxl") as writer:
    crudo.to_excel(writer, sheet_name="Resumen 4 Semanas - Pro + Rip", header=False, index=False)

with open(archivo_test, "rb") as f:
    archivo_bytes = f.read()

print("=== Test 1: detección de fila de encabezado ===")
fila_enc = detectar_fila_encabezado(archivo_bytes, "Resumen 4 Semanas - Pro + Rip")
print("Fila detectada (0-idx):", fila_enc, "-> esperado 3")
assert fila_enc == 3

print("\n=== Test 2: estructuración ancha ===")
datos_ancho, diag = cargar_y_estructurar(archivo_bytes, "Resumen 4 Semanas - Pro + Rip", fila_enc)
print(diag)
assert diag["n_tiendas_detectadas"] == len(TIENDAS), diag

print("\n=== Test 3: formato largo ===")
largo = construir_formato_largo(datos_ancho)
print("Filas formato largo:", len(largo), "(esperado", len(productos) * len(TIENDAS), ")")
assert len(largo) == len(productos) * len(TIENDAS)

print("\n=== Test 4: limpieza y filtros de negocio ===")
filtrado, embudo = limpiar_y_filtrar(largo, marcas=MARCAS_POR_DEFECTO, excluir_familias=FAMILIAS_EXCLUIR)
print("Embudo:", embudo)
print("Filas tras filtros:", len(filtrado))
print("Marcas presentes:", filtrado["Marca"].unique())
print("Familias presentes:", sorted(filtrado["Familia"].unique()))
print("Tiendas presentes:", sorted(filtrado["Tienda"].unique()))
assert "Rip Curl" not in filtrado["Marca"].unique()
assert "Bolsa de Papel" not in filtrado["Familia"].unique()
assert not any(t.startswith("3002") for t in filtrado["Tienda"].unique())
assert not any(t.startswith("4006") for t in filtrado["Tienda"].unique())
assert not any(t.startswith("5001") for t in filtrado["Tienda"].unique())
assert "Casaca" not in filtrado["Familia"].unique()  # año viejo -> no vigente
assert "Parka" not in filtrado["Familia"].unique()   # colegial -> excluido

print("\n=== Test 4b: filtro que deja 0 filas no debe romper columnas (bug real visto en producción) ===")
filtrado_vacio, embudo_vacio = limpiar_y_filtrar(largo, marcas=["Marca Que No Existe"], excluir_familias=FAMILIAS_EXCLUIR)
print("Embudo (marca inexistente):", embudo_vacio)
assert filtrado_vacio.empty
assert list(filtrado_vacio.columns) == list(largo.columns), (
    "Al quedar en 0 filas, el DataFrame perdió columnas — es el bug de pandas "
    ".apply(axis=1) sobre DataFrame vacío que ya se corrigió una vez; no debería repetirse."
)

print("\n=== Test 5: tabla de riesgo ===")
riesgo = calcular_tabla_riesgo(filtrado)
print(riesgo)
caso_forzado = riesgo[(riesgo["Tienda"].str.startswith("2208")) & (riesgo["Familia"] == "Poleron")]
assert len(caso_forzado) == 1, "El caso de riesgo forzado (Costanera, Poleron) no aparece en la tabla"
assert bool(caso_forzado.iloc[0]["Bodega Disponible"]) is True

print("\n=== Test 6: columna extra al inicio (bug real reportado por Francisco) ===")
# Reproduce exactamente lo que Francisco vio: una columna vacía antes de
# "Id Estilo" que corre todo el resto de las columnas fijas en 1 posición.
crudo_con_offset = crudo.copy()
valores_extra = [None] * len(crudo_con_offset)
valores_extra[1] = "grupo"  # 1 valor no nulo en algún lado, para que dropna(how="all") NO la elimine sola —
# así el test realmente ejercita la detección de offset, no el dropna (en el archivo real la columna
# extra tampoco era 100% NaN en todas las filas, por eso el diagnóstico la contó como columna #279).
crudo_con_offset.insert(0, "extra", valores_extra)
crudo_con_offset.columns = range(crudo_con_offset.shape[1])

archivo_test_offset = "test_4semtp_sintetico_con_offset.xlsx"
with pd.ExcelWriter(archivo_test_offset, engine="openpyxl") as writer:
    crudo_con_offset.to_excel(writer, sheet_name="Resumen 4 Semanas - Pro + Rip", header=False, index=False)

with open(archivo_test_offset, "rb") as f:
    archivo_bytes_offset = f.read()

fila_enc_offset = detectar_fila_encabezado(archivo_bytes_offset, "Resumen 4 Semanas - Pro + Rip")
datos_ancho_offset, diag_offset = cargar_y_estructurar(archivo_bytes_offset, "Resumen 4 Semanas - Pro + Rip", fila_enc_offset)
print("Offset detectado:", diag_offset["offset_columnas_aplicado"], "(esperado 1)")
print("Coincidencias por offset probado:", diag_offset["offset_diagnostico"])
assert diag_offset["offset_columnas_aplicado"] == 1
assert set(datos_ancho_offset["Marca"].unique()) == {"Volcom", "Rip Curl"}, (
    f"Marca no quedó bien alineada tras corregir el offset: {datos_ancho_offset['Marca'].unique()}"
)
os.remove(archivo_test_offset)

print("\n✅ TODOS LOS TESTS PASARON")
