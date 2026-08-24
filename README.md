# Tabla de Carga Semanal — Streamlit

Primera versión (estructura base) del script que reemplaza el notebook de
Colab para generar el entregable fijo de los lunes: la tabla de tiendas
críticas de carga (SKU/Pack + Tienda en riesgo de ruptura) para Volcom y
Rusty en Tienda Propia.

## Archivos

- `app.py` — interfaz Streamlit (sube el archivo, muestra diagnóstico, filtros y la tabla final).
- `logic.py` — toda la lógica de negocio (parsing del Excel, limpieza, reglas de vigencia/riesgo/bodega). Sin dependencias de la UI, así que se puede testear por separado.
- `test_sintetico.py` — prueba de humo con un archivo Excel sintético que imita la estructura real (138 columnas fijas + bloques de tienda). Verifica que el pipeline completo corre sin errores y que las reglas de negocio se aplican correctamente. **No reemplaza probar con el archivo real.**
- `requirements.txt` — dependencias para desplegar en Streamlit Community Cloud.

## Por qué Streamlit Cloud y no Streamlit local

Francisco no tiene permisos de administrador en su PC corporativo, así que
no puede instalar Python/Streamlit localmente. La idea es desplegar esta
carpeta como app en **Streamlit Community Cloud** (gratis, no requiere
instalación local, solo un repo de GitHub) para dejar de depender de correr
celdas en Colab cada lunes.

## Cómo probarlo localmente (para quien tenga Python)

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Cómo desplegarlo en Streamlit Community Cloud

1. Crear un repositorio de GitHub (puede ser privado) con estos 3 archivos: `app.py`, `logic.py`, `requirements.txt`.
2. Entrar a share.streamlit.io con la cuenta de GitHub.
3. "New app" → seleccionar el repo → archivo principal `app.py` → Deploy.
4. Queda una URL fija que se puede abrir cada lunes desde cualquier navegador, sin instalar nada.

## Qué hace (resumen)

1. Detecta automáticamente la fila de encabezado del Excel (busca la fila con `Marca` y `Familia`).
2. Lee las columnas 0-137 (dimensiones del producto, ventas, stock CD, etc.) por posición, según el orden confirmado en el Diccionario de Datos M300 (Parte B).
3. Detecta las 28 tiendas (columnas 138 en adelante) buscando el patrón `NNNN-NOMBRE TIENDA` en las filas de encabezado, y arma un formato largo (una fila por SKU/Pack × Tienda).
4. Aplica las reglas confirmadas el 24-08-2026:
   - Filtro de marca (Volcom/Rusty por defecto).
   - Exclusión permanente de la familia "Bolsa de Papel".
   - Exclusión de tiendas 50xx, Rip Curl (30xx) y multimarca (2022/4006) — configurable desde la barra lateral.
   - Vigencia de temporada (solo aplica a Tienda Propia): Toda Temporada siempre vigente; Invierno solo Año 2026; Verano Año 2026 o 2027; Colegial y el resto quedan fuera.
   - Negativos en stock se tratan como 0.
   - Cobertura semanal = Stock actual tienda / Venta Sem -1. Riesgo si Venta Sem -1 > 0 y Cobertura < 2 semanas.
   - Bodega disponible = columna `Stock Disponible CD - Pack` (el total agregado, no el desglose por Prosurf, por el bug de la fusión societaria Prosurf→Maui).
5. Muestra la tabla final con las columnas: Marca, Familia, Sub_Familia, Tienda, Desc Pack, Venta Sem -1, Stock Tienda, Bodega Disponible, Cobertura Sem, Motivo, Recomendación de carga, Prioridad, Semanas Stock Tienda — descargable en CSV.

## Estado del deploy

**Deploy en Streamlit Community Cloud completado y funcionando (24-08-2026).** El error inicial de acceso ("You do not have access to this app or it does not exist") se debía a permisos de la GitHub App de Streamlit sobre el repo — quedó resuelto.

Bugs encontrados y corregidos el mismo día ya probando contra el archivo real:
1. `ValueError: Duplicate column names found` — dos bloques de tienda detectados con el mismo nombre. Se agregó deduplicación automática + aviso visible en el diagnóstico.
2. `KeyError: '_vigente' not found in axis` — si el filtro de marca/familia/tienda dejaba el DataFrame en 0 filas, un bug de pandas al usar `.apply(axis=1)` sobre un DataFrame vacío hacía desaparecer todas las columnas. Se reescribió el filtro de vigencia de temporada de forma vectorizada (sin `.apply` fila por fila) y se agregó un panel "Embudo de filtrado" que muestra cuántas filas sobreviven en cada etapa, para detectar de inmediato si algún filtro se está comiendo todo el archivo.
3. **El embudo mostró 0 filas justo en el filtro de marca**, con 608.496 filas iniciales (= 21.732 SKU × 28 tiendas, calza exacto con lo ya validado en Colab). Se cambió el campo de texto libre por un **multiselect que se arma con los valores reales detectados en la columna "Marca"** y se agregó un panel que muestra los valores reales de columnas dimensión clave, lo que permitió ver el problema real: la columna "Marca" tenía códigos de estilo en vez de nombres de marca.
4. **Causa raíz encontrada:** el archivo real trae **1 columna extra al inicio** (vacía en las filas de muestra, pero no en el 100% del archivo — por eso no se elimina sola con el `dropna` de columnas vacías) antes de "Id Estilo", que corría TODAS las columnas fijas siguientes en una posición. Se agregó una **autocalibración**: el script prueba correr el bloque de columnas fijas por varios offsets (0, +1, +2, +3, -1) y elige el que hace que la columna "Marca" contenga más nombres de marca reconocibles (Volcom, Rusty, Rip Curl, Globe, Vans, Dragon, Electric, Creatures) en una muestra de 500 filas. El offset elegido queda visible en el panel de diagnóstico ("Columnas extra al inicio recortadas"), así se nota de inmediato si vuelve a pasar en una semana futura.

## Lo que falta validar con el archivo real (próxima sesión)

- **Esta versión no fue probada contra el archivo real** dentro de esta sesión (no había un `.xlsx` adjunto). Se probó con un archivo Excel sintético que imita la estructura documentada — el pipeline corre de punta a punta y aplica todas las reglas correctamente sobre datos ficticios, pero hay que confirmar con Francisco que:
  - La detección automática de tiendas (patrón `NNNN-NOMBRE`) efectivamente encuentra las 28 tiendas reales y no se cruza con otro dato.
  - El orden de columnas 0-137 sigue siendo el mismo que documentó el Diccionario de Datos (recordar el bug ya visto una vez de columnas con espacios en blanco al inicio del nombre — acá no debería afectar porque se usa posición, no nombre, pero vale la pena verificar el conteo total de columnas).
- La columna **"Prioridad"** es un criterio de orden que propuse yo (accionable primero — con bodega —, luego por cobertura ascendente dentro de cada grupo), **no es una regla que Francisco haya confirmado todavía** (ver pendiente #3 del documento de reglas del 24-08: "Definir si el umbral de 2 semanas de cobertura debe variar por familia").
- Falta extender la misma lógica a **4 Semanas Outlet** (sin el filtro de vigencia de temporada, que no aplica ahí).
- Falta decidir si el script debe conectarse directo a Google Drive (como en Colab) o si basta con que Francisco suba el archivo manualmente cada lunes vía el uploader de Streamlit — la versión actual usa uploader manual porque es más simple y evita el bug de caché de Drive documentado en la sesión del 24-08.
