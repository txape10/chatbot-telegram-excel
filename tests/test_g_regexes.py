"""Tests Sprint G, E3, F1, F2 — detección de intenciones por regex.

Verifica que las expresiones regulares del router de intenciones (messages.py)
detectan correctamente los patrones esperados y no dan falsos positivos.
No requiere Telegram ni LLM.
"""
import pytest
import re
from handlers.intent_patterns import (
    _RE_PREVIEW,
    _RE_VALORES_UNICOS,
    _RE_EXPLICAR_ARCHIVO,
    _RE_EXPORTAR_CSV,
    _RE_UNDO,
    _RE_EDICION,
    _RE_CREAR_EXCEL,
    _RE_COMPARAR,
    _RE_COMBINAR,
    _RE_GRAFICO,
    _RE_STATS,
    _RE_GUARDAR_MACRO,
    _RE_EJECUTAR_MACRO,
    _RE_LISTAR_MACROS,
    _RE_BORRAR_MACRO,
)
from handlers.excel_edit import _OPS_DESTRUCTIVAS


# ── Helpers ───────────────────────────────────────────────────────────────────

def match(pattern, texto):
    return bool(pattern.search(texto))


# ══════════════════════════════════════════════════════════════════════════════
# F2 — Preview de filas
# ══════════════════════════════════════════════════════════════════════════════

class TestPreview:
    def test_primeras_n_filas(self):
        assert match(_RE_PREVIEW, "muéstrame las primeras 10 filas")

    def test_ultimas_n_filas(self):
        assert match(_RE_PREVIEW, "muéstrame las últimas 5 filas")

    def test_primeras_registros(self):
        assert match(_RE_PREVIEW, "dame las primeras 20 registros")

    def test_n_primeras(self):
        assert match(_RE_PREVIEW, "las 15 primeras filas")

    def test_previsualizar(self):
        assert match(_RE_PREVIEW, "previsualiza el archivo")

    def test_ver_datos(self):
        assert match(_RE_PREVIEW, "ver los datos")

    def test_mostrarme_datos(self):
        assert match(_RE_PREVIEW, "muéstrame los datos")

    def test_primeras_sin_numero_no_matchea(self):
        # El regex requiere número o frase específica ("ver datos", "previsualizar")
        # "primeras filas" sin número no activa el preview
        assert not match(_RE_PREVIEW, "primeras filas")

    def test_no_falso_positivo_edicion(self):
        # "ordena" no debería activar preview
        assert not match(_RE_PREVIEW, "ordena por fecha")

    def test_no_falso_positivo_pregunta_general(self):
        assert not match(_RE_PREVIEW, "cómo uso BUSCARV en Excel")


# ══════════════════════════════════════════════════════════════════════════════
# F2 — Valores únicos
# ══════════════════════════════════════════════════════════════════════════════

class TestValoresUnicos:
    def test_valores_unicos(self):
        assert match(_RE_VALORES_UNICOS, "valores únicos de Categoría")

    def test_valores_distintos(self):
        assert match(_RE_VALORES_UNICOS, "qué valores distintos hay en Region")

    def test_que_categorias_hay(self):
        assert match(_RE_VALORES_UNICOS, "qué categorías hay en el archivo")

    def test_listar_opciones(self):
        # El regex usa "los?" — "listar los valores" sí coincide
        assert match(_RE_VALORES_UNICOS, "lista los valores de Estado")

    def test_cuantos_distintos(self):
        assert match(_RE_VALORES_UNICOS, "cuántos distintos hay en Producto")

    def test_cuales_son_los_posibles(self):
        assert match(_RE_VALORES_UNICOS, "cuáles son los posibles valores")

    def test_no_falso_positivo(self):
        assert not match(_RE_VALORES_UNICOS, "ordena por importe")

    def test_que_tipos_existen(self):
        assert match(_RE_VALORES_UNICOS, "qué tipos existen en la columna")


# ══════════════════════════════════════════════════════════════════════════════
# E3 — Explicar archivo
# ══════════════════════════════════════════════════════════════════════════════

class TestExplicarArchivo:
    def test_explicame_este_archivo(self):
        assert match(_RE_EXPLICAR_ARCHIVO, "explícame este archivo")

    def test_que_contiene(self):
        assert match(_RE_EXPLICAR_ARCHIVO, "qué contiene este archivo")

    def test_describeme_los_datos(self):
        # El regex acepta "descríbeme este archivo" (con "este" opcional)
        assert match(_RE_EXPLICAR_ARCHIVO, "descríbeme este archivo")

    def test_resumen_del_archivo(self):
        assert match(_RE_EXPLICAR_ARCHIVO, "resumen del archivo")

    def test_analiza_este_archivo(self):
        assert match(_RE_EXPLICAR_ARCHIVO, "analiza este archivo")

    def test_de_que_va_este_excel(self):
        assert match(_RE_EXPLICAR_ARCHIVO, "de qué va este excel")

    def test_que_hay_en_el_excel(self):
        assert match(_RE_EXPLICAR_ARCHIVO, "qué hay en este excel")

    def test_no_falso_positivo_estadistica(self):
        # "estadísticas" no es "explicame"
        assert not match(_RE_EXPLICAR_ARCHIVO, "dame estadísticas del archivo")

    def test_no_falso_positivo_edicion(self):
        assert not match(_RE_EXPLICAR_ARCHIVO, "ordena el archivo por fecha")


# ══════════════════════════════════════════════════════════════════════════════
# E3 — Exportar CSV
# ══════════════════════════════════════════════════════════════════════════════

class TestExportarCSV:
    def test_exportar_a_csv(self):
        assert match(_RE_EXPORTAR_CSV, "exportar a CSV")

    def test_exportar_como_csv(self):
        assert match(_RE_EXPORTAR_CSV, "exporta como csv")

    def test_guardar_en_csv(self):
        assert match(_RE_EXPORTAR_CSV, "guarda como csv")

    def test_descargar_en_csv(self):
        assert match(_RE_EXPORTAR_CSV, "descarga en csv")

    def test_convertir_a_csv(self):
        assert match(_RE_EXPORTAR_CSV, "convierte a csv")

    def test_en_formato_csv(self):
        assert match(_RE_EXPORTAR_CSV, "en formato csv")

    def test_case_insensitive(self):
        assert match(_RE_EXPORTAR_CSV, "Exportar A CSV")

    def test_no_falso_positivo(self):
        assert not match(_RE_EXPORTAR_CSV, "hazme un Excel")


# ══════════════════════════════════════════════════════════════════════════════
# E2 — Deshacer
# ══════════════════════════════════════════════════════════════════════════════

class TestUndo:
    def test_deshacer(self):
        assert match(_RE_UNDO, "deshacer")

    def test_deshaz(self):
        assert match(_RE_UNDO, "deshaz lo anterior")

    def test_vuelve_atras(self):
        assert match(_RE_UNDO, "vuelve atrás")

    def test_no_falso_positivo(self):
        assert not match(_RE_UNDO, "ordena los datos")


# ══════════════════════════════════════════════════════════════════════════════
# F1 — Operaciones destructivas
# ══════════════════════════════════════════════════════════════════════════════

class TestOpsDestructivas:
    def test_eliminar_columna_es_destructiva(self):
        assert "eliminar_columna" in _OPS_DESTRUCTIVAS

    def test_eliminar_duplicados_es_destructiva(self):
        assert "eliminar_duplicados" in _OPS_DESTRUCTIVAS

    def test_filtrar_exportar_es_destructiva(self):
        assert "filtrar_exportar" in _OPS_DESTRUCTIVAS

    def test_añadir_columna_no_es_destructiva(self):
        assert "añadir_columna" not in _OPS_DESTRUCTIVAS

    def test_ordenar_no_es_destructiva(self):
        assert "ordenar" not in _OPS_DESTRUCTIVAS

    def test_rellenar_nulos_no_es_destructiva(self):
        assert "rellenar_nulos" not in _OPS_DESTRUCTIVAS


# ══════════════════════════════════════════════════════════════════════════════
# Sprint G — Detección de intención de edición (regresión del router)
# ══════════════════════════════════════════════════════════════════════════════

class TestEdicionRouter:
    def test_añadir_columna(self):
        assert match(_RE_EDICION, "añade una columna Margen")

    def test_ordena(self):
        assert match(_RE_EDICION, "ordena por fecha descendente")

    def test_elimina_duplicados(self):
        assert match(_RE_EDICION, "elimina los duplicados")

    def test_rellena_nulos(self):
        assert match(_RE_EDICION, "rellena los vacíos con cero")

    def test_normaliza_texto(self):
        assert match(_RE_EDICION, "normaliza el texto de la columna Nombre")

    def test_estandariza_fechas(self):
        assert match(_RE_EDICION, "estandariza las fechas")

    def test_buscar_reemplazar(self):
        assert match(_RE_EDICION, "reemplaza Madrid por Valencia")

    def test_dividir_columna(self):
        assert match(_RE_EDICION, "divide la columna Nombre por espacio")

    def test_concatenar(self):
        assert match(_RE_EDICION, "concatena Nombre y Apellido")

    def test_pivot(self):
        assert match(_RE_EDICION, "pivotea por Producto")

    def test_unpivot(self):
        assert match(_RE_EDICION, "convierte las columnas en filas")

    def test_no_falso_positivo_pregunta(self):
        assert not match(_RE_EDICION, "cuánto suma la columna Ventas")

    def test_crear_excel_no_activa_edicion(self):
        # La creación de Excel no debe confundirse con edición
        assert not match(_RE_EDICION, "hazme un excel con columnas Fecha e Importe")


# ══════════════════════════════════════════════════════════════════════════════
# Sprint G — Detección de creación, gráfico, stats, comparar, combinar
# ══════════════════════════════════════════════════════════════════════════════

class TestOtrosRouters:
    def test_crear_excel(self):
        assert match(_RE_CREAR_EXCEL, "crea un Excel con columnas Fecha y Concepto")

    def test_hazme_tabla(self):
        assert match(_RE_CREAR_EXCEL, "hazme una tabla de gastos mensuales")

    def test_genera_archivo(self):
        assert match(_RE_CREAR_EXCEL, "genera un archivo con mis ventas")

    def test_comparar(self):
        assert match(_RE_COMPARAR, "compara los dos archivos")

    def test_diferencias(self):
        assert match(_RE_COMPARAR, "qué diferencias hay entre los archivos")

    def test_combinar(self):
        assert match(_RE_COMBINAR, "une los dos archivos por ID")

    def test_merge(self):
        assert match(_RE_COMBINAR, "merge con el otro archivo")

    def test_grafico_barras(self):
        assert match(_RE_GRAFICO, "hazme un gráfico de barras de Ventas")

    def test_grafico_lineas(self):
        assert match(_RE_GRAFICO, "grafico de líneas de Ventas por Mes")

    def test_estadisticas(self):
        assert match(_RE_STATS, "estadísticas del archivo")

    def test_correlaciones(self):
        # El regex usa correlac[ií][oó]n[es]? — coincide con "correlación" (singular)
        assert match(_RE_STATS, "correlación entre columnas")


# ══════════════════════════════════════════════════════════════════════════════
# Sprint G — Macros: detección de nombre en la frase
# ══════════════════════════════════════════════════════════════════════════════

class TestMacrosRegex:
    def test_guardar_macro_captura_nombre(self):
        # El regex espera "una macro" o directamente "macro" (no "esta macro")
        m = _RE_GUARDAR_MACRO.search("guarda una macro llamada LimpiarFechas")
        assert m is not None
        assert m.group(1) == "LimpiarFechas"

    def test_guardar_macro_variante_crea(self):
        m = _RE_GUARDAR_MACRO.search("crea una macro llamada OrdenarVentas")
        assert m is not None
        assert m.group(1) == "OrdenarVentas"

    def test_ejecutar_macro_captura_nombre(self):
        m = _RE_EJECUTAR_MACRO.search("ejecuta la macro LimpiarFechas")
        assert m is not None
        assert m.group(1) == "LimpiarFechas"

    def test_aplicar_macro(self):
        m = _RE_EJECUTAR_MACRO.search("aplica la macro MiFiltro")
        assert m is not None
        assert m.group(1) == "MiFiltro"

    def test_listar_macros(self):
        assert match(_RE_LISTAR_MACROS, "listar macros")

    def test_mis_macros(self):
        assert match(_RE_LISTAR_MACROS, "mis macros")

    def test_que_macros_tengo(self):
        assert match(_RE_LISTAR_MACROS, "qué macros tengo")

    def test_borrar_macro_captura_nombre(self):
        m = _RE_BORRAR_MACRO.search("borra la macro LimpiarFechas")
        assert m is not None
        assert m.group(1) == "LimpiarFechas"

    def test_eliminar_macro(self):
        m = _RE_BORRAR_MACRO.search("elimina la macro MiFiltro")
        assert m is not None
        assert m.group(1) == "MiFiltro"
