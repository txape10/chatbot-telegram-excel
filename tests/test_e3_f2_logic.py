"""Tests E3/F2 — lógica pura de preview, valores únicos y exportar CSV.

Extrae y verifica la lógica de datos de las funciones del handler,
sin necesitar Telegram ni LLM.
"""
import io
import re
import pytest
import pandas as pd


# ── Helpers que replican la lógica pura de los handlers ──────────────────────

def _extraer_preview_params(pregunta: str):
    """Extrae ultimas (bool) y n (int) de la pregunta, igual que _previsualizar."""
    ultimas = bool(re.search(r"\b[uú]ltimas?\b", pregunta, re.IGNORECASE))
    m = re.search(r"\b(\d+)\b", pregunta)
    n = int(m.group(1)) if m else 10
    n = min(n, 30)
    return ultimas, n


def _aplicar_preview(df: pd.DataFrame, pregunta: str):
    """Aplica la lógica de preview al DataFrame."""
    ultimas, n = _extraer_preview_params(pregunta)
    return df.tail(n) if ultimas else df.head(n)


def _encontrar_columna(df: pd.DataFrame, pregunta: str):
    """Replica la búsqueda de columna por nombre en la pregunta (_valores_unicos)."""
    pregunta_lower = pregunta.lower()
    for col in df.columns:
        if col.lower() in pregunta_lower:
            return col
    return None


def _crear_csv_bytes(df: pd.DataFrame) -> bytes:
    """Replica la lógica de _exportar_csv: UTF-8 con BOM."""
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding="utf-8-sig")
    buf.seek(0)
    return buf.read()


# ── Fixtures ─────────────────────────────────────────────────────────────────

@pytest.fixture
def df_grande():
    return pd.DataFrame({
        "ID":       range(1, 51),
        "Producto": [f"Prod_{i}" for i in range(1, 51)],
        "Ventas":   [i * 10.0 for i in range(1, 51)],
        "Región":   (["Norte", "Sur", "Este", "Oeste", "Centro"] * 10),
    })


@pytest.fixture
def df_categorias():
    return pd.DataFrame({
        "Estado":    ["Activo", "Inactivo", "Pendiente", "Activo", "Inactivo"],
        "Región":    ["Norte", "Sur", "Norte", "Este", "Sur"],
        "Importe":   [100, 200, 150, 300, 250],
    })


# ══════════════════════════════════════════════════════════════════════════════
# F2 — Preview: extracción de parámetros
# ══════════════════════════════════════════════════════════════════════════════

class TestPreviewParams:
    def test_primeras_10_por_defecto(self):
        ultimas, n = _extraer_preview_params("muéstrame los datos")
        assert not ultimas
        assert n == 10

    def test_primeras_15(self):
        ultimas, n = _extraer_preview_params("muéstrame las primeras 15 filas")
        assert not ultimas
        assert n == 15

    def test_ultimas_5(self):
        ultimas, n = _extraer_preview_params("dame las últimas 5 filas")
        assert ultimas
        assert n == 5

    def test_limite_30(self):
        """n se limita a 30 aunque se pida más."""
        ultimas, n = _extraer_preview_params("dame las primeras 100 filas")
        assert n == 30

    def test_ultimas_sin_numero(self):
        ultimas, n = _extraer_preview_params("muéstrame las últimas filas")
        assert ultimas
        assert n == 10  # default


class TestPreviewAplicado:
    def test_primeras_5(self, df_grande):
        resultado = _aplicar_preview(df_grande, "primeras 5 filas")
        assert len(resultado) == 5
        assert list(resultado["ID"]) == [1, 2, 3, 4, 5]

    def test_ultimas_3(self, df_grande):
        resultado = _aplicar_preview(df_grande, "últimas 3 filas")
        assert len(resultado) == 3
        assert list(resultado["ID"]) == [48, 49, 50]

    def test_default_10(self, df_grande):
        resultado = _aplicar_preview(df_grande, "ver datos")
        assert len(resultado) == 10

    def test_mas_filas_que_df(self):
        df = pd.DataFrame({"A": [1, 2, 3]})
        resultado = _aplicar_preview(df, "primeras 20 filas")
        # head(20) de un df de 3 filas devuelve las 3
        assert len(resultado) == 3

    def test_limite_30_aplicado(self, df_grande):
        resultado = _aplicar_preview(df_grande, "primeras 50 filas")
        assert len(resultado) == 30


# ══════════════════════════════════════════════════════════════════════════════
# F2 — Valores únicos: detección de columna
# ══════════════════════════════════════════════════════════════════════════════

class TestValoresUnicosLogica:
    def test_columna_detectada_exacta(self, df_categorias):
        col = _encontrar_columna(df_categorias, "valores únicos de Estado")
        assert col == "Estado"

    def test_columna_detectada_case_insensitive(self, df_categorias):
        col = _encontrar_columna(df_categorias, "valores únicos de estado")
        assert col == "Estado"

    def test_columna_no_detectada_devuelve_none(self, df_categorias):
        col = _encontrar_columna(df_categorias, "valores únicos del archivo")
        assert col is None

    def test_primera_columna_encontrada(self, df_categorias):
        """Si dos columnas están en la pregunta, devuelve la primera."""
        col = _encontrar_columna(df_categorias, "Estado y Región")
        assert col in ("Estado", "Región")

    def test_valores_unicos_correctos(self, df_categorias):
        unicos = set(df_categorias["Estado"].dropna().unique())
        assert unicos == {"Activo", "Inactivo", "Pendiente"}

    def test_n_valores_unicos(self, df_categorias):
        n = df_categorias["Región"].nunique()
        assert n == 3  # Norte, Sur, Este (la fixture tiene esos 3 únicos)

    def test_valores_unicos_ordenados(self, df_categorias):
        unicos = sorted(df_categorias["Estado"].dropna().unique(), key=str)
        assert unicos == ["Activo", "Inactivo", "Pendiente"]

    def test_nunique_por_columna(self, df_categorias):
        resumen = {col: df_categorias[col].nunique() for col in df_categorias.columns}
        assert resumen["Estado"] == 3  # Activo, Inactivo, Pendiente
        assert resumen["Región"] == 3  # Norte, Sur, Este


# ══════════════════════════════════════════════════════════════════════════════
# E3 — Exportar CSV
# ══════════════════════════════════════════════════════════════════════════════

class TestExportarCSV:
    def test_csv_tiene_bom_utf8(self, df_categorias):
        """El CSV debe empezar con BOM UTF-8 (0xEF 0xBB 0xBF) para Excel en Windows."""
        contenido = _crear_csv_bytes(df_categorias)
        assert contenido[:3] == b"\xef\xbb\xbf"

    def test_csv_tiene_cabeceras(self, df_categorias):
        contenido = _crear_csv_bytes(df_categorias).decode("utf-8-sig")
        primera_linea = contenido.splitlines()[0]
        assert "Estado" in primera_linea
        assert "Región" in primera_linea
        assert "Importe" in primera_linea

    def test_csv_tiene_datos(self, df_categorias):
        contenido = _crear_csv_bytes(df_categorias).decode("utf-8-sig")
        assert "Activo" in contenido
        assert "Norte" in contenido

    def test_csv_sin_indice(self, df_categorias):
        """No debe incluir el índice de pandas como columna."""
        contenido = _crear_csv_bytes(df_categorias).decode("utf-8-sig")
        primera_linea = contenido.splitlines()[0]
        # El índice añadiría un campo vacío o numérico al inicio
        assert not primera_linea.startswith(",")

    def test_csv_n_filas_correcto(self, df_categorias):
        contenido = _crear_csv_bytes(df_categorias).decode("utf-8-sig")
        lineas = [l for l in contenido.splitlines() if l]
        # 1 cabecera + 5 filas de datos
        assert len(lineas) == 6

    def test_csv_df_numerico(self):
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4.5, 5.5, 6.5]})
        contenido = _crear_csv_bytes(df).decode("utf-8-sig")
        assert "1" in contenido
        assert "4.5" in contenido

    def test_csv_df_vacio(self):
        df = pd.DataFrame({"Col1": [], "Col2": []})
        contenido = _crear_csv_bytes(df).decode("utf-8-sig")
        primera_linea = contenido.strip().splitlines()[0]
        assert "Col1" in primera_linea


# ══════════════════════════════════════════════════════════════════════════════
# llm.py — Funciones puras de utilidad
# ══════════════════════════════════════════════════════════════════════════════

class TestLLMUtils:
    def test_limpiar_json_sin_markdown(self):
        from services.llm import _limpiar_json
        texto = '{"op": "ordenar", "col": "Fecha"}'
        assert _limpiar_json(texto) == texto

    def test_limpiar_json_elimina_bloque_codigo(self):
        from services.llm import _limpiar_json
        texto = '```json\n{"op": "ordenar"}\n```'
        resultado = _limpiar_json(texto)
        assert "```" not in resultado
        assert '{"op": "ordenar"}' in resultado

    def test_limpiar_json_elimina_bloque_sin_lenguaje(self):
        from services.llm import _limpiar_json
        texto = '```\n{"op": "filtrar"}\n```'
        resultado = _limpiar_json(texto)
        assert "```" not in resultado

    def test_limpiar_json_mantiene_contenido(self):
        from services.llm import _limpiar_json
        inner = '{"aclaracion_necesaria": true, "pregunta": "¿qué columna?"}'
        texto = f'```json\n{inner}\n```'
        resultado = _limpiar_json(texto)
        import json
        parsed = json.loads(resultado)
        assert parsed["aclaracion_necesaria"] is True

    def test_estimar_tokens_proporcional(self):
        from services.llm import _estimar_tokens
        corto = _estimar_tokens("Hola")
        largo = _estimar_tokens("Hola " * 100)
        assert largo > corto

    def test_estimar_tokens_cadena_vacia(self):
        from services.llm import _estimar_tokens
        assert _estimar_tokens("") == 0

    def test_estimar_tokens_aprox_4_chars_por_token(self):
        from services.llm import _estimar_tokens
        # "1234" = 4 chars → ~1 token
        assert _estimar_tokens("1234") == 1

    def test_estimar_tokens_texto_largo(self):
        from services.llm import _estimar_tokens
        texto = "a" * 400   # 400 chars → ~100 tokens
        assert _estimar_tokens(texto) == 100
