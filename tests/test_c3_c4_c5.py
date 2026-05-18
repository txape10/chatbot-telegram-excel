"""Tests para Sprint C3 (tendencia), C4 (normalizar) y C5 (pivot/unpivot)."""
import io
import pytest
import numpy as np
import pandas as pd

from excel.editor import aplicar_edicion, EditorError
from excel.analyzer import analisis_tendencia


# ── Fixtures ─────────────────────────────────────────────────────────────────

@pytest.fixture
def df_ventas_mensuales():
    return pd.DataFrame({
        "Mes":     ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio"],
        "Ventas":  [1000, 1200, 1100, 1350, 1400, 1600],
        "Costes":  [600,  650,  620,  700,  710,  750],
    })


@pytest.fixture
def df_wide():
    """DataFrame en formato ancho: columnas de meses como valores."""
    return pd.DataFrame({
        "Producto":  ["Teclado", "Ratón", "Monitor"],
        "Enero":     [100, 200, 300],
        "Febrero":   [120, 210, 280],
        "Marzo":     [110, 230, 310],
    })


@pytest.fixture
def df_long():
    """DataFrame en formato largo para pivotear."""
    return pd.DataFrame({
        "Producto": ["Teclado", "Teclado", "Ratón", "Ratón"],
        "Mes":      ["Enero", "Febrero", "Enero", "Febrero"],
        "Ventas":   [100, 120, 200, 210],
    })


@pytest.fixture
def df_texto_sucio():
    return pd.DataFrame({
        "Nombre":    ["  Ana  ", "LUIS", "eva", "  María "],
        "Ciudad":    ["MADRID", "barcelona", "  Sevilla  ", "Valencia"],
        "Importe":   [100, 200, 150, 300],
    })


@pytest.fixture
def df_fechas_texto():
    return pd.DataFrame({
        "Fecha":  ["01/03/2024", "15-04-2024", "2024-05-20", "invalid"],
        "Valor":  [10, 20, 30, 40],
    })


# ── C3 — Análisis de tendencia ────────────────────────────────────────────────

def test_tendencia_detecta_crecimiento(df_ventas_mensuales):
    texto, buf = analisis_tendencia(df_ventas_mensuales)
    assert "creciente" in texto.lower()
    assert "Ventas" in texto


def test_tendencia_devuelve_r2(df_ventas_mensuales):
    texto, _ = analisis_tendencia(df_ventas_mensuales)
    assert "R²" in texto or "r²" in texto.lower()


def test_tendencia_genera_imagen(df_ventas_mensuales):
    _, buf = analisis_tendencia(df_ventas_mensuales)
    assert buf is not None
    data = buf.read(4)
    assert data[:4] == b'\x89PNG'


def test_tendencia_sin_columnas_numericas():
    df = pd.DataFrame({"A": ["x", "y", "z"], "B": ["p", "q", "r"]})
    texto, buf = analisis_tendencia(df)
    assert "⚠️" in texto
    assert buf is None


def test_tendencia_datos_insuficientes():
    df = pd.DataFrame({"Ventas": [100, 200]})   # solo 2 filas
    texto, buf = analisis_tendencia(df)
    assert "⚠️" in texto or buf is None  # puede no generar gráfico


def test_tendencia_detecta_decrecimiento():
    df = pd.DataFrame({"Ventas": [500, 400, 350, 300, 250, 200]})
    texto, _ = analisis_tendencia(df)
    assert "decreciente" in texto.lower()


def test_tendencia_usa_columna_fecha():
    df = pd.DataFrame({
        "Fecha":  pd.date_range("2024-01-01", periods=6, freq="ME"),
        "Ventas": [100, 120, 140, 160, 180, 200],
    })
    texto, buf = analisis_tendencia(df)
    assert "Ventas" in texto
    assert buf is not None


# ── C4 — Normalizar texto ─────────────────────────────────────────────────────

def test_normalizar_strip_elimina_espacios(df_texto_sucio):
    df_mod, desc, _ = aplicar_edicion(df_texto_sucio, {"op": "normalizar_texto", "accion": "strip"})
    assert df_mod["Nombre"].iloc[0] == "Ana"
    assert df_mod["Nombre"].iloc[3] == "María"
    assert "normalizado" in desc.lower()


def test_normalizar_upper(df_texto_sucio):
    df_mod, _, _ = aplicar_edicion(df_texto_sucio, {"op": "normalizar_texto", "accion": "upper"})
    assert df_mod["Ciudad"].iloc[1] == "BARCELONA"


def test_normalizar_lower(df_texto_sucio):
    df_mod, _, _ = aplicar_edicion(df_texto_sucio, {"op": "normalizar_texto", "accion": "lower"})
    assert df_mod["Ciudad"].iloc[0] == "madrid"


def test_normalizar_title(df_texto_sucio):
    df_mod, _, _ = aplicar_edicion(df_texto_sucio, {"op": "normalizar_texto", "accion": "title"})
    assert df_mod["Ciudad"].iloc[1] == "Barcelona"


def test_normalizar_todas(df_texto_sucio):
    df_mod, _, _ = aplicar_edicion(df_texto_sucio, {"op": "normalizar_texto", "accion": "todas"})
    # Strip + title
    assert df_mod["Nombre"].iloc[0] == "Ana"
    assert df_mod["Nombre"].iloc[1] == "Luis"


def test_normalizar_columna_especifica(df_texto_sucio):
    df_mod, desc, _ = aplicar_edicion(
        df_texto_sucio, {"op": "normalizar_texto", "col": "Nombre", "accion": "strip"}
    )
    assert "  Ana  ".strip() == df_mod["Nombre"].iloc[0]
    assert "Nombre" in desc


def test_normalizar_accion_invalida(df_texto_sucio):
    with pytest.raises(EditorError):
        aplicar_edicion(df_texto_sucio, {"op": "normalizar_texto", "accion": "rotar"})


def test_normalizar_no_toca_columnas_numericas(df_texto_sucio):
    df_mod, _, _ = aplicar_edicion(df_texto_sucio, {"op": "normalizar_texto", "accion": "todas"})
    assert df_mod["Importe"].tolist() == [100, 200, 150, 300]


# ── C4 — Estandarizar fechas ──────────────────────────────────────────────────

def test_estandarizar_fechas_columna_detectada(df_fechas_texto):
    df_mod, desc, _ = aplicar_edicion(df_fechas_texto, {"op": "estandarizar_fechas"})
    assert "estandarizad" in desc.lower()
    # Las 3 fechas válidas deben haberse convertido
    validas = df_mod["Fecha"].str.match(r"\d{2}/\d{2}/\d{4}", na=False)
    assert validas.sum() >= 3


def test_estandarizar_fechas_columna_explicita(df_fechas_texto):
    df_mod, desc, _ = aplicar_edicion(
        df_fechas_texto, {"op": "estandarizar_fechas", "col": "Fecha"}
    )
    assert "Fecha" in desc or "estandarizad" in desc.lower()


def test_estandarizar_fechas_sin_col_fecha_lanza_error():
    df = pd.DataFrame({"Nombre": ["Ana"], "Valor": [10]})
    with pytest.raises(EditorError):
        aplicar_edicion(df, {"op": "estandarizar_fechas"})


# ── C5 — Despivotear ──────────────────────────────────────────────────────────

def test_despivotear_columnas_en_filas(df_wide):
    df_mod, desc, _ = aplicar_edicion(df_wide, {
        "op": "despivotear",
        "columnas_valores": ["Enero", "Febrero", "Marzo"],
        "col_nombre": "Mes",
        "col_valor": "Ventas",
    })
    # 3 productos × 3 meses = 9 filas
    assert len(df_mod) == 9
    assert "Mes" in df_mod.columns
    assert "Ventas" in df_mod.columns
    assert "Producto" in df_mod.columns
    assert "despivotado" in desc.lower()


def test_despivotear_sin_columnas_valores_lanza_error(df_wide):
    with pytest.raises(EditorError):
        aplicar_edicion(df_wide, {"op": "despivotear", "columnas_valores": []})


def test_despivotear_columna_inexistente_lanza_error(df_wide):
    with pytest.raises(EditorError):
        aplicar_edicion(df_wide, {"op": "despivotear", "columnas_valores": ["NoExiste"]})


def test_despivotear_nombres_columnas_por_defecto(df_wide):
    df_mod, _, _ = aplicar_edicion(df_wide, {
        "op": "despivotear",
        "columnas_valores": ["Enero", "Febrero"],
    })
    assert "Variable" in df_mod.columns
    assert "Valor" in df_mod.columns


# ── C5 — Pivotear ─────────────────────────────────────────────────────────────

def test_pivotear_filas_en_columnas(df_long):
    df_mod, desc, _ = aplicar_edicion(df_long, {
        "op": "pivotear",
        "index": "Producto",
        "columns": "Mes",
        "values": "Ventas",
        "aggfunc": "suma",
    })
    assert "Enero" in df_mod.columns
    assert "Febrero" in df_mod.columns
    assert len(df_mod) == 2   # Teclado y Ratón
    assert "pivotado" in desc.lower()


def test_pivotear_sin_index_lanza_error(df_long):
    with pytest.raises(EditorError):
        aplicar_edicion(df_long, {
            "op": "pivotear",
            "columns": "Mes",
            "values": "Ventas",
        })


def test_pivotear_columna_inexistente_lanza_error(df_long):
    with pytest.raises(EditorError):
        aplicar_edicion(df_long, {
            "op": "pivotear",
            "index": "NoExiste",
            "columns": "Mes",
            "values": "Ventas",
        })


def test_pivotear_aggfunc_promedio(df_long):
    df_mod, _, _ = aplicar_edicion(df_long, {
        "op": "pivotear",
        "index": "Producto",
        "columns": "Mes",
        "values": "Ventas",
        "aggfunc": "promedio",
    })
    assert len(df_mod) == 2
