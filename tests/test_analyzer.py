import pytest
import pandas as pd
from excel.analyzer import resumir, construir_contexto, analizar_calidad


# ── resumir ──────────────────────────────────────────────────────────────────

def test_resumir_sin_problemas():
    df = pd.DataFrame({"Nombre": ["Ana", "Luis"], "Edad": [25, 30]})
    resumen = resumir(df, "test.xlsx")
    assert "test.xlsx" in resumen
    assert "Sin celdas vacías" in resumen
    assert "Sin duplicados" in resumen


def test_resumir_con_nulos():
    df = pd.DataFrame({"A": [1, None, 3], "B": ["x", "y", "z"]})
    resumen = resumir(df, "test.xlsx")
    assert "Celdas vacías" in resumen


def test_resumir_con_duplicados():
    df = pd.DataFrame({"A": [1, 1], "B": ["x", "x"]})
    resumen = resumir(df, "test.xlsx")
    assert "duplicada" in resumen


def test_resumir_con_errores_formula():
    df = pd.DataFrame({"A": [1, 2]})
    errores = ["Hoja1!B2: `#REF!`"]
    resumen = resumir(df, "test.xlsx", errores=errores)
    assert "Errores de fórmula" in resumen
    assert "#REF!" in resumen


def test_resumir_incluye_nombre_archivo():
    df = pd.DataFrame({"X": [1]})
    resumen = resumir(df, "mi_archivo.xlsx")
    assert "mi_archivo.xlsx" in resumen


# ── construir_contexto ───────────────────────────────────────────────────────

def test_construir_contexto_contiene_columnas():
    df = pd.DataFrame({"Producto": ["A"], "Precio": [10.0]})
    ctx = construir_contexto(df, "ventas.xlsx")
    assert "Producto" in ctx
    assert "Precio" in ctx
    assert "ventas.xlsx" in ctx


def test_construir_contexto_menciona_nulos():
    df = pd.DataFrame({"A": [1, None], "B": ["x", "y"]})
    ctx = construir_contexto(df, "test.xlsx")
    assert "vacías" in ctx or "nulos" in ctx.lower() or "A" in ctx


# ── analizar_calidad ─────────────────────────────────────────────────────────

def test_calidad_columna_constante():
    df = pd.DataFrame({"Estado": ["Activo"] * 5, "Valor": [1, 2, 3, 4, 5]})
    avisos = analizar_calidad(df)
    assert any("constante" in a for a in avisos)


def test_calidad_casi_vacia():
    # 9 de 10 nulos = 90% > umbral del 80%
    df = pd.DataFrame({"A": [None] * 9 + [1]})
    avisos = analizar_calidad(df)
    assert any("vacíos" in a for a in avisos)


def test_calidad_mezcla_texto_numero():
    df = pd.DataFrame({"Precio": ["10", "veinte", "30", "cuarenta", "50"]})
    avisos = analizar_calidad(df)
    assert any("mezcla" in a for a in avisos)


def test_calidad_outliers():
    valores = [10, 11, 10, 12, 11, 10, 1000]   # 1000 es outlier claro
    df = pd.DataFrame({"Ventas": valores})
    avisos = analizar_calidad(df)
    assert any("atípico" in a for a in avisos)


def test_calidad_sin_problemas():
    df = pd.DataFrame({"A": [1, 2, 3], "B": ["x", "y", "z"]})
    avisos = analizar_calidad(df)
    assert avisos == []
