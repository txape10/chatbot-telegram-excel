"""Tests Sprint F3 — comparar_dataframes."""
import pytest
import pandas as pd
from excel.analyzer import comparar_dataframes


@pytest.fixture
def df_base():
    return pd.DataFrame({
        "ID":     [1, 2, 3],
        "Nombre": ["Ana", "Luis", "Eva"],
        "Valor":  [100, 200, 300],
    })


@pytest.fixture
def df_igual(df_base):
    return df_base.copy()


@pytest.fixture
def df_con_fila_extra(df_base):
    extra = pd.DataFrame({"ID": [4], "Nombre": ["Mario"], "Valor": [400]})
    return pd.concat([df_base, extra], ignore_index=True)


@pytest.fixture
def df_fila_diferente():
    """Mismas columnas, distinto contenido en fila 2."""
    return pd.DataFrame({
        "ID":     [1, 2, 3],
        "Nombre": ["Ana", "CAMBIADO", "Eva"],
        "Valor":  [100, 200, 300],
    })


# ── Archivos idénticos ────────────────────────────────────────────────────────

def test_identicos_cero_diferencias(df_base, df_igual):
    texto, df_diff = comparar_dataframes(df_base, df_igual)
    assert "0" in texto or "idénticas" in texto.lower()
    assert df_diff is None


def test_identicos_texto_contiene_resumen(df_base, df_igual):
    texto, _ = comparar_dataframes(df_base, df_igual, "A", "B")
    assert "A" in texto
    assert "B" in texto


# ── Filas únicas en cada archivo ──────────────────────────────────────────────

def test_fila_extra_en_b(df_base, df_con_fila_extra):
    texto, df_diff = comparar_dataframes(df_base, df_con_fila_extra)
    assert df_diff is not None
    assert "_origen_" in df_diff.columns
    # La fila extra (Mario) está solo en B
    assert any("Archivo B" in str(v) or "solo en" in str(v).lower()
               for v in df_diff["_origen_"].values)


def test_fila_extra_en_a(df_con_fila_extra, df_base):
    texto, df_diff = comparar_dataframes(df_con_fila_extra, df_base)
    assert df_diff is not None
    assert len(df_diff) > 0


def test_df_diff_tiene_columna_origen(df_base, df_fila_diferente):
    _, df_diff = comparar_dataframes(df_base, df_fila_diferente)
    assert df_diff is not None
    assert "_origen_" in df_diff.columns


def test_df_diff_no_none_cuando_hay_diferencias(df_base, df_fila_diferente):
    _, df_diff = comparar_dataframes(df_base, df_fila_diferente)
    assert df_diff is not None
    assert len(df_diff) > 0


# ── Diferencias estructurales (columnas distintas) ────────────────────────────

def test_columna_solo_en_a():
    df_a = pd.DataFrame({"ID": [1, 2], "Extra": ["x", "y"], "Valor": [10, 20]})
    df_b = pd.DataFrame({"ID": [1, 2], "Valor": [10, 20]})
    texto, _ = comparar_dataframes(df_a, df_b)
    assert "Extra" in texto


def test_columna_solo_en_b():
    df_a = pd.DataFrame({"ID": [1], "Valor": [10]})
    df_b = pd.DataFrame({"ID": [1], "Valor": [10], "Nueva": ["z"]})
    texto, _ = comparar_dataframes(df_a, df_b)
    assert "Nueva" in texto


def test_sin_columnas_comunes():
    df_a = pd.DataFrame({"ColA": [1, 2]})
    df_b = pd.DataFrame({"ColB": [3, 4]})
    texto, df_diff = comparar_dataframes(df_a, df_b)
    assert "común" in texto.lower() or "comun" in texto.lower()
    assert df_diff is None


# ── Recuento de filas en el resumen ──────────────────────────────────────────

def test_texto_incluye_recuento_filas(df_base, df_igual):
    texto, _ = comparar_dataframes(df_base, df_igual)
    assert "3" in texto   # ambos tienen 3 filas


def test_texto_menciona_filas_identicas(df_base, df_igual):
    texto, _ = comparar_dataframes(df_base, df_igual)
    assert "idénticas" in texto.lower() or "3" in texto


# ── Nombres personalizados ────────────────────────────────────────────────────

def test_nombres_personalizados_en_texto(df_base, df_con_fila_extra):
    texto, _ = comparar_dataframes(df_base, df_con_fila_extra, "Enero", "Febrero")
    assert "Enero" in texto
    assert "Febrero" in texto


# ── DataFrames vacíos ─────────────────────────────────────────────────────────

def test_ambos_vacios():
    df_a = pd.DataFrame({"ID": [], "Valor": []})
    df_b = pd.DataFrame({"ID": [], "Valor": []})
    texto, df_diff = comparar_dataframes(df_a, df_b)
    assert texto  # debe devolver algo sin lanzar excepción
    assert df_diff is None


def test_uno_vacio_otro_con_filas():
    df_a = pd.DataFrame({"ID": [1, 2], "Valor": [10, 20]})
    df_b = pd.DataFrame({"ID": [], "Valor": []})
    texto, df_diff = comparar_dataframes(df_a, df_b)
    assert df_diff is not None or "0" in texto
