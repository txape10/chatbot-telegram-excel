"""Tests Sprint E1 — gráficos bajo demanda (generar_grafico_personalizado).

No hay llamadas a API. matplotlib usa el backend Agg (sin pantalla).
"""
import io
import pytest
import pandas as pd
from excel.charts import generar_grafico_personalizado, ChartError


@pytest.fixture
def df_ventas():
    return pd.DataFrame({
        "Mes":     ["Enero", "Febrero", "Marzo", "Abril", "Mayo"],
        "Ventas":  [1000.0, 1500.0, 1200.0, 1800.0, 1600.0],
        "Gastos":  [800.0,  900.0,  1000.0, 700.0,  950.0],
        "Región":  ["Norte", "Sur", "Norte", "Sur", "Norte"],
    })


@pytest.fixture
def df_numerico():
    return pd.DataFrame({
        "X": [1.0, 2.0, 3.0, 4.0, 5.0],
        "Y": [2.0, 4.0, 6.0, 8.0, 10.0],
    })


# ── Tipos de gráfico básicos ──────────────────────────────────────────────────

def test_grafico_barras_devuelve_buffer(df_ventas):
    buffer, titulo = generar_grafico_personalizado(df_ventas, col_y="Ventas", col_x="Mes", tipo="barras")
    assert isinstance(buffer, io.BytesIO)
    assert len(buffer.read()) > 0
    assert "Ventas" in titulo


def test_grafico_lineas_devuelve_buffer(df_ventas):
    buffer, titulo = generar_grafico_personalizado(df_ventas, col_y="Ventas", col_x="Mes", tipo="lineas")
    assert isinstance(buffer, io.BytesIO)
    assert len(buffer.read()) > 0


def test_grafico_sectores_devuelve_buffer(df_ventas):
    buffer, titulo = generar_grafico_personalizado(df_ventas, col_y="Ventas", col_x="Mes", tipo="sectores")
    assert isinstance(buffer, io.BytesIO)
    assert len(buffer.read()) > 0


def test_grafico_dispersion_devuelve_buffer(df_numerico):
    buffer, titulo = generar_grafico_personalizado(df_numerico, col_y="Y", col_x="X", tipo="dispersion")
    assert isinstance(buffer, io.BytesIO)
    assert len(buffer.read()) > 0


# ── Sin col_x (solo índice) ──────────────────────────────────────────────────

def test_grafico_sin_col_x(df_ventas):
    buffer, titulo = generar_grafico_personalizado(df_ventas, col_y="Ventas", col_x=None, tipo="barras")
    assert isinstance(buffer, io.BytesIO)
    assert len(buffer.read()) > 0


def test_grafico_tipo_por_defecto(df_ventas):
    """Tipo inválido cae a barras por defecto."""
    buffer, _ = generar_grafico_personalizado(df_ventas, col_y="Ventas", col_x="Mes", tipo="pizza")
    assert isinstance(buffer, io.BytesIO)


# ── Agregación ────────────────────────────────────────────────────────────────

def test_grafico_con_agregacion_suma(df_ventas):
    buffer, titulo = generar_grafico_personalizado(
        df_ventas, col_y="Ventas", col_x="Región", tipo="barras", agregar="suma"
    )
    assert isinstance(buffer, io.BytesIO)
    assert len(buffer.read()) > 0
    assert "suma" in titulo


def test_grafico_con_agregacion_promedio(df_ventas):
    buffer, titulo = generar_grafico_personalizado(
        df_ventas, col_y="Ventas", col_x="Región", tipo="barras", agregar="promedio"
    )
    assert "promedio" in titulo


def test_grafico_con_agregacion_contar(df_ventas):
    buffer, titulo = generar_grafico_personalizado(
        df_ventas, col_y="Ventas", col_x="Región", tipo="barras", agregar="contar"
    )
    assert "contar" in titulo


def test_grafico_con_agregacion_max(df_ventas):
    buffer, titulo = generar_grafico_personalizado(
        df_ventas, col_y="Ventas", col_x="Región", tipo="barras", agregar="max"
    )
    assert "max" in titulo


def test_grafico_con_agregacion_min(df_ventas):
    buffer, titulo = generar_grafico_personalizado(
        df_ventas, col_y="Ventas", col_x="Región", tipo="barras", agregar="min"
    )
    assert "min" in titulo


# ── Errores controlados ───────────────────────────────────────────────────────

def test_error_col_y_inexistente(df_ventas):
    with pytest.raises(ChartError, match="ColInexistente"):
        generar_grafico_personalizado(df_ventas, col_y="ColInexistente", col_x="Mes")


def test_error_col_x_inexistente(df_ventas):
    with pytest.raises(ChartError, match="ColInexistente"):
        generar_grafico_personalizado(df_ventas, col_y="Ventas", col_x="ColInexistente")


def test_error_col_y_no_numerica(df_ventas):
    """Columna de texto sin valores numéricos → ChartError."""
    with pytest.raises(ChartError):
        generar_grafico_personalizado(df_ventas, col_y="Mes", col_x="Región")


def test_error_dispersion_col_x_no_numerica(df_ventas):
    """Dispersión con col_x de texto → ChartError."""
    with pytest.raises(ChartError):
        generar_grafico_personalizado(df_ventas, col_y="Ventas", col_x="Mes", tipo="dispersion")


# ── Descripción del gráfico ───────────────────────────────────────────────────

def test_titulo_sin_agregacion(df_ventas):
    _, titulo = generar_grafico_personalizado(df_ventas, col_y="Ventas", col_x="Mes")
    assert "Ventas" in titulo
    # Sin agregación no incluye texto adicional de agrupación
    assert "por" not in titulo


def test_titulo_con_agregacion_incluye_columnas(df_ventas):
    _, titulo = generar_grafico_personalizado(
        df_ventas, col_y="Ventas", col_x="Región", agregar="suma"
    )
    assert "Ventas" in titulo
    assert "Región" in titulo
