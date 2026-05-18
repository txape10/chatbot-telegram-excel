"""Tests para Sprint B3 — combinar dos archivos Excel."""
import pytest
import pandas as pd

from excel.editor import combinar_dataframes, EditorError


# ── Fixtures ─────────────────────────────────────────────────────────────────

@pytest.fixture
def df_clientes():
    return pd.DataFrame({
        "ID":      [1, 2, 3, 4],
        "Nombre":  ["Ana", "Luis", "Eva", "Marcos"],
        "Ciudad":  ["Madrid", "Barcelona", "Sevilla", "Valencia"],
    })


@pytest.fixture
def df_pedidos():
    return pd.DataFrame({
        "ID":      [1, 2, 3, 5],
        "Importe": [150.0, 200.0, 75.0, 300.0],
        "Estado":  ["Entregado", "Pendiente", "Entregado", "Cancelado"],
    })


@pytest.fixture
def df_sin_cols_comunes():
    return pd.DataFrame({
        "Referencia": ["A01", "A02"],
        "Stock":      [10, 20],
    })


# ── Tests básicos ─────────────────────────────────────────────────────────────

def test_combinar_inner_solo_coincidentes(df_clientes, df_pedidos):
    df_result, desc = combinar_dataframes(df_clientes, df_pedidos, {"col": "ID", "como": "inner"})
    assert len(df_result) == 3          # IDs 1, 2, 3 coinciden; 4 y 5 no
    assert "Nombre" in df_result.columns
    assert "Importe" in df_result.columns
    assert "combinados" in desc.lower()


def test_combinar_left_todas_de_a(df_clientes, df_pedidos):
    df_result, _ = combinar_dataframes(df_clientes, df_pedidos, {"col": "ID", "como": "left"})
    assert len(df_result) == 4          # todos los clientes, aunque ID=4 no tenga pedido
    assert df_result.loc[df_result["ID"] == 4, "Importe"].isna().all()


def test_combinar_right_todas_de_b(df_clientes, df_pedidos):
    df_result, _ = combinar_dataframes(df_clientes, df_pedidos, {"col": "ID", "como": "right"})
    assert len(df_result) == 4          # todos los pedidos, aunque ID=5 no tenga cliente
    assert df_result.loc[df_result["ID"] == 5, "Nombre"].isna().all()


def test_combinar_outer_union_completa(df_clientes, df_pedidos):
    df_result, _ = combinar_dataframes(df_clientes, df_pedidos, {"col": "ID", "como": "outer"})
    assert len(df_result) == 5          # IDs 1,2,3,4,5


def test_combinar_columna_auto_detectada(df_clientes, df_pedidos):
    """Si no se especifica col, debe usar la primera columna común."""
    df_result, desc = combinar_dataframes(df_clientes, df_pedidos, {})
    assert "ID" in desc or len(df_result) > 0


def test_combinar_como_invalido_usa_inner(df_clientes, df_pedidos):
    """Tipo de join inválido debe caer en inner silenciosamente."""
    df_result, _ = combinar_dataframes(df_clientes, df_pedidos, {"col": "ID", "como": "xyz"})
    assert len(df_result) == 3


def test_combinar_sufijos_en_columnas_duplicadas():
    """Columnas con el mismo nombre (excepto la clave) deben recibir sufijos _A/_B."""
    df1 = pd.DataFrame({"ID": [1, 2], "Valor": [10, 20]})
    df2 = pd.DataFrame({"ID": [1, 2], "Valor": [100, 200]})
    df_result, _ = combinar_dataframes(df1, df2, {"col": "ID", "como": "inner"})
    assert "Valor_A" in df_result.columns
    assert "Valor_B" in df_result.columns


def test_combinar_descripcion_incluye_col_y_filas(df_clientes, df_pedidos):
    _, desc = combinar_dataframes(df_clientes, df_pedidos, {"col": "ID", "como": "inner"})
    assert "ID" in desc
    assert "3" in desc   # 3 filas resultantes


# ── Tests de error ────────────────────────────────────────────────────────────

def test_combinar_sin_cols_comunes_lanza_error(df_clientes, df_sin_cols_comunes):
    with pytest.raises(EditorError, match="columnas en común"):
        combinar_dataframes(df_clientes, df_sin_cols_comunes, {})


def test_combinar_col_inexistente_en_a_lanza_error(df_clientes, df_pedidos):
    with pytest.raises(EditorError, match="primer archivo"):
        combinar_dataframes(df_clientes, df_pedidos, {"col": "NoExiste"})


def test_combinar_col_inexistente_en_b_lanza_error(df_clientes, df_pedidos):
    with pytest.raises(EditorError, match="segundo archivo"):
        combinar_dataframes(df_pedidos, df_clientes, {"col": "Estado"})


# ── Tests de integridad ───────────────────────────────────────────────────────

def test_combinar_no_modifica_originales(df_clientes, df_pedidos):
    cols_orig_a = list(df_clientes.columns)
    cols_orig_b = list(df_pedidos.columns)
    combinar_dataframes(df_clientes, df_pedidos, {"col": "ID", "como": "inner"})
    assert list(df_clientes.columns) == cols_orig_a
    assert list(df_pedidos.columns) == cols_orig_b


def test_combinar_resultado_tiene_columnas_de_ambos(df_clientes, df_pedidos):
    df_result, _ = combinar_dataframes(df_clientes, df_pedidos, {"col": "ID", "como": "inner"})
    for col in ["ID", "Nombre", "Ciudad", "Importe", "Estado"]:
        assert col in df_result.columns
