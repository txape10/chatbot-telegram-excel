import pytest
import pandas as pd
from excel.query_engine import (ejecutar_query, formatear_resultado,
                                 QueryError, _limpiar_filas_resumen)


@pytest.fixture
def df_ventas():
    return pd.DataFrame({
        "Producto": ["A", "B", "A", "C", "B"],
        "Región":   ["Norte", "Sur", "Norte", "Sur", "Norte"],
        "Ventas":   [100, 200, 150, 300, 50],
    })


@pytest.fixture
def df_por_region():
    """DataFrame con 3 filas por región, para pruebas de top_n_por_grupo."""
    return pd.DataFrame({
        "Región":  ["Norte", "Norte", "Norte", "Sur", "Sur", "Sur"],
        "Ventas":  [100, 300, 200, 50, 400, 150],
        "Producto": ["A", "B", "C", "A", "B", "C"],
    })


# ── filtrar ──────────────────────────────────────────────────────────────────

def test_filtrar_por_igualdad(df_ventas):
    resultado, desc = ejecutar_query(df_ventas, {
        "op": "filtrar",
        "filtros": [{"col": "Región", "op": "==", "val": "Norte"}],
    })
    assert len(resultado) == 3
    assert "filas" in desc


def test_filtrar_por_mayor(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {
        "op": "filtrar",
        "filtros": [{"col": "Ventas", "op": ">", "val": 100}],
    })
    assert len(resultado) == 3


def test_filtrar_contiene(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {
        "op": "filtrar",
        "filtros": [{"col": "Producto", "op": "contiene", "val": "A"}],
    })
    assert len(resultado) == 2


def test_filtrar_multiples_condiciones(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {
        "op": "filtrar",
        "filtros": [
            {"col": "Región", "op": "==",  "val": "Norte"},
            {"col": "Ventas", "op": ">=", "val": 100},
        ],
    })
    # Norte con Ventas >= 100: filas 100 y 150 (no la de 50)
    assert len(resultado) == 2


def test_filtrar_sin_resultados(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {
        "op": "filtrar",
        "filtros": [{"col": "Región", "op": "==", "val": "Oeste"}],
    })
    assert len(resultado) == 0


# ── contar ───────────────────────────────────────────────────────────────────

def test_contar_total(df_ventas):
    resultado, desc = ejecutar_query(df_ventas, {"op": "contar"})
    assert resultado == 5
    assert "Total" in desc


def test_contar_por_grupo(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {"op": "contar", "por": "Producto"})
    assert len(resultado) == 3          # A, B, C
    assert "Recuento" in resultado.columns


# ── suma / promedio / max / min ───────────────────────────────────────────────

def test_suma_total(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {"op": "suma", "col": "Ventas"})
    assert resultado == 800.0


def test_suma_por_grupo(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {
        "op": "suma", "col": "Ventas", "por": "Región",
    })
    norte = resultado[resultado["Región"] == "Norte"].iloc[0, -1]
    assert norte == 300.0   # 100 + 150 + 50


def test_promedio(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {"op": "promedio", "col": "Ventas"})
    assert resultado == 160.0   # 800 / 5


def test_max(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {"op": "max", "col": "Ventas"})
    assert resultado == 300.0


def test_min(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {"op": "min", "col": "Ventas"})
    assert resultado == 50.0


# ── _limpiar_filas_resumen ───────────────────────────────────────────────────

def test_limpiar_fila_total_general(df_ventas):
    df_con_total = pd.concat([
        df_ventas,
        pd.DataFrame({"Producto": ["Total general"], "Región": [""], "Ventas": [800]}),
    ], ignore_index=True)
    resultado = _limpiar_filas_resumen(df_con_total)
    assert len(resultado) == 5
    assert "Total general" not in resultado["Producto"].values


def test_limpiar_fila_subtotal(df_ventas):
    df_con_sub = pd.concat([
        df_ventas,
        pd.DataFrame({"Producto": ["Subtotal"], "Región": [""], "Ventas": [250]}),
    ], ignore_index=True)
    resultado = _limpiar_filas_resumen(df_con_sub)
    assert len(resultado) == 5


def test_limpiar_no_elimina_filas_normales(df_ventas):
    resultado = _limpiar_filas_resumen(df_ventas)
    assert len(resultado) == 5


def test_limpiar_df_vacio():
    df_vacio = pd.DataFrame({"Col": []})
    resultado = _limpiar_filas_resumen(df_vacio)
    assert resultado.empty


# ── top_n ────────────────────────────────────────────────────────────────────

def test_top_n_desc(df_ventas):
    resultado, desc = ejecutar_query(df_ventas, {
        "op": "top_n", "col": "Ventas", "n": 3,
    })
    assert len(resultado) == 3
    assert resultado.iloc[0]["Ventas"] == 300   # el mayor


def test_top_n_asc(df_ventas):
    resultado, desc = ejecutar_query(df_ventas, {
        "op": "top_n", "col": "Ventas", "n": 2, "orden": "asc",
    })
    assert resultado.iloc[0]["Ventas"] == 50    # el menor
    assert "ltimos" in desc                     # "Últimos"


# ── top_n_por_grupo ──────────────────────────────────────────────────────────

def test_top_n_por_grupo_devuelve_n_por_region(df_por_region):
    resultado, desc = ejecutar_query(df_por_region, {
        "op": "top_n_por_grupo", "col": "Ventas", "por": "Región", "n": 2,
    })
    assert len(resultado) == 4          # 2 por cada una de las 2 regiones
    assert "Top" in desc


def test_top_n_por_grupo_ordenado_por_region(df_por_region):
    resultado, _ = ejecutar_query(df_por_region, {
        "op": "top_n_por_grupo", "col": "Ventas", "por": "Región", "n": 2,
    })
    regiones = resultado["Región"].tolist()
    # Todas las filas de Norte deben aparecer antes que las de Sur
    assert regiones.index("Norte") < regiones.index("Sur")


def test_top_n_por_grupo_top_correcto_por_grupo(df_por_region):
    resultado, _ = ejecutar_query(df_por_region, {
        "op": "top_n_por_grupo", "col": "Ventas", "por": "Región", "n": 1,
    })
    norte = resultado[resultado["Región"] == "Norte"]["Ventas"].iloc[0]
    sur   = resultado[resultado["Región"] == "Sur"]["Ventas"].iloc[0]
    assert norte == 300    # máximo de Norte
    assert sur   == 400    # máximo de Sur


def test_top_n_por_grupo_excluye_grupo_vacio():
    df = pd.DataFrame({
        "Región":  ["Norte", "Norte", "", None],
        "Ventas":  [100, 200, 999, 888],
    })
    resultado, _ = ejecutar_query(df, {
        "op": "top_n_por_grupo", "col": "Ventas", "por": "Región", "n": 2,
    })
    regiones = resultado["Región"].tolist()
    assert "" not in regiones
    assert None not in regiones


def test_top_n_por_grupo_asc(df_por_region):
    resultado, desc = ejecutar_query(df_por_region, {
        "op": "top_n_por_grupo", "col": "Ventas", "por": "Región", "n": 1, "orden": "asc",
    })
    norte = resultado[resultado["Región"] == "Norte"]["Ventas"].iloc[0]
    assert norte == 100    # mínimo de Norte
    assert "ltimos" in desc


# ── ordenar ──────────────────────────────────────────────────────────────────

def test_ordenar_asc(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {
        "op": "ordenar", "col": "Ventas", "orden": "asc",
    })
    assert resultado.iloc[0]["Ventas"] == 50
    assert len(resultado) == 5


def test_ordenar_desc(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {
        "op": "ordenar", "col": "Ventas", "orden": "desc",
    })
    assert resultado.iloc[0]["Ventas"] == 300


# ── agrupar ──────────────────────────────────────────────────────────────────

def test_agrupar_suma(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {
        "op": "agrupar", "por": "Producto", "col": "Ventas", "agg": "suma",
    })
    assert "Suma de Ventas" in resultado.columns
    assert len(resultado) == 3


def test_agrupar_sin_col(df_ventas):
    resultado, _ = ejecutar_query(df_ventas, {
        "op": "agrupar", "por": "Región",
    })
    assert "Recuento" in resultado.columns


# ── errores controlados ──────────────────────────────────────────────────────

def test_columna_inexistente_lanza_error(df_ventas):
    with pytest.raises(QueryError):
        ejecutar_query(df_ventas, {"op": "suma", "col": "NoExiste"})


def test_operacion_desconocida_lanza_error(df_ventas):
    with pytest.raises(QueryError):
        ejecutar_query(df_ventas, {"op": "inventar"})


def test_operador_filtro_invalido_lanza_error(df_ventas):
    with pytest.raises(QueryError):
        ejecutar_query(df_ventas, {
            "op": "filtrar",
            "filtros": [{"col": "Región", "op": "LIKE", "val": "Norte"}],
        })


# ── formatear_resultado ───────────────────────────────────────────────────────

def test_formatear_dataframe(df_ventas):
    texto = formatear_resultado(df_ventas, "Resultado test")
    assert "Resultado test" in texto
    assert "```" in texto


def test_formatear_escalar():
    texto = formatear_resultado(800.0, "Suma total")
    assert "800.0" in texto
    assert "Suma total" in texto


def test_formatear_dataframe_vacio():
    df_vacio = pd.DataFrame({"A": []})
    texto = formatear_resultado(df_vacio, "Sin datos")
    assert "Sin resultados" in texto
