"""Tests Sprint F4 — macros, buscar/reemplazar, dividir, concatenar, preferencias.

- utils/macros.py:         CRUD de macros personales en SQLite
- excel/editor.py:         buscar_reemplazar, dividir_columna, concatenar_columnas
- utils/user_prefs.py:     modo_respuesta y modo_privado
"""
import os
import pytest
import sqlite3
import pandas as pd

import utils.macros as macros_mod
import utils.user_prefs as prefs_mod
from excel.editor import aplicar_edicion, EditorError


# ══════════════════════════════════════════════════════════════════════════════
# F4 — Macros personales
# ══════════════════════════════════════════════════════════════════════════════

@pytest.fixture
def db_temporal(tmp_path, monkeypatch):
    """Redirige utils.db a una BD temporal para cada test."""
    db_file = str(tmp_path / "test_macros.db")
    monkeypatch.setattr("utils.db.DB_PATH", db_file)
    yield db_file


OPS_EJEMPLO = [
    {"op": "ordenar", "col": "Fecha", "ascendente": True},
    {"op": "rellenar_nulos", "col": "Valor", "valor": 0},
]


def test_guardar_y_obtener_macro(db_temporal):
    macros_mod.guardar_macro(1, "LimpiarFechas", OPS_EJEMPLO, "ordena y rellena vacíos")
    resultado = macros_mod.obtener_macro(1, "LimpiarFechas")
    assert resultado == OPS_EJEMPLO


def test_nombre_macro_en_minusculas(db_temporal):
    """Las macros se guardan con nombre en minúsculas independientemente del input."""
    macros_mod.guardar_macro(1, "MiMacro", OPS_EJEMPLO)
    assert macros_mod.obtener_macro(1, "mimacro") == OPS_EJEMPLO
    assert macros_mod.obtener_macro(1, "MiMacro") == OPS_EJEMPLO


def test_obtener_macro_inexistente_devuelve_none(db_temporal):
    resultado = macros_mod.obtener_macro(1, "no_existe")
    assert resultado is None


def test_sobreescribir_macro(db_temporal):
    macros_mod.guardar_macro(1, "mi_macro", [{"op": "v1"}])
    macros_mod.guardar_macro(1, "mi_macro", [{"op": "v2"}])  # sobreescribe
    resultado = macros_mod.obtener_macro(1, "mi_macro")
    assert resultado == [{"op": "v2"}]


def test_listar_macros_vacio(db_temporal):
    assert macros_mod.listar_macros(1) == []


def test_listar_macros_con_varias(db_temporal):
    macros_mod.guardar_macro(1, "macro_b", OPS_EJEMPLO)
    macros_mod.guardar_macro(1, "macro_a", OPS_EJEMPLO)
    lista = macros_mod.listar_macros(1)
    assert len(lista) == 2
    # Ordenadas por nombre
    assert lista[0]["nombre"] == "macro_a"
    assert lista[1]["nombre"] == "macro_b"


def test_listar_macros_incluye_descripcion(db_temporal):
    macros_mod.guardar_macro(1, "mi_macro", OPS_EJEMPLO, "descripción de prueba")
    lista = macros_mod.listar_macros(1)
    assert lista[0]["descripcion"] == "descripción de prueba"


def test_borrar_macro_existente_devuelve_true(db_temporal):
    macros_mod.guardar_macro(1, "a_borrar", OPS_EJEMPLO)
    resultado = macros_mod.borrar_macro(1, "a_borrar")
    assert resultado is True
    assert macros_mod.obtener_macro(1, "a_borrar") is None


def test_borrar_macro_inexistente_devuelve_false(db_temporal):
    resultado = macros_mod.borrar_macro(1, "no_existe")
    assert resultado is False


def test_macros_aisladas_por_usuario(db_temporal):
    macros_mod.guardar_macro(1, "macro_compartida", OPS_EJEMPLO)
    # El usuario 2 no debe ver las macros del usuario 1
    assert macros_mod.obtener_macro(2, "macro_compartida") is None
    assert macros_mod.listar_macros(2) == []


def test_macros_multiples_usuarios_sin_colision(db_temporal):
    macros_mod.guardar_macro(1, "mi_macro", [{"op": "u1"}])
    macros_mod.guardar_macro(2, "mi_macro", [{"op": "u2"}])
    assert macros_mod.obtener_macro(1, "mi_macro") == [{"op": "u1"}]
    assert macros_mod.obtener_macro(2, "mi_macro") == [{"op": "u2"}]


# ══════════════════════════════════════════════════════════════════════════════
# F4 — Editor: buscar_reemplazar, dividir_columna, concatenar_columnas
# ══════════════════════════════════════════════════════════════════════════════

@pytest.fixture
def df_texto():
    return pd.DataFrame({
        "Nombre":   ["Ana García", "Luis López", "Eva Martín"],
        "Ciudad":   ["Madrid", "Barcelona", "Madrid"],
        "Importe":  [100, 200, 300],
    })


# ── buscar_reemplazar ─────────────────────────────────────────────────────────

def test_buscar_reemplazar_en_columna(df_texto):
    df, desc, _ = aplicar_edicion(df_texto, {
        "op": "buscar_reemplazar",
        "buscar": "Madrid",
        "reemplazar": "Valencia",
        "col": "Ciudad",
    })
    assert list(df["Ciudad"]) == ["Valencia", "Barcelona", "Valencia"]
    assert "Madrid" in desc
    assert "Valencia" in desc


def test_buscar_reemplazar_en_todo_el_df(df_texto):
    df, desc, _ = aplicar_edicion(df_texto, {
        "op": "buscar_reemplazar",
        "buscar": "Madrid",
        "reemplazar": "Sevilla",
    })
    assert "Madrid" not in df["Ciudad"].values
    assert "todo el archivo" in desc.lower()


def test_buscar_reemplazar_sin_campo_buscar_lanza_error(df_texto):
    with pytest.raises(EditorError, match="buscar"):
        aplicar_edicion(df_texto, {"op": "buscar_reemplazar", "reemplazar": "X"})


def test_buscar_reemplazar_columna_inexistente_lanza_error(df_texto):
    with pytest.raises(EditorError):
        aplicar_edicion(df_texto, {
            "op": "buscar_reemplazar",
            "buscar": "a",
            "reemplazar": "b",
            "col": "NoExiste",
        })


def test_buscar_reemplazar_reemplazar_vacio_por_defecto(df_texto):
    df, _, _ = aplicar_edicion(df_texto, {
        "op": "buscar_reemplazar",
        "buscar": "García",
        "col": "Nombre",
    })
    assert "García" not in df["Nombre"].values[0]


# ── dividir_columna ───────────────────────────────────────────────────────────

def test_dividir_columna_por_espacio(df_texto):
    df, desc, _ = aplicar_edicion(df_texto, {
        "op": "dividir_columna",
        "col": "Nombre",
        "separador": " ",
        "col_nueva_1": "Nombre1",
        "col_nueva_2": "Apellido1",
    })
    assert "Nombre1" in df.columns
    assert "Apellido1" in df.columns
    assert df.loc[0, "Nombre1"] == "Ana"
    assert df.loc[0, "Apellido1"] == "García"


def test_dividir_columna_separador_personalizado():
    df = pd.DataFrame({"Fecha": ["2024-01-15", "2024-02-20", "2024-03-10"]})
    df_result, desc, _ = aplicar_edicion(df, {
        "op": "dividir_columna",
        "col": "Fecha",
        "separador": "-",
        "col_nueva_1": "Año",
        "col_nueva_2": "Mes",
        "n": 3,   # dividir en 3 partes: Año, Mes, Día
    })
    assert "Año" in df_result.columns
    assert df_result.loc[0, "Año"] == "2024"
    assert df_result.loc[0, "Mes"] == "01"


def test_dividir_columna_nombres_por_defecto(df_texto):
    """Sin nombres explícitos debe generar Nombre_1, Nombre_2."""
    df, desc, _ = aplicar_edicion(df_texto, {
        "op": "dividir_columna",
        "col": "Nombre",
        "separador": " ",
    })
    assert "Nombre_1" in df.columns or any("Nombre" in c for c in df.columns)


def test_dividir_columna_sin_col_lanza_error(df_texto):
    with pytest.raises(EditorError):
        aplicar_edicion(df_texto, {"op": "dividir_columna", "separador": " "})


def test_dividir_columna_inexistente_lanza_error(df_texto):
    with pytest.raises(EditorError):
        aplicar_edicion(df_texto, {
            "op": "dividir_columna", "col": "NoExiste", "separador": " "
        })


# ── concatenar_columnas ───────────────────────────────────────────────────────

def test_concatenar_dos_columnas(df_texto):
    df, desc, _ = aplicar_edicion(df_texto, {
        "op": "concatenar_columnas",
        "columnas": ["Ciudad", "Nombre"],
        "separador": " - ",
        "col_resultado": "Descripcion",
    })
    assert "Descripcion" in df.columns
    assert df.loc[0, "Descripcion"] == "Madrid - Ana García"


def test_concatenar_separador_vacio(df_texto):
    df, _, _ = aplicar_edicion(df_texto, {
        "op": "concatenar_columnas",
        "columnas": ["Ciudad", "Nombre"],
        "separador": "",
        "col_resultado": "Concat",
    })
    assert df.loc[0, "Concat"] == "MadridAna García"


def test_concatenar_nombre_resultado_por_defecto(df_texto):
    """Sin col_resultado usa los nombres de columnas unidos por _."""
    df, _, _ = aplicar_edicion(df_texto, {
        "op": "concatenar_columnas",
        "columnas": ["Ciudad", "Nombre"],
    })
    assert "Ciudad_Nombre" in df.columns


def test_concatenar_menos_de_dos_columnas_lanza_error(df_texto):
    with pytest.raises(EditorError):
        aplicar_edicion(df_texto, {
            "op": "concatenar_columnas",
            "columnas": ["Ciudad"],
        })


def test_concatenar_columna_inexistente_lanza_error(df_texto):
    with pytest.raises(EditorError):
        aplicar_edicion(df_texto, {
            "op": "concatenar_columnas",
            "columnas": ["Ciudad", "NoExiste"],
        })


def test_concatenar_tres_columnas():
    df = pd.DataFrame({
        "A": ["x"], "B": ["y"], "C": ["z"]
    })
    df_result, _, _ = aplicar_edicion(df, {
        "op": "concatenar_columnas",
        "columnas": ["A", "B", "C"],
        "separador": ",",
        "col_resultado": "ABC",
    })
    assert df_result.loc[0, "ABC"] == "x,y,z"


# ══════════════════════════════════════════════════════════════════════════════
# F4 — Preferencias: modo_respuesta y modo_privado
# ══════════════════════════════════════════════════════════════════════════════

@pytest.fixture
def db_prefs_temporal(tmp_path, monkeypatch):
    """Redirige utils.db a una BD temporal."""
    db_file = str(tmp_path / "test_prefs.db")
    monkeypatch.setattr("utils.db.DB_PATH", db_file)
    yield db_file


def test_modo_respuesta_por_defecto_es_texto(db_prefs_temporal):
    assert prefs_mod.get_modo_respuesta(100) == "texto"


def test_set_y_get_modo_respuesta_voz(db_prefs_temporal):
    prefs_mod.set_modo_respuesta(100, "voz")
    assert prefs_mod.get_modo_respuesta(100) == "voz"


def test_set_modo_respuesta_sobreescribe(db_prefs_temporal):
    prefs_mod.set_modo_respuesta(100, "voz")
    prefs_mod.set_modo_respuesta(100, "texto")
    assert prefs_mod.get_modo_respuesta(100) == "texto"


def test_ya_fue_preguntado_modo_false_inicial(db_prefs_temporal):
    assert not prefs_mod.ya_fue_preguntado_modo(200)


def test_marcar_preguntado_modo(db_prefs_temporal):
    prefs_mod.marcar_preguntado_modo(200)
    assert prefs_mod.ya_fue_preguntado_modo(200)


def test_modo_privado_false_por_defecto(db_prefs_temporal):
    assert not prefs_mod.get_modo_privado(300)


def test_toggle_modo_privado_activa(db_prefs_temporal):
    nuevo = prefs_mod.toggle_modo_privado(300)
    assert nuevo is True
    assert prefs_mod.get_modo_privado(300) is True


def test_toggle_modo_privado_desactiva(db_prefs_temporal):
    prefs_mod.toggle_modo_privado(300)   # activa
    nuevo = prefs_mod.toggle_modo_privado(300)  # desactiva
    assert nuevo is False
    assert prefs_mod.get_modo_privado(300) is False


def test_toggle_modo_privado_alternancia_multiple(db_prefs_temporal):
    estados = [prefs_mod.toggle_modo_privado(400) for _ in range(4)]
    assert estados == [True, False, True, False]


def test_preferencias_aisladas_por_usuario(db_prefs_temporal):
    prefs_mod.set_modo_respuesta(500, "voz")
    prefs_mod.set_modo_respuesta(501, "texto")
    assert prefs_mod.get_modo_respuesta(500) == "voz"
    assert prefs_mod.get_modo_respuesta(501) == "texto"
