"""Tests Sprint E2 — mecanismo undo y slots secundario en df_context.

Todos los tests usan user_ids distintos para evitar colisiones de estado.
"""
import pytest
import pandas as pd
import utils.df_context as ctx


# ── Fixtures ─────────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def limpiar_estado():
    """Limpia el estado global antes y después de cada test."""
    ctx._dataframes.clear()
    ctx._dataframes_undo.clear()
    ctx._dataframes_secundarios.clear()
    ctx._nombres_secundarios.clear()
    yield
    ctx._dataframes.clear()
    ctx._dataframes_undo.clear()
    ctx._dataframes_secundarios.clear()
    ctx._nombres_secundarios.clear()


def _df(vals):
    return pd.DataFrame({"v": vals})


# ── guardar_df y obtener_df ───────────────────────────────────────────────────

def test_guardar_y_obtener_df():
    ctx.guardar_df(1, _df([1, 2, 3]))
    df = ctx.obtener_df(1)
    assert list(df["v"]) == [1, 2, 3]


def test_obtener_df_inexistente_devuelve_none():
    assert ctx.obtener_df(99) is None


def test_guardar_df_hace_copia():
    original = _df([10])
    ctx.guardar_df(1, original)
    original["v"] = [99]  # mutar el original no afecta al almacenado
    assert list(ctx.obtener_df(1)["v"]) == [10]


# ── Slot undo ────────────────────────────────────────────────────────────────

def test_primer_guardar_no_crea_undo():
    """El primer guardar no tiene estado anterior — no debe crear undo."""
    ctx.guardar_df(2, _df([1]))
    assert not ctx.hay_undo(2)


def test_segundo_guardar_crea_undo():
    ctx.guardar_df(2, _df([1]))
    ctx.guardar_df(2, _df([2]))
    assert ctx.hay_undo(2)


def test_undo_recupera_estado_anterior():
    ctx.guardar_df(3, _df([1]))
    ctx.guardar_df(3, _df([2]))
    df_restaurado = ctx.restaurar_undo(3)
    assert list(df_restaurado["v"]) == [1]
    assert list(ctx.obtener_df(3)["v"]) == [1]


def test_restaurar_undo_sin_undo_devuelve_none():
    ctx.guardar_df(4, _df([1]))
    resultado = ctx.restaurar_undo(4)
    assert resultado is None


def test_doble_restaurar_equivale_a_redo():
    """Restaurar dos veces vuelve al estado más reciente (comportamiento redo)."""
    ctx.guardar_df(5, _df([1]))   # v1
    ctx.guardar_df(5, _df([2]))   # v2 — undo apunta a v1
    ctx.restaurar_undo(5)          # activo = v1, undo = v2
    df_redo = ctx.restaurar_undo(5)  # activo = v2, undo = v1
    assert list(df_redo["v"]) == [2]


def test_hay_undo_false_tras_restaurar_dos_veces_sin_mas_guardados():
    ctx.guardar_df(6, _df([1]))
    ctx.guardar_df(6, _df([2]))
    ctx.restaurar_undo(6)   # activo=1, undo=2
    ctx.restaurar_undo(6)   # activo=2, undo=1
    ctx.restaurar_undo(6)   # activo=1, undo=2
    # Siempre alterna mientras haya estados guardados
    assert ctx.hay_undo(6)


# ── Slot secundario ──────────────────────────────────────────────────────────

def test_guardar_y_obtener_secundario():
    ctx.guardar_df_secundario(7, _df([10, 20]), "ventas.xlsx")
    df = ctx.obtener_df_secundario(7)
    assert list(df["v"]) == [10, 20]


def test_nombre_secundario_guardado():
    ctx.guardar_df_secundario(8, _df([1]), "archivo_b.xlsx")
    assert ctx.obtener_nombre_secundario(8) == "archivo_b.xlsx"


def test_nombre_secundario_por_defecto():
    ctx.guardar_df_secundario(9, _df([1]))
    assert ctx.obtener_nombre_secundario(9) == "archivo"


def test_secundario_inexistente_devuelve_none():
    assert ctx.obtener_df_secundario(99) is None


def test_borrar_secundario():
    ctx.guardar_df_secundario(10, _df([1]), "a.xlsx")
    ctx.borrar_df_secundario(10)
    assert ctx.obtener_df_secundario(10) is None
    assert ctx.obtener_nombre_secundario(10) == "archivo"


def test_secundario_hace_copia():
    original = _df([5])
    ctx.guardar_df_secundario(11, original)
    original["v"] = [99]
    assert list(ctx.obtener_df_secundario(11)["v"]) == [5]


# ── borrar_todo ───────────────────────────────────────────────────────────────

def test_borrar_todo_limpia_activo_y_secundario_y_undo():
    ctx.guardar_df(12, _df([1]))
    ctx.guardar_df(12, _df([2]))   # crea undo
    ctx.guardar_df_secundario(12, _df([3]), "b.xlsx")
    ctx.borrar_todo(12)
    assert ctx.obtener_df(12) is None
    assert ctx.obtener_df_secundario(12) is None
    assert not ctx.hay_undo(12)


def test_borrar_todo_sin_estado_no_falla():
    ctx.borrar_todo(999)   # no debe lanzar excepción


# ── Aislamiento entre usuarios ───────────────────────────────────────────────

def test_usuarios_aislados():
    ctx.guardar_df(20, _df([100]))
    ctx.guardar_df(21, _df([200]))
    assert list(ctx.obtener_df(20)["v"]) == [100]
    assert list(ctx.obtener_df(21)["v"]) == [200]


def test_undo_aislado_por_usuario():
    ctx.guardar_df(30, _df([1]))
    ctx.guardar_df(30, _df([2]))
    ctx.guardar_df(31, _df([10]))
    # usuario 31 no tiene undo, usuario 30 sí
    assert ctx.hay_undo(30)
    assert not ctx.hay_undo(31)
