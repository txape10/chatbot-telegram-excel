# DataFrame activo, secundario y undo por usuario (en memoria, se pierde al reiniciar el bot)
import pandas as pd

_dataframes:            dict[int, pd.DataFrame] = {}
_dataframes_secundarios: dict[int, pd.DataFrame] = {}
_nombres_secundarios:   dict[int, str]           = {}
_dataframes_undo:       dict[int, pd.DataFrame] = {}   # snapshot previo a la última edición


def guardar_df(user_id: int, df: pd.DataFrame) -> None:
    # Guarda snapshot para undo ANTES de reemplazar el activo
    if user_id in _dataframes:
        _dataframes_undo[user_id] = _dataframes[user_id]   # ya es copia
    _dataframes[user_id] = df.copy()


# ── Slot de undo ─────────────────────────────────────────────────────────────

def obtener_df_undo(user_id: int) -> pd.DataFrame | None:
    return _dataframes_undo.get(user_id)


def restaurar_undo(user_id: int) -> pd.DataFrame | None:
    """Intercambia activo ↔ undo. Devuelve el df restaurado o None si no hay undo."""
    df_undo = _dataframes_undo.pop(user_id, None)
    if df_undo is None:
        return None
    # El activo actual pasa a ser el undo (para poder re-hacer si se quiere)
    if user_id in _dataframes:
        _dataframes_undo[user_id] = _dataframes[user_id]
    _dataframes[user_id] = df_undo
    return df_undo


def hay_undo(user_id: int) -> bool:
    return user_id in _dataframes_undo


def obtener_df(user_id: int) -> pd.DataFrame | None:
    return _dataframes.get(user_id)


def borrar_df(user_id: int) -> None:
    _dataframes.pop(user_id, None)


# ── Slot secundario (para combinar dos archivos) ──────────────────────────────

def guardar_df_secundario(user_id: int, df: pd.DataFrame, nombre: str = "archivo") -> None:
    _dataframes_secundarios[user_id] = df.copy()
    _nombres_secundarios[user_id] = nombre


def obtener_df_secundario(user_id: int) -> pd.DataFrame | None:
    return _dataframes_secundarios.get(user_id)


def obtener_nombre_secundario(user_id: int) -> str:
    return _nombres_secundarios.get(user_id, "archivo")


def borrar_df_secundario(user_id: int) -> None:
    _dataframes_secundarios.pop(user_id, None)
    _nombres_secundarios.pop(user_id, None)


def borrar_todo(user_id: int) -> None:
    """Limpia el df activo, el secundario y el undo (usado por /limpiar)."""
    borrar_df(user_id)
    borrar_df_secundario(user_id)
    _dataframes_undo.pop(user_id, None)
