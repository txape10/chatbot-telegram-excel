# DataFrame activo por usuario (en memoria, se pierde al reiniciar el bot)
import pandas as pd

_dataframes: dict[int, pd.DataFrame] = {}


def guardar_df(user_id: int, df: pd.DataFrame) -> None:
    _dataframes[user_id] = df.copy()


def obtener_df(user_id: int) -> pd.DataFrame | None:
    return _dataframes.get(user_id)


def borrar_df(user_id: int) -> None:
    _dataframes.pop(user_id, None)
