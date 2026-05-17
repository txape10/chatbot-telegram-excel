import pandas as pd

# Almacén en memoria: user_id → {df, nombre}
_datos: dict[int, dict] = {}


def guardar_datos_grafico(user_id: int, df: pd.DataFrame, nombre: str) -> None:
    _datos[user_id] = {"df": df.head(20).copy(), "nombre": nombre}


def obtener_datos_grafico(user_id: int) -> dict | None:
    return _datos.get(user_id)
