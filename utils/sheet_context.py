import pandas as pd

# Almacén en memoria: user_id → {hojas: {nombre: df}, nombres: [str]}
_datos: dict[int, dict] = {}


def guardar_hojas(user_id: int, sheets: dict[str, pd.DataFrame]) -> None:
    _datos[user_id] = {
        "hojas": {k: v for k, v in sheets.items()},
        "nombres": list(sheets.keys()),
    }


def obtener_hoja(user_id: int, indice: int) -> tuple[str, pd.DataFrame] | None:
    datos = _datos.get(user_id)
    if not datos or indice >= len(datos["nombres"]):
        return None
    nombre = datos["nombres"][indice]
    return nombre, datos["hojas"][nombre]


def listar_hojas(user_id: int) -> list[str]:
    datos = _datos.get(user_id)
    return datos["nombres"] if datos else []
