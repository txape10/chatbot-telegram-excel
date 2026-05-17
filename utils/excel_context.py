# Contexto del archivo Excel subido por cada usuario: {user_id: texto_contexto}
_contextos: dict[int, str] = {}


def guardar_contexto(user_id: int, contexto: str) -> None:
    _contextos[user_id] = contexto


def obtener_contexto(user_id: int) -> str | None:
    return _contextos.get(user_id)


def borrar_contexto(user_id: int) -> None:
    _contextos.pop(user_id, None)
