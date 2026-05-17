from config import HISTORIAL_MAX_MENSAJES

# Diccionario en memoria: {user_id: [{"role": ..., "parts": [...]}]}
_historial: dict[int, list[dict]] = {}


def obtener_historial(user_id: int) -> list[dict]:
    return _historial.get(user_id, [])


def agregar_mensaje(user_id: int, rol: str, texto: str) -> None:
    if user_id not in _historial:
        _historial[user_id] = []
    _historial[user_id].append({"role": rol, "parts": [texto]})
    # Conservar solo los últimos N turnos (cada turno = 2 mensajes: user + model)
    limite = HISTORIAL_MAX_MENSAJES * 2
    if len(_historial[user_id]) > limite:
        _historial[user_id] = _historial[user_id][-limite:]


def limpiar_historial(user_id: int) -> None:
    _historial.pop(user_id, None)
