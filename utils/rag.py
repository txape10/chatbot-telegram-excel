"""Almacén de ejemplos de respuestas valoradas positivamente (feedback RAG).

Tabla: feedback_rag(id, user_id, pregunta, respuesta, creado_en)
Los ejemplos se inyectan como few-shot en _construir_mensajes() para adaptar
el estilo y profundidad de las respuestas al usuario concreto.
"""
from utils.db import conectar as _db_conectar

_MAX_POR_USUARIO = 50

_CREATE = """
CREATE TABLE IF NOT EXISTS feedback_rag (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id   INTEGER NOT NULL,
    pregunta  TEXT    NOT NULL,
    respuesta TEXT    NOT NULL,
    creado_en DATETIME DEFAULT CURRENT_TIMESTAMP
)
"""


def _init(conn) -> None:
    conn.execute(_CREATE)


def guardar_ejemplo(user_id: int, pregunta: str, respuesta: str) -> None:
    with _db_conectar() as conn:
        _init(conn)
        conn.execute(
            "INSERT INTO feedback_rag (user_id, pregunta, respuesta) VALUES (?, ?, ?)",
            (user_id, pregunta[:2000], respuesta[:2000]),
        )
        # Mantener solo los últimos _MAX_POR_USUARIO por usuario
        conn.execute(
            """DELETE FROM feedback_rag WHERE user_id = ? AND id NOT IN (
                SELECT id FROM feedback_rag WHERE user_id = ?
                ORDER BY creado_en DESC LIMIT ?
            )""",
            (user_id, user_id, _MAX_POR_USUARIO),
        )


def obtener_ejemplos(user_id: int, limite: int = 3) -> list[dict]:
    """Devuelve los últimos `limite` ejemplos valorados positivamente por el usuario."""
    with _db_conectar() as conn:
        _init(conn)
        cur = conn.execute(
            """SELECT pregunta, respuesta FROM feedback_rag
               WHERE user_id = ? ORDER BY creado_en DESC LIMIT ?""",
            (user_id, limite),
        )
        return [{"pregunta": r[0], "respuesta": r[1]} for r in cur.fetchall()]
