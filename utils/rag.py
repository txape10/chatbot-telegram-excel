"""Almacén de ejemplos de respuestas con feedback del usuario (feedback RAG).

Tabla: feedback_rag(id, user_id, pregunta, respuesta, tipo, creado_en)
  tipo = 'positivo' → el ejemplo se inyecta como few-shot en _construir_mensajes()
  tipo = 'negativo' → se almacena para análisis; no se inyecta en el LLM
"""
from utils.db import conectar as _db_conectar

_MAX_POR_USUARIO = 50

_CREATE = """
CREATE TABLE IF NOT EXISTS feedback_rag (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id   INTEGER NOT NULL,
    pregunta  TEXT    NOT NULL,
    respuesta TEXT    NOT NULL,
    tipo      TEXT    NOT NULL DEFAULT 'positivo',
    creado_en DATETIME DEFAULT CURRENT_TIMESTAMP
)
"""


def _init(conn) -> None:
    conn.execute(_CREATE)
    # Migración: añadir columna tipo si la tabla ya existía sin ella
    try:
        conn.execute("ALTER TABLE feedback_rag ADD COLUMN tipo TEXT NOT NULL DEFAULT 'positivo'")
    except Exception:
        pass  # La columna ya existe


def guardar_ejemplo(user_id: int, pregunta: str, respuesta: str,
                    tipo: str = "positivo") -> None:
    with _db_conectar() as conn:
        _init(conn)
        conn.execute(
            "INSERT INTO feedback_rag (user_id, pregunta, respuesta, tipo) VALUES (?, ?, ?, ?)",
            (user_id, pregunta[:2000], respuesta[:2000], tipo),
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
    """Devuelve los últimos `limite` ejemplos valorados positivamente (few-shot)."""
    with _db_conectar() as conn:
        _init(conn)
        cur = conn.execute(
            """SELECT pregunta, respuesta FROM feedback_rag
               WHERE user_id = ? AND tipo = 'positivo'
               ORDER BY creado_en DESC LIMIT ?""",
            (user_id, limite),
        )
        return [{"pregunta": r[0], "respuesta": r[1]} for r in cur.fetchall()]
