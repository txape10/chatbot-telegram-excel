import sqlite3
import os
from config import HISTORIAL_MAX_MENSAJES

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "historial.db")


def _conectar() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS historial (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id   INTEGER NOT NULL,
            rol       TEXT    NOT NULL,
            texto     TEXT    NOT NULL,
            creado_en DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    return conn


def obtener_historial(user_id: int) -> list[dict]:
    limite = HISTORIAL_MAX_MENSAJES * 2
    with _conectar() as conn:
        filas = conn.execute(
            """
            SELECT rol, texto FROM (
                SELECT rol, texto, creado_en
                FROM historial
                WHERE user_id = ?
                ORDER BY creado_en DESC
                LIMIT ?
            ) ORDER BY creado_en ASC
            """,
            (user_id, limite),
        ).fetchall()
    return [{"role": rol, "parts": [texto]} for rol, texto in filas]


def agregar_mensaje(user_id: int, rol: str, texto: str) -> None:
    with _conectar() as conn:
        conn.execute(
            "INSERT INTO historial (user_id, rol, texto) VALUES (?, ?, ?)",
            (user_id, rol, texto),
        )
        conn.commit()


def limpiar_historial(user_id: int) -> None:
    with _conectar() as conn:
        conn.execute("DELETE FROM historial WHERE user_id = ?", (user_id,))
        conn.commit()
