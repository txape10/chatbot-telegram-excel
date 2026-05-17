import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "historial.db")


def _conectar() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS file_meta (
            user_id    INTEGER PRIMARY KEY,
            nombre     TEXT NOT NULL,
            hoja       TEXT DEFAULT '',
            timestamp  DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    return conn


def guardar_meta(user_id: int, nombre: str, hoja: str = "") -> None:
    """Guarda o actualiza el nombre del último archivo procesado por el usuario."""
    with _conectar() as conn:
        conn.execute("""
            INSERT INTO file_meta (user_id, nombre, hoja)
            VALUES (?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE
            SET nombre = excluded.nombre,
                hoja   = excluded.hoja,
                timestamp = CURRENT_TIMESTAMP
        """, (user_id, nombre, hoja))
        conn.commit()


def obtener_meta(user_id: int) -> dict | None:
    """Devuelve metadata del último archivo o None si no hay ninguno."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT nombre, hoja, timestamp FROM file_meta WHERE user_id = ?",
            (user_id,)
        ).fetchone()
    if fila:
        return {"nombre": fila[0], "hoja": fila[1], "timestamp": fila[2]}
    return None


def borrar_meta(user_id: int) -> None:
    with _conectar() as conn:
        conn.execute("DELETE FROM file_meta WHERE user_id = ?", (user_id,))
        conn.commit()
