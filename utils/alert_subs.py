"""Suscriptores de notificaciones del sistema (alertas RAM, disco, errores).

Tabla alert_subs en la misma BD que el historial. Cada fila es un Telegram ID
que recibirá los mensajes de alerta. Se puede activar y pausar individualmente.
"""
import sqlite3
from utils.db import conectar as _db_conectar

_CREATE = """
CREATE TABLE IF NOT EXISTS alert_subs (
    telegram_id INTEGER PRIMARY KEY,
    etiqueta    TEXT    NOT NULL DEFAULT '',
    activo      INTEGER NOT NULL DEFAULT 1,
    creado_en   TEXT    NOT NULL DEFAULT (datetime('now'))
)
"""


def _init(conn) -> None:
    conn.execute(_CREATE)


def obtener_subs() -> list[dict]:
    with _db_conectar() as conn:
        _init(conn)
        cur = conn.execute(
            "SELECT telegram_id, etiqueta, activo, creado_en FROM alert_subs ORDER BY creado_en"
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]


def agregar_sub(telegram_id: int, etiqueta: str = "") -> None:
    with _db_conectar() as conn:
        _init(conn)
        conn.execute(
            "INSERT OR REPLACE INTO alert_subs (telegram_id, etiqueta, activo) VALUES (?, ?, 1)",
            (telegram_id, etiqueta),
        )


def eliminar_sub(telegram_id: int) -> bool:
    with _db_conectar() as conn:
        _init(conn)
        cur = conn.execute("DELETE FROM alert_subs WHERE telegram_id = ?", (telegram_id,))
        return cur.rowcount > 0


def toggle_sub(telegram_id: int) -> bool | None:
    """Alterna activo/inactivo. Devuelve el nuevo estado, o None si no existe."""
    with _db_conectar() as conn:
        _init(conn)
        fila = conn.execute(
            "SELECT activo FROM alert_subs WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        if fila is None:
            return None
        nuevo = 0 if fila[0] else 1
        conn.execute(
            "UPDATE alert_subs SET activo = ? WHERE telegram_id = ?", (nuevo, telegram_id)
        )
        return bool(nuevo)


def ids_activos() -> list[int]:
    """Devuelve los telegram_id con activo=1."""
    with _db_conectar() as conn:
        _init(conn)
        cur = conn.execute("SELECT telegram_id FROM alert_subs WHERE activo = 1")
        return [r[0] for r in cur.fetchall()]
