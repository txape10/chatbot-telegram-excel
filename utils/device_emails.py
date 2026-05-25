"""Persiste el email Office del usuario del Add-in por device_id.

La tabla device_emails actúa como mapeo device_id → email para mostrar
el correo del usuario en el panel de admin aunque no haya vinculado Telegram.
"""
from utils.db import conectar

_CREATE = """
CREATE TABLE IF NOT EXISTS device_emails (
    device_id  TEXT     PRIMARY KEY,
    email      TEXT     NOT NULL,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
)
"""


def _init(conn) -> None:
    conn.execute(_CREATE)


def guardar_email(device_id: str, email: str) -> None:
    """Upsert device_id → email. Silencioso ante cualquier error."""
    if not device_id or not email:
        return
    try:
        with conectar() as conn:
            _init(conn)
            conn.execute(
                """INSERT INTO device_emails (device_id, email, updated_at)
                   VALUES (?, ?, CURRENT_TIMESTAMP)
                   ON CONFLICT(device_id) DO UPDATE SET
                       email      = excluded.email,
                       updated_at = excluded.updated_at""",
                (device_id, email.lower().strip()),
            )
    except Exception:
        pass


def obtener_emails_distintos() -> list[str]:
    """Devuelve todos los emails únicos del Add-in, ordenados alfabéticamente."""
    try:
        with conectar() as conn:
            _init(conn)
            cur = conn.execute(
                "SELECT DISTINCT email FROM device_emails ORDER BY email"
            )
            return [row[0] for row in cur.fetchall()]
    except Exception:
        return []
