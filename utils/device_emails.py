"""Persiste email y versión de Excel del usuario del Add-in por device_id.

La tabla device_emails mapea device_id → {email, excel_version} para mostrar
esta información en el panel de admin sin necesidad de vincular Telegram.
"""
from utils.db import conectar

_CREATE = """
CREATE TABLE IF NOT EXISTS device_emails (
    device_id     TEXT     PRIMARY KEY,
    email         TEXT     NOT NULL,
    excel_version TEXT,
    updated_at    DATETIME DEFAULT CURRENT_TIMESTAMP
)
"""

_ADD_VERSION_COL = "ALTER TABLE device_emails ADD COLUMN excel_version TEXT"


def _init(conn) -> None:
    conn.execute(_CREATE)
    try:
        conn.execute(_ADD_VERSION_COL)
    except Exception:
        pass  # la columna ya existe


def guardar_email(device_id: str, email: str, excel_version: str | None = None) -> None:
    """Upsert device_id → email + excel_version. Silencioso ante cualquier error."""
    if not device_id or not email:
        return
    try:
        with conectar() as conn:
            _init(conn)
            conn.execute(
                """INSERT INTO device_emails (device_id, email, excel_version, updated_at)
                   VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                   ON CONFLICT(device_id) DO UPDATE SET
                       email         = excluded.email,
                       excel_version = COALESCE(excluded.excel_version, excel_version),
                       updated_at    = excluded.updated_at""",
                (device_id, email.lower().strip(), excel_version or None),
            )
    except Exception:
        pass


def obtener_info_devices() -> list[dict]:
    """Devuelve todos los registros de device_emails con email y versión."""
    try:
        with conectar() as conn:
            _init(conn)
            cur = conn.execute(
                "SELECT device_id, email, excel_version FROM device_emails ORDER BY email"
            )
            cols = [d[0] for d in cur.description]
            return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception:
        return []


def obtener_emails_distintos() -> list[str]:
    """Devuelve todos los emails únicos del Add-in, ordenados alfabéticamente."""
    return [d["email"] for d in obtener_info_devices()]
