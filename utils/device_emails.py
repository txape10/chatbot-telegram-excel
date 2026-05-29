"""Persiste email, nombre y versión de Excel del usuario del Add-in por device_id.

Cada device_id recibe un user_id negativo estable (los IDs de Telegram son positivos,
así no hay colisión). Este user_id se usa como clave en historial, preferencias, etc.
"""
from utils.db import conectar

_CREATE = """
CREATE TABLE IF NOT EXISTS device_emails (
    device_id     TEXT     PRIMARY KEY,
    user_id       INTEGER  UNIQUE,
    email         TEXT     NOT NULL,
    display_name  TEXT,
    excel_version TEXT,
    updated_at    DATETIME DEFAULT CURRENT_TIMESTAMP
)
"""

_MIGRATIONS = [
    "ALTER TABLE device_emails ADD COLUMN excel_version TEXT",
    "ALTER TABLE device_emails ADD COLUMN user_id INTEGER",
    "ALTER TABLE device_emails ADD COLUMN display_name TEXT",
]


def _init(conn) -> None:
    conn.execute(_CREATE)
    for sql in _MIGRATIONS:
        try:
            conn.execute(sql)
        except Exception:
            pass
    # Backfill user_id para filas antiguas sin user_id (usa rowid como base negativa)
    conn.execute("UPDATE device_emails SET user_id = -rowid WHERE user_id IS NULL")


def obtener_o_crear_usuario(
    device_id: str,
    email: str | None = None,
    display_name: str | None = None,
    excel_version: str | None = None,
) -> int:
    """Devuelve el user_id estable para este device_id, creándolo si no existe.

    Los user_ids de Add-in son enteros negativos para no colisionar con Telegram IDs.
    Prioridad de identificador visible: email real > (Nombre) > (anon XXXXXXXX).
    """
    if not device_id:
        return 0  # fallback anónimo global

    nombre = (display_name or "").strip()
    email_efectivo = (
        (email or "").strip()
        or (f"({nombre})" if nombre else f"(anon {device_id[:8]})")
    )

    try:
        with conectar() as conn:
            _init(conn)
            fila = conn.execute(
                "SELECT user_id FROM device_emails WHERE device_id = ?", (device_id,)
            ).fetchone()

            if fila:
                user_id = fila[0]
                # Actualiza sin degradar: email real no se sobreescribe con label anón
                conn.execute(
                    """UPDATE device_emails SET
                        email         = CASE
                            WHEN ? NOT LIKE '(anon %)' THEN ?
                            ELSE email
                        END,
                        display_name  = COALESCE(NULLIF(?, ''), display_name),
                        excel_version = COALESCE(?, excel_version),
                        updated_at    = CURRENT_TIMESTAMP
                       WHERE device_id = ?""",
                    (email_efectivo, email_efectivo, nombre or None,
                     excel_version or None, device_id),
                )
                return user_id
            else:
                # Asigna el siguiente entero negativo disponible
                min_row = conn.execute(
                    "SELECT MIN(user_id) FROM device_emails"
                ).fetchone()
                new_uid = (min_row[0] or 0) - 1
                conn.execute(
                    """INSERT INTO device_emails
                       (device_id, user_id, email, display_name, excel_version)
                       VALUES (?, ?, ?, ?, ?)""",
                    (device_id, new_uid, email_efectivo,
                     nombre or None, excel_version or None),
                )
                return new_uid
    except Exception:
        return 0


def guardar_email(
    device_id: str,
    email: str | None,
    display_name: str | None = None,
    excel_version: str | None = None,
) -> None:
    """Wrapper de compatibilidad. Llama a obtener_o_crear_usuario."""
    obtener_o_crear_usuario(device_id, email, display_name, excel_version)


def obtener_info_devices() -> list[dict]:
    """Devuelve todos los registros con user_id, email, display_name y versión."""
    try:
        with conectar() as conn:
            _init(conn)
            cur = conn.execute(
                """SELECT device_id, user_id, email, display_name, excel_version
                   FROM device_emails ORDER BY email"""
            )
            cols = [d[0] for d in cur.description]
            return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception:
        return []


def obtener_info_por_user_id(user_id: int) -> dict | None:
    """Devuelve la info del device asociado a un user_id negativo."""
    try:
        with conectar() as conn:
            _init(conn)
            fila = conn.execute(
                """SELECT device_id, user_id, email, display_name, excel_version
                   FROM device_emails WHERE user_id = ?""",
                (user_id,),
            ).fetchone()
            if not fila:
                return None
            cols = ["device_id", "user_id", "email", "display_name", "excel_version"]
            return dict(zip(cols, fila))
    except Exception:
        return None


def obtener_emails_distintos() -> list[str]:
    """Devuelve todos los emails únicos del Add-in, ordenados alfabéticamente."""
    return [d["email"] for d in obtener_info_devices()]
