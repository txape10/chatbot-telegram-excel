"""Vinculación entre cuenta de Telegram (user_id) y emails/dispositivos del Add-in.

Un usuario de Telegram puede tener varios emails vinculados (p.ej. cuenta
personal y cuenta de empresa). Cualquiera de ellos sirve para que el Add-in
envíe archivos directamente a ese chat.

Flujos de vinculación:
  A) Por email (SSO / futuro Azure AD):
     1. /vincular email@empresa.com  →  user_links (telegram_id ↔ email)
     2. Add-in llama /tiene-vinculo?email=X  →  muestra botón si existe el link

  B) Por código efímero (dispositivo sin SSO):
     1. /vincular email@empresa.com  →  user_links
     2. /codigo  →  genera código de 6 dígitos en device_codes (5 min de vida)
     3. Usuario introduce el código en el Add-in
     4. Add-in llama /verificar-codigo  →  guarda device_links (device_id ↔ email)
     5. Add-in llama /tiene-vinculo?device_id=X  →  muestra botón
     6. "Enviar al bot" envía device_id; API busca telegram_id en device_links
"""
import sqlite3
from utils.db import conectar as _db_conectar


def _conectar() -> sqlite3.Connection:
    conn = _db_conectar()
    # Vínculos email ↔ Telegram (flujo A y prerequisito del flujo B)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_links (
            telegram_id  INTEGER  NOT NULL,
            email        TEXT     NOT NULL,
            creado_en    DATETIME DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (telegram_id, email),
            UNIQUE(email)
        )
    """)
    # Códigos efímeros de emparejamiento (flujo B, paso 2)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS device_codes (
            code        TEXT     PRIMARY KEY,
            telegram_id INTEGER  NOT NULL,
            email       TEXT     NOT NULL,
            expiry      TEXT     NOT NULL,
            usado       INTEGER  DEFAULT 0
        )
    """)
    # Vínculos device_id ↔ telegram_id/email (flujo B, tras verificar código)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS device_links (
            device_id   TEXT     PRIMARY KEY,
            telegram_id INTEGER  NOT NULL,
            email       TEXT     NOT NULL,
            creado_en   DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    return conn


def vincular(telegram_id: int, email: str) -> None:
    """Añade un email vinculado a este telegram_id.

    Si el email ya estaba vinculado a otro telegram_id, se reasigna.
    Si ya estaba vinculado a este mismo, no hace nada.
    """
    email = email.lower().strip()
    with _conectar() as conn:
        conn.execute("""
            INSERT OR REPLACE INTO user_links (telegram_id, email)
            VALUES (?, ?)
        """, (telegram_id, email))
        conn.commit()


def obtener_telegram_id(email: str) -> int | None:
    """Devuelve el telegram_id vinculado a ese email, o None si no existe."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT telegram_id FROM user_links WHERE email = ?",
            (email.lower().strip(),),
        ).fetchone()
    return fila[0] if fila else None


def obtener_emails(telegram_id: int) -> list[str]:
    """Devuelve todos los emails vinculados a este telegram_id (puede ser vacío)."""
    with _conectar() as conn:
        filas = conn.execute(
            "SELECT email FROM user_links WHERE telegram_id = ? ORDER BY creado_en",
            (telegram_id,),
        ).fetchall()
    return [f[0] for f in filas]


def obtener_todos_los_vinculos() -> list[dict]:
    """Devuelve todos los vínculos almacenados, ordenados por fecha descendente."""
    with _conectar() as conn:
        filas = conn.execute(
            "SELECT telegram_id, email, creado_en FROM user_links ORDER BY creado_en DESC"
        ).fetchall()
    return [{"telegram_id": f[0], "email": f[1], "creado_en": f[2]} for f in filas]


def desvincular(telegram_id: int, email: str | None = None) -> int:
    """Elimina uno o todos los emails vinculados a este telegram_id.

    Args:
        email: si se indica, elimina solo ese email; si es None, elimina todos.

    Returns:
        Número de filas eliminadas.
    """
    with _conectar() as conn:
        if email:
            cur = conn.execute(
                "DELETE FROM user_links WHERE telegram_id = ? AND email = ?",
                (telegram_id, email.lower().strip()),
            )
        else:
            cur = conn.execute(
                "DELETE FROM user_links WHERE telegram_id = ?",
                (telegram_id,),
            )
        conn.commit()
    return cur.rowcount


# ---------------------------------------------------------------------------
# Device codes — códigos efímeros de emparejamiento (flujo B)
# ---------------------------------------------------------------------------

def guardar_codigo_dispositivo(code: str, telegram_id: int, email: str, expiry: str) -> None:
    """Guarda un código efímero para emparejar un dispositivo Add-in con Telegram."""
    with _conectar() as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO device_codes (code, telegram_id, email, expiry, usado)
            VALUES (?, ?, ?, ?, 0)
            """,
            (code, telegram_id, email.lower().strip(), expiry),
        )
        conn.commit()


def obtener_codigo_dispositivo(code: str) -> dict | None:
    """Devuelve los datos del código (telegram_id, email, expiry, usado), o None si no existe."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT telegram_id, email, expiry, usado FROM device_codes WHERE code = ?",
            (code,),
        ).fetchone()
    if not fila:
        return None
    return {"telegram_id": fila[0], "email": fila[1], "expiry": fila[2], "usado": fila[3]}


def marcar_codigo_usado(code: str) -> None:
    """Marca el código como ya utilizado para que no pueda reutilizarse."""
    with _conectar() as conn:
        conn.execute("UPDATE device_codes SET usado = 1 WHERE code = ?", (code,))
        conn.commit()


# ---------------------------------------------------------------------------
# Device links — vínculos persistentes device_id ↔ email/telegram (flujo B)
# ---------------------------------------------------------------------------

def guardar_device_link(device_id: str, telegram_id: int, email: str) -> None:
    """Asocia un device_id de Add-in con un telegram_id y email."""
    with _conectar() as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO device_links (device_id, telegram_id, email)
            VALUES (?, ?, ?)
            """,
            (device_id, telegram_id, email.lower().strip()),
        )
        conn.commit()


def obtener_device_link(device_id: str) -> dict | None:
    """Devuelve el telegram_id y email asociados a este device_id, o None si no existe."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT telegram_id, email FROM device_links WHERE device_id = ?",
            (device_id,),
        ).fetchone()
    if not fila:
        return None
    return {"telegram_id": fila[0], "email": fila[1]}
