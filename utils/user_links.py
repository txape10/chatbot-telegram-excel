"""Vinculación entre cuenta de Telegram (user_id) y emails del Add-in.

Un usuario de Telegram puede tener varios emails vinculados (p.ej. cuenta
personal y cuenta de empresa). Cualquiera de ellos sirve para que el Add-in
envíe archivos directamente a ese chat.

Flujo:
  1. El usuario escribe /vincular email@empresa.com en Telegram.
  2. El bot inserta la asociación telegram_id ↔ email en user_links.
  3. Desde el Add-in, al pulsar "Enviar al bot" se llama a POST /enviar-al-bot
     con el email del usuario. La API recupera el telegram_id y reenvía el
     archivo al chat correspondiente.
"""
import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "historial.db")


def _conectar() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_links (
            telegram_id  INTEGER  NOT NULL,
            email        TEXT     NOT NULL,
            creado_en    DATETIME DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (telegram_id, email),
            UNIQUE(email)
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
