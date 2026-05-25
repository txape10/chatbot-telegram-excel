"""Vinculación entre cuenta de Telegram (user_id) y email del Add-in.

Permite que el usuario envíe archivos desde el Add-in de Excel directamente
a su chat de Telegram sin abandonar la aplicación.

Flujo:
  1. El usuario escribe /vincular email@empresa.com en Telegram.
  2. El bot almacena la asociación telegram_id ↔ email en user_links.
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
            telegram_id  INTEGER  PRIMARY KEY,
            email        TEXT     NOT NULL,
            creado_en    DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(email)
        )
    """)
    conn.commit()
    return conn


def vincular(telegram_id: int, email: str) -> None:
    """Vincula (o actualiza) el telegram_id con un email."""
    email = email.lower().strip()
    with _conectar() as conn:
        conn.execute("""
            INSERT INTO user_links (telegram_id, email)
            VALUES (?, ?)
            ON CONFLICT(telegram_id) DO UPDATE
                SET email     = excluded.email,
                    creado_en = CURRENT_TIMESTAMP
        """, (telegram_id, email))
        # Si el email ya estaba vinculado a otro telegram_id, reasignarlo
        conn.execute("""
            UPDATE user_links SET telegram_id = ?, creado_en = CURRENT_TIMESTAMP
            WHERE email = ? AND telegram_id != ?
        """, (telegram_id, email, telegram_id))
        conn.commit()


def obtener_telegram_id(email: str) -> int | None:
    """Devuelve el telegram_id vinculado a ese email, o None si no existe."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT telegram_id FROM user_links WHERE email = ?",
            (email.lower().strip(),),
        ).fetchone()
    return fila[0] if fila else None


def obtener_email(telegram_id: int) -> str | None:
    """Devuelve el email vinculado a ese telegram_id, o None si no existe."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT email FROM user_links WHERE telegram_id = ?",
            (telegram_id,),
        ).fetchone()
    return fila[0] if fila else None


def desvincular(telegram_id: int) -> bool:
    """Elimina el vínculo. Devuelve True si existía."""
    with _conectar() as conn:
        cur = conn.execute(
            "DELETE FROM user_links WHERE telegram_id = ?", (telegram_id,)
        )
        conn.commit()
    return cur.rowcount > 0
