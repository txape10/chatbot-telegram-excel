import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "historial.db")

VERSIONES = {
    "365":  "Microsoft 365",
    "2021": "Excel 2021",
    "2019": "Excel 2019",
    "2016": "Excel 2016 o anterior",
}


def _conectar() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_prefs (
            user_id        INTEGER PRIMARY KEY,
            version_excel  TEXT,
            actualizado_en DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # Migraciones: añadir columnas nuevas si aún no existen
    for sql in [
        "ALTER TABLE user_prefs ADD COLUMN modo_respuesta  TEXT    DEFAULT 'texto'",
        "ALTER TABLE user_prefs ADD COLUMN preguntado_modo INTEGER DEFAULT 0",
        "ALTER TABLE user_prefs ADD COLUMN modo_privado    INTEGER DEFAULT 0",
    ]:
        try:
            conn.execute(sql)
        except sqlite3.OperationalError:
            pass  # columna ya existe
    conn.commit()
    return conn


def get_version(user_id: int) -> str | None:
    """Devuelve la versión guardada o None si no está configurada."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT version_excel FROM user_prefs WHERE user_id = ?", (user_id,)
        ).fetchone()
    return fila[0] if fila else None


def set_version(user_id: int, version: str) -> None:
    with _conectar() as conn:
        conn.execute("""
            INSERT INTO user_prefs (user_id, version_excel)
            VALUES (?, ?)
            ON CONFLICT(user_id) DO UPDATE
            SET version_excel = excluded.version_excel,
                actualizado_en = CURRENT_TIMESTAMP
        """, (user_id, version))
        conn.commit()


def ya_fue_preguntado(user_id: int) -> bool:
    """True si ya le hemos pedido la versión al usuario (aunque no la haya dado)."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT 1 FROM user_prefs WHERE user_id = ?", (user_id,)
        ).fetchone()
    return fila is not None


def marcar_preguntado(user_id: int) -> None:
    """Registra que ya se le preguntó, para no volver a preguntar."""
    with _conectar() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO user_prefs (user_id, version_excel) VALUES (?, NULL)",
            (user_id,)
        )
        conn.commit()


# ── Modo de respuesta (texto / voz) ──────────────────────────────────────────

def get_modo_respuesta(user_id: int) -> str:
    """Devuelve 'voz' o 'texto' (por defecto 'texto')."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT modo_respuesta FROM user_prefs WHERE user_id = ?", (user_id,)
        ).fetchone()
    return (fila[0] or "texto") if fila else "texto"


def set_modo_respuesta(user_id: int, modo: str) -> None:
    """Guarda el modo de respuesta ('texto' o 'voz')."""
    with _conectar() as conn:
        conn.execute("""
            INSERT INTO user_prefs (user_id, modo_respuesta)
            VALUES (?, ?)
            ON CONFLICT(user_id) DO UPDATE
            SET modo_respuesta  = excluded.modo_respuesta,
                actualizado_en  = CURRENT_TIMESTAMP
        """, (user_id, modo))
        conn.commit()


def ya_fue_preguntado_modo(user_id: int) -> bool:
    """True si ya le preguntamos la preferencia de voz/texto."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT preguntado_modo FROM user_prefs WHERE user_id = ?", (user_id,)
        ).fetchone()
    return bool(fila and fila[0])


def marcar_preguntado_modo(user_id: int) -> None:
    """Registra que ya se le preguntó el modo de respuesta."""
    with _conectar() as conn:
        conn.execute("""
            INSERT INTO user_prefs (user_id, preguntado_modo)
            VALUES (?, 1)
            ON CONFLICT(user_id) DO UPDATE
            SET preguntado_modo = 1
        """, (user_id,))
        conn.commit()


# ── Modo privado ──────────────────────────────────────────────────────────────

def get_modo_privado(user_id: int) -> bool:
    """True si el usuario tiene el modo privado activo."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT modo_privado FROM user_prefs WHERE user_id = ?", (user_id,)
        ).fetchone()
    return bool(fila and fila[0])


def toggle_modo_privado(user_id: int) -> bool:
    """Alterna el modo privado. Devuelve el nuevo estado (True = activo)."""
    actual = get_modo_privado(user_id)
    nuevo  = not actual
    with _conectar() as conn:
        conn.execute("""
            INSERT INTO user_prefs (user_id, modo_privado)
            VALUES (?, ?)
            ON CONFLICT(user_id) DO UPDATE
            SET modo_privado = excluded.modo_privado
        """, (user_id, int(nuevo)))
        conn.commit()
    return nuevo
