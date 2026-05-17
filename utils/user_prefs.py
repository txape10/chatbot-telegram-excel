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
