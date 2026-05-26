"""Configuración por tipo de alerta del sistema (activo/inactivo).

Tabla: alert_config(tipo TEXT PK, activo INT, label TEXT)
Si el tipo no existe en BD se considera activo (fallo seguro).
"""
from utils.db import conectar as _db_conectar

_DEFAULTS = [
    ("arranque",  "Servidor online/reiniciado"),
    ("bot_error", "Error en bot Telegram"),
    ("disco",     "Disco data/ >400 MB"),
    ("error_500", "HTTP 500 en el servidor"),
    ("ram",       "RAM alta (>80%)"),
]

_CREATE = """
CREATE TABLE IF NOT EXISTS alert_config (
    tipo   TEXT PRIMARY KEY,
    activo INTEGER NOT NULL DEFAULT 1,
    label  TEXT NOT NULL DEFAULT ''
)
"""


def _init(conn) -> None:
    conn.execute(_CREATE)
    for tipo, label in _DEFAULTS:
        conn.execute(
            "INSERT OR IGNORE INTO alert_config (tipo, activo, label) VALUES (?, 1, ?)",
            (tipo, label),
        )


def obtener_config() -> list[dict]:
    with _db_conectar() as conn:
        _init(conn)
        cur = conn.execute(
            "SELECT tipo, activo, label FROM alert_config ORDER BY tipo"
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]


def esta_activo(tipo: str) -> bool:
    """Devuelve True si el tipo de alerta está activo (o si no existe — fallo seguro)."""
    with _db_conectar() as conn:
        _init(conn)
        row = conn.execute(
            "SELECT activo FROM alert_config WHERE tipo = ?", (tipo,)
        ).fetchone()
    return bool(row[0]) if row else True


def toggle_tipo(tipo: str) -> bool | None:
    """Alterna activo/inactivo. Devuelve el nuevo estado, o None si el tipo no existe."""
    with _db_conectar() as conn:
        _init(conn)
        row = conn.execute(
            "SELECT activo FROM alert_config WHERE tipo = ?", (tipo,)
        ).fetchone()
        if row is None:
            return None
        nuevo = 0 if row[0] else 1
        conn.execute(
            "UPDATE alert_config SET activo = ? WHERE tipo = ?", (nuevo, tipo)
        )
        return bool(nuevo)
