"""Configuración de módulos funcionales activables/desactivables desde el panel admin.

Tabla: feature_config(feature TEXT PK, activo INT, label TEXT)
Si el feature no existe en BD se considera activo (fallo seguro).
"""
from utils.db import conectar as _db_conectar

_DEFAULTS = [
    ("macros", "Macros personales (guardar y ejecutar secuencias de operaciones)"),
]

_CREATE = """
CREATE TABLE IF NOT EXISTS feature_config (
    feature TEXT PRIMARY KEY,
    activo  INTEGER NOT NULL DEFAULT 1,
    label   TEXT NOT NULL DEFAULT ''
)
"""


def _init(conn) -> None:
    conn.execute(_CREATE)
    for feature, label in _DEFAULTS:
        conn.execute(
            "INSERT OR IGNORE INTO feature_config (feature, activo, label) VALUES (?, 1, ?)",
            (feature, label),
        )


def obtener_config() -> list[dict]:
    with _db_conectar() as conn:
        _init(conn)
        cur = conn.execute(
            "SELECT feature, activo, label FROM feature_config ORDER BY feature"
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]


def esta_activo(feature: str) -> bool:
    """Devuelve True si el módulo está activo (o si no existe — fallo seguro)."""
    with _db_conectar() as conn:
        _init(conn)
        row = conn.execute(
            "SELECT activo FROM feature_config WHERE feature = ?", (feature,)
        ).fetchone()
    return bool(row[0]) if row else True


def toggle_feature(feature: str) -> bool | None:
    """Alterna activo/inactivo. Devuelve el nuevo estado, o None si el feature no existe."""
    with _db_conectar() as conn:
        _init(conn)
        row = conn.execute(
            "SELECT activo FROM feature_config WHERE feature = ?", (feature,)
        ).fetchone()
        if row is None:
            return None
        nuevo = 0 if row[0] else 1
        conn.execute(
            "UPDATE feature_config SET activo = ? WHERE feature = ?", (nuevo, feature)
        )
        return bool(nuevo)
