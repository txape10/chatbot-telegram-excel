"""Factory de conexión a base de datos.

- Render (cloud): TURSO_DATABASE_URL + TURSO_AUTH_TOKEN definidas
  → usa libsql-experimental (réplica embebida sincronizada con Turso).
- Servidor empresa / local: variables no definidas
  → SQLite estándar en data/historial.db.

Solo se importa libsql_experimental si Turso está configurado,
por lo que no hace falta instalar el paquete en entornos locales.
"""
import logging
import os
import sqlite3

logger = logging.getLogger(__name__)

_TURSO_URL   = os.getenv("TURSO_DATABASE_URL", "").strip()
_TURSO_TOKEN = os.getenv("TURSO_AUTH_TOKEN",   "").strip()

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "historial.db")


def conectar() -> sqlite3.Connection:
    """Devuelve una conexión lista para usar.

    En modo Turso:
      - La réplica local se sincroniza con la nube al abrir conexión.
      - Las escrituras se propagan automáticamente a Turso.
    En modo local:
      - SQLite estándar, sin dependencias extra.

    Si Turso falla (paquete no disponible, credenciales erróneas, red…)
    se registra el error y se cae a SQLite local para no crashear la app.
    """
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

    if _TURSO_URL and _TURSO_TOKEN:
        try:
            import libsql_experimental as libsql      # solo en cloud
            conn = libsql.connect(DB_PATH, sync_url=_TURSO_URL, auth_token=_TURSO_TOKEN)
            conn.sync()
            return conn  # type: ignore[return-value]
        except Exception as exc:
            logger.error("Turso no disponible (%s) — usando SQLite local como fallback", exc)

    return sqlite3.connect(DB_PATH)


def estado() -> dict:
    """Devuelve información sobre la conexión activa para el panel de admin.

    Retorna:
        modo:    "turso" | "sqlite"
        url:     URL de Turso o ruta local del archivo
        ok:      True si la conexión funciona, False si hay error
        error:   descripción del error si ok=False
    """
    ruta_local = os.path.abspath(DB_PATH)

    if _TURSO_URL and _TURSO_TOKEN:
        try:
            import libsql_experimental as libsql
            conn = libsql.connect(DB_PATH, sync_url=_TURSO_URL, auth_token=_TURSO_TOKEN)
            conn.sync()
            conn.execute("SELECT 1").fetchone()
            return {"modo": "turso", "url": _TURSO_URL, "ok": True, "error": ""}
        except Exception as exc:
            return {"modo": "turso", "url": _TURSO_URL, "ok": False, "error": str(exc)}

    try:
        conn = sqlite3.connect(ruta_local)
        conn.execute("SELECT 1").fetchone()
        return {"modo": "sqlite", "url": ruta_local, "ok": True, "error": ""}
    except Exception as exc:
        return {"modo": "sqlite", "url": ruta_local, "ok": False, "error": str(exc)}


def dict_row(cursor, row) -> dict:
    """Row factory que devuelve diccionarios en lugar de tuplas.

    Compatible con sqlite3 y libsql_experimental.
    Úsalo en módulos que necesiten acceso por nombre de columna.
    """
    return {description[0]: value for description, value in zip(cursor.description, row)}
