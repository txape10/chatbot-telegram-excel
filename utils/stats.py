"""Estadísticas de uso para el panel de administración.

Consulta directamente la base de datos SQLite sin depender de módulos
de Telegram ni LLM, por lo que puede usarse desde cualquier contexto.
"""
import sqlite3
from datetime import datetime
from utils.db import conectar as _db_conectar, dict_row


def _conectar() -> sqlite3.Connection:
    conn = _db_conectar()
    conn.row_factory = dict_row   # compatible sqlite3 y libsql_experimental
    return conn


def _tabla_existe(conn: sqlite3.Connection, nombre: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (nombre,)
    ).fetchone() is not None


def obtener_estadisticas() -> dict:
    """Devuelve un diccionario con estadísticas de uso del servicio.

    Estructura devuelta:
    {
        "total_mensajes":   int,
        "total_usuarios":   int,
        "mensajes_hoy":     int,
        "usuarios": [
            {
                "user_id":          int,
                "total_mensajes":   int,
                "mensajes_enviados":int,   # solo rol 'user'
                "ultima_actividad": str,   # ISO datetime
                "version_excel":    str | None,
                "modo_respuesta":   str,
                "modo_privado":     bool,
                "email":            str | None,   # si tiene Add-in vinculado
            },
            ...
        ],
        "mensajes_por_dia": [
            {"dia": "2025-05-20", "n": 42},
            ...
        ],
    }
    """
    with _conectar() as conn:
        if not _tabla_existe(conn, "historial"):
            return {
                "total_mensajes": 0,
                "total_usuarios": 0,
                "mensajes_hoy": 0,
                "usuarios": [],
                "mensajes_por_dia": [],
            }

        total_mensajes = conn.execute(
            "SELECT COUNT(*) FROM historial"
        ).fetchone()[0]

        total_usuarios = conn.execute(
            "SELECT COUNT(DISTINCT user_id) FROM historial"
        ).fetchone()[0]

        hoy = datetime.now().strftime("%Y-%m-%d")
        mensajes_hoy = conn.execute(
            "SELECT COUNT(*) FROM historial WHERE date(creado_en) = ?", (hoy,)
        ).fetchone()[0]

        # ── Por usuario ───────────────────────────────────────────────────────
        tiene_prefs = _tabla_existe(conn, "user_prefs")
        tiene_links = _tabla_existe(conn, "user_links")

        if tiene_prefs and tiene_links:
            query_usuarios = """
                SELECT
                    h.user_id,
                    COUNT(*) AS total_mensajes,
                    SUM(CASE WHEN h.rol = 'user' THEN 1 ELSE 0 END) AS mensajes_enviados,
                    MAX(h.creado_en) AS ultima_actividad,
                    p.version_excel,
                    COALESCE(p.modo_respuesta, 'texto') AS modo_respuesta,
                    COALESCE(p.modo_privado, 0) AS modo_privado,
                    l.email
                FROM historial h
                LEFT JOIN user_prefs  p ON h.user_id = p.user_id
                LEFT JOIN user_links  l ON h.user_id = l.telegram_id
                GROUP BY h.user_id
                ORDER BY ultima_actividad DESC
            """
        elif tiene_prefs:
            query_usuarios = """
                SELECT
                    h.user_id,
                    COUNT(*) AS total_mensajes,
                    SUM(CASE WHEN h.rol = 'user' THEN 1 ELSE 0 END) AS mensajes_enviados,
                    MAX(h.creado_en) AS ultima_actividad,
                    p.version_excel,
                    COALESCE(p.modo_respuesta, 'texto') AS modo_respuesta,
                    COALESCE(p.modo_privado, 0) AS modo_privado,
                    NULL AS email
                FROM historial h
                LEFT JOIN user_prefs p ON h.user_id = p.user_id
                GROUP BY h.user_id
                ORDER BY ultima_actividad DESC
            """
        else:
            query_usuarios = """
                SELECT
                    user_id,
                    COUNT(*) AS total_mensajes,
                    SUM(CASE WHEN rol = 'user' THEN 1 ELSE 0 END) AS mensajes_enviados,
                    MAX(creado_en) AS ultima_actividad,
                    NULL AS version_excel,
                    'texto' AS modo_respuesta,
                    0 AS modo_privado,
                    NULL AS email
                FROM historial
                GROUP BY user_id
                ORDER BY ultima_actividad DESC
            """

        filas_usuarios = conn.execute(query_usuarios).fetchall()

        # ── Mensajes por día (últimos 7 días) ─────────────────────────────────
        filas_dia = conn.execute("""
            SELECT date(creado_en) AS dia, COUNT(*) AS n
            FROM historial
            WHERE creado_en >= date('now', '-6 days')
            GROUP BY dia
            ORDER BY dia
        """).fetchall()

    usuarios = [
        {
            "user_id":          fila["user_id"],
            "total_mensajes":   fila["total_mensajes"],
            "mensajes_enviados":fila["mensajes_enviados"] or 0,
            "ultima_actividad": fila["ultima_actividad"] or "—",
            "version_excel":    fila["version_excel"],
            "modo_respuesta":   fila["modo_respuesta"],
            "modo_privado":     bool(fila["modo_privado"]),
            "email":            fila["email"],
        }
        for fila in filas_usuarios
    ]

    mensajes_por_dia = [
        {"dia": fila["dia"], "n": fila["n"]}
        for fila in filas_dia
    ]

    return {
        "total_mensajes":  total_mensajes,
        "total_usuarios":  total_usuarios,
        "mensajes_hoy":    mensajes_hoy,
        "usuarios":        usuarios,
        "mensajes_por_dia": mensajes_por_dia,
    }
