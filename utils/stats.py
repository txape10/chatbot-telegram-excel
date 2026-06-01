"""Estadísticas de uso para el panel de administración.

Consulta directamente la base de datos SQLite sin depender de módulos
de Telegram ni LLM, por lo que puede usarse desde cualquier contexto.
"""
import sqlite3
from datetime import datetime
from utils.db import conectar as _db_conectar


def _conectar() -> sqlite3.Connection:
    return _db_conectar()


def _filas_a_dicts(cursor) -> list[dict]:
    """Convierte los resultados del cursor a lista de dicts.
    Compatible con sqlite3 y libsql_experimental (que no soporta row_factory).
    """
    cols = [d[0] for d in cursor.description]
    return [dict(zip(cols, row)) for row in cursor.fetchall()]


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

        tiene_device_links = _tabla_existe(conn, "device_links")

        if tiene_prefs and tiene_links:
            # Unimos user_links y device_links para no perder usuarios que solo
            # completaron el emparejamiento por código (device_links) sin tener
            # entrada en user_links.
            fuente_links = """(
                SELECT telegram_id, MIN(email) AS email
                FROM (
                    SELECT telegram_id, email FROM user_links
                    {union_device}
                )
                GROUP BY telegram_id
            )""".format(
                union_device="UNION SELECT telegram_id, email FROM device_links"
                if tiene_device_links else ""
            )
            query_usuarios = f"""
                SELECT
                    h.user_id,
                    COUNT(*) AS total_mensajes,
                    SUM(CASE WHEN h.rol = 'user' THEN 1 ELSE 0 END) AS mensajes_enviados,
                    MAX(h.creado_en) AS ultima_actividad,
                    p.version_excel,
                    COALESCE(p.modo_respuesta, 'texto') AS modo_respuesta,
                    COALESCE(p.modo_privado, 0) AS modo_privado,
                    l.email,
                    p.display_name
                FROM historial h
                LEFT JOIN user_prefs p ON h.user_id = p.user_id
                LEFT JOIN {fuente_links} l ON h.user_id = l.telegram_id
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
                    NULL AS email,
                    p.display_name
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
                    NULL AS email,
                    NULL AS display_name
                FROM historial
                GROUP BY user_id
                ORDER BY ultima_actividad DESC
            """

        cur_usuarios = conn.execute(query_usuarios)
        filas_usuarios = _filas_a_dicts(cur_usuarios)

        # ── Mensajes por día (últimos 7 días) ─────────────────────────────────
        cur_dia = conn.execute("""
            SELECT date(creado_en) AS dia, COUNT(*) AS n
            FROM historial
            WHERE creado_en >= date('now', '-6 days')
            GROUP BY dia
            ORDER BY dia
        """)
        filas_dia = _filas_a_dicts(cur_dia)

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
            "display_name":     fila.get("display_name"),
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


def obtener_stats_usuarios_avanzadas() -> dict:
    """Estadísticas de sesiones por usuario.

    Una sesión es un bloque de mensajes con menos de 30 min de silencio entre ellos.
    Usa funciones de ventana (LAG, SUM OVER) disponibles en SQLite >= 3.25 y libSQL.

    Devuelve:
    {
        "por_usuario": [
            {
                "user_id":              int,
                "total_sesiones":       int,
                "duracion_media_min":   float,   # 0.0 si sesión de 1 solo mensaje
                "sesion_mas_larga_min": float,
                "primera_sesion":       str,     # ISO datetime
                "ultima_sesion":        str,
                "dias_inactivo":        float,   # 0.0 si la última sesión es hoy
                "es_ocasional":         bool,    # ≤ 2 sesiones y sin volver en > 7 d
            },
            ...  # ordenados por ultima_sesion desc
        ],
        "resumen": {
            "activos_7d":               int,
            "activos_30d":              int,
            "duracion_media_global_min":float,
            "sesion_mas_larga_min":     float,
            "total_sesiones":           int,
        }
    }
    """
    _vacio = {
        "por_usuario": [],
        "resumen": {
            "activos_7d": 0, "activos_30d": 0,
            "duracion_media_global_min": 0.0,
            "sesion_mas_larga_min": 0.0,
            "total_sesiones": 0,
        },
    }
    with _conectar() as conn:
        if not _tabla_existe(conn, "historial"):
            return _vacio

        sql = """
        WITH con_prev AS (
            SELECT
                user_id,
                creado_en,
                LAG(creado_en) OVER (
                    PARTITION BY user_id ORDER BY creado_en
                ) AS prev_msg
            FROM historial
            WHERE rol = 'user'
        ),
        con_sesion AS (
            SELECT
                user_id,
                creado_en,
                SUM(CASE
                    WHEN prev_msg IS NULL THEN 1
                    WHEN (julianday(creado_en) - julianday(prev_msg)) * 1440.0 > 30 THEN 1
                    ELSE 0
                END) OVER (
                    PARTITION BY user_id ORDER BY creado_en
                ) AS ses_num
            FROM con_prev
        ),
        sesiones AS (
            SELECT
                user_id,
                ses_num,
                MIN(creado_en) AS inicio,
                MAX(creado_en) AS fin,
                (julianday(MAX(creado_en)) - julianday(MIN(creado_en))) * 1440.0 AS dur_min
            FROM con_sesion
            GROUP BY user_id, ses_num
        )
        SELECT
            user_id,
            COUNT(*)                                           AS total_sesiones,
            ROUND(AVG(dur_min), 1)                             AS dur_media_min,
            ROUND(MAX(dur_min), 1)                             AS dur_max_min,
            MIN(inicio)                                        AS primera_sesion,
            MAX(fin)                                           AS ultima_sesion,
            ROUND(
                (julianday('now') - julianday(MAX(fin))), 1
            )                                                  AS dias_inactivo
        FROM sesiones
        GROUP BY user_id
        ORDER BY ultima_sesion DESC
        """
        try:
            cur = conn.execute(sql)
            filas = _filas_a_dicts(cur)
        except Exception:
            return _vacio

    por_usuario = []
    for f in filas:
        dias = float(f["dias_inactivo"] or 0.0)
        sesiones = int(f["total_sesiones"] or 1)
        por_usuario.append({
            "user_id":              f["user_id"],
            "total_sesiones":       sesiones,
            "duracion_media_min":   float(f["dur_media_min"] or 0.0),
            "sesion_mas_larga_min": float(f["dur_max_min"] or 0.0),
            "primera_sesion":       f["primera_sesion"] or "—",
            "ultima_sesion":        f["ultima_sesion"] or "—",
            "dias_inactivo":        dias,
            "es_ocasional":         sesiones <= 2 and dias > 7,
        })

    activos_7d  = sum(1 for u in por_usuario if u["dias_inactivo"] <= 7)
    activos_30d = sum(1 for u in por_usuario if u["dias_inactivo"] <= 30)
    total_ses   = sum(u["total_sesiones"] for u in por_usuario)
    if total_ses > 0:
        dur_media_global = round(
            sum(u["duracion_media_min"] * u["total_sesiones"] for u in por_usuario) / total_ses, 1
        )
    else:
        dur_media_global = 0.0
    dur_max_global = max((u["sesion_mas_larga_min"] for u in por_usuario), default=0.0)

    return {
        "por_usuario": por_usuario,
        "resumen": {
            "activos_7d":               activos_7d,
            "activos_30d":              activos_30d,
            "duracion_media_global_min": dur_media_global,
            "sesion_mas_larga_min":     dur_max_global,
            "total_sesiones":           total_ses,
        },
    }
