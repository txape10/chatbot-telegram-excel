"""Registro y consulta de estadísticas de llamadas al LLM.

Cada llamada al proveedor de IA queda anotada en la tabla llm_stats:
  - proveedor : nombre del proveedor que respondió ('groq', 'mistral', …)
  - tipo      : 'chat' | 'audio' | 'vision'
  - ok        : 1 éxito / 0 error
  - fallback  : 1 si fue el proveedor secundario quien respondió
  - error_tipo: 'rate_limit' | 'timeout' | 'conexion' | 'auth' | 'limite' | 'generico'

El módulo nunca lanza excepciones — si la BD falla, los stats se pierden
silenciosamente sin afectar al funcionamiento del bot.
"""
from utils.db import conectar as _db_conectar

_CREATE = """
CREATE TABLE IF NOT EXISTS llm_stats (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    creado_en  TEXT    NOT NULL DEFAULT (datetime('now')),
    proveedor  TEXT    NOT NULL,
    tipo       TEXT    NOT NULL DEFAULT 'chat',
    ok         INTEGER NOT NULL,
    fallback   INTEGER NOT NULL DEFAULT 0,
    error_tipo TEXT
)
"""


def _init(conn) -> None:
    conn.execute(_CREATE)


def registrar(proveedor: str, tipo: str, ok: bool,
              fallback: bool = False, error_tipo: str | None = None) -> None:
    """Registra una llamada al LLM. Falla silenciosamente."""
    try:
        with _db_conectar() as conn:
            _init(conn)
            conn.execute(
                "INSERT INTO llm_stats (proveedor, tipo, ok, fallback, error_tipo)"
                " VALUES (?, ?, ?, ?, ?)",
                (proveedor, tipo, int(ok), int(fallback), error_tipo),
            )
    except Exception:
        pass


def obtener_stats_ia() -> dict:
    """Devuelve estadísticas agregadas para el panel de administración."""
    _vacio = {
        "total": 0, "exitosas": 0, "fallbacks": 0, "errores": 0,
        "total_hoy": 0, "tasa_exito": 0.0,
        "errores_por_tipo": [], "por_proveedor": [], "por_dia": [],
    }
    try:
        with _db_conectar() as conn:
            _init(conn)

            fila = conn.execute("""
                SELECT
                    COUNT(*)                                          AS total,
                    COALESCE(SUM(ok), 0)                             AS exitosas,
                    COALESCE(SUM(fallback), 0)                       AS fallbacks,
                    COALESCE(SUM(CASE WHEN ok=0 THEN 1 ELSE 0 END), 0) AS errores
                FROM llm_stats
            """).fetchone()
            total, exitosas, fallbacks, errores = fila if fila else (0, 0, 0, 0)

            total_hoy = conn.execute(
                "SELECT COUNT(*) FROM llm_stats WHERE date(creado_en) = date('now')"
            ).fetchone()[0] or 0

            errores_por_tipo = [
                {"tipo": r[0], "n": r[1]}
                for r in conn.execute("""
                    SELECT error_tipo, COUNT(*) AS n
                    FROM llm_stats
                    WHERE ok=0 AND error_tipo IS NOT NULL
                    GROUP BY error_tipo
                    ORDER BY n DESC
                """).fetchall()
            ]

            por_proveedor = [
                {
                    "proveedor": r[0],
                    "total":    r[1],
                    "exitosas": r[2],
                    "errores":  r[1] - r[2],
                    "fallbacks": r[3],
                    "pct_ok":   round(r[2] / r[1] * 100, 1) if r[1] else 0,
                }
                for r in conn.execute("""
                    SELECT proveedor,
                           COUNT(*)         AS total,
                           COALESCE(SUM(ok), 0)      AS exitosas,
                           COALESCE(SUM(fallback), 0) AS fallbacks
                    FROM llm_stats
                    GROUP BY proveedor
                    ORDER BY total DESC
                """).fetchall()
            ]

            por_dia = [
                {"dia": r[0], "total": r[1], "exitosas": r[2]}
                for r in conn.execute("""
                    SELECT date(creado_en) AS dia,
                           COUNT(*)         AS total,
                           COALESCE(SUM(ok), 0) AS exitosas
                    FROM llm_stats
                    WHERE creado_en >= date('now', '-6 days')
                    GROUP BY dia
                    ORDER BY dia
                """).fetchall()
            ]

        return {
            "total":            total or 0,
            "exitosas":         exitosas or 0,
            "fallbacks":        fallbacks or 0,
            "errores":          errores or 0,
            "total_hoy":        total_hoy,
            "tasa_exito":       round((exitosas or 0) / total * 100, 1) if total else 0.0,
            "errores_por_tipo": errores_por_tipo,
            "por_proveedor":    por_proveedor,
            "por_dia":          por_dia,
        }
    except Exception:
        return _vacio
