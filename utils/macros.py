"""Almacenamiento de macros personales por usuario en SQLite.

Una macro es una secuencia de operaciones DSL con nombre que el usuario
puede guardar y reutilizar. Se describe en lenguaje natural y el LLM
la convierte a una lista de operaciones JSON.
"""
import json
import sqlite3
from utils.db import conectar as _db_conectar


def _conectar() -> sqlite3.Connection:
    conn = _db_conectar()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_macros (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id    INTEGER NOT NULL,
            nombre     TEXT    NOT NULL,
            descripcion TEXT,
            operaciones TEXT   NOT NULL,
            creado_en  DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, nombre)
        )
    """)
    conn.commit()
    return conn


def guardar_macro(user_id: int, nombre: str,
                  operaciones: list[dict], descripcion: str = "") -> None:
    """Guarda o sobreescribe una macro del usuario."""
    with _conectar() as conn:
        conn.execute("""
            INSERT INTO user_macros (user_id, nombre, descripcion, operaciones)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(user_id, nombre) DO UPDATE
            SET descripcion = excluded.descripcion,
                operaciones = excluded.operaciones,
                creado_en   = CURRENT_TIMESTAMP
        """, (user_id, nombre.lower(), descripcion, json.dumps(operaciones, ensure_ascii=False)))
        conn.commit()


def obtener_macro(user_id: int, nombre: str) -> list[dict] | None:
    """Devuelve las operaciones de la macro o None si no existe."""
    with _conectar() as conn:
        fila = conn.execute(
            "SELECT operaciones FROM user_macros WHERE user_id = ? AND nombre = ?",
            (user_id, nombre.lower()),
        ).fetchone()
    return json.loads(fila[0]) if fila else None


def listar_macros(user_id: int) -> list[dict]:
    """Devuelve la lista de macros del usuario [{nombre, descripcion, creado_en}]."""
    with _conectar() as conn:
        filas = conn.execute(
            "SELECT nombre, descripcion, creado_en FROM user_macros "
            "WHERE user_id = ? ORDER BY nombre",
            (user_id,),
        ).fetchall()
    return [{"nombre": f[0], "descripcion": f[1], "creado_en": f[2]} for f in filas]


def borrar_macro(user_id: int, nombre: str) -> bool:
    """Elimina una macro. Devuelve True si existía."""
    with _conectar() as conn:
        cur = conn.execute(
            "DELETE FROM user_macros WHERE user_id = ? AND nombre = ?",
            (user_id, nombre.lower()),
        )
        conn.commit()
    return cur.rowcount > 0
