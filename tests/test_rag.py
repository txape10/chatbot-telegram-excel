"""Tests para utils/rag.py — almacenamiento y recuperación de feedback."""
import pytest
from unittest.mock import patch, MagicMock
from utils.rag import guardar_ejemplo, obtener_ejemplos


def _conn_mock():
    """Conexión SQLite en memoria para tests aislados."""
    import sqlite3
    conn = sqlite3.connect(":memory:")
    conn.execute("""
        CREATE TABLE feedback_rag (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id   INTEGER NOT NULL,
            pregunta  TEXT    NOT NULL,
            respuesta TEXT    NOT NULL,
            tipo      TEXT    NOT NULL DEFAULT 'positivo',
            creado_en DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_rag_user ON feedback_rag(user_id, tipo)"
    )
    return conn


@pytest.fixture(autouse=True)
def _patch_db(monkeypatch):
    """Reemplaza conectar() por una conexión en memoria para cada test."""
    from contextlib import contextmanager
    conn = _conn_mock()

    @contextmanager
    def _fake_conectar():
        yield conn
        conn.commit()

    monkeypatch.setattr("utils.rag._INIT_HECHO", False)
    monkeypatch.setattr("utils.rag._db_conectar", _fake_conectar)
    yield conn


# ── guardar_ejemplo ───────────────────────────────────────────────────────────

def test_guardar_positivo_se_almacena(_patch_db):
    guardar_ejemplo(1, "¿Cómo sumo?", "Usa =SUMA()", tipo="positivo")
    cur = _patch_db.execute(
        "SELECT tipo FROM feedback_rag WHERE user_id = 1"
    )
    row = cur.fetchone()
    assert row is not None
    assert row[0] == "positivo"


def test_guardar_negativo_se_almacena(_patch_db):
    guardar_ejemplo(1, "¿Cómo filtro?", "Usa FILTRAR()", tipo="negativo")
    cur = _patch_db.execute(
        "SELECT tipo FROM feedback_rag WHERE user_id = 1"
    )
    row = cur.fetchone()
    assert row is not None
    assert row[0] == "negativo"


# ── obtener_ejemplos ──────────────────────────────────────────────────────────

def test_obtener_solo_devuelve_positivos(_patch_db):
    guardar_ejemplo(2, "Pregunta A", "Respuesta A", tipo="positivo")
    guardar_ejemplo(2, "Pregunta B", "Respuesta B", tipo="negativo")
    guardar_ejemplo(2, "Pregunta C", "Respuesta C", tipo="positivo")

    ejemplos = obtener_ejemplos(2, limite=10)
    assert len(ejemplos) == 2
    preguntas = [e["pregunta"] for e in ejemplos]
    assert "Pregunta A" in preguntas
    assert "Pregunta C" in preguntas
    assert "Pregunta B" not in preguntas


def test_feedback_negativo_no_contamina_few_shot(_patch_db):
    """El feedback negativo no debe inyectarse como few-shot."""
    for i in range(5):
        guardar_ejemplo(3, f"Mala pregunta {i}", f"Mala respuesta {i}", tipo="negativo")

    ejemplos = obtener_ejemplos(3, limite=10)
    assert ejemplos == []


def test_obtener_respeta_limite(_patch_db):
    for i in range(10):
        guardar_ejemplo(4, f"Pregunta {i}", f"Respuesta {i}", tipo="positivo")

    ejemplos = obtener_ejemplos(4, limite=3)
    assert len(ejemplos) == 3


def test_obtener_usuario_sin_feedback_devuelve_lista_vacia(_patch_db):
    ejemplos = obtener_ejemplos(999, limite=3)
    assert ejemplos == []


def test_obtener_devuelve_claves_correctas(_patch_db):
    guardar_ejemplo(5, "¿Qué es Excel?", "Una hoja de cálculo", tipo="positivo")
    ejemplos = obtener_ejemplos(5)
    assert "pregunta" in ejemplos[0]
    assert "respuesta" in ejemplos[0]
