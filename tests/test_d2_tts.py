"""Tests Sprint D2 — limpieza de Markdown para TTS.

No hace llamadas a edge-tts ni a ninguna API externa.
Solo prueba _limpiar_markdown, que es una función pura.
"""
import pytest
from services.tts import _limpiar_markdown


# ── Negrita e itálica ─────────────────────────────────────────────────────────

def test_negrita_doble_asterisco():
    resultado = _limpiar_markdown("Hola **mundo**")
    assert resultado == "Hola mundo"


def test_negrita_guion_bajo():
    resultado = _limpiar_markdown("Texto __importante__")
    assert resultado == "Texto importante"


def test_italica_un_asterisco():
    resultado = _limpiar_markdown("Texto *cursiva* aquí")
    assert resultado == "Texto cursiva aquí"


def test_italica_guion_bajo():
    resultado = _limpiar_markdown("Texto _cursiva_ aquí")
    assert resultado == "Texto cursiva aquí"


# ── Código ────────────────────────────────────────────────────────────────────

def test_codigo_inline():
    resultado = _limpiar_markdown("Usa `BUSCARV` para buscar")
    assert resultado == "Usa BUSCARV para buscar"


def test_bloque_codigo_eliminado():
    texto = "Mira esto:\n```python\nx = 1\n```\nListo."
    resultado = _limpiar_markdown(texto)
    assert "```" not in resultado
    assert "x = 1" not in resultado
    assert "Mira esto:" in resultado


# ── Cabeceras ─────────────────────────────────────────────────────────────────

def test_cabecera_h1():
    resultado = _limpiar_markdown("# Título principal")
    assert resultado == "Título principal"


def test_cabecera_h2():
    resultado = _limpiar_markdown("## Sección")
    assert resultado == "Sección"


def test_cabecera_h3():
    resultado = _limpiar_markdown("### Subsección")
    assert resultado == "Subsección"


# ── Viñetas ───────────────────────────────────────────────────────────────────

def test_vineta_guion():
    resultado = _limpiar_markdown("- Primer elemento")
    assert resultado == "Primer elemento"


def test_vineta_asterisco():
    resultado = _limpiar_markdown("* Segundo elemento")
    assert resultado == "Segundo elemento"


def test_vineta_punto():
    resultado = _limpiar_markdown("• Tercer elemento")
    assert resultado == "Tercer elemento"


# ── Tablas ────────────────────────────────────────────────────────────────────

def test_tabla_eliminada():
    texto = "| Col1 | Col2 |\n| A    | B    |"
    resultado = _limpiar_markdown(texto)
    assert "|" not in resultado


# ── Texto normal sin cambios ──────────────────────────────────────────────────

def test_texto_plano_sin_cambios():
    texto = "Esto es texto normal sin marcado."
    assert _limpiar_markdown(texto) == texto


def test_numeros_sin_cambios():
    texto = "El valor es 3.14 y el total es 100"
    assert _limpiar_markdown(texto) == texto


# ── Casos extremos ────────────────────────────────────────────────────────────

def test_cadena_vacia():
    assert _limpiar_markdown("") == ""


def test_texto_mixto():
    texto = "## Resultado\n- Usa **BUSCARV** con `=BUSCARV(A1,...)`"
    resultado = _limpiar_markdown(texto)
    assert "##" not in resultado
    assert "**" not in resultado
    assert "`" not in resultado
    assert "BUSCARV" in resultado


def test_saltos_multiples_reducidos():
    texto = "Línea 1\n\n\n\nLínea 2"
    resultado = _limpiar_markdown(texto)
    assert "\n\n\n" not in resultado
