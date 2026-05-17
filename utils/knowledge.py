import os
import logging

logger = logging.getLogger(__name__)

DIRECTORIO_KNOWLEDGE = os.path.join(os.path.dirname(__file__), "..", "knowledge")

# Orden de carga: primero el tono/estilo, luego el contenido técnico
_ORDEN = [
    "ejemplos_respuestas.md",
    "formulas_basicas.md",
    "formulas_avanzadas.md",
    "tablas_dinamicas.md",
    "formato_condicional.md",
    "graficos.md",
    "power_query.md",
    "vba_basico.md",
    "errores_comunes.md",
]


def cargar_base_conocimiento() -> str:
    """Carga solo la guía de tono/estilo para el system prompt.
    El contenido técnico no se inyecta en el prompt base para no superar
    el límite de tokens del tier gratuito de Groq (~12 000 TPM)."""
    contenido = _leer_archivo(
        os.path.join(DIRECTORIO_KNOWLEDGE, "ejemplos_respuestas.md")
    )
    if contenido:
        logger.info("Guía de estilo cargada en system prompt (%d chars)", len(contenido))
    else:
        logger.warning("No se encontró ejemplos_respuestas.md en knowledge/")
    return contenido


def cargar_conocimiento_completo() -> str:
    """Carga todos los archivos .md. Útil para referencia o RAG futuro."""
    bloques = []
    for nombre in _ORDEN:
        ruta = os.path.join(DIRECTORIO_KNOWLEDGE, nombre)
        if os.path.exists(ruta):
            contenido = _leer_archivo(ruta)
            if contenido:
                bloques.append(contenido)
    nombres_cargados = set(_ORDEN)
    for nombre in sorted(os.listdir(DIRECTORIO_KNOWLEDGE)):
        if nombre.endswith(".md") and nombre not in nombres_cargados:
            ruta = os.path.join(DIRECTORIO_KNOWLEDGE, nombre)
            contenido = _leer_archivo(ruta)
            if contenido:
                bloques.append(contenido)
    logger.info("Conocimiento completo cargado: %d archivos", len(bloques))
    return "\n\n---\n\n".join(bloques)


def _leer_archivo(ruta: str) -> str:
    try:
        with open(ruta, encoding="utf-8") as f:
            return f.read().strip()
    except Exception as error:
        logger.error("Error leyendo %s: %s", ruta, error)
        return ""
