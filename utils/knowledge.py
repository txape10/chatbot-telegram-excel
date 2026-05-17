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
    """Lee todos los archivos .md de knowledge/ y devuelve su contenido unificado."""
    bloques = []

    # Primero los archivos en el orden definido
    for nombre in _ORDEN:
        ruta = os.path.join(DIRECTORIO_KNOWLEDGE, nombre)
        if os.path.exists(ruta):
            contenido = _leer_archivo(ruta)
            if contenido:
                bloques.append(contenido)

    # Luego cualquier .md que no esté en la lista (por si se añaden nuevos)
    nombres_cargados = set(_ORDEN)
    for nombre in sorted(os.listdir(DIRECTORIO_KNOWLEDGE)):
        if nombre.endswith(".md") and nombre not in nombres_cargados:
            ruta = os.path.join(DIRECTORIO_KNOWLEDGE, nombre)
            contenido = _leer_archivo(ruta)
            if contenido:
                bloques.append(contenido)

    if bloques:
        logger.info("Base de conocimiento cargada: %d archivos", len(bloques))
    else:
        logger.warning("No se encontró contenido en knowledge/")

    return "\n\n---\n\n".join(bloques)


def _leer_archivo(ruta: str) -> str:
    try:
        with open(ruta, encoding="utf-8") as f:
            return f.read().strip()
    except Exception as error:
        logger.error("Error leyendo %s: %s", ruta, error)
        return ""
