import base64
import json
import logging
import pandas as pd

from config import SYSTEM_PROMPT
from services.llm_provider import (  # noqa: F401 — re-exportados
    LLMError, LLMProvider, obtener_proveedor, obtener_proveedor_privado,
)

__all__ = ["LLMError", "obtener_proveedor_privado"]

logger = logging.getLogger(__name__)


def _estimar_tokens(texto: str) -> int:
    """Estimación rápida: ~4 caracteres por token."""
    return len(texto) // 4


def _construir_mensajes(historial: list[dict], pregunta: str,
                         proveedor: LLMProvider | None = None,
                         system_override: str | None = None) -> list[dict]:
    """Construye la lista de mensajes ajustando el historial si la petición
    supera el presupuesto de tokens del proveedor activo."""
    proveedor = proveedor or obtener_proveedor()
    prompt_sistema = system_override or SYSTEM_PROMPT
    mensajes_sistema = [{"role": "system", "content": prompt_sistema}]
    tokens_fijos = _estimar_tokens(prompt_sistema) + _estimar_tokens(pregunta)
    presupuesto_historial = proveedor.max_tokens_peticion - tokens_fijos

    historial_valido = list(historial)
    while historial_valido and _estimar_tokens(str(historial_valido)) > presupuesto_historial:
        historial_valido = historial_valido[2:]
        logger.debug("Historial recortado a %d mensajes por límite de tokens", len(historial_valido))

    mensajes_historial = []
    for mensaje in historial_valido:
        rol = "assistant" if mensaje["role"] == "model" else "user"
        mensajes_historial.append({"role": rol, "content": mensaje["parts"][0]})

    return mensajes_sistema + mensajes_historial + [{"role": "user", "content": pregunta}]


def obtener_respuesta(historial: list[dict], pregunta: str,
                       proveedor: LLMProvider | None = None,
                       system_override: str | None = None) -> str:
    """Envía la pregunta al LLM y devuelve la respuesta.

    Args:
        proveedor: si se indica, usa ese proveedor en lugar del activo.
        system_override: reemplaza el system prompt por defecto.
    """
    proveedor_activo = proveedor or obtener_proveedor()
    mensajes = _construir_mensajes(historial, pregunta, proveedor_activo, system_override)
    tokens_estimados = _estimar_tokens(str(mensajes))
    logger.info("Petición LLM — tokens estimados: %d", tokens_estimados)
    return proveedor_activo.chat(mensajes)


def _limpiar_json(texto: str) -> str:
    """Elimina bloques de código markdown que el LLM añade por error."""
    if "```" in texto:
        lineas = [l for l in texto.splitlines() if not l.startswith("```")]
        texto = "\n".join(lineas).strip()
    return texto


def extraer_estructura_excel(pregunta: str) -> dict | None:
    """Interpreta la descripción del usuario y devuelve la estructura JSON para crear el xlsx."""
    from prompts.excel import CREAR_EXCEL_SISTEMA, CREAR_EXCEL_USUARIO
    try:
        texto = obtener_proveedor().chat(
            messages=[
                {"role": "system", "content": CREAR_EXCEL_SISTEMA},
                {"role": "user",   "content": CREAR_EXCEL_USUARIO.format(pregunta=pregunta)},
            ],
            temperature=0,
            max_tokens=800,
        )
        logger.debug("Estructura Excel del LLM: %s", texto)
        return json.loads(_limpiar_json(texto.strip()))
    except Exception as error:
        logger.warning("Error extrayendo estructura Excel: %s", error)
        return None


def extraer_operacion_edicion(df: pd.DataFrame, pregunta: str) -> dict | None:
    """Llama al LLM con el prompt de edición y parsea el JSON.

    Devuelve:
      - dict con la operación DSL     → aplicar edición
      - {"aclaracion_necesaria": True, "pregunta": ..., "opciones": [...]}  → pedir aclaración
      - None                           → RESPUESTA_LIBRE o error de parseo
    """
    from prompts.excel import EDITOR_DSL_SISTEMA, EDITOR_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {df[c].dtype}" for c in df.columns)
    muestra  = df.head(3).to_string(index=False)

    try:
        texto = obtener_proveedor().chat(
            messages=[
                {"role": "system", "content": EDITOR_DSL_SISTEMA},
                {"role": "user",   "content": EDITOR_DSL_USUARIO.format(
                    columnas=columnas, tipos=tipos, muestra=muestra, pregunta=pregunta,
                )},
            ],
            temperature=0,
            max_tokens=500,
        )
        logger.debug("Respuesta editor DSL del LLM: %s", texto)
        texto = texto.strip()
        if texto == "RESPUESTA_LIBRE":
            return None
        parsed = json.loads(_limpiar_json(texto))
        # Propagar aclaración tal cual — el handler decide qué hacer
        return parsed
    except Exception as error:
        logger.warning("Error extrayendo operación de edición: %s", error)
        return None


def extraer_query_dsl(df: pd.DataFrame, pregunta: str) -> dict | None:
    """Llama al LLM con el prompt DSL y parsea el JSON resultante.

    Devuelve:
      - dict con la query DSL          → ejecutar consulta
      - {"aclaracion_necesaria": True, "pregunta": ..., "opciones": [...]}  → pedir aclaración
      - None                           → RESPUESTA_LIBRE o error de parseo
    """
    from prompts.excel import QUERY_DSL_SISTEMA, QUERY_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {df[c].dtype}" for c in df.columns)
    muestra  = df.head(3).to_string(index=False)

    try:
        texto = obtener_proveedor().chat(
            messages=[
                {"role": "system", "content": QUERY_DSL_SISTEMA},
                {"role": "user",   "content": QUERY_DSL_USUARIO.format(
                    columnas=columnas, tipos=tipos, muestra=muestra, pregunta=pregunta,
                )},
            ],
            temperature=0,
            max_tokens=400,
        )
        logger.debug("Respuesta DSL del LLM: %s", texto)
        texto = texto.strip()
        if texto == "RESPUESTA_LIBRE":
            return None
        parsed = json.loads(_limpiar_json(texto))
        # Propagar aclaración tal cual — el handler decide qué hacer
        return parsed
    except Exception as error:
        logger.warning("Error extrayendo query DSL: %s", error)
        return None


def extraer_operacion_combinar(df1: pd.DataFrame, df2: pd.DataFrame,
                               pregunta: str) -> dict:
    """Extrae la columna clave y tipo de join para combinar dos DataFrames."""
    from prompts.excel import COMBINAR_DSL_SISTEMA, COMBINAR_DSL_USUARIO

    cols_a       = ", ".join(f"'{c}'" for c in df1.columns)
    cols_b       = ", ".join(f"'{c}'" for c in df2.columns)
    cols_comunes = ", ".join(f"'{c}'" for c in df1.columns if c in df2.columns) or "ninguna"

    try:
        texto = obtener_proveedor().chat(
            messages=[
                {"role": "system", "content": COMBINAR_DSL_SISTEMA},
                {"role": "user",   "content": COMBINAR_DSL_USUARIO.format(
                    cols_a=cols_a, cols_b=cols_b, cols_comunes=cols_comunes, pregunta=pregunta,
                )},
            ],
            temperature=0,
            max_tokens=100,
        )
        logger.debug("Respuesta combinar DSL del LLM: %s", texto)
        return json.loads(_limpiar_json(texto.strip()))
    except Exception as error:
        logger.warning("Error extrayendo operación combinar: %s", error)
        return {"col": None, "como": "inner"}


def extraer_peticion_grafico(df: pd.DataFrame, pregunta: str) -> dict | None:
    """Extrae los parámetros del gráfico pedido en lenguaje natural."""
    from prompts.excel import GRAFICO_DSL_SISTEMA, GRAFICO_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {df[c].dtype}" for c in df.columns)

    try:
        texto = obtener_proveedor().chat(
            messages=[
                {"role": "system", "content": GRAFICO_DSL_SISTEMA},
                {"role": "user",   "content": GRAFICO_DSL_USUARIO.format(
                    columnas=columnas, tipos=tipos, pregunta=pregunta,
                )},
            ],
            temperature=0,
            max_tokens=150,
        )
        logger.debug("Respuesta gráfico DSL del LLM: %s", texto)
        return json.loads(_limpiar_json(texto.strip()))
    except Exception as error:
        logger.warning("Error extrayendo petición de gráfico: %s", error)
        return None


def extraer_operaciones_macro(descripcion: str) -> list[dict] | None:
    """Convierte una descripción de macro en una lista de operaciones DSL."""
    from prompts.excel import MACRO_DSL_SISTEMA, MACRO_DSL_USUARIO

    try:
        texto = obtener_proveedor().chat(
            messages=[
                {"role": "system", "content": MACRO_DSL_SISTEMA},
                {"role": "user",   "content": MACRO_DSL_USUARIO.format(descripcion=descripcion)},
            ],
            temperature=0,
            max_tokens=400,
        )
        logger.debug("Operaciones macro del LLM: %s", texto)
        ops = json.loads(_limpiar_json(texto.strip()))
        return ops if isinstance(ops, list) else None
    except Exception as error:
        logger.warning("Error extrayendo operaciones de macro: %s", error)
        return None


def extraer_regla_formato(df: pd.DataFrame, instruccion: str) -> dict | None:
    """Extrae la regla de formato condicional desde lenguaje natural.

    Devuelve un dict con la regla DSL o None si no se pudo interpretar.
    """
    from prompts.excel import FORMATO_DSL_SISTEMA, FORMATO_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {df[c].dtype}" for c in df.columns)
    muestra  = df.head(3).to_string(index=False)

    try:
        texto = obtener_proveedor().chat(
            messages=[
                {"role": "system", "content": FORMATO_DSL_SISTEMA},
                {"role": "user",   "content": FORMATO_DSL_USUARIO.format(
                    columnas=columnas, tipos=tipos, muestra=muestra, pregunta=instruccion,
                )},
            ],
            temperature=0,
            max_tokens=200,
        )
        logger.debug("Regla formato del LLM: %s", texto)
        return json.loads(_limpiar_json(texto.strip()))
    except Exception as error:
        logger.warning("Error extrayendo regla de formato: %s", error)
        return None


def transcribir_audio(audio_bytes: bytes, filename: str = "audio.ogg") -> str:
    """Transcribe un mensaje de voz usando el proveedor activo."""
    return obtener_proveedor().transcribir(audio_bytes, filename)


def analizar_imagen(imagen_bytes: bytes, pregunta: str = "") -> str:
    """Analiza una captura de pantalla de Excel usando visión IA."""
    imagen_b64 = base64.standard_b64encode(imagen_bytes).decode("utf-8")
    texto_usuario = (
        pregunta if pregunta
        else "Analiza esta captura de Excel y explica qué hace, qué fórmulas usa y cómo podría mejorarla."
    )
    mensajes = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{imagen_b64}"}},
                {"type": "text", "text": texto_usuario},
            ],
        },
    ]
    return obtener_proveedor().vision(mensajes)
