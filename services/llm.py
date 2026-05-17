import base64
import json
import logging
import pandas as pd
from groq import Groq
from config import GROQ_API_KEY, SYSTEM_PROMPT

logger = logging.getLogger(__name__)

_cliente = Groq(api_key=GROQ_API_KEY, timeout=60.0)
MODELO        = "llama-3.3-70b-versatile"
MODELO_VISION = "meta-llama/llama-4-scout-17b-16e-instruct"

# Margen de seguridad bajo el límite de 12 000 TPM del tier gratuito de Groq
_MAX_TOKENS_PETICION = 9_000


def _estimar_tokens(texto: str) -> int:
    """Estimación rápida: ~4 caracteres por token."""
    return len(texto) // 4


def _construir_mensajes(historial: list[dict], pregunta: str) -> list[dict]:
    """Construye la lista de mensajes ajustando el historial si la petición
    supera el presupuesto de tokens."""
    mensajes_sistema = [{"role": "system", "content": SYSTEM_PROMPT}]
    tokens_fijos = _estimar_tokens(SYSTEM_PROMPT) + _estimar_tokens(pregunta)
    presupuesto_historial = _MAX_TOKENS_PETICION - tokens_fijos

    # Recortar historial por el principio (mensajes más antiguos) hasta que quepa
    historial_valido = list(historial)
    while historial_valido and _estimar_tokens(str(historial_valido)) > presupuesto_historial:
        historial_valido = historial_valido[2:]   # eliminar par user+model más antiguo
        logger.debug("Historial recortado a %d mensajes por límite de tokens", len(historial_valido))

    mensajes_historial = []
    for mensaje in historial_valido:
        rol = "assistant" if mensaje["role"] == "model" else "user"
        mensajes_historial.append({"role": rol, "content": mensaje["parts"][0]})

    return mensajes_sistema + mensajes_historial + [{"role": "user", "content": pregunta}]


def obtener_respuesta(historial: list[dict], pregunta: str) -> str:
    mensajes = _construir_mensajes(historial, pregunta)
    tokens_estimados = _estimar_tokens(str(mensajes))
    logger.info("Petición LLM — tokens estimados: %d", tokens_estimados)

    respuesta = _cliente.chat.completions.create(
        model=MODELO,
        messages=mensajes,
    )
    return respuesta.choices[0].message.content


def extraer_operacion_edicion(df: pd.DataFrame, pregunta: str) -> dict | None:
    """Llama al LLM con el prompt de edición y parsea el JSON.

    Devuelve el dict de operación o None si es RESPUESTA_LIBRE o falla el parseo.
    """
    from prompts.excel import EDITOR_DSL_SISTEMA, EDITOR_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {df[c].dtype}" for c in df.columns)
    muestra  = df.head(3).to_string(index=False)

    mensaje_usuario = EDITOR_DSL_USUARIO.format(
        columnas=columnas,
        tipos=tipos,
        muestra=muestra,
        pregunta=pregunta,
    )

    try:
        respuesta = _cliente.chat.completions.create(
            model=MODELO,
            messages=[
                {"role": "system", "content": EDITOR_DSL_SISTEMA},
                {"role": "user",   "content": mensaje_usuario},
            ],
            temperature=0,
            max_tokens=400,
        )
        texto = respuesta.choices[0].message.content.strip()
        logger.debug("Respuesta editor DSL del LLM: %s", texto)

        if texto == "RESPUESTA_LIBRE":
            return None

        if "```" in texto:
            lineas = [l for l in texto.splitlines() if not l.startswith("```")]
            texto = "\n".join(lineas).strip()

        return json.loads(texto)

    except Exception as error:
        logger.warning("Error extrayendo operación de edición: %s", error)
        return None


def extraer_query_dsl(df: pd.DataFrame, pregunta: str) -> dict | None:
    """Llama al LLM con el prompt DSL y parsea el JSON resultante.

    Devuelve el dict de query si la pregunta es una consulta de datos,
    o None si el LLM responde RESPUESTA_LIBRE o si falla el parseo.
    """
    from prompts.excel import QUERY_DSL_SISTEMA, QUERY_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {df[c].dtype}" for c in df.columns)
    muestra  = df.head(3).to_string(index=False)

    mensaje_usuario = QUERY_DSL_USUARIO.format(
        columnas=columnas,
        tipos=tipos,
        muestra=muestra,
        pregunta=pregunta,
    )

    try:
        respuesta = _cliente.chat.completions.create(
            model=MODELO,
            messages=[
                {"role": "system", "content": QUERY_DSL_SISTEMA},
                {"role": "user",   "content": mensaje_usuario},
            ],
            temperature=0,
            max_tokens=300,
        )
        texto = respuesta.choices[0].message.content.strip()
        logger.debug("Respuesta DSL del LLM: %s", texto)

        if texto == "RESPUESTA_LIBRE":
            return None

        # Limpiar posibles bloques de código markdown que el LLM añada por error
        if "```" in texto:
            lineas = [l for l in texto.splitlines() if not l.startswith("```")]
            texto = "\n".join(lineas).strip()

        return json.loads(texto)

    except Exception as error:
        logger.warning("Error extrayendo query DSL: %s", error)
        return None


def analizar_imagen(imagen_bytes: bytes, pregunta: str = "") -> str:
    """Analiza una captura de pantalla de Excel usando un modelo con visión."""
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

    respuesta = _cliente.chat.completions.create(
        model=MODELO_VISION,
        messages=mensajes,
    )
    return respuesta.choices[0].message.content
