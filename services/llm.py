import base64
import logging
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
