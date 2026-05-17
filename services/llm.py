import base64
from groq import Groq
from config import GROQ_API_KEY, SYSTEM_PROMPT

_cliente = Groq(api_key=GROQ_API_KEY, timeout=60.0)
MODELO = "llama-3.3-70b-versatile"
MODELO_VISION = "meta-llama/llama-4-scout-17b-16e-instruct"


def obtener_respuesta(historial: list[dict], pregunta: str) -> str:
    mensajes = [{"role": "system", "content": SYSTEM_PROMPT}]

    for mensaje in historial:
        rol = "assistant" if mensaje["role"] == "model" else "user"
        texto = mensaje["parts"][0]
        mensajes.append({"role": rol, "content": texto})

    mensajes.append({"role": "user", "content": pregunta})

    respuesta = _cliente.chat.completions.create(
        model=MODELO,
        messages=mensajes,
    )
    return respuesta.choices[0].message.content


def analizar_imagen(imagen_bytes: bytes, pregunta: str = "") -> str:
    """Analiza una captura de pantalla de Excel usando un modelo con visión."""
    imagen_b64 = base64.standard_b64encode(imagen_bytes).decode("utf-8")

    texto_usuario = pregunta if pregunta else "Analiza esta captura de Excel y explica qué hace, qué fórmulas usa y cómo podría mejorarla."

    mensajes = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": [
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{imagen_b64}"},
                },
                {"type": "text", "text": texto_usuario},
            ],
        },
    ]

    respuesta = _cliente.chat.completions.create(
        model=MODELO_VISION,
        messages=mensajes,
    )
    return respuesta.choices[0].message.content
