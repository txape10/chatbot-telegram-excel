from groq import Groq
from config import GROQ_API_KEY, SYSTEM_PROMPT

_cliente = Groq(api_key=GROQ_API_KEY)
MODELO = "llama-3.3-70b-versatile"


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
