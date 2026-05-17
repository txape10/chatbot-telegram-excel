import google.generativeai as genai
from config import GEMINI_API_KEY, SYSTEM_PROMPT

genai.configure(api_key=GEMINI_API_KEY)

_modelo = genai.GenerativeModel(
    model_name="gemini-2.0-flash",
    system_instruction=SYSTEM_PROMPT,
)


def obtener_respuesta(historial: list[dict], pregunta: str) -> str:
    chat = _modelo.start_chat(history=historial)
    respuesta = chat.send_message(pregunta)
    return respuesta.text
