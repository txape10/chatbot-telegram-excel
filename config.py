import os
from dotenv import load_dotenv
from utils.knowledge import cargar_base_conocimiento

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
AUTHORIZED_USERS = set(
    int(uid.strip())
    for uid in os.getenv("AUTHORIZED_USERS", "").split(",")
    if uid.strip()
)

HISTORIAL_MAX_MENSAJES = 10

_BASE_CONOCIMIENTO = cargar_base_conocimiento()

SYSTEM_PROMPT = """Eres un experto en Microsoft Excel con más de 20 años de experiencia.
Ante cada pregunta debes:
1. Dar una explicación clara y concisa
2. Incluir un ejemplo práctico con datos reales
3. Mostrar la fórmula o los pasos exactos
4. Añadir consejos o variantes útiles si los hay

Responde siempre en español.

A continuación tienes tu base de conocimiento de referencia. Úsala para dar respuestas precisas, con el tono y formato indicados:

""" + _BASE_CONOCIMIENTO
