import os
from dotenv import load_dotenv
from utils.knowledge import cargar_base_conocimiento
from prompts.excel import SYSTEM_BASE

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY   = os.getenv("GROQ_API_KEY")   # usado por GroqProvider
AUTHORIZED_USERS = set(
    int(uid.strip())
    for uid in os.getenv("AUTHORIZED_USERS", "").split(",")
    if uid.strip()
)

# ── Módulos activables ────────────────────────────────────────────────────────
# Permite instalar solo el bot, solo el Add-in o ambos
ENABLE_TELEGRAM = os.getenv("ENABLE_TELEGRAM", "true").lower() == "true"
ENABLE_ADDIN    = os.getenv("ENABLE_ADDIN",    "true").lower() == "true"

# ── Proveedor de IA ───────────────────────────────────────────────────────────
# groq | ollama | openai  (ver services/llm_provider.py)
LLM_PROVIDER = os.getenv("LLM_PROVIDER", "groq")

HISTORIAL_MAX_MENSAJES = 10  # Groq 6.000 TPM + Mistral como respaldo dan margen suficiente

# ── Límites de seguridad para archivos subidos ────────────────────────────────
MAX_FILAS    = 50_000
MAX_COLUMNAS = 100
MAX_HOJAS    = 10

_GUIA_ESTILO = cargar_base_conocimiento()

SYSTEM_PROMPT = (
    SYSTEM_BASE
    + "\n\nGuía de tono y formato:\n\n"
    + _GUIA_ESTILO
)
