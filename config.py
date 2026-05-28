import os
import platform
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

# ── Modo de despliegue ────────────────────────────────────────────────────────
# personal : bot Telegram personal + Add-in (defaults permisivos)
# empresa  : solo Add-in por defecto; Telegram desactivado; auth estricta
APP_MODE   = os.getenv("APP_MODE", "personal").lower()
IS_EMPRESA = APP_MODE == "empresa"

# ── Tablas dinámicas nativas (xlwings) ───────────────────────────────────────
# Solo disponible en Windows con Microsoft Excel instalado.
# En Linux (Render, empresa Linux) se usa el fallback de openpyxl automáticamente.
try:
    import xlwings as _xw  # noqa: F401
    PIVOT_NATIVO_DISPONIBLE = platform.system() == "Windows"
except ImportError:
    PIVOT_NATIVO_DISPONIBLE = False

# ── Módulos activables ────────────────────────────────────────────────────────
# Permite instalar solo el bot, solo el Add-in o ambos.
# En modo empresa, Telegram está desactivado a menos que se active explícitamente.
ENABLE_TELEGRAM = os.getenv("ENABLE_TELEGRAM", "false" if IS_EMPRESA else "true").lower() == "true"
ENABLE_ADDIN    = os.getenv("ENABLE_ADDIN", "true").lower() == "true"

# ── Proveedor de IA ───────────────────────────────────────────────────────────
# groq | ollama | openai  (ver services/llm_provider.py)
LLM_PROVIDER = os.getenv("LLM_PROVIDER", "groq")

HISTORIAL_MAX_MENSAJES = 10  # Groq 6.000 TPM + Mistral como respaldo dan margen suficiente

# ── Límites de seguridad para archivos subidos ────────────────────────────────
TAMANIO_MAXIMO_MB = 20        # tamaño máximo del archivo subido
MAX_FILAS    = 100_000
MAX_COLUMNAS = 100
MAX_HOJAS    = 10

_GUIA_ESTILO = cargar_base_conocimiento()

SYSTEM_PROMPT = (
    SYSTEM_BASE
    + "\n\nGuía de tono y formato:\n\n"
    + _GUIA_ESTILO
)

# Variante para el Add-in de Excel: mismo conocimiento + nota de contexto
SYSTEM_PROMPT_ADDIN = (
    SYSTEM_PROMPT
    + "\n\nCONTEXTO: Estás integrado en un panel lateral de Microsoft Excel (Add-in). "
    "El usuario trabaja directamente en su hoja de cálculo. "
    "Cuando el sistema devuelva datos tabulares, se escribirán automáticamente en las celdas. "
    "No menciones ni sugieras comandos del bot de Telegram."
)
