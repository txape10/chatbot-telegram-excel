from telegram import Update
from telegram.ext import ContextTypes
from utils.history import limpiar_historial
from utils.auth import solo_autorizados

MENSAJE_BIENVENIDA = (
    "👋 ¡Hola! Soy tu asistente personal de Excel.\n\n"
    "Hazme cualquier pregunta sobre Excel: fórmulas, tablas dinámicas, "
    "formato condicional, gráficos, macros/VBA, Power Query...\n\n"
    "Comandos disponibles:\n"
    "/ayuda — categorías de temas\n"
    "/limpiar — borrar el historial de conversación"
)

MENSAJE_AYUDA = (
    "📚 *Temas que puedo ayudarte:*\n\n"
    "• *Fórmulas y funciones* — BUSCARV, SUMAR.SI, CONTAR.SI, etc.\n"
    "• *Tablas dinámicas* — crear, agrupar, filtrar\n"
    "• *Formato condicional* — reglas, escalas de color, iconos\n"
    "• *Gráficos* — tipos, personalización, gráficos dinámicos\n"
    "• *Macros / VBA* — automatización, grabación, código\n"
    "• *Power Query* — importar, transformar y limpiar datos\n\n"
    "Escríbeme tu duda directamente 💬"
)


@solo_autorizados
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(MENSAJE_BIENVENIDA)


@solo_autorizados
async def ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(MENSAJE_AYUDA, parse_mode="Markdown")


@solo_autorizados
async def limpiar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    limpiar_historial(update.effective_user.id)
    await update.message.reply_text("🗑️ Historial borrado. ¡Empezamos de cero!")
