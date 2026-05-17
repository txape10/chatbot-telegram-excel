import logging
from telegram import Update
from telegram.ext import ContextTypes
from telegram.error import TelegramError
from services.llm import obtener_respuesta
from utils.history import obtener_historial, agregar_mensaje
from utils.excel_context import obtener_contexto
from utils.auth import solo_autorizados

logger = logging.getLogger(__name__)


@solo_autorizados
async def responder_mensaje(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    pregunta = update.message.text

    mensaje_carga = await update.message.reply_text("⏳ Pensando...")

    historial = obtener_historial(user_id)
    contexto_excel = obtener_contexto(user_id)
    pregunta_completa = (
        f"{contexto_excel}\n\nPregunta del usuario: {pregunta}"
        if contexto_excel else pregunta
    )

    try:
        respuesta = obtener_respuesta(historial, pregunta_completa)
        agregar_mensaje(user_id, "user", pregunta)
        agregar_mensaje(user_id, "model", respuesta)
        await mensaje_carga.edit_text(respuesta)
    except TelegramError as error:
        logger.error("Error de Telegram para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ Error al enviar la respuesta. Inténtalo de nuevo.")
    except Exception as error:
        logger.error("Error inesperado para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text(
            "⚠️ El asistente no está disponible en este momento. Inténtalo en unos segundos."
        )
