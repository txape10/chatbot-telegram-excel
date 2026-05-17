from telegram import Update
from telegram.ext import ContextTypes
from services.gemini import obtener_respuesta
from utils.history import obtener_historial, agregar_mensaje


async def responder_mensaje(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    pregunta = update.message.text

    await update.message.chat.send_action("typing")

    historial = obtener_historial(user_id)

    try:
        respuesta = obtener_respuesta(historial, pregunta)
        agregar_mensaje(user_id, "user", pregunta)
        agregar_mensaje(user_id, "model", respuesta)
        await update.message.reply_text(respuesta)
    except Exception as error:
        await update.message.reply_text(
            "⚠️ Ha ocurrido un error al procesar tu pregunta. Inténtalo de nuevo."
        )
        raise error
