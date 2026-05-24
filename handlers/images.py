import logging
from telegram import Update
from telegram.ext import ContextTypes
from utils.auth import solo_autorizados
from services.llm import LLMError, analizar_imagen
from prompts.excel import IMAGEN_SIN_CAPTION

logger = logging.getLogger(__name__)


@solo_autorizados
async def recibir_imagen(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Analiza una captura de pantalla de Excel enviada por el usuario."""
    mensaje_carga = await update.message.reply_text("⏳ Analizando la imagen...")

    try:
        # Descargar la foto en mayor resolución disponible
        foto = update.message.photo[-1]
        archivo = await foto.get_file()
        buffer = await archivo.download_as_bytearray()

        caption = update.message.caption or IMAGEN_SIN_CAPTION
        respuesta = analizar_imagen(bytes(buffer), caption)
        await mensaje_carga.edit_text(respuesta)

    except LLMError as error:
        logger.warning("Error LLM imagen para user_id %s [%s]: %s", update.effective_user.id, error.tipo, error)
        await mensaje_carga.edit_text(error.mensaje_usuario)
    except Exception as error:
        logger.error("Error analizando imagen para user_id %s: %s", update.effective_user.id, error)
        await mensaje_carga.edit_text("⚠️ No se pudo analizar la imagen. Inténtalo de nuevo.")
