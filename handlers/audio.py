"""Handler para mensajes de voz y archivos de audio.

Flujo:
1. Recibe el mensaje de voz grabado en Telegram (OGG/Opus).
2. Descarga los bytes del audio.
3. Transcribe con Groq Whisper (whisper-large-v3-turbo).
4. Muestra al usuario el texto reconocido para que pueda verificarlo.
5. Redirige al flujo normal de `procesar_pregunta` como si hubiera escrito el texto.
"""
import asyncio
import logging

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from utils.auth import solo_autorizados
from utils.user_prefs import ya_fue_preguntado_modo, marcar_preguntado_modo
from services.llm import LLMError, transcribir_audio
from handlers.messages import procesar_pregunta

logger = logging.getLogger(__name__)

_TECLADO_MODO = InlineKeyboardMarkup([
    [
        InlineKeyboardButton("🔊 Respóndeme por voz", callback_data="modo_voz"),
        InlineKeyboardButton("💬 Solo texto",          callback_data="modo_texto"),
    ]
])


@solo_autorizados
async def recibir_voz(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Procesa un mensaje de voz grabado desde el micrófono de Telegram."""
    mensaje_carga = await update.message.reply_text("🎧 Transcribiendo audio...")
    try:
        archivo = await update.message.voice.get_file()
        audio_bytes = bytes(await archivo.download_as_bytearray())

        texto = await asyncio.to_thread(transcribir_audio, audio_bytes, "audio.ogg")

        if not texto:
            await mensaje_carga.edit_text(
                "⚠️ No detecté ninguna voz en el audio. "
                "Intenta de nuevo hablando más cerca del micrófono."
            )
            return

        logger.info("Voz transcrita para user_id %s: %r", update.effective_user.id, texto)

        await mensaje_carga.delete()
        # Confirmar lo que se entendió para que el usuario pueda detectar errores
        await update.message.reply_text(f"🎤 Te escuché:\n_{texto}_", parse_mode="Markdown")

        # Procesar el texto transcrito exactamente igual que si lo hubiera escrito
        await procesar_pregunta(update, context, texto)

        # Una sola vez: preguntar si quiere respuestas por voz
        user_id = update.effective_user.id
        if not ya_fue_preguntado_modo(user_id):
            marcar_preguntado_modo(user_id)
            await update.message.reply_text(
                "💡 ¿Quieres que también te responda por voz?\n"
                "Puedes cambiarlo cuando quieras con /modo",
                reply_markup=_TECLADO_MODO,
            )

    except LLMError as error:
        logger.warning("Error LLM voz para user_id %s [%s]: %s", update.effective_user.id, error.tipo, error)
        try:
            await mensaje_carga.edit_text(error.mensaje_usuario)
        except Exception:
            pass
    except Exception as error:
        logger.error("Error transcribiendo voz para user_id %s: %s",
                     update.effective_user.id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text(
                "⚠️ No pude transcribir el audio. "
                "Comprueba que el micrófono funciona e inténtalo de nuevo."
            )
        except Exception:
            pass


@solo_autorizados
async def recibir_audio(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Procesa un archivo de audio enviado como adjunto (mp3, m4a, ogg, wav…)."""
    mensaje_carga = await update.message.reply_text("🎧 Transcribiendo archivo de audio...")
    try:
        audio = update.message.audio
        archivo = await audio.get_file()
        audio_bytes = bytes(await archivo.download_as_bytearray())

        # Intentar obtener el nombre original para que Groq infiera el codec
        nombre_archivo = audio.file_name or "audio.mp3"

        texto = await asyncio.to_thread(transcribir_audio, audio_bytes, nombre_archivo)

        if not texto:
            await mensaje_carga.edit_text(
                "⚠️ No pude extraer texto del archivo de audio."
            )
            return

        logger.info("Audio transcrito para user_id %s: %r", update.effective_user.id, texto)

        await mensaje_carga.delete()
        await update.message.reply_text(f"🎤 Transcripción:\n_{texto}_", parse_mode="Markdown")

        await procesar_pregunta(update, context, texto)

    except LLMError as error:
        logger.warning("Error LLM audio para user_id %s [%s]: %s", update.effective_user.id, error.tipo, error)
        try:
            await mensaje_carga.edit_text(error.mensaje_usuario)
        except Exception:
            pass
    except Exception as error:
        logger.error("Error transcribiendo audio para user_id %s: %s",
                     update.effective_user.id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text(
                "⚠️ No pude procesar el archivo de audio. "
                "Formatos soportados: ogg, mp3, m4a, wav, webm."
            )
        except Exception:
            pass
