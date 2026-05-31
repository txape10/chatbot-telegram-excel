import io
import logging

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from services.llm import LLMError, obtener_respuesta, obtener_proveedor_privado
from services.tts import texto_a_audio
from utils.history import obtener_historial, agregar_mensaje
from utils.excel_context import obtener_contexto
from utils.user_prefs import get_modo_respuesta, get_modo_privado, get_version, VERSIONES
from prompts.excel import PREGUNTA_CON_VERSION, PREGUNTA_CON_CONTEXTO
from handlers.intent_patterns import _CLAVE_ACLARACION

logger = logging.getLogger(__name__)


async def _pedir_aclaracion(update: Update, context: ContextTypes.DEFAULT_TYPE,
                             user_id: int, tipo: str, aclaracion: dict) -> None:
    """Muestra al usuario un InlineKeyboard con las opciones de aclaración.

    tipo: "edicion" | "dsl"
    aclaracion: {"aclaracion_necesaria": True, "pregunta": "...", "opciones": [...]}
    """
    pregunta_texto = aclaracion.get("pregunta", "¿Puedes concretar más?")
    opciones       = aclaracion.get("opciones", [])[:3]

    context.user_data["aclaracion_pendiente"] = {
        "tipo":    tipo,
        "opciones": opciones,
    }

    botones = [
        [InlineKeyboardButton(opt, callback_data=f"aclaracion_{i}")]
        for i, opt in enumerate(opciones)
    ]
    teclado = InlineKeyboardMarkup(botones)

    await update.message.reply_text(
        f"🤔 {pregunta_texto}",
        reply_markup=teclado,
    )


async def _enviar_respuesta(update: Update, user_id: int,
                             mensaje_carga, texto: str,
                             parse_mode: str = None) -> None:
    """Envía la respuesta en texto o texto+voz según la preferencia del usuario."""
    await mensaje_carga.edit_text(texto, parse_mode=parse_mode)

    if get_modo_respuesta(user_id) == "voz":
        try:
            audio_bytes = await texto_a_audio(texto)
            if audio_bytes:
                await update.message.reply_voice(voice=io.BytesIO(audio_bytes))
        except Exception as error:
            logger.warning("TTS falló para user_id %s: %s", user_id, error)


async def _responder_con_llm(update: Update, user_id: int, pregunta: str,
                              mensaje_carga=None) -> None:
    """Responde usando el LLM normal (sin contexto de edición)."""
    if mensaje_carga is None:
        mensaje_carga = await update.message.reply_text("⏳ Pensando...")
    historial = obtener_historial(user_id)
    contexto_excel = obtener_contexto(user_id)
    version = get_version(user_id)
    prefijo = PREGUNTA_CON_VERSION.format(version=VERSIONES.get(version, version)) if version else ""
    pregunta_completa = prefijo + (
        PREGUNTA_CON_CONTEXTO.format(contexto=contexto_excel, pregunta=pregunta)
        if contexto_excel else pregunta
    )
    try:
        _modo_privado = get_modo_privado(user_id)
        respuesta = obtener_respuesta(
            historial, pregunta_completa,
            obtener_proveedor_privado() if _modo_privado else None,
        )
        if not _modo_privado:
            agregar_mensaje(user_id, "user", pregunta)
            agregar_mensaje(user_id, "model", respuesta)
        await _enviar_respuesta(update, user_id, mensaje_carga, respuesta)
    except LLMError as error:
        logger.warning("Error LLM fallback para user_id %s [%s]: %s", user_id, error.tipo, error)
        await mensaje_carga.edit_text(error.mensaje_usuario)
    except Exception as error:
        logger.error("Error LLM fallback para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ Error al obtener respuesta. Inténtalo de nuevo.")
