import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from telegram.error import TelegramError
from services.llm import obtener_respuesta
from utils.history import obtener_historial, agregar_mensaje
from utils.excel_context import obtener_contexto
from utils.auth import solo_autorizados
from utils.user_prefs import get_version, ya_fue_preguntado, marcar_preguntado, VERSIONES

logger = logging.getLogger(__name__)

_TECLADO_VERSION = InlineKeyboardMarkup([
    [
        InlineKeyboardButton("Microsoft 365",      callback_data="version_365"),
        InlineKeyboardButton("Excel 2021",         callback_data="version_2021"),
    ],
    [
        InlineKeyboardButton("Excel 2019",         callback_data="version_2019"),
        InlineKeyboardButton("Excel 2016 o anterior", callback_data="version_2016"),
    ],
])


@solo_autorizados
async def responder_mensaje(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    pregunta = update.message.text.strip()

    # ── Explicador automático de fórmulas ─────────────────────────────────────
    if pregunta.startswith("="):
        await _explicar_formula(update, user_id, pregunta)
        return

    # ── Flujo normal ──────────────────────────────────────────────────────────
    mensaje_carga = await update.message.reply_text("⏳ Pensando...")

    historial = obtener_historial(user_id)
    contexto_excel = obtener_contexto(user_id)

    # Inyectar versión de Excel si está configurada
    version = get_version(user_id)
    prefijo_version = (
        f"[El usuario usa {VERSIONES.get(version, version)}. "
        "Adapta la respuesta a las funciones disponibles en esa versión.]\n\n"
        if version else ""
    )

    pregunta_completa = prefijo_version
    if contexto_excel:
        pregunta_completa += f"{contexto_excel}\n\nPregunta del usuario: {pregunta}"
    else:
        pregunta_completa += pregunta

    try:
        respuesta = obtener_respuesta(historial, pregunta_completa)
        agregar_mensaje(user_id, "user", pregunta)
        agregar_mensaje(user_id, "model", respuesta)
        await mensaje_carga.edit_text(respuesta)

        # Preguntar versión la primera vez (tras responder, no antes)
        if not ya_fue_preguntado(user_id):
            marcar_preguntado(user_id)
            await update.message.reply_text(
                "💡 Para darte respuestas más precisas, ¿qué versión de Excel usas?\n"
                "(Puedes cambiarlo siempre con /version)",
                reply_markup=_TECLADO_VERSION,
            )

    except TelegramError as error:
        logger.error("Error de Telegram para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ Error al enviar la respuesta. Inténtalo de nuevo.")
    except Exception as error:
        logger.error("Error inesperado para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text(
            "⚠️ El asistente no está disponible en este momento. Inténtalo en unos segundos."
        )


async def _explicar_formula(update: Update, user_id: int, formula: str) -> None:
    """Explica paso a paso una fórmula de Excel."""
    mensaje_carga = await update.message.reply_text("⏳ Analizando la fórmula...")
    prompt = (
        f"Explica paso a paso esta fórmula de Excel:\n\n{formula}\n\n"
        "Desglosa cada argumento o parte, explica qué hace cada uno y muestra "
        "un ejemplo práctico con datos reales de cuándo y cómo usarla."
    )
    try:
        historial = obtener_historial(user_id)
        respuesta = obtener_respuesta(historial, prompt)
        agregar_mensaje(user_id, "user", formula)
        agregar_mensaje(user_id, "model", respuesta)
        await mensaje_carga.edit_text(respuesta)
    except Exception as error:
        logger.error("Error explicando fórmula para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ No se pudo analizar la fórmula. Inténtalo de nuevo.")
