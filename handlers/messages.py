import asyncio
import logging
import re
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from telegram.error import TelegramError
from services.llm import obtener_respuesta, extraer_query_dsl
from utils.history import obtener_historial, agregar_mensaje
from utils.excel_context import obtener_contexto
from utils.df_context import obtener_df
from utils.auth import solo_autorizados
from utils.user_prefs import get_version, ya_fue_preguntado, marcar_preguntado, VERSIONES
from prompts.excel import EXPLICAR_FORMULA, PREGUNTA_CON_VERSION, PREGUNTA_CON_CONTEXTO
from excel.query_engine import ejecutar_query, formatear_resultado, QueryError
from excel.exporter import crear_tabla_dinamica

# Detección de tabla dinámica: basta con que mencione "tabla dinámica/s"
# salvo que sea una pregunta puramente informativa (qué es, cómo funciona, etc.)
_RE_TABLA_DINAMICA = re.compile(
    r"tabla[s]?\s+din[aá]mica[s]?",
    re.IGNORECASE,
)
_RE_SOLO_INFORMATIVA = re.compile(
    r"^[¿\s]*(qu[eé]\s+es\b|qu[eé]\s+son\b|c[oó]mo\s+(se\s+)?func|"
    r"para\s+qu[eé]\b|cu[aá]ndo\s+usar|por\s+qu[eé]\b|"
    r"explica[rm]?|qu[eé]\s+hace\b|qu[eé]\s+hacen\b|en\s+qu[eé]\s+consiste)",
    re.IGNORECASE,
)

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

    # ── Generador de tabla dinámica ───────────────────────────────────────────
    if _RE_TABLA_DINAMICA.search(pregunta) and not _RE_SOLO_INFORMATIVA.search(pregunta):
        logger.info("Intención tabla dinámica detectada para user_id %s: %r", user_id, pregunta)
        await _generar_tabla_dinamica(update, user_id)
        return

    # ── Flujo normal ──────────────────────────────────────────────────────────
    mensaje_carga = await update.message.reply_text("⏳ Pensando...")

    historial = obtener_historial(user_id)
    contexto_excel = obtener_contexto(user_id)

    # Inyectar versión de Excel si está configurada
    version = get_version(user_id)
    prefijo_version = (
        PREGUNTA_CON_VERSION.format(version=VERSIONES.get(version, version))
        if version else ""
    )

    if contexto_excel:
        pregunta_completa = prefijo_version + PREGUNTA_CON_CONTEXTO.format(
            contexto=contexto_excel, pregunta=pregunta
        )
    else:
        pregunta_completa = prefijo_version + pregunta

    try:
        # ── Intento DSL si hay DataFrame activo ───────────────────────────────
        df_activo = obtener_df(user_id)
        if df_activo is not None and contexto_excel:
            respuesta = await _intentar_dsl(
                update, user_id, df_activo, pregunta, historial,
                pregunta_completa, mensaje_carga,
            )
            if respuesta is not None:
                return

        # ── Flujo LLM normal ──────────────────────────────────────────────────
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


async def _generar_tabla_dinamica(update: Update, user_id: int) -> None:
    """Genera y envía un .xlsx con tabla dinámica usando los datos del usuario (si los tiene)."""
    mensaje_carga = await update.message.reply_text("⏳ Generando tabla dinámica...")
    try:
        df = obtener_df(user_id)
        usar_datos_usuario = df is not None and not df.empty
        logger.info("Generando tabla dinámica para user_id %s (datos propios: %s)",
                    user_id, usar_datos_usuario)

        buf, nombre = await asyncio.to_thread(crear_tabla_dinamica, df)

        if usar_datos_usuario:
            caption = (
                "📊 Tabla dinámica con tus datos\n\n"
                "El archivo tiene dos hojas:\n"
                "· Datos — tus datos originales\n"
                "· Tabla Dinámica — resúmenes agrupados\n\n"
                "💡 En Excel: Insertar → Tabla dinámica para una versión interactiva."
            )
        else:
            caption = (
                "📊 Ejemplo de tabla dinámica\n\n"
                "El archivo tiene dos hojas:\n"
                "· Datos — ventas ficticias de ejemplo\n"
                "· Tabla Dinámica — resúmenes y cruces calculados\n\n"
                "💡 Sube tu Excel y te la genero con tus propios datos."
            )

        # Enviar el archivo PRIMERO; borrar el mensaje de carga después
        await update.message.reply_document(document=buf, filename=nombre, caption=caption)
        try:
            await mensaje_carga.delete()
        except Exception:
            pass   # si ya no existe el mensaje de carga, no es crítico

    except Exception as error:
        logger.error("Error generando tabla dinámica para user_id %s: %s",
                     user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo generar el archivo. Inténtalo de nuevo.")
        except Exception:
            pass


async def _intentar_dsl(update, user_id, df, pregunta, historial,
                        pregunta_completa, mensaje_carga):
    """Intenta resolver la pregunta mediante el motor DSL.

    Devuelve la respuesta si tuvo éxito, o None si debe continuar por LLM normal.
    """
    try:
        query = await asyncio.to_thread(extraer_query_dsl, df, pregunta)
        if query is None:
            return None   # el LLM decidió RESPUESTA_LIBRE → flujo normal

        resultado, descripcion = await asyncio.to_thread(ejecutar_query, df, query)
        texto = formatear_resultado(resultado, descripcion)

        agregar_mensaje(user_id, "user", pregunta)
        agregar_mensaje(user_id, "model", texto)
        await mensaje_carga.edit_text(texto, parse_mode="Markdown")
        return texto

    except QueryError as error:
        logger.warning("QueryError para user_id %s: %s — usando LLM normal", user_id, error)
        return None   # error de datos → mejor responder con LLM normal
    except Exception as error:
        logger.warning("Error en DSL para user_id %s: %s — usando LLM normal", user_id, error)
        return None


async def _explicar_formula(update: Update, user_id: int, formula: str) -> None:
    """Explica paso a paso una fórmula de Excel."""
    mensaje_carga = await update.message.reply_text("⏳ Analizando la fórmula...")
    prompt = EXPLICAR_FORMULA.format(formula=formula)
    try:
        historial = obtener_historial(user_id)
        respuesta = obtener_respuesta(historial, prompt)
        agregar_mensaje(user_id, "user", formula)
        agregar_mensaje(user_id, "model", respuesta)
        await mensaje_carga.edit_text(respuesta)
    except Exception as error:
        logger.error("Error explicando fórmula para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ No se pudo analizar la fórmula. Inténtalo de nuevo.")
