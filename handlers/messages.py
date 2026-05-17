import asyncio
import logging
import re
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from telegram.error import TelegramError
from services.llm import obtener_respuesta, extraer_query_dsl, extraer_operacion_edicion
from utils.history import obtener_historial, agregar_mensaje
from utils.excel_context import obtener_contexto
from utils.df_context import obtener_df, guardar_df
from utils.file_meta import obtener_meta
from utils.auth import solo_autorizados
from utils.user_prefs import get_version, ya_fue_preguntado, marcar_preguntado, VERSIONES
from prompts.excel import EXPLICAR_FORMULA, PREGUNTA_CON_VERSION, PREGUNTA_CON_CONTEXTO
from excel.query_engine import ejecutar_query, formatear_resultado, QueryError
from excel.editor import aplicar_edicion, exportar_xlsx, EditorError
from excel.exporter import crear_tabla_dinamica

# Detección de intención de edición: verbos que implican modificar el archivo
_RE_EDICION = re.compile(
    r"\b("
    r"a[ñn]ade[r]?\s|agrega[r]?\s|crea\s+(?:una\s+)?columna|nueva\s+columna|"
    r"calcula\s+(?:una\s+)?columna|"
    r"ordena[r]?\s|ordena\s+(?:el|los|por)|"
    r"elimina[r]?\s+(?:los\s+|las\s+)?duplicados|"
    r"quita[r]?\s+(?:los\s+|las\s+)?duplicados|"
    r"borra[r]?\s+(?:los\s+|las\s+)?duplicados|"
    r"elimina[r]?\s+(?:la\s+|el\s+)?columna|"
    r"quita[r]?\s+(?:la\s+|el\s+)?columna|"
    r"borra[r]?\s+(?:la\s+|el\s+)?columna|"
    r"rellena[r]?\s+(?:los\s+|las\s+)?(?:vac[ií]os?|nulos?|huecos?)|"
    r"completa[r]?\s+(?:los\s+|las\s+)?(?:vac[ií]os?|nulos?)|"
    r"renombra[r]?\s|cambia[r]?\s+el\s+nombre\s+de\s|"
    r"aplica[r]?\s+formato\s+condicional|aplica[r]?\s+color|"
    r"pinta[r]?\s+(?:en\s+)?(?:rojo|verde|amarillo|naranja|azul)|colorea[r]?\s"
    r")\b",
    re.IGNORECASE,
)

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

    # ── Editor de archivo (si hay df activo y petición de modificación) ────────
    df_activo = obtener_df(user_id)
    if df_activo is not None and _RE_EDICION.search(pregunta):
        logger.info("Intención de edición detectada para user_id %s: %r", user_id, pregunta)
        await _intentar_edicion(update, user_id, df_activo, pregunta)
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


async def _intentar_edicion(update: Update, user_id: int, df, pregunta: str) -> None:
    """Extrae la operación de edición, aplica el cambio y envía el archivo modificado."""
    mensaje_carga = await update.message.reply_text("⏳ Aplicando modificación...")
    try:
        op = await asyncio.to_thread(extraer_operacion_edicion, df, pregunta)

        if op is None:
            # El LLM dijo RESPUESTA_LIBRE → dejar que el flujo normal responda
            await mensaje_carga.delete()
            # Re-lanzar como mensaje normal (sin edición)
            await _responder_con_llm(update, user_id, pregunta, mensaje_carga=None)
            return

        df_mod, descripcion, extras = await asyncio.to_thread(aplicar_edicion, df, op)

        # Obtener el nombre del archivo original si está disponible
        meta = obtener_meta(user_id)
        nombre_base = meta["nombre"] if meta else "archivo"

        fmt_cond = extras if op.get("op") == "formato_condicional" else None
        buf, nombre_archivo = await asyncio.to_thread(
            exportar_xlsx, df_mod, nombre_base, descripcion, fmt_cond
        )

        # Actualizar el df en memoria con la versión modificada
        guardar_df(user_id, df_mod)

        caption = f"✅ {descripcion}\n\nEl archivo incluye los cambios aplicados."
        await update.message.reply_document(document=buf, filename=nombre_archivo, caption=caption)
        try:
            await mensaje_carga.delete()
        except Exception:
            pass

    except EditorError as error:
        logger.warning("EditorError para user_id %s: %s", user_id, error)
        try:
            await mensaje_carga.edit_text(f"⚠️ No pude aplicar la modificación: {error}")
        except Exception:
            pass
    except Exception as error:
        logger.error("Error en edición para user_id %s: %s", user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo modificar el archivo. Inténtalo de nuevo.")
        except Exception:
            pass


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
        respuesta = obtener_respuesta(historial, pregunta_completa)
        agregar_mensaje(user_id, "user", pregunta)
        agregar_mensaje(user_id, "model", respuesta)
        await mensaje_carga.edit_text(respuesta)
    except Exception as error:
        logger.error("Error LLM fallback para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ Error al obtener respuesta. Inténtalo de nuevo.")


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
                "📊 Ejemplo de tabla dinámica (datos de muestra)\n\n"
                "No tenía tus datos en memoria — puede que hayas reiniciado el bot "
                "o aún no hayas subido ningún archivo.\n\n"
                "Sube tu Excel y repite la petición para generarla con tus propios datos."
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
