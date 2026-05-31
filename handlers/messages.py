import logging

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from telegram.error import TelegramError

from services.llm import LLMError, obtener_respuesta, obtener_proveedor_privado
from utils.history import obtener_historial, agregar_mensaje
from utils.excel_context import obtener_contexto
from utils.df_context import obtener_df, obtener_df_secundario
from utils.user_prefs import get_version, get_modo_privado, ya_fue_preguntado, marcar_preguntado, VERSIONES
from utils.feature_config import esta_activo as _feature_activa
from utils.auth import solo_autorizados
from prompts.excel import PREGUNTA_CON_VERSION, PREGUNTA_CON_CONTEXTO

from handlers.intent_patterns import (
    _CLAVE_ACLARACION,
    _RE_COMPARAR, _RE_COMBINAR, _RE_UNDO, _RE_EDICION, _RE_CREAR_EXCEL,
    _RE_GRAFICO, _RE_TABLA_DINAMICA, _RE_SOLO_INFORMATIVA, _RE_STATS,
    _RE_PREVIEW, _RE_VALORES_UNICOS, _RE_EXPLICAR_ARCHIVO, _RE_EXPORTAR_CSV,
    _RE_LISTAR_MACROS, _RE_GUARDAR_MACRO, _RE_BORRAR_MACRO, _RE_EJECUTAR_MACRO,
)
from handlers.bot_helpers import _pedir_aclaracion, _enviar_respuesta
from handlers.excel_edit import (
    callback_confirmacion,  # re-exportado para telegram_app.py
    _intentar_edicion, _intentar_combinar, _deshacer_operacion,
    _guardar_macro, _ejecutar_macro, _listar_macros, _borrar_macro,
)
from handlers.excel_query import (
    _intentar_dsl, _explicar_formula, _generar_tabla_dinamica,
    _crear_excel_desde_descripcion, _generar_grafico_bajo_demanda,
    _analizar_estadisticas, _previsualizar, _valores_unicos,
    _comparar_archivos, _explicar_archivo, _exportar_csv,
)

logger = logging.getLogger(__name__)

_TECLADO_VERSION = InlineKeyboardMarkup([
    [
        InlineKeyboardButton("Microsoft 365",         callback_data="version_365"),
        InlineKeyboardButton("Excel 2021",            callback_data="version_2021"),
    ],
    [
        InlineKeyboardButton("Excel 2019",            callback_data="version_2019"),
        InlineKeyboardButton("Excel 2016 o anterior", callback_data="version_2016"),
    ],
])


async def callback_aclaracion(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """El usuario eligió una opción de aclaración → re-procesar como nueva pregunta."""
    query = update.callback_query
    await query.answer()

    pendiente = context.user_data.get("aclaracion_pendiente")
    if not pendiente:
        await query.edit_message_text("⚠️ La sesión de aclaración expiró. Repite la pregunta.")
        return

    indice = int(query.data.split("_")[1])
    opciones = pendiente.get("opciones", [])

    if indice >= len(opciones):
        await query.edit_message_text("⚠️ Opción no válida.")
        return

    opcion_elegida = opciones[indice]
    context.user_data.pop("aclaracion_pendiente", None)

    await query.edit_message_text(f"✅ Entendido: *{opcion_elegida}*", parse_mode="Markdown")
    await procesar_pregunta(update, context, opcion_elegida)


@solo_autorizados
async def responder_mensaje(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await procesar_pregunta(update, context, update.message.text.strip())


async def procesar_pregunta(update: Update, context: ContextTypes.DEFAULT_TYPE,
                             pregunta: str) -> None:
    """Núcleo de procesamiento de una pregunta de texto.

    Reutilizable desde cualquier handler (texto, voz, etc.) una vez que
    el texto ya ha sido extraído o transcrito.
    """
    user_id = update.effective_user.id

    # ── Explicador automático de fórmulas ─────────────────────────────────────
    if pregunta.startswith("="):
        await _explicar_formula(update, user_id, pregunta)
        return

    df_activo     = obtener_df(user_id)
    df_secundario = obtener_df_secundario(user_id)

    # ── Comparar dos archivos ─────────────────────────────────────────────────
    if df_activo is not None and df_secundario is not None and _RE_COMPARAR.search(pregunta):
        await _comparar_archivos(update, user_id, df_secundario, df_activo)
        return

    # ── Combinar dos archivos ─────────────────────────────────────────────────
    if df_activo is not None and df_secundario is not None and _RE_COMBINAR.search(pregunta):
        await _intentar_combinar(update, user_id, df_secundario, df_activo, pregunta)
        return

    # ── Deshacer última operación ─────────────────────────────────────────────
    if df_activo is not None and _RE_UNDO.search(pregunta):
        await _deshacer_operacion(update, user_id)
        return

    # ── Edición de archivo ────────────────────────────────────────────────────
    if df_activo is not None and _RE_EDICION.search(pregunta):
        await _intentar_edicion(update, user_id, df_activo, pregunta, context)
        return

    # ── Crear Excel desde descripción ────────────────────────────────────────
    if _RE_CREAR_EXCEL.search(pregunta) and not _RE_EDICION.search(pregunta):
        await _crear_excel_desde_descripcion(update, user_id, pregunta)
        return

    # ── Gráfico bajo demanda ─────────────────────────────────────────────────
    if df_activo is not None and _RE_GRAFICO.search(pregunta):
        await _generar_grafico_bajo_demanda(update, user_id, df_activo, pregunta)
        return

    # ── Tabla dinámica ────────────────────────────────────────────────────────
    if _RE_TABLA_DINAMICA.search(pregunta) and not _RE_SOLO_INFORMATIVA.search(pregunta):
        await _generar_tabla_dinamica(update, user_id)
        return

    # ── Análisis estadístico ──────────────────────────────────────────────────
    if df_activo is not None and _RE_STATS.search(pregunta):
        await _analizar_estadisticas(update, user_id, df_activo, pregunta)
        return

    # ── Previsualizar filas ───────────────────────────────────────────────────
    if df_activo is not None and _RE_PREVIEW.search(pregunta):
        await _previsualizar(update, user_id, df_activo, pregunta)
        return

    # ── Valores únicos ────────────────────────────────────────────────────────
    if df_activo is not None and _RE_VALORES_UNICOS.search(pregunta):
        await _valores_unicos(update, user_id, df_activo, pregunta)
        return

    # ── Explicar archivo ──────────────────────────────────────────────────────
    if df_activo is not None and _RE_EXPLICAR_ARCHIVO.search(pregunta):
        await _explicar_archivo(update, user_id, df_activo)
        return

    # ── Exportar CSV ──────────────────────────────────────────────────────────
    if df_activo is not None and _RE_EXPORTAR_CSV.search(pregunta):
        await _exportar_csv(update, user_id, df_activo)
        return

    # ── Macros ────────────────────────────────────────────────────────────────
    _hay_peticion_macro = (
        _RE_LISTAR_MACROS.search(pregunta)
        or _RE_GUARDAR_MACRO.search(pregunta)
        or _RE_BORRAR_MACRO.search(pregunta)
        or _RE_EJECUTAR_MACRO.search(pregunta)
    )
    if _hay_peticion_macro and not _feature_activa("macros"):
        await update.message.reply_text("⚠️ El módulo de macros está desactivado en este momento.")
        return
    if _RE_LISTAR_MACROS.search(pregunta):
        await _listar_macros(update, user_id)
        return
    m_guardar = _RE_GUARDAR_MACRO.search(pregunta)
    if m_guardar:
        await _guardar_macro(update, user_id, m_guardar.group(1), pregunta)
        return
    m_borrar = _RE_BORRAR_MACRO.search(pregunta)
    if m_borrar:
        await _borrar_macro(update, user_id, m_borrar.group(1))
        return
    m_ejecutar = _RE_EJECUTAR_MACRO.search(pregunta)
    if m_ejecutar and df_activo is not None:
        await _ejecutar_macro(update, user_id, df_activo, m_ejecutar.group(1))
        return

    # ── Flujo normal (LLM) ────────────────────────────────────────────────────
    mensaje_carga = await update.message.reply_text("⏳ Pensando...")

    historial      = obtener_historial(user_id)
    contexto_excel = obtener_contexto(user_id)

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
                pregunta_completa, mensaje_carga, context=context,
            )
            if respuesta is not None:
                return

        # ── Flujo LLM normal ──────────────────────────────────────────────────
        _modo_privado = get_modo_privado(user_id)
        respuesta = obtener_respuesta(
            historial, pregunta_completa,
            obtener_proveedor_privado() if _modo_privado else None,
            user_id=user_id,
        )
        if not _modo_privado:
            agregar_mensaje(user_id, "user", pregunta)
            agregar_mensaje(user_id, "model", respuesta)
        await _enviar_respuesta(update, user_id, mensaje_carga, respuesta)

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
    except LLMError as error:
        logger.warning("Error LLM para user_id %s [%s]: %s", user_id, error.tipo, error)
        await mensaje_carga.edit_text(error.mensaje_usuario)
    except Exception as error:
        logger.error("Error inesperado para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text(
            "⚠️ El asistente no está disponible en este momento. Inténtalo en unos segundos."
        )
