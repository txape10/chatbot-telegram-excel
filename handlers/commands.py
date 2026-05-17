import random
import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from utils.history import limpiar_historial
from utils.excel_context import borrar_contexto
from utils.auth import solo_autorizados
from services.llm import obtener_respuesta
from excel.exporter import crear_ejemplo as crear_ejemplo_xlsx

logger = logging.getLogger(__name__)

MENSAJE_BIENVENIDA = (
    "👋 ¡Hola! Soy tu asistente personal de Excel.\n\n"
    "Hazme cualquier pregunta sobre Excel: fórmulas, tablas dinámicas, "
    "formato condicional, gráficos, macros/VBA, Power Query...\n\n"
    "Comandos disponibles:\n"
    "/ayuda — categorías de temas\n"
    "/ejemplo — función aleatoria de Excel con ejemplo\n"
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

FUNCIONES_EXCEL = [
    "BUSCARV", "BUSCARX", "ÍNDICE/COINCIDIR", "SUMAR.SI", "SUMAR.SI.CONJUNTO",
    "CONTAR.SI", "CONTAR.SI.CONJUNTO", "SI", "SI.ERROR", "SI.ND",
    "TEXTO", "CONCATENAR", "UNIRCADENAS", "IZQUIERDA/DERECHA/EXTRAE",
    "FECHA", "FECHANUMERO", "DÍAS", "AÑO/MES/DÍA",
    "LAMBDA", "LET", "BYROW", "BYCOL", "MAKEARRAY",
    "FILTRAR", "ORDENAR", "ÚNICO", "SECUENCIA",
    "TRANSPONER", "MMULT", "DESREF", "INDIRECTO",
    "MAX.SI.CONJUNTO", "MIN.SI.CONJUNTO", "PROMEDIO.SI",
]


@solo_autorizados
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(MENSAJE_BIENVENIDA)


@solo_autorizados
async def ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    teclado = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📐 Fórmulas",         callback_data="cat_formulas"),
            InlineKeyboardButton("🔄 Tablas dinámicas", callback_data="cat_tablas"),
        ],
        [
            InlineKeyboardButton("🎨 Formato cond.",    callback_data="cat_formato"),
            InlineKeyboardButton("📊 Gráficos",         callback_data="cat_graficos"),
        ],
        [
            InlineKeyboardButton("⚙️ Macros / VBA",     callback_data="cat_vba"),
            InlineKeyboardButton("🔗 Power Query",      callback_data="cat_powerquery"),
        ],
    ])
    await update.message.reply_text(MENSAJE_AYUDA, parse_mode="Markdown", reply_markup=teclado)


@solo_autorizados
async def limpiar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    limpiar_historial(user_id)
    borrar_contexto(user_id)
    await update.message.reply_text("🗑️ Historial y contexto Excel borrados. ¡Empezamos de cero!")


@solo_autorizados
async def ejemplo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    funcion = " ".join(context.args).upper() if context.args else random.choice(FUNCIONES_EXCEL)
    mensaje_carga = await update.message.reply_text(f"⏳ Generando ejemplo de {funcion}...")
    try:
        prompt = (
            f"Explícame la función {funcion} de Microsoft Excel con un ejemplo práctico. "
            "Usa datos reales y concretos, muestra la fórmula exacta y añade un consejo útil."
        )
        respuesta = obtener_respuesta([], prompt)
        await mensaje_carga.edit_text(respuesta)
    except Exception as error:
        logger.error("Error en /ejemplo con función %s: %s", funcion, error)
        await mensaje_carga.edit_text("⚠️ No se pudo generar el ejemplo. Inténtalo de nuevo.")


@solo_autorizados
async def generar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Genera y envía un archivo .xlsx de ejemplo para la función indicada."""
    if not context.args:
        await update.message.reply_text(
            "ℹ️ Indica la función para la que quieres el archivo.\n"
            "Ejemplo: /generar BUSCARV"
        )
        return

    funcion = " ".join(context.args).upper()
    mensaje_carga = await update.message.reply_text(f"⏳ Generando archivo de ejemplo para {funcion}...")
    try:
        buffer, nombre_archivo = crear_ejemplo_xlsx(funcion)
        await mensaje_carga.delete()
        await update.message.reply_document(
            document=buffer,
            filename=nombre_archivo,
            caption=f"📎 Ejemplo de {funcion} listo para abrir en Excel."
        )
    except Exception as error:
        logger.error("Error en /generar con función %s: %s", funcion, error)
        await mensaje_carga.edit_text("⚠️ No se pudo generar el archivo. Inténtalo de nuevo.")


# Textos para cada categoría del /ayuda
_TEXTOS_CATEGORIA = {
    "cat_formulas": (
        "📐 *Fórmulas y funciones*\n\n"
        "Algunas de las más útiles:\n"
        "• BUSCARV / BUSCARX — búsqueda vertical\n"
        "• SUMAR.SI / SUMAR.SI.CONJUNTO — sumas condicionadas\n"
        "• CONTAR.SI — contar por criterio\n"
        "• SI / SI.ERROR — lógica condicional\n"
        "• ÍNDICE + COINCIDIR — búsqueda flexible\n\n"
        "Escríbeme tu duda o usa /ejemplo BUSCARV para ver un ejemplo."
    ),
    "cat_tablas": (
        "🔄 *Tablas dinámicas*\n\n"
        "Te puedo ayudar con:\n"
        "• Crear y configurar una tabla dinámica\n"
        "• Agrupar por fechas, categorías o rangos\n"
        "• Filtros y segmentaciones\n"
        "• Campos calculados\n\n"
        "Escríbeme tu duda directamente 💬"
    ),
    "cat_formato": (
        "🎨 *Formato condicional*\n\n"
        "Te puedo ayudar con:\n"
        "• Reglas basadas en valores o fórmulas\n"
        "• Escalas de color y barras de datos\n"
        "• Conjuntos de iconos\n"
        "• Resaltar duplicados o fechas\n\n"
        "Escríbeme tu duda directamente 💬"
    ),
    "cat_graficos": (
        "📊 *Gráficos*\n\n"
        "Te puedo ayudar con:\n"
        "• Elegir el tipo de gráfico adecuado\n"
        "• Gráficos dinámicos vinculados a tablas\n"
        "• Personalización: colores, ejes, leyendas\n"
        "• Minigráficos (Sparklines)\n\n"
        "Escríbeme tu duda directamente 💬"
    ),
    "cat_vba": (
        "⚙️ *Macros / VBA*\n\n"
        "Te puedo ayudar con:\n"
        "• Grabar y editar macros básicas\n"
        "• Bucles, condiciones y variables en VBA\n"
        "• Automatizar tareas repetitivas\n"
        "• UserForms y cuadros de diálogo\n\n"
        "Escríbeme tu duda directamente 💬"
    ),
    "cat_powerquery": (
        "🔗 *Power Query*\n\n"
        "Te puedo ayudar con:\n"
        "• Importar datos desde Excel, CSV, web o BD\n"
        "• Limpiar y transformar datos\n"
        "• Combinar y anexar consultas\n"
        "• Lenguaje M básico\n\n"
        "Escríbeme tu duda directamente 💬"
    ),
}


async def callback_ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Responde a los botones del menú /ayuda."""
    query = update.callback_query
    await query.answer()
    texto = _TEXTOS_CATEGORIA.get(query.data, "Escríbeme tu duda directamente 💬")
    await query.edit_message_text(texto, parse_mode="Markdown")
