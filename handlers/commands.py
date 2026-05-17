import random
import logging
from telegram import Update
from telegram.ext import ContextTypes
from utils.history import limpiar_historial
from utils.auth import solo_autorizados
from services.gemini import obtener_respuesta

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
    await update.message.reply_text(MENSAJE_AYUDA, parse_mode="Markdown")


@solo_autorizados
async def limpiar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    limpiar_historial(update.effective_user.id)
    await update.message.reply_text("🗑️ Historial borrado. ¡Empezamos de cero!")


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
