"""
Configuración del bot de Telegram — compartida entre bot.py (polling local)
y api.py (webhook en cloud).

Cada modo importa crear_aplicacion() y lo usa a su manera.
"""
from telegram import BotCommand
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    MessageHandler,
    filters,
)

from config import TELEGRAM_TOKEN
from handlers.audio import recibir_audio, recibir_voz
from handlers.commands import (
    ayuda, callback_ayuda, callback_modo, callback_plantilla,
    callback_version, codigo, ejemplo, estado, generar, limpiar,
    modo, pivote, plantilla, privado, start, version,
    vincular, desvincular,
)
from handlers.documents import callback_chart, callback_sheet, recibir_documento
from handlers.images import recibir_imagen
from handlers.messages import callback_aclaracion, callback_confirmacion, responder_mensaje


async def _registrar_comandos(app):
    await app.bot.set_my_commands([
        BotCommand("start",     "Bienvenida e instrucciones"),
        BotCommand("ayuda",     "Categorías de temas disponibles"),
        BotCommand("ejemplo",   "Ejemplo de una función: /ejemplo BUSCARV"),
        BotCommand("generar",   "Genera un .xlsx de ejemplo: /generar BUSCARV"),
        BotCommand("plantilla", "Plantillas listas para usar (presupuesto, KPIs...)"),
        BotCommand("pivote",    "Genera un .xlsx con tabla dinámica de ejemplo"),
        BotCommand("version",   "Configura tu versión de Excel"),
        BotCommand("modo",      "Elige si quieres respuestas por voz o texto"),
        BotCommand("estado",    "Ver el estado actual de tu sesión"),
        BotCommand("privado",   "Activar/desactivar modo privado (sin historial)"),
        BotCommand("limpiar",    "Borrar el historial de conversación"),
        BotCommand("vincular",   "Vincular con el Add-in de Excel: /vincular email"),
        BotCommand("desvincular","Eliminar la vinculación con el Add-in"),
        BotCommand("codigo",     "Código de 6 dígitos para emparejar el Add-in"),
    ])


def crear_aplicacion():
    """Construye y devuelve la Application de PTB con todos los handlers registrados."""
    app = (
        ApplicationBuilder()
        .token(TELEGRAM_TOKEN)
        .post_init(_registrar_comandos)
        .build()
    )

    # Comandos
    app.add_handler(CommandHandler("start",     start))
    app.add_handler(CommandHandler("ayuda",     ayuda))
    app.add_handler(CommandHandler("limpiar",   limpiar))
    app.add_handler(CommandHandler("ejemplo",   ejemplo))
    app.add_handler(CommandHandler("generar",   generar))
    app.add_handler(CommandHandler("plantilla", plantilla))
    app.add_handler(CommandHandler("pivote",    pivote))
    app.add_handler(CommandHandler("version",   version))
    app.add_handler(CommandHandler("modo",      modo))
    app.add_handler(CommandHandler("estado",    estado))
    app.add_handler(CommandHandler("privado",     privado))
    app.add_handler(CommandHandler("vincular",    vincular))
    app.add_handler(CommandHandler("desvincular", desvincular))
    app.add_handler(CommandHandler("codigo",      codigo))

    # Callbacks InlineKeyboard
    app.add_handler(CallbackQueryHandler(callback_ayuda,         pattern="^cat_"))
    app.add_handler(CallbackQueryHandler(callback_version,       pattern="^version_"))
    app.add_handler(CallbackQueryHandler(callback_plantilla,     pattern="^plantilla_"))
    app.add_handler(CallbackQueryHandler(callback_modo,          pattern="^modo_"))
    app.add_handler(CallbackQueryHandler(callback_confirmacion,  pattern="^confirmar_op_"))
    app.add_handler(CallbackQueryHandler(callback_aclaracion,    pattern="^aclaracion_"))
    app.add_handler(CallbackQueryHandler(callback_sheet,         pattern="^sheet_"))
    app.add_handler(CallbackQueryHandler(callback_chart,         pattern="^chart_"))

    # Archivos Excel / CSV
    app.add_handler(MessageHandler(
        filters.Document.MimeType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        | filters.Document.MimeType("application/vnd.ms-excel")
        | filters.Document.MimeType("text/csv")
        | filters.Document.FileExtension("csv"),
        recibir_documento,
    ))

    # Imágenes, voz, texto
    app.add_handler(MessageHandler(filters.PHOTO,  recibir_imagen))
    app.add_handler(MessageHandler(filters.VOICE,  recibir_voz))
    app.add_handler(MessageHandler(filters.AUDIO,  recibir_audio))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, responder_mensaje))

    return app
