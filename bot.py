from logging_config import configurar_logging
configurar_logging()

from telegram import BotCommand
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, filters
from config import TELEGRAM_TOKEN
from handlers.commands import (
    start, ayuda, limpiar, ejemplo, generar, plantilla, version,
    callback_ayuda, callback_version, callback_plantilla,
)
from handlers.messages import responder_mensaje
from handlers.documents import recibir_documento, callback_sheet, callback_chart
from handlers.images import recibir_imagen


async def registrar_comandos(app):
    await app.bot.set_my_commands([
        BotCommand("start",     "Bienvenida e instrucciones"),
        BotCommand("ayuda",     "Categorías de temas disponibles"),
        BotCommand("ejemplo",   "Ejemplo de una función: /ejemplo BUSCARV"),
        BotCommand("generar",   "Genera un .xlsx de ejemplo: /generar BUSCARV"),
        BotCommand("plantilla", "Plantillas listas para usar (presupuesto, KPIs...)"),
        BotCommand("version",   "Configura tu versión de Excel"),
        BotCommand("limpiar",   "Borrar el historial de conversación"),
    ])


def main() -> None:
    app = (
        ApplicationBuilder()
        .token(TELEGRAM_TOKEN)
        .post_init(registrar_comandos)
        .build()
    )

    # Comandos
    app.add_handler(CommandHandler("start",     start))
    app.add_handler(CommandHandler("ayuda",     ayuda))
    app.add_handler(CommandHandler("limpiar",   limpiar))
    app.add_handler(CommandHandler("ejemplo",   ejemplo))
    app.add_handler(CommandHandler("generar",   generar))
    app.add_handler(CommandHandler("plantilla", plantilla))
    app.add_handler(CommandHandler("version",   version))

    # Callbacks de botones InlineKeyboard
    app.add_handler(CallbackQueryHandler(callback_ayuda,     pattern="^cat_"))
    app.add_handler(CallbackQueryHandler(callback_version,   pattern="^version_"))
    app.add_handler(CallbackQueryHandler(callback_plantilla, pattern="^plantilla_"))
    app.add_handler(CallbackQueryHandler(callback_sheet,     pattern="^sheet_"))
    app.add_handler(CallbackQueryHandler(callback_chart,     pattern="^chart_"))

    # Mensajes con archivos
    app.add_handler(MessageHandler(
        filters.Document.MimeType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        | filters.Document.MimeType("application/vnd.ms-excel")
        | filters.Document.MimeType("text/csv")
        | filters.Document.FileExtension("csv"),
        recibir_documento,
    ))

    # Imágenes y texto
    app.add_handler(MessageHandler(filters.PHOTO, recibir_imagen))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, responder_mensaje))

    print("Bot iniciado. Pulsa Ctrl+C para detenerlo.")
    app.run_polling()


if __name__ == "__main__":
    main()
