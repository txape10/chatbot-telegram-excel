from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters
from config import TELEGRAM_TOKEN
from handlers.commands import start, ayuda, limpiar, ejemplo
from handlers.messages import responder_mensaje


def main() -> None:
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("ayuda", ayuda))
    app.add_handler(CommandHandler("limpiar", limpiar))
    app.add_handler(CommandHandler("ejemplo", ejemplo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, responder_mensaje))

    print("Bot iniciado. Pulsa Ctrl+C para detenerlo.")
    app.run_polling()


if __name__ == "__main__":
    main()
