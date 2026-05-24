"""Punto de entrada local — arranca el bot en modo polling."""
from logging_config import configurar_logging
configurar_logging()

from telegram_app import crear_aplicacion


def main() -> None:
    app = crear_aplicacion()
    print("Bot iniciado en modo polling. Pulsa Ctrl+C para detenerlo.")
    app.run_polling()


if __name__ == "__main__":
    main()
