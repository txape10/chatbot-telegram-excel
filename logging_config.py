import logging
import os
from logging.handlers import RotatingFileHandler

LOG_DIR  = os.path.join(os.path.dirname(__file__), "data", "logs")
LOG_FILE = os.path.join(LOG_DIR, "bot.log")

_FORMATO   = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
_FECHA_FMT = "%Y-%m-%d %H:%M:%S"

# LOG_LEVEL=DEBUG en .env activa logs de diagnóstico (respuestas LLM completas, etc.)
_NIVEL = getattr(logging, os.getenv("LOG_LEVEL", "INFO").upper(), logging.INFO)


def configurar_logging() -> None:
    """Configura logging a consola y a fichero rotativo. Llamar una vez al arrancar."""
    os.makedirs(LOG_DIR, exist_ok=True)

    # Handler de fichero: máx. 5 MB × 3 backups
    file_handler = RotatingFileHandler(
        LOG_FILE,
        maxBytes=5 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
    file_handler.setLevel(_NIVEL)
    file_handler.setFormatter(logging.Formatter(_FORMATO, datefmt=_FECHA_FMT))

    # Handler de consola
    console_handler = logging.StreamHandler()
    console_handler.setLevel(_NIVEL)
    console_handler.setFormatter(logging.Formatter(_FORMATO, datefmt=_FECHA_FMT))

    # Logger raíz
    root = logging.getLogger()
    root.setLevel(_NIVEL)
    root.addHandler(file_handler)
    root.addHandler(console_handler)

    # Silenciar librerías muy verbosas
    for nombre in ("httpx", "httpcore", "telegram.ext.Application"):
        logging.getLogger(nombre).setLevel(logging.WARNING)

    logging.getLogger(__name__).info("Logging configurado → %s (nivel: %s)",
                                     LOG_FILE, logging.getLevelName(_NIVEL))
