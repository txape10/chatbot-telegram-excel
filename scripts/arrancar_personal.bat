@echo off
:: ============================================================
:: MODO PERSONAL — Arranque diario (Windows)
::
:: - Bot Telegram en modo polling (sin abrir puertos)
:: - Requiere: .env con TELEGRAM_TOKEN, GROQ_API_KEY, AUTHORIZED_USERS
:: - El bot funciona mientras esta ventana este abierta
:: - Pulsa Ctrl+C para detener
:: ============================================================

cd /d "%~dp0.."

if not exist .venv\Scripts\activate.bat (
    echo ERROR: Entorno virtual no encontrado.
    echo Ejecuta primero: scripts\instalar_personal.bat
    pause
    exit /b 1
)

call .venv\Scripts\activate
echo.
echo Arrancando Asistente Excel [modo personal - polling]...
echo Pulsa Ctrl+C para detener.
echo.
python bot.py
