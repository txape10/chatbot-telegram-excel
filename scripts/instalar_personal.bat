@echo off
:: ============================================================
:: MODO PERSONAL — Instalación inicial (Windows)
:: Ejecutar UNA SOLA VEZ antes de usar el bot
:: ============================================================

cd /d "%~dp0.."

echo.
echo === Asistente Excel — Instalacion personal (Windows) ===
echo.

:: Verificar que existe el .env
if not exist .env (
    echo ERROR: No se encuentra el archivo .env
    echo Copia .env.example como .env y rellena tus tokens.
    pause
    exit /b 1
)

echo [1/2] Creando entorno virtual Python...
if not exist .venv (
    python -m venv .venv
) else (
    echo      (ya existe, omitiendo)
)

echo [2/2] Instalando dependencias...
call .venv\Scripts\activate
pip install -r requirements.txt -q

echo.
echo === Instalacion completada ===
echo Para arrancar el bot ejecuta:  scripts\arrancar_personal.bat
echo.
pause
