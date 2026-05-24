#!/bin/bash
# ============================================================
# MODO EMPRESA — Arranque manual (servidor Linux)
#
# En PRODUCCIÓN usa el servicio systemd, no este script.
# Este script es para pruebas y mantenimiento.
#
# - Bot Telegram en modo polling (sin abrir puertos)
# - API REST en puerto 8000 (túnel Cloudflare lo expone como HTTPS)
# - Add-in servido como archivos estáticos desde /excel-addin/dist
# ============================================================

BASE="$(cd "$(dirname "$0")/.." && pwd)"
cd "$BASE"

if [ ! -d ".venv" ]; then
    echo "ERROR: Entorno virtual no encontrado."
    echo "Ejecuta primero: scripts/instalar_empresa.sh"
    exit 1
fi

source .venv/bin/activate
echo ""
echo "Arrancando Asistente Excel [modo empresa - polling + API REST]..."
echo "  Bot Telegram: polling (sin puertos entrantes)"
echo "  API Add-in:   puerto ${PORT:-8000} (tunelado por cloudflared)"
echo "  Pulsa Ctrl+C para detener"
echo ""
python api.py
