#!/bin/bash
# ============================================================
# MODO EMPRESA — Instalación inicial (servidor Linux)
# Ejecutar UNA SOLA VEZ en el servidor, como usuario del servicio
#
# Prereqs (instalar con apt antes de este script):
#   sudo apt install python3.11 python3.11-venv python3-pip git nodejs npm
# ============================================================

set -e
BASE="$(cd "$(dirname "$0")/.." && pwd)"
cd "$BASE"

echo ""
echo "=== Asistente Excel — Instalación servidor empresa ==="
echo ""

# Verificar .env
if [ ! -f ".env" ]; then
    echo "ERROR: No se encuentra el archivo .env"
    echo "Copia .env.example como .env y rellena tus tokens."
    exit 1
fi

echo "[1/4] Creando entorno virtual Python..."
if [ ! -d ".venv" ]; then
    python3.11 -m venv .venv
else
    echo "      (ya existe, omitiendo)"
fi

echo "[2/4] Instalando dependencias Python..."
source .venv/bin/activate
pip install -r requirements.txt -q

echo "[3/4] Instalando dependencias del Add-in..."
cd excel-addin
npm install --silent

echo "[4/4] Compilando Add-in (webpack)..."
npm run build
cd "$BASE"

# Permisos seguros para .env
chmod 600 .env
echo "      Permisos de .env ajustados a 600 (solo lectura del propietario)"

echo ""
echo "=== Instalación completada ==="
echo ""
echo "Próximos pasos:"
echo "  1. Configura cloudflared y obtén la URL del túnel"
echo "  2. Actualiza ADDIN_URL en el .env con esa URL"
echo "  3. Vuelve a compilar el Add-in:  cd excel-addin && npm run build"
echo "  4. Instala el servicio systemd: sudo cp scripts/asistente-excel.service /etc/systemd/system/"
echo "     sudo systemctl daemon-reload && sudo systemctl enable asistente-excel"
echo "  5. Arranca: sudo systemctl start asistente-excel"
echo ""
