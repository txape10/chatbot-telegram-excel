# ==============================================================================
# Instalador del Add-in "Asistente Excel"
# ==============================================================================
# Ejecutar en PowerShell como Administrador:
#   powershell -ExecutionPolicy Bypass -c "iex (iwr 'https://raw.githubusercontent.com/txape10/chatbot-telegram-excel/main/scripts/instalar_addin.ps1' -UseBasicParsing).Content"
# ==============================================================================

$carpeta    = "C:\Complementos\AsistenteExcel"
$destino    = "$carpeta\manifest.xml"
$urlManifest = "https://asistente-excel.onrender.com/manifest.xml"
$nombreShare = "AsistenteExcel"

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   Instalador -- Asistente Excel Add-in    " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# ── 1. Crear carpeta ──────────────────────────────────────────────────────────
if (-not (Test-Path $carpeta)) {
    New-Item -ItemType Directory -Path $carpeta -Force | Out-Null
    Write-Host "[OK] Carpeta creada: $carpeta" -ForegroundColor Green
} else {
    Write-Host "[OK] Carpeta ya existe: $carpeta" -ForegroundColor Green
}

# ── 2. Descargar manifest.xml ─────────────────────────────────────────────────
Write-Host "     Descargando manifest.xml desde Render..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $urlManifest -OutFile $destino -UseBasicParsing -ErrorAction Stop
    Write-Host "[OK] Archivo descargado correctamente" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] No se pudo descargar el archivo: $_" -ForegroundColor Red
    Write-Host "        Comprueba que tienes conexion a Internet e intentalo de nuevo." -ForegroundColor Red
    Read-Host "Pulsa Intro para salir"
    exit 1
}

# ── 3. Compartir la carpeta como recurso de red ───────────────────────────────
$rutaCatalogo = $carpeta   # ruta local como fallback

try {
    $shareExiste = Get-SmbShare -Name $nombreShare -ErrorAction SilentlyContinue
    if ($shareExiste) {
        Write-Host "[OK] Recurso compartido ya existe: \\$env:COMPUTERNAME\$nombreShare" -ForegroundColor Green
        $rutaCatalogo = "\\$env:COMPUTERNAME\$nombreShare"
    } else {
        New-SmbShare -Name $nombreShare -Path $carpeta -ReadAccess "Everyone" -ErrorAction Stop | Out-Null
        Write-Host "[OK] Carpeta compartida: \\$env:COMPUTERNAME\$nombreShare" -ForegroundColor Green
        $rutaCatalogo = "\\$env:COMPUTERNAME\$nombreShare"
    }
} catch {
    Write-Host "[AVISO] No se pudo compartir la carpeta automaticamente." -ForegroundColor Yellow
    Write-Host "        Se usara la ruta local: $carpeta" -ForegroundColor Yellow
    Write-Host "        (Esto funciona igualmente si solo lo usas tu en este PC)" -ForegroundColor Yellow
}

# ── 4. Instrucciones finales ──────────────────────────────────────────────────
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   Instalacion completada. Pasos en Excel  " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Abre Excel y ve a:" -ForegroundColor White
Write-Host "     Archivo > Opciones > Centro de confianza" -ForegroundColor White
Write-Host "     > Configuracion del Centro de confianza" -ForegroundColor White
Write-Host "     > Catalogos de complementos de confianza" -ForegroundColor White
Write-Host ""
Write-Host "  2. En 'URL del catalogo' escribe exactamente:" -ForegroundColor White
Write-Host "     $rutaCatalogo" -ForegroundColor Yellow
Write-Host ""
Write-Host "  3. Haz clic en 'Agregar catalogo'" -ForegroundColor White
Write-Host "     Marca 'Mostrar en menu' > Aceptar" -ForegroundColor White
Write-Host ""
Write-Host "  4. CIERRA Excel y vuelve a abrirlo" -ForegroundColor White
Write-Host ""
Write-Host "  5. Insertar > Mis complementos > Mi organizacion" -ForegroundColor White
Write-Host "     > Asistente Excel > Agregar" -ForegroundColor White
Write-Host ""
Write-Host "  El boton 'Abrir asistente' aparecera en la pestana Inicio." -ForegroundColor Green
Write-Host ""
Read-Host "Pulsa Intro para cerrar"
