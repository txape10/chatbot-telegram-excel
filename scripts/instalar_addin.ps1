# ==============================================================================
# Instalador automatico del Add-in "Asistente Excel"
# Ejecutado por instalar_addin.bat con privilegios de administrador
# ==============================================================================

$carpeta     = "C:\Complementos\AsistenteExcel"
$destino     = "$carpeta\manifest.xml"
$urlManifest = "https://asistente-excel.onrender.com/manifest.xml"
$shareName   = "AsistenteExcel"

# Versiones de Office a registrar (16.0 = 2016/2019/2021/365)
$officeVersions = @("16.0", "15.0")

function Write-Ok($msg)    { Write-Host "[OK] $msg"    -ForegroundColor Green  }
function Write-Warn($msg)  { Write-Host "[AVISO] $msg" -ForegroundColor Yellow }
function Write-Fail($msg)  { Write-Host "[ERROR] $msg" -ForegroundColor Red    }
function Write-Step($msg)  { Write-Host "`n--- $msg" -ForegroundColor Cyan    }

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   Instalador -- Asistente Excel Add-in   " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# ── 1. Crear carpeta ──────────────────────────────────────────────────────────
Write-Step "Creando carpeta"
if (-not (Test-Path $carpeta)) {
    New-Item -ItemType Directory -Path $carpeta -Force | Out-Null
    Write-Ok "Carpeta creada: $carpeta"
} else {
    Write-Ok "Carpeta ya existe: $carpeta"
}

# ── 2. Descargar manifest.xml ─────────────────────────────────────────────────
Write-Step "Descargando complemento desde Render"
try {
    Invoke-WebRequest -Uri $urlManifest -OutFile $destino -UseBasicParsing -ErrorAction Stop
    Write-Ok "manifest.xml descargado correctamente"
} catch {
    Write-Fail "No se pudo descargar: $_"
    Write-Host "`nComprueba que tienes conexion a Internet e intentalo de nuevo." -ForegroundColor Red
    Read-Host "`nPulsa Intro para salir"
    exit 1
}

# ── 3. Compartir la carpeta como recurso de red ───────────────────────────────
Write-Step "Compartiendo carpeta en red"
$rutaCatalogo = $carpeta   # fallback: ruta local

try {
    $shareExiste = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
    if ($shareExiste) {
        Write-Ok "Recurso compartido ya existe"
        $rutaCatalogo = "\\$env:COMPUTERNAME\$shareName"
    } else {
        New-SmbShare -Name $shareName -Path $carpeta -ReadAccess "Everyone" -ErrorAction Stop | Out-Null
        Write-Ok "Carpeta compartida como \\$env:COMPUTERNAME\$shareName"
        $rutaCatalogo = "\\$env:COMPUTERNAME\$shareName"
    }
} catch {
    Write-Warn "No se pudo compartir automaticamente. Se usara la ruta local."
    $rutaCatalogo = $carpeta
}

# ── 4. Centro de confianza — preguntar al usuario ─────────────────────────────
Write-Step "Centro de confianza de Excel"
Write-Host ""
Write-Host "  El instalador puede registrar el complemento automaticamente" -ForegroundColor White
Write-Host "  modificando el registro de Windows, o puedes hacerlo tu mismo" -ForegroundColor White
Write-Host "  desde los menus de Excel si lo prefieres." -ForegroundColor White
Write-Host ""
Write-Host "  [A] Automatico — el instalador lo configura solo (recomendado)" -ForegroundColor Green
Write-Host "  [M] Manual     — me indicas los pasos a seguir en Excel" -ForegroundColor Yellow
Write-Host ""

do {
    $opcion = (Read-Host "  Elige una opcion [A/M]").Trim().ToUpper()
} while ($opcion -ne "A" -and $opcion -ne "M")

if ($opcion -eq "A") {

    # ── 4a. Registro automatico ───────────────────────────────────────────────
    $registrado = $false
    foreach ($ver in $officeVersions) {
        $basePath = "HKCU:\Software\Microsoft\Office\$ver"
        if (-not (Test-Path $basePath)) { continue }

        $catalogPath = "$basePath\WEF\TrustedCatalogs"
        if (-not (Test-Path $catalogPath)) {
            New-Item -Path $catalogPath -Force | Out-Null
        }

        # Comprobar si ya esta registrado con esta URL
        $yaExiste = Get-ChildItem $catalogPath -ErrorAction SilentlyContinue | Where-Object {
            (Get-ItemProperty $_.PSPath -Name "Url" -ErrorAction SilentlyContinue).Url -eq $rutaCatalogo
        }

        if ($yaExiste) {
            Write-Ok "Catalogo ya registrado en Office $ver"
            $registrado = $true
            continue
        }

        # Crear nueva entrada con GUID unico
        $guid      = [System.Guid]::NewGuid().ToString("B").ToUpper()
        $entryPath = "$catalogPath\$guid"
        New-Item -Path $entryPath -Force | Out-Null
        Set-ItemProperty -Path $entryPath -Name "Id"    -Value $guid
        Set-ItemProperty -Path $entryPath -Name "Url"   -Value $rutaCatalogo
        Set-ItemProperty -Path $entryPath -Name "Flags" -Value 1 -Type DWord

        Write-Ok "Catalogo registrado en Office $ver"
        $registrado = $true
    }

    if (-not $registrado) {
        Write-Warn "No se encontro ninguna instalacion de Office en el registro."
        Write-Warn "Usa la opcion manual la proxima vez."
    }

} else {

    # ── 4b. Instrucciones manuales ────────────────────────────────────────────
    Write-Host ""
    Write-Host "  Sigue estos pasos en Excel:" -ForegroundColor White
    Write-Host ""
    Write-Host "  1. Abre Excel y ve a:" -ForegroundColor White
    Write-Host "        Archivo  >  Opciones  >  Centro de confianza" -ForegroundColor Yellow
    Write-Host "        >  Configuracion del Centro de confianza" -ForegroundColor Yellow
    Write-Host "        >  Catalogos de complementos de confianza" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  2. En el campo 'URL del catalogo' escribe exactamente:" -ForegroundColor White
    Write-Host "        $rutaCatalogo" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  3. Haz clic en 'Agregar catalogo'" -ForegroundColor White
    Write-Host "     Marca la casilla 'Mostrar en menu'  >  Aceptar" -ForegroundColor White
    Write-Host ""

}

# ── 5. Resultado final ────────────────────────────────────────────────────────
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   Instalacion completada correctamente   " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ultimo paso:" -ForegroundColor White
Write-Host ""
Write-Host "  1. CIERRA Excel completamente si esta abierto" -ForegroundColor Yellow
Write-Host "  2. Vuelve a abrir Excel" -ForegroundColor Yellow
Write-Host "  3. Insertar  >  Mis complementos  >  Mi organizacion" -ForegroundColor Yellow
Write-Host "     >  Asistente Excel  >  Agregar" -ForegroundColor Yellow
Write-Host ""
Write-Host "  El boton 'Abrir asistente' aparecera en la pestana Inicio." -ForegroundColor Green
Write-Host ""
Read-Host "Pulsa Intro para cerrar"
