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
function Write-Step($msg)  { Write-Host "`n--- $msg"   -ForegroundColor Cyan   }

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

# ── 3. Determinar la ruta del catalogo ───────────────────────────────────────
# Office acepta como catalogo de confianza:
#   - Ruta UNC:   \\NOMBRE_EQUIPO\recurso_compartido
#   - URL HTTPS:  https://servidor/carpeta
#   NO acepta rutas locales C:\ — no intentarlo nunca como fallback.
#
# Prioridad:
#   1. URL HTTPS (si el manifest viene de un servidor HTTPS)  ← siempre disponible con Render
#   2. Recurso SMB local (si el servidor HTTPS no esta disponible)
Write-Step "Preparando ruta del catalogo"
$rutaCatalogo = $null

# Opcion 1: URL HTTPS del servidor (preferida y siempre valida)
if ($urlManifest -match "^https://") {
    # Quitar el nombre del archivo; usar solo la carpeta base del servidor
    $rutaCatalogo = ($urlManifest -replace "/manifest\.xml$", "").TrimEnd("/")
    Write-Ok "Catalogo HTTPS: $rutaCatalogo"
}

# Opcion 2: Recurso SMB local (solo si no hay servidor HTTPS configurado)
if (-not $rutaCatalogo) {
    try {
        $shareExiste = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
        if ($shareExiste) {
            $rutaCatalogo = "\\$env:COMPUTERNAME\$shareName"
            Write-Ok "Recurso compartido existente: $rutaCatalogo"
        } else {
            New-SmbShare -Name $shareName -Path $carpeta -ReadAccess "Everyone" -ErrorAction Stop | Out-Null
            $rutaCatalogo = "\\$env:COMPUTERNAME\$shareName"
            Write-Ok "Carpeta compartida: $rutaCatalogo"
        }
    } catch {
        Write-Fail "No se pudo crear el recurso compartido."
        Write-Fail "Para usar un catalogo local, un administrador de IT debe compartir la carpeta $carpeta"
        Write-Fail "o configurar un servidor HTTPS con el manifest.xml."
        Read-Host "`nPulsa Intro para salir"
        exit 1
    }
}

# ── 4. Registrar catalogo en el Centro de confianza de Excel ─────────────────
Write-Step "Registrando en el Centro de confianza de Excel"
Write-Host ""
Write-Host "  [A] Automatico — el instalador modifica el registro de Windows" -ForegroundColor Green
Write-Host "  [M] Manual     — me indicas los pasos en Excel" -ForegroundColor Yellow
Write-Host "                   (si el antivirus bloquea la opcion A)" -ForegroundColor DarkYellow
Write-Host ""

do {
    $opcion = (Read-Host "  Elige una opcion [A/M]").Trim().ToUpper()
} while ($opcion -ne "A" -and $opcion -ne "M")

$registrado = $false

if ($opcion -eq "A") {

    # ── 4a. Registro automatico ───────────────────────────────────────────────
    foreach ($ver in $officeVersions) {
        $basePath = "HKCU:\Software\Microsoft\Office\$ver"
        if (-not (Test-Path $basePath)) { continue }

        $catalogPath = "$basePath\WEF\TrustedCatalogs"
        if (-not (Test-Path $catalogPath)) {
            New-Item -Path $catalogPath -Force | Out-Null
        }

        # Comprobar si ya esta registrado
        $yaExiste = Get-ChildItem $catalogPath -ErrorAction SilentlyContinue | Where-Object {
            (Get-ItemProperty $_.PSPath -Name "Url" -ErrorAction SilentlyContinue).Url -eq $rutaCatalogo
        }

        if ($yaExiste) {
            Write-Ok "Catalogo ya registrado en Office $ver"
            $registrado = $true
            continue
        }

        try {
            $guid      = [System.Guid]::NewGuid().ToString("B").ToUpper()
            $entryPath = "$catalogPath\$guid"
            New-Item -Path $entryPath -Force | Out-Null
            Set-ItemProperty -Path $entryPath -Name "Id"    -Value $guid
            Set-ItemProperty -Path $entryPath -Name "Url"   -Value $rutaCatalogo
            Set-ItemProperty -Path $entryPath -Name "Flags" -Value 1 -Type DWord
            Write-Ok "Catalogo registrado en Office $ver"
            $registrado = $true
        } catch {
            Write-Warn "No se pudo escribir en el registro para Office $ver"
        }
    }

    if (-not $registrado) {
        Write-Warn "No se pudo registrar automaticamente."
        Write-Warn "Sigue los pasos manuales que aparecen a continuacion."
        $opcion = "M"   # caer al bloque manual con la ruta correcta
    }
}

if ($opcion -eq "M") {

    # ── 4b. Instrucciones manuales ────────────────────────────────────────────
    Write-Host ""
    Write-Host "  Sigue estos pasos en Excel:" -ForegroundColor White
    Write-Host ""
    Write-Host "  1. Abre Excel" -ForegroundColor White
    Write-Host "     Archivo  >  Opciones  >  Centro de confianza" -ForegroundColor Yellow
    Write-Host "     >  Configuracion del Centro de confianza" -ForegroundColor Yellow
    Write-Host "     >  Catalogos de complementos de confianza" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  2. En 'URL del catalogo' copia exactamente esta direccion:" -ForegroundColor White
    Write-Host ""
    Write-Host "        $rutaCatalogo" -ForegroundColor Cyan
    Write-Host ""
    if ($rutaCatalogo -match "^https://") {
        Write-Host "     (es una URL https — el campo acepta esta direccion directamente)" -ForegroundColor DarkGray
    } else {
        Write-Host "     (es una ruta de red — debe empezar por \\ para ser valida)" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "  3. Clic en 'Agregar catalogo'" -ForegroundColor White
    Write-Host "     Marca 'Mostrar en menu'  >  Aceptar" -ForegroundColor White
    Write-Host ""
    $registrado = $true   # el usuario lo hara manualmente
}

# ── 5. Resultado final ────────────────────────────────────────────────────────
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
if ($registrado) {
    Write-Host "   Instalacion completada                  " -ForegroundColor Cyan
} else {
    Write-Host "   Complemento descargado (registro manual)" -ForegroundColor Yellow
}
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
Write-Host "  Ruta del catalogo usada: $rutaCatalogo" -ForegroundColor DarkGray
Write-Host ""
Read-Host "Pulsa Intro para cerrar"
