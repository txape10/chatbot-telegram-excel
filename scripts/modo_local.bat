@echo off
echo.
echo ========================================
echo   Cambiando a modo LOCAL (localhost:8000)
echo ========================================
echo.

set PROJECT=%~dp0..
set MANIFEST_DEST=C:\Complementos\AsistenteExcel\manifest.xml
set MANIFEST_LOCAL=%PROJECT%\excel-addin\manifest.xml

:: Verificar que el backend no está ya corriendo en 8000
netstat -ano | findstr ":8000" >nul 2>&1
if not errorlevel 1 (
    echo AVISO: El puerto 8000 ya está en uso. Cierra el backend anterior primero.
    pause
    exit /b 1
)

:: Instalar dependencias Python si faltan
echo Verificando dependencias Python...
cd /d "%PROJECT%"
pip install -r requirements.txt -q
echo [OK] Dependencias Python OK

:: Compilar el Add-in apuntando a localhost:8000
echo Compilando Add-in para localhost:8000...
cd /d "%PROJECT%\excel-addin"
set ADDIN_URL=http://localhost:8000/
call npm run build
if errorlevel 1 (
    echo ERROR: Fallo al compilar el Add-in.
    pause
    exit /b 1
)
echo [OK] Add-in compilado

:: Actualizar manifest local con localhost:8000
powershell -Command "(Get-Content '%MANIFEST_LOCAL%') -replace 'https?://localhost:\d+', 'http://localhost:8000' | Set-Content '%MANIFEST_LOCAL%' -Encoding utf8"
copy /Y "%MANIFEST_LOCAL%" "%MANIFEST_DEST%" >nul
echo [OK] Manifest actualizado: localhost:8000

:: Limpiar cache de Office
powershell -Command "Get-ChildItem '$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef' -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue"
echo [OK] Cache de Office limpiada

echo.
echo Iniciando backend (FastAPI en http://localhost:8000)...
echo Cuando arranque, abre Excel y recarga el panel del complemento.
echo.
python api.py
