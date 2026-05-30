@echo off
echo.
echo ========================================
echo   Cambiando a modo RENDER (produccion)
echo ========================================
echo.

set PROJECT=%~dp0..
set MANIFEST_DEST=C:\Complementos\AsistenteExcel\manifest.xml
set MANIFEST_RENDER=%PROJECT%\excel-addin\manifest.render.xml

:: Recompilar el Add-in apuntando a Render (IMPORTANTE: el dist/ va al repo)
echo Compilando Add-in para Render...
cd /d "%PROJECT%\excel-addin"
set ADDIN_URL=https://asistente-excel.onrender.com/
call npm run build
if errorlevel 1 (
    echo ERROR: Fallo al compilar el Add-in.
    pause
    exit /b 1
)
echo [OK] Add-in compilado para Render

:: Restaurar manifest de Render
copy /Y "%MANIFEST_RENDER%" "%MANIFEST_DEST%" >nul
echo [OK] Manifest restaurado: asistente-excel.onrender.com

:: Limpiar cache de Office
powershell -Command "Get-ChildItem '$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef' -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue"
echo [OK] Cache de Office limpiada

echo.
echo Para desplegar en Render:
echo   git add excel-addin/dist/ services/llm.py prompts/excel.py [otros cambios]
echo   git commit -m "descripcion del cambio"
echo   git push
echo.
echo Render redespliegue automaticamente en ~1-2 minutos.
echo Recuerda recargar el panel en Excel cuando Render termine.
echo.
pause
