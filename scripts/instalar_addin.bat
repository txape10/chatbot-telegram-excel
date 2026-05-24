@echo off
:: Instalador del Add-in "Asistente Excel"
:: Doble clic > Aceptar en UAC > listo

set SCRIPT=%TEMP%\instalar_addin_excel.ps1

:: Descargar el script de instalacion
powershell -NoProfile -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/txape10/chatbot-telegram-excel/main/scripts/instalar_addin.ps1','%SCRIPT%')"

:: Ejecutar el script con privilegios de administrador
powershell -NoProfile -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%SCRIPT%""' -Verb RunAs -Wait"

:: Limpiar el script temporal
del "%SCRIPT%" 2>nul
