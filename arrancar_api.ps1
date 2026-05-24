# Libera el puerto 8000 si está ocupado y arranca la API
Get-NetTCPConnection -LocalPort 8000 -ErrorAction SilentlyContinue |
    ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }

python api.py
