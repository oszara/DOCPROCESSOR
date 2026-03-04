@echo off
:: Detener NEXUS Omega
cd /d "%~dp0"
echo Deteniendo NEXUS Omega...
taskkill /F /IM python.exe /T >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq NEXUS*" >nul 2>&1
echo NEXUS detenido.
timeout /t 2 /nobreak >nul
