@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
title NEXUS Omega

if not exist ".nexus_listo" (
    cls
    echo.
    echo  Preparando NEXUS por primera vez, por favor espere...
    echo  Esto tardara varios minutos. NO cierre esta ventana.
    echo.
    wscript //nologo "_instalar.vbs"
    if not exist ".nexus_listo" (
        echo.
        echo  ERROR: No se pudo completar la instalacion.
        echo  Contacta al administrador del sistema.
        pause
        exit /b 1
    )
    timeout /t 2 /nobreak >nul
)

set PYTHON_CMD=python
if exist ".nexus_config" (
    for /f "tokens=2 delims==" %%a in ('findstr "python_cmd" ".nexus_config"') do set PYTHON_CMD=%%a
)

:: Detectar si curl está disponible, si no usar PowerShell
where curl >nul 2>&1
if %errorlevel% == 0 (
    set "CHECK=curl -s --max-time 1 http://127.0.0.1:8000 >nul 2>&1"
) else (
    set "CHECK=powershell -Command "try{$null=Invoke-WebRequest -Uri http://127.0.0.1:8000 -TimeoutSec 1 -UseBasicParsing;exit 0}catch{exit 1}""
)

%CHECK%
if %errorlevel% == 0 (
    start "" "http://127.0.0.1:8000"
    goto :fin
)

wscript //nologo "_lanzar_servidor.vbs"

cls
echo.
echo  Iniciando NEXUS Omega, por favor espere...
echo.
set intentos=0
:esperar
timeout /t 1 /nobreak >nul
set /a intentos+=1
%CHECK%
if %errorlevel% == 0 goto :abrir
if %intentos% lss 30 goto :esperar

:abrir
start "" "http://127.0.0.1:8000"

:fin
timeout /t 2 /nobreak >nul
exit /b 0
