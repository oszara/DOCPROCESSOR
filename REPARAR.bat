@echo off
cd /d "%~dp0"
title NEXUS — Reparación
color 0E
cls
echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║         NEXUS Omega — REPARACIÓN AUTOMÁTICA          ║
echo  ║  NO CIERRE ESTA VENTANA hasta que termine            ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

:: ── Detectar Python ──────────────────────────────────────────
set PYTHON_CMD=
python --version >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON_CMD=python
    goto :python_ok
)
py --version >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON_CMD=py
    goto :python_ok
)

:: Python no encontrado — instalar
echo  Python no encontrado. Descargando...
echo  (Esto puede tardar 5-10 minutos segun tu internet)
echo.
:: Descargar con PowerShell (disponible en todos los Windows modernos)
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe' -OutFile '_python_setup.exe'"
if not exist "_python_setup.exe" (
    echo  ERROR: No se pudo descargar Python.
    echo  Verifica tu conexion a internet e intenta de nuevo.
    pause
    exit /b 1
)
echo  Instalando Python...
_python_setup.exe /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1 Include_test=0
del /f /q "_python_setup.exe" >nul 2>&1
set PYTHON_CMD=python
echo  OK - Python instalado.

:python_ok
echo  Usando: %PYTHON_CMD%
echo.

:: ── Instalar/actualizar pip ────────────────────────────────────
echo  Actualizando pip...
%PYTHON_CMD% -m pip install --upgrade pip --quiet --disable-pip-version-check
echo.

:: ── Instalar paquetes principales ─────────────────────────────
echo  Instalando paquetes del sistema...
echo  (Puede tardar 3-5 minutos)
echo.

set PKGS=fastapi "uvicorn[standard]" pandas openpyxl python-docx pdfplumber python-multipart pystray pillow

for %%p in (%PKGS%) do (
    echo  Instalando %%p ...
    %PYTHON_CMD% -m pip install %%p --quiet --disable-pip-version-check
    if %errorlevel% == 0 (echo     OK) else (echo     ADVERTENCIA - puede no haberse instalado correctamente)
)

:: ── Instalar paquetes OCR (opcionales, no bloquean si fallan) ──
echo.
echo  Instalando componentes OCR (opcionales)...
%PYTHON_CMD% -m pip install pytesseract --quiet --disable-pip-version-check
%PYTHON_CMD% -m pip install opencv-python --quiet --disable-pip-version-check
if %errorlevel% == 0 (echo     OK - OCR disponible) else (echo     AVISO - OCR no instalado, el sistema funciona sin el)

:: ── Guardar configuración ─────────────────────────────────────
echo.
echo  Guardando configuracion...
echo python_cmd=%PYTHON_CMD%> .nexus_config
echo ok=1> .nexus_listo

:: ── Verificación final ────────────────────────────────────────
echo.
echo  Verificando instalacion...
%PYTHON_CMD% -c "import fastapi, uvicorn, pandas, pdfplumber, docx; print('  OK - Sistema listo')" 2>nul
if %errorlevel% neq 0 (
    echo  ADVERTENCIA: Algunos paquetes no se instalaron correctamente.
    echo  Intenta ejecutar este archivo de nuevo.
)

echo.
echo  ══════════════════════════════════════════════════════
echo  Reparacion completada.
echo  Ahora ejecuta NEXUS_INICIAR.bat para abrir el sistema.
echo  ══════════════════════════════════════════════════════
echo.
pause
