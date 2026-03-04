@echo off
cd /d "%~dp0"
title NEXUS — Ver Error del Servidor
color 0C
cls
echo.
echo  Iniciando servidor en modo visible para ver errores...
echo  NO cierre esta ventana.
echo.
echo  ─────────────────────────────────────────────────
echo.

:: Detectar Python
set PYTHON_CMD=python
if exist ".nexus_config" (
    for /f "tokens=2 delims==" %%a in ('findstr "python_cmd" ".nexus_config"') do set PYTHON_CMD=%%a
)

:: Correr el servidor con errores visibles (sin redirigir stderr)
%PYTHON_CMD% main_server_v3.py

echo.
echo  ─────────────────────────────────────────────────
echo  El servidor se detuvo. Copia el texto rojo de arriba
echo  y mandalo para saber que arreglar.
echo  ─────────────────────────────────────────────────
pause
