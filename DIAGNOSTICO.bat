@echo off
cd /d "%~dp0"
title NEXUS — Diagnóstico
color 0A
cls
echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║         NEXUS Omega — DIAGNÓSTICO DEL SISTEMA        ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

:: ── 1. Verificar Python ──────────────────────────────────────
echo  [1/5] Verificando Python...
python --version >nul 2>&1
if %errorlevel% == 0 (
    for /f "delims=" %%v in ('python --version 2^>^&1') do echo        OK - %%v
) else (
    py --version >nul 2>&1
    if %errorlevel% == 0 (
        for /f "delims=" %%v in ('py --version 2^>^&1') do echo        OK - %%v ^(usando py^)
    ) else (
        echo        ERROR - Python NO encontrado
        echo        Solucion: elimina el archivo .nexus_listo y vuelve a ejecutar NEXUS_INICIAR.bat
    )
)

:: ── 2. Verificar FastAPI/uvicorn ─────────────────────────────
echo.
echo  [2/5] Verificando paquetes instalados...
python -c "import fastapi; print('        OK - FastAPI', fastapi.__version__)" 2>nul || (
    py -c "import fastapi; print('        OK - FastAPI', fastapi.__version__)" 2>nul || echo        ERROR - FastAPI no instalado
)
python -c "import uvicorn; print('        OK - Uvicorn', uvicorn.__version__)" 2>nul || (
    py -c "import uvicorn; print('        OK - Uvicorn', uvicorn.__version__)" 2>nul || echo        ERROR - Uvicorn no instalado
)
python -c "import pandas; print('        OK - Pandas', pandas.__version__)" 2>nul || (
    py -c "import pandas; print('        OK - Pandas', pandas.__version__)" 2>nul || echo        ERROR - Pandas no instalado
)
python -c "import pdfplumber; print('        OK - pdfplumber')" 2>nul || (
    py -c "import pdfplumber; print('        OK - pdfplumber')" 2>nul || echo        ERROR - pdfplumber no instalado
)
python -c "import docx; print('        OK - python-docx')" 2>nul || (
    py -c "import docx; print('        OK - python-docx')" 2>nul || echo        ERROR - python-docx no instalado
)
python -c "import cv2; print('        OK - OpenCV', cv2.__version__)" 2>nul || (
    py -c "import cv2; print('        OK - OpenCV', cv2.__version__)" 2>nul || echo        AVISO - OpenCV no instalado ^(OCR desactivado^)
)
python -c "import pytesseract; print('        OK - pytesseract')" 2>nul || (
    py -c "import pytesseract; print('        OK - pytesseract')" 2>nul || echo        AVISO - pytesseract no instalado ^(OCR desactivado^)
)

:: ── 3. Verificar archivos del sistema ─────────────────────────
echo.
echo  [3/5] Verificando archivos del sistema...
if exist "main_server_v3.py"  (echo        OK - main_server_v3.py) else (echo        ERROR - main_server_v3.py NO encontrado)
if exist "panel.html"         (echo        OK - panel.html)        else (echo        ERROR - panel.html NO encontrado)
if exist "nexus_tray.py"      (echo        OK - nexus_tray.py)     else (echo        ERROR - nexus_tray.py NO encontrado)
if exist "nexus_setup.py"     (echo        OK - nexus_setup.py)    else (echo        ERROR - nexus_setup.py NO encontrado)
if exist ".nexus_listo"       (echo        OK - instalacion marcada como completa) else (echo        AVISO - .nexus_listo no encontrado ^(instalacion pendiente^))
if exist ".nexus_config"      (echo        OK - .nexus_config existe) else (echo        AVISO - .nexus_config no encontrado)

:: ── 4. Verificar Tesseract ────────────────────────────────────
echo.
echo  [4/5] Verificando Tesseract OCR...
if exist "tesseract\tesseract.exe" (
    echo        OK - Tesseract portable encontrado
) else (
    echo        AVISO - Tesseract NO instalado
    echo        Para instalarlo: elimina .nexus_listo y vuelve a ejecutar NEXUS_INICIAR.bat
)

:: ── 5. Intentar arrancar servidor de prueba ────────────────────
echo.
echo  [5/5] Intentando arrancar el servidor...
python -c "import fastapi, uvicorn, pandas, pdfplumber, docx; print('        OK - Todos los paquetes principales disponibles')" 2>nul || (
    py -c "import fastapi, uvicorn, pandas, pdfplumber, docx; print('        OK - Todos los paquetes principales disponibles')" 2>nul || (
        echo        ERROR - Faltan paquetes. Ejecuta REPARAR.bat para instalarlos
    )
)

echo.
echo  ══════════════════════════════════════════════════════
echo  Si ves ERROR arriba: ejecuta REPARAR.bat
echo  Si todo dice OK pero el sistema no abre:
echo    1. Cierra Firefox y Edge
echo    2. Ejecuta NEXUS_DETENER.bat
echo    3. Ejecuta NEXUS_INICIAR.bat de nuevo
echo  ══════════════════════════════════════════════════════
echo.
pause
