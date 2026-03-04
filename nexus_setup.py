"""
nexus_setup.py — Instalador automático de NEXUS Ω
Se ejecuta la primera vez y prepara todo lo necesario.
"""
import sys, os, subprocess, urllib.request, time, winreg, shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MARKER   = os.path.join(BASE_DIR, ".nexus_listo")

# ── Tesseract OCR portable ────────────────────────────────────────────────────
TESS_DIR = os.path.join(BASE_DIR, "tesseract")
TESS_ZIP = os.path.join(BASE_DIR, "_tesseract.zip")
TESS_EXE = os.path.join(TESS_DIR, "tesseract.exe")

DEPENDENCIAS = [
    "fastapi",
    "uvicorn[standard]",
    "pandas",
    "openpyxl",
    "python-docx",
    "pdfplumber",
    "python-multipart",
    "pystray",
    "pillow",
    "pytesseract",        # interfaz Python → Tesseract
    "opencv-python",      # preprocesamiento de imagen para mejor OCR
]

PYTHON_INSTALLER_URL = (
    "https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe"
)
PYTHON_INSTALLER_LOCAL = os.path.join(BASE_DIR, "_python_installer.exe")


def cls():
    os.system("cls" if os.name == "nt" else "clear")


def titulo(msg: str):
    print("=" * 60)
    print(f"  {msg}")
    print("=" * 60)


def paso(n: int, total: int, msg: str):
    bar_len = 30
    filled  = int(bar_len * n / total)
    bar     = "█" * filled + "░" * (bar_len - filled)
    print(f"\n  [{bar}]  Paso {n}/{total}")
    print(f"  ▶ {msg}")


def python_en_path() -> str | None:
    """Devuelve el ejecutable de Python si está disponible, o None."""
    for cmd in ["python", "python3", "py"]:
        try:
            r = subprocess.run(
                [cmd, "--version"],
                capture_output=True, text=True, timeout=5
            )
            if r.returncode == 0 and "Python 3" in r.stdout + r.stderr:
                return cmd
        except Exception:
            pass
    return None


def descargar_python():
    print("\n  Descargando Python 3.12 (puede tardar unos minutos)...")
    print("  Por favor no cierre esta ventana.\n")
    try:
        urllib.request.urlretrieve(
            PYTHON_INSTALLER_URL, PYTHON_INSTALLER_LOCAL,
            reporthook=lambda b, bs, ts: print(
                f"  {min(100, int(b * bs * 100 / ts))}%   ", end="\r"
            ) if ts > 0 else None
        )
        print("\n  Descarga completa.")
        return True
    except Exception as e:
        print(f"\n  ERROR al descargar Python: {e}")
        return False


def instalar_python():
    """Instala Python silenciosamente con pip incluido."""
    print("\n  Instalando Python 3.12 en tu computadora...")
    print("  (Se instalará solo para este usuario, sin necesitar permisos de administrador)")
    args = [
        PYTHON_INSTALLER_LOCAL,
        "/quiet",
        "InstallAllUsers=0",      # solo este usuario
        "PrependPath=1",           # agregar al PATH automáticamente
        "Include_pip=1",
        "Include_launcher=1",
        "Include_tcltk=0",
        "Include_test=0",
    ]
    r = subprocess.run(args, timeout=300)
    if r.returncode == 0:
        print("  ✓ Python instalado correctamente")
        # Limpiar instalador
        try:
            os.remove(PYTHON_INSTALLER_LOCAL)
        except Exception:
            pass
        return True
    else:
        print(f"  ERROR al instalar Python (código {r.returncode})")
        return False


def instalar_dependencias(python_cmd: str):
    print()
    total = len(DEPENDENCIAS)
    for i, paquete in enumerate(DEPENDENCIAS, 1):
        print(f"  Instalando {paquete}...  ({i}/{total})", end=" ", flush=True)
        try:
            r = subprocess.run(
                [python_cmd, "-m", "pip", "install", paquete,
                 "--quiet", "--disable-pip-version-check",
                 "--no-warn-script-location"],
                capture_output=True, text=True, timeout=120
            )
            if r.returncode == 0:
                print("✓")
            else:
                print(f"⚠ (advertencia)")
        except subprocess.TimeoutExpired:
            print("⚠ tiempo agotado, continuando...")
        except Exception as e:
            print(f"⚠ {e}")


# URLs alternativas — se intenta en orden hasta que una funcione
TESS_URLS = [
    # ZIP portable UB-Mannheim (GitHub)
    "https://github.com/UB-Mannheim/tesseract/releases/download/"
    "v5.3.3.20231005/tesseract-ocr-w64-portable-5.3.3.20231005.zip",
    # Espejo alternativo
    "https://digi.bib.uni-mannheim.de/tesseract/"
    "tesseract-ocr-w64-portable-5.3.3.20231005.zip",
]


def _descargar_con_progreso(url: str, destino: str) -> bool:
    """Descarga un archivo mostrando porcentaje. Retorna True si tuvo exito."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            total = int(resp.getheader("Content-Length", 0))
            descargado = 0
            with open(destino, "wb") as f:
                while True:
                    chunk = resp.read(65536)
                    if not chunk:
                        break
                    f.write(chunk)
                    descargado += len(chunk)
                    if total > 0:
                        pct = int(descargado * 100 / total)
                        sys.stdout.write("  " + str(pct) + "%   \r")
                        sys.stdout.flush()
        print()
        return True
    except Exception as e:
        print("  Fallo al descargar " + url[:60] + ": " + str(e))
        if os.path.exists(destino):
            os.remove(destino)
        return False


def descargar_tesseract_portable() -> bool:
    """Descarga e instala Tesseract OCR portable (no requiere permisos de administrador)."""
    if os.path.exists(TESS_EXE):
        print("  ✓ Tesseract OCR ya está instalado")
        return True

    print("\n  Descargando Tesseract OCR portable (~30 MB)...")
    print("  Esto permite leer fotos y PDFs escaneados sin instalar nada más.")

    # Intentar cada URL hasta que una funcione
    ok = False
    for url in TESS_URLS:
        print(f"  Intentando: {url[:55]}...")
        ok = _descargar_con_progreso(url, TESS_ZIP)
        if ok:
            break

    if not ok:
        print("  ⚠ No se pudo descargar Tesseract (sin internet o bloqueado por firewall).")
        print("  El sistema funcionará sin OCR — puedes pegar texto directamente.")
        return False

    # Descomprimir
    try:
        import zipfile
        print("  Descomprimiendo...")
        os.makedirs(TESS_DIR, exist_ok=True)
        with zipfile.ZipFile(TESS_ZIP, "r") as z:
            # Extraer y aplanar la estructura si viene dentro de una subcarpeta
            miembros = z.namelist()
            for m in miembros:
                datos = z.read(m)
                # Quitar el primer directorio del path si existe
                partes = m.split("/", 1)
                dest_rel = partes[1] if len(partes) > 1 else partes[0]
                if not dest_rel:
                    continue
                dest_full = os.path.join(TESS_DIR, dest_rel)
                os.makedirs(os.path.dirname(dest_full), exist_ok=True)
                if not m.endswith("/"):
                    with open(dest_full, "wb") as f:
                        f.write(datos)
        os.remove(TESS_ZIP)

        # Verificar que tesseract.exe existe
        if os.path.exists(TESS_EXE):
            print("  ✓ Tesseract OCR instalado correctamente en /tesseract")
            return True
        else:
            # Buscar el exe en subdirectorios
            for raiz, _, archivos in os.walk(TESS_DIR):
                for arch in archivos:
                    if arch.lower() == "tesseract.exe":
                        # Mover al lugar correcto
                        shutil.move(os.path.join(raiz, arch), TESS_EXE)
                        print("  ✓ Tesseract OCR instalado (reubicado)")
                        return True
            print("  ⚠ Tesseract descargado pero no se encontró tesseract.exe")
            print("  El sistema funcionará sin OCR.")
            return False
    except Exception as e:
        print(f"  ⚠ Error al descomprimir Tesseract: {e}")
        print("  El sistema funcionará sin OCR.")
        if os.path.exists(TESS_ZIP):
            os.remove(TESS_ZIP)
        return False


def crear_acceso_directo():
    """Crea acceso directo en el Escritorio apuntando a NEXUS_INICIAR.bat."""
    try:
        import winreg
        escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(escritorio):
            # Intentar obtener desde registro
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            ) as k:
                escritorio = winreg.QueryValueEx(k, "Desktop")[0]

        acceso = os.path.join(escritorio, "NEXUS Omega.lnk")
        bat    = os.path.join(BASE_DIR, "NEXUS_INICIAR.bat")

        # Crear .lnk con WScript
        vbs = f"""
Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = "{acceso}"
Set oLink = oWS.CreateShortcut(sLinkFile)
oLink.TargetPath = "{bat}"
oLink.WorkingDirectory = "{BASE_DIR}"
oLink.Description = "Abrir NEXUS Omega"
oLink.IconLocation = "{bat},0"
oLink.Save
"""
        vbs_path = os.path.join(BASE_DIR, "_crear_acceso.vbs")
        with open(vbs_path, "w", encoding="utf-8") as f:
            f.write(vbs)
        subprocess.run(["wscript", "//nologo", vbs_path],
                       capture_output=True, timeout=10)
        os.remove(vbs_path)
        print("  ✓ Acceso directo creado en el Escritorio")
        return True
    except Exception as e:
        print(f"  (No se pudo crear acceso directo: {e})")
        return False


def main():
    cls()
    titulo("NEXUS Ω — INSTALACIÓN AUTOMÁTICA")
    print("""
  Bienvenido al instalador de NEXUS Omega.
  Este proceso es completamente automático.
  No cierre esta ventana hasta que termine.
""")

    total_pasos = 5

    # ── PASO 1: Verificar / instalar Python ─────────────────────
    paso(1, total_pasos, "Verificando Python")
    python_cmd = python_en_path()

    if python_cmd:
        print(f"  ✓ Python ya está instalado ({python_cmd})")
    else:
        print("  Python no encontrado. Descargando e instalando...")
        ok = descargar_python()
        if not ok:
            print("\n  ⛔ No hay conexión a internet o falló la descarga.")
            print("  Instala Python manualmente desde: https://python.org")
            print("  Luego vuelve a hacer doble clic en NEXUS_INICIAR.bat")
            input("\n  Presiona ENTER para cerrar...")
            return

        ok = instalar_python()
        if not ok:
            print("\n  ⛔ La instalación de Python falló.")
            input("\n  Presiona ENTER para cerrar...")
            return

        # Refrescar PATH y buscar de nuevo
        python_cmd = python_en_path()
        if not python_cmd:
            python_cmd = "python"  # fallback

    # ── PASO 2: Instalar dependencias ───────────────────────────
    paso(2, total_pasos, "Instalando componentes del sistema")
    instalar_dependencias(python_cmd)

    # ── PASO 3: Descargar Tesseract OCR portable ────────────────
    paso(3, total_pasos, "Instalando Tesseract OCR (lector de fotos y PDFs)")
    descargar_tesseract_portable()

    # ── PASO 4: Crear acceso directo ─────────────────────────────
    paso(4, total_pasos, "Creando acceso directo en el Escritorio")
    crear_acceso_directo()

    # ── PASO 5: Marcar como listo ────────────────────────────────
    paso(5, total_pasos, "Finalizando configuración")
    with open(MARKER, "w") as f:
        f.write(f"python={python_cmd}\nok=1\n")

    # Guardar el comando Python para que el lanzador lo use
    cfg = os.path.join(BASE_DIR, ".nexus_config")
    with open(cfg, "w") as f:
        f.write(f"python_cmd={python_cmd}\n")

    print("\n" + "=" * 60)
    print("""
  ✅ NEXUS Omega instalado correctamente.

  Puedes cerrar esta ventana.
  El sistema se abrirá automáticamente en tu navegador.

  Para usar NEXUS en el futuro:
  → Haz doble clic en el ícono del Escritorio
    o en NEXUS_INICIAR.bat
""")
    print("=" * 60)
    time.sleep(3)


if __name__ == "__main__":
    main()
