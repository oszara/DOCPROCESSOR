"""
nexus_tray.py — Lanzador de NEXUS Ω con ícono en la bandeja del sistema.
Corre el servidor en segundo plano y muestra un ícono en la barra de tareas.
El usuario puede abrir el panel o detener NEXUS desde ahí.
"""

import sys, os, threading, time, webbrowser, subprocess, signal
import urllib.request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ── Leer comando Python configurado ──────────────────────────────
def _python_cmd() -> str:
    cfg = os.path.join(BASE_DIR, ".nexus_config")
    if os.path.exists(cfg):
        for line in open(cfg).readlines():
            if line.startswith("python_cmd="):
                return line.strip().split("=", 1)[1]
    return "python"


PYTHON = _python_cmd()
PORT = 8000
URL = f"http://127.0.0.1:{PORT}"

# ── Servidor ──────────────────────────────────────────────────────
_servidor_proc = None


def _iniciar_servidor():
    global _servidor_proc
    server_script = os.path.join(BASE_DIR, "main_server_v3.py")
    _servidor_proc = subprocess.Popen(
        [PYTHON, server_script],
        cwd=BASE_DIR,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
    )


def _esperar_servidor(timeout: int = 20) -> bool:
    """Espera hasta que el servidor responda o se agote el tiempo."""
    for _ in range(timeout * 2):
        try:
            urllib.request.urlopen(URL, timeout=1)
            return True
        except Exception:
            time.sleep(0.5)
    return False


def _detener_servidor():
    global _servidor_proc
    if _servidor_proc:
        try:
            _servidor_proc.terminate()
            _servidor_proc.wait(timeout=5)
        except Exception:
            try:
                _servidor_proc.kill()
            except Exception:
                pass
        _servidor_proc = None


# ── Ícono bandeja ─────────────────────────────────────────────────
def _crear_icono_imagen():
    """Crea el ícono PNG en memoria usando Pillow."""
    try:
        from PIL import Image, ImageDraw, ImageFont

        size = 64
        img = Image.new("RGBA", (size, size), (7, 11, 18, 255))
        d = ImageDraw.Draw(img)
        # Círculo de fondo
        d.ellipse([4, 4, 60, 60], fill=(0, 229, 255, 220))
        # Letra N
        try:
            fnt = ImageFont.truetype("arial.ttf", 36)
        except Exception:
            fnt = ImageFont.load_default()
        d.text((14, 10), "N", fill=(7, 11, 18, 255), font=fnt)
        return img
    except Exception:
        # Fallback: ícono mínimo si Pillow falla
        from PIL import Image

        img = Image.new("RGB", (32, 32), color=(0, 80, 160))
        return img


def _construir_menu(tray):
    import pystray

    return (
        pystray.MenuItem("🌐 Abrir NEXUS", lambda: webbrowser.open(URL), default=True),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            "📊 Tablero Inicios", lambda: webbrowser.open(URL + "/?tab=tablero")
        ),
        pystray.MenuItem(
            "🏛 Tablero Audiencias", lambda: webbrowser.open(URL + "/?tab=tablero-aud")
        ),
        pystray.MenuItem(
            "📋 Reporte Audiencias", lambda: webbrowser.open(URL + "/?tab=audiencias")
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("⛔ Detener NEXUS", lambda: _salir(tray)),
    )


def _salir(tray):
    tray.stop()
    _detener_servidor()
    sys.exit(0)


def _mostrar_notificacion(tray, titulo: str, mensaje: str):
    try:
        tray.notify(mensaje, titulo)
    except Exception:
        pass


def main():
    # Iniciar servidor en hilo separado
    hilo = threading.Thread(target=_iniciar_servidor, daemon=True)
    hilo.start()

    # Intentar mostrar ícono en bandeja (requiere pystray + pillow)
    try:
        import pystray
        from PIL import Image

        icono_img = _crear_icono_imagen()

        tray = pystray.Icon(
            "NEXUS Omega",
            icono_img,
            "NEXUS Ω — Iniciando...",
            menu=pystray.Menu(lambda: _construir_menu(tray)),
        )

        def _arrancar(t):
            ok = _esperar_servidor(timeout=20)
            if ok:
                t.title = "NEXUS Ω — Activo ✓"
                _mostrar_notificacion(
                    t, "NEXUS Ω", "Sistema listo. Haz clic para abrir."
                )
                webbrowser.open(URL)
            else:
                t.title = "NEXUS Ω — Error al iniciar"
                _mostrar_notificacion(
                    t, "NEXUS Ω ⚠", "El servidor tardó demasiado en iniciar."
                )

        tray.run(_arrancar)

    except ImportError:
        # pystray no disponible: funcionar sin ícono
        ok = _esperar_servidor(timeout=20)
        if ok:
            webbrowser.open(URL)
        # Mantener vivo el proceso hasta que el usuario cierre
        try:
            while True:
                time.sleep(60)
        except KeyboardInterrupt:
            _detener_servidor()


if __name__ == "__main__":
    main()
