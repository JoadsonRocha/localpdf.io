import ctypes
import os
import pathlib
import socket
import sys
import threading
import time
import webbrowser

ROOT = pathlib.Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
VENDOR = ROOT / "vendor"

# Keep native OCR/PDF dependencies inside the portable package.
os.environ.setdefault("LOCALPDF_MODE", "local")
os.environ.setdefault("LOCALPDF_HOST", "127.0.0.1")
os.environ.setdefault("LOCALPDF_PORT", "5000")
os.environ.setdefault("TESSDATA_PREFIX", str(VENDOR / "tesseract" / "tessdata"))

gs_bin = VENDOR / "ghostscript" / "bin"
tess_bin = VENDOR / "tesseract"

os.environ["PATH"] = os.pathsep.join(
    [
        str(tess_bin),
        str(gs_bin),
        os.environ.get("PATH", ""),
    ]
)

# Register DLL directory for Ghostscript on Python 3.8+ (PEP 578)
if gs_bin.is_dir() and hasattr(os, "add_dll_directory"):
    try:
        os.add_dll_directory(str(gs_bin))
    except OSError:
        pass

from waitress import serve

from app import app

try:
    import webview
except ImportError:
    webview = None

MUTEX_NAME = "LocalPDF_io_SingleInstance_Mutex"
ERROR_ALREADY_EXISTS = 183
WINDOW_TITLE = "LocalPDF.io"


def acquire_single_instance_mutex():
    """
    Ensures only one instance of LocalPDF runs at a time.
    If already running, restores and brings the existing native window to front
    or opens the browser to the existing instance and exits.
    """
    if sys.platform != "win32":
        return None

    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    last_error = kernel32.GetLastError()

    if last_error == ERROR_ALREADY_EXISTS:
        port = int(os.environ.get("LOCALPDF_PORT", "5000"))
        user32 = ctypes.windll.user32
        hwnd = user32.FindWindowW(None, WINDOW_TITLE)
        if hwnd:
            user32.ShowWindow(hwnd, 9)  # SW_RESTORE
            user32.SetForegroundWindow(hwnd)
        else:
            webbrowser.open(f"http://127.0.0.1:{port}/")
        sys.exit(0)

    return mutex


def find_port(start=5000):
    for port in range(start, start + 50):
        with socket.socket() as probe:
            try:
                probe.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("No local port is available.")


def open_browser(port):
    time.sleep(0.3)
    webbrowser.open(f"http://127.0.0.1:{port}/splash")


if __name__ == "__main__":
    _mutex = acquire_single_instance_mutex()
    port = find_port(int(os.environ.get("LOCALPDF_PORT", "5000")))

    # Start the local server in a background daemon thread
    server_thread = threading.Thread(
        target=serve,
        args=(app,),
        kwargs={"host": "127.0.0.1", "port": port, "threads": 4},
        daemon=True,
    )
    server_thread.start()

    # Launch native desktop application window if webview is available
    launched_native_window = False
    if webview is not None:
        try:
            webview.create_window(
                title=WINDOW_TITLE,
                url=f"http://127.0.0.1:{port}/splash",
                width=1280,
                height=840,
                min_size=(900, 600),
                background_color="#eff6ff",
            )
            launched_native_window = True
            webview.start(gui="edgechromium")
        except Exception:
            launched_native_window = False

    # Fallback to default browser if native window could not be opened
    if not launched_native_window:
        threading.Thread(target=open_browser, args=(port,), daemon=True).start()
        while True:
            time.sleep(1)
