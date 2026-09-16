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

MUTEX_NAME = "LocalPDF_io_SingleInstance_Mutex"
ERROR_ALREADY_EXISTS = 183


def acquire_single_instance_mutex():
    """
    Ensures only one instance of LocalPDF runs at a time.
    If already running, opens the browser to the existing instance and exits.
    """
    if sys.platform != "win32":
        return None

    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    last_error = kernel32.GetLastError()

    if last_error == ERROR_ALREADY_EXISTS:
        port = int(os.environ.get("LOCALPDF_PORT", "5000"))
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
    time.sleep(1)
    webbrowser.open(f"http://127.0.0.1:{port}/")


if __name__ == "__main__":
    _mutex = acquire_single_instance_mutex()
    port = find_port(int(os.environ.get("LOCALPDF_PORT", "5000")))
    threading.Thread(target=open_browser, args=(port,), daemon=True).start()
    serve(app, host="127.0.0.1", port=port, threads=4)
