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
os.environ["PATH"] = os.pathsep.join(
    [
        str(VENDOR / "tesseract"),
        str(VENDOR / "ghostscript" / "bin"),
        os.environ.get("PATH", ""),
    ]
)

from waitress import serve

from app import app


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
    port = find_port(int(os.environ.get("LOCALPDF_PORT", "5000")))
    threading.Thread(target=open_browser, args=(port,), daemon=True).start()
    serve(app, host="127.0.0.1", port=port, threads=4)
