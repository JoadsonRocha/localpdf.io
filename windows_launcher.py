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


class DesktopAPI:
    """Native desktop integration API exposed to JavaScript in PyWebView."""

    def __init__(self):
        self._window = None

    def set_window(self, window):
        self._window = window

    def get_downloads_dir(self):
        downloads = pathlib.Path.home() / "Downloads"
        return str(downloads) if downloads.exists() else str(pathlib.Path.home())

    def choose_save_path(self, filename=""):
        """
        Opens a native Windows SaveFileDialog and returns the chosen absolute path,
        or None if cancelled.
        """
        if not self._window:
            return None
        try:
            ext = pathlib.Path(filename).suffix.lower()
            filter_map = {
                ".pdf": "PDF Documents (*.pdf)|*.pdf",
                ".docx": "Word Documents (*.docx)|*.docx",
                ".xlsx": "Excel Spreadsheets (*.xlsx)|*.xlsx",
                ".zip": "ZIP Archives (*.zip)|*.zip",
                ".txt": "Text Files (*.txt)|*.txt",
                ".png": "PNG Images (*.png)|*.png",
                ".jpg": "JPEG Images (*.jpg)|*.jpg",
            }
            primary_filter = filter_map.get(ext, "All Files (*.*)|*.*")
            file_types = (primary_filter, "All Files (*.*)|*.*")

            initial_dir = self.get_downloads_dir()
            res = self._window.create_file_dialog(
                webview.FileDialog.SAVE,
                directory=initial_dir,
                save_filename=filename,
                file_types=file_types,
            )
            if res:
                chosen = res[0] if isinstance(res, (list, tuple)) else str(res)
                return os.path.abspath(chosen)
            return None
        except Exception:
            return None

    def save_file(self, filename, base64_data, target_path=None):
        """
        Fallback to save binary data directly via base64.
        """
        try:
            import base64

            file_bytes = base64.b64decode(base64_data)
            save_path = target_path
            if not save_path:
                save_path = self.choose_save_path(filename)
            if not save_path:
                downloads = pathlib.Path(self.get_downloads_dir())
                target = downloads / filename
                stem = target.stem
                suffix = target.suffix
                c = 1
                while target.exists():
                    target = downloads / f"{stem} ({c}){suffix}"
                    c += 1
                save_path = str(target)

            with open(save_path, "wb") as f:
                f.write(file_bytes)
            return {
                "success": True,
                "path": os.path.abspath(save_path),
                "filename": os.path.basename(save_path),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def open_saved_folder(self, file_path):
        """Opens Windows Explorer with the specified file selected."""
        try:
            if not file_path:
                return False
            path = os.path.abspath(file_path)
            if os.path.exists(path):
                import subprocess

                subprocess.Popen(["explorer.exe", f"/select,{path}"])
                return True
            elif os.path.exists(os.path.dirname(path)):
                import subprocess

                subprocess.Popen(["explorer.exe", os.path.dirname(path)])
                return True
        except Exception:
            pass
        return False

    def open_saved_file(self, file_path):
        """Opens the saved file in the default Windows application."""
        try:
            if not file_path:
                return False
            path = os.path.abspath(file_path)
            if os.path.exists(path):
                os.startfile(path)
                return True
        except Exception:
            pass
        return False


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
            webview.settings["ALLOW_DOWNLOADS"] = True
            desktop_api = DesktopAPI()
            window = webview.create_window(
                title=WINDOW_TITLE,
                url=f"http://127.0.0.1:{port}/splash",
                width=1280,
                height=840,
                min_size=(900, 600),
                background_color="#eff6ff",
                js_api=desktop_api,
            )
            desktop_api.set_window(window)
            webview.start(gui="edgechromium")
            launched_native_window = True
        except Exception:
            launched_native_window = False

    # Fallback to default browser if native window could not be opened
    if not launched_native_window:
        threading.Thread(target=open_browser, args=(port,), daemon=True).start()
        while True:
            time.sleep(1)
