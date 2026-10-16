import ctypes
import io
import os
import pathlib
import socket
import sys
import tempfile
import threading
import time
import webbrowser

# Ensure stdout and stderr exist even in PyInstaller --noconsole mode
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

ROOT = pathlib.Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
VENDOR = ROOT / "vendor"

os.environ.setdefault("LOCALPDF_MODE", "local")
os.environ.setdefault("LOCALPDF_HOST", "127.0.0.1")
os.environ.setdefault("LOCALPDF_PORT", "5000")


def configure_native_dependencies():
    """Auto-detects Tesseract OCR and Ghostscript in vendor or system directories."""
    # 1. Tesseract
    tess_candidates = [
        VENDOR / "tesseract",
        pathlib.Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR",
        pathlib.Path("C:/Program Files/Tesseract-OCR"),
        pathlib.Path("C:/Program Files (x86)/Tesseract-OCR"),
    ]
    detected_tess_dir = None
    for cand in tess_candidates:
        if cand and (cand / "tesseract.exe").exists():
            detected_tess_dir = cand
            break

    if detected_tess_dir:
        tess_exe = str(detected_tess_dir / "tesseract.exe")
        tessdata_dir = detected_tess_dir / "tessdata"
        if tessdata_dir.is_dir():
            os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)
        try:
            import pytesseract

            pytesseract.pytesseract.tesseract_cmd = tess_exe
        except Exception:
            pass

    # 2. Ghostscript
    gs_bin_candidates = [
        VENDOR / "ghostscript" / "bin",
    ]
    for base in [pathlib.Path("C:/Program Files/gs"), pathlib.Path("C:/Program Files (x86)/gs")]:
        if base.exists():
            try:
                for child in sorted(base.iterdir(), reverse=True):
                    if child.is_dir() and (child / "bin").is_dir():
                        gs_bin_candidates.append(child / "bin")
            except OSError:
                pass

    detected_gs_bin = None
    for cand in gs_bin_candidates:
        if cand and cand.is_dir():
            if (
                (cand / "gswin64c.exe").exists()
                or (cand / "gswin32c.exe").exists()
                or (cand / "gsdll64.dll").exists()
                or (cand / "gsdll32.dll").exists()
            ):
                detected_gs_bin = cand
                break

    # Register DLL directory for Ghostscript on Python 3.8+ (PEP 578)
    if detected_gs_bin and hasattr(os, "add_dll_directory"):
        try:
            os.add_dll_directory(str(detected_gs_bin))
        except OSError:
            pass

    path_parts = []
    if detected_tess_dir:
        path_parts.append(str(detected_tess_dir))
    if detected_gs_bin:
        path_parts.append(str(detected_gs_bin))
    if path_parts:
        os.environ["PATH"] = os.pathsep.join(path_parts + [os.environ.get("PATH", "")])


configure_native_dependencies()

from waitress import serve

from app import app

try:
    import webview
except ImportError:
    webview = None

MUTEX_NAME = "LocalPDF_io_SingleInstance_Mutex"
ERROR_ALREADY_EXISTS = 183
WINDOW_TITLE = "LocalPDF.io"
PORT_FILE = pathlib.Path(tempfile.gettempdir()) / "localpdf_active.port"


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
        Uses valid pywebview filter syntax: ('Description (*.ext)', ...).
        """
        if not self._window:
            return None
        try:
            ext = pathlib.Path(filename).suffix.lower()
            filter_map = {
                ".pdf": "PDF Documents (*.pdf)",
                ".docx": "Word Documents (*.docx)",
                ".xlsx": "Excel Spreadsheets (*.xlsx)",
                ".zip": "ZIP Archives (*.zip)",
                ".txt": "Text Files (*.txt)",
                ".png": "PNG Images (*.png)",
                ".jpg": "JPEG Images (*.jpg;*.jpeg)",
                ".jpeg": "JPEG Images (*.jpg;*.jpeg)",
            }
            primary_filter = filter_map.get(ext, "All Files (*.*)")
            file_types = (primary_filter, "All Files (*.*)")

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

                subprocess.Popen(f'explorer.exe /select,"{path}"')
                return True
            elif os.path.exists(os.path.dirname(path)):
                import subprocess

                subprocess.Popen(f'explorer.exe "{os.path.dirname(path)}"')
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


def get_active_port():
    """Reads the active port from the port lockfile or fallback to env var."""
    if PORT_FILE.exists():
        try:
            val = int(PORT_FILE.read_text().strip())
            if 1024 <= val <= 65535:
                return val
        except Exception:
            pass
    return int(os.environ.get("LOCALPDF_PORT", "5000"))


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
        port = get_active_port()
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
    os.environ["LOCALPDF_PORT"] = str(port)

    # Persist the active port for subsequent instances
    try:
        PORT_FILE.write_text(str(port))
    except Exception:
        pass

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
            try:
                webview.start(gui="edgechromium")
                launched_native_window = True
            except Exception:
                try:
                    webview.start()
                    launched_native_window = True
                except Exception:
                    launched_native_window = False
        except Exception:
            launched_native_window = False

    # Fallback to default browser if native window could not be opened
    if not launched_native_window:
        threading.Thread(target=open_browser, args=(port,), daemon=True).start()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass

    # Cleanup port file on exit
    try:
        if PORT_FILE.exists():
            PORT_FILE.unlink(missing_ok=True)
    except Exception:
        pass
