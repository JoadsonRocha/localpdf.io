# PyInstaller one-folder build for the Windows portable package.
from pathlib import Path

from PyInstaller.utils.hooks import collect_all

root = Path(SPECPATH).parents[1]
icon_file = root / "packaging" / "windows" / "icon.ico"

datas = [(str(root / "favicon.svg"), ".")]
if icon_file.exists():
    datas.append((str(icon_file), "."))

binaries = []
hiddenimports = [
    "fitz",
    "pytesseract",
    "pdf2docx",
    "ghostscript",
    "waitress",
    "openpyxl",
    "docx",
    "reportlab",
    "pdfplumber",
]

for package in ("fitz", "pdf2docx", "PIL", "docx", "reportlab", "openpyxl", "pdfplumber"):
    try:
        package_datas, package_binaries, package_hidden = collect_all(package)
        datas.extend(package_datas)
        binaries.extend(package_binaries)
        hiddenimports.extend(package_hidden)
    except ImportError:
        pass

analysis = Analysis(
    [str(root / "windows_launcher.py")],
    pathex=[str(root)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(analysis.pure)
exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="LocalPDF",
    icon=str(icon_file) if icon_file.exists() else None,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)
coll = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    name="LocalPDF",
)
