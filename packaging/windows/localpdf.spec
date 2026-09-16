# PyInstaller one-folder build for the Windows portable package.
from pathlib import Path

from PyInstaller.utils.hooks import collect_all

root = Path(SPECPATH).parents[1]
datas = [(str(root / "favicon.svg"), ".")]
binaries = []
hiddenimports = ["fitz", "pytesseract", "pdf2docx", "ghostscript"]

for package in ("fitz", "pdf2docx", "PIL"):
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
