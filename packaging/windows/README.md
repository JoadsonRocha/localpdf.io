# Windows distribution

LocalPDF.io has two planned Windows distributions:

- **Portable**: a folder containing `LocalPDF.exe`, bundled Python dependencies and local vendor folders.
- **MSI**: a WiX installer generated from the portable folder.

## Requirements

- Windows 10 or newer, 64-bit.
- Python 3.11 or newer for the build environment.
- PyInstaller and Waitress from `requirements_dev.txt`.
- WiX Toolset v4 on `PATH` for MSI generation.
- Windows SDK `signtool.exe` for release signing.
- Windows builds of Tesseract OCR and Ghostscript copied into `dist/LocalPDF/vendor/`.

## Build portable

From PowerShell at the repository root:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m pip install -r requirements_dev.txt
.\packaging\windows\build-portable.ps1
```

The executable is created at `dist/LocalPDF/LocalPDF.exe`. Running it listens only on `127.0.0.1`, chooses a free local port and opens the browser automatically.

The build script creates these vendor folders:

```text
dist/LocalPDF/vendor/tesseract/tessdata
dist/LocalPDF/vendor/ghostscript/bin
```

Copy the official Windows runtime files and language data into those folders before distributing the portable package. Review the licenses for every bundled dependency.

## Build MSI

After building and completing the vendor folders:

```powershell
.\packaging\windows\build-msi.ps1
```

The result is `dist/LocalPDF.msi`.

## Avoiding Windows warnings

A package cannot be made trusted by configuration alone. For public distribution:

1. Obtain an Authenticode code-signing certificate from a recognized certificate authority.
2. Sign the executable and MSI with `sign-release.ps1`.
3. Use SHA-256 and an RFC 3161 timestamp server.
4. Verify both artifacts with `signtool verify /pa /v`.
5. Publish SHA-256 hashes and the certificate information with the release.

Unsigned first releases may still show SmartScreen warnings even when the code is safe. Reputation is built over time and cannot be bypassed legitimately.

## Local security model

The Windows launcher forces `LOCALPDF_MODE=local`, binds only to `127.0.0.1`, does not configure firewall rules and does not use Railway. The MSI is not yet a release artifact; these scripts are the initial packaging foundation.
