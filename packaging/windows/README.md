# Windows distribution

LocalPDF.io has two Windows distribution targets:

- **Portable**: a folder containing `LocalPDF.exe`, bundled Python dependencies, application icons and local vendor folders.
- **MSI**: a WiX v4 installer (`LocalPDF.msi`) generated from the portable folder for system installation into `Program Files`.

## Requirements

- Windows 10 or newer, 64-bit.
- Python 3.11 or newer for the build environment.
- PyInstaller and Waitress from `requirements_dev.txt`.
- .NET 8 SDK and WiX Toolset v4 (`dotnet tool install --global wix`).
- WiX Heat extension (`wix extension add WixToolset.Heat --global`).
- Windows SDK `signtool.exe` for release signing (optional during local testing).
- Windows builds of Tesseract OCR and Ghostscript copied into `dist/LocalPDF/vendor/`.

## 1. Build portable

From PowerShell at the repository root:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m pip install -r requirements_dev.txt
.\packaging\windows\build-portable.ps1
```

The executable is created at `dist/LocalPDF/LocalPDF.exe` with `icon.ico` embedded. Running it listens only on `127.0.0.1`, ensures single-instance execution via a Windows Mutex, chooses a free local port and opens the browser automatically.

The build script creates these vendor folders:

```text
dist/LocalPDF/vendor/tesseract/tessdata
dist/LocalPDF/vendor/ghostscript/bin
```

Copy the official Windows runtime files (e.g. `gsdll64.dll`, `tesseract.exe`, language `.traineddata`) into those folders before distributing the package.

## 2. Assinatura Digital & Certificado (Code Signing)

Para evitar alertas do Windows Defender SmartScreen:

- **Configurar o certificado como confiável na máquina local**:
  ```powershell
  .\packaging\windows\install-cert.ps1
  ```
  Isso adiciona o certificado da máquina (`Cert:\CurrentUser\My`) às Autoridades Confiáveis do Windows sem armazenar arquivos de chaves no repositório.

- **Assinar binários e instalador**:
  O script `sign-release.ps1` detecta automaticamente o certificado de assinatura instalado no Windows (`Cert:\CurrentUser\My`) e utiliza `Set-AuthenticodeSignature` nativo do PowerShell com carimbo de tempo SHA256 (DigiCert):
  ```powershell
  # Assina automaticamente executáveis e DLLs
  .\packaging\windows\sign-release.ps1 -Target Binaries

  # Ou assine tudo de uma vez
  .\packaging\windows\sign-release.ps1 -Target All
  ```

## 3. Build MSI

Generate the installer package using WiX v4 (ele assina automaticamente os binários antes e o MSI após o build):

```powershell
.\packaging\windows\build-msi.ps1
```

The installer `dist/LocalPDF.msi` is created with:
- `Scope="perMachine"` for proper multi-user installation into `Program Files`.
- Start Menu shortcut with application icon and Start Menu grouping.
- Add/Remove Programs (ARP) metadata and icon.
- Automated cleanup on uninstall.

## 4. Verificação da Assinatura

Verifique o status da assinatura em qualquer arquivo:

```powershell
Get-AuthenticodeSignature -FilePath .\dist\LocalPDF\LocalPDF.exe
```

## Local security model

The Windows launcher forces `LOCALPDF_MODE=local`, binds only to `127.0.0.1`, creates temporary working files in the user's `%TEMP%` directory (never in `C:\Program Files`), does not configure firewall rules, and prevents multiple zombie background processes.
