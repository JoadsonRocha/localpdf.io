$ErrorActionPreference = "Stop"

$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\")).Path
$Dist = Join-Path $Root "dist"
$Portable = Join-Path $Dist "LocalPDF"
$VenvPython = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path $VenvPython)) {
    throw "Create .venv and install requirements_dev.txt first."
}

Push-Location $Root
try {
    & $VenvPython -m PyInstaller --noconfirm --clean packaging\windows\localpdf.spec
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed." }

    New-Item -ItemType Directory -Force (Join-Path $Portable "vendor\tesseract\tessdata") | Out-Null
    New-Item -ItemType Directory -Force (Join-Path $Portable "vendor\ghostscript\bin") | Out-Null

    Write-Host "Portable build created at $Portable"
    Write-Host "Copy the licensed Windows Tesseract and Ghostscript files into vendor before distribution."
} finally {
    Pop-Location
}
