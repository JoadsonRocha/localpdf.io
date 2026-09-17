param(
    [string]$Version = "1.0.0.0"
)

$ErrorActionPreference = "Stop"

$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\")).Path
$Portable = Join-Path $Root "dist\LocalPDF"
$Generated = Join-Path $PSScriptRoot "generated-files.wxs"
$Output = Join-Path $Root "dist\LocalPDF.msi"

$dotnetToolsPath = "$env:USERPROFILE\.dotnet\tools"
if ($env:PATH -notlike "*$dotnetToolsPath*") {
    $env:PATH = "$env:PATH;$dotnetToolsPath"
}

if (-not (Test-Path (Join-Path $Portable "LocalPDF.exe"))) {
    throw "Build the portable package first with build-portable.ps1."
}
if (-not (Get-Command wix -ErrorAction SilentlyContinue)) {
    throw "WiX Toolset v4 is required. Install it via 'dotnet tool install --global wix --version 4.0.6' and ensure wix is on PATH."
}

# Check if vendor binaries are populated
$TessData = Join-Path $Portable "vendor\tesseract\tessdata"
$GsBin = Join-Path $Portable "vendor\ghostscript\bin"
$hasTess = (Test-Path $TessData) -and ((Get-ChildItem $TessData -ErrorAction SilentlyContinue).Count -gt 0)
$hasGs = (Test-Path $GsBin) -and ((Get-ChildItem $GsBin -ErrorAction SilentlyContinue).Count -gt 0)

if (-not $hasTess -or -not $hasGs) {
    Write-Warning "Vendor binaries (Tesseract/Ghostscript) in '$Portable\vendor' appear empty. Bundle them before creating a final release."
}


Push-Location $Root
try {
    # Assinar binários portáteis antes de empacotar
    $signScript = Join-Path $PSScriptRoot "sign-release.ps1"
    if (Test-Path $signScript) {
        Write-Host "Assinando binários portáteis com Authenticode..." -ForegroundColor Cyan
        try {
            & $signScript -Target Binaries
        } catch {
            Write-Warning "Falha ao assinar binários: $_"
        }
    }

    Write-Host "Building MSI installer: $Output..."
    & wix build packaging\windows\installer.wxs -o $Output
    if ($LASTEXITCODE -ne 0) { throw "WiX build failed." }

    Write-Host "MSI created successfully at $Output"

    # Assinar o arquivo MSI resultante
    if (Test-Path $signScript) {
        Write-Host "Assinando instalador MSI com Authenticode..." -ForegroundColor Cyan
        try {
            & $signScript -Target Msi
        } catch {
            Write-Warning "Falha ao assinar MSI: $_"
        }
    }
} finally {
    if (Test-Path $Generated) {
        Remove-Item -Force $Generated -ErrorAction SilentlyContinue
    }
    Pop-Location
}
