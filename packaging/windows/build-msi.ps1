param(
    [string]$Version = "1.0.0.0"
)

$ErrorActionPreference = "Stop"

$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\")).Path
$Portable = Join-Path $Root "dist\LocalPDF"
$Generated = Join-Path $PSScriptRoot "generated-files.wxs"
$Output = Join-Path $Root "dist\LocalPDF.msi"

if (-not (Test-Path (Join-Path $Portable "LocalPDF.exe"))) {
    throw "Build the portable package first with build-portable.ps1."
}
if (-not (Get-Command wix -ErrorAction SilentlyContinue)) {
    throw "WiX Toolset v4 is required. Install it via 'dotnet tool install --global wix' and ensure wix is on PATH."
}

# Check if vendor binaries are populated
$TessData = Join-Path $Portable "vendor\tesseract\tessdata"
$GsBin = Join-Path $Portable "vendor\ghostscript\bin"
$hasTess = (Test-Path $TessData) -and ((Get-ChildItem $TessData -ErrorAction SilentlyContinue).Count -gt 0)
$hasGs = (Test-Path $GsBin) -and ((Get-ChildItem $GsBin -ErrorAction SilentlyContinue).Count -gt 0)

if (-not $hasTess -or -not $hasGs) {
    Write-Warning "Vendor binaries (Tesseract/Ghostscript) in '$Portable\vendor' appear empty. Bundle them before creating a final release."
}

# Ensure WixToolset.Heat extension is available for harvesting
$extensions = & wix extension list 2>&1
if ($LASTEXITCODE -eq 0 -and ($extensions -notmatch "WixToolset.Heat")) {
    Write-Host "Installing WixToolset.Heat extension for directory harvesting..."
    & wix extension add WixToolset.Heat --global
}

Push-Location $Root
try {
    Write-Host "Harvesting portable files from $Portable..."
    & wix extension add WixToolset.Heat --global 2>$null
    & wix harvest dir $Portable -ext WixToolset.Heat -o $Generated -dr INSTALLFOLDER -cg AppFiles -srd -sreg
    if ($LASTEXITCODE -ne 0) { throw "WiX harvest failed." }

    Write-Host "Building MSI installer: $Output..."
    & wix build packaging\windows\installer.wxs $Generated -ext WixToolset.Heat -o $Output
    if ($LASTEXITCODE -ne 0) { throw "WiX build failed." }

    Write-Host "MSI created successfully at $Output"
    Write-Host "Remember to sign $Output with Authenticode before public distribution."
} finally {
    if (Test-Path $Generated) {
        Remove-Item -Force $Generated -ErrorAction SilentlyContinue
    }
    Pop-Location
}
