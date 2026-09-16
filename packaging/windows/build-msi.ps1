$ErrorActionPreference = "Stop"

$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\")).Path
$Portable = Join-Path $Root "dist\LocalPDF"
$Generated = Join-Path $PSScriptRoot "generated-files.wxs"
$Output = Join-Path $Root "dist\LocalPDF.msi"

if (-not (Test-Path (Join-Path $Portable "LocalPDF.exe"))) {
    throw "Build the portable package first with build-portable.ps1."
}
if (-not (Get-Command wix -ErrorAction SilentlyContinue)) {
    throw "WiX Toolset v4 is required. Install it and ensure wix is on PATH."
}

Push-Location $Root
try {
    & wix harvest dir $Portable -o $Generated -dr INSTALLFOLDER -cg AppFiles -srd -sreg
    if ($LASTEXITCODE -ne 0) { throw "WiX harvest failed." }

    & wix build packaging\windows\installer.wxs $Generated -o $Output
    if ($LASTEXITCODE -ne 0) { throw "WiX build failed." }

    Write-Host "MSI created at $Output"
    Write-Host "Sign it with Authenticode before public distribution."
} finally {
    Pop-Location
}
