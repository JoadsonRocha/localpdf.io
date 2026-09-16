param(
    [Parameter(Mandatory = $true)]
    [string]$CertificateThumbprint,
    [ValidateSet("All", "Binaries", "Msi")]
    [string]$Target = "All",
    [string]$TimestampUrl = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\")).Path
$PortableDir = Join-Path $Root "dist\LocalPDF"
$MsiPath = Join-Path $Root "dist\LocalPDF.msi"

$signtool = Get-Command signtool.exe -ErrorAction SilentlyContinue
if (-not $signtool) {
    throw "Windows SDK signtool.exe is required."
}

function Sign-File([string]$FilePath) {
    if (Test-Path $FilePath) {
        Write-Host "Signing $FilePath ..."
        & $signtool.Source sign /sha1 $CertificateThumbprint /fd SHA256 /tr $TimestampUrl /td SHA256 $FilePath
        if ($LASTEXITCODE -ne 0) { throw "Signing failed for $FilePath" }
    }
}

# Sign portable binaries BEFORE creating MSI
if ($Target -in @("All", "Binaries")) {
    $exePath = Join-Path $PortableDir "LocalPDF.exe"
    if (Test-Path $exePath) {
        Sign-File $exePath
    } else {
        Write-Warning "LocalPDF.exe not found at $exePath."
    }

    # Also sign any vendor binaries if present
    Get-ChildItem -Path $PortableDir -Recurse -Include "*.exe", "*.dll" -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.FullName -ne $exePath) {
            Sign-File $_.FullName
        }
    }
}

# Sign MSI package AFTER creation
if ($Target -in @("All", "Msi")) {
    if (Test-Path $MsiPath) {
        Sign-File $MsiPath
    } else {
        if ($Target -eq "Msi") {
            Write-Warning "LocalPDF.msi not found at $MsiPath."
        }
    }
}

Write-Host "Signing complete. Verify with: signtool verify /pa /v <file>"
