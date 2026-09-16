param(
    [Parameter(Mandatory = $true)]
    [string]$CertificateThumbprint,
    [string]$TimestampUrl = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\")).Path
$Artifacts = @(
    (Join-Path $Root "dist\LocalPDF.msi"),
    (Join-Path $Root "dist\LocalPDF\LocalPDF.exe")
)

$signtool = Get-Command signtool.exe -ErrorAction SilentlyContinue
if (-not $signtool) {
    throw "Windows SDK signtool.exe is required."
}

foreach ($artifact in $Artifacts) {
    if (Test-Path $artifact) {
        & $signtool.Source sign /sha1 $CertificateThumbprint /fd SHA256 /tr $TimestampUrl /td SHA256 $artifact
        if ($LASTEXITCODE -ne 0) { throw "Signing failed for $artifact" }
    }
}

Write-Host "Artifacts signed. Verify with: signtool verify /pa /v <artifact>"
