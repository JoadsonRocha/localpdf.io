param(
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    [ValidateSet("All", "Binaries", "Msi")]
    [string]$Target = "All",
    [string]$TimestampUrl = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\")).Path
$PortableDir = Join-Path $Root "dist\LocalPDF"
$MsiPath = Join-Path $Root "dist\LocalPDF.msi"

# Auto-detect code signing certificate if thumbprint not provided
if (-not $CertificateThumbprint) {
    $cert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
        Where-Object { $_.HasPrivateKey -and ($_.EnhancedKeyUsageList.FriendlyName -contains "Assinatura do Código" -or $_.EnhancedKeyUsageList.FriendlyName -contains "Code Signing" -or $_.Thumbprint -eq "B5441819C4B2C689E5FD1ADF025F031E29709354") } |
        Select-Object -First 1

    if ($cert) {
        $CertificateThumbprint = $cert.Thumbprint
        Write-Host "Certificado de assinatura detectado automaticamente: $($cert.Subject) ($CertificateThumbprint)" -ForegroundColor Cyan
    } else {
        throw "Nenhum certificado de Code Signing encontrado em Cert:\CurrentUser\My ou Cert:\LocalMachine\My. Especifique -CertificateThumbprint."
    }
}

$certObj = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
if (-not $certObj) {
    $certObj = Get-Item "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
}

$signtool = Get-Command signtool.exe -ErrorAction SilentlyContinue

function Sign-File([string]$FilePath) {
    if (Test-Path $FilePath) {
        Write-Host "Assinando $FilePath ..." -ForegroundColor Yellow
        if ($signtool) {
            & $signtool.Source sign /sha1 $CertificateThumbprint /fd SHA256 /tr $TimestampUrl /td SHA256 $FilePath
            if ($LASTEXITCODE -ne 0) {
                # Tentar sem timestamp se o servidor estiver inacessível
                & $signtool.Source sign /sha1 $CertificateThumbprint /fd SHA256 $FilePath
            }
        } elseif ($certObj) {
            try {
                Set-AuthenticodeSignature -FilePath $FilePath -Certificate $certObj -TimestampServer $TimestampUrl -HashAlgorithm SHA256 | Out-Null
            } catch {
                # Fallback sem timestamp caso conexão de rede falhe
                Set-AuthenticodeSignature -FilePath $FilePath -Certificate $certObj -HashAlgorithm SHA256 | Out-Null
            }
        } else {
            throw "Não foi possível carregar o objeto do certificado para $CertificateThumbprint."
        }
        Write-Host "[OK] Assinado: $FilePath" -ForegroundColor Green
    }
}

# Sign portable binaries BEFORE creating MSI
if ($Target -in @("All", "Binaries")) {
    $exePath = Join-Path $PortableDir "LocalPDF.exe"
    if (Test-Path $exePath) {
        Sign-File $exePath
    } else {
        Write-Warning "LocalPDF.exe não encontrado em $exePath."
    }

    # Also sign unsigned vendor binaries if present (do not overwrite signatures of internal/system DLLs)
    $vendorDir = Join-Path $PortableDir "vendor"
    if (Test-Path $vendorDir) {
        Get-ChildItem -Path $vendorDir -Recurse -Include "*.exe", "*.dll" -ErrorAction SilentlyContinue | ForEach-Object {
            $existingSig = Get-AuthenticodeSignature $_.FullName -ErrorAction SilentlyContinue
            if (-not $existingSig -or $existingSig.Status -ne "Valid") {
                Sign-File $_.FullName
            }
        }
    }
}

# Sign MSI package AFTER creation
if ($Target -in @("All", "Msi")) {
    if (Test-Path $MsiPath) {
        Sign-File $MsiPath
    } else {
        if ($Target -eq "Msi") {
            Write-Warning "LocalPDF.msi não encontrado em $MsiPath."
        }
    }
}

Write-Host "`nAssinatura concluída com sucesso!" -ForegroundColor Green
Write-Host "Para verificar status: Get-AuthenticodeSignature -FilePath .\dist\LocalPDF\LocalPDF.exe" -ForegroundColor Cyan

