<#
.SYNOPSIS
    Instala o certificado de assinatura local nas Autoridades Confiáveis do Windows.
.DESCRIPTION
    Obtém o certificado local diretamente do repositório de chaves do Windows do desenvolvedor
    (Cert:\CurrentUser\My), SEM expor chaves ou arquivos no repositório Git.
#>
param(
    [string]$Thumbprint = "B5441819C4B2C689E5FD1ADF025F031E29709354",
    [string]$CertificatePath
)

$ErrorActionPreference = "Stop"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Instalador de Certificado Local Confiável - LocalPDF.io  " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

$cleanup = $false
if ($CertificatePath -and (Test-Path $CertificatePath)) {
    $tempCer = $CertificatePath
} else {
    # Busca dinamicamente no repositório protegido do Windows do usuário
    $cert = Get-Item "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction SilentlyContinue
    if (-not $cert) {
        $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { 
            $_.HasPrivateKey -and ($_.EnhancedKeyUsageList.FriendlyName -contains "Assinatura do Código" -or $_.EnhancedKeyUsageList.FriendlyName -contains "Code Signing")
        } | Select-Object -First 1
    }

    if (-not $cert) {
        Write-Error "Nenhum certificado de assinatura encontrado no seu Windows (Cert:\CurrentUser\My)."
        exit 1
    }

    # Gera arquivo temporário no diretório %TEMP% do usuário (fora do Git)
    $tempCer = Join-Path $env:TEMP "localpdf_temp_cert_$([System.Guid]::NewGuid().ToString('N')).cer"
    [System.IO.File]::WriteAllBytes($tempCer, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    $cleanup = $true
}

try {
    Write-Host "`nInstalando certificado nas Autoridades Confiáveis do seu Windows..." -ForegroundColor Yellow
    Write-Host "Certificado: $($cert.Subject)" -ForegroundColor Gray
    certutil -addstore -user -f "TrustedPublisher" $tempCer | Out-Null
    certutil -addstore -user -f "Root" $tempCer | Out-Null

    Write-Host "`n[OK] Certificado configurado como confiável com sucesso!" -ForegroundColor Green
    Write-Host "O Windows reconhecerá os executáveis e instaladores assinados nesta máquina." -ForegroundColor Green
} finally {
    if ($cleanup -and (Test-Path $tempCer)) {
        Remove-Item -Force $tempCer -ErrorAction SilentlyContinue
    }
}

