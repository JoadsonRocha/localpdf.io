<#
.SYNOPSIS
    Setup completo: instala dependências (.NET SDK, WiX v4), configura certificado e gera MSI assinado.
.DESCRIPTION
    Este script faz tudo em sequência:
    1. Instala .NET SDK 8 (se necessário)
    2. Instala WiX Toolset v4 via dotnet tool
    3. Instala o certificado de assinatura como confiável no Windows
    4. Assina os binários portáteis
    5. Gera o instalador MSI
    6. Assina o MSI

.PARAMETER SkipDotnet
    Pula a instalação do .NET SDK (se já estiver instalado manualmente).
.PARAMETER SkipCert
    Pula a instalação do certificado nas Autoridades Confiáveis.
.PARAMETER Version
    Versão do instalador (padrão: 1.0.0.0).
#>
param(
    [switch]$SkipDotnet,
    [switch]$SkipCert,
    [string]$Version = "1.0.0.0"
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path

function Write-Step([string]$msg) {
    Write-Host ""
    Write-Host "====================================================" -ForegroundColor Cyan
    Write-Host "  $msg" -ForegroundColor Cyan
    Write-Host "====================================================" -ForegroundColor Cyan
}
function Write-OK([string]$msg)   { Write-Host "[OK] $msg" -ForegroundColor Green }
function Write-Warn([string]$msg) { Write-Host "[AVISO] $msg" -ForegroundColor Yellow }
function Write-Fail([string]$msg) { Write-Host "[ERRO] $msg" -ForegroundColor Red }

# ──────────────────────────────────────────────────────────────
# PASSO 1: Instalar .NET SDK 8
# ──────────────────────────────────────────────────────────────
if (-not $SkipDotnet) {
    Write-Step "PASSO 1/5: Verificando .NET SDK"

    $dotnetExe = $null
    $dotnetCandidates = @(
        "dotnet",
        "$env:ProgramFiles\dotnet\dotnet.exe",
        "$env:LOCALAPPDATA\Microsoft\dotnet\dotnet.exe",
        "$env:USERPROFILE\.dotnet\dotnet.exe"
    )
    foreach ($c in $dotnetCandidates) {
        try {
            $ver = & $c --version 2>$null
            if ($LASTEXITCODE -eq 0) { $dotnetExe = $c; break }
        } catch {}
    }

    if ($dotnetExe) {
        $dotnetVersion = & $dotnetExe --version
        Write-OK ".NET SDK já instalado: $dotnetVersion (em: $dotnetExe)"
    } else {
        Write-Host ".NET SDK não encontrado. Baixando instalador do .NET SDK 8..." -ForegroundColor Yellow

        $installerUrl = "https://builds.dotnet.microsoft.com/dotnet/Sdk/8.0.406/dotnet-sdk-8.0.406-win-x64.exe"
        $installerPath = Join-Path $env:TEMP "dotnet-sdk-installer.exe"

        Write-Host "Baixando de: $installerUrl" -ForegroundColor Gray
        Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing

        Write-Host "Instalando .NET SDK 8 (modo silencioso)..." -ForegroundColor Yellow
        $proc = Start-Process -FilePath $installerPath -ArgumentList "/quiet", "/norestart", "ADDLOCAL=ALL" -Wait -PassThru
        if ($proc.ExitCode -ne 0 -and $proc.ExitCode -ne 3010) {
            Write-Fail "Falha ao instalar .NET SDK. Código: $($proc.ExitCode)"
            throw "Instale o .NET SDK 8 manualmente em https://dotnet.microsoft.com/download e execute novamente."
        }
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue

        $env:PATH = "$env:PATH;$env:ProgramFiles\dotnet"

        $dotnetExe = "$env:ProgramFiles\dotnet\dotnet.exe"
        if (Test-Path $dotnetExe) {
            $dotnetVersion = & $dotnetExe --version
            Write-OK ".NET SDK instalado com sucesso: $dotnetVersion"
        } else {
            throw ".NET SDK instalado mas não encontrado. Reinicie o PowerShell e execute novamente."
        }
    }
} else {
    Write-Warn "Pulando verificação do .NET SDK (-SkipDotnet)"
    $dotnetExe = "dotnet"
}

# ──────────────────────────────────────────────────────────────
# PASSO 2: Instalar WiX Toolset v4
# ──────────────────────────────────────────────────────────────
Write-Step "PASSO 2/5: Verificando WiX Toolset v4"

$wixExe = $null
$wixCandidates = @("wix", "$env:USERPROFILE\.dotnet\tools\wix.exe")
foreach ($c in $wixCandidates) {
    try {
        $ver = & $c --version 2>$null
        if ($LASTEXITCODE -eq 0) { $wixExe = $c; break }
    } catch {}
}

if ($wixExe) {
    $wixVersion = & $wixExe --version
    Write-OK "WiX já instalado: $wixVersion"
} else {
    Write-Host "Instalando WiX Toolset v4..." -ForegroundColor Yellow
    & $dotnetExe tool install --global wix --version 4.0.6
    if ($LASTEXITCODE -ne 0) {
        & $dotnetExe tool update --global wix --version 4.0.6
    }

    $dotnetToolsPath = "$env:USERPROFILE\.dotnet\tools"
    if ($env:PATH -notlike "*$dotnetToolsPath*") {
        $env:PATH = "$env:PATH;$dotnetToolsPath"
    }

    $wixExe = "$env:USERPROFILE\.dotnet\tools\wix.exe"
    if (Test-Path $wixExe) {
        Write-OK "WiX instalado: $(& $wixExe --version)"
    } else {
        throw "WiX não encontrado após instalação."
    }
}


# ──────────────────────────────────────────────────────────────
# PASSO 3: Certificado confiável
# ──────────────────────────────────────────────────────────────
if (-not $SkipCert) {
    Write-Step "PASSO 3/5: Configurando certificado como confiável"
    $installCertScript = Join-Path $PSScriptRoot "install-cert.ps1"
    if (Test-Path $installCertScript) {
        try {
            & $installCertScript
            Write-OK "Certificado configurado."
        } catch {
            Write-Warn "Erro ao instalar certificado: $_"
        }
    }
} else {
    Write-Warn "Pulando configuração do certificado (-SkipCert)"
}

# ──────────────────────────────────────────────────────────────
# PASSO 4 & 5: Build MSI (assina binários + MSI automaticamente)
# ──────────────────────────────────────────────────────────────
Write-Step "PASSO 4-5/5: Gerando instalador MSI assinado"

$dotnetToolsPath = "$env:USERPROFILE\.dotnet\tools"
if ($env:PATH -notlike "*$dotnetToolsPath*") {
    $env:PATH = "$env:PATH;$dotnetToolsPath"
}

$buildMsiScript = Join-Path $PSScriptRoot "build-msi.ps1"
Push-Location $Root
try {
    & $buildMsiScript -Version $Version
    if ($LASTEXITCODE -ne 0) { throw "build-msi.ps1 falhou com código $LASTEXITCODE" }
} finally {
    Pop-Location
}

# ──────────────────────────────────────────────────────────────
# RESULTADO FINAL
# ──────────────────────────────────────────────────────────────
$msiPath = Join-Path $Root "dist\LocalPDF.msi"
if (Test-Path $msiPath) {
    $msiSize = [Math]::Round((Get-Item $msiPath).Length / 1MB, 2)
    $sig = Get-AuthenticodeSignature -FilePath $msiPath

    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║        MSI GERADO COM SUCESSO!                   ║" -ForegroundColor Green
    Write-Host "╠══════════════════════════════════════════════════╣" -ForegroundColor Green
    Write-Host "║ Arquivo   : dist\LocalPDF.msi                    ║" -ForegroundColor Green
    Write-Host "║ Tamanho   : $msiSize MB" -ForegroundColor Green
    Write-Host "║ Assinatura: $($sig.Status)" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Green
} else {
    Write-Fail "MSI não encontrado após build. Verifique os erros acima."
}

