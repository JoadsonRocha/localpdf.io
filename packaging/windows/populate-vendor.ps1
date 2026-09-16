# populate-vendor.ps1
# Copia os binários do Tesseract OCR e Ghostscript instalados na máquina
# para a pasta vendor do build portátil LocalPDF.
#
# Uso:
#   .\populate-vendor.ps1
#   .\populate-vendor.ps1 -TesseractDir "C:\Meu\Tesseract" -GhostscriptDir "C:\Meu\gs\bin"
#
param(
    [string]$TesseractDir  = "",
    [string]$GhostscriptDir = ""
)

$ErrorActionPreference = "Stop"

$Root     = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$Vendor   = Join-Path $Root "dist\LocalPDF\vendor"
$DestTess = Join-Path $Vendor "tesseract"
$DestGs   = Join-Path $Vendor "ghostscript\bin"

# ─────────────────────────────────────────────────────────────────────────────
# Funções auxiliares
# ─────────────────────────────────────────────────────────────────────────────
function Find-Tesseract {
    # 1. Parâmetro explícito
    if ($TesseractDir -and (Test-Path (Join-Path $TesseractDir "tesseract.exe"))) {
        return $TesseractDir
    }
    # 2. Variável de ambiente TESSERACT_PATH
    if ($env:TESSERACT_PATH -and (Test-Path (Join-Path $env:TESSERACT_PATH "tesseract.exe"))) {
        return $env:TESSERACT_PATH
    }
    # 3. where.exe (se estiver no PATH)
    try {
        $whereResult = (where.exe tesseract 2>$null) 2>$null
        if ($LASTEXITCODE -eq 0 -and $whereResult) {
            return Split-Path $whereResult[0] -Parent
        }
    } catch { }
    # 4. Caminhos padrão de instalação
    $candidates = @(
        "C:\Program Files\Tesseract-OCR",
        "C:\Program Files (x86)\Tesseract-OCR",
        "$env:LOCALAPPDATA\Programs\Tesseract-OCR",
        "$env:ProgramFiles\Tesseract-OCR"
    )
    foreach ($c in $candidates) {
        if (Test-Path (Join-Path $c "tesseract.exe")) { return $c }
    }
    # 5. Busca automática em C:\Program Files (lento, fallback)
    Write-Host "  Buscando tesseract.exe automaticamente (pode demorar)..." -ForegroundColor DarkGray
    $found = Get-ChildItem "C:\Program Files", "C:\Program Files (x86)" `
        -Filter "tesseract.exe" -Recurse -Depth 6 -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($found) { return $found.DirectoryName }
    return $null
}

function Find-Ghostscript {
    # 1. Parâmetro explícito
    if ($GhostscriptDir -and (Test-Path $GhostscriptDir)) {
        $exe = Get-ChildItem $GhostscriptDir -Filter "gswin64c.exe" -ErrorAction SilentlyContinue |
               Select-Object -First 1
        if ($exe) { return $GhostscriptDir }
    }
    # 2. where.exe
    try {
        $whereResult = (where.exe gswin64c 2>$null) 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $whereResult) {
            $whereResult = (where.exe gswin32c 2>$null) 2>$null
        }
        if ($LASTEXITCODE -eq 0 -and $whereResult) {
            return Split-Path $whereResult[0] -Parent
        }
    } catch { }
    # 3. Caminhos padrão — versão mais recente do gs instalada
    $gsRoot = @("C:\Program Files\gs", "C:\Program Files (x86)\gs") |
        Where-Object { Test-Path $_ } |
        ForEach-Object { Get-ChildItem $_ -Directory -ErrorAction SilentlyContinue } |
        Sort-Object Name -Descending |
        Select-Object -First 1
    if ($gsRoot) {
        $bin = Join-Path $gsRoot.FullName "bin"
        if (Test-Path (Join-Path $bin "gswin64c.exe")) { return $bin }
        if (Test-Path (Join-Path $bin "gswin32c.exe")) { return $bin }
    }
    # 4. Busca automática
    Write-Host "  Buscando gswin64c.exe automaticamente (pode demorar)..." -ForegroundColor DarkGray
    $found = Get-ChildItem "C:\Program Files", "C:\Program Files (x86)" `
        -Filter "gswin64c.exe" -Recurse -Depth 6 -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($found) { return $found.DirectoryName }
    return $null
}

# ─────────────────────────────────────────────────────────────────────────────
# Verificação do build portátil
# ─────────────────────────────────────────────────────────────────────────────
if (-not (Test-Path (Join-Path $Root "dist\LocalPDF\LocalPDF.exe"))) {
    throw "Build portátil não encontrado em dist\LocalPDF\. Execute build-portable.ps1 primeiro."
}

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "  LocalPDF -- Popular binarios vendor"                  -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# TESSERACT
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[ Tesseract OCR ]" -ForegroundColor Yellow
$tessSource = Find-Tesseract

if (-not $tessSource) {
    Write-Warning "Tesseract nao encontrado. Instale em https://github.com/UB-Mannheim/tesseract/wiki"
    Write-Warning "Ou passe -TesseractDir 'C:\caminho\para\Tesseract-OCR'"
} else {
    Write-Host "  Fonte: $tessSource" -ForegroundColor Green

    # Cria destino
    New-Item -ItemType Directory -Force $DestTess | Out-Null

    # Copia tudo (exceto instalador e atalhos)
    $exclude = @("*.log", "unins*", "*.lnk")
    Get-ChildItem $tessSource -Recurse | Where-Object {
        foreach ($pat in $exclude) { if ($_.Name -like $pat) { return $false } }
        return $true
    } | ForEach-Object {
        $dest = Join-Path $DestTess ($_.FullName.Substring($tessSource.Length).TrimStart('\'))
        if ($_.PSIsContainer) {
            New-Item -ItemType Directory -Force $dest | Out-Null
        } else {
            Copy-Item -Force $_.FullName $dest
        }
    }

    $tessFiles = Get-ChildItem $DestTess -Recurse -File
    $tessCount = $tessFiles.Count
    $tessSize  = ($tessFiles | Measure-Object Length -Sum).Sum / 1MB
    Write-Host ("  Copiado: {0} arquivos ({1:N1} MB)" -f $tessCount, $tessSize) -ForegroundColor Green

    # Verifica tessdata
    $tessdata = Join-Path $DestTess "tessdata"
    $tdCount  = (Get-ChildItem $tessdata -Filter "*.traineddata" -ErrorAction SilentlyContinue).Count
    if ($tdCount -eq 0) {
        Write-Warning "Nenhum arquivo .traineddata encontrado em tessdata\. OCR pode nao funcionar."
        Write-Warning "Baixe em: https://github.com/tesseract-ocr/tessdata"
    } else {
        Write-Host "  tessdata: $tdCount idioma(s) disponivel(is)" -ForegroundColor Green
    }
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# GHOSTSCRIPT
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[ Ghostscript ]" -ForegroundColor Yellow
$gsSource = Find-Ghostscript

if (-not $gsSource) {
    Write-Warning "Ghostscript nao encontrado. Instale em https://www.ghostscript.com/releases/gsdnld.html"
    Write-Warning "Ou passe -GhostscriptDir 'C:\Program Files\gs\gs10.x.x\bin'"
} else {
    Write-Host "  Fonte: $gsSource" -ForegroundColor Green

    # Cria destino
    New-Item -ItemType Directory -Force $DestGs | Out-Null

    # Copia todos os arquivos do bin (exe + dlls)
    Get-ChildItem $gsSource -File | ForEach-Object {
        Copy-Item -Force $_.FullName (Join-Path $DestGs $_.Name)
    }

    $gsFiles = Get-ChildItem $DestGs -File
    $gsCount = $gsFiles.Count
    $gsSize  = ($gsFiles | Measure-Object Length -Sum).Sum / 1MB
    Write-Host ("  Copiado: {0} arquivos ({1:N1} MB)" -f $gsCount, $gsSize) -ForegroundColor Green
}

# ─────────────────────────────────────────────────────────────────────────────
# Resumo final
# ─────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan

$hasTess = (Test-Path (Join-Path $DestTess "tesseract.exe"))
$hasGs   = ($null -ne (Get-ChildItem $DestGs -Filter "gswin*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1))

Write-Host ""
if ($hasTess -and $hasGs) {
    Write-Host "  Build portátil 100% pronto!" -ForegroundColor Green
    Write-Host "  Pasta: $Root\dist\LocalPDF\" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Proximo passo: crie o ZIP ou gere o MSI com build-msi.ps1" -ForegroundColor Cyan
} elseif ($hasTess) {
    Write-Host "  Tesseract OK | Ghostscript AUSENTE" -ForegroundColor Yellow
    Write-Host "  Compressao via Ghostscript nao funcionara no portátil." -ForegroundColor Yellow
} elseif ($hasGs) {
    Write-Host "  Ghostscript OK | Tesseract AUSENTE" -ForegroundColor Yellow
    Write-Host "  OCR nao funcionara no portátil." -ForegroundColor Yellow
} else {
    Write-Host "  Nenhum binario vendor copiado." -ForegroundColor Red
    Write-Host "  Instale Tesseract e/ou Ghostscript e rode este script novamente." -ForegroundColor Red
}
Write-Host ""
