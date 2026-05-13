$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$DistDir = Join-Path $ProjectDir "dist"
$ReleaseDir = Join-Path $ProjectDir "release"

if (-not (Test-Path (Join-Path $DistDir "Script_Potenze.exe"))) {
    throw "Eseguibile non trovato. Esegui prima build_exe.ps1"
}

if (-not (Test-Path $ReleaseDir)) {
    New-Item -Path $ReleaseDir -ItemType Directory | Out-Null
}

$exeSrc = Join-Path $DistDir "Script_Potenze.exe"
$exeDst = Join-Path $ReleaseDir "Script_Potenze.exe"
try {
    Copy-Item $exeSrc $exeDst -Force -ErrorAction Stop
} catch {
    $exeAlt = Join-Path $ReleaseDir "Script_Potenze.exe.new"
    Copy-Item $exeSrc $exeAlt -Force
    Write-Host ""
    Write-Host "[ATTENZIONE] Impossibile sovrascrivere Script_Potenze.exe (file in uso?)." -ForegroundColor Yellow
    Write-Host "Copiato l'eseguibile aggiornato come: $exeAlt"
    Write-Host "Chiudi l'applicazione, elimina o rinomina release\Script_Potenze.exe, poi rinomina .exe.new in Script_Potenze.exe oppure rilancia prepare_release.ps1."
}

Copy-Item (Join-Path $ProjectDir "configurazione_schede.txt") (Join-Path $ReleaseDir "configurazione_schede.txt") -Force

Write-Host ""
Write-Host "Pacchetto release pronto in: $ReleaseDir"
Write-Host "Distribuire agli utenti contenuto di questa cartella."
