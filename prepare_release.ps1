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

Copy-Item (Join-Path $DistDir "Script_Potenze.exe") (Join-Path $ReleaseDir "Script_Potenze.exe") -Force
Copy-Item (Join-Path $ProjectDir "configurazione_schede.txt") (Join-Path $ReleaseDir "configurazione_schede.txt") -Force

Write-Host ""
Write-Host "Pacchetto release pronto in: $ReleaseDir"
Write-Host "Distribuire agli utenti contenuto di questa cartella."
