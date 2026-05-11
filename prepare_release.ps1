$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$DistDir = Join-Path $ProjectDir "dist"
$ReleaseDir = Join-Path $ProjectDir "release"

if (-not (Test-Path (Join-Path $DistDir "Script_Potenze.exe"))) {
    throw "Eseguibile non trovato. Esegui prima build_exe.ps1"
}

if (Test-Path $ReleaseDir) {
    Remove-Item $ReleaseDir -Recurse -Force
}
New-Item -Path $ReleaseDir -ItemType Directory | Out-Null

Copy-Item (Join-Path $DistDir "Script_Potenze.exe") (Join-Path $ReleaseDir "Script_Potenze.exe")
Copy-Item (Join-Path $ProjectDir "configurazione_schede.txt") (Join-Path $ReleaseDir "configurazione_schede.txt")
Copy-Item (Join-Path $ProjectDir "353880 - Smistamento potenza.xls") (Join-Path $ReleaseDir "353880 - Smistamento potenza.xls")

Write-Host ""
Write-Host "Pacchetto release pronto in: $ReleaseDir"
Write-Host "Distribuire agli utenti contenuto di questa cartella."
