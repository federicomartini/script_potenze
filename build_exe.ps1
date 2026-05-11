$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonExe = Join-Path $ProjectDir ".venv\Scripts\python.exe"
$PyInstallerExe = Join-Path $ProjectDir ".venv\Scripts\pyinstaller.exe"

if (-not (Test-Path $PythonExe)) {
    throw "Python del venv non trovato: $PythonExe"
}

if (-not (Test-Path $PyInstallerExe)) {
    throw "PyInstaller non trovato nel venv. Installa dipendenze prima del build."
}

Set-Location $ProjectDir

& $PyInstallerExe `
    --noconfirm `
    --clean `
    --onefile `
    --console `
    --name "Script_Potenze" `
    "script_potenze.py"

Write-Host ""
Write-Host "Build completato."
Write-Host "Eseguibile: $ProjectDir\dist\Script_Potenze.exe"
