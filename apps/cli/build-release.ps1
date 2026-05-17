# Build script for Outlook Organizer CLI
# Copies PS scripts and builds release executable

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$coreScripts = Join-Path $scriptDir "..\..\core\scripts-pwsh"
$srcDir = Join-Path $scriptDir "src"

Write-Host "Copiando scripts PowerShell..." -ForegroundColor Cyan
Copy-Item "$coreScripts\*.ps1" $srcDir -Force

Write-Host "Compilando ejecutable (ReleaseSmall)..." -ForegroundColor Cyan
Push-Location $scriptDir
zig build -Doptimize=ReleaseSmall
Pop-Location

$exe = Join-Path $scriptDir "zig-out\bin\outlook-organizer-cli.exe"
if (Test-Path $exe) {
    $size = [math]::Round((Get-Item $exe).Length / 1KB)
    Write-Host "`nBuild exitoso: $exe ($size KB)" -ForegroundColor Green
} else {
    Write-Host "`nError: no se encontro el ejecutable" -ForegroundColor Red
    exit 1
}
