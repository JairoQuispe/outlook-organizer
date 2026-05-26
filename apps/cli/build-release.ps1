# Build script for Outlook Organizer CLI
# Copies PS scripts, optimizes and builds release executable

param (
    [Parameter(Mandatory=$false)]
    [ValidateSet("ReleaseSmall", "ReleaseFast", "ReleaseSafe", "Debug")]
    [string]$Optimize = "ReleaseSmall",

    [Parameter(Mandatory=$false)]
    [switch]$Compress = $true
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$coreScripts = Join-Path $scriptDir "..\..\core\scripts-pwsh"
$srcDir = Join-Path $scriptDir "src"

# Crear directorio src si no existe
if (-not (Test-Path $srcDir)) {
    New-Item -ItemType Directory -Path $srcDir | Out-Null
}

Write-Host "Verificando scripts PowerShell..." -ForegroundColor Cyan
$copiedAny = $false
Get-ChildItem "$coreScripts\*.ps1" | ForEach-Object {
    $destFile = Join-Path $srcDir $_.Name
    $needsCopy = $true
    if (Test-Path $destFile) {
        $srcHash = (Get-FileHash $_.FullName -Algorithm MD5).Hash
        $destHash = (Get-FileHash $destFile -Algorithm MD5).Hash
        if ($srcHash -eq $destHash) {
            $needsCopy = $false
        }
    }
    if ($needsCopy) {
        Copy-Item $_.FullName $destFile -Force
        Write-Host "  Copiado: $_.Name (actualizado)" -ForegroundColor Gray
        $copiedAny = $true
    }
}
if (-not $copiedAny) {
    Write-Host "  Todos los scripts están actualizados." -ForegroundColor Gray
}

Write-Host "`nCompilando ejecutable ($Optimize)..." -ForegroundColor Cyan
Push-Location $scriptDir
zig build "-Doptimize=$Optimize"
Pop-Location

$exe = Join-Path $scriptDir "zig-out\bin\outlook-organizer-cli.exe"
if (Test-Path $exe) {
    $size = [math]::Round((Get-Item $exe).Length / 1KB)
    Write-Host "`nBuild exitoso: $exe ($size KB)" -ForegroundColor Green

    # Opcional: Comprimir con UPX si está habilitado y UPX está disponible
    if ($Compress -and $Optimize -ne "Debug") {
        $upxCmd = "upx"
        $hasUpx = $false
        if (Get-Command $upxCmd -ErrorAction SilentlyContinue) {
            $hasUpx = $true
        } else {
            # Intentar encontrar en la ruta de paquetes de WinGet por defecto si el PATH no se ha refrescado
            $wingetUpx = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\UPX.UPX_*" -Recurse -Filter "upx.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($wingetUpx) {
                $upxCmd = $wingetUpx.FullName
                $hasUpx = $true
            }
        }

        if ($hasUpx) {
            Write-Host "`nComprimiendo ejecutable con UPX..." -ForegroundColor Cyan
            $oldSize = (Get-Item $exe).Length
            & $upxCmd --best $exe | Out-Null
            $newSize = (Get-Item $exe).Length
            $pct = [math]::Round((1 - ($newSize / $oldSize)) * 100)
            $newSizeKB = [math]::Round($newSize / 1KB)
            Write-Host "  Compresión UPX exitosa: Reducido en $pct% ($newSizeKB KB)" -ForegroundColor Green
        } else {
            Write-Host "`n[UPX] No encontrado en el PATH ni en WinGet. Omitiendo compresión. (Instala UPX para reducir el tamaño a ~150KB)" -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "`nError: no se encontro el ejecutable" -ForegroundColor Red
    exit 1
}
