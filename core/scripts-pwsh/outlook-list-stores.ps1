param (
    [Parameter(Mandatory=$false)]
    [switch]$Json,
    [Parameter(Mandatory=$false)]
    [switch]$PreserveSession,
    [Parameter(Mandatory=$false)]
    [ValidateSet('ExchangeOnline','OST','PST')]
    [string]$StoreType,
    [Parameter(Mandatory=$false)]
    [ValidateSet('None','Json','Text')]
    [string]$ExportResult = 'None',
    [Parameter(Mandatory=$false)]
    [string]$ProfileName
)

function Format-StoreBytes {
    param($bytes)

    if (Get-Command Format-Bytes -ErrorAction SilentlyContinue) {
        return (Format-Bytes $bytes)
    }

    if ($bytes -ge 1GB) {
        return "{0:N2} GB" -f ($bytes / 1GB)
    } elseif ($bytes -ge 1MB) {
        return "{0:N2} MB" -f ($bytes / 1MB)
    } elseif ($bytes -ge 1KB) {
        return "{0:N2} KB" -f ($bytes / 1KB)
    } else {
        return "$bytes Bytes"
    }
}

function Get-SafePropertyValue {
    param(
        $object,
        [string]$PropertyName
    )

    try {
        return $object.$PropertyName
    } catch {
        return $null
    }
}

function Get-StoreInfo {
    param(
        $store
    )

    $path = Get-SafePropertyValue -object $store -PropertyName 'FilePath'

    $fileSize = $null
    $hasLocalPath = $false
    try {
        if ($path -and (Test-Path $path)) {
            $hasLocalPath = $true
            $fileInfo = Get-Item $path
            $fileSize = Format-StoreBytes $fileInfo.Length
        }
    } catch {}

    $exchangeStoreType = Get-SafePropertyValue -object $store -PropertyName 'ExchangeStoreType'
    $id = Get-SafePropertyValue -object $store -PropertyName 'StoreID'

    $storeCategory = Resolve-StoreCategory -ExchangeStoreType $exchangeStoreType -HasLocalPath $hasLocalPath -FilePath $path

    $displayName = Get-SafePropertyValue -object $store -PropertyName 'DisplayName'
    if (-not $displayName) { $displayName = 'Sin nombre' }

    $fileSizeDisplay = if ($fileSize) { $fileSize } else { "0 Bytes" }

    $storeInfo = [pscustomobject]@{
        displayName = $displayName
        storeId = $id
        filePath = $path
        fileSize = $fileSizeDisplay
        exchangeStoreType = $exchangeStoreType
        storeType = $storeCategory
    }

    return $storeInfo
}

function Get-ConnectedOutlookStores {
    param(
        $namespace
    )

    $result = New-Object System.Collections.Generic.List[object]
    $storesCollection = $namespace.Stores
    foreach ($store in $storesCollection) {
        try {
            $result.Add((Get-StoreInfo -store $store)) | Out-Null
        } finally {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($store) | Out-Null } catch {}
        }
    }

    if ($storesCollection) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($storesCollection) | Out-Null } catch {}
    }

    return $result.ToArray()
}

function Resolve-StoreCategory {
    param(
        $ExchangeStoreType,
        [bool]$HasLocalPath,
        [string]$FilePath
    )

    if ($null -ne $ExchangeStoreType -and -not $HasLocalPath) {
        return 'ExchangeOnline'
    }

    if ($HasLocalPath -and $FilePath) {
        $extension = [System.IO.Path]::GetExtension($FilePath)
        if ($extension) { $extension = $extension.ToLowerInvariant() }
        switch ($extension) {
            '.ost' { return 'OST' }
            '.pst' { return 'PST' }
        }
    }

    return 'Unknown'
}

function Get-OutlookExePath {
    $paths = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles}\Microsoft Office\root\Office15\OUTLOOK.EXE",
        "${env:ProgramFiles}\Microsoft Office\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles}\Microsoft Office\Office15\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15\OUTLOOK.EXE",
        "OUTLOOK.EXE"
    )
    foreach ($p in $paths) {
        try {
            $cmd = Get-Command $p -ErrorAction Stop
            if ($cmd) { return $cmd.Source }
        } catch {}
    }
    return "OUTLOOK.EXE"
}

if ($MyInvocation.InvocationName -ne ".") {
    $outlook = $null
    $namespace = $null
    $createdOutlook = $false
    $createdOutlookPid = $null
    $existingOutlookPids = @()
    if (-not $PreserveSession) {
        try {
            $existingOutlookPids = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
        } catch {}
    }
    $jsonErrorPayload = $null
    try {
        $maxRetries = 3
        $retryDelay = 2
        $progId = "Outlook.Application"

        $targetProfile = $ProfileName
        if ([string]::IsNullOrWhiteSpace($targetProfile)) {
            try {
                $targetProfile = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name "DefaultProfile" -ErrorAction SilentlyContinue).DefaultProfile
            } catch {}
            if (-not $targetProfile) {
                try {
                    $targetProfile = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles" -Name "DefaultProfile" -ErrorAction SilentlyContinue).DefaultProfile
                } catch {}
            }
            if (-not $targetProfile) { $targetProfile = "Outlook" }
        }

        if (-not $PreserveSession) {
            try {
                $outlookProbe = New-Object -ComObject $progId -ErrorAction Stop
                try { $outlookProbe.Quit() } catch {}
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlookProbe) | Out-Null
                $outlookProbe = $null
            } catch {
                throw "Outlook no parece estar instalado o registrado (ProgID '$progId')."
            }
        }

        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                try {
                    $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
                    $namespace = $outlook.GetNamespace("MAPI")
                    try { $namespace.Logon($ProfileName, "", $false, $false) } catch {}
                } catch {
                    if ($PreserveSession) {
                        throw "Outlook debe estar abierto para usar -PreserveSession."
                    }
                    $outlook = New-Object -ComObject Outlook.Application
                    $namespace = $outlook.GetNamespace("MAPI")
                    try { $namespace.Logon($ProfileName, "", $false, $true) } catch {}
                    $createdOutlook = $true
                    if (-not $createdOutlookPid) {
                        try {
                            $currentPids = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
                            if ($currentPids) {
                                $createdOutlookPid = ($currentPids | Where-Object { $existingOutlookPids -notcontains $_ }) | Select-Object -First 1
                            }
                        } catch {}
                    }
                }

                if (-not $namespace) {
                    throw "No se pudo obtener el Namespace MAPI de Outlook."
                }

                try {
                    $currentProfile = $namespace.CurrentProfileName
                    if ($currentProfile -ine $targetProfile) {
                        Write-Warning "Perfil actual '$currentProfile' != solicitado '$targetProfile'. Reiniciando Outlook con el perfil '$targetProfile'..."

                        if ($namespace) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null; $namespace = $null }
                        if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null; $outlook = $null }

                        $outlookPath = Get-OutlookExePath
                        Start-Process -FilePath $outlookPath -ArgumentList "/recycle", "/profile", "`"$targetProfile`"" -WindowStyle Hidden

                        $maxWait = 45
                        $waited = 0
                        while ($waited -lt $maxWait) {
                            Start-Sleep -Seconds 1
                            $waited++
                            try {
                                $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
                                $namespace = $outlook.GetNamespace("MAPI")
                                try { $namespace.Logon($ProfileName, "", $false, $false) } catch {}
                                $newProfile = $namespace.CurrentProfileName
                                if ($newProfile -ine $targetProfile) {
                                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null; $namespace = $null
                                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null; $outlook = $null
                                    continue
                                }
                                break
                            } catch {
                                $outlook = $null
                                $namespace = $null
                            }
                        }

                        if (-not $outlook -or -not $namespace) {
                            throw "No se pudo reiniciar Outlook con el perfil '$targetProfile'. Por favor, cierra Outlook manualmente e intenta de nuevo."
                        }

                        $createdOutlook = $true
                        Write-Verbose "Outlook reiniciado exitosamente con el perfil '$targetProfile'"
                    }
                } catch {
                    if ($_.Exception.Message -like "*No se pudo reiniciar Outlook*") { throw }
                }

                break
            } catch {
                $errorCode = $_.Exception.HResult
                if ($errorCode -eq 0xCC54011D -or $errorCode -eq -864313571) {
                    if ($attempt -lt $maxRetries) {
                        Write-Warning "Outlook ocupado (intento $attempt/$maxRetries). Reintentando en $retryDelay segundos..."
                        Start-Sleep -Seconds $retryDelay
                        continue
                    }
                }
                throw
            }
        }

        Write-Verbose "Enumerando stores conectados en Outlook..."
        $stores = Get-ConnectedOutlookStores -namespace $namespace
        $totalStores = if ($stores) { $stores.Count } else { 0 }
        Write-Verbose ("Stores encontrados: {0}" -f $totalStores)

        if ($StoreType) {
            $storesWithType = $stores | Where-Object { $_.storeType }
            $storesWithoutType = $stores | Where-Object { -not $_.storeType -or $_.storeType -eq 'Unknown' }
            $stores = $storesWithType | Where-Object { $_.storeType -eq $StoreType }
            Write-Verbose ("Filtrando StoreType='{0}'. Coincidencias: {1}. Sin clasificar: {2}." -f $StoreType, ($stores.Count), ($storesWithoutType.Count))
        }

        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $baseFileName = "list-store-$timestamp"
        $cwd = (Get-Location).Path
        $shouldExportJson = $ExportResult -eq 'Json'
        $shouldExportText = $ExportResult -eq 'Text'
        $shouldEmitJsonConsole = $Json
        $shouldEmitTextConsole = -not $Json

        if ($shouldEmitJsonConsole -or $shouldExportJson) {
            $payload = [pscustomobject]@{
                type = "stores"
                stores = @($stores)
            }
            $jsonOutput = $payload | ConvertTo-Json -Compress -Depth 6
            if ($shouldEmitJsonConsole) {
                Write-Output $jsonOutput
            }
            if ($shouldExportJson) {
                $jsonPath = Join-Path -Path $cwd -ChildPath "$baseFileName.json"
                $jsonOutput | Out-File -FilePath $jsonPath -Encoding UTF8
                Write-Verbose "Resultado exportado en JSON: $jsonPath"
            }
        }

        $plainLines = $null
        if ($shouldEmitTextConsole -or $shouldExportText) {
            $plainLines = New-Object System.Collections.Generic.List[string]
            $i = 1
            foreach ($s in $stores) {
                $pathDisplay = if ($s.filePath) { $s.filePath } else { "Modo Online / Sin ruta local" }
                $idDisplay = if ($s.storeId) { $s.storeId } else { "Sin StoreID" }
                $line = "[$i] $($s.displayName) - ID: $idDisplay - ($pathDisplay)"
                if ($shouldEmitTextConsole) {
                    Write-Output $line
                }
                if ($shouldExportText) {
                    $plainLines.Add($line) | Out-Null
                }
                $i++
            }
        }

        if ($shouldExportText -and $plainLines -and $plainLines.Count -gt 0) {
            $textPath = Join-Path -Path $cwd -ChildPath "$baseFileName.txt"
            $plainLines | Out-File -FilePath $textPath -Encoding UTF8
            Write-Verbose "Resultado exportado en texto: $textPath"
        }
    } catch {
        if ($Json) {
            $jsonErrorPayload = [pscustomobject]@{
                type = "error"
                message = $_.Exception.Message
            }
        } else {
            throw
        }
    } finally {
        if ($namespace) {
            if (-not $PreserveSession) {
                try { $namespace.Logoff() } catch {}
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
        }
        if ($outlook) {
            if ($createdOutlook -and -not $PreserveSession) {
                try {
                    $outlook.Quit()
                    Start-Sleep -Milliseconds 700
                } catch {}
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        if ($createdOutlook -and -not $PreserveSession -and $createdOutlookPid) {
            try {
                $proc = Get-Process -Id $createdOutlookPid -ErrorAction SilentlyContinue
                if ($proc) {
                    Stop-Process -Id $createdOutlookPid -ErrorAction SilentlyContinue
                }
            } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    if ($jsonErrorPayload) {
        Write-Output ($jsonErrorPayload | ConvertTo-Json -Compress -Depth 4)
        exit 1
    }
}
