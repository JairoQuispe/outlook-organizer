<#
.SYNOPSIS
  Escanea un PST y devuelve la lista plana de carpetas con conteos por año.

.DESCRIPTION
  Este script monta (si es necesario) un archivo PST en Outlook usando MAPI, recorre
  recursivamente sus carpetas y emite progreso/estadísticas en JSON o texto. Es el
  reemplazo del flag -ListFolders que vivía en outlook-import-pst.ps1.
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$PstPath,

    [Parameter(Mandatory = $false)]
    [string]$StoreId,

    [Parameter(Mandatory = $false)]
    [int]$FilterOnlyYear,

    [Parameter()]
    [switch]$IncludeSize,

    [Parameter()]
    [switch]$Summary,

    [Parameter()]
    [ValidateSet("json", "text")]
    [string]$ExportResult = "json",

    [Parameter()]
    [switch]$ExportFolders,

    [Parameter(Mandatory = $false)]
    [string]$ExportFoldersPath,

    [Parameter()]
    [switch]$Json,

    [Parameter()]
    [switch]$Headless,

    [Parameter()]
    [switch]$PreserveSession,

    [Parameter()]
    [Alias('ExportStadistics')]
    [switch]$ExportStatistics,

    [Parameter(Mandatory = $false)]
    [Alias('ExportStadisticsPath')]
    [string]$ExportStatisticsPath
)

# --- Initialization -----------------------------------------------------------

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
$ErrorActionPreference = "Stop"

$script:IsHeadlessOutput = ($Json -or $Headless -or $Summary)
if ($script:IsHeadlessOutput) {
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
}

$script:OutlookApplication = $null
$script:MainNamespace = $null
$script:PstStoreRef = $null
$script:PstRootRef = $null
$script:ScanState = $null
$script:PreserveSessionRequested = [bool]$PreserveSession
$script:CreatedOutlook = $false
$script:CreatedOutlookPid = $null
$script:ExistingOutlookPids = @()
$script:YearFilterEnabled = $PSBoundParameters.ContainsKey('FilterOnlyYear')
$script:IncludeSizeRequested = [bool]$IncludeSize
$script:SummaryRequested = [bool]$Summary
$script:ExportResultRequested = $PSBoundParameters.ContainsKey('ExportResult')
$script:ExportResultFormat = $ExportResult
$script:ExportFoldersRequested = [bool]$ExportFolders
$script:ExportFoldersPathValue = $ExportFoldersPath
$script:ExportStatisticsRequested = [bool]$ExportStatistics
$script:ExportStatisticsPathValue = $ExportStatisticsPath
if ($script:ExportStatisticsRequested) {
    $script:IncludeSizeRequested = $true
}

if (-not $script:PreserveSessionRequested) {
    try {
        $script:ExistingOutlookPids = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
    } catch {}
}

# --- Helper functions ---------------------------------------------------------

function Get-LogTimestamp { (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }

function Emit-Log {
    param([string]$Level, [string]$Message)
    $timestamp = Get-LogTimestamp
    $payload = @{ type = "log"; level = $Level; message = $Message; timestamp = $timestamp }
    if ($Json -or $script:IsHeadlessOutput) { [Console]::WriteLine(($payload | ConvertTo-Json -Compress -Depth 6)) }
    else { Write-Host "[$timestamp] [$Level] $Message" }
}

function Emit-ErrorPayload {
    param([string]$Message)
    $payload = @{ type = "error"; message = $Message }
    if ($Json -or $script:IsHeadlessOutput) { [Console]::WriteLine(($payload | ConvertTo-Json -Compress -Depth 6)) }
    else { Write-Error $Message }
}

function Format-SizeHuman {
    param([long]$Bytes)
    if (-not $Bytes) { return "0 B" }
    $KB = 1024
    $MB = $KB * 1024
    $GB = $MB * 1024
    $TB = $GB * 1024
    if ($Bytes -ge $TB) { return ("{0:N2} TB" -f ($Bytes / $TB)) }
    if ($Bytes -ge $GB) { return ("{0:N2} GB" -f ($Bytes / $GB)) }
    if ($Bytes -ge $MB) { return ("{0:N2} MB" -f ($Bytes / $MB)) }
    if ($Bytes -ge $KB) { return ("{0:N2} KB" -f ($Bytes / $KB)) }
    return ("{0} B" -f $Bytes)
}

function Emit-ScanProgress {
    param([string]$Phase, [string]$FolderPath, [int]$CurrentItemCount, [switch]$FolderCompleted)
    if (-not $script:ScanState) { return }

    $elapsedMs = [long](([DateTime]::UtcNow - $script:ScanState.startedAt).TotalMilliseconds)
    $totalFolders = [int]$script:ScanState.totalFolders
    $scannedFolders = [int]$script:ScanState.scannedFolders
    $percent = 0
    if ($totalFolders -gt 0) {
        $percent = [int][Math]::Floor(($scannedFolders * 100.0) / $totalFolders)
        if ($percent -gt 100) { $percent = 100 }
    }

    $payload = @{
        type = "scanProgress"
        phase = $Phase
        folderPath = $FolderPath
        currentItemCount = [int]$CurrentItemCount
        folderCompleted = [bool]$FolderCompleted
        scannedFolders = $scannedFolders
        totalFolders = $totalFolders
        accumulatedItems = [long]$script:ScanState.accumulatedItems
        percent = $percent
        elapsedMs = $elapsedMs
        pstSizeBytes = [long]$script:ScanState.pstSizeBytes
    }

    if ($Json -or $script:IsHeadlessOutput) { [Console]::WriteLine(($payload | ConvertTo-Json -Compress -Depth 6)) }
    else { Write-Host ("[scan] {0}% {1}/{2} folder={3} items={4}" -f $percent, $scannedFolders, $totalFolders, $FolderPath, $CurrentItemCount) }
}

function Format-SummaryText {
    param($Summary)
    $lines = @()
    $lines += "========================================"
    $lines += "  SCAN SUMMARY"
    $lines += "========================================"
    $lines += "Generated at: $($Summary.generatedAt)"
    $lines += ""
    $lines += "--- Inputs ---"
    $inputs = $Summary.inputs
    $lines += "  PST path:         $($inputs.pstPath)"
    $lines += "  Store ID:         $($inputs.storeId)"
    $lines += "  Filter year:      $(if ($inputs.filterOnlyYear) { $inputs.filterOnlyYear } else { '(all)' })"
    $lines += "  Include size:     $($inputs.includeSize)"
    $lines += "  Preserve session: $($inputs.preserveSession)"
    $lines += ""
    $lines += "--- Source ---"
    $source = $Summary.source
    $lines += "  Resolved path:    $($source.resolvedPstPath)"
    $lines += "  PST size:         $($source.pstSizeHuman) ($($source.pstSizeBytes) bytes)"
    $lines += "  Store ID:         $($source.storeId)"
    $lines += "  Display name:     $($source.storeDisplayName)"
    $lines += "  Already mounted:  $($source.alreadyMounted)"
    $lines += ""
    $lines += "--- Scan ---"
    $scan = $Summary.scan
    $lines += "  Estimated folders: $($scan.estimatedFolders)"
    $lines += "  Scanned folders:   $($scan.scannedFolders)"
    $lines += "  Matched folders:   $($scan.matchedFolders)"
    $lines += "  Elapsed:           $($scan.elapsedMs) ms"
    $lines += "  Accumulated items: $($scan.accumulatedItems)"
    $lines += "  Completed:         $($scan.completed)"
    $lines += ""
    $lines += "--- Totals ---"
    $totals = $Summary.totals
    $lines += "  Items:       $($totals.items)"
    $lines += "  Dated:       $($totals.datedItems)"
    $lines += "  Undated:     $($totals.undatedItems)"
    if ($totals.sizeBytes) {
        $lines += "  Size:        $($totals.sizeHuman) ($($totals.sizeBytes) bytes)"
    }
    if ($Summary.yearBreakdown -and $Summary.yearBreakdown.Count -gt 0) {
        $lines += ""
        $lines += "--- Year Breakdown ---"
        foreach ($row in $Summary.yearBreakdown) {
            $lines += "  $($row.year):  $($row.count) items"
        }
    }
    if ($Summary.topFoldersByItems -and $Summary.topFoldersByItems.Count -gt 0) {
        $lines += ""
        $lines += "--- Top Folders by Items ---"
        foreach ($f in $Summary.topFoldersByItems) {
            $sizePart = if ($f.sizeHuman) { " | $($f.sizeHuman)" } else { "" }
            $lines += "  $($f.path)  ($($f.itemCount) items$sizePart)"
        }
    }
    if ($Summary.topFoldersBySize -and $Summary.topFoldersBySize.Count -gt 0) {
        $lines += ""
        $lines += "--- Top Folders by Size ---"
        foreach ($f in $Summary.topFoldersBySize) {
            $lines += "  $($f.path)  ($($f.sizeHuman) | $($f.itemCount) items)"
        }
    }
    $lines += ""
    $lines += "========================================"
    return ($lines -join "`n")
}

function Write-ExportResult {
    param($Payload)
    $timestamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $ext = if ($script:ExportResultFormat -eq 'text') { 'txt' } else { 'json' }
    $filename = "scan-result-$timestamp.$ext"
    $target = Join-Path (Get-Location) $filename
    try {
        if ($script:ExportResultFormat -eq 'text') {
            if ($Payload.type -eq 'summary') {
                $formatted = Format-SummaryText -Summary $Payload
                $formatted | Out-File -FilePath $target -Encoding UTF8 -Force
            } else {
                $Payload | ConvertTo-Json -Depth 12 | Out-File -FilePath $target -Encoding UTF8 -Force
            }
        } else {
            $Payload | ConvertTo-Json -Compress -Depth 12 | Out-File -FilePath $target -Encoding UTF8 -Force
        }
        Emit-Log "info" "Resultado exportado a $target"
    } catch {
        Emit-Log "warn" "No se pudo exportar el resultado: $($_.Exception.Message)"
    }
}

function Get-FolderExportFileName {
    param([string]$YearFilter, [string]$SourceLabel, [bool]$WithSize)
    $yearPart = if ($YearFilter) { "year-$YearFilter" } else { "all-years" }
    $sizePart = if ($WithSize) { "with-size" } else { "no-size" }
    $safeSrc = $SourceLabel -replace '[\\/:*?"<>|]+', '-'
    if ($safeSrc.Length -gt 40) { $safeSrc = $safeSrc.Substring(0, 40) }
    $ts = (Get-Date).ToString('yyyyMMdd-HHmmss')
    return "scan-$yearPart-$safeSrc-$sizePart-$ts.json"
}

function Export-FolderList {
    param(
        [array]$Folders,
        [string]$PstPathInput,
        [string]$StoreIdInput,
        [int]$FilterOnlyYearInput,
        [bool]$IncludeSizeInput
    )

    $sourceLabel = if ($StoreIdInput) { "store-$StoreIdInput" } elseif ($PstPathInput) { "pst-$([System.IO.Path]::GetFileNameWithoutExtension($PstPathInput))" } else { "unknown" }
    $yearLabel = if ($script:YearFilterEnabled) { "$FilterOnlyYearInput" } else { "" }

    $payload = [ordered]@{
        type           = "folderExport"
        generatedAt    = (Get-Date).ToString('o')
        pstPath        = $PstPathInput
        storeId        = $StoreIdInput
        filterOnlyYear = if ($script:YearFilterEnabled) { [int]$FilterOnlyYearInput } else { $null }
        includeSize    = [bool]$IncludeSizeInput
        folderCount    = @($Folders).Count
        folders        = @($Folders)
    }

    $target = if ($script:ExportFoldersPathValue) {
        $script:ExportFoldersPathValue
    } else {
        Join-Path (Get-Location) (Get-FolderExportFileName -YearFilter $yearLabel -SourceLabel $sourceLabel -WithSize $IncludeSizeInput)
    }

    try {
        $payload | ConvertTo-Json -Depth 12 | Out-File -FilePath $target -Encoding UTF8 -Force
        Emit-Log "info" "Lista de carpetas exportada a $target"
    } catch {
        Emit-Log "warn" "No se pudo exportar las carpetas: $($_.Exception.Message)"
    }
}

function Export-PstStatistics {
    param(
        [array]$Folders,
        [string]$PstPathInput,
        [string]$StoreIdInput
    )

    $globalYearCounts = @{}
    $globalYearSizes = @{}
    $globalMonthCounts = @{}
    $globalMonthSizes = @{}
    
    $totalItems = 0
    $totalSizeBytes = 0
    
    foreach ($f in $Folders) {
        $totalItems += [int]$f.itemCount
        if ($null -ne $f.sizeBytes) {
            $totalSizeBytes += [long]$f.sizeBytes
        }
        
        foreach ($yRow in @($f.yearBreakdown)) {
            $y = [int]$yRow.year
            $c = [int]$yRow.count
            $s = 0
            if ($null -ne $yRow.sizeBytes) { $s = [long]$yRow.sizeBytes }
            
            if ($globalYearCounts.ContainsKey($y)) {
                $globalYearCounts[$y] += $c
                $globalYearSizes[$y] += $s
            } else {
                $globalYearCounts[$y] = $c
                $globalYearSizes[$y] = $s
            }
        }
        
        foreach ($mRow in @($f.monthBreakdown)) {
            $m = [string]$mRow.month
            $c = [int]$mRow.count
            $s = 0
            if ($null -ne $mRow.sizeBytes) { $s = [long]$mRow.sizeBytes }
            
            if ($globalMonthCounts.ContainsKey($m)) {
                $globalMonthCounts[$m] += $c
                $globalMonthSizes[$m] += $s
            } else {
                $globalMonthCounts[$m] = $c
                $globalMonthSizes[$m] = $s
            }
        }
    }
    
    $globalYearBreakdown = @()
    foreach ($y in ($globalYearCounts.Keys | Sort-Object -Descending)) {
        $globalYearBreakdown += [pscustomobject]@{
            year = [int]$y
            count = [int]$globalYearCounts[$y]
            sizeBytes = [long]$globalYearSizes[$y]
            sizeHuman = Format-SizeHuman -Bytes ([long]$globalYearSizes[$y])
        }
    }
    
    $globalMonthBreakdown = @()
    foreach ($m in ($globalMonthCounts.Keys | Sort-Object -Descending)) {
        $globalMonthBreakdown += [pscustomobject]@{
            month = [string]$m
            count = [int]$globalMonthCounts[$m]
            sizeBytes = [long]$globalMonthSizes[$m]
            sizeHuman = Format-SizeHuman -Bytes ([long]$globalMonthSizes[$m])
        }
    }
    
    $sourceLabel = if ($StoreIdInput) { "store-$StoreIdInput" } elseif ($PstPathInput) { "pst-$([System.IO.Path]::GetFileNameWithoutExtension($PstPathInput))" } else { "unknown" }
    
    $payload = [ordered]@{
        type           = "pstStatistics"
        generatedAt    = (Get-Date).ToString('o')
        pstPath        = $PstPathInput
        storeId        = $StoreIdInput
        pstSizeBytes   = [long]$script:ScanState.pstSizeBytes
        pstSizeHuman   = Format-SizeHuman -Bytes ([long]$script:ScanState.pstSizeBytes)
        foldersCount   = @($Folders).Count
        globalStats = [ordered]@{
            totalItems = [long]$totalItems
            totalSizeBytes = [long]$totalSizeBytes
            totalSizeHuman = Format-SizeHuman -Bytes $totalSizeBytes
            yearBreakdown = @($globalYearBreakdown)
            monthBreakdown = @($globalMonthBreakdown)
        }
        folders        = @($Folders)
    }
    
    $target = if ($script:ExportStatisticsPathValue) {
        $script:ExportStatisticsPathValue
    } else {
        $ts = (Get-Date).ToString('yyyyMMdd-HHmmss')
        $safeSrc = $sourceLabel -replace '[\\/:*?"<>|]+', '-'
        if ($safeSrc.Length -gt 40) { $safeSrc = $safeSrc.Substring(0, 40) }
        Join-Path (Get-Location) "pst-stats-$safeSrc-$ts.json"
    }
    
    try {
        $payload | ConvertTo-Json -Depth 12 | Out-File -FilePath $target -Encoding UTF8 -Force
        Emit-Log "info" "Estadisticas de PST exportadas a $target"
    } catch {
        Emit-Log "warn" "No se pudo exportar las estadisticas: $($_.Exception.Message)"
    }
}

# --- COM management -----------------------------------------------------------

function Release-ComObjectSafe {
    param([object]$ComObject)
    if ($null -eq $ComObject) { return }
    try {
        if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
    } catch {}
}

function Cleanup-ComResources {
    try {
        if ($null -ne $script:MainNamespace -and -not $script:PreserveSessionRequested) {
            try { $script:MainNamespace.Logoff() } catch {}
        }

        if ($null -ne $script:OutlookApplication -and $script:CreatedOutlook -and -not $script:PreserveSessionRequested) {
            try {
                $script:OutlookApplication.Quit()
                Start-Sleep -Milliseconds 700
            } catch {}
        }

        Release-ComObjectSafe $script:PstRootRef
        Release-ComObjectSafe $script:PstStoreRef
        Release-ComObjectSafe $script:MainNamespace
        Release-ComObjectSafe $script:OutlookApplication
    } catch {}

    if ($script:CreatedOutlook -and -not $script:PreserveSessionRequested) {
        try {
            if ($script:CreatedOutlookPid) {
                $proc = Get-Process -Id $script:CreatedOutlookPid -ErrorAction SilentlyContinue
                if ($proc) { Stop-Process -Id $script:CreatedOutlookPid -ErrorAction SilentlyContinue }
            } else {
                $currentPids = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
                foreach ($pid in $currentPids) {
                    if ($script:ExistingOutlookPids -notcontains $pid) {
                        try { Stop-Process -Id $pid -ErrorAction SilentlyContinue } catch {}
                    }
                }
            }
        } catch {}
    }

    $script:PstRootRef = $null
    $script:PstStoreRef = $null
    $script:MainNamespace = $null
    $script:OutlookApplication = $null
    $script:CreatedOutlookPid = $null
    $script:ExistingOutlookPids = @()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Exit-WithCleanup {
    param([int]$Code)
    Cleanup-ComResources
    exit $Code
}

function Get-OutlookNamespace {
    if ($script:MainNamespace -and $script:OutlookApplication) { return $script:MainNamespace }

    $progId = "Outlook.Application"
    $maxRetries = 3
    $retryDelay = 2

    if (-not $script:PreserveSessionRequested) {
        try {
            $outlookProbe = New-Object -ComObject $progId -ErrorAction Stop
            try { $outlookProbe.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlookProbe) | Out-Null
            $outlookProbe = $null
        } catch {
            Emit-ErrorPayload "Outlook no parece estar instalado o registrado (ProgID '$progId')."
            Exit-WithCleanup 1
        }
    }

    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        $outlook = $null
        $namespace = $null
        try {
            try {
                $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
                $namespace = $outlook.GetNamespace("MAPI")
                if (-not $script:PreserveSessionRequested) {
                    try { $namespace.Logon("", "", $false, $false) } catch {}
                }
            } catch {
                if ($script:PreserveSessionRequested) {
                    throw "Outlook debe estar abierto para usar -PreserveSession."
                }
                $outlook = New-Object -ComObject Outlook.Application
                $namespace = $outlook.GetNamespace("MAPI")
                try { $namespace.Logon("", "", $false, $true) } catch {}
                $script:CreatedOutlook = $true
                if (-not $script:CreatedOutlookPid) {
                    try {
                        $currentPids = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
                        if ($currentPids) {
                            $script:CreatedOutlookPid = ($currentPids | Where-Object { $script:ExistingOutlookPids -notcontains $_ }) | Select-Object -First 1
                        }
                    } catch {}
                }
            }

            if (-not $namespace) {
                throw "No se pudo obtener el Namespace MAPI de Outlook."
            }

            $script:OutlookApplication = $outlook
            $script:MainNamespace = $namespace
            return $script:MainNamespace
        } catch {
            if ($namespace) { Release-ComObjectSafe $namespace }
            if ($outlook) { Release-ComObjectSafe $outlook }

            $errorCode = $_.Exception.HResult
            if (($errorCode -eq 0xCC54011D -or $errorCode -eq -864313571) -and ($attempt -lt $maxRetries)) {
                Emit-Log "warn" "Outlook ocupado (intento $attempt/$maxRetries). Reintentando en $retryDelay s..."
                Start-Sleep -Seconds $retryDelay
                continue
            }

            Emit-ErrorPayload "No se pudo iniciar Outlook: $($_.Exception.Message)"
            Exit-WithCleanup 1
        }
    }
}

function Get-StoreByIdOrPath {
    param($namespace, [string]$StoreId, [string]$FilePath)
    foreach ($s in $namespace.Stores) {
        if ($StoreId -and ($s.StoreID -eq $StoreId)) { return $s }
        if ($FilePath) {
            try {
                if ($s.FilePath -and ($s.FilePath -eq $FilePath)) { return $s }
            } catch {}
        }
    }
    return $null
}

function Get-SubFolders-Safe {
    param($parentFolder)
    try { return $parentFolder.Folders } catch { return @() }
}

function Collect-PstFoldersRecursive {
    param($folder, [string]$pathPrefix, [ref]$out)
    $folderPath = $null
    $folderSizeBytes = 0
    try {
        $folderPath = if ($pathPrefix) { "$pathPrefix\$($folder.Name)" } else { $folder.Name }
        $count = 0
        $yearCounts = @{}
        $yearSizes = @{}
        $monthCounts = @{}
        $monthSizes = @{}
        $oldestDate = $null
        $newestDate = $null
        $usedTable = $false
        try {
            $table = $folder.GetTable("")
            $table.Columns.RemoveAll()
            $table.Columns.Add("ReceivedTime") | Out-Null
            $table.Columns.Add("CreationTime") | Out-Null
            try { $table.Columns.Add("SentOn") | Out-Null } catch {}
            try { $table.Columns.Add("LastModificationTime") | Out-Null } catch {}
            if ($script:IncludeSizeRequested) {
                try { $table.Columns.Add("Size") | Out-Null } catch {}
            }

            while (-not $table.EndOfTable) {
                $rows = $table.GetNextRows(300)
                foreach ($row in $rows) {
                    $d = $row["ReceivedTime"]
                    if (-not $d) { $d = $row["CreationTime"] }
                    if (-not $d) {
                        try { $d = $row["SentOn"] } catch {}
                    }
                    if (-not $d) {
                        try { $d = $row["LastModificationTime"] } catch {}
                    }
                    $includeItem = $true
                    if ($script:YearFilterEnabled) {
                        if (-not $d) {
                            $includeItem = $false
                        } else {
                            $includeItem = ([int]$d.Year -eq $FilterOnlyYear)
                        }
                    }
                    if (-not $includeItem) { continue }

                    $rowSize = 0
                    if ($script:IncludeSizeRequested) {
                        try { $rowSize = [long]$row["Size"] } catch {}
                        $folderSizeBytes += [long]$rowSize
                    }

                    $count++
                    if (($count % 2000) -eq 0) {
                        Emit-ScanProgress -Phase "folder_scan" -FolderPath $folderPath -CurrentItemCount $count
                    }
                    if ($d) {
                        $y = [int]$d.Year
                        $m = $d.ToString("yyyy-MM")
                        if ($yearCounts.ContainsKey($y)) { $yearCounts[$y]++ } else { $yearCounts[$y] = 1 }
                        if ($script:IncludeSizeRequested) {
                            if ($yearSizes.ContainsKey($y)) { $yearSizes[$y] += [long]$rowSize } else { $yearSizes[$y] = [long]$rowSize }
                        }
                        if ($monthCounts.ContainsKey($m)) { $monthCounts[$m]++ } else { $monthCounts[$m] = 1 }
                        if ($script:IncludeSizeRequested) {
                            if ($monthSizes.ContainsKey($m)) { $monthSizes[$m] += [long]$rowSize } else { $monthSizes[$m] = [long]$rowSize }
                        }
                        if ($null -eq $oldestDate -or $d -lt $oldestDate) { $oldestDate = $d }
                        if ($null -eq $newestDate -or $d -gt $newestDate) { $newestDate = $d }
                    }
                }
            }
            $usedTable = $true
        } catch {
            try {
                $count = 0
            } catch {}
        }

        if (-not $usedTable) {
            try {
                $scanCount = 0
                foreach ($item in $folder.Items) {
                    $d = $null
                    $scanCount++
                    if (($scanCount % 2000) -eq 0) {
                        Emit-ScanProgress -Phase "folder_scan" -FolderPath $folderPath -CurrentItemCount $scanCount
                    }
                    try { $d = $item.ReceivedTime } catch {}
                    if (-not $d) {
                        try { $d = $item.CreationTime } catch {}
                    }
                    if (-not $d) {
                        try { $d = $item.SentOn } catch {}
                    }
                    if (-not $d) {
                        try { $d = $item.LastModificationTime } catch {}
                    }
                    $includeItem = $true
                    if ($script:YearFilterEnabled) {
                        if (-not $d) {
                            $includeItem = $false
                        } else {
                            $includeItem = ([int]$d.Year -eq $FilterOnlyYear)
                        }
                    }
                    if (-not $includeItem) { continue }

                    $itemSize = 0
                    if ($script:IncludeSizeRequested) {
                        try { $itemSize = [long]$item.Size } catch {}
                        $folderSizeBytes += [long]$itemSize
                    }

                    $count++
                    if ($d) {
                        $y = [int]$d.Year
                        $m = $d.ToString("yyyy-MM")
                        if ($yearCounts.ContainsKey($y)) { $yearCounts[$y]++ } else { $yearCounts[$y] = 1 }
                        if ($script:IncludeSizeRequested) {
                            if ($yearSizes.ContainsKey($y)) { $yearSizes[$y] += [long]$itemSize } else { $yearSizes[$y] = [long]$itemSize }
                        }
                        if ($monthCounts.ContainsKey($m)) { $monthCounts[$m]++ } else { $monthCounts[$m] = 1 }
                        if ($script:IncludeSizeRequested) {
                            if ($monthSizes.ContainsKey($m)) { $monthSizes[$m] += [long]$itemSize } else { $monthSizes[$m] = [long]$itemSize }
                        }
                        if ($null -eq $oldestDate -or $d -lt $oldestDate) { $oldestDate = $d }
                        if ($null -eq $newestDate -or $d -gt $newestDate) { $newestDate = $d }
                    }
                }
            } catch {}
        }

        $yearBreakdown = @()
        foreach ($y in ($yearCounts.Keys | Sort-Object -Descending)) {
            $ySizeBytes = if ($script:IncludeSizeRequested) { [long]$yearSizes[$y] } else { $null }
            $ySizeHuman = if ($script:IncludeSizeRequested) { Format-SizeHuman -Bytes $ySizeBytes } else { $null }
            $yearBreakdown += [pscustomobject]@{
                year = [int]$y
                count = [int]$yearCounts[$y]
                sizeBytes = $ySizeBytes
                sizeHuman = $ySizeHuman
            }
        }

        $monthBreakdown = @()
        foreach ($m in ($monthCounts.Keys | Sort-Object -Descending)) {
            $mSizeBytes = if ($script:IncludeSizeRequested) { [long]$monthSizes[$m] } else { $null }
            $mSizeHuman = if ($script:IncludeSizeRequested) { Format-SizeHuman -Bytes $mSizeBytes } else { $null }
            $monthBreakdown += [pscustomobject]@{
                month = [string]$m
                count = [int]$monthCounts[$m]
                sizeBytes = $mSizeBytes
                sizeHuman = $mSizeHuman
            }
        }

        $datedCount = 0
        foreach ($k in $yearCounts.Keys) { $datedCount += [int]$yearCounts[$k] }
        $undatedCount = [int]([Math]::Max(0, $count - $datedCount))

        if ($script:YearFilterEnabled -and $count -le 0) {
            if ($script:ScanState) {
                $script:ScanState.scannedFolders = [int]$script:ScanState.scannedFolders + 1
                Emit-ScanProgress -Phase "folder_done" -FolderPath $folderPath -CurrentItemCount 0 -FolderCompleted
            }
            foreach ($sub in (Get-SubFolders-Safe -parentFolder $folder)) {
                Collect-PstFoldersRecursive -folder $sub -pathPrefix $folderPath -out $out
            }
            return
        }

        $folderObj = [ordered]@{
            type = "folder"
            path = $folderPath
            itemCount = $count
            sizeBytes = if ($script:IncludeSizeRequested) { [long]$folderSizeBytes } else { $null }
            sizeHuman = if ($script:IncludeSizeRequested) { Format-SizeHuman -Bytes $folderSizeBytes } else { $null }
            yearBreakdown = @($yearBreakdown)
            undatedCount = $undatedCount
        }

        if ($script:ExportStatisticsRequested) {
            $folderObj["monthBreakdown"] = @($monthBreakdown)
            $folderObj["oldestEmail"] = if ($oldestDate) { $oldestDate.ToString("yyyy-MM-ddTHH:mm:ssK") } else { $null }
            $folderObj["newestEmail"] = if ($newestDate) { $newestDate.ToString("yyyy-MM-ddTHH:mm:ssK") } else { $null }
        }

        $out.Value += [pscustomobject]$folderObj

        if ($script:ScanState) {
            $script:ScanState.scannedFolders = [int]$script:ScanState.scannedFolders + 1
            $script:ScanState.accumulatedItems = [long]$script:ScanState.accumulatedItems + [long]$count
            Emit-ScanProgress -Phase "folder_done" -FolderPath $folderPath -CurrentItemCount $count -FolderCompleted
        }

        foreach ($sub in (Get-SubFolders-Safe -parentFolder $folder)) {
            Collect-PstFoldersRecursive -folder $sub -pathPrefix $folderPath -out $out
        }
    } catch {
        $pathLabel = if ($folderPath) { " '$folderPath'" } else { "" }
        Emit-Log "warn" "Error recorriendo${pathLabel}: $($_.Exception.Message)"
    }
}

function Count-PstFoldersRecursive {
    param($folder)
    $count = 1
    foreach ($sub in (Get-SubFolders-Safe -parentFolder $folder)) {
        $count += Count-PstFoldersRecursive -folder $sub
    }
    return [int]$count
}

function Build-ScanSummaryPayload {
    param(
        [array]$Folders,
        [string]$PstPathInput,
        [string]$PstPathResolved,
        [string]$StoreIdInput,
        $Store,
        [bool]$AlreadyMounted,
        [int]$TotalFoldersEstimated
    )

    $elapsedMs = 0
    if ($script:ScanState -and $script:ScanState.startedAt) {
        $elapsedMs = [long](([DateTime]::UtcNow - $script:ScanState.startedAt).TotalMilliseconds)
    }

    $storeResolvedId = $null
    $storeDisplayName = $null
    try { $storeResolvedId = [string]$Store.StoreID } catch {}
    try { $storeDisplayName = [string]$Store.DisplayName } catch {}

    $totalItems = [long]0
    $totalUndatedItems = [long]0
    $totalSizeBytes = [long]0
    $globalYearCounts = @{}

    foreach ($folder in $Folders) {
        $folderItemCount = 0
        $folderUndated = 0
        try { $folderItemCount = [int]$folder.itemCount } catch {}
        try { $folderUndated = [int]$folder.undatedCount } catch {}

        $totalItems += [long]$folderItemCount
        $totalUndatedItems += [long]$folderUndated

        if ($script:IncludeSizeRequested) {
            $folderSize = 0
            try { $folderSize = [long]$folder.sizeBytes } catch {}
            $totalSizeBytes += [long]$folderSize
        }

        foreach ($yearRow in @($folder.yearBreakdown)) {
            $yearKey = 0
            $yearCount = 0
            try { $yearKey = [int]$yearRow.year } catch {}
            try { $yearCount = [int]$yearRow.count } catch {}
            if ($globalYearCounts.ContainsKey($yearKey)) {
                $globalYearCounts[$yearKey] += $yearCount
            } else {
                $globalYearCounts[$yearKey] = $yearCount
            }
        }
    }

    $globalYearBreakdown = @()
    foreach ($y in ($globalYearCounts.Keys | Sort-Object -Descending)) {
        $globalYearBreakdown += [pscustomobject]@{ year = [int]$y; count = [int]$globalYearCounts[$y] }
    }

    $topFoldersByItems = @(
        $Folders |
            Sort-Object -Property @{ Expression = { [int]$_.itemCount }; Descending = $true }, @{ Expression = { [string]$_.path }; Descending = $false } |
            Select-Object -First 25 |
            ForEach-Object {
                [pscustomobject]@{
                    path = [string]$_.path
                    itemCount = [int]$_.itemCount
                    sizeBytes = if ($script:IncludeSizeRequested) { [long]$_.sizeBytes } else { $null }
                    sizeHuman = if ($script:IncludeSizeRequested) { [string]$_.sizeHuman } else { $null }
                }
            }
    )

    $topFoldersBySize = @()
    if ($script:IncludeSizeRequested) {
        $topFoldersBySize = @(
            $Folders |
                Sort-Object -Property @{ Expression = { [long]$_.sizeBytes }; Descending = $true }, @{ Expression = { [string]$_.path }; Descending = $false } |
                Select-Object -First 25 |
                ForEach-Object {
                    [pscustomobject]@{
                        path = [string]$_.path
                        sizeBytes = [long]$_.sizeBytes
                        sizeHuman = [string]$_.sizeHuman
                        itemCount = [int]$_.itemCount
                    }
                }
        )
    }

    $matchedFolders = @($Folders).Count

    return [pscustomobject]@{
        type = "summary"
        generatedAt = (Get-Date).ToString("o")
        inputs = [pscustomobject]@{
            pstPath = $PstPathInput
            storeId = $StoreIdInput
            filterOnlyYear = if ($script:YearFilterEnabled) { [int]$FilterOnlyYear } else { $null }
            includeSize = [bool]$script:IncludeSizeRequested
            preserveSession = [bool]$script:PreserveSessionRequested
            summary = [bool]$script:SummaryRequested
            json = [bool]$Json
            headless = [bool]$Headless
        }
        source = [pscustomobject]@{
            resolvedPstPath = $PstPathResolved
            pstSizeBytes = [long]$script:ScanState.pstSizeBytes
            pstSizeHuman = Format-SizeHuman -Bytes ([long]$script:ScanState.pstSizeBytes)
            storeId = $storeResolvedId
            storeDisplayName = $storeDisplayName
            alreadyMounted = [bool]$AlreadyMounted
        }
        scan = [pscustomobject]@{
            estimatedFolders = [int]$TotalFoldersEstimated
            scannedFolders = [int]$script:ScanState.scannedFolders
            matchedFolders = [int]$matchedFolders
            elapsedMs = [long]$elapsedMs
            accumulatedItems = [long]$script:ScanState.accumulatedItems
            completed = $true
        }
        totals = [pscustomobject]@{
            items = [long]$totalItems
            datedItems = [long]([Math]::Max(0, $totalItems - $totalUndatedItems))
            undatedItems = [long]$totalUndatedItems
            sizeBytes = if ($script:IncludeSizeRequested) { [long]$totalSizeBytes } else { $null }
            sizeHuman = if ($script:IncludeSizeRequested) { Format-SizeHuman -Bytes $totalSizeBytes } else { $null }
        }
        yearBreakdown = @($globalYearBreakdown)
        topFoldersByItems = @($topFoldersByItems)
        topFoldersBySize = if ($script:IncludeSizeRequested) { @($topFoldersBySize) } else { @() }
        folders = @($Folders)
    }
}

if (-not $PstPath -and -not $StoreId) {
    Emit-ErrorPayload "Debe proporcionar -PstPath o -StoreId."
    Exit-WithCleanup 1
}

if ($script:YearFilterEnabled -and ($FilterOnlyYear -lt 1900 -or $FilterOnlyYear -gt 2100)) {
    Emit-ErrorPayload "-FilterOnlyYear debe estar entre 1900 y 2100."
    Exit-WithCleanup 1
}

if ($PstPath) {
    if (-not (Test-Path -LiteralPath $PstPath)) {
        Emit-ErrorPayload "PST no encontrado: $PstPath"
        Exit-WithCleanup 1
    }
}

$namespace = Get-OutlookNamespace

$alreadyMounted = $false
$pstStore = $null

if ($StoreId) {
    $pstStore = Get-StoreByIdOrPath -namespace $namespace -StoreId $StoreId
    if (-not $pstStore) {
        Emit-ErrorPayload "No se encontró el store con StoreId=$StoreId."
        Exit-WithCleanup 1
    }
    $alreadyMounted = $true
    Emit-Log "info" "Usando store existente (StoreId=$StoreId)."
} elseif ($PstPath) {
    $pstStore = Get-StoreByIdOrPath -namespace $namespace -FilePath $PstPath
    if ($pstStore) {
        $alreadyMounted = $true
        Emit-Log "info" "PST ya estaba montado."
    } else {
        Emit-Log "info" "Montando PST: $PstPath"
        try {
            $namespace.AddStoreEx($PstPath, 3)
        } catch {
            Emit-ErrorPayload "No se pudo montar el PST: $($_.Exception.Message)"
            Exit-WithCleanup 1
        }
        $pstStore = Get-StoreByIdOrPath -namespace $namespace -FilePath $PstPath
        if (-not $pstStore) {
            Emit-ErrorPayload "PST montado pero no localizado."
            Exit-WithCleanup 1
        }
    }
}

if (-not $pstStore) {
    Emit-ErrorPayload "No se pudo determinar el store a escanear."
    Exit-WithCleanup 1
}

$script:PstStoreRef = $pstStore
$pstRoot = $pstStore.GetRootFolder()
$script:PstRootRef = $pstRoot

$pstResolvedPath = $null
if ($PstPath) {
    $pstResolvedPath = $PstPath
} else {
    try { $pstResolvedPath = [string]$pstStore.FilePath } catch { $pstResolvedPath = $null }
}

$pstSizeBytes = 0
if ($pstResolvedPath -and (Test-Path -LiteralPath $pstResolvedPath)) {
    try { $pstSizeBytes = [long](Get-Item -LiteralPath $pstResolvedPath -ErrorAction Stop).Length } catch {}
}

$totalFoldersToScan = 0
foreach ($tf in (Get-SubFolders-Safe -parentFolder $pstRoot)) {
    $totalFoldersToScan += Count-PstFoldersRecursive -folder $tf
}

$script:ScanState = @{
    totalFolders = [int]$totalFoldersToScan
    scannedFolders = 0
    accumulatedItems = [long]0
    pstSizeBytes = [long]$pstSizeBytes
    startedAt = [DateTime]::UtcNow
}

if ($Json -or $script:IsHeadlessOutput) {
    [Console]::WriteLine((@{ type = "scanMeta"; pstPath = $PstPath; pstSizeBytes = [long]$pstSizeBytes; totalFolders = [int]$totalFoldersToScan; filterOnlyYear = if ($script:YearFilterEnabled) { [int]$FilterOnlyYear } else { $null }; includeSize = [bool]$script:IncludeSizeRequested; summary = [bool]$script:SummaryRequested } | ConvertTo-Json -Compress -Depth 6))
} else {
    $filterLabel = if ($script:YearFilterEnabled) { " | filtro año: $FilterOnlyYear" } else { "" }
    $sizeLabel = if ($script:IncludeSizeRequested) { " | con size" } else { "" }
    $summaryLabel = if ($script:SummaryRequested) { " | modo summary" } else { "" }
    Emit-Log "info" "Escaneando PST... Carpetas estimadas: $totalFoldersToScan$filterLabel$sizeLabel$summaryLabel"
}

$flat = [ref]@()
foreach ($tf in (Get-SubFolders-Safe -parentFolder $pstRoot)) {
    Collect-PstFoldersRecursive -folder $tf -pathPrefix "" -out $flat
}

Emit-ScanProgress -Phase "completed" -FolderPath "" -CurrentItemCount 0 -FolderCompleted

$finalPayload = $null
if ($script:SummaryRequested) {
    $finalPayload = Build-ScanSummaryPayload -Folders @($flat.Value) -PstPathInput $PstPath -PstPathResolved $pstResolvedPath -StoreIdInput $StoreId -Store $pstStore -AlreadyMounted $alreadyMounted -TotalFoldersEstimated $totalFoldersToScan
    [Console]::WriteLine(($finalPayload | ConvertTo-Json -Compress -Depth 12))
} else {
    $finalPayload = [pscustomobject]@{
        type = "folders"
        count = @($flat.Value).Count
        folders = @($flat.Value)
    }
    [Console]::WriteLine((@{ type = "folders"; count = @($flat.Value).Count } | ConvertTo-Json -Compress -Depth 6))
    foreach ($f in $flat.Value) {
        [Console]::WriteLine(($f | ConvertTo-Json -Compress -Depth 6))
    }
}

if ($script:ExportResultRequested -and $finalPayload) {
    Write-ExportResult -Payload $finalPayload
}

if ($script:ExportFoldersRequested -and @($flat.Value).Count -gt 0) {
    Export-FolderList -Folders @($flat.Value) -PstPathInput $PstPath -StoreIdInput $StoreId -FilterOnlyYearInput $FilterOnlyYear -IncludeSizeInput $script:IncludeSizeRequested
}

if ($script:ExportStatisticsRequested -and @($flat.Value).Count -gt 0) {
    Export-PstStatistics -Folders @($flat.Value) -PstPathInput $PstPath -StoreIdInput $StoreId
}

if (-not $alreadyMounted -and $PstPath) {
    try { $namespace.RemoveStore($pstRoot) } catch {}
}

Exit-WithCleanup 0
