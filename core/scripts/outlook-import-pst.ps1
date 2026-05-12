# Importar correos desde PST a buzón de Outlook/Exchange usando COM
param (
    [Parameter(Mandatory=$false)]
    [string]$PstPath,

    [Parameter(Mandatory=$false)]
    [string]$TargetStoreId,

    [Parameter(Mandatory=$false)]
    [ValidateSet("Copy", "Move", "copy", "move")]
    [string]$Action = "Copy",

    [Parameter(Mandatory=$false)]
    [int]$FilterOnlyYear,

    [Parameter(Mandatory=$false)]
    [string]$FilterOnlyMonths,

    [Parameter(Mandatory=$false)]
    [switch]$SkipDuplicates,

    [Parameter(Mandatory=$false)]
    [int]$ItemsPerMinute = 120,

    [Parameter(Mandatory=$false)]
    [int]$BurstSize = 20,

    [Parameter(Mandatory=$false)]
    [int]$MaxRetries = 5,

    [Parameter(Mandatory=$false)]
    [int]$InitialBackoffMs = 1000,

    [Parameter(Mandatory=$false)]
    [int]$MaxBackoffMs = 30000,

    [Parameter(Mandatory=$false)]
    [switch]$AdaptiveThrottling,

    [Parameter(Mandatory=$false)]
    [string]$IncludeFoldersJson,

    [Parameter(Mandatory=$false)]
    [string[]]$IncludeFolders,

    [Parameter(Mandatory=$false)]
    [switch]$ListFolders,

    [Parameter(Mandatory=$false)]
    [switch]$Json,

    [Parameter(Mandatory=$false)]
    [switch]$Headless,

    [Parameter(Mandatory=$false)]
    [switch]$DeepDuplicateCheck,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0, 7200)]
    [int]$DuplicateIndexInactivityTimeoutSec = 180,

    [Parameter(Mandatory=$false)]
    [long]$MaxItemSizeBytes = 157286400,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0, 100000)]
    [int]$MaxFailureRecords = 1000
)

$ErrorActionPreference = "Stop"

$script:IsHeadlessOutput = ($Json -or $Headless)
if ($script:IsHeadlessOutput) {
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
}

function Publish-Progress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete,
        [switch]$Completed,
        [int]$Copied,
        [int]$Moved,
        [int]$Skipped,
        [int]$Failed
    )

    if ($Json -or $script:IsHeadlessOutput) {
        $payload = @{
            type = "progress"
            activity = $Activity
            status = $Status
            percent = $PercentComplete
            completed = [bool]$Completed
            copied = $Copied
            moved = $Moved
            skipped = $Skipped
            failed = $Failed
        }
        [Console]::WriteLine(($payload | ConvertTo-Json -Compress -Depth 6))
    } else {
        $statusMessage = "$Status | Copiados=$Copied Movidos=$Moved Omitidos=$Skipped Fallidos=$Failed"
        if ($Completed) {
            Microsoft.PowerShell.Utility\Write-Progress -Activity $Activity -Status $statusMessage -PercentComplete 100 -Completed
        } else {
            Microsoft.PowerShell.Utility\Write-Progress -Activity $Activity -Status $statusMessage -PercentComplete $PercentComplete
        }
    }
}

function Release-ComObjectSafe {
    param([object]$ComObject)
    if ($null -eq $ComObject) { return }
    try {
        if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
    } catch {}
}

$script:OutlookApplication = $null
$script:MainNamespace = $null
$script:PstStoreRef = $null
$script:PstRootRef = $null
$script:TargetStoreRef = $null
$script:TargetRootRef = $null
$script:FilterOnlyMonthLookup = $null

function Cleanup-ComResources {
    if ($script:ChildFolderCache) {
        try {
            foreach ($childDict in $script:ChildFolderCache.Values) {
                if (-not $childDict) { continue }
                try {
                    foreach ($folderObj in $childDict.Values) {
                        Release-ComObjectSafe $folderObj
                    }
                } catch {}
                try { $childDict.Clear() } catch {}
            }
            $script:ChildFolderCache.Clear()
        } catch {}
    }
    $script:ChildFolderCache = $null
    $script:FilterOnlyMonthLookup = $null
    Release-ComObjectSafe $script:TargetRootRef
    Release-ComObjectSafe $script:TargetStoreRef
    Release-ComObjectSafe $script:PstRootRef
    Release-ComObjectSafe $script:PstStoreRef
    Release-ComObjectSafe $script:MainNamespace
    Release-ComObjectSafe $script:OutlookApplication
    $script:TargetRootRef = $null
    $script:TargetStoreRef = $null
    $script:PstRootRef = $null
    $script:PstStoreRef = $null
    $script:MainNamespace = $null
    $script:OutlookApplication = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Exit-WithCleanup {
    param([int]$Code)
    Cleanup-ComResources
    exit $Code
}

function Add-FailureRecord {
    param(
        [ref]$stats,
        [hashtable]$Record
    )

    if ($MaxFailureRecords -le 0 -or $stats.Value.failures.Count -lt $MaxFailureRecords) {
        $stats.Value.failures += $Record
    } else {
        $stats.Value.failureOverflow = [int]$stats.Value.failureOverflow + 1
    }
}

function Get-LogTimestamp {
    return (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
}

function Emit-Log {
    param([string]$Level, [string]$Message)
    $timestamp = Get-LogTimestamp
    $payload = @{ type = "log"; level = $Level; message = $Message; timestamp = $timestamp }
    if ($Json -or $script:IsHeadlessOutput) {
        [Console]::WriteLine(($payload | ConvertTo-Json -Compress -Depth 6))
    } else {
        Write-Host "[$timestamp] [$Level] $Message"
    }
}

function Emit-ErrorPayload {
    param([string]$Message)
    $payload = @{ type = "error"; message = $Message }
    if ($Json -or $script:IsHeadlessOutput) {
        [Console]::WriteLine(($payload | ConvertTo-Json -Compress -Depth 6))
    } else {
        Write-Error $Message
    }
}

function Emit-ScanProgress {
    param(
        [string]$Phase,
        [string]$FolderPath,
        [int]$CurrentItemCount,
        [switch]$FolderCompleted
    )

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

    if ($Json -or $script:IsHeadlessOutput) {
        [Console]::WriteLine(($payload | ConvertTo-Json -Compress -Depth 6))
    } else {
        Write-Host ("[scan] {0}% {1}/{2} folder={3} items={4}" -f $percent, $scannedFolders, $totalFolders, $FolderPath, $CurrentItemCount)
    }
}

function Format-Bytes {
    param ($bytes)
    if ($bytes -ge 1GB) { return "{0:N2} GB" -f ($bytes / 1GB) }
    elseif ($bytes -ge 1MB) { return "{0:N2} MB" -f ($bytes / 1MB) }
    elseif ($bytes -ge 1KB) { return "{0:N2} KB" -f ($bytes / 1KB) }
    else { return "$bytes Bytes" }
}

function Get-SafeSubject {
    param($item)
    try { return [string]$item.Subject } catch { return "(unknown)" }
}

function Normalize-FolderPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    return $Path.Trim().Replace("/", "\").Trim("\").ToLowerInvariant()
}

function Should-ProcessFolder {
    param([string]$FolderPath, [string[]]$SelectedFolders)
    if (-not $SelectedFolders -or $SelectedFolders.Count -eq 0) { return $true }
    $fp = Normalize-FolderPath $FolderPath
    foreach ($sel in $SelectedFolders) {
        if ($fp -eq $sel -or $fp.StartsWith("$sel\")) { return $true }
    }
    return $false
}

function Has-SelectedDescendant {
    param([string]$FolderPath, [string[]]$SelectedFolders)
    if (-not $SelectedFolders -or $SelectedFolders.Count -eq 0) { return $false }
    $fp = Normalize-FolderPath $FolderPath
    foreach ($sel in $SelectedFolders) {
        if ($sel.StartsWith("$fp\")) { return $true }
    }
    return $false
}

function Collect-PstFoldersRecursive {
    param($folder, [string]$pathPrefix, [ref]$out)
    $folderPath = $null
    try {
        $folderPath = if ($pathPrefix) { "$pathPrefix\$($folder.Name)" } else { $folder.Name }
        $count = 0
        $yearCounts = @{}
        $usedTable = $false
        try {
            $table = $folder.GetTable("")
            $table.Columns.RemoveAll()
            $table.Columns.Add("ReceivedTime") | Out-Null
            $table.Columns.Add("CreationTime") | Out-Null
            try { $table.Columns.Add("SentOn") | Out-Null } catch {}
            try { $table.Columns.Add("LastModificationTime") | Out-Null } catch {}

            while (-not $table.EndOfTable) {
                $rows = $table.GetNextRows(300)
                foreach ($row in $rows) {
                    $count++
                    if (($count % 2000) -eq 0) {
                        Emit-ScanProgress -Phase "folder_scan" -FolderPath $folderPath -CurrentItemCount $count
                    }
                    $d = $row["ReceivedTime"]
                    if (-not $d) { $d = $row["CreationTime"] }
                    if (-not $d) {
                        try { $d = $row["SentOn"] } catch {}
                    }
                    if (-not $d) {
                        try { $d = $row["LastModificationTime"] } catch {}
                    }
                    if ($d) {
                        $y = [int]$d.Year
                        if ($yearCounts.ContainsKey($y)) {
                            $yearCounts[$y]++
                        } else {
                            $yearCounts[$y] = 1
                        }
                    }
                }
            }
            $usedTable = $true
        } catch {
            try {
                $count = [int]$folder.Items.Count
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
                    if ($d) {
                        $y = [int]$d.Year
                        if ($yearCounts.ContainsKey($y)) {
                            $yearCounts[$y]++
                        } else {
                            $yearCounts[$y] = 1
                        }
                    }
                }
            } catch {}
        }

        $yearBreakdown = @()
        foreach ($y in ($yearCounts.Keys | Sort-Object -Descending)) {
            $yearBreakdown += [pscustomobject]@{
                year = [int]$y
                count = [int]$yearCounts[$y]
            }
        }

        $datedCount = 0
        foreach ($k in $yearCounts.Keys) {
            $datedCount += [int]$yearCounts[$k]
        }
        $undatedCount = [int]([Math]::Max(0, $count - $datedCount))

        $out.Value += [pscustomobject]@{
            type = "folder"
            path = $folderPath
            itemCount = $count
            yearBreakdown = @($yearBreakdown)
            undatedCount = $undatedCount
        }

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
        Emit-Log "warn" "Error recorriendo$pathLabel: $($_.Exception.Message)"
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

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null

if ($FilterOnlyYear -and ($FilterOnlyYear -lt 1900 -or $FilterOnlyYear -gt 9999)) {
    Emit-ErrorPayload "FilterOnlyYear inválido: $FilterOnlyYear. Debe estar entre 1900 y 9999."
    Exit-WithCleanup 1
}

function Resolve-MonthTokenToNumber {
    param([string]$Token)

    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }

    $t = $Token.Trim().ToLower()
    $t = $t.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $t.ToCharArray()) {
        $unicode = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($unicode -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $t = $sb.ToString()

    switch ($t) {
        "1" { return 1 }
        "2" { return 2 }
        "3" { return 3 }
        "4" { return 4 }
        "5" { return 5 }
        "6" { return 6 }
        "7" { return 7 }
        "8" { return 8 }
        "9" { return 9 }
        "10" { return 10 }
        "11" { return 11 }
        "12" { return 12 }
        "enero" { return 1 }
        "ene" { return 1 }
        "january" { return 1 }
        "jan" { return 1 }
        "febrero" { return 2 }
        "feb" { return 2 }
        "february" { return 2 }
        "marzo" { return 3 }
        "mar" { return 3 }
        "march" { return 3 }
        "abril" { return 4 }
        "abr" { return 4 }
        "april" { return 4 }
        "mayo" { return 5 }
        "may" { return 5 }
        "junio" { return 6 }
        "jun" { return 6 }
        "june" { return 6 }
        "julio" { return 7 }
        "jul" { return 7 }
        "july" { return 7 }
        "agosto" { return 8 }
        "ago" { return 8 }
        "august" { return 8 }
        "aug" { return 8 }
        "septiembre" { return 9 }
        "setiembre" { return 9 }
        "sep" { return 9 }
        "set" { return 9 }
        "september" { return 9 }
        "octubre" { return 10 }
        "oct" { return 10 }
        "october" { return 10 }
        "noviembre" { return 11 }
        "nov" { return 11 }
        "november" { return 11 }
        "diciembre" { return 12 }
        "dic" { return 12 }
        "december" { return 12 }
        "dec" { return 12 }
        default { return $null }
    }
}

[int[]]$FilterOnlyMonthNumbers = @()
$script:FilterOnlyMonthLookup = $null
if (-not [string]::IsNullOrWhiteSpace($FilterOnlyMonths)) {
    $tokens = $FilterOnlyMonths -split "," | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if (@($tokens).Count -eq 0) {
        Emit-ErrorPayload "FilterOnlyMonths no contiene meses válidos."
        Exit-WithCleanup 1
    }

    $monthLookup = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($tk in $tokens) {
        $m = Resolve-MonthTokenToNumber -Token $tk
        if (-not $m) {
            Emit-ErrorPayload "Mes inválido en FilterOnlyMonths: '$tk'. Use nombres (enero..diciembre) o números (1..12)."
            Exit-WithCleanup 1
        }
        [void]$monthLookup.Add([int]$m)
    }

    if ($monthLookup.Count -eq 0) {
        Emit-ErrorPayload "FilterOnlyMonths no contiene meses válidos."
        Exit-WithCleanup 1
    }

    $FilterOnlyMonthNumbers = @($monthLookup.ToArray() | Sort-Object)
    $script:FilterOnlyMonthLookup = $monthLookup
}

$SelectedFolderFilters = @()
if ($IncludeFoldersJson) {
    try {
        $decodedFolders = ConvertFrom-Json -InputObject $IncludeFoldersJson -ErrorAction Stop
    } catch {
        Emit-ErrorPayload "No se pudo interpretar -IncludeFoldersJson: $($_.Exception.Message)"
        Exit-WithCleanup 1
    }

    foreach ($f in @($decodedFolders)) {
        $nf = Normalize-FolderPath "$f"
        if (-not [string]::IsNullOrWhiteSpace($nf)) {
            $SelectedFolderFilters += $nf
        }
    }
}

if ($IncludeFolders -and @($IncludeFolders).Count -gt 0) {
    foreach ($f in $IncludeFolders) {
        $nf = Normalize-FolderPath $f
        if (-not [string]::IsNullOrWhiteSpace($nf)) {
            $SelectedFolderFilters += $nf
        }
    }
}

if ($SelectedFolderFilters.Count -gt 0) {
    $SelectedFolderFilters = @($SelectedFolderFilters | Sort-Object -Unique)
}

$script:TokenBucket = @{
    capacity    = [double]$BurstSize
    tokens      = [double]$BurstSize
    refillRate  = [double]($ItemsPerMinute / 60.0)
    lastRefill  = [DateTime]::UtcNow
    totalWaitedMs = [long]0
    throttleErrors = [int]0
    adaptiveMultiplier = [double]1.0
    lastStatsEmit = [DateTime]::UtcNow
}

$script:DupIndexCache = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[string]]'
$script:RuntimeDupKeys = New-Object 'System.Collections.Generic.HashSet[string]'
$script:ChildFolderCache = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.Generic.Dictionary[string, object]]'

function Wait-ForToken {
    while ($true) {
        $now = [DateTime]::UtcNow
        $elapsed = ($now - $script:TokenBucket.lastRefill).TotalSeconds
        $effectiveRate = $script:TokenBucket.refillRate * $script:TokenBucket.adaptiveMultiplier
        $script:TokenBucket.tokens = [Math]::Min(
            $script:TokenBucket.capacity,
            $script:TokenBucket.tokens + ($elapsed * $effectiveRate)
        )
        $script:TokenBucket.lastRefill = $now

        if ($script:TokenBucket.tokens -ge 1.0) {
            $script:TokenBucket.tokens -= 1.0
            return
        }

        $needed = 1.0 - $script:TokenBucket.tokens
        $waitMs = [int](($needed / [Math]::Max($effectiveRate, 0.001)) * 1000)
        $waitMs = [Math]::Max(50, [Math]::Min($waitMs, 5000))
        $script:TokenBucket.totalWaitedMs += $waitMs
        Start-Sleep -Milliseconds $waitMs
    }
}

function Emit-ThrottleStats {
    param([switch]$Force)
    $now = [DateTime]::UtcNow
    if (-not $Force) {
        if (($now - $script:TokenBucket.lastStatsEmit).TotalSeconds -lt 10) { return }
    }
    $script:TokenBucket.lastStatsEmit = $now
    $effRate = $script:TokenBucket.refillRate * $script:TokenBucket.adaptiveMultiplier * 60.0
    $payload = @{
        type = "throttleStats"
        effectiveRate = [double]([Math]::Round($effRate, 2))
        currentTokens = [double]([Math]::Round($script:TokenBucket.tokens, 2))
        burstCapacity = [int]$script:TokenBucket.capacity
        totalWaitedMs = [long]$script:TokenBucket.totalWaitedMs
        throttleErrors = [int]$script:TokenBucket.throttleErrors
        adaptiveMultiplier = [double]([Math]::Round($script:TokenBucket.adaptiveMultiplier, 3))
    }
    if ($Json -or $script:IsHeadlessOutput) {
        [Console]::WriteLine(($payload | ConvertTo-Json -Compress -Depth 6))
    } else {
        Write-Host ("[throttle] rate={0} tokens={1} waited={2}ms errs={3}" -f $payload.effectiveRate, $payload.currentTokens, $payload.totalWaitedMs, $payload.throttleErrors)
    }
}

function Is-ThrottlingError {
    param($errorRecord)
    $msg = "$errorRecord"
    # Hex HRESULT codes in error text
    if ($msg -match "0x80040115" -or $msg -match "0x8004011D" -or $msg -match "0x80040600") { return $true }
    # English patterns
    if ($msg -match "Server Busy" -or $msg -match "throttl" -or $msg -match "budget" -or $msg -match "too many requests" -or $msg -match "429") { return $true }
    # Spanish MAPI network/connection errors (locale-dependent messages from Outlook COM)
    if ($msg -match "problemas en la red" -or $msg -match "conexi.n con Microsoft Exchange" -or $msg -match "no se puede completar") { return $true }
    # Generic network/disconnection patterns
    if ($msg -match "network" -or $msg -match "disconnected" -or $msg -match "connection.*lost" -or $msg -match "RPC_E_DISCONNECTED") { return $true }
    # Check numeric HResult on the exception object
    try {
        $hr = $errorRecord.Exception.HResult
        if ($hr -eq [int]0x80040115 -or $hr -eq [int]0x8004011D -or $hr -eq [int]0x80040600 -or $hr -eq [int]0x800401FD) { return $true }
    } catch {}
    return $false
}

function Invoke-WithRetry {
    param(
        [ScriptBlock]$Operation,
        [string]$OperationName = "operation"
    )
    $attempt = 0
    while ($true) {
        try {
            return & $Operation
        } catch {
            if (-not (Is-ThrottlingError $_) -or $attempt -ge $MaxRetries) {
                throw
            }
            $script:TokenBucket.throttleErrors++
            $delay = [Math]::Min(
                $InitialBackoffMs * [Math]::Pow(2, $attempt),
                $MaxBackoffMs
            )
            $delay = [int]$delay
            # Network/MAPI disconnection errors need longer cooldown for reconnection
            $errMsg = "$($_.Exception.Message)"
            $isNetworkDrop = ($errMsg -match "problemas en la red" -or $errMsg -match "network" -or $errMsg -match "conexi.n con Microsoft Exchange" -or $errMsg -match "disconnected")
            if ($isNetworkDrop) {
                $delay = [Math]::Max($delay, 5000)
                $delay = [Math]::Min([int]($delay * 1.5), 60000)
                Emit-Log "warn" "Red/MAPI desconectado en $OperationName, esperando ${delay}ms para reconexión (intento $($attempt+1)/$MaxRetries): $errMsg"
            } else {
                Emit-Log "warn" "Throttling en $OperationName, backoff ${delay}ms (intento $($attempt+1)/$MaxRetries): $errMsg"
            }
            if ($AdaptiveThrottling) {
                $factor = if ($isNetworkDrop) { 0.5 } else { 0.7 }
                $script:TokenBucket.adaptiveMultiplier = [Math]::Max(0.1, $script:TokenBucket.adaptiveMultiplier * $factor)
                $newRate = [int]($script:TokenBucket.refillRate * $script:TokenBucket.adaptiveMultiplier * 60.0)
                Emit-Log "info" "Adaptive throttling: nueva tasa efectiva ~${newRate} items/min"
            }
            Start-Sleep -Milliseconds $delay
            $attempt++
        }
    }
}

function Get-OutlookNamespace {
    if ($script:MainNamespace -and $script:OutlookApplication) { return $script:MainNamespace }
    try {
        $outlook = $null
        $namespace = $null

        try {
            $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            $namespace = $outlook.GetNamespace("MAPI")
            try { $namespace.Logon("", "", $false, $false) } catch {}
        } catch {
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")
            try { $namespace.Logon("", "", $false, $true) } catch {}
        }

        $script:OutlookApplication = $outlook
        $script:MainNamespace = $namespace
        return $script:MainNamespace
    } catch {
        if ($namespace) { Release-ComObjectSafe $namespace }
        if ($outlook) { Release-ComObjectSafe $outlook }
        Emit-ErrorPayload "No se pudo iniciar Outlook: $($_.Exception.Message)"
        Exit-WithCleanup 1
    }
}

function Get-StoreByIdOrPath {
    param($namespace, [string]$StoreId, [string]$FilePath)
    foreach ($s in $namespace.Stores) {
        if ($StoreId -and ($s.StoreID -eq $StoreId)) { return $s }
        if ($FilePath) {
            try { if ($s.FilePath -and ($s.FilePath -eq $FilePath)) { return $s } } catch {}
        }
    }
    return $null
}

function Get-SubFolders-Safe {
    param($parentFolder)
    try { return $parentFolder.Folders } catch { return @() }
}

function Normalize-FolderName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }

    $v = $Name
    try { $v = $v.Normalize([Text.NormalizationForm]::FormKC) } catch {}
    # Remove zero-width chars, BOM, soft-hyphen, non-breaking spaces
    $v = [regex]::Replace($v, "[\u00AD\u200B-\u200F\u2028-\u202F\uFEFF\u00A0]", "")
    $v = [regex]::Replace($v, "\s+", " ")
    $v = $v.Trim().ToLowerInvariant()
    return $v
}

function Sanitize-FolderName {
    param([string]$Name)

    if ($null -eq $Name) { return "" }
    $v = [string]$Name
    $v = [regex]::Replace($v, "[\u00AD\u200B-\u200F\u2028-\u202F\uFEFF\u00A0]", "")
    $v = [regex]::Replace($v, "\s+", " ").Trim()
    if ([string]::IsNullOrWhiteSpace($v)) { return "(Sin nombre)" }
    return $v
}

function Get-ParentCacheKey {
    param($folder)
    try {
        $eid = $folder.EntryID
        if ($eid) { return $eid }
    } catch {}
    try {
        $sid = $folder.StoreID
        $path = $folder.FolderPath
        if ($sid -and $path) { return "$sid|$path" }
    } catch {}
    return $null
}

function Index-ChildFolders {
    param($parent)

    $cacheKey = Get-ParentCacheKey -folder $parent
    if ($cacheKey -and $script:ChildFolderCache.ContainsKey($cacheKey)) {
        return $script:ChildFolderCache[$cacheKey]
    }

    $index = New-Object 'System.Collections.Generic.Dictionary[string, object]'
    foreach ($f in (Get-SubFolders-Safe -parentFolder $parent)) {
        $childName = $null
        try { $childName = [string]$f.Name } catch { continue }
        $norm = Normalize-FolderName $childName
        if (-not [string]::IsNullOrWhiteSpace($norm) -and -not $index.ContainsKey($norm)) {
            $index[$norm] = $f
        }
    }

    if ($cacheKey) {
        $script:ChildFolderCache[$cacheKey] = $index
    }
    return $index
}

function Ensure-ChildFolder {
    param($parent, [string]$Name)

    $safeName = Sanitize-FolderName $Name
    $targetNorm = Normalize-FolderName $safeName

    if ([string]::IsNullOrWhiteSpace($targetNorm)) {
        $safeName = "(Sin nombre)"
        $targetNorm = Normalize-FolderName $safeName
    }

    $index = Index-ChildFolders -parent $parent

    if ($index.ContainsKey($targetNorm)) {
        Emit-Log "info" "Reutilizando carpeta existente: $safeName"
        return $index[$targetNorm]
    }

    # Double-check: re-enumerate live COM collection as final fallback
    # (handles rare case where cache was built before a sibling add)
    foreach ($f in (Get-SubFolders-Safe -parentFolder $parent)) {
        $childName = $null
        try { $childName = [string]$f.Name } catch { continue }
        $norm = Normalize-FolderName $childName
        if ($norm -eq $targetNorm) {
            $index[$norm] = $f
            Emit-Log "info" "Reutilizando carpeta existente (re-scan): $safeName"
            return $f
        }
    }

    Emit-Log "info" "Creando carpeta nueva: $safeName"
    $created = $parent.Folders.Add($safeName)
    $index[$targetNorm] = $created

    # Invalidate parent cache to force re-index on next sibling query
    $cacheKey = Get-ParentCacheKey -folder $parent
    if ($cacheKey) {
        $script:ChildFolderCache[$cacheKey] = $index
    }

    return $created
}

function Resolve-TargetTopFolder {
    param(
        $targetStore,
        $targetRoot,
        [string]$SourceTopName
    )

    $n = Normalize-FolderName $SourceTopName

    # OlDefaultFolders (valores COM)
    # 3=DeletedItems, 4=Outbox, 5=SentMail, 6=Inbox, 16=Drafts, 23=Junk
    if ($n -in @("bandeja de entrada", "inbox")) {
        try { return $targetStore.GetDefaultFolder(6) } catch {}
    } elseif ($n -in @("elementos eliminados", "deleted items")) {
        try { return $targetStore.GetDefaultFolder(3) } catch {}
    } elseif ($n -in @("elementos enviados", "sent items")) {
        try { return $targetStore.GetDefaultFolder(5) } catch {}
    } elseif ($n -in @("borradores", "drafts")) {
        try { return $targetStore.GetDefaultFolder(16) } catch {}
    } elseif ($n -in @("correo no deseado", "junk email")) {
        try { return $targetStore.GetDefaultFolder(23) } catch {}
    } elseif ($n -in @("bandeja de salida", "outbox")) {
        try { return $targetStore.GetDefaultFolder(4) } catch {}
    }

    return Ensure-ChildFolder -parent $targetRoot -Name $SourceTopName
}

function Analyze-PstFolderRecursive {
    param($folder, [string]$pathPrefix, [ref]$totalItems, [ref]$totalBytes, [ref]$folderList)
    try {
        $folderPath = if ($pathPrefix) { "$pathPrefix\$($folder.Name)" } else { $folder.Name }
        $itemCount = 0
        $sizeBytes = [long]0

        try {
            $items = $folder.Items
            $itemCount = [int]$items.Count
            try {
                $table = $folder.GetTable("")
                $table.Columns.RemoveAll()
                $table.Columns.Add("Size") | Out-Null
                while (-not $table.EndOfTable) {
                    $row = $table.GetNextRow()
                    try { $sizeBytes += [long]$row["Size"] } catch {}
                }
            } catch {}
        } catch {}

        $totalItems.Value += $itemCount
        $totalBytes.Value += $sizeBytes
        $folderList.Value += [pscustomobject]@{
            path = $folderPath
            itemCount = $itemCount
            sizeBytes = $sizeBytes
            sizeHuman = (Format-Bytes -bytes $sizeBytes)
        }

        foreach ($sub in (Get-SubFolders-Safe -parentFolder $folder)) {
            Analyze-PstFolderRecursive -folder $sub -pathPrefix $folderPath -totalItems $totalItems -totalBytes $totalBytes -folderList $folderList
        }
    } catch {
        Emit-Log "warn" "Error analizando carpeta '$($folder.Name)': $($_.Exception.Message)"
    }
}

function Normalize-MessageId {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $v = $Value.Trim().ToLowerInvariant()
    if ($v.StartsWith("<") -and $v.EndsWith(">") -and $v.Length -gt 2) {
        $v = $v.Substring(1, $v.Length - 2)
    }
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    return $v
}

function Normalize-Subject {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $v = $Value.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    $v = [regex]::Replace($v, "\s+", " ")
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    return $v
}

function Convert-BytesToHex {
    param($Bytes)
    if ($null -eq $Bytes) { return $null }

    $normalized = $null
    try {
        if ($Bytes -is [byte[]]) {
            $normalized = $Bytes
        } elseif ($Bytes -is [System.Array]) {
            $tmp = New-Object 'System.Collections.Generic.List[byte]'
            foreach ($b in $Bytes) {
                try { $tmp.Add([byte]$b) } catch {}
            }
            if ($tmp.Count -eq 0) { return $null }
            $normalized = $tmp.ToArray()
        } else {
            try { $normalized = [byte[]]$Bytes } catch { return $null }
        }
    } catch {
        return $null
    }

    if (-not $normalized -or $normalized.Length -eq 0) { return $null }
    return ([System.BitConverter]::ToString($normalized)).Replace("-", "").ToLowerInvariant()
}

function Get-DuplicateKeyFromRow {
    param($row)

    $msgId = $null
    try { $msgId = $row["http://schemas.microsoft.com/mapi/proptag/0x1035001F"] } catch {}
    if (-not $msgId) {
        try { $msgId = $row["urn:schemas:mailheader:message-id"] } catch {}
    }
    $normalizedMsgId = Normalize-MessageId $msgId
    if ($normalizedMsgId) { return "mid:$normalizedMsgId" }

    $searchKey = $null
    try { $searchKey = $row["http://schemas.microsoft.com/mapi/proptag/0x300B0102"] } catch {}
    if ($searchKey) {
        $hex = Convert-BytesToHex -Bytes $searchKey
        if ($hex) { return "sk:$hex" }
    }

    $subject = $null
    try { $subject = $row["Subject"] } catch {}
    $normalizedSubject = Normalize-Subject $subject
    $dateStr = "nodate"
    try { $d = $row["ReceivedTime"]; if ($d) { $dateStr = ([datetime]$d).ToString("yyyyMMddHHmmss") } } catch {}
    if ($dateStr -eq "nodate") {
        try { $d = $row["SentOn"]; if ($d) { $dateStr = ([datetime]$d).ToString("yyyyMMddHHmmss") } } catch {}
    }
    if ($dateStr -eq "nodate") {
        try { $d = $row["CreationTime"]; if ($d) { $dateStr = ([datetime]$d).ToString("yyyyMMddHHmmss") } } catch {}
    }
    $sender = "nosender"
    try { $s = $row["http://schemas.microsoft.com/mapi/proptag/0x0039001F"]; if ($s) { $sender = ([string]$s).Trim().ToLowerInvariant() } } catch {}

    if ($dateStr -ne "nodate") {
        $subjectPart = if ($normalizedSubject) { $normalizedSubject } else { "(nosubject)" }
        return "comp:$subjectPart|$dateStr|$sender"
    }

    return $null
}

function Get-DuplicateKeyFromItem {
    param($item)

    $msgId = $null
    try { $msgId = $item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F") } catch {}
    if (-not $msgId) {
        try { $msgId = $item.PropertyAccessor.GetProperty("urn:schemas:mailheader:message-id") } catch {}
    }
    $normalizedMsgId = Normalize-MessageId $msgId
    if ($normalizedMsgId) { return "mid:$normalizedMsgId" }

    $searchKey = $null
    try { $searchKey = $item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x300B0102") } catch {}
    if ($searchKey) {
        $hex = Convert-BytesToHex -Bytes $searchKey
        if ($hex) { return "sk:$hex" }
    }

    $subject = $null
    try { $subject = [string]$item.Subject } catch {}
    $normalizedSubject = Normalize-Subject $subject
    $dateStr = "nodate"
    try { $d = $item.ReceivedTime; if ($d) { $dateStr = ([datetime]$d).ToString("yyyyMMddHHmmss") } } catch {}
    if ($dateStr -eq "nodate") {
        try { $d = $item.SentOn; if ($d) { $dateStr = ([datetime]$d).ToString("yyyyMMddHHmmss") } } catch {}
    }
    if ($dateStr -eq "nodate") {
        try { $d = $item.CreationTime; if ($d) { $dateStr = ([datetime]$d).ToString("yyyyMMddHHmmss") } } catch {}
    }
    $sender = "nosender"
    try { $s = $item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0039001F"); if ($s) { $sender = ([string]$s).Trim().ToLowerInvariant() } } catch {}

    if ($dateStr -ne "nodate") {
        $subjectPart = if ($normalizedSubject) { $normalizedSubject } else { "(nosubject)" }
        return "comp:$subjectPart|$dateStr|$sender"
    }

    return $null
}

function Build-DuplicateIndexFromFolder {
    param($folder, [System.Collections.Generic.List[string]]$list)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $indexInactivityTimeout = [int]$DuplicateIndexInactivityTimeoutSec
    $inactivityTimeoutEnabled = ($indexInactivityTimeout -gt 0)
    $lastProgressAt = [DateTime]::UtcNow

    if ($inactivityTimeoutEnabled) {
        Emit-Log "info" "  Timeout por inactividad: ${indexInactivityTimeout}s"
    } else {
        Emit-Log "info" "  Timeout por inactividad deshabilitado"
    }

    # Strategy 1: Fast path using GetTable (batch reads)
    try {
        $tableFilter = ""
        if ($FilterOnlyYear) {
            $fromDate = Get-Date -Year $FilterOnlyYear -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
            $toDate = $fromDate.AddYears(1)
            $fromLiteral = $fromDate.ToString("yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
            $toLiteral = $toDate.ToString("yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
            $tableFilter = "@SQL=\"\"urn:schemas:httpmail:datereceived\"\" >= '$fromLiteral' AND \"\"urn:schemas:httpmail:datereceived\"\" < '$toLiteral'"
            Emit-Log "info" "  GetTable con filtro anio=$FilterOnlyYear..."
        } else {
            Emit-Log "info" "  GetTable iniciando..."
        }
        $table = $folder.GetTable($tableFilter)
        Emit-Log "info" "  GetTable listo, configurando columnas..."
        $table.Columns.RemoveAll()
        try { $table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x1035001F") | Out-Null } catch {}
        try { $table.Columns.Add("urn:schemas:mailheader:message-id") | Out-Null } catch {}
        try { $table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x300B0102") | Out-Null } catch {}
        try { $table.Columns.Add("Subject") | Out-Null } catch {}
        try { $table.Columns.Add("ReceivedTime") | Out-Null } catch {}
        try { $table.Columns.Add("SentOn") | Out-Null } catch {}
        try { $table.Columns.Add("CreationTime") | Out-Null } catch {}
        try { $table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0039001F") | Out-Null } catch {}

        Emit-Log "info" "  Leyendo filas..."
        $rowCount = 0
        while (-not $table.EndOfTable) {
            if ($inactivityTimeoutEnabled) {
                $idleSec = ([DateTime]::UtcNow - $lastProgressAt).TotalSeconds
                if ($idleSec -ge $indexInactivityTimeout) {
                    Emit-Log "warn" "  Timeout por inactividad en indexacion (${indexInactivityTimeout}s), usando indice parcial ($($list.Count) claves)"
                    return
                }
            }

            $row = $table.GetNextRow()
            $rowCount++
            $lastProgressAt = [DateTime]::UtcNow

            if ($rowCount % 500 -eq 0) {
                Emit-Log "info" "  indexado $($list.Count) claves ($($sw.Elapsed.TotalSeconds.ToString('F0'))s)..."
            }

            try {
                $k = Get-DuplicateKeyFromRow -row $row
                if ($k) { $list.Add([string]$k) }
            } catch {}
        }
        Emit-Log "info" "  Indexacion completa: $($list.Count) claves en $($sw.Elapsed.TotalSeconds.ToString('F1'))s"
        return
    } catch {
        Emit-Log "warn" "GetTable fallo ($($sw.Elapsed.TotalSeconds.ToString('F0'))s): $($_.Exception.Message)"
    }

    # Strategy 2: Fallback item-by-item (slow but reliable)
    try {
        $itemCount = 0
        try { $itemCount = [int]$folder.Items.Count } catch {}
        Emit-Log "warn" "Fallback: iterando $itemCount items individualmente..."
        $idx = 0
        $lastProgressAt = [DateTime]::UtcNow
        foreach ($it in $folder.Items) {
            if ($inactivityTimeoutEnabled) {
                $idleSec = ([DateTime]::UtcNow - $lastProgressAt).TotalSeconds
                if ($idleSec -ge $indexInactivityTimeout) {
                    Emit-Log "warn" "  Timeout por inactividad en fallback (${indexInactivityTimeout}s, $idx/$itemCount). Indice parcial: $($list.Count) claves"
                    return
                }
            }
            $k = $null
            try { $k = Get-DuplicateKeyFromItem -item $it } catch {}
            if ($k) { $list.Add([string]$k) }
            $idx++
            $lastProgressAt = [DateTime]::UtcNow
            if ($idx % 100 -eq 0) {
                Emit-Log "info" "  fallback: $idx/$itemCount ($($sw.Elapsed.TotalSeconds.ToString('F0'))s)..."
            }
        }
    } catch {}
}

function Build-DuplicateIndexRecursive {
    param($folder, [System.Collections.Generic.List[string]]$list)
    foreach ($sub in (Get-SubFolders-Safe -parentFolder $folder)) {
        $null = Build-DuplicateIndexFromFolder -folder $sub -list $list
        $null = Build-DuplicateIndexRecursive -folder $sub -list $list
    }
}

function Build-DuplicateIndex {
    param($targetFolder, [switch]$Deep)

    $folderId = $null
    try { $folderId = $targetFolder.EntryID } catch {}
    if ($folderId -and $script:DupIndexCache.ContainsKey($folderId)) {
        return ,$script:DupIndexCache[$folderId]
    }

    $list = New-Object 'System.Collections.Generic.List[string]'
    $null = Build-DuplicateIndexFromFolder -folder $targetFolder -list $list

    if ($Deep) {
        $null = Build-DuplicateIndexRecursive -folder $targetFolder -list $list
    }

    $readOnlySet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($key in $list) {
        [void]$readOnlySet.Add($key)
    }

    if ($folderId) {
        $script:DupIndexCache[$folderId] = $readOnlySet
    }

    return ,$readOnlySet
}

function Restore-FolderRecursive {
    param(
        $sourceFolder,
        $destFolder,
        [string]$pathPrefix,
        [ref]$stats
    )
    $folderPath = if ($pathPrefix) { "$pathPrefix\$($sourceFolder.Name)" } else { $sourceFolder.Name }
    $activity = if ($Action -ieq "Move") { "Moviendo" } else { "Copiando" }

    $processCurrent = Should-ProcessFolder -FolderPath $folderPath -SelectedFolders $SelectedFolderFilters
    $hasSelectedBelow = Has-SelectedDescendant -FolderPath $folderPath -SelectedFolders $SelectedFolderFilters
    if (-not $processCurrent -and -not $hasSelectedBelow) {
        return
    }

    $existingKeys = $null
    if ($SkipDuplicates) {
        Emit-Log "info" "Indexando duplicados en: $folderPath"
        $existingKeys = Build-DuplicateIndex -targetFolder $destFolder -Deep:$DeepDuplicateCheck
        Emit-Log "info" "Indice inicial: $($existingKeys.Count) claves existentes + $($script:RuntimeDupKeys.Count) runtime$(if ($DeepDuplicateCheck) { ' (deep)' })"
    }

    $items = $null
    try { $items = $sourceFolder.Items } catch {
        Emit-Log "error" "No se pudo acceder a Items en $folderPath : $($_.Exception.Message)"
        return
    }
    $itemCount = 0
    try { $itemCount = [int]$items.Count } catch {}

    Emit-Log "info" "$activity $itemCount ítems de: $folderPath"

    $hasMonthFilter = ($script:FilterOnlyMonthLookup -and $script:FilterOnlyMonthLookup.Count -gt 0)

    if ($processCurrent -and $itemCount -gt 0) {
        for ($i = $itemCount; $i -ge 1; $i--) {
            $item = $null
            try { $item = $items.Item($i) } catch { continue }
            if (-not $item) { continue }

            $itemDate = $null

            if ($FilterOnlyYear) {
                $itemDate = $null
                try { $itemDate = $item.ReceivedTime } catch {}
                if (-not $itemDate) {
                    try { $itemDate = $item.CreationTime } catch {}
                }
                if (-not $itemDate -or $itemDate.Year -ne $FilterOnlyYear) {
                    $stats.Value.skipped++
                    continue
                }
            }

            if ($hasMonthFilter) {
                if (-not $itemDate) {
                    try { $itemDate = $item.ReceivedTime } catch {}
                    if (-not $itemDate) {
                        try { $itemDate = $item.CreationTime } catch {}
                    }
                }
                $monthValue = $null
                if ($itemDate) {
                    try { $monthValue = [int]$itemDate.Month } catch { $monthValue = $null }
                }
                $monthMatches = $false
                if ($monthValue -ne $null) {
                    try { $monthMatches = $script:FilterOnlyMonthLookup.Contains($monthValue) } catch { $monthMatches = $false }
                }
                if (-not $itemDate -or -not $monthMatches) {
                    $stats.Value.skipped++
                    continue
                }
            }

            $dupKey = $null
            if ($existingKeys) {
                $dupKey = Get-DuplicateKeyFromItem -item $item
                if ($dupKey) {
                    $isDup = $false
                    $dupSource = $null
                    if ($existingKeys.Contains([string]$dupKey)) {
                        $isDup = $true
                        $dupSource = "existing"
                    } elseif ($script:RuntimeDupKeys.Contains([string]$dupKey)) {
                        $isDup = $true
                        $dupSource = "batch"
                    }
                    if ($isDup) {
                        $stats.Value.skipped++
                        $dupPayload = @{
                            type = "dupSkipped"
                            folder = $folderPath
                            subject = (Get-SafeSubject $item)
                            key = [string]$dupKey
                            source = $dupSource
                        }
                        if ($Json) { [Console]::WriteLine(($dupPayload | ConvertTo-Json -Compress -Depth 6)) }
                        else { Emit-Log "info" "Duplicado saltado [$dupSource]: $(Get-SafeSubject $item)" }
                        continue
                    }
                }
            }

            $itemSize = 0
            try { $itemSize = [long]$item.Size } catch {}
            if ($MaxItemSizeBytes -gt 0 -and $itemSize -gt $MaxItemSizeBytes) {
                $stats.Value.failed++
                Add-FailureRecord -stats $stats -Record @{
                    folder = $folderPath
                    subject = (Get-SafeSubject $item)
                    reason = "too_large"
                    sizeBytes = $itemSize
                }
                Emit-Log "warn" "Ítem > $(Format-Bytes -bytes $MaxItemSizeBytes) ignorado: $(Get-SafeSubject $item)"
                continue
            }

            Wait-ForToken

            try {
                Invoke-WithRetry -OperationName "$Action ítem" -Operation {
                    if ($Action -ieq "Move") {
                        [void]$item.Move($destFolder)
                    } else {
                        $copied = $item.Copy()
                        if ($copied) {
                            [void]$copied.Move($destFolder)
                        } else {
                            throw "No se pudo copiar el ítem."
                        }
                    }
                }
                if ($Action -ieq "Move") { $stats.Value.moved++ } else { $stats.Value.copied++ }
                if ($dupKey) {
                    [void]$script:RuntimeDupKeys.Add([string]$dupKey)
                    # Update folder cache so subsequent calls see this key
                    if ($existingKeys) {
                        [void]$existingKeys.Add([string]$dupKey)
                    }
                }
            } catch {
                $stats.Value.failed++
                $reason = if (Is-ThrottlingError $_) { "throttled_max_retries" } else { "error" }
                Add-FailureRecord -stats $stats -Record @{
                    folder = $folderPath
                    subject = (Get-SafeSubject $item)
                    reason = $reason
                    message = "$($_.Exception.Message)"
                }
                Emit-Log "error" "Falló ítem en $folderPath : $($_.Exception.Message)"
            }

            $stats.Value.processed++
            $pct = 0
            if ($stats.Value.total -gt 0) {
                $pct = [int][Math]::Round(($stats.Value.processed / $stats.Value.total) * 100)
                if ($pct -gt 99 -and $stats.Value.processed -lt $stats.Value.total) { $pct = 99 }
                if ($pct -eq 0 -and $stats.Value.processed -gt 0) { $pct = 1 }
            }
            Publish-Progress -Activity "Restauracion PST -> Buzon" -Status "$folderPath ($($stats.Value.processed)/$($stats.Value.total))" -PercentComplete $pct -Copied $stats.Value.copied -Moved $stats.Value.moved -Skipped $stats.Value.skipped -Failed $stats.Value.failed

            Emit-ThrottleStats
        }
    }

    # Pre-index destination children BEFORE iterating source subfolders
    # This ensures all existing dest subfolders are in cache before any .Add() calls
    [void](Index-ChildFolders -parent $destFolder)

    # Snapshot source subfolders into array to avoid COM enumerator invalidation
    $sourceSubFolders = @()
    foreach ($sub in (Get-SubFolders-Safe -parentFolder $sourceFolder)) {
        $sourceSubFolders += $sub
    }

    foreach ($sub in $sourceSubFolders) {
        $subName = $null
        try { $subName = [string]$sub.Name } catch { continue }
        $destSub = Ensure-ChildFolder -parent $destFolder -Name $subName
        Restore-FolderRecursive -sourceFolder $sub -destFolder $destSub -pathPrefix $folderPath -stats $stats
    }
}

if (-not $PstPath) { Emit-ErrorPayload "Se requiere -PstPath."; Exit-WithCleanup 1 }
if (-not (Test-Path $PstPath)) { Emit-ErrorPayload "PST no encontrado: $PstPath"; Exit-WithCleanup 1 }

$namespace = Get-OutlookNamespace

$alreadyMounted = $false
$pstStore = Get-StoreByIdOrPath -namespace $namespace -FilePath $PstPath
if ($pstStore) {
    $alreadyMounted = $true
    Emit-Log "info" "PST ya estaba montado."
} else {
    Emit-Log "info" "Montando PST: $PstPath"
    try { $namespace.AddStoreEx($PstPath, 3) } catch {
        Emit-ErrorPayload "No se pudo montar el PST: $($_.Exception.Message)"
        Exit-WithCleanup 1
    }
    $pstStore = Get-StoreByIdOrPath -namespace $namespace -FilePath $PstPath
    if (-not $pstStore) {
        Emit-ErrorPayload "PST montado pero no localizado."
        Exit-WithCleanup 1
    }
}

$script:PstStoreRef = $pstStore

$pstRoot = $pstStore.GetRootFolder()
$script:PstRootRef = $pstRoot

if ($ListFolders) {
    $pstSizeBytes = 0
    try {
        $pstSizeBytes = [long](Get-Item -LiteralPath $PstPath -ErrorAction Stop).Length
    } catch {}

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

    if ($Json) {
        Write-Output (@{
            type = "scanMeta"
            pstPath = $PstPath
            pstSizeBytes = [long]$pstSizeBytes
            totalFolders = [int]$totalFoldersToScan
        } | ConvertTo-Json -Compress -Depth 6)
    } else {
        Emit-Log "info" "Escaneando PST... Carpetas estimadas: $totalFoldersToScan"
    }

    $flat = [ref]@()
    foreach ($tf in (Get-SubFolders-Safe -parentFolder $pstRoot)) {
        Collect-PstFoldersRecursive -folder $tf -pathPrefix "" -out $flat
    }

    Emit-ScanProgress -Phase "completed" -FolderPath "" -CurrentItemCount 0 -FolderCompleted

    [Console]::WriteLine((@{ type = "folders"; count = @($flat.Value).Count } | ConvertTo-Json -Compress -Depth 6))
    foreach ($f in $flat.Value) {
        [Console]::WriteLine(($f | ConvertTo-Json -Compress -Depth 6))
    }

    if (-not $alreadyMounted) {
        try { $namespace.RemoveStore($pstRoot) } catch {}
    }
    Exit-WithCleanup 0
}

if (-not $TargetStoreId) { Emit-ErrorPayload "Se requiere -TargetStoreId."; Exit-WithCleanup 1 }

$targetStore = Get-StoreByIdOrPath -namespace $namespace -StoreId $TargetStoreId
if (-not $targetStore) {
    Emit-ErrorPayload "No se encontró el buzón destino (StoreId=$TargetStoreId)."
    Exit-WithCleanup 1
}
$script:TargetStoreRef = $targetStore
Emit-Log "info" "Buzón destino: $($targetStore.DisplayName)"

Emit-Log "info" "Contando ítems del PST (para progreso)..."
$totalItems = [ref]0
$totalBytes = [ref]([long]0)
$folderList = [ref]@()
foreach ($tf in (Get-SubFolders-Safe -parentFolder $pstRoot)) {
    Analyze-PstFolderRecursive -folder $tf -pathPrefix "" -totalItems $totalItems -totalBytes $totalBytes -folderList $folderList
}
Emit-Log "info" "Total ítems a procesar: $($totalItems.Value)"

$stats = [ref]@{
    copied = 0
    moved = 0
    skipped = 0
    failed = 0
    processed = 0
    total = [int]$totalItems.Value
    failures = @()
    failureOverflow = 0
}

$startTime = [DateTime]::UtcNow
$targetRoot = $targetStore.GetRootFolder()
$script:TargetRootRef = $targetRoot

# Pre-index all existing top-level folders in target to avoid duplicates
Emit-Log "info" "Pre-indexando carpetas destino..."
[void](Index-ChildFolders -parent $targetRoot)

foreach ($sourceTop in (Get-SubFolders-Safe -parentFolder $pstRoot)) {
    $destTop = Resolve-TargetTopFolder -targetStore $targetStore -targetRoot $targetRoot -SourceTopName $sourceTop.Name
    Restore-FolderRecursive -sourceFolder $sourceTop -destFolder $destTop -pathPrefix "" -stats $stats
}

Publish-Progress -Activity "Restauracion PST -> Buzon" -Status "Completado" -PercentComplete 100 -Completed -Copied $stats.Value.copied -Moved $stats.Value.moved -Skipped $stats.Value.skipped -Failed $stats.Value.failed
Emit-ThrottleStats -Force

if (-not $alreadyMounted) {
    try { $namespace.RemoveStore($pstRoot) } catch {
        Emit-Log "warn" "No se pudo desmontar el PST: $($_.Exception.Message)"
    }
}

$elapsed = [long](([DateTime]::UtcNow - $startTime).TotalMilliseconds)
$payload = @{
    type = "restoreResult"
    filterOnlyYear = if ($FilterOnlyYear) { [int]$FilterOnlyYear } else { $null }
    filterOnlyMonths = if ($FilterOnlyMonthNumbers.Count -gt 0) { @($FilterOnlyMonthNumbers) } else { @() }
    copied = [int]$stats.Value.copied
    moved = [int]$stats.Value.moved
    skipped = [int]$stats.Value.skipped
    failed = [int]$stats.Value.failed
    elapsedMs = $elapsed
    throttleEvents = [int]$script:TokenBucket.throttleErrors
    totalWaitedMs = [long]$script:TokenBucket.totalWaitedMs
    failures = @($stats.Value.failures)
    failureOverflow = [int]$stats.Value.failureOverflow
}
[Console]::WriteLine(($payload | ConvertTo-Json -Compress -Depth 10 -WarningAction SilentlyContinue))
Exit-WithCleanup 0
