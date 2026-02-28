<#
.SYNOPSIS
    Bulk enable/disable Exchange mailbox SOA (Exchange attribute SOA) for mailboxes in Exchange Online.

.DESCRIPTION
    Reads a CSV with Identity + optional Mode (Enable/Disable). For each row, the script:
      1) Reads mailbox state: IsDirSynced + IsExchangeCloudManaged
      2) If eligible, sets SOA using:
            Set-Mailbox -Identity <User> -IsExchangeCloudManaged $true / $false
      3) Logs every action (updates, skips, errors) to transcript + CSV
      4) Shows progress + colored status output

.CSV EXAMPLE
    Save as: .\MailboxSOA-Bulk.csv

    Identity,Mode
    user1@contoso.com,Enable
    user2@contoso.com,Enable
    shared.mailbox@contoso.com,Enable
    room101@contoso.com,Disable
    aliasOnlyUser,Enable

.NOTES
    Author: Peter
    Script Name: Enable-EXOMailboxSOA-Bulk.ps1
    Version: 1.7
    Updated: 2026-02-24
    Requires: ExchangeOnlineManagement module
#>

#region ========================== USER SETTINGS ==========================
$CsvPath = ".\MailboxSOA-Bulk.csv"
$WhatIfMode = $true
$DefaultMode = "Enable"
$AdminUPN = ""
$LogDir    = ".\Logs"
$ExportDir = ".\Exports"
$MaxRetries = 5
$InitialDelaySeconds = 2
$MaxDelaySeconds = 30
#endregion =================================================================

#region ========================== FUNCTIONS ==============================
function Write-Status {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet("INFO","OK","WARN","ERROR","SKIP","CHANGE")] [string] $Level = "INFO"
    )
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    switch ($Level) {
        "OK"     { Write-Host "[$ts] [ OK    ] $Message" -ForegroundColor Green }
        "CHANGE" { Write-Host "[$ts] [ CHANGE] $Message" -ForegroundColor Magenta }
        "WARN"   { Write-Host "[$ts] [ WARN  ] $Message" -ForegroundColor Yellow }
        "ERROR"  { Write-Host "[$ts] [ ERROR ] $Message" -ForegroundColor Red }
        "SKIP"   { Write-Host "[$ts] [ SKIP  ] $Message" -ForegroundColor DarkYellow }
        default  { Write-Host "[$ts] [ INFO  ] $Message" -ForegroundColor Cyan }
    }
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Show-FriendlyCsvMissing {
    param([Parameter(Mandatory)][string]$Path)

    Write-Host ""
    Write-Host "==================== INPUT FILE MISSING ====================" -ForegroundColor Red
    Write-Host "I couldn't find the CSV input file here:" -ForegroundColor Red
    Write-Host "  $Path" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Fix:" -ForegroundColor Cyan
    Write-Host "  1) Create the file (or copy it) to the folder you run the script from" -ForegroundColor Cyan
    Write-Host "  2) Or change `$CsvPath at the top of the script to the correct path" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Expected CSV format (example):" -ForegroundColor Cyan
    Write-Host "  Identity,Mode" -ForegroundColor Gray
    Write-Host "  user1@contoso.com,Enable" -ForegroundColor Gray
    Write-Host "  user2@contoso.com,Enable" -ForegroundColor Gray
    Write-Host "  shared.mailbox@contoso.com,Enable" -ForegroundColor Gray
    Write-Host "  room101@contoso.com,Disable" -ForegroundColor Gray
    Write-Host "  aliasOnlyUser,Enable" -ForegroundColor Gray
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host ""
}

function Assert-CsvExistsAndNotEmpty {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        Show-FriendlyCsvMissing -Path $Path
        exit 1
    }

    $content = Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $content -or $content.Count -lt 2) {
        Write-Host ""
        Write-Host "==================== INPUT FILE EMPTY ======================" -ForegroundColor Red
        Write-Host "The CSV exists but looks empty (or only has headers):" -ForegroundColor Red
        Write-Host "  $Path" -ForegroundColor Yellow
        Write-Host "Please add at least one user row under the header." -ForegroundColor Cyan
        Write-Host "============================================================" -ForegroundColor Red
        Write-Host ""
        exit 1
    }
}

function Normalize-Mode {
    param([string]$Mode)
    if ([string]::IsNullOrWhiteSpace($Mode)) { return $DefaultMode }
    $m = $Mode.Trim().ToLowerInvariant()
    switch ($m) {
        "enable"  { "Enable" }
        "disable" { "Disable" }
        default { throw "Invalid Mode '$Mode'. Allowed: Enable/Disable (or blank)." }
    }
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)] [scriptblock] $ScriptBlock,
        [Parameter(Mandatory)] [string] $OperationName
    )

    $attempt = 0
    $delay = [Math]::Max(1, $InitialDelaySeconds)

    while ($true) {
        try {
            $attempt++
            return & $ScriptBlock
        } catch {
            $msg = $_.Exception.Message
            $isTransient = (
                $msg -match "The server is busy" -or
                $msg -match "temporarily unavailable" -or
                $msg -match "throttl" -or
                $msg -match "Timeout" -or
                $msg -match "503" -or
                $msg -match "429"
            )

            if (-not $isTransient -or $attempt -ge $MaxRetries) {
                throw "Operation '$OperationName' failed after $attempt attempt(s). Last error: $msg"
            }

            Write-Status "Transient error on '$OperationName' (attempt $attempt/$MaxRetries). Retrying in ${delay}s..." "WARN"
            Start-Sleep -Seconds $delay
            $delay = [Math]::Min($delay * 2, $MaxDelaySeconds)
        }
    }
}
#endregion =================================================================

#region ========================== STARTUP ================================
$ErrorActionPreference = "Stop"

# FRIENDLY INPUT CHECK
Assert-CsvExistsAndNotEmpty -Path $CsvPath

Ensure-Folder -Path $LogDir
Ensure-Folder -Path $ExportDir

$runStamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
$TranscriptPath = Join-Path $LogDir    "Enable-EXOMailboxSOA_$runStamp.log.txt"
$ResultsPath    = Join-Path $ExportDir "Enable-EXOMailboxSOA_Results_$runStamp.csv"

Start-Transcript -Path $TranscriptPath -Force | Out-Null

Write-Status "Starting Enable-EXOMailboxSOA-Bulk.ps1 v1.7 (WhatIfMode=$WhatIfMode, DefaultMode=$DefaultMode)" "INFO"
Write-Status "CSV: $CsvPath" "INFO"
Write-Status "Transcript: $TranscriptPath" "INFO"
Write-Status "Results: $ResultsPath" "INFO"

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Status "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement" "ERROR"
    Stop-Transcript | Out-Null
    exit 1
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

try {
    if ([string]::IsNullOrWhiteSpace($AdminUPN)) {
        Write-Status "Connecting to Exchange Online (interactive)..." "INFO"
        Connect-ExchangeOnline -ShowBanner:$false
    } else {
        Write-Status "Connecting to Exchange Online as $AdminUPN ..." "INFO"
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
    }
    Write-Status "Connected to Exchange Online." "OK"
} catch {
    Write-Status "Failed to connect to Exchange Online: $($_.Exception.Message)" "ERROR"
    Stop-Transcript | Out-Null
    exit 1
}

$rows = Import-Csv -LiteralPath $CsvPath -ErrorAction Stop
if (-not ($rows | Get-Member -Name Identity)) {
    Write-Status "CSV missing required column: Identity" "ERROR"
    Stop-Transcript | Out-Null
    exit 1
}
#endregion =================================================================

#region ========================== MAIN LOOP ==============================
$results = New-Object System.Collections.Generic.List[object]

$total = $rows.Count
for ($i = 0; $i -lt $rows.Count; $i++) {
    $idx = $i + 1
    $id  = [string]$rows[$i].Identity
    $modeRaw = [string]$rows[$i].Mode

    $pct = if ($total -gt 0) { [int](($idx / $total) * 100) } else { 0 }
    $idText = if ([string]::IsNullOrWhiteSpace($id)) { "<empty>" } else { $id.Trim() }

    Write-Progress -Activity "Set Exchange Mailbox SOA (IsExchangeCloudManaged)" -Status "Processing ${idx} of ${total}: $idText" -PercentComplete $pct

    if ([string]::IsNullOrWhiteSpace($id)) {
        Write-Status "[${idx}/${total}] Empty Identity - skipped." "SKIP"
        $results.Add([pscustomobject]@{ Timestamp=(Get-Date).ToString("s"); Row=$idx; Identity=""; Mode=""; Result="Skipped"; Reason="Empty Identity"; Before=""; After=""; WhatIf=$WhatIfMode })
        continue
    }

    $id = $id.Trim()

    try { $mode = Normalize-Mode -Mode $modeRaw }
    catch {
        Write-Status "[${idx}/${total}] $id - invalid Mode '$modeRaw' - skipped." "SKIP"
        $results.Add([pscustomobject]@{ Timestamp=(Get-Date).ToString("s"); Row=$idx; Identity=$id; Mode=$modeRaw; Result="Skipped"; Reason="Invalid Mode"; Before=""; After=""; WhatIf=$WhatIfMode })
        continue
    }

    try {
        $mbx = Invoke-WithRetry -OperationName "Get-Mailbox $id" -ScriptBlock {
            Get-Mailbox -Identity $id -ErrorAction Stop
        }

        $isDirSynced = [string]$mbx.IsDirSynced
        $before = [string]$mbx.IsExchangeCloudManaged

        if ($isDirSynced -ne "True") {
            Write-Status "[${idx}/${total}] Skipped $id - not DirSynced (IsDirSynced=$isDirSynced)." "SKIP"
            $results.Add([pscustomobject]@{ Timestamp=(Get-Date).ToString("s"); Row=$idx; Identity=$id; Mode=$mode; Result="Skipped"; Reason="Not DirSynced"; Before=$before; After=$before; WhatIf=$WhatIfMode })
            continue
        }

        $target = if ($mode -eq "Enable") { "True" } else { "False" }

        if ($before -eq $target) {
            Write-Status "[${idx}/${total}] Skipped $id - already IsExchangeCloudManaged=$target" "SKIP"
            $results.Add([pscustomobject]@{ Timestamp=(Get-Date).ToString("s"); Row=$idx; Identity=$id; Mode=$mode; Result="Skipped"; Reason="Already $target"; Before=$before; After=$before; WhatIf=$WhatIfMode })
            continue
        }

        if ($WhatIfMode) {
            Write-Status "[${idx}/${total}] WHATIF $id -> set IsExchangeCloudManaged=$target" "WARN"
            $results.Add([pscustomobject]@{ Timestamp=(Get-Date).ToString("s"); Row=$idx; Identity=$id; Mode=$mode; Result="WhatIf"; Reason="Would set IsExchangeCloudManaged=$target"; Before=$before; After=$target; WhatIf=$WhatIfMode })
            continue
        }

        Invoke-WithRetry -OperationName "Set-Mailbox $id IsExchangeCloudManaged=$target" -ScriptBlock {
            if ($mode -eq "Enable") { Set-Mailbox -Identity $id -IsExchangeCloudManaged $true  -ErrorAction Stop }
            else                    { Set-Mailbox -Identity $id -IsExchangeCloudManaged $false -ErrorAction Stop }
        } | Out-Null

        Write-Status "[${idx}/${total}] Updated $id (Before=$before After=$target)" "CHANGE"
        $results.Add([pscustomobject]@{ Timestamp=(Get-Date).ToString("s"); Row=$idx; Identity=$id; Mode=$mode; Result="Updated"; Reason=""; Before=$before; After=$target; WhatIf=$WhatIfMode })
    }
    catch {
        $msg = $_.Exception.Message
        Write-Status "[${idx}/${total}] Error ${id}: $msg" "ERROR"
        $results.Add([pscustomobject]@{ Timestamp=(Get-Date).ToString("s"); Row=$idx; Identity=$id; Mode=$modeRaw; Result="Error"; Reason=$msg; Before=""; After=""; WhatIf=$WhatIfMode })
        continue
    }
}

Write-Progress -Activity "Set Exchange Mailbox SOA (IsExchangeCloudManaged)" -Completed
#endregion =================================================================

#region ========================== EXPORT + SUMMARY ========================
$results | Export-Csv -LiteralPath $ResultsPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-Status "Results exported: $ResultsPath" "OK"

$updated = ($results | Where-Object Result -eq "Updated").Count
$skipped = ($results | Where-Object Result -eq "Skipped").Count
$whatif  = ($results | Where-Object Result -eq "WhatIf").Count
$errors  = ($results | Where-Object Result -eq "Error").Count

Write-Host ""
Write-Status "Run complete." "OK"
Write-Status "Total:   $($results.Count)" "INFO"
Write-Status "Updated: $updated" "CHANGE"
Write-Status "WhatIf:  $whatif" "WARN"
Write-Status "Skipped: $skipped" "SKIP"
Write-Status "Errors:  $errors" "ERROR"

try {
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Status "Disconnected from Exchange Online." "OK"
} catch {
    Write-Status "Disconnect warning: $($_.Exception.Message)" "WARN"
}

Stop-Transcript | Out-Null
#endregion =================================================================
