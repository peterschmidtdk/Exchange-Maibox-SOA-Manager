<#
.SYNOPSIS
    Bulk enable/disable Exchange mailbox SOA (Exchange attribute SOA) for mailboxes in Exchange Online.

.DESCRIPTION
    Reads a CSV with Identity + optional Mode (Enable/Disable). For each row:
      - Reads mailbox: IsDirSynced + IsExchangeCloudManaged
      - If eligible, sets SOA with:
            Set-Mailbox -Identity <Identity> -IsExchangeCloudManaged $true / $false
      - Logs everything to transcript + results CSV
      - Shows progress + colored output

    Connection behavior:
      - If already connected to Exchange Online in the current PowerShell session, it will NOT reconnect.
      - It will display which tenant it is connected to (best-effort).
      - Optional disconnect at the end via $DisconnectWhenDone.

.CSV EXAMPLE
    Default path: .\MailboxSOA-Bulk.csv

    Identity,Mode
    user1@contoso.com,Enable
    user2@contoso.com,Enable
    shared.mailbox@contoso.com,Enable
    room101@contoso.com,Disable
    aliasOnlyUser,Enable

.NOTES
    Author: Peter Schmidt
    Script Name: Enable-EXOMailboxSOA-Bulk.ps1
    Version: 2.3
    Updated: 2026-02-28
    Requires: ExchangeOnlineManagement module

.CHANGELOG
    2.3 (2026-02-28) - CHANGE output color switched to Green.
    2.2 (2026-02-28) - Added $DisconnectWhenDone setting.
    2.1 (2026-02-28) - Skip Connect-ExchangeOnline if already connected; show tenant info.
    2.0 (2026-02-24) - Show version banner at runtime.
    1.9 (2026-02-24) - Missing/empty CSV messages changed from red to purple (Magenta).
    1.8 (2026-02-24) - Fully friendly CSV handling: no raw throw, auto-creates template CSV if missing.
    1.7 (2026-02-24) - Added friendly message on missing/empty CSV.
    1.6 (2026-02-24) - Fixed all "$var:" parsing issues by using ${var}.
    1.4 (2026-02-24) - Corrected command to Set-Mailbox -IsExchangeCloudManaged.
#>

#region ========================== SCRIPT META ==========================
$ScriptName    = "Enable-EXOMailboxSOA-Bulk.ps1"
$ScriptVersion = "2.3"
$ScriptUpdated = "2026-02-28"
#endregion =================================================================

#region ========================== USER SETTINGS ==========================
$CsvPath     = ".\MailboxSOA-Bulk.csv" # CSV with Identity,Mode (Enable/Disable)
$WhatIfMode  = $true # If true, will not make changes, only log what would be done. Set to $false to apply changes.
$DefaultMode = "Enable" # If Mode column is empty/whitespace, use this default. Allowed: Enable/Disable (case-insensitive)
$AdminUPN    = ""   # optional: specify UPN for Connect-ExchangeOnline
$LogDir      = ".\Logs"
$ExportDir   = ".\Exports"

$MaxRetries          = 5
$InitialDelaySeconds = 2
$MaxDelaySeconds     = 30

# Friendly input (missing/empty CSV) color
# Options: Magenta or DarkMagenta
$FriendlyColor = "Magenta" 

# Disconnect from Exchange Online when done?
$DisconnectWhenDone = $false # Set to $false to keep session open after script finishes (for testing/verification)
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
        "CHANGE" { Write-Host "[$ts] [ CHANGE] $Message" -ForegroundColor Green } # v2.3: was Magenta
        "WARN"   { Write-Host "[$ts] [ WARN  ] $Message" -ForegroundColor Yellow }
        "ERROR"  { Write-Host "[$ts] [ ERROR ] $Message" -ForegroundColor Red }
        "SKIP"   { Write-Host "[$ts] [ SKIP  ] $Message" -ForegroundColor DarkYellow }
        default  { Write-Host "[$ts] [ INFO  ] $Message" -ForegroundColor Cyan }
    }
}

function Show-VersionBanner {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  $ScriptName" -ForegroundColor Cyan
    Write-Host "  Version: $ScriptVersion   Updated: $ScriptUpdated   Author: Peter" -ForegroundColor Cyan
    Write-Host "  WhatIfMode: $WhatIfMode   DefaultMode: $DefaultMode   DisconnectWhenDone: $DisconnectWhenDone" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function New-TemplateCsv {
    param([Parameter(Mandatory)][string]$Path)

    $template = @(
        "Identity,Mode"
        "user1@contoso.com,Enable"
        "user2@contoso.com,Enable"
        "shared.mailbox@contoso.com,Enable"
        "room101@contoso.com,Disable"
        "aliasOnlyUser,Enable"
    )

    $dir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($dir)) { Ensure-Folder -Path $dir }

    Set-Content -LiteralPath $Path -Value $template -Encoding UTF8
}

function Ensure-FriendlyCsvReady {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Host ""
        Write-Host "==================== INPUT FILE MISSING ====================" -ForegroundColor $FriendlyColor
        Write-Host "I couldn't find the CSV input file here:" -ForegroundColor $FriendlyColor
        Write-Host "  $Path" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "I will create a template CSV for you now." -ForegroundColor $FriendlyColor
        Write-Host "Fill in your real users, then run the script again." -ForegroundColor $FriendlyColor
        Write-Host "============================================================" -ForegroundColor $FriendlyColor
        Write-Host ""

        New-TemplateCsv -Path $Path

        Write-Host "Template created:" -ForegroundColor Green
        Write-Host "  $Path" -ForegroundColor Green
        Write-Host ""
        Write-Host "Open it, edit identities, then rerun:" -ForegroundColor $FriendlyColor
        Write-Host "  .\Enable-EXOMailboxSOA-Bulk.ps1" -ForegroundColor Gray
        Write-Host ""
        exit 1
    }

    $content = Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $content -or $content.Count -lt 2) {
        Write-Host ""
        Write-Host "==================== INPUT FILE EMPTY ======================" -ForegroundColor $FriendlyColor
        Write-Host "The CSV exists but is empty (or only headers):" -ForegroundColor $FriendlyColor
        Write-Host "  $Path" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "I will overwrite it with a template CSV now." -ForegroundColor $FriendlyColor
        Write-Host "============================================================" -ForegroundColor $FriendlyColor
        Write-Host ""

        New-TemplateCsv -Path $Path

        Write-Host "Template written:" -ForegroundColor Green
        Write-Host "  $Path" -ForegroundColor Green
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

function Get-EXOConnectionInfo {
    try {
        if (Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            return (Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1)
        }
        return $null
    } catch {
        return $null
    }
}

function Test-EXOConnected {
    $ci = Get-EXOConnectionInfo
    if ($null -ne $ci) {
        if ($ci.PSObject.Properties.Name -contains "State") { return ([string]$ci.State -match "Connected") }
        if ($ci.PSObject.Properties.Name -contains "IsConnected") { return [bool]$ci.IsConnected }
        return $true
    }

    try { $null = Get-OrganizationConfig -ErrorAction Stop; return $true }
    catch { return $false }
}

function Get-EXOTenantLabel {
    try {
        $defaultDomain = (Get-AcceptedDomain -ErrorAction Stop | Where-Object { $_.Default -eq $true } | Select-Object -First 1).DomainName
        if (-not [string]::IsNullOrWhiteSpace($defaultDomain)) { return [string]$defaultDomain }
    } catch { }

    try {
        $org = Get-OrganizationConfig -ErrorAction Stop
        if ($org -and ($org.PSObject.Properties.Name -contains "Name") -and -not [string]::IsNullOrWhiteSpace($org.Name)) {
            return [string]$org.Name
        }
    } catch { }

    $ci = Get-EXOConnectionInfo
    if ($null -ne $ci) {
        foreach ($p in @("Organization","DelegatedOrganization","Tenant","TenantId","ConnectionUri","UserPrincipalName")) {
            if ($ci.PSObject.Properties.Name -contains $p) {
                $v = [string]$ci.$p
                if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
            }
        }
    }

    return "Unknown tenant"
}
#endregion =================================================================

#region ========================== STARTUP ================================
$ErrorActionPreference = "Stop"

Show-VersionBanner
Ensure-FriendlyCsvReady -Path $CsvPath

Ensure-Folder -Path $LogDir
Ensure-Folder -Path $ExportDir

$runStamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
$TranscriptPath = Join-Path $LogDir    "Enable-EXOMailboxSOA_${ScriptVersion}_$runStamp.log.txt"
$ResultsPath    = Join-Path $ExportDir "Enable-EXOMailboxSOA_Results_${ScriptVersion}_$runStamp.csv"

Start-Transcript -Path $TranscriptPath -Force | Out-Null

Write-Status "Transcript: $TranscriptPath" "INFO"
Write-Status "Results:    $ResultsPath" "INFO"
Write-Status "CSV:        $CsvPath" "INFO"

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Status "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement" "ERROR"
    Stop-Transcript | Out-Null
    exit 1
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

if (Test-EXOConnected) {
    $tenant = Get-EXOTenantLabel
    Write-Status "Already connected to Exchange Online. Tenant: $tenant" "OK"
} else {
    try {
        if ([string]::IsNullOrWhiteSpace($AdminUPN)) {
            Write-Status "Connecting to Exchange Online (interactive)..." "INFO"
            Connect-ExchangeOnline -ShowBanner:$false
        } else {
            Write-Status "Connecting to Exchange Online as $AdminUPN ..." "INFO"
            Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
        }
        $tenant = Get-EXOTenantLabel
        Write-Status "Connected to Exchange Online. Tenant: $tenant" "OK"
    } catch {
        Write-Status "Failed to connect to Exchange Online: $($_.Exception.Message)" "ERROR"
        Stop-Transcript | Out-Null
        exit 1
    }
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
        $target = if ($mode -eq "Enable") { "True" } else { "False" }

        if ($isDirSynced -ne "True") {
            Write-Status "[${idx}/${total}] Skipped $id - not DirSynced (IsDirSynced=$isDirSynced)." "SKIP"
            $results.Add([pscustomobject]@{ Timestamp=(Get-Date).ToString("s"); Row=$idx; Identity=$id; Mode=$mode; Result="Skipped"; Reason="Not DirSynced"; Before=$before; After=$before; WhatIf=$WhatIfMode })
            continue
        }

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

if ($DisconnectWhenDone) {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Status "Disconnected from Exchange Online (DisconnectWhenDone=True)." "OK"
    } catch {
        Write-Status "Disconnect warning: $($_.Exception.Message)" "WARN"
    }
} else {
    Write-Status "Leaving Exchange Online session connected (DisconnectWhenDone=False)." "INFO"
}

Stop-Transcript | Out-Null
#endregion =================================================================
