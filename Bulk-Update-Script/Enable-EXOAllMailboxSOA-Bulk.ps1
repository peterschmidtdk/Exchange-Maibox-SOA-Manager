<#
.SYNOPSIS
    Enable Exchange mailbox SOA (Exchange attribute SOA) for ALL mailboxes in the Exchange Online tenant.

.DESCRIPTION
    No CSV input. Reads every mailbox in the tenant and, for each one:
      - Reads IsDirSynced + IsExchangeCloudManaged
      - Classifies it: CloudManaged / OnPremManaged / CloudOnly
      - Depending on $RunMode (default "Ask" = menu at startup):
            Report   - only reports status on all mailboxes (no changes)
            WhatIf   - logs what would be changed (no changes)
            Apply    - sets SOA with: Set-Mailbox -Identity <Guid> -IsExchangeCloudManaged $true
            Rollback - reads the newest rollback CSV in $ExportDir and sets IsExchangeCloudManaged $false
                       on the mailboxes that the last Apply run updated
      - Logs everything to transcript + results CSV (written row by row, so a record exists even if the run is aborted)
      - Shows progress + colored output

    Safe to rerun: mailboxes that are already cloud managed are skipped.

    Apply mode also writes a rollback CSV (Identity,Mode=Disable) with the mailboxes that were updated.
    That file can be used directly as input for Enable-EXOMailboxSOA-Bulk.ps1.

    Connection behavior:
      - If already connected to Exchange Online in the current PowerShell session, it will NOT reconnect.
      - It will display which tenant it is connected to (best-effort).
      - Optional disconnect at the end via $DisconnectWhenDone.

.NOTES
    Author: Peter Schmidt
    Script Name: Enable-EXOAllMailboxSOA-Bulk.ps1
    Version: 1.1
    Updated: 2026-10-09
    Requires: ExchangeOnlineManagement module

.CHANGELOG
    1.1 (2026-10-09) - Startup menu (Report / WhatIf / Enable / Rollback). Added Rollback mode from last Apply run.
    1.0 (2026-10-09) - Initial version, based on Enable-EXOMailboxSOA-Bulk.ps1 v2.3.
#>

#region ========================== SCRIPT META ==========================
$ScriptName    = "Enable-EXOAllMailboxSOA-Bulk.ps1"
$ScriptVersion = "1.1"
$ScriptUpdated = "2026-10-09"
#endregion =================================================================

#region ========================== USER SETTINGS ==========================
$RunMode     = "Ask" # Ask = show menu at startup. Or fixed: Report = status on all mailboxes only. WhatIf = log what would be changed. Apply = enable SOA. Rollback = disable SOA on mailboxes updated in the last Apply run.
$AdminUPN    = ""   # optional: specify UPN for Connect-ExchangeOnline
$LogDir      = ".\Logs"
$ExportDir   = ".\Exports"

# Limit to specific mailbox types. Empty = all mailbox types.
# Example: @("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox")
$RecipientTypeDetails = @()

# Mailboxes to leave untouched (PrimarySmtpAddress, UPN or Alias). Logged as Skipped / Excluded.
$ExcludeIdentities = @()
$ExcludeFile       = "" # optional: text file with one identity per line

# Max number of mailboxes to change in one run (Apply/WhatIf). 0 = no limit. Use e.g. 10 for a pilot batch.
$MaxChanges = 0

# Ask for confirmation (tenant + number of mailboxes) before changing anything in Apply/Rollback mode?
$ConfirmBeforeApply = $true

$MaxRetries          = 5
$InitialDelaySeconds = 2
$MaxDelaySeconds     = 30

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
        "CHANGE" { Write-Host "[$ts] [ CHANGE] $Message" -ForegroundColor Green }
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
    Write-Host "  RunMode: $RunMode   MaxChanges: $MaxChanges   DisconnectWhenDone: $DisconnectWhenDone" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
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

function Add-Result {
    param(
        [Parameter(Mandatory)] $Mailbox,
        [Parameter(Mandatory)] [string] $SOAState,
        [Parameter(Mandatory)] [string] $Result,
        [string] $Reason = "",
        [string] $Before = "",
        [string] $After  = ""
    )

    $record = [pscustomobject]@{
        Timestamp            = (Get-Date).ToString("s")
        DisplayName          = [string]$Mailbox.DisplayName
        PrimarySmtpAddress   = [string]$Mailbox.PrimarySmtpAddress
        UserPrincipalName    = [string]$Mailbox.UserPrincipalName
        RecipientTypeDetails = [string]$Mailbox.RecipientTypeDetails
        IsDirSynced          = [string]$Mailbox.IsDirSynced
        SOAState             = $SOAState
        Before               = $Before
        After                = $After
        Result               = $Result
        Reason               = $Reason
        RunMode              = $RunMode
        Guid                 = [string]$Mailbox.Guid
    }

    $results.Add($record)
    # Appended row by row, so the record survives an aborted run.
    $record | Export-Csv -LiteralPath $ResultsPath -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Append
}
#endregion =================================================================

#region ========================== STARTUP ================================
$ErrorActionPreference = "Stop"

if ($RunMode -notin @("Ask","Report","WhatIf","Apply","Rollback")) {
    Write-Host "Invalid RunMode '$RunMode'. Allowed: Ask / Report / WhatIf / Apply / Rollback." -ForegroundColor Red
    exit 1
}

if ($RunMode -eq "Ask") {
    Write-Host ""
    Write-Host "==================== SELECT RUN MODE =======================" -ForegroundColor Cyan
    Write-Host "  1) Report   - show SOA status on all mailboxes (no changes)" -ForegroundColor Cyan
    Write-Host "  2) WhatIf   - show what Enable would change (no changes)" -ForegroundColor Cyan
    Write-Host "  3) Enable   - enable Exchange SOA on all eligible mailboxes" -ForegroundColor Cyan
    Write-Host "  4) Rollback - disable Exchange SOA on mailboxes updated in the last Enable run" -ForegroundColor Cyan
    Write-Host "  Q) Quit" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    while ($RunMode -eq "Ask") {
        switch ((Read-Host "Select 1-4 or Q").Trim().ToLowerInvariant()) {
            "1" { $RunMode = "Report" }
            "2" { $RunMode = "WhatIf" }
            "3" { $RunMode = "Apply" }
            "4" { $RunMode = "Rollback" }
            "q" { Write-Host "Quit - nothing done." -ForegroundColor Yellow; exit 0 }
            default { Write-Host "Invalid selection." -ForegroundColor Yellow }
        }
    }
}

Show-VersionBanner

Ensure-Folder -Path $LogDir
Ensure-Folder -Path $ExportDir

$runStamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
$TranscriptPath = Join-Path $LogDir    "Enable-EXOAllMailboxSOA-Bulk_${RunMode}_${ScriptVersion}_$runStamp.log.txt"
$ResultsPath    = Join-Path $ExportDir "Enable-EXOAllMailboxSOA-Bulk_${RunMode}_Results_${ScriptVersion}_$runStamp.csv"
$RollbackPath   = Join-Path $ExportDir "Enable-EXOAllMailboxSOA-Bulk_Rollback_${ScriptVersion}_$runStamp.csv"

Start-Transcript -Path $TranscriptPath -Force | Out-Null

Write-Status "Transcript: $TranscriptPath" "INFO"
Write-Status "Results:    $ResultsPath" "INFO"
Write-Status "RunMode:    $RunMode" "INFO"

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

# Exclusion list (settings + optional file), compared case-insensitively
$excluded = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($e in $ExcludeIdentities) {
    if (-not [string]::IsNullOrWhiteSpace($e)) { [void]$excluded.Add($e.Trim()) }
}
if (-not [string]::IsNullOrWhiteSpace($ExcludeFile)) {
    if (Test-Path -LiteralPath $ExcludeFile) {
        foreach ($e in (Get-Content -LiteralPath $ExcludeFile)) {
            if (-not [string]::IsNullOrWhiteSpace($e)) { [void]$excluded.Add($e.Trim()) }
        }
    } else {
        Write-Status "ExcludeFile not found: $ExcludeFile" "ERROR"
        Stop-Transcript | Out-Null
        exit 1
    }
}
Write-Status "Exclusions: $($excluded.Count)" "INFO"

$mailboxes    = @()
$rollbackRows = @()

if ($RunMode -eq "Rollback") {
    # Rollback only touches the mailboxes that the last Apply run updated
    $rollbackFile = Get-ChildItem -LiteralPath $ExportDir -Filter "Enable-EXOAllMailboxSOA-Bulk_Rollback_*.csv" -File |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $rollbackFile) {
        Write-Status "No rollback CSV found in $ExportDir - nothing to roll back." "ERROR"
        Stop-Transcript | Out-Null
        exit 1
    }

    $rollbackRows = @(Import-Csv -LiteralPath $rollbackFile.FullName)
    $total = $rollbackRows.Count
    Write-Status "Rollback file: $($rollbackFile.FullName)" "INFO"
    Write-Status "Mailboxes in rollback file: $total" "OK"

    if ($total -eq 0) {
        Write-Status "Rollback file is empty - nothing to roll back." "WARN"
        Stop-Transcript | Out-Null
        exit 1
    }

    if ($ConfirmBeforeApply) {
        Write-Host ""
        Write-Host "==================== CONFIRM ROLLBACK ======================" -ForegroundColor Yellow
        Write-Host "  Tenant:         $tenant" -ForegroundColor Yellow
        Write-Host "  Rollback file:  $($rollbackFile.Name)" -ForegroundColor Yellow
        Write-Host "  From run at:    $($rollbackFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Yellow
        Write-Host "  Mailboxes:      $total" -ForegroundColor Yellow
        Write-Host "============================================================" -ForegroundColor Yellow
        $answer = Read-Host "Type YES to disable Exchange SOA (IsExchangeCloudManaged=False) on these mailboxes"
        if ($answer -cne "YES") {
            Write-Status "Not confirmed - no changes made." "WARN"
            Stop-Transcript | Out-Null
            exit 1
        }
        Write-Status "Rollback confirmed for tenant $tenant." "OK"
    }
} else {
    try {
        Write-Status "Reading all mailboxes in the tenant (this can take a while)..." "INFO"
        $mailboxes = @(Invoke-WithRetry -OperationName "Get-Mailbox (all)" -ScriptBlock {
            if ($RecipientTypeDetails.Count -gt 0) {
                Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $RecipientTypeDetails -ErrorAction Stop
            } else {
                Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
            }
        })
    } catch {
        Write-Status "Failed to read mailboxes: $($_.Exception.Message)" "ERROR"
        Stop-Transcript | Out-Null
        exit 1
    }

    $total = $mailboxes.Count
    Write-Status "Mailboxes found: $total" "OK"
}

if ($mailboxes.Count -gt 0 -and -not ($mailboxes[0].PSObject.Properties.Name -contains "IsExchangeCloudManaged")) {
    Write-Status "Get-Mailbox does not return IsExchangeCloudManaged. Update the ExchangeOnlineManagement module / check that the feature is available in this tenant." "ERROR"
    Stop-Transcript | Out-Null
    exit 1
}

if ($RunMode -eq "Apply" -and $ConfirmBeforeApply) {
    $eligibleCount = @($mailboxes | Where-Object { $_.IsDirSynced -eq $true -and $_.IsExchangeCloudManaged -ne $true }).Count
    Write-Host ""
    Write-Host "==================== CONFIRM APPLY =========================" -ForegroundColor Yellow
    Write-Host "  Tenant:                   $tenant" -ForegroundColor Yellow
    Write-Host "  Mailboxes in scope:       $total" -ForegroundColor Yellow
    Write-Host "  Eligible (before excl.):  $eligibleCount" -ForegroundColor Yellow
    Write-Host "  MaxChanges:               $MaxChanges (0 = no limit)" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Yellow
    $answer = Read-Host "Type YES to enable Exchange SOA (IsExchangeCloudManaged=True) on these mailboxes"
    if ($answer -cne "YES") {
        Write-Status "Not confirmed - no changes made." "WARN"
        Stop-Transcript | Out-Null
        exit 1
    }
    Write-Status "Apply confirmed for tenant $tenant." "OK"
}
#endregion =================================================================

#region ========================== MAIN LOOP ==============================
$results = New-Object System.Collections.Generic.List[object]
$changeCount = 0 # Updated + WhatIf, counted against $MaxChanges

try {
    # Rollback: only the mailboxes from the last Apply run ($mailboxes is empty in this mode)
    for ($i = 0; $i -lt $rollbackRows.Count; $i++) {
        $idx = $i + 1
        $id  = ([string]$rollbackRows[$i].Identity).Trim()
        # Guid column is written by v1.1+; fall back to Identity for older rollback files
        $lookup = ([string]$rollbackRows[$i].Guid).Trim()
        if ([string]::IsNullOrWhiteSpace($lookup)) { $lookup = $id }

        $pct = [int](($idx / $total) * 100)
        Write-Progress -Activity "Exchange Mailbox SOA (IsExchangeCloudManaged) - $RunMode" -Status "Processing ${idx} of ${total}: $id" -PercentComplete $pct

        # Placeholder so the row is still logged if the mailbox cannot be read
        $mbx = [pscustomobject]@{ PrimarySmtpAddress = $id; Guid = [string]$rollbackRows[$i].Guid }

        if ([string]::IsNullOrWhiteSpace($lookup)) {
            Write-Status "[${idx}/${total}] Empty Identity - skipped." "SKIP"
            Add-Result -Mailbox $mbx -SOAState "Unknown" -Result "Skipped" -Reason "Empty Identity"
            continue
        }

        try {
            $mbx = Invoke-WithRetry -OperationName "Get-Mailbox $id" -ScriptBlock {
                Get-Mailbox -Identity $lookup -ErrorAction Stop
            }

            $before = [string]$mbx.IsExchangeCloudManaged
            $state  = if ([string]$mbx.IsDirSynced -ne "True") { "CloudOnly" } else { "OnPremManaged" }

            if ($before -ne "True") {
                Write-Status "[${idx}/${total}] Skipped $id - already IsExchangeCloudManaged=$before" "SKIP"
                Add-Result -Mailbox $mbx -SOAState $state -Result "Skipped" -Reason "Already False" -Before $before -After $before
                continue
            }

            Invoke-WithRetry -OperationName "Set-Mailbox $id IsExchangeCloudManaged=False" -ScriptBlock {
                Set-Mailbox -Identity $lookup -IsExchangeCloudManaged $false -ErrorAction Stop
            } | Out-Null

            Write-Status "[${idx}/${total}] Rolled back $id (Before=$before After=False)" "CHANGE"
            Add-Result -Mailbox $mbx -SOAState $state -Result "RolledBack" -Before $before -After "False"
        }
        catch {
            $msg = $_.Exception.Message
            Write-Status "[${idx}/${total}] Error ${id}: $msg" "ERROR"
            Add-Result -Mailbox $mbx -SOAState "Unknown" -Result "Error" -Reason $msg
            continue
        }
    }

    for ($i = 0; $i -lt $mailboxes.Count; $i++) {
        $idx  = $i + 1
        $mbx  = $mailboxes[$i]
        $id   = [string]$mbx.PrimarySmtpAddress
        $guid = [string]$mbx.Guid

        $pct = [int](($idx / $total) * 100)
        Write-Progress -Activity "Exchange Mailbox SOA (IsExchangeCloudManaged) - $RunMode" -Status "Processing ${idx} of ${total}: $id" -PercentComplete $pct

        try {
            $isDirSynced = [string]$mbx.IsDirSynced
            $before = [string]$mbx.IsExchangeCloudManaged

            $state = if ($isDirSynced -ne "True") { "CloudOnly" }
                     elseif ($before -eq "True")  { "CloudManaged" }
                     else                         { "OnPremManaged" }

            if ($RunMode -eq "Report") {
                switch ($state) {
                    "CloudManaged"  { Write-Status "[${idx}/${total}] $id - CloudManaged (SOA already enabled)" "OK" }
                    "OnPremManaged" { Write-Status "[${idx}/${total}] $id - OnPremManaged (eligible for SOA)" "WARN" }
                    default         { Write-Status "[${idx}/${total}] $id - CloudOnly (not DirSynced, SOA not applicable)" "INFO" }
                }
                Add-Result -Mailbox $mbx -SOAState $state -Result "Report" -Before $before -After $before
                continue
            }

            if ($state -eq "CloudOnly") {
                Write-Status "[${idx}/${total}] Skipped $id - not DirSynced (IsDirSynced=$isDirSynced)." "SKIP"
                Add-Result -Mailbox $mbx -SOAState $state -Result "Skipped" -Reason "Not DirSynced" -Before $before -After $before
                continue
            }

            if ($state -eq "CloudManaged") {
                Write-Status "[${idx}/${total}] Skipped $id - already IsExchangeCloudManaged=True" "SKIP"
                Add-Result -Mailbox $mbx -SOAState $state -Result "Skipped" -Reason "Already True" -Before $before -After $before
                continue
            }

            if ($excluded.Contains($id) -or $excluded.Contains([string]$mbx.UserPrincipalName) -or $excluded.Contains([string]$mbx.Alias)) {
                Write-Status "[${idx}/${total}] Skipped $id - on exclusion list." "SKIP"
                Add-Result -Mailbox $mbx -SOAState $state -Result "Skipped" -Reason "Excluded" -Before $before -After $before
                continue
            }

            if ($MaxChanges -gt 0 -and $changeCount -ge $MaxChanges) {
                Write-Status "[${idx}/${total}] Skipped $id - MaxChanges ($MaxChanges) reached." "SKIP"
                Add-Result -Mailbox $mbx -SOAState $state -Result "Skipped" -Reason "MaxChanges reached" -Before $before -After $before
                continue
            }

            if ($RunMode -eq "WhatIf") {
                $changeCount++
                Write-Status "[${idx}/${total}] WHATIF $id -> set IsExchangeCloudManaged=True" "WARN"
                Add-Result -Mailbox $mbx -SOAState $state -Result "WhatIf" -Reason "Would set IsExchangeCloudManaged=True" -Before $before -After "True"
                continue
            }

            # Counted before the call, so failed attempts also count against MaxChanges
            $changeCount++
            Invoke-WithRetry -OperationName "Set-Mailbox $id IsExchangeCloudManaged=True" -ScriptBlock {
                Set-Mailbox -Identity $guid -IsExchangeCloudManaged $true -ErrorAction Stop
            } | Out-Null

            Write-Status "[${idx}/${total}] Updated $id (Before=$before After=True)" "CHANGE"
            Add-Result -Mailbox $mbx -SOAState "CloudManaged" -Result "Updated" -Before $before -After "True"
        }
        catch {
            $msg = $_.Exception.Message
            Write-Status "[${idx}/${total}] Error ${id}: $msg" "ERROR"
            Add-Result -Mailbox $mbx -SOAState "Unknown" -Result "Error" -Reason $msg
            continue
        }
    }
}
finally {
    Write-Progress -Activity "Exchange Mailbox SOA (IsExchangeCloudManaged) - $RunMode" -Completed

    #region ====================== EXPORT + SUMMARY ========================
    if ($results.Count -gt 0) { Write-Status "Results exported: $ResultsPath" "OK" }

    $updatedRows = @($results | Where-Object Result -eq "Updated")
    if ($updatedRows.Count -gt 0) {
        # Same format as the input CSV for Enable-EXOMailboxSOA-Bulk.ps1 (extra Guid column is used by Rollback mode)
        $updatedRows |
            Select-Object @{n="Identity";e={$_.PrimarySmtpAddress}}, @{n="Mode";e={"Disable"}}, Guid |
            Export-Csv -LiteralPath $RollbackPath -NoTypeInformation -Encoding UTF8
        Write-Status "Rollback CSV (input for Enable-EXOMailboxSOA-Bulk.ps1): $RollbackPath" "OK"
    }

    $cloudManaged  = @($results | Where-Object SOAState -eq "CloudManaged").Count
    $onPremManaged = @($results | Where-Object SOAState -eq "OnPremManaged").Count
    $cloudOnly     = @($results | Where-Object SOAState -eq "CloudOnly").Count
    $whatif        = @($results | Where-Object Result -eq "WhatIf").Count
    $skipped       = @($results | Where-Object Result -eq "Skipped").Count
    $errors        = @($results | Where-Object Result -eq "Error").Count

    Write-Host ""
    if ($results.Count -lt $total) {
        Write-Status "Run ABORTED after $($results.Count) of $total mailboxes." "WARN"
    } else {
        Write-Status "Run complete." "OK"
    }
    Write-Status "Tenant:  $tenant" "INFO"
    Write-Status "RunMode: $RunMode" "INFO"
    Write-Status "Total:   $($results.Count)" "INFO"

    Write-Status "SOA status after this run:" "INFO"
    Write-Status "  CloudManaged  (SOA enabled):         $cloudManaged" "OK"
    Write-Status "  OnPremManaged (DirSynced, eligible): $onPremManaged" "WARN"
    Write-Status "  CloudOnly     (not DirSynced):       $cloudOnly" "INFO"

    if ($RunMode -ne "Report") {
        Write-Status "Actions:" "INFO"
        Write-Status "  Updated: $($updatedRows.Count)" "CHANGE"
        Write-Status "  RolledBack: $(@($results | Where-Object Result -eq 'RolledBack').Count)" "CHANGE"
        Write-Status "  WhatIf:  $whatif" "WARN"
        Write-Status "  Skipped: $skipped" "SKIP"
        $results | Where-Object Result -eq "Skipped" | Group-Object Reason | Sort-Object Name | ForEach-Object {
            Write-Status "    $($_.Name): $($_.Count)" "SKIP"
        }
    }
    Write-Status "  Errors:  $errors" "ERROR"

    Write-Status "By mailbox type:" "INFO"
    $results | Group-Object RecipientTypeDetails | Sort-Object Name | ForEach-Object {
        $cm = @($_.Group | Where-Object SOAState -eq "CloudManaged").Count
        $op = @($_.Group | Where-Object SOAState -eq "OnPremManaged").Count
        $co = @($_.Group | Where-Object SOAState -eq "CloudOnly").Count
        Write-Status "  $($_.Name): Total=$($_.Count) CloudManaged=$cm OnPremManaged=$op CloudOnly=$co" "INFO"
    }

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
    #endregion =============================================================
}
#endregion =================================================================
