<#
.SYNOPSIS
  Mailbox SOA Manager (GUI) for Exchange Online - Cloud-managed Exchange attributes (SOA) toggle.

.DESCRIPTION
  In Exchange Hybrid environments, Microsoft introduced a per-mailbox switch to transfer the
  Source of Authority (SOA) for Exchange attributes from on-premises to Exchange Online.

  This tool provides a Windows GUI to:
    - Connect to Exchange Online
    - Search for EXO mailboxes
    - Show SOA indicator in list: "SOA (Exchange Attributes)" = Online / On-Prem / Unknown
    - View IsDirSynced + IsExchangeCloudManaged state
    - Enable cloud management (IsExchangeCloudManaged = $true)
    - Revert to on-premises management (IsExchangeCloudManaged = $false)
    - Export a small backup (JSON) of mailbox + user properties before reverting

REFERENCE
  Microsoft Learn:
  https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

LOGGING
  - Single logfile only (append; never overwritten)
  - Timestamp on every line
  - All SOA changes logged with BEFORE/AFTER + Actor (Windows user + EXO identity if available)
  - RunId included for correlation

REQUIREMENTS
  - Windows PowerShell 5.1 OR PowerShell 7+ started with -STA
  - Module: ExchangeOnlineManagement

AUTHOR
  Peter

VERSION
  1.3 (2026-01-04)
#>

#region Safety / STA check
try {
    $apt = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apt -ne [System.Threading.ApartmentState]::STA) {
        Write-Warning "This GUI must run in STA mode. Start PowerShell with -STA and re-run."
        Write-Warning "Examples:"
        Write-Warning "  Windows PowerShell: powershell.exe -STA -File .\MailboxSOAManager-GUI.ps1"
        Write-Warning "  PowerShell 7+:      pwsh.exe -STA -File .\MailboxSOAManager-GUI.ps1"
        return
    }
} catch {
    # best effort
}
#endregion

#region Globals
$Script:ToolName     = "Mailbox SOA Manager"
$Script:RunId        = [Guid]::NewGuid().ToString()
$Script:LogDir       = Join-Path -Path (Get-Location) -ChildPath "Logs"
$Script:ExportDir    = Join-Path -Path (Get-Location) -ChildPath "Exports"
$Script:LogFile      = Join-Path -Path $Script:LogDir -ChildPath "MailboxSOAManager.log"
$Script:IsConnected  = $false
$Script:ExoActor     = $null  # populated after connect (best-effort)

New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null
New-Item -ItemType Directory -Path $Script:ExportDir -Force | Out-Null
#endregion

#region Logging (single logfile, timestamp per line)
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $winUser = "$env:USERDOMAIN\$env:USERNAME"
    $actor = if ($Script:ExoActor) { $Script:ExoActor } else { "EXO:unknown" }
    $line = "[$ts][$Level][RunId:$($Script:RunId)][Win:$winUser][$actor] $Message"
    Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8
}
#endregion

#region Module helpers
function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "Required module '$Name' is not installed.`n`nInstall it now (CurrentUser)?",
            "Missing Module",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
            throw "Module '$Name' not installed."
        }

        Write-Log "Installing module '$Name' (Scope=CurrentUser)..." "INFO"
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Log "Module '$Name' installed." "INFO"
    }

    Import-Module $Name -ErrorAction Stop
    Write-Log "Module loaded: $Name" "INFO"
}
#endregion

#region SOA indicator helper
function Get-SOAIndicator {
    param([object]$IsExchangeCloudManaged)
    if ($IsExchangeCloudManaged -eq $true)  { return "☁ Online" }
    if ($IsExchangeCloudManaged -eq $false) { return "🏢 On-Prem" }
    return "? Unknown"
}
#endregion

#region EXO connect/disconnect
function Get-ExoActorBestEffort {
    try {
        # Get-ConnectionInformation exists in newer ExchangeOnlineManagement builds.
        $ci = Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($ci) {
            $info = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
            if ($info -and $info.UserPrincipalName) {
                return "EXO:$($info.UserPrincipalName)"
            }
        }
    } catch {
        # ignore
    }
    return "EXO:unknown"
}

function Connect-EXO {
    try {
        Ensure-Module -Name "ExchangeOnlineManagement"

        Write-Log "Connecting to Exchange Online..." "INFO"
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
        $Script:IsConnected = $true
        $Script:ExoActor = Get-ExoActorBestEffort

        Write-Log "Connected to Exchange Online." "INFO"
        return $true
    } catch {
        $Script:IsConnected = $false
        $Script:ExoActor = "EXO:unknown"
        Write-Log "Connect-EXO failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Exchange Online.`n`n$($_.Exception.Message)",
            "Connect Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
}

function Disconnect-EXO {
    try {
        if ($Script:IsConnected) {
            Write-Log "Disconnecting from Exchange Online..." "INFO"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Disconnected from Exchange Online." "INFO"
        }
    } catch {
        Write-Log "Disconnect-EXO warning: $($_.Exception.Message)" "WARN"
    } finally {
        $Script:IsConnected = $false
        $Script:ExoActor = $null
    }
}
#endregion

#region Mailbox ops
function Search-Mailboxes {
    param(
        [Parameter(Mandatory)][string]$QueryText,
        [int]$Max = 200
    )
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $q = $QueryText.Trim()
    if ([string]::IsNullOrWhiteSpace($q)) { return @() }

    $filter = "DisplayName -like '*$q*' -or Alias -like '*$q*' -or PrimarySmtpAddress -like '*$q*'"
    Write-Log "Search-Mailboxes started. Query='$q' Max=$Max Filter='$filter'" "INFO"

    $items = Get-Mailbox -ResultSize $Max -Filter $filter -ErrorAction Stop |
        Select-Object `
            DisplayName,
            Alias,
            PrimarySmtpAddress,
            RecipientTypeDetails,
            IsDirSynced,
            IsExchangeCloudManaged,
            @{Name="SOA (Exchange Attributes)"; Expression={ Get-SOAIndicator $_.IsExchangeCloudManaged }}

    Write-Log "Search-Mailboxes completed. Results=$($items.Count)" "INFO"
    return @($items)
}

function Get-MailboxDetails {
    param([Parameter(Mandatory)][string]$Identity)

    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,Alias,PrimarySmtpAddress,RecipientTypeDetails,IsDirSynced,IsExchangeCloudManaged,ExchangeGuid,ExternalDirectoryObjectId

    $usr = $null
    try {
        $usr = Get-User -Identity $Identity -ErrorAction Stop |
            Select-Object DisplayName,UserPrincipalName,ImmutableId,RecipientTypeDetails,WhenChangedUTC
    } catch {
        Write-Log "Get-User failed (non-fatal) for '$Identity': $($_.Exception.Message)" "WARN"
    }

    [PSCustomObject]@{
        Mailbox = $mbx
        User    = $usr
    }
}

function Export-MailboxBackup {
    param([Parameter(Mandatory)][string]$
