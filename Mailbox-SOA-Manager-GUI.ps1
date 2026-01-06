<#
.SYNOPSIS
  SOA Mailbox Tool (GUI) for Exchange Online - Cloud-managed Exchange attributes (SOA) toggle.

.DESCRIPTION
  In Exchange Hybrid environments, Microsoft introduced a per-mailbox switch to transfer the
  Source of Authority (SOA) for Exchange attributes from on-premises to Exchange Online.

  This tool provides a Windows GUI to:
    - Connect/Disconnect to Exchange Online
    - Search for EXO mailboxes
    - View SOA status + DirSynced state
    - Enable cloud management (IsExchangeCloudManaged = $true)
    - Revert to on-premises management (IsExchangeCloudManaged = $false)
    - Open the logfile

  Microsoft reference:
    - https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

IMPORTANT NOTES
  - This does NOT migrate mailboxes. It only changes where Exchange *attributes* are managed.
  - For DirSynced users, next sync cycles can overwrite values depending on management state.

REQUIREMENTS
  - Windows PowerShell 5.1 OR PowerShell 7+ started with -STA
  - Module: ExchangeOnlineManagement
  - Appropriate Exchange Online permissions (Exchange Admin recommended)

AUTHOR
  Peter Schmidt

VERSION
  2.5.0 (2026-01-05)
#>

#region Safety / STA check
try {
    $apt = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apt -ne [System.Threading.ApartmentState]::STA) {
        Write-Warning "This GUI must run in STA mode. Start PowerShell with -STA and re-run."
        Write-Warning "Examples:"
        Write-Warning "  Windows PowerShell: powershell.exe -STA -File .\SOA-MailboxTool-GUI.ps1"
        Write-Warning "  PowerShell 7+:      pwsh.exe -STA -File .\SOA-MailboxTool-GUI.ps1"
        return
    }
} catch {
    # Best effort
}
#endregion

#region Globals
$Script:ToolName      = "SOA Mailbox Tool"
$Script:ScriptVersion = "2.5.0"

$Script:LogDir        = Join-Path -Path (Get-Location) -ChildPath "Logs"
$Script:LogFile       = Join-Path -Path $Script:LogDir -ChildPath "SOA-MailboxTool.log"
$Script:IsConnected   = $false

$Script:TenantName    = "Unknown"
$Script:MailboxCount  = 0

New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null
#endregion

#region Logging
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[{0}][{1}] {2}" -f $ts, $Level, $Message
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
            throw ("Module '{0}' not installed." -f $Name)
        }

        Write-Log ("Installing module '{0}' for CurrentUser..." -f $Name) "INFO"
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Write-Log ("Install-Module failed for '{0}': {1}" -f $Name, $_.Exception.Message) "ERROR"
            throw
        }
    }

    Import-Module $Name -ErrorAction Stop
    Write-Log ("Module loaded: {0}" -f $Name) "INFO"
}
#endregion

#region EXO connect/disconnect + Tenant info
function Get-TenantNameBestEffort {
    try {
        $org = Get-OrganizationConfig -ErrorAction Stop
        if ($org -and $org.Name) { return [string]$org.Name }
    } catch { }
    return "Unknown"
}

function Get-MailboxCountBestEffort {
    # Count can be slow in very large tenants; best effort with EXO cmdlet if available.
    try {
        $exoCmd = Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue
        if ($exoCmd) {
            # Count all mailbox types returned (within your RBAC scope)
            $c = (Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop | Measure-Object).Count
            return [int]$c
        }
    } catch {
        Write-Log ("Get-EXOMailbox count failed: {0}" -f $_.Exception.Message) "WARN"
    }

    try {
        # Fallback
        $c2 = (Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | Measure-Object).Count
        return [int]$c2
    } catch {
        Write-Log ("Get-Mailbox count failed: {0}" -f $_.Exception.Message) "WARN"
    }

    return 0
}

function Connect-EXO {
    try {
        Ensure-Module -Name "ExchangeOnlineManagement"

        Write-Log "Connecting to Exchange Online..." "INFO"
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
        $Script:IsConnected = $true

        # Tenant info
        $Script:TenantName   = Get-TenantNameBestEffort
        $Script:MailboxCount = Get-MailboxCountBestEffort

        Write-Log ("Connected to Exchange Online. Tenant='{0}' Mailboxes={1}" -f $Script:TenantName, $Script:MailboxCount) "INFO"
        return $true
    } catch {
        $Script:IsConnected = $false
        Write-Log ("Connect-EXO failed: {0}" -f $_.Exception.Message) "ERROR"
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
        }
    } catch {
        Write-Log ("Disconnect-EXO warning: {0}" -f $_.Exception.Message) "WARN"
    } finally {
        $Script:IsConnected  = $false
        $Script:TenantName   = "Unknown"
        $Script:MailboxCount = 0
        Write-Log "Disconnected (or no session)." "INFO"
    }
}
#endregion

#region Mailbox ops
function Get-SOAStatusText {
    param(
        [Parameter(Mandatory)][object]$MailboxRow
    )

    # Intended for DirSynced. If not DirSynced, show N/A.
    if ($MailboxRow.IsDirSynced -ne $true) {
        return "N/A (Not DirSynced)"
    }

    if ($MailboxRow.IsExchangeCloudManaged -eq $true)  { return "Online" }
    if ($MailboxRow.IsExchangeCloudManaged -eq $false) { return "On-Prem" }
    return "Unknown"
}

function Search-Mailboxes {
    param(
        [Parameter(Mandatory)][string]$QueryText,
        [int]$Max = 200
    )
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $q = $QueryText.Trim()
    if ([string]::IsNullOrWhiteSpace($q)) { return @() }

    # Use OPATH filter for performance (best effort)
    $filter = "DisplayName -like '*$q*' -or Alias -like '*$q*' -or PrimarySmtpAddress -like '*$q*'"

    $items = Get-Mailbox -ResultSize $Max -Filter $filter -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

    # Shape data for grid view (ONLY the requested columns)
    $view = foreach ($m in $items) {
        [PSCustomObject]@{
            DisplayName = [string]$m.DisplayName
            PrimarySMTP = [string]$m.PrimarySmtpAddress
            'SOA Status' = (Get-SOAStatusText -MailboxRow $m)
            DirSynced   = if ($m.IsDirSynced -eq $true) { "Yes" } else { "No" }
        }
    }

    return @($view)
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
        # Not fatal
    }

    [PSCustomObject]@{
        Mailbox = $mbx
        User    = $usr
    }
}

function Set-MailboxSOACloudManaged {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][bool]$EnableCloudManaged
    )
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $targetValue = [bool]$EnableCloudManaged

    $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

    if ($mbx.IsDirSynced -ne $true) {
        throw "Mailbox '$Identity' is not DirSynced (IsDirSynced=$($mbx.IsDirSynced)). This switch is intended for directory-synchronized users."
    }

    if ($mbx.IsExchangeCloudManaged -eq $targetValue) {
        return "No change needed. IsExchangeCloudManaged is already '$targetValue'."
    }

    Set-Mailbox -Identity $Identity -IsExchangeCloudManaged $targetValue -ErrorAction Stop
    Write-Log ("Set-Mailbox '{0}' IsExchangeCloudManaged={1}" -f $Identity, $targetValue) "INFO"
    return "Updated. IsExchangeCloudManaged is now '$targetValue'."
}
#endregion

#region GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "{0} v{1} - Exchange Online (IsExchangeCloudManaged)" -f $Script:ToolName, $Script:ScriptVersion
$form.Size = New-Object System.Drawing.Size(980, 700)
$form.StartPosition = "CenterScreen"

# Top bar
$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect"
$btnConnect.Location = New-Object System.Drawing.Point(12, 12)
$btnConnect.Size = New-Object System.Drawing.Size(110, 30)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Text = "Disconnect"
$btnDisconnect.Location = New-Object System.Drawing.Point(130, 12)
$btnDisconnect.Size = New-Object System.Drawing.Size(110, 30)
$btnDisconnect.Enabled = $false

$btnOpenLog = New-Object System.Windows.Forms.Button
$btnOpenLog.Text = "Open log"
$btnOpenLog.Location = New-Object System.Drawing.Point(248, 12)
$btnOpenLog.Size = New-Object System.Drawing.Size(110, 30)

$lblConn = New-Object System.Windows.Forms.Label
$lblConn.Text = "Status: Not connected"
$lblConn.Location = New-Object System.Drawing.Point(370, 18)
$lblConn.AutoSize = $true

# Search
$grpSearch = New-Object System.Windows.Forms.GroupBox
$grpSearch.Text = "Search mailboxes (DisplayName, Alias, Primary SMTP)"
$grpSearch.Location = New-Object System.Drawing.Point(12, 55)
$grpSearch.Size = New-Object System.Drawing.Size(940, 220)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(16, 30)
$txtSearch.Size = New-Object System.Drawing.Size(720, 25)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Search"
$btnSearch.Location = New-Object System.Drawing.Point(750, 27)
$btnSearch.Size = New-Object System.Drawing.Size(170, 30)
$btnSearch.Enabled = $false

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "Results: 0"
$lblCount.Location = New-Object System.Drawing.Point(16, 62)
$lblCount.AutoSize = $true

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(16, 85)
$grid.Size = New-Object System.Drawing.Size(904, 120)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = "Fill"
$grid.AutoGenerateColumns = $true

# Details
$grpDetails = New-Object System.Windows.Forms.GroupBox
$grpDetails.Text = "Selected mailbox details"
$grpDetails.Location = New-Object System.Drawing.Point(12, 285)
$grpDetails.Size = New-Object System.Drawing.Size(940, 270)

$txtDetails = New-Object System.Windows.Forms.TextBox
$txtDetails.Location = New-Object System.Drawing.Point(16, 30)
$txtDetails.Size = New-Object System.Drawing.Size(904, 165)
$txtDetails.Multiline = $true
$txtDetails.ScrollBars = "Vertical"
$txtDetails.ReadOnly = $true

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh details"
$btnRefresh.Location = New-Object System.Drawing.Point(16, 205)
$btnRefresh.Size = New-Object System.Drawing.Size(170, 32)
$btnRefresh.Enabled = $false

$btnEnableCloud = New-Object System.Windows.Forms.Button
$btnEnableCloud.Text = "Enable cloud SOA (true)"
$btnEnableCloud.Location = New-Object System.Drawing.Point(196, 205)
$btnEnableCloud.Size = New-Object System.Drawing.Size(240, 32)
$btnEnableCloud.Enabled = $false

$btnRevertOnPrem = New-Object System.Windows.Forms.Button
$btnRevertOnPrem.Text = "Revert to on-prem SOA (false)"
$btnRevertOnPrem.Location = New-Object System.Drawing.Point(446, 205)
$btnRevertOnPrem.Size = New-Object System.Drawing.Size(270, 32)
$btnRevertOnPrem.Enabled = $false

# Footer info
$lblFoot = New-Object System.Windows.Forms.Label
$lblFoot.Location = New-Object System.Drawing.Point(12, 565)
$lblFoot.Size = New-Object System.Drawing.Size(940, 90)
$lblFoot.Text =
("Notes:
- This tool changes IsExchangeCloudManaged only (does not migrate mailboxes).
- Intended primarily for DirSynced mailboxes.
Log: {0}
Version: {1}
" -f $Script:LogFile, $Script:ScriptVersion)
$lblFoot.AutoSize = $false

# Add controls
$form.Controls.AddRange(@($btnConnect,$btnDisconnect,$btnOpenLog,$lblConn,$grpSearch,$grpDetails,$lblFoot))
$grpSearch.Controls.AddRange(@($txtSearch,$btnSearch,$lblCount,$grid))
$grpDetails.Controls.AddRange(@($txtDetails,$btnRefresh,$btnEnableCloud,$btnRevertOnPrem))

# State tracking
$Script:SelectedIdentity = $null

function Set-UiConnectedState {
    param([bool]$Connected)

    $btnConnect.Enabled    = -not $Connected
    $btnDisconnect.Enabled = $Connected
    $btnSearch.Enabled     = $Connected

    $btnRefresh.Enabled     = $false
    $btnEnableCloud.Enabled = $false
    $btnRevertOnPrem.Enabled= $false

    $lblCount.Text = "Results: 0"

    if ($Connected) {
        $lblConn.Text = "Status: Connected to Exchange Online (Tenant: $($Script:TenantName)) | Mailboxes: $($Script:MailboxCount)"
    } else {
        $lblConn.Text = "Status: Not connected"
        $grid.DataSource = $null
        $txtDetails.Clear()
        $Script:SelectedIdentity = $null
    }
}

function Show-Details {
    param([string]$Identity)

    $details = Get-MailboxDetails -Identity $Identity
    $mbx = $details.Mailbox
    $usr = $details.User

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Mailbox:")
    $lines.Add("  DisplayName              : $($mbx.DisplayName)")
    $lines.Add("  PrimarySmtpAddress       : $($mbx.PrimarySmtpAddress)")
    $lines.Add("  RecipientTypeDetails     : $($mbx.RecipientTypeDetails)")
    $lines.Add("  IsDirSynced              : $($mbx.IsDirSynced)")
    $lines.Add("  IsExchangeCloudManaged   : $($mbx.IsExchangeCloudManaged)")
    $lines.Add("  ExchangeGuid             : $($mbx.ExchangeGuid)")
    $lines.Add("  ExternalDirectoryObjectId: $($mbx.ExternalDirectoryObjectId)")

    if ($usr) {
        $lines.Add("")
        $lines.Add("User:")
        $lines.Add("  UserPrincipalName        : $($usr.UserPrincipalName)")
        $lines.Add("  ImmutableId              : $($usr.ImmutableId)")
        $lines.Add("  WhenChangedUTC           : $($usr.WhenChangedUTC)")
    }

    $txtDetails.Lines = $lines.ToArray()

    $btnRefresh.Enabled      = $true
    $btnEnableCloud.Enabled  = $true
    $btnRevertOnPrem.Enabled = $true
}

# Events
$btnOpenLog.Add_Click({
    try {
        if (-not (Test-Path $Script:LogFile)) {
            New-Item -ItemType File -Path $Script:LogFile -Force | Out-Null
        }
        Start-Process -FilePath $Script:LogFile | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to open log file.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

$btnConnect.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        if (Connect-EXO) {
            Set-UiConnectedState -Connected $true
        }
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnDisconnect.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Disconnect-EXO
        Set-UiConnectedState -Connected $false
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnSearch.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $q = $txtSearch.Text
        Write-Log ("Search query: {0}" -f $q) "INFO"
        $results = Search-Mailboxes -QueryText $q -Max 200

        $lblCount.Text = "Results: $($results.Count)"

        if ($results.Count -eq 0) {
            $grid.DataSource = $null
            $txtDetails.Text = "No results."
            $Script:SelectedIdentity = $null
            $btnRefresh.Enabled = $false
            $btnEnableCloud.Enabled = $false
            $btnRevertOnPrem.Enabled= $false
            return
        }

        $grid.DataSource = $results
        $txtDetails.Text = "Select a mailbox row to see details."
    } catch {
        Write-Log ("Search failed: {0}" -f $_.Exception.Message) "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Search failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$grid.Add_SelectionChanged({
    try {
        if ($grid.SelectedRows.Count -gt 0) {
            $row = $grid.SelectedRows[0]
            # We renamed grid column from PrimarySmtpAddress to PrimarySMTP
            $smtp = $row.Cells["PrimarySMTP"].Value
            if ($smtp) {
                $Script:SelectedIdentity = $smtp.ToString()
                Show-Details -Identity $Script:SelectedIdentity
            }
        }
    } catch {
        Write-Log ("SelectionChanged warning: {0}" -f $_.Exception.Message) "WARN"
    }
})

$btnRefresh.Add_Click({
    if (-not $Script:SelectedIdentity) { return }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Show-Details -Identity $Script:SelectedIdentity
    } catch {
        Write-Log ("Refresh failed: {0}" -f $_.Exception.Message) "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Refresh failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnEnableCloud.Add_Click({
    if (-not $Script:SelectedIdentity) { return }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Enable cloud SOA for Exchange attributes for:`n`n$($Script:SelectedIdentity)`n`nThis sets IsExchangeCloudManaged = TRUE.",
        "Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $true
        [System.Windows.Forms.MessageBox]::Show(
            $msg,
            "Done",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        # Refresh details
        Show-Details -Identity $Script:SelectedIdentity

        # Re-run the current search so grid updates SOA Status
        if (-not [string]::IsNullOrWhiteSpace($txtSearch.Text)) {
            $btnSearch.PerformClick()
        }
    } catch {
        Write-Log ("Enable cloud SOA failed: {0}" -f $_.Exception.Message) "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Enable cloud SOA failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnRevertOnPrem.Add_Click({
    if (-not $Script:SelectedIdentity) { return }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Revert SOA back to on-prem for:`n`n$($Script:SelectedIdentity)`n`nThis sets IsExchangeCloudManaged = FALSE.`n`nWARNING: Next sync may overwrite cloud values with on-prem values.",
        "Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $false
        [System.Windows.Forms.MessageBox]::Show(
            $msg,
            "Done",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        # Refresh details
        Show-Details -Identity $Script:SelectedIdentity

        # Re-run the current search so grid updates SOA Status
        if (-not [string]::IsNullOrWhiteSpace($txtSearch.Text)) {
            $btnSearch.PerformClick()
        }
    } catch {
        Write-Log ("Revert to on-prem SOA failed: {0}" -f $_.Exception.Message) "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Revert failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$form.Add_FormClosing({
    try { Disconnect-EXO } catch {}
})

# Initialize UI
Set-UiConnectedState -Connected $false
Write-Log ("{0} v{1} started." -f $Script:ToolName, $Script:ScriptVersion) "INFO"

[System.Windows.Forms.Application]::Run($form)
#endregion
