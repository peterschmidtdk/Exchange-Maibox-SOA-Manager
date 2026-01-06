<#
.SYNOPSIS
  SOA Mailbox Tool (GUI) for Exchange Online - switch Exchange attribute SOA (IsExchangeCloudManaged).

.DESCRIPTION
  GUI tool to view and change mailbox Exchange attribute SOA state via IsExchangeCloudManaged:
    - Enable cloud management     : IsExchangeCloudManaged = $true
    - Revert to on-prem management: IsExchangeCloudManaged = $false

  Grid shows ONLY:
    - DisplayName
    - PrimarySMTP
    - SOA Status
    - DirSynced

  Also shows tenant name + mailbox count after connect.
  Includes "Open log" button.

REFERENCE
  https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

LOGGING
  - Single logfile (append-only)
  - Timestamp for each line

REQUIREMENTS
  - Windows PowerShell 5.1 or PowerShell 7+, must run in STA
  - ExchangeOnlineManagement module

AUTHOR
  Peter Schmidt

VERSION
  2.5.1 (2026-01-05)
    - Removed results/details panels; only GridView + actions
    - Grid columns: DisplayName, PrimarySMTP, SOA Status, DirSynced
    - Show tenant name + mailbox count after connect
    - Removed Export JSON
    - Added Open log button
#>

#region STA check
try {
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
        Write-Warning "This GUI must run in STA mode."
        Write-Warning "Run: powershell.exe -STA -File .\SOA-MailboxTool-GUI.ps1"
        Write-Warning "  or: pwsh.exe -STA -File .\SOA-MailboxTool-GUI.ps1"
        return
    }
} catch { }
#endregion

#region WinForms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region Globals
$Script:ToolName      = "SOA Mailbox Tool"
$Script:ScriptVersion = "2.5.1"

$Script:LogDir   = Join-Path -Path (Get-Location) -ChildPath "Logs"
$Script:LogFile  = Join-Path -Path $Script:LogDir -ChildPath "SOA-MailboxTool.log"

$Script:IsConnected  = $false
$Script:TenantName   = "Unknown"
$Script:MailboxCount = 0

$Script:LastQuery = ""

New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null
#endregion

#region Logging
function Write-Log {
    param(
        [Parameter(Mandatory)][ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level,
        [Parameter(Mandatory)][string]$Message
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $Script:LogFile -Value ("[{0}][{1}] {2}" -f $ts, $Level, $Message) -Encoding UTF8
}
#endregion

#region Module helper
function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            ("Required module '{0}' is not installed.`nInstall now (CurrentUser)?" -f $Name),
            "Missing Module",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
            throw ("Module '{0}' not installed." -f $Name)
        }

        Write-Log "INFO" ("Installing module '{0}' (CurrentUser)..." -f $Name)
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Log "INFO" ("Installed module '{0}'" -f $Name)
    }

    Import-Module $Name -ErrorAction Stop
    Write-Log "INFO" ("Imported module '{0}'" -f $Name)
}
#endregion

#region EXO connect/disconnect + tenant info
function Get-TenantNameBestEffort {
    try {
        $org = Get-OrganizationConfig -ErrorAction Stop
        if ($org -and $org.Name) { return [string]$org.Name }
    } catch { }
    return "Unknown"
}

function Get-MailboxCountBestEffort {
    try {
        $exoCmd = Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue
        if ($exoCmd) {
            return [int]((Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop | Measure-Object).Count)
        }
    } catch {
        Write-Log "WARN" ("Mailbox count via Get-EXOMailbox failed: {0}" -f $_.Exception.Message)
    }

    try {
        return [int]((Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | Measure-Object).Count)
    } catch {
        Write-Log "WARN" ("Mailbox count via Get-Mailbox failed: {0}" -f $_.Exception.Message)
    }

    return 0
}

function Connect-EXO {
    try {
        Ensure-Module -Name "ExchangeOnlineManagement"

        Write-Log "INFO" "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null

        $Script:IsConnected  = $true
        $Script:TenantName   = Get-TenantNameBestEffort
        $Script:MailboxCount = Get-MailboxCountBestEffort

        Write-Log "INFO" ("Connected. Tenant='{0}' Mailboxes={1}" -f $Script:TenantName, $Script:MailboxCount)
        return $true
    } catch {
        $Script:IsConnected = $false
        Write-Log "ERROR" ("Connect failed: {0}" -f $_.Exception.Message)
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
            Write-Log "INFO" "Disconnecting..."
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
    } catch {
        Write-Log "WARN" ("Disconnect warning: {0}" -f $_.Exception.Message)
    } finally {
        $Script:IsConnected  = $false
        $Script:TenantName   = "Unknown"
        $Script:MailboxCount = 0
        Write-Log "INFO" "Disconnected."
    }
}
#endregion

#region SOA helpers + mailbox search
function Get-SOAStatus {
    param($m)

    if ($m.IsDirSynced -ne $true) { return "N/A" }
    if ($m.IsExchangeCloudManaged -eq $true)  { return "Online" }
    if ($m.IsExchangeCloudManaged -eq $false) { return "On-Prem" }
    return "Unknown"
}

function Search-MailboxesForGrid {
    param(
        [Parameter(Mandatory)][string]$QueryText,
        [int]$Max = 500
    )
    if (-not $Script:IsConnected) { throw "Not connected." }

    $q = ($QueryText ?? "").Trim()
    if ([string]::IsNullOrWhiteSpace($q)) { return @() }

    # OPATH filter (best effort, faster than client-side)
    $filter = "DisplayName -like '*$q*' -or Alias -like '*$q*' -or PrimarySmtpAddress -like '*$q*'"

    $items = Get-Mailbox -ResultSize $Max -Filter $filter -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

    $rows = foreach ($m in $items) {
        [PSCustomObject]@{
            DisplayName  = [string]$m.DisplayName
            PrimarySMTP  = [string]$m.PrimarySmtpAddress
            'SOA Status' = (Get-SOAStatus $m)
            DirSynced    = if ($m.IsDirSynced -eq $true) { "Yes" } else { "No" }
        }
    }

    return @($rows)
}

function Set-MailboxSOA {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][bool]$EnableCloudManaged
    )
    if (-not $Script:IsConnected) { throw "Not connected." }

    $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

    if ($mbx.IsDirSynced -ne $true) {
        throw ("Mailbox '{0}' is not DirSynced. SOA switch is intended for DirSynced users." -f $Identity)
    }

    if ($mbx.IsExchangeCloudManaged -eq $EnableCloudManaged) {
        return ("No change needed. IsExchangeCloudManaged already '{0}'." -f $EnableCloudManaged)
    }

    Set-Mailbox -Identity $Identity -IsExchangeCloudManaged $EnableCloudManaged -ErrorAction Stop
    Write-Log "INFO" ("Set-Mailbox '{0}' IsExchangeCloudManaged={1}" -f $Identity, $EnableCloudManaged)
    return "Updated."
}
#endregion

#region GUI layout
$form = New-Object System.Windows.Forms.Form
$form.Text = "{0} v{1}" -f $Script:ToolName, $Script:ScriptVersion
$form.Size = New-Object System.Drawing.Size(1000, 650)
$form.StartPosition = "CenterScreen"

# Top controls
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

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Status: Not connected"
$lblStatus.Location = New-Object System.Drawing.Point(370, 18)
$lblStatus.AutoSize = $true

# Search controls
$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(12, 55)
$txtSearch.Size = New-Object System.Drawing.Size(740, 25)
$txtSearch.Enabled = $false

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Search"
$btnSearch.Location = New-Object System.Drawing.Point(760, 52)
$btnSearch.Size = New-Object System.Drawing.Size(100, 30)
$btnSearch.Enabled = $false

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(868, 52)
$btnClear.Size = New-Object System.Drawing.Size(100, 30)
$btnClear.Enabled = $false

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "Mailboxes shown: 0"
$lblCount.Location = New-Object System.Drawing.Point(12, 86)
$lblCount.AutoSize = $true

# GridView (DataGridView)
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 110)
$grid.Size = New-Object System.Drawing.Size(956, 420)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $true
$grid.AutoSizeColumnsMode = "Fill"
$grid.AutoGenerateColumns = $true

# Action buttons (operate on selected rows)
$btnEnable = New-Object System.Windows.Forms.Button
$btnEnable.Text = "Enable cloud SOA (Online)"
$btnEnable.Location = New-Object System.Drawing.Point(12, 545)
$btnEnable.Size = New-Object System.Drawing.Size(230, 35)
$btnEnable.Enabled = $false

$btnRevert = New-Object System.Windows.Forms.Button
$btnRevert.Text = "Revert to on-prem SOA"
$btnRevert.Location = New-Object System.Drawing.Point(250, 545)
$btnRevert.Size = New-Object System.Drawing.Size(230, 35)
$btnRevert.Enabled = $false

$form.Controls.AddRange(@(
    $btnConnect,$btnDisconnect,$btnOpenLog,$lblStatus,
    $txtSearch,$btnSearch,$btnClear,$lblCount,
    $grid,$btnEnable,$btnRevert
))
#endregion

#region UI helpers
function Set-ConnectedUI {
    param([bool]$Connected)

    $btnConnect.Enabled    = -not $Connected
    $btnDisconnect.Enabled = $Connected
    $txtSearch.Enabled     = $Connected
    $btnSearch.Enabled     = $Connected
    $btnClear.Enabled      = $Connected

    $btnEnable.Enabled = $false
    $btnRevert.Enabled = $false

    if ($Connected) {
        $lblStatus.Text = "Status: Connected (Tenant: $($Script:TenantName)) | Mailboxes in tenant: $($Script:MailboxCount)"
    } else {
        $lblStatus.Text = "Status: Not connected"
        $grid.DataSource = $null
        $lblCount.Text = "Mailboxes shown: 0"
        $Script:LastQuery = ""
    }
}

function Get-SelectedSMTPs {
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($r in $grid.SelectedRows) {
        try {
            $smtp = $r.Cells["PrimarySMTP"].Value
            if ($smtp) { [void]$list.Add($smtp.ToString()) }
        } catch { }
    }
    return @($list | Select-Object -Unique)
}

function Refresh-GridFromLastQuery {
    if ([string]::IsNullOrWhiteSpace($Script:LastQuery)) { return }
    $rows = Search-MailboxesForGrid -QueryText $Script:LastQuery -Max 500
    $grid.DataSource = $rows
    $lblCount.Text = "Mailboxes shown: $($rows.Count)"
}
#endregion

#region Events
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
            Set-ConnectedUI -Connected $true
        }
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnDisconnect.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Disconnect-EXO
        Set-ConnectedUI -Connected $false
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnSearch.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $q = ($txtSearch.Text ?? "").Trim()
        $Script:LastQuery = $q
        Write-Log "INFO" ("Search query: {0}" -f $q)

        $rows = Search-MailboxesForGrid -QueryText $q -Max 500
        $grid.DataSource = $rows
        $lblCount.Text = "Mailboxes shown: $($rows.Count)"

        $btnEnable.Enabled = ($rows.Count -gt 0)
        $btnRevert.Enabled = ($rows.Count -gt 0)
    } catch {
        Write-Log "ERROR" ("Search failed: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            "Search failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnClear.Add_Click({
    $txtSearch.Text = ""
    $Script:LastQuery = ""
    $grid.DataSource = $null
    $lblCount.Text = "Mailboxes shown: 0"
    $btnEnable.Enabled = $false
    $btnRevert.Enabled = $false
})

$grid.Add_SelectionChanged({
    # Enable action buttons only when at least one row is selected
    try {
        $hasSel = ($grid.SelectedRows.Count -gt 0)
        $btnEnable.Enabled = $hasSel
        $btnRevert.Enabled = $hasSel
    } catch { }
})

$btnEnable.Add_Click({
    try {
        $targets = Get-SelectedSMTPs
        if ($targets.Count -eq 0) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            ("Enable cloud SOA for {0} mailbox(es)?`n`nThis sets IsExchangeCloudManaged = TRUE." -f $targets.Count),
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        foreach ($smtp in $targets) {
            try {
                Write-Log "INFO" ("Enable cloud SOA: {0}" -f $smtp)
                [void](Set-MailboxSOA -Identity $smtp -EnableCloudManaged $true)
            } catch {
                Write-Log "ERROR" ("Enable cloud SOA failed for {0}: {1}" -f $smtp, $_.Exception.Message)
            }
        }

        Refresh-GridFromLastQuery
        [System.Windows.Forms.MessageBox]::Show(
            "Completed. Check log for details.",
            "Done",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnRevert.Add_Click({
    try {
        $targets = Get-SelectedSMTPs
        if ($targets.Count -eq 0) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            ("Revert SOA to on-prem for {0} mailbox(es)?`n`nThis sets IsExchangeCloudManaged = FALSE." -f $targets.Count),
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        foreach ($smtp in $targets) {
            try {
                Write-Log "INFO" ("Revert to on-prem SOA: {0}" -f $smtp)
                [void](Set-MailboxSOA -Identity $smtp -EnableCloudManaged $false)
            } catch {
                Write-Log "ERROR" ("Revert SOA failed for {0}: {1}" -f $smtp, $_.Exception.Message)
            }
        }

        Refresh-GridFromLastQuery
        [System.Windows.Forms.MessageBox]::Show(
            "Completed. Check log for details.",
            "Done",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$form.Add_FormClosing({
    try { Disconnect-EXO } catch { }
})

#endregion

# Init
Write-Log "INFO" ("{0} v{1} started." -f $Script:ToolName, $Script:ScriptVersion)
Set-ConnectedUI -Connected $false

[System.Windows.Forms.Application]::Run($form)
