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

  Adds:
    - Connect/Disconnect
    - Load all mailboxes into cache (browse/select)
    - Search/filter against cache (fast)
    - Show tenant name after connect
    - Show mailbox count in tenant
    - Open log file button

REFERENCE
  https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

LOGGING
  - Single logfile (append-only)
  - Timestamp for each log line

REQUIREMENTS
  - Windows PowerShell 5.1 or PowerShell 7+, must run in STA for WinForms
  - ExchangeOnlineManagement module

AUTHOR
  Peter Schmidt

VERSION
  2.5.2 (2026-01-05)
#>

#region STA check
try {
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
        Write-Warning "This GUI must run in STA mode."
        Write-Warning "Run:"
        Write-Warning "  powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\SOA-MailboxTool-GUI.ps1"
        Write-Warning "  pwsh.exe      -NoProfile -ExecutionPolicy Bypass -STA -File .\SOA-MailboxTool-GUI.ps1"
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
$Script:ScriptVersion = "2.5.2"

$Script:LogDir   = Join-Path -Path (Get-Location) -ChildPath "Logs"
$Script:LogFile  = Join-Path -Path $Script:LogDir -ChildPath "SOA-MailboxTool.log"

$Script:IsConnected  = $false
$Script:TenantName   = "Unknown"
$Script:MailboxCount = 0

$Script:MailboxCache = @()  # full cache objects
$Script:ViewRows     = @()  # current filtered view objects

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
    # Best effort count. May take time in large tenants.
    try {
        $cmd = Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue
        if ($cmd) {
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

#region SOA helpers
function Get-SOAStatusFromValues {
    param([bool]$IsDirSynced, $IsExchangeCloudManaged)

    if ($IsDirSynced -ne $true) { return "N/A" }
    if ($IsExchangeCloudManaged -eq $true)  { return "Online" }
    if ($IsExchangeCloudManaged -eq $false) { return "On-Prem" }
    return "Unknown"
}
#endregion

#region Mailbox cache/load/search + actions
function Load-AllMailboxesToCache {
    if (-not $Script:IsConnected) { throw "Not connected." }

    $cmdExo = Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue

    $raw = @()
    if ($cmdExo) {
        # Use Get-EXOMailbox when available (faster, modern)
        try {
            # Try with -Properties first (supported in modern module versions)
            $raw = @(Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged -ErrorAction Stop)
        } catch {
            # Fallback without -Properties
            $raw = @(Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop)
        }
    } else {
        # Fallback legacy
        $raw = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)
    }

    $cache = New-Object System.Collections.Generic.List[object]

    foreach ($m in $raw) {
        if ($null -eq $m) { continue }

        # Safe access
        $dn   = [string]$m.DisplayName
        $smtp = [string]$m.PrimarySmtpAddress
        $dir  = $false
        try { $dir = [bool]$m.IsDirSynced } catch { $dir = $false }

        $cloud = $null
        try { $cloud = $m.IsExchangeCloudManaged } catch { $cloud = $null }

        $soa = Get-SOAStatusFromValues -IsDirSynced $dir -IsExchangeCloudManaged $cloud

        # Keep extra fields for actions/refresh
        $obj = [pscustomobject]@{
            DisplayName          = $dn
            PrimarySMTP          = $smtp
            'SOA Status'         = $soa
            DirSynced            = if ($dir) { "Yes" } else { "No" }

            # hidden/useful fields
            _IsDirSynced         = $dir
            _IsExchangeCloudManaged = $cloud
            _Identity            = if ($smtp) { $smtp } else { [string]$m.Identity }
        }
        [void]$cache.Add($obj)
    }

    $Script:MailboxCache = @($cache)
    return $Script:MailboxCache
}

function Apply-FilterToCache {
    param([string]$FilterText)

    $q = ""
    if ($FilterText) { $q = $FilterText.Trim() }

    if ([string]::IsNullOrWhiteSpace($q)) {
        $Script:ViewRows = $Script:MailboxCache
        return $Script:ViewRows
    }

    $lower = $q.ToLowerInvariant()
    $Script:ViewRows = @(
        $Script:MailboxCache | Where-Object {
            ($_.DisplayName  -and $_.DisplayName.ToLowerInvariant().Contains($lower)) -or
            ($_.PrimarySMTP  -and $_.PrimarySMTP.ToLowerInvariant().Contains($lower)) -or
            ($_.('SOA Status') -and $_.('SOA Status').ToLowerInvariant().Contains($lower)) -or
            ($_.DirSynced    -and $_.DirSynced.ToLowerInvariant().Contains($lower))
        }
    )

    return $Script:ViewRows
}

function Get-SelectedIdentitiesFromGrid {
    param([System.Windows.Forms.DataGridView]$Grid)

    $list = New-Object System.Collections.Generic.List[string]
    foreach ($r in $Grid.SelectedRows) {
        try {
            # We bind to objects with _Identity
            $obj = $r.DataBoundItem
            if ($obj -and $obj._Identity) { [void]$list.Add([string]$obj._Identity) }
        } catch { }
    }
    return @($list | Select-Object -Unique)
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

#region GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = "{0} v{1}" -f $Script:ToolName, $Script:ScriptVersion
$form.Size = New-Object System.Drawing.Size(1100, 720)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)

# Layout container
$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = "Fill"
$layout.RowCount = 4
$layout.ColumnCount = 1
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 48)))  | Out-Null  # top bar
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46)))  | Out-Null  # search/load bar
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  | Out-Null  # grid
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 56)))  | Out-Null  # actions
$form.Controls.Add($layout)

# --- Top bar
$topBar = New-Object System.Windows.Forms.Panel
$topBar.Dock = "Fill"
$topBar.BackColor = [System.Drawing.Color]::White

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect"
$btnConnect.Size = New-Object System.Drawing.Size(110, 30)
$btnConnect.Location = New-Object System.Drawing.Point(10, 9)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Text = "Disconnect"
$btnDisconnect.Size = New-Object System.Drawing.Size(110, 30)
$btnDisconnect.Location = New-Object System.Drawing.Point(128, 9)
$btnDisconnect.Enabled = $false

$btnOpenLog = New-Object System.Windows.Forms.Button
$btnOpenLog.Text = "Open log"
$btnOpenLog.Size = New-Object System.Drawing.Size(110, 30)
$btnOpenLog.Location = New-Object System.Drawing.Point(246, 9)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Status: Not connected"
$lblStatus.AutoSize = $true
$lblStatus.Location = New-Object System.Drawing.Point(370, 14)

$topBar.Controls.AddRange(@($btnConnect,$btnDisconnect,$btnOpenLog,$lblStatus))
$layout.Controls.Add($topBar, 0, 0)

# --- Search/load bar
$searchBar = New-Object System.Windows.Forms.Panel
$searchBar.Dock = "Fill"
$searchBar.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)

$btnLoadAll = New-Object System.Windows.Forms.Button
$btnLoadAll.Text = "Load all mailboxes"
$btnLoadAll.Size = New-Object System.Drawing.Size(160, 30)
$btnLoadAll.Location = New-Object System.Drawing.Point(10, 8)
$btnLoadAll.Enabled = $false

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Size = New-Object System.Drawing.Size(520, 25)
$txtFilter.Location = New-Object System.Drawing.Point(182, 10)
$txtFilter.Enabled = $false

$btnApplyFilter = New-Object System.Windows.Forms.Button
$btnApplyFilter.Text = "Search"
$btnApplyFilter.Size = New-Object System.Drawing.Size(90, 30)
$btnApplyFilter.Location = New-Object System.Drawing.Point(710, 8)
$btnApplyFilter.Enabled = $false

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Size = New-Object System.Drawing.Size(90, 30)
$btnClear.Location = New-Object System.Drawing.Point(808, 8)
$btnClear.Enabled = $false

$lblCounts = New-Object System.Windows.Forms.Label
$lblCounts.Text = "Cached: 0 | Showing: 0"
$lblCounts.AutoSize = $true
$lblCounts.Location = New-Object System.Drawing.Point(910, 14)

$searchBar.Controls.AddRange(@($btnLoadAll,$txtFilter,$btnApplyFilter,$btnClear,$lblCounts))
$layout.Controls.Add($searchBar, 0, 1)

# --- Grid (DataGridView) (this is the real gridview)
$gridPanel = New-Object System.Windows.Forms.Panel
$gridPanel.Dock = "Fill"
$gridPanel.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = "FixedSingle"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $true
$grid.AutoGenerateColumns = $false
$grid.RowHeadersVisible = $false
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(243,243,243)
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# Define columns explicitly (prevents the "grey box with nothing useful")
function Add-GridColumn([string]$Header, [string]$Prop, [int]$Width = 200) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $Header
    $col.DataPropertyName = $Prop
    $col.Width = $Width
    $col.AutoSizeMode = "Fill"
    return $col
}
[void]$grid.Columns.Add((Add-GridColumn "DisplayName"  "DisplayName"  300))
[void]$grid.Columns.Add((Add-GridColumn "PrimarySMTP"  "PrimarySMTP"  260))
[void]$grid.Columns.Add((Add-GridColumn "SOA Status"   "SOA Status"   140))
[void]$grid.Columns.Add((Add-GridColumn "DirSynced"    "DirSynced"    100))

$gridPanel.Controls.Add($grid)
$layout.Controls.Add($gridPanel, 0, 2)

# --- Actions bar
$actionBar = New-Object System.Windows.Forms.Panel
$actionBar.Dock = "Fill"
$actionBar.BackColor = [System.Drawing.Color]::White

$btnEnable = New-Object System.Windows.Forms.Button
$btnEnable.Text = "Enable SOA -> Online"
$btnEnable.Size = New-Object System.Drawing.Size(190, 34)
$btnEnable.Location = New-Object System.Drawing.Point(10, 10)
$btnEnable.Enabled = $false

$btnRevert = New-Object System.Windows.Forms.Button
$btnRevert.Text = "Revert SOA -> On-Prem"
$btnRevert.Size = New-Object System.Drawing.Size(200, 34)
$btnRevert.Location = New-Object System.Drawing.Point(210, 10)
$btnRevert.Enabled = $false

$actionBar.Controls.AddRange(@($btnEnable,$btnRevert))
$layout.Controls.Add($actionBar, 0, 3)

# Helpers
function Update-Counts {
    $cached = 0
    $showing = 0
    if ($Script:MailboxCache) { $cached = $Script:MailboxCache.Count }
    if ($Script:ViewRows) { $showing = $Script:ViewRows.Count }
    $lblCounts.Text = ("Cached: {0} | Showing: {1}" -f $cached, $showing)
}

function Set-ConnectedUI {
    param([bool]$Connected)

    $btnConnect.Enabled    = -not $Connected
    $btnDisconnect.Enabled = $Connected

    $btnLoadAll.Enabled    = $Connected
    $txtFilter.Enabled     = $Connected
    $btnApplyFilter.Enabled= $Connected
    $btnClear.Enabled      = $Connected

    $btnEnable.Enabled     = $false
    $btnRevert.Enabled     = $false

    if ($Connected) {
        $lblStatus.Text = ("Status: Connected (Tenant: {0}) | Mailboxes in tenant: {1}" -f $Script:TenantName, $Script:MailboxCount)
    } else {
        $lblStatus.Text = "Status: Not connected"
        $grid.DataSource = $null
        $Script:MailboxCache = @()
        $Script:ViewRows = @()
        Update-Counts
    }
}

function Bind-Grid([object[]]$Rows) {
    $grid.DataSource = $null
    $grid.DataSource = $Rows
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

$btnLoadAll.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Write-Log "INFO" "Loading all mailboxes into cache..."
        $data = Load-AllMailboxesToCache
        Write-Log "INFO" ("Loaded cache: {0} mailboxes" -f $data.Count)

        # Default view = full cache
        $Script:ViewRows = $Script:MailboxCache
        Bind-Grid -Rows $Script:ViewRows
        Update-Counts
    } catch {
        Write-Log "ERROR" ("Load all mailboxes failed: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            "Load all mailboxes failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnApplyFilter.Add_Click({
    try {
        if (-not $Script:MailboxCache -or $Script:MailboxCache.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Load all mailboxes first.",
                "Info",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $q = $txtFilter.Text
        Write-Log "INFO" ("Filter applied: {0}" -f $q)

        $view = Apply-FilterToCache -FilterText $q
        Bind-Grid -Rows $view
        Update-Counts
    } catch {
        Write-Log "ERROR" ("Filter failed: {0}" -f $_.Exception.Message)
    }
})

$btnClear.Add_Click({
    $txtFilter.Text = ""
    if ($Script:MailboxCache) {
        $Script:ViewRows = $Script:MailboxCache
        Bind-Grid -Rows $Script:ViewRows
    } else {
        $grid.DataSource = $null
        $Script:ViewRows = @()
    }
    Update-Counts
})

$grid.Add_SelectionChanged({
    try {
        $hasSel = ($grid.SelectedRows.Count -gt 0)
        $btnEnable.Enabled = $hasSel
        $btnRevert.Enabled = $hasSel
    } catch { }
})

$btnEnable.Add_Click({
    try {
        $targets = Get-SelectedIdentitiesFromGrid -Grid $grid
        if ($targets.Count -eq 0) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            ("Enable SOA -> Online for {0} mailbox(es)?`nThis sets IsExchangeCloudManaged = TRUE." -f $targets.Count),
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        foreach ($id in $targets) {
            try {
                Write-Log "INFO" ("Enable SOA Online: {0}" -f $id)
                [void](Set-MailboxSOA -Identity $id -EnableCloudManaged $true)
            } catch {
                Write-Log "ERROR" ("Enable failed for {0}: {1}" -f $id, $_.Exception.Message)
            }
        }

        # Reload cache to reflect new status
        $data = Load-AllMailboxesToCache
        $Script:ViewRows = Apply-FilterToCache -FilterText $txtFilter.Text
        Bind-Grid -Rows $Script:ViewRows
        Update-Counts

        [System.Windows.Forms.MessageBox]::Show(
            "Completed. See log for details.",
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
        $targets = Get-SelectedIdentitiesFromGrid -Grid $grid
        if ($targets.Count -eq 0) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            ("Revert SOA -> On-Prem for {0} mailbox(es)?`nThis sets IsExchangeCloudManaged = FALSE." -f $targets.Count),
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        foreach ($id in $targets) {
            try {
                Write-Log "INFO" ("Revert SOA On-Prem: {0}" -f $id)
                [void](Set-MailboxSOA -Identity $id -EnableCloudManaged $false)
            } catch {
                Write-Log "ERROR" ("Revert failed for {0}: {1}" -f $id, $_.Exception.Message)
            }
        }

        # Reload cache to reflect new status
        $data = Load-AllMailboxesToCache
        $Script:ViewRows = Apply-FilterToCache -FilterText $txtFilter.Text
        Bind-Grid -Rows $Script:ViewRows
        Update-Counts

        [System.Windows.Forms.MessageBox]::Show(
            "Completed. See log for details.",
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

# Init
Write-Log "INFO" ("{0} v{1} started." -f $Script:ToolName, $Script:ScriptVersion)
Set-ConnectedUI -Connected $false
Update-Counts

[System.Windows.Forms.Application]::Run($form)
#endregion
