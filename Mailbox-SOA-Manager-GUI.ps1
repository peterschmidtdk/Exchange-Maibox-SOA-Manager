<#
.SYNOPSIS
  Mailbox SOA Manager (Exchange Online) - GUI tool to view and switch Exchange attribute SOA per mailbox.

.DESCRIPTION
  PowerShell 7+ only WinForms tool to:
    - Connect / Disconnect to Exchange Online
    - Load mailboxes (User + optional Shared)
    - Browse/search via local cache with paging
    - Show SOA indicator for Exchange Attributes management:
        ☁ Online  -> IsExchangeCloudManaged = True  (DirSynced mailboxes)
        🏢 On-Prem -> IsExchangeCloudManaged = False (DirSynced mailboxes)
        — N/A     -> Cloud-only / not DirSynced
    - Switch SOA:
        Set-Mailbox -IsExchangeCloudManaged $true/$false
    - Export current cached mailbox SOA settings to CSV
    - Log all actions to ONE logfile (timestamp per line)

.REFERENCE
  https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

.NOTES
  Name    : Mailbox SOA Manager
  Version : 1.9.0
  Date    : 2026-01-05
  Author  : Peter Schmidt

.CHANGELOG
  1.9.0 (2026-01-05)
    - PS7-only rewrite (simplified + optimized)
    - Fix: null-safe load/browse/search + stable grid binding
    - Adds: Tenant name in connected status (best-effort)
    - Keeps: single logfile with timestamp per line
#>

# =========================
# PS7+ requirement + STA relaunch
# =========================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        [System.Windows.Forms.MessageBox]::Show(
            "This tool requires PowerShell 7+.`n`nRun:`n  pwsh.exe -STA -File .\ExchangeSOAManager-GUI.ps1",
            "Mailbox SOA Manager",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } catch {}
    exit 1
}

function Get-ScriptPathSafe {
    if ($PSCommandPath) { return $PSCommandPath }
    if ($MyInvocation.MyCommand.Path) { return $MyInvocation.MyCommand.Path }
    return $null
}

function Ensure-STAOrRelaunch {
    try {
        $apt = [System.Threading.Thread]::CurrentThread.ApartmentState
        if ($apt -eq [System.Threading.ApartmentState]::STA) { return $true }

        $path = Get-ScriptPathSafe
        if (-not $path -or -not (Test-Path $path)) {
            Add-Type -AssemblyName System.Windows.Forms | Out-Null
            [System.Windows.Forms.MessageBox]::Show(
                "This GUI must run in STA mode.`n`nRun:`n  pwsh.exe -STA -File .\ExchangeSOAManager-GUI.ps1",
                "Mailbox SOA Manager",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return $false
        }

        $args = @("-NoProfile","-ExecutionPolicy","Bypass","-STA","-File", "`"$path`"") -join " "
        Start-Process -FilePath "pwsh.exe" -ArgumentList $args -WorkingDirectory (Split-Path -Parent $path) | Out-Null
        return $false
    } catch {
        return $false
    }
}

if (-not (Ensure-STAOrRelaunch)) { exit 0 }

# =========================
# Globals
# =========================
$Script:ToolName      = "Mailbox SOA Manager"
$Script:Version       = "1.9.0"
$Script:BuildDate     = "2026-01-05"
$Script:RunId         = [guid]::NewGuid().ToString()

$Script:Root          = (Get-Location).Path
$Script:LogDir        = Join-Path $Script:Root "Logs"
$Script:LogFile       = Join-Path $Script:LogDir "MailboxSOAManager.log"
$Script:ExportsDir    = Join-Path $Script:Root "Exports"

$Script:IsConnected   = $false
$Script:TenantName    = ""
$Script:ActorUpn      = ""

$Script:MailboxCache  = @()
$Script:View          = @()
$Script:PageSize      = 100
$Script:CurrentPage   = 1

# =========================
# Logging
# =========================
function Ensure-LogPath {
    if (-not (Test-Path $Script:LogDir)) { New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null }
    if (-not (Test-Path $Script:LogFile)) { New-Item -ItemType File -Path $Script:LogFile -Force | Out-Null }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level,
        [Parameter(Mandatory)][string]$Message
    )
    Ensure-LogPath
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tenant = if ($Script:TenantName) { $Script:TenantName } else { "Tenant:unknown" }
    $actor  = if ($Script:ActorUpn) { $Script:ActorUpn } else { "Actor:unknown" }
    $line = "[{0}][{1}][RunId:{2}][{3}][{4}] {5}" -f $ts, $Level, $Script:RunId, $tenant, $actor, $Message
    Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8
}

Write-Log "INFO" ("=== {0} v{1} ({2}) started ===" -f $Script:ToolName, $Script:Version, $Script:BuildDate)

# =========================
# Helpers
# =========================
function Show-ErrorBox {
    param([string]$Text, [string]$Title = "Error")
    [System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Show-InfoBox {
    param([string]$Text, [string]$Title = "Info")
    [System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)

    $found = Get-Module -ListAvailable -Name $Name | Select-Object -First 1
    if (-not $found) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            ("Required module '{0}' is not installed.`nInstall now for CurrentUser?" -f $Name),
            "Missing Module",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) {
            throw ("Module '{0}' is required." -f $Name)
        }

        Write-Log "INFO" ("Installing module: {0}" -f $Name)
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Log "INFO" ("Installed module: {0}" -f $Name)
    }

    Import-Module $Name -ErrorAction Stop
    Write-Log "INFO" ("Imported module: {0}" -f $Name)
}

function Get-SOAState {
    param([bool]$IsDirSynced, $IsExchangeCloudManaged)
    if (-not $IsDirSynced) { return "N/A" }
    if ($IsExchangeCloudManaged -eq $true)  { return "Online" }
    if ($IsExchangeCloudManaged -eq $false) { return "On-Prem" }
    return "Unknown"
}

function Get-SOAIndicator {
    param([string]$State)
    switch ($State) {
        "Online"  { "☁ Online" }
        "On-Prem" { "🏢 On-Prem" }
        default   { "— N/A" }
    }
}

function To-Row {
    param($m)

    # Robust string casts (no null method calls)
    $display = [string]$m.DisplayName
    $upn     = [string]$m.UserPrincipalName
    $smtp    = [string]$m.PrimarySmtpAddress
    $type    = [string]$m.RecipientTypeDetails
    $dir     = [bool]$m.IsDirSynced
    $cloud   = $m.IsExchangeCloudManaged

    $state = Get-SOAState -IsDirSynced $dir -IsExchangeCloudManaged $cloud

    [pscustomobject]@{
        Indicator              = (Get-SOAIndicator $state)
        SOAState               = $state
        DisplayName            = $display
        UserPrincipalName      = $upn
        PrimarySmtpAddress     = $smtp
        RecipientTypeDetails   = $type
        IsDirSynced            = $dir
        IsExchangeCloudManaged = $cloud
        Identity               = ([string]$m.Identity)
    }
}

function New-GridTable {
    $dt = New-Object System.Data.DataTable "Mailboxes"
    [void]$dt.Columns.Add("Indicator",[string])
    [void]$dt.Columns.Add("SOAState",[string])
    [void]$dt.Columns.Add("DisplayName",[string])
    [void]$dt.Columns.Add("UserPrincipalName",[string])
    [void]$dt.Columns.Add("PrimarySmtpAddress",[string])
    [void]$dt.Columns.Add("RecipientTypeDetails",[string])
    [void]$dt.Columns.Add("DirSynced",[string])
    return $dt
}

function Build-GridTable {
    param([array]$Items)

    $dt = New-GridTable
    foreach ($x in ($Items ?? @())) {
        if ($null -eq $x) { continue }
        $row = $dt.NewRow()
        $row["Indicator"]            = [string]$x.Indicator
        $row["SOAState"]             = [string]$x.SOAState
        $row["DisplayName"]          = [string]$x.DisplayName
        $row["UserPrincipalName"]    = [string]$x.UserPrincipalName
        $row["PrimarySmtpAddress"]   = [string]$x.PrimarySmtpAddress
        $row["RecipientTypeDetails"] = [string]$x.RecipientTypeDetails
        $row["DirSynced"]            = if ($x.IsDirSynced) { "Yes" } else { "No" }
        [void]$dt.Rows.Add($row)
    }
    return $dt
}

# =========================
# Exchange Online
# =========================
function Connect-EXO {
    try {
        Ensure-Module -Name "ExchangeOnlineManagement"

        Write-Log "INFO" "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null

        $Script:IsConnected = $true
        $Script:ActorUpn = ""
        $Script:TenantName = ""

        try {
            $ci = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
            if ($ci -and $ci.UserPrincipalName) { $Script:ActorUpn = [string]$ci.UserPrincipalName }
            if ($ci -and $ci.TenantId -and -not $Script:TenantName) { $Script:TenantName = [string]$ci.TenantId }
        } catch { }

        try {
            $org = Get-OrganizationConfig -ErrorAction Stop
            if ($org -and $org.Name) { $Script:TenantName = [string]$org.Name }
        } catch { }

        Write-Log "INFO" ("Connected. Tenant='{0}' Actor='{1}'" -f $Script:TenantName, $Script:ActorUpn)
        return $true
    } catch {
        Write-Log "ERROR" ("Connect failed: {0}" -f $_.Exception.Message)
        Show-ErrorBox -Text ("Connect failed:`n{0}" -f $_.Exception.Message) -Title $Script:ToolName
        $Script:IsConnected = $false
        $Script:TenantName = ""
        $Script:ActorUpn = ""
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
        $Script:IsConnected = $false
        $Script:TenantName = ""
        $Script:ActorUpn = ""
        Write-Log "INFO" "Disconnected."
    }
}

function Get-Mailboxes {
    param([bool]$IncludeShared)

    $cmd = Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue
    if (-not $cmd) { throw "Get-EXOMailbox is not available." }

    $props = @("DisplayName","UserPrincipalName","PrimarySmtpAddress","RecipientTypeDetails","IsDirSynced","IsExchangeCloudManaged","Identity")
    $types = @("UserMailbox")
    if ($IncludeShared) { $types += "SharedMailbox" }

    $all = New-Object System.Collections.Generic.List[object]

    foreach ($t in $types) {
        $splat = @{
            ResultSize            = "Unlimited"
            RecipientTypeDetails  = $t
            ErrorAction           = "Stop"
            PageSize              = 1000
        }

        if ($cmd.Parameters.ContainsKey("Properties")) {
            $splat["Properties"] = $props
        }

        $chunk = @(Get-EXOMailbox @splat)
        foreach ($m in $chunk) { if ($null -ne $m) { [void]$all.Add($m) } }
    }

    return @($all)
}

# =========================
# GUI
# =========================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form -Property @{
    Text = ("{0} v{1}" -f $Script:ToolName, $Script:Version)
    Size = New-Object System.Drawing.Size(1250, 900)
    MinimumSize = New-Object System.Drawing.Size(980, 650)
    StartPosition = "CenterScreen"
    BackColor = [System.Drawing.Color]::FromArgb(245,248,250)
    KeyPreview = $true
}

# Header
$header = New-Object System.Windows.Forms.Panel -Property @{
    Dock="Top"; Height=74; BackColor=[System.Drawing.Color]::White
}
$accent = New-Object System.Windows.Forms.Label -Property @{
    Dock="Bottom"; Height=2; BackColor=[System.Drawing.Color]::FromArgb(0,120,212)
}
$header.Controls.Add($accent)

$lblTitle = New-Object System.Windows.Forms.Label -Property @{
    Text=$Script:ToolName
    Location=New-Object System.Drawing.Point(18,18)
    Size=New-Object System.Drawing.Size(520,35)
    Font=New-Object System.Drawing.Font("Segoe UI",20,[System.Drawing.FontStyle]::Bold)
    ForeColor=[System.Drawing.Color]::FromArgb(0,78,146)
}
$header.Controls.Add($lblTitle)

$lblVer = New-Object System.Windows.Forms.Label -Property @{
    Text=("v{0} ({1})" -f $Script:Version, $Script:BuildDate)
    Location=New-Object System.Drawing.Point(540,28)
    Size=New-Object System.Drawing.Size(260,20)
    Font=New-Object System.Drawing.Font("Segoe UI",9)
    ForeColor=[System.Drawing.Color]::FromArgb(96,94,92)
}
$header.Controls.Add($lblVer)

$btnConnect = New-Object System.Windows.Forms.Button -Property @{
    Text="Connect"
    Size=New-Object System.Drawing.Size(120,35)
    Location=New-Object System.Drawing.Point(860,18)
    BackColor=[System.Drawing.Color]::White
    ForeColor=[System.Drawing.Color]::FromArgb(0,120,212)
    FlatStyle="Flat"
    Font=New-Object System.Drawing.Font("Segoe UI",10)
    Anchor="Top,Right"
}
$btnConnect.FlatAppearance.BorderSize = 1
$btnConnect.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0,120,212)
$header.Controls.Add($btnConnect)

$btnDisconnect = New-Object System.Windows.Forms.Button -Property @{
    Text="Disconnect"
    Size=New-Object System.Drawing.Size(120,35)
    Location=New-Object System.Drawing.Point(990,18)
    BackColor=[System.Drawing.Color]::White
    ForeColor=[System.Drawing.Color]::FromArgb(120,120,120)
    FlatStyle="Flat"
    Font=New-Object System.Drawing.Font("Segoe UI",10)
    Anchor="Top,Right"
    Enabled=$false
}
$btnDisconnect.FlatAppearance.BorderSize = 1
$btnDisconnect.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200,198,196)
$header.Controls.Add($btnDisconnect)

$form.Controls.Add($header)

# Control panel
$controls = New-Object System.Windows.Forms.Panel -Property @{
    Dock="Top"; Height=90; BackColor=[System.Drawing.Color]::White
}
$sep = New-Object System.Windows.Forms.Label -Property @{
    Dock="Bottom"; Height=1; BackColor=[System.Drawing.Color]::FromArgb(229,229,229)
}
$controls.Controls.Add($sep)

$lblSearch = New-Object System.Windows.Forms.Label -Property @{
    Text="Search:"
    Location=New-Object System.Drawing.Point(18,16)
    Size=New-Object System.Drawing.Size(60,20)
    Font=New-Object System.Drawing.Font("Segoe UI",10)
}
$controls.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox -Property @{
    Location=New-Object System.Drawing.Point(80,14)
    Size=New-Object System.Drawing.Size(360,25)
    Font=New-Object System.Drawing.Font("Segoe UI",10)
    Enabled=$false
}
$controls.Controls.Add($txtSearch)

$btnClear = New-Object System.Windows.Forms.Button -Property @{
    Text="Clear"
    Location=New-Object System.Drawing.Point(450,12)
    Size=New-Object System.Drawing.Size(80,30)
    BackColor=[System.Drawing.Color]::White
    FlatStyle="Flat"
    Font=New-Object System.Drawing.Font("Segoe UI",9)
    Enabled=$false
}
$btnClear.FlatAppearance.BorderSize = 1
$btnClear.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200,198,196)
$controls.Controls.Add($btnClear)

$chkDirSyncedOnly = New-Object System.Windows.Forms.CheckBox -Property @{
    Text="DirSynced only"
    Location=New-Object System.Drawing.Point(80,52)
    Size=New-Object System.Drawing.Size(140,22)
    Font=New-Object System.Drawing.Font("Segoe UI",9)
    Checked=$true
    Enabled=$false
}
$controls.Controls.Add($chkDirSyncedOnly)

$chkIncludeShared = New-Object System.Windows.Forms.CheckBox -Property @{
    Text="Include Shared"
    Location=New-Object System.Drawing.Point(230,52)
    Size=New-Object System.Drawing.Size(140,22)
    Font=New-Object System.Drawing.Font("Segoe UI",9)
    Checked=$false
    Enabled=$false
}
$controls.Controls.Add($chkIncludeShared)

$btnLoad = New-Object System.Windows.Forms.Button -Property @{
    Text="Load all mailboxes"
    Location=New-Object System.Drawing.Point(860,14)
    Size=New-Object System.Drawing.Size(250,34)
    BackColor=[System.Drawing.Color]::FromArgb(0,120,212)
    ForeColor=[System.Drawing.Color]::White
    FlatStyle="Flat"
    Font=New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
    Anchor="Top,Right"
    Enabled=$false
}
$btnLoad.FlatAppearance.BorderSize = 0
$controls.Controls.Add($btnLoad)

$btnExportAll = New-Object System.Windows.Forms.Button -Property @{
    Text="Export current mailbox SOA settings (CSV)"
    Location=New-Object System.Drawing.Point(860,52)
    Size=New-Object System.Drawing.Size(250,28)
    BackColor=[System.Drawing.Color]::White
    ForeColor=[System.Drawing.Color]::FromArgb(32,31,30)
    FlatStyle="Flat"
    Font=New-Object System.Drawing.Font("Segoe UI",9)
    Anchor="Top,Right"
    Enabled=$false
}
$btnExportAll.FlatAppearance.BorderSize = 1
$btnExportAll.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200,198,196)
$controls.Controls.Add($btnExportAll)

$form.Controls.Add($controls)

# Split layout
$split = New-Object System.Windows.Forms.SplitContainer -Property @{
    Dock="Fill"; Orientation="Vertical"; SplitterDistance=860; BackColor=[System.Drawing.Color]::FromArgb(245,248,250)
}

# Grid
$gridPanel = New-Object System.Windows.Forms.Panel -Property @{ Dock="Fill"; BackColor=[System.Drawing.Color]::FromArgb(245,248,250) }
$grid = New-Object System.Windows.Forms.DataGridView -Property @{
    Dock="Fill"
    ReadOnly=$true
    AllowUserToAddRows=$false
    AllowUserToDeleteRows=$false
    MultiSelect=$true
    SelectionMode="FullRowSelect"
    AutoSizeColumnsMode="Fill"
    BackgroundColor=[System.Drawing.Color]::White
    BorderStyle="None"
}
$grid.RowHeadersVisible = $false
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(243,242,241)
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$gridPanel.Controls.Add($grid)

$paging = New-Object System.Windows.Forms.Panel -Property @{ Dock="Bottom"; Height=42; BackColor=[System.Drawing.Color]::White }
$btnPrev = New-Object System.Windows.Forms.Button -Property @{
    Text="◀ Prev"; Location=New-Object System.Drawing.Point(10,8); Size=New-Object System.Drawing.Size(90,26)
    FlatStyle="Flat"; BackColor=[System.Drawing.Color]::White; Font=New-Object System.Drawing.Font("Segoe UI",9); Enabled=$false
}
$btnPrev.FlatAppearance.BorderSize = 1
$btnPrev.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200,198,196)
$paging.Controls.Add($btnPrev)

$btnNext = New-Object System.Windows.Forms.Button -Property @{
    Text="Next ▶"; Location=New-Object System.Drawing.Point(110,8); Size=New-Object System.Drawing.Size(90,26)
    FlatStyle="Flat"; BackColor=[System.Drawing.Color]::White; Font=New-Object System.Drawing.Font("Segoe UI",9); Enabled=$false
}
$btnNext.FlatAppearance.BorderSize = 1
$btnNext.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200,198,196)
$paging.Controls.Add($btnNext)

$lblPaging = New-Object System.Windows.Forms.Label -Property @{
    Text="Page 1/1 | Showing 0-0 of 0"
    Location=New-Object System.Drawing.Point(220,12)
    Size=New-Object System.Drawing.Size(520,20)
    Font=New-Object System.Drawing.Font("Segoe UI",9)
    ForeColor=[System.Drawing.Color]::FromArgb(96,94,92)
}
$paging.Controls.Add($lblPaging)

$lblPageSize = New-Object System.Windows.Forms.Label -Property @{
    Text="Page size:"
    Location=New-Object System.Drawing.Point(660,12)
    Size=New-Object System.Drawing.Size(70,20)
    Font=New-Object System.Drawing.Font("Segoe UI",9)
    Anchor="Top,Right"
}
$paging.Controls.Add($lblPageSize)

$cmbPageSize = New-Object System.Windows.Forms.ComboBox -Property @{
    Location=New-Object System.Drawing.Point(735,9)
    Size=New-Object System.Drawing.Size(110,24)
    DropDownStyle="DropDownList"
    Font=New-Object System.Drawing.Font("Segoe UI",9)
    Anchor="Top,Right"
    Enabled=$false
}
@("50","100","250","500","1000") | ForEach-Object { [void]$cmbPageSize.Items.Add($_) }
$cmbPageSize.SelectedItem = "100"
$paging.Controls.Add($cmbPageSize)

$gridPanel.Controls.Add($paging)
$split.Panel1.Controls.Add($gridPanel)

# Right panel: details + actions
$right = New-Object System.Windows.Forms.Panel -Property @{ Dock="Fill"; BackColor=[System.Drawing.Color]::White; Padding = New-Object System.Windows.Forms.Padding(14) }

$grpDetails = New-Object System.Windows.Forms.GroupBox -Property @{ Text="Selected mailbox"; Dock="Top"; Height=200; Font=New-Object System.Drawing.Font("Segoe UI",9) }
$txtDetails = New-Object System.Windows.Forms.TextBox -Property @{
    Dock="Fill"; Multiline=$true; ReadOnly=$true; ScrollBars="Vertical"
    Font=New-Object System.Drawing.Font("Consolas",9)
    Text="Select a mailbox row to see details."
}
$grpDetails.Controls.Add($txtDetails)
$right.Controls.Add($grpDetails)

$grpActions = New-Object System.Windows.Forms.GroupBox -Property @{ Text="Actions"; Dock="Top"; Height=210; Font=New-Object System.Drawing.Font("Segoe UI",9) }

$btnSetCloud = New-Object System.Windows.Forms.Button -Property @{
    Text="Set SOA -> Online (Enable)"
    Location=New-Object System.Drawing.Point(18,32)
    Size=New-Object System.Drawing.Size(320,40)
    BackColor=[System.Drawing.Color]::FromArgb(0,120,212)
    ForeColor=[System.Drawing.Color]::White
    FlatStyle="Flat"
    Font=New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
    Enabled=$false
}
$btnSetCloud.FlatAppearance.BorderSize = 0
$grpActions.Controls.Add($btnSetCloud)

$btnSetOnPrem = New-Object System.Windows.Forms.Button -Property @{
    Text="Set SOA -> On-Prem (Revert)"
    Location=New-Object System.Drawing.Point(18,82)
    Size=New-Object System.Drawing.Size(320,40)
    BackColor=[System.Drawing.Color]::FromArgb(252,80,34)
    ForeColor=[System.Drawing.Color]::White
    FlatStyle="Flat"
    Font=New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
    Enabled=$false
}
$btnSetOnPrem.FlatAppearance.BorderSize = 0
$grpActions.Controls.Add($btnSetOnPrem)

$lblHint = New-Object System.Windows.Forms.Label -Property @{
    Text="Only DirSynced mailboxes can switch SOA. Cloud-only mailboxes show '— N/A'."
    Location=New-Object System.Drawing.Point(18,135)
    Size=New-Object System.Drawing.Size(340,40)
    Font=New-Object System.Drawing.Font("Segoe UI",8)
    ForeColor=[System.Drawing.Color]::FromArgb(96,94,92)
}
$grpActions.Controls.Add($lblHint)

$right.Controls.Add($grpActions)
$split.Panel2.Controls.Add($right)

$form.Controls.Add($split)

# Status bar
$status = New-Object System.Windows.Forms.Panel -Property @{ Dock="Bottom"; Height=40; BackColor=[System.Drawing.Color]::White }
$lblStatus = New-Object System.Windows.Forms.Label -Property @{
    Text="Disconnected - click Connect to begin."
    Location=New-Object System.Drawing.Point(12,11)
    Size=New-Object System.Drawing.Size(780,20)
    Font=New-Object System.Drawing.Font("Segoe UI",9)
    ForeColor=[System.Drawing.Color]::FromArgb(32,31,30)
}
$status.Controls.Add($lblStatus)

$progress = New-Object System.Windows.Forms.ProgressBar -Property @{
    Location=New-Object System.Drawing.Point(820,10)
    Size=New-Object System.Drawing.Size(400,20)
    Anchor="Top,Right"
    Value=0
}
$status.Controls.Add($progress)
$form.Controls.Add($status)

# =========================
# View logic (filter + paging + binding)
# =========================
$binding = New-Object System.Windows.Forms.BindingSource
$grid.DataSource = $binding

function Set-Status {
    param([string]$Text, [int]$Pct = -1)
    $lblStatus.Text = $Text
    if ($Pct -ge 0 -and $Pct -le 100) { $progress.Value = $Pct }
}

function Apply-Filters {
    $q = ($txtSearch.Text ?? "").Trim()
    $dirOnly = [bool]$chkDirSyncedOnly.Checked

    $data = $Script:MailboxCache ?? @()
    if ($dirOnly) { $data = @($data | Where-Object { $_.IsDirSynced -eq $true }) }

    if ($q.Length -gt 0) {
        $data = @(
            $data | Where-Object {
                ([string]$_.DisplayName).Contains($q, [System.StringComparison]::OrdinalIgnoreCase) -or
                ([string]$_.UserPrincipalName).Contains($q, [System.StringComparison]::OrdinalIgnoreCase) -or
                ([string]$_.PrimarySmtpAddress).Contains($q, [System.StringComparison]::OrdinalIgnoreCase)
            }
        )
    }

    $Script:View = $data
}

function Render-Page {
    $items = $Script:View ?? @()
    $total = $items.Count
    $pages = [Math]::Max(1, [Math]::Ceiling($total / [double]$Script:PageSize))
    if ($Script:CurrentPage -lt 1) { $Script:CurrentPage = 1 }
    if ($Script:CurrentPage -gt $pages) { $Script:CurrentPage = $pages }

    $start = ($Script:CurrentPage - 1) * $Script:PageSize
    $endExcl = [Math]::Min($start + $Script:PageSize, $total)

    $pageItems = @()
    if ($total -gt 0 -and $start -lt $total) {
        $pageItems = $items[$start..($endExcl-1)]
    }

    $dt = Build-GridTable -Items $pageItems
    $binding.DataSource = $dt
    $binding.ResetBindings($true)

    $from = if ($total -eq 0) { 0 } else { $start + 1 }
    $lblPaging.Text = ("Page {0}/{1}  |  Showing {2}-{3} of {4}" -f $Script:CurrentPage, $pages, $from, $endExcl, $total)

    $btnPrev.Enabled = ($Script:CurrentPage -gt 1)
    $btnNext.Enabled = ($Script:CurrentPage -lt $pages)
}

function Refresh-Grid {
    Apply-Filters
    Render-Page
}

function Update-ConnectedUI {
    $connected = $Script:IsConnected
    $btnConnect.Enabled = -not $connected
    $btnDisconnect.Enabled = $connected
    $btnLoad.Enabled = $connected
    $btnExportAll.Enabled = $connected
    $txtSearch.Enabled = $connected
    $btnClear.Enabled = $connected
    $chkDirSyncedOnly.Enabled = $connected
    $chkIncludeShared.Enabled = $connected
    $cmbPageSize.Enabled = $connected

    $btnSetCloud.Enabled = $connected
    $btnSetOnPrem.Enabled = $connected

    if ($connected) {
        $t = if ($Script:TenantName) { $Script:TenantName } else { "Unknown" }
        Set-Status ("Connected to Exchange Online (Tenant: {0})" -f $t) 0
    } else {
        Set-Status "Disconnected - click Connect to begin." 0
        $Script:MailboxCache = @()
        $Script:View = @()
        $Script:CurrentPage = 1
        $binding.DataSource = New-GridTable
        $binding.ResetBindings($true)
        $txtDetails.Text = "Select a mailbox row to see details."
        $lblPaging.Text = "Page 1/1 | Showing 0-0 of 0"
        $btnPrev.Enabled = $false
        $btnNext.Enabled = $false
    }
}

# =========================
# Actions
# =========================
function Get-SelectedUpns {
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($r in $grid.SelectedRows) {
        $upn = [string]$r.Cells["UserPrincipalName"].Value
        if ($upn) { [void]$list.Add($upn) }
    }
    return @($list | Select-Object -Unique)
}

function Set-MailboxSOA {
    param([bool]$CloudManaged)

    if (-not $Script:IsConnected) { return }

    $targets = Get-SelectedUpns
    if ($targets.Count -eq 0) {
        Show-InfoBox "Select one or more mailboxes first." $Script:ToolName
        return
    }

    $targetLabel = if ($CloudManaged) { "Online (Enable)" } else { "On-Prem (Revert)" }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        ("Set SOA -> {0} for {1} mailbox(es)?" -f $targetLabel, $targets.Count),
        $Script:ToolName,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Log "INFO" ("User cancelled SOA change -> {0}" -f $targetLabel)
        return
    }

    $ok = 0; $skip = 0; $fail = 0
    Set-Status ("Applying SOA change -> {0}..." -f $targetLabel) 10

    foreach ($upn in $targets) {
        try {
            $row = $Script:MailboxCache | Where-Object { $_.UserPrincipalName -eq $upn } | Select-Object -First 1
            if (-not $row) { $skip++; Write-Log "WARN" ("Skip (not in cache): {0}" -f $upn); continue }

            if (-not $row.IsDirSynced) {
                $skip++
                Write-Log "WARN" ("Skip (not DirSynced): {0}" -f $upn)
                continue
            }

            if ($CloudManaged -and $row.IsExchangeCloudManaged -eq $true) {
                $skip++; Write-Log "INFO" ("No change needed (already Online): {0}" -f $upn); continue
            }
            if ((-not $CloudManaged) -and $row.IsExchangeCloudManaged -eq $false) {
                $skip++; Write-Log "INFO" ("No change needed (already On-Prem): {0}" -f $upn); continue
            }

            Write-Log "INFO" ("Set-Mailbox -Identity {0} -IsExchangeCloudManaged {1}" -f $upn, $CloudManaged)
            Set-Mailbox -Identity $upn -IsExchangeCloudManaged $CloudManaged -ErrorAction Stop
            $ok++
        } catch {
            $fail++
            Write-Log "ERROR" ("SOA update failed for {0}: {1}" -f $upn, $_.Exception.Message)
        }
    }

    # Refresh cache (lightweight): re-load all to keep it simple/reliable
    try {
        Set-Status "Refreshing mailbox list..." 70
        $raw = Get-Mailboxes -IncludeShared:$chkIncludeShared.Checked
        $Script:MailboxCache = @($raw | ForEach-Object { To-Row $_ })
        $Script:CurrentPage = 1
        Refresh-Grid
    } catch {
        Write-Log "WARN" ("Refresh after SOA change failed: {0}" -f $_.Exception.Message)
    }

    Set-Status ("Done. Success={0} Skipped={1} Failed={2}" -f $ok, $skip, $fail) 100
    Show-InfoBox ("Completed.`nSuccess: {0}`nSkipped: {1}`nFailed: {2}`n`nLog: {3}" -f $ok, $skip, $fail, $Script:LogFile) $Script:ToolName
}

function Export-CacheToCsv {
    if (-not $Script:MailboxCache -or $Script:MailboxCache.Count -eq 0) {
        Show-InfoBox "Nothing to export. Load mailboxes first." $Script:ToolName
        return
    }

    if (-not (Test-Path $Script:ExportsDir)) { New-Item -ItemType Directory -Path $Script:ExportsDir -Force | Out-Null }
    $ts = Get-Date -Format "yyyyMMdd-HHmmss"
    $file = Join-Path $Script:ExportsDir ("MailboxSOASettings-{0}.csv" -f $ts)

    try {
        Write-Log "INFO" ("Exporting cache to CSV: {0}" -f $file)
        $Script:MailboxCache |
            Select-Object Indicator, SOAState, DisplayName, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails, IsDirSynced, IsExchangeCloudManaged |
            Export-Csv -Path $file -NoTypeInformation -Encoding UTF8

        Show-InfoBox ("Exported:`n{0}" -f $file) $Script:ToolName
    } catch {
        Write-Log "ERROR" ("Export failed: {0}" -f $_.Exception.Message)
        Show-ErrorBox ("Export failed:`n{0}" -f $_.Exception.Message) $Script:ToolName
    }
}

# =========================
# Events
# =========================
$btnConnect.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        if (Connect-EXO) {
            Update-ConnectedUI
        }
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnDisconnect.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Disconnect-EXO
        Update-ConnectedUI
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnLoad.Add_Click({
    if (-not $Script:IsConnected) { return }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Set-Status "Loading mailboxes..." 15
        Write-Log "INFO" ("Load all mailboxes (IncludeShared={0})" -f $chkIncludeShared.Checked)

        $raw = Get-Mailboxes -IncludeShared:$chkIncludeShared.Checked
        Set-Status ("Processing {0} objects..." -f $raw.Count) 35

        $Script:MailboxCache = @($raw | ForEach-Object { To-Row $_ })
        $Script:CurrentPage = 1

        $txtSearch.Text = ""
        $cmbPageSize.SelectedItem = [string]$Script:PageSize

        Refresh-Grid
        Set-Status ("Loaded {0} mailboxes." -f $Script:MailboxCache.Count) 100
        Write-Log "INFO" ("Loaded into cache: {0}" -f $Script:MailboxCache.Count)
    } catch {
        Write-Log "ERROR" ("Load failed: {0}" -f $_.Exception.Message)
        Show-ErrorBox ("Load all mailboxes failed:`n{0}`n`nLog: {1}" -f $_.Exception.Message, $Script:LogFile) $Script:ToolName
        Set-Status "Load failed. See log." 0
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnExportAll.Add_Click({ Export-CacheToCsv })

$txtSearch.Add_TextChanged({
    $Script:CurrentPage = 1
    Refresh-Grid
})

$btnClear.Add_Click({
    $txtSearch.Text = ""
    $Script:CurrentPage = 1
    Refresh-Grid
})

$chkDirSyncedOnly.Add_CheckedChanged({
    $Script:CurrentPage = 1
    Refresh-Grid
})

$cmbPageSize.Add_SelectedIndexChanged({
    $Script:PageSize = [int]$cmbPageSize.SelectedItem
    $Script:CurrentPage = 1
    Refresh-Grid
})

$btnPrev.Add_Click({
    $Script:CurrentPage--
    if ($Script:CurrentPage -lt 1) { $Script:CurrentPage = 1 }
    Render-Page
})

$btnNext.Add_Click({
    $Script:CurrentPage++
    Render-Page
})

$grid.Add_SelectionChanged({
    try {
        if ($grid.SelectedRows.Count -eq 0) { return }
        $upn = [string]$grid.SelectedRows[0].Cells["UserPrincipalName"].Value
        if (-not $upn) { return }

        $obj = $Script:MailboxCache | Where-Object { $_.UserPrincipalName -eq $upn } | Select-Object -First 1
        if (-not $obj) { return }

        $txtDetails.Text = @"
DisplayName            : $($obj.DisplayName)
UPN                   : $($obj.UserPrincipalName)
Primary SMTP          : $($obj.PrimarySmtpAddress)
RecipientTypeDetails  : $($obj.RecipientTypeDetails)
DirSynced             : $($obj.IsDirSynced)
IsExchangeCloudManaged: $($obj.IsExchangeCloudManaged)
SOA                   : $($obj.Indicator)
"@
    } catch {
        Write-Log "WARN" ("SelectionChanged failed: {0}" -f $_.Exception.Message)
    }
})

$btnSetCloud.Add_Click({ Set-MailboxSOA -CloudManaged $true })
$btnSetOnPrem.Add_Click({ Set-MailboxSOA -CloudManaged $false })

$form.Add_FormClosing({
    try {
        Write-Log "INFO" "Form closing..."
        if ($Script:IsConnected) { Disconnect-EXO }
        Write-Log "INFO" "=== Closed ==="
    } catch { }
})

$form.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq "R") {
        if ($btnLoad.Enabled) { $btnLoad.PerformClick() }
    } elseif ($_.KeyCode -eq "Escape") {
        $form.Close()
    }
})

# Initial UI state
Update-ConnectedUI

[void]$form.ShowDialog()
