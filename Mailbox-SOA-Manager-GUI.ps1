<#
.SYNOPSIS
  Mailbox SOA Manager (GUI) for Exchange Online - Cloud-managed Exchange attributes (SOA) toggle.

.DESCRIPTION
  GUI tool to view and change mailbox Exchange attribute SOA state via IsExchangeCloudManaged:
    - Enable cloud management     : IsExchangeCloudManaged = $true
    - Revert to on-prem management: IsExchangeCloudManaged = $false

  Browsing/Search improvements (v1.6):
    - "Load all mailboxes" button (caches mailboxes locally for fast browsing/search)
    - Paging controls (Prev/Next + Page size + page indicator)
    - Search uses cached list when available (reliable + fast)
    - Clear search returns to full cached view
    - Server-side search remains as fallback (best-effort)

REFERENCE
  https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

LOGGING
  - Single logfile only (append; never overwritten)
  - Timestamp on every line
  - SOA changes logged with BEFORE/AFTER + Actor
  - RunId included for correlation

REQUIREMENTS
  - Windows PowerShell 5.1 OR PowerShell 7+ (must run in STA for WinForms)
  - Module: ExchangeOnlineManagement

AUTHOR
  Peter

VERSION
  1.6 (2026-01-04)
#>

# --- Load WinForms early (so we can show MessageBoxes even before GUI starts) ---
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    [System.Windows.Forms.Application]::EnableVisualStyles()
} catch {
    Write-Error "Failed to load WinForms assemblies. This tool must run on Windows with WinForms available. Error: $($_.Exception.Message)"
    return
}

#region Globals
$Script:ToolName     = "Mailbox SOA Manager"
$Script:RunId        = [Guid]::NewGuid().ToString()
$Script:LogDir       = Join-Path -Path (Get-Location) -ChildPath "Logs"
$Script:ExportDir    = Join-Path -Path (Get-Location) -ChildPath "Exports"   # created on-demand
$Script:LogFile      = Join-Path -Path $Script:LogDir -ChildPath "MailboxSOAManager.log"
$Script:IsConnected  = $false
$Script:ExoActor     = $null  # populated after connect (best-effort)

# Cache + paging state
$Script:MailboxCache     = @()   # full cache
$Script:CurrentView      = @()   # current view (cache or filtered)
$Script:CacheLoaded      = $false
$Script:PageSize         = 50
$Script:PageIndex        = 0
$Script:CurrentQueryText = ""

New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null
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

#region STA guard (AUTO RELAUNCH)
function Ensure-STA {
    try {
        $apt = [System.Threading.Thread]::CurrentThread.GetApartmentState()
        if ($apt -eq [System.Threading.ApartmentState]::STA) { return $true }

        Write-Log "Not running in STA mode (ApartmentState=$apt). Attempting self-relaunch in STA..." "WARN"

        $scriptPath = $MyInvocation.MyCommand.Path
        if ([string]::IsNullOrWhiteSpace($scriptPath) -or -not (Test-Path $scriptPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "This GUI must run in STA mode, but the script path could not be detected for auto-relaunch.`n`nPlease run it like:`n  powershell.exe -STA -File .\MailboxSOAManager-GUI.ps1",
                $Script:ToolName,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return $false
        }

        $exe = if ($PSVersionTable.PSEdition -eq "Core") { "pwsh.exe" } else { "powershell.exe" }
        $args = @("-NoProfile","-ExecutionPolicy","Bypass","-STA","-File","`"$scriptPath`"") -join " "

        Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory (Split-Path -Parent $scriptPath) | Out-Null
        Write-Log "Launched new process: $exe $args" "INFO"
        return $false
    } catch {
        Write-Log "Ensure-STA failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to validate STA mode.`n`n$($_.Exception.Message)",
            $Script:ToolName,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
}
if (-not (Ensure-STA)) { return }
#endregion

Write-Log "$($Script:ToolName) starting (GUI init)..." "INFO"

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

#region Helpers
function Get-SOAIndicator {
    param([object]$IsExchangeCloudManaged)
    if ($IsExchangeCloudManaged -eq $true)  { return "☁ Online" }
    if ($IsExchangeCloudManaged -eq $false) { return "🏢 On-Prem" }
    return "? Unknown"
}

function Escape-OPathValue {
    param([string]$Value)
    # OPATH string literal escaping: single quote doubled
    if ($null -eq $Value) { return "" }
    return ($Value -replace "'", "''")
}

function Get-MailboxCmdletMode {
    # Prefer EXO v3 cmdlets if present (better perf, less serialization)
    if (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue) { return "EXO" }
    return "Classic"
}

function Get-MailboxesServerSide {
    param(
        [int]$Max = 200,
        [string]$Filter = $null
    )

    $mode = Get-MailboxCmdletMode
    if ($mode -eq "EXO") {
        if ([string]::IsNullOrWhiteSpace($Filter)) {
            return Get-EXOMailbox -ResultSize $Max -PropertySets Minimum -ErrorAction Stop
        } else {
            return Get-EXOMailbox -ResultSize $Max -Filter $Filter -PropertySets Minimum -ErrorAction Stop
        }
    } else {
        if ([string]::IsNullOrWhiteSpace($Filter)) {
            return Get-Mailbox -ResultSize $Max -ErrorAction Stop
        } else {
            return Get-Mailbox -ResultSize $Max -Filter $Filter -ErrorAction Stop
        }
    }
}

function Convert-ToGridRow {
    param($MailboxObject)

    # Handle both EXO & classic types
    $primary = $MailboxObject.PrimarySmtpAddress
    if ($primary -and $primary.ToString) { $primary = $primary.ToString() }

    [PSCustomObject]@{
        DisplayName                = $MailboxObject.DisplayName
        Alias                      = $MailboxObject.Alias
        PrimarySmtpAddress         = $primary
        RecipientTypeDetails       = $MailboxObject.RecipientTypeDetails
        IsDirSynced                = $MailboxObject.IsDirSynced
        IsExchangeCloudManaged     = $MailboxObject.IsExchangeCloudManaged
        "SOA (Exchange Attributes)"= (Get-SOAIndicator $MailboxObject.IsExchangeCloudManaged)
    }
}
#endregion

#region EXO connect/disconnect
function Get-ExoActorBestEffort {
    try {
        $ci = Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($ci) {
            $info = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
            if ($info -and $info.UserPrincipalName) { return "EXO:$($info.UserPrincipalName)" }
        }
    } catch { }
    return "EXO:unknown"
}

function Connect-EXO {
    try {
        Ensure-Module -Name "ExchangeOnlineManagement"
        Write-Log "Connecting to Exchange Online..." "INFO"
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
        $Script:IsConnected = $true
        $Script:ExoActor = Get-ExoActorBestEffort
        Write-Log "Connected to Exchange Online. CmdletMode=$(Get-MailboxCmdletMode)" "INFO"
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

        # reset cache/view on disconnect
        $Script:MailboxCache     = @()
        $Script:CurrentView      = @()
        $Script:CacheLoaded      = $false
        $Script:PageIndex        = 0
        $Script:CurrentQueryText = ""
    }
}
#endregion

#region Mailbox ops
function Get-MailboxDetails {
    param([Parameter(Mandatory)][string]$Identity)

    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    # Use Get-Mailbox for details (classic has richer object; EXO ok too)
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

function Export-MailboxSOASettings {
    param([Parameter(Mandatory)][string]$Identity)

    New-Item -ItemType Directory -Path $Script:ExportDir -Force | Out-Null

    $details = Get-MailboxDetails -Identity $Identity
    $mbx = $details.Mailbox
    $soa = Get-SOAIndicator $mbx.IsExchangeCloudManaged

    $safeId  = ($mbx.PrimarySmtpAddress.ToString() -replace '[^a-zA-Z0-9\.\-_@]','_')
    $stamp   = Get-Date -Format "yyyyMMdd-HHmmss"
    $path    = Join-Path $Script:ExportDir "$safeId-MailboxSOASettings-$stamp.json"

    $export = [PSCustomObject]@{
        ExportType    = "Mailbox SOA Settings"
        ExportedAt    = (Get-Date).ToString("o")
        ToolName      = $Script:ToolName
        RunId         = $Script:RunId
        Identity      = $Identity
        SOASettings   = [PSCustomObject]@{
            PrimarySmtpAddress      = $mbx.PrimarySmtpAddress
            DisplayName             = $mbx.DisplayName
            IsDirSynced             = $mbx.IsDirSynced
            IsExchangeCloudManaged  = $mbx.IsExchangeCloudManaged
            SOAIndicator            = $soa
        }
        MailboxDetails = $mbx
        UserDetails    = $details.User
    }

    $export | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8
    Write-Log "Export-MailboxSOASettings completed for '$Identity'. Path='$path' SOA='$soa' IsExchangeCloudManaged='$($mbx.IsExchangeCloudManaged)'" "INFO"
    return $path
}

function Set-MailboxSOACloudManaged {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][bool]$EnableCloudManaged
    )
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $targetValue = [bool]$EnableCloudManaged

    $before = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

    Write-Log "SOA change requested for '$Identity'. TargetIsExchangeCloudManaged=$targetValue (Before=$($before.IsExchangeCloudManaged); IsDirSynced=$($before.IsDirSynced))" "INFO"

    if ($before.IsDirSynced -ne $true) {
        $msg = "Mailbox '$Identity' is not DirSynced (IsDirSynced=$($before.IsDirSynced)). Change blocked."
        Write-Log $msg "WARN"
        throw $msg
    }

    if ($before.IsExchangeCloudManaged -eq $targetValue) {
        $msg = "No change needed for '$Identity'. IsExchangeCloudManaged already '$targetValue'."
        Write-Log $msg "INFO"
        return $msg
    }

    try {
        Set-Mailbox -Identity $Identity -IsExchangeCloudManaged $targetValue -ErrorAction Stop
        Write-Log "Set-Mailbox executed for '$Identity' IsExchangeCloudManaged=$targetValue" "INFO"
    } catch {
        Write-Log "Set-Mailbox FAILED for '$Identity'. Error=$($_.Exception.Message)" "ERROR"
        throw
    }

    $after = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

    $changed = ($after.IsExchangeCloudManaged -eq $targetValue)
    Write-Log "SOA change result for '$Identity'. Before=$($before.IsExchangeCloudManaged) After=$($after.IsExchangeCloudManaged) Expected=$targetValue Success=$changed" "INFO"

    if (-not $changed) {
        return "Executed, but verification did not match expected value. Before='$($before.IsExchangeCloudManaged)' After='$($after.IsExchangeCloudManaged)' Expected='$targetValue'."
    }

    return "Updated. IsExchangeCloudManaged is now '$($after.IsExchangeCloudManaged)'."
}
#endregion

#region Cache + Paging
function Reset-ViewToCache {
    $Script:CurrentView = @($Script:MailboxCache)
    $Script:PageIndex = 0
    $Script:CurrentQueryText = ""
}

function Apply-SearchToCache {
    param([string]$QueryText)

    $q = ($QueryText ?? "").Trim()
    $Script:CurrentQueryText = $q
    $Script:PageIndex = 0

    if ([string]::IsNullOrWhiteSpace($q)) {
        $Script:CurrentView = @($Script:MailboxCache)
        return
    }

    $Script:CurrentView = @(
        $Script:MailboxCache | Where-Object {
            ($_.DisplayName -like "*$q*") -or
            ($_.Alias -like "*$q*") -or
            ($_.PrimarySmtpAddress -like "*$q*")
        }
    )
}

function Get-PageSlice {
    param(
        [array]$Items,
        [int]$PageIndex,
        [int]$PageSize
    )
    if (-not $Items) { return @() }
    if ($PageSize -le 0) { $PageSize = 50 }

    $count = $Items.Count
    $start = $PageIndex * $PageSize
    if ($start -ge $count) { return @() }

    $end = [Math]::Min($start + $PageSize - 1, $count - 1)
    if ($end -lt $start) { return @() }

    return @($Items[$start..$end])
}

function Get-TotalPages {
    param([array]$Items,[int]$PageSize)
    if (-not $Items -or $Items.Count -eq 0) { return 0 }
    if ($PageSize -le 0) { $PageSize = 50 }
    return [int][Math]::Ceiling($Items.Count / [double]$PageSize)
}
#endregion

#region GUI
try {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($Script:ToolName) - Exchange Online (IsExchangeCloudManaged)"
    $form.Size = New-Object System.Drawing.Size(1100, 700)
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

    $lblConn = New-Object System.Windows.Forms.Label
    $lblConn.Text = "Status: Not connected"
    $lblConn.Location = New-Object System.Drawing.Point(260, 18)
    $lblConn.AutoSize = $true

    # Search + browse group
    $grpBrowse = New-Object System.Windows.Forms.GroupBox
    $grpBrowse.Text = "Browse & Search"
    $grpBrowse.Location = New-Object System.Drawing.Point(12, 55)
    $grpBrowse.Size = New-Object System.Drawing.Size(1060, 170)

    $btnLoadAll = New-Object System.Windows.Forms.Button
    $btnLoadAll.Text = "Load all mailboxes"
    $btnLoadAll.Location = New-Object System.Drawing.Point(16, 30)
    $btnLoadAll.Size = New-Object System.Drawing.Size(170, 30)
    $btnLoadAll.Enabled = $false

    $lblLoadHint = New-Object System.Windows.Forms.Label
    $lblLoadHint.Text = "Tip: Load all first for fast browsing/search."
    $lblLoadHint.Location = New-Object System.Drawing.Point(200, 36)
    $lblLoadHint.AutoSize = $true

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(16, 70)
    $txtSearch.Size = New-Object System.Drawing.Size(520, 25)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = "Search"
    $btnSearch.Location = New-Object System.Drawing.Point(548, 67)
    $btnSearch.Size = New-Object System.Drawing.Size(110, 30)
    $btnSearch.Enabled = $false

    $btnClearSearch = New-Object System.Windows.Forms.Button
    $btnClearSearch.Text = "Clear"
    $btnClearSearch.Location = New-Object System.Drawing.Point(666, 67)
    $btnClearSearch.Size = New-Object System.Drawing.Size(110, 30)
    $btnClearSearch.Enabled = $false

    # Paging controls
    $btnPrev = New-Object System.Windows.Forms.Button
    $btnPrev.Text = "◀ Prev"
    $btnPrev.Location = New-Object System.Drawing.Point(16, 110)
    $btnPrev.Size = New-Object System.Drawing.Size(90, 30)
    $btnPrev.Enabled = $false

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = "Next ▶"
    $btnNext.Location = New-Object System.Drawing.Point(112, 110)
    $btnNext.Size = New-Object System.Drawing.Size(90, 30)
    $btnNext.Enabled = $false

    $lblPage = New-Object System.Windows.Forms.Label
    $lblPage.Text = "Page: -"
    $lblPage.Location = New-Object System.Drawing.Point(220, 116)
    $lblPage.AutoSize = $true

    $lblPageSize = New-Object System.Windows.Forms.Label
    $lblPageSize.Text = "Page size:"
    $lblPageSize.Location = New-Object System.Drawing.Point(360, 116)
    $lblPageSize.AutoSize = $true

    $cmbPageSize = New-Object System.Windows.Forms.ComboBox
    $cmbPageSize.Location = New-Object System.Drawing.Point(430, 112)
    $cmbPageSize.Size = New-Object System.Drawing.Size(90, 25)
    $cmbPageSize.DropDownStyle = 'DropDownList'
    [void]$cmbPageSize.Items.AddRange(@("25","50","100","200"))
    $cmbPageSize.SelectedItem = "50"
    $cmbPageSize.Enabled = $false

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "Count: -"
    $lblCount.Location = New-Object System.Drawing.Point(548, 116)
    $lblCount.AutoSize = $true

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(16, 145)
    $lblStatus.Size = New-Object System.Drawing.Size(1020, 20)

    # Grid (taller now)
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(16, 235)
    $grid.Size = New-Object System.Drawing.Size(1056, 220)
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
    $grpDetails.Location = New-Object System.Drawing.Point(12, 465)
    $grpDetails.Size = New-Object System.Drawing.Size(1060, 170)

    $txtDetails = New-Object System.Windows.Forms.TextBox
    $txtDetails.Location = New-Object System.Drawing.Point(16, 25)
    $txtDetails.Size = New-Object System.Drawing.Size(1028, 95)
    $txtDetails.Multiline = $true
    $txtDetails.ScrollBars = "Vertical"
    $txtDetails.ReadOnly = $true

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh details"
    $btnRefresh.Location = New-Object System.Drawing.Point(16, 130)
    $btnRefresh.Size = New-Object System.Drawing.Size(160, 30)
    $btnRefresh.Enabled = $false

    $btnExportSOA = New-Object System.Windows.Forms.Button
    $btnExportSOA.Text = "Export current mailbox SOA settings (JSON)"
    $btnExportSOA.Location = New-Object System.Drawing.Point(186, 130)
    $btnExportSOA.Size = New-Object System.Drawing.Size(300, 30)
    $btnExportSOA.Enabled = $false

    $btnEnableCloud = New-Object System.Windows.Forms.Button
    $btnEnableCloud.Text = "Enable cloud SOA (true)"
    $btnEnableCloud.Location = New-Object System.Drawing.Point(500, 130)
    $btnEnableCloud.Size = New-Object System.Drawing.Size(180, 30)
    $btnEnableCloud.Enabled = $false

    $btnRevertOnPrem = New-Object System.Windows.Forms.Button
    $btnRevertOnPrem.Text = "Revert to on-prem SOA (false)"
    $btnRevertOnPrem.Location = New-Object System.Drawing.Point(690, 130)
    $btnRevertOnPrem.Size = New-Object System.Drawing.Size(220, 30)
    $btnRevertOnPrem.Enabled = $false

    # Footer
    $lblFoot = New-Object System.Windows.Forms.Label
    $lblFoot.Location = New-Object System.Drawing.Point(12, 640)
    $lblFoot.Size = New-Object System.Drawing.Size(1060, 22)
    $lblFoot.Text = "Log: $($Script:LogFile)   |   Export: .\Exports (created only when exporting)"
    $lblFoot.AutoSize = $false

    # Add controls
    $form.Controls.AddRange(@($btnConnect,$btnDisconnect,$lblConn,$grpBrowse,$grid,$grpDetails,$lblFoot))
    $grpBrowse.Controls.AddRange(@(
        $btnLoadAll,$lblLoadHint,$txtSearch,$btnSearch,$btnClearSearch,
        $btnPrev,$btnNext,$lblPage,$lblPageSize,$cmbPageSize,$lblCount,$lblStatus
    ))
    $grpDetails.Controls.AddRange(@($txtDetails,$btnRefresh,$btnExportSOA,$btnEnableCloud,$btnRevertOnPrem))

    # UI state
    $Script:SelectedIdentity = $null

    function Update-PagingUI {
        $totalPages = Get-TotalPages -Items $Script:CurrentView -PageSize $Script:PageSize
        $totalItems = if ($Script:CurrentView) { $Script:CurrentView.Count } else { 0 }

        if ($totalPages -eq 0) {
            $lblPage.Text = "Page: -"
            $lblCount.Text = "Count: 0"
            $btnPrev.Enabled = $false
            $btnNext.Enabled = $false
            return
        }

        if ($Script:PageIndex -lt 0) { $Script:PageIndex = 0 }
        if ($Script:PageIndex -gt ($totalPages - 1)) { $Script:PageIndex = $totalPages - 1 }

        $lblPage.Text = "Page: $($Script:PageIndex + 1) / $totalPages"
        $lblCount.Text = "Count: $totalItems"

        $btnPrev.Enabled = ($Script:PageIndex -gt 0)
        $btnNext.Enabled = ($Script:PageIndex -lt ($totalPages - 1))
    }

    function Bind-GridFromCurrentView {
        $pageItems = Get-PageSlice -Items $Script:CurrentView -PageIndex $Script:PageIndex -PageSize $Script:PageSize
        $grid.DataSource = $null
        $grid.DataSource = $pageItems
        Update-PagingUI
    }

    function Reset-SelectionAndDetails {
        $Script:SelectedIdentity = $null
        $txtDetails.Clear()
        $btnRefresh.Enabled = $false
        $btnExportSOA.Enabled = $false
        $btnEnableCloud.Enabled = $false
        $btnRevertOnPrem.Enabled = $false
    }

    function Set-UiConnectedState {
        param([bool]$Connected)

        $btnConnect.Enabled    = -not $Connected
        $btnDisconnect.Enabled = $Connected

        $btnLoadAll.Enabled    = $Connected
        $btnSearch.Enabled     = $Connected
        $btnClearSearch.Enabled= $false

        $cmbPageSize.Enabled   = $Connected

        if (-not $Connected) {
            $lblConn.Text = "Status: Not connected"
            $grid.DataSource = $null
            Reset-SelectionAndDetails
            $lblStatus.Text = ""
            $lblPage.Text = "Page: -"
            $lblCount.Text = "Count: -"
            $btnPrev.Enabled = $false
            $btnNext.Enabled = $false
        } else {
            $lblConn.Text = "Status: Connected to Exchange Online"
            $lblStatus.Text = "Connected. Recommended: Click 'Load all mailboxes' to browse/search quickly."
        }
    }

    function Show-Details {
        param([string]$Identity)

        $details = Get-MailboxDetails -Identity $Identity
        $mbx = $details.Mailbox
        $usr = $details.User

        $soa = Get-SOAIndicator $mbx.IsExchangeCloudManaged

        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("Mailbox:")
        $lines.Add("  DisplayName               : $($mbx.DisplayName)")
        $lines.Add("  PrimarySmtpAddress        : $($mbx.PrimarySmtpAddress)")
        $lines.Add("  RecipientTypeDetails      : $($mbx.RecipientTypeDetails)")
        $lines.Add("  IsDirSynced               : $($mbx.IsDirSynced)")
        $lines.Add("  IsExchangeCloudManaged    : $($mbx.IsExchangeCloudManaged)")
        $lines.Add("  SOA (Exchange Attributes) : $soa")
        $lines.Add("  ExchangeGuid              : $($mbx.ExchangeGuid)")
        $lines.Add("  ExternalDirectoryObjectId : $($mbx.ExternalDirectoryObjectId)")

        if ($usr) {
            $lines.Add("")
            $lines.Add("User:")
            $lines.Add("  UserPrincipalName         : $($usr.UserPrincipalName)")
            $lines.Add("  ImmutableId               : $($usr.ImmutableId)")
            $lines.Add("  WhenChangedUTC            : $($usr.WhenChangedUTC)")
        }

        $lines.Add("")
        $lines.Add("Export:")
        $lines.Add("  Export current mailbox SOA settings (JSON) includes:")
        $lines.Add("    - IsDirSynced, IsExchangeCloudManaged, SOA indicator")
        $lines.Add("    - Mailbox/User identifiers for traceability")

        $txtDetails.Lines = $lines.ToArray()

        $btnRefresh.Enabled      = $true
        $btnExportSOA.Enabled    = $true
        $btnEnableCloud.Enabled  = $true
        $btnRevertOnPrem.Enabled = $true
    }

    # Events
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

    $cmbPageSize.Add_SelectedIndexChanged({
        try {
            $Script:PageSize = [int]$cmbPageSize.SelectedItem
            $Script:PageIndex = 0
            Write-Log "PageSize changed to $($Script:PageSize)" "INFO"
            if ($Script:CacheLoaded -and $Script:CurrentView) {
                Bind-GridFromCurrentView
            }
        } catch { }
    })

    $btnPrev.Add_Click({
        if ($Script:PageIndex -gt 0) {
            $Script:PageIndex--
            Write-Log "Paging Prev. PageIndex=$($Script:PageIndex)" "INFO"
            Bind-GridFromCurrentView
            Reset-SelectionAndDetails
        }
    })

    $btnNext.Add_Click({
        $totalPages = Get-TotalPages -Items $Script:CurrentView -PageSize $Script:PageSize
        if ($Script:PageIndex -lt ($totalPages - 1)) {
            $Script:PageIndex++
            Write-Log "Paging Next. PageIndex=$($Script:PageIndex)" "INFO"
            Bind-GridFromCurrentView
            Reset-SelectionAndDetails
        }
    })

    $btnLoadAll.Add_Click({
        if (-not $Script:IsConnected) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Load ALL mailboxes into local cache?`n`nThis enables fast browsing and searching.",
            "Load all mailboxes",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $lblStatus.Text = "Loading all mailboxes... (this may take a moment)"
            $form.Refresh()

            Write-Log "LoadAll started." "INFO"

            $mode = Get-MailboxCmdletMode
            $raw = @()

            if ($mode -eq "EXO") {
                # EXO cmdlet: ResultSize Unlimited supported
                $raw = @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum -ErrorAction Stop)
            } else {
                $raw = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)
            }

            $Script:MailboxCache = @($raw | ForEach-Object { Convert-ToGridRow $_ })
            $Script:CacheLoaded = $true

            Reset-ViewToCache
            Bind-GridFromCurrentView

            $btnClearSearch.Enabled = $true
            $btnPrev.Enabled = $true
            $btnNext.Enabled = $true

            $lblStatus.Text = "Loaded $($Script:MailboxCache.Count) mailboxes. Use paging + search."
            Write-Log "LoadAll completed. Count=$($Script:MailboxCache.Count)" "INFO"
            Reset-SelectionAndDetails
        } catch {
            Write-Log "LoadAll failed: $($_.Exception.Message)" "ERROR"
            $lblStatus.Text = "Load all failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "Load all mailboxes failed.`n`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $btnSearch.Add_Click({
        if (-not $Script:IsConnected) { return }

        $q = $txtSearch.Text
        $qTrim = ($q ?? "").Trim()
        Write-Log "Search clicked. Query='$qTrim' CacheLoaded=$($Script:CacheLoaded)" "INFO"

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            if ($Script:CacheLoaded) {
                Apply-SearchToCache -QueryText $qTrim
                Bind-GridFromCurrentView
                $lblStatus.Text = if ([string]::IsNullOrWhiteSpace($qTrim)) {
                    "Showing all cached mailboxes."
                } else {
                    "Filtered cached mailboxes by: '$qTrim'"
                }
                Reset-SelectionAndDetails
                $btnClearSearch.Enabled = $true
            } else {
                # Best-effort server-side search (limited), recommend Load All for reliability
                if ([string]::IsNullOrWhiteSpace($qTrim)) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Tip: Click 'Load all mailboxes' to browse/search reliably.`n`nIf you want server-side search, type a query first.",
                        "Search",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    ) | Out-Null
                    return
                }

                $safe = Escape-OPathValue $qTrim
                $filter = "DisplayName -like '*$safe*' -or Alias -like '*$safe*' -or PrimarySmtpAddress -like '*$safe*'"

                $raw = @(Get-MailboxesServerSide -Max 200 -Filter $filter)
                $rows = @($raw | ForEach-Object { Convert-ToGridRow $_ })

                # Use view paging even for server-side results
                $Script:CurrentView = $rows
                $Script:PageIndex = 0
                Bind-GridFromCurrentView

                $lblStatus.Text = "Server-side search returned $($rows.Count) results (max 200). Tip: Load all for full browsing."
                Reset-SelectionAndDetails
                $btnClearSearch.Enabled = $true
            }
        } catch {
            Write-Log "Search failed: $($_.Exception.Message)" "ERROR"
            $lblStatus.Text = "Search failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "Search failed.`n`n$($_.Exception.Message)`n`nTip: 'Load all mailboxes' for reliable search/browse.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $btnClearSearch.Add_Click({
        $txtSearch.Text = ""
        $Script:PageIndex = 0
        Write-Log "ClearSearch clicked. CacheLoaded=$($Script:CacheLoaded)" "INFO"

        if ($Script:CacheLoaded) {
            Reset-ViewToCache
            Bind-GridFromCurrentView
            $lblStatus.Text = "Showing all cached mailboxes."
        } else {
            $Script:CurrentView = @()
            $grid.DataSource = $null
            Update-PagingUI
            $lblStatus.Text = "Cleared results. Tip: Load all mailboxes to browse."
        }
        Reset-SelectionAndDetails
    })

    $grid.Add_SelectionChanged({
        try {
            if ($grid.SelectedRows.Count -gt 0) {
                $row = $grid.SelectedRows[0]
                $smtp = $row.Cells["PrimarySmtpAddress"].Value
                if ($smtp) {
                    $Script:SelectedIdentity = $smtp.ToString()
                    Write-Log "Selection changed. SelectedIdentity='$($Script:SelectedIdentity)'" "INFO"
                    Show-Details -Identity $Script:SelectedIdentity
                }
            }
        } catch {
            Write-Log "SelectionChanged warning: $($_.Exception.Message)" "WARN"
        }
    })

    $btnRefresh.Add_Click({
        if (-not $Script:SelectedIdentity) { return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Write-Log "Refresh clicked. Identity='$($Script:SelectedIdentity)'" "INFO"
            Show-Details -Identity $Script:SelectedIdentity
        } catch {
            Write-Log "Refresh failed: $($_.Exception.Message)" "ERROR"
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

    $btnExportSOA.Add_Click({
        if (-not $Script:SelectedIdentity) { return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Write-Log "Export current mailbox SOA settings clicked. Identity='$($Script:SelectedIdentity)'" "INFO"
            $path = Export-MailboxSOASettings -Identity $Script:SelectedIdentity
            [System.Windows.Forms.MessageBox]::Show(
                "Exported CURRENT mailbox SOA settings to:`n$path",
                "Export Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } catch {
            Write-Log "Export failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Export failed.`n`n$($_.Exception.Message)",
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
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Enable cloud SOA cancelled by user. Identity='$($Script:SelectedIdentity)'" "INFO"
            return
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Write-Log "Enable cloud SOA initiated. Identity='$($Script:SelectedIdentity)'" "INFO"
            $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $true
            [System.Windows.Forms.MessageBox]::Show(
                $msg,
                "Done",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            Show-Details -Identity $Script:SelectedIdentity

            # Refresh cache row if loaded
            if ($Script:CacheLoaded) {
                $idx = $Script:MailboxCache.FindIndex({ param($x) $x.PrimarySmtpAddress -eq $Script:SelectedIdentity })
                if ($idx -ge 0) {
                    $updated = Convert-ToGridRow (Get-Mailbox -Identity $Script:SelectedIdentity -ErrorAction Stop)
                    $Script:MailboxCache[$idx] = $updated

                    # Re-apply current search view if needed
                    if ([string]::IsNullOrWhiteSpace($txtSearch.Text)) {
                        Reset-ViewToCache
                    } else {
                        Apply-SearchToCache -QueryText $txtSearch.Text
                    }
                    Bind-GridFromCurrentView
                }
            }
        } catch {
            Write-Log "Enable cloud SOA failed: $($_.Exception.Message)" "ERROR"
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
            "Revert SOA back to on-prem for:`n`n$($Script:SelectedIdentity)`n`nThis sets IsExchangeCloudManaged = FALSE.",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Revert to on-prem SOA cancelled by user. Identity='$($Script:SelectedIdentity)'" "INFO"
            return
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Write-Log "Revert to on-prem SOA initiated. Identity='$($Script:SelectedIdentity)'" "INFO"

            $exportPrompt = [System.Windows.Forms.MessageBox]::Show(
                "Do you want to export CURRENT mailbox SOA settings (JSON) before reverting?",
                "Export recommended",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($exportPrompt -eq [System.Windows.Forms.DialogResult]::Yes) {
                $path = Export-MailboxSOASettings -Identity $Script:SelectedIdentity
                Write-Log "Export created before revert. Identity='$($Script:SelectedIdentity)' Path='$path'" "INFO"
            } else {
                Write-Log "Export skipped before revert. Identity='$($Script:SelectedIdentity)'" "WARN"
            }

            $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $false
            [System.Windows.Forms.MessageBox]::Show(
                $msg,
                "Done",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            Show-Details -Identity $Script:SelectedIdentity

            # Refresh cache row if loaded
            if ($Script:CacheLoaded) {
                $idx = $Script:MailboxCache.FindIndex({ param($x) $x.PrimarySmtpAddress -eq $Script:SelectedIdentity })
                if ($idx -ge 0) {
                    $updated = Convert-ToGridRow (Get-Mailbox -Identity $Script:SelectedIdentity -ErrorAction Stop)
                    $Script:MailboxCache[$idx] = $updated

                    if ([string]::IsNullOrWhiteSpace($txtSearch.Text)) {
                        Reset-ViewToCache
                    } else {
                        Apply-SearchToCache -QueryText $txtSearch.Text
                    }
                    Bind-GridFromCurrentView
                }
            }
        } catch {
            Write-Log "Revert to on-prem SOA failed: $($_.Exception.Message)" "ERROR"
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
        Write-Log "Application closing requested." "INFO"
        try { Disconnect-EXO } catch { }
        Write-Log "Application closed." "INFO"
    })

    # Init + Run
    Set-UiConnectedState -Connected $false
    Write-Log "$($Script:ToolName) GUI starting (Application.Run)..." "INFO"
    [System.Windows.Forms.Application]::Run($form)

} catch {
    Write-Log "FATAL: GUI failed to start. Error=$($_.Exception.Message)" "ERROR"
    [System.Windows.Forms.MessageBox]::Show(
        "GUI failed to start.`n`n$($_.Exception.Message)`n`nCheck log:`n$($Script:LogFile)",
        $Script:ToolName,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}
#endregion
