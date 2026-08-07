#Requires -Version 5.1
# =============================================================================
#  JumpCloud MTP Manager - GUI Tool
#  Altec Solutions Group, Inc.
# =============================================================================
# No default -- do not hardcode a live API key here. Paste it into the GUI
# field at runtime, or set the JC_API_KEY environment variable to prefill it
# (e.g. as a NinjaRMM script/environment variable) without committing it.
$DefaultApiKey = if ($env:JC_API_KEY) { $env:JC_API_KEY } else { "" }
# =============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$C = @{
    NavyDark  = [System.Drawing.Color]::FromArgb(15,  30,  60)
    NavyMid   = [System.Drawing.Color]::FromArgb(22,  40,  73)
    NavyLight = [System.Drawing.Color]::FromArgb(40,  60, 100)
    Green     = [System.Drawing.Color]::FromArgb(0,  163, 107)
    White     = [System.Drawing.Color]::White
    BgLight   = [System.Drawing.Color]::FromArgb(248, 250, 252)
    Border    = [System.Drawing.Color]::FromArgb(226, 232, 240)
    HeaderBg  = [System.Drawing.Color]::FromArgb(241, 245, 249)
    HeaderFg  = [System.Drawing.Color]::FromArgb(51,  65,  85)
    SubText   = [System.Drawing.Color]::FromArgb(100, 116, 139)
    SelBg     = [System.Drawing.Color]::FromArgb(219, 234, 254)
    SelFg     = [System.Drawing.Color]::FromArgb(30,  64, 175)
    Accent    = [System.Drawing.Color]::FromArgb(30,  64, 175)
    DangerBg  = [System.Drawing.Color]::FromArgb(254, 226, 226)
    DangerFg  = [System.Drawing.Color]::FromArgb(153,  27,  27)
    DangerBrd = [System.Drawing.Color]::FromArgb(252, 165, 165)
    WarnBg    = [System.Drawing.Color]::FromArgb(254, 243, 199)
    WarnFg    = [System.Drawing.Color]::FromArgb(120,  53,  15)
    WarnBrd   = [System.Drawing.Color]::FromArgb(253, 186, 116)
}

$script:ApiKey            = ""
$script:SelectedOrgId     = ""
$script:SelectedOrgName   = ""
$script:SelectedUserIds   = @()
$script:SelectedDeviceIds = @()

# =============================================================================
#  API HELPERS
# =============================================================================
function Get-JCHeaders {
    param([string]$OrgId = "")
    $h = @{
        "x-api-key"    = $script:ApiKey
        "Content-Type" = "application/json"
        "Accept"       = "application/json"
    }
    if ($OrgId) { $h["x-org-id"] = $OrgId }
    return $h
}

function Invoke-JCApi {
    param(
        [string]$Uri,
        [string]$Method  = "GET",
        [hashtable]$Headers,
        [string]$Body    = $null
    )
    try {
        $p = @{ Uri = $Uri; Method = $Method; Headers = $Headers; ErrorAction = "Stop" }
        if ($Body) { $p.Body = $Body }
        return Invoke-RestMethod @p
    }
    catch {
        $msg = $_.Exception.Message
        try {
            $stream = $_.Exception.Response.GetResponseStream()
            $msg   += "`n" + (New-Object System.IO.StreamReader($stream)).ReadToEnd()
        } catch {}
        throw $msg
    }
}

# Handles both wrapped { totalCount, results[] } and raw array responses
function Get-AllPages {
    param([string]$BaseUri, [hashtable]$Headers)
    $all   = [System.Collections.Generic.List[object]]::new()
    $skip  = 0
    $limit = 100
    do {
        $page    = Invoke-JCApi -Uri "$BaseUri`?limit=$limit&skip=$skip" -Headers $Headers
        $results = @()

        if ($page -is [System.Array]) {
            # Raw array response (some v1 endpoints, all v2 endpoints)
            $results = $page
            $total   = $page.Count
        }
        elseif ($null -ne $page -and
                $page.PSObject.Properties.Name -contains 'results') {
            # Wrapped { totalCount, results[] } response
            $results = @($page.results)
            $total   = if ($page.PSObject.Properties.Name -contains 'totalCount') {
                           [int]$page.totalCount
                       } else { $results.Count }
        }
        else {
            # Single object returned
            if ($null -ne $page) { $results = @($page) }
            $total = $results.Count
        }

        foreach ($r in $results) { if ($null -ne $r) { $all.Add($r) } }
        $skip += $limit
    } while ($all.Count -lt $total -and $results.Count -eq $limit)
    return $all
}

# =============================================================================
#  WINFORMS HELPERS
# =============================================================================
function New-StyledDGV {
    param([bool]$MultiSelect = $true)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock                  = "Fill"
    $g.ReadOnly              = $true
    $g.SelectionMode         = "FullRowSelect"
    $g.MultiSelect           = $MultiSelect
    $g.AllowUserToAddRows    = $false
    $g.AllowUserToDeleteRows = $false
    $g.RowHeadersVisible     = $false
    $g.BackgroundColor       = $C.White
    $g.GridColor             = $C.Border
    $g.BorderStyle           = "None"
    $g.AutoSizeColumnsMode   = "Fill"
    $g.CellBorderStyle       = "SingleHorizontal"
    $g.RowTemplate.Height         = 32
    $g.ColumnHeadersHeight        = 34
    $g.ColumnHeadersVisible       = $true
    # Let Windows render headers in its native visual style - always visible,
    # never wiped out by Columns.Clear()/rebuild cycles.
    $g.EnableHeadersVisualStyles  = $true
    $g.DefaultCellStyle.SelectionBackColor = $C.SelBg
    $g.DefaultCellStyle.SelectionForeColor = $C.SelFg
    $g.DefaultCellStyle.Padding            = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
    return $g
}

function New-ActionButton {
    param([string]$Text, [int]$X, [int]$Width = 155,
          $BgColor, $FgColor, $BrdColor, [bool]$Enabled = $true)
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Size      = New-Object System.Drawing.Size($Width, 30)
    $b.Location  = New-Object System.Drawing.Point($X, 8)
    $b.Enabled   = $Enabled
    $b.FlatStyle = "Flat"
    $b.BackColor = $BgColor
    $b.ForeColor = $FgColor
    $b.FlatAppearance.BorderColor = $BrdColor
    $b.FlatAppearance.BorderSize  = 1
    $b.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    return $b
}

function Set-Status {
    param([string]$Msg, [string]$Type = "normal")
    $script:statusLabel.Text = "  $Msg"
    $script:statusLabel.ForeColor = switch ($Type) {
        "ok"   { [System.Drawing.Color]::FromArgb(0, 130, 60)    }
        "err"  { [System.Drawing.Color]::FromArgb(180, 0, 0)     }
        "warn" { [System.Drawing.Color]::FromArgb(140, 80, 0)    }
        default{ [System.Drawing.Color]::FromArgb(160, 180, 210) }
    }
    $script:form.Refresh()
}

# Get selected IDs from a DGV using three fallback methods
function Get-SelectedIds {
    param([System.Windows.Forms.DataGridView]$Grid, [string[]]$Cached)
    # Method 1: cached selection (most reliable - immune to focus loss)
    if ($Cached.Count -gt 0) { return $Cached }
    # Method 2: SelectedRows collection
    $ids = @($Grid.SelectedRows | ForEach-Object { $_.Cells["_id"].Value } | Where-Object { $_ })
    if ($ids.Count -gt 0) { return $ids }
    # Method 3: walk every row checking .Selected property
    $ids = @($Grid.Rows | Where-Object { $_.Selected } | ForEach-Object { $_.Cells["_id"].Value } | Where-Object { $_ })
    return $ids
}

function Get-NameMap {
    param([System.Windows.Forms.DataGridView]$Grid, [string]$NameCol)
    $map = [System.Collections.Generic.Dictionary[string,string]]::new()
    foreach ($row in $Grid.Rows) {
        $id = $row.Cells["_id"].Value
        if ($id) { $map[$id] = $row.Cells[$NameCol].Value }
    }
    return $map
}

# =============================================================================
#  MAIN FORM
# =============================================================================
$script:form = New-Object System.Windows.Forms.Form
$script:form.Text          = "JumpCloud MTP Manager"
$script:form.Size          = New-Object System.Drawing.Size(1120, 700)
$script:form.MinimumSize   = New-Object System.Drawing.Size(960, 580)
$script:form.StartPosition = "CenterScreen"
$script:form.BackColor     = $C.HeaderBg
$script:form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

# Header
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock      = "Top"
$headerPanel.Height    = 58
$headerPanel.BackColor = $C.NavyMid
$script:form.Controls.Add($headerPanel)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = "JumpCloud MTP Manager"
$lblTitle.ForeColor = $C.White
$lblTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize  = $true
$lblTitle.Location  = New-Object System.Drawing.Point(16, 14)
$headerPanel.Controls.Add($lblTitle)

$lblApiLbl = New-Object System.Windows.Forms.Label
$lblApiLbl.Text      = "MTP API Key:"
$lblApiLbl.ForeColor = [System.Drawing.Color]::FromArgb(180, 200, 230)
$lblApiLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblApiLbl.AutoSize  = $true
$lblApiLbl.Location  = New-Object System.Drawing.Point(288, 20)
$headerPanel.Controls.Add($lblApiLbl)

$script:txtApiKey = New-Object System.Windows.Forms.TextBox
$script:txtApiKey.Size         = New-Object System.Drawing.Size(350, 24)
$script:txtApiKey.Location     = New-Object System.Drawing.Point(375, 17)
$script:txtApiKey.PasswordChar = [char]42
$script:txtApiKey.BackColor    = $C.NavyLight
$script:txtApiKey.ForeColor    = $C.White
$script:txtApiKey.BorderStyle  = "FixedSingle"
$script:txtApiKey.Font         = New-Object System.Drawing.Font("Segoe UI", 9)
$headerPanel.Controls.Add($script:txtApiKey)

$chkShowKey = New-Object System.Windows.Forms.CheckBox
$chkShowKey.Text      = "Show"
$chkShowKey.ForeColor = [System.Drawing.Color]::FromArgb(180, 200, 230)
$chkShowKey.AutoSize  = $true
$chkShowKey.Location  = New-Object System.Drawing.Point(732, 20)
$headerPanel.Controls.Add($chkShowKey)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text      = "Connect"
$btnConnect.Size      = New-Object System.Drawing.Size(100, 30)
$btnConnect.Location  = New-Object System.Drawing.Point(795, 14)
$btnConnect.BackColor = $C.Green
$btnConnect.ForeColor = $C.White
$btnConnect.FlatStyle = "Flat"
$btnConnect.FlatAppearance.BorderSize = 0
$btnConnect.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$headerPanel.Controls.Add($btnConnect)

$script:btnRefresh = New-Object System.Windows.Forms.Button
$script:btnRefresh.Text      = "Refresh"
$script:btnRefresh.Size      = New-Object System.Drawing.Size(90, 30)
$script:btnRefresh.Location  = New-Object System.Drawing.Point(903, 14)
$script:btnRefresh.BackColor = $C.NavyLight
$script:btnRefresh.ForeColor = [System.Drawing.Color]::FromArgb(180, 210, 255)
$script:btnRefresh.FlatStyle = "Flat"
$script:btnRefresh.FlatAppearance.BorderSize = 0
$script:btnRefresh.Enabled   = $false
$headerPanel.Controls.Add($script:btnRefresh)

# Status bar
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor  = $C.NavyDark
$statusStrip.SizingGrip = $false
$script:statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:statusLabel.Text      = "  Enter your MTP API key above and click Connect."
$script:statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(160, 180, 210)
$script:statusLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
$statusStrip.Items.Add($script:statusLabel) | Out-Null
$script:form.Controls.Add($statusStrip)

# Split container
$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Dock             = "Fill"
$splitMain.SplitterDistance = 340
$splitMain.SplitterWidth    = 5
$splitMain.BackColor        = $C.Border
$script:form.Controls.Add($splitMain)

# =============================================================================
#  LEFT PANE - Organizations  (Top -> Bottom -> Fill)
# =============================================================================
$leftOuter = New-Object System.Windows.Forms.Panel
$leftOuter.Dock = "Fill"; $leftOuter.BackColor = $C.White
$splitMain.Panel1.Controls.Add($leftOuter)

$leftHeader = New-Object System.Windows.Forms.Panel
$leftHeader.Dock = "Top"; $leftHeader.Height = 38; $leftHeader.BackColor = $C.HeaderBg
$leftOuter.Controls.Add($leftHeader)

$lblOrgsHeader = New-Object System.Windows.Forms.Label
$lblOrgsHeader.Text = "  ORGANIZATIONS"
$lblOrgsHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblOrgsHeader.ForeColor = $C.SubText; $lblOrgsHeader.Dock = "Fill"; $lblOrgsHeader.TextAlign = "MiddleLeft"
$leftHeader.Controls.Add($lblOrgsHeader)

$script:lblOrgCount = New-Object System.Windows.Forms.Label
$script:lblOrgCount.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$script:lblOrgCount.ForeColor = $C.SubText; $script:lblOrgCount.Dock = "Right"
$script:lblOrgCount.Width = 80; $script:lblOrgCount.TextAlign = "MiddleRight"
$leftHeader.Controls.Add($script:lblOrgCount)

$leftBottom = New-Object System.Windows.Forms.Panel
$leftBottom.Dock = "Bottom"; $leftBottom.Height = 48; $leftBottom.BackColor = $C.BgLight
$leftOuter.Controls.Add($leftBottom)

$script:btnDeleteOrg = New-ActionButton "Delete Org Workflow" 8 200 $C.DangerBg $C.DangerFg $C.DangerBrd $false
$leftBottom.Controls.Add($script:btnDeleteOrg)

$script:dgvOrgs = New-StyledDGV $false          # Fill - added LAST
$leftOuter.Controls.Add($script:dgvOrgs)

# =============================================================================
#  RIGHT PANE  (Top -> Fill)
# =============================================================================
$rightOuter = New-Object System.Windows.Forms.Panel
$rightOuter.Dock = "Fill"; $rightOuter.BackColor = $C.White
$splitMain.Panel2.Controls.Add($rightOuter)

$rightHeader = New-Object System.Windows.Forms.Panel
$rightHeader.Dock = "Top"; $rightHeader.Height = 38; $rightHeader.BackColor = $C.HeaderBg
$rightOuter.Controls.Add($rightHeader)

$script:lblOrgContext = New-Object System.Windows.Forms.Label
$script:lblOrgContext.Text = "  Select an organization from the left panel."
$script:lblOrgContext.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$script:lblOrgContext.ForeColor = $C.SubText; $script:lblOrgContext.Dock = "Fill"; $script:lblOrgContext.TextAlign = "MiddleLeft"
$rightHeader.Controls.Add($script:lblOrgContext)

$tabCtrl = New-Object System.Windows.Forms.TabControl
$tabCtrl.Dock = "Fill"; $tabCtrl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$rightOuter.Controls.Add($tabCtrl)                # Fill - added LAST

# =============================================================================
#  USERS TAB  (Bottom -> Fill)
# =============================================================================
$tabUsers = New-Object System.Windows.Forms.TabPage
$tabUsers.Text = "  Users  "; $tabUsers.BackColor = $C.White
$tabUsers.Padding = New-Object System.Windows.Forms.Padding(0)
$tabCtrl.TabPages.Add($tabUsers)

$userBottom = New-Object System.Windows.Forms.Panel
$userBottom.Dock = "Bottom"; $userBottom.Height = 48; $userBottom.BackColor = $C.BgLight
$tabUsers.Controls.Add($userBottom)                # Bottom - added FIRST

$script:btnDelSelUsers   = New-ActionButton "Delete Selected"        8   155 $C.DangerBg $C.DangerFg $C.DangerBrd $false
$script:btnDelAllUsers   = New-ActionButton "Delete ALL Users"       172  155 $C.WarnBg   $C.WarnFg   $C.WarnBrd   $false
$script:btnForceDelUsers = New-ActionButton "Force Delete (AD/LDAP)" 336  180 `
    ([System.Drawing.Color]::FromArgb(239, 68, 68)) `
    $C.White `
    ([System.Drawing.Color]::FromArgb(220, 38, 38)) $true   # always enabled
$userBottom.Controls.Add($script:btnDelSelUsers)
$userBottom.Controls.Add($script:btnDelAllUsers)
$userBottom.Controls.Add($script:btnForceDelUsers)

# Count label anchored to the right - no overlap with buttons
$script:lblUserCount = New-Object System.Windows.Forms.Label
$script:lblUserCount.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$script:lblUserCount.ForeColor = $C.SubText
$script:lblUserCount.Dock      = "Right"
$script:lblUserCount.Width     = 120
$script:lblUserCount.TextAlign = "MiddleRight"
$userBottom.Controls.Add($script:lblUserCount)

$script:dgvUsers = New-StyledDGV $true             # Fill - added LAST
$tabUsers.Controls.Add($script:dgvUsers)

# =============================================================================
#  DEVICES TAB  (Bottom -> Fill)
# =============================================================================
$tabDevices = New-Object System.Windows.Forms.TabPage
$tabDevices.Text = "  Devices  "; $tabDevices.BackColor = $C.White
$tabDevices.Padding = New-Object System.Windows.Forms.Padding(0)
$tabCtrl.TabPages.Add($tabDevices)

$devBottom = New-Object System.Windows.Forms.Panel
$devBottom.Dock = "Bottom"; $devBottom.Height = 48; $devBottom.BackColor = $C.BgLight
$tabDevices.Controls.Add($devBottom)               # Bottom - added FIRST

$script:lblDevCount = New-Object System.Windows.Forms.Label
$script:lblDevCount.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$script:lblDevCount.ForeColor = $C.SubText
$script:lblDevCount.Dock      = "Right"
$script:lblDevCount.Width     = 120
$script:lblDevCount.TextAlign = "MiddleRight"
$devBottom.Controls.Add($script:lblDevCount)

$script:btnDelSelDevices  = New-ActionButton "Delete Selected"    8  155 $C.DangerBg $C.DangerFg $C.DangerBrd $false
$script:btnDelAllDevices  = New-ActionButton "Delete ALL Devices" 172 155 $C.WarnBg   $C.WarnFg   $C.WarnBrd   $false
$devBottom.Controls.Add($script:btnDelSelDevices)
$devBottom.Controls.Add($script:btnDelAllDevices)

# "No devices" message shown when the org has no JumpCloud-managed devices
$script:lblNoDevices = New-Object System.Windows.Forms.Label
$script:lblNoDevices.Dock      = "Fill"
$script:lblNoDevices.TextAlign = "MiddleCenter"
$script:lblNoDevices.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$script:lblNoDevices.ForeColor = $C.SubText
$script:lblNoDevices.BackColor = $C.White
$script:lblNoDevices.Visible   = $false
$script:lblNoDevices.Text      = "No JumpCloud-managed devices found in this organization." + [System.Environment]::NewLine +
                                  "Devices only appear here if the JumpCloud agent is installed on them." + [System.Environment]::NewLine +
                                  "AD-synced users do not automatically register devices."
$tabDevices.Controls.Add($script:lblNoDevices)      # added before DGV so DGV is on top when visible

$script:dgvDevices = New-StyledDGV $true           # Fill - added LAST
$tabDevices.Controls.Add($script:dgvDevices)

# =============================================================================
#  DATA LOADING
# =============================================================================
function Load-Organizations {
    $script:dgvOrgs.Rows.Clear(); $script:dgvOrgs.Columns.Clear()
    Set-Status "Loading organizations..." "warn"
    try {
        $orgs = Get-AllPages "https://console.jumpcloud.com/api/organizations" (Get-JCHeaders)
        $script:dgvOrgs.Columns.Add("_id",        "ID")                | Out-Null
        $script:dgvOrgs.Columns.Add("displayName", "Organization Name") | Out-Null
        $script:dgvOrgs.Columns.Add("numUsers",    "Users")             | Out-Null
        $script:dgvOrgs.Columns["_id"].Visible            = $false
        $script:dgvOrgs.Columns["displayName"].FillWeight = 75
        $script:dgvOrgs.Columns["numUsers"].FillWeight    = 25
        foreach ($o in $orgs) { $script:dgvOrgs.Rows.Add($o._id, $o.displayName, $o.numUsers) | Out-Null }
        $script:lblOrgCount.Text   = "  $($orgs.Count) orgs  "
        $script:btnRefresh.Enabled = $true
        Set-Status "Loaded $($orgs.Count) organization(s). Click an org to load its users and devices." "ok"
    }
    catch {
        Set-Status "Connection error: $_" "err"
        [System.Windows.Forms.MessageBox]::Show("Could not connect:`n$_", "Connection Failed", "OK", "Error") | Out-Null
    }
}

function Load-Users {
    param([string]$OrgId)
    $script:dgvUsers.Rows.Clear(); $script:dgvUsers.Columns.Clear()
    $script:SelectedUserIds = @()
    $script:lblUserCount.Text = "Loading..."
    $script:btnDelSelUsers.Enabled = $false
    $script:btnDelAllUsers.Enabled = $false
    # Force Delete is always left enabled - it handles an empty grid gracefully
    Set-Status "Loading users for $script:SelectedOrgName..." "warn"
    try {
        $users = Get-AllPages "https://console.jumpcloud.com/api/systemusers" (Get-JCHeaders -OrgId $OrgId)
        $script:dgvUsers.Columns.Add("_id",           "ID")          | Out-Null
        $script:dgvUsers.Columns.Add("username",       "Username")    | Out-Null
        $script:dgvUsers.Columns.Add("firstname",      "First Name")  | Out-Null
        $script:dgvUsers.Columns.Add("lastname",       "Last Name")   | Out-Null
        $script:dgvUsers.Columns.Add("email",          "Email")       | Out-Null
        $script:dgvUsers.Columns.Add("ldapUser",       "AD/LDAP")     | Out-Null
        $script:dgvUsers.Columns.Add("account_locked", "Locked")      | Out-Null
        $script:dgvUsers.Columns["_id"].Visible              = $false
        $script:dgvUsers.Columns["username"].FillWeight       = 18
        $script:dgvUsers.Columns["firstname"].FillWeight      = 14
        $script:dgvUsers.Columns["lastname"].FillWeight       = 14
        $script:dgvUsers.Columns["email"].FillWeight          = 34
        $script:dgvUsers.Columns["ldapUser"].FillWeight       = 10
        $script:dgvUsers.Columns["account_locked"].FillWeight = 10
        foreach ($u in $users) {
            $isLdap  = if ($u.ldap_binding_user -or $u.externallyManaged) { "Yes" } else { "No" }
            $locked  = if ($u.account_locked) { "Yes" } else { "No" }
            $script:dgvUsers.Rows.Add($u._id, $u.username, $u.firstname, $u.lastname, $u.email, $isLdap, $locked) | Out-Null
        }
        $cnt = $users.Count
        $script:lblUserCount.Text      = "$cnt user(s)"
        $script:btnDelSelUsers.Enabled = $cnt -gt 0
        $script:btnDelAllUsers.Enabled = $cnt -gt 0
        Set-Status "Loaded $cnt user(s) in '$script:SelectedOrgName'." "ok"
    }
    catch {
        $script:lblUserCount.Text = "Error"
        Set-Status "Error loading users: $_" "err"
        [System.Windows.Forms.MessageBox]::Show("Error loading users:`n$_", "Error", "OK", "Error") | Out-Null
    }
}

function Load-Devices {
    param([string]$OrgId)
    $script:dgvDevices.Rows.Clear(); $script:dgvDevices.Columns.Clear()
    $script:SelectedDeviceIds = @()
    $script:lblDevCount.Text  = "Loading..."
    $script:lblNoDevices.Visible = $false
    $script:dgvDevices.Visible   = $true
    $script:btnDelSelDevices.Enabled = $false; $script:btnDelAllDevices.Enabled = $false
    Set-Status "Loading devices for $script:SelectedOrgName..." "warn"

    $devs  = $null
    $errV1 = ""
    $errV2 = ""

    # Attempt 1: v1 API
    try {
        $devs = Get-AllPages "https://console.jumpcloud.com/api/systems" (Get-JCHeaders -OrgId $OrgId)
    }
    catch { $errV1 = "$_"; $devs = $null }

    # Attempt 2: v2 API if v1 returned nothing or errored
    if ($null -eq $devs -or $devs.Count -eq 0) {
        try {
            $devs = Get-AllPages "https://console.jumpcloud.com/api/v2/systems" (Get-JCHeaders -OrgId $OrgId)
        }
        catch { $errV2 = "$_" }
    }

    if ($errV1 -and $errV2) {
        $script:lblDevCount.Text = "API Error"
        Set-Status "Device API error - see popup for details." "err"
        [System.Windows.Forms.MessageBox]::Show(
            "Both device API endpoints failed for '$script:SelectedOrgName'.`n`n" +
            "v1 error: $errV1`n`nv2 error: $errV2`n`n" +
            "Note: If this org has no JumpCloud agent installed on any device, 0 results is expected.",
            "Device Load Error", "OK", "Warning"
        ) | Out-Null
        return
    }

    if ($null -eq $devs) { $devs = @() }

    $script:dgvDevices.Columns.Add("_id",        "ID")           | Out-Null
    $script:dgvDevices.Columns.Add("displayName", "Device Name")  | Out-Null
    $script:dgvDevices.Columns.Add("os",          "OS")           | Out-Null
    $script:dgvDevices.Columns.Add("osVersion",   "Version")      | Out-Null
    $script:dgvDevices.Columns.Add("lastContact", "Last Contact") | Out-Null
    $script:dgvDevices.Columns.Add("active",      "Active")       | Out-Null
    $script:dgvDevices.Columns["_id"].Visible             = $false
    $script:dgvDevices.Columns["displayName"].FillWeight  = 26
    $script:dgvDevices.Columns["os"].FillWeight           = 18
    $script:dgvDevices.Columns["osVersion"].FillWeight    = 16
    $script:dgvDevices.Columns["lastContact"].FillWeight  = 30
    $script:dgvDevices.Columns["active"].FillWeight       = 10

    foreach ($d in $devs) {
        # Use hostname as fallback if displayName is absent
        $dname = if ($d.displayName)  { $d.displayName }
                 elseif ($d.hostname) { $d.hostname }
                 else                 { $d._id }
        $lc = if ($d.lastContact) {
            try { ([datetime]$d.lastContact).ToString("yyyy-MM-dd  HH:mm") } catch { $d.lastContact }
        } else { "Never" }
        $active = if ($d.active) { "Yes" } else { "No" }
        $script:dgvDevices.Rows.Add($d._id, $dname, $d.os, $d.osVersion, $lc, $active) | Out-Null
    }

    $cnt = $devs.Count
    $script:lblDevCount.Text             = "$cnt device(s)"
    $script:btnDelSelDevices.Enabled     = $cnt -gt 0
    $script:btnDelAllDevices.Enabled     = $cnt -gt 0

    # Show overlay message when no devices; hide DGV (it's empty anyway)
    $script:lblNoDevices.Visible = $cnt -eq 0
    $script:dgvDevices.Visible   = $cnt -gt 0

    if ($cnt -eq 0) {
        Set-Status "No JumpCloud-managed devices found in '$script:SelectedOrgName'." "warn"
    } else {
        Set-Status "Loaded $cnt device(s) in '$script:SelectedOrgName'." "ok"
    }
}

# =============================================================================
#  DELETE HELPERS
# =============================================================================

# Standard delete - returns $true on success, error string on failure
function Remove-JCUser {
    param([string]$UserId, [hashtable]$Headers)
    try {
        Invoke-JCApi -Uri "https://console.jumpcloud.com/api/systemusers/$UserId" `
                     -Method DELETE -Headers $Headers
        return $true
    }
    catch { return "$_" }
}

# Force delete: suspend user first (clears managed-user locks), then delete
function Remove-JCUserForce {
    param([string]$UserId, [hashtable]$Headers)

    # Step 1: GET current user object so we can PUT it back with flags cleared.
    # JumpCloud blocks DELETE on externally-managed (AD-synced) users until the
    # externallyManaged and ldap_binding_user flags are cleared first.
    $userObj = $null
    try {
        $userObj = Invoke-JCApi -Uri "https://console.jumpcloud.com/api/systemusers/$UserId" `
                                -Method GET -Headers $Headers
    }
    catch { <# proceed without GET data #> }

    # Step 2: PUT back the full object with external-management cleared.
    if ($null -ne $userObj) {
        try {
            # Clear the flags that block deletion
            $userObj.externallyManaged = $false
            if ($userObj.PSObject.Properties.Name -contains 'ldap_binding_user') {
                $userObj.ldap_binding_user = $false
            }
            if ($userObj.PSObject.Properties.Name -contains 'externally_managed') {
                $userObj.externally_managed = $false
            }
            $putBody = $userObj | ConvertTo-Json -Depth 10 -Compress
            Invoke-JCApi -Uri "https://console.jumpcloud.com/api/systemusers/$UserId" `
                         -Method PUT -Headers $Headers -Body $putBody | Out-Null
        }
        catch { <# ignore PUT failure - still attempt DELETE #> }
    }

    # Step 3: DELETE. cascade_manager=null handles users assigned as managers.
    try {
        Invoke-JCApi -Uri "https://console.jumpcloud.com/api/systemusers/$UserId`?cascade_manager=null" `
                     -Method DELETE -Headers $Headers
        return $true
    }
    catch {
        # If DELETE still fails, try once more without the query param
        try {
            Invoke-JCApi -Uri "https://console.jumpcloud.com/api/systemusers/$UserId" `
                         -Method DELETE -Headers $Headers
            return $true
        }
        catch { return "$_" }
    }
}

function Invoke-UserDelete {
    param(
        [string[]]$Ids,
        [System.Collections.Generic.Dictionary[string,string]]$NameMap,
        [bool]$Force = $false
    )
    $headers = Get-JCHeaders -OrgId $script:SelectedOrgId
    $ok = 0; $fail = 0; $failNames = @()
    $total = $Ids.Count
    foreach ($id in $Ids) {
        $name = if ($NameMap.ContainsKey($id)) { $NameMap[$id] } else { $id }
        $res  = if ($Force) { Remove-JCUserForce $id $headers } else { Remove-JCUser $id $headers }
        if ($res -eq $true) {
            $ok++
            Set-Status "Deleted user: $name  ($ok / $total)" "ok"
        } else {
            $fail++
            $failNames += "$name  -  $res"
            Set-Status "Failed: $name" "err"
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    return [PSCustomObject]@{ Ok = $ok; Fail = $fail; FailNames = $failNames }
}

function Invoke-DeviceDelete {
    param([string[]]$Ids, [System.Collections.Generic.Dictionary[string,string]]$NameMap)
    $headers = Get-JCHeaders -OrgId $script:SelectedOrgId
    $ok = 0; $fail = 0; $total = $Ids.Count
    foreach ($id in $Ids) {
        $name = if ($NameMap.ContainsKey($id)) { $NameMap[$id] } else { $id }
        try {
            Invoke-JCApi -Uri "https://console.jumpcloud.com/api/systems/$id" `
                         -Method DELETE -Headers $headers
            $ok++
            Set-Status "Deleted device: $name  ($ok / $total)" "ok"
        }
        catch {
            $fail++
            Set-Status "Failed: $name - $_" "err"
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    return [PSCustomObject]@{ Ok = $ok; Fail = $fail }
}

function Show-DeleteResult {
    param([PSCustomObject]$Result, [string]$ItemType)
    $msg = "Deleted: $($Result.Ok)   |   Failed: $($Result.Fail)"
    if ($Result.FailNames.Count -gt 0) {
        $msg += "`n`nFailed items:`n" + ($Result.FailNames -join "`n")
    }
    [System.Windows.Forms.MessageBox]::Show($msg, "$ItemType Deletion Complete", "OK", "Information") | Out-Null
}

function Invoke-Connect {
    $key = $script:txtApiKey.Text.Trim()
    if (-not $key) {
        [System.Windows.Forms.MessageBox]::Show("Please enter your MTP API key.", "Required", "OK", "Warning") | Out-Null
        return
    }
    $script:ApiKey = $key
    Load-Organizations
}

# =============================================================================
#  EVENT HANDLERS
# =============================================================================
$chkShowKey.Add_CheckedChanged({
    $script:txtApiKey.PasswordChar = if ($chkShowKey.Checked) { [char]0 } else { [char]42 }
})

$btnConnect.Add_Click({ Invoke-Connect })
$script:txtApiKey.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { Invoke-Connect } })

$script:btnRefresh.Add_Click({
    Load-Organizations
    if ($script:SelectedOrgId) {
        Load-Users   -OrgId $script:SelectedOrgId
        Load-Devices -OrgId $script:SelectedOrgId
    }
})

# Org selection
$script:dgvOrgs.Add_SelectionChanged({
    if ($script:dgvOrgs.SelectedRows.Count -eq 0) { return }
    $row      = $script:dgvOrgs.SelectedRows[0]
    $newOrgId = $row.Cells["_id"].Value
    if ($newOrgId -eq $script:SelectedOrgId) { return }
    $script:SelectedOrgId   = $newOrgId
    $script:SelectedOrgName = $row.Cells["displayName"].Value
    $script:lblOrgContext.Text      = "  [ $script:SelectedOrgName ]"
    $script:lblOrgContext.ForeColor = $C.Accent
    $script:lblOrgContext.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:btnDeleteOrg.Enabled    = $true
    Load-Users   -OrgId $script:SelectedOrgId
    Load-Devices -OrgId $script:SelectedOrgId
})

# Cache selections as they happen (prevents focus-loss clearing SelectedRows)
$script:dgvUsers.Add_SelectionChanged({
    $script:SelectedUserIds = @(
        $script:dgvUsers.SelectedRows | ForEach-Object { $_.Cells["_id"].Value } | Where-Object { $_ }
    )
})
$script:dgvDevices.Add_SelectionChanged({
    $script:SelectedDeviceIds = @(
        $script:dgvDevices.SelectedRows | ForEach-Object { $_.Cells["_id"].Value } | Where-Object { $_ }
    )
})

# --- Delete Selected Users ----------------------------------------------------
$script:btnDelSelUsers.Add_Click({
    $ids = Get-SelectedIds -Grid $script:dgvUsers -Cached $script:SelectedUserIds
    if ($ids.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No users selected.`nClick a row to select it. Hold Ctrl to select multiple.",
            "Nothing Selected", "OK", "Information") | Out-Null; return
    }
    $res = [System.Windows.Forms.MessageBox]::Show(
        "Permanently delete $($ids.Count) user(s) from '$script:SelectedOrgName'?`nThis cannot be undone.",
        "Confirm", "YesNo", "Warning")
    if ($res -ne "Yes") { return }
    $map = Get-NameMap $script:dgvUsers "username"
    $r   = Invoke-UserDelete -Ids $ids -NameMap $map -Force $false
    Show-DeleteResult $r "User"
    Load-Users -OrgId $script:SelectedOrgId
})

# --- Force Delete (AD/LDAP) ---------------------------------------------------
# Works on selected rows. If nothing is selected, targets ALL users in the grid
# so a single-user org doesn't require a prior row click.
$script:btnForceDelUsers.Add_Click({
    $ids = Get-SelectedIds -Grid $script:dgvUsers -Cached $script:SelectedUserIds

    # Fallback: no selection - use every user visible in the grid
    if ($ids.Count -eq 0) {
        $ids = @($script:dgvUsers.Rows |
                 ForEach-Object { $_.Cells["_id"].Value } |
                 Where-Object   { $_ })
    }

    if ($ids.Count -eq 0) { return }   # grid is truly empty

    $scope = if ($ids.Count -eq $script:dgvUsers.Rows.Count) {
                 "ALL $($ids.Count)"
             } else {
                 "the $($ids.Count) selected"
             }

    $res = [System.Windows.Forms.MessageBox]::Show(
        "FORCE DELETE $scope user(s) from '$script:SelectedOrgName'?`n`n" +
        "Each account is suspended first to break the AD/LDAP management`n" +
        "lock, then permanently deleted.`n`n" +
        "NOTE: If the AD sync integration is still active the user(s) will`n" +
        "be re-created on the next sync. Disable the AD integration first`n" +
        "to permanently prevent that.`n`n" +
        "This cannot be undone.",
        "Confirm Force Delete", "YesNo", "Warning")
    if ($res -ne "Yes") { return }

    $map = Get-NameMap $script:dgvUsers "username"
    $r   = Invoke-UserDelete -Ids $ids -NameMap $map -Force $true
    Show-DeleteResult $r "User (Force)"
    Load-Users -OrgId $script:SelectedOrgId
})

# --- Delete ALL Users ---------------------------------------------------------
$script:btnDelAllUsers.Add_Click({
    $cnt = $script:dgvUsers.Rows.Count
    if ($cnt -eq 0) { return }
    $res = [System.Windows.Forms.MessageBox]::Show(
        "WARNING - Delete ALL $cnt user(s) from '$script:SelectedOrgName'?`nThis cannot be undone.",
        "Confirm Bulk Delete", "YesNo", "Warning")
    if ($res -ne "Yes") { return }
    $res2 = [System.Windows.Forms.MessageBox]::Show(
        "Final confirmation: delete all $cnt users?", "Final Confirmation", "YesNo", "Warning")
    if ($res2 -ne "Yes") { return }
    $ids = @($script:dgvUsers.Rows | ForEach-Object { $_.Cells["_id"].Value })
    $map = Get-NameMap $script:dgvUsers "username"
    # Use force mode for bulk so AD-synced users aren't skipped
    $r   = Invoke-UserDelete -Ids $ids -NameMap $map -Force $true
    Show-DeleteResult $r "User (Bulk)"
    Load-Users -OrgId $script:SelectedOrgId
})

# --- Delete Selected Devices --------------------------------------------------
$script:btnDelSelDevices.Add_Click({
    $ids = Get-SelectedIds -Grid $script:dgvDevices -Cached $script:SelectedDeviceIds
    if ($ids.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No devices selected.`nClick a row to select it. Hold Ctrl to select multiple.",
            "Nothing Selected", "OK", "Information") | Out-Null; return
    }
    $res = [System.Windows.Forms.MessageBox]::Show(
        "Permanently delete $($ids.Count) device(s) from '$script:SelectedOrgName'?`n" +
        "The JumpCloud agent will be uninstalled. This cannot be undone.",
        "Confirm", "YesNo", "Warning")
    if ($res -ne "Yes") { return }
    $map = Get-NameMap $script:dgvDevices "displayName"
    $r   = Invoke-DeviceDelete -Ids $ids -NameMap $map
    Show-DeleteResult $r "Device"
    Load-Devices -OrgId $script:SelectedOrgId
})

# --- Delete ALL Devices -------------------------------------------------------
$script:btnDelAllDevices.Add_Click({
    $cnt = $script:dgvDevices.Rows.Count
    if ($cnt -eq 0) { return }
    $res = [System.Windows.Forms.MessageBox]::Show(
        "WARNING - Delete ALL $cnt device(s) from '$script:SelectedOrgName'?`n" +
        "The JumpCloud agent will be uninstalled from all devices.",
        "Confirm Bulk Delete", "YesNo", "Warning")
    if ($res -ne "Yes") { return }
    $res2 = [System.Windows.Forms.MessageBox]::Show(
        "Final confirmation: delete all $cnt devices?", "Final Confirmation", "YesNo", "Warning")
    if ($res2 -ne "Yes") { return }
    $ids = @($script:dgvDevices.Rows | ForEach-Object { $_.Cells["_id"].Value })
    $map = Get-NameMap $script:dgvDevices "displayName"
    $r   = Invoke-DeviceDelete -Ids $ids -NameMap $map
    Show-DeleteResult $r "Device (Bulk)"
    Load-Devices -OrgId $script:SelectedOrgId
})

# --- Delete Org Workflow ------------------------------------------------------
$script:btnDeleteOrg.Add_Click({
    $uCnt = $script:dgvUsers.Rows.Count
    $dCnt = $script:dgvDevices.Rows.Count
    $msg  = "ORG DELETION WORKFLOW`nOrganization: $script:SelectedOrgName`n`n"
    $msg += "Inventory:  Users: $uCnt   Devices: $dCnt`n`n"
    $msg += "Step 1 (auto)  : Force-delete all $dCnt device(s)`n"
    $msg += "Step 2 (auto)  : Force-delete all $uCnt user(s)`n"
    $msg += "Step 3 (manual): Settings > Request To Delete in Admin Portal`n"
    $msg += "Step 4 (manual): JumpCloud Support confirms removal from MTP`n`n"
    $msg += "Proceed with Steps 1 and 2?"
    $res = [System.Windows.Forms.MessageBox]::Show($msg, "Org Deletion Workflow", "YesNo", "Warning")
    if ($res -ne "Yes") { return }

    if ($dCnt -gt 0) {
        Set-Status "[1/2] Deleting all devices..." "warn"
        $dIds = @($script:dgvDevices.Rows | ForEach-Object { $_.Cells["_id"].Value })
        $dMap = Get-NameMap $script:dgvDevices "displayName"
        Invoke-DeviceDelete -Ids $dIds -NameMap $dMap | Out-Null
        Load-Devices -OrgId $script:SelectedOrgId
    }
    if ($uCnt -gt 0) {
        Set-Status "[2/2] Force-deleting all users..." "warn"
        $uIds = @($script:dgvUsers.Rows | ForEach-Object { $_.Cells["_id"].Value })
        $uMap = Get-NameMap $script:dgvUsers "username"
        Invoke-UserDelete -Ids $uIds -NameMap $uMap -Force $true | Out-Null
        Load-Users -OrgId $script:SelectedOrgId
    }

    $done  = "Automated cleanup complete for '$script:SelectedOrgName'.`n`n"
    $done += "Manual steps still required:`n"
    $done += "1. Launch this org from the MTP`n"
    $done += "2. Settings > Request To Delete`n"
    $done += "3. Enter the Org ID and submit`n"
    $done += "4. JumpCloud Support will contact you to finalize."
    [System.Windows.Forms.MessageBox]::Show($done, "Cleanup Complete", "OK", "Information") | Out-Null
    Set-Status "Cleanup complete. Submit the org deletion request in the JumpCloud portal." "ok"
})

$script:form.Add_Shown({
    if ($DefaultApiKey -ne "") { $script:txtApiKey.Text = $DefaultApiKey; Invoke-Connect }
})

$script:form.ShowDialog() | Out-Null