<#
.SYNOPSIS
    Entra ID Inactive Licensed User Scanner / Offboarding Tool (Self-Contained)

.DESCRIPTION
    Single script. On every launch it installs/updates required modules,
    verifies M365 admin roles, connects to a client's Entra ID tenant, scans
    every licensed member (non-guest) user's sign-in activity, and flags any
    account with no interactive or non-interactive sign-in in the last N days
    (default 90). Always writes a CSV report of what it found.

    A WinForms grid then lets you check off which flagged accounts to remediate.
    For each selected account:
      1. Convert mailbox to Shared            (Exchange Online)
      2. Rename  DisplayName -> "Historical - <Name>"
                 UPN/email   -> "historical-<original-localpart>@<domain>"
      3. Remove all assigned M365 licenses

    Accounts synced from on-prem AD (OnPremisesSyncEnabled) are flagged in the
    grid - the rename step is skipped for those (a cloud-side rename would just
    get reverted by the next AAD Connect sync cycle; use Offboard-ADUser.ps1 /
    on-prem AD tools for the rename instead). Shared-mailbox conversion and
    license removal still proceed for synced accounts since both are managed
    entirely in the cloud regardless of directory sync.

    Reading sign-in activity requires the Entra tenant to have Azure AD
    Premium P1/P2 (signInActivity is null tenant-wide otherwise) and the Graph
    application/delegated permission AuditLog.Read.All - "Setup Auth" below
    provisions an app registration with everything this tool needs.

    GUI mode    : Run with no parameters
    Report mode : Supply -ReportOnly to scan, write the CSV, and exit headless
                  (NinjaRMM / scheduled task friendly - no remediation is ever
                  performed in this mode)

.PARAMETER InactiveDays
    Sign-in inactivity threshold in days. Default 90.

.PARAMETER AdminUPN
    UPN of the M365 admin for role verification.

.PARAMETER ReportOnly
    Scan and export the CSV report, then exit. No GUI, no remediation.

.PARAMETER SkipRoleCheck
    Skip M365 role verification.

.PARAMETER SkipModuleCheck
    Skip module install/update check.

.PARAMETER WhatIf
    Simulate remediation steps - no changes committed. (Also toggleable in the GUI.)

.EXAMPLE
    # GUI, 90-day default threshold
    .\Find-InactiveLicensedUsers.ps1

.EXAMPLE
    # Headless report for NinjaRMM / scheduled task, 60-day threshold
    .\Find-InactiveLicensedUsers.ps1 -ReportOnly -InactiveDays 60
#>

[CmdletBinding()]
param(
    [int]      $InactiveDays   = 90,
    [string]   $AdminUPN       = "",
    [switch]   $ReportOnly,
    [switch]   $SkipRoleCheck,
    [switch]   $SkipModuleCheck,
    [switch]   $WhatIf
)

# ============================================================
# ADMIN ELEVATION - auto-relaunch as Administrator if needed
# ============================================================
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    if ($scriptPath) {
        $argString = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        $shell = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'PowerShell.exe' }
        try {
            Start-Process $shell -ArgumentList $argString -Verb RunAs
        } catch {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                "Administrator privileges are required to run this tool.`n`n" +
                "Please right-click the script and select 'Run as Administrator'.",
                "Elevation Required", "OK", "Warning") | Out-Null
        }
    } else {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "Administrator privileges are required.`n`nPlease right-click and select 'Run as Administrator'.",
            "Elevation Required", "OK", "Warning") | Out-Null
    }
    exit
}

# ============================================================
# POWERSHELL 7 REQUIREMENT
# Same rationale as Offboard-ADUser.ps1: keeps Microsoft.Graph.Authentication
# and ExchangeOnlineManagement module assemblies properly isolated.
# ============================================================

if ($PSVersionTable.PSVersion.Major -lt 7) {
    $pwsh7 = "$env:ProgramFiles\PowerShell\7\pwsh.exe"

    if (-not (Test-Path $pwsh7)) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] PowerShell 7 required but not installed. Installing now..." -ForegroundColor Yellow
        $ps7Installed = $false

        try {
            $wg = Get-Command winget -ErrorAction Stop
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Installing via winget..." -ForegroundColor Cyan
            $p = Start-Process winget `
                -ArgumentList "install --id Microsoft.PowerShell --silent --accept-source-agreements --accept-package-agreements" `
                -Wait -PassThru -WindowStyle Hidden
            if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) { $ps7Installed = $true }
        } catch {}

        if (-not $ps7Installed) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Downloading PowerShell 7 MSI from GitHub..." -ForegroundColor Cyan
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $release = Invoke-RestMethod -Uri "https://api.github.com/repos/PowerShell/PowerShell/releases/latest" -UseBasicParsing
                $msi = $release.assets | Where-Object { $_.name -match "win-x64\.msi$" } | Select-Object -First 1
                if (-not $msi) { throw "No win-x64.msi asset found in latest release." }
                $msiPath = Join-Path $env:TEMP $msi.name
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Downloading: $($msi.name)" -ForegroundColor Cyan
                Invoke-WebRequest -Uri $msi.browser_download_url -OutFile $msiPath -UseBasicParsing
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Running silent installer..." -ForegroundColor Cyan
                $p = Start-Process msiexec.exe `
                    -ArgumentList "/i `"$msiPath`" /quiet /norestart ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 REGISTER_MANIFEST=1" `
                    -Wait -PassThru
                Remove-Item $msiPath -Force -ErrorAction SilentlyContinue
                if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) { $ps7Installed = $true }
                else { throw "msiexec exit code $($p.ExitCode)" }
            } catch {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')][ERROR] PS7 install failed: $_" -ForegroundColor Red
                Write-Host "Install manually from https://aka.ms/powershell then re-run." -ForegroundColor Yellow
                if (-not [Environment]::UserInteractive) { exit 1 }
                Read-Host "Press Enter to exit"
                exit 1
            }
        }

        if ($ps7Installed) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')][OK] PowerShell 7 installed successfully." -ForegroundColor Green
        }
    }

    if (Test-Path $pwsh7) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Relaunching in PowerShell 7..." -ForegroundColor Cyan
        $relaunchArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        foreach ($key in $MyInvocation.BoundParameters.Keys) {
            $val = $MyInvocation.BoundParameters[$key]
            if ($val -is [System.Management.Automation.SwitchParameter]) {
                if ($val.IsPresent) { $relaunchArgs += " -$key" }
            } else {
                $relaunchArgs += " -$key `"$val`""
            }
        }
        Start-Process $pwsh7 -ArgumentList $relaunchArgs -Verb RunAs
        exit
    }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ============================================================
# SECTION 0a - MODULE BOOTSTRAP
# ============================================================

function Invoke-ModuleBootstrap {
    param([bool]$Silent = $false)

    $required = @(
        @{ Name = "ExchangeOnlineManagement";                     Min = "3.4.0" },
        @{ Name = "Microsoft.Graph.Authentication";               Min = "2.0.0" },
        @{ Name = "Microsoft.Graph.Users";                        Min = "2.0.0" },
        @{ Name = "Microsoft.Graph.Applications";                 Min = "2.0.0" },
        @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; Min = "2.0.0" }
    )

    function BL { param($m, $l = "INFO")
        if (-not $Silent) {
            $c = switch ($l) { "OK" { "Green" } "WARN" { "Yellow" } "ERR" { "Red" } default { "Cyan" } }
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$l] $m" -ForegroundColor $c
        }
        if ($script:StartupLog) { $script:StartupLog.Add([PSCustomObject]@{ Level = $l; Message = $m }) }
    }

    BL "Checking required modules (running as: $($env:USERNAME))..."

    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
        BL "NuGet provider ready." "OK"
    } catch { BL "NuGet provider: $_" "WARN" }

    try {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        BL "PSGallery trusted." "OK"
    } catch { BL "PSGallery trust failed: $_" "WARN" }

    $psget = Get-Module -ListAvailable -Name PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $psget -or $psget.Version -lt [version]"2.2.5") {
        BL "PowerShellGet is v$($psget.Version) - must update to v2.2.5+ before other installs..." "WARN"
        try {
            Install-Module -Name PowerShellGet -MinimumVersion 2.2.5 -Scope AllUsers -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Remove-Module -Name PowerShellGet -Force -ErrorAction SilentlyContinue
            Remove-Module -Name PackageManagement -Force -ErrorAction SilentlyContinue
            Import-Module -Name PowerShellGet -Force -ErrorAction SilentlyContinue
            BL "PowerShellGet updated. Version: $((Get-Module PowerShellGet).Version)" "OK"
        } catch { BL "PowerShellGet update failed: $_" "ERR" }
    } else { BL "PowerShellGet v$($psget.Version) - OK" "OK" }

    foreach ($mod in $required) {
        try {
            $inst = Get-Module -ListAvailable -Name $mod.Name | Sort-Object Version -Descending | Select-Object -First 1
            $needsInstall = (-not $inst)
            $needsUpdate  = $false
            if ($inst) {
                try { $needsUpdate = ($inst.Version -lt [version]$mod.Min) } catch { $needsInstall = $true }
            }
            if ($needsInstall -or $needsUpdate) {
                $action = if ($needsInstall) { "Installing" } else { "Updating" }
                BL "$action $($mod.Name) (AllUsers)..."
                Install-Module -Name $mod.Name -MinimumVersion $mod.Min -Scope AllUsers -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
                BL "$($mod.Name) - done." "OK"
            } else { BL "$($mod.Name) v$($inst.Version) - OK" "OK" }
        } catch { BL "$($mod.Name) failed: $_" "ERR" }
    }

    BL "Module check complete." "OK"
}

function Repair-GraphModuleVersionSkew {
    # A -MinimumVersion check per module (as above) only guarantees each Microsoft.Graph.*
    # submodule is AT LEAST some old floor version - it says nothing about whether the
    # submodules match EACH OTHER. On a machine where Authentication got upgraded (e.g. by
    # some other install/update) while Applications/Identity.DirectoryManagement did not,
    # each ends up on a different version in a different module path (PS7-native vs the
    # legacy Windows PowerShell path PS7 still sees for compatibility). Unversioned
    # Import-Module then resolves Authentication to the newest available version, but
    # Applications pulls in its OWN required (older) Authentication version internally -
    # two different physical DLLs claiming the same assembly identity, which throws
    # "Assembly with same name is already loaded" the instant both are imported, even in a
    # brand-new process.
    #
    # Uninstalling the mismatched old versions to force alignment sounds right but is
    # fragile in practice: on a server also running AAD Connect Sync / AAD App Proxy
    # Connector / SPO Management Shell (all of which showed up in this exact failure's
    # PSModulePath), some background service or scheduled task may already have those exact
    # DLLs open, and Windows won't let anyone - Admin included - delete or replace a file
    # another process holds a handle to. Confirmed on a live Server 2019 box: repeated
    # elevated repair attempts left the module versions completely unchanged.
    #
    # Side-by-side module versions are normal and safe in PowerShell, so instead of fighting
    # file locks, just make sure the target version is ALSO installed (never touching the
    # old ones) and expose it via $script:GraphModuleTargetVersion so every Import-Module
    # call for this family can pin -RequiredVersion explicitly. An explicit version request
    # is never ambiguous, so the stale older copies become permanently harmless.
    param([bool]$Silent = $false)

    function RGL { param($m, $l = "INFO")
        if (-not $Silent) {
            $c = switch ($l) { "OK" { "Green" } "WARN" { "Yellow" } "ERR" { "Red" } default { "Cyan" } }
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$l] $m" -ForegroundColor $c
        }
        if ($script:StartupLog) { $script:StartupLog.Add([PSCustomObject]@{ Level = $l; Message = $m }) }
    }

    $graphModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Applications", "Microsoft.Graph.Identity.DirectoryManagement")

    RGL "Resolving a pinned Microsoft.Graph module version for this family (avoids ambiguous version resolution instead of trying to remove old copies)..."

    $target = $null
    try {
        $target = (Find-Module -Name "Microsoft.Graph.Authentication" -Repository PSGallery -ErrorAction Stop).Version
        RGL "Latest on PSGallery: $target"
    } catch {
        RGL "Could not query PSGallery for the latest version ($_) - falling back to the highest version already installed locally." "WARN"
        $target = (Get-Module -ListAvailable -Name "Microsoft.Graph.Authentication" -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1).Version
        if (-not $target) {
            RGL "No local Microsoft.Graph.Authentication install found either - cannot pin a version. Import-Module calls will fall back to unpinned (may still hit the assembly conflict)." "ERR"
            return
        }
        RGL "Using locally-installed version: $target" "WARN"
    }

    foreach ($mn in $graphModules) {
        $hasTarget = Get-Module -ListAvailable -Name $mn -ErrorAction SilentlyContinue | Where-Object { $_.Version -eq $target }
        if ($hasTarget) { RGL "$mn $target - already present." "OK"; continue }
        try {
            RGL "Installing $mn $target (AllUsers, side-by-side with any existing versions)..."
            Install-Module -Name $mn -RequiredVersion $target -Scope AllUsers -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            RGL "$mn $target installed." "OK"
        } catch {
            RGL "Failed to install $mn ${target}: $_" "ERR"
            $target = $null
            break
        }
    }

    if ($target) {
        $script:GraphModuleTargetVersion = $target
        RGL "Pinned Graph module version for this run: $target" "OK"
    } else {
        RGL "No pinned version available - Import-Module calls will fall back to unpinned." "WARN"
    }
}

# ============================================================
# SECTION 0b - M365 ROLE VERIFICATION
# ============================================================

function Invoke-RoleBootstrap {
    param([string]$AdminUPN, [bool]$Silent = $false)

    function RL { param($m, $l = "INFO")
        if (-not $Silent) {
            $c = switch ($l) { "OK" { "Green" } "WARN" { "Yellow" } "ERR" { "Red" } default { "Cyan" } }
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$l] $m" -ForegroundColor $c
        }
        if ($script:StartupLog) { $script:StartupLog.Add([PSCustomObject]@{ Level = $l; Message = $m }) }
    }

    $requiredRoles = @("Exchange Administrator", "User Administrator", "License Administrator")

    RL "Connecting to Microsoft Graph for role verification..."

    if ((Test-Path variable:script:Config) -and $script:Config.AppId -and $script:Config.CertThumbprint -and $script:Config.TenantId) {
        RL "Cert-based auth configured - skipping interactive role bootstrap." "OK"
        return
    }

    try {
        # Pin an exact version when Repair-GraphModuleVersionSkew resolved one - see that
        # function for why: avoids ambiguous Import-Module resolution across module paths
        # entirely, rather than depending on old mismatched versions having been removed.
        $verArgs = @{}
        if ($script:GraphModuleTargetVersion) { $verArgs['RequiredVersion'] = $script:GraphModuleTargetVersion }
        $null = Import-Module Microsoft.Graph.Authentication @verArgs -ErrorAction Stop
        $null = Import-Module Microsoft.Graph.Users @verArgs -ErrorAction Stop
        $null = Import-Module Microsoft.Graph.Identity.DirectoryManagement @verArgs -ErrorAction Stop

        $null = Connect-MgGraph -Scopes @("RoleManagement.ReadWrite.Directory", "User.Read.All", "Directory.Read.All") -NoWelcome -ErrorAction Stop

        $mgAdmin = $null
        if ($AdminUPN -ne "") {
            $mgAdmin = Get-MgUser -Filter "userPrincipalName eq '$AdminUPN'" -Property Id, UserPrincipalName -ErrorAction SilentlyContinue
        }
        if (-not $mgAdmin) {
            $ctx = Get-MgContext
            if ($ctx -and $ctx.Account) {
                $mgAdmin = Get-MgUser -Filter "userPrincipalName eq '$($ctx.Account)'" -Property Id, UserPrincipalName -ErrorAction SilentlyContinue
            }
        }
        if (-not $mgAdmin) {
            RL "Could not resolve admin account - skipping role assignment." "WARN"
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
            return
        }

        RL "Checking roles for: $($mgAdmin.UserPrincipalName)"

        foreach ($roleName in $requiredRoles) {
            $activeRole = Get-MgDirectoryRole -Filter "displayName eq '$roleName'" -ErrorAction SilentlyContinue
            if (-not $activeRole) {
                $tmpl = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq $roleName } | Select-Object -First 1
                if ($tmpl) {
                    $activeRole = New-MgDirectoryRole -RoleTemplateId $tmpl.Id
                    RL "Activated role: $roleName" "OK"
                } else { RL "Role template not found: $roleName" "WARN"; continue }
            }
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $activeRole.Id -ErrorAction SilentlyContinue
            if ($members | Where-Object { $_.Id -eq $mgAdmin.Id }) {
                RL "$roleName - already assigned" "OK"
            } else {
                RL "Assigning $roleName to $($mgAdmin.UserPrincipalName)..."
                try {
                    New-MgDirectoryRoleMemberByRef -DirectoryRoleId $activeRole.Id -BodyParameter @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($mgAdmin.Id)"
                    }
                    RL "$roleName assigned." "OK"
                } catch { RL "Could not assign $roleName (may need Global Admin): $_" "WARN" }
            }
        }

        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        RL "Role verification complete." "OK"
    } catch {
        RL "Role bootstrap error: $_" "WARN"
        try { $null = Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    }
}

# ============================================================
# SECTION 0c - STARTUP / CONFIG
# ============================================================

$isSilent = $ReportOnly.IsPresent
# Startup runs (module bootstrap, Graph version-skew repair, role bootstrap) all happen
# before the GUI window exists, so their Write-Host output only ever reached whatever
# console launched the script - easy to miss, and useless for after-the-fact diagnosis.
# Buffer it here and replay it into the GUI log box once that exists (see SECTION 5).
$script:StartupLog = New-Object System.Collections.Generic.List[object]

Write-Host "[$(Get-Date -Format 'HH:mm:ss')][INFO] Entra ID Inactive Licensed User Scanner - v1.0" -ForegroundColor Magenta

if (-not $SkipModuleCheck) {
    Invoke-ModuleBootstrap -Silent:$isSilent
    Repair-GraphModuleVersionSkew -Silent:$isSilent
}
$null = Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

# Separate config file from Offboard-ADUser.ps1's ADM365Config.json on purpose -
# this tool provisions its own app registration so re-running "Setup Auth" here
# can never delete/replace the cert the offboarding tool depends on.
$script:ConfigPath = Join-Path (Split-Path $PSCommandPath -Parent) "InactiveScanConfig.json"
$script:Config = @{ AdminUPN = ""; TenantId = ""; AppId = ""; CertThumbprint = ""; ExchangeOrg = "" }

if (Test-Path $script:ConfigPath) {
    try {
        $loaded = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
        foreach ($key in @('AdminUPN','TenantId','AppId','CertThumbprint','ExchangeOrg')) {
            $val = $loaded.PSObject.Properties[$key]
            if ($val -and $val.Value) { $script:Config[$key] = $val.Value }
        }
    } catch {}
}

if ($AdminUPN -ne "") {
    $script:Config.AdminUPN = $AdminUPN
    $script:Config | ConvertTo-Json | Set-Content $script:ConfigPath -Encoding UTF8
} elseif ($script:Config.AdminUPN -ne "") {
    $AdminUPN = $script:Config.AdminUPN
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')][OK] Using saved admin UPN: $AdminUPN" -ForegroundColor Green
}

function Save-Config {
    try { $script:Config | ConvertTo-Json | Set-Content $script:ConfigPath -Encoding UTF8 } catch {}
}

if (-not $SkipRoleCheck) { Invoke-RoleBootstrap -AdminUPN $AdminUPN -Silent:$isSilent }

# ============================================================
# SECTION 1 - M365 AUTH HELPERS
# ============================================================

function Get-CertConfig {
    if (-not (Test-Path variable:script:Config)) { return $null }
    $appId = $script:Config['AppId']; $thumb = $script:Config['CertThumbprint']; $tid = $script:Config['TenantId']
    if ($appId -and $thumb -and $tid) {
        $c = Get-Item "Cert:\LocalMachine\My\$thumb" -ErrorAction SilentlyContinue
        if ($c -and $c.NotAfter -gt (Get-Date)) { return $script:Config }
    }
    return $null
}

function Connect-ToGraph {
    $appId = $script:Config['AppId']; $tid = $script:Config['TenantId']; $thumb = $script:Config['CertThumbprint']
    if (-not $appId -or -not $tid -or -not $thumb) {
        Write-Log 'InactiveScanConfig.json is missing AppId / TenantId / CertThumbprint.' 'ERR'
        return $false
    }

    $cert = $null
    foreach ($store in @('LocalMachine', 'CurrentUser')) {
        $c = Get-Item "Cert:\$store\My\$thumb" -ErrorAction SilentlyContinue
        if ($c) {
            try { $null = $c.PrivateKey; $cert = $c; Write-Log "Certificate found in $store\My store." 'INFO'; break }
            catch { Write-Log "Cert in $store\My - private key not accessible here (wrong context). Trying next store." 'WARN' }
        }
    }
    if (-not $cert) {
        Write-Log "Thumbprint $thumb not found in LocalMachine\My or CurrentUser\My (or private key inaccessible)." 'ERR'
        return $false
    }

    try {
        Write-Log 'Connecting to Microsoft Graph (cert auth)...' 'INFO'
        $null = Connect-MgGraph -ClientId $appId -TenantId $tid -Certificate $cert -NoWelcome -ErrorAction Stop
        Write-Log 'Graph connected.' 'OK'
        return $true
    } catch { Write-Log "Graph connect failed: $_" 'ERR'; return $false }
}

function Connect-M365Graph {
    param(
        [string[]]$FallbackScopes = @("User.ReadWrite.All", "AuditLog.Read.All", "Organization.Read.All", "Directory.Read.All"),
        [System.Windows.Forms.RichTextBox]$LogBox = $null
    )
    $verArgs = @{}
    if ($script:GraphModuleTargetVersion) { $verArgs['RequiredVersion'] = $script:GraphModuleTargetVersion }
    $null = Import-Module Microsoft.Graph.Authentication @verArgs -ErrorAction Stop

    # Force a clean slate. If an earlier interactive Connect-MgGraph happened in this same
    # process (Role Bootstrap runs one at startup whenever no cert config exists YET - i.e.
    # exactly the state the very first Setup Auth run starts from) and this call reconnects
    # with the cert instead, don't rely on the SDK to fully swap contexts mid-session -
    # disconnect explicitly first so there is no ambiguity about which identity's token the
    # next API call actually carries.
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    $cfg = Get-CertConfig
    if ($cfg) {
        $ok = Connect-ToGraph
        if (-not $ok) { throw "Certificate-based Graph connection failed - see log." }
    } else {
        Write-Log "Graph: interactive auth - run 'Setup Auth' to enable non-interactive + sign-in activity read." "WARN" $LogBox
        $null = Connect-MgGraph -Scopes $FallbackScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        if ($ctx -and $ctx.TenantId -and -not $script:Config.TenantId) { $script:Config.TenantId = $ctx.TenantId; Save-Config }
    }

    $finalCtx = Get-MgContext
    if ($finalCtx) {
        Write-Log "Graph context: AuthType=$($finalCtx.AuthType)  ClientId=$($finalCtx.ClientId)  Account=$($finalCtx.Account)  Scopes=$($finalCtx.Scopes -join ',')" "INFO" $LogBox
    } else {
        Write-Log "Graph context: Get-MgContext returned nothing after connecting - unexpected." "WARN" $LogBox
    }
}

function Invoke-EXOProcess {
    # Runs Exchange Online cmdlets in a separate powershell.exe child process to avoid
    # the .NET AppDomain-level assembly conflict between Microsoft.Graph.Authentication
    # (Azure.Core / MSAL.NET) and ExchangeOnlineManagement (bundles different versions).
    param(
        [hashtable]$Params,
        [string]$Commands,
        [System.Windows.Forms.RichTextBox]$LogBox = $null,
        [bool]$WhatIf = $false
    )

    if ($WhatIf) {
        Write-Log "  [WHATIF] EXO subprocess would run Exchange operations." "WARN" $LogBox
        return
    }

    $uid        = [System.Guid]::NewGuid().ToString('N').Substring(0, 10)
    $paramFile  = Join-Path $env:TEMP "exo_p_$uid.json"
    $scriptFile = Join-Path $env:TEMP "exo_s_$uid.ps1"
    $outputFile = Join-Path $env:TEMP "exo_o_$uid.txt"

    try {
        $Params | ConvertTo-Json -Depth 10 | Set-Content $paramFile -Encoding UTF8

        $scriptBody = @"
`$ErrorActionPreference = 'Stop'
`$p = Get-Content '$($paramFile -replace "'","''")' -Raw | ConvertFrom-Json
Import-Module ExchangeOnlineManagement -ErrorAction Stop
if (`$p.AppId -and `$p.Thumbprint -and `$p.Organization) {
    Connect-ExchangeOnline -AppId `$p.AppId ``
        -CertificateThumbprint `$p.Thumbprint ``
        -Organization `$p.Organization ``
        -ShowBanner:`$false -ErrorAction Stop
} else {
    Connect-ExchangeOnline -ShowBanner:`$false -ErrorAction Stop
}

$Commands

Disconnect-ExchangeOnline -Confirm:`$false -ErrorAction SilentlyContinue
"@
        $scriptBody | Set-Content $scriptFile -Encoding UTF8

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $exeShell = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $psi.FileName  = $exeShell
        $psi.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$scriptFile`" *> `"$outputFile`""
        $psi.UseShellExecute = $true
        $psi.WindowStyle     = [System.Diagnostics.ProcessWindowStyle]::Hidden
        $psi.CreateNoWindow  = $false

        $proc = [System.Diagnostics.Process]::Start($psi)
        $proc.WaitForExit(300000) | Out-Null

        if (Test-Path $outputFile) {
            $out = Get-Content $outputFile -Raw -ErrorAction SilentlyContinue
            foreach ($line in ($out -split "`r?`n")) {
                $l = $line.Trim()
                if ($l -and $l -notmatch '^WARNING:') {
                    if ($l -match 'error|exception|fail' -and $l -notmatch 'ErrorAction') { Write-Log "  EXO: $l" "WARN" $LogBox }
                    else { Write-Log "  EXO: $l" "INFO" $LogBox }
                }
            }
        }

        if ($proc.ExitCode -ne 0) { throw "EXO subprocess exited with code $($proc.ExitCode). See output above." }
    }
    finally {
        Remove-Item $paramFile  -Force -ErrorAction SilentlyContinue
        Remove-Item $scriptFile -Force -ErrorAction SilentlyContinue
        Remove-Item $outputFile -Force -ErrorAction SilentlyContinue
    }
}

function Initialize-M365Auth {
    # Runs entirely in an isolated child pwsh.exe process. The Microsoft.Graph PowerShell
    # SDK ships each Microsoft.Graph.* submodule with its OWN private copy of
    # Microsoft.Graph.Authentication.dll; importing several submodules across different
    # stages of one long-lived session (Role Bootstrap already connected to Graph at
    # startup) throws "Assembly with same name is already loaded" the moment this function
    # imports Applications/Identity.DirectoryManagement on top of that. A fresh process has
    # no assemblies loaded yet, so it can never collide - same isolation trick as
    # Invoke-EXOProcess above, just for the Graph app-registration flow instead of EXO.
    param([System.Windows.Forms.RichTextBox]$LogBox = $null)

    Write-Log "Launching auth setup in an isolated PowerShell window (avoids a known Microsoft.Graph module assembly conflict with this session)..." "INFO" $LogBox
    Write-Log "A new console window will open. Sign in with a GLOBAL ADMIN account when prompted, then return here." "WARN" $LogBox

    $uid        = [System.Guid]::NewGuid().ToString('N').Substring(0, 10)
    $scriptFile = Join-Path $env:TEMP "setupauth_s_$uid.ps1"
    $resultFile = Join-Path $env:TEMP "setupauth_r_$uid.json"

    $authCommands = @'
$ErrorActionPreference = "Stop"
$result = [ordered]@{ Success = $false; AppId = ""; TenantId = ""; CertThumbprint = ""; ExchangeOrg = ""; ErrorMessage = ""; Warnings = "" }
try {
    $verArgs = @{}
    if ($TargetGraphVersion) { $verArgs["RequiredVersion"] = $TargetGraphVersion }
    Import-Module Microsoft.Graph.Authentication @verArgs -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications @verArgs -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement @verArgs -ErrorAction Stop

    Write-Host "Connecting to Microsoft Graph - a sign-in window/browser should appear..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes @(
        "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All",
        "RoleManagement.ReadWrite.Directory", "Directory.ReadWrite.All", "Organization.Read.All"
    ) -NoWelcome -ErrorAction Stop

    $ctx = Get-MgContext
    $tenantId = $ctx.TenantId
    Write-Host "Connected. Tenant: $tenantId" -ForegroundColor Green

    Write-Host "Creating authentication certificate..." -ForegroundColor Cyan
    $appName = "ADM365InactiveScanTool"
    $cert = New-SelfSignedCertificate -Subject "CN=$appName" `
        -CertStoreLocation "Cert:\LocalMachine\My" `
        -KeyExportPolicy NonExportable -KeySpec Signature `
        -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears(2)
    Write-Host "Certificate: $($cert.Thumbprint)" -ForegroundColor Green

    $old = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($old) { Remove-MgApplication -ApplicationId $old.Id; Write-Host "Removed existing app." -ForegroundColor Yellow }

    Write-Host "Creating Entra ID app registration..." -ForegroundColor Cyan
    $app = New-MgApplication -DisplayName $appName -SignInAudience "AzureADMyOrg"
    Write-Host "App created: $($app.AppId)" -ForegroundColor Green

    Update-MgApplication -ApplicationId $app.Id -KeyCredentials @(@{
        Type = "AsymmetricX509Cert"; Usage = "Verify"; Key = $cert.RawData; DisplayName = "$appName Cert"
    }) | Out-Null
    Write-Host "Certificate uploaded to app." -ForegroundColor Green

    $graphSP = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
    $exoSP   = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'" -ErrorAction SilentlyContinue

    # Graph app roles: User.ReadWrite.All (rename + license removal), AuditLog.Read.All
    # (required to read signInActivity in app-only context), Organization/Directory.Read.All (lookups)
    $graphPermNames = @("User.ReadWrite.All", "AuditLog.Read.All", "Organization.Read.All", "Directory.Read.All")
    $graphRoles = @()
    foreach ($p in $graphPermNames) {
        $roleObj = $graphSP.AppRoles | Where-Object { $_.Value -eq $p } | Select-Object -First 1
        if ($roleObj) { $graphRoles += @{ Name = $p; Id = [string]$roleObj.Id; Type = "Role" } }
        else { Write-Host "Graph app role not found: $p" -ForegroundColor Yellow }
    }
    $reqAccess = @(@{ ResourceAppId = "00000003-0000-0000-c000-000000000000"; ResourceAccess = ($graphRoles | ForEach-Object { @{ Id = $_.Id; Type = $_.Type } }) })

    $exoRoleId = $null
    if ($exoSP) {
        $er = $exoSP.AppRoles | Where-Object { $_.Value -eq "Exchange.ManageAsApp" } | Select-Object -First 1
        if ($er) {
            $exoRoleId = [string]$er.Id
            $reqAccess += @{ ResourceAppId = "00000002-0000-0ff1-ce00-000000000000"; ResourceAccess = @(@{ Id = $exoRoleId; Type = "Role" }) }
        }
    }

    Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $reqAccess | Out-Null
    Write-Host "Permissions configured." -ForegroundColor Green

    Write-Host "Creating service principal..." -ForegroundColor Cyan
    $sp = New-MgServicePrincipal -AppId $app.AppId

    Write-Host "Waiting 15s for propagation, then granting admin consent..." -ForegroundColor Cyan
    Start-Sleep -Seconds 15

    # New-MgServicePrincipalAppRoleAssignment can fail transiently right after the SP was
    # just created (directory replication lag) even after the 15s wait above - retry each
    # one instead of the previous silent "try { } catch {}" (which meant a failed grant here
    # showed up later only as a confusing 403 Authorization_RequestDenied, with zero
    # indication which permission was actually missing).
    function Grant-AppRoleWithRetry {
        param($ServicePrincipalId, $ResourceId, $AppRoleId, $Label)
        for ($i = 1; $i -le 5; $i++) {
            try {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -PrincipalId $ServicePrincipalId -ResourceId $ResourceId -AppRoleId $AppRoleId | Out-Null
                Write-Host "  Granted: $Label" -ForegroundColor Green
                return $true
            } catch {
                if ($i -eq 5) {
                    Write-Host "  FAILED to grant ${Label}: $_" -ForegroundColor Red
                    return $false
                }
                Start-Sleep -Seconds 5
            }
        }
    }

    $roleGrantFailures = New-Object System.Collections.Generic.List[string]
    foreach ($role in $graphRoles) {
        if (-not (Grant-AppRoleWithRetry -ServicePrincipalId $sp.Id -ResourceId $graphSP.Id -AppRoleId $role.Id -Label $role.Name)) {
            $roleGrantFailures.Add($role.Name)
        }
    }
    if ($exoSP -and $exoRoleId) {
        if (-not (Grant-AppRoleWithRetry -ServicePrincipalId $sp.Id -ResourceId $exoSP.Id -AppRoleId $exoRoleId -Label "Exchange.ManageAsApp")) {
            $roleGrantFailures.Add("Exchange.ManageAsApp")
        }
    }
    if ($roleGrantFailures.Count -gt 0) {
        $result.Warnings = "Could not grant: $($roleGrantFailures -join ', '). Fix in Entra admin center > Enterprise Applications > ADM365InactiveScanTool > Permissions, or just re-run Setup Auth."
        Write-Host "WARNING: $($result.Warnings)" -ForegroundColor Yellow
    } else {
        Write-Host "Admin consent granted for all requested permissions." -ForegroundColor Green
    }

    # Exchange.ManageAsApp additionally requires the SP to hold Exchange Administrator
    try {
        $exoAdmRole = Get-MgDirectoryRole -Filter "displayName eq 'Exchange Administrator'" -ErrorAction SilentlyContinue
        if (-not $exoAdmRole) {
            $t = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq "Exchange Administrator" } | Select-Object -First 1
            $exoAdmRole = New-MgDirectoryRole -RoleTemplateId $t.Id
        }
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $exoAdmRole.Id -BodyParameter @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)"
        } | Out-Null
        Write-Host "Exchange Administrator role assigned to service principal." -ForegroundColor Green
    } catch { Write-Host "Exchange role assignment warning: $_" -ForegroundColor Yellow }

    $org = Get-MgOrganization | Select-Object -First 1
    $exoOrg = ($org.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1).Name

    $result.Success       = $true
    $result.AppId          = $app.AppId
    $result.TenantId       = $tenantId
    $result.CertThumbprint = $cert.Thumbprint
    $result.ExchangeOrg    = $exoOrg

    Disconnect-MgGraph -ErrorAction SilentlyContinue

    Write-Host "============================================" -ForegroundColor Green
    Write-Host " SETUP COMPLETE - non-interactive auth ready" -ForegroundColor Green
    Write-Host " App ID      : $($app.AppId)" -ForegroundColor Green
    Write-Host " Cert        : $($cert.Thumbprint)" -ForegroundColor Green
    Write-Host " Exchange Org: $exoOrg" -ForegroundColor Green
    Write-Host " Wait 5-10 minutes before first use, and confirm Azure AD Premium P1/P2" -ForegroundColor Yellow
    Write-Host " is licensed in this tenant - otherwise signInActivity reads back empty." -ForegroundColor Yellow
    Write-Host "============================================" -ForegroundColor Green
} catch {
    $errLine = $_.InvocationInfo.ScriptLineNumber
    $diag = New-Object System.Collections.Generic.List[string]
    try {
        $diag.Add("PSVersion: $($PSVersionTable.PSVersion)")
        $wi = [Security.Principal.WindowsIdentity]::GetCurrent()
        $diag.Add("Running as: $($wi.Name)")
        $wp = New-Object Security.Principal.WindowsPrincipal($wi)
        $diag.Add("Elevated: $($wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))")
        $diag.Add("PSModulePath entries:")
        ($env:PSModulePath -split ";") | ForEach-Object { $diag.Add("  $_") }
        $diag.Add("Get-Module -ListAvailable for the three imported modules:")
        foreach ($mn in @("Microsoft.Graph.Authentication","Microsoft.Graph.Applications","Microsoft.Graph.Identity.DirectoryManagement")) {
            $mods = Get-Module -ListAvailable -Name $mn -ErrorAction SilentlyContinue
            if ($mods) { foreach ($m in $mods) { $diag.Add("  $mn $($m.Version) -> $($m.ModuleBase)") } }
            else { $diag.Add("  $mn : NOT FOUND") }
        }
        $diag.Add("Loaded assemblies matching Microsoft.Graph*:")
        [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -like "Microsoft.Graph*" } | ForEach-Object {
            $diag.Add("  $($_.GetName().Name) v$($_.GetName().Version) -> $($_.Location)")
        }
    } catch { $diag.Add("diagnostic collection error: $_") }

    $result.ErrorMessage = "$_ (failed at line $errLine)`n---DIAGNOSTICS---`n" + ($diag -join "`n")
    Write-Host "[ERROR] Setup failed: $_ (line $errLine)" -ForegroundColor Red
    $diag | ForEach-Object { Write-Host $_ -ForegroundColor DarkGray }
    if ((Test-Path variable:cert) -and $cert) {
        try { Remove-Item "Cert:\LocalMachine\My\$($cert.Thumbprint)" -DeleteKey -ErrorAction SilentlyContinue } catch {}
    }
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}
'@

    $scriptBody = @"
`$ErrorActionPreference = 'Stop'
`$TargetGraphVersion = '$($script:GraphModuleTargetVersion)'

$authCommands

(`$result | ConvertTo-Json) | Set-Content '$($resultFile -replace "'","''")' -Encoding UTF8
Write-Host ''
Write-Host 'You may close this window now.' -ForegroundColor Cyan
"@
    $scriptBody | Set-Content $scriptFile -Encoding UTF8

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $exeShell = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    $psi.FileName        = $exeShell
    $psi.Arguments       = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptFile`""
    $psi.UseShellExecute = $true
    $psi.WindowStyle     = [System.Diagnostics.ProcessWindowStyle]::Normal

    try {
        $proc = [System.Diagnostics.Process]::Start($psi)
        $proc.WaitForExit()
    } catch {
        Write-Log "Could not launch setup process: $_" "ERROR" $LogBox
        Remove-Item $scriptFile -Force -ErrorAction SilentlyContinue
        return $false
    }

    if (-not (Test-Path $resultFile)) {
        Write-Log "Setup Auth window closed without producing a result - it may have been closed early, or sign-in was cancelled." "ERROR" $LogBox
        Remove-Item $scriptFile -Force -ErrorAction SilentlyContinue
        return $false
    }

    $r = Get-Content $resultFile -Raw | ConvertFrom-Json
    Remove-Item $resultFile -Force -ErrorAction SilentlyContinue
    Remove-Item $scriptFile -Force -ErrorAction SilentlyContinue

    if (-not $r.Success) {
        Write-Log "Setup Auth failed: $($r.ErrorMessage)" "ERROR" $LogBox
        return $false
    }

    $script:Config.AppId          = $r.AppId
    $script:Config.TenantId       = $r.TenantId
    $script:Config.CertThumbprint = $r.CertThumbprint
    $script:Config.ExchangeOrg    = $r.ExchangeOrg
    Save-Config

    Write-Log "Setup complete. App ID: $($r.AppId)  Cert: $($r.CertThumbprint)  Exchange Org: $($r.ExchangeOrg)" "SUCCESS" $LogBox
    if ($r.Warnings) { Write-Log $r.Warnings "WARN" $LogBox }
    Write-Log "Wait 5-10 minutes before first use, and confirm Azure AD Premium P1/P2 is licensed in this tenant." "WARN" $LogBox
    return $true
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO", [System.Windows.Forms.RichTextBox]$LogBox = $null)
    $ts = Get-Date -Format "HH:mm:ss"
    $cc = switch ($Level) { "WARN" { "Yellow" } "ERROR" { "Red" } "SUCCESS" { "Green" } "HEAD" { "Magenta" } default { "Cyan" } }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $cc

    if ($LogBox -and -not $LogBox.IsDisposed) {
        $rc = switch ($Level) {
            "WARN" { [System.Drawing.Color]::Gold } "ERROR" { [System.Drawing.Color]::Tomato }
            "SUCCESS" { [System.Drawing.Color]::LimeGreen } "HEAD" { [System.Drawing.Color]::Orchid }
            default { [System.Drawing.Color]::FromArgb(140, 200, 255) }
        }
        $LogBox.SelectionStart = $LogBox.TextLength
        $LogBox.SelectionColor = $rc
        $LogBox.AppendText("[$ts][$Level] $Message`r`n")
        $LogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ============================================================
# SECTION 2 - SCAN ENGINE
# ============================================================

function Get-InactiveLicensedUsers {
    param(
        [int]$InactiveDays,
        [System.Windows.Forms.RichTextBox]$LogBox = $null
    )

    Connect-M365Graph -LogBox $LogBox
    $cutoff = (Get-Date).ToUniversalTime().AddDays(-$InactiveDays)

    Write-Log "Retrieving licensed users (this can take a while on large tenants)..." "INFO" $LogBox
    $select = "id,displayName,userPrincipalName,mail,accountEnabled,assignedLicenses,onPremisesSyncEnabled,createdDateTime,userType,signInActivity"
    $uri = "https://graph.microsoft.com/v1.0/users?`$select=$select&`$top=999"

    $allUsers = New-Object System.Collections.Generic.List[object]
    $sawSignInField = $false
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        foreach ($u in $resp.value) {
            if ($u.PSObject.Properties['signInActivity']) { $sawSignInField = $true }
            $allUsers.Add($u)
        }
        $uri = $resp.'@odata.nextLink'
    } while ($uri)

    Write-Log "Retrieved $($allUsers.Count) total user objects." "INFO" $LogBox
    if (-not $sawSignInField) {
        Write-Log "signInActivity was not returned by Graph - falling back to account-creation-date heuristic only. Check that AuditLog.Read.All is consented and this tenant has Azure AD Premium P1/P2." "WARN" $LogBox
    }

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($u in $allUsers) {
        if ($u.userType -ne 'Member') { continue }
        if (-not $u.assignedLicenses -or $u.assignedLicenses.Count -eq 0) { continue }

        $lastInteractive    = $u.signInActivity.lastSignInDateTime
        $lastNonInteractive = $u.signInActivity.lastNonInteractiveSignInDateTime
        $lastActivity = @($lastInteractive, $lastNonInteractive) |
            Where-Object { $_ } | ForEach-Object { [datetime]$_ } |
            Sort-Object -Descending | Select-Object -First 1

        $created = if ($u.createdDateTime) { [datetime]$u.createdDateTime } else { $null }

        $isInactive = $false
        if ($lastActivity) {
            $isInactive = $lastActivity -lt $cutoff
        } elseif ($created) {
            $isInactive = $created -lt $cutoff
        }
        if (-not $isInactive) { continue }

        $refDate = if ($lastActivity) { $lastActivity } elseif ($created) { $created } else { $null }
        $daysInactive = if ($refDate) { [math]::Round(((Get-Date).ToUniversalTime() - $refDate).TotalDays) } else { -1 }

        $results.Add([PSCustomObject]@{
            Select            = $false
            Id                = $u.id
            DisplayName       = $u.displayName
            UserPrincipalName = $u.userPrincipalName
            Mail              = $u.mail
            AccountEnabled    = [bool]$u.accountEnabled
            Synced            = [bool]$u.onPremisesSyncEnabled
            LastActivity      = $lastActivity
            DaysInactive      = $daysInactive
            LicenseCount      = $u.assignedLicenses.Count
            SkuIds            = @($u.assignedLicenses | ForEach-Object { $_.skuId })
        })
    }

    Write-Log "$($results.Count) licensed member account(s) inactive $InactiveDays+ days." "SUCCESS" $LogBox
    return $results
}

function Export-ScanReport {
    param($Candidates, [string]$Path)
    $Candidates | Select-Object DisplayName, UserPrincipalName, Mail, AccountEnabled, Synced,
        @{N='LastActivityUTC';E={ if ($_.LastActivity) { $_.LastActivity.ToString('u') } else { '(never recorded)' } }},
        DaysInactive, LicenseCount |
        Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

# ============================================================
# SECTION 3 - REMEDIATION ENGINE
# ============================================================

function Invoke-UserRemediation {
    param(
        $Candidate,
        [bool]$WhatIfMode,
        [System.Windows.Forms.RichTextBox]$LogBox = $null
    )

    $result = [PSCustomObject]@{
        UserPrincipalName = $Candidate.UserPrincipalName
        MailboxConverted  = $false
        Renamed           = $false
        LicensesRemoved   = $false
        Error             = ""
    }

    function Step { param([string]$d, [scriptblock]$a)
        if ($WhatIfMode) { Write-Log "  [WHATIF] $d" "WARN" $LogBox }
        else             { Write-Log "  $d" "INFO" $LogBox; & $a }
    }

    $upn        = $Candidate.UserPrincipalName
    $origMail   = $Candidate.Mail
    $upnLocal   = $upn.Split("@")[0]
    $upnDomain  = $upn.Split("@")[1]
    $mailLocal  = if ($origMail) { $origMail.Split("@")[0] } else { $upnLocal }
    $mailDomain = if ($origMail) { $origMail.Split("@")[1] } else { $upnDomain }

    $alreadyRenamed = $upnLocal.ToLower().StartsWith("historical-")
    $newMail = "historical-$mailLocal@$mailDomain"
    $newUpn  = "historical-$upnLocal@$upnDomain"
    $newDisplayName = "Historical - $($Candidate.DisplayName)"

    try {
        Write-Log ">> $upn" "HEAD" $LogBox

        # 1. Convert mailbox to Shared + rename primary SMTP (subprocess - avoids Graph/EXO assembly conflict)
        if ($origMail) {
            $cfg = Get-CertConfig
            $eParams = @{
                AppId        = if ($cfg) { $cfg['AppId'] }         else { "" }
                Thumbprint   = if ($cfg) { $cfg['CertThumbprint'] } else { "" }
                Organization = if ($cfg) { $cfg['ExchangeOrg'] }   else { "" }
                OrigMail     = $origMail
                NewMail      = $newMail
                AlreadyRenamed = $alreadyRenamed
            }
            $exoCommands = @'
Write-Output "Converting mailbox to Shared: $($p.OrigMail)"
Set-Mailbox -Identity $p.OrigMail -Type Shared
Write-Output "Mailbox converted to Shared."

if (-not $p.AlreadyRenamed) {
    Write-Output "Adding new proxy address: $($p.NewMail)"
    Set-Mailbox -Identity $p.OrigMail -EmailAddresses @{Add="smtp:$($p.NewMail)"}
    Write-Output "Promoting to primary SMTP: $($p.NewMail)"
    Set-Mailbox -Identity $p.OrigMail -PrimarySmtpAddress $p.NewMail
    Write-Output "Primary SMTP updated."
} else {
    Write-Output "Address already carries historical- prefix - skipping SMTP rename."
}
'@
            Step "Convert mailbox to Shared + rename primary SMTP" {
                Invoke-EXOProcess -Params $eParams -Commands $exoCommands -LogBox $LogBox -WhatIf $WhatIfMode
            }
            $result.MailboxConverted = $true
        } else {
            Write-Log "  No mailbox (Mail property empty) - skipping shared-mailbox conversion." "WARN" $LogBox
        }

        # 2. Rename DisplayName + UPN (Graph) - skip for on-prem synced accounts
        if ($Candidate.Synced) {
            Write-Log "  Synced from on-prem AD - skipping DisplayName/UPN rename (would be reverted by next AAD Connect cycle). Use Offboard-ADUser.ps1 / on-prem AD to rename this account." "WARN" $LogBox
        } elseif ($alreadyRenamed) {
            Write-Log "  UPN already carries historical- prefix - skipping rename." "INFO" $LogBox
            $result.Renamed = $true
        } else {
            Step "Rename DisplayName -> `"$newDisplayName`", UPN -> `"$newUpn`"" {
                $body = @{ displayName = $newDisplayName; userPrincipalName = $newUpn } | ConvertTo-Json
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($Candidate.Id)" -Body $body -ContentType "application/json" | Out-Null
            }
            $result.Renamed = $true
        }

        # 3. Remove all assigned licenses
        if ($Candidate.SkuIds.Count -gt 0) {
            Step "Remove $($Candidate.SkuIds.Count) assigned license(s)" {
                $licBody = @{ addLicenses = @(); removeLicenses = @($Candidate.SkuIds) } | ConvertTo-Json -Depth 5
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($Candidate.Id)/assignLicense" -Body $licBody -ContentType "application/json" | Out-Null
            }
            $result.LicensesRemoved = $true
        } else {
            Write-Log "  No licenses assigned - nothing to remove." "INFO" $LogBox
        }

        Write-Log "  Done: $upn" "SUCCESS" $LogBox
    } catch {
        $result.Error = "$_"
        Write-Log "  ERROR on $upn - $_" "ERROR" $LogBox
    }

    return $result
}

# ============================================================
# SECTION 4 - REPORT-ONLY (HEADLESS) PATH
# ============================================================

if ($ReportOnly) {
    $candidates = Get-InactiveLicensedUsers -InactiveDays $InactiveDays
    $reportPath = Join-Path (Split-Path $PSCommandPath -Parent) "InactiveUserScan_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Export-ScanReport -Candidates $candidates -Path $reportPath
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')][OK] Report written: $reportPath ($($candidates.Count) accounts)" -ForegroundColor Green
    exit 0
}

# ============================================================
# SECTION 5 - GUI
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$BG      = [System.Drawing.Color]::FromArgb(20,  24,  33)
$PANEL   = [System.Drawing.Color]::FromArgb(30,  36,  50)
$CLR_INPUT = [System.Drawing.Color]::FromArgb(18,  22,  32)
$ACCENT  = [System.Drawing.Color]::FromArgb(0,  160, 230)
$GREEN   = [System.Drawing.Color]::FromArgb(0,  200, 130)
$WARN    = [System.Drawing.Color]::FromArgb(240, 165,   0)
$DANGER  = [System.Drawing.Color]::FromArgb(210,  55,  55)
$TEXT    = [System.Drawing.Color]::FromArgb(215, 225, 240)
$TEXTDIM = [System.Drawing.Color]::FromArgb(110, 128, 160)

$F_NORM  = New-Object System.Drawing.Font("Segoe UI", 9)
$F_BOLD  = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$F_MONO  = New-Object System.Drawing.Font("Consolas", 8.5)

function Lbl { param($t, $x, $y, $w, $h, $c = $TEXT, $f = $F_NORM)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $t; $l.Location = New-Object System.Drawing.Point($x, $y); $l.Size = New-Object System.Drawing.Size($w, $h)
    $l.Font = $f; $l.ForeColor = $c; $l.BackColor = [System.Drawing.Color]::Transparent; $l
}
function Btn { param($t, $x, $y, $w, $h, $bg = $ACCENT, $fg = [System.Drawing.Color]::White)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $t; $b.Location = New-Object System.Drawing.Point($x, $y); $b.Size = New-Object System.Drawing.Size($w, $h)
    $b.Font = $F_BOLD; $b.ForeColor = $fg; $b.BackColor = $bg
    $b.FlatStyle = "Flat"; $b.FlatAppearance.BorderSize = 0; $b
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Entra ID Inactive Licensed User Scanner"
$form.Size = New-Object System.Drawing.Size(1180, 760)
$form.StartPosition = "CenterScreen"; $form.BackColor = $BG
$form.FormBorderStyle = "Sizable"; $form.MinimumSize = New-Object System.Drawing.Size(900, 560); $form.Font = $F_NORM

$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Dock = "Top"; $pnlTop.Height = 90; $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)
$form.Controls.Add($pnlTop)

$lblTitle = Lbl "Entra ID Inactive Licensed User Scanner" 16 10 700 26 $TEXT (New-Object System.Drawing.Font("Segoe UI Semibold", 13))
$pnlTop.Controls.Add($lblTitle)

$lblDays = Lbl "Inactive threshold (days):" 16 46 170 22 $TEXT
$pnlTop.Controls.Add($lblDays)
$numDays = New-Object System.Windows.Forms.NumericUpDown
$numDays.Location = New-Object System.Drawing.Point(190, 44); $numDays.Size = New-Object System.Drawing.Size(70, 22)
$numDays.Minimum = 1; $numDays.Maximum = 3650; $numDays.Value = $InactiveDays
$numDays.Font = $F_NORM; $numDays.BackColor = $CLR_INPUT; $numDays.ForeColor = $TEXT
$pnlTop.Controls.Add($numDays)

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Text = "WhatIf - simulation only (no changes)"
$chkWhatIf.Location = New-Object System.Drawing.Point(280, 46); $chkWhatIf.Size = New-Object System.Drawing.Size(280, 22)
$chkWhatIf.Font = $F_BOLD; $chkWhatIf.ForeColor = $WARN; $chkWhatIf.BackColor = [System.Drawing.Color]::Transparent
$chkWhatIf.Checked = $true
$pnlTop.Controls.Add($chkWhatIf)

$btnScan       = Btn "Scan Tenant"       580 40 120 32 $ACCENT
$btnSelectAll  = Btn "Select All"        710 40 100 32 $PANEL $TEXT
$btnSelectNone = Btn "Select None"       818 40 100 32 $PANEL $TEXT
$btnSetupAuth  = Btn "Setup Auth"        926 40 110 32 $PANEL $WARN
$btnExport     = Btn "Export CSV"       1044 40 110 32 $PANEL $TEXT
$pnlTop.Controls.AddRange(@($btnScan, $btnSelectAll, $btnSelectNone, $btnSetupAuth, $btnExport))

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(0, 90)
$grid.Anchor = "Top,Bottom,Left,Right"
$grid.Size = New-Object System.Drawing.Size(1164, 400)
$grid.BackgroundColor = $PANEL
$grid.ForeColor = $TEXT
$grid.DefaultCellStyle.BackColor = $CLR_INPUT
$grid.DefaultCellStyle.ForeColor = $TEXT
$grid.DefaultCellStyle.SelectionBackColor = $ACCENT
$grid.ColumnHeadersDefaultCellStyle.BackColor = $PANEL
$grid.ColumnHeadersDefaultCellStyle.ForeColor = $TEXT
$grid.ColumnHeadersDefaultCellStyle.Font = $F_BOLD
$grid.EnableHeadersVisualStyles = $false
$grid.RowHeadersVisible = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"
$grid.Font = $F_NORM
$form.Controls.Add($grid)

$lblStatus = Lbl "Set a threshold and click Scan Tenant." 16 500 700 20 $TEXTDIM
$lblStatus.Anchor = "Bottom,Left"

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(16, 522); $progress.Size = New-Object System.Drawing.Size(1148, 16)
$progress.Anchor = "Bottom,Left,Right"; $progress.Minimum = 0; $progress.Maximum = 100
$progress.ForeColor = $GREEN; $progress.BackColor = $PANEL

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(16, 546); $logBox.Size = New-Object System.Drawing.Size(1148, 130)
$logBox.Anchor = "Bottom,Left,Right"
$logBox.Font = $F_MONO; $logBox.BackColor = [System.Drawing.Color]::FromArgb(10, 14, 22)
$logBox.ForeColor = $TEXT; $logBox.ReadOnly = $true; $logBox.BorderStyle = "None"; $logBox.ScrollBars = "Vertical"

if ($script:StartupLog -and $script:StartupLog.Count -gt 0) {
    Write-Log "--- Startup log (module bootstrap / Graph version-skew repair / role check) ---" "INFO" $logBox
    foreach ($entry in $script:StartupLog) {
        $mappedLevel = switch ($entry.Level) { "OK" { "SUCCESS" } "ERR" { "ERROR" } "WARN" { "WARN" } default { "INFO" } }
        Write-Log $entry.Message $mappedLevel $logBox
    }
    Write-Log "--- End startup log ---" "INFO" $logBox
}

$pnlBot = New-Object System.Windows.Forms.Panel
$pnlBot.Anchor = "Bottom,Left,Right"; $pnlBot.Location = New-Object System.Drawing.Point(0, 684); $pnlBot.Size = New-Object System.Drawing.Size(1180, 54)
$pnlBot.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)

$btnProcess = Btn "Process Selected Accounts" 16 10 220 34 $DANGER
$btnProcess.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$btnClose = Btn "Close" 1076 10 88 34 $PANEL $TEXT
$pnlBot.Controls.AddRange(@($btnProcess, $btnClose))

$form.Controls.AddRange(@($lblStatus, $progress, $logBox, $pnlBot))

$script:candidates = @()

function Refresh-Grid {
    $grid.Rows.Clear()
    if ($grid.Columns.Count -eq 0) {
        $grid.Columns.Add((New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = "Select"; HeaderText = "Sel"; Width = 40 })) | Out-Null
        $grid.Columns.Add("DisplayName", "Display Name") | Out-Null
        $grid.Columns.Add("UserPrincipalName", "UPN") | Out-Null
        $grid.Columns.Add("Mail", "Mail") | Out-Null
        $grid.Columns.Add("AccountEnabled", "Enabled") | Out-Null
        $grid.Columns.Add("Synced", "AD-Synced") | Out-Null
        $grid.Columns.Add("DaysInactive", "Days Inactive") | Out-Null
        $grid.Columns.Add("LicenseCount", "Licenses") | Out-Null
        foreach ($colName in @("DisplayName","UserPrincipalName","Mail","AccountEnabled","Synced","DaysInactive","LicenseCount")) {
            $grid.Columns[$colName].ReadOnly = $true
        }
    }
    foreach ($c in $script:candidates) {
        $idx = $grid.Rows.Add($false, $c.DisplayName, $c.UserPrincipalName, $c.Mail, $c.AccountEnabled, $c.Synced, $c.DaysInactive, $c.LicenseCount)
        if ($c.Synced) { $grid.Rows[$idx].DefaultCellStyle.ForeColor = $WARN }
        if (-not $c.AccountEnabled) { $grid.Rows[$idx].Cells['AccountEnabled'].Style.ForeColor = $DANGER }
    }
}

$grid.Add_CurrentCellDirtyStateChanged({
    if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
})

$btnScan.Add_Click({
    $btnScan.Enabled = $false
    $logBox.Clear()
    $lblStatus.ForeColor = $TEXTDIM; $lblStatus.Text = "Scanning tenant..."
    try {
        $script:candidates = Get-InactiveLicensedUsers -InactiveDays ([int]$numDays.Value) -LogBox $logBox
        Refresh-Grid
        $lblStatus.ForeColor = $GREEN
        $lblStatus.Text = "$($script:candidates.Count) inactive licensed account(s) found. Review and check accounts to process."
        $reportPath = Join-Path (Split-Path $PSCommandPath -Parent) "InactiveUserScan_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        Export-ScanReport -Candidates $script:candidates -Path $reportPath
        Write-Log "Scan report written: $reportPath" "SUCCESS" $logBox
    } catch {
        $lblStatus.ForeColor = $DANGER; $lblStatus.Text = "[X] Scan failed - see log."
        Write-Log "Scan error: $_" "ERROR" $logBox

        if ("$_" -match "403|Authorization_RequestDenied") {
            Write-Log "403 detected - running two follow-up probes to isolate the cause..." "WARN" $logBox
            try {
                $ctx = Get-MgContext
                if ($ctx) {
                    Write-Log "Context at failure time: AuthType=$($ctx.AuthType)  ClientId=$($ctx.ClientId)  Account=$($ctx.Account)  Scopes=$($ctx.Scopes -join ',')" "INFO" $logBox
                } else {
                    Write-Log "Context at failure time: Get-MgContext returned nothing." "WARN" $logBox
                }
            } catch { Write-Log "Could not read Get-MgContext: $_" "WARN" $logBox }

            try {
                $probe1 = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users?`$top=1&`$select=id,displayName" -ErrorAction Stop
                Write-Log "Probe 1 (basic /users query, no signInActivity): SUCCEEDED - $($probe1.value.Count) result(s). Confirms the token/permissions work for basic reads; the failure is specific to signInActivity." "SUCCESS" $logBox
            } catch {
                Write-Log "Probe 1 (basic /users query, no signInActivity): FAILED - $_. The token itself lacks basic Directory/User read - broader than just signInActivity." "ERROR" $logBox
            }

            try {
                $probe2 = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users?`$top=1&`$select=id,displayName,signInActivity" -Headers @{ ConsistencyLevel = "eventual" } -ErrorAction Stop
                Write-Log "Probe 2 (signInActivity + ConsistencyLevel:eventual header): SUCCEEDED - $($probe2.value.Count) result(s). The eventual-consistency header appears to be required for this tenant." "SUCCESS" $logBox
            } catch {
                Write-Log "Probe 2 (signInActivity + ConsistencyLevel:eventual header): FAILED - $_" "ERROR" $logBox
            }
        }
    } finally { $btnScan.Enabled = $true }
})

$btnSelectAll.Add_Click({ foreach ($row in $grid.Rows) { $row.Cells['Select'].Value = $true } })
$btnSelectNone.Add_Click({ foreach ($row in $grid.Rows) { $row.Cells['Select'].Value = $false } })

$btnExport.Add_Click({
    if ($script:candidates.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Run a scan first.", "Nothing to export") | Out-Null; return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV files (*.csv)|*.csv"
    $sfd.FileName = "InactiveUserScan_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($sfd.ShowDialog() -eq "OK") {
        Export-ScanReport -Candidates $script:candidates -Path $sfd.FileName
        Write-Log "Exported: $($sfd.FileName)" "SUCCESS" $logBox
    }
})

$btnSetupAuth.Add_Click({
    $r = [System.Windows.Forms.MessageBox]::Show(
        "This creates a certificate-based app registration in this tenant so the tool no longer needs interactive MFA sign-in each run, and grants it AuditLog.Read.All (needed to read sign-in activity) plus Exchange.ManageAsApp.`n`nSign in with a GLOBAL ADMIN account when prompted. Continue?",
        "Setup Auth", "YesNo", "Question")
    if ($r -ne "Yes") { return }
    $btnSetupAuth.Enabled = $false
    $ok = Initialize-M365Auth -LogBox $logBox
    $btnSetupAuth.Enabled = $true
    if ($ok) { $lblStatus.ForeColor = $GREEN; $lblStatus.Text = "[OK] Non-interactive auth configured. Wait 5-10 min then Scan Tenant." }
    else { $lblStatus.ForeColor = $DANGER; $lblStatus.Text = "[X] Setup Auth failed - see log." }
})

$btnProcess.Add_Click({
    # Resolve selection directly from the grid rows (authoritative for checkbox state)
    $selectedUpns = @()
    foreach ($row in $grid.Rows) {
        if ([bool]$row.Cells['Select'].Value) { $selectedUpns += [string]$row.Cells['UserPrincipalName'].Value }
    }
    $selected = @($script:candidates | Where-Object { $selectedUpns -contains $_.UserPrincipalName })

    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No accounts checked.", "Nothing selected") | Out-Null
        return
    }

    $whatIfMode = [bool]$chkWhatIf.Checked
    $preview = ($selected | Select-Object -First 25 | ForEach-Object { "  - $($_.UserPrincipalName)" }) -join "`n"
    if ($selected.Count -gt 25) { $preview += "`n  ... and $($selected.Count - 25) more" }

    $modeText = if ($whatIfMode) { "WHATIF SIMULATION (no changes will be made)" } else { "LIVE - THIS WILL MAKE REAL CHANGES" }
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Mode: $modeText`n`nFor each of the $($selected.Count) account(s) below:`n  1. Convert mailbox to Shared`n  2. Rename to `"Historical - <name>`" / `"historical-<upn>`"`n  3. Remove all assigned licenses`n`n$preview`n`nProceed?",
        "Confirm Remediation", "YesNo", "Warning")
    if ($r -ne "Yes") { return }

    $btnProcess.Enabled = $false
    $results = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($cand in $selected) {
        $i++
        $progress.Value = [int](($i / $selected.Count) * 100)
        $lblStatus.Text = "Processing $i of $($selected.Count): $($cand.UserPrincipalName)"
        [System.Windows.Forms.Application]::DoEvents()
        $results.Add((Invoke-UserRemediation -Candidate $cand -WhatIfMode $whatIfMode -LogBox $logBox))
    }
    $progress.Value = 100

    $logPath = Join-Path (Split-Path $PSCommandPath -Parent) "InactiveUserScan_Actions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $results | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8
    $failCount = @($results | Where-Object { $_.Error -ne "" }).Count

    if ($failCount -eq 0) { $lblStatus.ForeColor = $GREEN; $lblStatus.Text = "[OK] Processed $($results.Count) account(s), 0 errors. Log: $logPath" }
    else { $lblStatus.ForeColor = $WARN; $lblStatus.Text = "[!] Processed $($results.Count) account(s), $failCount error(s). Log: $logPath" }

    Write-Log "Action log written: $logPath" "SUCCESS" $logBox
    $btnProcess.Enabled = $true
})

$btnClose.Add_Click({ $form.Close() })

[void]$form.ShowDialog()
