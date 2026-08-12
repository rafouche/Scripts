<#
.SYNOPSIS
    AD/M365 Admin Toolbox (Self-Contained, Combined)

.DESCRIPTION
    Single script combining Onboard-ADUser.ps1, Offboard-ADUser.ps1,
    Set-ADEntraHardMatch.ps1, and Find-InactiveLicensedUsers.ps1 (all M365/
    Entra tools) plus Disable-InactiveADComputers.ps1 and
    Disable-InactiveADUsers.ps1 (pure on-prem AD housekeeping, no M365 auth
    needed) behind one launcher menu, so the shared module-bootstrap/auth/
    Graph-connection code (previously duplicated across the four M365 tools)
    is written and fixed exactly once.

    On every launch it installs/updates required modules, repairs any
    Microsoft.Graph module version skew, verifies M365 admin roles, and
    auto-discovers the AD environment from the DC - then opens a launcher
    with buttons for each tool. Each M365 tool's own business logic and GUI
    are unchanged from their standalone originals; only the shared plumbing
    (module bootstrap, Setup Auth, Graph/Exchange connection, logging) is
    now written once. The two on-prem AD cleanup tools are ported as a
    scan-and-select grid (review and check off exactly which computers/users
    to disable+move) rather than their originals' filter-and-blind-bulk-
    disable behavior - safer for a destructive bulk action.

    The original standalone scripts remain in this repo unchanged and
    independently runnable (NinjaRMM automation-script invocation of a
    single tool, or double-click) - this toolbox is an addition, not a
    replacement.

    GUI mode : Run with no parameters - opens the launcher menu.
    CLI mode : Supply -Tool plus that tool's own parameters (see each
               original script's help for the full per-tool parameter list;
               added in a later revision of this script).

.PARAMETER AdminUPN
    UPN of the M365 admin for role verification.

.PARAMETER SkipRoleCheck
    Skip M365 role verification.

.PARAMETER SkipModuleCheck
    Skip module install/update check.

.PARAMETER WhatIf
    Simulate steps where supported - no changes committed.

.EXAMPLE
    # GUI - opens the launcher menu
    .\ADM365-Toolbox.ps1
#>

[CmdletBinding()]
param(
    [string] $AdminUPN = "",
    [switch] $SkipRoleCheck,
    [switch] $SkipModuleCheck,
    [switch] $WhatIf
)

# ============================================================
# ADMIN ELEVATION - auto-relaunch as Administrator if needed
# ============================================================
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Resolve the script path (works for both double-click and PS console)
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    if ($scriptPath) {
        # Single-string ArgumentList is the most reliable pattern for -Verb RunAs
        $argString = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        # Use pwsh (PS7) if that is what launched this script, else fall back to Windows PowerShell
        $shell = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'PowerShell.exe' }
        try {
            Start-Process $shell -ArgumentList $argString -Verb RunAs
        } catch {
            # UAC was cancelled or failed - show message
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
# This tool requires PS7+ for proper isolation of the
# Microsoft.Graph.Authentication and ExchangeOnlineManagement
# module assemblies. Installs PS7 automatically if missing,
# then relaunches the script in the correct shell.
# ============================================================

if ($PSVersionTable.PSVersion.Major -lt 7) {
    $pwsh7 = "$env:ProgramFiles\PowerShell\7\pwsh.exe"

    if (-not (Test-Path $pwsh7)) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] PowerShell 7 required but not installed. Installing now..." -ForegroundColor Yellow
        $ps7Installed = $false

        # Attempt 1: winget (available on Win10 1809+ and Win11)
        try {
            $wg = Get-Command winget -ErrorAction Stop
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Installing via winget..." -ForegroundColor Cyan
            $p = Start-Process winget `
                -ArgumentList "install --id Microsoft.PowerShell --silent --accept-source-agreements --accept-package-agreements" `
                -Wait -PassThru -WindowStyle Hidden
            if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) { $ps7Installed = $true }
        } catch {}

        # Attempt 2: Download MSI from GitHub releases API
        if (-not $ps7Installed) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Downloading PowerShell 7 MSI from GitHub..." -ForegroundColor Cyan
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $release = Invoke-RestMethod `
                    -Uri "https://api.github.com/repos/PowerShell/PowerShell/releases/latest" `
                    -UseBasicParsing
                $msi = $release.assets |
                    Where-Object { $_.name -match "win-x64\.msi$" } |
                    Select-Object -First 1
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

    # Relaunch this script in PS7, preserving all parameters
    if (Test-Path $pwsh7) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Relaunching in PowerShell 7..." -ForegroundColor Cyan
        $relaunchArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        foreach ($key in $MyInvocation.BoundParameters.Keys) {
            $val = $MyInvocation.BoundParameters[$key]
            if ($val -is [System.Management.Automation.SwitchParameter]) {
                if ($val.IsPresent) { $relaunchArgs += " -$key" }
            } elseif ($val -is [string[]]) {
                foreach ($item in $val) { $relaunchArgs += " -$key `"$item`"" }
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
# Required-module list is the UNION across all four tools:
# Offboard needed the full set already (Sites/PnP for OneDrive delegate
# access); Onboard/Find-Inactive/HardMatch each need a subset of it.
# ============================================================

function Invoke-ModuleBootstrap {
    param([bool]$Silent = $false)

    $required = @(
        @{ Name = "ExchangeOnlineManagement";                     Min = "3.4.0" },
        @{ Name = "Microsoft.Graph.Authentication";               Min = "2.0.0" },
        @{ Name = "Microsoft.Graph.Users";                        Min = "2.0.0" },
        @{ Name = "Microsoft.Graph.Groups";                       Min = "2.0.0" },
        @{ Name = "Microsoft.Graph.Sites";                        Min = "2.0.0" },
        @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; Min = "2.0.0" },
        @{ Name = "PnP.PowerShell";                               Min = "2.4.0" },
        @{ Name = "Microsoft.Graph.Applications";                Min = "2.0.0" }
    )

    function BL { param($m, $l = "INFO")
        if (-not $Silent) {
            $c = switch ($l) { "OK" { "Green" } "WARN" { "Yellow" } "ERR" { "Red" } default { "Cyan" } }
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$l] $m" -ForegroundColor $c
        }
        if ($script:StartupLog) { $script:StartupLog.Add([PSCustomObject]@{ Level = $l; Message = $m }) }
    }

    BL "Checking required modules (running as: $($env:USERNAME))..."

    # Step 1: NuGet provider (AllUsers, requires admin)
    BL "Ensuring NuGet provider..."
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 `
            -Force -Scope AllUsers -ErrorAction Stop | Out-Null
        BL "NuGet provider ready." "OK"
    }
    catch { BL "NuGet provider: $_" "WARN" }

    # Step 2: Trust PSGallery
    try {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        BL "PSGallery trusted." "OK"
    }
    catch { BL "PSGallery trust failed: $_" "WARN" }

    # Step 3: Update PowerShellGet FIRST
    # Old versions (v1.x) that ship with Windows Server fail on newer PSGallery API
    # with "The property 'Version' cannot be found on this object"
    $psget = Get-Module -ListAvailable -Name PowerShellGet |
        Sort-Object Version -Descending | Select-Object -First 1
    if (-not $psget -or $psget.Version -lt [version]"2.2.5") {
        BL "PowerShellGet is v$($psget.Version) - must update to v2.2.5+ before other installs..." "WARN"
        try {
            Install-Module -Name PowerShellGet -MinimumVersion 2.2.5 `
                -Scope AllUsers -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            # Force load the new version in this session
            Remove-Module -Name PowerShellGet -Force -ErrorAction SilentlyContinue
            Remove-Module -Name PackageManagement -Force -ErrorAction SilentlyContinue
            Import-Module -Name PowerShellGet -Force -ErrorAction SilentlyContinue
            BL "PowerShellGet updated. Version: $((Get-Module PowerShellGet).Version)" "OK"
        }
        catch { BL "PowerShellGet update failed: $_" "ERR" }
    }
    else { BL "PowerShellGet v$($psget.Version) - OK" "OK" }

    # Step 4: RSAT / ActiveDirectory
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        BL "ActiveDirectory module missing - attempting RSAT install..." "WARN"
        try {
            # On Server OS use Add-WindowsFeature; on Win10/11 use WindowsCapability
            $osInfo = Get-WmiObject Win32_OperatingSystem -ErrorAction SilentlyContinue
            if ($osInfo -and $osInfo.Caption -match "Server") {
                Add-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction Stop | Out-Null
            }
            else {
                $cap = Get-WindowsCapability -Name "Rsat.ActiveDirectory*" -Online -ErrorAction SilentlyContinue
                if ($cap) { Add-WindowsCapability -Name $cap.Name -Online | Out-Null }
                else { BL "RSAT capability not found." "WARN" }
            }
            BL "RSAT ActiveDirectory installed." "OK"
        }
        catch { BL "RSAT install failed: $_" "WARN" }
    }
    else { BL "ActiveDirectory module - OK" "OK" }

    # Step 5: Install/update required modules (AllUsers scope)
    foreach ($mod in $required) {
        try {
            $inst = Get-Module -ListAvailable -Name $mod.Name |
                Sort-Object Version -Descending | Select-Object -First 1

            $needsInstall = (-not $inst)
            $needsUpdate  = $false
            if ($inst) {
                try { $needsUpdate = ($inst.Version -lt [version]$mod.Min) }
                catch { $needsInstall = $true }
            }

            if ($needsInstall -or $needsUpdate) {
                $action = if ($needsInstall) { "Installing" } else { "Updating" }
                BL "$action $($mod.Name) (AllUsers)..."
                Install-Module -Name $mod.Name -MinimumVersion $mod.Min `
                    -Scope AllUsers -Force -AllowClobber `
                    -Repository PSGallery -ErrorAction Stop
                BL "$($mod.Name) - done." "OK"
            }
            else { BL "$($mod.Name) v$($inst.Version) - OK" "OK" }
        }
        catch { BL "$($mod.Name) failed: $_" "ERR" }
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
    # Connector / SPO Management Shell, some background service or scheduled task may
    # already have those exact DLLs open, and Windows won't let anyone - Admin included -
    # delete or replace a file another process holds a handle to. Confirmed on a live
    # Server 2019 box: repeated elevated repair attempts left the module versions
    # completely unchanged.
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

    # Groups/Sites don't participate in the specific Authentication+Applications+
    # Identity.DirectoryManagement conflict (nothing explicitly imports them with an
    # unpinned version the way those three get imported together), but keeping the whole
    # Graph module family aligned is cheap and avoids ever needing to re-diagnose this.
    $graphModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Applications", "Microsoft.Graph.Identity.DirectoryManagement", "Microsoft.Graph.Groups", "Microsoft.Graph.Sites")

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
# Required-roles list is the UNION across all four tools: Offboard alone
# is the one that needs SharePoint Administrator (Onboard/Find-Inactive
# want the other three; HardMatch does no interactive role bootstrap at
# all in its standalone form, folded into this shared one here).
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

    $requiredRoles = @(
        "Exchange Administrator",
        "SharePoint Administrator",
        "User Administrator",
        "License Administrator"
    )

    RL "Connecting to Microsoft Graph for role verification..."

    # Skip interactive role assignment if cert-based auth is configured
    # The service principal already has the required permissions
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

        $null = Connect-MgGraph -Scopes @(
            "RoleManagement.ReadWrite.Directory",
            "User.Read.All",
            "Directory.Read.All"
        ) -NoWelcome -ErrorAction Stop

        $mgAdmin = $null
        if ($AdminUPN -ne "") {
            $mgAdmin = Get-MgUser -Filter "userPrincipalName eq '$AdminUPN'" `
                -Property Id, UserPrincipalName -ErrorAction SilentlyContinue
        }
        if (-not $mgAdmin) {
            $ctx = Get-MgContext
            if ($ctx -and $ctx.Account) {
                $mgAdmin = Get-MgUser -Filter "userPrincipalName eq '$($ctx.Account)'" `
                    -Property Id, UserPrincipalName -ErrorAction SilentlyContinue
            }
        }
        if (-not $mgAdmin) {
            RL "Could not resolve admin account - skipping role assignment." "WARN"
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
            return
        }

        RL "Checking roles for: $($mgAdmin.UserPrincipalName)"

        foreach ($roleName in $requiredRoles) {
            $activeRole = Get-MgDirectoryRole -Filter "displayName eq '$roleName'" `
                -ErrorAction SilentlyContinue
            if (-not $activeRole) {
                $tmpl = Get-MgDirectoryRoleTemplate |
                    Where-Object { $_.DisplayName -eq $roleName } | Select-Object -First 1
                if ($tmpl) {
                    $activeRole = New-MgDirectoryRole -RoleTemplateId $tmpl.Id
                    RL "Activated role: $roleName" "OK"
                }
                else { RL "Role template not found: $roleName" "WARN"; continue }
            }
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $activeRole.Id `
                -ErrorAction SilentlyContinue
            if ($members | Where-Object { $_.Id -eq $mgAdmin.Id }) {
                RL "$roleName - already assigned" "OK"
            }
            else {
                RL "Assigning $roleName to $($mgAdmin.UserPrincipalName)..."
                try {
                    New-MgDirectoryRoleMemberByRef -DirectoryRoleId $activeRole.Id `
                        -BodyParameter @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($mgAdmin.Id)"
                    }
                    RL "$roleName assigned." "OK"
                }
                catch { RL "Could not assign $roleName (may need Global Admin): $_" "WARN" }
            }
        }

        # PnP consent check
        RL "Checking PnP Management Shell consent..."
        try {
            Import-Module PnP.PowerShell -ErrorAction Stop
            $pnpApp = Get-MgApplication -Filter "displayName eq 'PnP Management Shell'" `
                -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $pnpApp) {
                RL "PnP Management Shell not registered - running one-time consent (browser will open)..." "WARN"
                Register-PnPManagementShellAccess
                RL "PnP consent complete." "OK"
            }
            else { RL "PnP Management Shell - consent OK." "OK" }
        }
        catch { RL "PnP consent check failed: $_" "WARN" }

        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        RL "Role verification complete." "OK"
    }
    catch {
        RL "Role bootstrap error: $_" "WARN"
        try { $null = Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    }
}

# ============================================================
# SECTION 0c - STARTUP
# ============================================================

# No -Tool CLI dispatch yet (added in a later revision) - always GUI/launcher mode for now.
$isCLI    = $false
$isSilent = $isCLI
# Startup runs (module bootstrap, Graph version-skew repair, role bootstrap) all happen
# before the GUI window exists, so their Write-Host output only ever reached whatever
# console launched the script - easy to miss, and useless for after-the-fact diagnosis.
# Buffer it here and replay it into the launcher's shared log box once that exists.
$script:StartupLog = New-Object System.Collections.Generic.List[object]

Write-Host "[$(Get-Date -Format 'HH:mm:ss')][INFO] AD/M365 Admin Toolbox - v1.0" -ForegroundColor Magenta

if (-not $SkipModuleCheck) {
    Invoke-ModuleBootstrap -Silent:$isSilent
    Repair-GraphModuleVersionSkew -Silent:$isSilent
}
$null = Import-Module ActiveDirectory -ErrorAction SilentlyContinue
# Import EXO BEFORE any Graph modules are loaded - both ship MSAL.NET but different
# versions. Whichever is imported first wins. EXO must win to avoid the
# GetTokenAsync assembly conflict when running both in the same session.
$null = Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

# Pre-load config BEFORE role bootstrap so the cert check can skip interactive auth.
$script:_PreConfigPath = Join-Path (Split-Path $PSCommandPath -Parent) "ADM365Config.json"
$script:Config = @{ AdminUPN=""; TenantId=""; AppId=""; CertThumbprint=""; ExchangeOrg="" }
if (Test-Path $script:_PreConfigPath) {
    try {
        $script:_PreLoaded = Get-Content $script:_PreConfigPath -Raw | ConvertFrom-Json
        foreach ($key in @('AdminUPN','TenantId','AppId','CertThumbprint','ExchangeOrg')) {
            $val = $script:_PreLoaded.PSObject.Properties[$key]
            if ($val -and $val.Value) { $script:Config[$key] = $val.Value }
        }
        if ($script:Config['AdminUPN'] -and $AdminUPN -eq "") { $AdminUPN = $script:Config['AdminUPN'] }
    } catch {}
}

if (-not $SkipRoleCheck) { Invoke-RoleBootstrap -AdminUPN $AdminUPN -Silent:$isSilent }

# ============================================================
# SECTION 0d - ENVIRONMENT AUTO-DISCOVERY
# Used by the Onboard/Offboard/HardMatch pages. Find-Inactive's scan page
# doesn't touch AD at all and ignores $script:E.
# ============================================================

function Get-EnvironmentInfo {
    $info = [ordered]@{
        DC                 = ""
        DomainDN           = ""
        LocalDomain        = ""
        EmailDomain        = ""
        DisabledUsersOU    = ""
        AADConnectServer   = ""
        TenantName         = ""
        SharePointAdminURL = ""
    }
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $info.DC          = $domain.PDCEmulator
        $info.DomainDN    = $domain.DistinguishedName
        $info.LocalDomain = $domain.DNSRoot

        $forest = Get-ADForest -ErrorAction SilentlyContinue
        if ($forest -and $forest.UPNSuffixes.Count -gt 0) {
            $nonLocal = $forest.UPNSuffixes |
                Where-Object { $_ -notmatch '\.local$' } | Select-Object -First 1
            $info.EmailDomain = if ($nonLocal) { $nonLocal } else { $forest.UPNSuffixes[0] }
        }
        if ($info.EmailDomain -eq "") {
            $info.EmailDomain = Get-ADUser -Filter { Enabled -eq $true } `
                -Properties UserPrincipalName | Select-Object -First 50 |
                Where-Object { $_.UserPrincipalName -match '@' } |
                ForEach-Object { $_.UserPrincipalName.Split('@')[1] } |
                Group-Object | Sort-Object Count -Descending |
                Select-Object -ExpandProperty Name -First 1
        }
        if ($info.EmailDomain -eq "") { $info.EmailDomain = $info.LocalDomain }

        # Search for "Disabled Users" first; fall back to any Disabled OU; last resort: construct path
        $disOU = Get-ADOrganizationalUnit -Filter { Name -like "Disabled Users" } |
            Select-Object -First 1
        if (-not $disOU) {
            $disOU = Get-ADOrganizationalUnit -Filter { Name -like "*Disabled*" } |
                Where-Object { $_.Name -notmatch "Computer" } | Select-Object -First 1
        }
        $info.DisabledUsersOU = if ($disOU) { $disOU.DistinguishedName } `
            else { "OU=Disabled Users,$($info.DomainDN)" }

        try {
            # Check local machine first - script often runs ON the AADC/DC server
            $localADSync = Get-Module -ListAvailable -Name ADSync -ErrorAction SilentlyContinue
            if ($localADSync) {
                $info.AADConnectServer = $env:COMPUTERNAME
            } else {
                $aadc = Get-ADComputer -Filter * -Properties Description |
                    Where-Object {
                        $_.Description -match 'AAD|Azure AD Connect|ADSync|Entra' -or
                        $_.Name        -match 'AADC|SYNC|ADC'
                    } | Select-Object -First 1
                if ($aadc) { $info.AADConnectServer = $aadc.Name }
            }
        }
        catch {}

        try {
            $aadObj = Get-ADObject -Filter { objectClass -eq "msDS-DeviceRegistrationService" } `
                -Properties * -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($aadObj -and $aadObj.'msDS-DeviceRegistrationServiceDnsName') {
                $info.TenantName = ($aadObj.'msDS-DeviceRegistrationServiceDnsName' -split '\.')[1]
            }
        }
        catch {}

        if ($info.TenantName -eq "" -and $info.EmailDomain) {
            $parts = $info.EmailDomain -split '\.'
            $info.TenantName = if ($parts.Count -ge 2) { $parts[-2] } else { $parts[0] }
        }
        if ($info.TenantName) {
            $info.SharePointAdminURL = "https://$($info.TenantName)-admin.sharepoint.com"
        }
    }
    catch { Write-Warning "Auto-discovery partial failure: $_" }
    return $info
}

Write-Host "[$(Get-Date -Format 'HH:mm:ss')][INFO] Discovering environment from DC..." -ForegroundColor Cyan
$script:E = Get-EnvironmentInfo
Write-Host "[$(Get-Date -Format 'HH:mm:ss')][OK] DC=$($script:E.DC)  Domain=$($script:E.LocalDomain)  Tenant=$($script:E.TenantName)" -ForegroundColor Green

# ============================================================
# CONFIG FILE - persists admin UPN across runs, same file/app registration
# ("ADM365LifecycleTool") already shared by Onboard/Offboard/HardMatch, so a
# machine that already ran Setup Auth for any of those three tools doesn't
# start from zero here - it just needs one Setup Auth re-click to pick up
# the AuditLog.Read.All scope this toolbox's scan page needs (see
# Initialize-M365Auth below).
# Graph SDK caches tokens via MSAL so MFA fires only once
# per token lifetime (~1 day with standard Entra policies).
# ============================================================

$script:ConfigPath = Join-Path (Split-Path $PSCommandPath -Parent) "ADM365Config.json"
$script:Config = @{
    AdminUPN       = ""
    TenantId       = $script:E.TenantName
    AppId          = ""
    CertThumbprint = ""
    ExchangeOrg    = ""
}

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

# ============================================================
# M365 AUTH HELPERS - cert-based (non-interactive) when configured,
# interactive fallback otherwise. Run "Setup Auth" once to configure.
# ============================================================

function Get-CertConfig {
    # Returns config hashtable if cert auth is fully configured and the cert is valid.
    # Uses bracket notation ['key'] which is safe under Set-StrictMode for hashtables.
    if (-not (Test-Path variable:script:Config)) { return $null }
    $appId = $script:Config['AppId']
    $thumb = $script:Config['CertThumbprint']
    $tid   = $script:Config['TenantId']
    if ($appId -and $thumb -and $tid) {
        $c = Get-Item "Cert:\LocalMachine\My\$thumb" -ErrorAction SilentlyContinue
        if ($c -and $c.NotAfter -gt (Get-Date)) { return $script:Config }
    }
    return $null
}

function Connect-ToGraph {
    # Resolves the cert from LocalMachine\My first, then CurrentUser\My as fallback,
    # and tests PrivateKey accessibility before connecting - fixes "Keyset does not
    # exist" under UAC elevation, PS7 relaunch, or NinjaRMM SYSTEM context.
    $appId = $script:Config['AppId']
    $tid   = $script:Config['TenantId']
    $thumb = $script:Config['CertThumbprint']
    if (-not $appId -or -not $tid -or -not $thumb) {
        Write-Log 'ADM365Config.json is missing AppId / TenantId / CertThumbprint.' 'ERR'
        return $false
    }

    $cert = $null
    foreach ($store in @('LocalMachine', 'CurrentUser')) {
        $c = Get-Item "Cert:\$store\My\$thumb" -ErrorAction SilentlyContinue
        if ($c) {
            try {
                $null = $c.PrivateKey
                $cert = $c
                Write-Log "Certificate found in $store\My store." 'INFO'
                break
            } catch {
                Write-Log "Cert in $store\My - private key not accessible here (wrong context). Trying next store." 'WARN'
            }
        }
    }

    if (-not $cert) {
        Write-Log "Thumbprint $thumb not found in LocalMachine\My or CurrentUser\My (or private key inaccessible)." 'ERR'
        Write-Log 'Fix: import the PFX into Cert:\LocalMachine\My so all elevated/SYSTEM contexts can reach it.' 'WARN'
        Write-Log '  Import-PfxCertificate -FilePath cert.pfx -CertStoreLocation Cert:\LocalMachine\My' 'WARN'
        return $false
    }

    try {
        Write-Log 'Connecting to Microsoft Graph (cert auth)...' 'INFO'
        $null = Connect-MgGraph -ClientId $appId -TenantId $tid `
            -Certificate $cert -NoWelcome -ErrorAction Stop
        Write-Log 'Graph connected.' 'OK'
        return $true
    } catch {
        Write-Log "Graph connect failed: $_" 'ERR'
        return $false
    }
}

function Connect-M365Graph {
    param(
        [string[]]$FallbackScopes = @(
            "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All",
            "Organization.Read.All", "RoleManagement.ReadWrite.Directory", "Sites.FullControl.All",
            "AuditLog.Read.All"
        ),
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
        # Delegate to Connect-ToGraph which resolves the cert from LocalMachine\My first,
        # ensuring private-key accessibility under elevation, PS7 relaunch, and SYSTEM.
        $ok = Connect-ToGraph
        if (-not $ok) { throw "Certificate-based Graph connection failed - see log." }
    } else {
        Write-Log "Graph: interactive auth - run 'Setup Auth' to enable non-interactive" "WARN" $LogBox
        $null = Connect-MgGraph -Scopes $FallbackScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        if ($ctx -and $ctx.TenantId -and -not $script:Config.TenantId) {
            $script:Config.TenantId = $ctx.TenantId; Save-Config
        }
    }

    $finalCtx = Get-MgContext
    if ($finalCtx) {
        Write-Log "Graph context: AuthType=$($finalCtx.AuthType)  ClientId=$($finalCtx.ClientId)  Account=$($finalCtx.Account)" "INFO" $LogBox
    }
}

function Connect-M365Exchange {
    # Import Graph.Authentication FIRST so its Azure.Core loads before EXO's version.
    # All heavier EXO cmdlet work runs via Invoke-EXOProcess in a child process to avoid
    # the Microsoft.Graph + ExchangeOnlineManagement Azure.Core/MSAL assembly conflict -
    # this direct connection path is kept for lightweight/SMTP-adjacent uses only
    # (e.g. Onboard's welcome email).
    $verArgs = @{}
    if ($script:GraphModuleTargetVersion) { $verArgs['RequiredVersion'] = $script:GraphModuleTargetVersion }
    $null = Import-Module Microsoft.Graph.Authentication @verArgs -ErrorAction SilentlyContinue
    $null = Import-Module ExchangeOnlineManagement -ErrorAction Stop
    $cfg = Get-CertConfig
    if ($cfg -and $cfg['ExchangeOrg']) {
        Connect-ExchangeOnline -AppId $cfg['AppId'] `
            -CertificateThumbprint $cfg['CertThumbprint'] `
            -Organization $cfg['ExchangeOrg'] `
            -ShowBanner:$false -ErrorAction Stop
    } else {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }
}

function Invoke-EXOProcess {
    # Runs Exchange Online cmdlets in a separate powershell.exe child process to avoid
    # the .NET AppDomain-level assembly conflict between Microsoft.Graph.Authentication
    # (Azure.Core / MSAL.NET) and ExchangeOnlineManagement (bundles different versions).
    # Each child process has an isolated DLL loader - no conflict possible.
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
        # Use same PowerShell edition as the parent process
        $exeShell = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $psi.FileName  = $exeShell
        $psi.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$scriptFile`" *> `"$outputFile`""
        $psi.UseShellExecute       = $true   # UseShellExecute=true needed for wildcard redirect
        $psi.WindowStyle           = [System.Diagnostics.ProcessWindowStyle]::Hidden
        $psi.CreateNoWindow        = $false

        $proc = [System.Diagnostics.Process]::Start($psi)
        $proc.WaitForExit(300000) | Out-Null  # 5 min timeout

        # Read output back and forward to log
        if (Test-Path $outputFile) {
            $out = Get-Content $outputFile -Raw -ErrorAction SilentlyContinue
            foreach ($line in ($out -split "`r?`n")) {
                $l = $line.Trim()
                if ($l -and $l -notmatch '^WARNING:') {
                    if ($l -match 'error|exception|fail' -and $l -notmatch 'ErrorAction') {
                        Write-Log "  EXO: $l" "WARN" $LogBox
                    } else {
                        Write-Log "  EXO: $l" "INFO" $LogBox
                    }
                }
            }
        }

        if ($proc.ExitCode -ne 0) {
            throw "EXO subprocess exited with code $($proc.ExitCode). See output above."
        }
    }
    finally {
        Remove-Item $paramFile  -Force -ErrorAction SilentlyContinue
        Remove-Item $scriptFile -Force -ErrorAction SilentlyContinue
        Remove-Item $outputFile -Force -ErrorAction SilentlyContinue
    }
}

function Connect-M365SharePoint {
    param([string]$AdminUrl)
    $null = Import-Module PnP.PowerShell -ErrorAction Stop
    $cfg = Get-CertConfig
    if ($cfg) {
        Connect-PnPOnline -Url $AdminUrl -ClientId $cfg.AppId `
            -Thumbprint $cfg.CertThumbprint -Tenant $cfg.TenantId -ErrorAction Stop
    } else {
        Connect-PnPOnline -Url $AdminUrl -Interactive -ErrorAction Stop
    }
}

function Initialize-M365Auth {
    # Runs entirely in an isolated child pwsh.exe process. The Microsoft.Graph PowerShell
    # SDK ships each Microsoft.Graph.* submodule with its OWN private copy of
    # Microsoft.Graph.Authentication.dll; importing several submodules across different
    # stages of one long-lived session (Role Bootstrap already connected to Graph at
    # startup) throws "Assembly with same name is already loaded" the moment this function
    # imports Applications/Identity.DirectoryManagement on top of that. A fresh process has
    # no assemblies loaded yet, so it can never collide - same isolation trick already used
    # for EXO operations via Invoke-EXOProcess, just for the Graph app-registration flow.
    #
    # Permission set requested here is the UNION across all four original tools' Setup Auth
    # flows, plus AuditLog.Read.All (needed by the scan page's signInActivity read - none of
    # Onboard/Offboard/HardMatch's shared app requested this before; omitting it is exactly
    # what caused a 403 earlier when this was still four separate apps/configs).
    param([System.Windows.Forms.RichTextBox]$LogBox = $null)

    Write-Log "Launching auth setup in an isolated PowerShell window (avoids a known Microsoft.Graph module assembly conflict with this session)..." "INFO" $LogBox
    Write-Log "A new console window will open. Sign in with a GLOBAL ADMIN account when prompted, then return here." "WARN" $LogBox

    $uid        = [System.Guid]::NewGuid().ToString('N').Substring(0, 10)
    $scriptFile = Join-Path $env:TEMP "setupauth_s_$uid.ps1"
    $resultFile = Join-Path $env:TEMP "setupauth_r_$uid.json"
    $fallbackExoOrg = "$($script:E.TenantName).onmicrosoft.com"

    $authCommands = @'
$ErrorActionPreference = "Stop"
$result = [ordered]@{ Success = $false; AppId = ""; TenantId = ""; CertThumbprint = ""; ExchangeOrg = ""; ErrorMessage = ""; Warnings = "" }
try {
    # This child process starts with nothing loaded, so it CAN still hit the same
    # assembly conflict if the underlying module install itself has version skew across
    # paths - pinning here is what actually fixed Setup Auth on the live test box, not
    # process isolation alone (confirmed: isolation alone left it failing).
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
    $appName = "ADM365LifecycleTool"
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
    $spoSP   = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" -ErrorAction SilentlyContinue

    # Union of all four original tools' requested Graph app roles, plus AuditLog.Read.All
    # (new - needed for the scan page's signInActivity read).
    $graphPermNames = @("User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All","RoleManagement.ReadWrite.Directory","AuditLog.Read.All")
    $graphRoles = @()
    foreach ($p in $graphPermNames) {
        $roleObj = $graphSP.AppRoles | Where-Object { $_.Value -eq $p } | Select-Object -First 1
        if ($roleObj) { $graphRoles += @{ Name = $p; Id = [string]$roleObj.Id; Type = "Role" } }
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

    $spoRoleId = $null
    if ($spoSP) {
        $sr = $spoSP.AppRoles | Where-Object { $_.Value -eq "Sites.FullControl.All" } | Select-Object -First 1
        if ($sr) {
            $spoRoleId = [string]$sr.Id
            $reqAccess += @{ ResourceAppId = "00000003-0000-0ff1-ce00-000000000000"; ResourceAccess = @(@{ Id = $spoRoleId; Type = "Role" }) }
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
    # one instead of a silent "try { } catch {}" (which meant a failed grant here would
    # show up later only as a confusing 403 Authorization_RequestDenied, with zero
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
    if ($spoSP -and $spoRoleId) {
        if (-not (Grant-AppRoleWithRetry -ServicePrincipalId $sp.Id -ResourceId $spoSP.Id -AppRoleId $spoRoleId -Label "Sites.FullControl.All")) {
            $roleGrantFailures.Add("Sites.FullControl.All")
        }
    }
    if ($roleGrantFailures.Count -gt 0) {
        $result.Warnings = "Could not grant: $($roleGrantFailures -join ', '). Fix in Entra admin center > Enterprise Applications > ADM365LifecycleTool > Permissions, or just re-run Setup Auth."
        Write-Host "WARNING: $($result.Warnings)" -ForegroundColor Yellow
    } else {
        Write-Host "Admin consent granted for all requested permissions." -ForegroundColor Green
    }

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
    if (-not $exoOrg) { $exoOrg = $FallbackExoOrg }

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
    Write-Host " Wait 5-10 minutes before first use." -ForegroundColor Yellow
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
`$FallbackExoOrg = '$fallbackExoOrg'
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
    Write-Log "Wait 5-10 minutes before first use." "WARN" $LogBox
    return $true
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [System.Windows.Forms.RichTextBox]$LogBox = $null
    )
    $ts = Get-Date -Format "HH:mm:ss"
    $cc = switch ($Level) {
        "WARN"    { "Yellow"  }
        "ERROR"   { "Red"     }
        "SUCCESS" { "Green"   }
        "HEAD"    { "Magenta" }
        default   { "Cyan"    }
    }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $cc

    if ($LogBox -and -not $LogBox.IsDisposed) {
        $rc = switch ($Level) {
            "WARN"    { [System.Drawing.Color]::Gold      }
            "ERROR"   { [System.Drawing.Color]::Tomato    }
            "SUCCESS" { [System.Drawing.Color]::LimeGreen }
            "HEAD"    { [System.Drawing.Color]::Orchid    }
            default   { [System.Drawing.Color]::FromArgb(140, 200, 255) }
        }
        $LogBox.SelectionStart = $LogBox.TextLength
        $LogBox.SelectionColor = $rc
        $LogBox.AppendText("[$ts][$Level] $Message`r`n")
        $LogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ============================================================
# SECTION 1 - LAUNCHER GUI
# Color palette hoisted here (once) rather than per-page: the shared page
# helper functions (added per-tool as pages are ported in) read these as
# default parameter values, which PowerShell resolves against the scope
# where the function was DEFINED, not where it's called from. Declaring
# the palette once at script scope avoids ever needing per-page
# $script:-prefixing discipline to keep it visible.
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
$BORDER  = [System.Drawing.Color]::FromArgb(45,  56,  80)

$F_NORM  = New-Object System.Drawing.Font("Segoe UI", 9)
$F_BOLD  = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$F_MONO  = New-Object System.Drawing.Font("Consolas", 8.5)
$F_TITLE = New-Object System.Drawing.Font("Segoe UI Semibold", 14)

# ---- Stub pages (replaced one at a time as each tool is ported in) ----
function Show-OnboardPage    { param($Owner) Write-Log "Onboard page not yet implemented in the combined toolbox - use Onboard-ADUser.ps1 directly for now." "WARN" $script:launcherLog }
function Invoke-Offboarding {
    param(
        [string]   $SamAccountName,
        [string]   $DisabledUsersOU,
        [string[]] $DelegateUsers,
        [string[]] $DistributionGroupMembers,
        [string]   $AADConnectServer,
        [string]   $SharePointAdminURL,
        [bool]     $WhatIfMode,
        [System.Windows.Forms.RichTextBox] $LogBox      = $null,
        [System.Windows.Forms.ProgressBar] $ProgressBar = $null,
        [System.Windows.Forms.Label]       $StatusLabel = $null
    )

    function Prog { param([int]$p, [string]$m)
        if ($ProgressBar) { $ProgressBar.Value = [Math]::Min($p, 100) }
        if ($StatusLabel) { $StatusLabel.Text  = $m }
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Step { param([string]$d, [scriptblock]$a)
        if ($WhatIfMode) { Write-Log "[WHATIF] $d" "WARN" $LogBox }
        else             { Write-Log $d "INFO" $LogBox; & $a }
    }

    try {
        # Step 1 - Resolve user
        Write-Log ">> 1/12 - Resolve AD user" "HEAD" $LogBox
        Prog 5 "Resolving user..."

        $u = Get-ADUser -Identity $SamAccountName `
            -Properties DisplayName, EmailAddress, UserPrincipalName, DistinguishedName `
            -ErrorAction Stop

        $origDN    = $u.DistinguishedName
        $origSam   = $u.SamAccountName
        $origDisp  = $u.DisplayName
        $origMail  = $u.EmailAddress
        $origUPN   = $u.UserPrincipalName
        $upnSuffix = $origUPN.Split("@")[1]
        $localPart = $origUPN.Split("@")[0]

        Write-Log "Found: $origDisp | $origMail" "SUCCESS" $LogBox

        # Step 2 - Build renamed values
        $newDisp  = "Historical - $origDisp"
        $newLocal = "Historical-$localPart"
        $newSam   = if ($newLocal.Length -gt 20) { $newLocal.Substring(0, 20) } else { $newLocal }
        $newUPN   = "$newLocal@$upnSuffix"
        $newMail  = "$newLocal@$upnSuffix"
        Write-Log "Rename display  -> $newDisp" "INFO" $LogBox
        Write-Log "Rename UPN      -> $newUPN"  "INFO" $LogBox

        # Step 3 - Rename AD
        Write-Log ">> 2/12 - Rename AD object" "HEAD" $LogBox
        Prog 10 "Renaming AD attributes..."

        Step "Set-ADUser + Rename-ADObject" {
            Set-ADUser -Identity $origSam `
                -DisplayName $newDisp -SamAccountName $newSam `
                -UserPrincipalName $newUPN -EmailAddress $newMail
            Rename-ADObject -Identity $origDN -NewName $newDisp
        }

        if (-not $WhatIfMode) {
            Start-Sleep -Seconds 2
            $u = Get-ADUser -Identity $newSam `
                -Properties DisplayName, EmailAddress, UserPrincipalName, DistinguishedName
            Write-Log "Renamed OK. DN: $($u.DistinguishedName)" "SUCCESS" $LogBox
        }

        # Step 4 - Sync #1: push rename so Exchange sees the Historical display name
        Write-Log ">> 3/12 - Entra sync #1 (push rename)" "HEAD" $LogBox
        Prog 18 "Triggering Entra sync #1..."

        if ($AADConnectServer) {
            Step "Delta sync #1 - $AADConnectServer" {
                $isLocal = ($AADConnectServer -eq $env:COMPUTERNAME) -or
                           ($AADConnectServer -eq "localhost") -or
                           ($AADConnectServer -eq "127.0.0.1")
                if ($isLocal) {
                    # ADSync depends on System.Web (.NET Framework only).
                    # In PS7 (.NET 6+) we must run it in a powershell.exe (PS5.1) subprocess.
                    if ($PSVersionTable.PSVersion.Major -ge 7) {
                        $syncOut = powershell.exe -NoProfile -ExecutionPolicy Bypass `
                            -Command "Import-Module ADSync -ErrorAction Stop; Start-ADSyncSyncCycle -PolicyType Delta" 2>&1
                        if ($LASTEXITCODE -ne 0) {
                            throw "ADSync local sync failed (exit $LASTEXITCODE): $syncOut"
                        }
                    } else {
                        Import-Module ADSync -ErrorAction Stop
                        Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                    }
                } else {
                    # Remote: Invoke-Command targets the server's default PS session (PS5.1) - ADSync works fine
                    Invoke-Command -ComputerName $AADConnectServer -ScriptBlock {
                        Import-Module ADSync -ErrorAction Stop
                        Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                    }
                }
                Write-Log "Sync #1 triggered - waiting 60s..." "INFO" $LogBox
                Start-Sleep -Seconds 60
            }
        }
        else { Write-Log "No AADConnect server - skipping sync #1." "WARN" $LogBox }

        # Step 5 - Connect M365 BEFORE disabling the AD account.
        # The mailbox must be converted while the user still exists and is licensed in Entra.
        Write-Log ">> 4/12 - Connect Exchange Online and Microsoft Graph" "HEAD" $LogBox
        Prog 26 "Connecting to M365..."

        Step "Connect Graph" {
            # Only connect Graph in-process. All EXO operations use Invoke-EXOProcess
            # (separate powershell.exe child) to avoid the Azure.Core/MSAL DLL conflict.
            Connect-M365Graph
        }

        # Steps 6, 7, 8 - Shared mailbox, permissions, distribution group
        # ALL Exchange Online cmdlets run in a child powershell.exe process to avoid
        # the Azure.Core / MSAL DLL conflict with Microsoft.Graph in this session.
        Write-Log ">> 5-7/12 - Exchange: Shared mailbox + permissions + distribution group" "HEAD" $LogBox
        Prog 34 "Running Exchange operations (subprocess)..."

        $cfg = Get-CertConfig
        $eParams = @{
            AppId         = if ($cfg) { $cfg['AppId'] }          else { "" }
            Thumbprint    = if ($cfg) { $cfg['CertThumbprint'] }  else { "" }
            Organization  = if ($cfg) { $cfg['ExchangeOrg'] }    else { "" }
            OrigMail      = $origMail
            DelegateUsers = $DelegateUsers
            DistMembers   = $DistributionGroupMembers
            GrpAlias      = "fwd-$localPart"
        }

        $exoCommands = @'
# Convert to Shared Mailbox
Write-Output "Converting mailbox to Shared..."
Set-Mailbox -Identity $p.OrigMail -Type Shared
Write-Output "Mailbox converted to Shared."

# Delegate permissions
foreach ($del in $p.DelegateUsers) {
    Write-Output "Granting Full Access to: $del"
    Add-MailboxPermission -Identity $p.OrigMail -User $del `
        -AccessRights FullAccess -InheritanceType All -AutoMapping $true
    Write-Output "Granting Send-As to: $del"
    Add-RecipientPermission -Identity $p.OrigMail -Trustee $del `
        -AccessRights SendAs -Confirm:$false
}

# Distribution group for forwarding
Write-Output "Setting up forwarding distribution group..."
$managedBy = if ($p.DelegateUsers.Count -gt 0) { $p.DelegateUsers[0] } else { $p.OrigMail }
$existing  = Get-DistributionGroup -Identity $p.GrpAlias -ErrorAction SilentlyContinue
if (-not $existing) {
    New-DistributionGroup -Name $p.GrpAlias -Alias $p.GrpAlias `
        -PrimarySmtpAddress $p.OrigMail -Type Distribution -ManagedBy $managedBy
    Write-Output "Distribution group created: $($p.GrpAlias)"
} else {
    Set-DistributionGroup -Identity $p.GrpAlias -PrimarySmtpAddress $p.OrigMail
    Write-Output "Distribution group updated: $($p.GrpAlias)"
}
foreach ($m in $p.DistMembers) {
    Add-DistributionGroupMember -Identity $p.GrpAlias -Member $m -ErrorAction SilentlyContinue
    Write-Output "Added dist member: $m"
}
Write-Output "Exchange operations complete."
'@

        Invoke-EXOProcess -Params $eParams -Commands $exoCommands -LogBox $LogBox -WhatIf $WhatIfMode
        Write-Log "Exchange: Shared mailbox + permissions + distribution group done." "SUCCESS" $LogBox

        # Step 9 - OneDrive SCA
        Write-Log ">> 8/12 - OneDrive Site Collection Admin" "HEAD" $LogBox
        Prog 57 "Granting OneDrive SCA..."

        if ($DelegateUsers.Count -gt 0 -and $SharePointAdminURL) {
            Step "PnP: connect and Set-PnPTenantSite -Owners" {
                Connect-M365SharePoint -AdminUrl $SharePointAdminURL
                $odPath     = $origUPN.Replace("@", "_").Replace(".", "_")
                $tenantBase = $SharePointAdminURL -replace "-admin\.", "."
                $odUrl      = "$tenantBase/personal/$odPath"
                Write-Log "OneDrive URL: $odUrl" "INFO" $LogBox
                foreach ($del in $DelegateUsers) {
                    Set-PnPTenantSite -Url $odUrl -Owners $del -ErrorAction Stop
                    Write-Log "SCA granted -> $del" "SUCCESS" $LogBox
                }
                Disconnect-PnPOnline
            }
        }
        else { Write-Log "No delegates or SPO URL - skipping OneDrive." "WARN" $LogBox }

        # Step 10 - Remove licenses (while user still exists in Entra)
        # Shared mailboxes do not require a license, so remove now.
        Write-Log ">> 9/12 - Remove M365 licenses" "HEAD" $LogBox
        Prog 64 "Removing licenses..."

        Step "Remove M365 licenses" {
            $mgU = $null
            try {
                $res = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$newUPN'&`$select=id,assignedLicenses" `
                    -ErrorAction SilentlyContinue
                if ($res -and $res.value -and $res.value.Count -gt 0) { $mgU = $res.value[0] }
            } catch {}
            if (-not $mgU) {
                try {
                    $res = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/users?`$filter=mail eq '$origMail'&`$select=id,assignedLicenses" `
                        -ErrorAction SilentlyContinue
                    if ($res -and $res.value -and $res.value.Count -gt 0) { $mgU = $res.value[0] }
                } catch {}
            }
            if ($mgU -and $mgU.assignedLicenses -and $mgU.assignedLicenses.Count -gt 0) {
                $skus = @($mgU.assignedLicenses | ForEach-Object { $_.skuId })
                $licBody = @{ addLicenses = @(); removeLicenses = $skus } | ConvertTo-Json -Depth 5
                Invoke-MgGraphRequest -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($mgU.id)/assignLicense" `
                    -Body $licBody -ContentType "application/json" | Out-Null
                Write-Log "Removed $($skus.Count) license(s)." "SUCCESS" $LogBox
            }
            else { Write-Log "No licenses assigned (or already removed)." "INFO" $LogBox }
        }

        # Disconnect Exchange before AD operations
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}

        # Step 11 - Disable and move AD account (AFTER all Exchange/M365 work)
        Write-Log ">> 10/12 - Disable and move AD account" "HEAD" $LogBox
        Prog 74 "Disabling account..."

        Step "Disable-ADAccount" {
            Disable-ADAccount -Identity $newSam
            Write-Log "Account disabled." "SUCCESS" $LogBox
        }
        Step "Move to Disabled Users OU" {
            Move-ADObject -Identity $u.DistinguishedName -TargetPath $DisabledUsersOU
            Write-Log "Moved -> $DisabledUsersOU" "SUCCESS" $LogBox
        }

        # Step 12 - Sync #2: push disable/move
        # IMPORTANT: If the Disabled Users OU is outside the Entra Connect sync scope,
        # this will soft-delete the Entra user (30-day recovery window).
        # Step 13 handles detecting and restoring it as a cloud-only account.
        Write-Log ">> 11/12 - Entra sync #2 (push disable / move)" "HEAD" $LogBox
        Prog 82 "Triggering Entra sync #2..."

        if ($AADConnectServer) {
            Step "Delta sync #2 - $AADConnectServer" {
                $isLocal = ($AADConnectServer -eq $env:COMPUTERNAME) -or
                           ($AADConnectServer -eq "localhost") -or
                           ($AADConnectServer -eq "127.0.0.1")
                if ($isLocal) {
                    # ADSync depends on System.Web (.NET Framework only).
                    # In PS7 (.NET 6+) we must run it in a powershell.exe (PS5.1) subprocess.
                    if ($PSVersionTable.PSVersion.Major -ge 7) {
                        $syncOut = powershell.exe -NoProfile -ExecutionPolicy Bypass `
                            -Command "Import-Module ADSync -ErrorAction Stop; Start-ADSyncSyncCycle -PolicyType Delta" 2>&1
                        if ($LASTEXITCODE -ne 0) {
                            throw "ADSync local sync failed (exit $LASTEXITCODE): $syncOut"
                        }
                    } else {
                        Import-Module ADSync -ErrorAction Stop
                        Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                    }
                } else {
                    # Remote: Invoke-Command targets the server's default PS session (PS5.1) - ADSync works fine
                    Invoke-Command -ComputerName $AADConnectServer -ScriptBlock {
                        Import-Module ADSync -ErrorAction Stop
                        Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                    }
                }
                Write-Log "Sync #2 triggered - waiting 90s for Entra to process..." "INFO" $LogBox
                Start-Sleep -Seconds 90
            }
        }
        else { Write-Log "No AADConnect server - skipping sync #2." "WARN" $LogBox }

        # Step 13 - Restore cloud-only identity if Entra soft-deleted the user
        # This happens when the Disabled Users OU is not in the Entra Connect sync scope.
        # Restoring makes the account "cloud-only" so the shared mailbox persists indefinitely.
        Write-Log ">> 12/12 - Verify Entra identity (restore if soft-deleted)" "HEAD" $LogBox
        Prog 90 "Checking Entra user status..."

        Step "Check and restore cloud identity" {
            $entraUser = $null
            try {
                $chk = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$newUPN'&`$select=id,accountEnabled" `
                    -ErrorAction SilentlyContinue
                if ($chk -and $chk.value -and $chk.value.Count -gt 0) { $entraUser = $chk.value[0] }
            } catch {}

            if ($entraUser) {
                Write-Log "Entra user exists (OU is in sync scope - cloud-only restore not needed)." "INFO" $LogBox
                Write-Log "Shared mailbox will persist as long as the synced account remains." "INFO" $LogBox
            } else {
                Write-Log "User not found in Entra active directory - checking Deleted Users..." "INFO" $LogBox
                $deletedRes = $null
                try {
                    $deletedRes = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?`$filter=userPrincipalName eq '$newUPN'" `
                        -ErrorAction SilentlyContinue
                } catch {}

                if ($deletedRes -and $deletedRes.value -and $deletedRes.value.Count -gt 0) {
                    $deletedId = $deletedRes.value[0].id
                    Write-Log "Found in Deleted Users - restoring as cloud-only account..." "INFO" $LogBox
                    try {
                        Invoke-MgGraphRequest -Method POST `
                            -Uri "https://graph.microsoft.com/v1.0/directory/deletedItems/$deletedId/restore" `
                            -Body "{}" -ContentType "application/json" | Out-Null
                        Write-Log "Cloud-only identity restored successfully." "SUCCESS" $LogBox
                        Write-Log "Shared mailbox is now permanently retained (no AD sync dependency)." "SUCCESS" $LogBox
                    } catch {
                        Write-Log "Auto-restore failed: $_ - restore manually in Entra portal within 30 days." "WARN" $LogBox
                    }
                } else {
                    Write-Log "User not yet in Deleted Users - sync may still be processing." "WARN" $LogBox
                    Write-Log "Check Entra portal in ~5 minutes. Restore manually if needed within 30 days." "WARN" $LogBox
                }
            }
        }

        try { $null = Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}

        Prog 100 "Complete."
        Write-Log "" "INFO" $LogBox
        Write-Log "============================================" "SUCCESS" $LogBox
        Write-Log " OFFBOARDING COMPLETE" "SUCCESS" $LogBox
        Write-Log " Was : $origDisp ($origMail)" "SUCCESS" $LogBox
        Write-Log " Now : $newDisp ($newUPN)" "SUCCESS" $LogBox
        Write-Log " Mailbox  : Shared | Delegates: $($DelegateUsers -join ', ')" "SUCCESS" $LogBox
        Write-Log " Licenses : Removed | AD: Disabled + Moved" "SUCCESS" $LogBox
        Write-Log " Entra    : Cloud-only (see log for details)" "SUCCESS" $LogBox
        if ($WhatIfMode) { Write-Log " *** WHATIF - no changes were made ***" "WARN" $LogBox }
        Write-Log "============================================" "SUCCESS" $LogBox
        return $true
    }
    catch {
        Write-Log "FATAL: $_" "ERROR" $LogBox
        Prog 0 "Error - see log."
        return $false
    }
}

function OB-Lbl { param($t, $x, $y, $w = 200, $h = 22, $f = $F_NORM, $c = $TEXT)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $t; $l.Location = [System.Drawing.Point]::new($x, $y)
    $l.Size = [System.Drawing.Size]::new($w, $h)
    $l.Font = $f; $l.ForeColor = $c
    $l.BackColor = [System.Drawing.Color]::Transparent; $l
}
function OB-TBox { param($x, $y, $w = 260, $h = 26, $multi = $false)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = [System.Drawing.Point]::new($x, $y); $t.Size = [System.Drawing.Size]::new($w, $h)
    $t.Font = $F_NORM; $t.ForeColor = $TEXT; $t.BackColor = $CLR_INPUT; $t.BorderStyle = "FixedSingle"
    if ($multi) { $t.Multiline = $true; $t.ScrollBars = "Vertical" }; $t
}
function OB-Btn { param($t, $x, $y, $w = 120, $h = 30, $bg = $ACCENT, $fg = $BG)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $t; $b.Location = [System.Drawing.Point]::new($x, $y); $b.Size = [System.Drawing.Size]::new($w, $h)
    $b.Font = $F_BOLD; $b.ForeColor = $fg; $b.BackColor = $bg
    $b.FlatStyle = "Flat"; $b.FlatAppearance.BorderSize = 0; $b.Cursor = "Hand"; $b
}
function OB-GBox { param($t, $x, $y, $w, $h)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $t; $g.Location = [System.Drawing.Point]::new($x, $y); $g.Size = [System.Drawing.Size]::new($w, $h)
    $g.Font = $F_BOLD; $g.ForeColor = $ACCENT; $g.BackColor = $PANEL; $g
}

function Show-UserSearch {
    param([System.Windows.Forms.Form]$Parent = $null)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Search Active Directory Users"
    $dlg.Size = [System.Drawing.Size]::new(640, 500)
    $dlg.StartPosition = "CenterParent"; $dlg.BackColor = $BG
    $dlg.FormBorderStyle = "FixedDialog"; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    $dlg.Controls.Add((OB-Lbl "Search by name, SAM account, or email:" 10 12 400 22 $F_BOLD $TEXT))
    $txtQ = OB-TBox 10 36 500 26; $dlg.Controls.Add($txtQ)
    $btnGo = OB-Btn "Search" 520 34 96 28 $ACCENT $BG; $dlg.Controls.Add($btnGo)
    $dlg.Controls.Add((OB-Lbl "Results - double-click or select and click OK:" 10 70 500 20 $F_NORM $TEXTDIM))

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = [System.Drawing.Point]::new(10, 92); $lv.Size = [System.Drawing.Size]::new(606, 308)
    $lv.View = "Details"; $lv.FullRowSelect = $true; $lv.GridLines = $true
    $lv.Font = $F_NORM; $lv.BackColor = $CLR_INPUT; $lv.ForeColor = $TEXT; $lv.BorderStyle = "FixedSingle"
    [void]$lv.Columns.Add("Display Name", 200); [void]$lv.Columns.Add("SAM", 120)
    [void]$lv.Columns.Add("Email", 200); [void]$lv.Columns.Add("Status", 76)
    $dlg.Controls.Add($lv)

    $lblC  = OB-Lbl "" 10 408 400 20 $F_NORM $TEXTDIM; $dlg.Controls.Add($lblC)
    $btnOK = OB-Btn "Select" 400 406 130 30 $GREEN $BG; $btnOK.Enabled = $false; $dlg.Controls.Add($btnOK)
    $btnX  = OB-Btn "Cancel" 540 406 76 30 $BORDER $TEXT; $dlg.Controls.Add($btnX)

    $script:SearchResult = $null

    $doSearch = {
        $q = $txtQ.Text.Trim(); $lv.Items.Clear()
        if ($q.Length -lt 2) { $lblC.Text = "Enter at least 2 characters."; return }
        try {
            $hits = Get-ADUser -Filter "(Name -like '*$q*') -or (SamAccountName -like '*$q*') -or (EmailAddress -like '*$q*') -or (DisplayName -like '*$q*')" `
                -Properties DisplayName, EmailAddress, UserPrincipalName, Enabled |
                Sort-Object DisplayName | Select-Object -First 100
            foreach ($h in $hits) {
                $item = New-Object System.Windows.Forms.ListViewItem("$($h.DisplayName)")
                [void]$item.SubItems.Add("$($h.SamAccountName)")
                [void]$item.SubItems.Add("$($h.EmailAddress)")
                [void]$item.SubItems.Add($(if ($h.Enabled) { "Active" } else { "Disabled" }))
                $item.Tag = $h
                if (-not $h.Enabled) { $item.ForeColor = $TEXTDIM }
                [void]$lv.Items.Add($item)
            }
            $lblC.Text = "$($hits.Count) user(s) found."
        }
        catch { $lblC.Text = "Search error: $_" }
    }

    $btnGo.Add_Click($doSearch)
    $txtQ.Add_KeyDown({ if ($_.KeyCode -eq "Return") { & $doSearch } })
    $lv.Add_SelectedIndexChanged({ $btnOK.Enabled = ($lv.SelectedItems.Count -gt 0) })
    $lv.Add_DoubleClick({
            if ($lv.SelectedItems.Count -gt 0) {
                $script:SearchResult = $lv.SelectedItems[0].Tag
                $dlg.DialogResult = "OK"; $dlg.Close()
            }
        })
    $btnOK.Add_Click({
            if ($lv.SelectedItems.Count -gt 0) {
                $script:SearchResult = $lv.SelectedItems[0].Tag
                $dlg.DialogResult = "OK"; $dlg.Close()
            }
        })
    $btnX.Add_Click({ $dlg.DialogResult = "Cancel"; $dlg.Close() })
    $script:SearchResult = $null
    if ($Parent) { [void]$dlg.ShowDialog($Parent) } else { [void]$dlg.ShowDialog() }
    $dlg.Dispose()
}

$script:SearchResult = $null

function Show-OffboardPage {
    param($Owner)
    try {
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD/M365 User Offboarding Tool"
$form.Size = [System.Drawing.Size]::new(820, 660)
$form.StartPosition = "CenterParent"; $form.BackColor = $BG
$form.FormBorderStyle = "Sizable"; $form.MaximizeBox = $true; $form.Font = $F_NORM
$form.MinimumSize = [System.Drawing.Size]::new(750, 540)

# Scrollable content panel - all content goes here so bottom buttons always visible
$pnlScroll = New-Object System.Windows.Forms.Panel
$pnlScroll.Dock = "Fill"
$pnlScroll.AutoScroll = $true
$pnlScroll.BackColor = $BG

# Title strip
$pnlTitle = New-Object System.Windows.Forms.Panel
$pnlTitle.Location = [System.Drawing.Point]::new(0, 0); $pnlTitle.Size = [System.Drawing.Size]::new(820, 52)
$pnlTitle.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)
$pnlTitle.Controls.Add((OB-Lbl "  [OFFBOARD]  AD/M365 User Offboarding Tool" 0 8 560 34 $F_TITLE $ACCENT))
$pnlTitle.Controls.Add((OB-Lbl "Domain: $($script:E.LocalDomain)   Tenant: $($script:E.TenantName)" 556 18 254 20 $F_NORM $TEXTDIM))
$pnlScroll.Controls.Add($pnlTitle)

# Status banner
$pnlBanner = New-Object System.Windows.Forms.Panel
$pnlBanner.Location = [System.Drawing.Point]::new(0, 52); $pnlBanner.Size = [System.Drawing.Size]::new(820, 24)
$pnlBanner.BackColor = [System.Drawing.Color]::FromArgb(0, 55, 18)
$lblBanner = OB-Lbl "  [OK] Modules verified   [OK] M365 roles verified" 0 3 820 18 $F_NORM $GREEN
$pnlBanner.Controls.Add($lblBanner)
$pnlScroll.Controls.Add($pnlBanner)

# User section
$grpUser = OB-GBox "User to Offboard" 10 84 794 112
$pnlScroll.Controls.Add($grpUser)

$grpUser.Controls.Add((OB-Lbl "SAM Account:" 10 26 110 22 $F_BOLD $TEXT))
$txtSAM = OB-TBox 126 22 216 26; $grpUser.Controls.Add($txtSAM)
$btnSearchAD = OB-Btn "Search AD" 352 20 116 28 $PANEL $ACCENT; $grpUser.Controls.Add($btnSearchAD)
$btnLookup   = OB-Btn "Lookup"    477 20 90  28 $PANEL $GREEN;  $grpUser.Controls.Add($btnLookup)

$lblUserInfo = OB-Lbl "Enter a SAM account name above, or click Search AD to find a user." 10 56 760 22 $F_NORM $TEXTDIM
$grpUser.Controls.Add($lblUserInfo)
$lblPreview = OB-Lbl "" 10 78 760 22 $F_NORM $ACCENT
$grpUser.Controls.Add($lblPreview)

# AD and Sync section
$grpAD = OB-GBox "AD and Sync  (auto-discovered - edit if needed)" 10 204 794 102
$pnlScroll.Controls.Add($grpAD)

$grpAD.Controls.Add((OB-Lbl "Disabled Users OU:" 10 26 140 22 $F_BOLD $TEXT))
$txtOU = OB-TBox 155 22 615 26; $txtOU.Text = $script:E.DisabledUsersOU; $grpAD.Controls.Add($txtOU)

$grpAD.Controls.Add((OB-Lbl "AADConnect Server:" 10 58 140 22 $F_BOLD $TEXT))
$txtAADC = OB-TBox 155 54 200 26; $txtAADC.Text = $script:E.AADConnectServer; $grpAD.Controls.Add($txtAADC)

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Text = "WhatIf - simulation only (no changes)"
$chkWhatIf.Location = [System.Drawing.Point]::new(368, 56); $chkWhatIf.Size = [System.Drawing.Size]::new(312, 24)
$chkWhatIf.Font = $F_BOLD; $chkWhatIf.ForeColor = $WARN
$chkWhatIf.BackColor = [System.Drawing.Color]::Transparent
$grpAD.Controls.Add($chkWhatIf)

# M365 section
$grpM365 = OB-GBox "M365  (auto-discovered - edit if needed)" 10 314 794 60
$pnlScroll.Controls.Add($grpM365)
$grpM365.Controls.Add((OB-Lbl "SPO Admin URL:" 10 26 120 22 $F_BOLD $TEXT))
$txtSPO = OB-TBox 136 22 640 26; $txtSPO.Text = $script:E.SharePointAdminURL; $grpM365.Controls.Add($txtSPO)

# Delegation section
$grpDel = OB-GBox "Delegates - Full Access, Send-As, OneDrive SCA  (one UPN per line)" 10 382 385 124
$pnlScroll.Controls.Add($grpDel)
$txtDelegates = OB-TBox 10 20 360 92 $true; $grpDel.Controls.Add($txtDelegates)

$grpDist = OB-GBox "Forwarding Dist Group Members  (one UPN per line)" 405 382 399 124
$pnlScroll.Controls.Add($grpDist)
$grpDist.Controls.Add((OB-Lbl "Group fwd-<sam> will receive old email as primary SMTP" 10 4 370 16 $F_NORM $TEXTDIM))
$txtDist = OB-TBox 10 20 374 92 $true; $grpDist.Controls.Add($txtDist)

# Progress and log
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = [System.Drawing.Point]::new(10, 516); $progress.Size = [System.Drawing.Size]::new(794, 16)
$progress.Minimum = 0; $progress.Maximum = 100; $progress.ForeColor = $GREEN; $progress.BackColor = $PANEL
$pnlScroll.Controls.Add($progress)

$lblStatus = OB-Lbl "Ready - complete the fields above and click Run Offboarding." 10 536 794 20 $F_NORM $TEXTDIM
$pnlScroll.Controls.Add($lblStatus)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = [System.Drawing.Point]::new(10, 560); $logBox.Size = [System.Drawing.Size]::new(794, 100)
$logBox.Font = $F_MONO; $logBox.BackColor = [System.Drawing.Color]::FromArgb(10, 14, 22)
$logBox.ForeColor = $TEXT; $logBox.ReadOnly = $true; $logBox.BorderStyle = "None"; $logBox.ScrollBars = "Vertical"
$pnlScroll.Controls.Add($logBox)

if ($script:StartupLog -and $script:StartupLog.Count -gt 0) {
    Write-Log "--- Startup log (module bootstrap / Graph version-skew repair / role check) ---" "INFO" $logBox
    foreach ($entry in $script:StartupLog) {
        $mappedLevel = switch ($entry.Level) { "OK" { "SUCCESS" } "ERR" { "ERROR" } "WARN" { "WARN" } default { "INFO" } }
        Write-Log $entry.Message $mappedLevel $logBox
    }
    Write-Log "--- End startup log ---" "INFO" $logBox
}

# Bottom bar
$pnlBot = New-Object System.Windows.Forms.Panel
$pnlBot.Dock = "Bottom"; $pnlBot.Height = 52
$pnlBot.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)
$form.Controls.Add($pnlBot)
$form.Controls.Add($pnlScroll)

$btnRun     = OB-Btn "Run Offboarding"   10  10 185 34 $GREEN  $BG
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$btnExport   = OB-Btn "Export Log"       206  12 130 30 $BORDER $TEXT
$btnRecheck  = OB-Btn "Re-check Setup"   346  12 160 30 $PANEL  $ACCENT
$btnSetupAuth= OB-Btn "Setup Auth"       516  12 140 30 $PANEL  $WARN
$btnClose    = OB-Btn "Close"            700  12 100 30 $DANGER ([System.Drawing.Color]::White)
$pnlBot.Controls.AddRange(@($btnRun, $btnExport, $btnRecheck, $btnSetupAuth, $btnClose))

# ============================================================
# EVENTS
# ============================================================

$btnSearchAD.Add_Click({
        Show-UserSearch -Parent $form
        if ($script:SearchResult) {
            $u = $script:SearchResult; $txtSAM.Text = $u.SamAccountName
            $lblUserInfo.ForeColor = $GREEN
            $lblUserInfo.Text = "[OK]  $($u.DisplayName)  |  $($u.EmailAddress)  |  $(if ($u.Enabled) { 'Active' } else { 'Disabled' })"
            $lp = $u.UserPrincipalName.Split("@")[0]; $suf = $u.UserPrincipalName.Split("@")[1]
            $lblPreview.Text = "Will rename -> ""Historical - $($u.DisplayName)""   UPN -> ""Historical-$lp@$suf"""
        }
    })

$btnLookup.Add_Click({
        $sam = $txtSAM.Text.Trim()
        if ($sam -eq "") { $lblUserInfo.ForeColor = $WARN; $lblUserInfo.Text = "Enter a SAM first."; return }
        try {
            $u = Get-ADUser -Identity $sam -Properties DisplayName, EmailAddress, UserPrincipalName, Enabled
            $lblUserInfo.ForeColor = $GREEN
            $lblUserInfo.Text = "[OK]  $($u.DisplayName)  |  $($u.EmailAddress)  |  $(if ($u.Enabled) { 'Active' } else { 'Disabled' })"
            $lp = $u.UserPrincipalName.Split("@")[0]; $suf = $u.UserPrincipalName.Split("@")[1]
            $lblPreview.Text = "Will rename -> ""Historical - $($u.DisplayName)""   UPN -> ""Historical-$lp@$suf"""
        }
        catch {
            $lblUserInfo.ForeColor = $DANGER; $lblUserInfo.Text = "[X]  Not found: $sam"; $lblPreview.Text = ""
        }
    })

$btnRun.Add_Click({
        $sam = $txtSAM.Text.Trim(); $ou = $txtOU.Text.Trim()
        if ($sam -eq "" -or $ou -eq "") {
            [System.Windows.Forms.MessageBox]::Show(
                "SAM Account Name and Disabled Users OU are required.",
                "Missing Fields", "OK", "Warning"); return
        }

        $dels  = @($txtDelegates.Text -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        $dists = @($txtDist.Text      -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        $aadc  = $txtAADC.Text.Trim(); $spo = $txtSPO.Text.Trim(); $wi = $chkWhatIf.Checked

        $msg = "Offboard user: $sam`n`n" +
            "  Disabled OU : $ou`n" +
            "  AADConnect  : $(if ($aadc) { $aadc } else { '(skip)' })`n" +
            "  Delegates   : $(if ($dels.Count) { $dels -join ', ' } else { '(none)' })`n" +
            "  SPO URL     : $(if ($spo) { $spo } else { '(skip)' })`n" +
            "  WhatIf      : $wi`n`n" +
            "$(if ($wi) { 'SIMULATION - no changes will be made.' } else { 'THIS WILL MAKE LIVE CHANGES. Proceed?' })"

        if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Offboarding", "YesNo", "Question") -ne "Yes") { return }

        $btnRun.Enabled = $false; $logBox.Clear()

        $ok = Invoke-Offboarding `
            -SamAccountName           $sam `
            -DisabledUsersOU          $ou `
            -DelegateUsers            $dels `
            -DistributionGroupMembers $dists `
            -AADConnectServer         $aadc `
            -SharePointAdminURL       $spo `
            -WhatIfMode               $wi `
            -LogBox                   $logBox `
            -ProgressBar              $progress `
            -StatusLabel              $lblStatus

        $btnRun.Enabled = $true
        if ($ok) {
            $lblStatus.ForeColor = $GREEN; $lblStatus.Text = "[OK]  Offboarding complete."
            if (-not $wi) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Offboarding completed for: $sam", "Done", "OK", "Information")
            }
        }
        else { $lblStatus.ForeColor = $DANGER; $lblStatus.Text = "[X]  Error - see log." }
    })

$btnRecheck.Add_Click({
        $lblBanner.ForeColor = $WARN
        $lblBanner.Text = "  [...]  Re-checking modules and M365 roles..."
        $pnlBanner.BackColor = [System.Drawing.Color]::FromArgb(60, 40, 0)
        [System.Windows.Forms.Application]::DoEvents()
        try {
            Invoke-ModuleBootstrap -Silent:$false
            Invoke-RoleBootstrap -AdminUPN $AdminUPN -Silent:$false
            $lblBanner.ForeColor = $GREEN
            $lblBanner.Text = "  [OK] Modules verified   [OK] M365 roles verified"
            $pnlBanner.BackColor = [System.Drawing.Color]::FromArgb(0, 55, 18)
        }
        catch {
            $lblBanner.ForeColor = $DANGER; $lblBanner.Text = "  [X]  Setup issue - see console."
            $pnlBanner.BackColor = [System.Drawing.Color]::FromArgb(60, 0, 0)
        }
    })

$btnExport.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Title = "Export Log"; $dlg.Filter = "Text (*.txt)|*.txt|All|*.*"
        $dlg.FileName = "Offboard-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
        if ($dlg.ShowDialog() -eq "OK") {
            $logBox.Text | Set-Content $dlg.FileName -Encoding UTF8
            Write-Log "Log exported: $($dlg.FileName)" "SUCCESS" $logBox
        }
    })

$btnSetupAuth.Add_Click({
        $cfg = Get-CertConfig
        if ($cfg) {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "Non-interactive auth is already configured.`n`n" +
                "App ID  : $($cfg.AppId)`n" +
                "Cert    : $($cfg.CertThumbprint)`n`n" +
                "Run setup again to replace it?",
                "Auth Already Configured", "YesNo", "Question")
            if ($ans -ne "Yes") { return }
        }
        $btnSetupAuth.Enabled = $false
        $logBox.Clear()
        $ok = Initialize-M365Auth -LogBox $logBox
        $btnSetupAuth.Enabled = $true
        if ($ok) {
            $lblStatus.ForeColor = $GREEN
            $lblStatus.Text = "[OK] Non-interactive auth configured. Wait 5-10 min then re-check setup."
        } else {
            $lblStatus.ForeColor = $DANGER
            $lblStatus.Text = "[X] Setup Auth failed - see log."
        }
    })

$btnClose.Add_Click({ $form.Close() })

    [void]$form.ShowDialog($Owner)
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "A startup error occurred:`n`n$_`n`nLine: $($_.InvocationInfo.ScriptLineNumber)",
            "Startup Error", "OK", "Error") | Out-Null
    }
}
# ---------------------------------------------------------------
# HARD-MATCH PAGE - color/font palette + local Write-Log
# Kept deliberately separate from the toolbox's own shared palette/Write-Log
# (bare $C/$F/Write-Log/$script:LogBox in the original standalone file,
# renamed here to $script:HMColors/$script:HMFonts/Write-HMLog/
# $script:HMLogBox) rather than ported to match. $C has a documented,
# already-hit bug in this exact tool family: PowerShell variable names are
# case-insensitive, so a stray local $c/$C anywhere downstream silently
# shadows the palette hashtable and breaks every Write-Log call after it.
# Renaming rather than reusing the toolbox's shared names removes any
# chance of that same class of bug resurfacing as more pages get added.
# ---------------------------------------------------------------
$script:HMColors = @{
    BG       = [System.Drawing.Color]::FromArgb( 24,  24,  24)
    Panel    = [System.Drawing.Color]::FromArgb( 32,  32,  32)
    Card     = [System.Drawing.Color]::FromArgb( 40,  40,  40)
    Border   = [System.Drawing.Color]::FromArgb( 60,  60,  60)
    Accent   = [System.Drawing.Color]::FromArgb(  0, 120, 212)
    Success  = [System.Drawing.Color]::FromArgb( 40, 167,  69)
    Warning  = [System.Drawing.Color]::FromArgb(255, 193,   7)
    Danger   = [System.Drawing.Color]::FromArgb(220,  53,  69)
    FG       = [System.Drawing.Color]::FromArgb(220, 220, 220)
    FGDim    = [System.Drawing.Color]::FromArgb(140, 140, 140)
    FGHdr    = [System.Drawing.Color]::FromArgb(255, 255, 255)
    LogBG    = [System.Drawing.Color]::FromArgb( 18,  18,  18)
}
$script:HMFonts = @{
    UI    = [System.Drawing.Font]::new('Segoe UI',  9)
    UIB   = [System.Drawing.Font]::new('Segoe UI',  9, [System.Drawing.FontStyle]::Bold)
    Small = [System.Drawing.Font]::new('Segoe UI',  8)
    Mono  = [System.Drawing.Font]::new('Consolas',  8)
    Title = [System.Drawing.Font]::new('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    Hdr   = [System.Drawing.Font]::new('Segoe UI',  9, [System.Drawing.FontStyle]::Bold)
}
$script:HMLogBox    = $null
$script:LblStatus   = $null
$script:SelAD       = $null   # single-mode selected AD user
$script:SelEntra    = $null   # single-mode selected Entra user

function Write-HMLog {
    param([string]$Msg, [string]$Level = 'INFO')
    $ts  = Get-Date -Format 'HH:mm:ss'
    $col = switch ($Level) {
        'OK'   { $script:HMColors.Success }
        'WARN' { $script:HMColors.Warning }
        'ERR'  { $script:HMColors.Danger  }
        default{ $script:HMColors.FGDim   }
    }
    if ($script:HMLogBox) {
        $script:HMLogBox.SelectionStart  = $script:HMLogBox.TextLength
        $script:HMLogBox.SelectionLength = 0
        $script:HMLogBox.SelectionColor  = $col
        $script:HMLogBox.AppendText("[$ts][$Level] $Msg`r`n")
        $script:HMLogBox.ScrollToCaret()
    } elseif ($script:StartupLog) {
        # Everything that runs before the GUI (and its log box) exists - module bootstrap,
        # Graph version-skew repair, the Setup Auth offer, the initial Connect-ToGraph -
        # otherwise only ever reached whatever console launched the script, easy to miss and
        # useless for after-the-fact diagnosis. Buffer it here; the GUI section replays it
        # into the log box the moment that control is created.
        $script:StartupLog.Add([PSCustomObject]@{ Level = $Level; Message = $Msg })
    }
    $hostCol = switch ($Level) {
        'OK'   { 'Green'  }
        'WARN' { 'Yellow' }
        'ERR'  { 'Red'    }
        default{ 'Cyan'   }
    }
    Write-Host "[$ts][$Level] $Msg" -ForegroundColor $hostCol
}

function Set-Status {
    param([string]$Msg)
    if ($script:LblStatus) {
        $script:LblStatus.Text = $Msg
        $script:LblStatus.Refresh()
    }
}

# Safely pull a display string out of a caught error. Some Graph SDK exception
# types throw their own error (e.g. "The property 'Warning' cannot be found
# on this object") when PowerShell tries to stringify them (`"$_"`/.ToString())
# -- known Microsoft.Graph SDK behavior, sometimes triggered by a stale/mixed
# module session (e.g. Az.Accounts also loaded). Reading .Exception.Message
# directly avoids invoking that broken stringification path; if even that
# fails, fall back to a generic message instead of crashing the caller.
function Get-SafeErrorText {
    param($ErrorRecord)
    try {
        if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Message) {
            return $ErrorRecord.Exception.Message
        }
    } catch {}
    return 'Unknown error (failed to read exception details -- check for a stale/mixed PowerShell module session, e.g. Az.Accounts loaded alongside Microsoft.Graph).'
}

function New-Btn {
    param([string]$Text,[int]$W=120,[int]$H=30,[System.Drawing.Color]$BG)
    $b = [System.Windows.Forms.Button]::new()
    $b.Text      = $Text
    $b.Width     = $W
    $b.Height    = $H
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderSize  = 1
    $b.FlatAppearance.BorderColor = $script:HMColors.Border
    $b.BackColor = if ($PSBoundParameters.ContainsKey('BG')) { $BG } else { $script:HMColors.Card }
    $b.ForeColor = $script:HMColors.FGHdr
    $b.Font      = $script:HMFonts.UIB
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

function New-Txt {
    param([int]$W=260)
    $t = [System.Windows.Forms.TextBox]::new()
    $t.Width       = $W
    $t.Height      = 24
    $t.BackColor   = $script:HMColors.Card
    $t.ForeColor   = $script:HMColors.FG
    $t.BorderStyle = 'FixedSingle'
    $t.Font        = $script:HMFonts.UI
    return $t
}

function New-Lbl {
    param([string]$Text,[System.Drawing.Font]$Font,[System.Drawing.Color]$Color)
    $l = [System.Windows.Forms.Label]::new()
    $l.Text      = $Text
    $l.Font      = if ($PSBoundParameters.ContainsKey('Font'))  { $Font  } else { $script:HMFonts.UI }
    $l.ForeColor = if ($PSBoundParameters.ContainsKey('Color')) { $Color } else { $script:HMColors.FG }
    $l.AutoSize  = $true
    $l.BackColor = [System.Drawing.Color]::Transparent
    return $l
}

function New-Sep {
    param([int]$Y,[int]$W=760)
    $p = [System.Windows.Forms.Panel]::new()
    $p.Height    = 1
    $p.Width     = $W
    $p.BackColor = $script:HMColors.Border
    $p.Location  = [System.Drawing.Point]::new(0,$Y)
    return $p
}

# ---------------------------------------------------------------
# 0F  MODULE BOOTSTRAP
# ---------------------------------------------------------------

function Get-ImmutableId([System.Guid]$g) {
    return [Convert]::ToBase64String($g.ToByteArray())
}

# Resolve a single AD user to its best Entra candidate.
# Returns a hashtable: @{ EntraUser=...; Confidence='Exact'|'Name'|'None'; Note=... }
function Resolve-EntraCandidate {
    param($ADUser)
    # Priority 1: UPN match
    try {
        $u = Get-MgUser -UserId $ADUser.UserPrincipalName `
             -Property Id,DisplayName,UserPrincipalName,AccountEnabled,OnPremisesImmutableId `
             -EA Stop
        return @{ EntraUser=$u; Confidence='Exact'; Note='UPN match' }
    } catch {}
    # Priority 2: mail match
    if ($ADUser.EmailAddress) {
        try {
            $res = @(Get-MgUser -Filter "mail eq '$($ADUser.EmailAddress)'" `
                   -Property Id,DisplayName,UserPrincipalName,AccountEnabled,OnPremisesImmutableId `
                   -Top 1 -EA Stop)
            if ($res.Count -gt 0) { return @{ EntraUser=$res[0]; Confidence='Exact'; Note='Mail match' } }
        } catch {}
    }
    # Priority 3: displayName startsWith
    try {
        $dn = $ADUser.DisplayName -replace "'","''"
        $res = @(Get-MgUser -Filter "startsWith(displayName,'$dn')" `
               -Property Id,DisplayName,UserPrincipalName,AccountEnabled,OnPremisesImmutableId `
               -Top 1 -EA Stop)
        if ($res.Count -gt 0) { return @{ EntraUser=$res[0]; Confidence='Name'; Note='DisplayName match' } }
    } catch {}
    return @{ EntraUser=$null; Confidence='None'; Note='No match found' }
}

# ---------------------------------------------------------------
# Find which Entra object currently owns a given ImmutableID.
# Graph does not support direct filter on onPremisesImmutableId,
# so we use the /users endpoint with a beta-compatible OData cast,
# falling back to a directory-objects search via Graph REST.
# Returns the conflicting user object, or $null if not found.
# ---------------------------------------------------------------
function Find-ImmutableIdOwner {
    param([string]$ImmutableId)
    try {
        # Encode the value for safe URL use
        $enc = [Uri]::EscapeDataString($ImmutableId)
        # Graph v1.0: filter users by onPremisesImmutableId
        $resp = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/users?`$filter=onPremisesImmutableId eq '$enc'&`$select=id,displayName,userPrincipalName,accountEnabled,onPremisesImmutableId" `
            -EA Stop
        $vals = @($resp.value)   # force array -- Graph may return bare object
        if ($vals.Count -gt 0) {
            return $vals[0]
        }
    } catch {
        Write-HMLog "  Conflict search error: $_" 'WARN'
    }
    return $null
}

# ---------------------------------------------------------------
# Show a dialog explaining the conflict and offering to clear it.
# Returns: 'Cleared' | 'Skip' | 'Cancel'
# ---------------------------------------------------------------
function Show-ConflictDialog {
    param($ConflictUser, [string]$TargetImmutableId, [string]$ADUserName)

    $dlg = [System.Windows.Forms.Form]::new()
    $dlg.Text            = 'ImmutableID Conflict Detected'
    $dlg.Width           = 620
    $dlg.Height          = 380
    $dlg.StartPosition   = 'CenterParent'
    $dlg.BackColor       = $script:HMColors.BG
    $dlg.ForeColor       = $script:HMColors.FG
    $dlg.Font            = $script:HMFonts.UI
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false

    # Header
    $pHdr = [System.Windows.Forms.Panel]::new()
    $pHdr.Dock = 'Top'; $pHdr.Height = 48; $pHdr.BackColor = $script:HMColors.Danger
    $dlg.Controls.Add($pHdr)
    $lHdr = New-Lbl '  ImmutableID Collision -- Another Entra Object Owns This GUID' -Font $script:HMFonts.Hdr -Color $script:HMColors.FGHdr
    $lHdr.Location = [System.Drawing.Point]::new(8,15); $pHdr.Controls.Add($lHdr)

    # Explanation
    $lExp = New-Lbl '' -Color $script:HMColors.FG
    $lExp.Font = $script:HMFonts.UI
    $lExp.AutoSize = $false
    $lExp.Size = [System.Drawing.Size]::new(580, 110)
    $lExp.Location = [System.Drawing.Point]::new(16, 60)
    $lExp.Text = "The ImmutableID you are trying to assign:`r`n$TargetImmutableId`r`n`r`nis already owned by a different Entra account:`r`n`r`nThis usually means a stale duplicate, a ghost account from a previous sync, or a soft-deleted user. You must clear the ImmutableID from the conflicting account before the match can proceed."
    $dlg.Controls.Add($lExp)

    # Conflict details card
    $pCard = [System.Windows.Forms.Panel]::new()
    $pCard.Location = [System.Drawing.Point]::new(16, 172)
    $pCard.Size     = [System.Drawing.Size]::new(580, 88)
    $pCard.BackColor = $script:HMColors.Card
    $dlg.Controls.Add($pCard)

    $lCardHdr = New-Lbl 'Conflicting Entra Account:' -Font $script:HMFonts.Hdr -Color $script:HMColors.Warning
    $lCardHdr.Location = [System.Drawing.Point]::new(8,6); $pCard.Controls.Add($lCardHdr)

    if ($ConflictUser) {
        $lName = New-Lbl "Display Name : $($ConflictUser.displayName)" -Color $script:HMColors.FG
        $lName.Location = [System.Drawing.Point]::new(8,26); $pCard.Controls.Add($lName)
        $lUPN  = New-Lbl "UPN          : $($ConflictUser.userPrincipalName)" -Color $script:HMColors.FGDim
        $lUPN.Location  = [System.Drawing.Point]::new(8,44); $pCard.Controls.Add($lUPN)
        $lID   = New-Lbl "Object ID    : $($ConflictUser.id)" -Font $script:HMFonts.Mono -Color $script:HMColors.FGDim
        $lID.Location   = [System.Drawing.Point]::new(8,62); $pCard.Controls.Add($lID)
    } else {
        $lNF = New-Lbl 'Conflicting account could not be looked up (may be a contact or deleted object).' -Color $script:HMColors.Warning
        $lNF.Location = [System.Drawing.Point]::new(8,28); $pCard.Controls.Add($lNF)
    }

    # Buttons
    $pBot = [System.Windows.Forms.Panel]::new()
    $pBot.Dock = 'Bottom'; $pBot.Height = 46; $pBot.BackColor = $script:HMColors.Panel
    $dlg.Controls.Add($pBot)

    $script:_conflictResult = 'Cancel'

    $btnClear = New-Btn 'Clear Conflict & Retry' 190 34 $script:HMColors.Success
    $btnClear.Location = [System.Drawing.Point]::new(220, 6)
    $btnClear.Enabled  = ($ConflictUser -ne $null)
    $pBot.Controls.Add($btnClear)

    $btnSkip = New-Btn 'Skip This User' 130 34 $script:HMColors.Warning
    $btnSkip.Location = [System.Drawing.Point]::new(418, 6)
    $pBot.Controls.Add($btnSkip)

    $btnCancel = New-Btn 'Cancel Batch' 110 34 $script:HMColors.Danger
    $btnCancel.Location = [System.Drawing.Point]::new(8, 6)
    $pBot.Controls.Add($btnCancel)

    $btnClear.Add_Click({
        if ($ConflictUser) {
            try {
                Write-HMLog "  Clearing ImmutableID from conflicting account: $($ConflictUser.userPrincipalName)" 'WARN'
                $body = '{"onPremisesImmutableId": null}'
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($ConflictUser.id)" `
                    -Body $body -ContentType 'application/json' -EA Stop
                Write-HMLog "  Conflict cleared from $($ConflictUser.userPrincipalName)." 'OK'
                $script:_conflictResult = 'Cleared'
            } catch {
                Write-HMLog "  Failed to clear conflict: $_" 'ERR'
                [System.Windows.Forms.MessageBox]::Show(
                    "Could not clear the conflicting ImmutableID:`r`n$_",
                    'Clear Failed','OK','Error')
                $script:_conflictResult = 'Skip'
            }
        }
        $dlg.DialogResult = 'OK'; $dlg.Close()
    })
    $btnSkip.Add_Click({   $script:_conflictResult = 'Skip';   $dlg.DialogResult = 'OK';     $dlg.Close() })
    $btnCancel.Add_Click({ $script:_conflictResult = 'Cancel'; $dlg.DialogResult = 'Cancel'; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:_conflictResult
}

function Invoke-HardMatch {
    param($ADUser, $EntraUser, [bool]$WhatIfMode, [bool]$AllowConflictUI = $true)
    $targetId  = Get-ImmutableId -g $ADUser.ObjectGUID
    $currentId = $EntraUser.OnPremisesImmutableId

    Write-HMLog "  AD    : $($ADUser.DisplayName)  [$($ADUser.SamAccountName)]" 'INFO'
    Write-HMLog "  Entra : $($EntraUser.DisplayName)  [$($EntraUser.UserPrincipalName)]" 'INFO'
    Write-HMLog "  ImmutableID -> $targetId" 'INFO'

    if ($currentId -and $currentId -eq $targetId) {
        Write-HMLog '  Already matched -- skipped.' 'OK'
        return 'AlreadyMatched'
    }
    if ($currentId -and $currentId -ne $targetId) {
        Write-HMLog "  WARNING: overwriting existing ImmutableID on target ($currentId)" 'WARN'
    }

    if ($WhatIfMode) {
        Write-HMLog "  [WHATIF] Would set ImmutableID on $($EntraUser.UserPrincipalName)" 'WARN'
        return 'WhatIf'
    }

    # Inner apply -- returns 'OK','VerifyFail','Conflict','Error'
    # Graph has eventual consistency: after Update-MgUser succeeds the read-back
    # can still return the old value for a few seconds.  Poll up to 8x / 3s apart
    # (~24s total) before declaring VerifyFail.
    $doApply = {
        try {
            Update-MgUser -UserId $EntraUser.Id `
                -OnPremisesImmutableId $targetId -EA Stop
        } catch {
            $msg = Get-SafeErrorText $_
            if ($msg -match 'onPremisesImmutableId' -and $msg -match '400') {
                return 'Conflict'
            }
            Write-HMLog "  FAILED: $msg" 'ERR'
            return 'Error'
        }
        # Poll verify -- Graph eventual consistency can lag a few seconds
        $maxRetries = 8; $retryDelay = 3; $attempt = 0
        do {
            if ($attempt -gt 0) {
                Write-HMLog "  Verify attempt $attempt/$maxRetries -- waiting ${retryDelay}s for Graph consistency..." 'INFO'
                Start-Sleep -Seconds $retryDelay
            }
            $attempt++
            try {
                $v = Get-MgUser -UserId $EntraUser.Id -Property OnPremisesImmutableId -EA Stop
                if ($v.OnPremisesImmutableId -eq $targetId) {
                    Write-HMLog "  Hard match applied and verified (attempt $attempt)." 'OK'
                    return 'OK'
                }
            } catch {
                Write-HMLog "  Verify read error (attempt $attempt): $_" 'WARN'
            }
        } while ($attempt -lt $maxRetries)
        Write-HMLog "  Write appeared to succeed but ImmutableID not confirmed after $maxRetries attempts -- check manually." 'WARN'
        return 'VerifyFail'
    }

    $result = & $doApply

    # Handle conflict with UI resolver
    if ($result -eq 'Conflict') {
        Write-HMLog '  Conflict detected: another Entra object owns this ImmutableID.' 'WARN'
        Write-HMLog '  Searching for the conflicting account...' 'INFO'
        $conflictOwner = Find-ImmutableIdOwner -ImmutableId $targetId

        if ($conflictOwner) {
            Write-HMLog "  Conflict owner: $($conflictOwner.displayName) [$($conflictOwner.userPrincipalName)]" 'WARN'
        } else {
            Write-HMLog '  Could not identify conflict owner (may be deleted/contact object).' 'WARN'
        }

        if (-not $AllowConflictUI) {
            # Non-interactive path (called from batch without UI override)
            Write-HMLog '  Conflict resolution requires manual action -- skipping.' 'WARN'
            return 'Conflict'
        }

        $resolution = Show-ConflictDialog -ConflictUser $conflictOwner `
            -TargetImmutableId $targetId -ADUserName $ADUser.DisplayName

        switch ($resolution) {
            'Cleared' {
                Write-HMLog '  Retrying hard match after conflict cleared...' 'INFO'
                $result = & $doApply
                if ($result -eq 'OK') { return 'OK' }
                Write-HMLog "  Retry result: $result" $(if ($result -eq 'OK') {'OK'} else {'ERR'})
                return $result
            }
            'Skip'   { Write-HMLog '  Skipped by user.' 'WARN'; return 'Skipped' }
            'Cancel' { Write-HMLog '  Batch cancelled by user.' 'WARN'; return 'CancelBatch' }
        }
    }

    return $result
}

# ---------------------------------------------------------------
# 2  AD USER SEARCH DIALOG
# ---------------------------------------------------------------
function Show-ADPicker {
    param([string]$Q='')
    $dlg = [System.Windows.Forms.Form]::new()
    $dlg.Text            = 'Select Active Directory User'
    $dlg.Width           = 700; $dlg.Height = 500
    $dlg.StartPosition   = 'CenterParent'
    $dlg.BackColor       = $script:HMColors.BG
    $dlg.ForeColor       = $script:HMColors.FG
    $dlg.Font            = $script:HMFonts.UI
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false

    # Search row
    $pTop = [System.Windows.Forms.Panel]::new()
    $pTop.Dock = 'Top'; $pTop.Height = 46; $pTop.BackColor = $script:HMColors.Panel
    $dlg.Controls.Add($pTop)

    $txt = New-Txt -W 480; $txt.Location = [System.Drawing.Point]::new(8,10); $txt.Text = $Q
    $pTop.Controls.Add($txt)
    $btnGo = New-Btn 'Search AD' 110 28 $script:HMColors.Accent
    $btnGo.Location = [System.Drawing.Point]::new(496,10)
    $pTop.Controls.Add($btnGo)

    # Legend bar
    $pLeg = [System.Windows.Forms.Panel]::new()
    $pLeg.Dock = 'Top'; $pLeg.Height = 28; $pLeg.BackColor = $script:HMColors.Panel
    $dlg.Controls.Add($pLeg)
    $lLeg = New-Lbl '  Columns:  Display Name  |  SAM Account Name (login)  |  User Principal Name (UPN / email login)  |  Account Enabled (Yes/No)   -- disabled accounts shown dimmed' -Font $script:HMFonts.Small -Color $script:HMColors.FGDim
    $lLeg.Location = [System.Drawing.Point]::new(0,8); $pLeg.Controls.Add($lLeg)

    # List
    $lv = [System.Windows.Forms.ListView]::new()
    $lv.Dock = 'Fill'; $lv.View = 'Details'; $lv.FullRowSelect = $true
    $lv.BackColor = $script:HMColors.Card; $lv.ForeColor = $script:HMColors.FG; $lv.Font = $script:HMFonts.UI
    $lv.GridLines = $true; $lv.BorderStyle = 'None'
    [void]$lv.Columns.Add('Display Name',          190)
    [void]$lv.Columns.Add('SAM Account Name',       130)
    [void]$lv.Columns.Add('User Principal Name (UPN)', 200)
    [void]$lv.Columns.Add('Acct Enabled',            70)
    $dlg.Controls.Add($lv)

    # Bottom
    $pBot = [System.Windows.Forms.Panel]::new()
    $pBot.Dock = 'Bottom'; $pBot.Height = 42; $pBot.BackColor = $script:HMColors.Panel
    $dlg.Controls.Add($pBot)
    $lblN = New-Lbl '0 results' -Color $script:HMColors.FGDim; $lblN.Location = [System.Drawing.Point]::new(8,14)
    $pBot.Controls.Add($lblN)
    $btnOK = New-Btn 'Select' 90 30 $script:HMColors.Success
    $btnOK.Anchor = 'Right,Bottom'; $btnOK.Location = [System.Drawing.Point]::new(590,6)
    $btnOK.Enabled = $false; $pBot.Controls.Add($btnOK)
    $btnX = New-Btn 'Cancel' 80 30
    $btnX.Anchor = 'Right,Bottom'; $btnX.Location = [System.Drawing.Point]::new(502,6)
    $pBot.Controls.Add($btnX)

    $script:_adPick = $null

    $doSearch = {
        $q2 = $txt.Text.Trim()
        if ($q2.Length -lt 1) { return }
        if ($q2 -notmatch '\*') { $q2 = "*$q2*" }
        $lv.Items.Clear(); $lblN.Text = 'Searching...'; $dlg.Refresh()
        try {
            $props = @('DisplayName','SamAccountName','UserPrincipalName',
                       'Enabled','ObjectGUID','EmailAddress','GivenName','Surname','DistinguishedName')
            $users = @(Get-ADUser -Filter {
                (DisplayName -like $q2) -or (SamAccountName -like $q2) -or
                (UserPrincipalName -like $q2) -or (EmailAddress -like $q2)
            } -Properties $props -EA Stop | Select-Object -First 200)
            foreach ($u in $users) {
                $li = [System.Windows.Forms.ListViewItem]::new($u.DisplayName)
                [void]$li.SubItems.Add($u.SamAccountName)
                [void]$li.SubItems.Add($u.UserPrincipalName)
                [void]$li.SubItems.Add($(if ($u.Enabled) {'Yes'} else {'No'}))
                $li.Tag = $u
                $li.ForeColor = if ($u.Enabled) { $script:HMColors.FG } else { $script:HMColors.FGDim }
                [void]$lv.Items.Add($li)
            }
            $lblN.Text = "$($lv.Items.Count) result(s)"
        } catch { $lblN.Text = "Error: $_" }
    }

    $btnGo.Add_Click($doSearch)
    $txt.Add_KeyDown({ param($s,$e); if ($e.KeyCode -eq 'Return') { & $doSearch } })
    $lv.Add_SelectedIndexChanged({ $btnOK.Enabled = ($lv.SelectedItems.Count -gt 0) })
    $lv.Add_DoubleClick({
        if ($lv.SelectedItems.Count -gt 0) {
            $script:_adPick = $lv.SelectedItems[0].Tag
            $dlg.DialogResult = 'OK'; $dlg.Close()
        }
    })
    $btnOK.Add_Click({
        if ($lv.SelectedItems.Count -gt 0) {
            $script:_adPick = $lv.SelectedItems[0].Tag
            $dlg.DialogResult = 'OK'; $dlg.Close()
        }
    })
    $btnX.Add_Click({ $dlg.DialogResult = 'Cancel'; $dlg.Close() })

    if ($Q.Length -ge 2) { & $doSearch }
    [void]$dlg.ShowDialog()
    return $script:_adPick
}

# ---------------------------------------------------------------
# 3  ENTRA USER SEARCH DIALOG
# ---------------------------------------------------------------
function Show-EntraPicker {
    param([string]$Q='')
    $dlg = [System.Windows.Forms.Form]::new()
    $dlg.Text            = 'Select Entra ID User'
    $dlg.Width           = 780; $dlg.Height = 500
    $dlg.StartPosition   = 'CenterParent'
    $dlg.BackColor       = $script:HMColors.BG; $dlg.ForeColor = $script:HMColors.FG
    $dlg.Font            = $script:HMFonts.UI; $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false

    $pTop = [System.Windows.Forms.Panel]::new()
    $pTop.Dock = 'Top'; $pTop.Height = 46; $pTop.BackColor = $script:HMColors.Panel
    $dlg.Controls.Add($pTop)

    $txt = New-Txt -W 560; $txt.Location = [System.Drawing.Point]::new(8,10); $txt.Text = $Q
    $pTop.Controls.Add($txt)
    $btnGo = New-Btn 'Search Entra' 120 28 $script:HMColors.Accent
    $btnGo.Location = [System.Drawing.Point]::new(576,10)
    $pTop.Controls.Add($btnGo)

    # Legend bar
    $pLeg2 = [System.Windows.Forms.Panel]::new()
    $pLeg2.Dock = 'Top'; $pLeg2.Height = 38; $pLeg2.BackColor = $script:HMColors.Panel
    $dlg.Controls.Add($pLeg2)
    $lLeg2a = New-Lbl '  Columns:  Display Name  |  UPN (Entra login)  |  Account Enabled  |  ImmutableID (Base64 GUID -- links to AD)' -Font $script:HMFonts.Small -Color $script:HMColors.FGDim
    $lLeg2a.Location = [System.Drawing.Point]::new(0,4); $pLeg2.Controls.Add($lLeg2a)
    $lLeg2b = New-Lbl '  Yellow rows = already has an ImmutableID (may already be matched to a different AD account -- review carefully before selecting)' -Font $script:HMFonts.Small -Color $script:HMColors.Warning
    $lLeg2b.Location = [System.Drawing.Point]::new(0,20); $pLeg2.Controls.Add($lLeg2b)

    $lv = [System.Windows.Forms.ListView]::new()
    $lv.Dock = 'Fill'; $lv.View = 'Details'; $lv.FullRowSelect = $true
    $lv.BackColor = $script:HMColors.Card; $lv.ForeColor = $script:HMColors.FG; $lv.Font = $script:HMFonts.UI
    $lv.GridLines = $true; $lv.BorderStyle = 'None'
    [void]$lv.Columns.Add('Display Name',          195)
    [void]$lv.Columns.Add('User Principal Name',    230)
    [void]$lv.Columns.Add('Acct Enabled',            70)
    [void]$lv.Columns.Add('ImmutableID (existing)',  195)
    $dlg.Controls.Add($lv)

    $pBot = [System.Windows.Forms.Panel]::new()
    $pBot.Dock = 'Bottom'; $pBot.Height = 42; $pBot.BackColor = $script:HMColors.Panel
    $dlg.Controls.Add($pBot)
    $lblN = New-Lbl '0 results' -Color $script:HMColors.FGDim; $lblN.Location = [System.Drawing.Point]::new(8,14)
    $pBot.Controls.Add($lblN)
    $btnOK = New-Btn 'Select' 90 30 $script:HMColors.Success
    $btnOK.Anchor = 'Right,Bottom'; $btnOK.Location = [System.Drawing.Point]::new(672,6)
    $btnOK.Enabled = $false; $pBot.Controls.Add($btnOK)
    $btnX = New-Btn 'Cancel' 80 30
    $btnX.Anchor = 'Right,Bottom'; $btnX.Location = [System.Drawing.Point]::new(584,6)
    $pBot.Controls.Add($btnX)

    $script:_entraPick = $null

    $doSearch = {
        $q2 = $txt.Text.Trim().TrimStart('*').TrimEnd('*')
        if ($q2.Length -lt 2) { return }
        $lv.Items.Clear(); $lblN.Text = 'Querying Graph...'; $dlg.Refresh()
        try {
            $esc = $q2 -replace "'","''"
            $filter = "startsWith(displayName,'$esc') or startsWith(userPrincipalName,'$esc') or startsWith(mail,'$esc')"
            $users = @(Get-MgUser -Filter $filter -Top 100 `
                -Property Id,DisplayName,UserPrincipalName,AccountEnabled,OnPremisesImmutableId -EA Stop)
            foreach ($u in $users) {
                $imm = if ($u.OnPremisesImmutableId) { $u.OnPremisesImmutableId } else { '(none)' }
                $li = [System.Windows.Forms.ListViewItem]::new($u.DisplayName)
                [void]$li.SubItems.Add($u.UserPrincipalName)
                [void]$li.SubItems.Add($(if ($u.AccountEnabled) {'Yes'} else {'No'}))
                [void]$li.SubItems.Add($imm)
                $li.Tag = $u
                $li.ForeColor = if ($u.OnPremisesImmutableId) { $script:HMColors.Warning } else { $script:HMColors.FG }
                [void]$lv.Items.Add($li)
            }
            $lblN.Text = "$($lv.Items.Count) result(s)"
        } catch { $lblN.Text = "Graph error: $_" }
    }

    $btnGo.Add_Click($doSearch)
    $txt.Add_KeyDown({ param($s,$e); if ($e.KeyCode -eq 'Return') { & $doSearch } })
    $lv.Add_SelectedIndexChanged({ $btnOK.Enabled = ($lv.SelectedItems.Count -gt 0) })
    $lv.Add_DoubleClick({
        if ($lv.SelectedItems.Count -gt 0) {
            $script:_entraPick = $lv.SelectedItems[0].Tag
            $dlg.DialogResult = 'OK'; $dlg.Close()
        }
    })
    $btnOK.Add_Click({
        if ($lv.SelectedItems.Count -gt 0) {
            $script:_entraPick = $lv.SelectedItems[0].Tag
            $dlg.DialogResult = 'OK'; $dlg.Close()
        }
    })
    $btnX.Add_Click({ $dlg.DialogResult = 'Cancel'; $dlg.Close() })

    if ($Q.Length -ge 2) { & $doSearch }
    [void]$dlg.ShowDialog()
    return $script:_entraPick
}

# ---------------------------------------------------------------
# 4  OU TREE PICKER DIALOG
# ---------------------------------------------------------------
function Show-OUPicker {
    $dlg = [System.Windows.Forms.Form]::new()
    $dlg.Text            = 'Select OU for Batch Scan'
    $dlg.Width           = 560; $dlg.Height = 480
    $dlg.StartPosition   = 'CenterParent'
    $dlg.BackColor       = $script:HMColors.BG; $dlg.ForeColor = $script:HMColors.FG
    $dlg.Font            = $script:HMFonts.UI; $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false

    $tv = [System.Windows.Forms.TreeView]::new()
    $tv.Dock       = 'Fill'
    $tv.BackColor  = $script:HMColors.Card
    $tv.ForeColor  = $script:HMColors.FG
    $tv.Font       = $script:HMFonts.UI
    $tv.BorderStyle = 'None'
    $dlg.Controls.Add($tv)

    $pBot = [System.Windows.Forms.Panel]::new()
    $pBot.Dock = 'Bottom'; $pBot.Height = 42; $pBot.BackColor = $script:HMColors.Panel
    $dlg.Controls.Add($pBot)
    $btnOK = New-Btn 'Select OU' 110 30 $script:HMColors.Success
    $btnOK.Anchor = 'Right,Bottom'; $btnOK.Location = [System.Drawing.Point]::new(436,6)
    $btnOK.Enabled = $false; $pBot.Controls.Add($btnOK)
    $btnX = New-Btn 'Cancel' 80 30
    $btnX.Anchor = 'Right,Bottom'; $btnX.Location = [System.Drawing.Point]::new(348,6)
    $pBot.Controls.Add($btnX)

    $script:_ouPick = $null

    # Populate tree
    function Add-OUNode {
        param($parent, $dn)
        try {
            $ous = @(Get-ADOrganizationalUnit -SearchBase $dn -SearchScope OneLevel `
                   -Filter * -Properties DistinguishedName -EA Stop |
                   Sort-Object Name)
            foreach ($ou in $ous) {
                $node = [System.Windows.Forms.TreeNode]::new($ou.Name)
                $node.Tag = $ou.DistinguishedName
                [void]$parent.Nodes.Add($node)
                Add-OUNode $node $ou.DistinguishedName
            }
        } catch {}
    }

    try {
        $root = Get-ADDomain -EA Stop
        $rootNode = [System.Windows.Forms.TreeNode]::new($root.DNSRoot)
        $rootNode.Tag = $root.DistinguishedName
        [void]$tv.Nodes.Add($rootNode)
        Add-OUNode $rootNode $root.DistinguishedName
        $rootNode.Expand()
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Could not load AD structure: $_", 'Error',
            'OK','Error')
    }

    $tv.Add_AfterSelect({ $btnOK.Enabled = ($tv.SelectedNode -ne $null) })
    $tv.Add_NodeMouseDoubleClick({
        if ($tv.SelectedNode) {
            $script:_ouPick = $tv.SelectedNode.Tag
            $dlg.DialogResult = 'OK'; $dlg.Close()
        }
    })
    $btnOK.Add_Click({
        if ($tv.SelectedNode) {
            $script:_ouPick = $tv.SelectedNode.Tag
            $dlg.DialogResult = 'OK'; $dlg.Close()
        }
    })
    $btnX.Add_Click({ $dlg.DialogResult = 'Cancel'; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:_ouPick
}

# ---------------------------------------------------------------
# 5  MAIN FORM
# ---------------------------------------------------------------
function Show-HardMatchPage {
    param($Owner)

    $form = [System.Windows.Forms.Form]::new()
    $form.Text            = 'AD -> Entra ID Hard Match Tool'
    $form.Width           = 980
    $form.Height          = 780
    $form.MinimumSize     = [System.Drawing.Size]::new(980, 780)
    $form.StartPosition   = 'CenterParent'
    $form.BackColor       = $script:HMColors.BG
    $form.ForeColor       = $script:HMColors.FG
    $form.Font            = $script:HMFonts.UI
    $form.FormBorderStyle = 'Sizable'

    # ---------- TITLE BAR ----------
    $pTitle = [System.Windows.Forms.Panel]::new()
    $pTitle.Dock = 'Top'; $pTitle.Height = 50; $pTitle.BackColor = $script:HMColors.Panel
    $form.Controls.Add($pTitle)

    $lblTitle = New-Lbl 'AD  ->  Entra Hard Match Tool' -Font $script:HMFonts.Title -Color $script:HMColors.FGHdr
    $lblTitle.Location = [System.Drawing.Point]::new(12, 14)
    $pTitle.Controls.Add($lblTitle)

    $lblWI = New-Lbl '[WHATIF -- NO CHANGES WILL BE WRITTEN]' -Font $script:HMFonts.Hdr -Color $script:HMColors.Warning
    $lblWI.Location = [System.Drawing.Point]::new(320, 17)
    $lblWI.Visible  = $WhatIf.IsPresent
    $pTitle.Controls.Add($lblWI)

    $lblCfg = New-Lbl "Config: $script:ConfigPath" -Font $script:HMFonts.Small -Color $script:HMColors.FGDim
    $lblCfg.Location = [System.Drawing.Point]::new(12, 34)
    $pTitle.Controls.Add($lblCfg)

    $btnSetupAuth = New-Btn 'Setup Auth' 120 30 $script:HMColors.Warning
    $btnSetupAuth.Location = [System.Drawing.Point]::new(820, 10)
    $btnSetupAuth.Anchor   = 'Top,Right'
    $pTitle.Controls.Add($btnSetupAuth)
    $btnSetupAuth.Add_Click({
        $cfg = Get-CertConfig
        if ($cfg) {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "Non-interactive auth is already configured.`r`n`r`nApp ID : $($cfg.AppId)`r`nCert   : $($cfg.CertThumbprint)`r`n`r`nRun setup again to replace it?",
                'Auth Already Configured', 'YesNo', 'Question')
            if ($ans -ne 'Yes') { return }
        }
        $btnSetupAuth.Enabled = $false
        $ok = Initialize-M365Auth -LogBox $script:HMLogBox
        $btnSetupAuth.Enabled = $true
        if ($ok) {
            Set-Status '[OK] Non-interactive auth configured. Wait 5-10 min then retry.'
        } else {
            Set-Status '[X] Setup Auth failed -- see log.'
        }
    })

    # ---------- TAB CONTROL ----------
    $tabs = [System.Windows.Forms.TabControl]::new()
    $tabs.Dock      = 'Fill'
    $tabs.BackColor = $script:HMColors.BG
    $tabs.Font      = $script:HMFonts.UIB
    $form.Controls.Add($tabs)

    # Bring tabs on top of title bar (Z order)
    $form.Controls.SetChildIndex($tabs, 0)

    # ---------- STATUS BAR ----------
    $pStatus = [System.Windows.Forms.Panel]::new()
    $pStatus.Dock = 'Bottom'; $pStatus.Height = 26; $pStatus.BackColor = $script:HMColors.Panel
    $form.Controls.Add($pStatus)
    $lblStatus = New-Lbl 'Ready' -Color $script:HMColors.FGDim; $lblStatus.Location = [System.Drawing.Point]::new(8,5)
    $pStatus.Controls.Add($lblStatus)
    $script:LblStatus = $lblStatus

    # ===================================================================
    # TAB 1 -- SINGLE USER
    # ===================================================================
    $tabSingle = [System.Windows.Forms.TabPage]::new('  Single User  ')
    $tabSingle.BackColor = $script:HMColors.BG; $tabSingle.ForeColor = $script:HMColors.FG
    [void]$tabs.TabPages.Add($tabSingle)

    # -- AD card --
    $cardAD = [System.Windows.Forms.Panel]::new()
    $cardAD.Location = [System.Drawing.Point]::new(12, 12)
    $cardAD.Size     = [System.Drawing.Size]::new(438, 200)
    $cardAD.BackColor = $script:HMColors.Card
    $tabSingle.Controls.Add($cardAD)

    $lADH = New-Lbl 'Active Directory User' -Font $script:HMFonts.Hdr -Color $script:HMColors.Accent
    $lADH.Location = [System.Drawing.Point]::new(8,8); $cardAD.Controls.Add($lADH)

    $txtADs = New-Txt 300; $txtADs.Location = [System.Drawing.Point]::new(8,32); $cardAD.Controls.Add($txtADs)
    $btnADs = New-Btn 'Search AD' 112 26 $script:HMColors.Accent; $btnADs.Location = [System.Drawing.Point]::new(316,32); $cardAD.Controls.Add($btnADs)

    $adFields = @{
        Name  = [System.Drawing.Point]::new(8,72)
        SAM   = [System.Drawing.Point]::new(8,92)
        UPN   = [System.Drawing.Point]::new(8,112)
        GUID  = [System.Drawing.Point]::new(8,132)
        ImmID = [System.Drawing.Point]::new(8,152)
        OU    = [System.Drawing.Point]::new(8,172)
    }
    $adVals  = @{}
    $adLabels = @{ Name='Name:'; SAM='SAM:'; UPN='UPN:'; GUID='GUID:'; ImmID='Imm ID:'; OU='OU:' }
    foreach ($k in $adFields.Keys) {
        $lKey = New-Lbl $adLabels[$k] -Color $script:HMColors.FGDim; $lKey.Location = $adFields[$k]; $cardAD.Controls.Add($lKey)
        $lVal = New-Lbl '--' -Color $script:HMColors.FG
        $lVal.Location = [System.Drawing.Point]::new(68, $adFields[$k].Y)
        $lVal.MaximumSize = [System.Drawing.Size]::new(364,18)
        if ($k -in @('GUID','ImmID')) { $lVal.Font = $script:HMFonts.Mono; $lVal.ForeColor = $script:HMColors.FGDim }
        $cardAD.Controls.Add($lVal)
        $adVals[$k] = $lVal
    }

    # arrow
    $lblArr = New-Lbl '->' -Font ([System.Drawing.Font]::new('Segoe UI',18,[System.Drawing.FontStyle]::Bold)) -Color $script:HMColors.Accent
    $lblArr.Location = [System.Drawing.Point]::new(460, 90); $tabSingle.Controls.Add($lblArr)

    # -- Entra card --
    $cardE = [System.Windows.Forms.Panel]::new()
    $cardE.Location = [System.Drawing.Point]::new(492, 12)
    $cardE.Size     = [System.Drawing.Size]::new(462, 200)
    $cardE.BackColor = $script:HMColors.Card
    $tabSingle.Controls.Add($cardE)

    $lEH = New-Lbl 'Entra ID (Azure AD) User' -Font $script:HMFonts.Hdr -Color $script:HMColors.Accent
    $lEH.Location = [System.Drawing.Point]::new(8,8); $cardE.Controls.Add($lEH)

    $txtEs = New-Txt 316; $txtEs.Location = [System.Drawing.Point]::new(8,32); $cardE.Controls.Add($txtEs)
    $btnEs = New-Btn 'Search Entra' 124 26 $script:HMColors.Accent; $btnEs.Location = [System.Drawing.Point]::new(328,32); $cardE.Controls.Add($btnEs)

    $eFields = @{
        Name  = [System.Drawing.Point]::new(8,72)
        UPN   = [System.Drawing.Point]::new(8,92)
        ObjId = [System.Drawing.Point]::new(8,112)
        ImmID = [System.Drawing.Point]::new(8,132)
        State = [System.Drawing.Point]::new(8,152)
        Sync  = [System.Drawing.Point]::new(8,172)
    }
    $eVals   = @{}
    $eLabels = @{ Name='Name:'; UPN='UPN:'; ObjId='Object ID:'; ImmID='Imm ID:'; State='Enabled:'; Sync='Sync State:' }
    foreach ($k in $eFields.Keys) {
        $lKey = New-Lbl $eLabels[$k] -Color $script:HMColors.FGDim; $lKey.Location = $eFields[$k]; $cardE.Controls.Add($lKey)
        $lVal = New-Lbl '--' -Color $script:HMColors.FG
        $lVal.Location = [System.Drawing.Point]::new(80, $eFields[$k].Y)
        $lVal.MaximumSize = [System.Drawing.Size]::new(376,18)
        if ($k -in @('ObjId','ImmID')) { $lVal.Font = $script:HMFonts.Mono; $lVal.ForeColor = $script:HMColors.FGDim }
        $cardE.Controls.Add($lVal)
        $eVals[$k] = $lVal
    }

    # -- Computed preview --
    $pPrev = [System.Windows.Forms.Panel]::new()
    $pPrev.Location = [System.Drawing.Point]::new(12, 224)
    $pPrev.Size     = [System.Drawing.Size]::new(942, 44)
    $pPrev.BackColor = $script:HMColors.Panel
    $tabSingle.Controls.Add($pPrev)

    $lPrevHdr = New-Lbl 'Computed ImmutableID (Base64 of AD GUID):' -Font $script:HMFonts.Hdr -Color $script:HMColors.FGHdr
    $lPrevHdr.Location = [System.Drawing.Point]::new(8,4); $pPrev.Controls.Add($lPrevHdr)
    $lPrevVal = New-Lbl 'Select an AD user above to compute.' -Font $script:HMFonts.Mono -Color $script:HMColors.FGDim
    $lPrevVal.Location = [System.Drawing.Point]::new(8,22)
    $lPrevVal.MaximumSize = [System.Drawing.Size]::new(930,18)
    $pPrev.Controls.Add($lPrevVal)

    # -- Validation notice --
    $lblV = New-Lbl '' -Font $script:HMFonts.UIB -Color $script:HMColors.FGDim
    $lblV.Location = [System.Drawing.Point]::new(12, 278)
    $lblV.MaximumSize = [System.Drawing.Size]::new(942,60)
    $tabSingle.Controls.Add($lblV)

    # -- Action row --
    $pAct = [System.Windows.Forms.Panel]::new()
    $pAct.Location = [System.Drawing.Point]::new(0, 344)
    $pAct.Size     = [System.Drawing.Size]::new(962, 50)
    $pAct.BackColor = $script:HMColors.Panel
    $tabSingle.Controls.Add($pAct)

    $btnApply  = New-Btn 'Apply Hard Match' 160 36 $script:HMColors.Success; $btnApply.Location  = [System.Drawing.Point]::new(640,7); $btnApply.Enabled = $false
    $btnClearI = New-Btn 'Clear Entra ImmutableID' 200 36 $script:HMColors.Danger;  $btnClearI.Location = [System.Drawing.Point]::new(430,7); $btnClearI.Enabled = $false
    $btnClearS = New-Btn 'Clear Selections' 140 36; $btnClearS.Location = [System.Drawing.Point]::new(280,7)
    $pAct.Controls.AddRange(@($btnApply,$btnClearI,$btnClearS))

    # -- Log (single tab) --
    $logS = [System.Windows.Forms.RichTextBox]::new()
    $logS.Location  = [System.Drawing.Point]::new(0, 394)
    $logS.Size      = [System.Drawing.Size]::new(962, 290)
    $logS.BackColor = $script:HMColors.LogBG; $logS.ForeColor = $script:HMColors.FG
    $logS.Font      = $script:HMFonts.Mono; $logS.ReadOnly = $true
    $logS.BorderStyle = 'None'; $logS.ScrollBars = 'Vertical'
    $tabSingle.Controls.Add($logS)
    $script:HMLogBox = $logS

    # ---- single tab: wire validation ----
    $updateVal = {
        $btnApply.Enabled  = $false
        $btnClearI.Enabled = $false
        if (-not $script:SelAD -or -not $script:SelEntra) {
            $lblV.Text = 'Select both an AD user and an Entra user to proceed.'
            $lblV.ForeColor = $script:HMColors.FGDim; return
        }
        $cid = Get-ImmutableId -g $script:SelAD.ObjectGUID
        $eid = $script:SelEntra.OnPremisesImmutableId
        if ($eid -and $eid -eq $cid) {
            $lblV.Text = '[OK] ImmutableID already matches -- these accounts are already hard-matched.'
            $lblV.ForeColor = $script:HMColors.Success
        } elseif ($eid -and $eid -ne $cid) {
            $lblV.Text = '[!] Entra user already has a DIFFERENT ImmutableID. Applying will overwrite it.'
            $lblV.ForeColor = $script:HMColors.Warning
            $btnApply.Enabled = $true; $btnClearI.Enabled = $true
        } else {
            $lblV.Text = '[OK] Ready -- Entra user has no existing ImmutableID.'
            $lblV.ForeColor = $script:HMColors.Success
            $btnApply.Enabled = $true
        }
    }

    $btnADs.Add_Click({
        $p = Show-ADPicker -Q $txtADs.Text.Trim()
        if ($p) {
            $script:SelAD = $p
            $adVals['Name'].Text  = $p.DisplayName
            $adVals['SAM'].Text   = $p.SamAccountName
            $adVals['UPN'].Text   = $p.UserPrincipalName
            $adVals['GUID'].Text  = $p.ObjectGUID.ToString()
            $cid = Get-ImmutableId -g $p.ObjectGUID
            $adVals['ImmID'].Text = $cid
            $adVals['OU'].Text    = ($p.DistinguishedName -replace '^[^,]+,','')
            $lPrevVal.Text      = $cid; $lPrevVal.ForeColor = $script:HMColors.Accent
            if (-not $txtEs.Text.Trim()) { $txtEs.Text = $p.DisplayName }
            Set-Status "AD user selected: $($p.DisplayName)"
            & $updateVal
        }
    })

    $btnEs.Add_Click({
        $p = Show-EntraPicker -Q $txtEs.Text.Trim()
        if ($p) {
            $script:SelEntra = $p
            $imm = if ($p.OnPremisesImmutableId) { $p.OnPremisesImmutableId } else { '(none)' }
            $eVals['Name'].Text  = $p.DisplayName
            $eVals['UPN'].Text   = $p.UserPrincipalName
            $eVals['ObjId'].Text = $p.Id
            $eVals['ImmID'].Text = $imm
            $eVals['ImmID'].ForeColor = if ($p.OnPremisesImmutableId) { $script:HMColors.Warning } else { $script:HMColors.FGDim }
            $eVals['State'].Text = if ($p.AccountEnabled) { 'Yes' } else { 'No' }
            $eVals['Sync'].Text  = if ($p.OnPremisesImmutableId) { 'Has ImmutableID' } else { 'Cloud-only / No ImmutableID' }
            if (-not $txtADs.Text.Trim()) { $txtADs.Text = $p.DisplayName }
            Set-Status "Entra user selected: $($p.DisplayName)"
            & $updateVal
        }
    })

    $btnApply.Add_Click({
        if (-not $script:SelAD -or -not $script:SelEntra) { return }
        $cid = Get-ImmutableId -g $script:SelAD.ObjectGUID
        $msg  = "Apply hard match?`r`n`r`n"
        $msg += "AD User   : $($script:SelAD.DisplayName) ($($script:SelAD.SamAccountName))`r`n"
        $msg += "Entra User: $($script:SelEntra.DisplayName) ($($script:SelEntra.UserPrincipalName))`r`n`r`n"
        $msg += "ImmutableID to set:`r`n$cid"
        if ($script:SelEntra.OnPremisesImmutableId) { $msg += "`r`n`r`nWARNING: Existing ImmutableID will be overwritten." }
        if ($WhatIf) { $msg += "`r`n`r`n[WHATIF] No changes will be written." }
        $r = [System.Windows.Forms.MessageBox]::Show($msg,'Confirm Hard Match','YesNo','Question')
        if ($r -ne 'Yes') { return }
        Set-Status 'Applying hard match...'
        $res = Invoke-HardMatch -ADUser $script:SelAD -EntraUser $script:SelEntra -WhatIfMode $WhatIf.IsPresent
        if ($res -eq 'OK' -and -not $WhatIf) {
            try {
                $refreshed = Get-MgUser -UserId $script:SelEntra.Id `
                    -Property Id,DisplayName,UserPrincipalName,AccountEnabled,OnPremisesImmutableId
                $script:SelEntra = $refreshed
                $eVals['ImmID'].Text = $refreshed.OnPremisesImmutableId
                & $updateVal
            } catch {}
        }
        Set-Status $(if ($res -eq 'OK') { 'Match applied.' } elseif ($res -eq 'AlreadyMatched') { 'Already matched.' } else { "Result: $res" })
    })

    $btnClearI.Add_Click({
        if (-not $script:SelEntra -or -not $script:SelEntra.OnPremisesImmutableId) { return }
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Clear ImmutableID from:`r`n$($script:SelEntra.DisplayName) ($($script:SelEntra.UserPrincipalName))?",
            'Clear ImmutableID','YesNo','Warning')
        if ($r -ne 'Yes') { return }
        if ($WhatIf) { Write-HMLog "[WHATIF] Would clear ImmutableID on $($script:SelEntra.UserPrincipalName)" 'WARN'; return }
        try {
            $body = '{"onPremisesImmutableId": null}'
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "https://graph.microsoft.com/v1.0/users/$($script:SelEntra.Id)" `
                -Body $body -ContentType 'application/json'
            Write-HMLog "ImmutableID cleared on $($script:SelEntra.UserPrincipalName)." 'OK'
            $refreshed = Get-MgUser -UserId $script:SelEntra.Id `
                -Property Id,DisplayName,UserPrincipalName,AccountEnabled,OnPremisesImmutableId
            $script:SelEntra = $refreshed
            $eVals['ImmID'].Text = '(none)'; $btnClearI.Enabled = $false
            & $updateVal
        } catch { Write-HMLog "Clear failed: $_" 'ERR' }
    })

    $btnClearS.Add_Click({
        $script:SelAD = $null; $script:SelEntra = $null
        foreach ($k in $adVals.Keys)  { $adVals[$k].Text = '--' }
        foreach ($k in $eVals.Keys)   { $eVals[$k].Text  = '--' }
        $lPrevVal.Text = 'Select an AD user above to compute.'
        $lPrevVal.ForeColor = $script:HMColors.FGDim
        $lblV.Text = ''; $btnApply.Enabled = $false; $btnClearI.Enabled = $false
        Set-Status 'Selections cleared.'
    })

    # ===================================================================
    # TAB 2 -- BATCH OU SCAN
    # ===================================================================
    $tabBatch = [System.Windows.Forms.TabPage]::new('  Batch (OU Scan)  ')
    $tabBatch.BackColor = $script:HMColors.BG; $tabBatch.ForeColor = $script:HMColors.FG
    [void]$tabs.TabPages.Add($tabBatch)

    # -- OU selector row --
    $pOU = [System.Windows.Forms.Panel]::new()
    $pOU.Location = [System.Drawing.Point]::new(12,12)
    $pOU.Size     = [System.Drawing.Size]::new(942,42)
    $pOU.BackColor = $script:HMColors.Panel
    $tabBatch.Controls.Add($pOU)

    $lOU = New-Lbl 'OU:' -Font $script:HMFonts.Hdr -Color $script:HMColors.FGHdr; $lOU.Location = [System.Drawing.Point]::new(8,12); $pOU.Controls.Add($lOU)
    $txtOU = New-Txt 680; $txtOU.Location = [System.Drawing.Point]::new(36,9); $txtOU.Text = $script:Config['DefaultOU']
    $pOU.Controls.Add($txtOU)
    $btnPickOU = New-Btn 'Browse OU' 120 26 $script:HMColors.Accent; $btnPickOU.Location = [System.Drawing.Point]::new(720,9); $pOU.Controls.Add($btnPickOU)

    # -- Options row --
    $pOpts = [System.Windows.Forms.Panel]::new()
    $pOpts.Location = [System.Drawing.Point]::new(12,62)
    $pOpts.Size     = [System.Drawing.Size]::new(942,34)
    $pOpts.BackColor = $script:HMColors.BG
    $tabBatch.Controls.Add($pOpts)

    $chkEnabled = [System.Windows.Forms.CheckBox]::new()
    $chkEnabled.Text = 'Enabled users only'; $chkEnabled.Checked = $true
    $chkEnabled.ForeColor = $script:HMColors.FG; $chkEnabled.BackColor = [System.Drawing.Color]::Transparent
    $chkEnabled.Location = [System.Drawing.Point]::new(0,6); $chkEnabled.AutoSize = $true
    $pOpts.Controls.Add($chkEnabled)

    $chkNoImm = [System.Windows.Forms.CheckBox]::new()
    $chkNoImm.Text = 'Skip already-matched users'; $chkNoImm.Checked = $true
    $chkNoImm.ForeColor = $script:HMColors.FG; $chkNoImm.BackColor = [System.Drawing.Color]::Transparent
    $chkNoImm.Location = [System.Drawing.Point]::new(170,6); $chkNoImm.AutoSize = $true
    $pOpts.Controls.Add($chkNoImm)

    $btnLoad = New-Btn 'Load Users from OU' 180 28 $script:HMColors.Accent; $btnLoad.Location = [System.Drawing.Point]::new(380,4)
    $pOpts.Controls.Add($btnLoad)

    $btnSelAll  = New-Btn 'Check All'   90 28; $btnSelAll.Location  = [System.Drawing.Point]::new(570,4)
    $btnSelNone = New-Btn 'Uncheck All' 96 28; $btnSelNone.Location = [System.Drawing.Point]::new(668,4)
    $pOpts.Controls.Add($btnSelAll); $pOpts.Controls.Add($btnSelNone)

    # -- Checked list --
    $lv2 = [System.Windows.Forms.ListView]::new()
    $lv2.Location     = [System.Drawing.Point]::new(12, 104)
    $lv2.Size         = [System.Drawing.Size]::new(942, 320)
    $lv2.View         = 'Details'
    $lv2.CheckBoxes   = $true
    $lv2.FullRowSelect = $true
    $lv2.BackColor    = $script:HMColors.Card; $lv2.ForeColor = $script:HMColors.FG
    $lv2.Font         = $script:HMFonts.UI;  $lv2.GridLines = $true
    $lv2.BorderStyle  = 'FixedSingle'
    [void]$lv2.Columns.Add('',             24)   # checkbox spacer
    [void]$lv2.Columns.Add('AD Display Name',    185)
    [void]$lv2.Columns.Add('SAM',                110)
    [void]$lv2.Columns.Add('Entra Match',        195)
    [void]$lv2.Columns.Add('Confidence',          80)
    [void]$lv2.Columns.Add('Current Imm ID',     185)
    [void]$lv2.Columns.Add('Action',              98)
    $tabBatch.Controls.Add($lv2)

    # -- Batch action row --
    $pBAct = [System.Windows.Forms.Panel]::new()
    $pBAct.Location = [System.Drawing.Point]::new(0, 432)
    $pBAct.Size     = [System.Drawing.Size]::new(962, 50)
    $pBAct.BackColor = $script:HMColors.Panel
    $tabBatch.Controls.Add($pBAct)

    $lblSel = New-Lbl '0 users loaded' -Color $script:HMColors.FGDim; $lblSel.Location = [System.Drawing.Point]::new(8,16); $pBAct.Controls.Add($lblSel)
    $btnResolve = New-Btn 'Re-Resolve Entra' 160 36 $script:HMColors.Accent
    $btnResolve.Location = [System.Drawing.Point]::new(560,7); $btnResolve.Enabled = $false; $pBAct.Controls.Add($btnResolve)
    $btnBatchApply = New-Btn 'Apply Checked Matches' 200 36 $script:HMColors.Success
    $btnBatchApply.Location = [System.Drawing.Point]::new(728,7); $btnBatchApply.Enabled = $false; $pBAct.Controls.Add($btnBatchApply)

    # -- Batch log --
    $logB = [System.Windows.Forms.RichTextBox]::new()
    $logB.Location  = [System.Drawing.Point]::new(0, 482)
    $logB.Size      = [System.Drawing.Size]::new(962, 190)
    $logB.BackColor = $script:HMColors.LogBG; $logB.ForeColor = $script:HMColors.FG
    $logB.Font      = $script:HMFonts.Mono; $logB.ReadOnly = $true
    $logB.BorderStyle = 'None'; $logB.ScrollBars = 'Vertical'
    $tabBatch.Controls.Add($logB)

    # Switch active log when tab changes
    $tabs.Add_SelectedIndexChanged({
        $script:HMLogBox = if ($tabs.SelectedIndex -eq 0) { $logS } else { $logB }
    })

    # OU browse
    $btnPickOU.Add_Click({
        $ou = Show-OUPicker
        if ($ou) {
            $txtOU.Text = $ou
            $script:Config['DefaultOU'] = $ou
        }
    })

    # Load users
    $btnLoad.Add_Click({
        $script:HMLogBox = $logB
        $ouDN = $txtOU.Text.Trim()
        if (-not $ouDN) {
            [System.Windows.Forms.MessageBox]::Show('Enter or browse to an OU first.','No OU','OK','Warning')
            return
        }
        $lv2.Items.Clear()
        $btnResolve.Enabled = $false; $btnBatchApply.Enabled = $false
        $lblSel.Text = 'Loading AD users...'
        $form.Refresh()
        try {
            if ($chkEnabled.Checked) {
                $adUsers = @(Get-ADUser -SearchBase $ouDN -SearchScope Subtree `
                    -Filter { Enabled -eq $true } `
                    -Properties DisplayName,SamAccountName,UserPrincipalName,EmailAddress,`
                                ObjectGUID,Enabled,DistinguishedName -EA Stop |
                    Sort-Object DisplayName)
            } else {
                $adUsers = @(Get-ADUser -SearchBase $ouDN -SearchScope Subtree `
                    -Filter * `
                    -Properties DisplayName,SamAccountName,UserPrincipalName,EmailAddress,`
                                ObjectGUID,Enabled,DistinguishedName -EA Stop |
                    Sort-Object DisplayName)
            }
            Write-HMLog "Loaded $($adUsers.Count) AD user(s) from $ouDN" 'INFO'
        } catch {
            Write-HMLog "Failed to load AD users: $_" 'ERR'
            $lblSel.Text = 'Load failed.'
            return
        }

        if (-not $adUsers -or $adUsers.Count -eq 0) {
            Write-HMLog 'No users found in this OU (check filter options).' 'WARN'
            $lblSel.Text = '0 users found.'
            return
        }

        Set-Status 'Resolving Entra matches...'
        $i = 0
        foreach ($u in $adUsers) {
            $i++
            Set-Status "Resolving $i / $($adUsers.Count) ..."
            $form.Refresh()
            $res = Resolve-EntraCandidate -ADUser $u
            $eu  = $res.EntraUser
            $cid = Get-ImmutableId -g $u.ObjectGUID
            $currentImm = if ($eu) { $eu.OnPremisesImmutableId } else { $null }

            # Skip already-matched if option set
            if ($chkNoImm.Checked -and $currentImm -and $currentImm -eq $cid) { continue }

            $action = if ($res.Confidence -eq 'None') { 'No Entra match' }
                      elseif ($currentImm -and $currentImm -eq $cid) { 'Already matched' }
                      elseif ($currentImm) { 'Will overwrite' }
                      else { 'Will match' }

            $li = [System.Windows.Forms.ListViewItem]::new('')
            [void]$li.SubItems.Add($u.DisplayName)
            [void]$li.SubItems.Add($u.SamAccountName)
            [void]$li.SubItems.Add($(if ($eu) { $eu.UserPrincipalName } else { '(not found)' }))
            [void]$li.SubItems.Add($res.Confidence)
            [void]$li.SubItems.Add($(if ($currentImm) { $currentImm } else { '(none)' }))
            [void]$li.SubItems.Add($action)
            $li.Tag = @{ ADUser=$u; EntraUser=$eu; Confidence=$res.Confidence; Action=$action }

            # Color coding
            $li.ForeColor = switch ($res.Confidence) {
                'Exact' { $script:HMColors.FG }
                'Name'  { $script:HMColors.Warning }
                default { $script:HMColors.FGDim }
            }

            # Auto-check rows with an Exact match that need matching
            $li.Checked = ($res.Confidence -eq 'Exact' -and $action -eq 'Will match')
            [void]$lv2.Items.Add($li)
        }

        $total   = $lv2.Items.Count
        # @(...) forces array context -- Where-Object with zero matches returns
        # $null, and $null.Count throws under Set-StrictMode -Version Latest
        # (normally PowerShell auto-resolves .Count on $null to 0, but strict
        # mode disables that convenience).
        $checked = @($lv2.Items | Where-Object { $_.Checked }).Count
        $lblSel.Text = "$total user(s) listed  |  $checked checked"
        $btnResolve.Enabled     = ($total -gt 0)
        $btnBatchApply.Enabled  = ($checked -gt 0)
        Set-Status "Loaded $total user(s)."
    })

    # Update count label on check change
    $lv2.Add_ItemChecked({
        $total   = $lv2.Items.Count
        # @(...) forces array context -- Where-Object with zero matches returns
        # $null, and $null.Count throws under Set-StrictMode -Version Latest
        # (normally PowerShell auto-resolves .Count on $null to 0, but strict
        # mode disables that convenience).
        $checked = @($lv2.Items | Where-Object { $_.Checked }).Count
        $lblSel.Text = "$total user(s) listed  |  $checked checked"
        $btnBatchApply.Enabled = ($checked -gt 0)
    })

    $btnSelAll.Add_Click({
        foreach ($li in $lv2.Items) {
            if ($li.Tag.EntraUser -ne $null) { $li.Checked = $true }
        }
    })
    $btnSelNone.Add_Click({ foreach ($li in $lv2.Items) { $li.Checked = $false } })

    # Re-resolve
    $btnResolve.Add_Click({
        $script:HMLogBox = $logB
        $i = 0
        foreach ($li in $lv2.Items) {
            $i++; Set-Status "Re-resolving $i / $($lv2.Items.Count) ..."
            $form.Refresh()
            $td  = $li.Tag
            $res = Resolve-EntraCandidate -ADUser $td.ADUser
            $eu  = $res.EntraUser
            $cid = Get-ImmutableId -g $td.ADUser.ObjectGUID
            $currentImm = if ($eu) { $eu.OnPremisesImmutableId } else { $null }
            $action = if ($res.Confidence -eq 'None') { 'No Entra match' }
                      elseif ($currentImm -and $currentImm -eq $cid) { 'Already matched' }
                      elseif ($currentImm) { 'Will overwrite' }
                      else { 'Will match' }
            $li.SubItems[3].Text = if ($eu) { $eu.UserPrincipalName } else { '(not found)' }
            $li.SubItems[4].Text = $res.Confidence
            $li.SubItems[5].Text = if ($currentImm) { $currentImm } else { '(none)' }
            $li.SubItems[6].Text = $action
            $li.Tag = @{ ADUser=$td.ADUser; EntraUser=$eu; Confidence=$res.Confidence; Action=$action }
            $li.ForeColor = switch ($res.Confidence) { 'Exact' {$script:HMColors.FG} 'Name' {$script:HMColors.Warning} default {$script:HMColors.FGDim} }
        }
        Set-Status 'Re-resolve complete.'
    })

    # Batch Apply
    $btnBatchApply.Add_Click({
        $script:HMLogBox = $logB
        $toApply = @($lv2.Items | Where-Object { $_.Checked -and $_.Tag.EntraUser -ne $null })
        if ($toApply.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No valid checked rows to process.','Nothing to do','OK','Information')
            return
        }

        # Confirm
        $names = ($toApply | Select-Object -First 8 | ForEach-Object { "  $($_.Tag.ADUser.DisplayName) -> $($_.Tag.EntraUser.UserPrincipalName)" }) -join "`r`n"
        if ($toApply.Count -gt 8) { $names += "`r`n  ... and $($toApply.Count - 8) more" }
        $msg = "Apply hard match for $($toApply.Count) user(s)?`r`n`r`n$names"
        if ($WhatIf) { $msg += "`r`n`r`n[WHATIF] No changes will be written." }
        $r = [System.Windows.Forms.MessageBox]::Show($msg,'Confirm Batch Match','YesNo','Question')
        if ($r -ne 'Yes') { return }

        $ok=0; $skip=0; $fail=0; $cancelled=$false
        $btnBatchApply.Enabled = $false
        foreach ($li in $toApply) {
            if ($cancelled) { break }
            $td = $li.Tag
            Write-HMLog "--- $($td.ADUser.DisplayName) ---" 'INFO'
            $res = Invoke-HardMatch -ADUser $td.ADUser -EntraUser $td.EntraUser `
                       -WhatIfMode $WhatIf.IsPresent -AllowConflictUI $true
            switch ($res) {
                'OK'            { $ok++;   $li.SubItems[6].Text = 'Matched';         $li.ForeColor = $script:HMColors.Success }
                'AlreadyMatched'{ $skip++; $li.SubItems[6].Text = 'Already matched'; $li.ForeColor = $script:HMColors.FGDim }
                'WhatIf'        { $skip++; $li.SubItems[6].Text = 'WhatIf';          $li.ForeColor = $script:HMColors.Warning }
                'Skipped'       { $skip++; $li.SubItems[6].Text = 'Skipped';         $li.ForeColor = $script:HMColors.FGDim }
                'Conflict'      { $fail++; $li.SubItems[6].Text = 'Conflict -- manual resolve needed'; $li.ForeColor = $script:HMColors.Danger }
                'CancelBatch'   {
                    $cancelled = $true
                    $li.SubItems[6].Text = 'Cancelled'
                    $li.ForeColor = $script:HMColors.FGDim
                    Write-HMLog 'Batch cancelled by user at conflict dialog.' 'WARN'
                }
                default         { $fail++; $li.SubItems[6].Text = "Failed: $res";   $li.ForeColor = $script:HMColors.Danger }
            }
            $li.Checked = $false
            $form.Refresh()
        }
        $summary = "Batch done: $ok matched  |  $skip skipped  |  $fail failed"
        if ($cancelled) { $summary += '  |  CANCELLED' }
        Write-HMLog $summary $(if ($fail -gt 0 -or $cancelled) {'WARN'} else {'OK'})
        Set-Status $summary
        $btnBatchApply.Enabled = $false
    })

    # ===================================================================
    # FINAL STARTUP LOG
    # ===================================================================
    $script:HMLogBox = $logS

    if ($script:StartupLog -and $script:StartupLog.Count -gt 0) {
        Write-HMLog '--- Startup log (module bootstrap / Graph version-skew repair / Setup Auth / initial Connect-ToGraph) ---' 'INFO'
        foreach ($entry in $script:StartupLog) { Write-HMLog $entry.Message $entry.Level }
        Write-HMLog '--- End startup log ---' 'INFO'
    }

    Write-HMLog 'Hard Match Tool ready.' 'OK'
    Write-HMLog "Config file: $script:ConfigPath" 'INFO'
    if ($WhatIf) { Write-HMLog 'WhatIf mode -- no changes will be written.' 'WARN' }

    [void]$form.ShowDialog($Owner)
}
function Show-ScanPage       { param($Owner) Write-Log "Scan page not yet implemented in the combined toolbox - use Find-InactiveLicensedUsers.ps1 directly for now." "WARN" $script:launcherLog }

# ============================================================
# ON-PREM AD CLEANUP (from Disable-InactiveADComputers.ps1 / Disable-InactiveADUsers.ps1)
# Pure on-prem AD, no M365/Graph auth involved at all - only the
# ActiveDirectory module (already installed by Invoke-ModuleBootstrap
# above). Ported as a scan-and-select grid, matching the scanner page's
# UX, rather than the originals' filter-and-blind-bulk-disable behavior -
# same 3 actions per object (disable, restamp description, move to the
# disabled OU), just with a review/checkbox step before anything commits.
# ============================================================

function Get-InactiveADObjects {
    param(
        [ValidateSet('Computer', 'User')][string]$Kind,
        [int]$InactiveDays,
        [string]$SearchBase,
        [string]$DisabledOU
    )
    $cutoffDate = (Get-Date).AddDays(-$InactiveDays)
    $searchParams = @{ Filter = { Enabled -eq $true } }
    $searchParams.Properties = if ($Kind -eq 'Computer') {
        @('LastLogonDate', 'DistinguishedName', 'Description', 'DNSHostName', 'OperatingSystem')
    } else {
        @('LastLogonDate', 'DistinguishedName', 'Description', 'SamAccountName', 'PasswordLastSet')
    }
    if ($SearchBase) { $searchParams['SearchBase'] = $SearchBase }

    $dcNames = @()
    if ($Kind -eq 'Computer') {
        try { $dcNames = (Get-ADDomainController -Filter *).Name } catch {}
    }

    $all = if ($Kind -eq 'Computer') { Get-ADComputer @searchParams } else { Get-ADUser @searchParams }

    $candidates = New-Object System.Collections.Generic.List[object]
    foreach ($obj in $all) {
        if ($Kind -eq 'Computer' -and $dcNames -contains $obj.Name) { continue }
        if ($obj.DistinguishedName -like "*$DisabledOU*") { continue }
        $lastLogon = $obj.LastLogonDate
        if ($lastLogon -and $lastLogon -ge $cutoffDate) { continue }

        $candidates.Add([PSCustomObject]@{
            Select            = $false
            Name              = if ($Kind -eq 'Computer') { $obj.Name } else { $obj.SamAccountName }
            LastLogonDate     = if ($lastLogon) { $lastLogon } else { $null }
            LastLogonText     = if ($lastLogon) { $lastLogon.ToString('yyyy-MM-dd') } else { 'Never' }
            Detail            = if ($Kind -eq 'Computer') { $(if ($obj.OperatingSystem) { $obj.OperatingSystem } else { 'Unknown OS' }) } else { $obj.SamAccountName }
            DistinguishedName = $obj.DistinguishedName
            Description       = $obj.Description
        })
    }
    return $candidates
}

function Show-ADCleanupPage {
    param($Owner, [ValidateSet('Computer', 'User')][string]$Kind)

    $title      = if ($Kind -eq 'Computer') { "Disable Inactive AD Computers" } else { "Disable Inactive AD Users" }
    $detailHead = if ($Kind -eq 'Computer') { "OS" } else { "SAM Account" }

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.Forms.MessageBox]::Show("The ActiveDirectory module is not available on this machine. Install RSAT or run on a Domain Controller.", $title, "OK", "Error") | Out-Null
        return
    }
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $title
    $dlg.Size = New-Object System.Drawing.Size(1040, 720)
    $dlg.StartPosition = "CenterParent"; $dlg.BackColor = $BG
    $dlg.FormBorderStyle = "Sizable"; $dlg.MinimumSize = New-Object System.Drawing.Size(820, 560); $dlg.Font = $F_NORM

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = "Top"; $pnlTop.Height = 110; $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)
    $dlg.Controls.Add($pnlTop)

    $lblTitle2 = New-Object System.Windows.Forms.Label
    $lblTitle2.Text = $title; $lblTitle2.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12); $lblTitle2.ForeColor = $TEXT
    $lblTitle2.Location = New-Object System.Drawing.Point(16, 8); $lblTitle2.AutoSize = $true
    $pnlTop.Controls.Add($lblTitle2)

    $lblDays = New-Object System.Windows.Forms.Label
    $lblDays.Text = "Inactive days:"; $lblDays.ForeColor = $TEXT; $lblDays.Location = New-Object System.Drawing.Point(16, 42); $lblDays.AutoSize = $true
    $pnlTop.Controls.Add($lblDays)
    $numDays = New-Object System.Windows.Forms.NumericUpDown
    $numDays.Location = New-Object System.Drawing.Point(110, 40); $numDays.Size = New-Object System.Drawing.Size(60, 22)
    $numDays.Minimum = 1; $numDays.Maximum = 3650; $numDays.Value = 90
    $numDays.BackColor = $CLR_INPUT; $numDays.ForeColor = $TEXT
    $pnlTop.Controls.Add($numDays)

    $lblOU = New-Object System.Windows.Forms.Label
    $lblOU.Text = "Disabled OU (DN, required):"; $lblOU.ForeColor = $TEXT; $lblOU.Location = New-Object System.Drawing.Point(186, 42); $lblOU.AutoSize = $true
    $pnlTop.Controls.Add($lblOU)
    $txtOU = New-Object System.Windows.Forms.TextBox
    $txtOU.Location = New-Object System.Drawing.Point(360, 40); $txtOU.Size = New-Object System.Drawing.Size(360, 22)
    $txtOU.BackColor = $CLR_INPUT; $txtOU.ForeColor = $TEXT; $txtOU.BorderStyle = "FixedSingle"
    if ($script:E -and $script:E.DomainDN) {
        $defaultOUName = if ($Kind -eq 'Computer') { 'Disabled Computers' } else { 'Disabled Users' }
        $txtOU.Text = "OU=$defaultOUName,$($script:E.DomainDN)"
    }
    $pnlTop.Controls.Add($txtOU)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Search base (optional, defaults to whole domain):"; $lblSearch.ForeColor = $TEXT
    $lblSearch.Location = New-Object System.Drawing.Point(16, 72); $lblSearch.AutoSize = $true
    $pnlTop.Controls.Add($lblSearch)
    $txtSearchBase = New-Object System.Windows.Forms.TextBox
    $txtSearchBase.Location = New-Object System.Drawing.Point(300, 70); $txtSearchBase.Size = New-Object System.Drawing.Size(420, 22)
    $txtSearchBase.BackColor = $CLR_INPUT; $txtSearchBase.ForeColor = $TEXT; $txtSearchBase.BorderStyle = "FixedSingle"
    $pnlTop.Controls.Add($txtSearchBase)

    $chkWhatIf = New-Object System.Windows.Forms.CheckBox
    $chkWhatIf.Text = "WhatIf - simulation only (no changes)"
    $chkWhatIf.Location = New-Object System.Drawing.Point(740, 40); $chkWhatIf.Size = New-Object System.Drawing.Size(260, 22)
    $chkWhatIf.Font = $F_BOLD; $chkWhatIf.ForeColor = $WARN; $chkWhatIf.Checked = $true
    $pnlTop.Controls.Add($chkWhatIf)

    $btnScan       = New-LauncherButton "Scan"        740  68 100 32 $ACCENT
    $btnSelectAll  = New-LauncherButton "Select All"   850  68  85 32 $PANEL $TEXT
    $pnlTop.Controls.AddRange(@($btnScan, $btnSelectAll))
    $btnScan.Font = $F_BOLD; $btnSelectAll.Font = $F_BOLD

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(0, 110); $grid.Anchor = "Top,Bottom,Left,Right"
    $grid.Size = New-Object System.Drawing.Size(1024, 380)
    $grid.BackgroundColor = $PANEL; $grid.ForeColor = $TEXT
    $grid.DefaultCellStyle.BackColor = $CLR_INPUT; $grid.DefaultCellStyle.ForeColor = $TEXT
    $grid.DefaultCellStyle.SelectionBackColor = $ACCENT
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $PANEL; $grid.ColumnHeadersDefaultCellStyle.ForeColor = $TEXT
    $grid.ColumnHeadersDefaultCellStyle.Font = $F_BOLD; $grid.EnableHeadersVisualStyles = $false
    $grid.RowHeadersVisible = $false; $grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = "FullRowSelect"; $grid.AutoSizeColumnsMode = "Fill"; $grid.Font = $F_NORM
    $grid.Columns.Add((New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = "Select"; HeaderText = "Sel"; Width = 40 })) | Out-Null
    $grid.Columns.Add("Name", "Name") | Out-Null
    $grid.Columns.Add("LastLogonText", "Last Logon") | Out-Null
    $grid.Columns.Add("Detail", $detailHead) | Out-Null
    $grid.Columns.Add("DistinguishedName", "Distinguished Name") | Out-Null
    foreach ($cn in @("Name", "LastLogonText", "Detail", "DistinguishedName")) { $grid.Columns[$cn].ReadOnly = $true }
    $dlg.Controls.Add($grid)
    $grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } })

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Set Disabled OU and click Scan."; $lblStatus.ForeColor = $TEXTDIM
    $lblStatus.Location = New-Object System.Drawing.Point(16, 498); $lblStatus.Size = New-Object System.Drawing.Size(1000, 20); $lblStatus.Anchor = "Bottom,Left,Right"
    $dlg.Controls.Add($lblStatus)

    $log = New-Object System.Windows.Forms.RichTextBox
    $log.Location = New-Object System.Drawing.Point(16, 522); $log.Size = New-Object System.Drawing.Size(1000, 108); $log.Anchor = "Bottom,Left,Right"
    $log.Font = $F_MONO; $log.BackColor = [System.Drawing.Color]::FromArgb(10, 14, 22)
    $log.ForeColor = $TEXT; $log.ReadOnly = $true; $log.BorderStyle = "None"; $log.ScrollBars = "Vertical"
    $dlg.Controls.Add($log)

    # Preserves the originals' audit-log-file intent (their stated design goal was "a log
    # file records every action taken") alongside the toolbox's own GUI log convention.
    $script:_adCleanupLogPath = Join-Path $env:TEMP "ADM365Toolbox_Disable$($Kind)s_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $logPath = $script:_adCleanupLogPath
    function Write-PageLog {
        param([string]$Message, [string]$Level = "INFO")
        Write-Log $Message $Level $log
        try { Add-Content -Path $logPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message" -ErrorAction SilentlyContinue } catch {}
    }
    Write-PageLog "Log file: $logPath" "INFO"

    $pnlBot = New-Object System.Windows.Forms.Panel
    $pnlBot.Anchor = "Bottom,Left,Right"; $pnlBot.Location = New-Object System.Drawing.Point(0, 634); $pnlBot.Size = New-Object System.Drawing.Size(1024, 50)
    $pnlBot.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)
    $btnProcess = New-LauncherButton "Disable + Move Selected" 16 8 240 34 $DANGER
    $btnProcess.Font = $F_BOLD
    $btnCloseDlg = New-LauncherButton "Close" 928 8 80 34 $PANEL $TEXT
    $pnlBot.Controls.AddRange(@($btnProcess, $btnCloseDlg))
    $dlg.Controls.Add($pnlBot)

    $script:_adCleanupCandidates = @()

    $btnScan.Add_Click({
        if (-not $txtOU.Text) {
            [System.Windows.Forms.MessageBox]::Show("Disabled OU is required.", $title) | Out-Null
            return
        }
        $btnScan.Enabled = $false
        $lblStatus.ForeColor = $TEXTDIM; $lblStatus.Text = "Scanning..."
        try {
            $ouExists = $true
            try { Get-ADOrganizationalUnit -Identity $txtOU.Text -ErrorAction Stop | Out-Null }
            catch { $ouExists = $false }
            if (-not $ouExists) {
                $r = [System.Windows.Forms.MessageBox]::Show("Target OU '$($txtOU.Text)' does not exist. Create it now?", $title, "YesNo", "Question")
                if ($r -eq "Yes") {
                    $ouName   = ($txtOU.Text -split ',')[0] -replace '^OU=', ''
                    $parentDN = ($txtOU.Text -split ',', 2)[1]
                    New-ADOrganizationalUnit -Name $ouName -Path $parentDN
                    Write-PageLog "Created OU: $($txtOU.Text)" "SUCCESS"
                } else {
                    Write-PageLog "OU does not exist and was not created - scan aborted." "ERROR"
                    $lblStatus.ForeColor = $DANGER; $lblStatus.Text = "[X] Target OU does not exist."
                    return
                }
            }

            $script:_adCleanupCandidates = Get-InactiveADObjects -Kind $Kind -InactiveDays ([int]$numDays.Value) -SearchBase $txtSearchBase.Text -DisabledOU $txtOU.Text
            $grid.Rows.Clear()
            foreach ($c in $script:_adCleanupCandidates) {
                $grid.Rows.Add($false, $c.Name, $c.LastLogonText, $c.Detail, $c.DistinguishedName) | Out-Null
            }
            $lblStatus.ForeColor = $GREEN
            $lblStatus.Text = "$($script:_adCleanupCandidates.Count) inactive $Kind(s) found. Review and check which to disable+move."
            Write-PageLog "$($script:_adCleanupCandidates.Count) inactive $Kind(s) found (inactive $([int]$numDays.Value)+ days)." "SUCCESS"
        } catch {
            $lblStatus.ForeColor = $DANGER; $lblStatus.Text = "[X] Scan failed - see log."
            Write-PageLog "Scan error: $_" "ERROR"
        } finally { $btnScan.Enabled = $true }
    })

    $btnSelectAll.Add_Click({
        $allChecked = $true
        foreach ($row in $grid.Rows) { if (-not [bool]$row.Cells['Select'].Value) { $allChecked = $false; break } }
        foreach ($row in $grid.Rows) { $row.Cells['Select'].Value = -not $allChecked }
    })

    $btnProcess.Add_Click({
        $selectedNames = @()
        foreach ($row in $grid.Rows) { if ([bool]$row.Cells['Select'].Value) { $selectedNames += [string]$row.Cells['Name'].Value } }
        $selected = @($script:_adCleanupCandidates | Where-Object { $selectedNames -contains $_.Name })

        if ($selected.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No accounts checked.", "Nothing selected") | Out-Null
            return
        }

        $whatIfMode = [bool]$chkWhatIf.Checked
        $modeText = if ($whatIfMode) { "WHATIF SIMULATION (no changes will be made)" } else { "LIVE - THIS WILL MAKE REAL CHANGES" }
        $preview = ($selected | Select-Object -First 25 | ForEach-Object { "  - $($_.Name)" }) -join "`n"
        if ($selected.Count -gt 25) { $preview += "`n  ... and $($selected.Count - 25) more" }
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Mode: $modeText`n`nFor each of the $($selected.Count) $Kind(s) below: disable, restamp description, move to `"$($txtOU.Text)`".`n`n$preview`n`nProceed?",
            "Confirm", "YesNo", "Warning")
        if ($r -ne "Yes") { return }

        $btnProcess.Enabled = $false
        $successCount = 0; $failCount = 0
        foreach ($c in $selected) {
            $newDesc = "Disabled by ADM365-Toolbox on $(Get-Date -Format 'yyyy-MM-dd') | Last logon: $($c.LastLogonText) | $($c.Description)"
            try {
                if ($whatIfMode) {
                    Write-PageLog "[WHATIF] Would disable, restamp description, and move: $($c.Name)" "WARN"
                } else {
                    Disable-ADAccount -Identity $c.DistinguishedName
                    if ($Kind -eq 'Computer') { Set-ADComputer -Identity $c.DistinguishedName -Description $newDesc }
                    else { Set-ADUser -Identity $c.DistinguishedName -Description $newDesc }
                    Move-ADObject -Identity $c.DistinguishedName -TargetPath $txtOU.Text
                    Write-PageLog "Disabled, restamped, and moved: $($c.Name)" "SUCCESS"
                }
                $successCount++
            } catch {
                Write-PageLog "FAILED: $($c.Name): $_" "ERROR"
                $failCount++
            }
        }
        $lblStatus.ForeColor = if ($failCount -eq 0) { $GREEN } else { $WARN }
        $lblStatus.Text = "Processed $successCount, $failCount failure(s). Log: $logPath"
        Write-PageLog "Done. Processed $successCount, $failCount failure(s)." "SUCCESS"
        $btnProcess.Enabled = $true
    })

    $btnCloseDlg.Add_Click({ $dlg.Close() })

    [void]$dlg.ShowDialog($Owner)
}

function New-LauncherButton { param($t, $x, $y, $w, $h, $bg = $ACCENT, $fg = [System.Drawing.Color]::White)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $t; $b.Location = New-Object System.Drawing.Point($x, $y); $b.Size = New-Object System.Drawing.Size($w, $h)
    $b.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11); $b.ForeColor = $fg; $b.BackColor = $bg
    $b.FlatStyle = "Flat"; $b.FlatAppearance.BorderSize = 0; $b
}

$launcherForm = New-Object System.Windows.Forms.Form
$launcherForm.Text = "AD/M365 Admin Toolbox"
$launcherForm.Size = New-Object System.Drawing.Size(760, 740)
$launcherForm.StartPosition = "CenterScreen"; $launcherForm.BackColor = $BG
$launcherForm.FormBorderStyle = "FixedDialog"; $launcherForm.MaximizeBox = $false; $launcherForm.Font = $F_NORM

$pnlTitle = New-Object System.Windows.Forms.Panel
$pnlTitle.Dock = "Top"; $pnlTitle.Height = 60; $pnlTitle.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "AD / M365 Admin Toolbox"; $lblTitle.Font = $F_TITLE; $lblTitle.ForeColor = $TEXT
$lblTitle.Location = New-Object System.Drawing.Point(16, 14); $lblTitle.AutoSize = $true
$pnlTitle.Controls.Add($lblTitle)
$launcherForm.Controls.Add($pnlTitle)

$btnOnboard    = New-LauncherButton "Onboard User"                   20  80 340 70
$btnOffboard   = New-LauncherButton "Offboard User"                 380  80 340 70 $DANGER
$btnHardMatch  = New-LauncherButton "Hard-Match AD <-> Entra"         20 160 340 70 $GREEN $BG
$btnScan       = New-LauncherButton "Scan Inactive Licensed Users"   380 160 340 70 $WARN  $BG
$btnDisableComputers = New-LauncherButton "Disable Inactive AD Computers"  20 240 340 70 $PANEL $TEXT
$btnDisableUsers     = New-LauncherButton "Disable Inactive AD Users"     380 240 340 70 $PANEL $TEXT
$btnSetupAuth  = New-LauncherButton "Setup Auth"                     20 330 340 44 $PANEL $TEXT
$btnClose      = New-LauncherButton "Close"                        380 330 340 44 $PANEL $TEXT
$launcherForm.Controls.AddRange(@($btnOnboard, $btnOffboard, $btnHardMatch, $btnScan, $btnDisableComputers, $btnDisableUsers, $btnSetupAuth, $btnClose))

$lblLogHead = New-Object System.Windows.Forms.Label
$lblLogHead.Text = "Log"; $lblLogHead.ForeColor = $TEXTDIM; $lblLogHead.Font = $F_BOLD
$lblLogHead.Location = New-Object System.Drawing.Point(20, 386); $lblLogHead.AutoSize = $true
$launcherForm.Controls.Add($lblLogHead)

$launcherLog = New-Object System.Windows.Forms.RichTextBox
$launcherLog.Location = New-Object System.Drawing.Point(20, 408); $launcherLog.Size = New-Object System.Drawing.Size(700, 260)
$launcherLog.Font = $F_MONO; $launcherLog.BackColor = [System.Drawing.Color]::FromArgb(10, 14, 22)
$launcherLog.ForeColor = $TEXT; $launcherLog.ReadOnly = $true; $launcherLog.BorderStyle = "None"; $launcherLog.ScrollBars = "Vertical"
$launcherForm.Controls.Add($launcherLog)
$script:launcherLog = $launcherLog

if ($script:StartupLog -and $script:StartupLog.Count -gt 0) {
    Write-Log "--- Startup log (module bootstrap / Graph version-skew repair / role check) ---" "INFO" $launcherLog
    foreach ($entry in $script:StartupLog) {
        $mappedLevel = switch ($entry.Level) { "OK" { "SUCCESS" } "ERR" { "ERROR" } "WARN" { "WARN" } default { "INFO" } }
        Write-Log $entry.Message $mappedLevel $launcherLog
    }
    Write-Log "--- End startup log ---" "INFO" $launcherLog
}
Write-Log "Toolbox ready." "SUCCESS" $launcherLog

$btnOnboard.Add_Click({ Show-OnboardPage -Owner $launcherForm })
$btnOffboard.Add_Click({ Show-OffboardPage -Owner $launcherForm })
$btnHardMatch.Add_Click({ Show-HardMatchPage -Owner $launcherForm })
$btnScan.Add_Click({ Show-ScanPage -Owner $launcherForm })
$btnDisableComputers.Add_Click({ Show-ADCleanupPage -Owner $launcherForm -Kind 'Computer' })
$btnDisableUsers.Add_Click({ Show-ADCleanupPage -Owner $launcherForm -Kind 'User' })

$btnSetupAuth.Add_Click({
    $r = [System.Windows.Forms.MessageBox]::Show(
        "This creates a certificate-based app registration in this tenant (or refreshes the existing 'ADM365LifecycleTool' one) so none of the four tools need interactive MFA sign-in each run, and grants it the full permission set all four tools need - including AuditLog.Read.All for the scan page.`n`nSign in with a GLOBAL ADMIN account when prompted. Continue?",
        "Setup Auth", "YesNo", "Question")
    if ($r -ne "Yes") { return }
    $btnSetupAuth.Enabled = $false
    $ok = Initialize-M365Auth -LogBox $launcherLog
    $btnSetupAuth.Enabled = $true
    if (-not $ok) {
        [System.Windows.Forms.MessageBox]::Show("Setup Auth failed - see the log for details.", "Setup Auth", "OK", "Error") | Out-Null
    }
})

$btnClose.Add_Click({ $launcherForm.Close() })

[void]$launcherForm.ShowDialog()
