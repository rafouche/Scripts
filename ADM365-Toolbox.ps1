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

$F_NORM  = New-Object System.Drawing.Font("Segoe UI", 9)
$F_BOLD  = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$F_MONO  = New-Object System.Drawing.Font("Consolas", 8.5)
$F_TITLE = New-Object System.Drawing.Font("Segoe UI Semibold", 14)

# ---- Stub pages (replaced one at a time as each tool is ported in) ----
function Show-OnboardPage    { param($Owner) Write-Log "Onboard page not yet implemented in the combined toolbox - use Onboard-ADUser.ps1 directly for now." "WARN" $script:launcherLog }
function Show-OffboardPage   { param($Owner) Write-Log "Offboard page not yet implemented in the combined toolbox - use Offboard-ADUser.ps1 directly for now." "WARN" $script:launcherLog }
function Show-HardMatchPage  { param($Owner) Write-Log "Hard-Match page not yet implemented in the combined toolbox - use Set-ADEntraHardMatch.ps1 directly for now." "WARN" $script:launcherLog }
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
