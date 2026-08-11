<#
# Copyright (c) 2025 Fouche Enterprises, LLC. All rights reserved. Licensed for use by authorized parties only.

.SYNOPSIS
    AD/M365 User Offboarding Tool (Self-Contained)

.DESCRIPTION
    Single script. On every launch it installs/updates required modules,
    verifies M365 admin roles, auto-discovers the environment from the DC,
    then runs the offboarding steps below.

    Steps:
      1. Rename AD user  (DisplayName -> "Historical - Name",
                          SAM / UPN / Email -> "Historical-<original>")
      2. Entra delta sync #1  (push rename)
      3. Disable account and move to Disabled Users OU
      4. Convert mailbox to Shared
      5. Full Access + Send-As to delegate(s)
      6. Forwarding distribution group for old email address
      7. OneDrive Site Collection Admin to delegate(s)
      8. Entra delta sync #2  (push disable / decouple)
      9. Remove all M365 licenses

    GUI mode  : Run with no parameters
    CLI mode  : Supply -SamAccountName  (NinjaRMM / automation)

.PARAMETER SamAccountName
    SAM of the user to offboard. Omit to open GUI.

.PARAMETER DisabledUsersOU
    Override auto-discovered Disabled Users OU (full DN).

.PARAMETER DelegateUsers
    UPN(s) to receive mailbox Full Access, Send-As, and OneDrive SCA.

.PARAMETER DistributionGroupMembers
    UPN(s) to add to the forwarding distribution group.

.PARAMETER AADConnectServer
    Override auto-discovered AADConnect server hostname.

.PARAMETER SharePointAdminURL
    Override auto-discovered SharePoint Admin URL.

.PARAMETER AdminUPN
    UPN of the M365 admin for role verification.

.PARAMETER SkipRoleCheck
    Skip M365 role verification.

.PARAMETER SkipModuleCheck
    Skip module install/update check.

.PARAMETER WhatIf
    Simulate all steps - no changes committed.

.EXAMPLE
    # GUI
    .\Offboard-ADUser.ps1

.EXAMPLE
    # CLI / NinjaRMM
    .\Offboard-ADUser.ps1 -SamAccountName "jsmith" -DelegateUsers "mgr@contoso.com"
#>

[CmdletBinding()]
param(
    [string]   $SamAccountName           = "",
    [string]   $DisabledUsersOU          = "",
    [string[]] $DelegateUsers            = @(),
    [string[]] $DistributionGroupMembers = @(),
    [string]   $AADConnectServer         = "",
    [string]   $SharePointAdminURL       = "",
    [string]   $AdminUPN                 = "",
    [switch]   $SkipRoleCheck,
    [switch]   $SkipModuleCheck,
    [switch]   $WhatIf
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
            "Administrator privileges are required.`n`n" +
            "Please right-click and select 'Run as Administrator'.",
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
        $null = Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $null = Import-Module Microsoft.Graph.Users -ErrorAction Stop
        $null = Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

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

$isCLI = ($SamAccountName -ne "")
$isSilent = $isCLI

Write-Host "[$(Get-Date -Format 'HH:mm:ss')][INFO] AD/M365 Offboarding Tool - v1.0 (c) Roger Fouche / Fouche Enterprises, LLC" -ForegroundColor Magenta

if (-not $SkipModuleCheck) { Invoke-ModuleBootstrap -Silent:$isSilent }
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

if ($DisabledUsersOU)    { $script:E.DisabledUsersOU    = $DisabledUsersOU }
if ($AADConnectServer)   { $script:E.AADConnectServer   = $AADConnectServer }
if ($SharePointAdminURL) { $script:E.SharePointAdminURL = $SharePointAdminURL }

Write-Host "[$(Get-Date -Format 'HH:mm:ss')][OK] DC=$($script:E.DC)  Domain=$($script:E.LocalDomain)  Tenant=$($script:E.TenantName)" -ForegroundColor Green

# ============================================================
# CONFIG FILE - persists admin UPN across runs
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
    param([string[]]$FallbackScopes = @(
        "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All",
        "Organization.Read.All", "RoleManagement.ReadWrite.Directory", "Sites.FullControl.All"
    ))
    $null = Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $cfg = Get-CertConfig
    if ($cfg) {
        # Delegate to Connect-ToGraph which resolves the cert from LocalMachine\My first,
        # ensuring private-key accessibility under elevation, PS7 relaunch, and SYSTEM.
        $ok = Connect-ToGraph
        if (-not $ok) { throw "Certificate-based Graph connection failed - see log." }
    } else {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')][WARN] Graph: interactive auth - run Setup Auth to enable non-interactive" -ForegroundColor Yellow
        $null = Connect-MgGraph -Scopes $FallbackScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        if ($ctx -and $ctx.TenantId -and -not $script:Config.TenantId) {
            $script:Config.TenantId = $ctx.TenantId; Save-Config
        }
    }
}

function Connect-M365Exchange {
    # NOTE: Connect-M365Exchange is kept for Onboard welcome email (SMTP path, no EXO cmdlets)
    # For Offboard, all EXO cmdlets run via Invoke-EXOProcess in a child process to avoid
    # the Microsoft.Graph + ExchangeOnlineManagement Azure.Core/MSAL assembly conflict.
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
    param([System.Windows.Forms.RichTextBox]$LogBox = $null)

    Write-Log "Launching auth setup in an isolated PowerShell window (avoids a known Microsoft.Graph module assembly conflict with this session)..." "INFO" $LogBox
    Write-Log "A new console window will open. Sign in with a GLOBAL ADMIN account when prompted, then return here." "WARN" $LogBox

    $uid        = [System.Guid]::NewGuid().ToString('N').Substring(0, 10)
    $scriptFile = Join-Path $env:TEMP "setupauth_s_$uid.ps1"
    $resultFile = Join-Path $env:TEMP "setupauth_r_$uid.json"
    $fallbackExoOrg = "$($script:E.TenantName).onmicrosoft.com"

    $authCommands = @'
$ErrorActionPreference = "Stop"
$result = [ordered]@{ Success = $false; AppId = ""; TenantId = ""; CertThumbprint = ""; ExchangeOrg = ""; ErrorMessage = "" }
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

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

    $graphPermNames = @("User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All","RoleManagement.ReadWrite.Directory")
    $graphRoles = @()
    foreach ($p in $graphPermNames) {
        $roleObj = $graphSP.AppRoles | Where-Object { $_.Value -eq $p } | Select-Object -First 1
        if ($roleObj) { $graphRoles += @{ Id = [string]$roleObj.Id; Type = "Role" } }
    }
    $reqAccess = @(@{ ResourceAppId = "00000003-0000-0000-c000-000000000000"; ResourceAccess = $graphRoles })

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

    foreach ($role in $graphRoles) {
        try { New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $graphSP.Id -AppRoleId $role.Id | Out-Null } catch {}
    }
    if ($exoSP -and $exoRoleId) {
        try { New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $exoSP.Id -AppRoleId $exoRoleId | Out-Null } catch {}
    }
    if ($spoSP -and $spoRoleId) {
        try { New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $spoSP.Id -AppRoleId $spoRoleId | Out-Null } catch {}
    }
    Write-Host "Admin consent granted." -ForegroundColor Green

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
    $result.ErrorMessage = "$_"
    Write-Host "[ERROR] Setup failed: $_" -ForegroundColor Red
    if ((Test-Path variable:cert) -and $cert) {
        try { Remove-Item "Cert:\LocalMachine\My\$($cert.Thumbprint)" -DeleteKey -ErrorAction SilentlyContinue } catch {}
    }
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}
'@

    $scriptBody = @"
`$ErrorActionPreference = 'Stop'
`$FallbackExoOrg = '$fallbackExoOrg'

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
    Write-Log "Wait 5-10 minutes before first use." "WARN" $LogBox
    return $true
}

# ============================================================

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
# SECTION 2 - CORE OFFBOARDING ENGINE
# ============================================================

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

# ============================================================
# CLI MODE
# ============================================================

if ($SamAccountName -ne "") {
    Invoke-Offboarding `
        -SamAccountName           $SamAccountName `
        -DisabledUsersOU          $script:E.DisabledUsersOU `
        -DelegateUsers            $DelegateUsers `
        -DistributionGroupMembers $DistributionGroupMembers `
        -AADConnectServer         $script:E.AADConnectServer `
        -SharePointAdminURL       $script:E.SharePointAdminURL `
        -WhatIfMode               $WhatIf.IsPresent
    return
}

# ============================================================
# GUI MODE
# ============================================================

# Top-level error catch - any failure before the form appears shows a
# MessageBox instead of silently closing the window.
try {

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Palette
$BG      = [System.Drawing.Color]::FromArgb(20,  24,  33)
$PANEL   = [System.Drawing.Color]::FromArgb(30,  36,  50)
$CLR_INPUT   = [System.Drawing.Color]::FromArgb(18,  22,  32)
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
$F_TITLE = New-Object System.Drawing.Font("Segoe UI Light", 13)

# Widget helpers
function Lbl { param($t, $x, $y, $w = 200, $h = 22, $f = $F_NORM, $c = $TEXT)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $t; $l.Location = [System.Drawing.Point]::new($x, $y)
    $l.Size = [System.Drawing.Size]::new($w, $h)
    $l.Font = $f; $l.ForeColor = $c
    $l.BackColor = [System.Drawing.Color]::Transparent; $l
}
function TBox { param($x, $y, $w = 260, $h = 26, $multi = $false)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = [System.Drawing.Point]::new($x, $y); $t.Size = [System.Drawing.Size]::new($w, $h)
    $t.Font = $F_NORM; $t.ForeColor = $TEXT; $t.BackColor = $CLR_INPUT; $t.BorderStyle = "FixedSingle"
    if ($multi) { $t.Multiline = $true; $t.ScrollBars = "Vertical" }; $t
}
function Btn { param($t, $x, $y, $w = 120, $h = 30, $bg = $ACCENT, $fg = $BG)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $t; $b.Location = [System.Drawing.Point]::new($x, $y); $b.Size = [System.Drawing.Size]::new($w, $h)
    $b.Font = $F_BOLD; $b.ForeColor = $fg; $b.BackColor = $bg
    $b.FlatStyle = "Flat"; $b.FlatAppearance.BorderSize = 0; $b.Cursor = "Hand"; $b
}
function GBox { param($t, $x, $y, $w, $h)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $t; $g.Location = [System.Drawing.Point]::new($x, $y); $g.Size = [System.Drawing.Size]::new($w, $h)
    $g.Font = $F_BOLD; $g.ForeColor = $ACCENT; $g.BackColor = $PANEL; $g
}

# ============================================================
# USER SEARCH DIALOG
# ============================================================

function Show-UserSearch {
    param([System.Windows.Forms.Form]$Parent = $null)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Search Active Directory Users"
    $dlg.Size = [System.Drawing.Size]::new(640, 500)
    $dlg.StartPosition = "CenterParent"; $dlg.BackColor = $BG
    $dlg.FormBorderStyle = "FixedDialog"; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    $dlg.Controls.Add((Lbl "Search by name, SAM account, or email:" 10 12 400 22 $F_BOLD $TEXT))
    $txtQ = TBox 10 36 500 26; $dlg.Controls.Add($txtQ)
    $btnGo = Btn "Search" 520 34 96 28 $ACCENT $BG; $dlg.Controls.Add($btnGo)
    $dlg.Controls.Add((Lbl "Results - double-click or select and click OK:" 10 70 500 20 $F_NORM $TEXTDIM))

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = [System.Drawing.Point]::new(10, 92); $lv.Size = [System.Drawing.Size]::new(606, 308)
    $lv.View = "Details"; $lv.FullRowSelect = $true; $lv.GridLines = $true
    $lv.Font = $F_NORM; $lv.BackColor = $CLR_INPUT; $lv.ForeColor = $TEXT; $lv.BorderStyle = "FixedSingle"
    [void]$lv.Columns.Add("Display Name", 200); [void]$lv.Columns.Add("SAM", 120)
    [void]$lv.Columns.Add("Email", 200); [void]$lv.Columns.Add("Status", 76)
    $dlg.Controls.Add($lv)

    $lblC  = Lbl "" 10 408 400 20 $F_NORM $TEXTDIM; $dlg.Controls.Add($lblC)
    $btnOK = Btn "Select" 400 406 130 30 $GREEN $BG; $btnOK.Enabled = $false; $dlg.Controls.Add($btnOK)
    $btnX  = Btn "Cancel" 540 406 76 30 $BORDER $TEXT; $dlg.Controls.Add($btnX)

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

# ============================================================
# MAIN FORM
# ============================================================

# Initialize selection variables so strict mode never throws
# if the user runs without opening the search dialog first
$script:SearchResult = $null

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD/M365 User Offboarding Tool"
$form.Size = [System.Drawing.Size]::new(820, 660)
$form.StartPosition = "CenterScreen"; $form.BackColor = $BG
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
$pnlTitle.Controls.Add((Lbl "  [OFFBOARD]  AD/M365 User Offboarding Tool" 0 8 560 34 $F_TITLE $ACCENT))
$pnlTitle.Controls.Add((Lbl "Domain: $($script:E.LocalDomain)   Tenant: $($script:E.TenantName)" 556 18 254 20 $F_NORM $TEXTDIM))
$pnlScroll.Controls.Add($pnlTitle)

# Status banner
$pnlBanner = New-Object System.Windows.Forms.Panel
$pnlBanner.Location = [System.Drawing.Point]::new(0, 52); $pnlBanner.Size = [System.Drawing.Size]::new(820, 24)
$pnlBanner.BackColor = [System.Drawing.Color]::FromArgb(0, 55, 18)
$lblBanner = Lbl "  [OK] Modules verified   [OK] M365 roles verified" 0 3 820 18 $F_NORM $GREEN
$pnlBanner.Controls.Add($lblBanner)
$pnlScroll.Controls.Add($pnlBanner)

# User section
$grpUser = GBox "User to Offboard" 10 84 794 112
$pnlScroll.Controls.Add($grpUser)

$grpUser.Controls.Add((Lbl "SAM Account:" 10 26 110 22 $F_BOLD $TEXT))
$txtSAM = TBox 126 22 216 26; $grpUser.Controls.Add($txtSAM)
$btnSearchAD = Btn "Search AD" 352 20 116 28 $PANEL $ACCENT; $grpUser.Controls.Add($btnSearchAD)
$btnLookup   = Btn "Lookup"    477 20 90  28 $PANEL $GREEN;  $grpUser.Controls.Add($btnLookup)

$lblUserInfo = Lbl "Enter a SAM account name above, or click Search AD to find a user." 10 56 760 22 $F_NORM $TEXTDIM
$grpUser.Controls.Add($lblUserInfo)
$lblPreview = Lbl "" 10 78 760 22 $F_NORM $ACCENT
$grpUser.Controls.Add($lblPreview)

# AD and Sync section
$grpAD = GBox "AD and Sync  (auto-discovered - edit if needed)" 10 204 794 102
$pnlScroll.Controls.Add($grpAD)

$grpAD.Controls.Add((Lbl "Disabled Users OU:" 10 26 140 22 $F_BOLD $TEXT))
$txtOU = TBox 155 22 615 26; $txtOU.Text = $script:E.DisabledUsersOU; $grpAD.Controls.Add($txtOU)

$grpAD.Controls.Add((Lbl "AADConnect Server:" 10 58 140 22 $F_BOLD $TEXT))
$txtAADC = TBox 155 54 200 26; $txtAADC.Text = $script:E.AADConnectServer; $grpAD.Controls.Add($txtAADC)

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Text = "WhatIf - simulation only (no changes)"
$chkWhatIf.Location = [System.Drawing.Point]::new(368, 56); $chkWhatIf.Size = [System.Drawing.Size]::new(312, 24)
$chkWhatIf.Font = $F_BOLD; $chkWhatIf.ForeColor = $WARN
$chkWhatIf.BackColor = [System.Drawing.Color]::Transparent
$grpAD.Controls.Add($chkWhatIf)

# M365 section
$grpM365 = GBox "M365  (auto-discovered - edit if needed)" 10 314 794 60
$pnlScroll.Controls.Add($grpM365)
$grpM365.Controls.Add((Lbl "SPO Admin URL:" 10 26 120 22 $F_BOLD $TEXT))
$txtSPO = TBox 136 22 640 26; $txtSPO.Text = $script:E.SharePointAdminURL; $grpM365.Controls.Add($txtSPO)

# Delegation section
$grpDel = GBox "Delegates - Full Access, Send-As, OneDrive SCA  (one UPN per line)" 10 382 385 124
$pnlScroll.Controls.Add($grpDel)
$txtDelegates = TBox 10 20 360 92 $true; $grpDel.Controls.Add($txtDelegates)

$grpDist = GBox "Forwarding Dist Group Members  (one UPN per line)" 405 382 399 124
$pnlScroll.Controls.Add($grpDist)
$grpDist.Controls.Add((Lbl "Group fwd-<sam> will receive old email as primary SMTP" 10 4 370 16 $F_NORM $TEXTDIM))
$txtDist = TBox 10 20 374 92 $true; $grpDist.Controls.Add($txtDist)

# Progress and log
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = [System.Drawing.Point]::new(10, 516); $progress.Size = [System.Drawing.Size]::new(794, 16)
$progress.Minimum = 0; $progress.Maximum = 100; $progress.ForeColor = $GREEN; $progress.BackColor = $PANEL
$pnlScroll.Controls.Add($progress)

$lblStatus = Lbl "Ready - complete the fields above and click Run Offboarding." 10 536 794 20 $F_NORM $TEXTDIM
$pnlScroll.Controls.Add($lblStatus)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = [System.Drawing.Point]::new(10, 560); $logBox.Size = [System.Drawing.Size]::new(794, 100)
$logBox.Font = $F_MONO; $logBox.BackColor = [System.Drawing.Color]::FromArgb(10, 14, 22)
$logBox.ForeColor = $TEXT; $logBox.ReadOnly = $true; $logBox.BorderStyle = "None"; $logBox.ScrollBars = "Vertical"
$pnlScroll.Controls.Add($logBox)

# Bottom bar
$pnlBot = New-Object System.Windows.Forms.Panel
$pnlBot.Dock = "Bottom"; $pnlBot.Height = 52
$pnlBot.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)
$form.Controls.Add($pnlBot)
$form.Controls.Add($pnlScroll)

$btnRun     = Btn "Run Offboarding"   10  10 185 34 $GREEN  $BG
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$btnExport   = Btn "Export Log"       206  12 130 30 $BORDER $TEXT
$btnRecheck  = Btn "Re-check Setup"   346  12 160 30 $PANEL  $ACCENT
$btnSetupAuth= Btn "Setup Auth"       516  12 140 30 $PANEL  $WARN
$btnClose    = Btn "Close"            700  12 100 30 $DANGER ([System.Drawing.Color]::White)
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

[System.Windows.Forms.Application]::Run($form)

} catch {
    # Catch any error that occurs before or during form launch
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "A startup error occurred:`n`n$_`n`nLine: $($_.InvocationInfo.ScriptLineNumber)",
        "Startup Error", "OK", "Error") | Out-Null
}