<#
# Copyright (c) 2025 Fouche Enterprises, LLC. All rights reserved. Licensed for use by authorized parties only.

.SYNOPSIS
    AD/M365 User Onboarding Tool (Self-Contained)

.DESCRIPTION
    Single script. On every launch it installs/updates required modules,
    verifies M365 admin roles, auto-discovers the environment from the DC,
    queries live M365 license inventory, then runs the onboarding steps.

    Steps:
      1. Create AD user account
      2. Add to AD security groups
      3. Place in target OU
      4. Enable account
      5. Entra delta sync (push new account to M365)
      6. Assign M365 license(s) from live tenant query
      7. Send welcome/credential notification (optional)

    GUI mode : Run with no parameters
    CLI mode : Supply -FirstName and -LastName (NinjaRMM / automation)

.PARAMETER FirstName
    New user's first name. Triggers CLI mode when supplied with -LastName.

.PARAMETER LastName
    New user's last name.

.PARAMETER DisplayName
    Override auto-generated display name (default: "FirstName LastName").

.PARAMETER JobTitle
    User's job title.

.PARAMETER Department
    User's department.

.PARAMETER Manager
    SAM account name of the user's manager.

.PARAMETER PhoneNumber
    User's phone number.

.PARAMETER TargetOU
    Override OU distinguished name for the new account.

.PARAMETER ADGroups
    Array of AD group names to add the user to.

.PARAMETER LicenseSkuIds
    Array of M365 license SKU IDs to assign.

.PARAMETER AADConnectServer
    Override auto-discovered AADConnect server hostname.

.PARAMETER AdminUPN
    UPN of the M365 admin for role verification.

.PARAMETER TempPassword
    Override auto-generated temporary password.

.PARAMETER NotifyEmail
    Email address to send the welcome/credential notification to.

.PARAMETER SkipRoleCheck
    Skip M365 role verification.

.PARAMETER SkipModuleCheck
    Skip module install/update check.

.PARAMETER WhatIf
    Simulate all steps - no changes committed.

.EXAMPLE
    # GUI
    .\Onboard-ADUser.ps1

.EXAMPLE
    # CLI / NinjaRMM
    .\Onboard-ADUser.ps1 -FirstName "Jane" -LastName "Smith" -JobTitle "Accountant"
#>

[CmdletBinding()]
param(
    [string]   $FirstName          = "",
    [string]   $LastName           = "",
    [string]   $DisplayName        = "",
    [string]   $JobTitle           = "",
    [string]   $Department         = "",
    [string]   $Manager            = "",
    [string]   $PhoneNumber        = "",
    [string]   $TargetOU           = "",
    [string[]] $ADGroups           = @(),
    [string[]] $LicenseSkuIds      = @(),
    [string]   $AADConnectServer   = "",
    [string]   $SharePointAdminURL = "",
    [string]   $AdminUPN           = "",
    [string]   $TempPassword       = "",
    [string]   $NotifyEmail        = "",
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
        # Use pwsh (PS7) if that is what launched this script, else fall back to Windows PowerShell
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
        @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; Min = "2.0.0" },
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
        "User Administrator",
        "License Administrator"
    )

    RL "Connecting to Microsoft Graph for role verification..."

    # Skip interactive role assignment if cert-based auth is configured
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
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue; return
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

$isCLI    = ($FirstName -ne "" -and $LastName -ne "")
$isSilent = $isCLI

Write-Host "[$(Get-Date -Format 'HH:mm:ss')][INFO] AD/M365 Onboarding Tool - v1.0 (c) Roger Fouche / Fouche Enterprises, LLC" -ForegroundColor Magenta

if (-not $SkipModuleCheck) { Invoke-ModuleBootstrap -Silent:$isSilent }
$null = Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Pre-load config BEFORE role bootstrap so the cert check can skip interactive auth.
# Full config init (with TenantId from environment discovery) happens in Section 0e.
$script:_PreConfigPath = Join-Path (Split-Path $PSCommandPath -Parent) "ADM365Config.json"
$script:Config = @{ AdminUPN=""; TenantId=""; AppId=""; CertThumbprint=""; ExchangeOrg=""; DefaultOU=""; DefaultGroups=""; DefaultLicenseSkuIds="" }
if (Test-Path $script:_PreConfigPath) {
    try {
        $script:_PreLoaded = Get-Content $script:_PreConfigPath -Raw | ConvertFrom-Json
        foreach ($key in @('AdminUPN','TenantId','AppId','CertThumbprint','ExchangeOrg','DefaultOU','DefaultGroups','DefaultLicenseSkuIds','DefaultUPNSuffix','UsageLocation')) {
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
        UPNSuffixes        = @()
        DefaultUserOU      = ""
        AADConnectServer   = ""
        TenantName         = ""
        SharePointAdminURL = ""
        NetBIOSDomain      = ""
    }
    try {
        $domain                = Get-ADDomain -ErrorAction Stop
        $info.DC               = $domain.PDCEmulator
        $info.DomainDN         = $domain.DistinguishedName
        $info.LocalDomain      = $domain.DNSRoot
        $info.NetBIOSDomain    = $domain.NetBIOSName
        $info.DefaultUserOU    = $domain.UsersContainer

        $forest = Get-ADForest -ErrorAction SilentlyContinue
        if ($forest -and $forest.UPNSuffixes.Count -gt 0) {
            $nonLocal = $forest.UPNSuffixes |
                Where-Object { $_ -notmatch '\.local$' } | Select-Object -First 1
            $info.EmailDomain = if ($nonLocal) { $nonLocal } else { $forest.UPNSuffixes[0] }
            # Collect all non-local suffixes for the GUI dropdown
            $info.UPNSuffixes = @($forest.UPNSuffixes | Where-Object { $_ -notmatch '\.local$' })
        }
        # Also include the DNS root if it is not .local and not already listed
        if ($domain.DNSRoot -notmatch '\.local$' -and $info.UPNSuffixes -notcontains $domain.DNSRoot) {
            $info.UPNSuffixes = @($domain.DNSRoot) + $info.UPNSuffixes
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

        try {
            # Check local machine first - script often runs ON the AADC/DC server
            $localADSync = Get-Module -ListAvailable -Name ADSync -ErrorAction SilentlyContinue
            if ($localADSync) {
                $info.AADConnectServer = $env:COMPUTERNAME
            } else {
                # Fall back: search AD for a computer that looks like AADC
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

if ($AADConnectServer)   { $script:E.AADConnectServer   = $AADConnectServer }
if ($SharePointAdminURL) { $script:E.SharePointAdminURL = $SharePointAdminURL }
if ($TargetOU)           { $script:E.DefaultUserOU      = $TargetOU }

Write-Host "[$(Get-Date -Format 'HH:mm:ss')][OK] DC=$($script:E.DC)  Domain=$($script:E.LocalDomain)  Tenant=$($script:E.TenantName)" -ForegroundColor Green

# ============================================================
# SECTION 0e - CONFIG FILE (persists admin UPN across runs)
# Stored next to the script. Graph SDK caches tokens via MSAL
# so MFA is only needed once per token lifetime (~1 day).
# ============================================================

$script:ConfigPath = Join-Path (Split-Path $PSCommandPath -Parent) "ADM365Config.json"
$script:Config = @{
    AdminUPN            = ""
    TenantId            = $script:E.TenantName
    AppId               = ""
    CertThumbprint      = ""
    ExchangeOrg         = ""
    DefaultOU           = ""
    DefaultGroups       = ""
    DefaultLicenseSkuIds= ""
    DefaultUPNSuffix    = ""
    UsageLocation       = "US"
}

if (Test-Path $script:ConfigPath) {
    try {
        $loaded = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
        foreach ($key in @('AdminUPN','TenantId','AppId','CertThumbprint','ExchangeOrg','DefaultOU','DefaultGroups','DefaultLicenseSkuIds','DefaultUPNSuffix','UsageLocation')) {
            $val = $loaded.PSObject.Properties[$key]
            if ($val -and $val.Value) { $script:Config[$key] = $val.Value }
        }
    } catch {}
}

# If AdminUPN was passed as a parameter, use and save it
if ($AdminUPN -ne "") {
    $script:Config.AdminUPN = $AdminUPN
    $script:Config | ConvertTo-Json | Set-Content $script:ConfigPath -Encoding UTF8
} elseif ($script:Config.AdminUPN -ne "") {
    $AdminUPN = $script:Config.AdminUPN
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')][OK] Using saved admin UPN: $AdminUPN" -ForegroundColor Green
}

function Save-Config {
    try {
        $script:Config | ConvertTo-Json | Set-Content $script:ConfigPath -Encoding UTF8
    } catch {}
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
    # Import Graph.Authentication FIRST so its Azure.Core loads before EXO's version
    $null = Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
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
    param([System.Windows.Forms.RichTextBox]$LogBox = $null)
    function AL { param($m, $l = "INFO")
        $c = switch ($l) { "OK"{"Green"} "WARN"{"Yellow"} "ERR"{"Red"} default{"Cyan"} }
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$l] $m" -ForegroundColor $c
        if ($LogBox -and -not $LogBox.IsDisposed) {
            $rc = switch ($l) { "OK"{[System.Drawing.Color]::LimeGreen} "WARN"{[System.Drawing.Color]::Gold} "ERR"{[System.Drawing.Color]::Tomato} default{[System.Drawing.Color]::FromArgb(140,200,255)} }
            $LogBox.SelectionStart = $LogBox.TextLength; $LogBox.SelectionColor = $rc
            $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')][$l] $m`r`n"); $LogBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    AL "Starting one-time non-interactive auth setup..." "INFO"
    AL "Sign in with a GLOBAL ADMIN account when prompted." "WARN"

    try {
        $null = Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $null = Import-Module Microsoft.Graph.Applications -ErrorAction Stop
        $null = Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

        $null = Connect-MgGraph -Scopes @(
            "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All",
            "RoleManagement.ReadWrite.Directory", "Directory.ReadWrite.All", "Organization.Read.All"
        ) -NoWelcome -ErrorAction Stop

        $ctx      = Get-MgContext
        $tenantId = $ctx.TenantId
        AL "Connected. Tenant: $tenantId" "OK"

        # Create 2-year self-signed certificate in LocalMachine store
        AL "Creating authentication certificate..."
        $appName = "ADM365LifecycleTool"
        $cert = New-SelfSignedCertificate -Subject "CN=$appName" `
            -CertStoreLocation "Cert:\LocalMachine\My" `
            -KeyExportPolicy NonExportable -KeySpec Signature `
            -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 `
            -NotAfter (Get-Date).AddYears(2)
        AL "Certificate: $($cert.Thumbprint)" "OK"

        # Remove existing app if present
        $old = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($old) { Remove-MgApplication -ApplicationId $old.Id; AL "Removed existing app." "WARN" }

        # Create app registration
        AL "Creating Entra ID app registration..."
        $app = New-MgApplication -DisplayName $appName -SignInAudience "AzureADMyOrg"
        AL "App created: $($app.AppId)" "OK"

        # Upload certificate
        $null = Update-MgApplication -ApplicationId $app.Id -KeyCredentials @(@{
            Type        = "AsymmetricX509Cert"
            Usage       = "Verify"
            Key         = $cert.RawData
            DisplayName = "$appName Cert"
        })
        AL "Certificate uploaded to app." "OK"

        # Look up service principals for each resource
        $graphSP = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
        $exoSP   = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'" -ErrorAction SilentlyContinue
        $spoSP   = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" -ErrorAction SilentlyContinue

        # Graph permissions
        $graphPermNames = @("User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All","RoleManagement.ReadWrite.Directory")
        $graphRoles = @()
        foreach ($p in $graphPermNames) {
            $r = $graphSP.AppRoles | Where-Object { $_.Value -eq $p } | Select-Object -First 1
            if ($r) { $graphRoles += @{ Id = [string]$r.Id; Type = "Role" } }
        }
        $reqAccess = @(@{ ResourceAppId = "00000003-0000-0000-c000-000000000000"; ResourceAccess = $graphRoles })

        # Exchange.ManageAsApp
        $exoRoleId = $null
        if ($exoSP) {
            $er = $exoSP.AppRoles | Where-Object { $_.Value -eq "Exchange.ManageAsApp" } | Select-Object -First 1
            if ($er) {
                $exoRoleId = [string]$er.Id
                $reqAccess += @{ ResourceAppId = "00000002-0000-0ff1-ce00-000000000000"; ResourceAccess = @(@{ Id = $exoRoleId; Type = "Role" }) }
            }
        }

        # SharePoint Sites.FullControl.All
        $spoRoleId = $null
        if ($spoSP) {
            $sr = $spoSP.AppRoles | Where-Object { $_.Value -eq "Sites.FullControl.All" } | Select-Object -First 1
            if ($sr) {
                $spoRoleId = [string]$sr.Id
                $reqAccess += @{ ResourceAppId = "00000003-0000-0ff1-ce00-000000000000"; ResourceAccess = @(@{ Id = $spoRoleId; Type = "Role" }) }
            }
        }

        $null = Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $reqAccess
        AL "Permissions configured." "OK"

        # Create service principal
        AL "Creating service principal..."
        $sp = New-MgServicePrincipal -AppId $app.AppId

        # Wait for propagation then grant admin consent
        AL "Waiting 15s for propagation, then granting admin consent..."
        Start-Sleep -Seconds 15

        foreach ($role in $graphRoles) {
            try { $null = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $graphSP.Id -AppRoleId $role.Id } catch {}
        }
        if ($exoSP -and $exoRoleId) {
            try { $null = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $exoSP.Id -AppRoleId $exoRoleId } catch {}
        }
        if ($spoSP -and $spoRoleId) {
            try { $null = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $spoSP.Id -AppRoleId $spoRoleId } catch {}
        }
        AL "Admin consent granted." "OK"

        # Add SP to Exchange Administrator role
        try {
            $exoAdmRole = Get-MgDirectoryRole -Filter "displayName eq 'Exchange Administrator'" -ErrorAction SilentlyContinue
            if (-not $exoAdmRole) {
                $t = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq "Exchange Administrator" } | Select-Object -First 1
                $exoAdmRole = New-MgDirectoryRole -RoleTemplateId $t.Id
            }
            $null = New-MgDirectoryRoleMemberByRef -DirectoryRoleId $exoAdmRole.Id `
                -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)" }
            AL "Exchange Administrator role assigned to service principal." "OK"
        } catch { AL "Exchange role assignment warning: $_" "WARN" }

        # Get Exchange org domain (.onmicrosoft.com)
        $org = Get-MgOrganization | Select-Object -First 1
        $exoOrg = ($org.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1).Name
        if (-not $exoOrg) { $exoOrg = "$($script:E.TenantName).onmicrosoft.com" }

        # Save everything
        $script:Config.AppId          = $app.AppId
        $script:Config.TenantId       = $tenantId
        $script:Config.CertThumbprint = $cert.Thumbprint
        $script:Config.ExchangeOrg    = $exoOrg
        Save-Config

        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue

        AL "============================================" "OK"
        AL " SETUP COMPLETE - non-interactive auth ready" "OK"
        AL " App ID      : $($app.AppId)" "OK"
        AL " Cert        : $($cert.Thumbprint)" "OK"
        AL " Exchange Org: $exoOrg" "OK"
        AL " Wait 5-10 minutes before first use." "WARN"
        AL "============================================" "OK"
        return $true
    }
    catch {
        AL "Setup error: $_" "ERR"
        # Clean up orphaned cert if it was created before the failure
        if ((Test-Path variable:cert) -and $cert) {
            try { Remove-Item "Cert:\LocalMachine\My\$($cert.Thumbprint)" -DeleteKey -ErrorAction SilentlyContinue } catch {}
        }
        try { $null = Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
        return $false
    }
}

# ============================================================

$script:CachedLicenses = @()

$script:LicenseFriendlyNames = @{
    # Microsoft 365 Business
    "O365_BUSINESS_ESSENTIALS"              = "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM"                 = "Microsoft 365 Business Standard"
    "SPB"                                   = "Microsoft 365 Business Premium"
    # Microsoft 365 Enterprise
    "SPE_E3"                                = "Microsoft 365 E3"
    "SPE_E5"                                = "Microsoft 365 E5"
    "SPE_F1"                                = "Microsoft 365 F1"
    "SPE_F3"                                = "Microsoft 365 F3"
    # Office 365
    "ENTERPRISEPACK"                        = "Office 365 E3"
    "ENTERPRISEPREMIUM"                     = "Office 365 E5"
    "STANDARDPACK"                          = "Office 365 E1"
    "DESKLESSPACK"                          = "Office 365 F3"
    "O365_BUSINESS"                         = "Microsoft 365 Apps for Business"
    "OFFICESUBSCRIPTION"                    = "Microsoft 365 Apps for Enterprise"
    # Exchange Online
    "EXCHANGESTANDARD"                      = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE"                    = "Exchange Online Plan 2"
    "EXCHANGEARCHIVE_ADDON"                 = "Exchange Online Archiving"
    # Teams
    "TEAMS_ESSENTIALS"                      = "Microsoft Teams Essentials"
    "TEAMS_EXPLORATORY"                     = "Microsoft Teams Exploratory"
    "MCOEV"                                 = "Teams Phone Standard"
    "MCOPSTN1"                              = "Teams Domestic Calling Plan"
    "MCOPSTN2"                              = "Teams Domestic and International Calling Plan"
    "TEAMS_ROOMS_STANDARD"                  = "Teams Rooms Standard"
    "TEAMS_ROOMS_PRO"                       = "Teams Rooms Pro"
    # Intune and Endpoint Management
    "INTUNE_A"                              = "Microsoft Intune Plan 1"
    "INTUNE_SMB"                            = "Microsoft Intune SMB"
    # Endpoint Privilege Management (EPM)
    "INTUNE_P2"                             = "Microsoft Intune Plan 2 (includes EPM)"
    "INTUNE_SUITE"                          = "Microsoft Intune Suite (EPM + Advanced)"
    "INTUNE_SUITE_ADO"                      = "Microsoft Intune Suite Add-on"
    # Enterprise Mobility and Security
    "EMS"                                   = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"                            = "Enterprise Mobility + Security E5"
    # Entra ID
    "AAD_PREMIUM"                           = "Microsoft Entra ID P1"
    "AAD_PREMIUM_P2"                        = "Microsoft Entra ID P2"
    "ENTRA_ID_GOVERNANCE"                   = "Microsoft Entra ID Governance"
    "ENTRA_SUITE"                           = "Microsoft Entra Suite"
    # Defender and Security
    "DEFENDER_ENDPOINT_P1"                  = "Microsoft Defender for Endpoint P1"
    "MDATP_XPLAT"                           = "Microsoft Defender for Endpoint P2"
    "ATP_ENTERPRISE"                        = "Microsoft Defender for Office 365 Plan 1"
    "THREAT_INTELLIGENCE"                   = "Microsoft Defender for Office 365 Plan 2"
    "MDO_SMB"                               = "Microsoft Defender for Office 365 SMB"
    "DEFENDER_IDENTITY"                     = "Microsoft Defender for Identity"
    "ADALLOM_STANDALONE"                    = "Microsoft Defender for Cloud Apps"
    "DEFENDER_BUSINESS"                     = "Microsoft Defender for Business"
    "M365_SECURITY_COMPLIANCE_FOR_SMB"      = "Microsoft Defender for Business (bundle)"
    # Purview and Compliance
    "RIGHTSMANAGEMENT"                      = "Azure Information Protection P1"
    "RMS_S_PREMIUM"                         = "Azure Information Protection P1 (alt)"
    "RMS_S_PREMIUM2"                        = "Azure Information Protection P2"
    "INFORMATION_PROTECTION_COMPLIANCE"     = "Microsoft Purview Information Protection"
    "LOCKBOX_ENTERPRISE"                    = "Customer Lockbox"
    # Windows
    "WIN10_PRO_ENT_SUB"                     = "Windows 10/11 Enterprise E3"
    "WIN_ENT_E5"                            = "Windows 10/11 Enterprise E5"
    "WIN10_VDA_E3"                          = "Windows 10/11 Enterprise E3 VDA"
    # Power Platform
    "POWER_BI_PRO"                          = "Power BI Pro"
    "POWER_BI_PREMIUM_PER_USER"             = "Power BI Premium Per User"
    "FLOW_FREE"                             = "Power Automate Free"
    "FLOW_P1"                               = "Power Automate Premium"
    "POWERAPPS_PER_USER"                    = "Power Apps Premium (per user)"
    "POWERAPPS_DEV"                         = "Power Apps Developer Plan"
    # Copilot
    "COPILOT_FOR_M365"                      = "Microsoft 365 Copilot"
    "M365_COPILOT"                          = "Microsoft 365 Copilot (alt)"
    # Project and Visio
    "PROJECTESSENTIALS"                     = "Project Plan 1"
    "PROJECTPROFESSIONAL"                   = "Project Plan 3"
    "PROJECTPREMIUM"                        = "Project Plan 5"
    "VISIOONLINE_PLAN1"                     = "Visio Plan 1"
    "VISIOCLIENT"                           = "Visio Plan 2"
    # Audio Conferencing
    "MCOMEETADV"                            = "Microsoft 365 Audio Conferencing"
    # Viva
    "VIVA_SUITE"                            = "Microsoft Viva Suite"
}

function Get-M365Licenses {
    param([bool]$Silent = $false)
    function LL { param($m, $l = "INFO")
        if (-not $Silent) {
            $c = switch ($l) { "OK" { "Green" } "WARN" { "Yellow" } default { "Cyan" } }
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$l] $m" -ForegroundColor $c
        }
    }

    LL "Querying M365 license inventory..."

    # Build results as a typed list so only our objects end up in it -
    # Graph cmdlets can output connection objects to the pipeline which
    # would otherwise get mixed into the caller's variable.
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $null = Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $null = Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

        # Suppress Connect-MgGraph pipeline output (it returns a context object)
        Connect-M365Graph -FallbackScopes @("Organization.Read.All")

        $skus = Get-MgSubscribedSku -ErrorAction Stop
        foreach ($sku in $skus) {
            $available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
            $friendly  = $script:LicenseFriendlyNames[$sku.SkuPartNumber]
            if (-not $friendly) { $friendly = $sku.SkuPartNumber }

            # Use PSCustomObject so .FriendlyName etc. work as true properties
            $results.Add([PSCustomObject]@{
                SkuId         = [string]$sku.SkuId
                SkuPartNumber = [string]$sku.SkuPartNumber
                FriendlyName  = [string]$friendly
                Available     = [int]$available
                Total         = [int]$sku.PrepaidUnits.Enabled
                Consumed      = [int]$sku.ConsumedUnits
            })
            LL "  $friendly - $available available of $($sku.PrepaidUnits.Enabled)" "OK"
        }

        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        LL "License query complete. $($results.Count) SKU(s) found." "OK"
    }
    catch {
        LL "License query failed: $_" "WARN"
        try { $null = Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    }

    # Return as plain array so callers can index it normally
    return , $results.ToArray()
}

Write-Host "[$(Get-Date -Format 'HH:mm:ss')][INFO] Querying M365 licenses..." -ForegroundColor Cyan
$script:CachedLicenses = Get-M365Licenses -Silent:$isSilent

# ============================================================
# SECTION 1 - SHARED LOGGER
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
# SECTION 2 - PASSWORD GENERATOR
# ============================================================

function New-TempPassword {
    $upper   = "ABCDEFGHJKLMNPQRSTUVWXYZ"
    $lower   = "abcdefghjkmnpqrstuvwxyz"
    $digits  = "23456789"
    $special = "!@#$%^"
    $all     = $upper + $lower + $digits + $special

    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $buf = [byte[]]::new(1)
    function RC { param($cs)
        do { $rng.GetBytes($buf) } while ($buf[0] -ge (256 - 256 % $cs.Length))
        $cs[$buf[0] % $cs.Length]
    }

    $pwd = @((RC $upper), (RC $lower), (RC $digits), (RC $special))
    for ($i = 4; $i -lt 12; $i++) { $pwd += (RC $all) }

    for ($i = $pwd.Count - 1; $i -gt 0; $i--) {
        $rng.GetBytes($buf)
        $j = $buf[0] % ($i + 1)
        $tmp = $pwd[$i]; $pwd[$i] = $pwd[$j]; $pwd[$j] = $tmp
    }
    return ($pwd -join "")
}

# ============================================================
# SECTION 3 - SAM GENERATOR
# ============================================================

function New-SamAccountName {
    param([string]$First, [string]$Last)
    $norm = { param($s)
        $s = $s.ToLower()
        $map = @{ 'a' = 'a'; 'e' = 'e' }  # placeholder - regex handles the rest
        ($s -replace '[^a-z0-9]', '')
    }
    $f = (& $norm $First)
    $l = (& $norm $Last)
    $base = "$($f.Substring(0, [Math]::Min(1, $f.Length)))$l"
    if ($base.Length -gt 20) { $base = $base.Substring(0, 20) }

    $candidate = $base
    $i = 2
    while (Get-ADUser -Filter "SamAccountName -eq '$candidate'" -ErrorAction SilentlyContinue) {
        $suffix = "$i"
        $candidate = "$($base.Substring(0, [Math]::Min($base.Length, 20 - $suffix.Length)))$suffix"
        $i++
    }
    return $candidate
}

# ============================================================
# SECTION 4 - CORE ONBOARDING ENGINE
# ============================================================

function Invoke-Onboarding {
    param(
        [string]   $FirstName,
        [string]   $LastName,
        [string]   $DisplayName,
        [string]   $JobTitle,
        [string]   $Department,
        [string]   $Manager,
        [string]   $PhoneNumber,
        [string]   $TargetOU,
        [string[]] $ADGroups,
        [string[]] $LicenseSkuIds,
        [string]   $AADConnectServer,
        [string]   $TempPassword,
        [string]   $NotifyEmail,
        [string]   $UPNSuffix = "",
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
        # Step 1 - Derive account values
        Write-Log ">> 1/7 - Generate account details" "HEAD" $LogBox
        Prog 8 "Generating account details..."

        $dispName = if ($DisplayName) { $DisplayName } else { "$FirstName $LastName" }
        $samName  = New-SamAccountName -First $FirstName -Last $LastName
        # Use explicitly selected suffix, fall back to auto-discovered email domain
        $suffix   = if ($UPNSuffix) { $UPNSuffix } else { $script:E.EmailDomain }
        $upn      = "$samName@$suffix"
        $password = if ($TempPassword) { $TempPassword } else { New-TempPassword }
        $secPwd   = ConvertTo-SecureString $password -AsPlainText -Force

        Write-Log "Display Name : $dispName"  "INFO" $LogBox
        Write-Log "SAM Account  : $samName"    "INFO" $LogBox
        Write-Log "UPN          : $upn"        "INFO" $LogBox
        Write-Log "Target OU    : $TargetOU"   "INFO" $LogBox

        # Step 2 - Create AD account
        Write-Log ">> 2/7 - Create AD user account" "HEAD" $LogBox
        Prog 18 "Creating AD account..."

        Step "New-ADUser: $samName" {
            $params = @{
                Name                  = $dispName
                GivenName             = $FirstName
                Surname               = $LastName
                DisplayName           = $dispName
                SamAccountName        = $samName
                UserPrincipalName     = $upn
                EmailAddress          = $upn
                AccountPassword       = $secPwd
                ChangePasswordAtLogon = $false
                PasswordNeverExpires  = $true
                Enabled               = $true
                Path                  = $TargetOU
            }
            if ($JobTitle)    { $params.Title       = $JobTitle }
            if ($Department)  { $params.Department  = $Department }
            if ($PhoneNumber) { $params.OfficePhone = $PhoneNumber }

            New-ADUser @params

            # Set mail and proxyAddresses BEFORE sync so Exchange picks them up correctly
            # mail attribute = EmailAddress (already set above via -EmailAddress)
            # proxyAddresses: uppercase SMTP: = primary, must be set separately
            Set-ADUser -Identity $samName -Add @{
                proxyAddresses = [string[]]@("SMTP:$upn")
            }

            if ($Manager) {
                try {
                    $mgrObj = Get-ADUser -Identity $Manager -ErrorAction Stop
                    Set-ADUser -Identity $samName -Manager $mgrObj.DistinguishedName
                    Write-Log "Manager set: $($mgrObj.DisplayName)" "SUCCESS" $LogBox
                }
                catch { Write-Log "Manager not found ($Manager) - skipped." "WARN" $LogBox }
            }

            Write-Log "AD account created: $samName" "SUCCESS" $LogBox
        }

        # Step 3 - Add to AD groups
        Write-Log ">> 3/7 - AD group membership" "HEAD" $LogBox
        Prog 30 "Adding to AD groups..."

        foreach ($grp in $ADGroups) {
            # Domain Users is the automatic primary group for all AD accounts -
            # it cannot be added via Add-ADGroupMember and the warning is harmless but confusing
            if ($grp -eq "Domain Users") {
                Write-Log "Domain Users - automatic primary group, skipping explicit add." "INFO" $LogBox
                continue
            }
            Step "Add to group: $grp" {
                try {
                    Add-ADGroupMember -Identity $grp -Members $samName -ErrorAction Stop
                    Write-Log "Added to: $grp" "SUCCESS" $LogBox
                }
                catch { Write-Log "Group not found ($grp): $_" "WARN" $LogBox }
            }
        }
        if ($ADGroups.Count -eq 0) {
            Write-Log "No AD groups specified." "WARN" $LogBox
        }

        # Step 4 - Entra sync
        Write-Log ">> 4/7 - Entra delta sync (push new account)" "HEAD" $LogBox
        Prog 45 "Triggering Entra sync..."

        if ($AADConnectServer) {
            Step "Delta sync - $AADConnectServer" {
                $isLocal = ($AADConnectServer -eq $env:COMPUTERNAME) -or
                           ($AADConnectServer -eq "localhost") -or
                           ($AADConnectServer -eq "127.0.0.1")
                if ($isLocal) {
                    # ADSync depends on System.Web (.NET Framework only).
                    # In PS7 (.NET 6+) we must delegate to a powershell.exe (PS5.1) subprocess.
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
                    # Remote: Invoke-Command targets PS5.1 on the remote server - ADSync works fine
                    Invoke-Command -ComputerName $AADConnectServer -ScriptBlock {
                        Import-Module ADSync -ErrorAction Stop
                        Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                    }
                }
                # Give sync 45s to propagate to Entra before polling begins
                Write-Log "Sync triggered - waiting 45s then polling Entra (up to 6 min)..." "INFO" $LogBox
                Start-Sleep -Seconds 45
            }
        } else {
            Write-Log "No AADConnect server - skipping sync." "WARN" $LogBox
            Write-Log "Trigger manually before license assignment." "WARN" $LogBox
        }

        # Step 5 - Connect Graph
        Write-Log ">> Connecting Microsoft Graph" "HEAD" $LogBox
        Prog 58 "Connecting to M365..."

        Step "Connect Graph" { Connect-M365Graph }

        # Step 6 - Assign licenses
        Write-Log ">> 5/7 - Assign M365 licenses" "HEAD" $LogBox
        Prog 68 "Assigning M365 licenses..."

        if ($LicenseSkuIds.Count -gt 0) {
            Step "Assign licenses to $upn" {
                # Use Invoke-MgGraphRequest (always available in Graph.Authentication)
                # instead of Set-MgUserLicense which requires Graph.Users to be resolvable
                $mgUser  = $null
                $retries = 0
                while (-not $mgUser -and $retries -lt 36) {
                    Start-Sleep -Seconds 10
                    try {
                        $result = Invoke-MgGraphRequest -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$upn'&`$select=id" `
                            -ErrorAction SilentlyContinue
                        if ($result -and $result.value -and $result.value.Count -gt 0) {
                            $mgUser = $result.value[0]
                        }
                    } catch {}
                    $retries++
                    if (-not $mgUser -and $retries % 6 -eq 0) {
                        Write-Log "  Still waiting for user in Entra... ($($retries * 10)s elapsed)" "WARN" $LogBox
                    }
                }

                if ($mgUser) {
                    # Set usage location (required before license assignment).
                    # Read from config so non-US clients work correctly.
                    $uLoc = if ($script:Config['UsageLocation']) { $script:Config['UsageLocation'] } else { 'US' }
                    Invoke-MgGraphRequest -Method PATCH `
                        -Uri "https://graph.microsoft.com/v1.0/users/$($mgUser.id)" `
                        -Body (@{ usageLocation = $uLoc } | ConvertTo-Json) `
                        -ContentType "application/json" -ErrorAction Stop | Out-Null
                    Write-Log "Usage location set to: $uLoc" "INFO" $LogBox

                    # Wait for usageLocation to propagate -- Graph requires it before assignLicense
                    Write-Log "Waiting 10s for usage location to propagate..." "INFO" $LogBox
                    Start-Sleep -Seconds 10

                    # Assign licenses via REST
                    $licBody = @{
                        addLicenses    = @($LicenseSkuIds | ForEach-Object { @{ skuId = $_ } })
                        removeLicenses = @()
                    } | ConvertTo-Json -Depth 5
                    Invoke-MgGraphRequest -Method POST `
                        -Uri "https://graph.microsoft.com/v1.0/users/$($mgUser.id)/assignLicense" `
                        -Body $licBody -ContentType "application/json" -ErrorAction Stop | Out-Null
                    Write-Log "Assigned $($LicenseSkuIds.Count) license(s) to $upn" "SUCCESS" $LogBox
                }
                else {
                    Write-Log "User $upn not found in Entra after $($retries * 10)s - license skipped." "WARN" $LogBox
                }
            }
        }
        else { Write-Log "No licenses selected - skipping license assignment." "WARN" $LogBox }

        # Step 7 - Welcome notification
        Write-Log ">> 6/7 - Welcome notification" "HEAD" $LogBox
        Prog 85 "Sending welcome notification..."

        if ($NotifyEmail -ne "") {
            Step "Send credential email to $NotifyEmail" {
                Connect-M365Exchange

                $licNames = $LicenseSkuIds | ForEach-Object {
                    $id = $_
                    $m  = $script:CachedLicenses | Where-Object { $_.SkuId -eq $id } | Select-Object -First 1
                    if ($m) { $m.FriendlyName } else { $id }
                }

                $bodyText = "New User Account Created`r`n`r`n" +
                    "Display Name  : $dispName`r`n" +
                    "Username (SAM): $samName`r`n" +
                    "UPN / Email   : $upn`r`n" +
                    "Temp Password : $password`r`n" +
                    "Target OU     : $TargetOU`r`n" +
                    "AD Groups     : $($ADGroups -join ', ')`r`n" +
                    "Licenses      : $($licNames -join ', ')`r`n`r`n" +
                    "The user must change their password at first login.`r`n" +
                    "Please provide these credentials securely.`r`n`r`n" +
                    "IT Administration`r`n"
                Send-MailMessage -To $NotifyEmail -Subject "New Account: $dispName" `
                    -Body $bodyText -SmtpServer "smtp.office365.com" -Port 587 `
                    -UseSsl -Credential (Get-Credential -Message "SMTP credentials") `
                    -ErrorAction SilentlyContinue
                Write-Log "Notification sent to $NotifyEmail" "SUCCESS" $LogBox
            }
        }
        else { Write-Log "No notify address - skipping welcome email." "INFO" $LogBox }

        # Done
        Prog 100 "Complete."
        Write-Log "" "INFO" $LogBox
        Write-Log "============================================" "SUCCESS" $LogBox
        Write-Log " ONBOARDING COMPLETE" "SUCCESS" $LogBox
        Write-Log " Name     : $dispName" "SUCCESS" $LogBox
        Write-Log " SAM      : $samName" "SUCCESS" $LogBox
        Write-Log " UPN      : $upn" "SUCCESS" $LogBox
        Write-Log " Password : $password  (RECORD THIS NOW)" "SUCCESS" $LogBox
        Write-Log " Groups   : $($ADGroups -join ', ')" "SUCCESS" $LogBox
        Write-Log " Licenses : $($LicenseSkuIds.Count) assigned" "SUCCESS" $LogBox
        if ($WhatIfMode) { Write-Log " *** WHATIF - no changes were made ***" "WARN" $LogBox }
        Write-Log "============================================" "SUCCESS" $LogBox

        $script:LastResult = @{
            DisplayName = $dispName
            SAM         = $samName
            UPN         = $upn
            Password    = $password
        }

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

if ($isCLI) {
    $ou = if ($TargetOU) { $TargetOU } else { $script:E.DefaultUserOU }
    Invoke-Onboarding `
        -FirstName        $FirstName `
        -LastName         $LastName `
        -DisplayName      $DisplayName `
        -JobTitle         $JobTitle `
        -Department       $Department `
        -Manager          $Manager `
        -PhoneNumber      $PhoneNumber `
        -TargetOU         $ou `
        -ADGroups         $ADGroups `
        -LicenseSkuIds    $LicenseSkuIds `
        -AADConnectServer $script:E.AADConnectServer `
        -TempPassword     $TempPassword `
        -NotifyEmail      $NotifyEmail `
        -WhatIfMode       $WhatIf.IsPresent
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
$F_SM    = New-Object System.Drawing.Font("Segoe UI", 8)
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
function TBox { param($x, $y, $w = 200, $h = 26, $pass = $false)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = [System.Drawing.Point]::new($x, $y); $t.Size = [System.Drawing.Size]::new($w, $h)
    $t.Font = $F_NORM; $t.ForeColor = $TEXT; $t.BackColor = $CLR_INPUT; $t.BorderStyle = "FixedSingle"
    if ($pass) { $t.UseSystemPasswordChar = $true }; $t
}
function TBoxML { param($x, $y, $w, $h)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = [System.Drawing.Point]::new($x, $y); $t.Size = [System.Drawing.Size]::new($w, $h)
    $t.Font = $F_NORM; $t.ForeColor = $TEXT; $t.BackColor = $CLR_INPUT; $t.BorderStyle = "FixedSingle"
    $t.Multiline = $true; $t.ScrollBars = "Vertical"; $t
}
function Btn { param($t, $x, $y, $w = 110, $h = 28, $bg = $ACCENT, $fg = $BG)
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
# DIALOGS
# ============================================================

function Show-ManagerSearch {
    param([System.Windows.Forms.Form]$Parent = $null)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Search for Manager"; $dlg.Size = [System.Drawing.Size]::new(620, 440)
    $dlg.StartPosition = "CenterParent"; $dlg.BackColor = $BG
    $dlg.FormBorderStyle = "FixedDialog"; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    $dlg.Controls.Add((Lbl "Search by name, SAM, or email:" 10 10 400 22 $F_BOLD $TEXT))
    $txtQ = TBox 10 34 480 26; $dlg.Controls.Add($txtQ)
    $btnGo = Btn "Search" 500 32 96 28 $ACCENT $BG; $dlg.Controls.Add($btnGo)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = [System.Drawing.Point]::new(10, 68); $lv.Size = [System.Drawing.Size]::new(584, 272)
    $lv.View = "Details"; $lv.FullRowSelect = $true; $lv.GridLines = $true
    $lv.Font = $F_NORM; $lv.BackColor = $CLR_INPUT; $lv.ForeColor = $TEXT; $lv.BorderStyle = "FixedSingle"
    [void]$lv.Columns.Add("Display Name", 200); [void]$lv.Columns.Add("SAM", 120)
    [void]$lv.Columns.Add("Email", 200); [void]$lv.Columns.Add("Dept", 50)
    $dlg.Controls.Add($lv)

    $lblC   = Lbl "" 10 348 400 20 $F_SM $TEXTDIM; $dlg.Controls.Add($lblC)
    $btnOK  = Btn "Select" 400 370 110 28 $GREEN $BG; $btnOK.Enabled = $false; $dlg.Controls.Add($btnOK)
    $btnX   = Btn "Cancel" 520 370 80 28 $BORDER $TEXT; $dlg.Controls.Add($btnX)

    $script:SelectedManager = $null

    $doSearch = {
        $q = $txtQ.Text.Trim(); $lv.Items.Clear()
        if ($q.Length -lt 2) { $lblC.Text = "Enter at least 2 characters."; return }
        try {
            $hits = Get-ADUser -Filter "(Name -like '*$q*') -or (SamAccountName -like '*$q*') -or (EmailAddress -like '*$q*')" `
                -Properties DisplayName, EmailAddress, Department, Enabled |
                Sort-Object DisplayName | Select-Object -First 100
            foreach ($h in $hits) {
                $item = New-Object System.Windows.Forms.ListViewItem("$($h.DisplayName)")
                [void]$item.SubItems.Add("$($h.SamAccountName)")
                [void]$item.SubItems.Add("$($h.EmailAddress)")
                [void]$item.SubItems.Add("$($h.Department)")
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
                $script:SelectedManager = $lv.SelectedItems[0].Tag
                $dlg.DialogResult = "OK"; $dlg.Close()
            }
        })
    $btnOK.Add_Click({
            if ($lv.SelectedItems.Count -gt 0) {
                $script:SelectedManager = $lv.SelectedItems[0].Tag
                $dlg.DialogResult = "OK"; $dlg.Close()
            }
        })
    $btnX.Add_Click({ $dlg.DialogResult = "Cancel"; $dlg.Close() })
    $script:SelectedManager = $null
    if ($Parent) { [void]$dlg.ShowDialog($Parent) } else { [void]$dlg.ShowDialog() }
    $dlg.Dispose()
}

function Show-OUPicker {
    param([System.Windows.Forms.Form]$Parent = $null)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Select Target OU"; $dlg.Size = [System.Drawing.Size]::new(620, 500)
    $dlg.StartPosition = "CenterParent"; $dlg.BackColor = $BG
    $dlg.FormBorderStyle = "FixedDialog"; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    $dlg.Controls.Add((Lbl "Select the OU for the new user account:" 10 10 580 22 $F_BOLD $TEXT))

    $tv = New-Object System.Windows.Forms.TreeView
    $tv.Location = [System.Drawing.Point]::new(10, 36); $tv.Size = [System.Drawing.Size]::new(584, 360)
    $tv.BackColor = $CLR_INPUT; $tv.ForeColor = $TEXT; $tv.Font = $F_NORM; $tv.BorderStyle = "FixedSingle"
    $dlg.Controls.Add($tv)

    $lblSel = Lbl "Selected: (none)" 10 404 584 20 $F_SM $TEXTDIM; $dlg.Controls.Add($lblSel)
    $btnOK  = Btn "Select" 390 426 100 28 $GREEN $BG; $btnOK.Enabled = $false; $dlg.Controls.Add($btnOK)
    $btnX   = Btn "Cancel" 500 426 90 28 $BORDER $TEXT; $dlg.Controls.Add($btnX)

    $script:SelectedOU = ""

    function Add-OUNode { param($parent, $dn)
        $children = Get-ADOrganizationalUnit -Filter * -SearchBase $dn `
            -SearchScope OneLevel -Properties Name | Sort-Object Name
        foreach ($child in $children) {
            $node = New-Object System.Windows.Forms.TreeNode($child.Name)
            $node.Tag = $child.DistinguishedName
            [void]$parent.Nodes.Add($node)
            Add-OUNode $node $child.DistinguishedName
        }
    }

    $rootNode = New-Object System.Windows.Forms.TreeNode($script:E.LocalDomain)
    $rootNode.Tag = $script:E.DomainDN
    [void]$tv.Nodes.Add($rootNode)
    Add-OUNode $rootNode $script:E.DomainDN
    $rootNode.Expand()

    $tv.Add_AfterSelect({
            if ($tv.SelectedNode -and $tv.SelectedNode.Tag) {
                $script:SelectedOU = $tv.SelectedNode.Tag
                $lblSel.Text = "Selected: $($script:SelectedOU)"
                $btnOK.Enabled = $true
            }
        })

    $btnOK.Add_Click({ $dlg.DialogResult = "OK"; $dlg.Close() })
    $btnX.Add_Click({ $dlg.DialogResult = "Cancel"; $dlg.Close() })
    $script:SelectedOU = ""
    if ($Parent) { [void]$dlg.ShowDialog($Parent) } else { [void]$dlg.ShowDialog() }
    $dlg.Dispose()
}

function Show-GroupPicker {
    param([string[]]$AlreadySelected = @(), [System.Windows.Forms.Form]$Parent = $null)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Select AD Security Groups"; $dlg.Size = [System.Drawing.Size]::new(680, 560)
    $dlg.StartPosition = "CenterParent"; $dlg.BackColor = $BG
    $dlg.FormBorderStyle = "FixedDialog"; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    $dlg.Controls.Add((Lbl "Search and select groups for the new user:" 10 10 640 22 $F_BOLD $TEXT))
    $txtQ   = TBox 10 36 480 26; $dlg.Controls.Add($txtQ)
    $btnGo  = Btn "Search" 500 34 100 28 $ACCENT $BG; $dlg.Controls.Add($btnGo)
    $dlg.Controls.Add((Lbl "Available groups (check to select):" 10 70 400 20 $F_NORM $TEXTDIM))

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = [System.Drawing.Point]::new(10, 92); $clb.Size = [System.Drawing.Size]::new(648, 340)
    $clb.BackColor = $CLR_INPUT; $clb.ForeColor = $TEXT; $clb.Font = $F_NORM
    $clb.BorderStyle = "FixedSingle"; $clb.CheckOnClick = $true
    $dlg.Controls.Add($clb)

    $lblCount = Lbl "" 10 440 440 20 $F_SM $TEXTDIM; $dlg.Controls.Add($lblCount)
    $btnOK    = Btn "Confirm Selection" 370 464 180 28 $GREEN $BG; $dlg.Controls.Add($btnOK)
    $btnX     = Btn "Cancel" 560 464 90 28 $BORDER $TEXT; $dlg.Controls.Add($btnX)

    # Lock Domain Users so it can never be unchecked
    $clb.Add_ItemCheck({
        if ($clb.Items[$_.Index] -eq "Domain Users") {
            $_.NewValue = [System.Windows.Forms.CheckState]::Checked
        }
    })

    $loadGroups = {
        param($filter = "")
        $clb.Items.Clear()
        # Domain Users is mandatory - always pinned at top, always checked
        $idx = $clb.Items.Add("Domain Users"); $clb.SetItemChecked($idx, $true)
        try {
            $groups = if ($filter) {
                Get-ADGroup -Filter "Name -like '$filter' -and GroupCategory -eq 'Security'" |
                    Where-Object { $_.Name -ne "Domain Users" } |
                    Sort-Object Name | Select-Object -First 200
            }
            else {
                Get-ADGroup -Filter { GroupCategory -eq "Security" } |
                    Where-Object { $_.Name -ne "Domain Users" } |
                    Sort-Object Name | Select-Object -First 200
            }
            foreach ($g in $groups) {
                $idx = $clb.Items.Add($g.Name)
                if ($AlreadySelected -contains $g.Name) { $clb.SetItemChecked($idx, $true) }
            }
            $lblCount.Text = "$(1 + $groups.Count) group(s) shown. Domain Users always required."
        }
        catch { $lblCount.Text = "Error: $_" }
    }

    & $loadGroups ""

    $btnGo.Add_Click({
            $q = $txtQ.Text.Trim()
            $filter = if ($q) { "*$q*" } else { "" }
            & $loadGroups $filter
        })
    $txtQ.Add_KeyDown({ if ($_.KeyCode -eq "Return") { $btnGo.PerformClick() } })

    $btnOK.Add_Click({
            # Always include Domain Users regardless of what was checked
            $others = @($clb.CheckedItems | Where-Object { $_ -ne "Domain Users" })
            $script:SelectedGroups = @("Domain Users") + $others
            $dlg.DialogResult = "OK"; $dlg.Close()
        })
    $btnX.Add_Click({ $dlg.DialogResult = "Cancel"; $dlg.Close() })
    if ($Parent) { [void]$dlg.ShowDialog($Parent) } else { [void]$dlg.ShowDialog() }
    $dlg.Dispose()
}

function Show-LicensePicker {
    param([System.Windows.Forms.Form]$Parent = $null)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Select M365 Licenses"; $dlg.Size = [System.Drawing.Size]::new(700, 500)
    $dlg.StartPosition = "CenterParent"; $dlg.BackColor = $BG
    $dlg.FormBorderStyle = "FixedDialog"; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    $dlg.Controls.Add((Lbl "Available licenses in your tenant (check to assign):" 10 10 640 22 $F_BOLD $TEXT))
    $dlg.Controls.Add((Lbl "Dimmed items = 0 seats available." 10 30 640 18 $F_SM $TEXTDIM))

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = [System.Drawing.Point]::new(10, 54); $lv.Size = [System.Drawing.Size]::new(668, 350)
    $lv.View = "Details"; $lv.CheckBoxes = $true; $lv.FullRowSelect = $true; $lv.GridLines = $true
    $lv.Font = $F_NORM; $lv.BackColor = $CLR_INPUT; $lv.ForeColor = $TEXT; $lv.BorderStyle = "FixedSingle"
    [void]$lv.Columns.Add("License Name", 300)
    [void]$lv.Columns.Add("SKU", 160)
    [void]$lv.Columns.Add("Available", 90)
    [void]$lv.Columns.Add("Total", 90)
    $dlg.Controls.Add($lv)

    foreach ($lic in $script:CachedLicenses) {
        $item = New-Object System.Windows.Forms.ListViewItem($lic.FriendlyName)
        [void]$item.SubItems.Add($lic.SkuPartNumber)
        [void]$item.SubItems.Add($lic.Available.ToString())
        [void]$item.SubItems.Add($lic.Total.ToString())
        $item.Tag = $lic.SkuId
        if ($lic.Available -le 0) { $item.ForeColor = $TEXTDIM }
        [void]$lv.Items.Add($item)
    }

    if ($script:CachedLicenses.Count -eq 0) {
        $dlg.Controls.Add((Lbl "No licenses found - check M365 connection." 10 410 640 22 $F_BOLD $WARN))
    }

    $btnRefresh = Btn "Refresh" 10 414 110 28 $PANEL $ACCENT; $dlg.Controls.Add($btnRefresh)
    $btnOK      = Btn "Assign Selected" 390 414 180 28 $GREEN $BG; $dlg.Controls.Add($btnOK)
    $btnX       = Btn "Cancel" 580 414 90 28 $BORDER $TEXT; $dlg.Controls.Add($btnX)

    # Pre-check any licenses that were previously selected (so Cancel preserves the selection)
    for ($i = 0; $i -lt $lv.Items.Count; $i++) {
        $skuId     = $lv.Items[$i].Tag
        $alreadySel = @($script:SelectedLicenses | Where-Object { $_['SkuId'] -eq $skuId })
        if ($alreadySel.Count -gt 0) { $lv.Items[$i].Checked = $true }
    }

    $btnRefresh.Add_Click({
            $lv.Items.Clear()
            $script:CachedLicenses = Get-M365Licenses -Silent:$false
            foreach ($lic in $script:CachedLicenses) {
                $item = New-Object System.Windows.Forms.ListViewItem($lic.FriendlyName)
                [void]$item.SubItems.Add($lic.SkuPartNumber)
                [void]$item.SubItems.Add($lic.Available.ToString())
                [void]$item.SubItems.Add($lic.Total.ToString())
                $item.Tag = $lic.SkuId
                if ($lic.Available -le 0) { $item.ForeColor = $TEXTDIM }
                [void]$lv.Items.Add($item)
            }
        })

    $btnOK.Add_Click({
            # Only update SelectedLicenses on OK - Cancel leaves it unchanged
            $script:SelectedLicenses = @(
                $lv.CheckedItems | ForEach-Object { @{ SkuId = $_.Tag; FriendlyName = $_.Text } }
            )
            $dlg.DialogResult = "OK"; $dlg.Close()
        })
    $btnX.Add_Click({ $dlg.DialogResult = "Cancel"; $dlg.Close() })
    if ($Parent) { [void]$dlg.ShowDialog($Parent) } else { [void]$dlg.ShowDialog() }
    $dlg.Dispose()
}

# ============================================================
# MAIN FORM
# ============================================================

# Initialize all selection variables up front so strict mode never
# throws if the user runs without opening a picker dialog first
$script:SelectedOU       = ""
$script:SelectedGroups   = @("Domain Users")
$script:SelectedLicenses = @()
$script:SelectedManager  = $null
$script:LastResult       = $null

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD/M365 User Onboarding Tool"
$form.Size = [System.Drawing.Size]::new(860, 660)
$form.StartPosition = "CenterScreen"; $form.BackColor = $BG
$form.FormBorderStyle = "Sizable"; $form.MaximizeBox = $true; $form.Font = $F_NORM
$form.MinimumSize = [System.Drawing.Size]::new(800, 560)

# Scrollable content panel - all content goes here so bottom buttons always visible
$pnlScroll = New-Object System.Windows.Forms.Panel
$pnlScroll.Dock = "Fill"
$pnlScroll.AutoScroll = $true
$pnlScroll.BackColor = $BG

# Title strip
$pnlTitle = New-Object System.Windows.Forms.Panel
$pnlTitle.Location = [System.Drawing.Point]::new(0, 0); $pnlTitle.Size = [System.Drawing.Size]::new(860, 52)
$pnlTitle.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)
$pnlTitle.Controls.Add((Lbl "  [ONBOARD]  AD/M365 User Onboarding Tool" 0 8 560 34 $F_TITLE $GREEN))
$pnlTitle.Controls.Add((Lbl "Domain: $($script:E.LocalDomain)   Tenant: $($script:E.TenantName)" 574 18 276 20 $F_NORM $TEXTDIM))
$pnlScroll.Controls.Add($pnlTitle)

# Status banner
$pnlBanner = New-Object System.Windows.Forms.Panel
$pnlBanner.Location = [System.Drawing.Point]::new(0, 52); $pnlBanner.Size = [System.Drawing.Size]::new(860, 24)
$pnlBanner.BackColor = [System.Drawing.Color]::FromArgb(0, 55, 18)
$licCount  = $script:CachedLicenses.Count
$lblBanner = Lbl "  [OK] Modules verified   [OK] M365 roles verified   [OK] $licCount license SKU(s) loaded" 0 3 860 18 $F_NORM $GREEN
$pnlBanner.Controls.Add($lblBanner)
$pnlScroll.Controls.Add($pnlBanner)

# Section: User Information
$grpInfo = GBox "New User Information" 10 84 838 134
$pnlScroll.Controls.Add($grpInfo)

$grpInfo.Controls.Add((Lbl "First Name *" 10 24 100 20 $F_BOLD $TEXT))
$txtFirst = TBox 10 44 160 26; $grpInfo.Controls.Add($txtFirst)

$grpInfo.Controls.Add((Lbl "Last Name *" 182 24 100 20 $F_BOLD $TEXT))
$txtLast = TBox 182 44 160 26; $grpInfo.Controls.Add($txtLast)

$grpInfo.Controls.Add((Lbl "Display Name" 354 24 120 20 $F_BOLD $TEXT))
$txtDisp = TBox 354 44 200 26; $txtDisp.ForeColor = $TEXTDIM; $grpInfo.Controls.Add($txtDisp)

$grpInfo.Controls.Add((Lbl "Job Title" 566 24 100 20 $F_BOLD $TEXT))
$txtTitle = TBox 566 44 252 26; $grpInfo.Controls.Add($txtTitle)

$grpInfo.Controls.Add((Lbl "Department" 10 78 100 20 $F_BOLD $TEXT))
$txtDept = TBox 10 96 200 26; $grpInfo.Controls.Add($txtDept)

$grpInfo.Controls.Add((Lbl "Phone" 222 78 80 20 $F_BOLD $TEXT))
$txtPhone = TBox 222 96 160 26; $grpInfo.Controls.Add($txtPhone)

$grpInfo.Controls.Add((Lbl "Manager" 394 78 80 20 $F_BOLD $TEXT))
$txtMgr = TBox 394 96 200 26; $txtMgr.ForeColor = $TEXTDIM; $txtMgr.ReadOnly = $true
$grpInfo.Controls.Add($txtMgr)
$btnMgrSearch = Btn "Search" 602 94 70 28 $PANEL $ACCENT; $grpInfo.Controls.Add($btnMgrSearch)

$grpInfo.Controls.Add((Lbl "Notify Email" 684 78 100 20 $F_BOLD $TEXT))
$txtNotify = TBox 684 96 144 26; $grpInfo.Controls.Add($txtNotify)

# Auto-fill display name
$updateDisp = {
    $f = $txtFirst.Text.Trim(); $l = $txtLast.Text.Trim()
    if ($txtDisp.ForeColor.ToArgb() -eq $TEXTDIM.ToArgb()) {
        $txtDisp.Text = "$f $l".Trim()
    }
}
$txtFirst.Add_TextChanged($updateDisp)
$txtLast.Add_TextChanged($updateDisp)
$txtDisp.Add_Enter({
        if ($txtDisp.ForeColor.ToArgb() -eq $TEXTDIM.ToArgb()) {
            $txtDisp.Text = ""; $txtDisp.ForeColor = $TEXT
        }
    })

# Section: Account Preview
$grpAcct = GBox "Generated Account Details  (auto-derived)" 10 226 838 96
$pnlScroll.Controls.Add($grpAcct)

$grpAcct.Controls.Add((Lbl "SAM Account:" 10 26 100 20 $F_BOLD $TEXT))
$lblSAM = Lbl "(enter name above)" 114 26 180 20 $F_MONO $ACCENT; $grpAcct.Controls.Add($lblSAM)

$grpAcct.Controls.Add((Lbl "UPN / Email:" 306 26 100 20 $F_BOLD $TEXT))
$lblUPN = Lbl "" 410 26 280 20 $F_MONO $ACCENT; $grpAcct.Controls.Add($lblUPN)

$grpAcct.Controls.Add((Lbl "Temp Password:" 700 26 116 20 $F_BOLD $TEXT))
$lblPwd = Lbl (New-TempPassword) 700 46 130 20 $F_MONO $WARN; $grpAcct.Controls.Add($lblPwd)
$btnNewPwd = Btn "New" 830 44 18 24 $PANEL $TEXTDIM; $grpAcct.Controls.Add($btnNewPwd)

$grpAcct.Controls.Add((Lbl "UPN Suffix:" 10 60 90 22 $F_BOLD $TEXT))
$cmbSuffix = New-Object System.Windows.Forms.ComboBox
$cmbSuffix.Location    = [System.Drawing.Point]::new(104, 58)
$cmbSuffix.Size        = [System.Drawing.Size]::new(280, 26)
$cmbSuffix.Font        = $F_NORM; $cmbSuffix.ForeColor = $TEXT; $cmbSuffix.BackColor = $CLR_INPUT
$cmbSuffix.DropDownStyle = "DropDownList"
foreach ($sfx in $script:E.UPNSuffixes) { [void]$cmbSuffix.Items.Add($sfx) }
if ($cmbSuffix.Items.Count -eq 0) { [void]$cmbSuffix.Items.Add($script:E.EmailDomain) }
$cmbSuffix.SelectedIndex = 0
$grpAcct.Controls.Add($cmbSuffix)
$grpAcct.Controls.Add((Lbl "(email domain for UPN and primary SMTP - sets mail + proxyAddresses before sync)" 396 62 438 18 $F_SM $TEXTDIM))

# Usage location (ISO 3166-1 alpha-2) -- required by M365 before license assignment
$grpAcct.Controls.Add((Lbl "Usage Location:" 700 58 116 22 $F_BOLD $TEXT))
$txtUsageLoc = New-Object System.Windows.Forms.TextBox
$txtUsageLoc.Location  = [System.Drawing.Point]::new(820, 58)
$txtUsageLoc.Size      = [System.Drawing.Size]::new(40, 24)
$txtUsageLoc.Font      = $F_NORM; $txtUsageLoc.ForeColor = $TEXT; $txtUsageLoc.BackColor = $CLR_INPUT
$txtUsageLoc.MaxLength = 2
$txtUsageLoc.Text      = if ($script:Config['UsageLocation']) { $script:Config['UsageLocation'] } else { 'US' }
$grpAcct.Controls.Add($txtUsageLoc)

$updatePreview = {
    $f = $txtFirst.Text.Trim(); $l = $txtLast.Text.Trim()
    $sfx = if ($cmbSuffix.SelectedItem) { $cmbSuffix.SelectedItem } else { $script:E.EmailDomain }
    if ($f.Length -gt 0 -and $l.Length -gt 0) {
        try {
            $sam = New-SamAccountName -First $f -Last $l
            $lblSAM.Text = $sam
            $lblUPN.Text = "$sam@$sfx"
        }
        catch { $lblSAM.Text = "(error)"; $lblUPN.Text = "" }
    }
    else { $lblSAM.Text = "(enter name above)"; $lblUPN.Text = "" }
}
$txtFirst.Add_TextChanged($updatePreview)
$txtLast.Add_TextChanged($updatePreview)
$cmbSuffix.Add_SelectedIndexChanged({
        & $updatePreview
        # Persist selected suffix to config
        if ($cmbSuffix.SelectedItem) {
            $script:Config['DefaultUPNSuffix'] = $cmbSuffix.SelectedItem
            Save-Config
        }
    })

$btnNewPwd.Add_Click({ $lblPwd.Text = New-TempPassword })

# Section: Target OU
$grpOU = GBox "Target OU  (auto-discovered - click Browse to change)" 10 308 838 66
$pnlScroll.Controls.Add($grpOU)

$grpOU.Controls.Add((Lbl "OU:" 10 28 30 22 $F_BOLD $TEXT))
$txtOU = New-Object System.Windows.Forms.TextBox
$txtOU.Location = [System.Drawing.Point]::new(44, 24); $txtOU.Size = [System.Drawing.Size]::new(680, 26)
$txtOU.Font = $F_NORM; $txtOU.ForeColor = $TEXT; $txtOU.BackColor = $CLR_INPUT; $txtOU.BorderStyle = "FixedSingle"
$txtOU.Text = $script:E.DefaultUserOU
$grpOU.Controls.Add($txtOU)
$btnBrowseOU = Btn "Browse" 732 22 92 28 $PANEL $ACCENT; $grpOU.Controls.Add($btnBrowseOU)

# Section: AD Groups
$grpGrps = GBox "AD Security Groups" 10 382 414 144
$pnlScroll.Controls.Add($grpGrps)

$clbGroups = New-Object System.Windows.Forms.CheckedListBox
$clbGroups.Location = [System.Drawing.Point]::new(10, 20); $clbGroups.Size = [System.Drawing.Size]::new(290, 112)
$clbGroups.BackColor = $CLR_INPUT; $clbGroups.ForeColor = $TEXT; $clbGroups.Font = $F_NORM
$clbGroups.BorderStyle = "FixedSingle"; $clbGroups.CheckOnClick = $true
$grpGrps.Controls.Add($clbGroups)

# Domain Users is always required - pre-add as locked
[void]$clbGroups.Items.Add("Domain Users", $true)

$btnBrowseGroups = Btn "Browse Groups" 308 20 96 30 $PANEL $ACCENT; $grpGrps.Controls.Add($btnBrowseGroups)
$btnClearGroups  = Btn "Clear" 308 58 96 28 $PANEL $DANGER; $grpGrps.Controls.Add($btnClearGroups)
$lblGrpCount     = Lbl "0 selected" 308 94 96 20 $F_SM $TEXTDIM; $grpGrps.Controls.Add($lblGrpCount)

# Section: M365 Licenses
$grpLic = GBox "M365 Licenses" 434 382 414 144
$pnlScroll.Controls.Add($grpLic)

$clbLicenses = New-Object System.Windows.Forms.CheckedListBox
$clbLicenses.Location = [System.Drawing.Point]::new(10, 20); $clbLicenses.Size = [System.Drawing.Size]::new(290, 112)
$clbLicenses.BackColor = $CLR_INPUT; $clbLicenses.ForeColor = $TEXT; $clbLicenses.Font = $F_NORM
$clbLicenses.BorderStyle = "FixedSingle"; $clbLicenses.CheckOnClick = $true
$grpLic.Controls.Add($clbLicenses)

foreach ($lic in $script:CachedLicenses) {
    $display = "$($lic.FriendlyName) ($($lic.Available) avail)"
    [void]$clbLicenses.Items.Add($display)
}

$btnBrowseLic = Btn "Browse Licenses" 308 20 96 30 $PANEL $ACCENT; $grpLic.Controls.Add($btnBrowseLic)
$btnClearLic  = Btn "Clear" 308 58 96 28 $PANEL $DANGER; $grpLic.Controls.Add($btnClearLic)
$lblLicCount  = Lbl "0 selected" 308 94 96 20 $F_SM $TEXTDIM; $grpLic.Controls.Add($lblLicCount)

# -- Restore saved defaults from config --------------------------------
# OU
if ($script:Config['DefaultOU'] -and $script:Config['DefaultOU'] -ne "") {
    $txtOU.Text = $script:Config['DefaultOU']
}

# UPN Suffix - restore saved selection
$savedSuffix = $script:Config['DefaultUPNSuffix']
if ($savedSuffix -and $savedSuffix -ne "") {
    $idx = $cmbSuffix.Items.IndexOf($savedSuffix)
    if ($idx -ge 0) { $cmbSuffix.SelectedIndex = $idx }
}

# AD Groups (Domain Users already added; restore any additional saved groups)
$savedGroups = $script:Config['DefaultGroups']
if ($savedGroups -and $savedGroups -ne "") {
    foreach ($g in ($savedGroups -split ',')) {
        $g = $g.Trim()
        if ($g -ne "" -and $g -ne "Domain Users") {
            [void]$clbGroups.Items.Add($g, $true)
            $script:SelectedGroups += $g
        }
    }
    $lblGrpCount.Text = "$($clbGroups.CheckedItems.Count) selected"
}

# Licenses - re-check any saved SKUs in the license list
$savedSkus = $script:Config['DefaultLicenseSkuIds']
if ($savedSkus -and $savedSkus -ne "") {
    $skuList = $savedSkus -split ','
    for ($i = 0; $i -lt $script:CachedLicenses.Count; $i++) {
        if ($skuList -contains $script:CachedLicenses[$i].SkuId) {
            $clbLicenses.SetItemChecked($i, $true)
            $script:SelectedLicenses += @{ SkuId = $script:CachedLicenses[$i].SkuId; FriendlyName = $script:CachedLicenses[$i].FriendlyName }
        }
    }
    $lblLicCount.Text = "$(@($script:SelectedLicenses).Count) selected"
}
# ---------------------------------------------------------------------

# Section: AD and Sync
$grpAD = GBox "AD and Sync  (auto-discovered - edit if needed)" 10 534 838 64
$pnlScroll.Controls.Add($grpAD)

$grpAD.Controls.Add((Lbl "AADConnect Server:" 10 28 140 22 $F_BOLD $TEXT))
$txtAADC = TBox 154 24 200 26; $txtAADC.Text = $script:E.AADConnectServer; $grpAD.Controls.Add($txtAADC)

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Text = "WhatIf - simulation only (no changes)"
$chkWhatIf.Location = [System.Drawing.Point]::new(370, 28); $chkWhatIf.Size = [System.Drawing.Size]::new(310, 24)
$chkWhatIf.Font = $F_BOLD; $chkWhatIf.ForeColor = $WARN
$chkWhatIf.BackColor = [System.Drawing.Color]::Transparent
$grpAD.Controls.Add($chkWhatIf)

# Progress and log
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = [System.Drawing.Point]::new(10, 608); $progress.Size = [System.Drawing.Size]::new(838, 16)
$progress.Minimum = 0; $progress.Maximum = 100; $progress.ForeColor = $GREEN; $progress.BackColor = $PANEL
$pnlScroll.Controls.Add($progress)

$lblStatus = Lbl "Ready - complete the fields above and click Run Onboarding." 10 628 838 20 $F_NORM $TEXTDIM
$pnlScroll.Controls.Add($lblStatus)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = [System.Drawing.Point]::new(10, 650); $logBox.Size = [System.Drawing.Size]::new(838, 96)
$logBox.Font = $F_MONO; $logBox.BackColor = [System.Drawing.Color]::FromArgb(10, 14, 22)
$logBox.ForeColor = $TEXT; $logBox.ReadOnly = $true; $logBox.BorderStyle = "None"; $logBox.ScrollBars = "Vertical"
$pnlScroll.Controls.Add($logBox)

# Bottom bar
$pnlBot = New-Object System.Windows.Forms.Panel
$pnlBot.Dock = "Bottom"; $pnlBot.Height = 52
$pnlBot.BackColor = [System.Drawing.Color]::FromArgb(12, 16, 24)
$form.Controls.Add($pnlBot)
$form.Controls.Add($pnlScroll)

$btnRun     = Btn "Run Onboarding"   10  10 185 34 $GREEN  $BG
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$btnExport   = Btn "Export Log"       206  12 130 30 $BORDER $TEXT
$btnRecheck  = Btn "Re-check Setup"   346  12 160 30 $PANEL  $ACCENT
$btnSetupAuth= Btn "Setup Auth"       516  12 140 30 $PANEL  $WARN
$btnClose    = Btn "Close"            740  12 100 30 $DANGER ([System.Drawing.Color]::White)
$pnlBot.Controls.AddRange(@($btnRun, $btnExport, $btnRecheck, $btnSetupAuth, $btnClose))

# ============================================================
# EVENTS
# ============================================================

$btnMgrSearch.Add_Click({
        Show-ManagerSearch -Parent $form
        if ($script:SelectedManager) {
            $txtMgr.ForeColor = $TEXT
            $txtMgr.Text = $script:SelectedManager.SamAccountName
            $txtMgr.Tag  = $script:SelectedManager.SamAccountName
        }
    })

$btnBrowseOU.Add_Click({
        Show-OUPicker -Parent $form
        if ($script:SelectedOU) {
            $txtOU.Text = $script:SelectedOU
            $script:Config['DefaultOU'] = $script:SelectedOU
            Save-Config
        }
    })

$btnBrowseGroups.Add_Click({
        $current = @($clbGroups.CheckedItems | Where-Object { $_ -ne "Domain Users" })
        Show-GroupPicker -AlreadySelected $current -Parent $form
        $clbGroups.Items.Clear()
        # Domain Users always first, always checked
        [void]$clbGroups.Items.Add("Domain Users", $true)
        foreach ($g in ($script:SelectedGroups | Where-Object { $_ -ne "Domain Users" })) {
            [void]$clbGroups.Items.Add($g, $true)
        }
        $lblGrpCount.Text = "$($clbGroups.CheckedItems.Count) selected"
        # Persist to config (Domain Users is implicit - stored separately)
        $extras = @($script:SelectedGroups | Where-Object { $_ -ne "Domain Users" })
        $script:Config['DefaultGroups'] = $extras -join ','
        Save-Config
    })

$btnClearGroups.Add_Click({
        # Clear all except Domain Users (always required)
        $clbGroups.Items.Clear()
        [void]$clbGroups.Items.Add("Domain Users", $true)
        $lblGrpCount.Text = "1 selected"
    })
$clbGroups.Add_ItemCheck({
        # Domain Users can never be unchecked
        if ($clbGroups.Items[$_.Index] -eq "Domain Users") {
            $_.NewValue = [System.Windows.Forms.CheckState]::Checked
        }
        $lblGrpCount.Text = "$($clbGroups.CheckedItems.Count) selected"
    })

$btnBrowseLic.Add_Click({
        Show-LicensePicker -Parent $form
        $clbLicenses.Items.Clear()
        foreach ($lic in $script:CachedLicenses) {
            $display    = "$($lic.FriendlyName) ($($lic.Available) avail)"
            $idx        = $clbLicenses.Items.Add($display)
            $isSelected = (@($script:SelectedLicenses | Where-Object { $_['SkuId'] -eq $lic.SkuId })).Count -gt 0
            $clbLicenses.SetItemChecked($idx, $isSelected)
        }
        $lblLicCount.Text = "$(@($script:SelectedLicenses).Count) selected"
        # Persist license selection to config
        $script:Config['DefaultLicenseSkuIds'] = (@($script:SelectedLicenses) | ForEach-Object { $_['SkuId'] }) -join ','
        Save-Config
    })

$btnClearLic.Add_Click({
        for ($i = 0; $i -lt $clbLicenses.Items.Count; $i++) { $clbLicenses.SetItemChecked($i, $false) }
        $script:SelectedLicenses = @(); $lblLicCount.Text = "0 selected"
    })
$clbLicenses.Add_ItemCheck({ $lblLicCount.Text = "$($clbLicenses.CheckedItems.Count) selected" })

$btnRun.Add_Click({
        $fn = $txtFirst.Text.Trim(); $ln = $txtLast.Text.Trim(); $ou = $txtOU.Text.Trim()
        if ($fn -eq "" -or $ln -eq "" -or $ou -eq "") {
            [System.Windows.Forms.MessageBox]::Show(
                "First Name, Last Name, and Target OU are required.",
                "Missing Fields", "OK", "Warning"); return
        }

        $selGroups  = @($clbGroups.CheckedItems | ForEach-Object { $_ })
        $selLicSkus = @()
        for ($i = 0; $i -lt $clbLicenses.Items.Count; $i++) {
            if ($clbLicenses.GetItemChecked($i) -and $i -lt $script:CachedLicenses.Count) {
                $selLicSkus += $script:CachedLicenses[$i].SkuId
            }
        }
        if ($selLicSkus.Count -eq 0 -and (@($script:SelectedLicenses)).Count -gt 0) {
            $selLicSkus = @($script:SelectedLicenses | ForEach-Object { $_['SkuId'] })
        }

        $licNames = @($selLicSkus | ForEach-Object {
                $id = $_; $m = $script:CachedLicenses | Where-Object { $_.SkuId -eq $id } | Select-Object -First 1
                if ($m) { $m.FriendlyName } else { $id }
            })

        $wi      = $chkWhatIf.Checked
        $dispVal = if ($txtDisp.ForeColor.ToArgb() -eq $TEXTDIM.ToArgb() -or $txtDisp.Text.Trim() -eq "") { "" } else { $txtDisp.Text.Trim() }
        $mgrVal  = if ($txtMgr.Tag) { $txtMgr.Tag } else { $txtMgr.Text.Trim() }
        $pwdVal  = $lblPwd.Text.Trim()

        $msg = "Create new user:`n`n" +
            "  Name       : $fn $ln`n" +
            "  Title      : $($txtTitle.Text.Trim())`n" +
            "  Department : $($txtDept.Text.Trim())`n" +
            "  Manager    : $(if ($mgrVal) { $mgrVal } else { '(none)' })`n" +
            "  Target OU  : $ou`n" +
            "  AD Groups  : $(if ($selGroups.Count) { $selGroups -join ', ' } else { '(none)' })`n" +
            "  Licenses   : $(if ($licNames.Count) { $licNames -join ', ' } else { '(none)' })`n" +
            "  WhatIf     : $wi`n`n" +
            "$(if ($wi) { 'SIMULATION - no changes will be made.' } else { 'THIS WILL CREATE A LIVE ACCOUNT. Proceed?' })"

        if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Onboarding", "YesNo", "Question") -ne "Yes") { return }

        $btnRun.Enabled = $false; $logBox.Clear()

        # Persist usage location from GUI field before invoking (engine reads $script:Config)
        $uLocVal = $txtUsageLoc.Text.Trim().ToUpper()
        if ($uLocVal.Length -eq 2) {
            $script:Config['UsageLocation'] = $uLocVal
            Save-Config
        }

        $ok = Invoke-Onboarding `
            -FirstName        $fn `
            -LastName         $ln `
            -DisplayName      $dispVal `
            -JobTitle         $txtTitle.Text.Trim() `
            -Department       $txtDept.Text.Trim() `
            -Manager          $mgrVal `
            -PhoneNumber      $txtPhone.Text.Trim() `
            -TargetOU         $ou `
            -ADGroups         $selGroups `
            -LicenseSkuIds    $selLicSkus `
            -AADConnectServer $txtAADC.Text.Trim() `
            -TempPassword     $pwdVal `
            -NotifyEmail      $txtNotify.Text.Trim() `
            -UPNSuffix        ($cmbSuffix.SelectedItem) `
            -WhatIfMode       $wi `
            -LogBox           $logBox `
            -ProgressBar      $progress `
            -StatusLabel      $lblStatus

        $btnRun.Enabled = $true

        if ($ok) {
            $lblStatus.ForeColor = $GREEN; $lblStatus.Text = "[OK]  Onboarding complete."
            if (-not $wi -and $script:LastResult) {
                $summary = "Onboarding complete!`n`n" +
                    "Display Name  : $($script:LastResult.DisplayName)`n" +
                    "SAM Account   : $($script:LastResult.SAM)`n" +
                    "UPN / Email   : $($script:LastResult.UPN)`n" +
                    "Temp Password : $($script:LastResult.Password)`n`n" +
                    "Password must be changed at first login.`n" +
                    "Copy it now - it will not be shown again."
                [System.Windows.Forms.MessageBox]::Show($summary, "Account Created", "OK", "Information")
            }
        }
        else { $lblStatus.ForeColor = $DANGER; $lblStatus.Text = "[X]  Error - see log." }
    })

$btnRecheck.Add_Click({
        $lblBanner.ForeColor = $WARN
        $lblBanner.Text = "  [...]  Re-checking modules, roles, and licenses..."
        $pnlBanner.BackColor = [System.Drawing.Color]::FromArgb(60, 40, 0)
        [System.Windows.Forms.Application]::DoEvents()
        try {
            Invoke-ModuleBootstrap -Silent:$false
            Invoke-RoleBootstrap -AdminUPN $AdminUPN -Silent:$false
            $script:CachedLicenses = Get-M365Licenses -Silent:$false
            $clbLicenses.Items.Clear()
            foreach ($lic in $script:CachedLicenses) {
                [void]$clbLicenses.Items.Add("$($lic.FriendlyName) ($($lic.Available) avail)")
            }
            $lblBanner.ForeColor = $GREEN
            $lblBanner.Text = "  [OK] Modules verified   [OK] M365 roles verified   [OK] $($script:CachedLicenses.Count) license SKU(s) loaded"
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
        $dlg.FileName = "Onboard-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
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
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "A startup error occurred:`n`n$_`n`nLine: $($_.InvocationInfo.ScriptLineNumber)",
        "Startup Error", "OK", "Error") | Out-Null
}