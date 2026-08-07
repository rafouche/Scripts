#Requires -Version 5.1
<#
================================================================================
  COPYRIGHT NOTICE AND LICENSE AGREEMENT
================================================================================

  Script:    Set-ADEntraHardMatch.ps1
  Author:    Roger Fouche
  Entity:    Fouche Enterprises, LLC
  Created:   2025

  Copyright (c) 2025 Roger Fouche / Fouche Enterprises, LLC.
  All rights reserved.

  This script and all associated intellectual property were developed by
  Roger Fouche independently, on personal time, using personal resources
  and accounts. This work is NOT a work-for-hire and is NOT the property
  of any employer or client.

  LICENSE GRANT:
  A revocable, non-exclusive, non-transferable license is granted to
  Altec Solutions Group, Inc. ("Licensee") to use this script solely
  for internal business operations, subject to the following conditions:

    1. LICENSE TERMINATION: This license terminates automatically and
       immediately upon the departure (voluntary or involuntary) of
       Roger Fouche from employment or contracted engagement with
       Altec Solutions Group, Inc.

    2. POST-TERMINATION USE: Any continued use of this script by
       Altec Solutions Group, Inc. after license termination requires
       a separate written license agreement with Roger Fouche /
       Fouche Enterprises, LLC. Unauthorized continued use constitutes
       copyright infringement.

    3. NO TRANSFER: This license may not be sublicensed, sold,
       transferred, or assigned to any third party without the express
       written consent of Roger Fouche / Fouche Enterprises, LLC.

    4. NO WARRANTY: This script is provided "as is" without warranty
       of any kind. Roger Fouche / Fouche Enterprises, LLC shall not
       be liable for any damages arising from its use.

    5. ATTRIBUTION: This copyright notice must be retained in all
       copies or modified versions of this script.

  LICENSING INQUIRIES:
  For licensing arrangements, contact: roger@fouche.dev

================================================================================

.SYNOPSIS
    AD to Entra ID Hard Match Tool
    Single user OR batch OU scan with checkbox selection.

.DESCRIPTION
    Reads credentials from ADM365Config.json (same file as Onboard/Offboard tools).
    Cert-based Graph auth -- no interactive sign-in prompt after first setup.

    MODES
    -----
    Single   : Search AD for one user, search Entra for one user, apply match.
    Batch    : Pick an OU (recurses sub-OUs), load all AD users, auto-resolve each
               to an Entra candidate, present a checked list -- select which ones
               to hard-match and apply in one pass.

    WHAT IS A HARD MATCH?
    ---------------------
    Sets OnPremisesImmutableId on the Entra user to Base64(ADUser.ObjectGUID).
    AAD Connect uses this value to bind the on-prem account to the cloud account
    permanently, even if UPNs or display names differ.

    WHATIF MODE
    -----------
    Pass -WhatIf to log every intended write without touching anything.

.PARAMETER SamAccountName
    CLI/Ninja: SAM account name of the AD user to hard-match (single-user mode).

.PARAMETER EntraUPN
    CLI/Ninja: Entra UPN to match to. If omitted, auto-resolved by UPN/mail match.

.PARAMETER OUPath
    CLI/Ninja: DistinguishedName of OU to batch-scan. Requires -BatchMode.

.PARAMETER BatchMode
    CLI/Ninja: Enable batch OU scan mode. Requires -OUPath.

.PARAMETER WhatIf
    Simulate all operations -- no changes written to Entra.

.PARAMETER Silent
    Suppress GUI and run headless (auto-set when CLI params are provided).

.PARAMETER SkipModuleCheck
    Skip module install/update check (faster on repeat runs).

.NOTES
    Author  : Roger Fouche / Fouche Enterprises, LLC
    Config  : ADM365Config.json  (same folder as this script)
    Modules : ActiveDirectory, Microsoft.Graph.Authentication,
              Microsoft.Graph.Users
#>

[CmdletBinding()]
param(
    # --- CLI / Ninja mode ---
    # Single-user hard match (non-interactive):
    #   -SamAccountName  AD SAM of the user to match
    #   -EntraUPN        Entra UPN to match to (if omitted, auto-resolved by UPN/mail)
    # Batch mode (non-interactive):
    #   -OUPath          DistinguishedName of OU to scan (recurses sub-OUs)
    #   -BatchMode       Must also be set for CLI batch
    # Common:
    #   -WhatIf          Log all intended changes but write nothing
    #   -Silent          Suppress GUI -- log to console only (auto-set when CLI params present)
    #   -SkipModuleCheck Skip module install check
    [string]$SamAccountName    = '',
    [string]$EntraUPN          = '',
    [string]$OUPath            = '',
    [switch]$BatchMode,
    [switch]$WhatIf,
    [switch]$Silent,
    [switch]$SkipModuleCheck
)

# ---------------------------------------------------------------
# BOOTSTRAP 1 -- ADMIN ELEVATION
# Must run before PS7 check (installing PS7 requires admin).
# ---------------------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host 'Not running as Administrator -- relaunching elevated...' -ForegroundColor Yellow
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    foreach ($key in $MyInvocation.BoundParameters.Keys) {
        $val = $MyInvocation.BoundParameters[$key]
        if     ($val -is [switch] -and $val.IsPresent) { $argList += " -$key" }
        elseif ($val -is [string] -and $val -ne '')    { $argList += " -$key `"$val`"" }
    }
    Start-Process (Get-Process -Id $PID).Path -ArgumentList $argList -Verb RunAs
    exit
}

# ---------------------------------------------------------------
# BOOTSTRAP 2 -- PS7 CHECK / INSTALL / RELAUNCH
# Run BEFORE Set-StrictMode -- bootstrap code needs permissive mode.
# ---------------------------------------------------------------
$pwsh7 = 'C:\Program Files\PowerShell\7\pwsh.exe'
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] PowerShell 7 required. Checking..." -ForegroundColor Cyan

    if (-not (Test-Path -LiteralPath $pwsh7)) {
        # Try winget first
        $winget = Get-Command winget -ErrorAction SilentlyContinue
        if ($winget) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Installing PS7 via winget..." -ForegroundColor Cyan
            try {
                $wg = Start-Process winget -ArgumentList 'install --id Microsoft.PowerShell --silent --accept-source-agreements --accept-package-agreements' `
                    -Wait -PassThru -WindowStyle Hidden
                if ($wg.ExitCode -eq 0) { Write-Host "[$(Get-Date -Format 'HH:mm:ss')] PS7 installed via winget." -ForegroundColor Green }
            } catch { Write-Warning "winget install failed: $_" }
        }

        # Fallback: MSI from GitHub releases API
        if (-not (Test-Path -LiteralPath $pwsh7)) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Trying MSI install from GitHub..." -ForegroundColor Cyan
            try {
                $rel  = Invoke-RestMethod 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest' -EA Stop
                $msi  = ($rel.assets | Where-Object { $_.name -match 'win-x64\.msi$' } | Select-Object -First 1).browser_download_url
                $tmp  = Join-Path $env:TEMP 'pwsh7.msi'
                Invoke-WebRequest $msi -OutFile $tmp -UseBasicParsing -EA Stop
                Start-Process msiexec -ArgumentList "/i `"$tmp`" /quiet /norestart" -Wait
                Remove-Item $tmp -Force -ErrorAction SilentlyContinue
                if (Test-Path -LiteralPath $pwsh7) { Write-Host "[$(Get-Date -Format 'HH:mm:ss')] PS7 installed via MSI." -ForegroundColor Green }
            } catch { Write-Warning "MSI install failed: $_" }
        }

        if (-not (Test-Path -LiteralPath $pwsh7)) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Could not install PS7 automatically." -ForegroundColor Red
            Write-Host "  Please install manually: https://aka.ms/powershell" -ForegroundColor Yellow
            if (-not $Silent -and -not $SamAccountName -and -not $OUPath) {
                Read-Host 'Press Enter to exit'
            }
            exit 1
        }
    }

    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Relaunching in PowerShell 7..." -ForegroundColor Cyan
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    foreach ($key in $MyInvocation.BoundParameters.Keys) {
        $val = $MyInvocation.BoundParameters[$key]
        if     ($val -is [switch] -and $val.IsPresent) { $argList += " -$key" }
        elseif ($val -is [string] -and $val -ne '')    { $argList += " -$key `"$val`"" }
    }
    Start-Process $pwsh7 -ArgumentList $argList -Verb RunAs -Wait
    exit
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Determine run mode
$script:IsCLI    = ($SamAccountName -ne '' -or ($OUPath -ne '' -and $BatchMode.IsPresent))
$script:IsSilent = $Silent.IsPresent -or $script:IsCLI

# ---------------------------------------------------------------
# 0A  ASSEMBLIES
# ---------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------
# 0B  COLOUR / FONT PALETTE  (matches Onboard / Offboard tools)
# ---------------------------------------------------------------
$C = @{
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
$F = @{
    UI    = [System.Drawing.Font]::new('Segoe UI',  9)
    UIB   = [System.Drawing.Font]::new('Segoe UI',  9, [System.Drawing.FontStyle]::Bold)
    Small = [System.Drawing.Font]::new('Segoe UI',  8)
    Mono  = [System.Drawing.Font]::new('Consolas',  8)
    Title = [System.Drawing.Font]::new('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    Hdr   = [System.Drawing.Font]::new('Segoe UI',  9, [System.Drawing.FontStyle]::Bold)
}

# ---------------------------------------------------------------
# 0C  SHARED SCRIPT STATE
# ---------------------------------------------------------------
$script:LogBox    = $null
$script:LblStatus = $null
$script:SelAD     = $null   # single-mode selected AD user
$script:SelEntra  = $null   # single-mode selected Entra user

# ---------------------------------------------------------------
# 0D  CONFIG LOAD
# ---------------------------------------------------------------
$script:ConfigPath = Join-Path (Split-Path $PSCommandPath -Parent) 'ADM365Config.json'
$script:Config = @{
    AdminUPN            = ''
    TenantId            = ''
    AppId               = ''
    CertThumbprint      = ''
    ExchangeOrg         = ''
    DefaultOU           = ''
    DefaultGroups       = ''
    DefaultLicenseSkuIds= ''
}

if (Test-Path $script:ConfigPath) {
    try {
        $loaded = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
        foreach ($key in @('AdminUPN','TenantId','AppId','CertThumbprint','ExchangeOrg',
                           'DefaultOU','DefaultGroups','DefaultLicenseSkuIds')) {
            $prop = $loaded.PSObject.Properties[$key]
            if ($prop -and $prop.Value) { $script:Config[$key] = $prop.Value }
        }
    } catch {
        Write-Warning "Could not parse ADM365Config.json: $_"
    }
} else {
    Write-Warning "ADM365Config.json not found at: $script:ConfigPath -- run Setup Auth (this tool's GUI, or Onboard-ADUser.ps1 / Offboard-ADUser.ps1) to create it."
}

function Save-Config {
    try {
        $script:Config | ConvertTo-Json | Set-Content $script:ConfigPath -Encoding UTF8
    } catch {}
}

function Get-CertConfig {
    # Same shape as Onboard-ADUser.ps1 / Offboard-ADUser.ps1 -- returns the config
    # hashtable if cert auth is fully configured and the cert is still valid.
    $appId = $script:Config['AppId']
    $thumb = $script:Config['CertThumbprint']
    $tid   = $script:Config['TenantId']
    if ($appId -and $thumb -and $tid) {
        $c = Get-Item "Cert:\LocalMachine\My\$thumb" -ErrorAction SilentlyContinue
        if ($c -and $c.NotAfter -gt (Get-Date)) { return $script:Config }
    }
    return $null
}

# ---------------------------------------------------------------
# 0E  HELPERS
# ---------------------------------------------------------------
function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $ts  = Get-Date -Format 'HH:mm:ss'
    $col = switch ($Level) {
        'OK'   { $C.Success }
        'WARN' { $C.Warning }
        'ERR'  { $C.Danger  }
        default{ $C.FGDim   }
    }
    if ($script:LogBox) {
        $script:LogBox.SelectionStart  = $script:LogBox.TextLength
        $script:LogBox.SelectionLength = 0
        $script:LogBox.SelectionColor  = $col
        $script:LogBox.AppendText("[$ts][$Level] $Msg`r`n")
        $script:LogBox.ScrollToCaret()
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
    $b.FlatAppearance.BorderColor = $C.Border
    $b.BackColor = if ($PSBoundParameters.ContainsKey('BG')) { $BG } else { $C.Card }
    $b.ForeColor = $C.FGHdr
    $b.Font      = $F.UIB
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

function New-Txt {
    param([int]$W=260)
    $t = [System.Windows.Forms.TextBox]::new()
    $t.Width       = $W
    $t.Height      = 24
    $t.BackColor   = $C.Card
    $t.ForeColor   = $C.FG
    $t.BorderStyle = 'FixedSingle'
    $t.Font        = $F.UI
    return $t
}

function New-Lbl {
    param([string]$Text,[System.Drawing.Font]$Font,[System.Drawing.Color]$Color)
    $l = [System.Windows.Forms.Label]::new()
    $l.Text      = $Text
    $l.Font      = if ($PSBoundParameters.ContainsKey('Font'))  { $Font  } else { $F.UI }
    $l.ForeColor = if ($PSBoundParameters.ContainsKey('Color')) { $Color } else { $C.FG }
    $l.AutoSize  = $true
    $l.BackColor = [System.Drawing.Color]::Transparent
    return $l
}

function New-Sep {
    param([int]$Y,[int]$W=760)
    $p = [System.Windows.Forms.Panel]::new()
    $p.Height    = 1
    $p.Width     = $W
    $p.BackColor = $C.Border
    $p.Location  = [System.Drawing.Point]::new(0,$Y)
    return $p
}

# ---------------------------------------------------------------
# 0F  MODULE BOOTSTRAP
# ---------------------------------------------------------------
function Invoke-ModuleBootstrap {
    # Matches the bootstrap in Onboard-ADUser.ps1 / Offboard-ADUser.ps1.
    # RSAT install path differs by OS type:
    #   Server OS   -> Add-WindowsFeature -Name RSAT-AD-PowerShell
    #   Workstation -> Add-WindowsCapability (Rsat.ActiveDirectory*)
    # Neither path is fatal -- if RSAT install fails we log and continue
    # so the GUI still launches and shows the error rather than hard-exiting.

    function BL { param($Msg, $Lvl = 'INFO')
        $col = switch ($Lvl) { 'OK' {'Green'} 'WARN' {'Yellow'} 'ERR' {'Red'} default {'Cyan'} }
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$Lvl] $Msg" -ForegroundColor $col
        if ($script:LogBox) { Write-Log $Msg $Lvl }
    }

    BL 'Checking required modules...'

    # -- NuGet provider / PSGallery trust --
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nuget -or $nuget.Version -lt [version]'2.8.5.201') {
        BL 'Installing NuGet provider...'
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 `
                -Force -Scope CurrentUser | Out-Null
            BL 'NuGet installed.' 'OK'
        } catch { BL "NuGet install failed: $_" 'WARN' }
    }
    $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

    # -- ActiveDirectory (RSAT) -- cannot use Install-Module --
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        BL 'ActiveDirectory module missing. Attempting RSAT install...' 'WARN'
        $rsatOk = $false

        # Detect Server vs workstation
        $isServer = $false
        try {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            if (-not $os) { $os = Get-WmiObject Win32_OperatingSystem -ErrorAction SilentlyContinue }
            $isServer = ($os -and $os.Caption -match 'Server')
        } catch {}

        if ($isServer) {
            # Server OS (including DC): Add-WindowsFeature
            BL 'Server OS detected -- using Add-WindowsFeature RSAT-AD-PowerShell...'
            try {
                $feat = Add-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction Stop
                if ($feat.Success) {
                    BL 'RSAT-AD-PowerShell installed.' 'OK'
                    $rsatOk = $true
                    if ($feat.RestartNeeded -eq 'Yes') {
                        BL 'Restart may be required for the feature to activate.' 'WARN'
                    }
                } else {
                    BL 'Add-WindowsFeature reported Success=$false.' 'WARN'
                }
            } catch { BL "Add-WindowsFeature failed: $_" 'WARN' }
        } else {
            # Workstation: Add-WindowsCapability
            BL 'Workstation OS detected -- using Add-WindowsCapability...'
            try {
                $cap = Get-WindowsCapability -Name 'Rsat.ActiveDirectory*' -Online -ErrorAction Stop
                if ($cap -and $cap.State -ne 'Installed') {
                    Add-WindowsCapability -Name $cap.Name -Online -ErrorAction Stop | Out-Null
                    BL 'RSAT installed via Add-WindowsCapability.' 'OK'
                    $rsatOk = $true
                } elseif ($cap -and $cap.State -eq 'Installed') {
                    BL 'RSAT capability already installed.' 'OK'
                    $rsatOk = $true
                }
            } catch { BL "Add-WindowsCapability failed: $_" 'WARN' }

            # Workstation DISM fallback
            if (-not $rsatOk) {
                BL 'Trying DISM fallback...'
                try {
                    $null = & dism.exe /Online /Add-Capability `
                        /CapabilityName:Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 `
                        /NoRestart 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        BL 'RSAT installed via DISM.' 'OK'; $rsatOk = $true
                    } else { BL "DISM exited $LASTEXITCODE." 'WARN' }
                } catch { BL "DISM failed: $_" 'WARN' }
            }
        }

        if (-not $rsatOk) {
            BL 'Automatic RSAT install failed -- continuing. AD features will not work until RSAT is installed.' 'ERR'
            BL '  Server  : Install-WindowsFeature RSAT-AD-PowerShell' 'WARN'
            BL '  Desktop : Settings > Optional Features > RSAT: Active Directory' 'WARN'
        }
    } else {
        BL 'ActiveDirectory module -- OK' 'OK'
    }

    # Import AD (non-fatal if still missing)
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    if (Get-Module -Name ActiveDirectory) {
        BL 'ActiveDirectory loaded.' 'OK'
    } else {
        BL 'ActiveDirectory NOT loaded -- AD features will fail.' 'ERR'
    }

    # -- Graph SDK (PowerShell Gallery) --
    $galleryMods = @(
        @{ Name = 'Microsoft.Graph.Authentication'; Min = '2.0.0' }
        @{ Name = 'Microsoft.Graph.Users';          Min = '2.0.0' }
    )
    foreach ($m in $galleryMods) {
        $inst = Get-Module -ListAvailable -Name $m.Name |
                Sort-Object Version -Descending | Select-Object -First 1
        if (-not $inst) {
            BL "Installing $($m.Name)..."
            try {
                Install-Module $m.Name -Scope AllUsers -Force -AllowClobber -EA Stop
                BL "$($m.Name) installed (AllUsers)." 'OK'
            } catch {
                try {
                    Install-Module $m.Name -Scope CurrentUser -Force -AllowClobber -EA Stop
                    BL "$($m.Name) installed (CurrentUser)." 'OK'
                } catch { BL "Install failed for $($m.Name): $_" 'ERR' }
            }
        } elseif ($inst.Version -lt [version]$m.Min) {
            BL "Updating $($m.Name) v$($inst.Version) to min $($m.Min)..."
            try {
                Update-Module $m.Name -Force -EA Stop
                BL "$($m.Name) updated." 'OK'
            } catch { BL "Update failed for $($m.Name): $_" 'WARN' }
        } else {
            BL "$($m.Name) v$($inst.Version) -- OK" 'OK'
        }
        Import-Module $m.Name -ErrorAction SilentlyContinue
    }

    BL 'Module check complete.' 'OK'
}

# ---------------------------------------------------------------
# 0G  GRAPH CONNECT  (cert-based, same as Onboard/Offboard)
# ---------------------------------------------------------------
function Connect-ToGraph {
    $appId = $script:Config['AppId']
    $tid   = $script:Config['TenantId']
    $thumb = $script:Config['CertThumbprint']
    if (-not $appId -or -not $tid -or -not $thumb) {
        Write-Log 'ADM365Config.json is missing AppId / TenantId / CertThumbprint.' 'ERR'
        return $false
    }

    # Locate the certificate object explicitly -- search LocalMachine first
    # (accessible to all elevated processes), then CurrentUser as fallback.
    # Passing -Certificate avoids the Graph SDK's own store walk which only
    # checks CurrentUser\My and fails when running elevated or as a different
    # token (e.g. after PS7 relaunch, UAC elevation, or Ninja SYSTEM context).
    $cert = $null
    foreach ($store in @('LocalMachine','CurrentUser')) {
        $c = Get-Item "Cert:\$store\My\$thumb" -ErrorAction SilentlyContinue
        if ($c) {
            try {
                $null = $c.PrivateKey   # throws if keyset inaccessible in this context
                $cert = $c
                Write-Log "Certificate found in $store\My store." 'INFO'
                break
            } catch {
                Write-Log "Cert in $store\My -- private key not accessible here (wrong context). Trying next store." 'WARN'
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
        Connect-MgGraph -ClientId $appId -TenantId $tid `
            -Certificate $cert -NoWelcome -EA Stop
        Write-Log 'Graph connected.' 'OK'
        return $true
    } catch {
        Write-Log "Graph connect failed: $(Get-SafeErrorText $_)" 'ERR'
        return $false
    }
}

# ---------------------------------------------------------------
# 0H  ONE-TIME M365 AUTH SETUP
# Identical flow to Onboard-ADUser.ps1 / Offboard-ADUser.ps1: creates a
# shared "ADM365LifecycleTool" app registration + cert and writes
# AppId/TenantId/CertThumbprint back into the shared ADM365Config.json,
# so all three tools can reuse the same non-interactive credential.
# ---------------------------------------------------------------
function Initialize-M365Auth {
    Write-Log 'Starting one-time non-interactive auth setup...' 'INFO'
    Write-Log 'Sign in with a GLOBAL ADMIN account when prompted.' 'WARN'

    try {
        $null = Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $null = Import-Module Microsoft.Graph.Applications -ErrorAction Stop
        $null = Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

        $null = Connect-MgGraph -Scopes @(
            'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All',
            'RoleManagement.ReadWrite.Directory', 'Directory.ReadWrite.All', 'Organization.Read.All'
        ) -NoWelcome -ErrorAction Stop

        $ctx      = Get-MgContext
        $tenantId = $ctx.TenantId
        Write-Log "Connected. Tenant: $tenantId" 'OK'

        Write-Log 'Creating authentication certificate...' 'INFO'
        $appName = 'ADM365LifecycleTool'
        $cert = New-SelfSignedCertificate -Subject "CN=$appName" `
            -CertStoreLocation 'Cert:\LocalMachine\My' `
            -KeyExportPolicy NonExportable -KeySpec Signature `
            -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 `
            -NotAfter (Get-Date).AddYears(2)
        Write-Log "Certificate: $($cert.Thumbprint)" 'OK'

        $old = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($old) { Remove-MgApplication -ApplicationId $old.Id; Write-Log 'Removed existing app.' 'WARN' }

        Write-Log 'Creating Entra ID app registration...' 'INFO'
        $app = New-MgApplication -DisplayName $appName -SignInAudience 'AzureADMyOrg'
        Write-Log "App created: $($app.AppId)" 'OK'

        $null = Update-MgApplication -ApplicationId $app.Id -KeyCredentials @(@{
            Type        = 'AsymmetricX509Cert'
            Usage       = 'Verify'
            Key         = $cert.RawData
            DisplayName = "$appName Cert"
        })
        Write-Log 'Certificate uploaded to app.' 'OK'

        $graphSP = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
        $exoSP   = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'" -ErrorAction SilentlyContinue
        $spoSP   = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" -ErrorAction SilentlyContinue

        $graphPermNames = @('User.ReadWrite.All','Group.ReadWrite.All','Directory.ReadWrite.All','Organization.Read.All','RoleManagement.ReadWrite.Directory')
        $graphRoles = @()
        foreach ($p in $graphPermNames) {
            $r = $graphSP.AppRoles | Where-Object { $_.Value -eq $p } | Select-Object -First 1
            if ($r) { $graphRoles += @{ Id = [string]$r.Id; Type = 'Role' } }
        }
        $reqAccess = @(@{ ResourceAppId = '00000003-0000-0000-c000-000000000000'; ResourceAccess = $graphRoles })

        $exoRoleId = $null
        if ($exoSP) {
            $er = $exoSP.AppRoles | Where-Object { $_.Value -eq 'Exchange.ManageAsApp' } | Select-Object -First 1
            if ($er) {
                $exoRoleId = [string]$er.Id
                $reqAccess += @{ ResourceAppId = '00000002-0000-0ff1-ce00-000000000000'; ResourceAccess = @(@{ Id = $exoRoleId; Type = 'Role' }) }
            }
        }

        $spoRoleId = $null
        if ($spoSP) {
            $sr = $spoSP.AppRoles | Where-Object { $_.Value -eq 'Sites.FullControl.All' } | Select-Object -First 1
            if ($sr) {
                $spoRoleId = [string]$sr.Id
                $reqAccess += @{ ResourceAppId = '00000003-0000-0ff1-ce00-000000000000'; ResourceAccess = @(@{ Id = $spoRoleId; Type = 'Role' }) }
            }
        }

        $null = Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $reqAccess
        Write-Log 'Permissions configured.' 'OK'

        Write-Log 'Creating service principal...' 'INFO'
        $sp = New-MgServicePrincipal -AppId $app.AppId

        Write-Log 'Waiting 15s for propagation, then granting admin consent...' 'INFO'
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
        Write-Log 'Admin consent granted.' 'OK'

        try {
            $exoAdmRole = Get-MgDirectoryRole -Filter "displayName eq 'Exchange Administrator'" -ErrorAction SilentlyContinue
            if (-not $exoAdmRole) {
                $t = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq 'Exchange Administrator' } | Select-Object -First 1
                $exoAdmRole = New-MgDirectoryRole -RoleTemplateId $t.Id
            }
            $null = New-MgDirectoryRoleMemberByRef -DirectoryRoleId $exoAdmRole.Id `
                -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)" }
            Write-Log 'Exchange Administrator role assigned to service principal.' 'OK'
        } catch { Write-Log "Exchange role assignment warning: $_" 'WARN' }

        $org = Get-MgOrganization | Select-Object -First 1
        $exoOrg = ($org.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1).Name
        if (-not $exoOrg) { Write-Log 'Could not determine initial .onmicrosoft.com domain -- ExchangeOrg left blank.' 'WARN' }

        $script:Config['AppId']          = $app.AppId
        $script:Config['TenantId']       = $tenantId
        $script:Config['CertThumbprint'] = $cert.Thumbprint
        $script:Config['ExchangeOrg']    = $exoOrg
        Save-Config

        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue

        Write-Log '============================================' 'OK'
        Write-Log ' SETUP COMPLETE - non-interactive auth ready' 'OK'
        Write-Log " App ID      : $($app.AppId)" 'OK'
        Write-Log " Cert        : $($cert.Thumbprint)" 'OK'
        Write-Log " Exchange Org: $exoOrg" 'OK'
        Write-Log ' Wait 5-10 minutes before first use.' 'WARN'
        Write-Log '============================================' 'OK'
        return $true
    }
    catch {
        Write-Log "Setup error: $(Get-SafeErrorText $_)" 'ERR'
        if ((Test-Path variable:cert) -and $cert) {
            try { Remove-Item "Cert:\LocalMachine\My\$($cert.Thumbprint)" -DeleteKey -ErrorAction SilentlyContinue } catch {}
        }
        try { $null = Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
        return $false
    }
}

# ---------------------------------------------------------------
# 1  CORE LOGIC
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
        Write-Log "  Conflict search error: $_" 'WARN'
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
    $dlg.BackColor       = $C.BG
    $dlg.ForeColor       = $C.FG
    $dlg.Font            = $F.UI
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false

    # Header
    $pHdr = [System.Windows.Forms.Panel]::new()
    $pHdr.Dock = 'Top'; $pHdr.Height = 48; $pHdr.BackColor = $C.Danger
    $dlg.Controls.Add($pHdr)
    $lHdr = New-Lbl '  ImmutableID Collision -- Another Entra Object Owns This GUID' -Font $F.Hdr -Color $C.FGHdr
    $lHdr.Location = [System.Drawing.Point]::new(8,15); $pHdr.Controls.Add($lHdr)

    # Explanation
    $lExp = New-Lbl '' -Color $C.FG
    $lExp.Font = $F.UI
    $lExp.AutoSize = $false
    $lExp.Size = [System.Drawing.Size]::new(580, 110)
    $lExp.Location = [System.Drawing.Point]::new(16, 60)
    $lExp.Text = "The ImmutableID you are trying to assign:`r`n$TargetImmutableId`r`n`r`nis already owned by a different Entra account:`r`n`r`nThis usually means a stale duplicate, a ghost account from a previous sync, or a soft-deleted user. You must clear the ImmutableID from the conflicting account before the match can proceed."
    $dlg.Controls.Add($lExp)

    # Conflict details card
    $pCard = [System.Windows.Forms.Panel]::new()
    $pCard.Location = [System.Drawing.Point]::new(16, 172)
    $pCard.Size     = [System.Drawing.Size]::new(580, 88)
    $pCard.BackColor = $C.Card
    $dlg.Controls.Add($pCard)

    $lCardHdr = New-Lbl 'Conflicting Entra Account:' -Font $F.Hdr -Color $C.Warning
    $lCardHdr.Location = [System.Drawing.Point]::new(8,6); $pCard.Controls.Add($lCardHdr)

    if ($ConflictUser) {
        $lName = New-Lbl "Display Name : $($ConflictUser.displayName)" -Color $C.FG
        $lName.Location = [System.Drawing.Point]::new(8,26); $pCard.Controls.Add($lName)
        $lUPN  = New-Lbl "UPN          : $($ConflictUser.userPrincipalName)" -Color $C.FGDim
        $lUPN.Location  = [System.Drawing.Point]::new(8,44); $pCard.Controls.Add($lUPN)
        $lID   = New-Lbl "Object ID    : $($ConflictUser.id)" -Font $F.Mono -Color $C.FGDim
        $lID.Location   = [System.Drawing.Point]::new(8,62); $pCard.Controls.Add($lID)
    } else {
        $lNF = New-Lbl 'Conflicting account could not be looked up (may be a contact or deleted object).' -Color $C.Warning
        $lNF.Location = [System.Drawing.Point]::new(8,28); $pCard.Controls.Add($lNF)
    }

    # Buttons
    $pBot = [System.Windows.Forms.Panel]::new()
    $pBot.Dock = 'Bottom'; $pBot.Height = 46; $pBot.BackColor = $C.Panel
    $dlg.Controls.Add($pBot)

    $script:_conflictResult = 'Cancel'

    $btnClear = New-Btn 'Clear Conflict & Retry' 190 34 $C.Success
    $btnClear.Location = [System.Drawing.Point]::new(220, 6)
    $btnClear.Enabled  = ($ConflictUser -ne $null)
    $pBot.Controls.Add($btnClear)

    $btnSkip = New-Btn 'Skip This User' 130 34 $C.Warning
    $btnSkip.Location = [System.Drawing.Point]::new(418, 6)
    $pBot.Controls.Add($btnSkip)

    $btnCancel = New-Btn 'Cancel Batch' 110 34 $C.Danger
    $btnCancel.Location = [System.Drawing.Point]::new(8, 6)
    $pBot.Controls.Add($btnCancel)

    $btnClear.Add_Click({
        if ($ConflictUser) {
            try {
                Write-Log "  Clearing ImmutableID from conflicting account: $($ConflictUser.userPrincipalName)" 'WARN'
                $body = '{"onPremisesImmutableId": null}'
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($ConflictUser.id)" `
                    -Body $body -ContentType 'application/json' -EA Stop
                Write-Log "  Conflict cleared from $($ConflictUser.userPrincipalName)." 'OK'
                $script:_conflictResult = 'Cleared'
            } catch {
                Write-Log "  Failed to clear conflict: $_" 'ERR'
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

    Write-Log "  AD    : $($ADUser.DisplayName)  [$($ADUser.SamAccountName)]" 'INFO'
    Write-Log "  Entra : $($EntraUser.DisplayName)  [$($EntraUser.UserPrincipalName)]" 'INFO'
    Write-Log "  ImmutableID -> $targetId" 'INFO'

    if ($currentId -and $currentId -eq $targetId) {
        Write-Log '  Already matched -- skipped.' 'OK'
        return 'AlreadyMatched'
    }
    if ($currentId -and $currentId -ne $targetId) {
        Write-Log "  WARNING: overwriting existing ImmutableID on target ($currentId)" 'WARN'
    }

    if ($WhatIfMode) {
        Write-Log "  [WHATIF] Would set ImmutableID on $($EntraUser.UserPrincipalName)" 'WARN'
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
            Write-Log "  FAILED: $msg" 'ERR'
            return 'Error'
        }
        # Poll verify -- Graph eventual consistency can lag a few seconds
        $maxRetries = 8; $retryDelay = 3; $attempt = 0
        do {
            if ($attempt -gt 0) {
                Write-Log "  Verify attempt $attempt/$maxRetries -- waiting ${retryDelay}s for Graph consistency..." 'INFO'
                Start-Sleep -Seconds $retryDelay
            }
            $attempt++
            try {
                $v = Get-MgUser -UserId $EntraUser.Id -Property OnPremisesImmutableId -EA Stop
                if ($v.OnPremisesImmutableId -eq $targetId) {
                    Write-Log "  Hard match applied and verified (attempt $attempt)." 'OK'
                    return 'OK'
                }
            } catch {
                Write-Log "  Verify read error (attempt $attempt): $_" 'WARN'
            }
        } while ($attempt -lt $maxRetries)
        Write-Log "  Write appeared to succeed but ImmutableID not confirmed after $maxRetries attempts -- check manually." 'WARN'
        return 'VerifyFail'
    }

    $result = & $doApply

    # Handle conflict with UI resolver
    if ($result -eq 'Conflict') {
        Write-Log '  Conflict detected: another Entra object owns this ImmutableID.' 'WARN'
        Write-Log '  Searching for the conflicting account...' 'INFO'
        $conflictOwner = Find-ImmutableIdOwner -ImmutableId $targetId

        if ($conflictOwner) {
            Write-Log "  Conflict owner: $($conflictOwner.displayName) [$($conflictOwner.userPrincipalName)]" 'WARN'
        } else {
            Write-Log '  Could not identify conflict owner (may be deleted/contact object).' 'WARN'
        }

        if (-not $AllowConflictUI) {
            # Non-interactive path (called from batch without UI override)
            Write-Log '  Conflict resolution requires manual action -- skipping.' 'WARN'
            return 'Conflict'
        }

        $resolution = Show-ConflictDialog -ConflictUser $conflictOwner `
            -TargetImmutableId $targetId -ADUserName $ADUser.DisplayName

        switch ($resolution) {
            'Cleared' {
                Write-Log '  Retrying hard match after conflict cleared...' 'INFO'
                $result = & $doApply
                if ($result -eq 'OK') { return 'OK' }
                Write-Log "  Retry result: $result" $(if ($result -eq 'OK') {'OK'} else {'ERR'})
                return $result
            }
            'Skip'   { Write-Log '  Skipped by user.' 'WARN'; return 'Skipped' }
            'Cancel' { Write-Log '  Batch cancelled by user.' 'WARN'; return 'CancelBatch' }
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
    $dlg.BackColor       = $C.BG
    $dlg.ForeColor       = $C.FG
    $dlg.Font            = $F.UI
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false

    # Search row
    $pTop = [System.Windows.Forms.Panel]::new()
    $pTop.Dock = 'Top'; $pTop.Height = 46; $pTop.BackColor = $C.Panel
    $dlg.Controls.Add($pTop)

    $txt = New-Txt -W 480; $txt.Location = [System.Drawing.Point]::new(8,10); $txt.Text = $Q
    $pTop.Controls.Add($txt)
    $btnGo = New-Btn 'Search AD' 110 28 $C.Accent
    $btnGo.Location = [System.Drawing.Point]::new(496,10)
    $pTop.Controls.Add($btnGo)

    # Legend bar
    $pLeg = [System.Windows.Forms.Panel]::new()
    $pLeg.Dock = 'Top'; $pLeg.Height = 28; $pLeg.BackColor = $C.Panel
    $dlg.Controls.Add($pLeg)
    $lLeg = New-Lbl '  Columns:  Display Name  |  SAM Account Name (login)  |  User Principal Name (UPN / email login)  |  Account Enabled (Yes/No)   -- disabled accounts shown dimmed' -Font $F.Small -Color $C.FGDim
    $lLeg.Location = [System.Drawing.Point]::new(0,8); $pLeg.Controls.Add($lLeg)

    # List
    $lv = [System.Windows.Forms.ListView]::new()
    $lv.Dock = 'Fill'; $lv.View = 'Details'; $lv.FullRowSelect = $true
    $lv.BackColor = $C.Card; $lv.ForeColor = $C.FG; $lv.Font = $F.UI
    $lv.GridLines = $true; $lv.BorderStyle = 'None'
    [void]$lv.Columns.Add('Display Name',          190)
    [void]$lv.Columns.Add('SAM Account Name',       130)
    [void]$lv.Columns.Add('User Principal Name (UPN)', 200)
    [void]$lv.Columns.Add('Acct Enabled',            70)
    $dlg.Controls.Add($lv)

    # Bottom
    $pBot = [System.Windows.Forms.Panel]::new()
    $pBot.Dock = 'Bottom'; $pBot.Height = 42; $pBot.BackColor = $C.Panel
    $dlg.Controls.Add($pBot)
    $lblN = New-Lbl '0 results' -Color $C.FGDim; $lblN.Location = [System.Drawing.Point]::new(8,14)
    $pBot.Controls.Add($lblN)
    $btnOK = New-Btn 'Select' 90 30 $C.Success
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
                $li.ForeColor = if ($u.Enabled) { $C.FG } else { $C.FGDim }
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
    $dlg.BackColor       = $C.BG; $dlg.ForeColor = $C.FG
    $dlg.Font            = $F.UI; $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false

    $pTop = [System.Windows.Forms.Panel]::new()
    $pTop.Dock = 'Top'; $pTop.Height = 46; $pTop.BackColor = $C.Panel
    $dlg.Controls.Add($pTop)

    $txt = New-Txt -W 560; $txt.Location = [System.Drawing.Point]::new(8,10); $txt.Text = $Q
    $pTop.Controls.Add($txt)
    $btnGo = New-Btn 'Search Entra' 120 28 $C.Accent
    $btnGo.Location = [System.Drawing.Point]::new(576,10)
    $pTop.Controls.Add($btnGo)

    # Legend bar
    $pLeg2 = [System.Windows.Forms.Panel]::new()
    $pLeg2.Dock = 'Top'; $pLeg2.Height = 38; $pLeg2.BackColor = $C.Panel
    $dlg.Controls.Add($pLeg2)
    $lLeg2a = New-Lbl '  Columns:  Display Name  |  UPN (Entra login)  |  Account Enabled  |  ImmutableID (Base64 GUID -- links to AD)' -Font $F.Small -Color $C.FGDim
    $lLeg2a.Location = [System.Drawing.Point]::new(0,4); $pLeg2.Controls.Add($lLeg2a)
    $lLeg2b = New-Lbl '  Yellow rows = already has an ImmutableID (may already be matched to a different AD account -- review carefully before selecting)' -Font $F.Small -Color $C.Warning
    $lLeg2b.Location = [System.Drawing.Point]::new(0,20); $pLeg2.Controls.Add($lLeg2b)

    $lv = [System.Windows.Forms.ListView]::new()
    $lv.Dock = 'Fill'; $lv.View = 'Details'; $lv.FullRowSelect = $true
    $lv.BackColor = $C.Card; $lv.ForeColor = $C.FG; $lv.Font = $F.UI
    $lv.GridLines = $true; $lv.BorderStyle = 'None'
    [void]$lv.Columns.Add('Display Name',          195)
    [void]$lv.Columns.Add('User Principal Name',    230)
    [void]$lv.Columns.Add('Acct Enabled',            70)
    [void]$lv.Columns.Add('ImmutableID (existing)',  195)
    $dlg.Controls.Add($lv)

    $pBot = [System.Windows.Forms.Panel]::new()
    $pBot.Dock = 'Bottom'; $pBot.Height = 42; $pBot.BackColor = $C.Panel
    $dlg.Controls.Add($pBot)
    $lblN = New-Lbl '0 results' -Color $C.FGDim; $lblN.Location = [System.Drawing.Point]::new(8,14)
    $pBot.Controls.Add($lblN)
    $btnOK = New-Btn 'Select' 90 30 $C.Success
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
                $li.ForeColor = if ($u.OnPremisesImmutableId) { $C.Warning } else { $C.FG }
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
    $dlg.BackColor       = $C.BG; $dlg.ForeColor = $C.FG
    $dlg.Font            = $F.UI; $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false

    $tv = [System.Windows.Forms.TreeView]::new()
    $tv.Dock       = 'Fill'
    $tv.BackColor  = $C.Card
    $tv.ForeColor  = $C.FG
    $tv.Font       = $F.UI
    $tv.BorderStyle = 'None'
    $dlg.Controls.Add($tv)

    $pBot = [System.Windows.Forms.Panel]::new()
    $pBot.Dock = 'Bottom'; $pBot.Height = 42; $pBot.BackColor = $C.Panel
    $dlg.Controls.Add($pBot)
    $btnOK = New-Btn 'Select OU' 110 30 $C.Success
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
function Show-MainForm {

    $form = [System.Windows.Forms.Form]::new()
    $form.Text            = 'AD -> Entra ID Hard Match Tool'
    $form.Width           = 980
    $form.Height          = 780
    $form.MinimumSize     = [System.Drawing.Size]::new(980, 780)
    $form.StartPosition   = 'CenterScreen'
    $form.BackColor       = $C.BG
    $form.ForeColor       = $C.FG
    $form.Font            = $F.UI
    $form.FormBorderStyle = 'Sizable'

    # ---------- TITLE BAR ----------
    $pTitle = [System.Windows.Forms.Panel]::new()
    $pTitle.Dock = 'Top'; $pTitle.Height = 50; $pTitle.BackColor = $C.Panel
    $form.Controls.Add($pTitle)

    $lblTitle = New-Lbl 'AD  ->  Entra Hard Match Tool' -Font $F.Title -Color $C.FGHdr
    $lblTitle.Location = [System.Drawing.Point]::new(12, 14)
    $pTitle.Controls.Add($lblTitle)

    $lblWI = New-Lbl '[WHATIF -- NO CHANGES WILL BE WRITTEN]' -Font $F.Hdr -Color $C.Warning
    $lblWI.Location = [System.Drawing.Point]::new(320, 17)
    $lblWI.Visible  = $WhatIf.IsPresent
    $pTitle.Controls.Add($lblWI)

    $lblCfg = New-Lbl "Config: $script:ConfigPath" -Font $F.Small -Color $C.FGDim
    $lblCfg.Location = [System.Drawing.Point]::new(12, 34)
    $pTitle.Controls.Add($lblCfg)

    $btnSetupAuth = New-Btn 'Setup Auth' 120 30 $C.Warning
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
        $ok = Initialize-M365Auth
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
    $tabs.BackColor = $C.BG
    $tabs.Font      = $F.UIB
    $form.Controls.Add($tabs)

    # Bring tabs on top of title bar (Z order)
    $form.Controls.SetChildIndex($tabs, 0)

    # ---------- STATUS BAR ----------
    $pStatus = [System.Windows.Forms.Panel]::new()
    $pStatus.Dock = 'Bottom'; $pStatus.Height = 26; $pStatus.BackColor = $C.Panel
    $form.Controls.Add($pStatus)
    $lblStatus = New-Lbl 'Ready' -Color $C.FGDim; $lblStatus.Location = [System.Drawing.Point]::new(8,5)
    $pStatus.Controls.Add($lblStatus)
    $script:LblStatus = $lblStatus

    # ===================================================================
    # TAB 1 -- SINGLE USER
    # ===================================================================
    $tabSingle = [System.Windows.Forms.TabPage]::new('  Single User  ')
    $tabSingle.BackColor = $C.BG; $tabSingle.ForeColor = $C.FG
    [void]$tabs.TabPages.Add($tabSingle)

    # -- AD card --
    $cardAD = [System.Windows.Forms.Panel]::new()
    $cardAD.Location = [System.Drawing.Point]::new(12, 12)
    $cardAD.Size     = [System.Drawing.Size]::new(438, 200)
    $cardAD.BackColor = $C.Card
    $tabSingle.Controls.Add($cardAD)

    $lADH = New-Lbl 'Active Directory User' -Font $F.Hdr -Color $C.Accent
    $lADH.Location = [System.Drawing.Point]::new(8,8); $cardAD.Controls.Add($lADH)

    $txtADs = New-Txt 300; $txtADs.Location = [System.Drawing.Point]::new(8,32); $cardAD.Controls.Add($txtADs)
    $btnADs = New-Btn 'Search AD' 112 26 $C.Accent; $btnADs.Location = [System.Drawing.Point]::new(316,32); $cardAD.Controls.Add($btnADs)

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
        $lKey = New-Lbl $adLabels[$k] -Color $C.FGDim; $lKey.Location = $adFields[$k]; $cardAD.Controls.Add($lKey)
        $lVal = New-Lbl '--' -Color $C.FG
        $lVal.Location = [System.Drawing.Point]::new(68, $adFields[$k].Y)
        $lVal.MaximumSize = [System.Drawing.Size]::new(364,18)
        if ($k -in @('GUID','ImmID')) { $lVal.Font = $F.Mono; $lVal.ForeColor = $C.FGDim }
        $cardAD.Controls.Add($lVal)
        $adVals[$k] = $lVal
    }

    # arrow
    $lblArr = New-Lbl '->' -Font ([System.Drawing.Font]::new('Segoe UI',18,[System.Drawing.FontStyle]::Bold)) -Color $C.Accent
    $lblArr.Location = [System.Drawing.Point]::new(460, 90); $tabSingle.Controls.Add($lblArr)

    # -- Entra card --
    $cardE = [System.Windows.Forms.Panel]::new()
    $cardE.Location = [System.Drawing.Point]::new(492, 12)
    $cardE.Size     = [System.Drawing.Size]::new(462, 200)
    $cardE.BackColor = $C.Card
    $tabSingle.Controls.Add($cardE)

    $lEH = New-Lbl 'Entra ID (Azure AD) User' -Font $F.Hdr -Color $C.Accent
    $lEH.Location = [System.Drawing.Point]::new(8,8); $cardE.Controls.Add($lEH)

    $txtEs = New-Txt 316; $txtEs.Location = [System.Drawing.Point]::new(8,32); $cardE.Controls.Add($txtEs)
    $btnEs = New-Btn 'Search Entra' 124 26 $C.Accent; $btnEs.Location = [System.Drawing.Point]::new(328,32); $cardE.Controls.Add($btnEs)

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
        $lKey = New-Lbl $eLabels[$k] -Color $C.FGDim; $lKey.Location = $eFields[$k]; $cardE.Controls.Add($lKey)
        $lVal = New-Lbl '--' -Color $C.FG
        $lVal.Location = [System.Drawing.Point]::new(80, $eFields[$k].Y)
        $lVal.MaximumSize = [System.Drawing.Size]::new(376,18)
        if ($k -in @('ObjId','ImmID')) { $lVal.Font = $F.Mono; $lVal.ForeColor = $C.FGDim }
        $cardE.Controls.Add($lVal)
        $eVals[$k] = $lVal
    }

    # -- Computed preview --
    $pPrev = [System.Windows.Forms.Panel]::new()
    $pPrev.Location = [System.Drawing.Point]::new(12, 224)
    $pPrev.Size     = [System.Drawing.Size]::new(942, 44)
    $pPrev.BackColor = $C.Panel
    $tabSingle.Controls.Add($pPrev)

    $lPrevHdr = New-Lbl 'Computed ImmutableID (Base64 of AD GUID):' -Font $F.Hdr -Color $C.FGHdr
    $lPrevHdr.Location = [System.Drawing.Point]::new(8,4); $pPrev.Controls.Add($lPrevHdr)
    $lPrevVal = New-Lbl 'Select an AD user above to compute.' -Font $F.Mono -Color $C.FGDim
    $lPrevVal.Location = [System.Drawing.Point]::new(8,22)
    $lPrevVal.MaximumSize = [System.Drawing.Size]::new(930,18)
    $pPrev.Controls.Add($lPrevVal)

    # -- Validation notice --
    $lblV = New-Lbl '' -Font $F.UIB -Color $C.FGDim
    $lblV.Location = [System.Drawing.Point]::new(12, 278)
    $lblV.MaximumSize = [System.Drawing.Size]::new(942,60)
    $tabSingle.Controls.Add($lblV)

    # -- Action row --
    $pAct = [System.Windows.Forms.Panel]::new()
    $pAct.Location = [System.Drawing.Point]::new(0, 344)
    $pAct.Size     = [System.Drawing.Size]::new(962, 50)
    $pAct.BackColor = $C.Panel
    $tabSingle.Controls.Add($pAct)

    $btnApply  = New-Btn 'Apply Hard Match' 160 36 $C.Success; $btnApply.Location  = [System.Drawing.Point]::new(640,7); $btnApply.Enabled = $false
    $btnClearI = New-Btn 'Clear Entra ImmutableID' 200 36 $C.Danger;  $btnClearI.Location = [System.Drawing.Point]::new(430,7); $btnClearI.Enabled = $false
    $btnClearS = New-Btn 'Clear Selections' 140 36; $btnClearS.Location = [System.Drawing.Point]::new(280,7)
    $pAct.Controls.AddRange(@($btnApply,$btnClearI,$btnClearS))

    # -- Log (single tab) --
    $logS = [System.Windows.Forms.RichTextBox]::new()
    $logS.Location  = [System.Drawing.Point]::new(0, 394)
    $logS.Size      = [System.Drawing.Size]::new(962, 290)
    $logS.BackColor = $C.LogBG; $logS.ForeColor = $C.FG
    $logS.Font      = $F.Mono; $logS.ReadOnly = $true
    $logS.BorderStyle = 'None'; $logS.ScrollBars = 'Vertical'
    $tabSingle.Controls.Add($logS)
    $script:LogBox = $logS

    # ---- single tab: wire validation ----
    $updateVal = {
        $btnApply.Enabled  = $false
        $btnClearI.Enabled = $false
        if (-not $script:SelAD -or -not $script:SelEntra) {
            $lblV.Text = 'Select both an AD user and an Entra user to proceed.'
            $lblV.ForeColor = $C.FGDim; return
        }
        $cid = Get-ImmutableId -g $script:SelAD.ObjectGUID
        $eid = $script:SelEntra.OnPremisesImmutableId
        if ($eid -and $eid -eq $cid) {
            $lblV.Text = '[OK] ImmutableID already matches -- these accounts are already hard-matched.'
            $lblV.ForeColor = $C.Success
        } elseif ($eid -and $eid -ne $cid) {
            $lblV.Text = '[!] Entra user already has a DIFFERENT ImmutableID. Applying will overwrite it.'
            $lblV.ForeColor = $C.Warning
            $btnApply.Enabled = $true; $btnClearI.Enabled = $true
        } else {
            $lblV.Text = '[OK] Ready -- Entra user has no existing ImmutableID.'
            $lblV.ForeColor = $C.Success
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
            $lPrevVal.Text      = $cid; $lPrevVal.ForeColor = $C.Accent
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
            $eVals['ImmID'].ForeColor = if ($p.OnPremisesImmutableId) { $C.Warning } else { $C.FGDim }
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
        if ($WhatIf) { Write-Log "[WHATIF] Would clear ImmutableID on $($script:SelEntra.UserPrincipalName)" 'WARN'; return }
        try {
            $body = '{"onPremisesImmutableId": null}'
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "https://graph.microsoft.com/v1.0/users/$($script:SelEntra.Id)" `
                -Body $body -ContentType 'application/json'
            Write-Log "ImmutableID cleared on $($script:SelEntra.UserPrincipalName)." 'OK'
            $refreshed = Get-MgUser -UserId $script:SelEntra.Id `
                -Property Id,DisplayName,UserPrincipalName,AccountEnabled,OnPremisesImmutableId
            $script:SelEntra = $refreshed
            $eVals['ImmID'].Text = '(none)'; $btnClearI.Enabled = $false
            & $updateVal
        } catch { Write-Log "Clear failed: $_" 'ERR' }
    })

    $btnClearS.Add_Click({
        $script:SelAD = $null; $script:SelEntra = $null
        foreach ($k in $adVals.Keys)  { $adVals[$k].Text = '--' }
        foreach ($k in $eVals.Keys)   { $eVals[$k].Text  = '--' }
        $lPrevVal.Text = 'Select an AD user above to compute.'
        $lPrevVal.ForeColor = $C.FGDim
        $lblV.Text = ''; $btnApply.Enabled = $false; $btnClearI.Enabled = $false
        Set-Status 'Selections cleared.'
    })

    # ===================================================================
    # TAB 2 -- BATCH OU SCAN
    # ===================================================================
    $tabBatch = [System.Windows.Forms.TabPage]::new('  Batch (OU Scan)  ')
    $tabBatch.BackColor = $C.BG; $tabBatch.ForeColor = $C.FG
    [void]$tabs.TabPages.Add($tabBatch)

    # -- OU selector row --
    $pOU = [System.Windows.Forms.Panel]::new()
    $pOU.Location = [System.Drawing.Point]::new(12,12)
    $pOU.Size     = [System.Drawing.Size]::new(942,42)
    $pOU.BackColor = $C.Panel
    $tabBatch.Controls.Add($pOU)

    $lOU = New-Lbl 'OU:' -Font $F.Hdr -Color $C.FGHdr; $lOU.Location = [System.Drawing.Point]::new(8,12); $pOU.Controls.Add($lOU)
    $txtOU = New-Txt 680; $txtOU.Location = [System.Drawing.Point]::new(36,9); $txtOU.Text = $script:Config['DefaultOU']
    $pOU.Controls.Add($txtOU)
    $btnPickOU = New-Btn 'Browse OU' 120 26 $C.Accent; $btnPickOU.Location = [System.Drawing.Point]::new(720,9); $pOU.Controls.Add($btnPickOU)

    # -- Options row --
    $pOpts = [System.Windows.Forms.Panel]::new()
    $pOpts.Location = [System.Drawing.Point]::new(12,62)
    $pOpts.Size     = [System.Drawing.Size]::new(942,34)
    $pOpts.BackColor = $C.BG
    $tabBatch.Controls.Add($pOpts)

    $chkEnabled = [System.Windows.Forms.CheckBox]::new()
    $chkEnabled.Text = 'Enabled users only'; $chkEnabled.Checked = $true
    $chkEnabled.ForeColor = $C.FG; $chkEnabled.BackColor = [System.Drawing.Color]::Transparent
    $chkEnabled.Location = [System.Drawing.Point]::new(0,6); $chkEnabled.AutoSize = $true
    $pOpts.Controls.Add($chkEnabled)

    $chkNoImm = [System.Windows.Forms.CheckBox]::new()
    $chkNoImm.Text = 'Skip already-matched users'; $chkNoImm.Checked = $true
    $chkNoImm.ForeColor = $C.FG; $chkNoImm.BackColor = [System.Drawing.Color]::Transparent
    $chkNoImm.Location = [System.Drawing.Point]::new(170,6); $chkNoImm.AutoSize = $true
    $pOpts.Controls.Add($chkNoImm)

    $btnLoad = New-Btn 'Load Users from OU' 180 28 $C.Accent; $btnLoad.Location = [System.Drawing.Point]::new(380,4)
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
    $lv2.BackColor    = $C.Card; $lv2.ForeColor = $C.FG
    $lv2.Font         = $F.UI;  $lv2.GridLines = $true
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
    $pBAct.BackColor = $C.Panel
    $tabBatch.Controls.Add($pBAct)

    $lblSel = New-Lbl '0 users loaded' -Color $C.FGDim; $lblSel.Location = [System.Drawing.Point]::new(8,16); $pBAct.Controls.Add($lblSel)
    $btnResolve = New-Btn 'Re-Resolve Entra' 160 36 $C.Accent
    $btnResolve.Location = [System.Drawing.Point]::new(560,7); $btnResolve.Enabled = $false; $pBAct.Controls.Add($btnResolve)
    $btnBatchApply = New-Btn 'Apply Checked Matches' 200 36 $C.Success
    $btnBatchApply.Location = [System.Drawing.Point]::new(728,7); $btnBatchApply.Enabled = $false; $pBAct.Controls.Add($btnBatchApply)

    # -- Batch log --
    $logB = [System.Windows.Forms.RichTextBox]::new()
    $logB.Location  = [System.Drawing.Point]::new(0, 482)
    $logB.Size      = [System.Drawing.Size]::new(962, 190)
    $logB.BackColor = $C.LogBG; $logB.ForeColor = $C.FG
    $logB.Font      = $F.Mono; $logB.ReadOnly = $true
    $logB.BorderStyle = 'None'; $logB.ScrollBars = 'Vertical'
    $tabBatch.Controls.Add($logB)

    # Switch active log when tab changes
    $tabs.Add_SelectedIndexChanged({
        $script:LogBox = if ($tabs.SelectedIndex -eq 0) { $logS } else { $logB }
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
        $script:LogBox = $logB
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
            Write-Log "Loaded $($adUsers.Count) AD user(s) from $ouDN" 'INFO'
        } catch {
            Write-Log "Failed to load AD users: $_" 'ERR'
            $lblSel.Text = 'Load failed.'
            return
        }

        if (-not $adUsers -or $adUsers.Count -eq 0) {
            Write-Log 'No users found in this OU (check filter options).' 'WARN'
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
                'Exact' { $C.FG }
                'Name'  { $C.Warning }
                default { $C.FGDim }
            }

            # Auto-check rows with an Exact match that need matching
            $li.Checked = ($res.Confidence -eq 'Exact' -and $action -eq 'Will match')
            [void]$lv2.Items.Add($li)
        }

        $total   = $lv2.Items.Count
        $checked = ($lv2.Items | Where-Object { $_.Checked }).Count
        $lblSel.Text = "$total user(s) listed  |  $checked checked"
        $btnResolve.Enabled     = ($total -gt 0)
        $btnBatchApply.Enabled  = ($checked -gt 0)
        Set-Status "Loaded $total user(s)."
    })

    # Update count label on check change
    $lv2.Add_ItemChecked({
        $total   = $lv2.Items.Count
        $checked = ($lv2.Items | Where-Object { $_.Checked }).Count
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
        $script:LogBox = $logB
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
            $li.ForeColor = switch ($res.Confidence) { 'Exact' {$C.FG} 'Name' {$C.Warning} default {$C.FGDim} }
        }
        Set-Status 'Re-resolve complete.'
    })

    # Batch Apply
    $btnBatchApply.Add_Click({
        $script:LogBox = $logB
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
            Write-Log "--- $($td.ADUser.DisplayName) ---" 'INFO'
            $res = Invoke-HardMatch -ADUser $td.ADUser -EntraUser $td.EntraUser `
                       -WhatIfMode $WhatIf.IsPresent -AllowConflictUI $true
            switch ($res) {
                'OK'            { $ok++;   $li.SubItems[6].Text = 'Matched';         $li.ForeColor = $C.Success }
                'AlreadyMatched'{ $skip++; $li.SubItems[6].Text = 'Already matched'; $li.ForeColor = $C.FGDim }
                'WhatIf'        { $skip++; $li.SubItems[6].Text = 'WhatIf';          $li.ForeColor = $C.Warning }
                'Skipped'       { $skip++; $li.SubItems[6].Text = 'Skipped';         $li.ForeColor = $C.FGDim }
                'Conflict'      { $fail++; $li.SubItems[6].Text = 'Conflict -- manual resolve needed'; $li.ForeColor = $C.Danger }
                'CancelBatch'   {
                    $cancelled = $true
                    $li.SubItems[6].Text = 'Cancelled'
                    $li.ForeColor = $C.FGDim
                    Write-Log 'Batch cancelled by user at conflict dialog.' 'WARN'
                }
                default         { $fail++; $li.SubItems[6].Text = "Failed: $res";   $li.ForeColor = $C.Danger }
            }
            $li.Checked = $false
            $form.Refresh()
        }
        $summary = "Batch done: $ok matched  |  $skip skipped  |  $fail failed"
        if ($cancelled) { $summary += '  |  CANCELLED' }
        Write-Log $summary $(if ($fail -gt 0 -or $cancelled) {'WARN'} else {'OK'})
        Set-Status $summary
        $btnBatchApply.Enabled = $false
    })

    # ===================================================================
    # FINAL STARTUP LOG
    # ===================================================================
    $script:LogBox = $logS
    Write-Log 'Hard Match Tool ready.' 'OK'
    Write-Log "Config file: $script:ConfigPath" 'INFO'
    if ($WhatIf) { Write-Log 'WhatIf mode -- no changes will be written.' 'WARN' }

    [void]$form.ShowDialog()
}

# ---------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------
if (-not $SkipModuleCheck) { Invoke-ModuleBootstrap }

# Az.Accounts and Microsoft.Graph share underlying auth/token-cache assemblies;
# having both loaded in the same session is a known cause of the SDK throwing
# odd PowerShell errors (e.g. "The property 'Warning' cannot be found on this
# object") when Connect-MgGraph is called. Flag it early rather than let it
# surface as a confusing crash later.
if (Get-Module -Name Az.Accounts -ErrorAction SilentlyContinue) {
    Write-Log "Az.Accounts module is also loaded in this session -- this is a known cause of odd Microsoft.Graph SDK errors (e.g. \"property 'Warning' cannot be found\"). If Graph auth fails below, close this session and re-run in a fresh PowerShell 7 window." 'WARN'
}

# If ADM365Config.json is missing or incomplete, offer the same one-time Setup
# Auth wizard used by Onboard-ADUser.ps1 / Offboard-ADUser.ps1 before trying to
# connect -- it creates the shared "ADM365LifecycleTool" app + cert and writes
# AppId/TenantId/CertThumbprint back into the shared config file.
$script:_cfgReady = [bool]($script:Config['AppId'] -and $script:Config['TenantId'] -and $script:Config['CertThumbprint'])
if (-not $script:_cfgReady) {
    if ($script:IsCLI) {
        Write-Log "ADM365Config.json missing or incomplete at $script:ConfigPath -- run this tool in GUI mode and use 'Setup Auth' (or run the same step from Onboard-ADUser.ps1 / Offboard-ADUser.ps1) first." 'ERR'
        exit 2
    }
    $r = [System.Windows.Forms.MessageBox]::Show(
        "No M365 auth is configured yet ($script:ConfigPath is missing or incomplete).`r`n`r`nRun the one-time Setup Auth wizard now? This is the same setup used by Onboard-ADUser.ps1 / Offboard-ADUser.ps1 -- it creates a shared app registration + certificate and writes them to ADM365Config.json.`r`n`r`nRequires a Global Admin sign-in.",
        'M365 Auth Not Configured', 'YesNo', 'Question')
    if ($r -eq 'Yes') {
        try { $null = Initialize-M365Auth }
        catch { Write-Log "Setup Auth crashed unexpectedly: $(Get-SafeErrorText $_)" 'ERR' }
    }
}

try {
    $graphOk = Connect-ToGraph
} catch {
    Write-Log "Connect-ToGraph crashed unexpectedly: $(Get-SafeErrorText $_)" 'ERR'
    $graphOk = $false
}
if (-not $graphOk) {
    if ($script:IsCLI) {
        Write-Log 'Graph connection failed -- cannot proceed in CLI mode.' 'ERR'
        exit 2
    }
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Could not connect to Microsoft Graph.`r`nCheck ADM365Config.json (AppId / TenantId / CertThumbprint), or use the 'Setup Auth' button in the main window.`r`n`r`nContinue anyway? (Entra features will fail)",
        'Graph Connection Failed','YesNo','Warning')
    if ($r -ne 'Yes') { exit 1 }
}

# ---- CLI: single-user mode ----
if ($SamAccountName -ne '' -and -not $BatchMode.IsPresent) {
    Write-Log "CLI single-user mode: SAM=$SamAccountName" 'INFO'
    try {
        $adUser = Get-ADUser -Identity $SamAccountName `
            -Properties DisplayName,SamAccountName,UserPrincipalName,EmailAddress,ObjectGUID,DistinguishedName `
            -EA Stop
    } catch {
        Write-Log "AD user not found: $SamAccountName -- $_" 'ERR'; exit 2
    }

    $res = Resolve-EntraCandidate -ADUser $adUser
    if ($res.Confidence -eq 'None' -and $EntraUPN -eq '') {
        Write-Log "No Entra match found for $SamAccountName and no -EntraUPN supplied." 'ERR'; exit 2
    }

    $entraUser = if ($EntraUPN -ne '') {
        try { Get-MgUser -UserId $EntraUPN -Property Id,DisplayName,UserPrincipalName,AccountEnabled,OnPremisesImmutableId -EA Stop }
        catch { Write-Log "Entra user not found: $EntraUPN -- $_" 'ERR'; exit 2 }
    } else { $res.EntraUser }

    $result = Invoke-HardMatch -ADUser $adUser -EntraUser $entraUser -WhatIfMode $WhatIf.IsPresent -AllowConflictUI $false
    Write-Log "Result: $result" $(if ($result -eq 'OK' -or $result -eq 'AlreadyMatched') {'OK'} else {'ERR'})
    exit $(if ($result -eq 'OK' -or $result -eq 'AlreadyMatched') {0} elseif ($result -eq 'WhatIf') {0} else {1})
}

# ---- CLI: batch OU mode ----
if ($OUPath -ne '' -and $BatchMode.IsPresent) {
    Write-Log "CLI batch mode: OU=$OUPath" 'INFO'
    try {
        $adUsers = @(Get-ADUser -SearchBase $OUPath -SearchScope Subtree -Filter { Enabled -eq $true } `
            -Properties DisplayName,SamAccountName,UserPrincipalName,EmailAddress,ObjectGUID,DistinguishedName `
            -EA Stop | Sort-Object DisplayName)
    } catch {
        Write-Log "Failed to load AD users from ${OUPath}: $_" 'ERR'; exit 2
    }
    Write-Log "Loaded $($adUsers.Count) enabled AD user(s)." 'INFO'
    $ok=0; $skip=0; $fail=0
    foreach ($u in $adUsers) {
        Write-Log "--- $($u.DisplayName) [$($u.SamAccountName)] ---" 'INFO'
        $res = Resolve-EntraCandidate -ADUser $u
        if ($res.Confidence -eq 'None') { Write-Log '  No Entra match -- skipped.' 'WARN'; $skip++; continue }
        $result = Invoke-HardMatch -ADUser $u -EntraUser $res.EntraUser -WhatIfMode $WhatIf.IsPresent -AllowConflictUI $false
        switch ($result) {
            'OK'            { $ok++ }
            'AlreadyMatched'{ $skip++ }
            'WhatIf'        { $skip++ }
            default         { $fail++ }
        }
    }
    Write-Log "CLI batch complete: $ok matched, $skip skipped, $fail failed." $(if ($fail -gt 0) {'WARN'} else {'OK'})
    exit $(if ($fail -gt 0) {1} else {0})
}

# ---- GUI mode ----
Show-MainForm
