# ==============================================================================
#  VPN Connection Setup Script
#  Supports: IKEv2 | L2TP/IPsec
#  Version:  1.4
#  Author:   Altec Solutions Group, Inc.
#
#  USAGE EXAMPLES:
#    Run with all defaults embedded below:
#      .\Setup-VPN.ps1
#
#    Override specific parameters:
#      .\Setup-VPN.ps1 -ConnectionName "Acme VPN" -ServerAddress "vpn.acme.com" -L2tpPsk "secretkey"
#
#    Full override, IKEv2:
#      .\Setup-VPN.ps1 -VpnType IKEv2 -ConnectionName "Acme VPN" -ServerAddress "vpn.acme.com" -AuthMethod MachineCertificate
#
#    Suppress reboot prompt (useful for RMM deployment):
#      .\Setup-VPN.ps1 -PromptReboot:$false
#
#    Skip DNS suffix:
#      .\Setup-VPN.ps1 -ConnectionName "Acme VPN" -ServerAddress "vpn.acme.com" -DnsSuffix ""
# ==============================================================================

[CmdletBinding()]
param(
    # --- General ---
    [string]$ConnectionName    = "Gold VPN",
    [string]$ServerAddress     = "gmi-dqzzqdtvrv.dynamic-m.com",
    [bool]  $AllUserConnection = $true,
    [bool]  $SplitTunneling    = $false,

    # --- VPN Type ---
    [ValidateSet("IKEv2", "L2TP")]
    [string]$VpnType           = "L2TP",

    # --- Authentication ---
    [ValidateSet("MSChapv2", "Chap", "Pap", "Eap", "MachineCertificate")]
    [string]$AuthMethod        = "Pap",

    # --- L2TP Settings ---
    # No default -- do not hardcode a real pre-shared key here. Pass it via
    # -L2tpPsk (e.g. as a NinjaRMM script parameter/secure custom field).
    [string]$L2tpPsk           = "",

    # --- IKEv2 Settings ---
    [ValidateSet("NoEncryption", "Optional", "Required", "Maximum", "Custom")]
    [string]$IKEv2Encryption   = "Optional",

    # --- DNS Suffix (leave blank to skip) ---
    [string]$DnsSuffix         = "gold.local",

    # --- Reboot Prompt ---
    [bool]  $PromptReboot      = $true
)

# ==============================================================================
#  SCRIPT BODY — Do not edit below unless customizing behavior
# ==============================================================================

# ------------------------------------------------------------------------------
# 1. Privilege Check
# ------------------------------------------------------------------------------
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Error "This script must be run as Administrator. Right-click > Run as Administrator."
    exit 1
}

if ($VpnType -eq "L2TP" -and [string]::IsNullOrWhiteSpace($L2tpPsk)) {
    Write-Error "L2tpPsk was not supplied. Pass it with -L2tpPsk '<key>' (e.g. as a NinjaRMM script parameter) -- it is no longer hardcoded in this file."
    exit 1
}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  VPN Setup: $ConnectionName"                           -ForegroundColor Cyan
Write-Host "  Type     : $VpnType"                                  -ForegroundColor Cyan
Write-Host "  Server   : $ServerAddress"                            -ForegroundColor Cyan
Write-Host "  Scope    : $(if ($AllUserConnection) { 'All Users' } else { 'Current User' })" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# ------------------------------------------------------------------------------
# 2. L2TP-Specific: Registry and Service Prerequisites
#    Skipped entirely for IKEv2
# ------------------------------------------------------------------------------
if ($VpnType -eq "L2TP") {

    Write-Host "[L2TP] Applying registry fixes and service configuration..." -ForegroundColor Yellow
    Write-Host ""

    # --- Fix for Error 809: NAT Traversal ---
    $RegNat = @{
        Path  = "HKLM:\SYSTEM\CurrentControlSet\Services\PolicyAgent"
        Name  = "AssumeUDPEncapsulationContextOnSendRule"
        Value = 2
    }
    Write-Host "  Configuring NAT Traversal (Error 809 fix)..." -ForegroundColor Cyan
    if (-not (Test-Path $RegNat.Path)) { New-Item -Path $RegNat.Path -Force | Out-Null }
    Set-ItemProperty -Path $RegNat.Path -Name $RegNat.Name -Value $RegNat.Value -Type DWord
    Write-Host "  OK: $($RegNat.Name) = $($RegNat.Value)" -ForegroundColor Green

    # --- Enable IPsec (ProhibitIpSec = 0) ---
    $RegIpSec = @{
        Path  = "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Parameters"
        Name  = "ProhibitIpSec"
        Value = 0
    }
    Write-Host "  Enabling IPsec..." -ForegroundColor Cyan
    if (-not (Test-Path $RegIpSec.Path)) { New-Item -Path $RegIpSec.Path -Force | Out-Null }
    Set-ItemProperty -Path $RegIpSec.Path -Name $RegIpSec.Name -Value $RegIpSec.Value -Type DWord
    Write-Host "  OK: $($RegIpSec.Name) = $($RegIpSec.Value)" -ForegroundColor Green

    Write-Host ""

    # --- Ensure IPsec Services are Running ---
    $Services = @(
        @{ Name = "IKEEXT";      DisplayName = "IKE and AuthIP IPsec Keying Modules" },
        @{ Name = "PolicyAgent"; DisplayName = "IPsec Policy Agent" }
    )

    foreach ($Svc in $Services) {
        Write-Host "  Service: $($Svc.DisplayName)..." -ForegroundColor Cyan
        $ServiceObj = Get-Service -Name $Svc.Name -ErrorAction SilentlyContinue

        if ($null -eq $ServiceObj) {
            Write-Warning "  Service '$($Svc.Name)' not found on this system."
            continue
        }

        Set-Service -Name $Svc.Name -StartupType Automatic

        $ServiceObj = Get-Service -Name $Svc.Name
        if ($ServiceObj.Status -ne 'Running') {
            Write-Host "  Starting $($Svc.Name)..." -ForegroundColor Yellow
            Start-Service -Name $Svc.Name -ErrorAction SilentlyContinue
            $ServiceObj = Get-Service -Name $Svc.Name
        }

        Write-Host "  OK: Status=$($ServiceObj.Status) | Startup=Automatic" -ForegroundColor Green
    }

    Write-Host ""
}

# ------------------------------------------------------------------------------
# 3. Remove Existing VPN Connection (if it exists) to ensure clean state
# ------------------------------------------------------------------------------
$ExistingVpn = Get-VpnConnection -Name $ConnectionName -AllUserConnection:$AllUserConnection -ErrorAction SilentlyContinue
if ($ExistingVpn) {
    Write-Host "Removing existing VPN connection '$ConnectionName'..." -ForegroundColor Yellow
    Remove-VpnConnection -Name $ConnectionName -AllUserConnection:$AllUserConnection -Force
    Write-Host "OK: Removed." -ForegroundColor Green
    Write-Host ""
}

# ------------------------------------------------------------------------------
# 4. Create the VPN Connection
# ------------------------------------------------------------------------------
Write-Host "Creating VPN connection '$ConnectionName'..." -ForegroundColor Cyan

try {
    if ($VpnType -eq "L2TP") {

        Add-VpnConnection `
            -AllUserConnection:$AllUserConnection `
            -Name                  $ConnectionName `
            -ServerAddress         $ServerAddress `
            -TunnelType            "L2tp" `
            -L2tpPsk               $L2tpPsk `
            -AuthenticationMethod  $AuthMethod `
            -SplitTunneling:$SplitTunneling `
            -RememberCredential `
            -Force

    } elseif ($VpnType -eq "IKEv2") {

        Add-VpnConnection `
            -AllUserConnection:$AllUserConnection `
            -Name                  $ConnectionName `
            -ServerAddress         $ServerAddress `
            -TunnelType            "IKEv2" `
            -AuthenticationMethod  $AuthMethod `
            -EncryptionLevel       $IKEv2Encryption `
            -SplitTunneling:$SplitTunneling `
            -RememberCredential `
            -Force

    }

    Write-Host "OK: VPN connection created successfully." -ForegroundColor Green

} catch {
    Write-Error "Failed to create VPN connection: $_"
    exit 1
}

# ------------------------------------------------------------------------------
# 5. Optional: Set DNS Suffix
# ------------------------------------------------------------------------------
if ($DnsSuffix -ne "") {
    Write-Host ""
    Write-Host "Setting DNS suffix to '$DnsSuffix'..." -ForegroundColor Cyan
    try {
        Set-VpnConnection -Name $ConnectionName -DnsSuffix $DnsSuffix -AllUserConnection:$AllUserConnection
        Write-Host "OK: DNS suffix set." -ForegroundColor Green
    } catch {
        Write-Warning "Could not set DNS suffix: $_"
    }
}

# ------------------------------------------------------------------------------
# 6. Summary
# ------------------------------------------------------------------------------
Write-Host ""
Write-Host "======================================================" -ForegroundColor Green
Write-Host "  Setup Complete!" -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green
Write-Host ""

$FinalVpn = Get-VpnConnection -Name $ConnectionName -AllUserConnection:$AllUserConnection -ErrorAction SilentlyContinue
if ($FinalVpn) {
    Write-Host "  Connection Name : $($FinalVpn.Name)"
    Write-Host "  Server          : $($FinalVpn.ServerAddress)"
    Write-Host "  Tunnel Type     : $($FinalVpn.TunnelType)"
    Write-Host "  Auth Method     : $($FinalVpn.AuthenticationMethod)"
    Write-Host "  Split Tunneling : $($FinalVpn.SplitTunneling)"
    Write-Host "  DNS Suffix      : $($FinalVpn.DnsSuffix)"
    Write-Host "  Scope           : $(if ($AllUserConnection) { 'All Users' } else { 'Current User' })"
}

Write-Host ""

# --- Reboot Warning (L2TP only) ---
if ($VpnType -eq "L2TP" -and $PromptReboot) {
    Write-Host "  !! REBOOT REQUIRED !!" -ForegroundColor Red -BackgroundColor Black
    Write-Host "  Registry changes for L2TP/IPsec will not take effect until" -ForegroundColor Red
    Write-Host "  this computer is restarted. Please reboot before testing." -ForegroundColor Red
    Write-Host ""
    $Reboot = Read-Host "  Reboot now? (Y/N)"
    if ($Reboot -match "^[Yy]") {
        Write-Host "Rebooting in 15 seconds... Press Ctrl+C to cancel." -ForegroundColor Yellow
        Start-Sleep -Seconds 15
        Restart-Computer -Force
    } else {
        Write-Host "Reboot skipped. Remember to restart before connecting." -ForegroundColor Yellow
    }
}
