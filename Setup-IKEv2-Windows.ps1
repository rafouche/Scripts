# 1. Check for Administrative Privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as an Administrator."
    break
}

# 2. Set Registry Key for NAT Traversal (Error 809 fix)
$RegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\PolicyAgent"
$Name         = "AssumeUDPEncapsulationContextOnSendRule"
$Value        = 2

Write-Host "Configuring Registry Key for NAT Traversal..." -ForegroundColor Cyan
if (-not (Test-Path $RegistryPath)) {
    New-Item -Path $RegistryPath -Force | Out-Null
}

Set-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -Type DWord
Write-Host "Success: $Name set to $Value." -ForegroundColor Green

# 3. Set Registry Key for enabling IPSec
$RegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Parameters"
$Name         = "ProhibitIpSec"
$Value        = 0

Write-Host "Configuring Registry Key to enable IPSec..." -ForegroundColor Cyan
if (-not (Test-Path $RegistryPath)) {
    New-Item -Path $RegistryPath -Force | Out-Null
}

Set-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -Type DWord
Write-Host "Success: $Name set to $Value." -ForegroundColor Green

# 4. Configure and Start IPsec Services
$Services = @("IKEEXT", "PolicyAgent") # IKE and AuthIP, IPsec Policy Agent

foreach ($Service in $Services) {
    Write-Host "Checking service: $Service..." -ForegroundColor Cyan
    $ServiceObj = Get-Service -Name $Service -ErrorAction SilentlyContinue
    
    if ($ServiceObj) {
        # Set to Automatic
        Set-Service -Name $Service -StartupType Automatic
        
        # Start if not running
        if ($ServiceObj.Status -ne 'Running') {
            Write-Host "Starting $Service..." -ForegroundColor Yellow
            Start-Service -Name $Service
        }
        Write-Host "Status: $($ServiceObj.Status) | Startup: Automatic" -ForegroundColor Green
    } else {
        Write-Warning "Service $Service was not found on this system."
    }
}

Write-Host "`nConfiguration Complete." -ForegroundColor Cyan
Write-Host "CRITICAL: You must REBOOT this computer for the registry changes to take effect." -ForegroundColor Red -BackgroundColor Black
