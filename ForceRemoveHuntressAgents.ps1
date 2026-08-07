<#
.SYNOPSIS
    Uninstalls Huntress EDR, ITDR, or all Huntress agents from a Windows workstation.

.DESCRIPTION
    This script searches installed programs for Huntress-related products and uninstalls them.
    It uses the uninstall command from the registry or falls back to msiexec with the product code.
    Requires administrative privileges.

.NOTES
    Author: Senior Systems Engineer
    Date: 2025-12-15
#>

# Ensure script is running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

Write-Host "Searching for Huntress agents..." -ForegroundColor Cyan

# Registry paths for installed software
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

$found = @()

foreach ($path in $regPaths) {
    if (Test-Path $path) {
        Get-ChildItem $path | ForEach-Object {
            $dispName = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).DisplayName
            $uninstallStr = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).UninstallString
            $prodCode = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).PSChildName

            if ($dispName -and $dispName -match "Huntress") {
                $found += [PSCustomObject]@{
                    Name          = $dispName
                    UninstallCmd  = $uninstallStr
                    ProductCode   = $prodCode
                }
            }
        }
    }
}

if (-not $found) {
    Write-Host "No Huntress agents found on this system." -ForegroundColor Yellow
    exit 0
}

Write-Host "Found the following Huntress products:" -ForegroundColor Green
$found | Format-Table -AutoSize

foreach ($app in $found) {
    Write-Host "`nUninstalling: $($app.Name)" -ForegroundColor Cyan
    try {
        if ($app.UninstallCmd) {
            # Some uninstall strings have quotes and extra args
            $cmd, $args = $null, $null
            if ($app.UninstallCmd -match '^"(.+?)"\s*(.*)$') {
                $cmd = $matches[1]
                $args = $matches[2]
            } else {
                $parts = $app.UninstallCmd -split "\s+", 2
                $cmd = $parts[0]
                $args = if ($parts.Count -gt 1) { $parts[1] } else { "" }
            }

            Start-Process -FilePath $cmd -ArgumentList $args -Wait -NoNewWindow
        }
        elseif ($app.ProductCode -match "^\{.*\}$") {
            # Fallback to msiexec if only product code is available
            Start-Process "msiexec.exe" -ArgumentList "/x $($app.ProductCode) /qn /norestart" -Wait -NoNewWindow
        }
        else {
            Write-Warning "No uninstall method found for $($app.Name)"
        }
    }
    catch {
        Write-Error "Failed to uninstall $($app.Name): $_"
    }
}

Write-Host "`nUninstallation process completed." -ForegroundColor Green
