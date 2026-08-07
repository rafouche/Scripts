<#
.SYNOPSIS
Bypasses Windows 11 hardware requirements for in-place upgrades.

.DESCRIPTION
This script modifies specific registry values to override system compatibility checks performed during
Windows 11 upgrades. It removes legacy upgrade failure entries, simulates compatible hardware state,
enables Microsoft's documented bypass policy for unsupported TPM or CPU configurations, and sets the
UpgradeEligibility flag required by the Windows 11 Upgrade Assistant.

This is intended for lab, evaluation, or controlled environments where hardware policy allows.

.NOTES
Author: asheroto
Source: https://gist.github.com/asheroto/5087d2a38b311b0c92be2a4f23f92d3e
Required: Run as Administrator

.LICENSE
Use at your own risk. No warranty expressed or implied.
#>

function Write-Section {
<#
.SYNOPSIS
Displays a section header with borders using Write-Host and optional color.

.DESCRIPTION
Prints multi-line text surrounded by a hash border for readability.
Supports output coloring via the Color parameter.

.PARAMETER Text
The text to display. Can include multiple lines.

.PARAMETER Color
(Optional) The color to use for the text and border. Defaults to White.

.EXAMPLE
Write-Section -Text "Starting Process"

.EXAMPLE
Write-Section -Text "Line 1`nLine 2" -Color Green
#>
    param (
        [Parameter(Mandatory)]
        [string]$Text,

        [string]$Color = "White"
    )

    $lines = $Text -split "`n"
    $maxLength = ($lines | Measure-Object -Property Length -Maximum).Maximum
    $border = "#" * ($maxLength + 4)

    Write-Host ""
    Write-Host $border -ForegroundColor $Color
    foreach ($line in $lines) {
        Write-Host ("# " + $line.PadRight($maxLength) + " #") -ForegroundColor $Color
    }
    Write-Host $border -ForegroundColor $Color
    Write-Host ""
}

function Set-RegistryValueForced {
    <#
.SYNOPSIS
Adds or updates a registry value with error handling.

.DESCRIPTION
Creates the specified registry key if it does not exist and sets the provided value.
Supports String, DWord, QWord, Binary, and MultiString types.
Outputs an error message if the operation fails.

.PARAMETER Path
The full registry path (e.g., HKLM:\Software\Example).

.PARAMETER Name
The name of the registry value to create or update.

.PARAMETER Type
The type of the registry value (String, DWord, QWord, Binary, MultiString).

.PARAMETER Value
The value to set. For MultiString, provide an array of strings.

.EXAMPLE
Set-RegistryValueForced -Path "HKLM:\Software\Test" -Name "TestValue" -Type String -Value "OK"

.EXAMPLE
Set-RegistryValueForced -Path "HKLM:\Software\Test" -Name "Flags" -Type DWord -Value 1
#>

    param (
        [string]$Path,
        [string]$Name,
        [string]$Type,
        [object]$Value
    )

    try {
        # Ensure the key exists
        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }

        # Set the registry value
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force
    } catch {
        Write-Output "Failed to set $Name in ${Path}: $($_.Exception.Message)"
    }
}

# Step 1: Clear old upgrade failure records
Write-Host "Step 1: Clearing old upgrade failure records..." -ForegroundColor Yellow
Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\CompatMarkers" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Shared" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\TargetVersionUpgradeExperienceIndicators" -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Cleanup complete." -ForegroundColor Green

# Step 2: Simulating hardware compatibility
Write-Host "Step 2: Simulating hardware compatibility..." -ForegroundColor Yellow
Set-RegistryValueForced -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\HwReqChk" -Name "HwReqChkVars" -Type MultiString -Value @(
    "SQ_SecureBootCapable=TRUE",
    "SQ_SecureBootEnabled=TRUE",
    "SQ_TpmVersion=2",
    "SQ_RamMB=8192"
)
Write-Host "Hardware compatibility values applied." -ForegroundColor Green

# Step 3: Allow upgrades on unsupported TPM or CPU
Write-Host "Step 3: Allowing upgrades on unsupported TPM or CPU..." -ForegroundColor Yellow
Set-RegistryValueForced -Path "HKLM:\SYSTEM\Setup\MoSetup" -Name "AllowUpgradesWithUnsupportedTPMOrCPU" -Type DWord -Value 1
Write-Host "Upgrade policy for unsupported hardware enabled." -ForegroundColor Green

# Step 4: Set Upgrade Eligibility flag in HKCU
Write-Host "Step 4: Setting upgrade eligibility flag..." -ForegroundColor Yellow
Set-RegistryValueForced -Path "HKCU:\Software\Microsoft\PCHC" -Name "UpgradeEligibility" -Type DWord -Value 1
Write-Host "Eligibility flag set." -ForegroundColor Green

# Done
Write-Section -Text "All operations completed successfully!`nYou can now upgrade using the Windows 11 Upgrade Assistant or setup.exe from installation media.`nNo restart required." -Color Cyan

