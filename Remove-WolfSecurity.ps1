#Requires -RunAsAdministrator
<#
    Remove-WolfSecurity.ps1
    Diagnoses HP Wolf Pro Security state, runs the SUPPORTED uninstall sequence,
    and removes the Defender exclusions Wolf leaves behind.

    This does NOT brute-force kill Wolf's drivers/services. Wolf hooks the boot
    path, and forcing it out of order leaves the machine with boot-security errors.
    Use HP's official uninstaller for the actual removal; this orchestrates it safely.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$AttemptUninstall   # off by default - diagnose first
)

$ErrorActionPreference = 'Stop'
function Write-Step { param($m) Write-Host "[*] $m" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!] $m" -ForegroundColor Yellow }

# --- 1. Detect what's actually installed -----------------------------------
Write-Step "Detecting HP Wolf / Bromium components"

$bromiumKey = 'HKLM:\SOFTWARE\Bromium\vSentry'
if (Test-Path $bromiumKey) {
    $versions = Get-ChildItem $bromiumKey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName
    Write-Warn "Bromium vSentry versions registered: $($versions -join ', ')"
    if ($versions.Count -gt 1) {
        Write-Warn "Multiple versions present - this is the usual cause of the"
        Write-Warn "'uninstall of active version not supported' error and the missing"
        Write-Warn "Add/Remove entry. The HP uninstall tool resolves this; manually you"
        Write-Warn "would prefix the STALE version key names with 'old' before retrying."
    }
} else {
    Write-Ok "No Bromium vSentry key found."
}

# Enumerate from the uninstall hive (where Add/Remove reads), incl. orphans
$uninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$wolf = Get-ItemProperty $uninstallPaths -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match 'HP Wolf|HP Sure Click|Bromium|HP Security Update' } |
    Select-Object DisplayName, DisplayVersion, UninstallString, QuietUninstallString

if ($wolf) {
    Write-Step "Registered Wolf components (note missing/blank UninstallString = orphaned MSI):"
    $wolf | Format-List
} else {
    Write-Warn "Nothing in the Uninstall hive - entry is orphaned. Use HP's uninstall tool."
}

# --- 2. Locate HP's official uninstaller -----------------------------------
Write-Step "Searching for HP's official uninstall tool"

$candidates = Get-ChildItem -Path 'C:\Program Files*\HP\*','C:\Program Files*\Bromium\*' `
    -Recurse -Include 'HPWolfSecurityUninstall.exe','BrManage.exe','setup.exe','uninstall.exe' `
    -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName

if ($candidates) {
    Write-Ok "Possible uninstaller(s) found:"
    $candidates | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
} else {
    Write-Warn "No local uninstaller found. Download the HP Wolf Pro Security"
    Write-Warn "Uninstall tool from the Wolf Pro Security console or HP support."
}

# --- 3. Supported uninstall sequence (opt-in) ------------------------------
if ($AttemptUninstall) {
    Write-Step "Running supported uninstall sequence (correct order, with reboots)"
    Write-Warn "Order: 1) HP Sure Click / Wolf Pro Security  2) Wolf Console  3) Security Update Service"

    $order = 'HP Wolf Pro Security|HP Sure Click','HP Wolf Security - Console','HP Security Update Service'
    foreach ($pattern in $order) {
        $target = $wolf | Where-Object { $_.DisplayName -match $pattern } | Select-Object -First 1
        if ($target -and $target.QuietUninstallString) {
            if ($PSCmdlet.ShouldProcess($target.DisplayName, 'Uninstall')) {
                Write-Ok "Uninstalling $($target.DisplayName)"
                cmd /c $target.QuietUninstallString
            }
        } elseif ($target) {
            Write-Warn "$($target.DisplayName): no quiet string - run HP's tool for this one."
        }
    }
    Write-Warn "REBOOT after this completes before verifying. Wolf removes boot components on restart."
}

# --- 4. Clean up Defender exclusions Wolf leaves behind --------------------
Write-Step "Checking Defender exclusions left by Wolf (NOT removed by its uninstaller)"

if (Get-Command Get-MpPreference -ErrorAction SilentlyContinue) {
    $pref = Get-MpPreference
    $suspect = '(?i)bromium|hp wolf|sure click|vSentry'

    $exPath = $pref.ExclusionPath      | Where-Object { $_ -match $suspect }
    $exProc = $pref.ExclusionProcess   | Where-Object { $_ -match $suspect }
    $exExt  = $pref.ExclusionExtension | Where-Object { $_ -match $suspect }

    if ($exPath -or $exProc -or $exExt) {
        Write-Warn "Wolf-related Defender exclusions found:"
        ($exPath + $exProc + $exExt) | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
        foreach ($p in $exPath) {
            if ($PSCmdlet.ShouldProcess($p, 'Remove path exclusion')) {
                Remove-MpPreference -ExclusionPath $p -ErrorAction SilentlyContinue
            }
        }
        foreach ($p in $exProc) {
            if ($PSCmdlet.ShouldProcess($p, 'Remove process exclusion')) {
                Remove-MpPreference -ExclusionProcess $p -ErrorAction SilentlyContinue
            }
        }
        foreach ($p in $exExt) {
            if ($PSCmdlet.ShouldProcess($p, 'Remove ext exclusion')) {
                Remove-MpPreference -ExclusionExtension $p -ErrorAction SilentlyContinue
            }
        }
        Write-Ok "Reviewed/removed Wolf exclusions. Verify with Get-MpPreference."
    } else {
        Write-Ok "No Wolf-related exclusions detected."
    }
} else {
    Write-Warn "Get-MpPreference unavailable."
}

Write-Step "Done. Run your Enable-Defender hardening script AFTER the reboot."