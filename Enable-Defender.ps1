#Requires -RunAsAdministrator
<#
    Enable-Defender.ps1
    Ensures Microsoft Defender Antivirus is enabled and running.
    Clears known policy keys that disable it, enables the firewall and Security Center,
    and reports any LIVE third-party AV registered with the system.

    Does NOT force-remove third-party AV. The third-party check decodes productState
    so it only flags products that are actually enabled - tombstoned/stale registrations
    (e.g. a half-removed product left in SecurityCenter2) are ignored.
#>

[CmdletBinding(SupportsShouldProcess)]
param()

$ErrorActionPreference = 'Stop'

function Write-Step { param($m) Write-Host "[*] $m" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!] $m" -ForegroundColor Yellow }

# --- 1. Clear policy keys that disable Defender ----------------------------
Write-Step "Clearing Defender-disabling policy keys"

$policyKeys = @(
    @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'DisableAntiSpyware' },
    @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender'; Name = 'DisableAntiVirus' },
    @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableRealtimeMonitoring' },
    @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableBehaviorMonitoring' },
    @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableOnAccessProtection' },
    @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name = 'DisableScanOnRealtimeEnable' }
)

foreach ($k in $policyKeys) {
    if (Test-Path $k.Path) {
        $existing = Get-ItemProperty -Path $k.Path -Name $k.Name -ErrorAction SilentlyContinue
        if ($null -ne $existing) {
            if ($PSCmdlet.ShouldProcess("$($k.Path)\$($k.Name)", 'Remove')) {
                Remove-ItemProperty -Path $k.Path -Name $k.Name -Force -ErrorAction SilentlyContinue
                Write-Ok "Removed $($k.Name)"
            }
        }
    }
}

# --- 2. Restore Defender service start values ------------------------------
Write-Step "Ensuring Defender services are set to start"

$services = @{
    'WinDefend'             = 2   # Microsoft Defender Antivirus Service (Automatic)
    'WdNisSvc'              = 3   # Network Inspection Service (Manual)
    'Sense'                 = 3   # Defender for Endpoint (Manual; harmless if absent)
    'SecurityHealthService' = 2
    'wscsvc'                = 2   # Security Center (Automatic)
}

foreach ($svc in $services.GetEnumerator()) {
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$($svc.Key)"
    if (Test-Path $regPath) {
        if ($PSCmdlet.ShouldProcess($svc.Key, "Set Start=$($svc.Value)")) {
            Set-ItemProperty -Path $regPath -Name 'Start' -Value $svc.Value -ErrorAction SilentlyContinue
            Write-Ok "$($svc.Key) start value set"
        }
    }
}

# --- 3. Start the services -------------------------------------------------
Write-Step "Starting Security Center and Defender services"

foreach ($name in 'wscsvc','WinDefend','SecurityHealthService') {
    try {
        $s = Get-Service -Name $name -ErrorAction Stop
        if ($s.Status -ne 'Running') {
            if ($PSCmdlet.ShouldProcess($name, 'Start')) {
                Start-Service -Name $name -ErrorAction SilentlyContinue
            }
        }
        Write-Ok "$name : $((Get-Service $name).Status)"
    } catch { Write-Warn "$name not present or could not start" }
}

# --- 4. Re-enable Defender real-time settings ------------------------------
Write-Step "Re-enabling Defender protection settings"

if (Get-Command Set-MpPreference -ErrorAction SilentlyContinue) {
    if ($PSCmdlet.ShouldProcess('Defender', 'Enable real-time protection')) {
        Set-MpPreference -DisableRealtimeMonitoring $false     -ErrorAction SilentlyContinue
        Set-MpPreference -DisableBehaviorMonitoring $false     -ErrorAction SilentlyContinue
        Set-MpPreference -DisableIOAVProtection $false         -ErrorAction SilentlyContinue
        Set-MpPreference -DisableScriptScanning $false         -ErrorAction SilentlyContinue
        Set-MpPreference -MAPSReporting Advanced               -ErrorAction SilentlyContinue
        Set-MpPreference -SubmitSamplesConsent SendSafeSamples -ErrorAction SilentlyContinue
        Write-Ok "Real-time protection settings applied"
        Write-Warn "If Tamper Protection is ON, some of the above are expected to no-op (by design)."
    }
} else {
    Write-Warn "Set-MpPreference unavailable (Defender feature may be removed)"
}

# --- 5. Enable Windows Firewall --------------------------------------------
Write-Step "Enabling Windows Firewall (all profiles)"

if ($PSCmdlet.ShouldProcess('Firewall', 'Enable all profiles')) {
    Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True -ErrorAction SilentlyContinue
    Write-Ok "Firewall profiles enabled"
}

# --- 6. Report Defender status ---------------------------------------------
Write-Step "Current Defender status"

if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
    Get-MpComputerStatus |
        Select-Object AMRunningMode, AMServiceEnabled, AntivirusEnabled,
                      RealTimeProtectionEnabled, AntispywareEnabled, IsTamperProtected |
        Format-List
}

# --- 7. Report (do NOT remove) LIVE third-party AV -------------------------
# productState is a bitmask. The "enabled" bit is 0x1000. Stale/tombstoned
# registrations (e.g. a half-removed product) have that bit clear and a frozen
# timestamp; we skip those so the script only warns about AV that is actually live.
Write-Step "Checking for LIVE third-party antivirus registered with Security Center"

try {
    $avProducts = Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop |
        Where-Object {
            $_.displayName -notmatch 'Windows Defender|Microsoft Defender' -and
            ([int]$_.productState -band 0x1000)   # only entries whose "enabled" bit is set
        }

    if ($avProducts) {
        Write-Warn "LIVE third-party AV detected. Defender will stay passive while these are active:"
        $avProducts | ForEach-Object {
            $stateHex = ('0x{0:X}' -f [int]$_.productState)
            Write-Host "    - $($_.displayName)  (productState $stateHex, reported $($_.timestamp))" -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "    To remove one cleanly:" -ForegroundColor Gray
        Write-Host "      1. Use the vendor's official removal/uninstall tool, or" -ForegroundColor Gray
        Write-Host "      2. Get the uninstall string and run it interactively:" -ForegroundColor Gray
        Write-Host '         Get-CimInstance Win32_Product | Where-Object Name -match "<vendor>" | Select Name,IdentifyingNumber' -ForegroundColor Gray
        Write-Host "      Disable the vendor's tamper protection in its own console first." -ForegroundColor Gray
        Write-Host "    Defender automatically reactivates once the third-party AV is gone." -ForegroundColor Gray
    } else {
        Write-Ok "No LIVE third-party AV registered. Defender is the active provider."

        # Surface stale/tombstoned entries as informational only (no action needed).
        $stale = Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction SilentlyContinue |
            Where-Object {
                $_.displayName -notmatch 'Windows Defender|Microsoft Defender' -and
                -not ([int]$_.productState -band 0x1000)
            }
        if ($stale) {
            Write-Warn "Note: stale/disabled AV registrations are present (cosmetic, not blocking Defender):"
            $stale | ForEach-Object {
                $stateHex = ('0x{0:X}' -f [int]$_.productState)
                Write-Host "    - $($_.displayName)  (productState $stateHex, reported $($_.timestamp))" -ForegroundColor Gray
            }
            Write-Host "    These usually age out, or clear after 'Restart-Service wscsvc -Force' or a reboot." -ForegroundColor Gray
        }
    }
} catch {
    Write-Warn "Could not query SecurityCenter2 (normal on Server SKUs)."
}

Write-Step "Done."