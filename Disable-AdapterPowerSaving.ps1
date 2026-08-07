#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Sets the High Performance power plan, disables hibernate, disables all
    adapter power-saving settings, and enables Wake-on-LAN on supported adapters.

.DESCRIPTION
    Power Plan
    - Activates the built-in High Performance plan (8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c)
    - Disables hibernate (powercfg -hibernate off)
    - Sets Wireless Adapter power saving to Maximum Performance (AC + DC)

    Per Adapter
    - Disables "Allow the computer to turn off this device to save power" via PnP (PnPCapabilities = 24)
    - Disables EEE, ULP, Selective Suspend, and all idle/low-power features
    - Explicitly ENABLES WakeOnMagicPacket and WakeOnPattern where supported
    - Sets Wi-Fi adapters to CAM (Constantly Awake Mode) via PowerSaveMode = 0

.NOTES
    Run as Administrator. Reboot recommended after running.
    Tested on Windows 10/11.
#>

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
$LogFile = Join-Path $PSScriptRoot ("Disable-AdapterPowerSaving_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

$HighPerfPlanGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"  # Built-in High Performance
$WifiSubGroup     = "19cbb8fa-5279-450e-9fac-8a3d5fedd0c1"  # Wireless Adapter Settings
$WifiSetting      = "12bbebe6-58d6-4636-95bb-3217ef867c1a"  # Power Saving Mode (0 = Max Performance)

# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    $Color = switch ($Level) {
        "ERROR" { "Red"    }
        "WARN"  { "Yellow" }
        default { "Cyan"   }
    }
    Write-Host $Entry -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $Entry
}

# --------------------------------------------------------------------------
# Properties to DISABLE (set to "0" unless overridden)
# --------------------------------------------------------------------------
$PowerSavingProperties = @(
    "AutoPowerSaveModeEnabled",
    "PowerSavingMode",
    "EnablePME",
    "EEELinkAdvertisement",     # Energy Efficient Ethernet (802.3az)
    "EEE",
    "ReduceSpeedOnPowerDown",
    "ULPMode",                  # Ultra Low Power
    "SelectiveSuspend",
    "SelectiveSuspendStatus",
    "bLowPowerEnable",
    "LPSE",                     # Low Power Single Ended
    "GigaLite",                 # Realtek reduced-power Gigabit
    "PnPCapabilities",          # 24 = disable Windows power management for device
    "PowerSaveMode",            # Wi-Fi: 0 = CAM (Constantly Awake Mode)
    "PowerSavingLevel",
    "MIMOPowerSaveMode",
    "MIMOPowerSave",
    "uAPSDSupport",
    "RoamAggressiveness"
)

$PropertyValueOverrides = @{
    "PnPCapabilities" = "24"    # 24 = fully disable Windows power management
}

# --------------------------------------------------------------------------
# WoL properties to ENABLE (set to "1")
# --------------------------------------------------------------------------
$WolProperties = @(
    "WakeOnMagicPacket",
    "WakeOnPattern",
    "PMWakeOnMagicPacket",
    "PMWakeOnPattern",
    "WakeOnLink"
)

# ==========================================================================
# STEP 0 - Power Plan
# ==========================================================================
Write-Log "===== Disable Network Adapter Power Saving + Enable WoL ====="
Write-Log "Log file: $LogFile"
Write-Log ""
Write-Log "STEP 0 - Configuring Power Plan"
Write-Log "----------------------------------------------"

# Verify the High Performance plan exists before activating
$planExists = (powercfg /list) -match $HighPerfPlanGuid
if ($planExists) {
    powercfg /setactive $HighPerfPlanGuid
    Write-Log "SET  Active power plan -> High Performance ($HighPerfPlanGuid)"
}
else {
    Write-Log "WARN High Performance plan not found on this machine. Power plan unchanged." "WARN"
}

# Disable hibernate - frees hiberfil.sys and prevents sleep-related adapter drops
powercfg -hibernate OFF
Write-Log "SET  Hibernate -> OFF"

# Apply Wi-Fi Maximum Performance to the now-active plan via SCHEME_CURRENT
# No PowerShell equivalent for these powercfg index commands
powercfg /SETACVALUEINDEX SCHEME_CURRENT $WifiSubGroup $WifiSetting 0
powercfg /SETDCVALUEINDEX SCHEME_CURRENT $WifiSubGroup $WifiSetting 0
powercfg /setactive SCHEME_CURRENT
Write-Log "SET  Wireless Adapter Power Saving Mode -> Maximum Performance (AC + DC)"

# ==========================================================================
# STEP 1-4 - Network Adapters
# ==========================================================================
Write-Log ""
Write-Log "STEP 1-4 - Processing Network Adapters"
Write-Log "----------------------------------------------"

$Adapters = Get-NetAdapter | Where-Object {
    $_.HardwareInterface -eq $true -and $_.InterfaceType -ne 24
} | Sort-Object -Property MediaType, Name

if (-not $Adapters) {
    Write-Log "No physical network adapters found. Exiting." "WARN"
    exit 1
}

Write-Log ("Found {0} physical adapter(s)." -f $Adapters.Count)

foreach ($Adapter in $Adapters) {
    $isWifi    = $Adapter.MediaType -eq "802.11" -or $Adapter.Name -match "Wi[-\s]?Fi|Wireless|WLAN|802\.11"
    $typeLabel = if ($isWifi) { "Wi-Fi" } else { "Ethernet/Other" }

    Write-Log ""
    Write-Log "----------------------------------------------"
    Write-Log ("Adapter : {0}" -f $Adapter.Name)
    Write-Log ("Type    : {0}" -f $typeLabel)
    Write-Log ("Status  : {0}" -f $Adapter.Status)

    $AdvProps = Get-NetAdapterAdvancedProperty -Name $Adapter.Name -ErrorAction SilentlyContinue

    # ------------------------------------------------------------------
    # 1. Disable power-saving advanced properties
    # ------------------------------------------------------------------
    Write-Log "  [1] Disabling power-saving properties..."
    foreach ($PropName in $PowerSavingProperties) {
        $Prop = $AdvProps | Where-Object { $_.RegistryKeyword -eq $PropName }
        if ($Prop) {
            $TargetValue = if ($PropertyValueOverrides.ContainsKey($PropName)) {
                $PropertyValueOverrides[$PropName]
            }
            else {
                "0"
            }

            if ($Prop.RegistryValue -ne $TargetValue) {
                try {
                    Set-NetAdapterAdvancedProperty -Name $Adapter.Name `
                        -RegistryKeyword $PropName -RegistryValue $TargetValue -ErrorAction Stop
                    Write-Log ("      SET  {0} -> {1}" -f $PropName, $TargetValue)
                }
                catch {
                    Write-Log ("      FAIL {0}: {1}" -f $PropName, $_.Exception.Message) "WARN"
                }
            }
            else {
                Write-Log ("      OK   {0} already = {1}" -f $PropName, $TargetValue)
            }
        }
    }

    # ------------------------------------------------------------------
    # 2. Enable WoL where supported
    # ------------------------------------------------------------------
    Write-Log "  [2] Enabling Wake-on-LAN (where supported)..."
    $wolFound = $false

    foreach ($PropName in $WolProperties) {
        $Prop = $AdvProps | Where-Object { $_.RegistryKeyword -eq $PropName }
        if ($Prop) {
            $wolFound = $true
            if ($Prop.RegistryValue -ne "1") {
                try {
                    Set-NetAdapterAdvancedProperty -Name $Adapter.Name `
                        -RegistryKeyword $PropName -RegistryValue "1" -ErrorAction Stop
                    Write-Log ("      SET  {0} -> 1 (Enabled)" -f $PropName)
                }
                catch {
                    Write-Log ("      FAIL {0}: {1}" -f $PropName, $_.Exception.Message) "WARN"
                }
            }
            else {
                Write-Log ("      OK   {0} already = 1 (Enabled)" -f $PropName)
            }
        }
    }

    try {
        $pmCaps    = Get-NetAdapterPowerManagement -Name $Adapter.Name -ErrorAction Stop
        $setParams = @{ Name = $Adapter.Name; ErrorAction = "Stop" }
        $wolCmdlet = $false

        if ($pmCaps.WakeOnMagicPacket -ne "Unsupported") {
            $setParams["WakeOnMagicPacket"] = "Enabled"
            $wolCmdlet = $true
        }
        if ($pmCaps.WakeOnPattern -ne "Unsupported") {
            $setParams["WakeOnPattern"] = "Enabled"
            $wolCmdlet = $true
        }

        if ($wolCmdlet) {
            Set-NetAdapterPowerManagement @setParams
            Write-Log "      SET  NetAdapterPowerManagement WoL -> Enabled (supported capabilities only)"
            $wolFound = $true
        }
    }
    catch {
        # Adapter does not support this cmdlet - not an error
    }

    if (-not $wolFound) {
        Write-Log "      INFO This adapter does not advertise WoL support - skipped." "WARN"
    }

    # ------------------------------------------------------------------
    # 3. PnP - disable "Allow computer to turn off this device"
    # ------------------------------------------------------------------
    Write-Log "  [3] Disabling PnP power management (Device Manager checkbox)..."
    try {
        $PnpDev = Get-PnpDevice | Where-Object {
            $_.FriendlyName -eq $Adapter.InterfaceDescription -or
            $_.FriendlyName -like "*$($Adapter.InterfaceDescription)*"
        } | Select-Object -First 1

        if ($PnpDev) {
            Write-Log ("      Device: {0}" -f $PnpDev.FriendlyName)
            $RegPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($PnpDev.InstanceId)"
            if (Test-Path $RegPath) {
                Set-ItemProperty -Path $RegPath -Name "PnPCapabilities" -Value 24 -Type DWord -Force
                Write-Log "      SET  PnPCapabilities = 24 -> power management disabled"
            }
            else {
                Write-Log ("      SKIP Registry path not found: $RegPath") "WARN"
            }
        }
        else {
            Write-Log "      SKIP No matching PnP device found." "WARN"
        }
    }
    catch {
        Write-Log ("      ERROR: {0}" -f $_.Exception.Message) "ERROR"
    }

    # ------------------------------------------------------------------
    # 4. Wi-Fi: CAM mode (power plan Wi-Fi setting handled in Step 0)
    # ------------------------------------------------------------------
    if ($isWifi) {
        Write-Log "  [4] Applying Wi-Fi CAM mode..."
        try {
            Set-NetAdapterAdvancedProperty -Name $Adapter.Name `
                -RegistryKeyword "PowerSaveMode" -RegistryValue "0" -ErrorAction SilentlyContinue
            Write-Log "      SET  PowerSaveMode -> 0 (CAM / Constantly Awake Mode)"
        }
        catch {
            Write-Log ("      WARN: {0}" -f $_.Exception.Message) "WARN"
        }
    }

    Write-Log ("  Done: {0}" -f $Adapter.Name)
}

# ==========================================================================
# Summary
# ==========================================================================
Write-Log ""
Write-Log "===== Complete ====="
Write-Log "Power plan: High Performance | Hibernate: OFF | WoL: Enabled where supported"
Write-Log "A reboot is recommended for PnP registry changes to take full effect."
Write-Log "Log saved to: $LogFile"
