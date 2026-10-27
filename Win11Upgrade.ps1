<#
.SYNOPSIS
    Altec - Windows 11 in-place upgrade: download media, stage, finalize, post-upgrade.

.DESCRIPTION
    Two-phase silent in-place upgrade using setup.exe /SkipFinalize and /Finalize.
      Stage    = all the slow work (compat, file copy, image prep) while the user keeps working.
      Finalize = short step that commits the staged upgrade, then a controlled reboot.

    Modes (default is Check, which is read-only):
      Check        Read-only readiness checks (TPM, Secure Boot, UEFI, disk, RAM, pending reboot, etc.)
      Download     Build local media. Default source is the Microsoft media catalog (products.cab):
                   finds the newest ESD for this device's language/arch/edition type, downloads it
                   directly from Microsoft, verifies the catalog hash, and builds media with DISM.
                   No ISO needed. Skips everything if the device is already on the newest build.
      Scan         Read-only setup.exe compatibility scan (/compat scanonly) against the local media
      Stage        Quick readiness + catalog check (seconds), then hands off to a SYSTEM scheduled task
                   (StageWorker) that downloads, builds media and runs setup.exe /SkipFinalize. The RMM
                   job needs no long timeout. The worker updates the Ninja status field as it goes:
                   Queued, Downloading, Building media, Staging, Staged (or Stage failed / Current).
                   Safe to schedule broadly: current devices exit 0, devices failing readiness exit 10.
      Upgrade      One-shot alternative to Stage + Finalize: same checks and background download/media
                   build, then a single setup run (no /SkipFinalize) and an automatic reboot with the
                   standard warning as soon as setup finishes. Field: Queued, Downloading, Building media,
                   Upgrading, Upgraded - rebooting. The reboot time is not controlled; use Stage + Finalize
                   when the reboot must land in a maintenance window.
      Finalize     setup.exe /Finalize, registers the post-upgrade task, schedules the reboot.
                   Devices that are not staged are skipped with exit 0, so it is safe to schedule broadly.
      PostUpgrade  Run by the startup task after the upgrade: re-applies OOBE suppression, cleans up
      Status       Read-only report of where this device is in the process
      Cleanup      Removes Altec media, state and scheduled tasks (does not touch $WINDOWS.~BT)
      StageWorker  Internal. Run by the Stage scheduled task; do not run by hand.

    Self-contained: all settings live in the CONFIGURATION block below. No external files.
    Setting resolution order: CLI parameter, then NinjaOne script variable (environment variable),
    then the CONFIGURATION block.
      NinjaOne script variables:  mode (dropdown), force (checkbox), catalogUrl (text, optional override),
                                  migrateDrivers (dropdown All/None, optional per-run override for Stage)

.EXIT CODES
     0  Success, nothing to do, or in progress (see Ninja status field for detail)
     1  Unhandled error
     2  Configuration error
    10  Readiness check failed
    11  Compatibility scan found a blocking issue
    20  Catalog, download or hash verification failed
    21  Media problem (no install image, edition or language mismatch)
    30  Stage failed
    40  Finalize failed
    41  Staged files missing (staged state recorded but $WINDOWS.~BT or media gone)
    50  Post-upgrade failed or upgrade not detected (likely rollback)

.NOTES
    Altec Solutions Group, Inc.
    Pilot on a few devices before fleet use. setup.exe exit codes for /SkipFinalize are
    configurable (Setup.StageSuccessCodes) so you can adjust without editing code.
#>

[CmdletBinding()]
param(
    [string]$Mode,
    [string]$CatalogUrl,
    [string]$MigrateDrivers,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ---------------------------------------------------------------------------
# Setting resolution: CLI -> Ninja env var -> default
# ---------------------------------------------------------------------------
function Resolve-Setting {
    param($CliValue, [string]$EnvName, $Default)
    if ($null -ne $CliValue -and "$CliValue".Trim() -ne '') { return "$CliValue".Trim() }
    $ev = [Environment]::GetEnvironmentVariable($EnvName)
    if ($null -ne $ev -and $ev.Trim() -ne '') { return $ev.Trim() }
    return $Default
}

$ValidModes = @('Check','Download','Scan','Stage','Upgrade','StageWorker','Finalize','PostUpgrade','Status','Cleanup')
$RawMode = Resolve-Setting -CliValue $Mode -EnvName 'mode' -Default 'Check'
# Tolerate preset-parameter styles such as "mode=Stage", "-Mode Stage" or "Mode: Stage".
$m = [regex]::Match("$RawMode", '^(?i)\s*-?mode\s*[=: ]\s*([A-Za-z]+)')
if ($m.Success) { $RawMode = $m.Groups[1].Value }
$Mode = ($ValidModes | Where-Object { $_ -ieq "$RawMode".Trim() } | Select-Object -First 1)
if (-not $Mode) {
    Write-Output "ERROR: Invalid mode '$RawMode'. Valid modes: $($ValidModes -join ', ')"
    exit 2
}


if ($PSBoundParameters.ContainsKey('Force')) {
    $ForceFlag = [bool]$Force
} else {
    $ForceFlag = ((Resolve-Setting -CliValue $null -EnvName 'force' -Default 'false') -in @('true','1','yes'))
}

# ===========================================================================
# CONFIGURATION - edit this block. Everything the script needs is here.
# ===========================================================================
$Config = @{
    WorkRoot = 'C:\ProgramData\Altec\Win11Upgrade' 
    Esd = @{
        # Catalog lookup order:
        #   1. Microsoft Update metadata service (what the current Media Creation Tool uses; returns the
        #      live catalog, e.g. 26H2). Undocumented endpoint, so it can change without notice.
        #   2. Static catalog link below - ONLY if AllowStaticFallback is true. Microsoft stopped updating
        #      it after 24H2, so falling back would silently target old media (e.g. 26100.4349).
        # -CatalogUrl / Ninja var catalogUrl skips step 1 and uses only that URL.
        CatalogApiUrl       = 'https://fe3.delivery.mp.microsoft.com/UpdateMetadataService/updates/search/v1/bydeviceinfo/'
        CatalogApiOsVersion = '10.0.26100.1'
        CatalogUrl          = 'https://go.microsoft.com/fwlink/?LinkId=2156292'
        UseCatalogApi       = $true
        CatalogApiTimeoutSec = 30
        CatalogApiRetries   = 3
        AllowStaticFallback = $false
        MediaType           = 'Auto'      # Auto | Consumer | Business (Auto = Business for Enterprise editions)
        Language            = 'Auto'      # Auto = device install language, or e.g. en-us
        Architecture        = 'Auto'      # Auto | x64 | ARM64
        ExportEditions      = 'Current'   # Current = only this device's edition (smaller, faster) | All
        DeleteEsdAfterBuild = $true
    }
    Requirements = @{
        MinFreeGB         = 35
        MinRamGB          = 4
        RequireTpm2       = $true
        RequireSecureBoot = $true
        RequireUefi       = $true
        BlockOnPendingReboot = $true
    }
    Setup = @{
        DynamicUpdate          = 'Enable'         # Enable | Disable | NoDrivers | NoLCU | NoDriversNoLCU
        Priority               = 'Normal'         # Normal | Low
        BitLocker              = 'AlwaysSuspend'  # AlwaysSuspend | TryKeepActive | ForceKeepActive
        ShowOOBE               = 'None'           # None | Full
        Telemetry              = 'Disable'        # Enable | Disable
        CompatIgnoreWarning    = $true
        MigrateDrivers         = 'All'            # All | None
        ExtraStageArgs         = ''
        ExtraFinalizeArgs      = ''
        AutoDownload           = $true
        StageTimeoutMinutes    = 180       # setup.exe /SkipFinalize limit (runs in the background task)
        FinalizeTimeoutMinutes = 60
        StageSuccessCodes      = @('0x00000000')
        FinalizeSuccessCodes   = @('0x00000000')
        UpgradeSuccessCodes    = @('0x00000000', '0x00000BC2', '0x80070BC2')   # one-shot Upgrade mode (0xBC2 = reboot required)
    }
    Customization = @{
        UnattendPath                = ''   # local, UNC or https
        PostOOBEScriptPath          = ''   # local, UNC or https (.cmd) -> passed as SetupComplete via /PostOOBE
        SuppressPrivacyExperience   = $true
        SuppressScoobe              = $true
        SuppressFirstLogonAnimation = $true
    }
    Finalize = @{
        NoReboot                    = $false
        RebootDelaySeconds          = 600
        RebootMessage               = 'Altec is completing a scheduled Windows upgrade. Please save your work. Your computer will restart shortly.'
        SuspendBitLockerRebootCount = 2
    }
    PostUpgrade = @{
        MaxBootAttempts       = 3
        CleanupMediaOnSuccess = $true
    }
    Ninja = @{
        StatusField = 'win11UpgradeStatus'   # device Text custom field (script read/write); '' to disable
        CliPath     = 'C:\ProgramData\NinjaRMMAgent\ninjarmm-cli.exe'
    }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function ConvertTo-Hashtable {
    param($InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) { return $InputObject }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $InputObject.PSObject.Properties) { $h[$p.Name] = ConvertTo-Hashtable $p.Value }
        return $h
    }
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        return ,@($InputObject | ForEach-Object { ConvertTo-Hashtable $_ })
    }
    return $InputObject
}

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    # Write-Host (not Write-Output) so log lines never leak into function return values.
    Write-Host $line
}

function ConvertTo-HexCode {
    param([long]$Code)
    if ($Code -lt 0) { $Code = $Code + 4294967296 }
    return ('0x{0:X8}' -f $Code)
}

$SetupCodeMap = @{
    '0x00000000' = 'Success'
    '0xC1900101' = 'Driver or boot-phase failure; upgrade rolled back. Check Panther logs and update drivers/BIOS.'
    '0xC1900107' = 'Cleanup from a previous install attempt is pending. Reboot and retry.'
    '0xC1900132' = '/Finalize ran without a prior /SkipFinalize stage, or stage did not complete.'
    '0xC1900200' = 'Device does not meet minimum Windows 11 requirements.'
    '0xC1900202' = 'Device does not meet minimum Windows 11 requirements.'
    '0xC1900204' = 'Migration choice not available. Usually edition or language mismatch with the media.'
    '0xC1900208' = 'Incompatible app blocking upgrade. See compat XML in setup logs.'
    '0xC190020E' = 'Not enough free disk space.'
    '0xC1900210' = 'Compatibility scan: no issues found.'
    '0x80070003' = 'Path not found. Usually a damaged driver store or component store on the device (see setuperr.log via Status). Repair with DISM/SFC, or retry with migrateDrivers = None.'
}

function Get-CodeDescription {
    param([string]$Hex)
    if ($SetupCodeMap.ContainsKey($Hex)) { return $SetupCodeMap[$Hex] }
    return 'Unknown code. Review setupact.log / setuperr.log in the SetupLogs folder.'
}

$script:NinjaCli = $null
function Get-NinjaCli {
    if ($script:NinjaCli) { return $script:NinjaCli }
    $cands = @($Config.Ninja.CliPath)
    if ($env:NINJA_DATA_PATH) { $cands += (Join-Path $env:NINJA_DATA_PATH 'ninjarmm-cli.exe') }
    $cands += 'C:\ProgramData\NinjaRMMAgent\ninjarmm-cli.exe'
    # Agent install folder (where the NinjaRMMAgent service runs from)
    try {
        $svc = Get-CimInstance Win32_Service -Filter "Name='NinjaRMMAgent'" -ErrorAction Stop
        if ($svc.PathName) {
            $exe = ($svc.PathName -replace '^"([^"]+)".*$', '$1') -replace '^(\S+\.exe).*$', '$1'
            $cands += (Join-Path (Split-Path $exe -Parent) 'ninjarmm-cli.exe')
        }
    } catch { }
    foreach ($c in $cands) { if ($c -and (Test-Path $c)) { $script:NinjaCli = $c; return $c } }
    return $null
}

function Set-NinjaStatus {
    param([string]$Text)
    $field = $Config.Ninja.StatusField
    if (-not $field -or -not $Text) { return }
    $val = '{0} ({1})' -f $Text, (Get-Date -Format 'yyyy-MM-dd HH:mm')
    try {
        $cli = Get-NinjaCli
        if ($cli) {
            $oldEap = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
            try { $out = & $cli set $field $val 2>&1; $code = $LASTEXITCODE } finally { $ErrorActionPreference = $oldEap }
            if ($code -eq 0) { Write-Log "Ninja field $field = $val"; return }
            Write-Log "ninjarmm-cli set failed (exit $code): $(($out | Out-String).Trim())" 'WARN'
        }
        if (Get-Command Ninja-Property-Set -ErrorAction SilentlyContinue) {
            Ninja-Property-Set $field $val
            Write-Log "Ninja field $field = $val (Ninja-Property-Set)"
            return
        }
        Write-Log "Could not update Ninja field $($field): ninjarmm-cli.exe not found and Ninja-Property-Set unavailable." 'WARN'
    } catch { Write-Log "Could not update Ninja field $field : $($_.Exception.Message)" 'WARN' }
}

function Exit-Script {
    param([int]$Code, [string]$Message, [string]$NinjaStatus)
    if ($Message) { Write-Log $Message ($(if ($Code -eq 0) { 'INFO' } else { 'ERROR' })) }
    if ($NinjaStatus) { Set-NinjaStatus $NinjaStatus }
    try { Stop-Transcript | Out-Null } catch { }
    exit $Code
}

function Get-State {
    if (Test-Path $StatePath) {
        try { return (ConvertTo-Hashtable (Get-Content -Path $StatePath -Raw | ConvertFrom-Json)) } catch { }
    }
    return @{}
}

function Save-State {
    param([hashtable]$State)
    $State['Updated'] = (Get-Date).ToString('s')
    $State | ConvertTo-Json -Depth 8 | Set-Content -Path $StatePath -Encoding ASCII
}

function Get-OsBuildInfo {
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $name = "$($cv.ProductName)"
    # ProductName still reads "Windows 10" on Windows 11; correct it by build number.
    if ([int]$cv.CurrentBuildNumber -ge 22000) { $name = $name -replace 'Windows 10', 'Windows 11' }
    return @{
        Build          = [int]$cv.CurrentBuildNumber
        UBR            = [int]$cv.UBR
        DisplayVersion = "$($cv.DisplayVersion)"
        EditionID      = "$($cv.EditionID)"
        ProductName    = $name
    }
}

function Get-FirmwareType {
    # kernel32!GetFirmwareType is authoritative (PEFirmwareType registry value only exists in WinPE).
    try {
        if (-not ('AltecFirmware' -as [type])) {
            Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class AltecFirmware { [DllImport("kernel32.dll")] public static extern bool GetFirmwareType(ref uint FirmwareType); }'
        }
        [uint32]$t = 0
        if ([AltecFirmware]::GetFirmwareType([ref]$t)) {
            switch ($t) { 1 { return 'BIOS' } 2 { return 'UEFI' } }
        }
    } catch { }
    # Fallback: Confirm-SecureBootUEFI throws on legacy BIOS, returns true/false on UEFI.
    try { Confirm-SecureBootUEFI | Out-Null; return 'UEFI' } catch [System.PlatformNotSupportedException] { return 'BIOS' } catch { return 'Unknown' }
}

function Get-InstallLanguage {
    try {
        $lcid = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language').InstallLanguage
        return ([System.Globalization.CultureInfo]::new([Convert]::ToInt32($lcid, 16))).Name
    } catch { return '' }
}

function Invoke-FileDownload {
    param([string]$Url, [string]$Destination)
    # Microsoft's CDN hands out http links; try https first (hashes are verified either way),
    # then the original link. Stalled connections are abandoned instead of hanging for minutes.
    $urls = @($Url)
    if ($Url -match '^http://[^/]*\.delivery\.mp\.microsoft\.com/') { $urls = @(($Url -replace '^http://', 'https://'), $Url) }
    $curl = Join-Path $env:SystemRoot 'System32\curl.exe'
    $lastErr = ''
    foreach ($u in $urls) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        if (Test-Path $curl) {
            $oldEap = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            try {
                & $curl -L --fail --silent --connect-timeout 20 --speed-limit 20480 --speed-time 60 --retry 3 --retry-delay 10 -C - -o $Destination $u
                $code = $LASTEXITCODE
            } finally { $ErrorActionPreference = $oldEap }
            if ($code -eq 0) {
                $len = (Get-Item $Destination).Length
                $size = $(if ($len -ge 1MB) { '{0:N1} MB' -f ($len / 1MB) } else { '{0:N0} KB' -f ($len / 1KB) })
                Write-Log ("Downloaded {0} in {1:N0}s via {2}" -f $size, $sw.Elapsed.TotalSeconds, ([Uri]$u).Scheme)
                return
            }
            $lastErr = "curl exit $code on $(([Uri]$u).Scheme) (6=DNS, 7=connect, 28=timeout/stall, 35/60=TLS, 22=HTTP error)"
            if ($u -ne $urls[-1]) { Write-Log "https not usable for this host ($lastErr); using http. File is still hash-verified." }
            else { Write-Log "Download attempt failed: $lastErr" 'WARN' }
            Remove-Item $Destination -Force -ErrorAction SilentlyContinue
        } else {
            try { Start-BitsTransfer -Source $u -Destination $Destination -ErrorAction Stop; return }
            catch { $lastErr = $_.Exception.Message; Write-Log "BITS download failed: $lastErr" 'WARN' }
        }
    }
    throw "Download failed: $lastErr"
}

function Copy-FromUnc {
    param([string]$Source, [string]$Destination)
    $srcDir = Split-Path $Source -Parent
    $srcFile = Split-Path $Source -Leaf
    $dstDir = Split-Path $Destination -Parent
    & robocopy.exe $srcDir $dstDir $srcFile /Z /R:3 /W:10 /NP /NJH /NJS | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed with exit code $LASTEXITCODE (SYSTEM uses the computer account; share needs Domain Computers read)" }
    $copied = Join-Path $dstDir $srcFile
    if ($copied -ne $Destination) { Move-Item -Path $copied -Destination $Destination -Force }
}

# Pulls a customization asset (local, UNC or https) into WorkRoot\Custom. Returns local path or ''.
function Get-Asset {
    param([string]$Source, [string]$DestName)
    if (-not $Source) { return '' }
    $dest = Join-Path $CustomDir $DestName
    if ($Source -match '^https?://') {
        Invoke-FileDownload -Url $Source -Destination $dest
    } elseif ($Source -like '\\*') {
        Copy-FromUnc -Source $Source -Destination $dest
    } else {
        if (-not (Test-Path $Source)) { throw "Asset not found: $Source" }
        Copy-Item -Path $Source -Destination $dest -Force
    }
    return $dest
}

# Prints the tail of setup's error log and the last StageWorker transcript (read-only).
function Show-StageDiagnostics {
    $cands = @(
        (Join-Path $BtFolder 'Sources\Panther\setuperr.log'),
        (Join-Path $LogsDir 'SetupLogs\Panther\setuperr.log'),
        (Join-Path $LogsDir 'SetupLogs\setuperr.log')
    )
    foreach ($c in $cands) {
        if ((Test-Path $c) -and (Get-Item $c).Length -gt 0) {
            Write-Log "---- last 25 lines of $c ----"
            Get-Content -Path $c -Tail 25 | ForEach-Object { Write-Host "  $_" }
            break
        }
    }
    $w = Get-ChildItem -Path $LogsDir -Filter 'Win11Upgrade_StageWorker_*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($w) {
        Write-Log "---- last 20 lines of $($w.Name) ----"
        Get-Content -Path $w.FullName -Tail 20 | ForEach-Object { Write-Host "  $_" }
    }
}

function Get-InstallImagePath {
    foreach ($n in @('install.wim','install.esd','install.swm')) {
        $p = Join-Path $MediaDir "sources\$n"
        if (Test-Path $p) { return $p }
    }
    return ''
}

function Get-MediaInfo {
    $img = Get-InstallImagePath
    if (-not $img) { return $null }
    $editions = @()
    $build = 0; $ubr = 0; $langs = @()
    foreach ($i in (Get-WindowsImage -ImagePath $img)) {
        $d = Get-WindowsImage -ImagePath $img -Index $i.ImageIndex
        $editions += "$($d.EditionId)"
        if ($build -eq 0) {
            $v = "$($d.Version)".Split('.')
            $build = [int]$v[2]; $ubr = [int]$v[3]
            $langs = @($d.Languages | ForEach-Object { "$_" })
        }
    }
    return @{ ImagePath = $img; Editions = $editions; Build = $build; UBR = $ubr; Languages = $langs }
}

function Get-UninstallMatch {
    param([string]$Pattern)
    $keys = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
              'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    return @(Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like $Pattern })
}

function Set-DwordForAllUsers {
    param([string]$SubKey, [string]$Name, [int]$Value)
    # Loaded user hives
    Get-ChildItem 'Registry::HKEY_USERS' | Where-Object { $_.PSChildName -match '^S-1-5-21-[\d-]+$' } | ForEach-Object {
        $k = "Registry::HKEY_USERS\$($_.PSChildName)\$SubKey"
        if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
        New-ItemProperty -Path $k -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
    }
    # Unloaded profiles + Default user
    $loaded = @(Get-ChildItem 'Registry::HKEY_USERS' | ForEach-Object { $_.PSChildName })
    $targets = @()
    Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object {
        $sid = $_.PSChildName
        if ($sid -match '^S-1-5-21-[\d-]+$' -and $loaded -notcontains $sid) {
            $targets += (Join-Path (Get-ItemProperty $_.PSPath).ProfileImagePath 'NTUSER.DAT')
        }
    }
    $targets += (Join-Path $env:SystemDrive 'Users\Default\NTUSER.DAT')
    foreach ($hive in $targets) {
        if (-not (Test-Path $hive)) { continue }
        $tmp = 'AltecTmpHive'
        & reg.exe load "HKU\$tmp" "$hive" 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) { Write-Log "Skipped hive (locked or unreadable): $hive" 'WARN'; continue }
        try {
            & reg.exe add "HKU\$tmp\$SubKey" /v $Name /t REG_DWORD /d $Value /f 2>&1 | Out-Null
        } finally {
            [GC]::Collect(); Start-Sleep -Milliseconds 500
            & reg.exe unload "HKU\$tmp" 2>&1 | Out-Null
        }
    }
}

function Set-OobeSuppression {
    $c = $Config.Customization
    if ($c.SuppressPrivacyExperience) {
        $k = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE'
        if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
        New-ItemProperty -Path $k -Name 'DisablePrivacyExperience' -Value 1 -PropertyType DWord -Force | Out-Null
        Write-Log 'Set DisablePrivacyExperience policy.'
    }
    if ($c.SuppressFirstLogonAnimation) {
        $k = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
        New-ItemProperty -Path $k -Name 'EnableFirstLogonAnimation' -Value 0 -PropertyType DWord -Force | Out-Null
        Write-Log 'Disabled first logon animation.'
    }
    if ($c.SuppressScoobe) {
        Set-DwordForAllUsers -SubKey 'Software\Microsoft\Windows\CurrentVersion\UserProfileEngagement' -Name 'ScoobeSystemSettingEnabled' -Value 0
        Write-Log 'Disabled "Finish setting up your device" (SCOOBE) for all profiles.'
    }
}

function Register-AltecTask {
    param([string]$Name, [string]$Execute, [string]$Arguments, [string]$WorkingDir, [switch]$AtStartup, [int]$LimitHours = 6)
    $action = New-ScheduledTaskAction -Execute $Execute -Argument $Arguments -WorkingDirectory $WorkingDir
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours $LimitHours) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew
    if ($AtStartup) {
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $trigger.Delay = 'PT3M'
        Register-ScheduledTask -TaskName $Name -TaskPath $TaskPath -Action $action -Principal $principal -Settings $settings -Trigger $trigger -Force | Out-Null
    } else {
        Register-ScheduledTask -TaskName $Name -TaskPath $TaskPath -Action $action -Principal $principal -Settings $settings -Force | Out-Null
    }
}

function Remove-AltecTask {
    param([string]$Name)
    Unregister-ScheduledTask -TaskName $Name -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
}

function Copy-SelfToWorkRoot {
    # The startup task runs a local copy of this script (RMM temp copies get deleted).
    if (-not $PSCommandPath) { throw 'Cannot determine script path to copy for the post-upgrade task.' }
    if ($PSCommandPath -ne $LocalScriptPath) {
        Copy-Item -Path $PSCommandPath -Destination $LocalScriptPath -Force
    }
}

# Optional override of the media catalog location (CLI -> Ninja var -> CONFIGURATION block)
$ovMig = Resolve-Setting -CliValue $MigrateDrivers -EnvName 'migrateDrivers' -Default ''
if ($ovMig -in @('All','None')) { $Config.Setup.MigrateDrivers = $ovMig }
$ovCatalog = Resolve-Setting -CliValue $CatalogUrl -EnvName 'catalogUrl' -Default ''
if ($ovCatalog) { $Config.Esd.CatalogUrl = $ovCatalog; $Config.Esd.UseCatalogApi = $false }

$WorkRoot        = $Config.WorkRoot
$MediaDir        = Join-Path $WorkRoot 'Media'
$CustomDir       = Join-Path $WorkRoot 'Custom'
$LogsDir         = Join-Path $WorkRoot 'Logs'
$StatePath       = Join-Path $WorkRoot 'state.json'
$LocalScriptPath = Join-Path $WorkRoot 'Win11Upgrade.ps1'
$TaskPath        = '\Altec\'
$StageTaskName   = 'Win11Upgrade-Stage'
$PostTaskName    = 'Win11Upgrade-PostUpgrade'
$BtFolder        = Join-Path $env:SystemDrive '$WINDOWS.~BT'

foreach ($d in @($WorkRoot, $MediaDir, $CustomDir, $LogsDir)) {
    if (-not (Test-Path $d)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
}

try { Start-Transcript -Path (Join-Path $LogsDir ("Win11Upgrade_{0}_{1}.log" -f $Mode, (Get-Date -Format 'yyyyMMdd_HHmmss'))) -Force | Out-Null } catch { }

Write-Log "Mode: $Mode | Force: $ForceFlag"

# Must be elevated
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Exit-Script 2 'Must run as SYSTEM or elevated administrator.'
}

# ---------------------------------------------------------------------------
# Readiness (read-only)
# ---------------------------------------------------------------------------
function Invoke-Readiness {
    $r = $Config.Requirements
    $results = New-Object System.Collections.Generic.List[object]
    function Add-Result($Name, $Status, $Detail) { $results.Add([pscustomobject]@{ Check = $Name; Status = $Status; Detail = $Detail }) }

    $os = Get-OsBuildInfo
    Add-Result 'OS' 'Info' "$($os.ProductName) $($os.DisplayVersion) build $($os.Build).$($os.UBR) ($($os.EditionID))"

    $pt = (Get-CimInstance Win32_OperatingSystem).ProductType
    if ($pt -ne 1) { Add-Result 'Client OS' 'Fail' 'This is a server OS.' } else { Add-Result 'Client OS' 'Pass' 'Workstation' }

    $fw = Get-FirmwareType
    $isUefi = ($fw -eq 'UEFI')
    Add-Result 'UEFI' ($(if ($isUefi) { 'Pass' } elseif ($r.RequireUefi) { 'Fail' } else { 'Warn' })) ($(if ($isUefi) { 'UEFI' } elseif ($fw -eq 'BIOS') { 'Legacy BIOS (MBR2GPT + firmware change needed)' } else { 'Could not determine firmware type' }))

    $sb = $false
    try { $sb = [bool](Confirm-SecureBootUEFI) } catch { $sb = $false }
    Add-Result 'Secure Boot' ($(if ($sb) { 'Pass' } elseif ($r.RequireSecureBoot) { 'Fail' } else { 'Warn' })) ($(if ($sb) { 'Enabled' } else { 'Disabled or unsupported' }))

    try {
        $tpm = Get-CimInstance -Namespace 'root/cimv2/security/microsofttpm' -ClassName Win32_Tpm -ErrorAction Stop
        $spec = ("$($tpm.SpecVersion)" -split ',')[0].Trim()
        $ok = ($spec -like '2.*') -and $tpm.IsEnabled_InitialValue
        Add-Result 'TPM' ($(if ($ok) { 'Pass' } elseif ($r.RequireTpm2) { 'Fail' } else { 'Warn' })) "Spec $spec, Enabled=$($tpm.IsEnabled_InitialValue)"
    } catch {
        Add-Result 'TPM' ($(if ($r.RequireTpm2) { 'Fail' } else { 'Warn' })) 'No TPM detected'
    }

    $ramGB = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 1)
    Add-Result 'RAM' ($(if ($ramGB -ge $r.MinRamGB) { 'Pass' } else { 'Fail' })) "$ramGB GB"

    $freeGB = [math]::Round((Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':'))).Free / 1GB, 1)
    Add-Result 'Free space' ($(if ($freeGB -ge $r.MinFreeGB) { 'Pass' } else { 'Fail' })) "$freeGB GB free on $env:SystemDrive (need $($r.MinFreeGB))"

    $pending = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
               (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
    Add-Result 'Pending reboot' ($(if (-not $pending) { 'Pass' } elseif ($r.BlockOnPendingReboot) { 'Fail' } else { 'Warn' })) ($(if ($pending) { 'Reboot pending' } else { 'None' }))

    Add-Result 'CPU' 'Info' "$((Get-CimInstance Win32_Processor | Select-Object -First 1).Name) - CPU support is validated by Scan mode, not here."
    Add-Result 'Install language' 'Info' (Get-InstallLanguage)

    try {
        $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        Add-Result 'BitLocker' 'Info' "Protection $($bl.ProtectionStatus), $($bl.VolumeStatus)"
    } catch { Add-Result 'BitLocker' 'Info' 'Unknown' }

    if (Test-Path $BtFolder) { Add-Result 'Prior setup files' 'Warn' "$BtFolder exists (previous or in-progress attempt)" }

    if ((Get-UninstallMatch 'Duo Authentication for Windows Logon*').Count -gt 0) {
        Add-Result 'Duo Windows Logon' 'Warn' 'Installed. Confirm a Win11-supported version before upgrade. Altec standard: replace with Windows Hello for Business (Cloud Kerberos Trust).'
    }
    if ((Get-UninstallMatch '*AutoElevate*').Count -gt 0) {
        Add-Result 'AutoElevate' 'Warn' 'Installed. Altec standard: replace with Intune Endpoint Privilege Management.'
    }

    Write-Host ($results | Format-Table -AutoSize | Out-String -Width 250)
    return ,$results
}

# ---------------------------------------------------------------------------
# Media
# ---------------------------------------------------------------------------
function Get-HostArch {
    if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64' -or $env:PROCESSOR_ARCHITEW6432 -eq 'ARM64') { return 'ARM64' }
    return 'x64'
}

# Extracts a CAB. expand.exe silently COPIES a cab it cannot parse (exit 0), so each method is
# verified by looking for an extracted .xml, then the next method is tried.
function Expand-CatalogCab {
    param([string]$CabPath, [string]$Destination)
    $sys = Join-Path $env:SystemRoot 'System32'
    $methods = @(
        @{ Name = 'tar.exe';    Run = { & (Join-Path $sys 'tar.exe') -xf $CabPath -C $Destination 2>&1 } },
        @{ Name = 'expand.exe'; Run = { & (Join-Path $sys 'expand.exe') $CabPath "-F:*" $Destination 2>&1 } },
        @{ Name = 'Shell COM';  Run = {
            $sh = New-Object -ComObject Shell.Application
            $sh.Namespace($Destination).CopyHere($sh.Namespace($CabPath).Items(), 0x14)
            Start-Sleep -Seconds 2 } }
    )
    foreach ($m in $methods) {
        Get-ChildItem -Path $Destination -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        try {
            $out = & $m.Run
            if (@(Get-ChildItem -Path $Destination -Recurse -File -Filter '*.xml').Count -gt 0) {
                Write-Log "Catalog extracted with $($m.Name)."
                return
            }
            Write-Log "$($m.Name) did not extract the catalog. $(($out | Out-String).Trim())" 'WARN'
        } catch {
            Write-Log "$($m.Name) failed: $($_.Exception.Message)" 'WARN'
        }
    }
    throw 'Could not extract products.xml from the catalog cab with tar.exe, expand.exe or Shell COM.'
}

# Asks the Microsoft Update metadata service for the current products.cab (same call the current
# Media Creation Tool makes). The CDN link is http, so the file is verified against the size and
# SHA256 returned over https before use.
function Get-CatalogCabFromApi {
    param([string]$Destination)
    $e = $Config.Esd
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $pn = $(if ((Get-HostArch) -eq 'ARM64') { 'Windows.Products.Cab.arm64' } else { 'Windows.Products.Cab.amd64' })
    $body = '{"Products":"PN=' + $pn + '&V=0.0.0.0","DeviceAttributes":"DUScan=1;OSVersion=' + $e.CatalogApiOsVersion + '"}'
    Write-Log "Querying Microsoft catalog service ($pn)"
    $resp = $null
    $errors = @()
    $curl = Join-Path $env:SystemRoot 'System32\curl.exe'
    $bodyFile = Join-Path $WorkRoot 'catalog_request.json'
    [IO.File]::WriteAllText($bodyFile, $body, [Text.Encoding]::ASCII)
    for ($try = 1; $try -le [int]$e.CatalogApiRetries -and -not $resp; $try++) {
        # curl.exe first: no WPAD/proxy auto-detect delay and a hard timeout. Invoke-RestMethod as backup.
        if (Test-Path $curl) {
            # Stderr is not redirected: under Windows PowerShell 5.1 with ErrorAction Stop, 2>&1 on a
            # native command turns any stderr line into a terminating error.
            $oldEap = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            try {
                $raw = & $curl -s --fail --connect-timeout 15 --max-time ([int]$e.CatalogApiTimeoutSec) -X POST `
                    -H 'Content-Type: application/json' -H 'Accept: application/json' --data-binary "@$bodyFile" $e.CatalogApiUrl
                $code = $LASTEXITCODE
            } finally { $ErrorActionPreference = $oldEap }
            if ($code -eq 0) {
                try { $resp = ($raw -join "`n") | ConvertFrom-Json } catch { $errors += "try $try curl: bad JSON" }
            } else {
                $errors += "try $try curl exit $code (6=DNS, 7=connect refused, 28=timeout, 35/60=TLS, 22=HTTP error)"
            }
        }
        if (-not $resp) {
            try {
                $resp = Invoke-RestMethod -Method Post -Uri $e.CatalogApiUrl -ContentType 'application/json' -Headers @{ Accept = 'application/json' } -Body $body -TimeoutSec ([int]$e.CatalogApiTimeoutSec) -UseBasicParsing
            } catch { $errors += "try $try Invoke-RestMethod: $($_.Exception.Message)" }
        }
        if (-not $resp -and $try -lt [int]$e.CatalogApiRetries) {
            Write-Log "Catalog service attempt $try failed; retrying in 10s." 'WARN'
            Start-Sleep -Seconds 10
        }
    }
    Remove-Item $bodyFile -Force -ErrorAction SilentlyContinue
    if (-not $resp) { throw "No response from catalog service. $($errors -join ' | ')" }
    $locs = @($resp.FileLocations | Where-Object { $_.FileName -eq 'products.cab' })
    if ($locs.Count -ne 1) { throw "Service returned $($locs.Count) products.cab entries." }
    $loc = $locs[0]
    $uri = [Uri]$loc.Url
    if ($uri.Host -notlike '*.delivery.mp.microsoft.com') { throw "Unexpected catalog host $($uri.Host)." }
    Write-Log "Catalog service returned $($uri.AbsoluteUri)"
    Remove-Item $Destination -Force -ErrorAction SilentlyContinue
    Invoke-FileDownload -Url $uri.AbsoluteUri -Destination $Destination

    $len = (Get-Item $Destination).Length
    if ($loc.Size -and $len -ne [long]$loc.Size) { throw "Catalog size mismatch ($len vs $($loc.Size))." }
    if ($loc.Digest) {
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $digest = [Convert]::ToBase64String($sha.ComputeHash([IO.File]::ReadAllBytes($Destination))) } finally { $sha.Dispose() }
        if ($digest -cne "$($loc.Digest)") { throw 'Catalog SHA256 does not match the service digest.' }
        Write-Log 'Catalog verified against service digest.'
    } else {
        Write-Log 'Service returned no digest for the catalog; size check only.' 'WARN'
    }
}

# Reads the Microsoft media catalog and returns the newest matching ESD entry.
function Get-EsdCatalogEntry {
    $e = $Config.Esd
    $cabPath = Join-Path $WorkRoot 'products.cab'
    $xmlDir = Join-Path $WorkRoot 'Catalog'
    if (Test-Path $xmlDir) { Remove-Item $xmlDir -Recurse -Force }
    New-Item -Path $xmlDir -ItemType Directory -Force | Out-Null
    Remove-Item $cabPath -Force -ErrorAction SilentlyContinue

    $got = $false
    if ($e.UseCatalogApi) {
        try {
            Get-CatalogCabFromApi -Destination $cabPath
            $got = $true
        } catch {
            if (-not $e.AllowStaticFallback) {
                throw "Catalog service lookup failed: $($_.Exception.Message) Static fallback is disabled because it serves old media. Check that this device can reach fe3.delivery.mp.microsoft.com on 443 (firewall, DNS filter, SSL inspection)."
            }
            Write-Log "Catalog service lookup failed: $($_.Exception.Message). Falling back to static catalog link (old media)." 'WARN'
        }
    }
    if (-not $got) {
        Write-Log "Reading media catalog: $($e.CatalogUrl)"
        Invoke-FileDownload -Url $e.CatalogUrl -Destination $cabPath
    }
    $len = (Get-Item $cabPath).Length
    $hdr = New-Object byte[] 8
    $fs = [System.IO.File]::OpenRead($cabPath)
    try { [void]$fs.Read($hdr, 0, 8) } finally { $fs.Close() }
    $magic = [System.Text.Encoding]::ASCII.GetString($hdr, 0, 4)
    $hex = ($hdr | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
    Write-Log "Catalog downloaded: $len bytes, header $hex"

    if ($magic -eq 'MSCF') {
        Expand-CatalogCab -CabPath $cabPath -Destination $xmlDir
    } elseif ([System.Text.Encoding]::ASCII.GetString($hdr) -match '<') {
        Copy-Item -Path $cabPath -Destination (Join-Path $xmlDir 'products.xml') -Force
    } else {
        throw "Catalog is not a CAB or XML file ($len bytes, header $hex). Check the catalog settings."
    }

    $allFiles = @(Get-ChildItem -Path $xmlDir -Recurse -File)
    $xmlFile = $allFiles | Where-Object { $_.Extension -ieq '.xml' -and (Select-String -Path $_.FullName -Pattern '<MCT' -SimpleMatch -Quiet) } | Select-Object -First 1
    if (-not $xmlFile) {
        throw "No MCT catalog XML found. Files extracted: $(($allFiles.Name) -join ', ')"
    }
    Write-Log "Using catalog file $($xmlFile.Name)"
    [xml]$xml = Get-Content -Path $xmlFile.FullName -Raw
    $files = @($xml.MCT.Catalogs.Catalog.PublishedMedia.Files.File)
    if ($files.Count -eq 0) { throw 'Catalog has no file entries (format may have changed).' }

    $arch = $(if ($e.Architecture -ieq 'Auto') { Get-HostArch } else { $e.Architecture })
    $lang = $(if ($e.Language -ieq 'Auto') { (Get-InstallLanguage).ToLower() } else { "$($e.Language)".ToLower() })
    if (-not $lang) { throw 'Could not determine install language; set Esd.Language.' }
    $type = $e.MediaType
    if ($type -ieq 'Auto') { $type = $(if ((Get-OsBuildInfo).EditionID -like 'Enterprise*') { 'Business' } else { 'Consumer' }) }
    $pattern = $(if ($type -ieq 'Business') { 'CLIENTBUSINESS_VOL' } else { 'CLIENTCONSUMER_RET' })

    $matched = @($files | Where-Object {
        "$($_.Architecture)" -ieq $arch -and "$($_.LanguageCode)" -ieq $lang -and
        "$($_.FileName)" -match $pattern -and "$($_.FilePath)" -match '^https?://'
    })
    $cands = @()
    foreach ($g in ($matched | Group-Object { "$($_.FilePath)" })) {
        $f = $g.Group[0]
        $m = [regex]::Match("$($f.FileName)", '^(\d{5})\.(\d+)\.')
        if (-not $m.Success) { continue }
        $cands += [pscustomobject]@{
            FileName = "$($f.FileName)"; Url = "$($f.FilePath)"
            Sha1 = "$($f.Sha1)"; Sha256 = "$($f.Sha256)"; Size = "$($f.Size)"
            Build = [int]$m.Groups[1].Value; UBR = [int]$m.Groups[2].Value
            Arch = $arch; Lang = $lang; Type = $type
        }
    }
    if ($cands.Count -eq 0) { throw "No catalog entry for $type / $lang / $arch." }
    return ($cands | Sort-Object Build, UBR -Descending | Select-Object -First 1)
}

function Test-CatalogHash {
    param([string]$Path, $Entry)
    if ($Entry.Sha256) { return ((Get-FileHash -Path $Path -Algorithm SHA256).Hash -ieq $Entry.Sha256.Trim()) }
    if ($Entry.Sha1)   { return ((Get-FileHash -Path $Path -Algorithm SHA1).Hash -ieq $Entry.Sha1.Trim()) }
    Write-Log 'Catalog entry has no hash; checking size only.' 'WARN'
    return ($Entry.Size -and ((Get-Item $Path).Length -eq [long]$Entry.Size))
}

# Turns an MCT ESD into a setup media folder:
#   "Windows Setup Media" image -> applied to Media
#   WinPE + Windows Setup images -> sources\boot.wim
#   edition image(s) -> sources\install.wim
function Build-MediaFromEsd {
    param([string]$EsdPath)
    if (Test-Path $MediaDir) { Remove-Item $MediaDir -Recurse -Force }
    New-Item -Path $MediaDir -ItemType Directory -Force | Out-Null

    $images = @(Get-WindowsImage -ImagePath $EsdPath)
    $setupImg = $images | Where-Object { $_.ImageName -match 'Setup Media' } | Select-Object -First 1
    $peImg    = $images | Where-Object { $_.ImageName -match 'Windows PE' } | Select-Object -First 1
    $wsImg    = $images | Where-Object { $_.ImageName -match 'Windows Setup' -and $_.ImageName -notmatch 'Media' } | Select-Object -First 1
    $setupIdx = $(if ($setupImg) { $setupImg.ImageIndex } else { 1 })

    Write-Log "Applying setup media (index $setupIdx)"
    Expand-WindowsImage -ImagePath $EsdPath -Index $setupIdx -ApplyPath $MediaDir | Out-Null

    $bootWim = Join-Path $MediaDir 'sources\boot.wim'
    foreach ($b in @($peImg, $wsImg)) {
        if ($b) {
            Write-Log "Exporting $($b.ImageName) to boot.wim"
            Export-WindowsImage -SourceImagePath $EsdPath -SourceIndex $b.ImageIndex -DestinationImagePath $bootWim -CompressionType max | Out-Null
        }
    }

    $skip = @($setupIdx, $peImg.ImageIndex, $wsImg.ImageIndex)
    $editions = @()
    foreach ($i in $images) {
        if ($skip -contains $i.ImageIndex) { continue }
        $d = Get-WindowsImage -ImagePath $EsdPath -Index $i.ImageIndex
        $editions += [pscustomobject]@{ Index = $i.ImageIndex; EditionId = "$($d.EditionId)"; Name = $i.ImageName }
    }
    $current = (Get-OsBuildInfo).EditionID
    if ($Config.Esd.ExportEditions -ieq 'All') { $pick = $editions } else { $pick = @($editions | Where-Object { $_.EditionId -ieq $current }) }
    if ($pick.Count -eq 0) { throw "Edition $current not in ESD (available: $(($editions.EditionId) -join ', ')). Enterprise needs Esd.MediaType Business." }

    $installWim = Join-Path $MediaDir 'sources\install.wim'
    foreach ($p in $pick) {
        Write-Log "Exporting $($p.Name) ($($p.EditionId)) to install.wim - this takes a while"
        Export-WindowsImage -SourceImagePath $EsdPath -SourceIndex $p.Index -DestinationImagePath $installWim -CompressionType max | Out-Null
    }
    if (-not (Test-Path (Join-Path $MediaDir 'setup.exe'))) { throw 'setup.exe missing after applying setup media.' }
}

# Returns @{ AlreadyCurrent = bool }
function Invoke-EsdMedia {
    $entry = Get-EsdCatalogEntry
    Write-Log "Newest in catalog: $($entry.FileName) (build $($entry.Build).$($entry.UBR), $($entry.Type), $($entry.Lang), $($entry.Arch))"

    $os = Get-OsBuildInfo
    if (-not $ForceFlag -and $os.Build -ge $entry.Build) {
        if ($os.Build -eq $entry.Build -and $os.UBR -lt $entry.UBR) {
            Write-Log "Device is on the same version ($($os.Build).$($os.UBR)). The newer $($entry.Build).$($entry.UBR) is a cumulative update; Windows Update installs it. No upgrade needed (use -Force for a repair install)."
        } else {
            Write-Log "Device is $($os.Build).$($os.UBR), already at or above catalog build. No media needed."
        }
        return @{ AlreadyCurrent = $true }
    }

    $state = Get-State
    if ((Get-InstallImagePath) -and $state.Media -and $state.Media.SourceId -eq $entry.FileName -and -not $ForceFlag) {
        Write-Log 'Media for this build is already prepared.'
        return @{ AlreadyCurrent = $false }
    }

    $esdPath = Join-Path $WorkRoot $entry.FileName
    Get-ChildItem -Path $WorkRoot -Filter '*.esd' -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne $entry.FileName } | Remove-Item -Force
    $ok = $false
    if (Test-Path $esdPath) {
        $ok = Test-CatalogHash -Path $esdPath -Entry $entry
        if (-not $ok -and $entry.Size -and (Get-Item $esdPath).Length -ge [long]$entry.Size) { Remove-Item $esdPath -Force }
    }
    if (-not $ok) {
        Write-Log "Downloading $($entry.Url)"
        Invoke-FileDownload -Url $entry.Url -Destination $esdPath
        if (-not (Test-CatalogHash -Path $esdPath -Entry $entry)) {
            Remove-Item $esdPath -Force -ErrorAction SilentlyContinue
            throw 'ESD hash does not match the catalog. File deleted.'
        }
        Write-Log 'ESD downloaded and verified against catalog hash.'
    }

    if ($Mode -eq 'StageWorker') { $st = Get-State; if ($st.Stage) { $st.Stage['Step'] = 'Building media'; Save-State $st } }
    Set-NinjaStatus 'Building media'
    Build-MediaFromEsd -EsdPath $esdPath
    $info = Get-MediaInfo
    if (-not $info) { throw 'No install image found after building media.' }
    Write-Log "Media build $($info.Build).$($info.UBR) | Editions: $($info.Editions -join ', ') | Languages: $($info.Languages -join ', ')"

    $state = Get-State
    $state['Media'] = @{ SourceType = 'Esd'; SourceId = $entry.FileName; Sha256 = ''; Build = $info.Build; UBR = $info.UBR; Editions = $info.Editions; Languages = $info.Languages; Extracted = (Get-Date).ToString('s') }
    Save-State $state
    if ($Config.Esd.DeleteEsdAfterBuild) { Remove-Item $esdPath -Force -ErrorAction SilentlyContinue; Write-Log 'Deleted ESD after building media.' }
    return @{ AlreadyCurrent = $false }
}

# Dispatcher. Returns $true if the device is already current (nothing to do).
function Invoke-MediaDownload {
    $state = Get-State
    if ($Mode -ne 'StageWorker' -and $state.Stage -and $state.Stage.Status -in @('Running','Staged') -and (Get-InstallImagePath) -and -not $ForceFlag) {
        Write-Log 'An upgrade is staged from the current media; not replacing it (Finalize needs the same media). Use -Force to override.' 'WARN'
        return $false
    }
    return [bool](Invoke-EsdMedia).AlreadyCurrent
}

function Test-MediaFit {
    # Returns @{ Ok; Reason; Code }
    $state = Get-State
    $m = $state.Media
    if (-not $m -or -not (Get-InstallImagePath)) { return @{ Ok = $false; Code = 21; Reason = 'No media. Run Download first.' } }
    $os = Get-OsBuildInfo
    if (($os.Build -gt $m.Build) -or ($os.Build -eq $m.Build -and $os.UBR -ge $m.UBR) -or ($os.Build -eq $m.Build -and -not $ForceFlag)) {
        return @{ Ok = $false; Code = 0; Reason = "Already current (OS $($os.Build).$($os.UBR), media $($m.Build).$($m.UBR)). Same-build upgrade is a repair install; use -Force if that is intended." }
    }
    if ($m.Editions -notcontains $os.EditionID) {
        return @{ Ok = $false; Code = 21; Reason = "Edition $($os.EditionID) not in media ($($m.Editions -join ', ')). Enterprise needs business (VL) media." }
    }
    $lang = Get-InstallLanguage
    $mediaLangs = @($m.Languages | ForEach-Object { ($_ -split ' ')[0] })
    if ($lang -and ($mediaLangs -notcontains $lang) -and -not $ForceFlag) {
        return @{ Ok = $false; Code = 21; Reason = "Install language $lang not in media ($($mediaLangs -join ', ')). Keep-everything upgrade will fail." }
    }
    return @{ Ok = $true; Code = 0; Reason = "OK: $($os.Build).$($os.UBR) -> $($m.Build).$($m.UBR)" }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
try {
    switch ($Mode) {

        'Check' {
            $res = Invoke-Readiness
            try {
                $latest = Get-EsdCatalogEntry
                $os = Get-OsBuildInfo
                $cur = ($os.Build -ge $latest.Build)
                $note = ''
                if ($os.Build -eq $latest.Build -and $os.UBR -lt $latest.UBR) { $note = " (newer cumulative update $($latest.Build).$($latest.UBR) comes through Windows Update, not this script)" }
                Write-Log "Latest available: $($latest.Build).$($latest.UBR) ($($latest.FileName)). Upgrade needed: $(-not $cur)$note"
            } catch { Write-Log "Could not read media catalog: $($_.Exception.Message)" 'WARN' }
            $fails = @($res | Where-Object { $_.Status -eq 'Fail' })
            if ($fails.Count -gt 0) { Exit-Script 10 "Readiness FAILED: $($fails.Check -join ', ')" 'Not ready' }
            Exit-Script 0 'Readiness passed.' 'Ready'
        }

        'Download' {
            try { $isCurrent = Invoke-MediaDownload } catch { Exit-Script 20 "Download failed: $($_.Exception.Message)" 'Download failed' }
            if ($isCurrent) { Exit-Script 0 'Device is already on the newest build. Nothing to download.' 'Current' }
            Exit-Script 0 'Media ready.' 'Media ready'
        }

        'Scan' {
            if (-not (Get-InstallImagePath)) { Exit-Script 21 'No media. Run Download first.' }
            $setupExe = Join-Path $MediaDir 'setup.exe'
            $scanLogs = Join-Path $LogsDir 'ScanLogs'
            $scanArgs = "/auto upgrade /quiet /eula accept /compat scanonly /dynamicupdate disable /copylogs `"$scanLogs`""
            Write-Log "Running compat scan: setup.exe $scanArgs"
            $p = Start-Process -FilePath $setupExe -ArgumentList $scanArgs -WorkingDirectory $MediaDir -PassThru -Wait -WindowStyle Hidden
            $hex = ConvertTo-HexCode $p.ExitCode
            Write-Log "Scan result $hex : $(Get-CodeDescription $hex)"
            $state = Get-State; $state['Scan'] = @{ Code = $hex; When = (Get-Date).ToString('s') }; Save-State $state
            if ($hex -eq '0xC1900210') { Exit-Script 0 'Compatibility scan passed.' 'Scan passed' }
            Exit-Script 11 "Compatibility scan blocked: $hex" "Scan blocked $hex"
        }

        { $_ -in @('Stage','Upgrade') } {
            $oneShot = ($Mode -eq 'Upgrade')
            # Fast checks only (seconds). Download, media build and setup run in the background
            # StageWorker task, so the RMM job needs no long timeout.
            $task = Get-ScheduledTask -TaskName $StageTaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
            if ($task -and $task.State -eq 'Running') {
                $st = (Get-State).Stage
                Exit-Script 0 "Stage already running in the background ($($st.Step))." $st.Step
            }

            $state = Get-State
            if ($state.Stage -and $state.Stage.Status -eq 'Staged' -and -not $ForceFlag) {
                Exit-Script 0 'Already staged. Run Finalize (or use -Force to re-stage / re-run).' 'Staged'
            }

            if (-not $ForceFlag) {
                $res = Invoke-Readiness
                $fails = @($res | Where-Object { $_.Status -eq 'Fail' })
                if ($fails.Count -gt 0) { Exit-Script 10 "Readiness FAILED: $($fails.Check -join ', ')" 'Not ready' }
            }

            # Catalog lookup is quick; lets devices that are already current exit without a background job.
            try { $entry = Get-EsdCatalogEntry } catch { Exit-Script 20 "Catalog lookup failed: $($_.Exception.Message)" 'Catalog failed' }
            $os = Get-OsBuildInfo
            Write-Log "Device $($os.Build).$($os.UBR), newest in catalog $($entry.Build).$($entry.UBR)."
            if (-not $ForceFlag -and $os.Build -ge $entry.Build) {
                Exit-Script 0 'Device is already on the newest version. Nothing to stage.' 'Current'
            }

            # Customization assets (fail fast on a bad path)
            $c = $Config.Customization
            $unattend = ''; $postOobeDir = ''
            try {
                if ($c.UnattendPath) { $unattend = Get-Asset -Source $c.UnattendPath -DestName 'unattend.xml'; Write-Log "Unattend: $unattend" }
                if ($c.PostOOBEScriptPath) {
                    $postOobeDir = Join-Path $CustomDir 'PostOOBE'
                    New-Item -Path $postOobeDir -ItemType Directory -Force | Out-Null
                    $src = Get-Asset -Source $c.PostOOBEScriptPath -DestName 'postoobe_source.cmd'
                    Copy-Item $src (Join-Path $postOobeDir 'setupcomplete.cmd') -Force
                    Write-Log "PostOOBE SetupComplete: $postOobeDir\setupcomplete.cmd"
                }
            } catch { Exit-Script 2 "Customization asset failed: $($_.Exception.Message)" 'Config error' }

            Set-OobeSuppression
            Copy-SelfToWorkRoot

            $state = Get-State
            $state['Stage'] = @{ Status = 'Running'; Step = 'Queued'; Started = (Get-Date).ToString('s'); Unattend = $unattend; PostOobeDir = $postOobeDir; MigrateDrivers = $Config.Setup.MigrateDrivers; OneShot = $oneShot }
            Write-Log "Driver migration for this run: $($Config.Setup.MigrateDrivers)"
            Save-State $state

            $workerArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$LocalScriptPath`" -Mode StageWorker"
            if ($ForceFlag) { $workerArgs += ' -Force' }
            Register-AltecTask -Name $StageTaskName -Execute 'powershell.exe' -Arguments $workerArgs -WorkingDir $WorkRoot -LimitHours ([math]::Ceiling($Config.Setup.StageTimeoutMinutes / 60) + 2)
            Start-ScheduledTask -TaskName $StageTaskName -TaskPath $TaskPath
            if ($oneShot) {
                Exit-Script 0 "Upgrade started in the background (target $($entry.Build).$($entry.UBR)). Watch the status field: Downloading, Building media, Upgrading, Upgraded - rebooting. The device restarts automatically with a $($Config.Finalize.RebootDelaySeconds)s warning when setup finishes." 'Queued'
            }
            Exit-Script 0 "Stage started in the background (target $($entry.Build).$($entry.UBR)). Watch the status field: Downloading, Building media, Staging, Staged." 'Queued'
        }

        'StageWorker' {
            # Runs inside the SYSTEM scheduled task: download, build media, run setup /SkipFinalize,
            # record the result and update the Ninja status field.
            function Set-Step([string]$Step) {
                $st = Get-State; if ($st.Stage) { $st.Stage['Step'] = $Step; Save-State $st }
                Set-NinjaStatus $Step
            }
            function Stop-Stage([int]$Code, [string]$Msg, [string]$Field, [string]$Result) {
                $st = Get-State
                if (-not $st.Stage) { $st['Stage'] = @{} }
                $st.Stage['Status'] = 'Failed'; $st.Stage['Code'] = $Result; $st.Stage['Step'] = $Field
                $st.Stage['Finished'] = (Get-Date).ToString('s'); Save-State $st
                Exit-Script $Code $Msg $Field
            }

            $state = Get-State
            if (-not $state.Stage) { Exit-Script 30 'StageWorker: no stage request in state. Run -Mode Stage.' 'Stage failed' }

            Set-Step 'Downloading'
            try { $isCurrent = Invoke-MediaDownload } catch { Stop-Stage 20 "Download failed: $($_.Exception.Message)" 'Download failed' 'Download' }
            if ($isCurrent) {
                $st = Get-State; $st.Stage['Status'] = 'Current'; $st.Stage['Step'] = 'Current'; Save-State $st
                Exit-Script 0 'Device is already on the newest version.' 'Current'
            }

            $fit = Test-MediaFit
            Write-Log $fit.Reason
            if (-not $fit.Ok) {
                if ($fit.Code -eq 0) {
                    $st = Get-State; $st.Stage['Status'] = 'Current'; $st.Stage['Step'] = 'Current'; Save-State $st
                    Exit-Script 0 $fit.Reason 'Current'
                }
                Stop-Stage $fit.Code $fit.Reason 'Media mismatch' 'Media'
            }

            $s = $Config.Setup
            $state = Get-State
            $setupLogDir = Join-Path $LogsDir 'SetupLogs'
            if (-not (Test-Path $setupLogDir)) { New-Item -Path $setupLogDir -ItemType Directory -Force | Out-Null }
            $a = New-Object System.Collections.Generic.List[string]
            $oneShot = [bool]$state.Stage.OneShot
            $phaseArg = $(if ($oneShot) { '/noreboot' } else { '/skipfinalize' })
            foreach ($x in @('/auto upgrade','/quiet','/eula accept',$phaseArg,
                             "/dynamicupdate $($s.DynamicUpdate)", "/priority $($s.Priority)",
                             "/bitlocker $($s.BitLocker)", "/migratedrivers $(if ($state.Stage.MigrateDrivers) { $state.Stage.MigrateDrivers } else { $s.MigrateDrivers })",
                             "/copylogs `"$(Join-Path $LogsDir 'SetupLogs')`"")) { $a.Add($x) }
            if ($s.ShowOOBE)              { $a.Add("/showoobe $($s.ShowOOBE)") }
            if ($s.Telemetry)             { $a.Add("/telemetry $($s.Telemetry)") }
            if ($s.CompatIgnoreWarning)   { $a.Add('/compat ignorewarning') }
            if ($state.Stage.Unattend)    { $a.Add("/unattend:`"$($state.Stage.Unattend)`"") }
            if ($state.Stage.PostOobeDir) { $a.Add("/postoobe:`"$($state.Stage.PostOobeDir)`"") }
            if ($s.ExtraStageArgs)        { $a.Add($s.ExtraStageArgs) }
            $stageArgs = $a -join ' '
            $state.Stage['Args'] = $stageArgs
            Save-State $state

            if ($oneShot) {
                # Register post-upgrade task before setup so first boot on the new OS is handled.
                $postArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$LocalScriptPath`" -Mode PostUpgrade"
                Register-AltecTask -Name $PostTaskName -Execute 'powershell.exe' -Arguments $postArgs -WorkingDir $WorkRoot -AtStartup -LimitHours 2
                Set-Step 'Upgrading'
            } else {
                Set-Step 'Staging'
            }
            $setupExe = Join-Path $MediaDir 'setup.exe'
            Write-Log "StageWorker: setup.exe $stageArgs"
            $p = Start-Process -FilePath $setupExe -ArgumentList $stageArgs -WorkingDirectory $MediaDir -PassThru -WindowStyle Hidden
            $done = $p.WaitForExit([int]$s.StageTimeoutMinutes * 60000)
            if (-not $done) {
                if ($oneShot) { Remove-AltecTask $PostTaskName }
                Stop-Stage 30 "Setup did not finish within $($s.StageTimeoutMinutes) min." 'Stage failed - timeout' 'Timeout'
            }

            $hex = ConvertTo-HexCode $p.ExitCode
            if ($oneShot) {
                if ($s.UpgradeSuccessCodes -contains $hex) {
                    $st = Get-State
                    $st.Stage['Status'] = 'Upgraded'; $st.Stage['Code'] = $hex; $st.Stage['Step'] = 'Upgraded - rebooting'
                    $st.Stage['Finished'] = (Get-Date).ToString('s')
                    $st['Finalize'] = @{ Status = 'Finalized'; Code = $hex; When = (Get-Date).ToString('s'); OneShot = $true }
                    $st['PostUpgrade'] = @{ Attempts = 0 }
                    Save-State $st
                    if ($Config.Finalize.NoReboot) { Exit-Script 0 "Upgrade complete ($hex). Reboot required (NoReboot set)." 'Upgraded - reboot pending' }
                    $delay = [int]$Config.Finalize.RebootDelaySeconds
                    & shutdown.exe /r /t $delay /c $Config.Finalize.RebootMessage /d p:2:4 | Out-Null
                    Exit-Script 0 "Upgrade complete ($hex). Reboot scheduled in $delay seconds." 'Upgraded - rebooting'
                }
                Remove-AltecTask $PostTaskName
                Stop-Stage 30 "Upgrade failed $hex : $(Get-CodeDescription $hex)" "Upgrade failed $hex" $hex
            }
            if ($s.StageSuccessCodes -contains $hex) {
                $st = Get-State
                $st.Stage['Status'] = 'Staged'; $st.Stage['Code'] = $hex; $st.Stage['Step'] = 'Staged'
                $st.Stage['Finished'] = (Get-Date).ToString('s'); Save-State $st
                Exit-Script 0 "Stage complete ($hex). Ready to Finalize." 'Staged'
            }
            Stop-Stage 30 "Stage failed $hex : $(Get-CodeDescription $hex)" "Stage failed $hex" $hex
        }

        'Finalize' {
            $state = Get-State
            $staged = $state.Stage -and $state.Stage.Status -eq 'Staged'
            if (-not $staged -and -not $ForceFlag) { Exit-Script 0 'Nothing staged on this device. Skipping.' $null }
            if (-not (Test-Path $BtFolder)) { Exit-Script 41 "$BtFolder not found; staged files are gone. Re-run Stage." 'Not staged' }
            $setupExe = Join-Path $MediaDir 'setup.exe'
            if (-not (Test-Path $setupExe)) { Exit-Script 41 'Local media missing; Finalize must run from the same media used to stage.' 'Not staged' }

            Copy-SelfToWorkRoot
            $postArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$LocalScriptPath`" -Mode PostUpgrade"
            Register-AltecTask -Name $PostTaskName -Execute 'powershell.exe' -Arguments $postArgs -WorkingDir $WorkRoot -AtStartup -LimitHours 2
            Write-Log 'Registered post-upgrade startup task.'

            try {
                $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
                if ($bl.ProtectionStatus -eq 'On') {
                    Suspend-BitLocker -MountPoint $env:SystemDrive -RebootCount $Config.Finalize.SuspendBitLockerRebootCount | Out-Null
                    Write-Log "BitLocker suspended for $($Config.Finalize.SuspendBitLockerRebootCount) reboots."
                }
            } catch { Write-Log "BitLocker check skipped: $($_.Exception.Message)" 'WARN' }

            $fArgs = '/auto upgrade /quiet /eula accept /finalize /noreboot'
            if ($Config.Setup.ExtraFinalizeArgs) { $fArgs = "$fArgs $($Config.Setup.ExtraFinalizeArgs)" }
            Write-Log "Finalize command: setup.exe $fArgs"
            $p = Start-Process -FilePath $setupExe -ArgumentList $fArgs -WorkingDirectory $MediaDir -PassThru -WindowStyle Hidden
            if (-not $p.WaitForExit($Config.Setup.FinalizeTimeoutMinutes * 60000)) {
                Exit-Script 40 "Finalize did not finish within $($Config.Setup.FinalizeTimeoutMinutes) min." 'Finalize timeout'
            }
            $hex = ConvertTo-HexCode $p.ExitCode
            $state = Get-State
            $state['Finalize'] = @{ Code = $hex; When = (Get-Date).ToString('s') }
            if ($Config.Setup.FinalizeSuccessCodes -notcontains $hex) {
                $state.Finalize['Status'] = 'Failed'; Save-State $state
                Remove-AltecTask $PostTaskName
                Exit-Script 40 "Finalize failed $hex : $(Get-CodeDescription $hex)" "Finalize failed $hex"
            }
            $state.Finalize['Status'] = 'Finalized'
            $state['PostUpgrade'] = @{ Attempts = 0 }
            Save-State $state

            if ($Config.Finalize.NoReboot) { Exit-Script 0 'Finalized. Reboot required to complete (NoReboot set).' 'Finalized - reboot pending' }
            $delay = [int]$Config.Finalize.RebootDelaySeconds
            & shutdown.exe /r /t $delay /c $Config.Finalize.RebootMessage /d p:2:4 | Out-Null
            Exit-Script 0 "Finalized. Reboot scheduled in $delay seconds." 'Finalized - rebooting'
        }

        'PostUpgrade' {
            $state = Get-State
            if (-not $state.Media) { Remove-AltecTask $PostTaskName; Exit-Script 50 'No state found.' 'Post-upgrade failed' }
            $os = Get-OsBuildInfo
            $upgraded = ($os.Build -gt $state.Media.Build) -or ($os.Build -eq $state.Media.Build)
            if (-not $upgraded) {
                if (-not $state.PostUpgrade) { $state['PostUpgrade'] = @{ Attempts = 0 } }
                $state.PostUpgrade['Attempts'] = [int]$state.PostUpgrade.Attempts + 1
                Save-State $state
                if ($state.PostUpgrade.Attempts -ge $Config.PostUpgrade.MaxBootAttempts) {
                    Remove-AltecTask $PostTaskName
                    Exit-Script 50 "Upgrade not detected after $($state.PostUpgrade.Attempts) boots (still $($os.Build)). Likely rollback; check $env:SystemDrive\`$WINDOWS.~BT\Sources\Rollback and SetupLogs." 'Rolled back'
                }
                Exit-Script 0 "Build $($os.Build) not yet at target $($state.Media.Build). Will check next boot." 'Upgrade pending'
            }

            Write-Log "Upgrade detected: $($os.ProductName) $($os.DisplayVersion) build $($os.Build).$($os.UBR)"
            Set-OobeSuppression

            if ($Config.PostUpgrade.CleanupMediaOnSuccess) {
                Remove-Item $MediaDir -Recurse -Force -ErrorAction SilentlyContinue
                Get-ChildItem -Path $WorkRoot -Filter '*.esd' -ErrorAction SilentlyContinue | Remove-Item -Force
                Write-Log 'Removed local media.'
            }
            Remove-AltecTask $StageTaskName
            Remove-AltecTask $PostTaskName
            $state['PostUpgrade'] = @{ Status = 'Completed'; When = (Get-Date).ToString('s'); Build = "$($os.Build).$($os.UBR)" }
            Save-State $state
            Exit-Script 0 'Post-upgrade complete.' "Completed $($os.DisplayVersion)"
        }

        'Status' {
            $state = Get-State
            $os = Get-OsBuildInfo
            Write-Log "OS: $($os.ProductName) $($os.DisplayVersion) build $($os.Build).$($os.UBR) ($($os.EditionID))"
            if ($state.Media) { Write-Log "Media: build $($state.Media.Build).$($state.Media.UBR) ($($state.Media.Editions -join ', '))" } else { Write-Log 'Media: none' }
            Write-Log "Setup working folder present: $(Test-Path $BtFolder)"

            if ($state.PostUpgrade -and $state.PostUpgrade.Status -eq 'Completed') { Exit-Script 0 "Completed: $($state.PostUpgrade.Build)" "Completed $($os.DisplayVersion)" }
            if ($state.PostUpgrade -and $state.PostUpgrade.Status) { Exit-Script 50 "Post-upgrade status: $($state.PostUpgrade.Status)" 'Post-upgrade failed' }
            if ($state.Finalize -and $state.Finalize.Status -eq 'Finalized') { Exit-Script 0 'Finalized; waiting for reboot / post-upgrade.' 'Finalized - reboot pending' }
            if ($state.Finalize -and $state.Finalize.Status -eq 'Failed') { Exit-Script 40 "Finalize failed $($state.Finalize.Code)" 'Finalize failed' }

            $t = Get-ScheduledTask -TaskName $StageTaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
            if ($t -and $t.State -eq 'Running') { Exit-Script 0 "Stage in progress since $($state.Stage.Started): $($state.Stage.Step)." $state.Stage.Step }
            if ($state.Stage -and $state.Stage.Status -eq 'Running') {
                # Worker is gone but never recorded a result (crash, reboot during staging).
                $state.Stage['Status'] = 'Failed'; $state.Stage['Code'] = 'WorkerStopped'
                Save-State $state
            }
            if ($state.Stage -and $state.Stage.Status -eq 'Current') { Exit-Script 0 'Device is already on the newest version.' 'Current' }
            if ($state.Stage -and $state.Stage.Status -eq 'Staged') { Exit-Script 0 "Staged ($($state.Stage.Code)). Ready to Finalize." 'Staged' }
            if ($state.Stage -and $state.Stage.Status -eq 'Failed') { Show-StageDiagnostics; Exit-Script 30 "Stage failed $($state.Stage.Code): $(Get-CodeDescription $state.Stage.Code)" "Stage failed $($state.Stage.Code)" }
            Exit-Script 0 'Nothing staged.' 'Not staged'
        }

        'Cleanup' {
            Remove-AltecTask $StageTaskName
            Remove-AltecTask $PostTaskName
            Remove-Item $MediaDir -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item $CustomDir -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item (Join-Path $WorkRoot 'Catalog') -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item (Join-Path $WorkRoot 'products.cab') -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path $WorkRoot -Filter '*.esd' -ErrorAction SilentlyContinue | Remove-Item -Force
            Remove-Item $StatePath -Force -ErrorAction SilentlyContinue
            Write-Log "Note: $BtFolder is managed by Windows and is not removed here."
            Exit-Script 0 'Cleanup complete.' 'Cleaned up'
        }
    }
} catch {
    Exit-Script 1 "Unhandled error: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" 'Script error'
}
