<#
.SYNOPSIS
    Altec - NinjaOne DNS resilience (NRPT). Keeps the NinjaOne agent, patcher,
    and NinjaOne Remote resolving their cloud endpoints when the device's
    normal DNS (DC, router, DNS filter) is down.

.DESCRIPTION
    Creates ONE Name Resolution Policy Table (NRPT) rule, tagged by Comment,
    that sends lookups for NinjaOne namespaces to public resolvers instead of
    the adapter DNS servers. All other name resolution is untouched.

    READ-ONLY BY DEFAULT. Action defaults to Audit.

    Allowed remediation actions (nothing else is ever changed):
      Apply  - create or update the single Altec-tagged NRPT rule, then flush
               the DNS client cache. Auto-rolls back if post-change lookups fail.
      Remove - delete Altec-tagged NRPT rules only, then flush the DNS cache.

    Setting resolution order (first non-empty wins):
      1. CLI parameter
      2. NinjaOne script variable (environment variable)
      3. JSON config file (-ConfigPath / $env:nrptConfigPath)
      4. Built-in default

    NinjaOne script variables (create these on the script in Ninja):
      nrptAction            Drop-down: Audit | Apply | Remove
      nrptNameServers       Text, comma separated, e.g. 1.1.1.1,8.8.8.8
      nrptExtraNamespaces   Text, comma separated, added to the built-in list
      nrptConfigPath        Text, path to JSON config on the endpoint
      nrptStatusField       Text, optional device custom field to write status

Exit codes:
    0   Compliant / action completed and verified
    1   Audit: drift detected (rule missing or different)
    2   GPO-delivered NRPT present; local rules are ignored by Windows
    3   Pre-check failed: not enough configured resolvers answered
    4   Apply/Remove failed (changes rolled back where possible)
    5   Post-change verification failed; changes rolled back
    10  Not elevated, or NRPT cmdlets unavailable on this OS
    11  Invalid configuration

.NOTES
    Version : 1.0.0  (2026-10-28)
    Author  : Altec Solutions Group
    Requires: Windows 8.1 / Server 2012 R2 or later, PowerShell 5.1, SYSTEM/admin
    ASCII-clean by design. Do not paste smart quotes or em dashes into this file.
#>

[CmdletBinding()]
param(
    [ValidateSet('', 'Audit', 'Apply', 'Remove')]
    [string]$Action = '',
    [string[]]$NameServers,
    [string[]]$ExtraNamespaces,
    [string]$ConfigPath = '',
    [string]$StatusField = ''
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'

# ---------------------------------------------------------------------------
# Built-in defaults (used only when no CLI, Ninja variable, or config value)
# ---------------------------------------------------------------------------
$Defaults = [ordered]@{
    Action              = 'Audit'
    TagComment          = 'Altec-NinjaOne-NRPT'
    DisplayName         = 'Altec - NinjaOne DNS resilience'
    # Two different providers on purpose. NRPT does NOT fall back to the
    # adapter DNS if every server listed here is unreachable.
    NameServers         = @('1.1.1.1', '8.8.8.8')
    # Leading dot = suffix rule (matches every subdomain at any depth).
    # No leading dot = exact FQDN rule.
    Namespaces          = @(
        # Core agent, patcher, websocket/connect shards, RTC, tunnels, file
        # transfer (fts-*), resources, MDM/SCEP, NMS, and every regional
        # instance (app, us2, eu, ca, oc) - all live under ninjarmm.com
        '.ninjarmm.com',
        # NinjaOne Remote relays (nc-*-<region>.ninjarmm.net)
        '.ninjarmm.net',
        # Current NinjaOne brand domain
        '.ninjaone.com',
        # Branded / white-label console URLs, all regions
        '.rmmservice.com',
        '.rmmservice.eu',
        '.rmmservice.ca',
        '.rmmservice.com.au',
        # US2 instance branded URLs
        '.rmmservices.net',
        # NinjaOne Backup
        '.ninja-backup.com',
        # Agent crash reporting (NinjaOne lists it as a global requirement)
        'ingest.sentry.io',
        '.ingest.sentry.io',
        # NinjaOne-owned S3 buckets (exact names only; never all of AWS)
        'ninjauploads.s3.amazonaws.com',
        'ninjaresources.s3.amazonaws.com',
        'ninja-catalog-release.s3.us-east-1.amazonaws.com',
        'ninja-attachments.s3.us-west-2.amazonaws.com'
    )
    # Hosts resolved after a change to prove the rule works end to end.
    # Must ALL resolve or the change is rolled back.
    VerifyHosts         = @('app.ninjarmm.com', 'resources.ninjarmm.com')
    # Extra hosts checked and reported, but never trigger a rollback.
    InfoHosts           = @('nc-1-us-west-2.ninjarmm.net', 'agent-app.ninjarmm.com')
    # Minimum resolvers that must answer before Apply is allowed.
    MinHealthyResolvers = 1
    StatusField         = ''
    LogDir              = 'C:\ProgramData\Altec\Logs'
}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
$script:LogFile = $null
function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    if ($script:LogFile) {
        try { Add-Content -Path $script:LogFile -Value $line -Encoding ASCII } catch { }
    }
}

function Set-NinjaStatus {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($script:Cfg.StatusField)) { return }
    if (-not (Get-Command -Name 'Ninja-Property-Set' -ErrorAction SilentlyContinue)) {
        Write-Log "Status field '$($script:Cfg.StatusField)' set but Ninja-Property-Set is not available; skipping." 'WARN'
        return
    }
    try {
        $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
        Ninja-Property-Set $script:Cfg.StatusField ("{0} | {1}" -f $stamp, $Text) | Out-Null
    } catch {
        Write-Log "Could not write status field: $($_.Exception.Message)" 'WARN'
    }
}

function Exit-Script {
    param([int]$Code, [string]$Summary)
    $lvl = 'OK'
    if ($Code -ne 0) { $lvl = 'ERROR' }
    if ($Code -eq 1) { $lvl = 'WARN' }
    Write-Log ("RESULT: exit {0} - {1}" -f $Code, $Summary) $lvl
    Set-NinjaStatus ("exit {0}: {1}" -f $Code, $Summary)
    exit $Code
}

# ---------------------------------------------------------------------------
# Config resolution: CLI > Ninja env var > config file > default
# ---------------------------------------------------------------------------
function Split-List {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
    return @($Value -split '[,;\s]+' | Where-Object { $_ -ne '' } | ForEach-Object { $_.Trim() })
}

function Resolve-Setting {
    param($CliValue, [string]$EnvName, $FileValue, $DefaultValue, [switch]$IsList)
    if ($null -ne $CliValue -and @($CliValue).Count -gt 0 -and -not ([string]::IsNullOrWhiteSpace(($CliValue -join '')))) {
        if ($IsList) { return @($CliValue | ForEach-Object { Split-List $_ }) }
        return $CliValue
    }
    $envVal = [Environment]::GetEnvironmentVariable($EnvName)
    if (-not [string]::IsNullOrWhiteSpace($envVal)) {
        if ($IsList) { return @(Split-List $envVal) }
        return $envVal.Trim()
    }
    if ($null -ne $FileValue -and -not ([string]::IsNullOrWhiteSpace((@($FileValue) -join '')))) {
        if ($IsList) { return @($FileValue) }
        return $FileValue
    }
    return $DefaultValue
}

$cfgPath = Resolve-Setting -CliValue $ConfigPath -EnvName 'nrptConfigPath' -FileValue $null -DefaultValue ''
$file = $null
if (-not [string]::IsNullOrWhiteSpace($cfgPath)) {
    if (-not (Test-Path -LiteralPath $cfgPath)) {
        Write-Output "Config file not found: $cfgPath"
        exit 11
    }
    try {
        $file = Get-Content -LiteralPath $cfgPath -Raw | ConvertFrom-Json
    } catch {
        Write-Output "Config file is not valid JSON: $cfgPath - $($_.Exception.Message)"
        exit 11
    }
}
function Get-FileValue { param([string]$Name) if ($null -eq $file) { return $null }; $p = $file.PSObject.Properties[$Name]; if ($p) { return $p.Value }; return $null }

$script:Cfg = [ordered]@{
    Action              = Resolve-Setting $Action 'nrptAction' (Get-FileValue 'Action') $Defaults.Action
    TagComment          = Resolve-Setting $null 'nrptTagComment' (Get-FileValue 'TagComment') $Defaults.TagComment
    DisplayName         = Resolve-Setting $null 'nrptDisplayName' (Get-FileValue 'DisplayName') $Defaults.DisplayName
    NameServers         = Resolve-Setting $NameServers 'nrptNameServers' (Get-FileValue 'NameServers') $Defaults.NameServers -IsList
    Namespaces          = Resolve-Setting $null 'nrptNamespaces' (Get-FileValue 'Namespaces') $Defaults.Namespaces -IsList
    ExtraNamespaces     = Resolve-Setting $ExtraNamespaces 'nrptExtraNamespaces' (Get-FileValue 'ExtraNamespaces') @() -IsList
    VerifyHosts         = Resolve-Setting $null 'nrptVerifyHosts' (Get-FileValue 'VerifyHosts') $Defaults.VerifyHosts -IsList
    InfoHosts           = Resolve-Setting $null 'nrptInfoHosts' (Get-FileValue 'InfoHosts') $Defaults.InfoHosts -IsList
    MinHealthyResolvers = [int](Resolve-Setting $null 'nrptMinHealthyResolvers' (Get-FileValue 'MinHealthyResolvers') $Defaults.MinHealthyResolvers)
    StatusField         = Resolve-Setting $StatusField 'nrptStatusField' (Get-FileValue 'StatusField') $Defaults.StatusField
    LogDir              = Resolve-Setting $null 'nrptLogDir' (Get-FileValue 'LogDir') $Defaults.LogDir
}

# Normalize lists
$script:Cfg.Namespaces = @(@($script:Cfg.Namespaces) + @($script:Cfg.ExtraNamespaces) |
    Where-Object { $null -ne $_ } | ForEach-Object { $_.ToString().Trim().ToLowerInvariant() } |
    Where-Object { $_ -ne '' } | Sort-Object -Unique)
$script:Cfg.NameServers = @(@($script:Cfg.NameServers) | Where-Object { $null -ne $_ } | ForEach-Object { $_.ToString().Trim() } | Where-Object { $_ -ne '' } | Select-Object -Unique)

# Logging setup
try {
    if (-not (Test-Path -LiteralPath $script:Cfg.LogDir)) { New-Item -ItemType Directory -Path $script:Cfg.LogDir -Force | Out-Null }
    $script:LogFile = Join-Path $script:Cfg.LogDir 'NinjaDnsResilience.log'
} catch { $script:LogFile = $null }

Write-Log "Altec NinjaOne DNS resilience v$ScriptVersion on $env:COMPUTERNAME"
Write-Log ("Action={0}  NameServers={1}" -f $script:Cfg.Action, ($script:Cfg.NameServers -join ','))
Write-Log ("Namespaces ({0}): {1}" -f $script:Cfg.Namespaces.Count, ($script:Cfg.Namespaces -join ' '))

# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
if ($script:Cfg.Action -notin @('Audit', 'Apply', 'Remove')) { Exit-Script 11 "Invalid Action '$($script:Cfg.Action)'" }
if ($script:Cfg.Namespaces.Count -eq 0) { Exit-Script 11 'Namespace list is empty' }
if ($script:Cfg.NameServers.Count -eq 0) { Exit-Script 11 'NameServer list is empty' }
foreach ($ns in $script:Cfg.NameServers) {
    $ip = $null
    if (-not [System.Net.IPAddress]::TryParse($ns, [ref]$ip)) { Exit-Script 11 "NameServer '$ns' is not an IP address" }
}
foreach ($n in $script:Cfg.Namespaces) {
    if ($n -notmatch '^\.?[a-z0-9]([a-z0-9\-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9\-]*[a-z0-9])?)+$') { Exit-Script 11 "Namespace '$n' is not a valid DNS name or suffix" }
}
if ($script:Cfg.MinHealthyResolvers -lt 1) { $script:Cfg.MinHealthyResolvers = 1 }
$script:Cfg.VerifyHosts = @(@($script:Cfg.VerifyHosts) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$script:Cfg.InfoHosts = @(@($script:Cfg.InfoHosts) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
if ($script:Cfg.VerifyHosts.Count -eq 0) { Exit-Script 11 'VerifyHosts list is empty' }

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Exit-Script 10 'Must run elevated (SYSTEM or admin)' }
foreach ($cmd in 'Get-DnsClientNrptRule', 'Add-DnsClientNrptRule', 'Set-DnsClientNrptRule', 'Remove-DnsClientNrptRule', 'Get-DnsClientNrptPolicy', 'Resolve-DnsName') {
    if (-not (Get-Command -Name $cmd -ErrorAction SilentlyContinue)) { Exit-Script 10 "Cmdlet $cmd not available (unsupported OS)" }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Get-TaggedRules {
    @(Get-DnsClientNrptRule | Where-Object { $_.Comment -eq $script:Cfg.TagComment })
}

function Test-GpoNrpt {
    # When any NRPT rule is delivered by Group Policy, Windows ignores ALL
    # locally configured NRPT rules.
    $gpoKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient\DnsPolicyConfig'
    if (Test-Path -LiteralPath $gpoKey) {
        $children = @(Get-ChildItem -LiteralPath $gpoKey -ErrorAction SilentlyContinue)
        if ($children.Count -gt 0) { return $true }
    }
    return $false
}

function Test-Resolver {
    param([string]$Server, [string]$HostName)
    try {
        $r = Resolve-DnsName -Name $HostName -Server $Server -DnsOnly -NoHostsFile -QuickTimeout -ErrorAction Stop
        return (@($r | Where-Object { $_.Type -in 'A', 'AAAA', 'CNAME' }).Count -gt 0)
    } catch { return $false }
}

function Test-NormalPath {
    # Resolves through the normal client path, so NRPT is honored.
    param([string]$HostName)
    try {
        $r = Resolve-DnsName -Name $HostName -DnsOnly -NoHostsFile -ErrorAction Stop
        return (@($r | Where-Object { $_.Type -in 'A', 'AAAA' }).Count -gt 0)
    } catch { return $false }
}

function Compare-Sets {
    param([string[]]$A, [string[]]$B, [switch]$Ordered)
    $a1 = @($A | ForEach-Object { $_.ToLowerInvariant() })
    $b1 = @($B | ForEach-Object { $_.ToLowerInvariant() })
    if (-not $Ordered) { $a1 = @($a1 | Sort-Object -Unique); $b1 = @($b1 | Sort-Object -Unique) }
    return (($a1 -join '|') -eq ($b1 -join '|'))
}

function Get-DriftReport {
    $result = [ordered]@{ Compliant = $false; Reasons = @(); Warnings = @(); Rules = @(Get-TaggedRules) }
    $rules = $result.Rules
    if ($rules.Count -eq 0) { $result.Reasons += 'No Altec-tagged NRPT rule present'; return $result }
    if ($rules.Count -gt 1) { $result.Reasons += "Found $($rules.Count) Altec-tagged rules; expected 1" }
    $r = $rules[0]
    if (-not (Compare-Sets -A @($r.Namespace) -B $script:Cfg.Namespaces)) {
        $missing = @($script:Cfg.Namespaces | Where-Object { @($r.Namespace) -notcontains $_ })
        $extra = @(@($r.Namespace) | Where-Object { $script:Cfg.Namespaces -notcontains $_ })
        if ($missing.Count) { $result.Reasons += ('Missing namespaces: ' + ($missing -join ' ')) }
        if ($extra.Count) { $result.Reasons += ('Unexpected namespaces: ' + ($extra -join ' ')) }
    }
    if (-not (Compare-Sets -A @($r.NameServers) -B $script:Cfg.NameServers -Ordered)) {
        $result.Reasons += ('NameServers differ: have ' + (@($r.NameServers) -join ',') + ' want ' + ($script:Cfg.NameServers -join ','))
    }
    # Cross-check the effective policy. Reported as a warning only, so a
    # formatting difference in the effective view can never force a rollback.
    # (GPO override is detected separately by Test-GpoNrpt.)
    $eff = @(Get-DnsClientNrptPolicy -Effective -ErrorAction SilentlyContinue)
    $effNs = @($eff | ForEach-Object { $_.Namespace } | ForEach-Object { $_.ToLowerInvariant() })
    $notEff = @($script:Cfg.Namespaces | Where-Object { $effNs -notcontains $_ })
    if ($notEff.Count) { $result.Warnings += ('Not shown in effective policy: ' + ($notEff -join ' ')) }
    if ($result.Reasons.Count -eq 0) { $result.Compliant = $true }
    return $result
}

function Invoke-Verification {
    $ok = $true
    foreach ($h in $script:Cfg.VerifyHosts) {
        if (Test-NormalPath $h) { Write-Log "Verify: $h resolves" 'OK' }
        else { Write-Log "Verify: $h FAILED to resolve" 'ERROR'; $ok = $false }
    }
    foreach ($h in $script:Cfg.InfoHosts) {
        if (Test-NormalPath $h) { Write-Log "Info: $h resolves" 'OK' }
        else { Write-Log "Info: $h did not resolve (informational only)" 'WARN' }
    }
    return $ok
}

# ---------------------------------------------------------------------------
# Common pre-checks
# ---------------------------------------------------------------------------
$gpoActive = Test-GpoNrpt
if ($gpoActive) {
    Write-Log 'Group Policy NRPT rules are present. Windows ignores local NRPT rules while any GPO NRPT exists.' 'WARN'
    Write-Log 'Fix: deliver this rule through the same GPO (Computer Config > Policies > Windows Settings > Name Resolution Policy).' 'WARN'
}

# Resolver health (reported in every mode, enforced for Apply)
$probeHost = $script:Cfg.VerifyHosts[0]
$healthy = 0
foreach ($ns in $script:Cfg.NameServers) {
    if (Test-Resolver -Server $ns -HostName $probeHost) { Write-Log "Resolver $ns answered for $probeHost" 'OK'; $healthy++ }
    else { Write-Log "Resolver $ns did NOT answer for $probeHost (outbound 53 blocked or intercepted?)" 'WARN' }
}

# ---------------------------------------------------------------------------
# Actions
# ---------------------------------------------------------------------------
switch ($script:Cfg.Action) {

    'Audit' {
        $drift = Get-DriftReport
        foreach ($reason in $drift.Reasons) { Write-Log $reason 'WARN' }
        foreach ($w in $drift.Warnings) { Write-Log $w 'WARN' }
        Invoke-Verification | Out-Null
        if ($gpoActive -and $drift.Rules.Count -gt 0) { Exit-Script 2 'Local rule present but overridden by GPO NRPT' }
        if ($healthy -lt $script:Cfg.MinHealthyResolvers) { Exit-Script 3 "Only $healthy configured resolver(s) answered" }
        if ($drift.Compliant) { Exit-Script 0 'Compliant' }
        if ($gpoActive) { Exit-Script 2 'Drift, and GPO NRPT present (local rules would be ignored)' }
        Exit-Script 1 ('Drift: ' + ($drift.Reasons -join '; '))
    }

    'Apply' {
        if ($gpoActive) { Exit-Script 2 'GPO NRPT present; not applying a local rule that Windows would ignore' }
        if ($healthy -lt $script:Cfg.MinHealthyResolvers) {
            Exit-Script 3 "Refusing to apply: only $healthy of $($script:Cfg.NameServers.Count) resolvers answered. Applying would take Ninja offline."
        }

        $drift = Get-DriftReport
        if ($drift.Compliant) {
            Invoke-Verification | Out-Null
            Exit-Script 0 'Already compliant; no change made'
        }

        # Snapshot for rollback
        $before = @($drift.Rules | ForEach-Object {
                [pscustomobject]@{ Namespace = @($_.Namespace); NameServers = @($_.NameServers); DisplayName = $_.DisplayName; Comment = $_.Comment }
            })
        $changed = $false
        try {
            $rules = @($drift.Rules)
            # Remove duplicates beyond the first tagged rule
            if ($rules.Count -gt 1) {
                foreach ($dup in $rules[1..($rules.Count - 1)]) {
                    Write-Log "Removing duplicate tagged rule $($dup.Name)"
                    Remove-DnsClientNrptRule -Name $dup.Name -Force
                }
                $changed = $true
            }
            if ($rules.Count -eq 0) {
                Write-Log 'Creating NRPT rule'
                Add-DnsClientNrptRule -Namespace $script:Cfg.Namespaces -NameServers $script:Cfg.NameServers `
                    -DisplayName $script:Cfg.DisplayName -Comment $script:Cfg.TagComment | Out-Null
            } else {
                Write-Log "Updating NRPT rule $($rules[0].Name)"
                Set-DnsClientNrptRule -Name $rules[0].Name -Namespace $script:Cfg.Namespaces -NameServers $script:Cfg.NameServers `
                    -DisplayName $script:Cfg.DisplayName -Comment $script:Cfg.TagComment
            }
            $changed = $true
            Clear-DnsClientCache
            Start-Sleep -Seconds 2
        } catch {
            Write-Log "Apply error: $($_.Exception.Message)" 'ERROR'
            if (-not $changed -and @(Get-TaggedRules).Count -eq $before.Count) { Exit-Script 4 'Apply failed before any change was made' }
        }

        $verifyOk = Invoke-Verification
        $post = Get-DriftReport
        if ($verifyOk -and $post.Compliant) { Exit-Script 0 'Rule applied and verified' }

        # Roll back to the exact prior state
        Write-Log 'Verification failed; rolling back' 'ERROR'
        foreach ($reason in $post.Reasons) { Write-Log $reason 'WARN' }
        foreach ($w in $post.Warnings) { Write-Log $w 'WARN' }
        try {
            Get-TaggedRules | ForEach-Object { Remove-DnsClientNrptRule -Name $_.Name -Force }
            foreach ($b in $before) {
                Add-DnsClientNrptRule -Namespace $b.Namespace -NameServers $b.NameServers -DisplayName $b.DisplayName -Comment $b.Comment | Out-Null
            }
            Clear-DnsClientCache
            Write-Log 'Rollback complete' 'OK'
        } catch {
            Exit-Script 4 "Verification failed AND rollback failed: $($_.Exception.Message). Run with Action=Remove."
        }
        Exit-Script 5 'Post-change verification failed; restored previous state'
    }

    'Remove' {
        $rules = @(Get-TaggedRules)
        if ($rules.Count -eq 0) {
            Invoke-Verification | Out-Null
            Exit-Script 0 'No Altec-tagged rule present; nothing to remove'
        }
        try {
            foreach ($r in $rules) {
                Write-Log "Removing NRPT rule $($r.Name)"
                Remove-DnsClientNrptRule -Name $r.Name -Force
            }
            Clear-DnsClientCache
        } catch {
            Exit-Script 4 "Remove failed: $($_.Exception.Message)"
        }
        if (@(Get-TaggedRules).Count -ne 0) { Exit-Script 4 'Rule still present after removal' }
        $ok = Invoke-Verification
        if (-not $ok) { Exit-Script 5 'Rule removed, but Ninja hosts do not resolve through normal DNS (expected if local DNS is broken)' }
        Exit-Script 0 'Rule removed; Ninja hosts resolve through normal DNS'
    }
}
