<#
.SYNOPSIS
    Safely removes the DNSFilter roaming client (or white-labeled "DNS Agent") without
    stranding the machine on loopback DNS.

.DESCRIPTION
    The DNSFilter agent points adapter DNS at its own loopback listener (127.0.0.2 on current
    agents, 127.0.0.1 on some older/custom builds). Uninstalling first can leave adapters
    pointed at a listener that no longer exists. This script does it in a safe order:

      1. Discover    - agent products, services, and DNS config of EVERY interface
                       (present, disconnected, and hidden/ghost adapters via the registry).
      2. Snapshot    - writes a JSON snapshot of DNS + service state to the log folder.
      3. Stop agent  - disables and stops the agent services (watchdog "Service Manager"
                       first) so nothing re-applies loopback DNS. Nothing is uninstalled yet.
      4. Reset DNS   - DHCP adapters: static DNS cleared so DHCP-provided DNS takes over.
                       Static-IP adapters are NEVER blanked: loopback entries are removed,
                       remaining static servers are kept, or FallbackDnsServers are applied.
      5. Validate    - no loopback DNS on any Up adapter, N test names resolve through the
                       system resolver, and at least one DNS server answers a direct query.
                       FAIL -> automatic rollback (DNS + services restored), agent left
                       installed, script exits non-zero.
      6. Uninstall   - MSI uninstall by product code (not Win32_Product).
      7. Re-check    - the agent's uninstaller can restore its own pre-install DNS backup,
                       so DNS is re-planned, re-reset if needed, and re-validated.
      8. Residue     - leftover services, install folders, and HKLM\SOFTWARE\DNSFilter.

    SERVER MODE (automatic): on a Domain Controller (or any host running the DNS Server
    role) adapters are NOT reset to DHCP. Every adapter that carries the DNSFilter loopback
    or has a default gateway is set statically to the server's OWN IP (DcDnsMode=Self), or
    to partner DC(s) first and then itself (DcDnsMode=PartnerFirst). Partner DCs are found
    automatically from the _ldap._tcp.dc._msdcs SRV records while the agent is still
    resolving. IPv6 ::1 is left alone on servers. Before uninstall, the DC must also
    resolve its own AD domain through its own IP, or the script rolls back.

    Also fixes "stranded" machines: if the agent is already gone but loopback DNS remains,
    only the DNS reset + validation phases run.

    Default Mode is Audit (read-only). Nothing is changed unless Mode = Remediate.

.PARAMETER Mode
    Audit (default) or Remediate.

.PARAMETER ResetScope
    All (default): clear static DNS on every DHCP-enabled adapter, loopback or not.
    LoopbackOnly : only touch adapters whose DNS contains a loopback address.
    Static-IP adapters follow the static rules above in both scopes.

.PARAMETER FallbackDnsServers
    OPTIONAL override. Normally leave blank: DC IPs are discovered automatically.
    - Member servers with a static IP and loopback-only DNS: these servers are used
      (default: discovered DCs). If neither exists, the script stops before changes (exit 3).
    - DC with DcDnsMode=PartnerFirst: used as the partner list (default: discovered DCs).
    The legacy Ninja variable "Verify DC IP" ($env:verifydcip) is accepted as an alias.

.PARAMETER TestHostnames
    Comma-separated names used for validation.

.PARAMETER MinSuccessCount
    How many TestHostnames must resolve for validation to pass.

.PARAMETER ValidationRetries / RetryDelaySeconds
    Validation attempts and delay between them (gives DHCP DNS time to take effect).

.PARAMETER UninstallTimeoutMinutes
    Max time to wait for each msiexec uninstall.

.PARAMETER CleanupResidue
    true (default) / false. Remove leftover services, folders, and registry keys after a
    successful uninstall.

.PARAMETER DcDnsMode
    Server mode only. Self (default): DNS = this server's own IP.
    PartnerFirst: DNS = partner DC(s), then this server's own IP (Microsoft's recommended
    order for multi-DC domains; avoids a DC waiting on itself at boot).

.PARAMETER StatusCustomField
    Optional NinjaOne device custom field (text) to receive a one-line result.

.PARAMETER AgentDnsAddresses
    Loopback DNS addresses that belong to DNSFilter. Default: 127.0.0.1,127.0.0.2,::1
    Only these are treated as DNSFilter residue. Other loopback resolvers (Cloudflare WARP
    uses 127.0.2.2 / 127.0.2.3) are left alone.

.PARAMETER ExcludeAdapterRegex
    Regex matched against adapter name and description. Matching adapters are never
    changed and their DNS servers are not used for the "server answered" check. They
    still serve normal name resolution. Default covers Cloudflare WARP and common
    VPN/tunnel clients. Set to "none" to disable.

.PARAMETER LogDirectory
    Default: C:\ProgramData\Altec\Logs\DNSFilterRemoval

.NOTES
    Variable resolution: CLI parameter -> Ninja script variable (environment) -> default.
    Ninja script variables should use the parameter names above (mode, resetScope, etc.).

    EXIT CODES
      0  Success, or nothing to do
      1  Fatal: bad parameter, not elevated, or unexpected error
      2  Audit: remediation needed (agent present and/or loopback DNS found)
      3  Blocked: static-IP adapter with loopback-only DNS and no DC/fallback DNS available.
         No net change (services restored if they had been stopped).
      4  (unused; DCs are now handled automatically)
      5  DNS validation failed (or agent could not be stopped). Rolled back. Agent still
         installed and DNS restored to original state.
      6  CRITICAL: validation failed AND rollback failed or DNS still broken. Check now.
      7  Uninstall failed. DNS is on DHCP and validated; agent services left DISABLED.
      8  CRITICAL: DNS validation failed after uninstall. Agent is gone; fix DNS manually.
      9  Success, reboot required to finish uninstall.
     10  Uninstalled and DNS validated, but residue remains (service, process, folder,
         registry key, or scheduled task). See log.
#>

[CmdletBinding()]
param(
    [string]$Mode,
    [string]$ResetScope,
    [string]$FallbackDnsServers,
    [string]$TestHostnames,
    [string]$MinSuccessCount,
    [string]$ValidationRetries,
    [string]$RetryDelaySeconds,
    [string]$UninstallTimeoutMinutes,
    [string]$CleanupResidue,
    [string]$DcDnsMode,
    [string]$StatusCustomField,
    [string]$AgentDnsAddresses,
    [string]$ExcludeAdapterRegex,
    [string]$LogDirectory
)

# ---------------------------------------------------------------------------------------
# Relaunch in 64-bit PowerShell if started from a 32-bit host. A 32-bit process is
# redirected to WOW6432Node and would miss the 64-bit uninstall entries.
# ---------------------------------------------------------------------------------------
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $ps64 = Join-Path $env:WINDIR 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
    foreach ($k in $PSBoundParameters.Keys) {
        $argList += ('-{0}' -f $k)
        $argList += ('"{0}"' -f $PSBoundParameters[$k])
    }
    & $ps64 @argList
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Continue'
$script:CanRollback = $false

# ---------------------------------------------------------------------------------------
# Exit codes
# ---------------------------------------------------------------------------------------
$EXIT_OK              = 0
$EXIT_FATAL           = 1
$EXIT_AUDIT_ACTION    = 2
$EXIT_BLOCKED_STATIC  = 3
$EXIT_BLOCKED_DNSROLE = 4
$EXIT_ROLLED_BACK     = 5
$EXIT_ROLLBACK_FAILED = 6
$EXIT_UNINSTALL_FAIL  = 7
$EXIT_POST_DNS_FAIL   = 8
$EXIT_OK_REBOOT       = 9
$EXIT_OK_RESIDUE      = 10

# ---------------------------------------------------------------------------------------
# Vendor identifiers (edit here if DNSFilter renames anything)
# ---------------------------------------------------------------------------------------
$ProductNameRegex    = '^(DNSFilter Agent|DNS Agent)(\s|$)'
$KnownServiceNames   = @('DNSFilter Agent', 'DNSFilter Agent Service Manager', 'DNS Agent', 'DNS Agent Service Manager')
$ServiceDisplayRegex = '^(DNSFilter Agent|DNS Agent)'
$ResidueFolderNames = @('DNSFilter Agent', 'DNS Agent', 'DNSFilter')
$ResidueFolders = @()
foreach ($n in $ResidueFolderNames) {
    $ResidueFolders += (Join-Path $env:ProgramFiles $n)
    if (${env:ProgramFiles(x86)}) { $ResidueFolders += (Join-Path ${env:ProgramFiles(x86)} $n) }
}
$ResidueFolders += (Join-Path $env:ProgramData 'DNSFilter')
$ResidueTaskPattern    = '*dnsfilter*'
$ResidueProcessPattern = '*dnsfilter*'
$ResidueRegistryKeys = @('HKLM:\SOFTWARE\DNSFilter', 'HKLM:\SOFTWARE\WOW6432Node\DNSFilter')

$script:V4Root = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces'
$script:V6Root = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\Interfaces'
$script:NetClassRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}'

# ---------------------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------------------
function Resolve-Setting {
    param([string]$Name, [string]$CliValue, [string]$Default)
    if (-not [string]::IsNullOrWhiteSpace($CliValue)) { return $CliValue.Trim() }
    $envVal = [Environment]::GetEnvironmentVariable($Name)
    if (-not [string]::IsNullOrWhiteSpace($envVal)) { return $envVal.Trim() }
    return $Default
}

function ConvertTo-Bool {
    param([string]$Value)
    return ($Value -match '^(1|true|yes|y|on)$')
}

function ConvertTo-IntSetting {
    param([string]$Name, [string]$Value, [int]$Min)
    $n = 0
    if (-not [int]::TryParse($Value, [ref]$n) -or $n -lt $Min) {
        Write-Host ('[FATAL] {0} must be an integer >= {1}. Got: {2}' -f $Name, $Min, $Value)
        exit $EXIT_FATAL
    }
    return $n
}

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
    if ([string]::IsNullOrWhiteSpace($script:StatusField)) { return }
    if (Get-Command -Name 'Ninja-Property-Set' -ErrorAction SilentlyContinue) {
        try { Ninja-Property-Set $script:StatusField $Text | Out-Null } catch { Write-Log ('Could not set Ninja field {0}: {1}' -f $script:StatusField, $_.Exception.Message) 'WARN' }
    }
}

function Exit-Script {
    param([int]$Code, [string]$Summary)
    $lvl = 'OK'
    if ($Code -ne $EXIT_OK -and $Code -ne $EXIT_OK_REBOOT) { $lvl = 'ERROR' }
    if ($Code -eq $EXIT_AUDIT_ACTION) { $lvl = 'WARN' }
    Write-Log ('RESULT (exit {0}): {1}' -f $Code, $Summary) $lvl
    Set-NinjaStatus ('{0} | exit {1} | {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $Code, $Summary)
    exit $Code
}

function Test-IsLoopback {
    param([string]$Address)
    $a = $Address.Trim()
    if ($a -match '^(::ffff:)?127\.') { return $true }
    if ($a -eq '::1' -or $a -eq '0:0:0:0:0:0:0:1') { return $true }
    return $false
}

$script:AgentDns = @('127.0.0.1', '127.0.0.2', '::1')
$script:ExcludeRx = ''
function Test-IsAgentDns {
    # True only for DNSFilter's own loopback listeners, not other local resolvers (WARP etc.)
    param([string]$Address)
    return ($script:AgentDns -contains $Address.Trim())
}

function Test-IsExcludedAdapter {
    param([string]$Name, [string]$Description)
    if ([string]::IsNullOrWhiteSpace($script:ExcludeRx)) { return $false }
    return (($Name -match $script:ExcludeRx) -or ($Description -match $script:ExcludeRx))
}

function Get-ExcludedIfIndex {
    @(Get-NetAdapter -IncludeHidden -ErrorAction SilentlyContinue |
        Where-Object { Test-IsExcludedAdapter -Name $_.Name -Description $_.InterfaceDescription } |
        ForEach-Object { $_.ifIndex })
}

function Split-DnsList {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    $Value -split '[,\s]+' | Where-Object { $_ }
}

function Test-IsDefaultSiteLocal {
    # Windows lists fec0:0:0:ffff::1/2/3 when no IPv6 DNS is configured. Not real servers.
    param([string]$Address)
    return ($Address -match '^fec0:0:0:ffff::[123]$')
}

# ---------------------------------------------------------------------------------------
# Discovery
# ---------------------------------------------------------------------------------------
function Get-AgentProducts {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -and $_.DisplayName -match $ProductNameRegex } |
        ForEach-Object {
            [pscustomobject]@{
                DisplayName     = $_.DisplayName
                DisplayVersion  = $_.DisplayVersion
                KeyName         = $_.PSChildName
                UninstallString = $_.UninstallString
            }
        }
}

function Get-AgentServiceState {
    $svcs = @(Get-CimInstance -ClassName Win32_Service -ErrorAction SilentlyContinue |
        Where-Object { ($KnownServiceNames -contains $_.Name) -or ($_.DisplayName -match $ServiceDisplayRegex) -or ($_.PathName -like '*dnsfilter*') })
    foreach ($s in $svcs) {
        $delayed = $false
        $d = Get-ItemProperty -Path ('HKLM:\SYSTEM\CurrentControlSet\Services\{0}' -f $s.Name) -Name 'DelayedAutostart' -ErrorAction SilentlyContinue
        if ($d -and $d.DelayedAutostart -eq 1) { $delayed = $true }
        [pscustomobject]@{
            Name        = $s.Name
            DisplayName = $s.DisplayName
            StartMode   = $s.StartMode
            Delayed     = $delayed
            State       = $s.State
            IsManager   = ($s.Name -match 'Manager' -or $s.DisplayName -match 'Manager')
        }
    }
}

function Get-DnsInterfaceState {
    $adapters = @{}
    foreach ($a in @(Get-NetAdapter -IncludeHidden -ErrorAction SilentlyContinue)) {
        if ($a.InterfaceGuid) { $adapters[$a.InterfaceGuid.ToUpper()] = $a }
    }

    $guids = New-Object System.Collections.Generic.List[string]
    foreach ($root in @($script:V4Root, $script:V6Root)) {
        if (Test-Path $root) {
            foreach ($k in @(Get-ChildItem -Path $root -ErrorAction SilentlyContinue)) {
                $g = $k.PSChildName.ToUpper()
                if ($g -match '^\{[0-9A-F-]{36}\}$' -and -not $guids.Contains($g)) { $guids.Add($g) }
            }
        }
    }

    foreach ($g in $guids) {
        $p4 = Join-Path $script:V4Root $g
        $p6 = Join-Path $script:V6Root $g
        $r4 = $null; $r6 = $null
        if (Test-Path $p4) { $r4 = Get-ItemProperty -Path $p4 -ErrorAction SilentlyContinue }
        if (Test-Path $p6) { $r6 = Get-ItemProperty -Path $p6 -ErrorAction SilentlyContinue }

        $a = $adapters[$g]
        $name = $null
        if ($a) { $name = $a.Name }
        else {
            $c = Get-ItemProperty -Path ('{0}\{1}\Connection' -f $script:NetClassRoot, $g) -Name 'Name' -ErrorAction SilentlyContinue
            if ($c) { $name = $c.Name }
        }
        if (-not $name) { $name = '(unnamed)' }

        $enableDhcp = $true
        if ($r4 -and $null -ne $r4.EnableDHCP) { $enableDhcp = ([int]$r4.EnableDHCP -eq 1) }

        $desc = ''; $status = 'NotPresent'; $ifIndex = $null
        if ($a) { $desc = $a.InterfaceDescription; $status = [string]$a.Status; $ifIndex = $a.ifIndex }

        $v4ns = ''; $v4dhcp = ''; $v6ns = ''
        if ($r4) { $v4ns = [string]$r4.NameServer; $v4dhcp = [string]$r4.DhcpNameServer }
        if ($r6) { $v6ns = [string]$r6.NameServer }

        [pscustomobject]@{
            Guid             = $g
            Name             = $name
            Description      = $desc
            Present          = [bool]$a
            Status           = $status
            IfIndex          = $ifIndex
            EnableDhcp       = $enableDhcp
            HasV4Key         = [bool]$r4
            HasV6Key         = [bool]$r6
            V4NameServer     = $v4ns
            V6NameServer     = $v6ns
            V4DhcpNameServer = $v4dhcp
        }
    }
}

# ---------------------------------------------------------------------------------------
# Domain / server-mode helpers
# ---------------------------------------------------------------------------------------
function Get-DomainDcIps {
    # Must run while DNS still works (agent running), so it is called before any change.
    param([string]$Domain)
    $ips = @()
    if ([string]::IsNullOrWhiteSpace($Domain)) { return }
    try {
        $srv = @(Resolve-DnsName -Name ('_ldap._tcp.dc._msdcs.{0}' -f $Domain) -Type SRV -DnsOnly -QuickTimeout -ErrorAction Stop)
        $ips += @($srv | Where-Object { $_.Type -eq 'A' } | ForEach-Object { $_.IPAddress })
        foreach ($t in @($srv | Where-Object { $_.Type -eq 'SRV' } | ForEach-Object { $_.NameTarget } | Sort-Object -Unique)) {
            try { $ips += @(Resolve-DnsName -Name $t -Type A -DnsOnly -QuickTimeout -ErrorAction Stop | Where-Object { $_.Type -eq 'A' } | ForEach-Object { $_.IPAddress }) } catch { }
        }
    }
    catch { Write-Log ('DC discovery for {0} failed: {1}' -f $Domain, $_.Exception.Message) 'WARN' }
    $ips | Where-Object { $_ -and -not (Test-IsLoopback $_) } | Sort-Object -Unique
}

function Get-LocalIPv4 {
    Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notmatch '^(127\.|169\.254\.)' } |
        ForEach-Object { $_.IPAddress } | Sort-Object -Unique
}

function Get-SelfDnsMap {
    # Server mode: target DNS list per adapter GUID. Only adapters that have a default
    # gateway or currently carry loopback DNS are targeted (skips iSCSI/heartbeat NICs).
    param([string]$Mode, [string[]]$Partners)
    $map = @{}
    $allLocal = @(Get-LocalIPv4)
    foreach ($a in @(Get-NetAdapter -ErrorAction SilentlyContinue)) {
        $ips = @(Get-NetIPAddress -InterfaceIndex $a.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -notmatch '^(127\.|169\.254\.)' } | ForEach-Object { $_.IPAddress })
        if ($ips.Count -eq 0) { continue }
        $hasGw = @(Get-NetRoute -InterfaceIndex $a.ifIndex -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue).Count -gt 0
        $dns = @(Get-DnsClientServerAddress -InterfaceIndex $a.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | ForEach-Object { $_.ServerAddresses })
        if (Test-IsExcludedAdapter -Name $a.Name -Description $a.InterfaceDescription) { continue }
        $hasLoop = @($dns | Where-Object { Test-IsAgentDns $_ }).Count -gt 0
        if (-not ($hasGw -or $hasLoop)) { continue }
        $list = @()
        if ($Mode -eq 'PartnerFirst') { $list += @($Partners | Where-Object { $allLocal -notcontains $_ }) }
        $list += $ips[0]
        $map[$a.InterfaceGuid.ToUpper()] = $list
        Write-Log ('Server mode target: "{0}" -> [{1}]' -f $a.Name, ($list -join ','))
    }
    return $map
}

# ---------------------------------------------------------------------------------------
# Planning: decide the new DNS list per interface, per address family
# ---------------------------------------------------------------------------------------
function Get-DnsPlan {
    param($State, [string]$Scope, [string[]]$Fallback, [bool]$ServerMode, $SelfMap)
    foreach ($i in $State) {
        $v4 = @(Split-DnsList $i.V4NameServer)
        $v6 = @(Split-DnsList $i.V6NameServer)
        $v4Loop = @($v4 | Where-Object { Test-IsAgentDns $_ }).Count -gt 0
        # ::1 is normal on a DC / DNS server; only treat it as DNSFilter residue on clients
        $v6Loop = @($v6 | Where-Object { (Test-IsAgentDns $_) -and -not ($ServerMode -and $_ -eq '::1') }).Count -gt 0
        # Another local resolver (e.g. Cloudflare WARP 127.0.2.2/.3) owns DNS on this adapter
        $v4Other = @($v4 | Where-Object { (Test-IsLoopback $_) -and -not (Test-IsAgentDns $_) }).Count -gt 0
        $v6Other = @($v6 | Where-Object { (Test-IsLoopback $_) -and -not (Test-IsAgentDns $_) }).Count -gt 0
        $excluded = Test-IsExcludedAdapter -Name $i.Name -Description $i.Description
        if ($excluded) {
            [pscustomobject]@{
                Iface = $i; Guid = $i.Guid
                V4Old = $v4; V4New = $v4; V4Action = 'Excluded'
                V6Old = $v6; V6New = $v6; V6Action = 'Excluded'
                V4Loopback = $false; V6Loopback = $false; Changed = $false; Blocked = $false
            }
            continue
        }

        $self = $null
        if ($ServerMode -and $SelfMap -and $SelfMap.ContainsKey($i.Guid)) { $self = @($SelfMap[$i.Guid]) }

        # IPv4
        $newV4 = $v4; $act4 = 'None'
        if ($self) {
            if (($v4 -join ',') -ne ($self -join ',')) { $newV4 = $self; $act4 = 'SetSelf' } else { $act4 = 'Keep' }
        }
        elseif ($v4.Count -gt 0) {
            if ($v4Other) {
                # Leave the other resolver's entries; strip only DNSFilter's
                if ($v4Loop) { $newV4 = @($v4 | Where-Object { -not (Test-IsAgentDns $_) }); $act4 = 'RemoveLoopback' } else { $act4 = 'KeepOtherResolver' }
            }
            elseif ((-not $ServerMode) -and $i.EnableDhcp -and ($Scope -eq 'All' -or $v4Loop)) {
                $newV4 = @(); $act4 = 'ResetToDhcp'
            }
            elseif ($v4Loop) {
                # Static adapter (or non-targeted server adapter) carrying loopback
                $remain = @($v4 | Where-Object { -not (Test-IsAgentDns $_) })
                if ($remain.Count -gt 0)                  { $newV4 = $remain;   $act4 = 'RemoveLoopback' }
                elseif ($ServerMode -and -not $i.Present) { $newV4 = @();       $act4 = 'RemoveLoopback' }
                elseif ($Fallback.Count -gt 0)            { $newV4 = $Fallback; $act4 = 'SetFallback' }
                else                                      { $act4 = 'BLOCKED' }
            }
            else { $act4 = 'Keep' }
        }

        # IPv6
        $newV6 = $v6; $act6 = 'None'
        if ($v6.Count -gt 0) {
            if ($v6Loop) {
                $newV6 = @($v6 | Where-Object { -not ((Test-IsAgentDns $_) -and -not ($ServerMode -and $_ -eq '::1')) })
                if ($newV6.Count -gt 0) { $act6 = 'RemoveLoopback' } else { $act6 = 'ResetToAuto' }
            }
            elseif ($v6Other) { $act6 = 'KeepOtherResolver' }
            elseif ((-not $ServerMode) -and $Scope -eq 'All' -and $i.EnableDhcp) { $newV6 = @(); $act6 = 'ResetToAuto' }
            else { $act6 = 'Keep' }
        }

        $changed = ($act4 -in @('ResetToDhcp', 'RemoveLoopback', 'SetFallback', 'SetSelf')) -or ($act6 -in @('ResetToAuto', 'RemoveLoopback'))

        [pscustomobject]@{
            Iface      = $i
            Guid       = $i.Guid
            V4Old      = $v4
            V4New      = $newV4
            V4Action   = $act4
            V6Old      = $v6
            V6New      = $newV6
            V6Action   = $act6
            V4Loopback = $v4Loop
            V6Loopback = $v6Loop
            Changed    = $changed
            Blocked    = ($act4 -eq 'BLOCKED')
        }
    }
}

function Write-Plan {
    param($Plan, [string]$Title)
    Write-Log ('--- {0} ---' -f $Title)
    $shown = 0
    foreach ($p in $Plan) {
        $i = $p.Iface
        if ($p.V4Old.Count -eq 0 -and $p.V6Old.Count -eq 0 -and -not $i.Present) { continue }
        $shown++
        $lvl = 'INFO'
        if ($p.Blocked) { $lvl = 'ERROR' } elseif ($p.V4Loopback -or $p.V6Loopback) { $lvl = 'WARN' }
        Write-Log ('[{0}] "{1}" ({2}) DHCP={3}' -f $i.Status, $i.Name, $i.Description, $i.EnableDhcp) $lvl
        Write-Log ('      IPv4: [{0}] -> {1} [{2}]   DHCP offers: [{3}]' -f ($p.V4Old -join ','), $p.V4Action, ($p.V4New -join ','), $i.V4DhcpNameServer) $lvl
        if ($p.V6Old.Count -gt 0) {
            Write-Log ('      IPv6: [{0}] -> {1} [{2}]' -f ($p.V6Old -join ','), $p.V6Action, ($p.V6New -join ',')) $lvl
        }
        if ($p.V4Action -eq 'ResetToDhcp' -and $i.Status -eq 'Up' -and [string]::IsNullOrWhiteSpace($i.V4DhcpNameServer)) {
            Write-Log '      NOTE: Up adapter has no DHCP-provided DNS yet. Validation will decide.' 'WARN'
        }
    }
    if ($shown -eq 0) { Write-Log 'No interfaces with DNS configuration found.' }
}

# ---------------------------------------------------------------------------------------
# Apply DNS: registry first (covers hidden/ghost adapters), then the cmdlet for present
# adapters so the TCP/IP stack picks up the change immediately.
# ---------------------------------------------------------------------------------------
function Set-InterfaceDns {
    param($Iface, [string[]]$V4, [string[]]$V6)
    $V4 = @($V4 | Where-Object { $_ })
    $V6 = @($V6 | Where-Object { $_ })

    if ($Iface.HasV4Key) {
        New-ItemProperty -Path (Join-Path $script:V4Root $Iface.Guid) -Name 'NameServer' -Value ($V4 -join ',') -PropertyType String -Force -ErrorAction Stop | Out-Null
    }
    if ($Iface.HasV6Key) {
        New-ItemProperty -Path (Join-Path $script:V6Root $Iface.Guid) -Name 'NameServer' -Value ($V6 -join ',') -PropertyType String -Force -ErrorAction Stop | Out-Null
    }

    if ($Iface.Present -and $null -ne $Iface.IfIndex) {
        try {
            $all = @($V4) + @($V6)
            if ($all.Count -eq 0) {
                Set-DnsClientServerAddress -InterfaceIndex $Iface.IfIndex -ResetServerAddresses -ErrorAction Stop
            }
            else {
                Set-DnsClientServerAddress -InterfaceIndex $Iface.IfIndex -ServerAddresses $all -ErrorAction Stop
            }
        }
        catch {
            Write-Log ('      Cmdlet update failed on "{0}" (registry already updated): {1}' -f $Iface.Name, $_.Exception.Message) 'WARN'
        }
    }
}

function Invoke-DnsPlan {
    param($Plan)
    $ok = $true
    foreach ($p in @($Plan | Where-Object { $_.Changed })) {
        try {
            Set-InterfaceDns -Iface $p.Iface -V4 $p.V4New -V6 $p.V6New
            Write-Log ('Applied: "{0}" IPv4 {1} [{2}] / IPv6 {3} [{4}]' -f $p.Iface.Name, $p.V4Action, ($p.V4New -join ','), $p.V6Action, ($p.V6New -join ',')) 'OK'
        }
        catch {
            Write-Log ('FAILED to apply DNS on "{0}": {1}' -f $p.Iface.Name, $_.Exception.Message) 'ERROR'
            $ok = $false
        }
    }
    Clear-DnsClientCache -ErrorAction SilentlyContinue
    return $ok
}

# ---------------------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------------------
function Test-DnsHealth {
    param(
        [string[]]$Hosts,
        [int]$MinSuccess,
        [int]$Retries,
        [int]$DelaySec,
        [bool]$RequireNoLoopback,
        [string]$Label,
        [string[]]$AllowLoopback = @()
    )
    if ($MinSuccess -gt $Hosts.Count) { $MinSuccess = $Hosts.Count }

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        Clear-DnsClientCache -ErrorAction SilentlyContinue
        $exIdx = @(Get-ExcludedIfIndex)
        $upIdx = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' -and $exIdx -notcontains $_.ifIndex } | ForEach-Object { $_.ifIndex })
        $servers = @(Get-DnsClientServerAddress -ErrorAction SilentlyContinue |
            Where-Object { $upIdx -contains $_.InterfaceIndex } |
            ForEach-Object { $_.ServerAddresses } |
            Where-Object { $_ -and -not (Test-IsDefaultSiteLocal $_) } |
            Sort-Object -Unique)
        $loop = @($servers | Where-Object { (Test-IsAgentDns $_) -and ($AllowLoopback -notcontains $_) })

        $prefix = '[{0} {1}/{2}]' -f $Label, $attempt, $Retries
        if ($upIdx.Count -eq 0) {
            Write-Log ('{0} No non-tunnel adapters are Up.' -f $prefix) 'WARN'
        }
        elseif ($RequireNoLoopback -and $loop.Count -gt 0) {
            Write-Log ('{0} Loopback DNS still active on an Up adapter: {1}' -f $prefix, ($loop -join ',')) 'WARN'
        }
        elseif ($servers.Count -eq 0) {
            Write-Log ('{0} Up adapters have no DNS servers (waiting on DHCP?).' -f $prefix) 'WARN'
        }
        else {
            $resolved = 0
            foreach ($h in $Hosts) {
                try {
                    $r = @(Resolve-DnsName -Name $h -DnsOnly -NoHostsFile -QuickTimeout -ErrorAction Stop | Where-Object { $_.IPAddress })
                    if ($r.Count -gt 0) { $resolved++; Write-Log ('{0}   OK   {1} -> {2}' -f $prefix, $h, $r[0].IPAddress) }
                    else { Write-Log ('{0}   FAIL {1} (no address records)' -f $prefix, $h) 'WARN' }
                }
                catch { Write-Log ('{0}   FAIL {1} ({2})' -f $prefix, $h, $_.Exception.Message) 'WARN' }
            }

            $serverOk = 0
            foreach ($s in $servers) {
                try {
                    Resolve-DnsName -Name $Hosts[0] -Server $s -DnsOnly -QuickTimeout -ErrorAction Stop | Out-Null
                    $serverOk++
                    Write-Log ('{0}   Server {1} answered' -f $prefix, $s)
                }
                catch { Write-Log ('{0}   Server {1} did not answer' -f $prefix, $s) 'WARN' }
            }

            Write-Log ('{0} Resolved {1}/{2} (need {3}); servers answering {4}/{5}' -f $prefix, $resolved, $Hosts.Count, $MinSuccess, $serverOk, $servers.Count)
            if ($resolved -ge $MinSuccess -and $serverOk -ge 1) {
                Write-Log ('{0} PASS. Active DNS: {1}' -f $prefix, ($servers -join ',')) 'OK'
                return $true
            }
        }
        if ($attempt -lt $Retries) { Start-Sleep -Seconds $DelaySec }
    }
    Write-Log ('[{0}] FAIL after {1} attempt(s).' -f $Label, $Retries) 'ERROR'
    return $false
}

# ---------------------------------------------------------------------------------------
# Services
# ---------------------------------------------------------------------------------------
function Stop-AgentServices {
    param($Services)
    # Watchdog/manager first so it cannot restart the agent
    $ordered = @($Services | Sort-Object { if ($_.IsManager) { 0 } else { 1 } })
    foreach ($s in $ordered) {
        try { Set-Service -Name $s.Name -StartupType Disabled -ErrorAction Stop; Write-Log ('Disabled service "{0}"' -f $s.Name) }
        catch { Write-Log ('Could not disable "{0}": {1}' -f $s.Name, $_.Exception.Message) 'WARN' }
    }
    foreach ($s in $ordered) {
        Stop-Service -Name $s.Name -Force -ErrorAction SilentlyContinue
    }

    $deadline = (Get-Date).AddSeconds(45)
    do {
        $running = @($ordered | Where-Object { (Get-Service -Name $_.Name -ErrorAction SilentlyContinue).Status -ne 'Stopped' })
        if ($running.Count -eq 0) { break }
        Start-Sleep -Seconds 2
    } while ((Get-Date) -lt $deadline)

    foreach ($s in $running) {
        $svc = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $s.Name) -ErrorAction SilentlyContinue
        if ($svc -and $svc.ProcessId -gt 0) {
            Write-Log ('Service "{0}" did not stop; terminating PID {1}' -f $s.Name, $svc.ProcessId) 'WARN'
            Stop-Process -Id $svc.ProcessId -Force -ErrorAction SilentlyContinue
        }
    }
    Start-Sleep -Seconds 3
    $still = @($ordered | Where-Object { (Get-Service -Name $_.Name -ErrorAction SilentlyContinue).Status -ne 'Stopped' })
    foreach ($s in $still) { Write-Log ('Service "{0}" is still running.' -f $s.Name) 'ERROR' }
    return ($still.Count -eq 0)
}

function Restore-AgentServices {
    param($Services)
    $ok = $true
    foreach ($s in $Services) {
        $start = 'auto'
        switch ($s.StartMode) {
            'Auto'     { if ($s.Delayed) { $start = 'delayed-auto' } else { $start = 'auto' } }
            'Manual'   { $start = 'demand' }
            'Disabled' { $start = 'disabled' }
        }
        & sc.exe config "$($s.Name)" start= $start | Out-Null
        if ($LASTEXITCODE -ne 0) { $ok = $false; Write-Log ('sc config failed for "{0}"' -f $s.Name) 'ERROR' }
    }
    # Agent first, then manager
    foreach ($s in @($Services | Sort-Object { if ($_.IsManager) { 1 } else { 0 } })) {
        if ($s.State -eq 'Running') {
            try { Start-Service -Name $s.Name -ErrorAction Stop; Write-Log ('Started "{0}"' -f $s.Name) 'OK' }
            catch { $ok = $false; Write-Log ('Could not start "{0}": {1}' -f $s.Name, $_.Exception.Message) 'ERROR' }
        }
    }
    return $ok
}

# ---------------------------------------------------------------------------------------
# Rollback: put DNS back exactly as snapshotted, restore services, re-check DNS
# ---------------------------------------------------------------------------------------
function Invoke-Rollback {
    param($Snapshot, $Services, [bool]$BaselineOk)
    Write-Log '=== ROLLBACK: restoring original DNS settings and agent services ===' 'WARN'
    $ok = $true
    $snapIndex = @{}
    foreach ($s in $Snapshot) { $snapIndex[$s.Guid] = $s }

    foreach ($c in @(Get-DnsInterfaceState)) {
        $o = $snapIndex[$c.Guid]
        if (-not $o) { continue }
        if (Test-IsExcludedAdapter -Name $c.Name -Description $c.Description) { continue }   # tunnel manages its own DNS
        if ($c.V4NameServer -ne $o.V4NameServer -or $c.V6NameServer -ne $o.V6NameServer) {
            try {
                Set-InterfaceDns -Iface $c -V4 @(Split-DnsList $o.V4NameServer) -V6 @(Split-DnsList $o.V6NameServer)
                Write-Log ('Restored "{0}" IPv4 [{1}] IPv6 [{2}]' -f $c.Name, $o.V4NameServer, $o.V6NameServer) 'OK'
            }
            catch { $ok = $false; Write-Log ('Restore FAILED on "{0}": {1}' -f $c.Name, $_.Exception.Message) 'ERROR' }
        }
    }

    if ($Services.Count -gt 0) {
        if (-not (Restore-AgentServices -Services $Services)) { $ok = $false }
        Start-Sleep -Seconds 10
    }
    Clear-DnsClientCache -ErrorAction SilentlyContinue

    $dnsOk = Test-DnsHealth -Hosts $script:Hosts -MinSuccess $script:MinSuccess -Retries 3 -DelaySec 5 -RequireNoLoopback $false -Label 'Post-rollback'
    if (-not $dnsOk -and -not $BaselineOk) {
        Write-Log 'DNS was already failing before changes; rollback returned the machine to its original state.' 'WARN'
        $dnsOk = $true
    }
    return ($ok -and $dnsOk)
}

# =======================================================================================
# MAIN
# =======================================================================================
$Mode                    = Resolve-Setting 'mode'                    $Mode                    'Audit'
$ResetScope              = Resolve-Setting 'resetScope'              $ResetScope              'All'
$FallbackDnsServers      = Resolve-Setting 'fallbackDnsServers'      $FallbackDnsServers      ''
if ([string]::IsNullOrWhiteSpace($FallbackDnsServers)) { $FallbackDnsServers = Resolve-Setting 'verifydcip' '' '' }
$TestHostnames           = Resolve-Setting 'testHostnames'           $TestHostnames           'www.microsoft.com,www.google.com,www.cloudflare.com'
$MinSuccessCount         = Resolve-Setting 'minSuccessCount'         $MinSuccessCount         '2'
$ValidationRetries       = Resolve-Setting 'validationRetries'       $ValidationRetries       '6'
$RetryDelaySeconds       = Resolve-Setting 'retryDelaySeconds'       $RetryDelaySeconds       '5'
$UninstallTimeoutMinutes = Resolve-Setting 'uninstallTimeoutMinutes' $UninstallTimeoutMinutes '10'
$CleanupResidue          = Resolve-Setting 'cleanupResidue'          $CleanupResidue          'true'
$DcDnsMode               = Resolve-Setting 'dcDnsMode'               $DcDnsMode               'Self'
$script:StatusField      = Resolve-Setting 'statusCustomField'       $StatusCustomField       ''
$AgentDnsAddresses       = Resolve-Setting 'agentDnsAddresses'       $AgentDnsAddresses       '127.0.0.1,127.0.0.2,::1'
$ExcludeAdapterRegex     = Resolve-Setting 'excludeAdapterRegex'     $ExcludeAdapterRegex     'Cloudflare WARP|CloudflareWARP|WireGuard|Wintun|Tailscale|ZeroTier|OpenVPN|TAP-Windows|PANGP|GlobalProtect|AnyConnect|Fortinet|FortiClient|SonicWall|NetExtender|Juniper|Pulse Secure|Ivanti|Zscaler|NordLynx'
$LogDirectory            = Resolve-Setting 'logDirectory'            $LogDirectory            (Join-Path $env:ProgramData 'Altec\Logs\DNSFilterRemoval')

if ($Mode -notin @('Audit', 'Remediate'))          { Write-Host "[FATAL] Mode must be Audit or Remediate. Got: $Mode"; exit $EXIT_FATAL }
if ($ResetScope -notin @('All', 'LoopbackOnly'))   { Write-Host "[FATAL] ResetScope must be All or LoopbackOnly. Got: $ResetScope"; exit $EXIT_FATAL }

$script:AgentDns = @($AgentDnsAddresses -split '[,;\s]+' | Where-Object { $_ })
if ($script:AgentDns.Count -eq 0) { Write-Host '[FATAL] AgentDnsAddresses is empty.'; exit $EXIT_FATAL }
foreach ($a in $script:AgentDns) { if (-not (Test-IsLoopback $a)) { Write-Host "[FATAL] AgentDnsAddresses must be loopback addresses. Got: $a"; exit $EXIT_FATAL } }
if ($ExcludeAdapterRegex -eq 'none') { $script:ExcludeRx = '' } else {
    try { [void][regex]::new($ExcludeAdapterRegex); $script:ExcludeRx = $ExcludeAdapterRegex }
    catch { Write-Host "[FATAL] ExcludeAdapterRegex is not a valid regex: $ExcludeAdapterRegex"; exit $EXIT_FATAL }
}

$fallbackList = @()
foreach ($s in @($FallbackDnsServers -split '[,;\s]+' | Where-Object { $_ })) {
    $ip = $null
    if ([System.Net.IPAddress]::TryParse($s, [ref]$ip) -and $ip.AddressFamily -eq 'InterNetwork' -and -not (Test-IsLoopback $s)) { $fallbackList += $s }
    else { Write-Host "[FATAL] Invalid FallbackDnsServers entry (IPv4, non-loopback only): $s"; exit $EXIT_FATAL }
}

$script:Hosts      = @($TestHostnames -split '[,;\s]+' | Where-Object { $_ })
if ($script:Hosts.Count -eq 0) { Write-Host '[FATAL] TestHostnames is empty.'; exit $EXIT_FATAL }
$script:MinSuccess = ConvertTo-IntSetting 'MinSuccessCount' $MinSuccessCount 1
$retries           = ConvertTo-IntSetting 'ValidationRetries' $ValidationRetries 1
$delay             = ConvertTo-IntSetting 'RetryDelaySeconds' $RetryDelaySeconds 1
$uninstallTimeout  = ConvertTo-IntSetting 'UninstallTimeoutMinutes' $UninstallTimeoutMinutes 1
$doCleanup         = ConvertTo-Bool $CleanupResidue
if ($DcDnsMode -notin @('Self', 'PartnerFirst')) { Write-Host "[FATAL] DcDnsMode must be Self or PartnerFirst. Got: $DcDnsMode"; exit $EXIT_FATAL }

try {
    if (-not (Test-Path $LogDirectory)) { New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:LogFile = Join-Path $LogDirectory ('DNSFilterRemoval-{0}-{1}.log' -f $env:COMPUTERNAME, $stamp)
}
catch { Write-Host "[WARN] Could not create log directory $LogDirectory. Console logging only." }

try {
    Write-Log ('=== DNSFilter safe removal | {0} | Mode={1} | Scope={2} ===' -f $env:COMPUTERNAME, $Mode, $ResetScope)
    Write-Log ('DNSFilter DNS addresses: [{0}] | Excluded adapters regex: [{1}]' -f ($script:AgentDns -join ','), $script:ExcludeRx)
    Write-Log ('Test names: {0} (need {1}) | Retries {2} x {3}s | Fallback DNS: [{4}]' -f ($script:Hosts -join ','), $script:MinSuccess, $retries, $delay, ($fallbackList -join ','))

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($Mode -eq 'Remediate' -and -not $isAdmin) { Exit-Script $EXIT_FATAL 'Remediate requires elevation (run as SYSTEM or admin).' }

    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem
    $isDC    = ($os.ProductType -eq 2)
    $dnsRole = [bool](Get-Service -Name 'DNS' -ErrorAction SilentlyContinue)
    $script:ServerMode = ($isDC -or $dnsRole)
    $script:AllowLoop = @()
    if ($script:ServerMode) {
        $script:AllowLoop = @('::1')
        Write-Log ('SERVER MODE: DC={0}, DNS Server role={1}. Adapters will point to this server (DcDnsMode={2}), not DHCP.' -f $isDC, $dnsRole, $DcDnsMode) 'WARN'
    }

    # DC discovery while DNS still works through the agent
    $discoveredDcs = @()
    if ($cs.PartOfDomain) {
        $discoveredDcs = @(Get-DomainDcIps -Domain $cs.Domain)
        Write-Log ('Discovered DCs for {0}: [{1}]' -f $cs.Domain, ($discoveredDcs -join ','))
    }
    $selfIps = @(Get-LocalIPv4)
    $partners = @($fallbackList)
    if ($partners.Count -eq 0) { $partners = @($discoveredDcs | Where-Object { $selfIps -notcontains $_ }) }
    $memberFallback = @($fallbackList)
    if ($memberFallback.Count -eq 0 -and -not $script:ServerMode) { $memberFallback = @($discoveredDcs) }
    $selfMap = @{}
    if ($script:ServerMode) {
        if ($DcDnsMode -eq 'PartnerFirst' -and $partners.Count -eq 0) { Write-Log 'PartnerFirst requested but no partner DCs found; using self only.' 'WARN' }
        $selfMap = Get-SelfDnsMap -Mode $DcDnsMode -Partners $partners
        if ($selfMap.Count -eq 0) { Write-Log 'Server mode: no adapter with an IPv4 address and gateway/loopback DNS was found.' 'WARN' }
    }

    # ---- Discovery -------------------------------------------------------------------
    $products = @(Get-AgentProducts)
    $services = @(Get-AgentServiceState)
    $snapshot = @(Get-DnsInterfaceState)

    if ($products.Count -eq 0) { Write-Log 'No DNSFilter/DNS Agent product registered.' }
    foreach ($p in $products) { Write-Log ('Product: {0} {1} [{2}]' -f $p.DisplayName, $p.DisplayVersion, $p.KeyName) }
    foreach ($s in $services) { Write-Log ('Service: "{0}" ({1}) State={2} Start={3} Delayed={4}' -f $s.Name, $s.DisplayName, $s.State, $s.StartMode, $s.Delayed) }

    $plan = @(Get-DnsPlan -State $snapshot -Scope $ResetScope -Fallback $memberFallback -ServerMode $script:ServerMode -SelfMap $selfMap)
    Write-Plan -Plan $plan -Title 'DNS plan (current state)'

    $loopbackFound = @($plan | Where-Object { $_.V4Loopback -or $_.V6Loopback }).Count -gt 0
    $blocked       = @($plan | Where-Object { $_.Blocked })

    if ($script:LogFile) {
        $snapFile = Join-Path $LogDirectory ('DNSFilterRemoval-{0}-{1}-snapshot.json' -f $env:COMPUTERNAME, $stamp)
        [pscustomobject]@{ Computer = $env:COMPUTERNAME; Taken = (Get-Date).ToString('s'); Products = $products; Services = $services; Interfaces = $snapshot } |
            ConvertTo-Json -Depth 5 | Out-File -FilePath $snapFile -Encoding ASCII
        Write-Log ('Snapshot saved: {0}' -f $snapFile)
    }

    if ($products.Count -eq 0 -and $services.Count -eq 0 -and ((-not $loopbackFound) -or $script:ServerMode)) {
        # On a DC/DNS server with no agent, 127.0.0.1 is a legitimate setting; leave it alone.
        Exit-Script $EXIT_OK 'DNSFilter not present and no DNSFilter loopback DNS. Nothing to do.'
    }

    # ---- Audit -----------------------------------------------------------------------
    if ($Mode -eq 'Audit') {
        if ($blocked.Count -gt 0) { Write-Log ('Remediate would be BLOCKED: static-IP adapter(s) with loopback-only DNS and no DC/fallback DNS available: {0}' -f (($blocked | ForEach-Object { $_.Iface.Name }) -join ', ')) 'WARN' }
        Exit-Script $EXIT_AUDIT_ACTION ('Audit: agent products={0}, services={1}, loopback DNS={2}. Run with Mode=Remediate.' -f $products.Count, $services.Count, $loopbackFound)
    }

    # ---- Remediate guards ------------------------------------------------------------
    if ($blocked.Count -gt 0) {
        Exit-Script $EXIT_BLOCKED_STATIC ('Static-IP adapter(s) with loopback-only DNS and no DC/fallback DNS available: {0}. No changes made.' -f (($blocked | ForEach-Object { $_.Iface.Name }) -join ', '))
    }

    # ---- Baseline (informational) ----------------------------------------------------
    $baselineOk = Test-DnsHealth -Hosts $script:Hosts -MinSuccess $script:MinSuccess -Retries 1 -DelaySec 1 -RequireNoLoopback $false -Label 'Baseline'
    if (-not $baselineOk) { Write-Log 'Baseline resolution is already failing through the agent. Continuing; post-change validation is the gate.' 'WARN' }

    # ---- Stop agent ------------------------------------------------------------------
    $script:CanRollback = $true
    $script:RbSnapshot = $snapshot; $script:RbServices = $services; $script:RbBaseline = $baselineOk
    if ($services.Count -gt 0) {
        Write-Log '=== Phase: stop and disable agent services ==='
        if (-not (Stop-AgentServices -Services $services)) {
            $rb = Invoke-Rollback -Snapshot $snapshot -Services $services -BaselineOk $baselineOk
            if ($rb) { Exit-Script $EXIT_ROLLED_BACK 'Could not stop agent services. Rolled back; agent still installed.' }
            Exit-Script $EXIT_ROLLBACK_FAILED 'Could not stop agent services AND rollback had errors. Check DNS now.'
        }
        Start-Sleep -Seconds 5   # agent may restore its own DNS backup on stop; let it settle
    }

    # ---- Reset DNS -------------------------------------------------------------------
    Write-Log '=== Phase: reset DNS ==='
    $plan2 = @(Get-DnsPlan -State @(Get-DnsInterfaceState) -Scope $ResetScope -Fallback $memberFallback -ServerMode $script:ServerMode -SelfMap $selfMap)
    Write-Plan -Plan $plan2 -Title 'DNS plan (after agent stop)'
    if (@($plan2 | Where-Object { $_.Blocked }).Count -gt 0) {
        $rb = Invoke-Rollback -Snapshot $snapshot -Services $services -BaselineOk $baselineOk
        if ($rb) { Exit-Script $EXIT_BLOCKED_STATIC 'Static-IP adapter with loopback-only DNS appeared after agent stop and no DC/fallback DNS available. Rolled back.' }
        Exit-Script $EXIT_ROLLBACK_FAILED 'Blocked on static adapter AND rollback had errors. Check DNS now.'
    }

    $applyOk = Invoke-DnsPlan -Plan $plan2
    if (-not $applyOk) {
        $rb = Invoke-Rollback -Snapshot $snapshot -Services $services -BaselineOk $baselineOk
        if ($rb) { Exit-Script $EXIT_ROLLED_BACK 'Failed to apply DNS changes. Rolled back; agent still installed.' }
        Exit-Script $EXIT_ROLLBACK_FAILED 'Failed to apply DNS changes AND rollback had errors. Check DNS now.'
    }

    # ---- Validate (gate) -------------------------------------------------------------
    Write-Log '=== Phase: validate DNS before uninstall ==='
    $valid = Test-DnsHealth -Hosts $script:Hosts -MinSuccess $script:MinSuccess -Retries $retries -DelaySec $delay -RequireNoLoopback $true -Label 'Pre-uninstall' -AllowLoopback $script:AllowLoop
    if (-not $valid) {
        $rb = Invoke-Rollback -Snapshot $snapshot -Services $services -BaselineOk $baselineOk
        if ($rb) { Exit-Script $EXIT_ROLLED_BACK 'DNS validation failed after reset. Rolled back; agent still installed.' }
        Exit-Script $EXIT_ROLLBACK_FAILED 'DNS validation failed AND rollback failed. Check DNS now.'
    }

    if ($isDC) {
        # A DC pointed at itself must answer for its own AD zone and DC locator records
        $selfOk = $false
        foreach ($ip in @($selfMap.Values | ForEach-Object { $_ } | Where-Object { $selfIps -contains $_ } | Sort-Object -Unique)) {
            try {
                Resolve-DnsName -Name $cs.Domain -Server $ip -DnsOnly -QuickTimeout -ErrorAction Stop | Out-Null
                Resolve-DnsName -Name ('_ldap._tcp.dc._msdcs.{0}' -f $cs.Domain) -Type SRV -Server $ip -DnsOnly -QuickTimeout -ErrorAction Stop | Out-Null
                Write-Log ('DC self-check PASS: {0} answers for {1} and its DC SRV records.' -f $ip, $cs.Domain) 'OK'
                $selfOk = $true
            }
            catch { Write-Log ('DC self-check FAIL on {0}: {1}' -f $ip, $_.Exception.Message) 'ERROR' }
        }
        if (-not $selfOk) {
            $rb = Invoke-Rollback -Snapshot $snapshot -Services $services -BaselineOk $baselineOk
            if ($rb) { Exit-Script $EXIT_ROLLED_BACK 'DC could not resolve its own AD domain via its own IP. Rolled back; agent still installed.' }
            Exit-Script $EXIT_ROLLBACK_FAILED 'DC self-resolution failed AND rollback failed. Check DNS now.'
        }
    }
    elseif ($cs.PartOfDomain) {
        try { Resolve-DnsName -Name $cs.Domain -DnsOnly -QuickTimeout -ErrorAction Stop | Out-Null; Write-Log ('AD domain {0} resolves.' -f $cs.Domain) 'OK' }
        catch { Write-Log ('AD domain {0} does not resolve (normal if off-network; check DHCP DNS if on-site).' -f $cs.Domain) 'WARN' }
    }
    $expectedDcs = @($memberFallback)
    if ((-not $script:ServerMode) -and $expectedDcs.Count -gt 0) {
        $upIdx = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' } | ForEach-Object { $_.ifIndex })
        $active = @(Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $upIdx -contains $_.InterfaceIndex } | ForEach-Object { $_.ServerAddresses })
        if (@($active | Where-Object { $expectedDcs -contains $_ }).Count -gt 0) { Write-Log ('Active DNS includes a DC: {0}' -f ($expectedDcs -join ',')) 'OK' }
        else { Write-Log ('Active DNS [{0}] does not include a DC [{1}]. Normal if off-network; check DHCP scope DNS if on-site.' -f ($active -join ','), ($expectedDcs -join ',')) 'WARN' }
    }

    # ---- Uninstall -------------------------------------------------------------------
    $script:CanRollback = $false   # past the gate; uninstall may remove what rollback needs
    $uninstallOk = $true
    $rebootNeeded = $false
    if ($products.Count -gt 0) {
        Write-Log '=== Phase: uninstall agent ==='
        foreach ($p in $products) {
            $guid = $null
            if ($p.KeyName -match '^\{[0-9A-Fa-f-]{36}\}$') { $guid = $p.KeyName }
            elseif ($p.UninstallString -match '\{[0-9A-Fa-f-]{36}\}') { $guid = $Matches[0] }
            if (-not $guid) {
                Write-Log ('No MSI product code for "{0}". UninstallString: {1}' -f $p.DisplayName, $p.UninstallString) 'ERROR'
                $uninstallOk = $false
                continue
            }
            $msiLog = Join-Path $LogDirectory ('msi-uninstall-{0}-{1}.log' -f $guid.Trim('{}'), $stamp)
            Write-Log ('Uninstalling {0} {1} ({2})' -f $p.DisplayName, $p.DisplayVersion, $guid)
            $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList ('/x {0} /qn /norestart /l*v "{1}"' -f $guid, $msiLog) -PassThru -WindowStyle Hidden
            $null = $proc.Handle
            if (-not $proc.WaitForExit($uninstallTimeout * 60000)) {
                Write-Log ('msiexec timed out after {0} min; terminating.' -f $uninstallTimeout) 'ERROR'
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                $uninstallOk = $false
                continue
            }
            switch ($proc.ExitCode) {
                0       { Write-Log 'msiexec exit 0 (success)' 'OK' }
                1605    { Write-Log 'msiexec exit 1605 (product not installed)' 'OK' }
                3010    { Write-Log 'msiexec exit 3010 (success, reboot required)' 'OK'; $rebootNeeded = $true }
                1641    { Write-Log 'msiexec exit 1641 (success, reboot initiated)' 'OK'; $rebootNeeded = $true }
                1618    { Write-Log 'msiexec exit 1618 (another install in progress)' 'ERROR'; $uninstallOk = $false }
                default { Write-Log ('msiexec exit {0}. See {1}' -f $proc.ExitCode, $msiLog) 'ERROR'; $uninstallOk = $false }
            }
        }
        if (@(Get-AgentProducts).Count -gt 0) {
            Write-Log 'Agent product still registered after uninstall.' 'ERROR'
            $uninstallOk = $false
        }
    }
    else {
        Write-Log 'No product registered (stranded DNS cleanup only). Skipping uninstall.'
    }

    # ---- Post-uninstall DNS re-check (uninstaller may restore its own DNS backup) ------
    Write-Log '=== Phase: post-uninstall DNS re-check ==='
    Start-Sleep -Seconds 5
    $plan3 = @(Get-DnsPlan -State @(Get-DnsInterfaceState) -Scope $ResetScope -Fallback $memberFallback -ServerMode $script:ServerMode -SelfMap $selfMap)
    $reapply = @($plan3 | Where-Object { $_.Changed })
    if ($reapply.Count -gt 0) {
        Write-Log 'DNS changed during uninstall. Re-applying reset.' 'WARN'
        Write-Plan -Plan $plan3 -Title 'DNS plan (post-uninstall)'
        if (@($plan3 | Where-Object { $_.Blocked }).Count -gt 0) {
            Exit-Script $EXIT_POST_DNS_FAIL 'Uninstall restored loopback DNS on a static-IP adapter and no fallback is set. Fix DNS manually.'
        }
        $null = Invoke-DnsPlan -Plan $plan3
    }
    $postValid = Test-DnsHealth -Hosts $script:Hosts -MinSuccess $script:MinSuccess -Retries $retries -DelaySec $delay -RequireNoLoopback $true -Label 'Post-uninstall' -AllowLoopback $script:AllowLoop
    if (-not $postValid) {
        Exit-Script $EXIT_POST_DNS_FAIL ('DNS validation failed after uninstall. Snapshot/log in {0}.' -f $LogDirectory)
    }

    if (-not $uninstallOk) {
        Exit-Script $EXIT_UNINSTALL_FAIL 'Uninstall failed. DNS is on DHCP and working; agent services left disabled.'
    }

    # ---- Residue cleanup -------------------------------------------------------------
    if ($doCleanup) {
        Write-Log '=== Phase: residue cleanup ==='
        foreach ($folder in $ResidueFolders) {
            if (-not (Test-Path $folder)) { continue }
            Get-Process -ErrorAction SilentlyContinue |
                Where-Object { $_.Path -and $_.Path.StartsWith($folder, [StringComparison]::OrdinalIgnoreCase) } |
                ForEach-Object { Write-Log ('Stopping leftover process {0} ({1})' -f $_.Name, $_.Id) 'WARN'; Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
        }
        Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like $ResidueProcessPattern } |
            ForEach-Object { Write-Log ('Stopping leftover process {0} ({1})' -f $_.Name, $_.Id) 'WARN'; Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
        foreach ($t in @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like $ResidueTaskPattern })) {
            try { Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction Stop; Write-Log ('Removed scheduled task {0}{1}' -f $t.TaskPath, $t.TaskName) 'OK' }
            catch { Write-Log ('Could not remove scheduled task {0}: {1}' -f $t.TaskName, $_.Exception.Message) 'WARN' }
        }
        foreach ($s in @(Get-AgentServiceState)) {
            Stop-Service -Name $s.Name -Force -ErrorAction SilentlyContinue
            & sc.exe delete "$($s.Name)" | Out-Null
            if ($LASTEXITCODE -eq 0) { Write-Log ('Deleted leftover service "{0}"' -f $s.Name) 'OK' }
            else { Write-Log ('Could not delete service "{0}" (sc exit {1}); may clear after reboot' -f $s.Name, $LASTEXITCODE) 'WARN' }
        }
        foreach ($folder in $ResidueFolders) {
            if (Test-Path $folder) {
                try { Remove-Item -Path $folder -Recurse -Force -ErrorAction Stop; Write-Log ('Removed folder {0}' -f $folder) 'OK' }
                catch { Write-Log ('Could not remove {0}: {1}' -f $folder, $_.Exception.Message) 'WARN' }
            }
        }
        foreach ($key in $ResidueRegistryKeys) {
            if (Test-Path $key) {
                try { Remove-Item -Path $key -Recurse -Force -ErrorAction Stop; Write-Log ('Removed registry key {0}' -f $key) 'OK' }
                catch { Write-Log ('Could not remove {0}: {1}' -f $key, $_.Exception.Message) 'WARN' }
            }
        }
    }

    # ---- Final removal verification ----------------------------------------------------
    Write-Log '=== Phase: verify removal ==='
    $leftovers = @()
    foreach ($s in @(Get-AgentServiceState)) { $leftovers += ('service:{0}' -f $s.Name) }
    foreach ($pr in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like $ResidueProcessPattern })) { $leftovers += ('process:{0}' -f $pr.ProcessName) }
    foreach ($f in $ResidueFolders) { if (Test-Path $f) { $leftovers += ('folder:{0}' -f $f) } }
    foreach ($k in $ResidueRegistryKeys) { if (Test-Path $k) { $leftovers += ('registry:{0}' -f $k) } }
    foreach ($t in @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like $ResidueTaskPattern })) { $leftovers += ('task:{0}' -f $t.TaskName) }
    if (@(Get-AgentProducts).Count -gt 0) { $leftovers += 'product:still registered' }
    if ($leftovers.Count -eq 0) { Write-Log 'PASS: no services, processes, folders, registry keys, or tasks remain.' 'OK' }
    else { foreach ($l in $leftovers) { Write-Log ('Leftover: {0}' -f $l) 'WARN' } }

    if ($rebootNeeded) { Exit-Script $EXIT_OK_REBOOT 'DNSFilter removed, DNS on DHCP and validated. Reboot required to finish.' }
    if ($leftovers.Count -gt 0) { Exit-Script $EXIT_OK_RESIDUE ('DNSFilter uninstalled and DNS validated, but residue remains: {0}' -f ($leftovers -join '; ')) }
    Exit-Script $EXIT_OK 'DNSFilter removed, DNS on DHCP and validated.'
}
catch {
    Write-Log ('Unhandled error: {0} at line {1}' -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber) 'ERROR'
    if ($script:CanRollback) {
        if (Invoke-Rollback -Snapshot $script:RbSnapshot -Services $script:RbServices -BaselineOk $script:RbBaseline) {
            Exit-Script $EXIT_ROLLED_BACK 'Unhandled error before uninstall. Rolled back; agent still installed.'
        }
        Exit-Script $EXIT_ROLLBACK_FAILED 'Unhandled error AND rollback failed. Check DNS now.'
    }
    Exit-Script $EXIT_FATAL 'Unhandled error. See log; DNS snapshot saved for manual rollback.'
}
