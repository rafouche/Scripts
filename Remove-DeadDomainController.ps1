#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Forcefully removes a dead/offline Domain Controller from Active Directory and cleans up associated DNS and AD metadata.

.DESCRIPTION
    This script performs a metadata cleanup of a Domain Controller that is no longer online
    and cannot be demoted gracefully. It handles:
      - AD metadata cleanup (ntdsutil-style via PowerShell)
      - Removal of the DC computer account from AD
      - Removal of associated DNS records (A, AAAA, SRV, NS, CNAME)
      - Removal of AD site/replication link objects
      - Removal of FRS/DFSR objects if present
      - Detection and forceful seizure of any FSMO roles held by the dead DC

    Designed to be run directly on the new/primary Domain Controller that will
    assume FSMO roles. Must be run as a Domain Admin (or higher).

.PARAMETER DCName
    The hostname (NetBIOS name) of the dead Domain Controller to remove.
    Example: DC02

.PARAMETER DomainFQDN
    The fully qualified domain name of the domain.
    Example: corp.altecusa.com

.PARAMETER SiteName
    (Optional) The AD site name the dead DC belongs to. If not specified, the script
    will attempt to discover it automatically.

.PARAMETER DCIPAddress
    (Optional) The IP address of the dead DC. Used to assist in DNS cleanup.
    If not provided, DNS cleanup will rely on name-based lookups only.

.PARAMETER SkipDNSCleanup
    Switch to skip DNS record removal (e.g., if DNS is hosted externally).

.PARAMETER SkipFSMOSeize
    Switch to skip automatic FSMO role seizure. Use this if roles were already
    transferred, or if you want to handle seizure manually after the script runs.

.PARAMETER WhatIf
    Shows what would be done without making any changes.

.EXAMPLE
    .\Remove-DeadDomainController.ps1 -DCName "DC02" -DomainFQDN "corp.altecusa.com"

.EXAMPLE
    .\Remove-DeadDomainController.ps1 -DCName "DC02" -DomainFQDN "corp.altecusa.com" -DCIPAddress "10.1.1.20" -WhatIf

.EXAMPLE
    .\Remove-DeadDomainController.ps1 -DCName "DC02" -DomainFQDN "corp.altecusa.com" -SkipFSMOSeize

.NOTES
    Author      : Altec Solutions Group, Inc.
    Version     : 1.1
    Requires    : PowerShell 5.1+, ActiveDirectory module, DNS Server module (for DNS cleanup)
                  Run directly on the new primary/target DC as Domain Admin or higher.

    WARNING: This operation is NOT reversible. Verify the DC is truly offline and
             will NOT be brought back online before running this script.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [Parameter(Mandatory = $true)]
    [string]$DCName,

    [Parameter(Mandatory = $true)]
    [string]$DomainFQDN,

    [Parameter(Mandatory = $false)]
    [string]$SiteName,

    [Parameter(Mandatory = $false)]
    [string]$DCIPAddress,

    [Parameter(Mandatory = $false)]
    [switch]$SkipDNSCleanup,

    [Parameter(Mandatory = $false)]
    [switch]$SkipFSMOSeize
)

#region --- Initialization ---

$ErrorActionPreference = "Stop"
$ScriptVersion         = "1.1"
$Timestamp             = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$LogFile               = "$PSScriptRoot\Remove-DC_${DCName}_${Timestamp}.log"
$DCNameUpper           = $DCName.ToUpper()

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS","SECTION")]
        [string]$Level = "INFO"
    )
    $Entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $Entry
    switch ($Level) {
        "INFO"    { Write-Host $Entry -ForegroundColor Cyan }
        "WARN"    { Write-Host $Entry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $Entry -ForegroundColor Red }
        "SUCCESS" { Write-Host $Entry -ForegroundColor Green }
        "SECTION" { Write-Host "`n$Entry`n$('='*80)" -ForegroundColor Magenta }
    }
}

function Confirm-Action {
    param([string]$ActionDescription)
    if ($WhatIfPreference) {
        Write-Log "WHATIF: $ActionDescription" -Level "WARN"
        return $false
    }
    return $true
}

#endregion

#region --- Pre-flight Checks ---

Write-Log "====== Remove-DeadDomainController v$ScriptVersion ======" -Level "SECTION"
Write-Log "Target DC    : $DCNameUpper"
Write-Log "Domain       : $DomainFQDN"
Write-Log "Initiated by : $env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME"

# Verify AD module
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Log "ActiveDirectory PowerShell module not found. Install RSAT or run from a DC." -Level "ERROR"
    exit 1
}
Import-Module ActiveDirectory -ErrorAction Stop
Write-Log "ActiveDirectory module loaded." -Level "INFO"

# Confirm the target DC is NOT reachable (safety check)
Write-Log "Verifying target DC is offline..." -Level "INFO"
$IsReachable = Test-Connection -ComputerName $DCNameUpper -Count 2 -Quiet -ErrorAction SilentlyContinue
if ($IsReachable) {
    Write-Log "WARNING: $DCNameUpper appears to be ONLINE (ping responded). Metadata cleanup should only be performed on truly offline DCs." -Level "WARN"
    $Confirm = Read-Host "The DC appears online. Are you SURE you want to force-remove it? Type YES to continue"
    if ($Confirm -ne "YES") {
        Write-Log "Aborted by user." -Level "WARN"
        exit 0
    }
} else {
    Write-Log "$DCNameUpper is not reachable (offline). Proceeding with metadata cleanup." -Level "SUCCESS"
}

# Resolve domain DN
try {
    $DomainDN = (Get-ADDomain -Identity $DomainFQDN).DistinguishedName
    Write-Log "Domain DN: $DomainDN" -Level "INFO"
} catch {
    Write-Log "Failed to resolve domain '$DomainFQDN'. Ensure you are connected to the domain and have appropriate permissions. Error: $_" -Level "ERROR"
    exit 1
}

# Locate the DC object in AD
try {
    $DCObject = Get-ADDomainController -Filter { Name -eq $DCNameUpper } -Server $DomainFQDN -ErrorAction SilentlyContinue
    if (-not $DCObject) {
        # Try finding via computer account
        $DCComputer = Get-ADComputer -Filter { Name -eq $DCNameUpper } -Server $DomainFQDN -Properties * -ErrorAction SilentlyContinue
        if (-not $DCComputer) {
            Write-Log "Could not locate '$DCNameUpper' as a DC or computer object in AD. It may have already been removed." -Level "WARN"
        } else {
            Write-Log "Found computer account for $DCNameUpper (may already be partially removed as DC)." -Level "INFO"
        }
    } else {
        Write-Log "Located DC object: $($DCObject.DistinguishedName)" -Level "INFO"
        if (-not $SiteName) {
            $SiteName = $DCObject.Site
            Write-Log "Auto-detected AD Site: $SiteName" -Level "INFO"
        }
    }
} catch {
    Write-Log "Error querying AD for DC object: $_" -Level "WARN"
}

#endregion

#region --- Step 1: AD Metadata Cleanup (Remove-ADDomainController) ---

Write-Log "STEP 1: AD Metadata Cleanup" -Level "SECTION"

try {
    $DCObject = Get-ADDomainController -Filter { Name -eq $DCNameUpper } -Server $DomainFQDN -ErrorAction SilentlyContinue

    if ($DCObject) {
        if ($PSCmdlet.ShouldProcess($DCNameUpper, "Remove-ADDomainController (Force metadata cleanup)")) {
            Write-Log "Running Remove-ADDomainController -ForceRemoval..." -Level "INFO"
            Remove-ADDomainController -Identity $DCObject -ForceRemoval -Confirm:$false
            Write-Log "DC metadata removed successfully from Active Directory." -Level "SUCCESS"
        }
    } else {
        Write-Log "DC object not found via Get-ADDomainController - may already be removed. Skipping Remove-ADDomainController." -Level "WARN"
    }
} catch {
    Write-Log "Remove-ADDomainController failed: $_" -Level "ERROR"
    Write-Log "Attempting fallback via ntdsutil..." -Level "WARN"

    # Fallback: ntdsutil metadata cleanup
    $NtdsutilScript = @"
activate instance ntds
metadata cleanup
connections
connect to server $env:COMPUTERNAME
quit
select operation target
list sites
select site 0
list servers in site
quit
quit
quit
"@
    Write-Log 'ntdsutil fallback requires manual interaction. Please run ntdsutil manually:' -Level "WARN"
    Write-Log '  ntdsutil -> metadata cleanup -> connections -> connect to server [LiveDC]' -Level "WARN"
    Write-Log '  -> select operation target -> select the site/server -> remove selected server' -Level "WARN"
}

#endregion

#region --- Step 2: Remove Computer Account from AD ---

Write-Log "STEP 2: Remove Computer Account from AD" -Level "SECTION"

try {
    $DCComputer = Get-ADComputer -Filter { Name -eq $DCNameUpper } -Server $DomainFQDN -ErrorAction SilentlyContinue
    if ($DCComputer) {
        if ($PSCmdlet.ShouldProcess($DCComputer.DistinguishedName, "Remove-ADComputer")) {
            Remove-ADComputer -Identity $DCComputer -Confirm:$false
            Write-Log "Computer account '$DCNameUpper' removed from AD." -Level "SUCCESS"
        }
    } else {
        Write-Log "Computer account '$DCNameUpper' not found (may already be removed)." -Level "INFO"
    }
} catch {
    Write-Log "Failed to remove computer account: $_" -Level "ERROR"
}

#endregion

#region --- Step 3: Remove Replication/NTDS Settings Objects ---

Write-Log "STEP 3: Remove Orphaned Replication Objects" -Level "SECTION"

try {
    # Search for lingering NTDS Settings objects
    $NTDSPath = "CN=NTDS Settings,CN=$DCNameUpper,CN=Servers"
    $SearchBase = "CN=Sites,CN=Configuration,$DomainDN"
    $NTDSObjects = Get-ADObject -Filter { Name -eq "NTDS Settings" } -SearchBase $SearchBase `
        -SearchScope Subtree -Server $DomainFQDN -ErrorAction SilentlyContinue |
        Where-Object { $_.DistinguishedName -like "*CN=$DCNameUpper,*" }

    if ($NTDSObjects) {
        foreach ($obj in $NTDSObjects) {
            if ($PSCmdlet.ShouldProcess($obj.DistinguishedName, "Remove orphaned NTDS Settings object")) {
                Remove-ADObject -Identity $obj -Recursive -Confirm:$false
                Write-Log "Removed NTDS Settings object: $($obj.DistinguishedName)" -Level "SUCCESS"
            }
        }
    } else {
        Write-Log "No orphaned NTDS Settings objects found." -Level "INFO"
    }

    # Remove the server object from Sites and Services
    $ServerObjects = Get-ADObject -Filter { Name -eq $DCNameUpper } -SearchBase $SearchBase `
        -SearchScope Subtree -Server $DomainFQDN -ErrorAction SilentlyContinue |
        Where-Object { $_.ObjectClass -eq "server" }

    if ($ServerObjects) {
        foreach ($obj in $ServerObjects) {
            if ($PSCmdlet.ShouldProcess($obj.DistinguishedName, "Remove server object from Sites and Services")) {
                Remove-ADObject -Identity $obj -Recursive -Confirm:$false
                Write-Log "Removed server object from Sites and Services: $($obj.DistinguishedName)" -Level "SUCCESS"
            }
        }
    } else {
        Write-Log "No lingering server objects found in Sites and Services." -Level "INFO"
    }
} catch {
    Write-Log "Error during replication object cleanup: $_" -Level "ERROR"
}

#endregion

#region --- Step 4: Remove FRS / DFSR Subscriber Objects ---

Write-Log "STEP 4: Remove FRS / DFSR Subscriber Objects" -Level "SECTION"

try {
    # DFSR subscriber objects
    $DFSRBase = "CN=DFSR-LocalSettings,CN=$DCNameUpper,OU=Domain Controllers,$DomainDN"
    $DFSRObject = Get-ADObject -Identity $DFSRBase -Server $DomainFQDN -ErrorAction SilentlyContinue
    if ($DFSRObject) {
        if ($PSCmdlet.ShouldProcess($DFSRBase, "Remove DFSR subscriber object")) {
            Remove-ADObject -Identity $DFSRObject -Recursive -Confirm:$false
            Write-Log "Removed DFSR subscriber object." -Level "SUCCESS"
        }
    } else {
        Write-Log "No DFSR subscriber objects found for $DCNameUpper." -Level "INFO"
    }

    # FRS subscriber objects
    $FRSBase = "CN=Domain System Volume (SYSVOL share),CN=NTFRS Subscriptions,CN=$DCNameUpper,OU=Domain Controllers,$DomainDN"
    $FRSObject = Get-ADObject -Identity $FRSBase -Server $DomainFQDN -ErrorAction SilentlyContinue
    if ($FRSObject) {
        if ($PSCmdlet.ShouldProcess($FRSBase, "Remove FRS subscriber object")) {
            Remove-ADObject -Identity $FRSObject -Recursive -Confirm:$false
            Write-Log "Removed FRS subscriber object." -Level "SUCCESS"
        }
    } else {
        Write-Log "No FRS subscriber objects found for $DCNameUpper." -Level "INFO"
    }
} catch {
    Write-Log "Error during FRS/DFSR cleanup: $_" -Level "WARN"
}

#endregion

#region --- Step 5: DNS Cleanup ---

Write-Log "STEP 5: DNS Cleanup" -Level "SECTION"

if ($SkipDNSCleanup) {
    Write-Log "DNS cleanup skipped (-SkipDNSCleanup specified)." -Level "WARN"
} else {
    # Check for DNS Server module
    if (-not (Get-Module -ListAvailable -Name DnsServer)) {
        Write-Log "DnsServer PowerShell module not found. DNS cleanup will be skipped. Install RSAT DNS Tools to enable this." -Level "WARN"
    } else {
        Import-Module DnsServer -ErrorAction SilentlyContinue

        # Determine which DNS server to query/modify (use current DC or PDCe)
        try {
            $PDCe = (Get-ADDomain -Identity $DomainFQDN).PDCEmulator
            $DNSServer = $PDCe
            Write-Log "Using PDC Emulator for DNS operations: $PDCe" -Level "INFO"
        } catch {
            $DNSServer = $env:COMPUTERNAME
            Write-Log "Could not determine PDCe - using local machine for DNS: $DNSServer" -Level "WARN"
        }

        # --- Remove A/AAAA records ---
        Write-Log "Removing A/AAAA host records for $DCNameUpper in zone '$DomainFQDN'..." -Level "INFO"
        try {
            $ARecords = Get-DnsServerResourceRecord -ZoneName $DomainFQDN -ComputerName $DNSServer `
                -RRType A -ErrorAction SilentlyContinue |
                Where-Object { $_.HostName -eq $DCNameUpper }

            foreach ($rec in $ARecords) {
                if ($PSCmdlet.ShouldProcess("$($rec.HostName) -> $($rec.RecordData.IPv4Address)", "Remove DNS A record")) {
                    Remove-DnsServerResourceRecord -ZoneName $DomainFQDN -ComputerName $DNSServer `
                        -InputObject $rec -Force
                    Write-Log "Removed A record: $($rec.HostName) -> $($rec.RecordData.IPv4Address)" -Level "SUCCESS"
                }
            }

            if ($DCIPAddress) {
                # Also remove by IP in case hostname differs
                $ARecordsByIP = Get-DnsServerResourceRecord -ZoneName $DomainFQDN -ComputerName $DNSServer `
                    -RRType A -ErrorAction SilentlyContinue |
                    Where-Object { $_.RecordData.IPv4Address -eq $DCIPAddress }
                foreach ($rec in $ARecordsByIP) {
                    if ($PSCmdlet.ShouldProcess("$($rec.HostName) -> $($rec.RecordData.IPv4Address)", "Remove DNS A record by IP")) {
                        Remove-DnsServerResourceRecord -ZoneName $DomainFQDN -ComputerName $DNSServer `
                            -InputObject $rec -Force
                        Write-Log "Removed A record by IP: $($rec.HostName) -> $($rec.RecordData.IPv4Address)" -Level "SUCCESS"
                    }
                }
            }
        } catch {
            Write-Log "Error removing A records: $_" -Level "WARN"
        }

        # --- Remove SRV records referencing the dead DC ---
        Write-Log "Scanning for SRV records referencing $DCNameUpper..." -Level "INFO"
        $SRVZones = @(
            "_msdcs.$DomainFQDN",
            $DomainFQDN
        )
        foreach ($zone in $SRVZones) {
            try {
                $SRVRecords = Get-DnsServerResourceRecord -ZoneName $zone -ComputerName $DNSServer `
                    -RRType SRV -ErrorAction SilentlyContinue |
                    Where-Object { $_.RecordData.NameTarget -like "$DCNameUpper*" }

                foreach ($rec in $SRVRecords) {
                    if ($PSCmdlet.ShouldProcess("SRV: $($rec.HostName) -> $($rec.RecordData.NameTarget)", "Remove DNS SRV record")) {
                        Remove-DnsServerResourceRecord -ZoneName $zone -ComputerName $DNSServer `
                            -InputObject $rec -Force
                        Write-Log "Removed SRV record [$zone]: $($rec.HostName) -> $($rec.RecordData.NameTarget)" -Level "SUCCESS"
                    }
                }
            } catch {
                Write-Log "Error scanning SRV records in zone '$zone': $_" -Level "WARN"
            }
        }

        # --- Remove NS records referencing the dead DC ---
        Write-Log "Scanning for NS records referencing $DCNameUpper..." -Level "INFO"
        try {
            $NSRecords = Get-DnsServerResourceRecord -ZoneName $DomainFQDN -ComputerName $DNSServer `
                -RRType NS -ErrorAction SilentlyContinue |
                Where-Object { $_.RecordData.NameServer -like "$DCNameUpper*" }

            foreach ($rec in $NSRecords) {
                if ($PSCmdlet.ShouldProcess("NS: $($rec.RecordData.NameServer)", "Remove DNS NS record")) {
                    Remove-DnsServerResourceRecord -ZoneName $DomainFQDN -ComputerName $DNSServer `
                        -InputObject $rec -Force
                    Write-Log "Removed NS record: $($rec.RecordData.NameServer)" -Level "SUCCESS"
                }
            }
        } catch {
            Write-Log "Error scanning NS records: $_" -Level "WARN"
        }

        # --- Remove CNAME records referencing the dead DC ---
        Write-Log "Scanning for CNAME records referencing $DCNameUpper..." -Level "INFO"
        try {
            $CNAMERecords = Get-DnsServerResourceRecord -ZoneName $DomainFQDN -ComputerName $DNSServer `
                -RRType CNAME -ErrorAction SilentlyContinue |
                Where-Object { $_.RecordData.HostNameAlias -like "$DCNameUpper*" }

            foreach ($rec in $CNAMERecords) {
                if ($PSCmdlet.ShouldProcess("CNAME: $($rec.HostName) -> $($rec.RecordData.HostNameAlias)", "Remove DNS CNAME record")) {
                    Remove-DnsServerResourceRecord -ZoneName $DomainFQDN -ComputerName $DNSServer `
                        -InputObject $rec -Force
                    Write-Log "Removed CNAME record: $($rec.HostName) -> $($rec.RecordData.HostNameAlias)" -Level "SUCCESS"
                }
            }
        } catch {
            Write-Log "Error scanning CNAME records: $_" -Level "WARN"
        }

        # --- Cleanup _msdcs subdelegation ---
        Write-Log "Checking for DC GUID record in _msdcs zone..." -Level "INFO"
        try {
            $MsdcsZone = "_msdcs.$DomainFQDN"
            $MsdcsARecords = Get-DnsServerResourceRecord -ZoneName $MsdcsZone -ComputerName $DNSServer `
                -RRType A -ErrorAction SilentlyContinue |
                Where-Object { $_.RecordData.IPv4Address -eq $DCIPAddress }
            foreach ($rec in $MsdcsARecords) {
                if ($PSCmdlet.ShouldProcess("_msdcs A: $($rec.HostName)", "Remove _msdcs A record")) {
                    Remove-DnsServerResourceRecord -ZoneName $MsdcsZone -ComputerName $DNSServer `
                        -InputObject $rec -Force
                    Write-Log "Removed _msdcs A record: $($rec.HostName)" -Level "SUCCESS"
                }
            }
        } catch {
            Write-Log "_msdcs zone cleanup skipped or failed: $_" -Level "WARN"
        }
    }
}

#endregion

#region --- Step 6: Verify Replication Health on Remaining DCs ---

Write-Log "STEP 6: Post-Cleanup Replication Check" -Level "SECTION"

try {
    Write-Log "Running repadmin /showrepl to verify remaining DC replication health..." -Level "INFO"
    if (-not $WhatIfPreference) {
        $ReplOutput = & repadmin /showrepl 2>&1
        $ReplOutput | ForEach-Object { Write-Log $_ -Level "INFO" }
        Write-Log "Replication check complete. Review output above for any errors." -Level "INFO"
    } else {
        Write-Log "WHATIF: Would run repadmin /showrepl" -Level "WARN"
    }
} catch {
    Write-Log "repadmin not available or failed: $_" -Level "WARN"
}

#endregion

#region --- Step 7: FSMO Role Detection and Seizure ---

Write-Log "STEP 7: FSMO Role Detection and Seizure" -Level "SECTION"

if ($SkipFSMOSeize) {
    Write-Log "FSMO seizure skipped (-SkipFSMOSeize specified). Run 'netdom query fsmo' manually to check role holders." -Level "WARN"
} else {
    # This server is the seizure target (script is meant to run on the new primary DC)
    $SeizureTarget = $env:COMPUTERNAME

    Write-Log "Checking current FSMO role holders..." -Level "INFO"
    Write-Log "Roles will be seized to this server: $SeizureTarget" -Level "INFO"

    try {
        $Forest     = Get-ADForest  -Server $DomainFQDN
        $Domain     = Get-ADDomain  -Server $DomainFQDN

        # Map each role to its current holder (just the hostname portion)
        $RoleMap = [ordered]@{
            "SchemaMaster"          = ($Forest.SchemaMaster       -split '\.')[0].ToUpper()
            "DomainNamingMaster"    = ($Forest.DomainNamingMaster -split '\.')[0].ToUpper()
            "PDCEmulator"           = ($Domain.PDCEmulator        -split '\.')[0].ToUpper()
            "RIDMaster"             = ($Domain.RIDMaster          -split '\.')[0].ToUpper()
            "InfrastructureMaster"  = ($Domain.InfrastructureMaster -split '\.')[0].ToUpper()
        }

        Write-Log "Current FSMO role holders:" -Level "INFO"
        foreach ($role in $RoleMap.Keys) {
            Write-Log ("  {0,-25} -> {1}" -f $role, $RoleMap[$role]) -Level "INFO"
        }

        # Determine which roles are held by the dead DC
        $RolesToSeize = $RoleMap.Keys | Where-Object { $RoleMap[$_] -eq $DCNameUpper }

        if (-not $RolesToSeize) {
            Write-Log "No FSMO roles are held by '$DCNameUpper'. No seizure required." -Level "SUCCESS"
        } else {
            Write-Log "The following FSMO roles are held by the dead DC and need to be seized:" -Level "WARN"
            foreach ($role in $RolesToSeize) {
                Write-Log "  -> $role" -Level "WARN"
            }

            # --- Seize Domain-level roles (PDCEmulator, RIDMaster, InfrastructureMaster) ---
            $DomainRoles      = @("PDCEmulator", "RIDMaster", "InfrastructureMaster")
            $DomainRoleTarget = $DomainRoles | Where-Object { $RolesToSeize -contains $_ }

            if ($DomainRoleTarget) {
                Write-Log "Seizing domain-level FSMO roles on '$SeizureTarget'..." -Level "INFO"
                if ($PSCmdlet.ShouldProcess($SeizureTarget, "Seize domain FSMO roles: $($DomainRoleTarget -join ', ')")) {
                    try {
                        # Move-ADDirectoryServerOperationMasterRole with -Force seizes without needing the old DC
                        Move-ADDirectoryServerOperationMasterRole `
                            -Identity $SeizureTarget `
                            -OperationMasterRole $DomainRoleTarget `
                            -Force `
                            -Confirm:$false
                        foreach ($role in $DomainRoleTarget) {
                            Write-Log "Seized role '$role' -> $SeizureTarget" -Level "SUCCESS"
                        }
                    } catch {
                        Write-Log "Failed to seize domain FSMO roles via PowerShell: $_" -Level "ERROR"
                        Write-Log "Attempting fallback via ntdsutil for domain roles..." -Level "WARN"
                        # ntdsutil seizure fallback (outputs instructions)
                        $NtdsLines = @(
                            "ntdsutil",
                            "  roles",
                            "  connections",
                            "    connect to server $SeizureTarget",
                            "    quit",
                            "  seize pdc",
                            "  seize rid master",
                            "  seize infrastructure master",
                            "  quit",
                            "quit"
                        )
                        Write-Log "Manual ntdsutil seizure commands for domain roles:" -Level "WARN"
                        $NtdsLines | ForEach-Object { Write-Log "  $_" -Level "WARN" }
                    }
                }
            }

            # --- Seize Forest-level roles (SchemaMaster, DomainNamingMaster) ---
            $ForestRoles      = @("SchemaMaster", "DomainNamingMaster")
            $ForestRoleTarget = $ForestRoles | Where-Object { $RolesToSeize -contains $_ }

            if ($ForestRoleTarget) {
                Write-Log "Seizing forest-level FSMO roles on '$SeizureTarget'..." -Level "INFO"
                Write-Log "NOTE: Seizing SchemaMaster or DomainNamingMaster requires Schema Admin / Enterprise Admin membership." -Level "WARN"
                if ($PSCmdlet.ShouldProcess($SeizureTarget, "Seize forest FSMO roles: $($ForestRoleTarget -join ', ')")) {
                    try {
                        Move-ADDirectoryServerOperationMasterRole `
                            -Identity $SeizureTarget `
                            -OperationMasterRole $ForestRoleTarget `
                            -Force `
                            -Confirm:$false
                        foreach ($role in $ForestRoleTarget) {
                            Write-Log "Seized role '$role' -> $SeizureTarget" -Level "SUCCESS"
                        }
                    } catch {
                        Write-Log "Failed to seize forest FSMO roles via PowerShell: $_" -Level "ERROR"
                        Write-Log "Attempting fallback via ntdsutil for forest roles..." -Level "WARN"
                        $NtdsLines = @(
                            "ntdsutil",
                            "  roles",
                            "  connections",
                            "    connect to server $SeizureTarget",
                            "    quit",
                            "  seize schema master",
                            "  seize domain naming master",
                            "  quit",
                            "quit"
                        )
                        Write-Log "Manual ntdsutil seizure commands for forest roles:" -Level "WARN"
                        $NtdsLines | ForEach-Object { Write-Log "  $_" -Level "WARN" }
                    }
                }
            }

            # --- Verify roles after seizure ---
            Write-Log "Verifying FSMO roles after seizure..." -Level "INFO"
            try {
                $ForestPost = Get-ADForest  -Server $DomainFQDN
                $DomainPost = Get-ADDomain  -Server $DomainFQDN

                $PostRoleMap = [ordered]@{
                    "SchemaMaster"         = ($ForestPost.SchemaMaster          -split '\.')[0].ToUpper()
                    "DomainNamingMaster"   = ($ForestPost.DomainNamingMaster    -split '\.')[0].ToUpper()
                    "PDCEmulator"          = ($DomainPost.PDCEmulator           -split '\.')[0].ToUpper()
                    "RIDMaster"            = ($DomainPost.RIDMaster             -split '\.')[0].ToUpper()
                    "InfrastructureMaster" = ($DomainPost.InfrastructureMaster  -split '\.')[0].ToUpper()
                }

                Write-Log "FSMO role holders after seizure:" -Level "INFO"
                foreach ($role in $PostRoleMap.Keys) {
                    $holder = $PostRoleMap[$role]
                    $status = if ($holder -eq $DCNameUpper) { "ERROR" } else { "SUCCESS" }
                    Write-Log ("  {0,-25} -> {1}" -f $role, $holder) -Level $status
                }

                $StillOnDead = $PostRoleMap.Keys | Where-Object { $PostRoleMap[$_] -eq $DCNameUpper }
                if ($StillOnDead) {
                    Write-Log "WARNING: The following roles still show the dead DC as holder. Manual ntdsutil seizure may be required: $($StillOnDead -join ', ')" -Level "ERROR"
                } else {
                    Write-Log "All FSMO roles successfully moved off of '$DCNameUpper'." -Level "SUCCESS"
                }
            } catch {
                Write-Log "Could not re-query FSMO roles for verification: $_" -Level "WARN"
            }
        }
    } catch {
        Write-Log "Failed to query FSMO role holders: $_" -Level "ERROR"
        Write-Log "Run 'netdom query fsmo' manually and seize any roles held by '$DCNameUpper'." -Level "WARN"
    }
}

#endregion

#region --- Summary ---

Write-Log "====== Cleanup Complete ======" -Level "SECTION"
Write-Log "Dead DC '$DCNameUpper' has been processed for removal from domain '$DomainFQDN'." -Level "SUCCESS"
Write-Log "Log file saved to: $LogFile" -Level "INFO"
Write-Log "" -Level "INFO"
Write-Log "RECOMMENDED POST-CLEANUP STEPS:" -Level "INFO"
Write-Log "  1. Run 'repadmin /replsummary' on all remaining DCs to confirm clean replication." -Level "INFO"
Write-Log "  2. Run 'dcdiag /test:replications /s:$env:COMPUTERNAME' on this DC." -Level "INFO"
Write-Log "  3. Run 'netdom query fsmo' to confirm all 5 roles are on live DCs." -Level "INFO"
Write-Log "  4. Open 'Active Directory Sites and Services' and confirm the server object is gone." -Level "INFO"
Write-Log "  5. Open 'DNS Manager' and verify no stale records remain for $DCNameUpper." -Level "INFO"
Write-Log "  6. If $DCNameUpper was a Global Catalog server, verify another GC exists:" -Level "INFO"
Write-Log "     repadmin /options $env:COMPUTERNAME  -- look for IS_GC flag" -Level "INFO"

#endregion
