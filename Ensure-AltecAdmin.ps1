#Requires -Version 5.1
<#
.SYNOPSIS
  Ensure a local break-glass admin account exists, is enabled, has the
  supplied password, is a member of the local Administrators group, and
  (by default) is hidden from the Windows sign-in screen.

.DESCRIPTION
  Companion to "Run Autoelevate Removal" for hosts that uninstalled AE but
  failed to create the break-glass account.

  Account name defaults to "Altec". Override with -AccountName or the Ninja
  script variable AltecAdminUsername.

  Value resolution order (each value):
    1. CLI parameter (-AccountName / -Password / -HideFromLogon)
    2. Ninja script variable (environment variable)
    3. Default ("Altec" for the name, $true for HideFromLogon,
       the password has no default)

  Hide from sign-in screen:
    Writes a DWORD named after the account (value 0) to
      HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList
    The account still works. Sign in with "Other user" and type .\Altec
    HideFromLogon = false removes the value if present (unhides the account).

  Does NOT remove other Administrators members.
  Does NOT touch the built-in Administrator (RID 500). Refuses if the target
    name resolves to it.
  Does NOT run on domain controllers (no local SAM accounts there).
  Does NOT remove or rename legacy break-glass accounts (warns only).
  Does NOT touch AutoElevate.

  Ninja library setup:
    Name:         Ensure AltecAdmin Break-Glass
    Language:     PowerShell
    Architecture: 64-bit   (required; the LocalAccounts module does not load
                            in 32-bit PowerShell on 64-bit Windows)
    Run as:       SYSTEM
    Script variables:
      AltecAdminPassword       TEXT, required.
        Ninja rejects "!" in values. Type {EXCL} for each "!".
        The script converts {EXCL} to "!" on the endpoint.
      AltecAdminUsername       TEXT, optional. Blank = "Altec".
        1-20 chars. Not allowed: " / \ [ ] : ; | = , + * ? < > @
      AltecAdminHideFromLogon  CHECKBOX, optional. Default checked (true).
        Accepted values: true/false, 1/0, yes/no, on/off.

  Full name / description:
    Default name ("Altec")  -> created with the Altec full name and description.
    Any other name          -> created with blank full name and description.
    Existing accounts keep whatever full name and description they already have.

.NOTES
  Log: C:\ProgramData\Altec\Ensure-AltecAdmin.log

  Exit codes:
    0  = SUCCESS
    10 = Password missing
    11 = Invalid account name
    12 = Target name is the built-in Administrator (RID 500), refused
    13 = Host is a domain controller, refused
    14 = Running in 32-bit PowerShell on 64-bit Windows, refused
    20 = Create / password reset / enable failed
    30 = Add to Administrators failed
    40 = Verification failed
    50 = Hide/unhide from sign-in screen failed
#>
[CmdletBinding()]
param(
    [string]$AccountName,
    [string]$Password,
    [string]$HideFromLogon
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# ---------------- config ----------------
$DefaultAccountName    = 'Altec'
$DefaultHideFromLogon  = $true
$NinjaVarAccountName   = 'AltecAdminUsername'
$NinjaVarPassword      = 'AltecAdminPassword'
$NinjaVarHideFromLogon = 'AltecAdminHideFromLogon'
$LegacyAccountNames    = @('altecadmin', 'altec')   # warn-only if present alongside target
$AccountFullName       = 'Altec Break Glass'
$AccountDescription    = 'Altec MSP break-glass local administrator'
$AdminsSid             = 'S-1-5-32-544'             # locale-independent Administrators
$UF_DONT_EXPIRE_PASSWD = 0x10000
$UserListPath          = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList'

$script:Computer = $env:COMPUTERNAME
$script:LogDir   = 'C:\ProgramData\Altec'
$script:LogFile  = Join-Path $script:LogDir 'Ensure-AltecAdmin.log'
$script:Warnings = 0
$script:Errors   = 0
$script:ExitCode = 0
$script:Account  = ''
$script:UseDefaultLabels = $true
$script:HideState = 'NOT_CHECKED'
$script:Actions  = New-Object System.Collections.Generic.List[string]

# ---------------- helpers ----------------
function Write-Log {
    param(
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO',
        [Parameter(Mandatory)][string]$Message
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    try {
        if (-not (Test-Path -LiteralPath $script:LogDir)) {
            New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null
        }
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8
    } catch {}
    # Write-Host (not Write-Output) so log lines never leak into function return values
    Write-Host $line
    switch ($Level) {
        'WARN'  { $script:Warnings++ }
        'ERROR' { $script:Errors++ }
    }
}

function Set-Failed {
    param([int]$Code, [string]$Message)
    Write-Log ERROR $Message
    if ($script:ExitCode -eq 0) { $script:ExitCode = $Code }
}

function Resolve-Value {
    # CLI parameter -> Ninja env var -> default
    param([string]$CliValue, [string]$EnvName, [string]$Default)
    if (-not [string]::IsNullOrWhiteSpace($CliValue)) {
        return [pscustomobject]@{ Value = ($CliValue -replace '\{EXCL\}', '!').Trim(); Source = 'parameter' }
    }
    $envVal = [Environment]::GetEnvironmentVariable($EnvName)
    if (-not [string]::IsNullOrWhiteSpace($envVal)) {
        return [pscustomobject]@{ Value = ($envVal -replace '\{EXCL\}', '!').Trim(); Source = 'Ninja variable' }
    }
    return [pscustomobject]@{ Value = $Default; Source = 'default' }
}

function ConvertTo-BoolOrNull {
    # Returns $true / $false, or $null if the text is not a recognized boolean
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    switch -Regex ($Text.Trim()) {
        '^(1|true|yes|y|on)$'   { return $true }
        '^(0|false|no|n|off)$'  { return $false }
        default                 { return $null }
    }
}

function Test-AccountName {
    # Returns $null if valid, otherwise a reason string
    param([string]$Name)
    if ($Name.Length -lt 1 -or $Name.Length -gt 20) { return 'must be 1-20 characters' }
    if ($Name -match '["/\\\[\]:;|=,+*?<>@]')      { return 'contains an invalid character' }
    if ($Name -match '^[\s\.]+$')                   { return 'cannot be only spaces or periods' }
    if ($Name -ieq $env:COMPUTERNAME)               { return 'cannot match the computer name' }
    return $null
}

function Get-LocalAccount {
    param([string]$Name)
    try { return Get-LocalUser -Name $Name -ErrorAction Stop } catch { return $null }
}

function Get-AdminsGroup {
    # ADSI instead of Get-LocalGroupMember: the cmdlet throws on groups that
    # contain orphaned or Entra (AzureAD) SIDs, which is common on our fleet.
    $sid  = New-Object System.Security.Principal.SecurityIdentifier($AdminsSid)
    $name = $sid.Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
    return [ADSI]('WinNT://{0}/{1},group' -f $env:COMPUTERNAME, $name)
}

function Test-InAdmins {
    param([string]$UserSid)
    $grp = Get-AdminsGroup
    foreach ($m in @($grp.psbase.Invoke('Members'))) {
        try {
            $bytes = $m.GetType().InvokeMember('objectSid', 'GetProperty', $null, $m, $null)
            $s = (New-Object System.Security.Principal.SecurityIdentifier($bytes, 0)).Value
            if ($s -eq $UserSid) { return $true }
        } catch {}
    }
    return $false
}

function Add-ToAdmins {
    param([string]$Name)
    $grp = Get-AdminsGroup
    [void]$grp.psbase.Invoke('Add', ('WinNT://{0}/{1},user' -f $env:COMPUTERNAME, $Name))
}

function New-LocalAccountAdsi {
    # Fallback if New-LocalUser fails. Avoids "net user" so the password
    # never appears on a process command line.
    param([string]$Name, [string]$Plain)
    $c = [ADSI]('WinNT://{0},computer' -f $env:COMPUTERNAME)
    $u = $c.Create('User', $Name)
    $u.SetPassword($Plain)
    $u.SetInfo()
    if ($script:UseDefaultLabels) {
        $u.Put('FullName', $AccountFullName)
        $u.Put('Description', $AccountDescription)
    }
    $flags = [int]$u.InvokeGet('UserFlags')
    $u.Put('UserFlags', ($flags -bor $UF_DONT_EXPIRE_PASSWD))
    $u.SetInfo()
}

function Clear-Lockout {
    # Get-LocalUser has no LockedOut property and Unlock-LocalUser does not
    # exist, so lockout is handled through ADSI.
    param([string]$Name)
    $u = [ADSI]('WinNT://{0}/{1},user' -f $env:COMPUTERNAME, $Name)
    if ([bool]$u.InvokeGet('IsAccountLocked')) {
        $u.InvokeSet('IsAccountLocked', $false)
        $u.SetInfo()
        return $true
    }
    return $false
}

function Get-UserListEntry {
    # Returns @{ Exists; Value; Kind } for the account's UserList value
    param([string]$Name)
    $r = [pscustomobject]@{ Exists = $false; Value = $null; Kind = $null }
    if (-not (Test-Path -LiteralPath $UserListPath)) { return $r }
    $key = Get-Item -LiteralPath $UserListPath -ErrorAction Stop
    if ($key.GetValueNames() -icontains $Name) {
        $r.Exists = $true
        $r.Value  = $key.GetValue($Name)
        $r.Kind   = $key.GetValueKind($Name).ToString()
    }
    return $r
}

function Test-IsHidden {
    param([string]$Name)
    $e = Get-UserListEntry -Name $Name
    return ($e.Exists -and $e.Kind -eq 'DWord' -and [int]$e.Value -eq 0)
}

function Set-HiddenFromLogon {
    param([string]$Name)
    if (Test-IsHidden -Name $Name) { return $false }
    if (-not (Test-Path -LiteralPath $UserListPath)) {
        New-Item -Path $UserListPath -Force -ErrorAction Stop | Out-Null
    }
    # -Force overwrites a wrong-type value (e.g. REG_SZ "0") with a DWORD
    New-ItemProperty -LiteralPath $UserListPath -Name $Name -PropertyType DWord -Value 0 -Force -ErrorAction Stop | Out-Null
    return $true
}

function Clear-HiddenFromLogon {
    param([string]$Name)
    $e = Get-UserListEntry -Name $Name
    if (-not $e.Exists) { return $false }
    Remove-ItemProperty -LiteralPath $UserListPath -Name $Name -ErrorAction Stop
    return $true
}

# ---------------- main ----------------
Write-Log INFO ('===== Ensure-AltecAdmin starting on {0} =====' -f $script:Computer)

# Environment guards
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    Set-Failed 14 'Running in 32-bit PowerShell on 64-bit Windows. Set the Ninja script architecture to 64-bit.'
}
if ($script:ExitCode -eq 0) {
    try {
        $role = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).DomainRole
        if ($role -ge 4) {
            Set-Failed 13 'Host is a domain controller. Local accounts do not exist on DCs. Refusing.'
        }
    } catch {
        Write-Log WARN ('Could not determine domain role: {0}' -f $_.Exception.Message)
    }
}

$acctRes = Resolve-Value -CliValue $AccountName   -EnvName $NinjaVarAccountName   -Default $DefaultAccountName
$pwRes   = Resolve-Value -CliValue $Password      -EnvName $NinjaVarPassword      -Default ''
$hideRes = Resolve-Value -CliValue $HideFromLogon -EnvName $NinjaVarHideFromLogon -Default ([string]$DefaultHideFromLogon)
$script:Account = $acctRes.Value
$plain          = $pwRes.Value

$hideBool = ConvertTo-BoolOrNull -Text $hideRes.Value
if ($null -eq $hideBool) {
    Write-Log WARN ("HideFromLogon value '{0}' not recognized. Using default ({1})." -f $hideRes.Value, $DefaultHideFromLogon)
    $hideBool = $DefaultHideFromLogon
}

# Altec full name/description only apply to the default account name
$script:UseDefaultLabels = ($script:Account -ieq $DefaultAccountName)

Write-Log INFO ('Target account: {0} (source: {1}, Altec labels: {2})' -f $script:Account, $acctRes.Source, $script:UseDefaultLabels)
Write-Log INFO ('Hide from sign-in screen: {0} (source: {1})' -f $hideBool, $hideRes.Source)

if ($script:ExitCode -eq 0) {
    $nameProblem = Test-AccountName -Name $script:Account
    if ($nameProblem) {
        Set-Failed 11 ("Account name '{0}' {1}." -f $script:Account, $nameProblem)
    } elseif ([string]::IsNullOrWhiteSpace($plain)) {
        Set-Failed 10 'AltecAdminPassword is empty. Set the Ninja script variable (type {EXCL} for each !).'
    } else {
        Write-Log INFO ('Password resolved from {0} (length={1}).' -f $pwRes.Source, $plain.Length)
    }
}

# Create or reset
if ($script:ExitCode -eq 0) {
    $sec      = ConvertTo-SecureString $plain -AsPlainText -Force
    $existing = Get-LocalAccount -Name $script:Account

    if ($existing -and $existing.SID.Value -match '-500$') {
        Set-Failed 12 ('{0} is the built-in Administrator (RID 500). Refusing to modify it. Use a different name.' -f $script:Account)
    } elseif ($existing) {
        Write-Log INFO ('Account {0} exists. Resetting password and ensuring enabled.' -f $script:Account)
        try {
            Set-LocalUser -Name $script:Account -Password $sec -PasswordNeverExpires $true -AccountNeverExpires -ErrorAction Stop
            Write-Log OK 'Password reset (password and account set to never expire)'
            $script:Actions.Add('PASSWORD_RESET')
        } catch {
            Set-Failed 20 ('Password reset failed: {0}' -f $_.Exception.Message)
        }
        if (-not $existing.Enabled) {
            try {
                Enable-LocalUser -Name $script:Account -ErrorAction Stop
                Write-Log OK 'Account enabled'
                $script:Actions.Add('ENABLED')
            } catch {
                Set-Failed 20 ('Enable failed: {0}' -f $_.Exception.Message)
            }
        }
        try {
            if (Clear-Lockout -Name $script:Account) {
                Write-Log OK 'Lockout cleared'
                $script:Actions.Add('UNLOCKED')
            }
        } catch {
            Write-Log WARN ('Lockout check failed: {0}' -f $_.Exception.Message)
        }
    } else {
        Write-Log INFO ('Creating local account {0}' -f $script:Account)
        try {
            $newUserArgs = @{
                Name                 = $script:Account
                Password             = $sec
                PasswordNeverExpires = $true
                AccountNeverExpires  = $true
                ErrorAction          = 'Stop'
            }
            if ($script:UseDefaultLabels) {
                $newUserArgs['FullName']    = $AccountFullName
                $newUserArgs['Description'] = $AccountDescription
            }
            New-LocalUser @newUserArgs | Out-Null
            Write-Log OK 'Account created'
            $script:Actions.Add('CREATED')
        } catch {
            Write-Log WARN ('New-LocalUser failed ({0}). Trying ADSI.' -f $_.Exception.Message)
            try {
                New-LocalAccountAdsi -Name $script:Account -Plain $plain
                Write-Log OK 'Account created via ADSI'
                $script:Actions.Add('CREATED_ADSI')
            } catch {
                Set-Failed 20 ('ADSI create failed: {0}' -f $_.Exception.Message)
            }
        }
    }
}

# Administrators membership + verify
if ($script:ExitCode -eq 0) {
    $user = Get-LocalAccount -Name $script:Account
    if (-not $user) {
        Set-Failed 40 'Account not found after create/reset.'
    } else {
        $userSid = $user.SID.Value
        try {
            if (Test-InAdmins -UserSid $userSid) {
                Write-Log INFO 'Already in Administrators'
            } else {
                Add-ToAdmins -Name $script:Account
                Write-Log OK ('Added {0} to Administrators' -f $script:Account)
                $script:Actions.Add('ADDED_TO_ADMINS')
            }
        } catch {
            Set-Failed 30 ('Add to Administrators failed: {0}' -f $_.Exception.Message)
        }

        if ($script:ExitCode -eq 0) {
            $inAdmins = $false
            try { $inAdmins = Test-InAdmins -UserSid $userSid } catch {}
            $u2 = Get-LocalAccount -Name $script:Account
            $enabled = [bool]($u2 -and $u2.Enabled)
            if ($inAdmins -and $enabled) {
                Write-Log OK 'Verify OK: account exists, is enabled, and is in Administrators'
                $script:Actions.Add('VERIFIED')
            } else {
                Set-Failed 40 ('Verify failed: InAdministrators={0} Enabled={1}' -f $inAdmins, $enabled)
            }
        }
    }
}

# Hide / unhide from sign-in screen + verify
if ($script:ExitCode -eq 0) {
    try {
        if ($hideBool) {
            if (Set-HiddenFromLogon -Name $script:Account) {
                Write-Log OK ('Hidden from sign-in screen (UserList\{0} = 0)' -f $script:Account)
                $script:Actions.Add('HIDDEN')
            } else {
                Write-Log INFO 'Already hidden from sign-in screen'
            }
            if (Test-IsHidden -Name $script:Account) {
                $script:HideState = 'HIDDEN'
            } else {
                Set-Failed 50 'Verify failed: UserList value missing or not DWORD 0 after write.'
                $script:HideState = 'FAILED'
            }
        } else {
            if (Clear-HiddenFromLogon -Name $script:Account) {
                Write-Log OK ('Removed UserList\{0} (account visible on sign-in screen)' -f $script:Account)
                $script:Actions.Add('UNHIDDEN')
            } else {
                Write-Log INFO 'Not hidden; nothing to remove'
            }
            if ((Get-UserListEntry -Name $script:Account).Exists) {
                Set-Failed 50 'Verify failed: UserList value still present after removal.'
                $script:HideState = 'FAILED'
            } else {
                $script:HideState = 'VISIBLE'
            }
        }
    } catch {
        Set-Failed 50 ('Hide/unhide failed: {0}' -f $_.Exception.Message)
        $script:HideState = 'FAILED'
    }
}

# Legacy account notice (read-only)
foreach ($legacy in $LegacyAccountNames) {
    if ($legacy -ieq $script:Account) { continue }
    if (Get-LocalAccount -Name $legacy) {
        Write-Log WARN ('Legacy break-glass account {0} is also present. Review and remove manually if retired.' -f $legacy)
    }
}

# Best-effort cleanup of plaintext
$plain = $null
$Password = $null
$pwRes = $null
$sec = $null
[GC]::Collect()

$status = if ($script:ExitCode -eq 0) { 'SUCCESS' } else { 'FAILED' }

$block = @"

========== RESULT ==========
STATUS              : $status
COMPUTER            : $($script:Computer)
ACCOUNT             : $($script:Account)
SIGN-IN SCREEN      : $($script:HideState)
ACTIONS             : $([string]::Join(', ', $script:Actions))
WARNINGS            : $($script:Warnings)
ERRORS              : $($script:Errors)
LOG FILE            : $($script:LogFile)
EXIT CODE           : $($script:ExitCode)
============================
"@
Write-Host $block
Write-Log $(if ($script:ExitCode -eq 0) { 'OK' } else { 'ERROR' }) ('===== Done  STATUS={0}  EXIT={1} =====' -f $status, $script:ExitCode)

exit $script:ExitCode
