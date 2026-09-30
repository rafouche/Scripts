<#
Latest Wrike Desktop Install - NinjaOne library script (PowerShell, Windows, run as SYSTEM)

Same steps as the "Latest Wrike Desktop Install" Install Application automation,
as a plain library script so it can be run through the NinjaOne API by script id
(the API key the Help Desk agent uses cannot start Install Application
automations - NinjaOne answers "user_context_required" for those).

Steps: record the installed Wrike version, close Wrike if it is running,
download Wrike's current MSI from Wrike's own download link, install it silently
(same msiexec switches as the automation: /quiet /qn /norestart, REINSTALL=ALL
REINSTALLMODE=ecmus so an older or equal version is refreshed in place), then
print the before/after version so the ticket note can quote it.
Exit code 0 = installed, non-zero = failed (the log path is printed).
#>
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$msiUrl = if ([Environment]::Is64BitOperatingSystem) { 'https://dl.wrike.com/download/WrikeDesktopApp.latest.msi' } else { 'https://dl.wrike.com/download/WrikeDesktopApp.latest.x32.msi' }
$work   = Join-Path $env:ProgramData 'NinjaRMMAgent\download\wrike-latest'
$msi    = Join-Path $work 'WrikeDesktopApp.latest.msi'
$log    = Join-Path $work ('wrike-install-{0:yyyyMMdd-HHmmss}.log' -f (Get-Date))

function Get-WrikeVersion {
    $keys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    $hit = Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like 'Wrike for Windows*' } | Select-Object -First 1
    if ($hit) { return $hit.DisplayVersion } else { return $null }
}

$before = Get-WrikeVersion
Write-Output ("Wrike for Windows before: {0}" -f ($(if ($before) { $before } else { 'not installed' })))

New-Item -ItemType Directory -Path $work -Force | Out-Null
Write-Output "Downloading $msiUrl"
Invoke-WebRequest -Uri $msiUrl -OutFile $msi -UseBasicParsing
$size = (Get-Item $msi).Length
if ($size -lt 5MB) { throw "Download looks wrong ($size bytes) - not installing." }
Write-Output ("Downloaded {0:N1} MB" -f ($size / 1MB))

$running = Get-Process -Name 'Wrike' -ErrorAction SilentlyContinue
if ($running) { Write-Output "Closing running Wrike ($($running.Count) process(es))"; $running | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 3 }
else { Write-Output "Wrike is not running" }

$args = '/i "{0}" /quiet /qn /norestart REINSTALLMODE="ecmus" REINSTALL="ALL" /log "{1}"' -f $msi, $log
Write-Output "Running: msiexec $args"
$p = Start-Process -FilePath 'msiexec.exe' -ArgumentList $args -Wait -PassThru
Write-Output "msiexec exit code: $($p.ExitCode) (log: $log)"

$after = Get-WrikeVersion
Write-Output ("Wrike for Windows after: {0}" -f ($(if ($after) { $after } else { 'not installed' })))
if ($p.ExitCode -notin 0, 1641, 3010) { Write-Output "RESULT: FAILED"; exit $p.ExitCode }
Write-Output ("RESULT: SUCCESS - Wrike {0} -> {1}{2}" -f ($(if ($before) { $before } else { 'none' })), $after, $(if ($p.ExitCode -eq 3010 -or $p.ExitCode -eq 1641) { ' (reboot requested by installer)' } else { '' }))
Remove-Item $msi -Force -ErrorAction SilentlyContinue
exit 0
