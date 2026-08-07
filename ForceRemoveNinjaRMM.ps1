# Force Remove NinjaRMM / NinjaOne Agent
# Run as Administrator

Write-Host "Stopping and removing NinjaRMM Agent..." -ForegroundColor Yellow

try {
    # Stop NinjaRMM services if they exist
    $services = @("NinjaRMMAgent", "NinjaRMMAgentService")
    foreach ($svc in $services) {
        if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
            Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
            sc.exe delete $svc | Out-Null
            Write-Host "Service $svc stopped and deleted."
        }
    }

    # Kill any running NinjaRMM processes
    $processes = @("NinjaRMMAgent", "NinjaRMMAgentService", "NinjaRMMAgentUpdater")
    foreach ($proc in $processes) {
        Get-Process -Name $proc -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }

    # Attempt to run the uninstaller if present
    $uninstallPaths = @(
        "C:\Program Files\NinjaRMMAgent\uninstall.exe",
        "C:\Program Files (x86)\NinjaRMMAgent\uninstall.exe"
    )
    foreach ($path in $uninstallPaths) {
        if (Test-Path $path) {
            Write-Host "Running uninstaller at $path..."
            Start-Process $path "/S" -Wait -ErrorAction SilentlyContinue
        }
    }

    # Remove leftover folders
    $folders = @(
        "C:\Program Files\NinjaRMMAgent",
        "C:\Program Files (x86)\NinjaRMMAgent",
        "C:\ProgramData\NinjaRMMAgent"
    )
    foreach ($folder in $folders) {
        if (Test-Path $folder) {
            Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Deleted folder: $folder"
        }
    }

    # Remove registry uninstall entries
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($regPath in $regPaths) {
        Get-ChildItem $regPath -ErrorAction SilentlyContinue | ForEach-Object {
            $dispName = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).DisplayName
            if ($dispName -like "*NinjaRMM*") {
                Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Removed registry entry for: $dispName"
            }
        }
    }

    Write-Host "NinjaRMM Agent removal complete." -ForegroundColor Green
}
catch {
    Write-Host "Error during removal: $_" -ForegroundColor Red
}
