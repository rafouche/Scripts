[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [string]$OfflinePath,
    
    [switch]$Online
)

# ========== CONFIGURABLE ARRAYS ==========

# Appx packages to KEEP (all others will be removed)
$AppxKeepList = @(
    "Microsoft.WindowsStore*",
    "Microsoft.Windows.Photos*",
    "Microsoft.WindowsCalculator*",
    "Microsoft.Web*",
    "Microsoft.OneDrive*",
    "microsoft.windowscommunicationsapps*",  # Mail & Calendar
    "Microsoft.VCLibs*",
    "Microsoft.VP9VideoExtensions*",
    "Microsoft.HEIF*",
    "Microsoft.HEVC*",
    "Microsoft.MicrosoftEdge*",
    "Microsoft.WindowsNotepad*",
    "Microsoft.MSPaint*",
    "Microsoft.SecHealth*",
    "Microsoft.Windows.Secure*",
    "Microsoft.Terminal*",
    "Microsoft.DesktopAppInstaller*",
    "Microsoft.MicrosoftStickyNotes*",
    "Microsoft.ScreenSketch*",
    "Microsoft.WindowsCamera*",
    "Microsoft.WindowsAlarms*",
    "Microsoft.WindowsSoundRecorder*",
    "Microsoft.Zune*",
    "Microsoft.BingWeather*",
    "Microsoft.People*",
    "Microsoft.549981C3F5F10*",  # Cortana
    "Microsoft.Clipchamp*",
    "Microsoft.Office.OneNote*",
    "Microsoft.StorePurchaseApp*",
    "Microsoft.ApplicationCompatibilityEnhancements",
    "aimgr",
    "Microsoft.AV1VideoExtension";
    "Microsoft.AVCEncoderVideoExtension",
    "Microsoft.Office.ActionsServer",
    "Microsoft.OfficePushNotificationUtility",
    "Microsoft.YourPhone",
    "MicrosoftWindows.CrossDevice"
)

# Manufacturer-specific packages to ALWAYS REMOVE (HP, Dell, Lenovo, etc.)
$ManufacturerAppxRemoveList = @(
    "*HP*",
    "*Hewlett*",
    "*Dell*",
    "*Lenovo*",
    "*Acer*",
    "*ASUS*",
    "*Toshiba*",
    "*Samsung*",
    "*Sony*",
    "*MSI*",
    "*Gateway*",
    "*Panasonic*",
    "*Fujitsu*",
    "*LG*",
    "*Xiaomi*",
    "*Huawei*"
)

# OEM OOBE XML files to remove
$OOBEXmlFiles = @(
    "\System32\OEM\OOBE.xml",
    "\System32\OEM\OOBE\OOBE.xml",
    "\System32\OEM\Registration.xml",
    "\System32\OEM\*\OOBE.xml",
    "\Panther\OOBE.xml",
    "\System32\Sysprep\OOBE.xml",
    "\System32\Sysprep\Panther\OOBE.xml"
)

# Manufacturer directories to clean
$ManufacturerDirectories = @(
    "\Program Files\HP",
    "\Program Files (x86)\HP",
    "\Program Files\Dell",
    "\Program Files (x86)\Dell",
    "\Program Files\Lenovo",
    "\Program Files (x86)\Lenovo",
    "\Program Files\Acer",
    "\Program Files (x86)\Acer",
    "\Program Files\ASUS",
    "\Program Files (x86)\ASUS",
    "\ProgramData\HP",
    "\ProgramData\Dell",
    "\ProgramData\Lenovo",
    "\Windows\System32\OEM",
    "\Windows\OEM",
    "\Users\Public\Desktop\HP*.lnk",
    "\Users\Public\Desktop\Dell*.lnk",
    "\Users\Public\Desktop\Lenovo*.lnk"
)

# ========== REGISTRY TWEAKS CONFIGURATION ==========

# Default User Registry Tweaks
$DefaultUserTweaks = @(
    # Content Delivery Manager settings
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SystemPaneSuggestionsEnabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-338393Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-353694Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-338388Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-353698Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-338387Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-310093Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-338389Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-353696Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContent-314559Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SoftLandingEnabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "PreInstalledAppsEnabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "PreInstalledAppsEverEnabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "OEMPreInstalledAppsEnabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SilentInstalledAppsEnabled"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "ContentDeliveryAllowed"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name = "SubscribedContentEnabled"; Value = 0; Type = "DWORD"},
    
    # Privacy settings
    @{Path = "Software\Policies\Microsoft\Windows\CloudContent"; Name = "DisableWindowsConsumerFeatures"; Value = 1; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name = "ShowSyncProviderNotifications"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Policies\Microsoft\Windows\DataCollection"; Name = "DoNotShowFeedbackNotifications"; Value = 1; Type = "DWORD"},
    @{Path = "Software\Microsoft\Siuf\Rules"; Name = "NumberOfSIUFInPeriod"; Value = 0; Type = "DWORD"},
    
    # UI customization
    @{Path = "Control Panel\International\User Profile"; Name = "HttpAcceptLanguageOptOut"; Value = 1; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People"; Name = "PeopleBand"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"; Name = "AppsUseLightTheme"; Value = 1; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name = "TaskbarMn"; Value = 0; Type = "DWORD"},
    @{Path = "Control Panel\UnsupportedHardwareNotificationCache"; Name = "SV1"; Value = 0; Type = "DWORD"},
    @{Path = "Control Panel\UnsupportedHardwareNotificationCache"; Name = "SV2"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\UserProfileEngagement"; Name = "ScoobeSystemSettingEnabled"; Value = 0; Type = "DWORD"},
    
    # Additional privacy
    @{Path = "SOFTWARE\Microsoft\Personalization\Settings"; Name = "AcceptedPrivacyPolicy"; Value = 0; Type = "DWORD"},
    @{Path = "SOFTWARE\Microsoft\Windows\CurrentVersion\SettingSync\Groups\Language"; Name = "Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "SOFTWARE\Microsoft\InputPersonalization"; Name = "RestrictImplicitTextCollection"; Value = 1; Type = "DWORD"},
    @{Path = "SOFTWARE\Microsoft\InputPersonalization"; Name = "RestrictImplicitInkCollection"; Value = 1; Type = "DWORD"},
    @{Path = "SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore"; Name = "HarvestContacts"; Value = 0; Type = "DWORD"},
    @{Path = "SOFTWARE\Microsoft\Windows\CurrentVersion\AppHost"; Name = "EnableWebContentEvaluation"; Value = 1; Type = "DWORD"},
    
    # Taskbar customization
    @{Path = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name = "TaskbarDa"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name = "ShowTaskViewButton"; Value = 0; Type = "DWORD"},
    @{Path = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name = "TaskbarAl"; Value = 0; Type = "DWORD"},
    
    # Classic right-click menu
    @{Path = "Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"; Name = "(Default)"; Value = ""; Type = "SZ"}
)

# Software (Machine) Registry Tweaks
$SoftwareTweaks = @(
    # Cloud Content policies
    @{Path = "Policies\Microsoft\Windows\CloudContent"; Name = "DisableSoftLanding"; Value = 1; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\CloudContent"; Name = "DisableWindowsConsumerFeatures"; Value = 1; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\CloudContent"; Name = "DisableThirdPartySuggestions"; Value = 1; Type = "DWORD"},
    
    # Data Collection
    @{Path = "Policies\Microsoft\Windows\DataCollection"; Name = "AllowTelemetry"; Value = 0; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\PreviewBuilds"; Name = "EnableConfigFlighting"; Value = 0; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\DataCollection"; Name = "DoNotShowFeedbackNotifications"; Value = 1; Type = "DWORD"},
    
    # Delivery Optimization
    @{Path = "Policies\Microsoft\Windows\DeliveryOptimization"; Name = "DODownloadMode"; Value = 1; Type = "DWORD"},
    
    # Edge policies
    @{Path = "Policies\Microsoft\MicrosoftEdge\Main"; Name = "DoNotTrack"; Value = 1; Type = "DWORD"},
    
    # OneDrive policies
    @{Path = "Policies\Microsoft\Windows\OneDrive"; Name = "DehydrateSyncedTeamSites"; Value = 1; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\OneDrive"; Name = "FilesOnDemandEnabled"; Value = 1; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\OneDrive"; Name = "SilentAccountConfig"; Value = 1; Type = "DWORD"},
    
    # Windows Update policies
    @{Path = "Policies\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"; Name = "EnableFeaturedSoftware"; Value = 0; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\WindowsUpdate"; Name = "ExcludeWUDriversInQualityUpdate"; Value = 0; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "AllowMUUpdateService"; Value = 1; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "AUOptions"; Value = 4; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "AutomaticMaintenanceEnabled"; Value = 1; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "NoAutoUpdate"; Value = 0; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "ScheduledInstallDay"; Value = 7; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "ScheduledInstallThirdWeek"; Value = 1; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "ScheduledInstallTime"; Value = 1; Type = "DWORD"},
    
    # Other software settings
    @{Path = "Microsoft\Windows\CurrentVersion\AdvertisingInfo"; Name = "Enabled"; Value = 0; Type = "DWORD"},
    @{Path = "Microsoft\Windows\CurrentVersion\Device Metadata"; Name = "PreventDeviceMetadataFromNetwork"; Value = 1; Type = "DWORD"},
    @{Path = "Microsoft\Windows\CurrentVersion\Explorer"; Name = "DisableEdgeDesktopShortcutCreation"; Value = 0; Type = "DWORD"},
    @{Path = "Microsoft\Windows\CurrentVersion\AppHost"; Name = "EnableWebContentEvaluation"; Value = 1; Type = "DWORD"},
    @{Path = "Microsoft\Windows\CurrentVersion\Themes\Personalize"; Name = "AppsUseLightTheme"; Value = 1; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\Windows Chat"; Name = "ChatIcon"; Value = 3; Type = "DWORD"},
    
    # OOBE settings
    @{Path = "Microsoft\Windows\CurrentVersion\OOBE"; Name = "BypassNRO"; Value = 1; Type = "DWORD"},
    @{Path = "Microsoft\Windows\CurrentVersion\OOBE"; Name = "SkipMachineOOBE"; Value = 1; Type = "DWORD"},
    @{Path = "Microsoft\Windows\CurrentVersion\OOBE"; Name = "SkipUserOOBE"; Value = 1; Type = "DWORD"},
    @{Path = "Microsoft\Windows\CurrentVersion\OOBE"; Name = "DisableOEMRegistration"; Value = 1; Type = "DWORD"},
    
    # Cortana/Web Search
    @{Path = "Policies\Microsoft\Windows\Windows Search"; Name = "AllowCortana"; Value = 0; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\Windows Search"; Name = "DisableWebSearch"; Value = 1; Type = "DWORD"},

    # Remove Start Menu Suggestions
    @{Path = "Policies\Microsoft\Windows\Windows\Explorer"; Name = "HideRecommendedSection"; Value = 1; Type = "DWORD"},

    # Turn off Bing suggestions and predictive searches
    @{Path = "Policies\Microsoft\Windows\Windows\Explorer"; Name = "DisableSearchBoxSuggestions"; Value= 1; Type = "DWORD"},

    # Remove weather, news, and Copilot widgets from the Start menu
    @{Path = "Policies\Microsoft\Windows\Windows\Explorer"; Name = "AllowNewsAndInterest"; Value = 0; Type = "DWORD"},
    @{Path = "Policies\Microsoft\Windows\Windows\Explorer"; Name = "TurnOffWindowsCopilot"; Value = 1; Type = "DWORD"}

)

# System Registry Tweaks
$SystemTweaks = @(
    # Windows 11 upgrade bypasses
    @{Path = "Setup\LabConfig"; Name = "BypassCPUCheck"; Value = 1; Type = "DWORD"},
    @{Path = "Setup\LabConfig"; Name = "BypassRAMCheck"; Value = 1; Type = "DWORD"},
    @{Path = "Setup\LabConfig"; Name = "BypassSecureBootCheck"; Value = 1; Type = "DWORD"},
    @{Path = "Setup\LabConfig"; Name = "BypassStorageCheck"; Value = 1; Type = "DWORD"},
    @{Path = "Setup\LabConfig"; Name = "BypassTPMCheck"; Value = 1; Type = "DWORD"},
    @{Path = "Setup\MoSetup"; Name = "AllowUpgradesWithUnsupportedTPMOrCPU"; Value = 1; Type = "DWORD"},
    
    # Disable unwanted services
    @{Path = "ControlSet001\Services\DiagTrack"; Name = "Start"; Value = 4; Type = "DWORD"},
    @{Path = "ControlSet001\Services\XblAuthManager"; Name = "Start"; Value = 4; Type = "DWORD"},
    @{Path = "ControlSet001\Services\XblGameSave"; Name = "Start"; Value = 4; Type = "DWORD"},
    @{Path = "ControlSet001\Services\XboxNetApiSvc"; Name = "Start"; Value = 4; Type = "DWORD"},
    @{Path = "ControlSet001\Services\WMPNetworkSvc"; Name = "Start"; Value = 4; Type = "DWORD"}
)

# ========== HELPER FUNCTIONS ==========

function Test-IsOfflineImage {
    param([string]$Path)
    
    if (-not $Path) { return $false }
    
    # Check if path is a mounted Windows directory
    if (Test-Path "$Path\Windows\System32") {
        return $true
    }
    
    # Check if path is a WIM file
    if ($Path -match '\.wim$' -and (Test-Path $Path)) {
        return $true
    }
    
    return $false
}

function Mount-WimImage {
    param([string]$WimPath)
    
    $MountDir = "$env:TEMP\ImageMount_$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -ItemType Directory -Path $MountDir -Force | Out-Null
    
    try {
        Write-Host "Mounting WIM image: $WimPath" -ForegroundColor Cyan
        dism /Mount-Image /ImageFile:$WimPath /Index:1 /MountDir:$MountDir | Out-Null
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount WIM image"
        }
        
        return $MountDir
    } catch {
        Write-Error "Error mounting image: $_"
        if (Test-Path $MountDir) {
            Remove-Item -Path $MountDir -Force -Recurse
        }
        return $null
    }
}

function Unmount-WimImage {
    param([string]$MountDir, [switch]$Commit)
    
    if ($Commit) {
        Write-Host "Committing changes and unmounting image..." -ForegroundColor Cyan
        dism /Unmount-Image /MountDir:$MountDir /Commit | Out-Null
    } else {
        Write-Host "Discarding changes and unmounting image..." -ForegroundColor Yellow
        dism /Unmount-Image /MountDir:$MountDir /Discard | Out-Null
    }
    
    if (Test-Path $MountDir) {
        Remove-Item -Path $MountDir -Force -Recurse -ErrorAction SilentlyContinue
    }
}

function Invoke-OfflineRegistryOperation {
    param(
        [string]$MountPath,
        [scriptblock]$ScriptBlock
    )
    
    $Hives = @{
        'DefaultUser' = @{Path = "$MountPath\Users\Default\ntuser.dat"; Loaded = $null}
        'Software' = @{Path = "$MountPath\Windows\System32\config\SOFTWARE"; Loaded = $null}
        'System' = @{Path = "$MountPath\Windows\System32\config\SYSTEM"; Loaded = $null}
    }
    
    try {
        # Load registry hives
        foreach ($key in $Hives.Keys) {
            if (Test-Path $Hives[$key].Path) {
                $HiveName = "Offline$key"
                reg load "HKLM\$HiveName" $Hives[$key].Path 2>&1 | Out-Null
                $Hives[$key].Loaded = $HiveName
            }
        }
        
        # Execute scriptblock with loaded hive names
        & $ScriptBlock $Hives['DefaultUser'].Loaded $Hives['Software'].Loaded $Hives['System'].Loaded
        
    } finally {
        # Unload registry hives
        [gc]::Collect()
        foreach ($key in $Hives.Keys) {
            if ($Hives[$key].Loaded) {
                reg unload "HKLM\$($Hives[$key].Loaded)" 2>&1 | Out-Null
            }
        }
    }
}

function Apply-RegistryTweaks {
    param(
        [string]$HiveRoot,
        [array]$Tweaks,
        [string]$HiveType = "DEFAULT"
    )
    
    Write-Host "Applying $HiveType registry tweaks..." -ForegroundColor Cyan
    
    foreach ($tweak in $Tweaks) {
        $FullPath = "$HiveRoot\$($tweak.Path)"
        
        # Ensure the registry path exists
        $ParentPath = Split-Path $FullPath -Parent
        $KeyName = Split-Path $FullPath -Leaf
        
        try {
            # Create parent key if it doesn't exist
            reg add "$ParentPath" /f 2>&1 | Out-Null
            
            # Apply the registry value
            switch ($tweak.Type) {
                "DWORD" {
                    reg add "$ParentPath" /t REG_DWORD /v $tweak.Name /d $tweak.Value /f 2>&1 | Out-Null
                }
                "SZ" {
                    reg add "$ParentPath" /t REG_SZ /v $tweak.Name /d $tweak.Value /f 2>&1 | Out-Null
                }
                "DELETE" {
                    reg delete "$ParentPath" /v $tweak.Name /f 2>&1 | Out-Null
                }
            }
            
            Write-Verbose "Applied: $FullPath\$($tweak.Name) = $($tweak.Value)"
            
        } catch {
            Write-Warning "Failed to apply registry tweak: $FullPath\$($tweak.Name) - $_"
        }
    }
}

function Remove-AppxPackages {
    param([string]$MountPath, [bool]$IsOnline)
    
    Write-Host "Processing Appx packages..." -ForegroundColor Cyan
    
    if ($IsOnline) {
        # Online mode - remove from current system
        
        # First, remove manufacturer packages
        foreach ($pattern in $ManufacturerAppxRemoveList) {
            try {
                $Packages = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like $pattern -or $_.Publisher -like $pattern }
                foreach ($pkg in $Packages) {
                    Write-Host "Removing manufacturer package: $($pkg.Name)" -ForegroundColor Green
                    Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction SilentlyContinue
                }
            } catch {}
        }
        
        # Remove provisioned manufacturer packages
        $Provisioned = Get-AppxProvisionedPackage -Online
        foreach ($pkg in $Provisioned) {
            $ShouldRemove = $false
            
            # Check if it's a manufacturer package
            foreach ($pattern in $ManufacturerAppxRemoveList) {
                if ($pkg.DisplayName -like $pattern -or $pkg.PublisherId -like $pattern) {
                    $ShouldRemove = $true
                    break
                }
            }
            
            # Check if it's NOT in the keep list
            if (-not $ShouldRemove) {
                $ShouldKeep = $false
                foreach ($keepPattern in $AppxKeepList) {
                    if ($pkg.DisplayName -like $keepPattern) {
                        $ShouldKeep = $true
                        break
                    }
                }
                $ShouldRemove = -not $ShouldKeep
            }
            
            if ($ShouldRemove) {
                try {
                    Write-Host "Removing provisioned: $($pkg.DisplayName)" -ForegroundColor Yellow
                    Remove-AppxProvisionedPackage -Online -PackageName $pkg.PackageName -ErrorAction SilentlyContinue
                } catch {}
            }
        }
        
    } else {
        # Offline mode - remove from mounted image
        
        $Provisioned = Get-AppxProvisionedPackage -Path $MountPath
        foreach ($pkg in $Provisioned) {
            $ShouldRemove = $false
            
            # Check if it's a manufacturer package
            foreach ($pattern in $ManufacturerAppxRemoveList) {
                if ($pkg.DisplayName -like $pattern -or $pkg.PublisherId -like $pattern) {
                    $ShouldRemove = $true
                    break
                }
            }
            
            # Check if it's NOT in the keep list
            if (-not $ShouldRemove) {
                $ShouldKeep = $false
                foreach ($keepPattern in $AppxKeepList) {
                    if ($pkg.DisplayName -like $keepPattern) {
                        $ShouldKeep = $true
                        break
                    }
                }
                $ShouldRemove = -not $ShouldKeep
            }
            
            if ($ShouldRemove) {
                try {
                    Write-Host "Removing provisioned: $($pkg.DisplayName)" -ForegroundColor Yellow
                    Remove-AppxProvisionedPackage -Path $MountPath -PackageName $pkg.PackageName -ErrorAction SilentlyContinue
                } catch {}
            }
        }
    }
}

function Remove-OOBEXmlFiles {
    param([string]$WindowsPath)
    
    Write-Host "Cleaning OOBE XML files..." -ForegroundColor Cyan
    
    foreach ($xmlFile in $OOBEXmlFiles) {
        $FullPath = "$WindowsPath$xmlFile"
        if (Test-Path $FullPath) {
            try {
                Write-Host "Removing: $FullPath" -ForegroundColor Green
                Remove-Item -Path $FullPath -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Warning "Could not remove $FullPath : $_"
            }
        }
    }
    
    # Clean OEM directory
    $OEMDir = "$WindowsPath\System32\OEM"
    if (Test-Path $OEMDir) {
        Write-Host "Cleaning OEM directory..." -ForegroundColor Cyan
        Get-ChildItem -Path $OEMDir -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse
    }
}

function Remove-ManufacturerDirectories {
    param([string]$RootPath)
    
    Write-Host "Cleaning manufacturer directories..." -ForegroundColor Cyan
    
    foreach ($dir in $ManufacturerDirectories) {
        $FullPath = "$RootPath$dir"
        
        if (Test-Path $FullPath) {
            try {
                if ($dir -like "*.lnk") {
                    # Remove shortcut files
                    Get-ChildItem -Path $FullPath -ErrorAction SilentlyContinue | Remove-Item -Force
                    Write-Host "Removed shortcuts: $FullPath" -ForegroundColor Green
                } else {
                    # Remove directories
                    Remove-Item -Path $FullPath -Force -Recurse -ErrorAction SilentlyContinue
                    Write-Host "Removed directory: $FullPath" -ForegroundColor Green
                }
            } catch {
                Write-Warning "Could not clean $FullPath : $_"
            }
        }
    }
}

function Remove-ScheduledTasks {
    param([string]$MountPath, [bool]$IsOnline)
    
    Write-Host "Cleaning scheduled tasks..." -ForegroundColor Cyan
    
    if ($IsOnline) {
        # Online mode
        $ManufacturerTasks = Get-ScheduledTask | Where-Object {
            foreach ($pattern in $ManufacturerAppxRemoveList) {
                if ($_.TaskName -like $pattern -or $_.Description -like $pattern) {
                    return $true
                }
            }
            return $false
        }
        
        foreach ($task in $ManufacturerTasks) {
            try {
                Write-Host "Removing task: $($task.TaskName)" -ForegroundColor Yellow
                Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
            } catch {}
        }
        
    } else {
        # Offline mode
        $TasksPath = "$MountPath\Windows\System32\Tasks"
        if (Test-Path $TasksPath) {
            foreach ($pattern in $ManufacturerAppxRemoveList) {
                $CleanPattern = $pattern.Replace("*", "")
                if ($CleanPattern) {
                    Get-ChildItem -Path $TasksPath -Filter "*$CleanPattern*" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force
                }
            }
        }
    }
}

function Set-StartMenuLayout {
    param([string]$MountPath)
    
    Write-Host "Setting start menu layout..." -ForegroundColor Cyan
    
    $WindowsPath = if ($MountPath) { $MountPath } else { "C:" }
    $LayoutFile = "$WindowsPath\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml"
    $JsonLayoutFile = "$WindowsPath\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.json"
    
    # Remove existing layout files
    if (Test-Path $LayoutFile) { Remove-Item $LayoutFile -Force }
    if (Test-Path $JsonLayoutFile) { Remove-Item $JsonLayoutFile -Force }
    
    # Create directory if it doesn't exist
    $LayoutDir = Split-Path $LayoutFile -Parent
    if (!(Test-Path $LayoutDir)) {
        New-Item -ItemType Directory -Path $LayoutDir -Force | Out-Null
    }
    
    # Create Windows 11 JSON layout file
    $Start11Layout = @'
{
    "pinnedList": [
        {"desktopAppLink": "%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\System Tools\\File Explorer.lnk"},
        {"packagedAppId": "Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge"},
        {"packagedAppId": "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App"},
        {"packagedAppId": "Microsoft.Windows.Photos_8wekyb3d8bbwe!App"}
    ]
}
'@
    
    $Start11Layout | Set-Content -Path $JsonLayoutFile -Force
    Write-Host "Created start menu layout at: $JsonLayoutFile" -ForegroundColor Green
}

function Create-UnattendXml {
    param([string]$MountPath)
    
    Write-Host "Creating unattend.xml..." -ForegroundColor Cyan
    
    $WindowsPath = if ($MountPath) { $MountPath } else { "C:" }
    $UnattendPath = "$WindowsPath\Windows\System32\Sysprep\unattend.xml"
    
    $UnattendContent = @'
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
    <settings pass="oobeSystem">
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <OOBE>
                <HideEULAPage>true</HideEULAPage>
                <SkipMachineOOBE>true</SkipMachineOOBE>
                <SkipUserOOBE>true</SkipUserOOBE>
                <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
                <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
                <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
                <NetworkLocation>Work</NetworkLocation>
                <ProtectYourPC>1</ProtectYourPC>
            </OOBE>
        </component>
    </settings>
</unattend>
'@
    
    # Ensure directory exists
    $UnattendDir = Split-Path $UnattendPath -Parent
    if (!(Test-Path $UnattendDir)) {
        New-Item -ItemType Directory -Path $UnattendDir -Force | Out-Null
    }
    
    $UnattendContent | Set-Content -Path $UnattendPath -Force
    Write-Host "Created unattend.xml at: $UnattendPath" -ForegroundColor Green
}

# ========== MAIN SCRIPT LOGIC ==========

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  Generic Windows Image Cleanup Tool" -ForegroundColor Yellow
Write-Host "=========================================" -ForegroundColor Cyan

# Determine operation mode
$IsMounted = $false
$OriginalPath = $OfflinePath

if ($OfflinePath) {
    if ($OfflinePath -match '\.wim$') {
        # It's a WIM file, need to mount it
        $MountPath = Mount-WimImage -WimPath $OfflinePath
        if (-not $MountPath) {
            Write-Error "Failed to mount WIM image"
            exit 1
        }
        $IsMounted = $true
    } elseif (Test-Path "$OfflinePath\Windows\System32") {
        # It's already a mounted directory
        $MountPath = $OfflinePath
        $IsOnline = $false
    } else {
        Write-Error "Invalid offline path. Must be either a WIM file or mounted Windows directory."
        exit 1
    }
} elseif ($Online) {
    $MountPath = $null
    $IsOnline = $true
} else {
    # Default to current system
    $MountPath = $null
    $IsOnline = $true
}

try {
    # Get Windows path
    $WindowsPath = if ($MountPath) { "$MountPath\Windows" } else { "C:\Windows" }
    $RootPath = if ($MountPath) { $MountPath } else { "C:" }
    
    # 1. Remove OOBE XML files
    Remove-OOBEXmlFiles -WindowsPath $WindowsPath
    
    # 2. Apply registry tweaks
    if ($IsOnline) {
        # Online mode - apply directly
        Apply-RegistryTweaks -HiveRoot "HKCU" -Tweaks $DefaultUserTweaks -HiveType "Default User"
        Apply-RegistryTweaks -HiveRoot "HKLM" -Tweaks $SoftwareTweaks -HiveType "Software"
        Apply-RegistryTweaks -HiveRoot "HKLM" -Tweaks $SystemTweaks -HiveType "System"
    } else {
        # Offline mode - load hives and apply
        Invoke-OfflineRegistryOperation -MountPath $MountPath -ScriptBlock {
            param($DefHive, $SoftHive, $SysHive)
            
            if ($DefHive) {
                Apply-RegistryTweaks -HiveRoot "HKLM:\$DefHive" -Tweaks $DefaultUserTweaks -HiveType "Default User"
            }
            if ($SoftHive) {
                Apply-RegistryTweaks -HiveRoot "HKLM:\$SoftHive" -Tweaks $SoftwareTweaks -HiveType "Software"
            }
            if ($SysHive) {
                Apply-RegistryTweaks -HiveRoot "HKLM:\$SysHive" -Tweaks $SystemTweaks -HiveType "System"
            }
        }
    }
    
    # 3. Remove Appx packages
    Remove-AppxPackages -MountPath $MountPath -IsOnline $IsOnline
    
    # 4. Remove manufacturer directories
    Remove-ManufacturerDirectories -RootPath $RootPath
    
    # 5. Remove scheduled tasks
    Remove-ScheduledTasks -MountPath $MountPath -IsOnline $IsOnline
    
    # 6. Set start menu layout
    Set-StartMenuLayout -MountPath $MountPath
    
    # 7. Create unattend.xml
    Create-UnattendXml -MountPath $MountPath
    
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "  Image cleanup completed successfully!" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Cyan
    
    # Display summary
    Write-Host "`nSummary of actions performed:" -ForegroundColor Yellow
    Write-Host "- Removed manufacturer OOBE files" -ForegroundColor White
    Write-Host "- Applied privacy and UI registry tweaks" -ForegroundColor White
    Write-Host "- Removed unwanted Appx packages" -ForegroundColor White
    Write-Host "- Cleaned manufacturer directories" -ForegroundColor White
    Write-Host "- Removed manufacturer scheduled tasks" -ForegroundColor White
    Write-Host "- Set clean start menu layout" -ForegroundColor White
    Write-Host "- Created unattend.xml for OOBE" -ForegroundColor White
    
} finally {
    # Clean up mounted image if we mounted it
    if ($IsMounted -and $MountPath) {
        $Commit = Read-Host "`nCommit changes to WIM image? (Y/N)"
        if ($Commit -eq 'Y' -or $Commit -eq 'y') {
            Unmount-WimImage -MountDir $MountPath -Commit
            Write-Host "Changes committed to: $OriginalPath" -ForegroundColor Green
        } else {
            Unmount-WimImage -MountDir $MountPath
            Write-Host "Changes discarded" -ForegroundColor Yellow
        }
    }
}

Write-Host "`nScript completed!" -ForegroundColor Green