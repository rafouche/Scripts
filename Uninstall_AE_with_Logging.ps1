# Copyright (c) 2023 CyberFOX LLC
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#    * Redistributions of source code must retain the above copyright
#      notice, this list of conditions and the following disclaimer.
#    * Redistributions in binary form must reproduce the above copyright
#      notice, this list of conditions and the following disclaimer in the
#      documentation and/or other materials provided with the distribution.
#    * Neither the name of the AutoElevate nor the names of its contributors
#      may be used to endorse or promote products derived from this software
#      without specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL OPENDNS BE LIABLE FOR ANY DIRECT, INDIRECT,
# INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
# LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA,
# OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
# LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
# NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE,
# EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

<#
.SYNOPSIS
  Uninstalls the AutoElevate agent and creates a log file on the root of the C: drive
  #>

$softwareDisplayName = "AutoElevate"
$serviceName = "AutoElevateAgent"
$logFilePath = "C:\AutoElevateUninstallLog.txt"    
$registryKeyPath = "HKLM:\Software\autoelevate"
$folderPath = "C:\Program Files*\AutoElevate"

$ScriptFailed = "Script Failed!"

function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

function Debug-Print ($msg) {
    if ($DebugPrintEnabled -eq 1) {
        Write-Host "$(Get-TimeStamp) [DEBUG] $msg"
    }
}

function Kill-Service {
# Check if the script is running with administrator privileges
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as an administrator."
    exit
}
taskkill /F /FI "SERVICES eq AutoElevateAgent"
}

function Del-Registry-Key {  
# Check if the registry key exists before attempting to delete it
if (Test-Path -Path $registryKeyPath) {
    # Prompt for confirmation (uncomment this line if you want to prompt for confirmation)
    # $confirmation = Read-Host "Are you sure you want to delete $registryKeyPath? (Y/N)"
    
    # If the user confirms (or if you skip confirmation), delete the registry key
    # if ($confirmation -eq "Y" -or $confirmation -eq "y") {
    Remove-Item -Path $registryKeyPath -Recurse -Force
    Write-Host "Registry key $registryKeyPath deleted successfully."
    # }
} else {
    Write-Host "Registry key $registryKeyPath not found."
}
}

function Del-AEFolder {  
# Check if the folder exists
if (Test-Path -Path $folderPath -PathType Container) {
    # Delete the folder and its contents
    Remove-Item -Path $folderPath -Recurse -Force
    Write-Host "Folder '$folderPath' deleted successfully."
} else {
    Write-Host "Folder '$folderPath' does not exist."
}
# Function to log messages to the log file
function Write-Host {
    param (
        [string]$message
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    $logMessage | Out-File -Append -FilePath $logFilePath
}
}

function UninstallAE {  
# Check if the log file exists, if not, create it
if (-not (Test-Path -Path $logFilePath)) {
    New-Item -Path $logFilePath -ItemType File
}

Write-Host "Uninstallation process started for '$softwareDisplayName'."

$uninstallKey = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $softwareDisplayName }

if ($uninstallKey) {
    Write-Host "Uninstalling '$softwareDisplayName'..."
    $uninstallResult = $uninstallKey.Uninstall()
    
    if ($uninstallResult.ReturnValue -eq 0) {
        Write-Host "'$softwareDisplayName' was successfully uninstalled."
    } else {
        Write-Host "Failed to uninstall '$softwareDisplayName'. Return code: $($uninstallResult.ReturnValue)"
    }
} else {
    Write-Host "Software '$softwareDisplayName' not found in the list of installed programs."
}
}

function main () {
    Debug-Print("Checking for Begining uninstall...")
    
    Write-Host "$(Get-TimeStamp) SofwareName: " $softwareDisplayName
    Write-Host "$(Get-TimeStamp) ServiceName: " $serviceName
    Write-Host "$(Get-TimeStamp) RegistryName: " $registryKeyPath
    Write-Host "$(Get-TimeStamp) DirectoryName: " $folderPath
    Write-Host "$(Get-TimeStamp) RegistryName: " $registryKeyPath
    Write-Host "$(Get-TimeStamp) LogFilePath: " $logFilePath
    
    Kill-Service
    Del-Registry-Key
    Del-AEFolder
    UninstallAE
    
    Write-Host "$(Get-TimeStamp) AutoElevate Agent successfully uninstalled!"
}

try
{
    main
} catch {
    $ErrorMessage = $_.Exception.Message
    Write-Host "$(Get-TimeStamp) $ErrorMessage"
    exit 1
}

Write-Host "Uninstallation process completed. Log file: $logFilePath"