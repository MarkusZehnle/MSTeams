<#
.SYNOPSIS
    Parses the Teams VDI connection info JSON file and outputs key session details including VDI optimization status, connected state, software versions, and peripheral device info.

.DESCRIPTION
    This script reads the 'vdi_connection_info.json' file from the Teams cache directory in a VDI environment.
    It extracts details such as VDI Optimization version, connection state, Teams and plugin versions, client platform, OS version, available and selected peripherals, and secondary ringer.
    Additionally, it converts timestamps to human-readable format and resolves OS build numbers to friendly Windows or Windows Server version names.

.EXAMPLE
    PS C:\> .\Get-TeamsOptimization.ps1
    Outputs all current session info from the Teams VDI JSON file in a readable format.

.EXAMPLE
    PS C:\> .\Get-TeamsOptimization.ps1 | Out-File .\TeamsVDIReport.txt
    Exports the session info to a text file for review.

.NOTES
    Version:        1.0
    Author:         Markus Zehnle (Chat-GPT tbh)
    Created:        2025-05-19
    Requires:       PowerShell 5.1 or later
    JSON Source:    C:\Users\<username>\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\tfw\vdi_connection_info.json

.LINK
    Microsoft Teams VDI Monitoring API documentation (internal/private)
    https://learn.microsoft.com/en-us/microsoftteams/vdi-2#monitoring-api
#>

# Path to JSON file
$tfwPath = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\tfw\vdi_connection_info.json"

# Check if the file exists
if (-Not (Test-Path -Path $tfwPath)) {
    Write-Error "The file vdi_connection_info.json cannot be found: $tfwPath"
    exit 1
}

# Get and convert JSON
$jsonContent = Get-Content -Path $tfwPath -Raw | ConvertFrom-Json

# Get the last entry for the session (new session)
$session = $jsonContent.VdiConnectionInfo | Select-Object -Last 1

# Extract data from JSON
$vdiState = $session.vdiConnectedState
$vdiVersion = $session.vdiVersionInfo
$devices = $session.devices

# Timestamp in human readbale format
$epoch = Get-Date -Date "1970-01-01 00:00:00Z"
$timestampReadable = $epoch.AddMilliseconds($vdiState.timestamp).ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss")

# Functions
# Get Windows version from JSON and registry...
function Get-WindowsVersionName {

    try {
        $osInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $actualBuild = $osInfo.CurrentBuildNumber
        $actualMinor = $osInfo.UBR
        $sku = $osInfo.EditionID
        $productName = $osInfo.ProductName
    } catch {
        $actualBuild = "Unknown"
        $sku = ""
        $productName = ""
        $actualMinor = ""
    }

    # Determine version name
    $versionName = switch ($true) {
        ($actualbuild -eq 26100 -and $productName -like "*Server*") { "Windows Server 2025 24H2"; break }
        ($actualbuild -eq 26100)                                     { "Windows 11 24H2"; break }
        ($actualbuild -eq 25300)                                     { "Windows 11 Insider Dev"; break }
        ($actualbuild -eq 22631)                                     { "Windows 11 23H2"; break }
        ($actualbuild -eq 22621)                                     { "Windows 11 22H2"; break }
        ($actualbuild -eq 22000)                                     { "Windows 11 21H2"; break }
        ($actualbuild -eq 20348)                                     { "Windows Server 2022 21H2"; break }
        ($actualbuild -eq 19045)                                     { "Windows 10 22H2"; break }
        ($actualbuild -eq 19044)                                     { "Windows 10 21H2"; break }
        ($actualbuild -eq 17763)                                     { "Windows Server 2019 1809"; break }
        default                                                      { "Unknown Windows Version"; break }
    }

    if ($sku) {
        return "$versionName (OS Build $actualBuild.$actualMinor) [$sku]"
    } elseif ($productName) {
        return "$versionName (OS Build $actualBuild.$actualMinor) [$productName]"
    } else {
        return "$versionName (OS Build $actualBuild.$actualMinor)"
    }
}

# Extract available peripheral devices
function Get-DeviceLabels {
    param ($deviceList)
    return ($deviceList | ForEach-Object { $_.label }) -join ', '
}

# Output
Write-Host "`n=== VDI Monitoring Informationen ===" -ForegroundColor Cyan
Write-Host "Timestamp                : $timestampReadable"
Write-Host "VM OS Version            : $(Get-WindowsVersionName $vdiVersion.vmVersion)"
Write-Host "Connected Stack          : $($vdiState.connectedStack)"

switch ($vdiState.vdiMode) {
    # 1122 seems WebRTC; 5100 seems AVD Media Optimized; 5200 seems Slimcore
    "1122" {
        Write-Host "`nVDI Optimization         : WebRTC Optimized"
    }
    "5100" {
        Write-Host "`nVDI Optimization         : AVD Media Optimized"
    }
    "5200" {
        Write-Host "`nVDI Optimization         : SlimCore Optimized"
        # The following values are only available in SlimCore mode
        Write-Host "`nSlimCore Version         : $($vdiVersion.remoteSlimcoreVersion)"
        Write-Host "VDIBridge Version        : $($vdiVersion.bridgeVersion)"
        Write-Host "MS Teams Plugin Version  : $($vdiVersion.pluginVersion)"
        Write-Host "`nTeams Version            : $($vdiVersion.teamsVersion)"
        #Write-Host "`nClient Platform          : $($vdiVersion.clientPlatform)"
        Write-Host "`nRD Client                : $($vdiVersion.rdClientProductName)"
        Write-Host "RD Client Version        : $($vdiVersion.rdClientVersion)"

        Write-Host "`n--- Available peripheral devices ---" -ForegroundColor Cyan
        Write-Host "Speakers     : $(Get-DeviceLabels $devices.speaker.available)"
        Write-Host "Cameras      : $(Get-DeviceLabels $devices.camera.available)"
        Write-Host "Microphones  : $(Get-DeviceLabels $devices.microphone.available)"

        Write-Host "`n--- Selected peripheral devices ---" -ForegroundColor Cyan
        Write-Host "Speaker      : $($devices.speaker.selected)"
        Write-Host "Camera       : $($devices.camera.selected)"
        Write-Host "Microphone   : $($devices.microphone.selected)"

        Write-Host "`nSecondary Ringtone : $($devices.secondaryRinger)"
    }
    default {
        Write-Host "VDI Optimization         : Unknown Mode ($($vdiState.vdiMode))"
    }
}
