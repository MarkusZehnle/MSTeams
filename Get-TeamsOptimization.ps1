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
    Version:        1.1
    Author:         Markus Zehnle (ChatGPT tbh)
    Created:        2025-05-19
    Updated:        2025-05-23
    Requires:       PowerShell 5.1 or later
    JSON Source:    C:\Users\<username>\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\tfw\vdi_connection_info.json

.VERSION HISTORY
    1.0     2025-05-19    Initial version
    1.1     2025-05-23    Added detailed vdiMode parsing (Thanks Fernando K. and Kenny W. from Microsoft)

.LINK
    Microsoft Teams VDI Monitoring API documentation
    https://learn.microsoft.com/en-us/microsoftteams/vdi-2#monitoring-api
#>

$tfwPath = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\tfw\vdi_connection_info.json"
if (-Not (Test-Path -Path $tfwPath)) {
    Write-Error "The file vdi_connection_info.json cannot be found: $tfwPath"
    exit 1
}

$jsonContent = Get-Content -Path $tfwPath -Raw | ConvertFrom-Json
$session = $jsonContent.VdiConnectionInfo | Select-Object -Last 1
$vdiState = $session.vdiConnectedState
$vdiVersion = $session.vdiVersionInfo
$devices = $session.devices
$timestampReadable = (Get-Date -Date "1970-01-01 00:00:00Z").AddMilliseconds($vdiState.timestamp).ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss")

function Get-WindowsVersionName {
    try {
        $osInfo = Get-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"
        $actualBuild = $osInfo.CurrentBuildNumber
        $actualMinor = $osInfo.UBR
        $sku = $osInfo.EditionID
        $productName = $osInfo.ProductName
    } catch {
        return "Unknown Windows Version"
    }

    $versionName = switch ($true) {
        ($actualBuild -eq 26100 -and $productName -like "*Server*") { "Windows Server 2025" }
        ($actualBuild -eq 26100) { "Windows 11 24H2" }
        ($actualBuild -eq 25300) { "Windows 11 Insider Dev" }
        ($actualBuild -eq 22631) { "Windows 11 23H2" }
        ($actualBuild -eq 22621) { "Windows 11 22H2" }
        ($actualBuild -eq 22000) { "Windows 11 21H2" }
        ($actualBuild -eq 20348) { "Windows Server 2022" }
        ($actualBuild -eq 19045) { "Windows 10 22H2" }
        ($actualBuild -eq 19044) { "Windows 10 21H2" }
        ($actualBuild -eq 17763) { "Windows Server 2019" }
        default { "Unknown Windows Version" }
    }

    return "$versionName (OS Build $actualBuild.$actualMinor) [$sku]"
}

function Get-DeviceLabels {
    param ($deviceList)
    return ($deviceList | ForEach-Object { $_.label }) -join ', '
}

function Get-VdiModeDetails {
    param ([string]$vdiMode)
    if (-not $vdiMode -or $vdiMode.Length -lt 4) { return "Unknown (Invalid Code)" }

    $platformCode = $vdiMode.Substring(0, 1)
    $optimizationCode = $vdiMode.Substring(1, 1)

    $platform = switch ($platformCode) {
        "1" { "Citrix" }
        "2" { "Citrix" }
        "3" { "Horizon" }
        "5" { "AVD/W365" }
        default { "Unknown Platform" }
    }

    $optimization = switch ($optimizationCode) {
        "0" { "not optimized" }
        "1" { "WebRTC optimized" }
        "2" { "SlimCore optimized" }
        default { "Unknown Optimization" }
    }

    return "$platform $optimization ($vdiMode)"
}

Write-Host "`n=== Microsoft Teams VDI Monitoring ===" -ForegroundColor Cyan
Write-Host ("{0,-25}: {1}" -f "Session Timestamp", $timestampReadable)
Write-Host ("{0,-25}: {1}" -f "VM OS Version", (Get-WindowsVersionName))
#Write-Host ("{0,-25}: {1}" -f "Connected Stack", $vdiState.connectedStack)
Write-Host ("{0,-25}: {1}" -f "VDI Mode Details", (Get-VdiModeDetails $vdiState.vdiMode))

# SlimCore-specific section with correct condition
if ($vdiState.vdiMode.Substring(1,1) -eq "2" -and $vdiState.connectedStack -eq "remote") {
    Write-Host "`n--- SlimCore Optimization Information ---" -ForegroundColor Cyan
    Write-Host ("{0,-25}: {1}" -f "SlimCore Version", $vdiVersion.remoteSlimcoreVersion)
    Write-Host ("{0,-25}: {1}" -f "VDIBridge Version", $vdiVersion.bridgeVersion)
    Write-Host ("{0,-25}: {1}" -f "Teams Plugin Version", $vdiVersion.pluginVersion)
    Write-Host ("{0,-25}: {1}" -f "Teams Version", $vdiVersion.teamsVersion)

    Write-Host "`n--- Remote Client Information ---" -ForegroundColor Cyan
    Write-Host ("{0,-25}: {1}" -f "Remote Client", $vdiVersion.rdClientProductName)
    Write-Host ("{0,-25}: {1}" -f "Remote Client Version", $vdiVersion.rdClientVersion)

    Write-Host "`n--- Available Peripheral Devices ---" -ForegroundColor Cyan
    Write-Host ("{0,-25}: {1}" -f "Speakers", (Get-DeviceLabels $devices.speaker.available))
    Write-Host ("{0,-25}: {1}" -f "Cameras", (Get-DeviceLabels $devices.camera.available))
    Write-Host ("{0,-25}: {1}" -f "Microphones", (Get-DeviceLabels $devices.microphone.available))

    Write-Host "`n--- Selected Peripheral Devices ---" -ForegroundColor Cyan
    Write-Host ("{0,-25}: {1}" -f "Speaker", $devices.speaker.selected)
    Write-Host ("{0,-25}: {1}" -f "Camera", $devices.camera.selected)
    Write-Host ("{0,-25}: {1}" -f "Microphone", $devices.microphone.selected)
    Write-Host ("{0,-25}: {1}" -f "Secondary Ringtone", $devices.secondaryRinger)
}
elseif ($vdiState.vdiMode.Substring(1,1) -eq "2") {
    Write-Host "`n--- SlimCore Optimization Information ---" -ForegroundColor Cyan
    Write-Host ("{0,-25}: {1}" -f "SlimCore Optimization", "Failed to load SlimCore optimization. Please restart Teams.") -ForegroundColor Red
}
