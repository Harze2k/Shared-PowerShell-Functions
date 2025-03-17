<#
.DESCRIPTION
This PowerShell script is designed to update Plex HTPC and its MPV (libmpv) component on Windows systems. It checks for updates, downloads them if available, and installs them while handling the process of stopping and restarting the Plex HTPC application.
The script requires administrator privileges to run successfully.
It uses the New-Log function for logging, which should be defined or imported separately.
The script relies on NanaZip for extracting files, and will attempt to install it using winget if not present.
The script backs up the current MPV client before updating, unless the -Force parameter is used and the backup fails.
The script handles stopping and restarting the Plex HTPC process during updates.
Main Functions:
Get-PlexHTPCFileVersion: Retrieves the current version of Plex HTPC installed on the system.
Check-PlexHTPCUpdate: Checks if there's an update available for Plex HTPC.
Compare-MPVVersion: Compares MPV versions.
Get-MPVVersion: Retrieves the current and latest versions of MPV.
Check-MPVUpdate: Checks if there's an update available for MPV.
Stop-RestartPlexHTPC: Stops or restarts the Plex HTPC process.
Update-PlexHTPC: Performs the update process for Plex HTPC.
Update-MPV: Performs the update process for MPV.
Ensure-NanaZipInstalled: Ensures that NanaZip is installed, which is required for extracting update files.
.PARAMETER CompareVersion
CompareVersion (switch): Used in Check-MPVUpdate to download and compare MPV versions based on file version instead of release date.
.PARAMETER KeepCurrentVersion
KeepCurrentVersion (switch): Used in Check-MPVUpdate to skip the update and keep the current MPV version.
.PARAMETER Restart
Restart (switch): Used in Stop-RestartPlexHTPC to restart Plex HTPC after stopping it.
.PARAMETER Force
Force (switch): Used in Update-MPV to continue the update process even if backing up the current MPV client fails.
.OUTPUTS
Example output:
[2024-08-22 21:51:41.925][SUCCESS] NanaZip version [3.1.1080.0] is already installed. [Function: Ensure-NanaZipInstalled]
[2024-08-22 21:51:43.332][SUCCESS] Successfully downloaded the latest mpv client to compare against. [Function: Get-MPVVersion]
[2024-08-22 21:51:44.130][SUCCESS] Successfully extraced the mpv update... [Function: Get-MPVVersion]
[2024-08-22 21:51:44.163][SUCCESS] Local mpv version: [v0.38.0_4-243-g718fd0d8] is up to date. [Function: Check-MPVUpdate]
[2024-08-22 21:51:44.164][INFO] You are using a custom modified PlexHTPC mpv and not the standard PlexHTPC mpv client. [Function: Check-MPVUpdate]
[2024-08-22 21:51:44.193][INFO] No MPV update required.
[2024-08-22 21:51:44.347][SUCCESS] Successfully got data from plex.tv. [Function: Check-PlexHTPCUpdate]
[2024-08-22 21:51:44.394][SUCCESS] Plex HTPC is up-to-date. [Function: Check-PlexHTPCUpdate]
[2024-08-22 21:51:44.395][SUCCESS] Current Plex HTPC version: [1.66.1.215] same as latest online version: [1.66.1.215] [Function: Check-PlexHTPCUpdate]
[2024-08-22 21:51:44.407][INFO] No Plex HTPC update required.
[2024-08-22 21:51:44.411][SUCCESS] Update process completed.
#>
function Get-PlexHTPCFileVersion {
    [CmdletBinding()]
    param (
        [string]$LatestOnlineVersion
    )
    $plexPath = Join-Path -Path $localInstallPath -ChildPath 'Plex HTPC.exe'
    if (Test-Path $plexPath) {
        try {
            $content = [System.IO.File]::ReadAllBytes($plexPath)
            $text = [System.Text.Encoding]::ASCII.GetString($content)
            $cleanText = $text -replace "`0", " "
            $versionPattern = "Plex HTPC\s+(\d+\.\d+\.\d+\.\d+)-[a-f0-9]+"
            if ($cleanText -match $versionPattern) {
                $foundVersion = $matches[1]
                $updateNeeded = $foundVersion -ne $LatestOnlineVersion
                return [PSCustomObject]@{
                    UpdateNeeded = $updateNeeded
                    LocalVersion = $foundVersion
                }
            }
            else {
                return [PSCustomObject]@{
                    UpdateNeeded = $true
                    LocalVersion = $null
                }
            }
        }
        catch {
            New-Log "An error occurred while processing the file." -Level ERROR
            return [PSCustomObject]@{
                UpdateNeeded = $true
                LocalVersion = $null
            }
        }
    }
    else {
        New-Log "Plex HTPC executable not found at the specified path." -Level WARNING
        return [PSCustomObject]@{
            UpdateNeeded = $true
            LocalVersion = $null
        }
    }
}
function Check-PlexHTPCUpdate {
    $plexExePath = Join-Path $localInstallPath -ChildPath 'Plex HTPC.exe'
    if (-not (Test-Path $plexExePath)) {
        New-Log "Plex HTPC not installed. Install manually." -Level WARNING
        return $null
    }
    try {
        $json = Invoke-RestMethod -Uri "https://plex.tv/api/downloads/7.json"
        $download = $json.computer.Windows.releases.url
        New-Log "Successfully got data from plex.tv." -Level SUCCESS
    }
    catch {
        New-Log "Could not fetching data from plex.tv." -Level ERROR
        return $null
    }
    $downloadVersion = $json.computer.Windows.version -replace '-.*$'
    $localVersion = Get-PlexHTPCFileVersion -LatestOnlineVersion $downloadVersion
    if ($localVersion.UpdateNeeded -eq $false) {
        New-Log "Plex HTPC is up-to-date." -Level SUCCESS
        New-Log "Current Plex HTPC version: [$($localVersion.LocalVersion)] same as latest online version: [$downloadVersion]" -Level SUCCESS
        return $null
    }
    else {
        New-Log "Plex HTPC is NOT up-to-date."
        New-Log "Current Plex HTPC version: [$($localVersion.LocalVersion)] is NOT same as latest online version: [$downloadVersion]. Will update!" -Level SUCCESS
        return @{
            Download = $download
            Version  = $downloadVersion
        }
    }
}
function Compare-MPVVersion {
    [CmdletBinding()]
    param (
        [string]$Version
    )
    $versionPattern = '(?<minor>\d+)\.(?<patch>\d+)[_-](?<build>\d+)[_-](?<revision>\d+)'
    if ($Version -match $versionPattern) {
        $minor = $matches['minor']
        $patch = $matches['patch']
        $build = $matches['build']
        $revision = $matches['revision']
        $patch = if ($patch) {
            $patch
        }
        else {
            '0'
        }
        $build = if ($build) {
            $build
        }
        else {
            '0'
        }
        $revision = if ($revision) {
            $revision
        }
        else {
            '0'
        }
        $customVersionString = "$minor.$patch.$build.$revision"
        return $customVersionString
    }
    if ($Version -match '-UNKNOWN') {
        return ($version -split '-')[0]
    }
    return $null
}
function Get-MPVVersion {
    [CmdletBinding()]
    param (
        [switch]$CompareVersion
    )
    try {
        $mpvOnlineResponse = Invoke-RestMethod -Uri "https://api.github.com/repos/mitzsch/mpv-winbuild/releases/latest"
        $mpvDownloadUrl = ($mpvOnlineResponse.assets | Where-Object name -Like "mpv-dev-x86_64-v3-*").browser_download_url
    }
    catch {
        New-Log "Could not get MPV version info from github." -Level ERROR
        return @{
            NeedUpdate        = $false
            Version           = $null
            VersionDownloaded = $false
            URL               = $null
        }
    }
    if ($CompareVersion.IsPresent) {
        New-Item -Path $tempPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        try {
            $httpclient = [System.Net.Http.HttpClient]::new()
            $response = $httpclient.GetAsync($mpvDownloadUrl).Result
            [System.IO.File]::WriteAllBytes($tempPath + "\$(($mpvOnlineResponse.assets | Where-Object name -Like "mpv-dev-x86_64-v3-*").name)", $response.Content.ReadAsByteArrayAsync().Result)
            $httpclient.Dispose()
            New-Log "Successfully downloaded the latest mpv client to compare against." -Level SUCCESS
            & $zipTool x (Join-Path $tempPath ([System.IO.Path]::GetFileName($mpvDownloadUrl))) ("-o" + (Join-Path $tempPath "lib\")) -y | Out-Null
            New-Log "Successfully extraced the mpv update..." -Level SUCCESS
        }
        catch {
            New-Log "Something went wrong when downloading or extracting the latest mpv client." -Level ERROR
            return @{
                NeedUpdate        = $false
                Version           = $null
                VersionDownloaded = $false
                URL               = $null
            }
        }
        try {
            $online = (Get-Item -Path "$tempPath\lib\libmpv-2.dll").VersionInfo.FileVersion
            $local = (Get-Item (Join-Path $localInstallPath "libmpv-2.dll")).VersionInfo.FileVersion
            [version]$localVersion = Compare-MPVVersion -Version $local
            [version]$onlineVersion = Compare-MPVVersion -Version $online
            if ($onlineVersion -and $localVersion -and $onlineVersion -gt $localVersion) {
                return @{
                    NeedUpdate        = $true
                    Version           = $onlineVersion
                    VersionDownloaded = $true
                    URL               = $mpvDownloadUrl
                }
            }
            else {
                return @{
                    NeedUpdate        = $false
                    Version           = $onlineVersion
                    VersionDownloaded = $true
                    URL               = $mpvDownloadUrl
                }
            }
        }
        catch {
            New-Log "Could not compare versions mpv version." -Level ERROR
            return @{
                NeedUpdate        = $false
                Version           = $null
                VersionDownloaded = $false
                URL               = $null
            }
        }
    }
    try {
        [datetime]$localVersionDate = (Get-Item (Join-Path $localInstallPath "libmpv-2.dll")).LastWriteTime.ToString("yyyy-MM-dd")
        [datetime]$onlineVersionDate = ($mpvOnlineResponse.tag_name).Substring(0, 10)
        if ($onlineversiondate -and $localVersionDate -and $onlineVersionDate -gt $localVersionDate) {
            return @{
                NeedUpdate        = $true
                Version           = $onlineVersionDate
                VersionDownloaded = $false
                URL               = $mpvDownloadUrl
            }
        }
        else {
            return @{
                NeedUpdate        = $false
                Version           = $null
                VersionDownloaded = $false
                URL               = $mpvDownloadUrl
            }
        }
    }
    catch {
        New-Log "Could not compare versions mpv version." -Level ERROR
        return @{
            NeedUpdate        = $false
            Version           = $null
            VersionDownloaded = $false
            URL               = $null
        }
    }
}
function Check-MPVUpdate {
    [CmdletBinding()]
    param(
        [switch]$CompareVersion,
        [switch]$KeepCurrentVersion
    )
    if ($KeepCurrentVersion.IsPresent) {
        New-Log "Keeping current MPV version."
        New-Log "Current local mpv version is: [$((Get-Item (Join-Path $localInstallPath "libmpv-2.dll")).VersionInfo.FileVersion)]"
        return
    }
    $needUpdate = if ($CompareVersion.IsPresent) {
        Get-MPVVersion -CompareVersion
    }
    else {
        Get-MPVVersion
    }
    if ($needUpdate.NeedUpdate -eq $true -and $needUpdate.URL ) {
        New-Log "New MPV version available:[v0.$($needUpdate.Version)]" -Level SUCCESS
        return @{
            Download          = $needUpdate.URL
            Version           = $needUpdate.Version
            VersionDownloaded = $needUpdate.VersionDownloaded
        }
    }
    else {
        New-Log "Local mpv version: [$((Get-Item (Join-Path $localInstallPath "libmpv-2.dll")).VersionInfo.FileVersion)] is up to date." -Level SUCCESS
    }
    if ((Get-Item (Join-Path $localInstallPath "libmpv-2.dll")).Length -lt 50000000) {
        New-Log "You are using the original PlexHTPC mpv client."
    }
    else {
        New-Log "You are using a custom modified PlexHTPC mpv and not the standard PlexHTPC mpv client."
    }
}
function Stop-RestartPlexHTPC {
    [CmdletBinding()]
    param(
        [switch]$Restart
    )
    if ($Restart.IsPresent) {
        New-Log "Restarting Plex HTPC..."
        try {
            Start-Process -FilePath (Join-Path $localInstallPath "Plex HTPC.exe") -WindowStyle Hidden -ErrorAction Stop
            New-Log "Plex HTPC successfully restarted!" -Level SUCCESS
        }
        catch {
            New-Log "Something went wrong when restarting Plex HTPC." -Level ERROR
        }
        return
    }
    $plexProcess = Get-Process "Plex HTPC" -ErrorAction SilentlyContinue
    if ($plexProcess) {
        New-Log "Stopping Plex HTPC process..."
        try {
            $plexProcess | Stop-Process -Force -Confirm:$false -ErrorAction Stop
            New-Log "Successfully stopped Plex HTPC before the update." -Level SUCCESS
        }
        catch {
            New-Log "Something went wrong when stopping Plex HTPC." -Level ERROR
        }
    }
    else {
        New-Log "Plex HTPC is not running. Nothing to stop."
    }
}
function Update-PlexHTPC {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$Download,
        [Parameter(Mandatory)][string]$Version
    )
    try {
        New-Log "Starting Plex HTPC update process..."
        New-Item -Path $tempPath -ItemType Directory -Force | Out-Null
        $plexFilename = Split-Path $Download -Leaf
        New-Log "Downloading Plex HTPC update: $plexFilename"
        try {
            $httpclient = [System.Net.Http.HttpClient]::new()
            $response = $httpclient.GetAsync($Download).Result
            [System.IO.File]::WriteAllBytes((Join-Path $tempPath $plexFilename), $response.Content.ReadAsByteArrayAsync().Result)
            $httpclient.Dispose()
            New-Log "Successfully downloaded the new Plex HTPC client." -Level SUCCESS
            & $zipTool x (Join-Path $tempPath $plexFilename) ("-o" + (Join-Path $tempPath "app\")) -y | Out-Null
            New-Log "Successfully extraced the Plex HTPC update..." -Level SUCCESS
        }
        catch {
            New-Log "Something went wrong when downloading or extracting the new Plex HTPC client. Will abort the update." -Level ERROR
            return
        }
        Stop-RestartPlexHTPC
        New-Log "Copying updated Plex HTPC files..."
        try {
            Copy-Item -Path (Join-Path $tempPath "app\*") -Destination $localInstallPath -Recurse -Exclude '$PLUGINSDIR', '$TEMP', "*.nsi", "*.nsis" -Force -ErrorAction Stop
            New-Log "Plex HTPC update completed successfully! New version: [$Version]" -Level SUCCESS
        }
        catch {
            New-Log "Something went wrong during Plex HTPC update. Will abort the update." -Level ERROR
            return
        }
        Stop-RestartPlexHTPC -Restart
    }
    catch {
        New-Log "An error occurred during the Plex HTPC update process." -Level ERROR
    }
    finally {
        Remove-Item $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}
function Update-MPV {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$Download,
        [Parameter(Mandatory)][string]$Version,
        [switch]$VersionDownloaded,
        [switch]$Force
    )
    try {
        if ($VersionDownloaded.IsPresent) {
            try {
                New-Log "Starting MPV update process... Latest version is already downloaded. Copying new libmpv-2.dll to $localInstallPath..."
                Copy-Item -Path (Join-Path $tempPath "lib\libmpv-2.dll") -Destination $localInstallPath -Force -Recurse -Confirm:$false -ErrorAction Stop | Out-Null
                New-Log "MPV update completed successfully! New version:[v0.$Version]" -Level SUCCESS
            }
            catch {
                New-Log "Something went wrong during MPV update." -Level ERROR
            }
            return
        }
        New-Log "Starting MPV update process..."
        New-Item -Path $tempPath -ItemType Directory -Force | Out-Null
        $mpvFilename = [System.IO.Path]::GetFileName($Download)
        New-Log "Downloading MPV update: $mpvFilename"
        try {
            $httpclient = [System.Net.Http.HttpClient]::new()
            $response = $httpclient.GetAsync($Download).Result
            [System.IO.File]::WriteAllBytes((Join-Path $tempPath $mpvFilename), $response.Content.ReadAsByteArrayAsync().Result)
            $httpclient.Dispose()
            New-Log "Successfully downloaded the new mpv client." -Level SUCCESS
            & $zipTool x (Join-Path $tempPath $mpvFilename) ("-o" + (Join-Path $tempPath "lib\")) -y | Out-Null
            New-Log "Successfully extraced the mpv update..." -Level SUCCESS
        }
        catch {
            New-Log "Something went wrong when downloading or extracting the new mpv client. Will abort the update." -Level ERROR
            return
        }
        Stop-RestartPlexHTPC
        New-Log "Copying updated MPV files..."
        New-Item -Path $backupDir -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        try {
            Copy-Item -Path ("$localInstallPath" + "libmpv-2.dll") -Destination $backupDir -Force -Confirm:$false -ErrorAction Stop | Out-Null
            New-Log "Successfully backed up the current mpv client to: [$backupDir\libmpv-2.dll]." -Level SUCCESS
        }
        catch {
            if($Force.IsPresent) {
                New-Log "Something went wrong when trying to backup the current mpv client. Will contune since -Force was specified." -Level ERROR
            }
            else {
                New-Log "Something went wrong when trying to backup the current mpv client. Will abort since -Force was not specified." -Level ERROR
                return
            }
        }
        try {
            Copy-Item -Path (Join-Path $tempPath "lib\libmpv-2.dll") -Destination $localInstallPath -Force -Recurse -Confirm:$false -ErrorAction Stop | Out-Null
            New-Log "MPV update completed successfully! New version: [v0.$Version]" -Level SUCCESS
        }
        catch {
            New-Log "Something went wrong during MPV update. Will abort the update." -Level ERROR
            return
        }
        Stop-RestartPlexHTPC -Restart
        Remove-Item -Path (Join-Path $tempPath $mpvFilename) -Force -ErrorAction SilentlyContinue
    }
    catch {
        New-Log "An error occurred during the MPV update process. Will abort the update." -Level ERROR
    }
    finally {
        Remove-Item $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}
function Ensure-NanaZipInstalled {
    $nanaZipInstalled = Get-Command nanazip
    if ($nanaZipInstalled) {
        New-Log "NanaZip is already installed." -Level SUCCESS
        return $true
    }
    $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $wingetPath) {
        New-Log "Winget is not installed. Please install the Windows Package Manager (winget) to proceed." -Level WARNING
        return $false
    }
    New-Log "Attempting to install NanaZip..."
    try {
        Start-Process -FilePath $wingetPath.Source -ArgumentList "install -e --force --id M2Team.NanaZip --accept-package-agreements --accept-source-agreements --silent" -Wait -NoNewWindow -ErrorAction Stop
        New-Log "NanaZip version has been successfully installed. Restart powershell so its found." -Level SUCCESS
        Exit 0
    }
    catch {
        New-Log "An error occurred while trying to install NanaZip." -Level ERROR
        return $false
    }
}
# Main script & Invoke Log Function
### OBS: New-Log Function is neede otherwise remove all New-Log and replace with Write-Host. New-Log is vastly better though, check the link below:
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
# Global variables
$localInstallPath = "C:\Program Files\Plex\Plex HTPC\"
$backupDir = "$localInstallPath" + "Backup"
$tempPath = "$env:TEMP\plexhtpc\"
$zipTool = "nanazipc"
#Start update check
if (!(Ensure-NanaZipInstalled)) {
    New-Log "Failed to make sure nanazip is installed and usable. The script will not work without it."
    Exit 1
}
$mpvUpdateInfo = Check-MPVUpdate -CompareVersion #With -CompareVersion it will download the latest mpv client and check the dll version to compare with. Without -CompareVersion it will just compare the date of the local dll with latest on Github.
if ($mpvUpdateInfo) {
    New-Log "Updating MPV..."
    Update-MPV -Download $mpvUpdateInfo.Download -Version $mpvUpdateInfo.Version -VersionDownloaded:$mpvUpdateInfo.VersionDownloaded -Force #-Force is used to ignore if the script is unabled to backup the current mvp client.
} else {
    New-Log "No MPV update required."
}
$plexUpdateInfo = Check-PlexHTPCUpdate
if ($plexUpdateInfo) {
    New-Log "Updating Plex HTPC..."
    Update-PlexHTPC -Download $plexUpdateInfo.Download -Version $plexUpdateInfo.Version
} else {
    New-Log "No Plex HTPC update required."
}
New-Log "Update process completed." -level SUCCESS