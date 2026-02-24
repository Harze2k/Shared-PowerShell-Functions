#Requires -Version 5.1
<#
.SYNOPSIS
    Updates a portable VS Code Insiders installation.
.DESCRIPTION
    Checks for the latest VS Code Insiders build via the update API,
    downloads the Windows x64 zip archive, and extracts it to the
    portable installation folder. The 'data' folder (user settings
    and extensions) is preserved across updates.
.PARAMETER InstallPath
    The path to the portable VS Code Insiders installation folder.
    Defaults to 'C:\Temp\VSCode-Insider'.
.PARAMETER LogFilePath
    Optional path for a log file. When specified, all output is also
    written to disk via New-Log.
.EXAMPLE
    .\Update-VSCodeInsiders.ps1
    Checks for updates and installs the latest VS Code Insiders build.
.EXAMPLE
    .\Update-VSCodeInsiders.ps1 -InstallPath 'D:\Tools\VSCode-Insider'
    Updates a portable installation at a custom path.
.EXAMPLE
    .\Update-VSCodeInsiders.ps1 -LogFilePath 'C:\Logs\vscode-update.log'
    Updates and writes a log file to disk.
.NOTES
    Author: Martin
    Requires: New-Log (https://github.com/Harze2k/Shared-PowerShell-Functions)
    The portable installation must already exist at the specified path.
.LINK
    https://github.com/Harze2k/Shared-PowerShell-Functions
#>
[CmdletBinding()]
param(
    [Parameter()][ValidateNotNullOrEmpty()][string]$InstallPath = 'C:\Temp\VSCode-Insider',
    [Parameter()][string]$LogFilePath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$LogScriptUrl = "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/New-Log.ps1"
#region Dependencies
if (-not (Get-Command "New-Log" -ErrorAction SilentlyContinue) -and $LogScriptUrl) {
    $LogScriptPath = Join-Path $env:TEMP "New-Log.ps1"
    if (-not (Test-Path $LogScriptPath)) {
        try {
            Invoke-WebRequest -Uri $LogScriptUrl -OutFile $LogScriptPath -UseBasicParsing -MaximumRedirection 1 -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to download New-Log.ps1. Proceeding with basic fallback logger."
        }
    }
    if (Test-Path $LogScriptPath) {
        . $LogScriptPath
    }
}
if (-not (Get-Command "New-Log" -ErrorAction SilentlyContinue)) {
    function New-Log {
        [CmdletBinding()]
        param(
            [Parameter(ValueFromPipeline, Position = 0)]$Message,
            [Parameter(Position = 1)][string]$Level = "INFO"
        )
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        $color = switch ($Level) {
            "ERROR" { 'Red' }
            "WARNING" { 'Yellow' }
            "SUCCESS" { 'Green' }
            "DEBUG" { 'Blue' }
            "VERBOSE" { 'Cyan' }
            default { 'White' }
        }
        Write-Host "[$timestamp][$Level] $Message" -ForegroundColor $color
    }
}
$logParams = @{}
if ($LogFilePath) { $logParams['LogFilePath'] = $LogFilePath }
#endregion Dependencies
#region Functions
function Test-VSCodeInsidersRunning {
    <#
    .SYNOPSIS
        Checks if VS Code Insiders is currently running.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    $process = Get-Process -Name 'Code - Insiders' -ErrorAction SilentlyContinue
    return ($null -ne $process)
}
function Get-InstalledVersion {
    <#
    .SYNOPSIS
        Gets the currently installed version and commit hash.
    .DESCRIPTION
        Reads version info from product.json and cross-checks against the exe's
        ProductVersion. The exe version is authoritative because VS Code's own
        auto-updater can leave product.json stale (files in a hash-subfolder).
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)][string]$InstallPath
    )
    $jsonVersion = $null
    $jsonCommit = $null
    $hashFolder = Remove-HashFolders -GetHashFolder -InstallPath $InstallPath
    $productJsonPath = Join-Path -Path "$InstallPath\$hashFolder" -ChildPath 'resources\app\product.json'
    if (Test-Path -Path $productJsonPath -PathType Leaf) {
        $productJson = Get-Content -Path $productJsonPath -Raw -ErrorAction Stop | ConvertFrom-Json
        $jsonVersion = $productJson.version
        $jsonCommit = $productJson.commit
    }
    $exeVersion = $null
    $exePath = Join-Path -Path $InstallPath -ChildPath 'Code - Insiders.exe'
    if (Test-Path -Path $exePath -PathType Leaf) {
        $exeVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exePath).ProductVersion
    }
    if (-not $jsonVersion -and -not $exeVersion) {
        New-Log "No installed version found at '$InstallPath'. Treating as fresh install." -Level WARNING @logParams
        return [PSCustomObject]@{
            Version    = $null
            Commit     = $null
            ExeVersion = $null
        }
    }
    if ($exeVersion -and $jsonVersion -and $exeVersion -ne $jsonVersion) {
        New-Log "product.json version ($jsonVersion) differs from exe version ($exeVersion). On-disk files need refresh." -Level WARNING @logParams
    }
    return [PSCustomObject]@{
        Version    = $jsonVersion
        Commit     = $jsonCommit
        ExeVersion = $exeVersion
    }
}
function Get-LatestInsiderBuild {
    <#
    .SYNOPSIS
        Queries the VS Code update API for the latest Insider build.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()
    $apiUrl = 'https://update.code.visualstudio.com/api/update/win32-x64-archive/insider/latest'
    New-Log 'Querying VS Code update API for the latest Insider build...' @logParams
    $response = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing -ErrorAction Stop
    return [PSCustomObject]@{
        Version     = $response.productVersion
        Commit      = $response.version
        DownloadUrl = $response.url
        Sha256Hash  = $response.sha256hash
    }
}
function Remove-HashFolders {
    <#
    .SYNOPSIS
        Removes old commit-hash folders from the VS Code portable installation directory.
    .DESCRIPTION
        Identifies and removes folders whose names look like short alphanumeric hashes
        (commit ID fragments). Preserves all known VS Code folders and the data folder.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$InstallPath,
        [Parameter()][switch]$GetHashFolder
    )
    # Known VS Code folders that must never be removed
    $knownFolders = @(
        'appx'
        'bin'
        'data'
        'locales'
        'policies'
        'resources'
        'tools'
    )
    $directories = Get-ChildItem -Path $InstallPath -Directory -ErrorAction Stop
    foreach ($dir in $directories) {
        $name = $dir.Name
        if ($knownFolders -contains $name) {
            continue
        }
        # Match folder names that look like hex commit hashes (8+ hex characters, no spaces)
        if ($name -match '^[0-9a-f]{8,}$') {
            if ($GetHashFolder.IsPresent) {
                return $name
            }
            if ($PSCmdlet.ShouldProcess($dir.FullName, 'Remove hash folder')) {
                New-Log "Removing old hash folder: $name" -Level WARNING @logParams
                Remove-Item -Path $dir.FullName -Recurse -Force -ErrorAction Stop
            }
        }
    }
}
function Get-VSCodeInsidersZip {
    <#
    .SYNOPSIS
        Downloads the VS Code Insiders zip archive.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$DownloadUrl,
        [Parameter(Mandatory)][string]$Sha256Hash
    )
    $tempFile = Join-Path -Path $env:TEMP -ChildPath "VSCode-Insiders-$(Get-Date -Format 'yyyyMMdd-HHmmss').zip"
    New-Log 'Downloading VS Code Insiders zip...' @logParams
    New-Log "URL: $DownloadUrl" -Level VERBOSE @logParams
    New-Log "Destination: $tempFile" -Level VERBOSE @logParams
    try {
        Import-Module BitsTransfer -ErrorAction Stop
        Start-BitsTransfer -Source $DownloadUrl -Destination $tempFile -ErrorAction Stop
    }
    catch {
        New-Log 'BITS transfer failed, falling back to Invoke-WebRequest...' -Level WARNING @logParams
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $tempFile -UseBasicParsing -ErrorAction Stop
    }
    New-Log 'Verifying SHA256 hash...' @logParams
    $fileHash = (Get-FileHash -Path $tempFile -Algorithm SHA256 -ErrorAction Stop).Hash
    if ($fileHash -ne $Sha256Hash) {
        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        throw "Hash mismatch! Expected '$Sha256Hash' but got '$fileHash'. The download may be corrupted."
    }
    New-Log 'Hash verified successfully.' -Level SUCCESS @logParams
    return $tempFile
}
function Expand-VSCodeInsidersZip {
    <#
    .SYNOPSIS
        Extracts the VS Code Insiders zip to the portable installation folder.
    .DESCRIPTION
        The zip contains a top-level folder (e.g. VSCode-win32-x64-1.110.0-insider).
        This function extracts to a temp directory first, then copies the inner
        contents directly to the installation path, avoiding nested subfolder issues.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [Parameter(Mandatory)][string]$InstallPath
    )
    $tempExtractPath = Join-Path -Path $env:TEMP -ChildPath "VSCode-Insiders-Extract-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    try {
        New-Log 'Extracting zip to temporary location...' @logParams
        Expand-Archive -Path $ZipPath -DestinationPath $tempExtractPath -Force -ErrorAction Stop
        $extractedItems = Get-ChildItem -Path $tempExtractPath -ErrorAction Stop
        $sourcePath = $tempExtractPath
        if ($extractedItems.Count -eq 1 -and $extractedItems[0].PSIsContainer) {
            $sourcePath = $extractedItems[0].FullName
            New-Log "Detected top-level folder: $($extractedItems[0].Name)" -Level VERBOSE @logParams
        }
        New-Log "Copying files to '$InstallPath'..." @logParams
        $sourceItems = Get-ChildItem -Path $sourcePath -ErrorAction Stop
        foreach ($item in $sourceItems) {
            if ($item.Name -eq 'data') {
                New-Log "Skipping 'data' folder (preserving user settings)." -Level WARNING @logParams
                continue
            }
            $destinationItem = Join-Path -Path $InstallPath -ChildPath $item.Name
            if (Test-Path -Path $destinationItem) {
                Remove-Item -Path $destinationItem -Recurse -Force -ErrorAction Stop
            }
            Copy-Item -Path $item.FullName -Destination $destinationItem -Recurse -Force -ErrorAction Stop
        }
        New-Log 'Files copied successfully.' -Level SUCCESS @logParams
    }
    finally {
        if (Test-Path -Path $tempExtractPath) {
            Remove-Item -Path $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
#endregion Functions
#region Main
try {
    New-Log '=== VS Code Insiders Portable Updater ===' @logParams
    # Step 1: Validate installation path
    if (-not (Test-Path -Path $InstallPath -PathType Container)) {
        throw "Installation path not found: '$InstallPath'"
    }
    # Step 2: Check if VS Code Insiders is running
    if (Test-VSCodeInsidersRunning) {
        New-Log "'Code - Insiders.exe' is currently running. Please close it before updating." -Level WARNING @logParams
        exit 1
    }
    # Step 3: Get installed version
    $installed = Get-InstalledVersion -InstallPath $InstallPath
    if ($installed.Version) {
        New-Log "Installed version: $($installed.Version)" @logParams
        New-Log "Installed commit:  $($installed.Commit)" -Level VERBOSE @logParams
        if ($installed.ExeVersion -and $installed.ExeVersion -ne $installed.Version) {
            New-Log "Exe version:       $($installed.ExeVersion)" @logParams
        }
    }
    else {
        New-Log 'No existing version detected. Will perform initial extraction.' -Level WARNING @logParams
    }
    # Step 4: Query for the latest build
    $latest = Get-LatestInsiderBuild
    New-Log "Latest version:    $($latest.Version)" @logParams
    New-Log "Latest commit:     $($latest.Commit)" -Level VERBOSE @logParams
    # Step 5: Compare commits — only a commit match means the on-disk files are current.
    # The exe version alone is unreliable because VS Code's auto-updater can update the
    # exe while leaving the rest of the installation (product.json, resources, etc.) stale.
    if ($installed.Commit -and $installed.Commit -eq $latest.Commit) {
        New-Log 'Already up to date. No update needed.' -Level SUCCESS @logParams
        exit 0
    }
    $fromVersion = if ($installed.ExeVersion) { $installed.ExeVersion } else { $installed.Version }
    New-Log "Update available: $fromVersion.$($installed.Commit) -> $($latest.Version).$($latest.Commit))" -Level SUCCESS @logParams
    # Step 6: Download the zip
    $zipPath = Get-VSCodeInsidersZip -DownloadUrl $latest.DownloadUrl -Sha256Hash $latest.Sha256Hash
    try {
        # Step 7: Clean up old hash folders
        New-Log 'Cleaning up old hash folders...' @logParams
        Remove-HashFolders -InstallPath $InstallPath
        # Step 8: Extract and install
        Expand-VSCodeInsidersZip -ZipPath $zipPath -InstallPath $InstallPath
    }
    finally {
        if (Test-Path -Path $zipPath) {
            New-Log 'Cleaning up downloaded zip...' @logParams
            Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
        }
    }
    # Step 9: Verify the update
    $updatedVersion = Get-InstalledVersion -InstallPath $InstallPath
    if ($updatedVersion.Commit -eq $latest.Commit) {
        New-Log 'Update completed successfully!' -Level SUCCESS @logParams
        New-Log "Version: $($updatedVersion.Version)" @logParams
        New-Log "Commit:  $($updatedVersion.Commit)" -Level VERBOSE @logParams
    }
    else {
        New-Log 'Update may not have applied correctly. Please verify manually.' -Level WARNING @logParams
    }
}
catch {
    New-Log "Update failed: $_" -Level ERROR @logParams
    exit 1
}
#endregion Main
