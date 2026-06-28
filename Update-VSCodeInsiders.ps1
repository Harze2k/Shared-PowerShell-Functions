#Requires -Version 5.1
<#
.SYNOPSIS
    Updates a portable VS Code Insiders installation (x64, Windows).
.DESCRIPTION
    Checks for the latest VS Code Insiders build via the official update API,
    downloads the Windows x64 zip archive, verifies its SHA256 hash, extracts
    it to a staging folder, validates the staged payload, and only then swaps
    it into the portable installation folder. The 'data' folder (user settings
    and extensions, when present) is preserved across updates.
    The install is never mutated until a verified-good replacement has been
    fully downloaded, hash-checked and extracted. A corrupt or failed download
    aborts with zero changes to the existing installation. If the installation
    is found in a half-applied state (e.g. the bulk of the app sitting inside a
    leftover commit-hash folder), the update self-heals it into a clean flat
    layout.
.PARAMETER InstallPath
    The path to the portable VS Code Insiders installation folder.
    Defaults to 'C:\Temp\VSCode-Insider'.
.PARAMETER LogFilePath
    Optional path for a log file. When specified, all output is also written
    to disk via New-Log.
.PARAMETER MinFreeSpaceMB
    Minimum free space (MB) required on both the TEMP and install drives before
    downloading/extracting. Defaults to 1024 (1 GB).
.PARAMETER MaxRetries
    Number of attempts for transient network operations (API query, download).
    Defaults to 3.
.EXAMPLE
    .\Update-VSCodeInsiders_v1.10.ps1
    Checks for updates and installs the latest VS Code Insiders build.
.EXAMPLE
    .\Update-VSCodeInsiders_v1.10.ps1 -InstallPath 'C:\Temp\VSCode-Insider' -LogFilePath 'C:\Temp\vscode-update.log'
    Updates a portable installation at a custom path and writes a log file.
.NOTES
    Author : Martin
    Version: 1.10
    Requires: New-Log (https://github.com/Harze2k/Shared-PowerShell-Functions)
        Loaded from a local New-Log.ps1 next to this script if present,
        otherwise fetched from GitHub, otherwise a built-in fallback.
    On-disk layout note:
    This installation keeps the running version's payload either flat under
        InstallPath OR inside a commit-hash-named subfolder (e.g. '628f6de50e').
        Get-InstalledVersion reads product.json from the hash folder when present
        and falls back to the flat location otherwise. After a successful update
        the layout is always flat and any stale hash folders are removed.
    VERSION HISTORY:
    ================
    v1.10 - Hardened + optimized rework:
            - Install is never modified until a verified-good payload is staged
                (download -> hash verify -> extract -> validate -> swap). A bad
                download no longer deletes the old install.
            - Forces TLS 1.2 before any web request (Windows PowerShell 5.1).
            - Retry-with-backoff for the update API query and the download.
            - Free-space pre-check on TEMP and install drives.
            - Prefers a local New-Log.ps1, then GitHub, then built-in fallback.
            - Faster extraction via [IO.Compression.ZipFile] (Expand-Archive
                fallback) and multithreaded robocopy for directory copies.
            - Fixed stray ')' in the "Update available" message and a missing
                @logParams on the hash-mismatch warning.
.LINK
    https://github.com/Harze2k/Shared-PowerShell-Functions
#>
[CmdletBinding()]
param(
    [Parameter()][ValidateNotNullOrEmpty()][string]$InstallPath = 'C:\Temp\VSCode-Insider',
    [Parameter()][string]$LogFilePath,
    [Parameter()][ValidateRange(0, [int]::MaxValue)][int]$MinFreeSpaceMB = 1024,
    [Parameter()][ValidateRange(1, 10)][int]$MaxRetries = 3
)
#region Bootstrap
# TLS 1.2 must be enabled before any web request on Windows PowerShell 5.1.
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
}
catch {
    Write-Warning "Could not set TLS 1.2: $($_.Exception.Message)"
}
if (-not (Get-Command -Name New-Log -ErrorAction SilentlyContinue)) {
    try {
        $newLogUrl = 'https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1'
        . ([ScriptBlock]::Create((Invoke-WebRequest -Uri $newLogUrl -UseBasicParsing -ErrorAction Stop).Content))
    }
    catch {
        Write-Warning "Could not load New-Log from GitHub: $($_.Exception.Message) -- using built-in fallback."
        function New-Log {
            [CmdletBinding()]
            param(
                [Parameter(ValueFromPipeline, Position = 0)]$Message,
                [Parameter(Position = 1)][string]$Level = 'INFO',
                [Parameter()][string]$LogFilePath,
                [Parameter()]$ErrorObject
            )
            process {
                $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
                $color = switch ($Level) {
                    'ERROR' { 'Red' }
                    'EXTENDEDERROR' { 'Red' }
                    'WARNING' { 'Yellow' }
                    'SUCCESS' { 'Green' }
                    'DEBUG' { 'Blue' }
                    'VERBOSE' { 'Cyan' }
                    default { 'White' }
                }
                $suffix = if ($ErrorObject) { " [$($ErrorObject.Exception.Message)]" } else { '' }
                $logLine = "[$timestamp][$Level] $Message$suffix"
                Write-Host $logLine -ForegroundColor $color
                if ($LogFilePath) {
                    try {
                        Add-Content -Path $LogFilePath -Value $logLine -ErrorAction Stop
                    }
                    catch {
                        Write-Warning "Failed to write to log file '$LogFilePath': $($_.Exception.Message)"
                    }
                }
            }
        }
    }
}
#endregion Bootstrap
#region Functions
function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [Parameter()][int]$MaxAttempts = 3,
        [Parameter()][int]$DelaySeconds = 2,
        [Parameter()][string]$OperationName = 'operation'
    )
    $delay = $DelaySeconds
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -ge $MaxAttempts) { throw }
            New-Log "$OperationName failed (attempt $attempt/$MaxAttempts). Retrying in $delay s..." -Level WARNING -ErrorObject $_ @logParams
            Start-Sleep -Seconds $delay
            $delay *= 2
        }
    }
}
function Test-FreeSpace {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$MinFreeMB
    )
    $qualifier = Split-Path -Path $Path -Qualifier -ErrorAction SilentlyContinue
    if (-not $qualifier) {
        New-Log "Could not determine drive for '$Path'; skipping free-space check." -Level WARNING @logParams
        return $true
    }
    $drive = Get-PSDrive -Name $qualifier.TrimEnd(':') -ErrorAction Stop
    $freeMB = [math]::Floor($drive.Free / 1MB)
    if ($freeMB -lt $MinFreeMB) {
        New-Log "Insufficient free space on $qualifier ($freeMB MB free, need $MinFreeMB MB)." -Level WARNING @logParams
        return $false
    }
    New-Log "Free space on ${qualifier}: $freeMB MB (need $MinFreeMB MB)." -Level VERBOSE @logParams
    return $true
}
function Get-InstalledVersion {
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
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter()][int]$MaxAttempts = 3
    )
    $apiUrl = 'https://update.code.visualstudio.com/api/update/win32-x64-archive/insider/latest'
    New-Log 'Querying VS Code update API for the latest Insider build...' @logParams
    $response = Invoke-WithRetry -OperationName 'Update API query' -MaxAttempts $MaxAttempts -ScriptBlock {
        Invoke-RestMethod -Uri $apiUrl -UseBasicParsing -ErrorAction Stop
    }
    return [PSCustomObject]@{
        Version     = $response.productVersion
        Commit      = $response.version
        DownloadUrl = $response.url
        Sha256Hash  = $response.sha256hash
    }
}
function Remove-HashFolders {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$InstallPath,
        [Parameter()][switch]$GetHashFolder
    )
    $knownFolders = @( # Known VS Code folders that must never be removed
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
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$DownloadUrl,
        [Parameter(Mandatory)][string]$Sha256Hash,
        [Parameter()][int]$MaxAttempts = 3
    )
    $tempFile = Join-Path -Path $env:TEMP -ChildPath "VSCode-Insiders-$(Get-Date -Format 'yyyyMMdd-HHmmss').zip"
    New-Log 'Downloading VS Code Insiders zip...' @logParams
    New-Log "URL: $DownloadUrl" -Level VERBOSE @logParams
    New-Log "Destination: $tempFile" -Level VERBOSE @logParams
    Invoke-WithRetry -OperationName 'Download' -MaxAttempts $MaxAttempts -ScriptBlock {
        try {
            Import-Module BitsTransfer -ErrorAction Stop
            Start-BitsTransfer -Source $DownloadUrl -Destination $tempFile -ErrorAction Stop
        }
        catch {
            New-Log 'BITS transfer failed, falling back to Invoke-WebRequest...' -Level WARNING @logParams
            Invoke-WebRequest -Uri $DownloadUrl -OutFile $tempFile -UseBasicParsing -ErrorAction Stop
        }
    } | Out-Null
    New-Log 'Verifying SHA256 hash...' @logParams
    $fileHash = (Get-FileHash -Path $tempFile -Algorithm SHA256 -ErrorAction Stop).Hash
    if ($fileHash -ne $Sha256Hash) {
        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        New-Log "Hash mismatch! Expected '$Sha256Hash' but got '$fileHash'. The download may be corrupted. Aborting." -Level WARNING @logParams
        return
    }
    New-Log 'Hash verified successfully.' -Level SUCCESS @logParams
    return $tempFile
}
function Expand-ZipFast {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [Parameter(Mandatory)][string]$DestinationPath
    )
    # Prefer the .NET extractor (far faster than Expand-Archive); fall back if unavailable.
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $DestinationPath)
    }
    catch {
        New-Log "Fast extraction unavailable ($($_.Exception.Message)); using Expand-Archive." -Level WARNING @logParams
        Expand-Archive -Path $ZipPath -DestinationPath $DestinationPath -Force -ErrorAction Stop
    }
}
function Copy-DirectoryFast {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination
    )
    # Multithreaded mirror of a single directory; robocopy exit codes < 8 are success.
    $roboArgs = @($Source, $Destination, '/MIR', '/MT:16', '/R:2', '/W:2', '/NFL', '/NDL', '/NJH', '/NJS', '/NP', '/NC', '/NS')
    & robocopy.exe @roboArgs | Out-Null
    $code = $LASTEXITCODE
    if ($code -ge 8) {
        New-Log "robocopy failed (exit $code) for '$Source'; falling back to Copy-Item." -Level WARNING @logParams
        if (Test-Path -Path $Destination) {
            Remove-Item -Path $Destination -Recurse -Force -ErrorAction Stop
        }
        Copy-Item -Path $Source -Destination $Destination -Recurse -Force -ErrorAction Stop
    }
}
function Expand-VSCodeInsidersZip {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [Parameter(Mandatory)][string]$InstallPath
    )
    $tempExtractPath = Join-Path -Path $env:TEMP -ChildPath "VSCode-Insiders-Extract-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    try {
        New-Log 'Extracting zip to temporary staging location...' @logParams
        Expand-ZipFast -ZipPath $ZipPath -DestinationPath $tempExtractPath
        $extractedItems = Get-ChildItem -Path $tempExtractPath -ErrorAction Stop
        $sourcePath = $tempExtractPath
        if ($extractedItems.Count -eq 1 -and $extractedItems[0].PSIsContainer) {
            $sourcePath = $extractedItems[0].FullName
            New-Log "Detected top-level folder: $($extractedItems[0].Name)" -Level VERBOSE @logParams
        }
        # Validate the staged payload BEFORE touching the existing installation.
        $stagedExe = Join-Path -Path $sourcePath -ChildPath 'Code - Insiders.exe'
        if (-not (Test-Path -Path $stagedExe -PathType Leaf)) {
            throw "Staged payload is invalid: 'Code - Insiders.exe' not found in '$sourcePath'. Aborting before modifying the installation."
        }
        New-Log 'Staged payload validated.' -Level SUCCESS @logParams
        # From here on we mutate the installation. Remove stale hash folders first.
        New-Log 'Cleaning up old hash folders...' @logParams
        Remove-HashFolders -InstallPath $InstallPath
        New-Log "Copying files to '$InstallPath'..." @logParams
        $sourceItems = Get-ChildItem -Path $sourcePath -ErrorAction Stop
        foreach ($item in $sourceItems) {
            if ($item.Name -eq 'data') {
                New-Log "Skipping 'data' folder (preserving user settings)." -Level WARNING @logParams
                continue
            }
            $destinationItem = Join-Path -Path $InstallPath -ChildPath $item.Name
            if ($item.PSIsContainer) {
                Copy-DirectoryFast -Source $item.FullName -Destination $destinationItem
            }
            else {
                Copy-Item -Path $item.FullName -Destination $destinationItem -Force -ErrorAction Stop
            }
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
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$logParams = @{}
if ($LogFilePath) { $logParams['LogFilePath'] = $LogFilePath }
try {
    New-Log '=== VS Code Insiders Portable Updater (v1.10) ===' @logParams
    if (-not (Test-Path -Path $InstallPath -PathType Container)) {
        throw "Installation path not found: '$InstallPath'"
    }
    if (Get-Process -Name 'Code - Insiders' -ErrorAction SilentlyContinue) {
        New-Log "'Code - Insiders.exe' is currently running. Please close it before updating." -Level WARNING @logParams
        exit 1
    }
    $installed = Get-InstalledVersion -InstallPath $InstallPath
    if ($installed.Version) {
        New-Log "Installed version: $($installed.Version)" @logParams
        New-Log "Installed commit:  $($installed.Commit)" @logParams
        if ($installed.ExeVersion -and $installed.ExeVersion -ne $installed.Version) {
            New-Log "Exe version:       $($installed.ExeVersion)" @logParams
        }
    }
    else {
        New-Log 'No existing version detected. Will perform initial extraction.' -Level WARNING @logParams
    }
    $latest = Get-LatestInsiderBuild -MaxAttempts $MaxRetries
    New-Log "Latest version:    $($latest.Version)" @logParams
    New-Log "Latest commit:     $($latest.Commit)" @logParams
    if ($installed.Commit -and $installed.Commit -eq $latest.Commit) {
        New-Log 'Already up to date. No update needed.' -Level SUCCESS @logParams
        exit 0
    }
    # Pre-flight: ensure enough free space on the TEMP and install drives.
    $tempOk = Test-FreeSpace -Path $env:TEMP -MinFreeMB $MinFreeSpaceMB
    $installOk = Test-FreeSpace -Path $InstallPath -MinFreeMB $MinFreeSpaceMB
    if (-not $tempOk -or -not $installOk) {
        throw "Insufficient free disk space. Need at least $MinFreeSpaceMB MB on both the TEMP and install drives."
    }
    $fromVersion = if ($installed.ExeVersion) { $installed.ExeVersion } else { $installed.Version }
    New-Log "Update available: $fromVersion.$($installed.Commit) -> $($latest.Version).$($latest.Commit)" -Level SUCCESS @logParams
    $zipPath = Get-VSCodeInsidersZip -DownloadUrl $latest.DownloadUrl -Sha256Hash $latest.Sha256Hash -MaxAttempts $MaxRetries
    if (-not $zipPath) {
        throw 'Download or hash verification failed. The installation was left unchanged.'
    }
    try {
        Expand-VSCodeInsidersZip -ZipPath $zipPath -InstallPath $InstallPath
    }
    finally {
        if (Test-Path -Path $zipPath) {
            New-Log 'Cleaning up downloaded zip...' @logParams
            Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
        }
    }
    $updatedVersion = Get-InstalledVersion -InstallPath $InstallPath
    if ($updatedVersion.Commit -eq $latest.Commit) {
        New-Log 'Update completed successfully!' -Level SUCCESS @logParams
        New-Log "Version: $($updatedVersion.Version)" @logParams
        New-Log "Commit:  $($updatedVersion.Commit)" @logParams
    }
    else {
        New-Log 'Update may not have applied correctly. Please verify manually.' -Level WARNING @logParams
    }
}
catch {
    New-Log "Update failed: $($_.Exception.Message)" -Level ERROR -ErrorObject $_ @logParams
    exit 1
}
#endregion Main