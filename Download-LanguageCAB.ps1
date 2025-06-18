function Download-LanguageCAB {
<#
.SYNOPSIS
    Downloads Windows Language Packs for a specific OS build and optionally converts them from ESD to CAB format.
.DESCRIPTION
    This function automates the process of fetching and managing Windows Language Pack files from UUP Dump, a well-known repository for Windows Update files. It is designed for IT professionals and system administrators who create custom OS images.
    The function targets a specific Windows version (10 or 11) and build number, then downloads all associated language and feature-on-demand packs for a given language code (e.g., 'de-de').
    Key Features:
    - Automatically finds the correct download links from UUP Dump for a given build.
    - Downloads Language Pack and Feature on Demand files (.esd or .cab).
    - Optionally converts the main Language Pack ESD file into a CAB file, which is often required for DISM servicing.
    - Highly optimized for performance using:
        - Parallel processing for the most intensive conversion steps (SxSExpand).
        - Asynchronous downloads with a reusable HttpClient.
        - Optional high-speed RAM Disk support (via ImDisk Toolkit) to minimize I/O bottlenecks during conversion.
    - Automatically downloads necessary third-party tools (ESD2CAB utility, ImDisk Toolkit) if needed.
.PARAMETER FolderPath
    The root destination path where language pack files and supporting tools will be saved. The function will create a sub-directory within this path named after the specified build number (e.g., C:\LPs\22621.1702).
    This parameter is mandatory and accepts pipeline input.
.PARAMETER Os
    The target Windows operating system version.
    Accepted values: "10", "11"
.PARAMETER Build
    The specific Windows build number to download language packs for, in the format 'xxxxx.yyyy' (e.g., '22621.1702').
    If not specified, the function will automatically detect the build and UBR (Update Build Revision) of the local machine it is running on.
    Default value: (Current system build and UBR)
.PARAMETER Language
    The language code for the packs you want to download, in the 'll-CC' format (e.g., 'en-us', 'de-de', 'ja-jp').
.PARAMETER UUPUrls
    An array of UUP Dump website hostnames to try. The function will attempt to connect to them in order until it gets a successful response. This provides resilience if one of the sites is down.
    Default value: @('www.uupdump.net', 'www.uupdump.cn')
.PARAMETER ESDToCAB
    If this switch is specified, the function will find the downloaded main Language Pack ESD file and convert it into a CAB file. This process requires the 'ESD2CAB' utility, which will be downloaded automatically if not present.
.PARAMETER RemoveESD
    If this switch is specified along with -ESDToCAB, the original ESD file will be deleted after a successful conversion to a CAB file, saving disk space.
.PARAMETER MaxParallelJobs
    Specifies the maximum number of parallel threads to use during the SxSExpand portion of the ESD-to-CAB conversion. This is the most CPU-intensive part of the process.
    Default value: The number of logical processors on the machine ([Environment]::ProcessorCount).
.PARAMETER UseRAMDisk
    If specified, the function will attempt to create and use a temporary RAM disk for the ESD-to-CAB conversion process. This dramatically improves performance by eliminating disk I/O latency. This feature requires the ImDisk Virtual Disk Driver.
.PARAMETER RemoveRAMDisk
    If specified along with -UseRAMDisk, the function will automatically dismount and remove the RAM disk upon completion of the script, freeing up the memory.
.PARAMETER DriveLetter
    Specifies the drive letter to assign to the RAM disk when -UseRAMDisk is enabled.
    Default value: 'R'
.PARAMETER ForceInstallImDisk
    If specified, the script will attempt to automatically download and install 'imdisk.exe' to 'C:\Windows\System32' if it is not found in the system's PATH.
    Note: The script may need to be re-run after the installation completes. Administrative privileges are required for this operation.
.EXAMPLE
    # Example 1: Basic Download for the Current System's Build
    # This command downloads German (de-de) language packs for the same build as the machine running the script
    # into the C:\LanguagePacks folder.
    Download-LanguageCAB -FolderPath C:\LanguagePacks -Os 11 -Language de-de
.EXAMPLE
    # Example 2: Download for a Specific Build and Convert ESD to CAB
    # This downloads French (fr-fr) packs for Windows 11 build 22621.1702, converts the main ESD to a CAB,
    # and then deletes the original ESD file.
    Download-LanguageCAB -FolderPath D:\OS_Creation\LPs -Os 11 -Build 22621.1702 -Language fr-fr -ESDToCAB -RemoveESD
.EXAMPLE
    # Example 3: Maximum Performance Conversion using a RAM Disk
    # This command downloads Japanese (ja-jp) packs, then performs the ESD-to-CAB conversion using a RAM disk
    # on drive Z: for maximum speed. It also automatically cleans up the RAM disk when finished.
    # Administrative privileges are recommended for this.
    Download-LanguageCAB -FolderPath C:\Temp\LPs -Os 11 -Language ja-jp -ESDToCAB -UseRAMDisk -RemoveRAMDisk -DriveLetter Z
.EXAMPLE
    # Example 4: First Time Use with Automatic Tool Installation
    # If ImDisk is not installed, this command will attempt to download and install it, then proceed with
    # the download and conversion using the RAM disk.
    Download-LanguageCAB -FolderPath C:\LPs -Os 10 -Language es-es -ESDToCAB -UseRAMDisk -ForceInstallImDisk
.EXAMPLE
    # Example 5: Using the Pipeline for the Folder Path
    # This demonstrates piping the destination folder path to the function.
    'C:\MyLanguagePacks' | Download-LanguageCAB -Os 11 -Language it-it -ESDToCAB
.INPUTS
    System.String
    You can pipe a string containing the destination folder path to the -FolderPath parameter.
.OUTPUTS
    None
    This function does not return any objects to the pipeline. It writes files to the disk and outputs status and log messages to the console.
.NOTES
    Author:	Harze2k
    Date: 2025-06-18
	Version: 2.0 (Complete remake)
        -Parallel processing!
        -RAMDisk usage!
        -Lots other smaller fixes.
.LINK
    UUP Dump: https://www.uupdump.net
    ESD2CAB Tool: https://github.com/abbodi1406/WHD/
    ImDisk Toolkit: https://sourceforge.net/projects/imdisk-toolkit/
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)][ValidateNotNullOrEmpty()][string]$FolderPath,
        [Parameter(Mandatory)][ValidateSet("10", "11")][string]$Os,
        [ValidatePattern("^\d{5}\.\d+$")][string]$Build = ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name 'CurrentBuild').CurrentBuild) + '.' + ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name 'UBR').UBR),
        [Parameter(Mandatory)][ValidatePattern("^[a-z]{2}-[a-z]{2}$")][string]$Language,
        [ValidateNotNullOrEmpty()][string[]]$UUPUrls = @('www.uupdump.net', 'www.uupdump.cn'),
        [switch]$ESDToCAB,
        [switch]$RemoveESD,
        [int]$MaxParallelJobs = ([Environment]::ProcessorCount),
        [switch]$UseRAMDisk,
        [switch]$RemoveRAMDisk,
        [string]$DriveLetter = 'R',
        [switch]$ForceInstallImDisk
    )
    begin {
        function Install-ImDiskToolkit {
            [CmdletBinding()]
            param (
                [string]$DownloadUrl = 'https://sourceforge.net/projects/imdisk-toolkit/files/latest/download',
                [string]$TempZip = "$env:Temp\ImDiskTk-x64.zip",
                [string]$ExtractDir = "$env:Temp\imdiskDir",
                [string]$CabDir = "$env:Temp\imdiskDir\cab",
                [string]$Destination = 'C:\Windows\System32\'
            )
            try {
                if (-not (Get-Command expand.exe -ErrorAction SilentlyContinue)) {
                    New-Log 'expand.exe not found in PATH' -Level WARNING
                    return
                }
                Invoke-WebRequest -UserAgent 'Wget' -Uri $DownloadUrl -OutFile $TempZip -ErrorAction Stop
                Expand-Archive -Path $TempZip -DestinationPath $ExtractDir -Force -ErrorAction Stop
                $imdiskcab = (Get-ChildItem -Path $ExtractDir -Include '*.cab' -Recurse | Select-Object -First 1).FullName
                if (-not $imdiskcab) {
                    New-Log "CAB file not found in $ExtractDir"  -Level WARNING
                    return
                }
                New-Item -Path $CabDir -ItemType Directory -Force | Out-Null
                & expand.exe -F:* $imdiskcab $CabDir > $null 2>&1
                if ($LASTEXITCODE -ne 0) {
                    New-Log 'expand.exe failed to extract CAB' -Level WARNING
                }
                $imdiskExe = (Get-ChildItem -Path $CabDir -Filter 'imdisk.exe' -Recurse).FullName | Where-Object { $_ -match 'amd64' }
                if (-not $imdiskExe) {
                    New-Log "imdisk.exe for amd64 not found in $CabDir" -Level WARNING
                    return
                }
                Copy-Item -Path $imdiskExe -Destination $Destination -Force -ErrorAction Stop
                New-Log "Successfully placed imdisk.exe in $Destination" -Level SUCCESS
            }
            catch {
                New-Log "Failed to install ImDisk Toolkit" -Level ERROR
            }
            Remove-ItemSilent -Path $ExtractDir -Force -Recurse
        }
        function Remove-ItemSilent {
            [CmdletBinding()]
            param (
                [string]$Path,
                [switch]$Recurse,
                [switch]$Force
            )
            $oldProgressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            try {
                if ($Recurse -and $Force) {
                    Remove-Item $Path -Recurse -Force -ErrorAction Stop
                }
                elseif ($Recurse) {
                    Remove-Item $Path -Recurse -ErrorAction Stop
                }
                elseif ($Force) {
                    Remove-Item $Path -Force -ErrorAction Stop
                }
                else {
                    Remove-Item $Path -ErrorAction Stop
                }
                New-Log "Successfully removed path $path" -Level SUCCESS
            }
            catch {
                New-Log "Failed to remove path $path" -Level ERROR
            }
            finally {
                $ProgressPreference = $oldProgressPreference
            }
        }
        function Download-File {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DownloadLink,
                [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OutputFile,
                [Parameter(Mandatory)][string]$FileName,
                [Parameter(Mandatory)][System.Net.Http.HttpClient]$HttpClient
            )
            if ([string]::IsNullOrEmpty($fileName)) {
                [string]$fileName = $(Split-Path $outputFile -Leaf)
            }
            try {
                $response = $HttpClient.GetAsync($downloadLink).Result
                $response.EnsureSuccessStatusCode()
                $fileStream = $null
                try {
                    $fileStream = [System.IO.File]::Create($outputFile)
                    $response.Content.CopyToAsync($fileStream).Wait()
                }
                finally {
                    if ($fileStream) {
                        $fileStream.Dispose()
                    }
                }
                $response.Dispose()
                New-Log "Successfully downloaded $fileName with HttpClient." -Level SUCCESS
                return $true
            }
            catch {
                New-Log "Failed to download $fileName with HttpClient: $($_.Exception.Message)" -Level ERROR
                return $false
            }
        }
        function New-TempRAMDisk {
            [CmdletBinding()]
            param (
                [int]$SizeMB = 2048,
                [string]$DriveLetter = "R",
                [switch]$ForceInstallImDisk
            )
            try {
                $imdiskPath = Get-Command "imdisk.exe" -ErrorAction SilentlyContinue
                if (-not $imdiskPath -and $ForceInstallImDisk.IsPresent) {
                    Install-ImDiskToolkit
                    New-Log "Just installed imdisk.exe, need to rerun the script."
                    exit 0
                }
                elseif ($ForceInstallImDisk.IsPresent) {
                    Install-ImDiskToolkit
                }
                $imdiskPath = Get-Command "imdisk.exe" -ErrorAction SilentlyContinue
                if ($imdiskPath) {
                    New-Log "Creating RAM disk using imdisk..." -Level INFO
                    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
                    if (-not $isAdmin) {
                        New-Log "Warning: Not running as administrator - imdisk may fail" -Level WARNING
                    }
                    New-Log "Creating unformatted RAM disk..." -Level DEBUG
                    $imdiskexe = $imdiskPath.Source
                    $cmdArgs = "/c `"$imdiskexe -a -t vm -s ${SizeMB}M -m ${DriveLetter}: >nul 2>&1`""
                    $process = Start-Process -FilePath "cmd.exe" -ArgumentList $cmdArgs -Wait -PassThru -WindowStyle Hidden
                    if ($process.ExitCode -eq 0) {
                        New-Log "RAM disk created, formatting with NTFS..." -Level INFO
                        try {
                            $formatCmdArgs = "/c `"format.com ${DriveLetter}: /fs:ntfs /q /y /v:RAMDISK >nul 2>&1`""
                            $formatProcess = Start-Process -FilePath "cmd.exe" -ArgumentList $formatCmdArgs -Wait -PassThru -WindowStyle Hidden
                            if ($formatProcess.ExitCode -eq 0) {
                                New-Log "Successfully created and formatted RAM disk ${DriveLetter}: using imdisk" -Level SUCCESS
                                return $DriveLetter
                            }
                            else {
                                New-Log "RAM disk created but formatting failed. Drive may still be usable." -Level WARNING
                                return $DriveLetter
                            }
                        }
                        catch {
                            New-Log "Error during formatting: $($_.Exception.Message). Drive may still be usable." -Level WARNING
                            return $DriveLetter
                        }
                    }
                }
            }
            catch {
                New-Log "Error creating RAM disk: $($_.Exception.Message)" -Level ERROR
            }
            return $null
        }
        function Remove-TempRAMDisk {
            [CmdletBinding()]
            param (
                [string]$DriveLetter = 'R'
            )
            if (-not $DriveLetter) { return }
            try {
                if ($DriveLetter.Length -eq 1) {
                    $drivePath = "${DriveLetter}:"
                    New-Log "Cleaning up RAM disk $drivePath..." -Level INFO
                    $imdiskPath = Get-Command "imdisk.exe" -ErrorAction SilentlyContinue
                    if ($imdiskPath) {
                        $imdiskexe = $imdiskPath.Source
                        $cmdArgs = "/c `"$imdiskexe -D -m $drivePath >nul 2>&1`""
                        $process = Start-Process -FilePath "cmd.exe" -ArgumentList $cmdArgs -Wait -PassThru -WindowStyle Hidden
                        if ($process.ExitCode -eq 0) {
                            New-Log "Successfully removed RAM disk using imdisk" -Level SUCCESS
                        }
                        else {
                            New-Log "Failed to force-remove RAM disk, exit code: $($forceProcess.ExitCode)" -Level ERROR
                        }
                    }
                }
            }
            catch {
                New-Log "Error during RAM disk cleanup: $($_.Exception.Message)" -Level WARNING
            }
        }
        function Convert-EsdToCab-Optimized {
            [CmdletBinding()]
            param (
                [string]$EsdFile,
                [string]$WorkingDir,
                [string]$ImageX,
                [string]$Cabarc,
                [string]$Sxs,
                [switch]$RemoveESD,
                [int]$MaxParallelJobs = 4,
                [switch]$UseRAMDisk,
                [string]$DriveLetter = 'R',
                [switch]$ForceInstallImDisk
            )
            $pack = [System.IO.Path]::GetFileNameWithoutExtension($EsdFile)
            $cabFile = Join-Path $WorkingDir "$pack.cab"
            if (Test-Path $cabFile) {
                New-Log "$cabFile already exists, skipping."
                return
            }
            if ($UseRAMDisk) {
                New-Log "Creating temporary RAM disk for optimal performance..." -Level INFO
                if (!(Test-Path "${DriveLetter}:\")) {
                    $ramDiskCreated = New-TempRAMDisk -DriveLetter $DriveLetter -ForceInstallImDisk:$ForceInstallImDisk.IsPresent
                }
                else {
                    $ramDiskCreated = $DriveLetter
                }
                if ($ramDiskCreated) {
                    if ($ramDiskCreated.Length -eq 1) {
                        $tempDir = Join-Path "${ramDiskCreated}:" "temp$([System.IO.Path]::GetRandomFileName())"
                    }
                }
                else {
                    New-Log "Failed to create RAM disk, using regular temp directory" -Level WARNING
                    $tempDir = Join-Path $env:TEMP "ESDConvert$([System.IO.Path]::GetRandomFileName())"
                }
            }
            else {
                $tempDir = Join-Path $env:TEMP "ESDConvert$([System.IO.Path]::GetRandomFileName())"
            }
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                New-Log "Extracting ESD file..." -Level INFO
                $extractProcess = & $ImageX /APPLY $EsdFile 1 $tempDir /NOACL ALL /NOTADMIN /TEMP $env:TEMP #| Out-Null
                if ($LASTEXITCODE -eq 0 -or $extractProcess -match 'Successfully') {
                    New-Log "Successfully extracted esd file $EsdFile with ImageX in $($stopwatch.Elapsed.TotalSeconds) seconds." -Level SUCCESS
                    $stopwatch.Restart()
                }
                else {
                    New-Log "ImageX extraction failed with exit code $LASTEXITCODE" -Level WARNING
                }
            }
            catch {
                New-Log "Extracting with ImageX had an error: $($_.Exception.Message)" -Level ERROR
                return
            }
            Push-Location $tempDir
            New-Item -ItemType Directory -Path "_sxs" -Force | Out-Null
            $manifestFiles = Get-ChildItem -Filter "*.manifest"
            if ($manifestFiles.Count -eq 0) {
                New-Log "No manifest files found to process" -Level WARNING
                Pop-Location
                return
            }
            New-Log "Converting $($manifestFiles.Count) manifest files using $MaxParallelJobs parallel jobs..." -Level INFO
            $manifestFiles | ForEach-Object -Parallel {
                $manifest = $_
                $sxsPath = $using:Sxs
                $tempDir = $using:tempDir
                try {
                    $outputPath = Join-Path $tempDir "_sxs\$($manifest.Name)"
                    & $sxsPath $manifest.Name $outputPath | Out-Null
                }
                catch {
                    Write-Error "Error processing $($manifest.Name): $($_.Exception.Message)"
                }
            } -ThrottleLimit $MaxParallelJobs
            $conversionTime = $stopwatch.Elapsed.TotalSeconds
            $stopwatch.Restart()
            if (Test-Path "_sxs\*.manifest") {
                Move-Item "_sxs\*" . -Force
            }
            Remove-ItemSilent -Path "_sxs" -Recurse -Force
            New-Log "Creating CAB archive..." -Level INFO
            try {
                & $Cabarc -m LZX:15 -r -p N "$cabFile" *.* | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    $cabTime = $stopwatch.Elapsed.TotalSeconds
                    New-Log "Successfully added all files to $pack.cab with Cabarc in $cabTime seconds." -Level SUCCESS
                    New-Log "=========================================================="
                    New-Log "Successfully expanded, converted and created the cab file."
                    New-Log "Total conversion time: $($conversionTime + $cabTime) seconds"
                    New-Log "=========================================================="
                }
                else {
                    New-Log "Cabarc failed with exit code: $LASTEXITCODE" -Level WARNING
                }
            }
            catch {
                New-Log "Adding to cab archive with Cabarc had an error: $($_.Exception.Message)" -Level ERROR
            }
            Pop-Location
            Remove-ItemSilent -Path $tempDir -Recurse -Force
            if ($RemoveESD.IsPresent -and (Test-Path $cabFile)) {
                Remove-ItemSilent -Path ".\$pack.esd" -Recurse -Force
            }
        }
        $script:ToolsCached = $false
        ##############################################################################################################################
        #Calling custom logging function from: https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1#
        ##############################################################################################################################
        if (-not (Get-Command -Name "New-Log" -ErrorAction SilentlyContinue)) {
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
        }
        ################################################################################################################################################
        #Calling custom Get-RandomHeader function from: https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/Get-RandomHeader.ps1#
        ################################################################################################################################################
        if (-not (Get-Command -Name "Get-RandomHeader" -ErrorAction SilentlyContinue)) {
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/Get-RandomHeader.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
        }
        try {
            if (-not (Test-Path "$folderPath\$build")) {
                New-Log "Creating folder: $folderPath\$build"
                New-Item -Path "$folderPath\$build" -ItemType Directory -Force | Out-Null
            }
        }
        catch {
            New-Log "Failed to create or access folder: $folderPath\$build" -Level ERROR
            return
        }
    }
    process {
        $totalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $lang = ($Language -split '-')[0]
        $headers = Get-RandomHeader
        $WebResponse = $null
        foreach ($UUPUrl in $UUPUrls) {
            try {
                $uri = "https://$UUPUrl/known.php?q=windows+$($Os)+$($Build)"
                $WebResponse = (Invoke-WebRequest -Uri $uri -UseBasicParsing -MaximumRedirection 1 -Method GET -Headers $headers -TimeoutSec 30 -ErrorAction Stop).Links
                if ($WebResponse.Count -gt 0) {
                    New-Log "Got webResponse from $UUPUrl." -Level SUCCESS
                    break
                }
                else {
                    continue
                }
            }
            catch {
                New-Log "Failed to get WebResponse from $UUPUrl - $($_.Exception.Message)" -Level ERROR
            }
        }
        if ($null -eq $WebResponse -or $WebResponse.Count -eq 0) {
            New-Log "Failed to get WebResponse from any UUP site. Aborting." -Level WARNING
            return
        }
        $Links = $null
        foreach ($UUPUrl in $UUPUrls) {
            try {
                $UpdateID = (($WebResponse | Where-Object { $_.href -match "(./selectlang.php\?id|selectlang.php\?id)" }).href).split("=")[1]
                New-Log "Using UpdateId: $UpdateID" -Level DEBUG
                $uri = "https://$UUPUrl/get.php?id=$UpdateID&pack=$Language&edition=core"
                $Links = (Invoke-WebRequest -Uri $uri -UseBasicParsing -MaximumRedirection 1 -Method GET -Headers $headers -TimeoutSec 30 -ErrorAction Stop).Links
                if ($Links.Count -gt 0) {
                    New-Log "Got links response from $UUPUrl." -Level SUCCESS
                    break
                }
            }
            catch {
                New-Log "Failed to get download links from $UUPUrl - $($_.Exception.Message)" -Level ERROR
                continue
            }
        }
        if ($null -eq $Links -or $Links.Count -eq 0) {
            New-Log "Failed to get links to parse. Will not download any files. Aborting." -Level WARNING
            return
        }
        $downloadLinks = $true
        if ($Links) {
            New-Log "========================================================"
            New-Log "Start of finding and downloading language cab/esd files."
            New-Log "========================================================"
            try {
                $httpClient = Get-RandomHeader -GetHTTPClient -ErrorAction Stop
                $httpClient.Timeout = [TimeSpan]::FromMinutes(10)
                New-Log "Created reusable HttpClient for downloads" -Level DEBUG
                foreach ($link in ($Links.outerHTML -match ".*(LanguagePack|LanguageFeatures).*$lang-.*")) {
                    $URL = $link.Split('"')[1]
                    $Filename = $link.Split('>')[1].Split('<')[0] -replace '\s', ''
                    $FilenameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($Filename)
                    $FileExtension = [IO.Path]::GetExtension($Filename)
                    $NewFilename = "$FilenameWithoutExtension-$Build$FileExtension"
                    $OutFile = Join-Path "$folderPath\$build" $NewFilename
                    if (!(Test-Path $OutFile)) {
                        $downloadLinks = Download-File -DownloadLink $URL -OutputFile $OutFile -FileName $NewFilename -HttpClient $httpClient
                        if (-not $downloadLinks) { break }
                    }
                    else {
                        New-Log "$NewFilename was already downloaded, skipping." -Level DEBUG
                    }
                }
            }
            finally {
                if ($httpClient) {
                    $httpClient.Dispose()
                    New-Log "Disposed HttpClient after downloads" -Level DEBUG
                }
            }
        }
        if (!(Test-Path "$folderPath\$build\esd2cab_CLI.cmd") -and !$script:ToolsCached) {
            New-Log "Downloading and caching ESD2CAB cmd tools..." -Level DEBUG
            try {
                $toolsZip = "$folderPath\$build\ESD2CAB.zip"
                Invoke-WebRequest -Uri 'https://github.com/abbodi1406/WHD/raw/master/scripts/ESD2CAB-CAB2ESD-2.zip' -UseBasicParsing -MaximumRedirection 1 -Method GET -OutFile $toolsZip -ErrorAction Stop
                Expand-Archive -Path $toolsZip -DestinationPath "$folderPath\$build" -Force -ErrorAction Stop
                Remove-ItemSilent -Path $toolsZip -Force -Recurse
                $script:ToolsCached = $true
            }
            catch {
                New-Log "Failed to download ESD2CAB cmd tool: $($_.Exception.Message)" -Level ERROR
                $downloadLinks = $false
            }
        }
        if ($ESDToCAB -and $downloadLinks) {
            Set-Location -Path "$folderPath\$build"
            $bits = if ($env:PROCESSOR_ARCHITECTURE -eq "AMD64") { "x64" } else { "x86" }
            $requiredFiles = @("image$bits.exe", "cabarc.exe", "SxSExpand.exe")
            $missingFiles = $requiredFiles | Where-Object { -not (Test-Path "bin\$_") }
            if ($missingFiles.Count -gt 0) {
                New-Log "Missing required files: $($missingFiles -join ', ')" -Level ERROR
                return
            }
            $ImageX = Resolve-Path "bin\image$bits.exe"
            $Cabarc = Resolve-Path "bin\cabarc.exe"
            $Sxs = Resolve-Path "bin\SxSExpand.exe"
            $esdFile = Get-ChildItem -Path . -Include "*Windows*LanguagePack*$language*.esd" -Recurse
            if ($null -eq $esdFile) {
                New-Log "No .esd file detected." -Level WARNING
                return
            }
            Convert-EsdToCab-Optimized -EsdFile $esdFile.Name -WorkingDir "$folderPath\$build" -ImageX $imageX -Cabarc $cabarc -Sxs $sxs -MaxParallelJobs $MaxParallelJobs -UseRAMDisk:$UseRAMDisk -RemoveESD:$RemoveESD -DriveLetter $DriveLetter -ForceInstallImDisk:$ForceInstallImDisk.IsPresent
            $totalTime = $totalStopwatch.Elapsed.TotalSeconds
            New-Log "All downloads and conversions completed in $totalTime seconds." -Level SUCCESS
        }
        elseif ($ESDToCAB -and !$downloadLinks) {
            New-Log "Either the download of the CABs failed or the esd to cab tool failed to download. Will abort." -Level WARNING
        }
        if ($RemoveRAMDisk) {
            Remove-TempRAMDisk -DriveLetter $DriveLetter
        }
    }
}
#Some more examples:
<#
$currentBuild = ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name 'CurrentBuild').CurrentBuild) + '.' + ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name 'UBR').UBR)
Download-LanguageCAB -FolderPath "C:\Toolkit\Toolkit_v13.7\Data\Temp\Intune\LanguageCAB" -Os "11" -Language "sv-se" -RemoveESD -ESDToCAB -UUPUrls @('www.uupdump.net', 'www.uupdump.cn') -Build $currentBuild -UseRAMDisk
Download-LanguageCAB -FolderPath "C:\Toolkit\Toolkit_v13.7\Data\Temp\Intune\LanguageCAB" -Os "11" -Language "fi-fi" -RemoveESD -ESDToCAB -UUPUrls @('www.uupdump.net', 'www.uupdump.cn') -Build $currentBuild -UseRAMDisk
Download-LanguageCAB -FolderPath "C:\Toolkit\Toolkit_v13.7\Data\Temp\Intune\LanguageCAB" -Os "11" -Language "nb-no" -RemoveESD -ESDToCAB -UUPUrls @('www.uupdump.net', 'www.uupdump.cn') -Build $currentBuild -UseRAMDisk
Download-LanguageCAB -FolderPath "C:\Toolkit\Toolkit_v13.7\Data\Temp\Intune\LanguageCAB" -Os "11" -Language "ja-jp" -RemoveESD -ESDToCAB -UUPUrls @('www.uupdump.net', 'www.uupdump.cn') -Build $currentBuild -UseRAMDisk -RemoveRAMDisk
#>