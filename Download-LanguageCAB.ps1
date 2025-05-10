function Download-LanguageCAB {
	<#
	.SYNOPSIS
		Downloads and manages Windows language pack CAB/ESD files for OS image creation.
	.DESCRIPTION
		The Download-LanguageCAB function is a comprehensive tool for downloading and managing Windows language pack files.
		It's useful when creating language installation scripts and OS images.
		The function supports downloading language-specific CAB/ESD files for Windows 10 or 11 from UUP dump sites.
		Key capabilities:
		- Downloads language packs from multiple configurable UUP dump sites
		- Converts ESD files to CAB format for easier deployment
		- Optionally removes original ESD files after successful conversion
		- Uses randomized user agents and request headers to prevent bot detection
		- Automatically manages the required conversion tools (imagex, cabarc, SxSExpand)
		DEPENDENCIES:
		- Requires New-Log function: https://github.com/Harze2k/Shared-PowerShell-Functions/blob/main/New-Log.ps1
		- Requires Get-RandomHeader function: https://github.com/Harze2k/Shared-PowerShell-Functions/blob/main/Get-RandomHeader.ps1
	.PARAMETER FolderPath
		Specifies the directory where the downloaded files will be saved. This parameter is mandatory.
	.PARAMETER Os
		Specifies the Windows version for which to download the language files.
		This parameter is mandatory and accepts either "10" or "11".
	.PARAMETER Build
		Allows specification of a particular Windows build number for which to download language files.
		If not provided, it defaults to the current system's LCU version.
		The build number should match the pattern ^\d{5}\.\d+$ (e.g., "26100.994").
	.PARAMETER Language
		Defines the language code for which to download files.
		This parameter is mandatory and should be in the format of two lowercase letters,
		a hyphen, and two more lowercase letters (e.g., "en-us" for English-US).
	.PARAMETER UUPUrls
		An optional array of strings representing the URLs of UUP dump sites to use for downloading.
		It defaults to @('www.uupdump.net', 'www.uupdump.cn').
	.PARAMETER ESDToCAB
		A switch parameter that, when specified, triggers the conversion of downloaded ESD files to CAB format.
	.PARAMETER RemoveESD
		A switch parameter that, when specified along with ESDToCAB, removes the original ESD files after successful conversion to CAB.
	.EXAMPLE
		# Example 1: Download language files for Windows 11 using current system build
		Download-LanguageCAB -FolderPath "C:\LanguageFiles" -Os "11" -Language "en-us"
	.EXAMPLE
		# Example 2: Download and convert ESD files to CAB for Windows 11 with specific build
		Download-LanguageCAB -FolderPath "C:\LanguageFiles" -Os "11" -Build "26100.994" -Language "de-de" -ESDToCAB
	.EXAMPLE
		# Example 3: Download, convert to CAB, and remove original ESD files
		Download-LanguageCAB -FolderPath "C:\LanguageFiles" -Os "10" -Language "ja-jp" -ESDToCAB -RemoveESD
	.EXAMPLE
		# Example 4: Use a custom UUP dump site for downloading
		Download-LanguageCAB -FolderPath "C:\LanguageFiles" -Os "11" -Language "es-es" -UUPUrls @('custom.uupdump.site')
	.NOTES
		Author: Harze2k
		Last Updated: May 10, 2025
		Version: 1.2 (Fixed some bugs and made sure it worked properly.)
			-Added some sample output.
		Sample output:
		[2025-05-10 06:39:17.932][INFO] Creating folder: C:\Temp\LanguageCAB\22631.5189
		[2025-05-10 06:39:18.467][SUCCESS] Got webResponse from www.uupdump.net.
		[2025-05-10 06:39:18.469][DEBUG] Using UpdateId: dcec0074-275a-494b-bc14-fa62bc11828b
		[2025-05-10 06:39:21.924][SUCCESS] Got links response from www.uupdump.net.
		[2025-05-10 06:39:21.924][INFO] ========================================================
		[2025-05-10 06:39:21.925][INFO] Start of finding and downloading language cab/esd files.
		[2025-05-10 06:39:21.925][INFO] ========================================================
		[2025-05-10 06:39:22.930][SUCCESS] Successfully downloaded Microsoft-Windows-Client-LanguagePack-Package_sv-se-amd64-sv-se-22631.5189.esd with HttpClient.
		[2025-05-10 06:39:23.485][SUCCESS] Successfully downloaded Microsoft-Windows-LanguageFeatures-Basic-sv-se-Package-amd64-22631.5189.cab with HttpClient.
		[2025-05-10 06:39:23.966][SUCCESS] Successfully downloaded Microsoft-Windows-LanguageFeatures-Handwriting-sv-se-Package-amd64-22631.5189.cab with HttpClient.
		[2025-05-10 06:39:24.196][SUCCESS] Successfully downloaded Microsoft-Windows-LanguageFeatures-OCR-sv-se-Package-amd64-22631.5189.cab with HttpClient.
		[2025-05-10 06:39:24.662][SUCCESS] Successfully downloaded Microsoft-Windows-LanguageFeatures-TextToSpeech-sv-se-Package-amd64-22631.5189.cab with HttpClient.
		[2025-05-10 06:39:24.663][DEBUG] Downloading ESD2CAB cmd tool.
		[2025-05-10 06:39:31.308][SUCCESS] Successfully extracted esd file Microsoft-Windows-Client-LanguagePack-Package_sv-se-amd64-sv-se-22631.5189.esd with ImageX.
		[2025-05-10 06:39:31.309][INFO] Next step (converting) will take some time..
		[2025-05-10 06:40:47.103][SUCCESS] Successfully converted all files in Microsoft-Windows-Client-LanguagePack-Package_sv-se-amd64-sv-se-22631.5189.esd with SxSExpand.
		[2025-05-10 06:40:47.104][INFO] Next step (adding to cab archive) will take some time..
		[2025-05-10 06:41:29.805][SUCCESS] Successfully added all files to Microsoft-Windows-Client-LanguagePack-Package_sv-se-amd64-sv-se-22631.5189.cab with Cabarc.
		[2025-05-10 06:41:29.806][INFO] ==========================================================
		[2025-05-10 06:41:29.806][INFO] Successfully expanded, converted and created the cab file.
		[2025-05-10 06:41:29.806][INFO] ==========================================================
		[2025-05-10 06:41:35.714][SUCCESS] All downloads and conversions completed.
#>
	[CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)][ValidateNotNullOrEmpty()][string]$FolderPath,
        [Parameter(Mandatory)][ValidateSet("10", "11")][string]$Os,
        [ValidatePattern("^\d{5}\.\d+$")][string]$Build = ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name 'CurrentBuild').CurrentBuild) + '.' + ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name 'UBR').UBR),
        [Parameter(Mandatory)][ValidatePattern("^[a-z]{2}-[a-z]{2}$")][string]$Language,
        [ValidateNotNullOrEmpty()][string[]]$UUPUrls = @('www.uupdump.net', 'www.uupdump.cn'),
        [switch]$ESDToCAB,
        [switch]$RemoveESD
    )
    begin {
        function Download-File {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DownloadLink,
                [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OutputFile,
                [Parameter(Mandatory)][string]$FileName
            )
            if ([string]::IsNullOrEmpty($fileName)) {
                [string]$fileName = $(Split-Path $outputFile -Leaf)
            }
            try {
                $httpClient = Get-RandomHeader -GetHTTPClient -ErrorAction Stop
                [System.IO.File]::WriteAllBytes($outputFile, (($httpClient.SendAsync((New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Get, $downloadLink)))).Result).Content.ReadAsByteArrayAsync().Result)
                $httpClient.Dispose()
                New-Log "Successfully downloaded $fileName with HttpClient." -Level SUCCESS
                return $true
            }
            catch {
                New-Log "Failed to download $fileName with HttpClient." -Level ERROR
                return $false
            }
        }
        function Convert-EsdToCab {
            param (
                [string]$EsdFile,
                [string]$WorkingDir,
                [string]$ImageX,
                [string]$Cabarc,
                [string]$Sxs,
                [switch]$RemoveESD
            )
            $pack = [System.IO.Path]::GetFileNameWithoutExtension($EsdFile)
            $cabFile = Join-Path $WorkingDir "$pack.cab"
            if (Test-Path $cabFile) {
                New-Log "$cabFile already exists, skipping."
                return
            }
            $tempDir = Join-Path "C:\" "temp$([System.IO.Path]::GetRandomFileName())"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            try {
                & $ImageX /APPLY $EsdFile 1 $tempDir /NOACL ALL /NOTADMIN /TEMP $env:TEMP | Out-Null
                New-Log "Successfully extracted esd file $EsdFile with ImageX." -Level SUCCESS
                New-Log "Next step (converting) will take some time.."
            }
            catch {
                New-Log "============================================================"
                New-Log "Extracting with ImageX had an error." -Level ERROR
                New-Log "============================================================"
            }
            Push-Location $tempDir
            New-Item -ItemType Directory -Path "_sxs" -Force | Out-Null
            Get-ChildItem -Filter "*.manifest" | ForEach-Object {
                try {
                    & $Sxs $_.Name "_sxs\$($_.Name)" | Out-Null
                }
                catch {
                    [bool]$sxSExpandError = $true
                    New-Log "============================================================"
                    New-Log "Converting with SxSExpand had an error." -Level ERROR
                    New-Log "============================================================"
                }
            }
            if (!($sxSExpandError)) {
                New-Log "Successfully converted all files in $EsdFile with SxSExpand." -Level SUCCESS
                New-Log "Next step (adding to cab archive) will take some time.."
            }
            if (Test-Path "_sxs\*.manifest") {
                Move-Item "_sxs\*" . -Force
            }
            Remove-Item "_sxs" -Recurse -Force
            try {
                & $Cabarc -m LZX:21 -r -p N "$cabFile" *.* | Out-Null
            }
            catch {
                [bool]$cabarcError = $true
                New-Log "============================================================"
                New-Log "Adding to cab archive with Cabarc had an error." -Level ERROR
                New-Log "============================================================"
            }
            if (!($cabarcError)) {
                New-Log "Successfully added all files to $pack.cab with Cabarc." -Level SUCCESS
            }
            if ($LASTEXITCODE -eq 0) {
                New-Log "=========================================================="
                New-Log "Successfully expanded, converted and created the cab file."
                New-Log "=========================================================="
            }
            Set-Location -Path $WorkingDir
            Start-Sleep 5
            Remove-Item $tempDir -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            if ($RemoveESD.IsPresent) {
                Remove-Item -Path ".\$pack.esd" -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            }
        }
        ##############################################################################################################################
        #Calling custom logging function from: https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1#
        ##############################################################################################################################
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
        ################################################################################################################################################
        #Calling custom Get-RandomHeader function from: https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/Get-RandomHeader.ps1#
        ################################################################################################################################################
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/Get-RandomHeader.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
        try {
            if (-not (Test-Path "$folderPath\$build")) {
                New-Log "Creating folder: $folderPath\$build"
                New-Item -Path "$folderPath\$build" -ItemType Directory -Force | Out-Null
            }
        }
        catch {
            New-Log "Failed to create or access folder: "$folderPath\$build"." -Level ERROR
            return
        }
    }
    process {
        $lang = ($Language -split '-')[0]
        $headers = Get-RandomHeader
        foreach ($UUPUrl in $UUPUrls) {
            try {
                $WebResponse = (Invoke-WebRequest -Uri "https://$UUPUrl/known.php?q=windows+$($Os)+$($Build)" -UseBasicParsing -MaximumRedirection 1 -Method GET -Headers $headers -ErrorAction Stop).Links
                if ($WebResponse.Count -eq 0 -or [string]::IsNullOrEmpty($WebResponse)) {
                    continue
                }
                New-Log "Got webResponse from $UUPUrl." -Level SUCCESS
                Break
            }
            catch {
                New-Log "Failed to get WebResponse from $UUPUrl." -Level ERROR
            }
        }
        if ($null -eq $WebResponse -or $WebResponse.Count -eq 0) {
            New-Log "Failed to get WebResponse. Aborting." -Level WARNING
            return
        }
        foreach ($UUPUrl in $UUPUrls) {
            try {
                $UpdateID = (($WebResponse | Where-Object { $_.href -match "(./selectlang.php\?id|selectlang.php\?id)" }).href).split("=")[1]
                New-Log "Using UpdateId: $UpdateID" -Level DEBUG
                $Links = (Invoke-WebRequest -Uri "https://$UUPUrl/get.php?id=$UpdateID&pack=$Language&edition=core" -UseBasicParsing -MaximumRedirection 1 -Method GET -Headers $headers -ErrorAction Stop).Links
                if ($Links.Count -eq 0 -or [string]::IsNullOrEmpty($Links)) {
                    continue
                }
                New-Log "Got links response from $UUPUrl." -Level SUCCESS
                Break
            }
            catch {
                New-Log "Failed to get download links from $UUPUrl." -Level ERROR
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
            foreach ($link in ($Links.outerHTML -match ".*(LanguagePack|LanguageFeatures).*$lang-.*")) {
                $URL = $link.Split('"')[1]
                $Filename = $link.Split('>')[1].Split('<')[0] -replace '\s', ''
                $FilenameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($Filename)
                $FileExtension = [IO.Path]::GetExtension($Filename)
                $NewFilename = "$FilenameWithoutExtension-$Build$FileExtension"
                $OutFile = Join-Path "$folderPath\$build" $NewFilename
                if (!(Test-Path $OutFile)) {
                    $downloadLinks = Download-File -DownloadLink $URL -OutputFile $OutFile -FileName $NewFilename
                }
                else {
                    New-Log "$NewFilename was already downloaded, skipping." -Level DEBUG
                }
            }
        }
        else {
            New-Log "No links found to download. Aborting." -Level WARNING
            return
        }
        if (!(Test-Path "$folderPath\$build\esd2cab_CLI.cmd")) {
            New-Log "Downloading ESD2CAB cmd tool." -Level DEBUG
            try {
                Invoke-WebRequest -Uri 'https://github.com/abbodi1406/WHD/raw/master/scripts/ESD2CAB-CAB2ESD-2.zip' -UseBasicParsing -MaximumRedirection 1 -Method GET -OutFile "$folderPath\$build\ESD2CAB.zip" -ErrorAction Stop
                Start-Sleep 1
                Expand-Archive -Path "$folderPath\$build\ESD2CAB.zip" -DestinationPath $folderPath\$build -Force -ErrorAction Stop
            }
            catch {
                New-Log "Failed to download ESD2CAB cmd tool." -Level ERROR
                $downloadLinks = $false
            }
        }
        if ($ESDToCAB -and $downloadLinks) {
            Set-Location -Path "$folderPath\$build"
            $bits = if ($env:PROCESSOR_ARCHITECTURE -eq "AMD64") {
                "x64"
            }
            else {
                "x86"
            }
            $requiredFiles = @("image$bits.exe", "cabarc.exe", "SxSExpand.exe")
            foreach ($file in $requiredFiles) {
                if (-not (Test-Path "bin\$file")) {
                    New-Log -Message "$file is not detected." -Level ERROR
                    return
                }
            }
            $ImageX = Resolve-Path "bin\image$bits.exe"
            $Cabarc = Resolve-Path "bin\cabarc.exe"
            $Sxs = Resolve-Path "bin\SxSExpand.exe"
            $esdFile = Get-ChildItem -Path . -Include "*Windows*LanguagePack*$language*.esd" -Recurse
            if ([string]::IsNullOrEmpty($esdFile)) {
                New-Log -Message "No .esd file detected." -Level WARNING
                return
            }
            if ($RemoveESD.IsPresent) {
                Convert-EsdToCab -EsdFile $esdFile.Name -WorkingDir "$folderPath\$build" -ImageX $imageX -Cabarc $cabarc -Sxs $sxs -RemoveESD
            }
            else {
                Convert-EsdToCab -EsdFile $esdFile.Name -WorkingDir "$folderPath\$build" -ImageX $imageX -Cabarc $cabarc -Sxs $sxs
            }
            New-Log -Message "All downloads and conversions completed." -Level SUCCESS
        }
        elseif ($ESDToCAB -and !($downloadLinks)) {
            New-Log "Either the download of the CABs failed or the esd to cab tool failed to download. Will abort." -Level WARNING
        }
    }
}
#$currentBuild = ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name 'CurrentBuild').CurrentBuild) + '.' + ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name 'UBR').UBR)
#Download-LanguageCAB -FolderPath "C:\Temp\LanguageCAB" -Os "11" -Language "sv-se" -RemoveESD -ESDToCAB -UUPUrls @('www.uupdump.net', 'www.uupdump.cn') -Build $currentBuild