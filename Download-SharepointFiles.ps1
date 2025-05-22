<#
.SYNOPSIS
    SharePoint Online File Management Script
.DESCRIPTION
    This script provides comprehensive functionality for interacting with SharePoint Online
    document libraries through Microsoft Graph API. It includes functions for authentication,
    file listing, and downloading with robust error handling and retry logic.
.NOTES
    Author:		Harze2k
    Date:   	2025-05-23
	Version:	1.3 (Public release.)
    Purpose:    Automated SharePoint file management via Microsoft Graph API

    Prerequisites:
    - PowerShell 5.1 or higher
    - Microsoft Graph API permissions (Sites.Read.All or Files.Read.All)
    - Valid authentication token or certificate

	Example setting up the enconded command parameter:

	$sb = {
		Import-Module "$env:SystemRoot\System32\WindowsPowerShell\v1.0\Modules\PKI\pki.psd1" -Force -ea SilentlyContinue
		Import-Module "$env:SystemRoot\System32\WindowsPowerShell\v1.0\Modules\MSAL.PS\4.37.0.0\MSAL.PS.psd1" -Force -ea SilentlyContinue
		$certThumbprint = '' # Add this
		$pfxFilePath = "C:\path\to\cert\cert.pfx"
		$pfxPassword = ConvertTo-SecureString -String "password_used_on_pfx_cert" -AsPlainText -Force
		$clientId = '' # Add this
		$tenantId = '' # Add this
		if (!(Get-ChildItem -LiteralPath "Cert:\CurrentUser\My" | Where-Object { $_.Thumbprint -eq $certThumbprint })) {
			Import-PfxCertificate -FilePath $pfxFilePath -CertStoreLocation Cert:\CurrentUser\My -Password $pfxPassword -Confirm:$false -ea SilentlyContinue | Out-Null
		}
		if (!(Get-ChildItem -LiteralPath "Cert:\LocalMachine\My" | Where-Object { $_.Thumbprint -eq $certThumbprint })) {
			Import-PfxCertificate -FilePath $pfxFilePath -CertStoreLocation Cert:\LocalMachine\My -Password $pfxPassword -Confirm:$false -ea SilentlyContinue | Out-Null
		}
		$cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $certThumbprint }
		$token = Get-MsalToken -ClientId $clientId -TenantId $tenantId -Scopes "https://graph.microsoft.com/.default" -ClientCertificate $cert
		$token | ConvertTo-Json
	}
	$bytes = [System.Text.Encoding]::Unicode.GetBytes($sb)
	$encodedCommand = [Convert]::ToBase64String($bytes)
	$encodedCommand # Use this as input.

.LINK
    https://docs.microsoft.com/en-us/graph/api/resources/driveitem
#>
#region SharePoint Drive Functions
function Get-SharePointDriveInfo {
	<#
    .SYNOPSIS
        Retrieves Site ID and Drive ID for a SharePoint site and document library.
    .DESCRIPTION
        This function connects to Microsoft Graph API to obtain the Site ID and Drive ID
        for a specified SharePoint site and document library. It handles common naming
        variations like "Shared Documents" vs "Documents".
    .PARAMETER SiteUrl
        The SharePoint site URL in the format "hostname:/sites/sitename" or "hostname:/teams/teamname".
        Example: "contoso.sharepoint.com:/sites/MySite"
    .PARAMETER Token
        A valid Microsoft Graph API access token with Sites.Read.All or Files.Read.All permissions.
    .PARAMETER FileArea
        The display name of the Document Library (e.g., "Shared Documents", "Documents", "MyCustomLibrary").
        The function automatically handles "Shared Documents" vs "Documents" naming variations.
    .OUTPUTS
        PSCustomObject with properties:
        - SiteId: The GUID of the SharePoint site
        - DriveId: The GUID of the document library
        - DriveName: The actual name of the drive as returned by the API
    .EXAMPLE
        $driveInfo = Get-SharePointDriveInfo -SiteUrl "contoso.sharepoint.com:/sites/IT" -Token $token -FileArea "Shared Documents"
    .EXAMPLE
        $driveInfo = Get-SharePointDriveInfo -SiteUrl "contoso.sharepoint.com:/teams/Marketing" -Token $token -FileArea "Project Files"
    .NOTES
        The function implements intelligent drive matching to handle common SharePoint naming inconsistencies.
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SiteUrl,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Token,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$FileArea
	)
	$headers = @{
		"Authorization" = "Bearer $Token"
		"Accept"        = "application/json"
		"Content-Type"  = "application/json"
	}
	try {
		# Construct site URL for Graph API
		$graphSiteUrlPart = $SiteUrl.Replace('https://', '')
		$siteRequestUri = "https://graph.microsoft.com/v1.0/sites/$graphSiteUrlPart"
		New-Log "Requesting Site ID from: $siteRequestUri" -Level VERBOSE
		# Get site information
		$site = Invoke-RestMethod -Uri $siteRequestUri -Headers $headers -Method Get -UseBasicParsing -ErrorAction Stop -Verbose:$false
		$siteId = $site.id
		if (-not $siteId) {
			New-Log "Failed to retrieve Site ID" -Level ERROR
			return $null
		}
		New-Log "Site ID: $siteId" -Level VERBOSE
		# Get drives (document libraries)
		$drivesRequestUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
		$drivesResponse = Invoke-RestMethod -Uri $drivesRequestUri -Headers $headers -Method Get -UseBasicParsing -ErrorAction Stop -Verbose:$false
		# Drive matching logic to handle common naming variations
		$driveMapping = @{
			'Shared Documents' = @('Shared Documents', 'Documents', 'Shared%20Documents')
			'Documents'        = @('Documents', 'Shared Documents')
		}
		$searchPatterns = if ($driveMapping.ContainsKey($FileArea)) {
			$driveMapping[$FileArea]
		}
		else {
			@($FileArea)
		}
		$foundDrive = $null
		foreach ($pattern in $searchPatterns) {
			$foundDrive = $drivesResponse.value | Where-Object { $_.name -eq $pattern } | Select-Object -First 1
			if ($foundDrive) {
				break
			}
		}
		if ($foundDrive) {
			$driveId = $foundDrive.Id
			New-Log "Drive ID for '$($foundDrive.name)': $driveId" -Level SUCCESS
			return @{
				SiteId    = $siteId
				DriveId   = $driveId
				DriveName = $foundDrive.name
			}
		}
		else {
			$availableDrives = ($drivesResponse.value | ForEach-Object { "'$($_.name)'" }) -join ", "
			New-Log "Drive '$FileArea' not found. Available drives: $availableDrives" -Level ERROR
			return $null
		}
	}
	catch {
		New-Log "Failed to get Site/Drive info." -Level ERROR
		return $null
	}
}
#endregion SharePoint Drive Functions
#region File Retrieval Functions
function Get-SharePointFiles {
	<#
    .SYNOPSIS
        Retrieves a list of files and folders from a SharePoint Online document library.
    .DESCRIPTION
        This function connects to Microsoft Graph API to list items within a SharePoint folder.
        It handles pagination for large folders and provides detailed file information including
        size, modification dates, and download URLs.
    .PARAMETER Token
        A valid Microsoft Graph API access token with Sites.Read.All or Files.Read.All permissions.
    .PARAMETER SiteURL
        The SharePoint site URL in the format "hostname:/sites/sitename".
        Example: "contoso.sharepoint.com:/sites/ProjectAlpha"
    .PARAMETER FileArea
        The display name of the Document Library (e.g., "Shared Documents", "Documents").
    .PARAMETER Folder
        The path to the folder within the Document Library. Use forward slashes (/).
        If omitted, items from the root of the Document Library are listed.
        Example: "General/Reports/2024"
    .PARAMETER PageSize
        The number of items to retrieve per page. Default is 200.
        Larger values may improve performance for folders with many items.
    .OUTPUTS
        Array of PSCustomObject with properties:
        - name: File or folder name
        - id: Unique identifier
        - ItemType: "File" or "Folder"
        - IsFolder: Boolean indicating if item is a folder
        - IsFile: Boolean indicating if item is a file
        - FileExtension: File extension (lowercase, null for folders)
        - CreatedBy: Display name of creator
        - ModifiedBy: Display name of last modifier
        - LastModified: DateTime object in local time
        - SizeBytes: Size in bytes
        - SizeMB: Size in megabytes (rounded to 2 decimals)
        - SizeDisplay: Human-readable size (B, KB, MB, GB)
        - webUrl: SharePoint web URL
        - @microsoft.graph.downloadUrl: Direct download URL (valid for limited time)
    .EXAMPLE
        # Get all files from root of document library
        $files = Get-SharePointFiles -Token $token -SiteURL "contoso.sharepoint.com:/sites/IT" -FileArea "Shared Documents"
    .EXAMPLE
        # Get files from specific folder
        $files = Get-SharePointFiles -Token $token -SiteURL "contoso.sharepoint.com:/sites/IT" -FileArea "Documents" -Folder "Policies/2024"
    .EXAMPLE
        # Get files with custom page size
        $files = Get-SharePointFiles -Token $token -SiteURL "contoso.sharepoint.com:/sites/IT" -FileArea "Documents" -PageSize 500
    .EXAMPLE
        # Filter results for Excel files modified in last 7 days
        $recentExcel = Get-SharePointFiles @params | Where-Object {
            $_.FileExtension -eq 'xlsx' -and
            $_.LastModified -gt (Get-Date).AddDays(-7)
        }
    .NOTES
        - Download URLs are temporary and typically expire after a few hours
        - The function automatically handles pagination for large folders
        - File sizes are provided in multiple formats for convenience
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Token,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SiteURL,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$FileArea,
		[Parameter()][string]$Folder = "",
		[Parameter()][ValidateRange(1, 1000)][int]$PageSize = 200
	)
	# Get drive information
	$driveInfo = Get-SharePointDriveInfo -SiteUrl $SiteURL -Token $Token -FileArea $FileArea
	if (-not $driveInfo) {
		return $null
	}
	$headers = @{
		"Authorization" = "Bearer $Token"
		"Accept"        = "application/json"
		"Prefer"        = "odata.maxpagesize=$PageSize"
	}
	$allItems = [System.Collections.Generic.List[PSObject]]::new()
	try {
		# Build the request URL
		$folderPathSegment = if ($Folder) {
			$trimmedFolder = $Folder.Trim('/')
			"root:/${trimmedFolder}:/children"
		}
		else {
			"root/children"
		}
		$baseUrl = "https://graph.microsoft.com/v1.0/sites/$($driveInfo.SiteId)/drives/$($driveInfo.DriveId)/$folderPathSegment"
		$listUrl = $baseUrl
		$pageCount = 0
		# Retrieve all pages
		do {
			$pageCount++
			New-Log "Fetching page $pageCount from: $listUrl" -Level VERBOSE
			$response = Invoke-RestMethod -Uri $listUrl -Headers $headers -Method Get -UseBasicParsing -ErrorAction Stop -Verbose:$false
			if ($response.value) {
				$response.value | ForEach-Object { $allItems.Add($_) }
				New-Log "Retrieved $($response.value.Count) items in page $pageCount" -Level VERBOSE
			}
			$listUrl = $response.'@odata.nextLink'
		} while ($listUrl)
		New-Log "Total items retrieved: $($allItems.Count)" -Level SUCCESS
		# Return enhanced output with calculated properties
		return $allItems | Select-Object `
			name,
		id,
		@{ Name = 'ItemType'; Expression = { if ($_.folder) { 'Folder' } else { 'File' } } },
		@{ Name = 'IsFolder'; Expression = { $null -ne $_.folder } },
		@{ Name = 'IsFile'; Expression = { $null -ne $_.file } },
		@{ Name = 'FileExtension'; Expression = {
				if ($_.file -and $_.name -match '\.([^.]+)$') { $matches[1].ToLower() } else { $null }
			}
		},
		@{ Name = 'CreatedBy'; Expression = { $_.createdBy.user.displayName } },
		@{ Name = 'ModifiedBy'; Expression = { $_.lastModifiedBy.user.displayName } },
		createdDateTime,
		lastModifiedDateTime,
		@{ Name = 'LastModified'; Expression = {
				if ($_.lastModifiedDateTime) {
					try { [DateTime]::Parse($_.lastModifiedDateTime).ToLocalTime() }
					catch { $null }
				}
				else { $null }
			}
		},
		webUrl,
		'@microsoft.graph.downloadUrl',
		@{ Name = 'SizeBytes'; Expression = { $_.size } },
		@{ Name = 'SizeMB'; Expression = {
				if ($_.size) { [Math]::Round($_.size / 1MB, 2) } else { 0 }
			}
		},
		@{ Name = 'SizeDisplay'; Expression = {
				if ($_.size) {
					if ($_.size -lt 1KB) { "$($_.size) B" }
					elseif ($_.size -lt 1MB) { "{0:N2} KB" -f ($_.size / 1KB) }
					elseif ($_.size -lt 1GB) { "{0:N2} MB" -f ($_.size / 1MB) }
					else { "{0:N2} GB" -f ($_.size / 1GB) }
				}
				else { "N/A" }
			}
		},
		@{ Name = 'MimeType'; Expression = { $_.file.mimeType } },
		@{ Name = 'Path'; Expression = {
				if ($Folder) { "$Folder/$($_.name)" } else { $_.name }
			}
		}
	}
	catch {
		New-Log "Failed to retrieve files" -Level ERROR
		return $null
	}
}
#endregion File Retrieval Functions
#region Download Functions
function Download-SharePointFile {
	<#
    .SYNOPSIS
        Downloads a file from SharePoint using Microsoft Graph download URL.
    .DESCRIPTION
        This function downloads files from SharePoint using the pre-authenticated download URL
        obtained from Get-SharePointFiles. It uses HttpClient for efficient downloading and
        includes error handling and optional overwrite functionality.
    .PARAMETER DownloadUrl
        The direct download URL from the '@microsoft.graph.downloadUrl' property.
        This URL is pre-authenticated and typically expires after a few hours.
    .PARAMETER FileName
        The name to save the file as on the local system.
    .PARAMETER DestinationPath
        The local directory where the file should be saved.
        Defaults to the current working directory.
    .PARAMETER Force
        If specified, overwrites existing files without prompting.
        If not specified and file exists, the function will display a warning and skip.
    .OUTPUTS
        PSCustomObject with properties:
        - FullPath: The complete local path of the downloaded file
        - DownloadUrl: The URL used for download (for reference)
    .EXAMPLE
        # Simple download
        $file = $sharePointFiles | Where-Object {$_.Name -eq "Report.xlsx"} | Select-Object -First 1
        Download-SharePointFile -DownloadUrl $file.'@microsoft.graph.downloadUrl' -FileName $file.Name
    .EXAMPLE
        # Download with custom destination
        Download-SharePointFile -DownloadUrl $url -FileName "Budget.xlsx" -DestinationPath "C:\Reports"
    .EXAMPLE
        # Force overwrite existing file
        Download-SharePointFile -DownloadUrl $url -FileName "Data.csv" -DestinationPath "C:\Data" -Force
    .EXAMPLE
        # Download multiple files
        $files | ForEach-Object {
            Download-SharePointFile -DownloadUrl $_.'@microsoft.graph.downloadUrl' -FileName $_.Name -Force
        }
    .NOTES
        - Uses System.Net.Http.HttpClient for efficient downloading
        - Creates destination directory if it doesn't exist
        - Download URLs expire after a few hours, so download promptly after retrieval
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DownloadUrl,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$FileName,
		[Parameter()][ValidateScript({ Test-Path $_ -PathType Container -IsValid })][string]$DestinationPath = (Get-Location).Path,
		[Parameter()][switch]$Force
	)
	# Create destination directory if it doesn't exist
	if (-not (Test-Path $DestinationPath -PathType Container)) {
		try {
			New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
			New-Log "Created destination directory: $DestinationPath" -Level SUCCESS
		}
		catch {
			New-Log "Failed to create destination directory." -Level ERROR
			return $null
		}
	}
	$fullPath = Join-Path -Path $DestinationPath -ChildPath $FileName
	# Check if file exists and handle accordingly
	if (Test-Path $fullPath -PathType Leaf) {
		if (-not $Force) {
			New-Log "Use Download-SharePointFile with the -Force parameter to overwrite the existing file. Will abort this download." -Level WARNING
			return
		}
		Remove-Item $fullPath -Force | Out-Null
	}
	try {
		New-Log "Downloading '$FileName' to '$fullPath'"
		# Download using HttpClient
		$httpClient = [System.Net.Http.HttpClient]::new()
		[System.IO.File]::WriteAllBytes(
			$fullPath,
			(($httpClient.SendAsync(
					(New-Object System.Net.Http.HttpRequestMessage(
						[System.Net.Http.HttpMethod]::Get,
						$DownloadUrl
					))
				)).Result).Content.ReadAsByteArrayAsync().Result
		)
		$httpClient.Dispose()
		# Verify download
		if (Test-Path $fullPath -PathType Leaf) {
			$fileInfo = Get-Item $fullPath
			New-Log "Successfully downloaded: $($fileInfo.Name) ($('{0:N2} MB' -f ($fileInfo.Length / 1MB)))" -Level SUCCESS
			return [pscustomObject]@{
				FullPath    = $fullPath
				DownloadUrl = $DownloadUrl
			}
		}
		else {
			New-Log "File not found after download" -Level WARNING
			return $null
		}
	}
	catch {
		New-Log "Download failed." -Level ERROR
		# Clean up partial download
		if (Test-Path $fullPath -PathType Leaf) {
			Remove-Item $fullPath -Force -ErrorAction SilentlyContinue
		}
		return $null
	}
}
#endregion Download Functions
#region Authentication Functions
function Get-GraphToken {
	<#
    .SYNOPSIS
        Acquires a Microsoft Graph API access token using an encoded authentication command.
    .DESCRIPTION
        This function decodes and executes an encoded PowerShell command that returns
        a Microsoft Graph API token. It includes retry logic with exponential backoff
        and provides token expiration information when available.
    .PARAMETER EncodedCommand
        A Base64-encoded PowerShell command that returns a token object with an AccessToken property.
        The command should return either a JSON string or an object that can be converted to JSON.
    .PARAMETER MaxRetries
        Maximum number of attempts to acquire the token. Default is 3.
    .PARAMETER RetryDelaySeconds
        Base delay in seconds between retry attempts. The actual delay increases with each attempt
        using exponential backoff (delay * attempt number). Default is 10 seconds.
    .OUTPUTS
        PSCustomObject containing:
        - AccessToken: The bearer token for Graph API calls
        - ExpiresOn: Token expiration time (if provided by the auth command)
        - Additional properties as returned by the authentication command
    .EXAMPLE
        # Get token with default retry settings
        $token = Get-GraphToken -EncodedCommand $encodedCmd
    .EXAMPLE
        # Get token with custom retry configuration
        $token = Get-GraphToken -EncodedCommand $encodedCmd -MaxRetries 5 -RetryDelaySeconds 30
    .EXAMPLE
        # Use token in Graph API call
        $token = Get-GraphToken -EncodedCommand $encodedCmd
        $headers = @{ "Authorization" = "Bearer $($token.AccessToken)" }
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me" -Headers $headers
    .NOTES
        - The encoded command should handle its own authentication logic
        - Token expiration information is logged if available
        - Uses exponential backoff for retry delays
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$EncodedCommand,
		[Parameter()][ValidateRange(1, 10)][int]$MaxRetries = 3,
		[Parameter()][ValidateRange(1, 300)][int]$RetryDelaySeconds = 10
	)
	$attempt = 0
	while ($attempt -lt $MaxRetries) {
		$attempt++
		try {
			New-Log "Token acquisition attempt $attempt of $MaxRetries" -Level VERBOSE
			# Decode and execute the command
			$decodedCommand = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($EncodedCommand))
			$tokenResult = Invoke-Expression $decodedCommand
			# Handle different token response formats
			if ($tokenResult -is [string]) {
				$tokenContainer = $tokenResult | ConvertFrom-Json
			}
			else {
				$tokenContainer = $tokenResult
			}
			if ($tokenContainer -and $tokenContainer.AccessToken) {
				New-Log "Successfully acquired Microsoft Graph API token" -Level SUCCESS
				# Log token expiration if available
				if ($tokenContainer.ExpiresOn) {
					$expiresIn = [DateTime]$tokenContainer.ExpiresOn - [DateTime]::Now
					New-Log "Token expires in: $($expiresIn.TotalMinutes) minutes" -Level VERBOSE
				}
				return $tokenContainer
			}
			New-Log "Token response did not contain AccessToken." -Level WARNING
		}
		catch {
			New-Log "Token acquisition failed." -Level ERROR
			if ($attempt -lt $MaxRetries) {
				$delay = $RetryDelaySeconds * $attempt  # Exponential backoff
				New-Log "Retrying in $delay seconds..."
				Start-Sleep -Seconds $delay
			}
		}
	}
	New-Log "Failed to acquire token after $MaxRetries attempts" -Level WARNING
	return $null
}
#endregion Authentication Functions
#region Main Orchestration Function
function Download-SharepointFiles {
	<#
    .SYNOPSIS
        Main orchestration function for downloading files from SharePoint Online.
    .DESCRIPTION
        This function orchestrates the complete process of connecting to SharePoint,
        retrieving file listings, filtering based on patterns, and downloading files
        with retry logic. It combines all other functions in this module to provide
        a complete SharePoint file download solution.
    .PARAMETER EncodedCommand
        Base64-encoded authentication command for acquiring Graph API token.
    .PARAMETER SiteURL
        The SharePoint site URL (e.g., "contoso.sharepoint.com:/sites/IT").
    .PARAMETER FileArea
        The document library name (e.g., "Shared Documents").
    .PARAMETER FolderPath
        Path within the document library (e.g., "Reports/Monthly").
    .PARAMETER FilePattern
        Regular expression pattern to match files for download.
    .PARAMETER DestinationPath
        Local directory where files will be downloaded. Default is "C:\Temp".
    .PARAMETER DownloadAll
        If specified, downloads all matching files. Otherwise, only the most recent file is downloaded.
    .PARAMETER MaxDownloadRetries
        Maximum number of retry attempts for each file download. Default is 3.
    .PARAMETER DownloadRetryDelay
        Delay in seconds between download retry attempts. Default is 30.
    .OUTPUTS
        Array of PSCustomObject containing:
        - FullPath: Local path of downloaded file
        - DownloadUrl: Source URL used for download
    .EXAMPLE
        # Download latest file matching pattern
        $config = @{
            EncodedCommand = $encodedAuth
            SiteURL = "contoso.sharepoint.com:/sites/Finance"
            FileArea = "Shared Documents"
            FolderPath = "Reports/2024"
            FilePattern = "Monthly.*\.xlsx$"
            DestinationPath = "C:\Reports"
        }
        $results = Download-SharepointFiles @config
    .EXAMPLE
        # Download all Excel files from folder
        Download-SharepointFiles -EncodedCommand $auth -SiteURL $site -FileArea "Documents" `
            -FolderPath "Data" -FilePattern "\.xlsx$" -DownloadAll
    .EXAMPLE
        # Download with custom retry settings
        $results = Download-SharepointFiles @params -MaxDownloadRetries 5 -DownloadRetryDelay 60
    .NOTES
        - Requires valid Graph API permissions
        - Download URLs expire after a few hours
        - Files are downloaded sequentially with retry logic
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$EncodedCommand,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SiteURL,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$FileArea,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$FolderPath,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$FilePattern,
		[Parameter()][ValidateScript({ Test-Path $_ -PathType Container -IsValid })][string]$DestinationPath = "C:\Temp",
		[Parameter()][switch]$DownloadAll,
		[Parameter()][ValidateRange(1, 10)][int]$MaxDownloadRetries = 3,
		[Parameter()][ValidateRange(1, 300)][int]$DownloadRetryDelay = 30
	)
	# Step 1: Acquire authentication token
	$tokenContainer = Get-GraphToken -EncodedCommand $EncodedCommand
	if (-not $tokenContainer) {
		New-Log "Failed to acquire authentication token" -Level ERROR
		return
	}
	# Step 2: Retrieve files from SharePoint
	New-Log "Fetching files from SharePoint: Site='$SiteURL', Library='$FileArea', Folder='$FolderPath'"
	$spFiles = Get-SharePointFiles -Token $tokenContainer.AccessToken -SiteURL $SiteURL -FileArea $FileArea -Folder $FolderPath -ErrorAction Stop
	if (-not $spFiles) {
		New-Log "No files retrieved from SharePoint" -Level ERROR
		return
	}
	# Step 3: Filter files based on pattern
	$matchingFiles = $spFiles | Where-Object {
		$_.IsFile -and $_.Name -match $FilePattern
	}
	if (-not $matchingFiles) {
		New-Log "No files matching pattern '$FilePattern' found" -Level WARNING
		return
	}
	New-Log "Found $($matchingFiles.Count) files matching pattern '$FilePattern'" -Level SUCCESS
	# Step 4: Determine which files to download
	$filesToDownload = if ($DownloadAll) {
		$matchingFiles
	}
	else {
		$matchingFiles | Sort-Object -Property LastModified -Descending | Select-Object -First 1
	}
	# Step 5: Download files with retry logic
	$downloadedFiles = @()
	foreach ($file in $filesToDownload) {
		New-Log "Processing: $($file.Name) (Modified: $($file.LastModified), Size: $($file.SizeDisplay))"
		$downloaded = $false
		$retryCount = 0
		while (-not $downloaded -and $retryCount -lt $MaxDownloadRetries) {
			$retryCount++
			try {
				$result = Download-SharePointFile -DownloadUrl $file.'@microsoft.graph.downloadUrl' -FileName $file.Name -DestinationPath $DestinationPath -Force -ErrorAction Stop
				if ($result) {
					$downloadedFiles += $result
					$downloaded = $true
				}
			}
			catch {
				New-Log "Download attempt $retryCount failed." -Level ERROR
				if ($retryCount -lt $MaxDownloadRetries) {
					New-Log "Waiting $DownloadRetryDelay seconds before retry..." -Level WARNING
					Start-Sleep -Seconds $DownloadRetryDelay
				}
			}
		}
		if (-not $downloaded) {
			New-Log "Failed to download '$($file.Name)' after $MaxDownloadRetries attempts" -Level ERROR
		}
	}
	# Step 6: Summary
	if ($downloadedFiles.Count -gt 0) {
		New-Log "Successfully downloaded $($downloadedFiles.Count) file(s):" -Level SUCCESS
		$downloadedFiles | ForEach-Object {
			New-Log " - $($_.FullPath)" -Level SUCCESS
		}
	}
	else {
		New-Log "No files were successfully downloaded" -Level WARNING
	}
	return $downloadedFiles
}
#endregion Main Orchestration Function
#region Script Execution
#####################################################################################################################################################
### OBS: New-Log Function is needed otherwise remove all New-Log and replace with Write-Host. New-Log is vastly better though, check the link below:#
#####################################################################################################################################################
#Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
<#
#Example:
try {
	# Load required assemblies
	Add-Type -AssemblyName System.Web
	# Configuration
	$config = @{
		EncodedCommand     = '' # A scriptblock that returns a Token Container. See example at the top of the file.
		SiteURL            = "contoso.sharepoint.com:/sites/msteams_3dc3d3"
		FileArea           = "Shared Documents"
		FolderPath         = "ISD/CRM/Field Service Management"
		FilePattern        = 'Email Groups and Members\.xlsx$' # File(s) to match
		DestinationPath    = "C:\Temp"
		MaxDownloadRetries = 3
		DownloadRetryDelay = 30
	}
	# Execute download
	$results = Download-SharepointFiles @config -ErrorAction Stop
	# Return results
	$results
}
catch {
	New-Log "Script execution failed." -Level ERROR
	exit 1
}
#>