#Requires -Version 5.1
<#
.SYNOPSIS
	Downloads a file from a specified URL with support for progress, resume, retries, custom headers, and HttpClient integration.
.DESCRIPTION
	The Download-File function provides a robust mechanism for downloading files over HTTP/HTTPS.
	It handles various input types for the URL (string, URI, Hashtable, PSObject), automatically detects filenames if not provided, and saves the file to a specified path (defaulting to the current location).
	Key features include:
	- Streaming download suitable for large files.
	- Download progress reporting using Write-Progress.
	- Resuming interrupted downloads using the -Resume switch.
	- Automatic retries on transient failures.
	- Support for custom HTTP headers.
	- Option to provide and reuse an existing System.Net.Http.HttpClient instance.
	- Option to ignore SSL/TLS certificate errors (use with caution).
	- Overwriting existing files using the -Force switch.
	- Adherence to -WhatIf and confirmation prompts via SupportsShouldProcess.
	- Automatic buffer size optimization based on file size and available memory.
	- TLS 1.2/1.3 enforced where possible.
	The function outputs a PS Custo mObject detailing the download outcome, including status, file path, size, time taken, speed, and any errors encountered.
.PARAMETER UrlInput
	The URL of the file to download. This can be a simple string, a [System.Uri] object,
	or a Hashtable/PSObject containing a 'Url' property (and optionally 'FileName', 'FilePath', 'Headers', etc. to override parameters).
	This parameter is mandatory and accepts pipeline input.
.PARAMETER FileName
	The desired name for the downloaded file. If omitted, the function attempts to automatically determine the filename from the URL.
	If auto-detection fails or results in a name without an extension, a fallback name (like 'unknown-<guid>.tmp') or a '.tmp' extension might be added.
.PARAMETER FilePath
	The directory path where the downloaded file should be saved. Defaults to the current working directory (Get-Location).
	If the directory does not exist, the function will attempt to create it (subject to -WhatIf/-Confirm).
.PARAMETER HTTPClient
	An optional, pre-configured [System.Net.Http.HttpClient] instance to use for the download.
	If provided, parameters like -TimeoutSeconds, -IgnoreSSLErrors, and default headers might be ignored in favor of the client's configuration.
	Custom headers passed via -Headers will still be added to the specific request if this client is used.
.PARAMETER BufferFactor
	A multiplier (1-10) applied to the dynamically calculated optimal buffer size. Default is determined automatically based on file size (typically 1 for small, 2 for medium, 4 for large files).
	Adjust this only if you have specific performance tuning needs. The buffer size is capped based on available memory.
.PARAMETER TimeoutSeconds
	The timeout duration in seconds for the HTTP request. Defaults to 100 seconds.
	This is ignored if a custom -HTTPClient is provided.
.PARAMETER DisposeClient
	If specified along with a provided -HTTPClient, the function will dispose of the provided HttpClient instance in the 'end' block after all pipeline input has been processed. Use with caution if the client is intended for reuse elsewhere.
.PARAMETER Resume
	If specified, attempts to resume the download if the target file already exists and is partially downloaded.
	The server must support byte range requests (HTTP 206 Partial Content) for resume to work.
.PARAMETER RetryCount
	The number of times to retry the download if it fails. Defaults to 1 (meaning one initial attempt + one retry = 2 total attempts). Set to 0 for no retries.
.PARAMETER RetryDelaySeconds
	The delay in seconds between download retries. Defaults to 5 seconds.
.PARAMETER Headers
	A hashtable containing custom HTTP headers to add to the download request (e.g., @{'Authorization'='Bearer token'; 'X-Custom-ID'='123'}).
.PARAMETER IgnoreSSLErrors
	If specified, bypasses SSL/TLS certificate validation errors. Use with caution, as this can be insecure.
	This is ignored if a custom -HTTPClient is provided (use the client's handler configuration instead).
.PARAMETER Force
	If specified, overwrites the destination file if it already exists, even if -Resume is not used. Suppresses the overwrite confirmation prompt.
.EXAMPLE
	PS C:\> Download-File -UrlInput "https://example.com/largefile.zip" -FilePath "C:\Downloads"
	Downloads the file to C:\Downloads\largefile.zip, showing progress.
.EXAMPLE
	PS C:\> Download-File "https://example.com/document" -FileName "mydoc.pdf" -RetryCount 3
	Downloads the file, renaming it to mydoc.pdf, and retries up to 3 times on failure.
.EXAMPLE
	PS C:\> $headers = @{ "User-Agent"="MyCustomAgent/1.0" }
	PS C:\> Download-File -UrlInput "https://api.example.com/data.json" -Headers $headers -IgnoreSSLErrors
	Downloads data.json using a custom User-Agent and ignoring SSL errors.
.EXAMPLE
	PS C:\> Get-Content "urls.txt" | Download-File -FilePath "D:\Output" -Resume
	Downloads each URL listed in urls.txt into the D:\Output directory, attempting to resume if files exist.
.EXAMPLE
	PS C:\> $client = [System.Net.Http.HttpClient]::new()
	PS C:\> # (Configure $client headers, timeout, handler etc. as needed)
	PS C:\> $downloadTask = @{
	>>   Url = "https://secure.example.com/archive.tar.gz"
	>>   FileName = "backup.tar.gz"
	>>   HTTPClient = $client
	>> }
	PS C:\> $downloadTask | Download-File -DisposeClient
	Downloads using the pre-configured HttpClient instance and disposes of the client afterwards.
.NOTES
    Author: Harze2k
    Date:   2025-04-27 (Updated)
    Version: 1.6 (First public release.)
	- Requires PowerShell 5.1 or later.
	- Depends on the custom 'New-Log' function for logging. Ensure New-Log is available in the scope.
	- Uses System.Net.Http.HttpClient for downloads, which is generally more performant than older methods like WebClient or Invoke-WebRequest for large files.
	- Automatic filename detection relies on the URL path and may not always be perfect. Specify -FileName for guaranteed results.
	- Buffer size optimization attempts to balance throughput and memory usage but is not guaranteed to be optimal in all network/system conditions.
.LINK
	System.Net.Http.HttpClient
	https://docs.microsoft.com/en-us/dotnet/api/system.net.http.httpclient
#>
function Download-File {
	[CmdletBinding(SupportsShouldProcess = $true)]
	param(
		[Parameter(Mandatory, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)][object]$UrlInput,
		[Parameter(Position = 1, ValueFromPipelineByPropertyName = $true)][string]$FileName,
		[Parameter(Position = 2, ValueFromPipelineByPropertyName = $true)][string]$FilePath = (Get-Location).Path,
		[Parameter(Position = 3, ValueFromPipelineByPropertyName = $true)][System.Net.Http.HttpClient]$HTTPClient,
		[Parameter(Position = 4, ValueFromPipelineByPropertyName = $true)][ValidateRange(1, 10)][int]$BufferFactor = 0, # Default 0 means auto-calculate based on file size later
		[Parameter(ValueFromPipelineByPropertyName = $true)][ValidateRange(1, [int]::MaxValue)][int]$TimeoutSeconds = 100,
		[Parameter(ValueFromPipelineByPropertyName = $true)][switch]$DisposeClient,
		[Parameter(ValueFromPipelineByPropertyName = $true)][switch]$Resume,
		[Parameter(ValueFromPipelineByPropertyName = $true)][ValidateRange(0, [int]::MaxValue)][int]$RetryCount = 1,
		[Parameter(ValueFromPipelineByPropertyName = $true)][ValidateRange(0, [int]::MaxValue)][int]$RetryDelaySeconds = 5,
		[Parameter(ValueFromPipelineByPropertyName = $true)][hashtable]$Headers,
		[Parameter(ValueFromPipelineByPropertyName = $true)][switch]$IgnoreSSLErrors,
		[Parameter(ValueFromPipelineByPropertyName = $true)][switch]$Force
	)
	begin {
		# Helper function to format file sizes nicely
		function Format-FileSize {
			[CmdletBinding()]
			param (
				[Parameter(Mandatory)][long]$Bytes,
				[Parameter()][int]$Precision = 2 # Default to 2 decimal places
			)
			process {
				if ($Bytes -lt 0) { return "N/A" }
				if ($Bytes -eq 0) { return "0 B" }
				$sizes = 'B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB'
				$base = 1024
				$index = [Math]::Floor([Math]::Log($Bytes, $base))
				$index = [Math]::Min($index, $sizes.Count - 1)
				$index = [Math]::Max(0, $index)
				$value = $Bytes / [Math]::Pow($base, $index)
				# Use the Precision parameter in the format string
				$formatString = "{0:N$($Precision)} {1}"
				return $formatString -f $value, $sizes[$index]
			}
		}
		# Helper function to determine a reasonable buffer size
		function Get-OptimalBufferSize {
			[CmdletBinding()]
			param (
				[Parameter(Mandatory)][long]$FileSize, # -1 if unknown
				[Parameter()][int]$Factor = 0 # User override factor (1-10)
			)
			process {
				$KB = 1024L; $MB = 1024L * $KB; $GB = 1024L * $MB
				$smallFactor = 1   # For files <= 100MB
				$mediumFactor = 2  # For files <= 1000MB
				$largeFactor = 4   # For files > 1000MB
				[int]$bufferFactor = if ($Factor -ge 1 -and $Factor -le 10) {
					$Factor
				}
				else {
					if ($FileSize -lt 0) {
						$mediumFactor
					}
					elseif ($FileSize -le (100L * $MB)) {
						$smallFactor
					}
					elseif ($FileSize -le (1000L * $MB)) {
						$mediumFactor
					}
					else {
						$largeFactor
					}
				}
				# Determine base buffer size based on file size
				[long]$baseBufferSize = 64L * $KB # Default
				if ($FileSize -ge 0) {
					if ($FileSize -lt (1L * $MB)) { $baseBufferSize = 16L * $KB }
					elseif ($FileSize -lt (10L * $MB)) { $baseBufferSize = 64L * $KB }
					elseif ($FileSize -lt (100L * $MB)) { $baseBufferSize = 128L * $KB }
					elseif ($FileSize -lt (500L * $MB)) { $baseBufferSize = 256L * $KB }
					elseif ($FileSize -lt (1L * $GB)) { $baseBufferSize = 512L * $KB }
					else { $baseBufferSize = 1L * $MB }
				}
				[long]$availableMemory = 1L * $GB # Default fallback
				[long]$maxBufferCap = 8L * $MB
				try {
					$osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
					if ($null -ne $osInfo -and $null -ne $osInfo.FreePhysicalMemory) {
						$availableMemory = $osInfo.FreePhysicalMemory * $KB
					}
				}
				catch {
					New-Log "Failed getting memory information. Using default assumption for buffer calculation." -Level ERROR
				}
				[long]$ramPercentage = [long]($availableMemory * 0.005) # 0.5% of free RAM
				[long]$maxBufferBasedOnRam = [Math]::Min($maxBufferCap, $ramPercentage)
				[long]$minBuffer = 4L * $KB # Absolute minimum buffer size
				$maxBufferBasedOnRam = [Math]::Max($minBuffer, $maxBufferBasedOnRam) # Ensure RAM cap is not below min
				[long]$calculatedBuffer = $baseBufferSize * $bufferFactor
				$calculatedBuffer = [Math]::Min($calculatedBuffer, $maxBufferBasedOnRam) # Apply RAM cap
				[long]$finalBuffer = [Math]::Max($minBuffer, $calculatedBuffer) # Apply minimum floor
				New-Log "Optimal buffer size: $(Format-FileSize $finalBuffer) (Factor: $bufferFactor, FileSize: $(Format-FileSize $FileSize), RAM Cap: $(Format-FileSize $maxBufferBasedOnRam))" -Level SUCCESS
				return $finalBuffer
			}
		}
		$originalSecurityProtocol = $null
		try {
			$originalSecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol
			$securityProtocols = [System.Net.SecurityProtocolType]::Tls12
			if ([System.Enum]::TryParse('Tls13', [ref]$securityProtocols)) {
				$securityProtocols = $securityProtocols -bor [System.Net.SecurityProtocolType]::Tls13
			}
			[System.Net.ServicePointManager]::SecurityProtocol = $securityProtocols
		}
		catch {
			New-Log "Unable to set preferred TLS 1.2/1.3. Using system defaults." -Level ERROR
		}
		$disposeHttpClientInternally = $false
		$sharedHttpClient = $null # Will hold the client used in the process block
		$providedHttpClient = $HTTPClient # Keep track if one was provided initially for end block disposal
	}
	process {
		$result = [pscustomobject]@{
			Status       = "Pending"
			FileName     = $null
			FileSize     = $null # Formatted string
			FilePath     = $null
			TotalBytes   = 0L    # Raw long value
			TimeTaken    = $null # TimeSpan string
			AverageSpeed = $null # Formatted string/s
			URL          = $null
			Retries      = 0
			ResumeUsed   = $false
			Error        = $null
		}
		$targetUrl = $null
		$targetFileName = $FileName # Use parameter if provided
		$targetFilePath = $FilePath # Use parameter if provided
		$targetHeaders = $Headers   # Use parameter if provided
		try {
			if ($UrlInput -is [uri]) {
				$targetUrl = $UrlInput.AbsoluteUri
			}
			elseif ($UrlInput -is [string]) {
				$targetUrl = $UrlInput
			}
			elseif ($UrlInput -is [hashtable] -or $UrlInput -is [psobject]) {
				$urlProp = $UrlInput.PSObject.Properties | Where-Object { $_.Name -eq 'Url' -or $_.Name -eq 'Uri' } | Select-Object -First 1
				if ($urlProp) {
					$targetUrl = $urlProp.Value
				}
				$targetUrl = $UrlInput.Url
				if (-not $PSBoundParameters.ContainsKey('FileName') -and $UrlInput.PSObject.Properties['FileName']) { $targetFileName = $UrlInput.FileName }
				if (-not $PSBoundParameters.ContainsKey('FilePath') -and $UrlInput.PSObject.Properties['FilePath']) { $targetFilePath = $UrlInput.FilePath }
				if (-not $PSBoundParameters.ContainsKey('Headers') -and $UrlInput.PSObject.Properties['Headers']) { $targetHeaders = $UrlInput.Headers }
			}
			else {
				New-Log "Unsupported UrlInput type: $($UrlInput.GetType().FullName)" -Level WARNING
			}
			if (-not $targetUrl) {
				$result.Status = "Failed"
				$result.Error = "Phase 1 (Input Validation/Param) Failed."
				New-Log $result.Error -Level WARNING
				Write-Output $result
				return
			}
			$result.URL = $targetUrl
			if ([string]::IsNullOrWhiteSpace($targetFileName)) {
				try {
					$uri = [System.Uri]$targetUrl
					$decodedPath = [System.Web.HttpUtility]::UrlDecode($uri.AbsolutePath)
					$fn = [System.IO.Path]::GetFileName($decodedPath)
					if ([string]::IsNullOrWhiteSpace($fn) -or $fn -eq "/") {
						# If no filename in path, create one from host + .download
						$invalidChars = [RegEx]::Escape([System.IO.Path]::GetInvalidFileNameChars() -join '')
						$fallbackName = ($uri.Host -replace "[$invalidChars]", "_") + ".download"
						if ([string]::IsNullOrWhiteSpace($fallbackName)) { $fallbackName = "unknown.download" } # Absolute fallback
						$targetFileName = $fallbackName
						New-Log "Auto-detected FileName: '$targetFileName' (using fallback from host)" -Level SUCCESS
					}
					else {
						$targetFileName = $fn
						New-Log "Auto-detected FileName: '$targetFileName' (from URL path)" -Level SUCCESS
					}
					if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($targetFileName))) {
						$targetFileName = "$targetFileName.tmp"
						New-Log "Auto-detected filename lacked extension. Added '.tmp': '$targetFileName'" -Level WARNING
					}
				}
				catch {
					$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
					$guid = [System.Guid]::NewGuid().ToString().Substring(0, 8)
					$targetFileName = "download-${timestamp}-${guid}.tmp"
					New-Log "FileName auto-detection failed. Using generated name: '$targetFileName'. Error: $($_.Exception.Message)" -Level ERROR
				}
			}
			elseif ([System.IO.Path]::GetExtension($targetFileName) -eq "") {
				$randomPart = [System.Guid]::NewGuid().ToString().Substring(0, 8)
				$itemFileName = "unknown-$randomPart.tmp"
				New-Log "Emergency filename fallback used: '$itemFileName'" -Level WARNING
			}
			if ([string]::IsNullOrWhiteSpace($targetFileName)) {
				$result.Status = "Failed"
				$result.Error = "Phase 2 (Path Setup) Failed for filename '$targetFileName'."
				New-Log $result.Error -Level WARNING
				Write-Output $result
				return
			}
			$result.FileName = $targetFileName
			$progressId = [Math]::Abs($targetFileName.GetHashCode())
			New-Log "Validated Input: URL='$($result.URL)', FileName='$($result.FileName)', FilePath='$targetFilePath'" -Level INFO
		}
		catch {
			$result.Status = "Failed"
			$result.Error = "Phase 1 (Input Validation/Parameter Resolution) Failed: $($_.Exception.Message)"
			New-Log $result.Error -Level ERROR
			Write-Output $result
			return
		}
		$outputFile = $null
		try {
			$resolvedFilePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($targetFilePath)
			if (-not [System.IO.Path]::IsPathRooted($resolvedFilePath)) {
				$resolvedFilePath = Join-Path -Path (Get-Location).Path -ChildPath $resolvedFilePath
			}
			if (-not (Test-Path -Path $resolvedFilePath -PathType Container)) {
				if ($PSCmdlet.ShouldProcess($resolvedFilePath, "Create Directory")) {
					New-Item -Path $resolvedFilePath -ItemType Directory -Force -ErrorAction Stop | Out-Null
				}
			}
			$outputFile = Join-Path -Path $resolvedFilePath -ChildPath $result.FileName
			if ($outputFile) {
				$result.FilePath = $outputFile
				New-Log "Output file path set: $outputFile" -Level VERBOSE
			}
		}
		catch {
			$result.Status = "Failed"
			$result.Error = "Phase 2 (Path Setup) Failed."
			if ($outputFile) { $result.FilePath = $outputFile }
			New-Log $result.Error -Level ERROR
			Write-Output $result
			return
		}
		$currentHttpClient = $null
		$clientHandler = $null
		$disposeHttpClientLocally = $false
		try {
			if ($HTTPClient -eq $null) {
				$clientHandler = [System.Net.Http.HttpClientHandler]::new()
				if ($IgnoreSSLErrors) {
					New-Log "SSL/TLS certificate errors will be ignored for this download." -Level WARNING
					$clientHandler.ServerCertificateCustomValidationCallback = { $true }
				}
				$clientHandler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
				$currentHttpClient = [System.Net.Http.HttpClient]::new($clientHandler, $true)
				$currentHttpClient.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
				if ($targetHeaders) {
					New-Log "Adding custom headers to new HttpClient: $($targetHeaders.Keys -join ', ')" -Level VERBOSE
					foreach ($key in $targetHeaders.Keys) {
						$currentHttpClient.DefaultRequestHeaders.Remove($key) | Out-Null
						$currentHttpClient.DefaultRequestHeaders.TryAddWithoutValidation($key, $targetHeaders[$key]) | Out-Null
					}
				}
				else {
					$userAgent = "PowerShell/Download-File (PSVersion=$($PSVersionTable.PSVersion.Major); Runtime=$([System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription))"
					try { $currentHttpClient.DefaultRequestHeaders.UserAgent.ParseAdd($userAgent) } catch {}
				}
				$disposeHttpClientLocally = $true
			}
			else {
				$currentHttpClient = $HTTPClient
				New-Log "Using provided HttpClient instance."
				if ($IgnoreSSLErrors -and -not $disposeHttpClientInternally) {
					New-Log "Ignoring -IgnoreSSLErrors parameter; using provided HttpClient's handler configuration." -Level WARNING
				}
				if ($PSBoundParameters.ContainsKey('TimeoutSeconds') -and -not $disposeHttpClientInternally) {
					New-Log "Ignoring -TimeoutSeconds parameter; using provided HttpClient's timeout configuration."
				}
				if ($targetHeaders) {
					New-Log "Custom headers from -Headers will be applied to the HttpRequestMessage (using provided client)."
				}
			}
			$sharedHttpClient = $currentHttpClient # Reference the client being used
		}
		catch {
			$result.Status = "Failed"
			$result.Error = "Phase 3 (HttpClient Setup) Failed."
			New-Log $result.Error -Level ERROR
			if ($disposeHttpClientLocally -and $currentHttpClient -ne $null) {
				try { $currentHttpClient.Dispose() } catch {}
			}
			Write-Output $result
			return
		}
		$performDownload = $true
		$existingFileSize = 0L
		$fileMode = [System.IO.FileMode]::Create # Default: Overwrite or Create new
		$fileAccess = [System.IO.FileAccess]::Write
		try {
			if (Test-Path $outputFile -PathType Leaf) {
				$existingFileInfo = Get-Item $outputFile
				$existingFileSizeOnDisk = $existingFileInfo.Length
				if ($Resume) {
					if ($existingFileSizeOnDisk -gt 0) {
						New-Log "Resume requested. Existing file '$($result.FileName)' found with size: $(Format-FileSize $existingFileSizeOnDisk)" -Level DEBUG
						$existingFileSize = $existingFileSizeOnDisk # Start download from this offset
						$result.ResumeUsed = $true
						$fileMode = [System.IO.FileMode]::Append # Append to existing file
					}
					else {
						New-Log "Resume requested, but existing file '$($result.FileName)' is empty. Starting download from scratch." -Level WARNING
						$existingFileSize = 0L
						$fileMode = [System.IO.FileMode]::Create
						$result.ResumeUsed = $false
					}
				}
				else {
					if (-not $Force) {
						if (-not $PSCmdlet.ShouldProcess($outputFile, "Overwrite existing file")) {
							New-Log "Skipping download: Target file '$outputFile' already exists and overwrite was not confirmed (use -Force or -Resume)." -Level WARNING
							$result.Status = "Skipped"
							$result.Error = "Target file exists. Use -Resume or -Force to overwrite."
							$result.TotalBytes = $existingFileSizeOnDisk
							$result.FileSize = Format-FileSize $existingFileSizeOnDisk
							$performDownload = $false
						}
						else {
							New-Log "Proceeding to overwrite existing file '$outputFile' as confirmed." -Level VERBOSE
							$existingFileSize = 0L
							$fileMode = [System.IO.FileMode]::Create
						}
					}
					else {
						New-Log "Overwriting existing file '$outputFile' due to -Force parameter." -Level VERBOSE
						$existingFileSize = 0L
						$fileMode = [System.IO.FileMode]::Create
					}
				}
			}
			else {
				if (-not $PSCmdlet.ShouldProcess($outputFile, "Download file from '$($result.URL)'")) {
					New-Log "Skipping download: Action cancelled by user or -WhatIf." -Level WARNING
					$result.Status = "Skipped"
					$result.Error = "-WhatIf specified or user cancelled download operation."
					$performDownload = $false
				}
			}
			# --- Download Execution (if not skipped) ---
			if ($performDownload) {
				$retriesLeft = $RetryCount
				$downloadSuccess = $false
				$stopwatch = [System.Diagnostics.Stopwatch]::new()
				$currentUrlToDownload = $result.Url # Can change due to redirects
				while (-not $downloadSuccess -and $retriesLeft -ge 0) {
					$currentTry = $RetryCount - $retriesLeft + 1
					$totalAttempts = $RetryCount + 1
					New-Log "Download Attempt $currentTry/$totalAttempts : URL='$currentUrlToDownload', File='$outputFile'" -Level INFO
					$request = $null
					$response = $null
					$contentStream = $null
					$fileStream = $null
					$attemptResumeThisTry = $Resume -and ($existingFileSize -gt 0) # Check if resume is applicable for *this* attempt
					try {
						$request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $currentUrlToDownload)
						if ($attemptResumeThisTry) {
							$request.Headers.Range = [System.Net.Http.Headers.RangeHeaderValue]::new($existingFileSize, $null)
							New-Log "Adding Range header for resume: bytes=$existingFileSize-" -Level VERBOSE
						}
						if ($targetHeaders -and $HTTPClient -ne $null) {
							New-Log "Adding custom headers to HttpRequestMessage (using provided client): $($targetHeaders.Keys -join ', ')" -Level VERBOSE
							foreach ($key in $targetHeaders.Keys) {
								if ($key -ne 'Range' -or !$attemptResumeThisTry) {
									$request.Headers.TryAddWithoutValidation($key, $targetHeaders[$key]) | Out-Null
								}
							}
						}
						$response = $currentHttpClient.SendAsync($request, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
						$actualUrlUsed = $response.RequestMessage.RequestUri.AbsoluteUri
						if ($currentUrlToDownload -ne $actualUrlUsed) {
							New-Log "Request was redirected. Final URL: $actualUrlUsed" -Level VERBOSE
							$currentUrlToDownload = $actualUrlUsed # Use the final URL for potential retries
							$result.URL = $actualUrlUsed # Update result object
						}
						New-Log "Response status code: $($response.StatusCode) ($([int]$response.StatusCode))" -Level VERBOSE
						$effectiveTotalLength = -1L # Total expected size of the file (-1 if unknown)
						$contentLengthHeader = if ($response.Content.Headers.ContentLength.HasValue) { $response.Content.Headers.ContentLength.Value } else { -1L }
						if ($attemptResumeThisTry) {
							if ($response.StatusCode -eq [System.Net.HttpStatusCode]::PartialContent) {
								$contentRange = $response.Content.Headers.ContentRange
								if ($contentRange -ne $null -and $contentRange.HasLength -and $contentRange.From -eq $existingFileSize) {
									$effectiveTotalLength = $contentRange.Length.Value
									New-Log "Resume accepted (206 Partial Content). Continuing download. Expected total size: $(Format-FileSize $effectiveTotalLength)" -Level SUCCESS
									$result.ResumeUsed = $true
								}
								elseif ($contentRange -ne $null -and $contentRange.From -ne $existingFileSize) {
									New-Log "Resume failed: Server returned 206 Partial Content but for unexpected range (From $($contentRange.From) != Requested $existingFileSize). Restarting download from scratch for this attempt." -Level WARNING
									$attemptResumeThisTry = $false
									$existingFileSize = 0L
									$fileMode = [System.IO.FileMode]::Create
									$effectiveTotalLength = $contentLengthHeader
									$result.ResumeUsed = $false
								}
								else {
									New-Log "Resume failed: Server returned 206 Partial Content but Content-Range header was missing or invalid. Restarting download from scratch for this attempt." -Level WARNING
									$attemptResumeThisTry = $false
									$existingFileSize = 0L
									$fileMode = [System.IO.FileMode]::Create
									$effectiveTotalLength = $contentLengthHeader
									$result.ResumeUsed = $false
								}
							}
							else {
								New-Log "Resume failed: Server did not return 206 Partial Content (Status: $($response.StatusCode)). Restarting download from scratch for this attempt." -Level WARNING
								$response.EnsureSuccessStatusCode() | Out-Null # Ensure it's at least a successful code (e.g., 200 OK)
								$attemptResumeThisTry = $false
								$existingFileSize = 0L
								$fileMode = [System.IO.FileMode]::Create
								$effectiveTotalLength = $contentLengthHeader
								$result.ResumeUsed = $false
							}
						}
						else {
							$response.EnsureSuccessStatusCode() | Out-Null
							$effectiveTotalLength = $contentLengthHeader
							if ($effectiveTotalLength -ge 0) {
								New-Log "Starting new download. Expected total size: $(Format-FileSize $effectiveTotalLength)" -Level VERBOSE
							}
							else {
								New-Log "Starting new download. Content-Length header not provided by server." -Level VERBOSE
							}
						}
						$contentStream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
						$bufferSize = Get-OptimalBufferSize -FileSize $effectiveTotalLength -Factor $BufferFactor
						# Use FileStream with appropriate Mode, Access, and Share; buffer size passed for potential OS optimization
						$fileStream = [System.IO.FileStream]::new($outputFile, $fileMode, $fileAccess, [System.IO.FileShare]::Read, $bufferSize)
						New-Log "Opened file stream with Mode '$fileMode', Access '$fileAccess', Buffer hint $bufferSize." -Level VERBOSE
						$buffer = New-Object byte[] $bufferSize
						$totalBytesReadThisSession = 0L
						$lastProgressTime = Get-Date
						$progressUpdateInterval = [TimeSpan]::FromMilliseconds(200) # Update progress every 200ms
						$stopwatch.Restart() # Start timing this attempt
						$bytesRead = 0
						while (($bytesRead = $contentStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
							$fileStream.Write($buffer, 0, $bytesRead)
							$totalBytesReadThisSession += $bytesRead
							$now = Get-Date
							if (($now - $lastProgressTime) -ge $progressUpdateInterval) {
								$currentTotalBytes = $existingFileSize + $totalBytesReadThisSession
								$elapsedSeconds = $stopwatch.Elapsed.TotalSeconds
								$speed = if ($elapsedSeconds -gt 0.01) { $totalBytesReadThisSession / $elapsedSeconds } else { 0 }
								$progressParams = @{
									Activity         = "Downloading $($result.FileName) (Try $currentTry/$totalAttempts)"
									CurrentOperation = "Receiving..."
									Id               = $progressId
								}
								if ($effectiveTotalLength -gt 0) {
									$percent = [math]::Min(100, [math]::Floor(($currentTotalBytes / $effectiveTotalLength) * 100))
									$progressParams.Status = "{0} / {1} ({2}%) - {3}/s" -f (Format-FileSize $currentTotalBytes), (Format-FileSize $effectiveTotalLength), $percent, (Format-FileSize $speed)
									$progressParams.PercentComplete = $percent
								}
								else {
									$progressParams.Status = "{0} downloaded - {1}/s" -f (Format-FileSize $currentTotalBytes), (Format-FileSize $speed)
								}
								Write-Progress @progressParams
								$lastProgressTime = $now
							}
						}
						$stopwatch.Stop()
						$downloadSuccess = $true
						Write-Progress -Activity "Downloading $($result.FileName) (Try $currentTry/$totalAttempts)" -Completed -Id $progressId
						$finalFileSize = $existingFileSize + $totalBytesReadThisSession
						$result.Status = "Completed"
						$result.FileSize = Format-FileSize $finalFileSize
						$result.TotalBytes = $finalFileSize
						$result.TimeTaken = $stopwatch.Elapsed.ToString('hh\:mm\:ss\.fff')
						# Calculate average speed based on bytes downloaded *this session* and time taken *this session*
						$avgSpeed = if ($stopwatch.Elapsed.TotalSeconds -gt 0.01) { $totalBytesReadThisSession / $stopwatch.Elapsed.TotalSeconds } else { 0 }
						$result.AverageSpeed = "$(Format-FileSize -Bytes $avgSpeed -Precision 3)/s"
						$result.Retries = $RetryCount - $retriesLeft # Number of retries used
						$result.Error = $null # Clear any previous retry errors
						New-Log "Download completed successfully for '$($result.FileName)'." -Level SUCCESS
					} # End Try block for download attempt
					catch {
						$stopwatch.Stop()
						try { Write-Progress -Activity "Downloading $($result.FileName) (Try $currentTry/$totalAttempts)" -Completed -Id $progressId } catch {}
						$errorMessage = $_.Exception.Message
						if ($_.Exception.InnerException) { $errorMessage += " | Inner: $($_.Exception.InnerException.Message)" }
						New-Log "Error during download attempt $currentTry for '$($result.FileName)': $errorMessage" -Level WARNING
						$result.Error = "Attempt $currentTry failed: $errorMessage" # Store last error
						$retriesLeft--
						if ($retriesLeft -ge 0) {
							New-Log "Retrying download for '$($result.FileName)' in $RetryDelaySeconds second(s)... ($retriesLeft retries remaining)" -Level DEBUG
							Start-Sleep -Seconds $RetryDelaySeconds
							if ($Resume) {
								try {
									if ($fileStream -ne $null) { try { $fileStream.Dispose() } catch {} finally { $fileStream = $null } }
									if ($contentStream -ne $null) { try { $contentStream.Dispose() } catch {} finally { $contentStream = $null } }
									if (Test-Path $outputFile -PathType Leaf) {
										$existingFileSize = (Get-Item $outputFile).Length
										New-Log "Re-checked file size for resume retry: $(Format-FileSize $existingFileSize)" -Level VERBOSE
										$fileMode = [System.IO.FileMode]::Append
									}
									else {
										New-Log "File '$outputFile' not found before retry, resetting resume state." -Level WARNING
										$existingFileSize = 0L
										$fileMode = [System.IO.FileMode]::Create
										$result.ResumeUsed = $false
									}
								}
								catch {
									New-Log "Error re-checking file size before retry, resetting resume state. Error: $($_.Exception.Message)" -Level ERROR
									$existingFileSize = 0L
									$fileMode = [System.IO.FileMode]::Create
									$result.ResumeUsed = $false
								}
							}
							else {
								$existingFileSize = 0L
								$fileMode = [System.IO.FileMode]::Create
							}
						}
						else {
							New-Log "Download FAILED for '$($result.FileName)' after $totalAttempts attempts. Last error: $errorMessage" -Level WARNING
							$result.Status = "Failed"
							$result.Retries = $RetryCount + 1 # Indicate all attempts used
							try {
								if (Test-Path $outputFile -PathType Leaf) {
									if ($fileStream -ne $null) { try { $fileStream.Dispose() } catch {} finally { $fileStream = $null } }
									$finalFileInfo = Get-Item $outputFile
									$result.TotalBytes = $finalFileInfo.Length
									$result.FileSize = Format-FileSize $finalFileInfo.Length
									New-Log "Failed download left file '$outputFile' with final size: $($result.FileSize)." -Level WARNING
								}
								else {
									$result.TotalBytes = 0L
									$result.FileSize = Format-FileSize 0L
								}
							}
							catch {
								New-Log "Could not determine final file size after failure for '$outputFile'." -Level WARNING
								$result.TotalBytes = -1L # Indicate unknown size
								$result.FileSize = "N/A"
							}
						}
					} # End Catch block for download attempt
					finally {
						if ($contentStream) { try { $contentStream.Dispose() } catch { New-Log "Error disposing content stream: $($_.Exception.Message)" -Level ERROR } }
						if ($fileStream) { try { $fileStream.Dispose() } catch { New-Log "Error disposing file stream: $($_.Exception.Message)" -Level ERROR } }
						if ($response) { try { $response.Dispose() } catch { New-Log "Error disposing response message: $($_.Exception.Message)" -Level ERROR } }
						if ($request) { try { $request.Dispose() } catch { New-Log "Error disposing request message: $($_.Exception.Message)" -Level ERROR } }
						New-Log "Cleaned up resources for attempt $currentTry." -Level VERBOSE
					} # End Finally block for download attempt
				} # End While loop for retries
				if (-not $downloadSuccess -and $result.Status -ne 'Failed') {
					$result.Status = "Failed"
					$result.Error = if ($result.Error) { "Download did not complete after all retries for an unknown reason." }
					New-Log "Download loop finished for '$($result.FileName)' but success flag not set and status not Failed. Marking as Failed." -Level WARNING
				}
			} # End If PerformDownload
		}
		catch {
			# Catch errors during the pre-download checks or unexpected errors wrapping the download loop
			$result.Status = "Failed"
			$result.Error = "Phase 4 (Download Execution) Failed unexpectedly."
			New-Log $result.Error -Level ERROR
			try { Write-Progress -Activity "Downloading $($result.FileName)" -Completed -Id $progressId } catch {}
		}
		finally {
			if ($disposeHttpClientLocally -and $sharedHttpClient -ne $null) {
				New-Log "Disposing internally created HttpClient for '$($result.FileName)'." -Level VERBOSE
				try { $sharedHttpClient.Dispose() } catch { New-Log "Error disposing internal HttpClient" -Level ERROR }
			}
		}
		Write-Output $result
	}
	end {
		New-Log "Finished processing all pipeline input for Download-File." -Level VERBOSE
		if ($originalSecurityProtocol -ne $null) {
			try {
				New-Log "Restoring original SecurityProtocol settings: $originalSecurityProtocol" -Level VERBOSE
				[System.Net.ServicePointManager]::SecurityProtocol = $originalSecurityProtocol
			}
			catch {
				New-Log "Failed to restore original SecurityProtocol settings." -Level ERROR
			}
		}
		if ($providedHttpClient -ne $null -and $DisposeClient) {
			New-Log "Disposing the provided external HttpClient instance as requested by -DisposeClient." -Level WARNING
			try {
				$providedHttpClient.Dispose()
			}
			catch {	}
		}
		elseif ($providedHttpClient -ne $null) {
			New-Log "Leaving the provided external HttpClient instance undisposed (as -DisposeClient was not specified)." -Level VERBOSE
		}
	}
}
#Can use this function to generate a random header:
#Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/Get-RandomHeader.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
###############################################################################################################################
#Needs custom logging function from: https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1#
###############################################################################################################################
#Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
#$header = Get-RandomHeader
#$result = Download-File -Url "https://github.com/git-for-windows/git/releases/download/v2.39.1.windows.1/Git-2.39.1-32-bit.exe" -FilePath "$env:TEMP\DownloadTests" -Headers $header