function New-Log {
	<#
    .SYNOPSIS
        Writes formatted log messages to the console and/or a file.
    .DESCRIPTION
        New-Log provides a flexible way to log messages with different levels (INFO, ERROR, EXTENDEDERROR, WARNING, etc.),
        formats (TEXT, JSON), and destinations (Console, File). It supports pipeline input,
        automatic grouping of pipeline objects, log file rotation, custom date formats,
        and intelligent error context retrieval. Handles null/empty messages and complex nested objects.
        Internal verbose messages only appear if -Verbose is used directly on New-Log.
        Returns the exact same object that was passed in when -ReturnObject is used.
        EXTENDEDERROR level provides comprehensive error details including inner exceptions, category info, and error IDs.
    .PARAMETER Message
        The message or object to log. Can be piped. Objects will be formatted appropriately.
        Handles $null and empty strings gracefully.
	.PARAMETER DontClearErrorVariable
		Stops the function from clearing the $Error variable. Default: clears it.
    .PARAMETER Level
        The severity level of the log message. Defaults to "INFO".
        Valid values: ERROR, EXTENDEDERROR, WARNING, INFO, SUCCESS, DEBUG, VERBOSE.
        EXTENDEDERROR shows full error context including extended details.
    .PARAMETER NoConsole
        Switch parameter. If present, suppresses output to the console.
    .PARAMETER ReturnObject
        Switch parameter. If present, returns the exact same object that was passed in.
    .PARAMETER LogFilePath
        Specifies the full path to the log file. If not provided, logging only occurs to the console (unless -NoConsole is specified).
    .PARAMETER ForcedLogFile
        Switch parameter. If present, the log file will be overwritten with the new entry instead of appending.
        Use with caution. (Consider renaming to -OverwriteLogFile in future versions)
    .PARAMETER AppendTimestampToFile
        Switch parameter. If present, appends a timestamp (yyyyMMdd_HHmmss) to the log file name before the extension.
        Example: mylog.log -> mylog_20231027_153000.log
    .PARAMETER LogRotationSizeMB
        Specifies the maximum size in megabytes (MB) for the log file before it's rotated.
        If the log file exceeds this size, it's renamed with a timestamp, and a new log file is started.
        Defaults to 0 (no rotation).
    .PARAMETER LogFormat
        Specifies the format for log entries written to the file. Defaults to "TEXT".
        Valid values: TEXT, JSON.
    .PARAMETER NoAutoGroup
        Switch parameter. If present, disables the automatic grouping of multiple pipeline input objects.
        Each pipeline item will be logged individually as it arrives.
    .PARAMETER NoErrorLookup
        Switch parameter. If present, prevents the function from automatically looking for error details ($_, $Error[0])
        when the Level is set to ERROR.
    .PARAMETER DateFormat
        Specifies the date and time format string for timestamps in the log.
        Defaults to 'yyyy-MM-dd HH:mm:ss.fff'.
        Common formats: 'yyyy-MM-dd HH:mm:ss.fff', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm:ss', 'MM/dd/yyyy HH:mm:ss', 'dd.MM.yyyy HH:mm:ss', 'yyyyMMdd_HHmmss'.
    .PARAMETER ErrorObject
        Allows explicitly passing an ErrorRecord object (e.g., from a catch block: $_ | New-Log -Level ERROR -ErrorObject $_).
        This takes precedence over automatic error lookup. Defaults to trying to capture $_ from the caller's scope if available.
    .EXAMPLE
		$ht = @{ Name = "John"; Age = 30; Role = "Developer" }
		$result = $ht | New-Log -Level INFO -ReturnObject
		# Console shows clean table format with prefixes, $result contains the original hashtable
    .EXAMPLE
		$complexObj = [PSCustomObject]@{ Name = "Complex"; Nested = @{ Level = 1; Data = @(1,2) } }
		$result = $complexObj | New-Log -Level INFO -ReturnObject
		# Console shows clean table format with prefixes, $result contains the original PSCustomObject
    .EXAMPLE
		Get-Process | select -first 3 | New-Log -GroupObjects
		# Output includes header and properly indented table rows
    .EXAMPLE
		New-Log "JSON test" -LogFilePath "test.log" -LogFormat JSON
		# Creates test.log in current directory
	.EXAMPLE
		$megaObject = [PSCustomObject]@{
			Level1  = @{
				Level2A = @{
					Level3A = @{
						Data    = @("Item1", "Item2", "Item3")
						Numbers = @(1..10)
						Nested  = @{
							DeepValue = "VeryDeep"
							DeepArray = @(
								@{ SubItem1 = "Value1"; SubData = @(1, 2, 3) }, @{ SubItem2 = "Value2"; SubData = @(4, 5, 6) }
							)
						}
					}
					Level3B = "Simple string at level 3"
				}
				Level2B = @{
					Dates    = @((Get-Date), (Get-Date).AddDays(-1), (Get-Date).AddDays(1))
					Booleans = @($true, $false, $true)
					Mixed    = @("String", 42, $true, @{ InnerHash = "Value" })
				}
			}
			Level1B = [PSCustomObject]@{
				Property1 = "PSCustomObject in main object"
				Property2 = @{
					Array = @(1..5)
					Hash  = @{ Key1 = "Value1"; Key2 = "Value2" }
				}
			}
		}
		$result = $megaObject | New-Log -Level INFO -ReturnObject
		[2025-07-06 11:02:19.322][INFO] Level1             Level1B
		[2025-07-06 11:02:19.322][INFO] ------             -------
		[2025-07-06 11:02:19.322][INFO] {Level2A, Level2B} @{Property1=PSCustomObject in main object; Property2=System.Collections.Hashtable}
	.EXAMPLE
		try {
			Get-ChildItem -Path C:\Nonexistingpath -ea Stop
		}
		catch {
			New-Log "Failed" -Level ERROR
		}
		[2025-07-06 07:10:30.174][ERROR] Failed [Function: test-error][CodeRow: (2,4) (Function,Script)][FailedCode: Get-ChildItem -Path C:\Nonexistingpath -ea Stop][ExceptionMessage: Cannot find path 'C:\Nonexistingpath' because it does not exist.]
    .NOTES
        Author: Harze2k
        Date:   2025-07-06
        Version: 4.0 (Completely redesigned complex object formatting to use Format-Table for clean output)
		- MAJOR: -ReturnObject now returns the ORIGINAL input object instead of log metadata
        - MAJOR: Completely redesigned complex object formatting to use Format-Table for clean output
		- MAJOR: Added -Level EXTENDEDERROR for even more error details.
		- Removed all internal Write-Verbose messeges. (Not needed).
        - Hashtables and custom objects now display as proper tables in console output
        - Table output (multiple lines) now counts as single logical entry
        - Improved user experience: pipeline operations with -ReturnObject now work intuitively
        - Console output for complex objects is now much more readable and professional
        - Fixed GroupedItems being empty in grouped pipeline results by properly converting List[object] to array
        - Added validation for DateFormat parameter with 6 most common datetime formats
        - Added automatic conversion of relative log file paths to absolute paths using current location
        - Fixed multi-line object output to show each line with proper timestamp and level prefix
        - Added validation for error line numbers and column offsets to prevent negative values
        - Internal Write-Verbose messages only show if -Verbose is passed *directly* to New-Log.
        - Handles null/empty string input gracefully.
        - Uses UTF8 encoding without BOM for log files.
        - Fixed log rotation to properly track rotated files
        - Removed unused IsPSCore parameter from Write-PrefixedLinesToConsole helper.
        - Improved object formatting to match natural PowerShell output with timestamp/level prefixes
    #>
	[CmdletBinding()]
	param(
		[Parameter(ValueFromPipeline, Position = 0)]$Message,
		[Parameter(Position = 1)][ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "EXTENDEDERROR", "VERBOSE", "DEBUG")][string]$Level = "INFO",
		[Parameter(Position = 2)][switch]$DontClearErrorVariable,
		[Parameter(Position = 3)][switch]$NoConsole,
		[Parameter(Position = 4)][switch]$ReturnObject,
		[Parameter(Position = 5)][string]$LogFilePath,
		[Parameter(Position = 6)][switch]$ForcedLogFile,
		[Parameter(Position = 7)][switch]$AppendTimestampToFile,
		[Parameter(Position = 8)][ValidateRange(0, [double]::MaxValue)][double]$LogRotationSizeMB = 1,
		[Parameter(Position = 9)][ValidateSet("TEXT", "JSON")][string]$LogFormat = "TEXT",
		[Parameter(Position = 11)][switch]$NoAutoGroup,
		[Parameter(Position = 12)][switch]$NoErrorLookup,
		[Parameter(Position = 13)][ValidateSet('yyyy-MM-dd HH:mm:ss.fff', 'yyyy-MM-dd HH:mm:ss.fff', 'yyyy-MM-ddTHH:mm:ss.fff', 'MM/dd/yyyy HH:mm:ss.fff', 'dd.MM.yyyy HH:mm:ss.fff', 'yyyyMMdd_HHmmss.fff')][string]$DateFormat = 'yyyy-MM-dd HH:mm:ss.fff',
		[Parameter()]$ErrorObject = $(if ($global:error.Count -gt 0) { $global:error[0] } else { $null })
	)
	begin {
		$script:isPSCore = $PSVersionTable.PSVersion.Major -ge 6
		$script:LevelColors = @{
			ERROR         = @{ ANSI = 91; PS = 'Red' }
			EXTENDEDERROR = @{ ANSI = 91; PS = 'Red' }
			WARNING       = @{ ANSI = 93; PS = 'Yellow' }
			INFO          = @{ ANSI = 37; PS = 'White' }
			SUCCESS       = @{ ANSI = 92; PS = 'Green' }
			DEBUG         = @{ ANSI = 94; PS = 'Blue' }
			VERBOSE       = @{ ANSI = 96; PS = 'Cyan' }
		}
		$script:Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($false)
		$script:OriginalConsoleEncoding = $null
		$script:ObjectCollection = [System.Collections.Generic.List[object]]::new()
		$script:IsPipelineInput = $MyInvocation.ExpectingInput
		$script:PipelineItemCounter = 0
		$script:InitialErrorCount = $Error.Count
		$script:RotatedFiles = @()    # Track rotated files for testing
		$script:LogRotated = $false   # Flag to indicate if rotation occurred
		try {
			$script:OriginalConsoleEncoding = [Console]::OutputEncoding
			[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
		}
		catch {
			Write-Warning "Failed to set console output encoding to UTF8: $($_.Exception.Message)"
		}
		if ($LogFilePath -and $AppendTimestampToFile) {
			try {
				# Convert relative path to absolute path if needed before processing timestamp
				$resolvedLogFilePath = $LogFilePath
				if (-not [System.IO.Path]::IsPathRooted($LogFilePath)) {
					$resolvedLogFilePath = Join-Path -Path (Get-Location -PSProvider FileSystem).Path -ChildPath $LogFilePath
				}
				$fileInfo = [System.IO.FileInfo]::new($resolvedLogFilePath)
				$logDir = $fileInfo.DirectoryName
				$logBaseName = $fileInfo.BaseName
				$logExtension = $fileInfo.Extension
				$timestampSuffix = Get-Date -Format 'yyyyMMdd_HHmmss'
				$LogFilePath = Join-Path -Path $logDir -ChildPath ($logBaseName + "_" + $timestampSuffix + $logExtension)
			}
			catch {
				Write-Error "Failed to append timestamp to LogFilePath '$LogFilePath' : $($_.Exception.Message)"
			}
		}
		#region Write-PrefixedLinesToConsole
		function Write-PrefixedLinesToConsole {
			[CmdletBinding()]
			param(
				[string]$Prefix,
				[string]$Content,
				$Color,
				[string]$CurrentTimestamp,
				[string]$CurrentLevel,
				[hashtable]$CurrentLevelColors,
				[bool]$CurrentIsPSCore,
				[bool]$UseIndividualPrefixes = $false
			)
			if ([string]::IsNullOrEmpty($Content)) {
				Write-Host $Prefix -ForegroundColor $Color
				return
			}
			$lines = $Content -split "\r?\n"
			$nonEmptyLines = $lines | Where-Object { $_ -ne '' }
			if ($nonEmptyLines.Count -eq 0) {
				Write-Host $Prefix -ForegroundColor $Color
				return
			}
			# For table output (multiple lines), show each line with prefix
			if ($UseIndividualPrefixes -and $nonEmptyLines.Count -gt 1) {
				foreach ($line in $nonEmptyLines) {
					$colorCode = $CurrentLevelColors[$CurrentLevel].ANSI
					$ansiReset = "`e[0m"
					$individualPrefix = if ($CurrentIsPSCore) {
						"`e[34m[$CurrentTimestamp]$ansiReset`e[${colorCode}m[$CurrentLevel]$ansiReset"
					}
					else {
						"[$CurrentTimestamp][$CurrentLevel]"
					}
					$outputLine = $individualPrefix + " " + $line
					Write-Host $outputLine -ForegroundColor $Color
				}
			}
			else {
				# For single lines or when individual prefixes not requested
				foreach ($line in $nonEmptyLines) {
					$outputLine = $Prefix + " " + $line
					Write-Host $outputLine -ForegroundColor $Color
				}
			}
		}
		#endregion Write-PrefixedLinesToConsole
		#region Get-ErrorToProcess
		function Get-ErrorToProcess {
			[CmdletBinding()]
			param(
				$CurrentItem,
				[bool]$NoErrorLookupParameter
			)
			if ($NoErrorLookupParameter) {
				return $null
			}
			# Check explicit ErrorObject parameter first
			if ($null -ne $ErrorObject -and $ErrorObject -is [System.Management.Automation.ErrorRecord]) {
				return $ErrorObject
			}
			# Check if CurrentItem is an ErrorRecord
			if ($CurrentItem -is [System.Management.Automation.ErrorRecord]) {
				return $CurrentItem
			}
			# Check for $_ in caller scope (most common in catch blocks)
			try {
				$scopedErrorVar = Get-Variable -Name '_' -Scope 1 -ErrorAction Stop
				if ($null -ne $scopedErrorVar.Value -and $scopedErrorVar.Value -is [System.Management.Automation.ErrorRecord]) {
					return $scopedErrorVar.Value
				}
			}
			catch {
				Write-Warning "Error objaect: $_ not available in caller scope"
			}
			if ($Error.Count -gt $script:InitialErrorCount) {
				return $Error[0]
			}
			if ($global:Error.Count -gt 0) {
				return $global:Error[0]
			}
			return $null
		}
		#endregion Get-ErrorToProcess
		#region Get-ExtendedErrorDetails
		function Get-ExtendedErrorDetails {
			[CmdletBinding()]
			param(
				[System.Management.Automation.ErrorRecord]$ErrorRecord
			)
			if ($null -eq $ErrorRecord) {
				return $null
			}
			$extendedDetails = [System.Collections.Generic.List[string]]::new()
			if ($ErrorRecord.Exception.InnerException) {
				$innerMsg = $ErrorRecord.Exception.InnerException.Message
				if ($innerMsg -and $innerMsg -ne $ErrorRecord.Exception.Message) {
					$extendedDetails.Add("InnerException: $innerMsg")
				}
			}
			if ($ErrorRecord.CategoryInfo) {
				$categoryInfo = $ErrorRecord.CategoryInfo.ToString()
				if ($categoryInfo -and $categoryInfo -ne 'NotSpecified') {
					$extendedDetails.Add("CategoryInfo: $categoryInfo")
				}
			}
			if ($ErrorRecord.FullyQualifiedErrorId -and $ErrorRecord.FullyQualifiedErrorId -ne 'NativeCommandError') {
				$extendedDetails.Add("FullyQualifiedErrorId: $($ErrorRecord.FullyQualifiedErrorId)")
			}
			if ($ErrorRecord.TargetObject) {
				$targetObjStr = try {
					$ErrorRecord.TargetObject.ToString()
				}
				catch {
					$ErrorRecord.TargetObject.GetType().Name
				}
				if ($targetObjStr -and $targetObjStr.Length -lt 100) {
					$extendedDetails.Add("TargetObject: $targetObjStr")
				}
			}
			if ($ErrorRecord.Exception.HResult -and $ErrorRecord.Exception.HResult -ne 0 -and $ErrorRecord.Exception.HResult -ne -2146233088) {
				$hresultHex = "0x{0:X8}" -f $ErrorRecord.Exception.HResult
				$extendedDetails.Add("HResult: $hresultHex")
			}
			if ($ErrorRecord.Exception.Source -and $ErrorRecord.Exception.Source -ne 'System.Management.Automation') {
				$extendedDetails.Add("Source: $($ErrorRecord.Exception.Source)")
			}
			if ($extendedDetails.Count -gt 0) {
				return ($extendedDetails -join " | ")
			}
			return $null
		}
		#endregion Get-ExtendedErrorDetails
		#region Format-ItemForDisplay
		function Format-ItemForDisplay {
			[CmdletBinding()]
			param(
				$Item
			)
			if ($null -eq $Item) {
				return ""
			}
			if ($Item -is [string]) {
				return $Item
			}
			# For hashtables, custom objects, etc. - choose format based on property count
			if ($Item -is [System.Collections.IDictionary] -or
				($Item -isnot [ValueType] -and $Item -isnot [string] -and $Item -isnot [System.Management.Automation.ErrorRecord])) {
				try {
					$propertyCount = 0
					if ($Item -is [System.Collections.IDictionary]) {
						$propertyCount = $Item.Keys.Count
					}
					elseif ($Item -is [PSCustomObject]) {
						$propertyCount = ($Item.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' }).Count
					}
					else {
						# For other objects, try to get visible properties
						$propertyCount = ($Item | Get-Member -MemberType Properties | Where-Object { $_.Name -notlike '__*' }).Count
					}
					# Use Format-List for 4+ properties, Format-Table for fewer
					if ($propertyCount -ge 4) {
						$listString = ($Item | Format-List | Out-String).Trim()
						return $listString
					}
					else {
						$tableString = ($Item | Format-Table -AutoSize | Out-String).Trim()
						return $tableString
					}
				}
				catch {
					return ($Item | Out-String -Width 4096).Trim()
				}
			}
			if ($Item -is [System.Management.Automation.ErrorRecord]) {
				return ($Item | Format-List * -Force | Out-String).Trim()
			}
			return ($Item | Out-String -Width 4096).Trim()
		}
		#endregion Format-ItemForDisplay
		#region Write-LogToFile
		function Write-LogToFile {
			[CmdletBinding()]
			param(
				[Parameter(Mandatory)][string]$ContentToWrite,
				[double]$LogRotationSizeMB = 1
			)
			if (-not $LogFilePath) {
				Write-Warning "LogFilePath not specified."
				return
			}
			try {
				# Convert relative path to absolute path if needed
				$resolvedLogFilePath = $LogFilePath
				if (-not [System.IO.Path]::IsPathRooted($LogFilePath)) {
					$resolvedLogFilePath = Join-Path -Path (Get-Location -PSProvider FileSystem).Path -ChildPath $LogFilePath
				}
				$parentDir = Split-Path -Path $resolvedLogFilePath -Parent -ErrorAction Stop
				if ([string]::IsNullOrEmpty($parentDir)) {
					$parentDir = (Get-Location -PSProvider FileSystem).Path
				}
				if (-not (Test-Path -Path $parentDir -PathType Container)) {
					New-Item -Path $parentDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
				}
				$LogFilePath = $resolvedLogFilePath
				$fileExists = Test-Path -Path $LogFilePath -PathType Leaf
				if ($LogRotationSizeMB -gt 0 -and $fileExists) {
					$logFileItem = Get-Item -Path $LogFilePath -ErrorAction SilentlyContinue
					# Convert bytes to MB for comparison - using raw bytes is more accurate
					$fileSizeInBytes = $logFileItem.Length
					$thresholdInBytes = $LogRotationSizeMB * 1MB
					if ($logFileItem -and $fileSizeInBytes -ge $thresholdInBytes) {
						$backupTimestamp = Get-Date -Format 'yyyyMMdd_HHmmssfff'
						$fileInfo = [System.IO.FileInfo]::new($LogFilePath)
						$backupPath = Join-Path -Path $fileInfo.DirectoryName -ChildPath ($fileInfo.BaseName + "_" + $backupTimestamp + $fileInfo.Extension)
						Copy-Item -Path $LogFilePath -Destination $backupPath -Force -ErrorAction SilentlyContinue
						Remove-Item -Path $LogFilePath -Force -ErrorAction SilentlyContinue
						$script:RotatedFiles += $backupPath
						$script:LogRotated = $true
						$fileExists = $false
					}
				}
				if ($ForcedLogFile -or -not $fileExists) {
					[System.IO.File]::WriteAllText($LogFilePath, $ContentToWrite, $script:Utf8NoBomEncoding)
				}
				else {
					$contentToAppend = if ($fileExists -and (Get-Item $LogFilePath).Length -gt 0) { "`r`n" + $ContentToWrite } else { $ContentToWrite }
					[System.IO.File]::AppendAllText($LogFilePath, $contentToAppend, $script:Utf8NoBomEncoding)
				}
			}
			catch {
				Write-Error "Failed write log '$LogFilePath': $($_.Exception.Message)"
			}
		}
		#endregion Write-LogToFile
		#region Process-SingleLogItem
		function Process-SingleLogItem {
			[CmdletBinding()]
			param(
				$ItemToProcess,
				[string]$CurrentLevel,
				[string]$CurrentDateFormat,
				[hashtable]$CurrentLevelColors,
				[bool]$CurrentIsPSCore,
				[bool]$CurrentNoConsole,
				[bool]$CurrentReturnObject,
				[bool]$CurrentNoErrorLookup,
				[bool]$DontClearErrorVariable
			)
			if ($null -eq $ItemToProcess -and $CurrentLevel -ne 'ERROR') {
			}
			# Prepare basic log entry
			$timestamp = Get-Date -Format $CurrentDateFormat
			$originalItem = $ItemToProcess
			$errorDetails = $null
			$effectiveItem = $ItemToProcess
			if (($null -eq $effectiveItem -or $effectiveItem -eq "") -and $CurrentLevel -ne 'ERROR' -and $CurrentLevel -ne 'EXTENDEDERROR') {
				$effectiveItem = ""
			}
			# Special handling for ERROR and EXTENDEDERROR levels
			$err = $null
			if ($CurrentLevel -eq 'ERROR' -or $CurrentLevel -eq 'EXTENDEDERROR') {
				$err = Get-ErrorToProcess -CurrentItem $effectiveItem -NoErrorLookupParameter $CurrentNoErrorLookup
				if ($err) {
					$errorDetails = [PSCustomObject]@{
						ExceptionType   = $err.Exception.GetType().FullName
						Exception       = $err.Exception.Message
						FullError       = $err
						FailedCode      = $null
						CallerInfo      = $null
						ExtendedDetails = $null
					}
					$extendedErrorInfo = Get-ExtendedErrorDetails -ErrorRecord $err
					if ($extendedErrorInfo) {
						$errorDetails.ExtendedDetails = $extendedErrorInfo
					}
					if (-not $DontClearErrorVariable) {
						$Error.Clear()
					}
					# Set failed code if available
					if ($err.InvocationInfo -and $err.InvocationInfo.Line) {
						$errorDetails.FailedCode = $err.InvocationInfo.Line.Trim()
					}
					else {
						$errorDetails.FailedCode = "N/A"
					}
					# Try to get caller info with improved validation and better fallback logic
					try {
						$cs = Get-PSCallStack -EA SilentlyContinue
						if ($cs -and $cs.Count -gt 1) {
							$cf = $null
							$fallbackFrame = $null
							# Primary pass: Look for named functions.
							for ($i = 1; $i -lt $cs.Count; $i++) {
								$frame = $cs[$i]
								$command = $frame.Command
								$functionName = $frame.FunctionName
								# Skip New-Log calls
								if ($command -eq 'New-Log') {
									continue
								}
								# Store first non-New-Log frame as fallback
								if (-not $fallbackFrame) {
									$fallbackFrame = $frame
								}
								# Prefer frames with actual function names
								if ($functionName -and $functionName -ne '<ScriptBlock>' -and $functionName -ne 'New-Log') {
									$cf = $frame
									break
								}
							}
							# Use fallback if no named function found
							if (-not $cf -and $fallbackFrame) {
								$cf = $fallbackFrame
							}
							if ($cf) {
								$sn = if ($cf.ScriptName) {
									[IO.Path]::GetFileName($cf.ScriptName)
								}
								else {
									'<NoScript>'
								}
								$fn = if ($cf.FunctionName -and $cf.FunctionName -ne '<ScriptBlock>') {
									$cf.FunctionName
								}
								else {
									"<Script>"
								}
								# Determine context type
								$contextType = if ($cf.FunctionName -and $cf.FunctionName -ne '<ScriptBlock>') {
									"Function"
								}
								else {
									"Script"
								}
								$location = if ($cf.ScriptName) {
									"Script"
								}
								else {
									"Interactive"
								}
								$errorDetails.CallerInfo = "File [$sn], Line [$($cf.ScriptLineNumber)], Context [Function [$fn]]"
							}
						}
					}
					catch {
						Write-Warning "Could not get caller info: $($_.Exception.Message)"
					}
					if (($null -eq $originalItem -or $originalItem -eq "") -or $originalItem -is [System.Management.Automation.ErrorRecord]) {
						$effectiveItem = $errorDetails.Exception
					}
				}
				elseif (-not $CurrentNoErrorLookup) {
					# No error context found AND NoErrorLookup was NOT specified
					if ($null -eq $effectiveItem -or $effectiveItem -eq "") {
						$effectiveItem = "<ERROR Logged with Null/Empty Message>"
						Write-Warning "ERROR level specified but no specific error context found for null/empty input."
					}
				}
			}
			$formattedMessage = Format-ItemForDisplay -Item $effectiveItem
			# Console output
			if (-not $CurrentNoConsole) {
				$colorCode = $CurrentLevelColors[$CurrentLevel].ANSI
				$psColor = $CurrentLevelColors[$CurrentLevel].PS
				$ansiReset = "`e[0m"
				$consolePrefix = if ($CurrentIsPSCore) {
					"`e[34m[$timestamp]$ansiReset`e[${colorCode}m[$CurrentLevel]$ansiReset"
				}
				else {
					"[$timestamp][$CurrentLevel]"
				}
				if (($CurrentLevel -eq 'ERROR' -or $CurrentLevel -eq 'EXTENDEDERROR') -and $errorDetails) {
					$functionNameFromCaller = "N/A"
					$coderowContext = "Unknown"
					if ($errorDetails.CallerInfo) {
						# Extract function name from CallerInfo string: "File [...], Line [...], Context [Function [functionname]]"
						$contextMatch = [regex]::Match($errorDetails.CallerInfo, "Context \[Function \[(.*?)\]\]$")
						if ($contextMatch.Success) {
							$functionNameFromCaller = $contextMatch.Groups[1].Value
							$coderowContext = if ($functionNameFromCaller -eq "<Script>") { "Script" } else { "Function" }
						}
						else {
							# Try alternate pattern for script context
							$scriptMatch = [regex]::Match($errorDetails.CallerInfo, "Context \[Function \[<Script>\]\]$")
							if ($scriptMatch.Success) {
								$functionNameFromCaller = "<Script>"
								$coderowContext = "Script"
							}
						}
					}
					$failedCodeLineNum = "N/A"
					$failedCodeColNum = "N/A"
					if ($err -and $err.InvocationInfo) {
						# Validate and sanitize line number
						if ($null -ne $err.InvocationInfo.ScriptLineNumber -and $err.InvocationInfo.ScriptLineNumber -gt 0) {
							$failedCodeLineNum = [Math]::Max(1, $err.InvocationInfo.ScriptLineNumber)
						}
						# Validate and sanitize column number
						if ($null -ne $err.InvocationInfo.OffsetInLine -and $err.InvocationInfo.OffsetInLine -ge 0) {
							$failedCodeColNum = [Math]::Max(0, $err.InvocationInfo.OffsetInLine)
						}
					}
					$coderowLocation = "Unknown"
					if ($err -and $err.InvocationInfo -and $err.InvocationInfo.ScriptName) {
						$coderowLocation = "Script"
					}
					elseif ($err -and $err.InvocationInfo) {
						$coderowLocation = "Interactive"
					}
					$failedCodeText = if ($errorDetails.FailedCode -ne 'N/A') { $errorDetails.FailedCode } else { "N/A" }
					$exceptionMsgText = $errorDetails.Exception
					if ($CurrentIsPSCore) {
						$ansiReset = "`e[0m"
						$blue = "`e[94m"
						$errorSuffix = " [${blue}Function${ansiReset}: $functionNameFromCaller]"
						$errorSuffix += "[${blue}CodeRow${ansiReset}: ($failedCodeLineNum,$failedCodeColNum) ($coderowContext,$coderowLocation)]"
						$errorSuffix += "[${blue}FailedCode${ansiReset}: $failedCodeText]"
						$errorSuffix += "[${blue}ExceptionMessage${ansiReset}: ${ansireset}`e[$($CurrentLevelColors[$CurrentLevel].ANSI)m$exceptionMsgText$ansireset]"
						if ($CurrentLevel -eq 'EXTENDEDERROR' -and $errorDetails.ExtendedDetails) {
							$errorSuffix += "[${blue}ExtendedErrorDetail${ansiReset}: `e[$($CurrentLevelColors[$CurrentLevel].ANSI)m$($errorDetails.ExtendedDetails)$ansireset]"
						}
					}
					else {
						$errorSuffix = " [Function: $functionNameFromCaller]"
						$errorSuffix += "[CodeRow: ($failedCodeLineNum,$failedCodeColNum) ($coderowContext,$coderowLocation)]"
						$errorSuffix += "[FailedCode: $failedCodeText]"
						$errorSuffix += "[ExceptionMessage: $exceptionMsgText]"
						if ($CurrentLevel -eq 'EXTENDEDERROR' -and $errorDetails.ExtendedDetails) {
							$errorSuffix += "[ExtendedErrorDetail: $($errorDetails.ExtendedDetails)]"
						}
					}
					Write-PrefixedLinesToConsole -Prefix $consolePrefix -Content ($formattedMessage + $errorSuffix) -Color $psColor -CurrentTimestamp $timestamp -CurrentLevel $CurrentLevel -CurrentLevelColors $CurrentLevelColors -CurrentIsPSCore $CurrentIsPSCore -UseIndividualPrefixes $false
				}
				else {
					# Check if the formatted message is multi-line (like table output) and needs individual prefixes
					$needsIndividualPrefixes = $formattedMessage -and $formattedMessage.Contains("`n") -and ($formattedMessage -split "`n").Count -gt 1
					Write-PrefixedLinesToConsole -Prefix $consolePrefix -Content $formattedMessage -Color $psColor -CurrentTimestamp $timestamp -CurrentLevel $CurrentLevel -CurrentLevelColors $CurrentLevelColors -CurrentIsPSCore $CurrentIsPSCore -UseIndividualPrefixes $needsIndividualPrefixes
				}
			}
			if ($LogFilePath) {
				$fileContent = $null
				if ($LogFormat -eq "JSON") {
					$logEntryBase = [PSCustomObject]@{
						Timestamp = $timestamp
						Level     = $CurrentLevel
						Message   = if ($originalItem -is [Management.Automation.ErrorRecord]) { $errorDetails.Exception } else { $originalItem }
					}
					if ($errorDetails) {
						$logEntryBase | Add-Member -MemberType NoteProperty -Name ErrorDetails -Value $errorDetails -Force
					}
					$logEntryForJson = $logEntryBase | Select-Object * -ExcludeProperty FullError
					if ($logEntryBase.ErrorDetails) {
						$logEntryForJson.ErrorDetails | Add-Member -NotePropertyName StackTrace -Value $logEntryBase.ErrorDetails.FullError.ScriptStackTrace -Force
					}
					try {
						$fileContent = $logEntryForJson | ConvertTo-Json -Depth 5 -Compress -ErrorAction Stop
					}
					catch {
						Write-Warning "JSON conversion failed: $($_.Exception.Message). Fallback."
						$fileContent = "[$timestamp][$CurrentLevel] JSON_Error: $($_.Exception.Message) - Item: ($($originalItem | Out-String -Width 200).Trim())"
						if ($errorDetails) {
							$fileContent += " [Err: $($errorDetails.Exception)]"
						}
					}
				}
				else {
					$fileContent = "[$timestamp][$CurrentLevel]"
					if (-not [string]::IsNullOrEmpty($formattedMessage)) {
						$fileContent += " $formattedMessage"
					}
					if (($CurrentLevel -eq 'ERROR' -or $CurrentLevel -eq 'EXTENDEDERROR') -and $errorDetails) {
						# Add error details in same format as console with enhanced caller detection
						$functionNameFromCaller = "N/A"
						$coderowContext = "Unknown"
						if ($errorDetails.CallerInfo) {
							$contextMatch = [regex]::Match($errorDetails.CallerInfo, "Context \[Function \[(.*?)\]\]$")
							if ($contextMatch.Success) {
								$functionNameFromCaller = $contextMatch.Groups[1].Value
								$coderowContext = "Function"
							}
							else {
								$scriptMatch = [regex]::Match($errorDetails.CallerInfo, "Context \[Function \[<Script>\]\]$")
								if ($scriptMatch.Success) {
									$functionNameFromCaller = "<Script>"
									$coderowContext = "Script"
								}
							}
						}
						$failedCodeLineNum = "N/A"
						$failedCodeColNum = "N/A"
						if ($err.InvocationInfo) {
							if ($null -ne $err.InvocationInfo.ScriptLineNumber -and $err.InvocationInfo.ScriptLineNumber -gt 0) {
								$failedCodeLineNum = [Math]::Max(1, $err.InvocationInfo.ScriptLineNumber)
							}
							if ($null -ne $err.InvocationInfo.OffsetInLine -and $err.InvocationInfo.OffsetInLine -ge 0) {
								$failedCodeColNum = [Math]::Max(0, $err.InvocationInfo.OffsetInLine)
							}
						}
						$coderowLocation = "Unknown"
						if ($err.InvocationInfo -and $err.InvocationInfo.ScriptName) {
							$coderowLocation = "Script"
						}
						elseif ($err.InvocationInfo) {
							$coderowLocation = "Interactive"
						}
						$failedCodeText = if ($errorDetails.FailedCode -ne 'N/A') { $errorDetails.FailedCode } else { "N/A" }
						$exceptionMsgText = $errorDetails.Exception
						$fileContent += " [Function: $functionNameFromCaller]"
						$fileContent += "[CodeRow: ($failedCodeLineNum,$failedCodeColNum) ($coderowContext,$coderowLocation)]"
						$fileContent += "[FailedCode: $failedCodeText]"
						$fileContent += "[ExceptionMessage: $exceptionMsgText]"
						if ($CurrentLevel -eq 'EXTENDEDERROR' -and $errorDetails.ExtendedDetails) {
							$fileContent += "[ExtendedErrorDetail: $($errorDetails.ExtendedDetails)]"
						}
					}
				}
				Write-LogToFile -ContentToWrite $fileContent -LogRotationSizeMB $LogRotationSizeMB
			}
			# Return the EXACT original object if requested
			if ($CurrentReturnObject) {
				return $originalItem
			}
			else {
				return $null
			}
		}
		#endregion Process-SingleLogItem
		#region Process-GroupedLogItems
		function Process-GroupedLogItems {
			[CmdletBinding()]
			param(
				[Parameter(Mandatory)]$ItemsToProcess,
				[string]$CurrentLevel,
				[string]$CurrentDateFormat,
				[hashtable]$CurrentLevelColors,
				[bool]$CurrentIsPSCore,
				[bool]$CurrentNoConsole,
				[bool]$CurrentReturnObject
			)
			if ($null -eq $ItemsToProcess -or $ItemsToProcess.Count -eq 0) {
				return $null
			}
			$timestamp = Get-Date -Format $CurrentDateFormat -ErrorAction Stop
			$objectCount = $ItemsToProcess.Count
			if ($LogFilePath) {
				if ($LogFormat -eq "JSON") {
					foreach ($item in $ItemsToProcess) {
						$entryBase = [PSCustomObject]@{
							Timestamp = $timestamp
							Level     = $CurrentLevel
							Message   = $item
						}
						$fileContent = $entryBase | ConvertTo-Json -Depth 5 -Compress -ErrorAction SilentlyContinue
						if ($fileContent) {
							Write-LogToFile -ContentToWrite $fileContent
						}
						else {
							Write-Warning "Failed JSON group item: Type '$($item.GetType().Name)'."
							$fallbackContent = @{
								Timestamp = $timestamp
								Level     = $CurrentLevel
								Message   = "JSON_Error grouped type $($item.GetType().Name)"
							} | ConvertTo-Json -Compress
							Write-LogToFile -ContentToWrite $fallbackContent
						}
					}
				}
				else {
					# Choose format based on item count: 4+ items = vertical (Format-List), fewer = horizontal (Format-Table)
					if ($objectCount -ge 4) {
						$formattedString = ($ItemsToProcess | Format-List | Out-String).Trim()
					}
					else {
						$formattedString = ($ItemsToProcess | Format-Table -AutoSize | Out-String).Trim()
					}
					$fileContent = "[$timestamp][$CurrentLevel] Start of grouped items ($objectCount):`r`n"
					$formattedLines = $formattedString -split "`r?`n"
					foreach ($line in $formattedLines) {
						$fileContent += "[$timestamp][$CurrentLevel] $line`r`n"
					}
					$fileContent += "[$timestamp][$CurrentLevel] End of grouped items."
					Write-LogToFile -ContentToWrite $fileContent.TrimEnd()
				}
			}
			# Console output for grouped items
			if (-not $CurrentNoConsole) {
				$colorCode = $CurrentLevelColors[$CurrentLevel].ANSI
				$psColor = $CurrentLevelColors[$CurrentLevel].PS
				$ansiReset = "`e[0m"
				$consolePrefix = if ($CurrentIsPSCore) {
					"`e[34m[$timestamp]$ansiReset`e[${colorCode}m[$CurrentLevel]$ansiReset"
				}
				else {
					"[$timestamp][$CurrentLevel]"
				}
				Write-PrefixedLinesToConsole -Prefix $consolePrefix -Content "Displaying $objectCount grouped items:" -Color $psColor -CurrentTimestamp $timestamp -CurrentLevel $CurrentLevel -CurrentLevelColors $CurrentLevelColors -CurrentIsPSCore $CurrentIsPSCore -UseIndividualPrefixes $false
				if ($objectCount -ge 4) {
					$formattedString = ($ItemsToProcess | Format-List | Out-String).Trim()
				}
				else {
					$formattedString = ($ItemsToProcess | Format-Table -AutoSize | Out-String).Trim()
				}
				# For grouped output, we want each line to have its own prefix
				Write-PrefixedLinesToConsole -Prefix $consolePrefix -Content $formattedString -Color $psColor -CurrentTimestamp $timestamp -CurrentLevel $CurrentLevel -CurrentLevelColors $CurrentLevelColors -CurrentIsPSCore $CurrentIsPSCore -UseIndividualPrefixes $true
			}
			if ($CurrentReturnObject) {
				# Return the original objects as an array
				if ($ItemsToProcess -is [System.Collections.Generic.List[object]]) {
					return $ItemsToProcess.ToArray()
				}
				else {
					return @($ItemsToProcess)
				}
			}
			else {
				return $null
			}
		}
		#endregion Process-GroupedLogItems
	}
	process {
		$currentItem = $Message
		# Skip VERBOSE level messages if not enabled
		if ($Level -eq "VERBOSE") {
			$prefs = try { Get-Variable -Name VerbosePreference -Scope 1 -ValueOnly -ErrorAction Stop } catch { $null }
			if ($prefs -ne [Management.Automation.ActionPreference]::Continue) {
				return
			}
		}
		# Handle pipeline input
		if ($script:IsPipelineInput) {
			$script:PipelineItemCounter++
			if (-not $NoAutoGroup) {
				$script:ObjectCollection.Add($currentItem)
				return
			}
		}
		# Process immediately if not pipeline OR pipeline with -NoAutoGroup
		try {
			$result = Process-SingleLogItem -ItemToProcess $currentItem -CurrentLevel $Level -CurrentDateFormat $DateFormat -CurrentLevelColors $script:LevelColors -CurrentIsPSCore $script:isPSCore -CurrentNoConsole $NoConsole -CurrentReturnObject $ReturnObject -CurrentNoErrorLookup $NoErrorLookup -DontClearErrorVariable $DontClearErrorVariable.IsPresent -ErrorAction Stop
			if ($ReturnObject -and $null -ne $result) {
				Write-Output $result  # Explicitly use Write-Output to ensure clean pipeline behavior
			}
		}
		catch {
			Write-Error "Error processing single item Process: $($_.Exception.Message)"
		}
	} # End Process block
	end {
		try {
			if ($Level -eq "VERBOSE") {
				$prefs = try { Get-Variable -Name VerbosePreference -Scope 1 -ValueOnly -ErrorAction Stop } catch { $null }
				if ($prefs -ne [Management.Automation.ActionPreference]::Continue) {
					return
				}
			}
			$processedGroup = $false
			$result = $null
			# Handle auto-grouped pipeline input
			if ($script:IsPipelineInput -and -not $NoAutoGroup -and $script:PipelineItemCounter -gt 1) {
				try {
					$result = Process-GroupedLogItems -ItemsToProcess $script:ObjectCollection -CurrentLevel $Level -CurrentDateFormat $DateFormat -CurrentLevelColors $script:LevelColors -CurrentIsPSCore $script:isPSCore -CurrentNoConsole $NoConsole -CurrentReturnObject $ReturnObject -ErrorAction Stop
					$processedGroup = $true
				}
				catch {
					Write-Error "Error processing GroupedLogItems: $($_.Exception.Message)"
					$processedGroup = $false
				}
			}
			# Handle single pipeline item with auto-grouping enabled
			elseif ($script:IsPipelineInput -and -not $NoAutoGroup -and $script:PipelineItemCounter -eq 1) {
				try {
					$result = Process-SingleLogItem -ItemToProcess $script:ObjectCollection[0] -CurrentLevel $Level -CurrentDateFormat $DateFormat -CurrentLevelColors $script:LevelColors -CurrentIsPSCore $script:isPSCore -CurrentNoConsole $NoConsole -CurrentReturnObject $ReturnObject -CurrentNoErrorLookup $NoErrorLookup -DontClearErrorVariable $DontClearErrorVariable.IsPresent -ErrorAction Stop
					$processedGroup = $true
				}
				catch {
					Write-Error "Error processing single item Process: $($_.Exception.Message)"
					$processedGroup = $false
				}
			}
			if ($processedGroup -and $ReturnObject -and $null -ne $result) {
				Write-Output $result
			}
		}
		catch {
			Write-Error "Error during End block: $($_.Exception.Message)"
		}
		finally {
			if ($null -ne $script:OriginalConsoleEncoding -and [Console]::OutputEncoding -ne $script:OriginalConsoleEncoding) {
				try {
					[Console]::OutputEncoding = $script:OriginalConsoleEncoding
				}
				catch {
					Write-Warning "Failed restore console encode: $($_.Exception.Message)"
				}
			}
			if ($script:ObjectCollection) {
				$script:ObjectCollection.Clear()
			}
		}
	} # End End block
}