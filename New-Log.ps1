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
        ERROR level shows: function name, exception type, code row (line/col), failed code, exception message.
        EXTENDEDERROR additionally shows: full InnerException chain, extended category info, HResult, source, and stack trace.
    .PARAMETER Message
        The message or object to log. Can be piped. Objects will be formatted appropriately.
        Handles $null and empty strings gracefully.
    .PARAMETER Level
        The severity level of the log message. Defaults to "INFO".
        Valid values: ERROR, EXTENDEDERROR, WARNING, INFO, SUCCESS, DEBUG, VERBOSE.
        ERROR shows function/line/code/exception context.
        EXTENDEDERROR additionally shows InnerException chain, HResult, CategoryInfo, and PowerShell stack trace.
    .PARAMETER PreserveErrorVariable
        Switch parameter. If present, prevents the function from clearing $Error after logging an error.
        Default behavior: clears $Error after capturing context to prevent stale errors from
        contaminating future ERROR-level log calls.
        Alias: DontClearErrorVariable (backward compatibility)
    .PARAMETER NoConsole
        Switch parameter. If present, suppresses output to the console.
    .PARAMETER ReturnObject
        Switch parameter. If present, returns the exact same object that was passed in.
    .PARAMETER LogFilePath
        Specifies the full path to the log file. If not provided, logging only occurs to the console
        (unless -NoConsole is also specified).
    .PARAMETER OverwriteLogFile
        Switch parameter. If present, the log file will be overwritten with the new entry instead of appending.
        Alias: ForcedLogFile (backward compatibility)
    .PARAMETER AppendTimestampToFile
        Switch parameter. If present, appends a timestamp (yyyyMMdd_HHmmss) to the log file name before the extension.
        Example: mylog.log -> mylog_20231027_153000.log
    .PARAMETER LogRotationSizeMB
        Specifies the maximum size in megabytes (MB) for the log file before it is rotated.
        If the log file exceeds this size, it is renamed with a timestamp and a new log file is started.
        Defaults to 1. Set to 0 to disable rotation.
    .PARAMETER LogFormat
        Specifies the format for log entries written to the file. Defaults to "TEXT".
        Valid values: TEXT, JSON.
        JSON format includes structured CallerInfo and full stack trace for ERROR/EXTENDEDERROR levels.
    .PARAMETER NoAutoGroup
        Switch parameter. If present, disables the automatic grouping of multiple pipeline input objects.
        Each pipeline item will be logged individually as it arrives.
    .PARAMETER NoErrorLookup
        Switch parameter. If present, prevents the function from automatically looking for error details
        ($_, $Error[0]) when the Level is set to ERROR or EXTENDEDERROR.
    .PARAMETER DateFormat
        Specifies the date and time format string for timestamps in the log.
        Defaults to 'yyyy-MM-dd HH:mm:ss.fff'.
    .PARAMETER ErrorObject
        Allows explicitly passing an ErrorRecord object (e.g., from a catch block: New-Log "Msg" -Level ERROR -ErrorObject $_).
        This takes precedence over automatic error lookup.
        Default: captures $global:Error[0] at parameter-bind time to ensure errors are available before any function code executes.
    .EXAMPLE
        New-Log "Application started" -Level INFO
    .EXAMPLE
        try {
            Get-ChildItem -Path C:\Nonexistingpath -ea Stop
        }
        catch {
            New-Log "Failed to list directory" -Level ERROR
        }
        # Output: [timestamp][ERROR] Failed to list directory [Function: <caller>][ExceptionType: ...][CodeRow: (line,col)...][FailedCode: ...][ExceptionMessage: ...]
    .EXAMPLE
        try {
            [System.IO.File]::ReadAllText("C:\missing.txt")
        }
        catch {
            New-Log "File read error" -Level EXTENDEDERROR
        }
        # Output includes InnerException chain, HResult, CategoryInfo, and StackTrace
    .EXAMPLE
        $ht = @{ Name = "John"; Age = 30; Role = "Developer" }
        $result = $ht | New-Log -Level INFO -ReturnObject
        # Console shows clean table format; $result contains the original hashtable
    .EXAMPLE
        Get-Process | Select-Object -First 3 | New-Log -Level INFO
        # Output includes header and properly indented table rows
    .EXAMPLE
        New-Log "JSON test" -LogFilePath "C:\logs\app.log" -LogFormat JSON
        # Creates or appends to app.log in JSON format with structured error details
    .EXAMPLE
        New-Log "Session start" -LogFilePath "C:\logs\app.log" -AppendTimestampToFile
        # Creates app_20231027_153000.log
    .NOTES
        Author: Harze2k
        Version: 5.2
        Date: 2026-02-24

        VERSION HISTORY:
        ================
        v5.2 - Bug fixes, DRY refactor, EXTENDEDERROR improvements
        ----------------------------------------------------------------
        BUG FIXES:
        - Fixed file TEXT output always writing coderowContext="Function" even for script-root callers
        (console section had the correct conditional; file section always wrote "Function")
        - Fixed missing "$err -and" null guard before $err.InvocationInfo in file TEXT section
        (would throw NullReferenceException if error lookup returned $null)
        - Removed spurious Write-Warning in Get-ErrorToProcess that fired whenever $_ was not
        defined at scope 1 - this is completely normal and produced false positive warnings
        - Fixed typo "objaect" -> "object" in error message
        - Filled parameter Position gap (was: 9,11,12; now: 9,10,11 with no gap)
        - Fixed Write-LogToFile mutating outer-scope $LogFilePath via local $resolvedLogFilePath variable
        (inner function was writing to enclosing scope variable, now self-contained)
        - Fixed $script:RotatedFiles still using array concatenation (+= @()) despite v5.0 notes
        claiming it was changed to List[string]; now correctly uses List[string].Add()
        - Fixed Get-ErrorToProcess to actually check scopes 1-5 for $_ (only checked scope 1 before,
        v5.1 notes claimed multi-scope was implemented but it was not)
        - Removed redundant second $global:Error check in Get-ErrorToProcess ($Error and $global:Error
        are the same reference in PowerShell; the second check returned stale pre-call errors)
        - Removed empty no-op if block in Process-SingleLogItem (was: if ($null -eq ...) { })
        - Fixed Process-GroupedLogItems JSON loop not passing LogRotationSizeMB to Write-LogToFile
        - Fixed duplicate 'yyyy-MM-dd HH:mm:ss.fff' in DateFormat ValidateSet
        REFACTORING:
        - CallerInfo stored as hashtable @{FunctionName; ScriptName; LineNumber; ContextType; LocationType}
        instead of a formatted string that was immediately re-parsed with regex (fragile)
        - Extracted Get-ErrorSuffixText helper: console and file TEXT now share one suffix-building
        function (DRY). Accepts -UseANSI switch for console vs file output.
        - Write-LogToFile now self-contained: accepts ResolvedFilePath, OverwriteFile, RotationSizeMB
        as parameters instead of reading/mutating outer-scope variables
        - Log file path resolved once in begin block ($script:ResolvedLogFilePath) for all output
        EXTENDEDERROR IMPROVEMENTS:
        - ExceptionType now shown in TEXT output (was captured in PSCustomObject but never displayed)
        - Full InnerException chain walked (not just the first level), with depth labels:
        InnerException, InnerException[1], InnerException[2], etc.
        - PowerShell stack trace (ScriptStackTrace) included in EXTENDEDERROR TEXT output as
        condensed single line: [StackTrace: at Foo ... | at Bar ...]
        - Stack trace stored directly in errorDetails object (no more dynamic Add-Member for JSON)
        - Distinct warning message when EXTENDEDERROR requested but no error record is available
        - JSON output now uses clean separate ErrorDetails object (no FullError property leaking)
        PARAMETER RENAMES (backward-compatible aliases retained for all existing callers):
        - DontClearErrorVariable -> PreserveErrorVariable (removes confusing double negative)
        - ForcedLogFile -> OverwriteLogFile (clearer intent)
        DATEFORMAT IMPROVEMENTS:
        - Added 'yyyy-MM-dd HH:mm:ss' (no milliseconds) to the ValidateSet
	#>
	[CmdletBinding()]
	param(
		[Parameter(ValueFromPipeline, Position = 0)]$Message,
		[Parameter(Position = 1)][ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "EXTENDEDERROR", "VERBOSE", "DEBUG")][string]$Level = "INFO",
		[Parameter(Position = 2)][Alias('DontClearErrorVariable')][switch]$PreserveErrorVariable,
		[Parameter(Position = 3)][switch]$NoConsole,
		[Parameter(Position = 4)][switch]$ReturnObject,
		[Parameter(Position = 5)][string]$LogFilePath,
		[Parameter(Position = 6)][Alias('ForcedLogFile')][switch]$OverwriteLogFile,
		[Parameter(Position = 7)][switch]$AppendTimestampToFile,
		[Parameter(Position = 8)][ValidateRange(0, [double]::MaxValue)][double]$LogRotationSizeMB = 1,
		[Parameter(Position = 9)][ValidateSet("TEXT", "JSON")][string]$LogFormat = "TEXT",
		[Parameter(Position = 10)][switch]$NoAutoGroup,
		[Parameter(Position = 11)][switch]$NoErrorLookup,
		[Parameter(Position = 12)][ValidateSet(
			'yyyy-MM-dd HH:mm:ss.fff',
			'yyyy-MM-dd HH:mm:ss',
			'yyyy-MM-ddTHH:mm:ss.fff',
			'MM/dd/yyyy HH:mm:ss.fff',
			'dd.MM.yyyy HH:mm:ss.fff',
			'yyyyMMdd_HHmmss.fff'
		)][string]$DateFormat = 'yyyy-MM-dd HH:mm:ss.fff',
		[Parameter()]$ErrorObject = $(if ($global:error.Count -and $global:error.Count -gt 0) { $global:error[0] } else { $null })
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
		$script:RotatedFiles = [System.Collections.Generic.List[string]]::new()
		$script:LogRotated = $false
		# Resolve log file path once in begin block so all inner functions use the same path
		$script:ResolvedLogFilePath = $null
		if ($LogFilePath) {
			try {
				$rawPath = $LogFilePath
				if (-not [System.IO.Path]::IsPathRooted($rawPath)) {
					$rawPath = Join-Path -Path (Get-Location -PSProvider FileSystem).Path -ChildPath $rawPath
				}
				if ($AppendTimestampToFile) {
					$fi = [System.IO.FileInfo]::new($rawPath)
					$tsSuffix = Get-Date -Format 'yyyyMMdd_HHmmss'
					$rawPath = Join-Path -Path $fi.DirectoryName -ChildPath ($fi.BaseName + '_' + $tsSuffix + $fi.Extension)
				}
				$script:ResolvedLogFilePath = $rawPath
			}
			catch {
				Write-Error "Failed to resolve LogFilePath '$LogFilePath': $($_.Exception.Message)"
			}
		}
		try {
			if ($Host.Name -ne 'Windows PowerShell ISE Host' -and $Host.Name -ne 'ServerRemoteHost') {
				$script:OriginalConsoleEncoding = [Console]::OutputEncoding
				[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
			}
		}
		catch {
			# Silently ignore - no console handle available (ISE, scheduled tasks, etc.)
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
			$nonEmptyLines = @($lines | Where-Object { $_ -ne '' })
			if ($nonEmptyLines.Count -eq 0) {
				Write-Host $Prefix -ForegroundColor $Color
				return
			}
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
					Write-Host ($individualPrefix + ' ' + $line) -ForegroundColor $Color
				}
			}
			else {
				foreach ($line in $nonEmptyLines) {
					Write-Host ($Prefix + ' ' + $line) -ForegroundColor $Color
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
			# Explicit ErrorObject parameter takes highest precedence
			if ($null -ne $ErrorObject -and $ErrorObject -is [System.Management.Automation.ErrorRecord]) {
				return $ErrorObject
			}
			# Current pipeline item may itself be an ErrorRecord
			if ($CurrentItem -is [System.Management.Automation.ErrorRecord]) {
				return $CurrentItem
			}
			# Search scopes 1-5 for $_ that is an ErrorRecord.
			# Scope 1 is the direct caller; higher scopes handle nested helper function scenarios
			# where the catch block $_ may be several frames up the call stack.
			for ($scope = 1; $scope -le 5; $scope++) {
				try {
					$scopedVar = Get-Variable -Name '_' -Scope $scope -ErrorAction Stop
					if ($null -ne $scopedVar.Value -and $scopedVar.Value -is [System.Management.Automation.ErrorRecord]) {
						return $scopedVar.Value
					}
				}
				catch {
					# Variable not accessible at this scope - continue to next
				}
			}
			# Fall back to $Error only if a NEW error appeared since this function call started
			if ($Error.Count -gt $script:InitialErrorCount) {
				return $Error[0]
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
			# Walk the full InnerException chain (not just the first level)
			$inner = $ErrorRecord.Exception.InnerException
			$innerDepth = 0
			while ($null -ne $inner) {
				$innerMsg = $inner.Message
				if ($innerMsg -and $innerMsg -ne $ErrorRecord.Exception.Message) {
					$label = if ($innerDepth -eq 0) { 'InnerException' } else { "InnerException[$innerDepth]" }
					$extendedDetails.Add("${label}: $innerMsg")
				}
				$inner = $inner.InnerException
				$innerDepth++
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
				$targetObjStr = try { $ErrorRecord.TargetObject.ToString() } catch { $ErrorRecord.TargetObject.GetType().Name }
				if ($targetObjStr -and $targetObjStr.Length -lt 100) {
					$extendedDetails.Add("TargetObject: $targetObjStr")
				}
			}
			if ($ErrorRecord.Exception.HResult -and $ErrorRecord.Exception.HResult -ne 0 -and $ErrorRecord.Exception.HResult -ne -2146233088) {
				$extendedDetails.Add("HResult: 0x{0:X8}" -f $ErrorRecord.Exception.HResult)
			}
			if ($ErrorRecord.Exception.Source -and $ErrorRecord.Exception.Source -ne 'System.Management.Automation') {
				$extendedDetails.Add("Source: $($ErrorRecord.Exception.Source)")
			}
			if ($extendedDetails.Count -gt 0) { return ($extendedDetails -join ' | ') } else { return $null }
		}
		#endregion Get-ExtendedErrorDetails
		#region Get-ErrorSuffixText
		function Get-ErrorSuffixText {
			# Builds the structured error context suffix string.
			# Used by both console output (UseANSI=$true) and file TEXT output (UseANSI=$false).
			# This is the single source of truth for the error suffix format.
			[CmdletBinding()]
			param(
				[PSCustomObject]$ErrorDetails,
				[System.Management.Automation.ErrorRecord]$ErrorRecord,
				[string]$CurrentLevel,
				[bool]$UseANSI = $false,
				[int]$ANSIColorCode = 91
			)
			if ($null -eq $ErrorDetails) { return '' }
			# Extract from structured CallerInfo hashtable (no regex needed)
			$callerInfo = $ErrorDetails.CallerInfo
			$fnName = if ($callerInfo -and $callerInfo.FunctionName) { $callerInfo.FunctionName } else { 'N/A' }
			$contextType = if ($callerInfo -and $callerInfo.ContextType) { $callerInfo.ContextType } else { 'Unknown' }
			$locType = if ($callerInfo -and $callerInfo.LocationType) { $callerInfo.LocationType } else { 'Unknown' }
			$lineNum = 'N/A'
			$colNum = 'N/A'
			if ($ErrorRecord -and $ErrorRecord.InvocationInfo) {
				if ($null -ne $ErrorRecord.InvocationInfo.ScriptLineNumber -and $ErrorRecord.InvocationInfo.ScriptLineNumber -gt 0) {
					$lineNum = [Math]::Max(1, $ErrorRecord.InvocationInfo.ScriptLineNumber)
				}
				if ($null -ne $ErrorRecord.InvocationInfo.OffsetInLine -and $ErrorRecord.InvocationInfo.OffsetInLine -ge 0) {
					$colNum = [Math]::Max(0, $ErrorRecord.InvocationInfo.OffsetInLine)
				}
			}
			$failedCode = if ($ErrorDetails.FailedCode -and $ErrorDetails.FailedCode -ne 'N/A') { $ErrorDetails.FailedCode } else { 'N/A' }
			$exMsg = $ErrorDetails.Exception
			$exType = $ErrorDetails.ExceptionType
			if ($UseANSI) {
				$reset = "`e[0m"
				$blue = "`e[94m"
				$lvlColor = "`e[${ANSIColorCode}m"
				$suffix = " [${blue}Function${reset}: $fnName]"
				$suffix += "[${blue}ExceptionType${reset}: $exType]"
				$suffix += "[${blue}CodeRow${reset}: ($lineNum,$colNum) ($contextType,$locType)]"
				$suffix += "[${blue}FailedCode${reset}: $failedCode]"
				$suffix += "[${blue}ExceptionMessage${reset}: ${lvlColor}$exMsg${reset}]"
				if ($CurrentLevel -eq 'EXTENDEDERROR') {
					if ($ErrorDetails.ExtendedDetails) {
						$suffix += "[${blue}ExtendedErrorDetail${reset}: ${lvlColor}$($ErrorDetails.ExtendedDetails)${reset}]"
					}
					if ($ErrorDetails.StackTrace) {
						$condensed = ($ErrorDetails.StackTrace -replace '\r?\n', ' | ').Trim()
						$suffix += "[${blue}StackTrace${reset}: $condensed]"
					}
				}
			}
			else {
				$suffix = " [Function: $fnName]"
				$suffix += "[ExceptionType: $exType]"
				$suffix += "[CodeRow: ($lineNum,$colNum) ($contextType,$locType)]"
				$suffix += "[FailedCode: $failedCode]"
				$suffix += "[ExceptionMessage: $exMsg]"
				if ($CurrentLevel -eq 'EXTENDEDERROR') {
					if ($ErrorDetails.ExtendedDetails) {
						$suffix += "[ExtendedErrorDetail: $($ErrorDetails.ExtendedDetails)]"
					}
					if ($ErrorDetails.StackTrace) {
						$condensed = ($ErrorDetails.StackTrace -replace '\r?\n', ' | ').Trim()
						$suffix += "[StackTrace: $condensed]"
					}
				}
			}
			return $suffix
		}
		#endregion Get-ErrorSuffixText
		#region Format-ItemForDisplay
		function Format-ItemForDisplay {
			[CmdletBinding()]
			param($Item)
			if ($null -eq $Item) { return '' }
			if ($Item -is [string]) { return $Item }
			if ($Item -is [System.Collections.IDictionary] -or
				($Item -isnot [ValueType] -and $Item -isnot [string] -and $Item -isnot [System.Management.Automation.ErrorRecord])) {
				try {
					$propertyCount = 0
					if ($Item -is [System.Collections.IDictionary]) {
						$propertyCount = $Item.Keys.Count
					}
					elseif ($Item -is [PSCustomObject]) {
						$propertyCount = @($Item.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' }).Count
					}
					else {
						$propertyCount = @($Item | Get-Member -MemberType Properties | Where-Object { $_.Name -notlike '__*' }).Count
					}
					if ($propertyCount -ge 4) {
						return ($Item | Format-List | Out-String).Trim()
					}
					else {
						return ($Item | Format-Table -AutoSize | Out-String).Trim()
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
			# Self-contained file writer. Accepts resolved path as a parameter.
			# Does not read or mutate any outer-scope variables.
			[CmdletBinding()]
			param(
				[Parameter(Mandatory)][string]$ContentToWrite,
				[Parameter(Mandatory)][string]$ResolvedFilePath,
				[bool]$OverwriteFile = $false,
				[double]$RotationSizeMB = 1
			)
			try {
				$parentDir = Split-Path -Path $ResolvedFilePath -Parent -ErrorAction Stop
				if ([string]::IsNullOrEmpty($parentDir)) {
					$parentDir = (Get-Location -PSProvider FileSystem).Path
				}
				if (-not (Test-Path -Path $parentDir -PathType Container)) {
					New-Item -Path $parentDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
				}
				$fileExists = Test-Path -Path $ResolvedFilePath -PathType Leaf
				if ($RotationSizeMB -gt 0 -and $fileExists) {
					$logFileItem = Get-Item -Path $ResolvedFilePath -ErrorAction SilentlyContinue
					if ($logFileItem -and $logFileItem.Length -ge ($RotationSizeMB * 1MB)) {
						$backupTimestamp = Get-Date -Format 'yyyyMMdd_HHmmssfff'
						$fi = [System.IO.FileInfo]::new($ResolvedFilePath)
						$backupPath = Join-Path -Path $fi.DirectoryName -ChildPath ($fi.BaseName + '_' + $backupTimestamp + $fi.Extension)
						Copy-Item -Path $ResolvedFilePath -Destination $backupPath -Force -ErrorAction SilentlyContinue
						Remove-Item -Path $ResolvedFilePath -Force -ErrorAction SilentlyContinue
						$script:RotatedFiles.Add($backupPath)
						$script:LogRotated = $true
						$fileExists = $false
					}
				}
				if ($OverwriteFile -or -not $fileExists) {
					[System.IO.File]::WriteAllText($ResolvedFilePath, $ContentToWrite, $script:Utf8NoBomEncoding)
				}
				else {
					$appendContent = if ((Get-Item -Path $ResolvedFilePath -ErrorAction SilentlyContinue).Length -gt 0) {
						"`r`n" + $ContentToWrite
					}
					else {
						$ContentToWrite
					}
					[System.IO.File]::AppendAllText($ResolvedFilePath, $appendContent, $script:Utf8NoBomEncoding)
				}
			}
			catch {
				Write-Error "Failed to write log '$ResolvedFilePath': $($_.Exception.Message)"
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
				[bool]$PreserveErrorVariable
			)
			$timestamp = Get-Date -Format $CurrentDateFormat
			$originalItem = $ItemToProcess
			$effectiveItem = $ItemToProcess
			$errorDetails = $null
			$err = $null
			if ($CurrentLevel -in @('ERROR', 'EXTENDEDERROR')) {
				$err = Get-ErrorToProcess -CurrentItem $effectiveItem -NoErrorLookupParameter $CurrentNoErrorLookup
				if ($err) {
					$errorDetails = [PSCustomObject]@{
						ExceptionType   = $err.Exception.GetType().FullName
						Exception       = $err.Exception.Message
						FullError       = $err
						FailedCode      = $null
						CallerInfo      = $null
						ExtendedDetails = $null
						StackTrace      = $err.ScriptStackTrace
					}
					$errorDetails.ExtendedDetails = Get-ExtendedErrorDetails -ErrorRecord $err
					if (-not $PreserveErrorVariable) {
						$Error.Clear()
					}
					$errorDetails.FailedCode = if ($err.InvocationInfo -and $err.InvocationInfo.Line) {
						$err.InvocationInfo.Line.Trim()
					}
					else { 'N/A' }
					# Build structured CallerInfo hashtable - direct property access later, no regex
					try {
						$cs = @(Get-PSCallStack -EA SilentlyContinue)
						if ($cs.Count -gt 1) {
							$cf = $null
							$fallbackFrame = $null
							for ($i = 1; $i -lt $cs.Count; $i++) {
								$frame = $cs[$i]
								if ($frame.Command -eq 'New-Log') { continue }
								if (-not $fallbackFrame) { $fallbackFrame = $frame }
								if ($frame.FunctionName -and $frame.FunctionName -ne '<ScriptBlock>' -and $frame.FunctionName -ne 'New-Log') {
									$cf = $frame
									break
								}
							}
							if (-not $cf -and $fallbackFrame) { $cf = $fallbackFrame }
							if ($cf) {
								$isNamedFunction = $cf.FunctionName -and $cf.FunctionName -ne '<ScriptBlock>'
								$errorDetails.CallerInfo = @{
									FunctionName = if ($isNamedFunction) { $cf.FunctionName } else { '<Script>' }
									ScriptName   = if ($cf.ScriptName) { [IO.Path]::GetFileName($cf.ScriptName) } else { '<NoScript>' }
									LineNumber   = $cf.ScriptLineNumber
									ContextType  = if ($isNamedFunction) { 'Function' } else { 'Script' }
									LocationType = if ($cf.ScriptName) { 'Script' } else { 'Interactive' }
								}
							}
						}
					}
					catch {
						Write-Warning "Could not get caller info: $($_.Exception.Message)"
					}
					# Use error exception message as the display message when no explicit message was given
					if ($null -eq $originalItem -or $originalItem -eq '' -or $originalItem -is [System.Management.Automation.ErrorRecord]) {
						$effectiveItem = $errorDetails.Exception
					}
				}
				elseif (-not $CurrentNoErrorLookup) {
					if ($null -eq $effectiveItem -or $effectiveItem -eq '') {
						$levelLabel = if ($CurrentLevel -eq 'EXTENDEDERROR') { 'EXTENDEDERROR' } else { 'ERROR' }
						$effectiveItem = "<$levelLabel Logged with Null/Empty Message>"
						Write-Warning "$levelLabel level specified but no error context was found for null/empty input."
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
				if ($CurrentLevel -in @('ERROR', 'EXTENDEDERROR') -and $errorDetails) {
					$errorSuffix = Get-ErrorSuffixText -ErrorDetails $errorDetails -ErrorRecord $err -CurrentLevel $CurrentLevel -UseANSI $CurrentIsPSCore -ANSIColorCode $colorCode
					Write-PrefixedLinesToConsole -Prefix $consolePrefix -Content ($formattedMessage + $errorSuffix) -Color $psColor -CurrentTimestamp $timestamp -CurrentLevel $CurrentLevel -CurrentLevelColors $CurrentLevelColors -CurrentIsPSCore $CurrentIsPSCore -UseIndividualPrefixes $false
				}
				else {
					$needsIndividualPrefixes = $formattedMessage -and $formattedMessage.Contains("`n") -and ($formattedMessage -split "`n").Count -gt 1
					Write-PrefixedLinesToConsole -Prefix $consolePrefix -Content $formattedMessage -Color $psColor -CurrentTimestamp $timestamp -CurrentLevel $CurrentLevel -CurrentLevelColors $CurrentLevelColors -CurrentIsPSCore $CurrentIsPSCore -UseIndividualPrefixes $needsIndividualPrefixes
				}
			}
			# File output
			if ($script:ResolvedLogFilePath) {
				$fileContent = $null
				if ($LogFormat -eq 'JSON') {
					# Build a clean ErrorDetails object for JSON - no FullError reference leaking
					$jsonErrorDetails = $null
					if ($errorDetails) {
						$jsonErrorDetails = [PSCustomObject]@{
							ExceptionType   = $errorDetails.ExceptionType
							Exception       = $errorDetails.Exception
							FailedCode      = $errorDetails.FailedCode
							CallerInfo      = $errorDetails.CallerInfo
							ExtendedDetails = $errorDetails.ExtendedDetails
							StackTrace      = $errorDetails.StackTrace
						}
					}
					$logEntry = [PSCustomObject]@{
						Timestamp    = $timestamp
						Level        = $CurrentLevel
						Message      = if ($originalItem -is [System.Management.Automation.ErrorRecord]) { $errorDetails.Exception } else { $originalItem }
						ErrorDetails = $jsonErrorDetails
					}
					try {
						$fileContent = $logEntry | ConvertTo-Json -Depth 5 -Compress -ErrorAction Stop
					}
					catch {
						Write-Warning "JSON conversion failed: $($_.Exception.Message). Using text fallback."
						$fileContent = "[$timestamp][$CurrentLevel] JSON_Error: $($_.Exception.Message) - Item: $(($originalItem | Out-String -Width 200).Trim())"
						if ($errorDetails) { $fileContent += " [Err: $($errorDetails.Exception)]" }
					}
				}
				else {
					$fileContent = "[$timestamp][$CurrentLevel]"
					if (-not [string]::IsNullOrEmpty($formattedMessage)) {
						$fileContent += " $formattedMessage"
					}
					if ($CurrentLevel -in @('ERROR', 'EXTENDEDERROR') -and $errorDetails) {
						$fileContent += Get-ErrorSuffixText -ErrorDetails $errorDetails -ErrorRecord $err -CurrentLevel $CurrentLevel -UseANSI $false
					}
				}
				Write-LogToFile -ContentToWrite $fileContent -ResolvedFilePath $script:ResolvedLogFilePath -OverwriteFile $OverwriteLogFile.IsPresent -RotationSizeMB $LogRotationSizeMB
			}
			if ($CurrentReturnObject) {
				if ($errorDetails) {
					return [PSCustomObject]@{
						Message         = $originalItem
						ExceptionType   = $errorDetails.ExceptionType
						Exception       = $errorDetails.Exception
						FailedCode      = $errorDetails.FailedCode
						CallerInfo      = $errorDetails.CallerInfo
						ExtendedDetails = $errorDetails.ExtendedDetails
						StackTrace      = $errorDetails.StackTrace
					}
				}
				return $originalItem
			}
			else { return $null }
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
			if ($script:ResolvedLogFilePath) {
				if ($LogFormat -eq 'JSON') {
					foreach ($item in $ItemsToProcess) {
						$entryBase = [PSCustomObject]@{
							Timestamp = $timestamp
							Level     = $CurrentLevel
							Message   = $item
						}
						$fileContent = $entryBase | ConvertTo-Json -Depth 5 -Compress -ErrorAction SilentlyContinue
						if ($fileContent) {
							Write-LogToFile -ContentToWrite $fileContent -ResolvedFilePath $script:ResolvedLogFilePath -OverwriteFile $OverwriteLogFile.IsPresent -RotationSizeMB $LogRotationSizeMB
						}
						else {
							Write-Warning "Failed JSON group item: Type '$($item.GetType().Name)'."
							$fallbackContent = @{
								Timestamp = $timestamp
								Level     = $CurrentLevel
								Message   = "JSON_Error grouped type $($item.GetType().Name)"
							} | ConvertTo-Json -Compress
							Write-LogToFile -ContentToWrite $fallbackContent -ResolvedFilePath $script:ResolvedLogFilePath -OverwriteFile $OverwriteLogFile.IsPresent -RotationSizeMB $LogRotationSizeMB
						}
					}
				}
				else {
					$formattedString = if ($objectCount -ge 4) {
						($ItemsToProcess | Format-List | Out-String).Trim()
					}
					else {
						($ItemsToProcess | Format-Table -AutoSize | Out-String).Trim()
					}
					$fileContent = "[$timestamp][$CurrentLevel] Start of grouped items ($objectCount):`r`n"
					foreach ($line in ($formattedString -split "`r?`n")) {
						$fileContent += "[$timestamp][$CurrentLevel] $line`r`n"
					}
					$fileContent += "[$timestamp][$CurrentLevel] End of grouped items."
					Write-LogToFile -ContentToWrite $fileContent.TrimEnd() -ResolvedFilePath $script:ResolvedLogFilePath -OverwriteFile $OverwriteLogFile.IsPresent -RotationSizeMB $LogRotationSizeMB
				}
			}
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
				$formattedString = if ($objectCount -ge 4) {
					($ItemsToProcess | Format-List | Out-String).Trim()
				}
				else {
					($ItemsToProcess | Format-Table -AutoSize | Out-String).Trim()
				}
				Write-PrefixedLinesToConsole -Prefix $consolePrefix -Content $formattedString -Color $psColor -CurrentTimestamp $timestamp -CurrentLevel $CurrentLevel -CurrentLevelColors $CurrentLevelColors -CurrentIsPSCore $CurrentIsPSCore -UseIndividualPrefixes $true
			}
			if ($CurrentReturnObject) {
				return $(if ($ItemsToProcess -is [System.Collections.Generic.List[object]]) { $ItemsToProcess.ToArray() } else { @($ItemsToProcess) })
			}
			else { return $null }
		}
		#endregion Process-GroupedLogItems
	}
	process {
		$currentItem = $Message
		if ($Level -eq 'VERBOSE') {
			$prefs = try { Get-Variable -Name VerbosePreference -Scope 1 -ValueOnly -ErrorAction Stop } catch { $null }
			if ($prefs -ne [Management.Automation.ActionPreference]::Continue) { return }
		}
		if ($script:IsPipelineInput) {
			$script:PipelineItemCounter++
			if (-not $NoAutoGroup) {
				$script:ObjectCollection.Add($currentItem)
				return
			}
		}
		try {
			$result = Process-SingleLogItem -ItemToProcess $currentItem -CurrentLevel $Level -CurrentDateFormat $DateFormat -CurrentLevelColors $script:LevelColors -CurrentIsPSCore $script:isPSCore -CurrentNoConsole $NoConsole -CurrentReturnObject $ReturnObject -CurrentNoErrorLookup $NoErrorLookup -PreserveErrorVariable $PreserveErrorVariable.IsPresent -ErrorAction Stop
			if ($ReturnObject -and $null -ne $result) { Write-Output $result }
		}
		catch {
			Write-Error "Error processing single item: $($_.Exception.Message)"
		}
	}
	end {
		try {
			if ($Level -eq 'VERBOSE') {
				$prefs = try { Get-Variable -Name VerbosePreference -Scope 1 -ValueOnly -ErrorAction Stop } catch { $null }
				if ($prefs -ne [Management.Automation.ActionPreference]::Continue) { return }
			}
			$processedGroup = $false
			$result = $null
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
			elseif ($script:IsPipelineInput -and -not $NoAutoGroup -and $script:PipelineItemCounter -eq 1) {
				try {
					$result = Process-SingleLogItem -ItemToProcess $script:ObjectCollection[0] -CurrentLevel $Level -CurrentDateFormat $DateFormat -CurrentLevelColors $script:LevelColors -CurrentIsPSCore $script:isPSCore -CurrentNoConsole $NoConsole -CurrentReturnObject $ReturnObject -CurrentNoErrorLookup $NoErrorLookup -PreserveErrorVariable $PreserveErrorVariable.IsPresent -ErrorAction Stop
					$processedGroup = $true
				}
				catch {
					Write-Error "Error processing single pipeline item in end block: $($_.Exception.Message)"
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
				try { [Console]::OutputEncoding = $script:OriginalConsoleEncoding }
				catch { Write-Warning "Failed to restore console encoding: $($_.Exception.Message)" }
			}
			if ($script:ObjectCollection) { $script:ObjectCollection.Clear() }
		}
	}
}