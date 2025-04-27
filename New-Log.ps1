#Requires -Version 5.1
function New-Log {
	<#
    .SYNOPSIS
        Writes formatted log messages to the console and/or a file.
    .DESCRIPTION
        New-Log provides a flexible way to log messages with different levels (INFO, ERROR, WARNING, etc.),
        formats (TEXT, JSON), and destinations (Console, File). It supports pipeline input,
        automatic grouping of pipeline objects, log file rotation, custom date formats,
        and intelligent error context retrieval. Handles null/empty messages and complex nested objects.
        Internal verbose messages only appear if -Verbose is used directly on New-Log.
        Adds type information and respects original content indentation for multi-line console output.
    .PARAMETER Message
        The message or object to log. Can be piped. Objects will be formatted appropriately.
        Handles $null and empty strings gracefully.
    .PARAMETER Level
        The severity level of the log message. Defaults to "INFO".
        Valid values: ERROR, WARNING, INFO, SUCCESS, DEBUG, VERBOSE.
    .PARAMETER NoConsole
        Switch parameter. If present, suppresses output to the console.
    .PARAMETER ReturnObject
        Switch parameter. If present, returns a PSCustomObject representing the log entry.
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
    .PARAMETER GroupObjects
        Switch parameter. If present, collects all pipeline input objects and logs them as a single formatted table at the end.
        Overrides automatic grouping behavior.
    .PARAMETER NoAutoGroup
        Switch parameter. If present, disables the automatic grouping of multiple pipeline input objects.
        Each pipeline item will be logged individually as it arrives.
    .PARAMETER NoErrorLookup
        Switch parameter. If present, prevents the function from automatically looking for error details ($_, $Error[0])
        when the Level is set to ERROR.
    .PARAMETER DateFormat
        Specifies the date and time format string for timestamps in the log.
        Defaults to 'yyyy-MM-dd HH:mm:ss.fff'.
    .PARAMETER ErrorObject
        Allows explicitly passing an ErrorRecord object (e.g., from a catch block: $_ | New-Log -Level ERROR -ErrorObject $_).
        This takes precedence over automatic error lookup. Defaults to trying to capture $_ from the caller's scope if available.
    .EXAMPLE
		$ht = @{ Name = "John"; Age = 30; Role = "Developer" }
		$ht | New-Log -Level INFO
		# Output includes type and properly indented key-value pairs
    .EXAMPLE
		$complexObj = [PSCustomObject]@{ Name = "Complex"; Nested = @{ Level = 1; Data = @(1,2) } }
		$complexObj | New-Log -Level INFO
		# Output includes type and properly indented JSON representation
    .EXAMPLE
		Get-Process | select -first 3 | New-Log -GroupObjects
		# Output includes header and properly indented table rows
	.EXAMPLE
		$complexObj = [PSCustomObject]@{
			Name   = "Complex Object"
			Nested = [PSCustomObject]@{
				Level     = 1
				Data      = @(1, 2, 3)
				SubNested = [PSCustomObject]@{
					Level      = 2
					Enabled    = $true
					Properties = @{
						A = "Value A"
						B = "Value B"
					}
				}
			}
		}
		$complexObj | New-Log
		Output:
		[2025-04-23 05:32:38.661][INFO] Type is [management.automation.pscustomobject]
		[2025-04-23 05:32:38.661][INFO] {
		[2025-04-23 05:32:38.661][INFO]   "Name": "Complex Object",
		[2025-04-23 05:32:38.661][INFO]   "Nested": {
		[2025-04-23 05:32:38.661][INFO]     "Level": 1,
		[2025-04-23 05:32:38.661][INFO]     "Data": [
		[2025-04-23 05:32:38.661][INFO]       1,
		[2025-04-23 05:32:38.661][INFO]       2,
		[2025-04-23 05:32:38.661][INFO]       3
		[2025-04-23 05:32:38.661][INFO]     ],
		[2025-04-23 05:32:38.661][INFO]     "SubNested": {
		[2025-04-23 05:32:38.661][INFO]       "Level": 2,
		[2025-04-23 05:32:38.661][INFO]       "Enabled": true,
		[2025-04-23 05:32:38.661][INFO]       "Properties": {
		[2025-04-23 05:32:38.661][INFO]         "A": "Value A",
		[2025-04-23 05:32:38.661][INFO]         "B": "Value B"
		[2025-04-23 05:32:38.661][INFO]       }
		[2025-04-23 05:32:38.661][INFO]     }
		[2025-04-23 05:32:38.661][INFO]   }
		[2025-04-23 05:32:38.661][INFO] }
	.EXAMPLE
		try {
			1 / 0
		}
		catch {
			New-Log "Failed" -Level ERROR
		}
		[2025-04-23 05:37:32.118][ERROR] Failed
		[2025-04-23 05:37:32.118][ERROR Detail] Exception: Attempted to divide by zero.
		[2025-04-23 05:37:32.118][ERROR Detail] Caller: File [New-Log.ps1], Line [848], Context [New-Log.ps1]
		[2025-04-23 05:37:32.118][ERROR Detail] Code: 1 / 0
    .NOTES
        Author: Harze2k
        Date:   2025-04-27 (Updated)
        Version: 3.4 (Just some cleanup.)
        - Console writer now respects and preserves the original indentation from formatters (like ConvertTo-Json, Format-Table) instead of applying a secondary, fixed indent.
        - Added 'Type is [typename]' line for complex objects in console TEXT output.
        - Internal Write-Verbose messages only show if -Verbose is passed *directly* to New-Log.
        - Handles null/empty string input gracefully.
        - Uses UTF8 encoding without BOM for log files.
        - Added enhanced output to make function more testable
        - Fixed log rotation to properly track rotated files
    #>
	[CmdletBinding()]
	param(
		[Parameter(ValueFromPipeline, Position = 0)]$Message,
		[Parameter(Position = 1)][ValidateSet("ERROR", "WARNING", "INFO", "SUCCESS", "DEBUG", "VERBOSE")][string]$Level = "INFO",
		[Parameter(Position = 2)][switch]$NoConsole,
		[Parameter(Position = 3)][switch]$ReturnObject,
		[Parameter(Position = 4)][string]$LogFilePath,
		[Parameter(Position = 5)][switch]$ForcedLogFile,
		[Parameter(Position = 6)][switch]$AppendTimestampToFile,
		[Parameter(Position = 7)][ValidateRange(0, [double]::MaxValue)][double]$LogRotationSizeMB = 1,
		[Parameter(Position = 8)][ValidateSet("TEXT", "JSON")][string]$LogFormat = "TEXT",
		[Parameter(Position = 9)][switch]$GroupObjects,
		[Parameter(Position = 10)][switch]$NoAutoGroup,
		[Parameter(Position = 11)][switch]$NoErrorLookup,
		[Parameter(Position = 12)][string]$DateFormat = 'yyyy-MM-dd HH:mm:ss.fff',
		[Parameter()]$ErrorObject = $(if ($global:error.Count -gt 0) { $global:error[0] } else {$null })
	)
	Begin {
		#region Initialize Variables
		$script:isPSCore = $PSVersionTable.PSVersion.Major -ge 6
		$script:LevelColors = @{
			ERROR   = @{ ANSI = 91; PS = 'Red' }
			WARNING = @{ ANSI = 93; PS = 'Yellow' }
			INFO    = @{ ANSI = 37; PS = 'White' }
			SUCCESS = @{ ANSI = 92; PS = 'Green' }
			DEBUG   = @{ ANSI = 94; PS = 'Blue' }
			VERBOSE = @{ ANSI = 96; PS = 'Cyan' }
		}
		$script:Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($false)
		$script:OriginalConsoleEncoding = $null
		$script:ShowInternalVerbose = $PSBoundParameters.ContainsKey('Verbose') -and ($VerbosePreference -eq [System.Management.Automation.ActionPreference]::Continue)
		$script:ObjectCollection = [System.Collections.Generic.List[object]]::new()
		$script:IsPipelineInput = $MyInvocation.ExpectingInput
		$script:PipelineItemCounter = 0
		$script:InitialErrorCount = $Error.Count
        $script:LogRecordCount = 0    # Added to track log entries for testing
        $script:RotatedFiles = @()    # Added to track rotated files for testing
        $script:LogRotated = $false   # Added flag to indicate if rotation occurred
		#endregion Initialize Variables
		if ($script:ShowInternalVerbose) { Write-Verbose "Internal verbose logging enabled for this call." }
		#region Setup Console Encoding
		try {
			$script:OriginalConsoleEncoding = [Console]::OutputEncoding
			[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
		}
		catch {
			Write-Warning "Failed to set console output encoding to UTF8: $($_.Exception.Message)"
		}
		#endregion Setup Console Encoding
		#region Handle Timestamped Log File
		if ($LogFilePath -and $AppendTimestampToFile) {
			try {
				$fileInfo = [System.IO.FileInfo]::new($LogFilePath)
				$logDir = $fileInfo.DirectoryName
				$logBaseName = $fileInfo.BaseName
				$logExtension = $fileInfo.Extension
				$timestampSuffix = Get-Date -Format 'yyyyMMdd_HHmmss'
				$LogFilePath = Join-Path -Path $logDir -ChildPath ($logBaseName + "_" + $timestampSuffix + $logExtension)
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Appended timestamp, new LogFilePath: $LogFilePath"
				}
			}
			catch {
				Write-Error "Failed to append timestamp to LogFilePath '$LogFilePath' : $($_.Exception.Message)"
			}
		}
		#endregion Handle Timestamped Log File
		#region Helper Functions
		#region Console Output Functions
		function Write-PrefixedLinesToConsole {
			[CmdletBinding()]
			param(
				[string]$Prefix,
				[string]$Content, # Allow empty
				$Color,
				[bool]$IsPSCore
			)
			$lines = $Content -split "\r?\n"
			$nonEmptyLines = $lines | Where-Object { $_ -ne '' }
			if ($nonEmptyLines.Count -eq 0) {
				Write-Host $Prefix -ForegroundColor $Color
				return
			}
			foreach ($line in $nonEmptyLines) {
				# Construct the output string before calling Write-Host
				$outputLine = $Prefix + " " + $line
				Write-Host $outputLine -ForegroundColor $Color
                # Track that console output was generated (for testing)
                $script:LogRecordCount++
			}
		}
		#endregion Console Output Functions
		#region Error Handling Functions
		function Get-ErrorToProcess {
			[CmdletBinding()]
			param(
				$CurrentItem,
				[bool]$NoErrorLookupParameter
			)
			if ($ErrorObject -ne $null -and $ErrorObject -is [System.Management.Automation.ErrorRecord]) {
				return $ErrorObject
			}
			if ($CurrentItem -is [System.Management.Automation.ErrorRecord]) {
				return $CurrentItem
			}
			if ($NoErrorLookupParameter) {
				return $null
			}
			$scopedErrorVar = try { Get-Variable -Name '_' -Scope 1 -ErrorAction Stop } catch { $null }
			if ($scopedErrorVar -ne $null -and $scopedErrorVar.Value -is [System.Management.Automation.ErrorRecord]) {
				return $scopedErrorVar.Value
			}
			if ($Error.Count -gt $script:InitialErrorCount) {
				return $Error[0]
			}
			try {
				$callStack = Get-PSCallStack -ErrorAction SilentlyContinue
				if ($callStack -match 'catch') {
					if ($Error.Count -gt 0) {
						return $Error[0]
					}
				}
			}
			catch {
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Could not examine call stack: $($_.Exception.Message)"
				}
			}
			return $null
		}
		#endregion Error Handling Functions
		#region Formatting Functions
		function Format-ItemForTextOutput {
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
			if ($Item -is [System.Collections.IDictionary]) {
				$typeName = $Item.GetType().Name.ToLower()
				$outputLines = [System.Collections.Generic.List[string]]::new()
				$sortedKeys = $Item.Keys | Sort-Object
				foreach ($key in $sortedKeys) {
					$valueString = try {
						($Item[$key] | Out-String -Width 120).Trim()
					}
					catch {
						"<Error Formatting Value>"
					}
					$outputLines.Add("$key : $valueString")
				}
				$keyValueString = $outputLines -join "`r`n"
				return "Type is [$typeName]`r`n$keyValueString"
			}
			if ($Item -is [System.Management.Automation.ErrorRecord]) {
				return ($Item | Format-List * -Force | Out-String).Trim()
			}
			if ($Item -isnot [ValueType]) {
				$typeName = ($Item.PSObject.TypeNames[0] -replace '^System\.', '').ToLower()
				try {
					$jsonString = ($Item | ConvertTo-Json -Depth 5 -WarningAction SilentlyContinue | Out-String).Trim()
					return "Type is [$typeName]`r`n$jsonString"
				}
				catch {
					Write-Warning "Failed JSON format: $($_.Exception.Message). Fallback."
					$fallbackString = ($Item | Out-String -Width 4096).Trim()
					return "Type is [$typeName]`r`n$fallbackString"
				}
			}
			return ($Item | Out-String -Width 4096).Trim()
		}
		#endregion Formatting Functions
		#region File Operations Functions
		function Write-LogToFile {
			[CmdletBinding()]
			param(
				[Parameter(Mandatory)][string]$ContentToWrite,
				[double]$LogRotationSizeMB = 1
			)
			if (-not $LogFilePath) {
				if ($script:ShowInternalVerbose) {
					Write-Verbose "LogFilePath not specified."
				}
				return
			}
			try {
				$parentDir = Split-Path -Path $LogFilePath -Parent -ErrorAction Stop
				if (-not (Test-Path -Path $parentDir -PathType Container)) {
					if ($script:ShowInternalVerbose) {
						Write-Verbose "Creating log directory: $parentDir"
					}
					New-Item -Path $parentDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
				}
				$fileExists = Test-Path -Path $LogFilePath -PathType Leaf
				# Handle log rotation if needed
				if ($LogRotationSizeMB -gt 0 -and $fileExists) {
					$logFileItem = Get-Item -Path $LogFilePath -ErrorAction SilentlyContinue
					# Convert bytes to MB for comparison - using raw bytes is more accurate
					$fileSizeInBytes = $logFileItem.Length
					$thresholdInBytes = $LogRotationSizeMB * 1MB
					if ($script:ShowInternalVerbose) {
						Write-Verbose "File size: $fileSizeInBytes bytes, threshold: $thresholdInBytes bytes"
					}
					if ($logFileItem -and $fileSizeInBytes -ge $thresholdInBytes) {
						$backupTimestamp = Get-Date -Format 'yyyyMMdd_HHmmssfff'
						$fileInfo = [System.IO.FileInfo]::new($LogFilePath)
						$backupPath = Join-Path -Path $fileInfo.DirectoryName -ChildPath ($fileInfo.BaseName + "_" + $backupTimestamp + $fileInfo.Extension)
						if ($script:ShowInternalVerbose) {
							Write-Verbose "Log rotation. Rotating '$LogFilePath' to '$backupPath'"
						}
                        # Perform the rotation
						Copy-Item -Path $LogFilePath -Destination $backupPath -Force -ErrorAction SilentlyContinue
						Remove-Item -Path $LogFilePath -Force -ErrorAction SilentlyContinue
                        # Track rotated files for testing
                        $script:RotatedFiles += $backupPath
                        $script:LogRotated = $true
                        if ($script:ShowInternalVerbose) {
                            Write-Verbose "Rotated file count: $($script:RotatedFiles.Count)"
                        }
						$fileExists = $false
					}
				}
				# Write or append to file
				if ($ForcedLogFile -or -not $fileExists) {
					if ($script:ShowInternalVerbose) {
						Write-Verbose "Writing (overwrite/new): $LogFilePath"
					}
					[System.IO.File]::WriteAllText($LogFilePath, $ContentToWrite, $script:Utf8NoBomEncoding)
				}
				else {
					if ($script:ShowInternalVerbose) {
						Write-Verbose "Appending: $LogFilePath"
					}
					$contentToAppend = if ($fileExists -and (Get-Item $LogFilePath).Length -gt 0) { "`r`n" + $ContentToWrite } else { $ContentToWrite }
					[System.IO.File]::AppendAllText($LogFilePath, $contentToAppend, $script:Utf8NoBomEncoding)
				}
                # For testing purposes, track the write
                $script:LogRecordCount++
			}
			catch {
				Write-Warning "Failed write log '$LogFilePath': $($_.Exception.Message)"
			}
		}
		#endregion File Operations Functions
		#region Item Processing Functions
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
				[bool]$CurrentNoErrorLookup
			)
			if ($null -eq $ItemToProcess -and $CurrentLevel -ne 'ERROR') {
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Processing null input."
				}
			}
			# Prepare basic log entry
			$timestamp = Get-Date -Format $CurrentDateFormat
			$logEntryBase = [PSCustomObject]@{
				Timestamp = $timestamp
				Level     = $CurrentLevel
				Message   = $null
			}
			$originalItem = $ItemToProcess
			$errorDetails = $null
			$effectiveItem = $ItemToProcess
			# Handle null/empty items
			if (($null -eq $effectiveItem -or $effectiveItem -eq "") -and $CurrentLevel -ne 'ERROR') {
				$effectiveItem = ""
			}
			# Special handling for ERROR level
			if ($CurrentLevel -eq 'ERROR') {
				$err = Get-ErrorToProcess -CurrentItem $effectiveItem -NoErrorLookupParameter $CurrentNoErrorLookup
				if ($err) {
					$errorDetails = [PSCustomObject]@{
						ExceptionType = $err.Exception.GetType().FullName
						Exception     = $err.Exception.Message
						FullError     = $err
						FailedCode    = $null
						CallerInfo    = $null
					}
					# Set failed code if available
					if ($err.InvocationInfo -and $err.InvocationInfo.Line) {
						$errorDetails.FailedCode = $err.InvocationInfo.Line.Trim()
					}
					else {
						$errorDetails.FailedCode = "N/A"
					}
					# Try to get caller info
					try {
						$cs = Get-PSCallStack -EA SilentlyContinue
						if ($cs) {
							$cf = $null
							for ($i = 1; $i -lt $cs.Count; $i++) {
								if ($cs[$i].Command -ne 'New-Log' -and $cs[$i].Command -ne '<ScriptBlock>') {
									$cf = $cs[$i]
									break
								}
							}
							if ($cf) {
								$sn = if ($cf.ScriptName) {
									[IO.Path]::GetFileName($cf.ScriptName)
								}
								else {
									'<NoScript>'
								}
								$fn = if ($cf.FunctionName -ne '<ScriptBlock>') {
									"Function [$($cf.FunctionName)]"
								}
								else {
									$sn
								}
								$errorDetails.CallerInfo = "File [$sn], Line [$($cf.ScriptLineNumber)], Context [$fn]"
							}
						}
					}
					catch {
						if ($script:ShowInternalVerbose) {
							Write-Verbose "Could not get caller info: $($_.Exception.Message)"
						}
					}
					$logEntryBase | Add-Member -MemberType NoteProperty -Name ErrorDetails -Value $errorDetails -Force
					if (($originalItem -eq $null -or $originalItem -eq "") -or $originalItem -is [System.Management.Automation.ErrorRecord]) {
						$effectiveItem = $errorDetails.Exception
					}
				}
				elseif ($null -eq $effectiveItem -or $effectiveItem -eq "") {
					$effectiveItem = "<ERROR Logged with Null/Empty Message>"
					Write-Warning "ERROR level specified but no specific error context found for null/empty input."
				}
			}
			# Format the message
			$formattedTextMessage = Format-ItemForTextOutput -Item $effectiveItem
			if ($originalItem -is [string] -or $null -eq $originalItem -or $CurrentLevel -eq 'ERROR') {
				$logEntryBase.Message = $formattedTextMessage
			}
			else {
				$logEntryBase.Message = $originalItem
			}
			# Console output
			if (-not $CurrentNoConsole) {
				$colorCode = $CurrentLevelColors[$CurrentLevel].ANSI
				$psColor = $CurrentLevelColors[$CurrentLevel].PS
				$ansiReset = "`e[0m"
				$consolePrefix = if ($CurrentIsPSCore) {
					"`e[34m[$timestamp]$ansiReset`e[${colorCode}m[$CurrentLevel]$ansiReset"
				} else {
					"[$timestamp][$CurrentLevel]"
				}
				Write-PrefixedLinesToConsole -Prefix $consolePrefix `
					-Content $formattedTextMessage `
					-Color $psColor `
					-IsPSCore $CurrentIsPSCore
				# Additional error details for console
				if ($CurrentLevel -eq 'ERROR' -and $errorDetails) {
					$errorPrefix = if ($CurrentIsPSCore) {
						"`e[34m[$timestamp]$ansiReset`e[${colorCode}m[ERROR Detail]$ansiReset"
					} else {
						"[$timestamp][ERROR Detail]"
					}
					Write-PrefixedLinesToConsole -Prefix $errorPrefix `
						-Content "Exception: $($errorDetails.Exception)" `
						-Color $psColor `
						-IsPSCore $CurrentIsPSCore
					if ($errorDetails.CallerInfo) {
						Write-PrefixedLinesToConsole -Prefix $errorPrefix `
							-Content "Caller: $($errorDetails.CallerInfo)" `
							-Color $psColor `
							-IsPSCore $CurrentIsPSCore
					}
					if ($errorDetails.FailedCode -ne 'N/A') {
						Write-PrefixedLinesToConsole -Prefix $errorPrefix `
							-Content "Code: $($errorDetails.FailedCode)" `
							-Color $psColor `
							-IsPSCore $CurrentIsPSCore
					}
				}
			}
			# File logging
			if ($LogFilePath) {
				$fileContent = $null
				if ($LogFormat -eq "JSON") {
					if ($originalItem -is [Management.Automation.ErrorRecord]) {
						$logEntryBase.Message = $errorDetails.Exception
					}
					else {
						$logEntryBase.Message = $originalItem
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
				else { # TEXT format
					$fileContent = "[$timestamp][$CurrentLevel]"
					if (-not [string]::IsNullOrEmpty($formattedTextMessage)) {
						$fileContent += " $formattedTextMessage"
					}
					if ($CurrentLevel -eq 'ERROR' -and $errorDetails) {
						$fileContent += "`r`n[$timestamp][ERROR Detail] Exception: $($errorDetails.Exception)"
						if ($errorDetails.CallerInfo) {
							$fileContent += "`r`n[$timestamp][ERROR Detail] CallerInfo: $($errorDetails.CallerInfo)"
						}
						if ($errorDetails.FailedCode -ne 'N/A') {
							$fileContent += "`r`n[$timestamp][ERROR Detail] FailedCode: $($errorDetails.FailedCode)"
						}
						if ($errorDetails.FullError.ScriptStackTrace) {
							$fileContent += "`r`n[$timestamp][ERROR Detail] StackTrace: $($errorDetails.FullError.ScriptStackTrace -replace '[\r\n]+',' | ')"
						}
					}
				}
				Write-LogToFile -ContentToWrite $fileContent -LogRotationSizeMB $LogRotationSizeMB
			}
			# Return object if requested
			if ($CurrentReturnObject) {
                # Add UsedANSI property for testing
                if (-not $CurrentNoConsole -and $CurrentIsPSCore) {
                    $logEntryBase | Add-Member -MemberType NoteProperty -Name UsedANSI -Value $true -Force
                }
				$logEntryBase.Message = $originalItem
                # Add properties for testing
                $logEntryBase | Add-Member -MemberType NoteProperty -Name LogRecordCount -Value $script:LogRecordCount -Force
                # Always add the RotatedFiles property if log rotation is enabled
                if ($LogRotationSizeMB -gt 0) {
                    if ($script:LogRotated -or $script:RotatedFiles.Count -gt 0) {
                        $logEntryBase | Add-Member -MemberType NoteProperty -Name RotatedFiles -Value $script:RotatedFiles -Force
                    }
                }
				return $logEntryBase
			}
			else {
				return $null
			}
		}
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
			$timestamp = Get-Date -Format $CurrentDateFormat
			$objectCount = $ItemsToProcess.Count
			# File logging for grouped items
			if ($LogFilePath) {
				if ($LogFormat -eq "JSON") {
					if ($script:ShowInternalVerbose) {
						Write-Verbose "Logging $objectCount grouped items as JSON Lines."
					}
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
				else { # TEXT format
					if ($script:ShowInternalVerbose) {
						Write-Verbose "Logging $objectCount grouped items as TEXT table."
					}
					$tableString = ($ItemsToProcess | Format-Table -AutoSize | Out-String).Trim()
					$fileContent = "[$timestamp][$CurrentLevel] Start of grouped items ($objectCount):`r`n"
					$tableLines = $tableString -split "`r?`n"
					foreach ($line in $tableLines) {
						$fileContent += "[$timestamp][$CurrentLevel] $line`r`n"
					}
					$fileContent += "[$timestamp][$CurrentLevel] End of grouped items."
					Write-LogToFile -ContentToWrite $fileContent.TrimEnd()
				}
			}
			# Console output for grouped items
			if (-not $CurrentNoConsole) {
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Displaying $objectCount grouped items."
				}
				$colorCode = $CurrentLevelColors[$CurrentLevel].ANSI
				$psColor = $CurrentLevelColors[$CurrentLevel].PS
				$ansiReset = "`e[0m"
				$consolePrefix = if ($CurrentIsPSCore) {
					"`e[34m[$timestamp]$ansiReset`e[${colorCode}m[$CurrentLevel]$ansiReset"
				} else {
					"[$timestamp][$CurrentLevel]"
				}
				Write-PrefixedLinesToConsole -Prefix $consolePrefix `
					-Content "Displaying $objectCount grouped items:" `
					-Color $psColor `
					-IsPSCore $CurrentIsPSCore
				$tableString = ($ItemsToProcess | Format-Table -AutoSize | Out-String).Trim()
				Write-PrefixedLinesToConsole -Prefix $consolePrefix `
					-Content $tableString `
					-Color $psColor `
					-IsPSCore $CurrentIsPSCore
			}
			# Return summary object if requested
			if ($CurrentReturnObject) {
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Returning summary object for $objectCount grouped items."
				}
				$groupedResult = [PSCustomObject]@{
					Timestamp    = $timestamp
					Level        = $CurrentLevel
					Message      = "Collection of $objectCount objects."
					GroupedItems = $ItemsToProcess
				}
                # Add properties for testing
                $groupedResult | Add-Member -MemberType NoteProperty -Name LogRecordCount -Value $script:LogRecordCount -Force
                # Always add the RotatedFiles property if log rotation is enabled
                if ($LogRotationSizeMB -gt 0) {
                    if ($script:LogRotated -or $script:RotatedFiles.Count -gt 0) {
                        $groupedResult | Add-Member -MemberType NoteProperty -Name RotatedFiles -Value $script:RotatedFiles -Force
                    }
                }
                if ($CurrentIsPSCore) {
                    $groupedResult | Add-Member -MemberType NoteProperty -Name UsedANSI -Value $true -Force
                }
				return $groupedResult
			}
			else {
				return $null
			}
		}
		#endregion Item Processing Functions
		#endregion Helper Functions
	} # End Begin block
	Process {
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
			# Collection modes
			if ($GroupObjects) {
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Collect G #$($script:PipelineItemCounter)."
				}
				$script:ObjectCollection.Add($currentItem)
				return
			}
			elseif (-not $NoAutoGroup) {
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Collect A #$($script:PipelineItemCounter)."
				}
				$script:ObjectCollection.Add($currentItem)
				return
			}
			# Direct processing for pipeline with -NoAutoGroup
			if ($script:ShowInternalVerbose) {
				Write-Verbose "Proc NAG #$($script:PipelineItemCounter)."
			}
		}
		# Process immediately if not pipeline OR pipeline with -NoAutoGroup
		try {
			if ($script:ShowInternalVerbose) {
				Write-Verbose "Proc single item Process."
			}
			$result = Process-SingleLogItem -ItemToProcess $currentItem `
				-CurrentLevel $Level `
				-CurrentDateFormat $DateFormat `
				-CurrentLevelColors $script:LevelColors `
				-CurrentIsPSCore $script:isPSCore `
				-CurrentNoConsole $NoConsole `
				-CurrentReturnObject $ReturnObject `
				-CurrentNoErrorLookup $NoErrorLookup -ErrorAction Stop
			if ($ReturnObject -and $null -ne $result) {
				Write-Output $result
			}
		}
		catch {
			Write-Error "Error processing single item Process: $($_.Exception.Message)"
		}
	} # End Process block
	End {
		if ($script:ShowInternalVerbose) {
			Write-Verbose "Entering End block. Items: $($script:PipelineItemCounter)."
		}
		try {
			# Skip VERBOSE level messages if not enabled
			if ($Level -eq "VERBOSE") {
				$prefs = try { Get-Variable -Name VerbosePreference -Scope 1 -ValueOnly -ErrorAction Stop } catch { $null }
				if ($prefs -ne [Management.Automation.ActionPreference]::Continue) {
					if ($script:ShowInternalVerbose) {
						Write-Verbose "Skip final VERBOSE."
					}
					return
				}
			}
			$processedGroup = $false
			$result = $null
			# Handle group objects
			if ($GroupObjects -and $script:ObjectCollection.Count -gt 0) {
				$itemCount = $script:ObjectCollection.Count
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Proc G ($itemCount) End."
				}
				try {
					$result = Process-GroupedLogItems -ItemsToProcess $script:ObjectCollection `
						-CurrentLevel $Level `
						-CurrentDateFormat $DateFormat `
						-CurrentLevelColors $script:LevelColors `
						-CurrentIsPSCore $script:isPSCore `
						-CurrentNoConsole $NoConsole `
						-CurrentReturnObject $ReturnObject -ErrorAction Stop
					$processedGroup = $true
				}
				catch {
					Write-Error "Error processing GroupedLogItems: $($_.Exception.Message)"
					$processedGroup = $false
				}
			}
			# Handle auto-grouped pipeline input
			elseif ($script:IsPipelineInput -and -not $NoAutoGroup -and $script:PipelineItemCounter -gt 1) {
				$itemCount = $script:ObjectCollection.Count
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Proc A>1 ($itemCount) End."
				}
				try {
					$result = Process-GroupedLogItems -ItemsToProcess $script:ObjectCollection `
						-CurrentLevel $Level `
						-CurrentDateFormat $DateFormat `
						-CurrentLevelColors $script:LevelColors `
						-CurrentIsPSCore $script:isPSCore `
						-CurrentNoConsole $NoConsole `
						-CurrentReturnObject $ReturnObject -ErrorAction Stop
					$processedGroup = $true
				}
				catch {
					Write-Error "Error processing GroupedLogItems: $($_.Exception.Message)"
					$processedGroup = $false
				}
			}
			# Handle single pipeline item with auto-grouping enabled
			elseif ($script:IsPipelineInput -and -not $NoAutoGroup -and $script:PipelineItemCounter -eq 1) {
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Proc A=1 (1) End."
				}
				try {
					$result = Process-SingleLogItem -ItemToProcess $script:ObjectCollection[0] `
						-CurrentLevel $Level `
						-CurrentDateFormat $DateFormat `
						-CurrentLevelColors $script:LevelColors `
						-CurrentIsPSCore $script:isPSCore `
						-CurrentNoConsole $NoConsole `
						-CurrentReturnObject $ReturnObject `
						-CurrentNoErrorLookup $NoErrorLookup -ErrorAction Stop
					$processedGroup = $true
				}
				catch {
					Write-Error "Error processing single item Process: $($_.Exception.Message)"
					$processedGroup = $false
				}
			}
			# Return result object if requested
			if ($processedGroup -and $ReturnObject -and $null -ne $result) {
				Write-Output $result
			}
		}
		catch {
			Write-Error "Error during End block: $($_.Exception.Message)"
		}
		finally {
			# Restore console encoding
			if ($script:OriginalConsoleEncoding -ne $null -and [Console]::OutputEncoding -ne $script:OriginalConsoleEncoding) {
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Restore console encoding."
				}
				try {
					[Console]::OutputEncoding = $script:OriginalConsoleEncoding
				}
				catch {
					Write-Warning "Failed restore console encode: $($_.Exception.Message)"
				}
			}
			# Clear object collection
			if ($script:ObjectCollection) {
				$script:ObjectCollection.Clear()
				if ($script:ShowInternalVerbose) {
					Write-Verbose "Cleared object collection."
				}
			}
		}
	} # End End block
}