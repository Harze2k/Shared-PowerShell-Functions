function Start-WithProgress {
	<#
	.SYNOPSIS
		Executes a command or waits for a condition with a progress display and timeout,
		optimizing for transcript cleanliness and cross-version PowerShell compatibility.
	.DESCRIPTION
		This function provides two main modes: executing an external program/script (`FilePath` mode)
		or waiting for a PowerShell scriptblock condition to become true (`Condition` mode).
		It displays a progress bar representing elapsed time versus a specified timeout.
		Key Features & Behavior:
		- Transcript-Aware Progress:
		- In PowerShell 7+, if a transcript is active, it uses the native `Write-Progress` cmdlet. This ensures
			the transcript file remains clean of progress bar visual updates, as `Write-Progress` visuals are
			not logged by the host. The appearance of the native `Write-Progress` bar is standard for PowerShell.
		- In PowerShell 5.1, if a transcript is active AND running in a console (not ISE) AND ANSI escape codes are enabled,
			it uses a custom `Write-ProgressBar` that employs `[System.Console]::Write()`. This method also aims
			to keep the PS5.1 transcript clean of visual updates. The custom bar provides rich color.
			The *final* state of the progress bar in this scenario is explicitly written to the transcript via Write-Host.
		- In all other scenarios (e.g., no transcript active, PowerShell ISE, or when `DontUseEscapeCodes` is specified),
			the custom `Write-ProgressBar` (using `Write-Host`) is used, providing colored or basic visuals directly
			in the console. In these cases, if a transcript is active, it will capture the `Write-Host` output,
			including the final progress bar state.
		- Asynchronous Operations: Uses background jobs for `FilePath` mode to run processes asynchronously,
			allowing the main thread to update the progress display.
		- Customizable: Timeout, sleep interval, progress bar width, and process window styles are configurable.
		- Logging: Relies on a `New-Log` function (must be defined externally) for status messages,
			which are always written to the transcript if active.
		- Return Values: Provides status codes (e.g., process exit code, 0 for condition met, -1 for timeout).
			For `JustWait` mode with a timeout, it returns no value to the pipeline to keep transcripts clean.
	.PARAMETER FilePath
		The path to the executable file or script to be run (for the 'FilePath' parameter set).
		Aliases: Path
	.PARAMETER Arguments
		A string containing the arguments to pass to the executable specified by FilePath.
		Aliases: Parameters, ArgumentList
	.PARAMETER TimeoutInSeconds
		The maximum time in seconds to wait for the process to complete or the condition to become true.
		If the timeout is reached, the process/job is terminated, or the condition wait stops. Default is 60 seconds.
	.PARAMETER SleepInSeconds
		The interval in seconds at which the progress display updates and checks are performed.
		Must be less than or equal to TimeoutInSeconds. Default is 1 second.
	.PARAMETER NoNewWindow
		If specified, the process started by Start-Process (internally) will not create a new window.
		Cannot be used with -WindowStyle if -WindowStyle is anything other than 'Hidden'.
	.PARAMETER PassThru
		If specified when using FilePath mode, this function will attempt to return the exit code
		of the executed process. This function *always* returns an exit code or a status code.
	.PARAMETER RedirectStandardOutput
		Specifies a file path to redirect the standard output (stdout) of the process to.
	.PARAMETER JustWait
		Used with the 'Condition' parameter set. If specified, when the condition times out, it logs a
		"wait is over" message and returns no value to the pipeline (to keep transcripts clean).
		If the condition is met, it behaves like normal Condition mode.
	.PARAMETER RedirectStandardError
		Specifies a file path to redirect the standard error (stderr) of the process to.
	.PARAMETER WindowStyle
		Specifies the window style for the process. Valid values: Normal, Hidden, Minimized, Maximized. Default: Hidden.
	.PARAMETER ProgressBarWidth
		The width of the character-based progress bar (used by the custom bar and to format the native bar's status text). Default: 50.
	.PARAMETER ISEMode
		[switch] Forces the custom progress bar to use a mode compatible with the PowerShell ISE.
		If not specified, the function attempts to auto-detect ISE.
	.PARAMETER Condition
		A scriptblock that evaluates to $true or $false (for the 'Condition' parameter set).
		The function will wait until this condition is $true or the timeout is reached.
	.PARAMETER Message
		A message to display (via New-Log) when starting the wait. For `Condition` mode, this is also used
		as the `-Activity` for `Write-Progress` if native progress is used.
	.PARAMETER SuccessMessage
		A message to display (via New-Log) when the condition in `Condition` mode successfully evaluates to $true.
	.PARAMETER DontUseEscapeCodes
		[switch] If specified, forces the custom progress bar to use simple block characters ('#', '░') and
		disables ANSI escape codes, even in consoles that might support them.
	.OUTPUTS
		System.Int32
		- In FilePath mode: Returns the exit code of the executed process.
			Returns -1 if the process timed out or an internal error occurred.
			Returns -11 if there was a fatal error generating the PS5.1 batch file wrapper.
		System.Int32 or $null
		- In Condition mode:
		- Returns 0 if the condition became true.
		- Returns -1 if the timeout was reached before the condition became true (unless -JustWait is used).
		- If -JustWait is used and timeout is reached, returns $null (no pipeline output).
	.EXAMPLE
		# Assuming New-Log, Test-TranscriptActive, and Write-ProgressBar are defined.
		# Example 1: Run cmd.exe, custom bar if no transcript, native if transcript (PS7+).
		Start-WithProgress -FilePath 'cmd.exe' -Arguments '/c timeout /t 3 /nobreak >nul && echo CMD Done' -TimeoutInSeconds 5
	.EXAMPLE
		# Example 2: Wait for a file, showing progress. Transcript behavior as above.
		# $null = Start-Job { Start-Sleep 3; Set-Content -Path C:\Temp\myfile.txt -Value "created" }
		Start-WithProgress -Condition { Test-Path C:\Temp\myfile.txt } -Message "Waiting for C:\Temp\myfile.txt" -TimeoutInSeconds 10
	.EXAMPLE
		# Example 3: Just wait, with custom message. Output to transcript is clean if JustWait times out.
		# Pipe to Out-Null if you don't want the $null from JustWait timeout on console either.
		Start-WithProgress -JustWait -Message "Pausing for 5 seconds..." -TimeoutInSeconds 5 | Out-Null
	.NOTES
		AUTHOR: Harze2k
		VERSION: 4.3 (Ensures final progress bar is logged to PS5.1 transcript when System.Console.Write was used for live updates. Not for PS7+ atm.)
			-Now work with running transcript!
			-Fixed to maintain the actual progress state and remaining time when the progress bar ends.
		Dependencies:
		- `New-Log`: Must be defined in the calling scope for logging.
		- `Test-TranscriptActive`: Must be defined (uses reflection method) for transcript-aware behavior.
		- `Write-ProgressBar`: Must be defined for custom progress display.
		The `Start-WithProgress` function itself is designed to orchestrate these components.
		The PS5.1 batch file wrapper for `FilePath` mode uses OEM encoding for stdout/stderr redirection;
		PS7+ uses UTF-8.
	#>
	[CmdletBinding(DefaultParameterSetName = 'FilePath')]
	param (
		[Parameter(Position = 0, ParameterSetName = 'FilePath')][Alias('Path')][string]$FilePath,
		[Parameter(Position = 1)][Alias('Parameters', 'ArgumentList')][string]$Arguments = '',
		[Parameter(Position = 2)][int]$TimeoutInSeconds = 60,
		[Parameter(Position = 3)][ValidateRange(1, [int]::MaxValue)][int]$SleepInSeconds = 1,
		[switch]$NoNewWindow,
		[switch]$PassThru,
		[string]$RedirectStandardOutput,
		[switch]$JustWait,
		[string]$RedirectStandardError,
		[Parameter(Position = 4)][ValidateSet('Normal', 'Hidden', 'Minimized', 'Maximized')][string]$WindowStyle = 'Hidden',
		[Parameter(Position = 5)][int]$ProgressBarWidth = 50,
		[switch]$ISEMode,
		[Parameter(Position = 0, ParameterSetName = 'Condition')][scriptblock]$Condition,
		[Parameter(Position = 1, ParameterSetName = 'Condition')][string]$Message = "Awaiting condition...",
		[Parameter(Position = 2, ParameterSetName = 'Condition')][string]$SuccessMessage,
		[switch]$DontUseEscapeCodes
	)
	function Write-ProgressBar {
		<#
		.SYNOPSIS
			Displays a customizable, single-line progress bar in the console.
		.DESCRIPTION
			This function renders a progress bar that updates on the same line using carriage returns.
			It supports ANSI true-color for modern terminals and falls back to basic characters and
			standard console colors for the PowerShell ISE or when escape codes are disabled.
			Crucially, for PowerShell 5.1 when a transcript is active and running in a non-ISE console,
			it attempts to use [System.Console]::Write() for its output. This typically prevents
			the progress bar's visual updates from being logged to the PS5.1 transcript file, keeping it clean.
			In other scenarios (e.g., no transcript, PowerShell 7+, ISE), it uses Write-Host.
			This function relies on an external 'Test-TranscriptActive' function to determine transcript status.
		.PARAMETER Completed
			The number of completed units for the progress bar.
		.PARAMETER IsPsISE
			[switch] Indicates if the current host is PowerShell ISE. This affects color and character rendering.
			This is typically passed by a calling function like Start-WithProgress.
		.PARAMETER Width
			[int] The character width of the progress bar itself (excluding text like percentage, time). Default is 40.
		.PARAMETER TimeElapsed
			[TimeSpan] The time elapsed so far for the operation being monitored.
		.PARAMETER RemainingTime
			[TimeSpan] The estimated time remaining for the operation.
		.PARAMETER DontUseEscapeCodes
			[switch] If specified, forces the use of basic characters ('#', '░', ' ') and disables ANSI escape codes,
			even in terminals that might support them.
		.NOTES
			- Requires 'Test-TranscriptActive' function to be available in the scope for optimal transcript-aware behavior in PS5.1.
			- The visual appearance and transcript behavior are highly dependent on the PowerShell version, host (console vs. ISE),
			- active transcript status, and the -DontUseEscapeCodes switch.
			- ANSI true-color sequences (`\e[38;2;r;g;bm`) are used for rich color display.
			- Includes an internal helper 'Get-ConsoleColorForISE' to map percentages to basic System.ConsoleColor names for ISE compatibility.
		#>
		param(
			[int]$Completed,
			[bool]$IsPsISE,
			[int]$Width = 40,
			[TimeSpan]$TimeElapsed,
			[TimeSpan]$RemainingTime,
			[switch]$DontUseEscapeCodes,
			[switch]$PassThruStringForLog
		)
		function Get-ConsoleColorForISE {
			param([double]$Percentage)
			switch ($Percentage) {
				{ $_ -lt 16.7 } { return 'Green' }
				{ $_ -lt 33.3 } { return 'DarkGreen' }
				{ $_ -lt 50 } { return 'Yellow' }
				{ $_ -lt 66.7 } { return 'DarkYellow' }
				{ $_ -lt 83.3 } { return 'DarkRed' }
				default { return 'Red' }
			}
		}
		$completedWidth = [Math]::Max(0, [Math]::Min($Completed, $Width))
		$remainingWidth = $Width - $completedWidth
		$percentage = 0.0
		if ($Width -gt 0) {
			$percentage = [Math]::Round(($completedWidth / $Width) * 100, 1)
		}
		# --- Build the text part of the progress string (common for all outputs) ---
		$progressDetailsString = "] {0:0.0}% Elapsed: {1:hh\:mm\:ss} (Remaining: {2:hh\:mm\:ss})     " -f $percentage, $TimeElapsed, $RemainingTime
		# --- Logic for -PassThruStringForLog ---
		if ($PassThruStringForLog.IsPresent) {
			# For logging, always return a simple character representation, no ANSI.
			$logBarChars = ('█' * $completedWidth) + ('░' * $remainingWidth) # Or use '#' and ' ' if preferred for logs
			return "[{0}{1}" -f $logBarChars, $progressDetailsString # Return the simple string
		}
		# --- Console Printing Logic (if not -PassThruStringForLog) ---
		$isTranscriptCurrentlyActive = $false
		if (Get-Command -Name Test-TranscriptActive -ErrorAction SilentlyContinue) {
			$isTranscriptCurrentlyActive = Test-TranscriptActive
		}
		$isPSCoreVersion = $PSVersionTable.PSVersion.Major -ge 6
		$useSystemConsoleWrite = $isTranscriptCurrentlyActive -and (-not $isPSCoreVersion) -and (-not $IsPsISE) -and (-not $DontUseEscapeCodes.IsPresent)
		$barCharsForDisplay = "" # Different variable for display bar
		$escapeChar = [char]27
		if ($useSystemConsoleWrite) {
			# PS5.1, Transcript, Console, ANSI -> System.Console.Write
			for ($i = 0; $i -lt $Width; $i++) {
				if ($i -lt $completedWidth) {
					$colorPercentage = ($i / $Width)
					$greenValue = [Math]::Max(0, [Math]::Min(255, [int](255 * (1 - $colorPercentage))))
					$redValue = [Math]::Max(0, [Math]::Min(255, [int](255 * $colorPercentage)))
					$barCharsForDisplay += "${escapeChar}[38;2;${redValue};${greenValue};0m█${escapeChar}[0m"
				}
				else {
					$barCharsForDisplay += "${escapeChar}[38;2;100;100;100m░${escapeChar}[0m"
				}
			}
			$fullLineToPrint = "`r${escapeChar}[0m[{0}{1}" -f $barCharsForDisplay, $progressDetailsString
			[System.Console]::Write($fullLineToPrint)
		}
		else {
			# All other scenarios -> Write-Host
			if ($IsPsISE -or $DontUseEscapeCodes.IsPresent) {
				if ($IsPsISE) {
					Write-Host -NoNewline "`r${escapeChar}[0m["
					for ($i = 0; $i -lt $completedWidth; $i++) {
						$consoleColor = Get-ConsoleColorForISE(($i / $Width) * 100)
						Write-Host -NoNewline -ForegroundColor $consoleColor '#'
					}
					for ($i = 0; $i -lt $remainingWidth; $i++) {
						Write-Host -NoNewline ' '
					}
					Write-Host -NoNewline $progressDetailsString
					return
				}
				else {
					# DontUseEscapeCodes but not ISE
					for ($i = 0; $i -lt $completedWidth; $i++) {
						$barCharsForDisplay += '#'
					}
					for ($i = 0; $i -lt $remainingWidth; $i++) {
						$barCharsForDisplay += '░'
					}
				}
			}
			else {
				# ANSI for Write-Host
				for ($i = 0; $i -lt $Width; $i++) {
					if ($i -lt $completedWidth) {
						$colorPercentage = ($i / $Width)
						$greenValue = [Math]::Max(0, [Math]::Min(255, [int](255 * (1 - $colorPercentage))))
						$redValue = [Math]::Max(0, [Math]::Min(255, [int](255 * $colorPercentage)))
						$barCharsForDisplay += "${escapeChar}[38;2;${redValue};${greenValue};0m█${escapeChar}[0m"
					}
					else {
						$barCharsForDisplay += "${escapeChar}[38;2;100;100;100m░${escapeChar}[0m"
					}
				}
			}
			$fullLineToPrint = "`r${escapeChar}[0m[{0}{1}" -f $barCharsForDisplay, $progressDetailsString
			Write-Host -NoNewline $fullLineToPrint
		}
	}
	$tempResources = @{ TempFiles = @(); Job = $null }
	if (-not $PSBoundParameters.ContainsKey('ISEMode')) {
		$ISEMode = $Host.Name -eq 'Windows PowerShell ISE Host'
	}
	$timeoutTimeSpan = [TimeSpan]::FromSeconds($TimeoutInSeconds)
	$isTranscriptActiveForContext = $false
	if (Get-Command -Name Test-TranscriptActive -ErrorAction SilentlyContinue) {
		$isTranscriptActiveForContext = Test-TranscriptActive
	}
	else {
		New-Log "Start-WithProgress: Test-TranscriptActive function not found. Will abort." -Level WARNING
		return
	}
	$isPSCoreVersion = $PSVersionTable.PSVersion.Major -ge 6
	$useNativeWriteProgress = $isTranscriptActiveForContext -and $isPSCoreVersion
	$progressIdForNative = Get-Random
	# Condition under which Write-ProgressBar uses System.Console.Write (and thus bypasses PS5.1 transcript for live updates)
	$customBarVisualsBypassTranscript = $isTranscriptActiveForContext -and (-not $isPSCoreVersion) -and (-not $ISEMode) -and (-not $DontUseEscapeCodes.IsPresent)
	if ($JustWait) {
		$Condition = { $false }
		if ($PSCmdlet.ParameterSetName -eq 'Condition' -and (-not $PSBoundParameters.ContainsKey('Message') -or [string]::IsNullOrWhiteSpace($Message))) {
			$Message = "Waiting for $TimeoutInSeconds seconds..."
		}
	}
	#endregion
	try {
		# --- Parameter Validation ---
		if ($SleepInSeconds -gt $TimeoutInSeconds) {
			New-Log 'SleepInSeconds must be <= TimeoutInSeconds.' -Level WARNING
			return -1
		}
		if (-not (Get-Command -Name New-Log -ErrorAction SilentlyContinue)) {
			Write-Error "New-Log function not found. Will abort."
			return -1
		}
		if (-not $useNativeWriteProgress -and -not (Get-Command -Name Write-ProgressBar -ErrorAction SilentlyContinue)) {
			New-Log "Write-ProgressBar function not found. Will abort." -Level WARNING
			return -1
		}
		# --- Condition-Based Progress ---
		if ($PSCmdlet.ParameterSetName -eq 'Condition') {
			if ($null -eq $Condition) {
				New-Log "Condition parameter is null." -Level WARNING
				if ($useNativeWriteProgress) {
					Write-Progress -Id $progressIdForNative -Activity $Message -Completed -ErrorAction SilentlyContinue
				}
				return -1
			}
			New-Log "$Message"
			$startTime = Get-Date
			$conditionErrorCount = 0
			$hasShownProgress = $false
			$currentCompletion = 0
			$lastRemainingTime = [TimeSpan]::FromSeconds($TimeoutInSeconds)
			while ($true) {
				$currentTime = Get-Date
				$timeElapsed = $currentTime - $startTime
				$remainingTime = $timeoutTimeSpan - $timeElapsed
				if ($remainingTime.TotalSeconds -lt 0) {
					$remainingTime = [TimeSpan]::Zero
				}
				$lastRemainingTime = $remainingTime
				$percentComplete = if ($TimeoutInSeconds -gt 0) {
					[Math]::Min(100, ($timeElapsed.TotalSeconds / $TimeoutInSeconds) * 100)
				}
				else {
					100
				}
				$currentCompletion = [Math]::Round(($percentComplete / 100) * $ProgressBarWidth)
				$hasShownProgress = $true
				if ($useNativeWriteProgress) {
					$barTextNative = '[' + ('█' * $currentCompletion) + ('░' * ($ProgressBarWidth - $currentCompletion)) + ']'
					$statusStringNative = "{0} {1:N1}% Elapsed: {2:hh\:mm\:ss} (Remaining: {3:hh\:mm\:ss})" -f $barTextNative, $percentComplete, $timeElapsed, $remainingTime
					Write-Progress -Id $progressIdForNative -Activity $Message -Status $statusStringNative -PercentComplete ([int]$percentComplete) -SecondsRemaining ([int]$remainingTime.TotalSeconds) -EA SilentlyContinue
				}
				elseif (Get-Command -Name Write-ProgressBar -ErrorAction SilentlyContinue) {
					Write-ProgressBar -Completed $currentCompletion -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $timeElapsed -RemainingTime $remainingTime -DontUseEscapeCodes:$DontUseEscapeCodes.IsPresent
				}
				if ($timeElapsed.TotalSeconds -ge $TimeoutInSeconds) {
					if ($useNativeWriteProgress) {
						# Keep current progress rather than marking as completed
						$finalBarTextNative = '[' + ('█' * $currentCompletion) + ('░' * ($ProgressBarWidth - $currentCompletion)) + ']'
						$finalStatusStringNative = "{0} {1:N1}% Elapsed: {2:hh\:mm\:ss} (Remaining: {3:hh\:mm\:ss})" -f $finalBarTextNative, $percentComplete, $timeElapsed, $lastRemainingTime
						Write-Progress -Id $progressIdForNative -Activity $Message -Status $finalStatusStringNative -PercentComplete ([int]$percentComplete) -Completed -ErrorAction SilentlyContinue
					}
					elseif ($hasShownProgress -and (Get-Command -Name Write-ProgressBar -ErrorAction SilentlyContinue)) {
						if (-not $useNativeWriteProgress) {
							# Final visual update - maintain current completion level
							Write-ProgressBar -Completed $currentCompletion -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $timeElapsed -RemainingTime $lastRemainingTime -DontUseEscapeCodes:$DontUseEscapeCodes.IsPresent
							# If visual update bypassed transcript, explicitly log final bar state
							if ($customBarVisualsBypassTranscript) {
								$barChars = ('█' * $currentCompletion) + ('░' * ($ProgressBarWidth - $currentCompletion))
								$finalBarText = "[$barChars] $([Math]::Round($percentComplete, 1))% Elapsed: " + $timeElapsed.ToString("hh\:mm\:ss") + " (Remaining: " + $lastRemainingTime.ToString("hh\:mm\:ss") + ")"
								$res = New-Log "" -Level SUCCESS -NoConsole -ReturnObject
								Write-Information "[$($res.TimeStamp)][$($res.Level)] $finalBarText"
							}
						}
						if ($customBarVisualsBypassTranscript) {
							[System.Console]::WriteLine()
						}
						else {
							Write-Host
						}
					}
					if ($JustWait) {
						New-Log 'The wait is over!' -Level SUCCESS
						return
					}
					else {
						New-Log "Condition timed out: $Message" -Level WARNING
						return -1
					}
				}
				try {
					if ((& $Condition) -eq $true) {
						if ($useNativeWriteProgress) {
							# Keep current progress rather than marking as completed
							$finalBarTextNative = '[' + ('█' * $currentCompletion) + ('░' * ($ProgressBarWidth - $currentCompletion)) + ']'
							$finalStatusStringNative = "{0} {1:N1}% Elapsed: {2:hh\:mm\:ss} (Remaining: {3:hh\:mm\:ss})" -f $finalBarTextNative, $percentComplete, $timeElapsed, $lastRemainingTime
							Write-Progress -Id $progressIdForNative -Activity $Message -Status $finalStatusStringNative -PercentComplete ([int]$percentComplete) -Completed -ErrorAction SilentlyContinue
						}
						elseif ($hasShownProgress -and (Get-Command -Name Write-ProgressBar -ErrorAction SilentlyContinue)) {
							if (-not $useNativeWriteProgress) {
								# Final visual update - maintain current completion
								Write-ProgressBar -Completed $currentCompletion -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $timeElapsed -RemainingTime $lastRemainingTime -DontUseEscapeCodes:$DontUseEscapeCodes.IsPresent
								# If visual update bypassed transcript, explicitly log final bar state
								if ($customBarVisualsBypassTranscript) {
									$finalBarLogString = Write-ProgressBar -Completed $currentCompletion -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $timeElapsed -RemainingTime $lastRemainingTime -DontUseEscapeCodes:$DontUseEscapeCodes.IsPresent -PassThruStringForLog
									if ($finalBarLogString) {
										Write-Host $finalBarLogString
									}
								}
							}
							if ($customBarVisualsBypassTranscript) {
								[System.Console]::WriteLine()
							}
							else {
								Write-Host
							}
						}
						New-Log "Condition met: $Message" -Level SUCCESS
						if ($SuccessMessage) {
							New-Log "$SuccessMessage" -Level SUCCESS
						}
						return 0
					}
				}
				catch {
					$conditionErrorCount++
					if ($conditionErrorCount -le 3) {
						New-Log "Error evaluating condition." -Level ERROR
					}
				}
				Start-Sleep -Seconds $SleepInSeconds
			}
		}
		# --- FilePath-Based Progress (Process Execution) ---
		if ($PSCmdlet.ParameterSetName -eq 'FilePath') {
			if ([string]::IsNullOrWhiteSpace($FilePath)) {
				New-Log "FilePath missing." -Level WARNING
				return -1
			}
			$OriginalFilePath = $FilePath
			if (-not (Test-Path $FilePath -PathType Leaf)) {
				$resolvedCommand = Get-Command $FilePath -EA SilentlyContinue
				if ($null -eq $resolvedCommand) {
					$pathMatch = $env:PATH.Split(';') | ForEach-Object {
						Join-Path $_ $FilePath
					} |	Where-Object {
						Test-Path $_ -PathType Leaf
					} | Select-Object -First 1
					if ($pathMatch) {
						$FilePath = $pathMatch
					}
					else {
						New-Log "FilePath '$OriginalFilePath' not found." -Level WARNING
						return -1
					}
				}
				else {
					$FilePath = $resolvedCommand.Source
				}
			}
			if ($NoNewWindow -and $PSBoundParameters.ContainsKey('WindowStyle') -and $WindowStyle -ne 'Hidden') {
				New-Log "NoNewWindow and WindowStyle conflict. Will abort." -Level WARNING
				return -1
			}
			$executableFileName = Split-Path -Path $FilePath -Leaf
			$progressActivityMessage = "Executing: $executableFileName"
			if ($Arguments) {
				$progressActivityMessage += " ($($Arguments.Substring(0, [Math]::Min($Arguments.Length, 30)))...)"
			}
			$tempOutput = if ($RedirectStandardOutput) {
				$RedirectStandardOutput
			}
			else {
				$tempFile = [System.IO.Path]::GetTempFileName()
				$tempResources.TempFiles += $tempFile
				$tempFile
			}
			$tempError = if ($RedirectStandardError) {
				$RedirectStandardError
			}
			else {
				$tempFile = [System.IO.Path]::GetTempFileName()
				$tempResources.TempFiles += $tempFile
				$tempFile
			}
			if (-not $PSBoundParameters.ContainsKey('RedirectStandardOutput')) {
				Set-Content -Path $tempOutput -Value $null -Encoding UTF8 -ErrorAction SilentlyContinue
			}
			if (-not $PSBoundParameters.ContainsKey('RedirectStandardError')) {
				Set-Content -Path $tempError -Value $null -Encoding UTF8 -ErrorAction SilentlyContinue
			}
			$newLogDefForJob = ''
			try {
				$newLogDefForJob = ${function:New-Log}.ToString()
			}
			catch {
				Write-Error "Error getting New-Log definition."
				return -1
			}
			$scriptBlock = {
				param(
					$ProcessFilePath,
					$ProcessArguments,
					$ProcessOutputFile,
					$ProcessErrorFile,
					$ProcessNoNewWindow,
					$ProcessWindowStyle,
					$IsPs5,
					$PassedNewLogDefinition
				)
				if (-not [string]::IsNullOrWhiteSpace($PassedNewLogDefinition)) {
					Invoke-Expression -Command "function New-Log { $PassedNewLogDefinition }"
				}
				else {
					Write-Error "Job: 'New-Log' definition was not passed to the job."
					"Job: 'New-Log' definition was not passed to the job." | Out-File -FilePath $ProcessErrorFile -Encoding OEM -Append
					return -1
				}
				if ($IsPs5) {
					$tempBatchFile = $null
					try {
						$tempBatchFile = [System.IO.Path]::GetTempFileName() + ".bat"
						$batchLine1 = "@echo off"
						$batchLine2 = ""
						$batchLine3 = "exit /b %errorlevel%"
						$ExecutableNameOnly = (Split-Path -Path $ProcessFilePath -Leaf).ToLowerInvariant()
						$isPotentiallyTimeoutInCmdArgs = $false
						if ($ExecutableNameOnly -eq "cmd.exe" -and $ProcessArguments -match '\btimeout(\.exe)?\b') {
							$isPotentiallyTimeoutInCmdArgs = $true
						}
						if ($ExecutableNameOnly -eq "cmd.exe") {
							if ($isPotentiallyTimeoutInCmdArgs) {
								$batchLine2 = """$ProcessFilePath"" $ProcessArguments 1>""$ProcessOutputFile"" 2>""$ProcessErrorFile"""
							}
							else {
								$batchLine2 = """$ProcessFilePath"" $ProcessArguments <NUL 1>""$ProcessOutputFile"" 2>""$ProcessErrorFile"""
							}
						}
						else {
							$CmdPath = (Get-Command cmd.exe -ErrorAction Stop).Source
							$InnerTargetCommandForCmdC = """""$ProcessFilePath"" $ProcessArguments"""
							$batchLine2 = """$CmdPath"" /S /C $InnerTargetCommandForCmdC <NUL 1>""$ProcessOutputFile"" 2>""$ProcessErrorFile"""
						}
						$BatchContent = "$batchLine1`r`n$batchLine2`r`n$batchLine3`r`n"
						if ([string]::IsNullOrWhiteSpace($batchLine2) -or $BatchContent.Length -lt 20) {
							$fatalMsg = "Job FATAL: Batch file gen failed. Line2: '$batchLine2'"
							New-Log $fatalMsg -Level WARNING
							$fatalMsg | Out-File -FilePath $ProcessErrorFile -Encoding OEM -Append
							return -11
						}
						Set-Content -Path $tempBatchFile -Value $BatchContent -Encoding ASCII -Force
						$startInfoArgs = @{
							FilePath    = $tempBatchFile
							Wait        = $true
							PassThru    = $true
							ErrorAction = 'Stop'
						}
						if ($ProcessNoNewWindow) {
							$startInfoArgs.NoNewWindow = $true
						}
						elseif ($ProcessWindowStyle) {
							$startInfoArgs.WindowStyle = $ProcessWindowStyle
						}
						$process = Start-Process @startInfoArgs
						return $process.ExitCode
					}
					catch {
						$errorMessage = "Job ERROR (PS5 Path): $($_.Exception.Message). TempBatFile: $tempBatchFile"
						if ($tempBatchFile -and (Test-Path $tempBatchFile)) {
							$batchContentOnError = Get-Content $tempBatchFile -Raw -ErrorAction SilentlyContinue
							$errorMessage += "`r`nBatch File Content (`"$tempBatchFile`"):`r`n$batchContentOnError"
						}
						New-Log $errorMessage -Level ERROR
						$errorMessage | Out-File -FilePath $ProcessErrorFile -Encoding OEM -Append
						return -1
					}
					finally {
						if ($tempBatchFile -and (Test-Path $tempBatchFile)) {
							Remove-Item -Path $tempBatchFile -Force -ErrorAction SilentlyContinue
						}
					}
				}
				else {
					# PS 7+
					try {
						$startInfoArgs = @{
							FilePath               = $ProcessFilePath
							ArgumentList           = $ProcessArguments
							Wait                   = $true
							PassThru               = $true
							RedirectStandardOutput = $ProcessOutputFile
							RedirectStandardError  = $ProcessErrorFile
							ErrorAction            = 'Stop'
						}
						if ($ProcessNoNewWindow) {
							$startInfoArgs.NoNewWindow = $true
						}
						elseif ($ProcessWindowStyle) {
							$startInfoArgs.WindowStyle = $ProcessWindowStyle
						}
						$process = Start-Process @startInfoArgs
						return $process.ExitCode
					}
					catch {
						$errMsg = "Job ERROR (PS7+): '$ProcessFilePath'. $($_.Exception.Message)"
						New-Log "Job ERROR (PS7+): Starting process. Details: $errMsg" -Level ERROR
						$errMsg | Out-File -FilePath $ProcessErrorFile -Encoding UTF8 -Append
						return -1
					}
				}
			} # End of $scriptBlock
			New-Log "Executing: `"$FilePath`" ($($Arguments)) Timeout=$TimeoutInSeconds" -Level INFO
			try {
				$job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $FilePath, $Arguments, $tempOutput, $tempError, $NoNewWindow, $WindowStyle, (-not $isPSCoreVersion), $newLogDefForJob
				$tempResources.Job = $job
			}
			catch {
				New-Log "Failed to start job." -Level ERROR
				return -1
			}
			$startTime = Get-Date
			$exitCode = $null
			$hasShownProgress = $false
			$currentCompletion = 0
			$lastRemainingTime = [TimeSpan]::FromSeconds($TimeoutInSeconds)
			while ($job.State -eq 'Running') {
				$currentTime = Get-Date
				$timeElapsed = $currentTime - $startTime
				$remainingTime = $timeoutTimeSpan - $timeElapsed
				if ($remainingTime.TotalSeconds -lt 0) {
					$remainingTime = [TimeSpan]::Zero
				}
				$lastRemainingTime = $remainingTime
				$percentComplete = if ($TimeoutInSeconds -gt 0) {
					[Math]::Min(100, ($timeElapsed.TotalSeconds / $TimeoutInSeconds) * 100)
				}
				else {
					100
				}
				$currentCompletion = [Math]::Round(($percentComplete / 100) * $ProgressBarWidth)
				$hasShownProgress = $true
				if ($useNativeWriteProgress) {
					$barTextNative = '[' + ('█' * $currentCompletion) + ('░' * ($ProgressBarWidth - $currentCompletion)) + ']'
					$statusStringNative = "{0} {1:N1}% Elapsed: {2:hh\:mm\:ss} (Remaining: {3:hh\:mm\:ss})" -f $barTextNative, $percentComplete, $timeElapsed, $remainingTime
					Write-Progress -Id $progressIdForNative -Activity $progressActivityMessage -Status $statusStringNative -PercentComplete ([int]$percentComplete) -SecondsRemaining ([int]$remainingTime.TotalSeconds) -ErrorAction SilentlyContinue
				}
				elseif (Get-Command -Name Write-ProgressBar -ErrorAction SilentlyContinue) {
					Write-ProgressBar -Completed $currentCompletion -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $timeElapsed -RemainingTime $remainingTime -DontUseEscapeCodes:$DontUseEscapeCodes.IsPresent
				}
				if ($timeElapsed.TotalSeconds -ge $TimeoutInSeconds) {
					if ($useNativeWriteProgress) {
						# Keep current progress rather than marking as completed
						$finalBarTextNative = '[' + ('█' * $currentCompletion) + ('░' * ($ProgressBarWidth - $currentCompletion)) + ']'
						$finalStatusStringNative = "{0} {1:N1}% Elapsed: {2:hh\:mm\:ss} (Remaining: {3:hh\:mm\:ss})" -f $finalBarTextNative, $percentComplete, $timeElapsed, $lastRemainingTime
						Write-Progress -Id $progressIdForNative -Activity $progressActivityMessage -Status $finalStatusStringNative -PercentComplete ([int]$percentComplete) -Completed -ErrorAction SilentlyContinue
					}
					elseif ($hasShownProgress -and (Get-Command -Name Write-ProgressBar -ErrorAction SilentlyContinue)) {
						if (-not $useNativeWriteProgress) {
							# Final visual update with current completion level
							Write-ProgressBar -Completed $currentCompletion -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $timeElapsed -RemainingTime $lastRemainingTime -DontUseEscapeCodes:$DontUseEscapeCodes.IsPresent
							# If visual update bypassed transcript, explicitly log final bar state
							if ($customBarVisualsBypassTranscript) {
								$finalBarLogString = Write-ProgressBar -Completed $currentCompletion -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $timeElapsed -RemainingTime $lastRemainingTime -DontUseEscapeCodes:$DontUseEscapeCodes.IsPresent -PassThruStringForLog
								if ($finalBarLogString) {
									Write-Host $finalBarLogString
								}
							}
						}
						if ($customBarVisualsBypassTranscript) {
							[System.Console]::WriteLine()
						}
						else {
							Write-Host
						}
					}
					New-Log "$executableFileName process timed out." -Level WARNING
					Stop-Job -Job $job
					$exitCode = -1
					break
				}
				Start-Sleep -Seconds $SleepInSeconds
			}
			if ($null -eq $exitCode) {
				# Process did not timeout in the loop above
				try {
					$jobResult = Receive-Job -Job $job -Wait -EA Stop
					$exitCode = $jobResult
					$finalTimeElapsed = (Get-Date) - $startTime
					$finalPercentComplete = if ($TimeoutInSeconds -gt 0) {
						[Math]::Min(100, ($finalTimeElapsed.TotalSeconds / $TimeoutInSeconds) * 100)
					}
					else {
						100
					}
					# Calculate final completion based on current time, not forcing to 100%
					$finalCompletion = [Math]::Round(($finalPercentComplete / 100) * $ProgressBarWidth)
					# Use the last known remaining time
					$remainingTime = $timeoutTimeSpan - $finalTimeElapsed
					if ($remainingTime.TotalSeconds -lt 0) {
						$remainingTime = [TimeSpan]::Zero
					}
					$lastRemainingTime = $remainingTime
					if ($hasShownProgress) {
						if ($useNativeWriteProgress) {
							# Keep current progress rather than marking as completed
							$finalBar = '[' + ('█' * $finalCompletion) + ('░' * ($ProgressBarWidth - $finalCompletion)) + ']'
							$finalStatus = "{0} {1:N1}% Elapsed: {2:hh\:mm\:ss} (Remaining: {3:hh\:mm\:ss})" -f $finalBar, $finalPercentComplete, $finalTimeElapsed, $lastRemainingTime
							Write-Progress -Id $progressIdForNative -Activity $progressActivityMessage -Status $finalStatus -PercentComplete ([int]$finalPercentComplete) -Completed -EA SilentlyContinue
						}
						elseif (Get-Command -Name Write-ProgressBar -ErrorAction SilentlyContinue) {
							if (-not $useNativeWriteProgress) {
								# Final visual update with current completion
								Write-ProgressBar -Completed $finalCompletion -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $finalTimeElapsed -RemainingTime $lastRemainingTime -DontUseEscapeCodes:$DontUseEscapeCodes.IsPresent
								# If visual update bypassed transcript, explicitly log final bar state
								if ($customBarVisualsBypassTranscript) {
									$finalBarLogString = Write-ProgressBar -Completed $finalCompletion -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $finalTimeElapsed -RemainingTime $lastRemainingTime -DontUseEscapeCodes:$DontUseEscapeCodes.IsPresent -PassThruStringForLog
									if ($finalBarLogString) {
										Write-Host $finalBarLogString
									}
								}
							}
							if ($customBarVisualsBypassTranscript) {
								[System.Console]::WriteLine()
							}
							else {
								Write-Host
							}
						}
					}
					elseif ($useNativeWriteProgress) {
						Write-Progress -Id $progressIdForNative -Activity $progressActivityMessage -Completed -ErrorAction SilentlyContinue
					}
					if ($exitCode -eq 0) {
						New-Log "$executableFileName completed successfully (Code $exitCode)." -Level SUCCESS
					}
					else {
						New-Log "$executableFileName exited (Code $exitCode)." -Level WARNING
						if (-not $PSBoundParameters.ContainsKey('RedirectStandardError') -and (Test-Path $tempError) -and (Get-Item $tempError).Length -gt 0) {
							$errContent = Get-Content $tempError -Raw -ErrorAction SilentlyContinue
							if ($errContent) {
								New-Log "StdErr: $errContent" -Level WARNING
							}
						}
					}
				}
				catch {
					if ($useNativeWriteProgress) {
						Write-Progress -Id $progressIdForNative -Activity $progressActivityMessage -Completed -ErrorAction SilentlyContinue
					}
					elseif ($hasShownProgress -and (Get-Command -Name Write-ProgressBar -ErrorAction SilentlyContinue)) {
						# Console newline (even on error, to prevent log message on same line as last bar update)
						if ($customBarVisualsBypassTranscript) {
							[System.Console]::WriteLine()
						}
						else {
							Write-Host
						}
					}
					New-Log "Error retrieving job result for $executableFileName." -Level ERROR
					$exitCode = -1
				}
			}
			return $exitCode
		}
	}
	catch {
		New-Log "Unhandled exception in Start-WithProgress." -Level ERROR
		if ($useNativeWriteProgress) {
			Write-Progress -Id $progressIdForNative -Activity "Error Occurred in Start-WithProgress" -Completed -ErrorAction SilentlyContinue
		}
		return -1
	}
	finally {
		if ($tempResources.Job) {
			Remove-Job -Job $tempResources.Job -Force -ErrorAction SilentlyContinue
		}
		$tempResources.TempFiles | ForEach-Object {
			if (Test-Path $_ -PathType Leaf) {
				Remove-Item $_ -Force -ErrorAction SilentlyContinue
			}
		}
		if ($tempOutput -and -not $PSBoundParameters.ContainsKey('RedirectStandardOutput') -and (Test-Path $tempOutput -PathType Leaf)) {
			Remove-Item $tempOutput -Force -ErrorAction SilentlyContinue
		}
		if ($tempError -and -not $PSBoundParameters.ContainsKey('RedirectStandardError') -and (Test-Path $tempError -PathType Leaf)) {
			Remove-Item $tempError -Force -ErrorAction SilentlyContinue
		}
		if ($PSCmdlet.ParameterSetName -ne 'Unknown' -and $useNativeWriteProgress) {
			try {
				Write-Progress -Id $progressIdForNative -Activity "Finalizing..." -Completed -ErrorAction SilentlyContinue
			}
			catch { }
		}
	}
}
function Test-TranscriptActive {
	<#
.SYNOPSIS
    Checks if a PowerShell transcript is currently active.
.DESCRIPTION
    This function determines if a transcript is active by inspecting internal runspace data via reflection.
    It is designed to be non-intrusive and does not start or stop any transcripts.
    This method is generally more reliable than checking global variables or attempting to stop/start transcripts.
.OUTPUTS
    System.Boolean
    $true if one or more transcripts are active in the current runspace; $false otherwise or if an error occurs during detection.
.NOTES
    Relies on reflecting into $Host.Runspace.TranscriptionData.Transcripts.
    While robust, the internal structure of these objects could theoretically change in future PowerShell versions,
    though the "TranscriptionData" and "Transcripts" properties have been stable.
#>
	[CmdletBinding()]
	param()
	process {
		try {
			$flags = [System.Reflection.BindingFlags]'Instance, NonPublic, Public'
			$transcriptionData = $Host.Runspace.GetType().GetProperty('TranscriptionData', $flags).GetValue($Host.Runspace)
			if (-not $transcriptionData) {
				return $false
			}
			$transcripts = $transcriptionData.GetType().GetProperty('Transcripts', $flags).GetValue($transcriptionData)
			if (-not $transcripts) {
				return $false
			}
			if (($transcripts -is [System.Collections.ICollection] -and $transcripts.Count -gt 0) -or ($transcripts -isnot [System.Collections.ICollection] -and ($transcripts | ForEach-Object { $_ } | Measure-Object).Count -gt 0)) {
				return $true
			}
			else {
				return $false
			}
		}
		catch {
			# Silently return false on any reflection error, assuming no transcript.
			return $false
		}
	}
}
#####################################################################################################################################################
### OBS: New-Log Function is needed otherwise remove all New-Log and replace with Write-Host. New-Log is vastly better though, check the link below:#
#####################################################################################################################################################
#Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
# --- MORE EXAMPLES ---
# Ensure New-Log is defined in your session before running these examples.
<#
Start-WithProgress -FilePath 'C:\Windows\System32\cmd.exe' -Arguments '/c echo Hello From CMD && timeout /t 3 && echo CMD Finished' -TimeoutInSeconds 10
###
Start-WithProgress -FilePath 'C:\Windows\System32\ping.exe' -Arguments 'google.com -n 7' -TimeoutInSeconds 15 -SleepInSeconds 2
###
Start-WithProgress -FilePath 'C:\Windows\System32\ping.exe' -Arguments 'google.com -n 10' -TimeoutInSeconds 5
###
$testFile = Join-Path $env:TEMP "StartWithProgress_TestFile.txt"
if(Test-Path $testFile) { Remove-Item $testFile -Force }
Start-Job -ScriptBlock { param($file) Start-Sleep -Seconds 3; "Created!" | Out-File -FilePath $file } -ArgumentList $testFile | Out-Null
Start-WithProgress -Condition { Test-Path $testFile } -Message "Waiting for '$testFile'..." -SuccessMessage "'$testFile' has been created!" -TimeoutInSeconds 10
if(Test-Path $testFile) {
	Write-Host "Content of ${testFile}: $(Get-Content $testFile)"
	Remove-Item $testFile -Force
}
Get-Job | Where-Object {$_.Name -like 'Job*'} | Remove-Job -Force # Clean up example background job
###
Start-Transcript -Path "C:\Temp\Test_TranscriptActive.txt" -Force
Start-WithProgress -JustWait -TimeoutInSeconds 10 -SleepInSeconds 2 -Message "Test: PS5.1 Transcript Active"
Stop-Transcript
Get-Content "C:\Temp\Test_TranscriptActive.txt"
#>