function Start-WithProgress {
	<#
	.SYNOPSIS
		Executes a command or waits for a condition with a progress bar and timeout.
		Compatible with PowerShell 5.1 and PowerShell 7+.
	.DESCRIPTION
		This function provides two main modes of operation:
		1. FilePath Mode (Default): Executes an external program or script. It displays a progress bar
		representing the elapsed time versus the specified timeout. Standard output and standard error
		can be redirected.
		2. Condition Mode: Waits for a PowerShell scriptblock condition to become true. It displays
		a progress bar based on elapsed time versus timeout.
		The function uses background jobs to run processes or monitor conditions asynchronously,
		allowing the progress bar to update in the main thread. It handles differences between
		PowerShell 5.1 and newer versions, especially for process execution and I/O redirection,
		to ensure consistent behavior. For PowerShell 5.1, process execution involves creating a
		temporary batch file to robustly handle I/O and special characters, particularly for
		sensitive executables like `timeout.exe`.
		It features a customizable progress bar, including width and color schemes (ANSI for modern
		consoles, fallback for ISE/older systems).
		Logging is performed using a 'New-Log' function, which **must** be available in the
		calling scope. This function is intended to be part of a logging module.
	.PARAMETER FilePath
		The path to the executable file or script to be run. This is for the 'FilePath' parameter set.
		Aliases: Path
	.PARAMETER Arguments
		A string containing the arguments to pass to the executable specified by FilePath.
		Aliases: Parameters, ArgumentList
	.PARAMETER TimeoutInSeconds
		The maximum time in seconds to wait for the process to complete or the condition to become true.
		If the timeout is reached, the process/job is terminated, or the condition wait stops.
		Default is 60 seconds.
	.PARAMETER SleepInSeconds
		The interval in seconds at which the progress bar updates and checks are performed.
		Must be less than or equal to TimeoutInSeconds.
		Default is 1 second.
	.PARAMETER NoNewWindow
		If specified, the process started by Start-Process (internally) will not create a new window.
		Cannot be used with -WindowStyle if -WindowStyle is anything other than 'Hidden'.
	.PARAMETER PassThru
		If specified when using FilePath mode, this function will attempt to return the exit code
		of the executed process. This mimics the -PassThru behavior of Start-Process, but this
		function *always* returns an exit code or a status code (-1 for timeout/error).
	.PARAMETER RedirectStandardOutput
		Specifies a file path to redirect the standard output (stdout) of the process to.
		If not specified, stdout is redirected to a temporary file and discarded unless an error occurs.
		Encoding for PS5.1 redirection via batch file is OEM; for PS6+ it is UTF8.
	.PARAMETER JustWait
		Used with the 'Condition' parameter set. If specified, when the condition becomes true or
		the timeout is reached, it logs a "wait is over" message instead of success/failure messages
		related to the condition itself.
	.PARAMETER RedirectStandardError
		Specifies a file path to redirect the standard error (stderr) of the process to.
		If not specified, stderr is redirected to a temporary file. If the process exits with a
		non-zero code, the content of this temporary stderr file is displayed.
		Encoding for PS5.1 redirection via batch file is OEM; for PS6+ it is UTF8.
	.PARAMETER WindowStyle
		Specifies the window style for the process. Valid values are: Normal, Hidden, Minimized, Maximized.
		Default is 'Hidden'. If -NoNewWindow is used, WindowStyle is effectively Hidden.
	.PARAMETER ProgressBarWidth
		The width of the character-based progress bar. Default is 50.
	.PARAMETER ISEMode
		Forces the progress bar to use a mode compatible with the PowerShell ISE (which does not
		support ANSI escape codes for colors). If not specified, the function attempts to auto-detect ISE.
	.PARAMETER Condition
		A scriptblock that evaluates to $true or $false. The function will wait until this condition
		is $true or the timeout is reached. This is for the 'Condition' parameter set.
	.PARAMETER Message
		A message to display (via New-Log) when starting the wait in 'Condition' mode.
	.PARAMETER SuccessMessage
		A message to display (via New-Log) when the condition in 'Condition' mode successfully evaluates to $true.
	.PARAMETER DontUseEscapeCodes
		If specified, forces the progress bar to use the simple block character display, even in
		consoles that might support ANSI escape codes. Useful for compatibility or preference.
	.OUTPUTS
		System.Int32
		- In FilePath mode: Returns the exit code of the executed process.
		- Returns -1 if the process timed out or an internal error occurred in Start-WithProgress.
		- Returns -11 if there was a fatal error generating the PS5.1 batch file wrapper.
		- In Condition mode:
		- Returns 0 if the condition became true.
		- Returns -1 if the timeout was reached before the condition became true.
	.EXAMPLE
		# Assuming New-Log function is defined and available.
		# Example 1: Run cmd.exe to echo and then timeout, wait up to 10 seconds.
		Start-WithProgress -FilePath 'C:\Windows\System32\cmd.exe' -Arguments '/c echo Hello World && timeout /t 3' -TimeoutInSeconds 10
	.EXAMPLE
		# Example 2: Ping google.com 5 times, but kill the process if it takes longer than 7 seconds.
		Start-WithProgress -FilePath 'C:\Windows\System32\ping.exe' -Arguments 'google.com -n 5' -TimeoutInSeconds 7
	.EXAMPLE
		# Example 3: Wait for a specific file to exist, checking every 2 seconds, for up to 30 seconds.
		$myFile = "C:\temp\waitForMe.txt"
		# Remove-Item $myFile -ErrorAction SilentlyContinue # Ensure it doesn't exist initially
		# Start-Job { Start-Sleep 5; "content" | Out-File $using:myFile } # Create it after 5s
		Start-WithProgress -Condition { Test-Path $myFile } -Message "Waiting for $myFile..." -SuccessMessage "$myFile created!" -TimeoutInSeconds 30 -SleepInSeconds 2
	.NOTES
		AUTHOR: Harze2k
		Date:   2025-05-10
		VERSION: 2.8 (Fixed so that -FilePath is not required when using -JustWait)

		-Requires a `New-Log` function to be defined in the scope where Start-WithProgress is called.

		The definition of this `New-Log` function will be passed into the background job.
		If `New-Log` is not found in the calling scope when Start-WithProgress is invoked, this function
		will likely error when trying to capture its definition.
		For PowerShell 5.1, process execution involves creating a temporary batch file wrapper. This
		means stdout/stderr from processes in PS5.1 will typically be in the system's OEM encoding.
		In PowerShell 7+, redirection uses UTF-8.
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
		[Parameter(Position = 1, ParameterSetName = 'Condition')][string]$Message,
		[Parameter(Position = 2, ParameterSetName = 'Condition')][string]$SuccessMessage,
		[switch]$DontUseEscapeCodes
	)
	#region Helper-functions and initialization
	#region Write-ProgressBar
	function Write-ProgressBar {
		param(
			[int]$Completed,
			[bool]$IsPsISE,
			[int]$Width = 40,
			[TimeSpan]$TimeElapsed,
			[TimeSpan]$RemainingTime,
			[switch]$DontUseEscapeCodes
		)
		#region Get-ConsoleColor
		function Get-ConsoleColor {
			param([double]$Percentage)
			switch ($Percentage) {
				{ $_ -lt 16.7 } { return [System.ConsoleColor]::Green }
				{ $_ -lt 33.3 } { return [System.ConsoleColor]::DarkGreen }
				{ $_ -lt 50 } { return [System.ConsoleColor]::Yellow }
				{ $_ -lt 66.7 } { return [System.ConsoleColor]::DarkYellow }
				{ $_ -lt 83.3 } { return [System.ConsoleColor]::DarkRed }
				default { return [System.ConsoleColor]::Red }
			}
		}
		if (!(Get-Command New-Log -ErrorAction SilentlyContinue)) {
			Write-Host "Please make sure the New-Log function is loaded. Will abort. Or remove New-Log with Write-Host." -ForegroundColor Red
			return
		}
		#endregion Get-ConsoleColor
		$completedWidth = [Math]::Max(0, [Math]::Min($Completed, $Width))
		$remainingWidth = $Width - $completedWidth
		$percentage = [Math]::Round(($completedWidth / $Width) * 100, 1)
		Write-Host -NoNewline "`r"
		Write-Host -NoNewline "["
		if ($IsPsISE -or $DontUseEscapeCodes) {
			for ($i = 0; $i -lt $completedWidth; $i++) {
				$color = Get-ConsoleColor(($i / $Width) * 100)
				Write-Host -NoNewline -ForegroundColor $color '#'
			}
			for ($i = 0; $i -lt $remainingWidth; $i++) {
				Write-Host -NoNewline ' '
			}
		}
		else {
			$esc = [char]27
			for ($i = 0; $i -lt $Width; $i++) {
				$colorPercentage = ($i / $Width)
				$greenVal = [Math]::Max(0, [Math]::Min(255, [int](255 * (1 - $colorPercentage))))
				$redVal = [Math]::Max(0, [Math]::Min(255, [int](255 * $colorPercentage)))
				if ($i -lt $completedWidth) {
					Write-Host -NoNewline "${esc}[38;2;${redVal};${greenVal};0m█${esc}[0m"
				}
				else {
					Write-Host -NoNewline "${esc}[38;2;100;100;100m░${esc}[0m"
				}
			}
		}
		Write-Host -NoNewline "] "
		Write-Host -NoNewline ("{0:0.0}%" -f $percentage)
		Write-Host -NoNewline (" Elapsed: {0:hh\:mm\:ss} (Remaining: {1:hh\:mm\:ss})" -f $TimeElapsed, $RemainingTime)
		Write-Host -NoNewline "     "
	}
	if ($JustWait) {
		$Condition = { $false }
	}
	#endregion Write-ProgressBar
	# Create a hashtable to track temporary resources for cleanup
	$tempResources = @{
		TempFiles = @()
		Job       = $null
	}
	# Set ISE detection if not explicitly specified
	if (-not $PSBoundParameters.ContainsKey('ISEMode')) {
		$ISEMode = $Host.Name -eq 'Windows PowerShell ISE Host'
	}
	$timeoutTimeSpan = [TimeSpan]::FromSeconds($TimeoutInSeconds)
	#endregion Helper-functions and initialization
	try {
		# --- Parameter Validation ---
		# Verify SleepInSeconds vs TimeoutInSeconds
		if ($SleepInSeconds -gt $TimeoutInSeconds) {
			New-Log 'SleepInSeconds must be less than or equal to TimeoutInSeconds. Aborting.' -Level WARNING
			return -1
		}
		# Early validation of New-Log function
		if (-not (Get-Command -Name New-Log -ErrorAction SilentlyContinue)) {
			Write-Error "Start-WithProgress: The 'New-Log' function was not found in the current scope. It is required for this function to operate."
			return -1
		}
		# --- Condition-Based Progress ---
		if ($PSCmdlet.ParameterSetName -eq 'Condition') {
			if ($null -eq $Condition) {
				New-Log "Error: Condition parameter is null or empty" -Level ERROR
				return -1
			}
			New-Log "$Message"
			$startTime = Get-Date
			$conditionErrorCount = 0  # Track condition evaluation errors
			while ($true) {
				$currentTime = Get-Date
				$timeElapsed = $currentTime - $startTime
				$remainingTime = $timeoutTimeSpan - $timeElapsed
				if ($remainingTime.TotalSeconds -lt 0) { $remainingTime = [TimeSpan]::Zero }
				$completedPercentage = [Math]::Min(100, ($timeElapsed.TotalSeconds / $TimeoutInSeconds) * 100)
				$completedBlocks = [Math]::Round(($completedPercentage / 100) * $ProgressBarWidth)
				Write-ProgressBar -Completed $completedBlocks -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $timeElapsed -RemainingTime $remainingTime -DontUseEscapeCodes:$DontUseEscapeCodes
				# Check for timeout
				if ($timeElapsed.TotalSeconds -ge $TimeoutInSeconds) {
					Write-Host
					if ($JustWait) {
						New-Log 'The wait is over!' -Level SUCCESS
					}
					else {
						New-Log "Condition wait timed out after $TimeoutInSeconds seconds: $Message" -Level WARNING
					}
					return -1
				}
				# Try to evaluate the condition with error handling
				try {
					$conditionResult = & $Condition
					if ($conditionResult -eq $true) {
						Write-Host
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
						# Only log first few errors to avoid spamming
						New-Log "Error evaluating condition: $($_.Exception.Message)" -Level ERROR
					}
				}
				Start-Sleep -Seconds $SleepInSeconds
			}
		}
		# --- FilePath-Based Progress (Process Execution) ---
		# Validate FilePath parameter
		if ($PSCmdlet.ParameterSetName -eq 'FilePath') {
			if ([string]::IsNullOrWhiteSpace($FilePath)) {
				New-Log "FilePath parameter is required for process execution. Aborting." -Level WARNING
				return -1
			}
			$OriginalFilePath = $FilePath # Keep original for logging
			# Advanced FilePath validation and resolution
			$pathExists = Test-Path $FilePath -PathType Leaf -ErrorAction SilentlyContinue
			if (-not $pathExists) {
				# If not a direct file, try to resolve as a command
				$resolvedCommand = Get-Command $FilePath -ErrorAction SilentlyContinue
				if ($null -eq $resolvedCommand) {
					# Try to find in the Path
					$pathMatch = $env:PATH.Split(';') |
						ForEach-Object { Join-Path $_ $FilePath } |
						Where-Object { Test-Path $_ -PathType Leaf } |
						Select-Object -First 1
					if ($pathMatch) {
						$FilePath = $pathMatch
					}
					else {
						New-Log "FilePath '$OriginalFilePath' is not a valid file path and could not be resolved as a command. Aborting." -Level WARNING
						return -1
					}
				}
				else {
					$FilePath = $resolvedCommand.Source
				}
			}
			# Validate WindowStyle and NoNewWindow compatibility
			if ($NoNewWindow -and $PSBoundParameters.ContainsKey('WindowStyle') -and $WindowStyle -ne 'Hidden') {
				New-Log "Parameters 'NoNewWindow' and explicit 'WindowStyle' (other than default 'Hidden') cannot be specified at the same time. Aborting." -Level WARNING
				return -1
			}
			$executableFileName = Split-Path -Path $FilePath -Leaf # Use resolved FilePath
			# Prepare temporary files for redirection
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
			# Initialize output and error files
			if (-not $PSBoundParameters.ContainsKey('RedirectStandardOutput')) {
				Set-Content -Path $tempOutput -Value $null -Encoding UTF8 -ErrorAction SilentlyContinue
			}
			if (-not $PSBoundParameters.ContainsKey('RedirectStandardError')) {
				Set-Content -Path $tempError -Value $null -Encoding UTF8 -ErrorAction SilentlyContinue
			}
			$isPowerShellV5OrLower = ($PSVersionTable.PSVersion.Major -le 5)
			# Get New-Log function definition for the background job
			$newLogDefForJob = ''
			try {
				$newLogDefForJob = ${function:New-Log}.ToString()
			}
			catch {
				New-Log "Error getting New-Log function definition: $($_.Exception.Message)" -Level ERROR
				return -1 # Critical dependency missing
			}
			# --- SCRIPTBLOCK FOR BACKGROUND JOB ---
			$scriptBlock = {
				param($ProcFilePath, $ProcArguments, $ProcOutputFile, $ProcErrorFile, $ProcNoNewWindow, $ProcWindowStyle, $IsPs5, $PassedNewLogDefinition)
				# Define New-Log within the job's script scope using the passed definition
				if (-not [string]::IsNullOrWhiteSpace($PassedNewLogDefinition)) {
					Invoke-Expression -Command "function New-Log { $PassedNewLogDefinition }"
				}
				else {
					Write-Error "Start-WithProgress: The 'New-Log' function was not found in the current scope. It is required for this function to operate. Please ensure it is defined or imported."
					return -1
				}
				# --- PowerShell 5.1 Process Execution Path ---
				if ($IsPs5) {
					$tempBatFile = $null
					try {
						$tempBatFile = [System.IO.Path]::GetTempFileName() + ".bat"
						$Line1 = "@echo off"
						$Line2 = ""
						$Line3 = "exit /b %errorlevel%"
						$ExecutableNameOnly = (Split-Path -Path $ProcFilePath -Leaf).ToLowerInvariant()
						$isTimeoutScenario = $false
						if ($ExecutableNameOnly -eq "cmd.exe" -and $ProcArguments -match 'timeout(\.exe)?\s') {
							$isTimeoutScenario = $true
						}
						if ($ExecutableNameOnly -eq "cmd.exe") {
							if ($isTimeoutScenario) {
								$Line2 = """$ProcFilePath"" $ProcArguments 1>""$ProcOutputFile"" 2>""$ProcErrorFile"""
							}
							else {
								$Line2 = """$ProcFilePath"" $ProcArguments <NUL 1>""$ProcOutputFile"" 2>""$ProcErrorFile"""
							}
						}
						else {
							$InnerTargetCommandForCmdC = """""$ProcFilePath"" $ProcArguments"""
							$Line2 = "cmd.exe /S /C $InnerTargetCommandForCmdC <NUL 1>""$ProcOutputFile"" 2>""$ProcErrorFile"""
						}
						$BatContent = "$Line1`r`n$Line2`r`n$Line3`r`n"
						if ([string]::IsNullOrWhiteSpace($Line2) -or [string]::IsNullOrWhiteSpace($BatContent) -or $BatContent.Length -lt 20) {
							$fatalMsg = "Job FATAL: Batch file content generation failed. Line2: '$Line2'"
							New-Log $fatalMsg -Level WARNING
							$fatalMsg | Out-File -FilePath $ProcErrorFile -Encoding OEM -Append
							return -11
						}
						Set-Content -Path $tempBatFile -Value $BatContent -Encoding ASCII -Force
						$startInfoArgs = @{
							FilePath    = $tempBatFile
							Wait        = $true
							PassThru    = $true
							ErrorAction = 'Stop'
						}
						if ($ProcNoNewWindow) {
							$startInfoArgs.NoNewWindow = $true
						}
						elseif ($ProcWindowStyle) {
							$startInfoArgs.WindowStyle = $ProcWindowStyle
						}
						$process = Start-Process @startInfoArgs
						return $process.ExitCode
					}
					catch {
						$errorMessage = "Job ERROR (PS5 Path): $($_.Exception.Message). TempBatFile: $tempBatFile"
						if ($tempBatFile -and (Test-Path $tempBatFile)) {
							$batContentOnError = Get-Content $tempBatFile -Raw -ErrorAction SilentlyContinue
							$errorMessage += "`r`nBatch File Content on Error (`"$tempBatFile`"):`r`n$batContentOnError"
						}
						New-Log $errorMessage -Level ERROR
						$errorMessage | Out-File -FilePath $ProcErrorFile -Encoding OEM -Append
						return -1
					}
					finally {
						if ($tempBatFile -and (Test-Path $tempBatFile)) {
							Remove-Item -Path $tempBatFile -Force -ErrorAction SilentlyContinue
						}
					}
				}
				# --- PowerShell 7+ Process Execution Path ---
				else {
					try {
						$startInfoArgs = @{
							FilePath               = $ProcFilePath
							ArgumentList           = $ProcArguments
							Wait                   = $true
							PassThru               = $true
							RedirectStandardOutput = $ProcOutputFile
							RedirectStandardError  = $ProcErrorFile
							ErrorAction            = 'Stop'
						}
						if ($ProcNoNewWindow) {
							$startInfoArgs.NoNewWindow = $true
						}
						elseif ($ProcWindowStyle) {
							$startInfoArgs.WindowStyle = $ProcWindowStyle
						}
						$process = Start-Process @startInfoArgs
						return $process.ExitCode
					}
					catch {
						$errorMessage = "Job ERROR (PS7+ Path): Error starting process '$ProcFilePath'. $($_.Exception.Message)"
						New-Log "Job ERROR (PS7+ Path): Error starting process." -Level ERROR
						$errorMessage | Out-File -FilePath $ProcErrorFile -Encoding UTF8 -Append
						return -1
					}
				}
			} # --- END SCRIPTBLOCK FOR BACKGROUND JOB ---
			$argsAsStringForLog = "`"$FilePath`"" # Use resolved FilePath for log
			if ($Arguments) {
				$argsAsStringForLog += " Arguments=`"$Arguments`""
			}
			$argsAsStringForLog += " TimeoutInSeconds=`"$TimeoutInSeconds`" SleepInSeconds=`"$SleepInSeconds`""
			New-Log "Executing: $argsAsStringForLog" -Level INFO
			# Start the job
			try {
				$job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $FilePath, $Arguments, $tempOutput, $tempError, $NoNewWindow, $WindowStyle, $isPowerShellV5OrLower, $newLogDefForJob
				$tempResources.Job = $job
			}
			catch {
				New-Log "Failed to start job: $($_.Exception.Message)" -Level ERROR
				return -1
			}
			$startTime = Get-Date
			$exitCode = $null
			$hasShownProgress = $false
			# --- Job Monitoring Loop ---
			while ($job.State -eq 'Running') {
				$currentTime = Get-Date
				$timeElapsed = $currentTime - $startTime
				$remainingTime = $timeoutTimeSpan - $timeElapsed
				if ($remainingTime.TotalSeconds -lt 0) {
					$remainingTime = [TimeSpan]::Zero
				}
				$completedPercentage = [Math]::Min(100, ($timeElapsed.TotalSeconds / $TimeoutInSeconds) * 100)
				$completedBlocks = [Math]::Round(($completedPercentage / 100) * $ProgressBarWidth)
				Write-ProgressBar -Completed $completedBlocks -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $timeElapsed -RemainingTime $remainingTime -DontUseEscapeCodes:$DontUseEscapeCodes
				$hasShownProgress = $true
				if ($timeElapsed.TotalSeconds -ge $TimeoutInSeconds) {
					Write-Host
					New-Log "$executableFileName process was terminated due to timeout after $TimeoutInSeconds seconds." -Level WARNING
					Stop-Job -Job $job
					$exitCode = -1
					break
				}
				Start-Sleep -Seconds $SleepInSeconds
			} # --- End Job Monitoring Loop ---
			# --- Process Job Result ---
			if ($null -eq $exitCode) {
				try {
					$jobResult = Receive-Job -Job $job -Wait -ErrorAction Stop
					$exitCode = $jobResult
					$finalTimeElapsed = (Get-Date) - $startTime
					$finalCompletedPercentage = if ($job.State -eq [System.Management.Automation.JobState]::Completed -or $job.State -eq [System.Management.Automation.JobState]::Failed) {
						if ($finalTimeElapsed.TotalSeconds -lt $TimeoutInSeconds) {
							100
						}
						else {
							[Math]::Min(100, ($finalTimeElapsed.TotalSeconds / $TimeoutInSeconds) * 100)
						}
					}
					else {
						[Math]::Min(100, ($finalTimeElapsed.TotalSeconds / $TimeoutInSeconds) * 100)
					}
					$finalCompletedBlocks = [Math]::Min($ProgressBarWidth, [Math]::Round(($finalCompletedPercentage / 100) * $ProgressBarWidth))
					if ($hasShownProgress -or $finalTimeElapsed.TotalSeconds -lt $SleepInSeconds) {
						Write-ProgressBar -Completed $finalCompletedBlocks -IsPsISE $ISEMode -Width $ProgressBarWidth -TimeElapsed $finalTimeElapsed -RemainingTime ([TimeSpan]::Zero) -DontUseEscapeCodes:$DontUseEscapeCodes
					}
					Write-Host
					if ($exitCode -eq 0) {
						New-Log "$executableFileName completed successfully. Exit code $($exitCode)." -Level SUCCESS
					}
					else {
						New-Log "$executableFileName exited with code $($exitCode)." -Level WARNING
						if (-not $PSBoundParameters.ContainsKey('RedirectStandardError') -and (Test-Path $tempError) -and (Get-Item $tempError).Length -gt 0) {
							$errorContent = Get-Content $tempError -Raw -ErrorAction SilentlyContinue
							if ($errorContent) {
								New-Log "Standard Error from ${executableFileName}. ErrorContent: $errorContent" -Level WARNING
							}
						}
					}
				}
				catch {
					New-Log "Error retrieving job result: $($_.Exception.Message)" -Level ERROR
					$exitCode = -1
					if ($hasShownProgress) {
						Write-Host
					}
				}
			}
			return $exitCode
		}
	}
	catch {
		$errorMessage = "Unhandled exception in Start-WithProgress: $($_.Exception.Message)"
		New-Log $errorMessage -Level ERROR
		New-Log "Stack Trace: $($_.Exception.StackTrace)" -Level ERROR
		return -1
	}
	finally {
		# Ensure cleanup of temporary resources
		if ($tempResources.Job) {
			Remove-Job -Job $tempResources.Job -Force -ErrorAction SilentlyContinue
		}
		# Clean up temp files
		foreach ($tempFile in $tempResources.TempFiles) {
			if (Test-Path $tempFile -PathType Leaf) {
				Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
			}
		}
		# Clean up any additional temp files created outside the tracking mechanism
		if ($tempOutput -and -not $PSBoundParameters.ContainsKey('RedirectStandardOutput') -and (Test-Path $tempOutput -PathType Leaf)) {
			Remove-Item $tempOutput -Force -ErrorAction SilentlyContinue
		}
		if ($tempError -and -not $PSBoundParameters.ContainsKey('RedirectStandardError') -and (Test-Path $tempError -PathType Leaf)) {
			Remove-Item $tempError -Force -ErrorAction SilentlyContinue
		}
	}
}
#####################################################################################################################################################
### OBS: New-Log Function is needed otherwise remove all New-Log and replace with Write-Host. New-Log is vastly better though, check the link below:#
#####################################################################################################################################################
#Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
<#
# --- MORE EXAMPLES ---
# Ensure New-Log is defined in your session before running these examples.
Start-WithProgress -FilePath 'C:\Windows\System32\cmd.exe' -Arguments '/c echo Hello From CMD && timeout /t 3 && echo CMD Finished' -TimeoutInSeconds 10
Start-WithProgress -FilePath 'C:\Windows\System32\ping.exe' -Arguments 'google.com -n 7' -TimeoutInSeconds 15 -SleepInSeconds 2
Start-WithProgress -FilePath 'C:\Windows\System32\ping.exe' -Arguments 'google.com -n 10' -TimeoutInSeconds 5
$testFile = Join-Path $env:TEMP "StartWithProgress_TestFile.txt"
if(Test-Path $testFile) { Remove-Item $testFile -Force }
Start-Job -ScriptBlock { param($file) Start-Sleep -Seconds 3; "Created!" | Out-File -FilePath $file } -ArgumentList $testFile | Out-Null
Start-WithProgress -Condition { Test-Path $testFile } -Message "Waiting for '$testFile'..." -SuccessMessage "'$testFile' has been created!" -TimeoutInSeconds 10
if(Test-Path $testFile) {
	Write-Host "Content of ${testFile}: $(Get-Content $testFile)"
	Remove-Item $testFile -Force
}
Get-Job | Where-Object {$_.Name -like 'Job*'} | Remove-Job -Force # Clean up example background job
#>