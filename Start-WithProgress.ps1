function Start-WithProgress {
    <#
    .SYNOPSIS
        Executes external commands or waits for conditions with a visual progress bar and timeout.
    .DESCRIPTION
        Start-WithProgress provides three operation modes with visual progress feedback:
        FILEPATH MODE (Default)
        Executes an external program or script while displaying a progress bar. The process
        runs in the background and is monitored for completion or timeout. Supports output
        redirection for capturing stdout and stderr.
        CONDITION MODE
        Waits for a PowerShell scriptblock condition to evaluate to $true. The condition is
        checked repeatedly at the specified sleep interval until it becomes true or times out.
        Useful for waiting on file creation, service states, or other testable conditions.
        JUSTWAIT MODE
        Simply displays a progress bar for the specified duration. Useful for adding visual
        feedback during deliberate pauses in scripts.
        KEY FEATURES
        - Transcript-Aware: Progress updates bypass transcript recording, keeping logs clean while still displaying the final state.
        - Color Gradient: Progress bar shows green→yellow→red gradient for intuitive feedback.
        - Cross-Platform: Compatible with PowerShell 5.1 and 7+ on Windows 10/11.
        - Timeout Protection: All operations have configurable timeouts to prevent hangs.
        - Clean Cleanup: Automatic cleanup of background jobs and temporary files.
        REQUIREMENTS
        - The New-Log function must be defined in the calling scope for logging output.
        - Windows operating system (uses Windows-specific process handling).
    .PARAMETER FilePath
        Path to the executable file, script, or command to run. Can be:
        - Full path: "C:\Windows\System32\ping.exe"
        - Relative path: ".\script.bat"
        - Command in PATH: "ping" or "python"
        The function automatically resolves commands found in the system PATH.
        [Parameter Set: FilePath]
    .PARAMETER Arguments
        Arguments to pass to the executable as a single string.
        Special characters (|, &, <, >) are handled correctly for cmd.exe.
        Aliases: Parameters, ArgumentList
    .PARAMETER Condition
        A scriptblock that evaluates to $true or $false. The function waits until
        this condition returns $true or the timeout is reached.
        The scriptblock is evaluated fresh on each check interval.
        [Parameter Set: Condition]
    .PARAMETER JustWait
        Switch to enable simple wait mode. Displays a progress bar for the duration
        specified by TimeoutInSeconds without executing any command or condition.
        [Parameter Set: JustWait]
    .PARAMETER Message
        Message to log when starting the operation. Displayed via New-Log at INFO level.
        Default: "Awaiting condition..." (for Condition and JustWait modes)
    .PARAMETER SuccessMessage
        Custom message to log when a condition is successfully met.
        If not specified, logs "Condition met: <Message>" by default.
        [Parameter Set: Condition only]
    .PARAMETER TimeoutInSeconds
        Maximum time to wait for completion in seconds.
        - FilePath mode: Process is terminated if still running after timeout
        - Condition mode: Returns -1 if condition not met within timeout
        - JustWait mode: Total duration to wait
        Default: 60 seconds. Minimum: 1 second.
    .PARAMETER SleepInSeconds
        Interval between progress updates and condition checks in seconds.
        Smaller values give smoother progress but use more CPU.
        Must not exceed TimeoutInSeconds.
        Default: 1 second. Minimum: 1 second.
    .PARAMETER ProgressBarWidth
        Width of the progress bar in characters (filled + unfilled portion).
        Does not include brackets, percentage, or time displays.
        Default: 50 characters.
    .PARAMETER WindowStyle
        Window style for the launched process (FilePath mode only).
        Valid values: Normal, Hidden, Minimized, Maximized
        Default: Hidden (process window not visible)
    .PARAMETER NoNewWindow
        When specified, runs the process in the current console window.
        Cannot be combined with WindowStyle other than Hidden.
        [FilePath mode only]
    .PARAMETER PassThru
        When specified, returns the result code to the pipeline.
        - FilePath mode: Returns the process exit code (0 = success, non-zero = error, -1 = timeout/error)
        - Condition mode: Returns 0 if condition met, -1 if timed out
        Without this switch, no output is written to the pipeline.
    .PARAMETER RedirectStandardOutput
        File path to capture the process's standard output (stdout).
        If not specified, stdout is captured to a temp file and discarded.
        [FilePath mode only]
    .PARAMETER RedirectStandardError
        File path to capture the process's standard error (stderr).
        If not specified and process exits with non-zero code, stderr is logged as warning.
        [FilePath mode only]
    .PARAMETER DontUseEscapeCodes
        When specified, uses basic ASCII characters instead of ANSI color codes.
        Use this for terminals that don't support ANSI escape sequences.
    .PARAMETER ISEMode
        When specified, forces ISE-compatible display mode using ConsoleColor.
        Auto-detected when running in PowerShell ISE.
    .OUTPUTS
        System.Int32
            FilePath mode:
                Returns the process exit code on success (usually 0).
                Returns -1 on timeout, if file not found, or on execution error.
            Condition mode:
                Returns 0 when condition becomes true.
                Returns -1 on timeout.
            JustWait mode:
                Returns $null (no output to pipeline).
    .EXAMPLE
        Start-WithProgress -FilePath 'ping.exe' -Arguments 'google.com -n 4' -TimeoutInSeconds 10
        Description: Ping google.com 4 times with a 10-second timeout.
        Shows progress bar while waiting for ping to complete.
    .EXAMPLE
        Start-WithProgress -FilePath 'cmd.exe' -Arguments '/c echo Hello World' -TimeoutInSeconds 5
        Description: Run a simple cmd.exe command with progress display.
    .EXAMPLE
        $exitCode = Start-WithProgress -FilePath 'robocopy.exe' -Arguments 'C:\Source D:\Backup /MIR' -TimeoutInSeconds 300
        if ($exitCode -lt 8) { Write-Host "Backup completed successfully" }
        Description: Run robocopy with 5-minute timeout and check exit code.
        Robocopy exit codes 0-7 indicate success with varying conditions.
    .EXAMPLE
        Start-WithProgress -FilePath 'cmd.exe' -Arguments '/c dir C:\Windows\System32 /s' `
            -TimeoutInSeconds 60 -RedirectStandardOutput 'C:\Temp\dirlist.txt'
        Description: Capture directory listing to a file while showing progress.
    .EXAMPLE
        Start-WithProgress -FilePath 'python.exe' -Arguments 'script.py --verbose' `
            -TimeoutInSeconds 120 -RedirectStandardError 'C:\Logs\errors.txt'
        Description: Run Python script and capture any errors to a log file.
    .EXAMPLE
        Start-WithProgress -Condition { Test-Path 'C:\Temp\ready.txt' } `
            -Message 'Waiting for ready signal...' -TimeoutInSeconds 60 -ISEMode
        Description: Wait up to 60 seconds for a file to be created.
        Returns 0 when file appears, -1 on timeout.
    .EXAMPLE
        Start-WithProgress -Condition { (Get-Service 'Spooler').Status -eq 'Running' } `
            -Message 'Waiting for Print Spooler...' -SuccessMessage 'Spooler is running!' `
            -TimeoutInSeconds 30 -SleepInSeconds 2
        Description: Wait for a Windows service to start, checking every 2 seconds.
    .EXAMPLE
        $script:counter = 0
        Start-WithProgress -Condition { $script:counter++; $script:counter -ge 5 } `
            -Message 'Counting iterations...' -TimeoutInSeconds 10 -SleepInSeconds 1
        Description: Wait for a counter to reach a threshold.
        Demonstrates using script-scoped variables in conditions.
    .EXAMPLE
        Start-WithProgress -Condition { 
            $response = Invoke-WebRequest -Uri 'http://localhost:8080/health' -UseBasicParsing -ErrorAction SilentlyContinue
            $response.StatusCode -eq 200
        } -Message 'Waiting for web server...' -TimeoutInSeconds 120 -SleepInSeconds 5
        Description: Wait for a web service to become healthy, checking every 5 seconds.
    .EXAMPLE
        Start-WithProgress -Condition {
            -not (Get-Process -Name 'notepad' -ErrorAction SilentlyContinue)
        } -Message 'Waiting for Notepad to close...' -TimeoutInSeconds 300
        Description: Wait for a specific process to exit.
    .EXAMPLE
        Start-WithProgress -JustWait -TimeoutInSeconds 5 -Message 'Pausing before next step...'
        Description: Simple 5-second pause with visual progress feedback.
    .EXAMPLE
        Write-Host "Deploying application..."
        Start-WithProgress -JustWait -TimeoutInSeconds 10 -Message 'Waiting for services to stabilize...'
        Write-Host "Running health checks..."
        Description: Add visual delay between deployment steps.
    .EXAMPLE
        Start-WithProgress -JustWait -TimeoutInSeconds 30 -SleepInSeconds 1 -ProgressBarWidth 80
        Description: Wide progress bar with 1-second update intervals for smooth animation.
    .EXAMPLE
        # Running with transcript - progress bar won't spam the log
        Start-Transcript -Path 'C:\Logs\deploy.txt'
        Start-WithProgress -FilePath 'msiexec.exe' -Arguments '/i setup.msi /qn' -TimeoutInSeconds 300
        Stop-Transcript
        Description: Progress displays on screen but only final state is logged to transcript.
    .EXAMPLE
        # Chaining multiple operations
        $result1 = Start-WithProgress -FilePath 'cmd.exe' -Arguments '/c setup1.bat' -TimeoutInSeconds 60
        if ($result1 -eq 0) {
            $result2 = Start-WithProgress -Condition { Test-Path 'C:\App\installed.flag' } `
                -Message 'Verifying installation...' -TimeoutInSeconds 30
        }
        Description: Run installer then wait for installation verification file.
    .NOTES
        Author:         Harze2k
        Version:        3.2
        Date:           2025-11-29
        Compatibility:  PowerShell 5.1, PowerShell 7.x
        Platform:       Windows 10, Windows 11, Windows Server 2016+
        CHANGES (v3.2)
        - Fixed ISEMode color gradient to use smooth green→yellow→red transition
        - Fixed potential output leaks causing array returns instead of single values
        - Improved temp file cleanup reliability for batch wrapper files
        - All internal operations now properly suppress output
        DEPENDENCIES
        Please use my New-Log function for loggin to the console, otherwise this wont work. Check it out on my GitHub: 
        https://github.com/Harze2k/Shared-PowerShell-Functions/blob/main/New-Log.ps1
        One-liner to use it:
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
        TRANSCRIPT BEHAVIOR
        When a PowerShell transcript is active, progress bar updates use [System.Console]::Write()
        which bypasses the transcript. Only the final progress state and log messages are recorded,
        keeping transcript files clean and readable.
        EXIT CODES
        For FilePath mode, the function returns the actual process exit code. Common conventions:
        - 0: Success
        - 1: General error
        - -1: Timeout, file not found, or internal error (set by this function)
        ERROR HANDLING
        - Invalid file paths log a WARNING and return -1
        - Condition evaluation errors are logged (max 3 times) and continue
        - Timeout always logs a WARNING with the duration exceeded
        SPECIAL CHARACTERS IN ARGUMENTS
        For cmd.exe arguments containing |, &, <, >, the function automatically wraps
        commands to preserve shell interpretation. Example:
            -Arguments '/c echo test | findstr test'  # Works correctly
    .LINK
        Write-ProgressBar
    .LINK
        Test-TranscriptActive
    #>
    [CmdletBinding(DefaultParameterSetName = 'FilePath')]
    [OutputType([int])]
    param(
        [Parameter(Position = 0, ParameterSetName = 'FilePath')][Alias('Path')][string]$FilePath,
        [Parameter(Position = 1)][Alias('Parameters', 'ArgumentList')][string]$Arguments = '',
        [Parameter(Position = 0, ParameterSetName = 'Condition', Mandatory)][scriptblock]$Condition,
        [Parameter(Position = 0, ParameterSetName = 'JustWait', Mandatory)][switch]$JustWait,
        [Parameter(Position = 1, ParameterSetName = 'Condition')]
        [Parameter(Position = 1, ParameterSetName = 'JustWait')]
        [string]$Message,
        [Parameter(Position = 2, ParameterSetName = 'Condition')][string]$SuccessMessage,
        [Parameter(Position = 2)][ValidateRange([float]0.05, [int]::MaxValue)][float]$TimeoutInSeconds = 60,
        [Parameter(Position = 3)][ValidateRange([float]0.05, [int]::MaxValue)][float]$SleepInSeconds = 1,
        [Parameter(Position = 4)][int]$ProgressBarWidth = 50,
        [Parameter(Position = 5)][ValidateSet('Normal', 'Hidden', 'Minimized', 'Maximized')][string]$WindowStyle = 'Hidden',
        [switch]$NoNewWindow,
        [switch]$PassThru,
        [string]$RedirectStandardOutput,
        [string]$RedirectStandardError,
        [switch]$DontUseEscapeCodes,
        [switch]$ISEMode
    )
    function Test-TranscriptActive {
        <#
    .SYNOPSIS
        Checks if a PowerShell transcript is currently active using reflection.
    .DESCRIPTION
        Inspects internal PowerShell runspace data via reflection to determine if a transcript
        session is currently recording. This function is non-intrusive and does not start,
        stop, or modify any transcript sessions.
        The function uses .NET reflection to access the internal TranscriptionData property
        of the current runspace, which is not exposed through public PowerShell APIs.
        This is primarily used by Start-WithProgress to determine whether to use
        Console.Write (which bypasses transcripts) for progress bar updates.
    .OUTPUTS
        System.Boolean
            Returns $true if a transcript is currently active.
            Returns $false if no transcript is active or if any error occurs during detection.
    .EXAMPLE
        if (Test-TranscriptActive) {
            Write-Host "A transcript is currently recording."
        }
        Description: Check if transcript is active before performing an operation.
    .EXAMPLE
        Start-Transcript -Path "C:\Logs\session.txt"
        Test-TranscriptActive  # Returns $true
        Stop-Transcript
        Test-TranscriptActive  # Returns $false
        Description: Demonstrates transcript detection before and after starting a transcript.
    .EXAMPLE
        $wasTranscribing = Test-TranscriptActive
        if (-not $wasTranscribing) { Start-Transcript -Path $logPath }
        # ... do work ...
        if (-not $wasTranscribing) { Stop-Transcript }
        Description: Conditionally start transcript only if one isn't already running.
    .NOTES
        Author:         Harze2k
        Version:        1.0
        Compatibility:  PowerShell 5.1 and PowerShell 7+
        This function uses reflection to access internal PowerShell properties which may
        change in future PowerShell versions. If reflection fails, the function safely
        returns $false rather than throwing an error.
    .LINK
        Start-Transcript
    .LINK
        Stop-Transcript
    .LINK
        Start-WithProgress
    #>
        [CmdletBinding()]
        [OutputType([bool])]
        param()
        try {
            $flags = [System.Reflection.BindingFlags]'Instance, NonPublic, Public'
            $transcriptionData = $Host.Runspace.GetType().GetProperty('TranscriptionData', $flags).GetValue($Host.Runspace)
            if (-not $transcriptionData) { return $false }
            $transcripts = $transcriptionData.GetType().GetProperty('Transcripts', $flags).GetValue($transcriptionData)
            if (-not $transcripts) { return $false }
            if ($transcripts -is [System.Collections.ICollection]) { return $transcripts.Count -gt 0 }
            return (@($transcripts) | Measure-Object).Count -gt 0
        }
        catch { return $false }
    }
    function Write-ProgressBar {
        <#
    .SYNOPSIS
        Displays a customizable, single-line progress bar with color gradient in the console.
    .DESCRIPTION
        Renders a visually appealing progress bar that updates in place on a single line using
        carriage returns. The progress bar features a smooth color gradient from green (0%)
        through yellow (50%) to red (100%), providing intuitive visual feedback.
        Display Modes:
        - ANSI True-Color Mode: Uses 24-bit RGB colors for smooth gradients in modern terminals
        - ISE Mode: Falls back to ConsoleColor for PowerShell ISE compatibility
        - Basic Mode: Uses simple # and ░ characters when ANSI is disabled
        Transcript Handling:
        When UseConsoleWrite is enabled, the function uses [System.Console]::Write() instead
        of Write-Host. This bypasses PowerShell's transcript mechanism, preventing progress
        bar updates from cluttering transcript files while still displaying on screen.
        The progress bar format is:
        [████████████░░░░░░░░] 50.0% Elapsed: 00:00:30 Remaining: 00:00:30
    .PARAMETER PercentComplete
        The completion percentage to display, from 0 to 100.
        Values are clamped to this range automatically.
    .PARAMETER Width
        The character width of the progress bar (the filled/unfilled portion only).
        Does not include brackets, percentage, or time displays.
        Default is 50 characters.
    .PARAMETER TimeElapsed
        A TimeSpan representing how much time has elapsed.
        Displayed in hh:mm:ss format.
    .PARAMETER RemainingTime
        A TimeSpan representing the estimated time remaining.
        Displayed in hh:mm:ss format.
    .PARAMETER IsPsISE
        When $true, uses Write-Host with ConsoleColor instead of ANSI escape codes.
        PowerShell ISE does not support ANSI escape sequences.
    .PARAMETER UseConsoleWrite
        When specified, uses [System.Console]::Write() instead of Write-Host.
        This bypasses transcript recording, keeping transcripts clean during progress updates.
        Used automatically by Start-WithProgress when a transcript is active.
    .PARAMETER DontUseEscapeCodes
        When specified, uses basic ASCII characters (# and ░) instead of ANSI colors.
        Useful for terminals that don't support ANSI escape sequences or when piping output.
    .PARAMETER ReturnString
        When specified, returns the progress bar as a plain string instead of writing to console.
        The returned string contains no ANSI codes and is suitable for logging.
        Used by Start-WithProgress to log the final progress state to transcripts.
    .OUTPUTS
        System.String
            When -ReturnString is specified, returns the progress bar as a plain text string.
        System.Void
            When -ReturnString is not specified, writes directly to console and returns nothing.
    .EXAMPLE
        Write-ProgressBar -PercentComplete 50 -Width 40
        Description: Display a simple 50% progress bar with 40 character width.
    .EXAMPLE
        $elapsed = [TimeSpan]::FromSeconds(30)
        $remaining = [TimeSpan]::FromSeconds(30)
        Write-ProgressBar -PercentComplete 50 -TimeElapsed $elapsed -RemainingTime $remaining
        Description: Display progress bar with elapsed and remaining time.
    .EXAMPLE
        for ($i = 0; $i -le 100; $i += 10) {
            Write-ProgressBar -PercentComplete $i -Width 30
            Start-Sleep -Milliseconds 200
        }
        Write-Host ""  # Newline after progress completes
        Description: Animate a progress bar from 0% to 100%.
    .EXAMPLE
        $barString = Write-ProgressBar -PercentComplete 75 -ReturnString
        Add-Content -Path "log.txt" -Value $barString
        Description: Get progress bar as string for logging purposes.
    .EXAMPLE
        Write-ProgressBar -PercentComplete 100 -DontUseEscapeCodes
        Description: Display progress bar without ANSI colors for basic terminals.
    .NOTES
        Author:         Harze2k
        Version:        1.2
        Compatibility:  PowerShell 5.1 and PowerShell 7+
        The color gradient uses ANSI 24-bit true color escape sequences with three phases:
        - 0-50%:   Green → Yellow  (Red increases 0→255, Green stays 255)
        - 50-75%:  Yellow → Orange (Red stays 255, Green decreases 255→127)
        - 75-100%: Orange → Red    (Red stays 255, Green decreases 127→0)
        - Unfilled portion uses dark gray (RGB 100,100,100)
        For ISE compatibility, the gradient is approximated using four ConsoleColors:
        0-25% Green → 25-50% Yellow → 50-75% DarkYellow → 75-100% Red
    .LINK
        Start-WithProgress
    #>
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][ValidateRange(0, 100)][double]$PercentComplete,
            [int]$Width = 50,
            [TimeSpan]$TimeElapsed = [TimeSpan]::Zero,
            [TimeSpan]$RemainingTime = [TimeSpan]::Zero,
            [bool]$IsPsISE = $false,
            [switch]$UseConsoleWrite,
            [switch]$DontUseEscapeCodes,
            [switch]$ReturnString
        )
        $percentValue = [Math]::Max(0.0, [Math]::Min([Math]::Round($PercentComplete, 1), 100)) # Clamp and round to 2 decimals
        $completedWidth = [Math]::Max(0, [Math]::Min([int][Math]::Round(($percentValue / 100) * $Width), $Width)) # Width stays integer
        $remainingWidth = $Width - $completedWidth
        $progressDetails = "] {0,6:N1}% Elapsed: {1:hh\:mm\:ss} Remaining: {2:hh\:mm\:ss}   " -f $percentValue, $TimeElapsed, $RemainingTime
        if ($ReturnString) {
            $barChars = ('█' * $completedWidth) + ('░' * $remainingWidth)
            return "[{0}{1}" -f $barChars, $progressDetails
        }
        $escapeChar = [char]27
        $useAnsi = (-not $DontUseEscapeCodes)
        $isActuallyInISE = $Host.Name -eq 'Windows PowerShell ISE Host'
        if ($IsPsISE -and $isActuallyInISE) {
            Write-Host -NoNewline "`r["
            for ($i = 0; $i -lt $completedWidth; $i++) {
                $colorPct = if ($Width -gt 0) { ($i / $Width) * 100 } else { 0 }
                $color = if ($colorPct -lt 25) { 'Green' }
                elseif ($colorPct -lt 50) { 'Yellow' }
                elseif ($colorPct -lt 75) { 'DarkYellow' }
                else { 'Red' }
                Write-Host -NoNewline -ForegroundColor $color '#'
            }
            Write-Host -NoNewline -ForegroundColor DarkGray ('░' * $remainingWidth)
            Write-Host -NoNewline $progressDetails
            return
        }
        if ($IsPsISE -and -not $isActuallyInISE -and -not $DontUseEscapeCodes) {
            $barDisplay = [System.Text.StringBuilder]::new(512)
            for ($i = 0; $i -lt $Width; $i++) {
                if ($i -lt $completedWidth) {
                    $colorPct = if ($Width -gt 0) { $i / $Width } else { 0 }
                    if ($colorPct -le 0.5) {
                        $adjustedPct = $colorPct * 2
                        $red = [int](255 * $adjustedPct)
                        $green = 255
                    }
                    elseif ($colorPct -le 0.75) {
                        $adjustedPct = ($colorPct - 0.5) * 4
                        $red = 255
                        $green = [int](255 - (128 * $adjustedPct))
                    }
                    else {
                        $adjustedPct = ($colorPct - 0.75) * 4
                        $red = 255
                        $green = [int](127 * (1 - $adjustedPct))
                    }
                    [void]$barDisplay.Append("${escapeChar}[38;2;${red};${green};0m#${escapeChar}[0m")
                }
                else {
                    [void]$barDisplay.Append("${escapeChar}[38;2;100;100;100m░${escapeChar}[0m")
                }
            }
            $fullLine = "`r${escapeChar}[0m[{0}{1}" -f $barDisplay.ToString(), $progressDetails
            if ($UseConsoleWrite) { [System.Console]::Write($fullLine) }
            else { Write-Host -NoNewline $fullLine }
            return
        }
        $barDisplay = [System.Text.StringBuilder]::new(512)
        if ($useAnsi) {
            for ($i = 0; $i -lt $Width; $i++) {
                if ($i -lt $completedWidth) {
                    $colorPct = if ($Width -gt 0) { $i / $Width } else { 0 }
                    if ($colorPct -le 0.5) {
                        $adjustedPct = $colorPct * 2
                        $red = [int](255 * $adjustedPct)
                        $green = 255
                    }
                    elseif ($colorPct -le 0.75) {
                        $adjustedPct = ($colorPct - 0.5) * 4
                        $red = 255
                        $green = [int](255 - (128 * $adjustedPct))
                    }
                    else {
                        $adjustedPct = ($colorPct - 0.75) * 4
                        $red = 255
                        $green = [int](127 * (1 - $adjustedPct))
                    }
                    [void]$barDisplay.Append("${escapeChar}[38;2;${red};${green};0m█${escapeChar}[0m")
                }
                else {
                    [void]$barDisplay.Append("${escapeChar}[38;2;100;100;100m░${escapeChar}[0m")
                }
            }
        }
        else {
            [void]$barDisplay.Append('#' * $completedWidth)
            [void]$barDisplay.Append('░' * $remainingWidth)
        }
        $fullLine = "`r${escapeChar}[0m[{0}{1}" -f $barDisplay.ToString(), $progressDetails
        if ($UseConsoleWrite) { [System.Console]::Write($fullLine) }
        else { Write-Host -NoNewline $fullLine }
    }
    #region Initialization
    $hasNewLog = $null -ne (Get-Command -Name 'New-Log' -ErrorAction SilentlyContinue)
    if (-not $hasNewLog) {
        Write-Error "New-Log function not found. This function requires New-Log to be defined."
        return -1
    }
    if ($SleepInSeconds -gt $TimeoutInSeconds) {
        New-Log 'SleepInSeconds cannot be greater than TimeoutInSeconds.' -Level WARNING
        return -1
    }
    if (-not $PSBoundParameters.ContainsKey('ISEMode')) { $ISEMode = $Host.Name -eq 'Windows PowerShell ISE Host' }
    $isTranscriptActive = Test-TranscriptActive
    $useConsoleWrite = $isTranscriptActive -and (-not $ISEMode) -and (-not $DontUseEscapeCodes)
    $timeoutSpan = [TimeSpan]::FromSeconds($TimeoutInSeconds)
    $tempResources = @{ TempFiles = [System.Collections.Generic.List[string]]::new(); Process = $null }
    #endregion
    #region Helper Functions
    $showProgressBlock = {
        param([double]$Pct, [TimeSpan]$Elap, [TimeSpan]$Rem)
        Write-ProgressBar -PercentComplete $Pct -Width $ProgressBarWidth -TimeElapsed $Elap -RemainingTime $Rem -IsPsISE $ISEMode -UseConsoleWrite:$useConsoleWrite -DontUseEscapeCodes:$DontUseEscapeCodes
    }
    $finalizeProgressBlock = {
        param([double]$Pct, [TimeSpan]$Elap, [TimeSpan]$Rem, [bool]$HasShown)
        if (-not $HasShown) { return }
        if ($useConsoleWrite) {
            $finalBarString = Write-ProgressBar -PercentComplete $Pct -Width $ProgressBarWidth -TimeElapsed $Elap -RemainingTime $Rem -ReturnString
            Write-Host -NoNewline "`r$finalBarString"
            Write-ProgressBar -PercentComplete $Pct -Width $ProgressBarWidth -TimeElapsed $Elap -RemainingTime $Rem -UseConsoleWrite
            [System.Console]::WriteLine()
        }
        else {
            Write-ProgressBar -PercentComplete $Pct -Width $ProgressBarWidth -TimeElapsed $Elap -RemainingTime $Rem -IsPsISE $ISEMode -DontUseEscapeCodes:$DontUseEscapeCodes
            Write-Host
        }
    }
    #endregion
    try {
        #region JustWait Mode
        if ($PSCmdlet.ParameterSetName -eq 'JustWait') {
            if ([string]::IsNullOrWhiteSpace($Message)) { 
                New-Log "Waiting for $TimeoutInSeconds seconds..."
            }
            else {
                New-Log $Message
            }
            $startTime = Get-Date
            $hasShownProgress = $false
            while ($true) {
                $elapsed = (Get-Date) - $startTime
                $remaining = $timeoutSpan - $elapsed
                if ($remaining.TotalSeconds -lt 0) { $remaining = [TimeSpan]::Zero }
                $percent = if ($TimeoutInSeconds -gt 0) { 
                    [Math]::Round(($elapsed.TotalSeconds / $TimeoutInSeconds) * 100.0, 1) # 1 decimal
                }
                else { 
                    100 
                }
                if ($elapsed.TotalSeconds -ge $TimeoutInSeconds) {
                    if ($hasShownProgress) { & $finalizeProgressBlock $percent $elapsed $remaining $true }
                    New-Log "Successfully waited for $TimeoutInSeconds seconds" -Level SUCCESS
                    return
                }
                & $showProgressBlock $percent $elapsed $remaining
                $hasShownProgress = $true
                Start-Sleep -Seconds $SleepInSeconds
            }
        }
        #endregion
        #region Condition Mode
        if ($PSCmdlet.ParameterSetName -eq 'Condition') {
            if ($null -eq $Condition) {
                New-Log "Condition parameter is null." -Level WARNING
                return -1
            }
            if ($null -eq $Message) {
                New-Log "Starting condition.."
            }
            else {
                New-Log "Starting condition: [$Message]"
            }
            $startTime = Get-Date
            $conditionErrorCount = 0
            $hasShownProgress = $false
            while ($true) {
                $elapsed = (Get-Date) - $startTime
                $remaining = $timeoutSpan - $elapsed
                if ($remaining.TotalSeconds -lt 0) { $remaining = [TimeSpan]::Zero }
                $percent = if ($TimeoutInSeconds -gt 0) { 
                    [Math]::Round(($elapsed.TotalSeconds / $TimeoutInSeconds) * 100.0, 1) # 1 decimal
                }
                else { 
                    100 
                }
                try {
                    if ((& $Condition) -eq $true) {
                        if ($hasShownProgress) { & $finalizeProgressBlock $percent $elapsed $remaining $true }
                        New-Log "The condition: [$Message] was successfull!" -Level SUCCESS
                        if ($SuccessMessage) { New-Log $SuccessMessage -Level SUCCESS }
                        if ($PassThru) { return 0 } else { return }
                    }
                }
                catch {
                    $conditionErrorCount++
                    if ($conditionErrorCount -le 3) { New-Log "Error evaluating condition" -Level ERROR }
                }
                if ($elapsed.TotalSeconds -ge $TimeoutInSeconds) {
                    if ($hasShownProgress) { & $finalizeProgressBlock $percent $elapsed $remaining $true }
                    New-Log "The condition: [$Message] timed out." -Level WARNING
                    if ($PassThru) { return -1 } else { return }
                }
                & $showProgressBlock $percent $elapsed $remaining
                $hasShownProgress = $true
                Start-Sleep -Seconds $SleepInSeconds
            }
        }
        #endregion
        #region FilePath Mode
        if ($PSCmdlet.ParameterSetName -eq 'FilePath') {
            if ([string]::IsNullOrWhiteSpace($FilePath)) {
                New-Log "FilePath is required." -Level WARNING
                return -1
            }
            $originalFilePath = $FilePath
            if (-not (Test-Path $FilePath -PathType Leaf)) {
                $resolved = Get-Command $FilePath -ErrorAction SilentlyContinue
                if ($resolved -and $resolved.Source) { $FilePath = $resolved.Source }
                else {
                    $pathEntries = @($env:PATH -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                    $pathMatch = $pathEntries | ForEach-Object {
                        $candidate = Join-Path $_ $originalFilePath
                        if (Test-Path $candidate -PathType Leaf) { $candidate }
                    } | Select-Object -First 1
                    if ($pathMatch) { $FilePath = $pathMatch }
                    else {
                        New-Log "FilePath '$originalFilePath' not found." -Level WARNING
                        return -1
                    }
                }
            }
            if ($NoNewWindow -and $PSBoundParameters.ContainsKey('WindowStyle') -and $WindowStyle -ne 'Hidden') {
                New-Log "NoNewWindow and WindowStyle (non-Hidden) cannot be used together." -Level WARNING
                return -1
            }
            $executableName = Split-Path -Path $FilePath -Leaf
            $tempOutput = if ($RedirectStandardOutput) { $RedirectStandardOutput }
            else {
                $tf = [System.IO.Path]::GetTempFileName()
                $tempResources.TempFiles.Add($tf)
                $tf
            }
            $tempError = if ($RedirectStandardError) { $RedirectStandardError }
            else {
                $tf = [System.IO.Path]::GetTempFileName()
                $tempResources.TempFiles.Add($tf)
                $tf
            }
            if (-not $PSBoundParameters.ContainsKey('RedirectStandardOutput')) { [void](Set-Content -Path $tempOutput -Value $null -Encoding UTF8 -ErrorAction SilentlyContinue) }
            if (-not $PSBoundParameters.ContainsKey('RedirectStandardError')) { [void](Set-Content -Path $tempError -Value $null -Encoding UTF8 -ErrorAction SilentlyContinue) }
            New-Log "Executing: `"$FilePath`" $Arguments (Timeout: ${TimeoutInSeconds}s)"
            $proc = $null
            $tempBatch = $null
            try {
                $exeName = (Split-Path -Path $FilePath -Leaf).ToLowerInvariant()
                $hasTimeout = ($exeName -eq "cmd.exe") -and ($Arguments -match '\btimeout(\.exe)?\b')
                $tempBatch = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.bat'
                $tempResources.TempFiles.Add($tempBatch)
                $cmdLine = if ($exeName -eq "cmd.exe") {
                    if ($Arguments -match '^\s*/[ck]\s+(.+)$') {
                        $cmdSwitch = if ($Arguments -match '^\s*/c\s+') { '/c' } else { '/k' }
                        $cmdPart = $Arguments -replace '^\s*/[ck]\s+', ''
                        if ($hasTimeout) { """$FilePath"" $cmdSwitch ""$cmdPart"" 1>""$tempOutput"" 2>""$tempError""" }
                        else { """$FilePath"" $cmdSwitch ""$cmdPart"" <NUL 1>""$tempOutput"" 2>""$tempError""" }
                    }
                    else {
                        if ($hasTimeout) { """$FilePath"" $Arguments 1>""$tempOutput"" 2>""$tempError""" }
                        else { """$FilePath"" $Arguments <NUL 1>""$tempOutput"" 2>""$tempError""" }
                    }
                }
                else {
                    $cmdPath = (Get-Command cmd.exe -ErrorAction Stop).Source
                    """$cmdPath"" /S /C """"$FilePath"" $Arguments"" <NUL 1>""$tempOutput"" 2>""$tempError"""
                }
                $batchContent = "@echo off`r`n$cmdLine`r`nexit /b %errorlevel%`r`n"
                [void](Set-Content -Path $tempBatch -Value $batchContent -Encoding ASCII -Force)
                $psi = [System.Diagnostics.ProcessStartInfo]::new()
                $psi.FileName = $tempBatch
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                if ($NoNewWindow) { $psi.CreateNoWindow = $true }
                elseif ($WindowStyle -eq 'Hidden') { $psi.CreateNoWindow = $true }
                else {
                    $psi.UseShellExecute = $true
                    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::$WindowStyle
                }
                $proc = [System.Diagnostics.Process]::new()
                $proc.StartInfo = $psi
                $tempResources.Process = $proc
                [void]$proc.Start()
            }
            catch {
                New-Log "Failed to start process." -Level ERROR
                return -1
            }
            $startTime = Get-Date
            $exitCode = $null
            $hasShownProgress = $false
            while (-not $proc.HasExited) {
                $elapsed = (Get-Date) - $startTime
                $remaining = $timeoutSpan - $elapsed
                if ($remaining.TotalSeconds -lt 0) { $remaining = [TimeSpan]::Zero }
                $percent = if ($TimeoutInSeconds -gt 0) { [Math]::Min(100, ($elapsed.TotalSeconds / $TimeoutInSeconds) * 100) } else { 100 }
                if ($elapsed.TotalSeconds -ge $TimeoutInSeconds) {
                    if ($hasShownProgress) { & $finalizeProgressBlock $percent $elapsed $remaining $true }
                    New-Log "Process: [$executableName] timed out after ${TimeoutInSeconds}s." -Level WARNING
                    try {
                        $proc.Kill()
                        [void]$proc.WaitForExit(5000)
                    }
                    catch { }
                    $exitCode = -1
                    break
                }
                & $showProgressBlock $percent $elapsed $remaining
                $hasShownProgress = $true
                Start-Sleep -Seconds $SleepInSeconds
            }
            if ($null -eq $exitCode) {
                try { [void]$proc.WaitForExit() }
                catch { }
                $exitCode = $proc.ExitCode
                $finalElapsed = (Get-Date) - $startTime
                $finalRemaining = $timeoutSpan - $finalElapsed
                if ($finalRemaining.TotalSeconds -lt 0) { $finalRemaining = [TimeSpan]::Zero }
                $finalPercent = if ($TimeoutInSeconds -gt 0) { [Math]::Min(100, ($finalElapsed.TotalSeconds / $TimeoutInSeconds) * 100) } else { 100 }
                & $finalizeProgressBlock $finalPercent $finalElapsed $finalRemaining $hasShownProgress
                if ($exitCode -eq 0) { New-Log "Process: [$executableName] completed successfully, exit code: [$exitCode]." -Level SUCCESS }
                else {
                    New-Log "Process: [$executableName] exited with code: [$exitCode]." -Level WARNING
                    if (-not $PSBoundParameters.ContainsKey('RedirectStandardError') -and (Test-Path $tempError)) {
                        $errContent = Get-Content $tempError -Raw -ErrorAction SilentlyContinue
                        if ($errContent -and $errContent.Trim()) { New-Log "StdErr: $($errContent.Trim())" -Level WARNING }
                    }
                }
            }
            if ($PassThru) { return $exitCode } else { return }
        }
        #endregion
    }
    catch {
        New-Log "Unhandled exception in Start-WithProgress" -Level ERROR
        return -1
    }
    finally {
        if ($tempResources.Process) {
            try {
                if (-not $tempResources.Process.HasExited) {
                    $tempResources.Process.Kill()
                    [void]$tempResources.Process.WaitForExit(2000)
                }
                $tempResources.Process.Dispose()
            }
            catch { }
        }
        foreach ($tf in $tempResources.TempFiles) {
            if ($tf -and (Test-Path $tf -PathType Leaf)) {
                $isUserOutput = $PSBoundParameters.ContainsKey('RedirectStandardOutput') -and $tf -eq $RedirectStandardOutput
                $isUserError = $PSBoundParameters.ContainsKey('RedirectStandardError') -and $tf -eq $RedirectStandardError
                if (-not $isUserOutput -and -not $isUserError) {
                    try { Remove-Item $tf -Force -ErrorAction Stop } catch { Start-Sleep -Milliseconds 100; Remove-Item $tf -Force -ErrorAction SilentlyContinue }
                }
            }
        }
    }
}
<#
Please use my New-Log function for loggin to the console, otherwise this wont work. Check it out on my GitHub: 
https://github.com/Harze2k/Shared-PowerShell-Functions/blob/main/New-Log.ps1
One-liner to use it:
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
#>