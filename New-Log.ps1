<#
.SYNOPSIS
    Writes a log entry to the console and/or to a log file, with optional formatting and error information.
.DESCRIPTION
    The New-Log function allows you to write log entries at various levels (ERROR, WARNING, INFO, SUCCESS, DEBUG, VERBOSE).
    You can pass any type of message: string, hashtable, PSCustomObject, or even ErrorRecord objects. The function can:
    - Display logs in the console with color-coding for different log levels.
    - Write logs to a file, with optional file rotation based on size.
    - Output logs in JSON or plain-text format.
    - Create a new log file with a timestamp appended to its name.
    - Overwrite an existing file if desired (ForcedLogFile).
    - Return the log entry as an object for further processing.
    - Skip console output completely if needed.
.PARAMETER Message
    The main content of your log entry. Can be a string, hashtable, PSCustomObject, or ErrorRecord.
.PARAMETER Level
    Specifies the severity or type of message to log.
    Valid options: ERROR, WARNING, INFO, SUCCESS, DEBUG, VERBOSE.
    Default is INFO.
.PARAMETER NoConsole
    If set, no output is written to the console.
.PARAMETER ReturnObject
    If set, returns the constructed log object instead of writing nothing or just logging to file/console.
.PARAMETER LogFilePath
    Specifies the full path to the log file. If not provided, logs will only appear in the console (unless NoConsole is set).
.PARAMETER ForcedLogFile
    If set, overwrites the entire log file instead of appending.
.PARAMETER AppendTimestampToFile
    If set, appends a timestamp to the log file name before writing.
    For example, LogFilePath "C:\Logs\MyLog.txt" might become "C:\Logs\MyLog_20250225_153000.txt".
.PARAMETER LogRotationSizeMB
    When specified with a nonzero value, checks the size of the current log file and renames (rotates) it if it exceeds the size in MB.
.PARAMETER LogFormat
    Controls how the log message is written to the file. Accepts "TEXT" (default) or "JSON".
.EXAMPLE
    # Simple usage (INFO level by default):
    New-Log "This is a test message"
.EXAMPLE
    # Log an ERROR. If an ErrorRecord is passed, exception details are included automatically:
    try {
        1/0
    }
    catch {
        New-Log $_ -Level ERROR
    }
.EXAMPLE
    # Write a WARNING to a file and the console:
    New-Log "This will go to the log file" -Level WARNING -LogFilePath "C:\Logs\ExampleLog.txt"
.EXAMPLE
    # Overwrite the file each time (-ForcedLogFile) and return the log object (-ReturnObject):
    $logEntry = New-Log "Overwriting file with a new entry" -Level INFO -ForcedLogFile -LogFilePath "C:\Logs\Overwrite.txt" -ReturnObject
    # $logEntry now holds the PSCustomObject with all log details.
.EXAMPLE
    # Append a timestamp to the log file name:
    New-Log "Log with time-based filename" -Level DEBUG -LogFilePath "C:\Logs\DailyLog.txt" -AppendTimestampToFile
.EXAMPLE
    # Skip console output and only write JSON logs to file:
    New-Log "Background log entry" -Level VERBOSE -NoConsole -LogFilePath "C:\Logs\MyLog.json" -LogFormat JSON
.NOTES
    Written for Windows PowerShell 5.1 or newer. Automatically handles console color adjustments.
    If running on PowerShell 7 or newer, advanced ANSI coloring is used for console messages.
#>
function New-Log {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, Position = 0)]$Message,
        [Parameter(Position = 1)][ValidateSet("ERROR", "WARNING", "INFO", "SUCCESS", "DEBUG", "VERBOSE")][string]$Level = "INFO",
        [Parameter(Position = 2)][switch]$NoConsole,
        [Parameter(Position = 3)][switch]$ReturnObject,
        [Parameter(Position = 4)][string]$LogFilePath,
        [Parameter(Position = 5)][switch]$ForcedLogFile,
        [Parameter(Position = 6)][switch]$AppendTimestampToFile,
        [Parameter(Position = 7)][int]$LogRotationSizeMB = 0,
        [Parameter(Position = 8)][ValidateSet("TEXT", "JSON")][string]$LogFormat = "TEXT"
    )
    Begin {
        $script:isPSCore = $PSVersionTable.PSVersion.Major -ge 6
        $levelColors = @{
            ERROR   = @{ ANSI = 91; PS = 'Red' }
            WARNING = @{ ANSI = 93; PS = 'Yellow' }
            INFO    = @{ ANSI = 37; PS = 'White' }
            SUCCESS = @{ ANSI = 92; PS = 'Green' }
            DEBUG   = @{ ANSI = 94; PS = 'Blue' }
            VERBOSE = @{ ANSI = 96; PS = 'Cyan' }
        }
        $utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $true
        try {
            if ($script:isPSCore) {
                $originalEncoding = [Console]::OutputEncoding
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            }
            else {
                $originalEncoding = [Console]::OutputEncoding
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            }
        }
        catch {
            Write-Verbose "Failed to set console encoding: $($_.Exception.Message)"
        }
        if ($LogFilePath -and $AppendTimestampToFile) {
            $fileInfo = [System.IO.FileInfo]::new($LogFilePath)
            $LogFilePath = "{0}\{1}_{2}{3}" -f $fileInfo.Directory, $fileInfo.BaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'), $fileInfo.Extension
        }
    }
    Process {
        try {
            if ($null -eq $Message -and $Level -ne 'ERROR') { return }
            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
            $callStack = Get-PSCallStack -ErrorAction SilentlyContinue
            $colorCode = $levelColors[$Level].ANSI
            $ansiReset = if ($script:isPSCore) { "`e[0m" } else { '' }
            $msgData = switch ($Message) {
                { $_ -is [System.Collections.Hashtable] -or $_ -is [PSCustomObject] } {
                    $originalMessage = $Message
                    try {
                        if ($_ -is [System.Collections.Hashtable]) {
                            [PSCustomObject]$_ | ConvertTo-Json -Compress -Depth 5 -ErrorAction Stop
                        }
                        else {
                            $_ | ConvertTo-Json -Compress -Depth 5 -ErrorAction Stop
                        }
                    }
                    catch { ($_ | Out-String -Width 4096).Trim() }
                }
                { $_ -is [string] } { $_ }
                { $_ -is [System.Management.Automation.ErrorRecord] } {
                    $($_.Exception.Message)
                }
                default { ($_ | Out-String -Width 4096).Trim() }
            }
            $exceptionMsg = $failedCode = $callerInfo = $callerInfoPSCore = $null
            if ($Level -eq 'ERROR') {
                $err = if ($Message -is [System.Management.Automation.ErrorRecord]) {
                    $Message
                }
                elseif ($Error.Count -gt 0) {
                    $Error[0]
                }
                else {
                    $null
                }
                if ($err) {
                    $exceptionMsg = $err.Exception.Message
                    $failedCode = if ($err.InvocationInfo) { $err.InvocationInfo.Line.Trim() } else { "N/A" }
                    $relevantCall = @()
                    $callStack = $callStack | Sort-Object -Property ScriptLineNumber -Descending
                    for ($i = 0; $i -lt $callStack.Count; $i++) {
                        $prop = $callStack[$i]
                        if ($prop.Command -eq "New-Log") {
                            for ($j = $i - 1; $j -ge 0; $j--) {
                                if ($callStack[$j].FunctionName -notmatch '<ScriptBlock>' -or $callStack[$j].Command -ne $callStack[0].Command) {
                                    $relevantCall = [PSCustomObject]@{
                                        CalledFromLine   = $callStack[0].ScriptLineNumber
                                        FunctionName     = $callStack[$j].FunctionName
                                        ScriptLineNumber = $callStack[$j].ScriptLineNumber
                                    }
                                }
                                else {
                                    $relevantCall = [PSCustomObject]@{
                                        CalledFromLine = $callStack[0].ScriptLineNumber
                                    }
                                }
                                break
                            }
                            break
                        }
                    }
                    if ($relevantCall.CalledFromLine) {
                        $callerInfo = "Called from line: $($relevantCall.CalledFromLine)"
                        $callerInfoPSCore = "Called from file line: `e[34m$($relevantCall.CalledFromLine)$ansiReset"
                    }
                    if ($relevantCall.FunctionName) {
                        $callerInfo += ", in function [$($relevantCall.FunctionName)] on file line: $($relevantCall.ScriptLineNumber)"
                        $callerInfoPSCore += ", in function `e[34m[$($relevantCall.FunctionName)]$ansiReset on file line: `e[34m$($relevantCall.ScriptLineNumber)$ansiReset"
                    }
                }
            }
            $logEntry = [PSCustomObject]@{
                Timestamp = $timestamp
                Level     = $Level
                Message   = $msgData
            }
            if ($Level -eq 'ERROR') {
                if ($callerInfo) { $logEntry | Add-Member -NotePropertyName CallerInfo -NotePropertyValue $callerInfo }
                if ($failedCode) { $logEntry | Add-Member -NotePropertyName FailedCode -NotePropertyValue $failedCode }
                if ($exceptionMsg) { $logEntry | Add-Member -NotePropertyName Exception -NotePropertyValue $exceptionMsg }
            }
            if ($LogFilePath) {
                $parentDir = Split-Path $LogFilePath -Parent
                if (-not (Test-Path $parentDir)) {
                    New-Item $parentDir -ItemType Directory -Force | Out-Null
                }
                if ($LogRotationSizeMB -gt 0 -and (Test-Path $LogFilePath)) {
                    $logFile = Get-Item $LogFilePath
                    if ($logFile.Length -gt ($LogRotationSizeMB * 1MB)) {
                        $backupPath = "{0}\{1}_{2}{3}" -f $logFile.DirectoryName, $logFile.BaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'), $logFile.Extension
                        Move-Item -Path $LogFilePath -Destination $backupPath -Force
                    }
                }
                if ($LogFormat -eq "JSON") {
                    $fileMessage = $logEntry | ConvertTo-Json -Compress
                }
                else {
                    $fileMessage = "[$timestamp][$Level] $msgData"
                    if ($originalMessage) {
                        $fileMessage = "[$timestamp][$Level] `n$(($originalMessage | Format-Table -AutoSize | Out-String).Trim())"
                    }
                    if ($Level -eq 'ERROR' -and $callerInfo) {
                        $fileMessage += " [CallerInfo: $callerInfo]"
                        if ($failedCode) { $fileMessage += "[FailedCode: $failedCode]" }
                        if ($exceptionMsg) { $fileMessage += "[Exception: $exceptionMsg]" }
                    }
                }
                if ($ForcedLogFile.IsPresent) {
                    [System.IO.File]::WriteAllText($LogFilePath, $fileMessage, $utf8NoBomEncoding)
                }
                else {
                    $content = [System.IO.File]::ReadAllText($LogFilePath, $utf8NoBomEncoding)
                    $content += "`r`n" + $fileMessage
                    [System.IO.File]::WriteAllText($LogFilePath, $content, $utf8NoBomEncoding)
                }
            }
            if (-not $NoConsole) {
                if ($originalMessage) {
                    $prefixMessage = if ($script:isPSCore) {
                        "`e[34m[$timestamp]$ansiReset`e[${colorCode}m[$Level]$ansiReset"
                    }
                    else {
                        "[$timestamp][$Level]"
                    }
                    Write-Host $prefixMessage -ForegroundColor $levelColors[$Level].PS -NoNewline
                    $tableOutput = ($originalMessage | Format-Table -AutoSize | Out-String).Trim()
                    Write-Host "`n$tableOutput" -ForegroundColor $levelColors[$Level].PS
                }
                else {
                    $consoleMessage = if ($script:isPSCore) {
                        "`e[34m[$timestamp]$ansiReset`e[${colorCode}m[$Level]$ansiReset `e[37m$msgData"
                    }
                    else {
                        "[$timestamp][$Level] $msgData"
                    }
                    if ($Level -eq 'ERROR' -and $script:isPSCore) {
                        if ($callerInfoPSCore) { $consoleMessage += " `e[${colorCode}m[CallerInfo:$ansiReset $callerInfoPSCore`e[${colorCode}m]" }
                        if ($failedCode) { $consoleMessage += "`e[${colorCode}m[FailedCode:$ansiReset `e[37m$failedCode$ansiReset`e[${colorCode}m]" }
                        if ($exceptionMsg) { $consoleMessage += "`e[${colorCode}m[Exception:$ansiReset `e[37m$exceptionMsg$ansiReset`e[${colorCode}m]" }
                    }
                    elseif ($Level -eq 'ERROR') {
                        if ($callerInfo) { $consoleMessage += " [CallerInfo: $callerInfo]" }
                        if ($failedCode) { $consoleMessage += "[FailedCode: $failedCode]" }
                        if ($exceptionMsg) { $consoleMessage += "[Exception: $exceptionMsg]" }
                    }
                    Write-Host $consoleMessage -ForegroundColor $levelColors[$Level].PS
                }
            }
            if ($ReturnObject) {
                return $logEntry
            }
        }
        catch {
            Write-Error "Logging failed: $($_.Exception.Message)"
        }
    }
    End {
        try {
            [Console]::OutputEncoding = $originalEncoding
        }
        catch {
            [void]$null
        }
    }
}