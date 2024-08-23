<#
.DESCRIPTION
The New-Log function is a versatile logging utility designed to provide detailed logs with customizable formatting for PowerShell scripts.
It supports logging messages of different severity levels, including ERROR, WARNING, INFO, SUCCESS, and DEBUG.
The function also handles different PowerShell versions, including PowerShell Core (6 and above), and adjusts the console output for color-coded messages based on the severity level.
Additionally, it provides detailed information for ERROR logs, including the script line number, function name, and the exact code that failed, enhancing the debugging process.

.PARAMETER Message
-Message: Accepts the log message, which can be a string, hashtable, or PSCustomObject.

.PARAMETER Level
-Level: Specifies the severity level of the log message, defaulting to INFO.

.PARAMETER NoConsole
-NoConsole: Suppresses console output if specified, useful for silent logging.

.PARAMETER PassThru
-PassThru: Returns the formatted log message as a string instead of writing it to the console or file.

.PARAMETER AsObject
-AsObject: Returns the log entry as a PSCustomObject, which can be useful for further processing or outputting structured data.

.PARAMETER ForcedLogFile
-ForcedLogFile: Overwrites the existing log file with the new log entry if specified, otherwise appends to the log file.

.PARAMETER LogFilePath
-LogFilePath: Specifies the path to the log file where the message should be written. If the directory does not exist, it is created automatically.

.EXAMPLE
Example 1: Log an informational message to the console
New-Log -Message "The process completed successfully." -Level "INFO"

Example 2: Log an error message to a log file. These messeges are standard:
Timestamp      : 2024-08-16 09:56:02.178
Level          : ERROR
Message        : A critical error occurred in the script.
Exception      : Cannot find path 'C:\ttmm' because it does not exist.
CallerFunction : <Name of function used>
CodeRow        : (2,433) (Function,Script)
FailedCode     : Get-ChildItem -Path C:\ttmm -ErrorAction Stop

try {
    Get-ChildItem -Path C:\ttmm -ErrorAction Stop
}
catch {
    New-Log -Message "A critical error occurred in the script." -Level "ERROR" -LogFilePath "C:\Logs\error.log"
}
Example 3: Log a debug message without console output but returning the log as a string
$logEntry = New-Log -Message "Debugging the script." -Level "DEBUG" -NoConsole -PassThru

Example 4: Log a success message and return it as a PSCustomObject for further processing
$logObject = New-Log -Message "Operation completed successfully." -Level "SUCCESS" -AsObject

Example 5: Overwrite the existing log file with a warning message
New-Log -Message "This will overwrite the log file." -Level "WARNING" -LogFilePath "C:\Logs\warning.log" -ForcedLogFile

Example 6: Log a message with a PSCustomObject and output it both to the console and as a PSCustomObject
$customMessage = [PSCustomObject]@{
    UserName = "Admin"
    Action   = "Login"
    Status   = "Success"
}
$returnedObject = $customMessage | New-Log -Level "INFO" -PassThru -AsObject
#>
function New-Log {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true)]$Message,
        [Parameter(Position = 1)][ValidateSet("ERROR", "WARNING", "INFO", "SUCCESS", "DEBUG")][string]$Level = "INFO",
        [Parameter(Position = 2)][switch]$NoConsole,
        [Parameter(Position = 3)][switch]$PassThru,
        [Parameter(Position = 4)][switch]$AsObject,
        [Parameter(Position = 5)][switch]$ForcedLogFile,
        [Parameter(Position = 6)][string]$LogFilePath
    )
    Begin {
        $isPSCore = $PSVersionTable.PSVersion.Major -ge 6
        $levelColors = @{
            "ERROR"   = @{ANSI = "31"; PS = "Red" }
            "WARNING" = @{ANSI = "33"; PS = "Yellow" }
            "SUCCESS" = @{ANSI = "32"; PS = "Green" }
            "DEBUG"   = @{ANSI = "34"; PS = "Blue" }
            "INFO"    = @{ANSI = "37"; PS = "White" }
        }
        $reset = if ($isPSCore) {
            "`e[0m"
        }
        else {
            ""
        }
        $blue = if ($isPSCore) {
            "`e[34m"
        }
        else {
            ""
        }
        try {
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        }
        catch {
            Write-Host "Unable to set console encoding to UTF8" -ForegroundColor Yellow
        }
        function Write-MessageToConsole {
            if ($LogSentToConsole -eq $true) {
                return
            }
            if (!($NoConsole.IsPresent)) {
                if ($isPSCore) {
                    Write-Host $logMessage
                }
                else {
                    $logMessage | ForEach-Object { Write-Host $_ -ForegroundColor $levelColors[$Level].PS }
                }
            }
            return $true
        }
    }
    Process {
        if ($null -eq $Message -and $Level -ne "ERROR") {
            return
        }
        try {
            if ($Message -and $Message.GetType().Name -eq 'Hashtable') {
                $Message = New-Object -TypeName PSObject -Property $Message
            }
            if ($Message -and $Message.GetType().Name -notin @("PSCustomObject", "Hashtable", "String", "Software")) {
                Write-Host "Unsupported message type: $($Message.GetType().Name). Must be PSCustomObject, Hashtable or string" -ForegroundColor Red
                return
            }
            $logSentToConsole = $false
            $logMessage = ''
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
            $callerInfo = (Get-PSCallStack)[1]
            $originalMessage = $Message
            $levelColor = if ($isPSCore) {
                $levelColors[$Level].ANSI
            }
            else {
                $levelColors[$Level].PS
            }
            $headerPrefix = if ($isPSCore) {
                "$reset[$blue$timestamp$reset][`e[$($levelColor)m$Level$reset]"
            }
            else {
                "[$timestamp][$Level]"
            }
            if ($Message -isnot [string]) {
                $Message = ($Message | Format-List | Out-String).Trim()
            }
            if ($callerInfo.FunctionName -ne '<ScriptBlock>') {
                $functionInfo, $messageLines = if ($isPSCore) {
                    if (!($Message)) {
                        "[${blue}Function${reset}: $($callerInfo.FunctionName)]"; "$headerPrefix${_}"
                    }
                    else {
                        " [${blue}Function${reset}: $($callerInfo.FunctionName)]"; $Message -split "`n" | ForEach-Object { "$headerPrefix $_" }
                    }
                }
                else {
                    if (!($Message)) {
                        "[Function: $($callerInfo.FunctionName)]"; "$headerPrefix${_}"
                    }
                    else {
                        " [Function: $($callerInfo.FunctionName)]"; $Message -split "`n" | ForEach-Object { "$headerPrefix $_" }
                    }
                }
                $logMessage += ($messageLines -join "`n") + $functionInfo
            }
            else {
                $messageLines = if (!($Message)) {
                    "$headerPrefix${_}"
                }
                else {
                    $Message -split "`n" | ForEach-Object { "$headerPrefix $_" }
                }
                $logMessage += $messageLines -join "`n"
            }
            if ($Level -eq "ERROR" -and $Error[0]) {
                $errorRecord = $Error[0]
                $invocationInfo = $errorRecord.InvocationInfo
                try {
                    if ($ErrorRecord.InvocationInfo.PSCommandPath -and (Test-Path -Path $ErrorRecord.InvocationInfo.PSCommandPath)) {
                        $scriptLines = Get-Content -Path "$($ErrorRecord.InvocationInfo.PSCommandPath)" -ErrorAction Stop
                    }
                    elseif ($ErrorRecord.InvocationInfo.ScriptName -and (Test-Path -Path $ErrorRecord.InvocationInfo.ScriptName)) {
                        $scriptLines = Get-Content -Path "$($ErrorRecord.InvocationInfo.ScriptName)" -ErrorAction Stop
                    }
                }
                catch {
                    if ($isPSCore) {
                        Write-Host "$reset[$blue$timestamp$reset][$($reset)e[31mERROR$reset] An error occurred in New-Log function. $($reset)e[31m$($_.Exception.Message)$reset"
                    }
                    else {
                        Write-Host "[$timestamp][ERROR] An error occurred in New-Log function. $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                $functionName = $callerInfo.Command
                $failedCode = $invocationInfo.Line.Trim()
                [int]$errorLine = $errorRecord.InvocationInfo.ScriptLineNumber
                if ([string]::IsNullOrEmpty($errorLine)) {
                    [int]$errorLine = $invocationInfo.ScriptLineNumber
                }
                if ($null -ne $scriptLines) {
                    [int]$functionStartLine = ($scriptLines | Select-String -Pattern "function\s+$functionName" | Select-Object -First 1).LineNumber
                    $lineNumberInFunction = $errorLine - $functionStartLine
                    $lineInfo = "($lineNumberInFunction,$errorLine) (Function,Script)"
                    if ($callerInfo.FunctionName -eq '<ScriptBlock>') {
                        $lineInfo = "$errorLine (Script)"
                    }
                }
                else {
                    $lineNumberInFunction = $errorLine - ([int]$callerInfo.ScriptLineNumber - [int]$invocationInfo.OffsetInLine) - 1
                    $lineInfo = "($lineNumberInFunction,$errorLine) (Function,Script)"
                    if ($callerInfo.FunctionName -eq '<ScriptBlock>') {
                        $lineInfo = "$errorLine (Script)"
                    }
                }
                $exceptionMessage = $($errorRecord.Exception.Message)
                if ($isPSCore) {
                    $logMessage += "[${blue}CodeRow${reset}: $lineInfo]"
                    $logMessage += "[${blue}FailedCode${reset}: $failedCode]"
                    $logMessage += "[${blue}ExceptionMessage${reset}: ${reset}`e[$($levelColors[$Level].ANSI)m$exceptionMessage$reset]"
                }
                else {
                    $logMessage += "[CodeRow: $lineInfo]"
                    $logMessage += "[FailedCode: $failedCode]"
                    $logMessage += "[ExceptionMessage: $exceptionMessage]"
                }
            }
            if (!($NoConsole.IsPresent) -and !($PassThru.IsPresent) -and !($AsObject.IsPresent) -and !($LogFilePath)) {
                $LogSentToConsole = Write-MessageToConsole
            }
            if ($LogFilePath) {
                $LogSentToConsole = Write-MessageToConsole
                $logMessage = [regex]::Replace($logMessage, $([regex]::Escape("`e") + '\[[0-9;]*[mGKHF]'), '')
                if (!(Test-Path -Path (Split-Path -Path $LogFilePath -Parent))) {
                    New-Item -Path (Split-Path -Path $LogFilePath -Parent) -ItemType Directory -Force -ErrorAction Stop | Out-Null
                }
                if ($ForcedLogFile.IsPresent) {
                    Remove-Item -Path $LogFilePath -Force -ErrorAction SilentlyContinue | Out-Null
                    Set-Content -Value $logMessage -Path $LogFilePath -Force -Encoding utf8
                }
                else {
                    $logMessage | Out-File -FilePath $LogFilePath -Append -Encoding utf8
                }
            }
            $object = [PSCustomObject]@{
                Timestamp      = $timestamp
                Level          = $Level
                Message        = if ($originalMessage -and $originalMessage.GetType().Name -eq 'String') {
                    $Message
                }
                else {
                    [pscustomobject](($Message | Format-List | Out-String).Trim()) -split "`n"
                }
                Exception      = if (Get-Variable -Name 'exceptionMessage' -ErrorAction SilentlyContinue) {
                    $exceptionMessage
                }
                else {
                    $null
                }
                CallerFunction = if ($callerInfo.FunctionName -eq '<ScriptBlock>') {
                    $null
                }
                else {
                    $callerInfo.FunctionName
                }
                CodeRow        = if (Get-Variable -Name 'lineInfo' -ErrorAction SilentlyContinue) {
                    $lineInfo
                }
                else {
                    $null
                }
                FailedCode     = if (Get-Variable -Name 'failedCode' -ErrorAction SilentlyContinue) {
                    $failedCode
                }
                else {
                    $null
                }
            }
            if ($PassThru.IsPresent -and $AsObject.IsPresent) {
                $LogSentToConsole = Write-MessageToConsole
                return $object
            }
            elseif ($PassThru.IsPresent -and !($AsObject.IsPresent)) {
                $LogSentToConsole = Write-MessageToConsole
                return $logMessage
            }
            elseif (!($NoConsole.IsPresent) -and $AsObject.IsPresent) {
                $object | Out-Host
            }
        }
        catch {
            if ($isPSCore) {
                Write-Host "$reset[$blue$timestamp$reset][`e[31mERROR$reset] An error occurred in New-Log function. `e[31m$($_.Exception.Message)$reset"
            }
            else {
                Write-Host "[$timestamp][ERROR] An error occurred in New-Log function. $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}