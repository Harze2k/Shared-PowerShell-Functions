#Requires -Version 7.0
<#
.SYNOPSIS
    Automates multi-threaded folder backups to a local, NAS, or Cloud destination using Robocopy.
.DESCRIPTION
    Run-AutoBackup is a high-performance backup wrapper for Robocopy. 
    It leverages PowerShell 7 parallel processing to backup multiple directories simultaneously 
    while using Robocopy's multi-threading for individual files. 
    It is pre-configured with parameters to handle common NAS/Cloud quirks, 
    such as timestamp rounding (/FFT) and daylight saving time shifts (/DST). 
    It also supports granular file and directory exclusions to optimize backup times.
.PARAMETER SourcePaths
    An array of local directory paths you want to backup.
.PARAMETER DestinationBase
    The root directory where the backups will be stored (e.g., a NAS path, external drive, or virtual cloud drive).
    The original folder structure (minus the drive letter) will be recreated here.
.PARAMETER ExcludeDirectories
    An array of specific folder paths to exclude from the backup.
.PARAMETER ExcludeFiles
    An array of file names or wildcard extensions to ignore (e.g., "*.log", "*.tmp").
    By default, it ignores common temporary, cache, and system files.
.PARAMETER ThrottleLimit
    The maximum number of directories to backup at the exact same time. Default is 3.
.PARAMETER RobocopyThreads
    The number of threads Robocopy uses per directory task (/MT). Default is 4.
.PARAMETER LogScriptUrl
    The URL to raw code for the advanced New-Log function. If New-Log is not already loaded, 
    it will be downloaded. If it fails, a lightweight fallback logger is created in-memory.
.EXAMPLE
    $sources = @("C:\Users\Martin\Documents", "C:\Temp")
    Run-AutoBackup -SourcePaths $sources -DestinationBase "G:\AutoBackup"
.EXAMPLE
    $sources = @(
		"C:\Users\Martin\GitHub-Harze2k"
		"C:\Users\Martin\AppData\Local\zen"
		"C:\Users\Martin\AppData\Local\Plex Media Server"
		"C:\Users\Martin\AppData\Local\qBittorrent"
		"C:\Users\Martin\AppData\Roaming\zen"
		"C:\Users\Martin\AppData\Local\Zen Browser"
		"C:\Users\Martin\AppData\Roaming\mpv.net"
		"C:\Users\Martin\AppData\Roaming\Code - Insiders"
		"C:\Users\Martin\AppData\Roaming\SVP4"
		"C:\Users\Martin\Documents\PowerShell"
		"C:\Toolkit\Toolkit_v13.7\Data"
		"C:\Toolkit\Toolkit_v13.7\Custom"
		"C:\Temp"
		"C:\Program Files\PowerShell\Modules"
		"C:\Program Files\totalcmd"
	)
	$excludeDirs = @(
		"C:\Users\Martin\GitHub-Harze2k\NodeJs"
		"C:\Users\Martin\AppData\Local\Plex Media Server\Cache"
	)
	$excludeFiles = @("parent.lock", "*.log", "*.tmp", "*.temp", "*.cache", "*.bak", "~*", "Thumbs.db", "desktop.ini")
	Run-AutoBackup -SourcePaths $sources -DestinationBase "G:\AutoBackup" -ExcludeDirectories $excludeDirs -ExcludeFiles $excludeFiles
.NOTES
    Author: Harze2k (Martin)
    Requires: PowerShell 7.0 or newer (uses ForEach-Object -Parallel)
	
	Example output (when everything has been copied at least once, first run takes longer ofc):

	[2026-02-24 20:38:37.320][SUCCESS] Starting backup of 15 tasks
	[2026-02-24 20:38:52.389][SUCCESS] [1/15] OK (exit 1): C:\Users\Martin\AppData\Local\zen                                
	[2026-02-24 20:38:55.228][SUCCESS] [2/15] OK (exit 1): C:\Users\Martin\AppData\Local\qBittorrent
	[2026-02-24 20:38:56.138][SUCCESS] [3/15] OK (exit 1): C:\Users\Martin\GitHub-Harze2k
	[2026-02-24 20:38:56.194][SUCCESS] [4/15] OK (exit 0): C:\Users\Martin\AppData\Local\Zen Browser
	[2026-02-24 20:38:56.355][SUCCESS] [5/15] OK (exit 1): C:\Users\Martin\AppData\Roaming\mpv.net
	[2026-02-24 20:38:56.404][SUCCESS] [6/15] OK (exit 0): C:\Users\Martin\AppData\Roaming\Code - Insiders
	[2026-02-24 20:38:56.442][SUCCESS] [7/15] OK (exit 1): C:\Users\Martin\AppData\Roaming\SVP4
	[2026-02-24 20:38:56.478][SUCCESS] [8/15] OK (exit 0): C:\Users\Martin\Documents\PowerShell
	[2026-02-24 20:38:57.168][SUCCESS] [9/15] OK (exit 0): C:\Toolkit\Toolkit_v13.7\Data
	[2026-02-24 20:38:57.519][SUCCESS] [10/15] OK (exit 0): C:\Toolkit\Toolkit_v13.7\Custom
	[2026-02-24 20:39:05.300][SUCCESS] [11/15] OK (exit 3): C:\Temp
	[2026-02-24 20:39:06.320][SUCCESS] [12/15] OK (exit 0): C:\Program Files\PowerShell\Modules
	[2026-02-24 20:39:06.370][SUCCESS] [13/15] OK (exit 1): C:\Program Files\totalcmd
	[2026-02-24 20:39:06.396][SUCCESS] [14/15] OK (exit 3): C:\Users\Martin\AppData\Roaming\zen
	[2026-02-24 20:39:15.029][SUCCESS] [15/15] OK (exit 1): C:\Users\Martin\AppData\Local\Plex Media Server
	[2026-02-24 20:39:19.858][INFO] Total runtime: 00:00:39
#>
function Run-AutoBackup {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, HelpMessage = "Array of folder paths to backup.")][string[]]$SourcePaths,
		[Parameter(Mandatory = $true, HelpMessage = "The base destination directory where files will be copied.")][string]$DestinationBase,
		[Parameter(Mandatory = $false, HelpMessage = "Array of directory paths to exclude.")][string[]]$ExcludeDirectories = @(),
		[Parameter(Mandatory = $false, HelpMessage = "Array of file names or extensions to exclude.")][string[]]$ExcludeFiles = @(),
		[Parameter(Mandatory = $false, HelpMessage = "Number of concurrent folder backup tasks.")][int]$ThrottleLimit = 3,
		[Parameter(Mandatory = $false, HelpMessage = "Number of Robocopy threads per task (/MT).")][int]$RobocopyThreads = 4,
		[Parameter(Mandatory = $false, HelpMessage = "URL for the New-Log function script.")][string]$LogScriptUrl = "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/refs/heads/main/New-Log.ps1"
	)
	$ErrorActionPreference = 'Stop'
	# 1. Load Advanced Logging Function
	if (-not (Get-Command "New-Log" -ErrorAction SilentlyContinue) -and $LogScriptUrl) {
		$LogScriptPath = Join-Path $env:TEMP "New-Log.ps1"
		if (-not (Test-Path $LogScriptPath)) {
			try {
				Invoke-WebRequest -Uri $LogScriptUrl -OutFile $LogScriptPath -UseBasicParsing -MaximumRedirection 1 -ErrorAction Stop
			}
			catch {
				Write-Warning "Failed to download New-Log.ps1. Proceeding with basic fallback logger."
			}
		}
		if (Test-Path $LogScriptPath) {
			. $LogScriptPath
		}
	}
	# 2. Basic Fallback Logger (Activates if the download above failed or LogScriptUrl is not provided)
	if (-not (Get-Command "New-Log" -ErrorAction SilentlyContinue)) {
		function New-Log {
			[CmdletBinding()]
			param(
				[Parameter(ValueFromPipeline, Position = 0)]$Message,
				[Parameter(Position = 1)][string]$Level = "INFO"
			)
			$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
			$color = switch ($Level) {
				"ERROR" { 'Red' }
				"WARNING" { 'Yellow' }
				"SUCCESS" { 'Green' }
				"DEBUG" { 'Blue' }
				"VERBOSE" { 'Cyan' }
				default { 'White' }
			}
			Write-Host "[$timestamp][$Level] $Message" -ForegroundColor $color
		}
	}
	# 3. Build the Task List dynamically
	$copyTasks = foreach ($src in $SourcePaths) {
		# Dynamically removes any drive letter (e.g. C:\ or D:\) so it works for all drives
		$relativePath = $src -replace '^.:\\', ''
		[pscustomobject]@{
			Source      = $src
			Destination = Join-Path -Path $DestinationBase -ChildPath $relativePath
		}
	}
	$copyTasks = @($copyTasks)
	$completed = 0
	$total = $copyTasks.Count
	$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
	New-Log "Starting backup of $($total) tasks" -Level SUCCESS
	# 4. Parallel Execution Pipeline
	$copyTasks | ForEach-Object -Parallel {
		$source = $_.Source
		$dest = $_.Destination
		$localExcludeDirs = $using:ExcludeDirectories
		$localExcludeFiles = $using:ExcludeFiles
		$localThreads = $using:RobocopyThreads
		if (-not (Test-Path -Path $source)) { 
			return [pscustomobject]@{ Source = $source; ExitCode = -1; Status = 'SOURCE_MISSING' } 
		}
		if (-not (Test-Path -Path $dest)) {
			New-Item -Path $dest -ItemType Directory -Force | Out-Null
		}
		$roboArgs = @(
			$source, $dest, 
			"/MIR", "/R:1", "/W:0", "/MT:$localThreads", "/J", 
			"/FFT", "/DST", 
			"/NP", "/NFL", "/NDL", "/NJH", "/NJS"
		)
		# Append Directory Exclusions (/XD)
		if ($localExcludeDirs.Count -gt 0) {
			$roboArgs += "/XD"
			$roboArgs += $localExcludeDirs
		}
		# Append File Exclusions (/XF)
		if ($localExcludeFiles.Count -gt 0) {
			$roboArgs += "/XF"
			$roboArgs += $localExcludeFiles
		}
		# Execute Robocopy and capture output
		$roboOut = & robocopy $roboArgs
		$exitCode = $LASTEXITCODE
		return [pscustomobject]@{ 
			Source   = $source; 
			ExitCode = $exitCode; 
			Status   = if ($exitCode -lt 8) { 'OK' } else { 'FAILED' };
			Output   = $roboOut -join "`n"
		}
	} -ThrottleLimit $ThrottleLimit | ForEach-Object { 
		# 5. Handle Results as they finish
		$completed++
		Write-Progress -Activity "Backup in progress" -Status "[$completed/$total] $($_.Source)" -PercentComplete (($completed / $total) * 100)
		if ($_.Status -eq 'OK') {
			New-Log "[$completed/$total] OK (exit $($_.ExitCode)): $($_.Source)" -Level SUCCESS
		}
		elseif ($_.Status -eq 'SOURCE_MISSING') {
			New-Log "[$completed/$total] Source not found: $($_.Source)" -Level WARNING
		}
		else {
			New-Log "[$completed/$total] FAILED (exit $($_.ExitCode)): $($_.Source) `nDetails: $($_.Output)" -Level ERROR
		}
	}
	$stopwatch.Stop()
	Write-Progress -Activity "Backup in progress" -Completed
	New-Log "Total runtime: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
}