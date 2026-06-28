#Requires -Version 7.0
<#
.SYNOPSIS
    Automated backup script with parallel incremental change detection.
.DESCRIPTION
    Backs up configured source directories and files to a network destination using Robocopy.
    v1.2 improvements over v1.1:
        - PARALLEL change detection across all sources (was sequential) - large speed-up
            when most folders are unchanged and must be fully scanned to prove it.
        - Skips junctions / reparse points in both the scan and Robocopy (/XJ) - avoids
            AppData junction loops and redundant traversal.
        - Additive backup: Robocopy /E instead of /MIR, so deletions are NOT mirrored to
            the destination (files removed from a source are kept on the NAS).
        - Faster retry policy (/R:1 /W:2) - files locked by running apps fail fast instead
            of stalling ~15s each; anything genuinely transient is caught on the next run.
        - Correct change watermark: recorded at run START, not after the backup finishes,
            so files changed mid-backup are no longer missed on the next run.
        - Individual-file change check uses a 2-second tolerance (matches Robocopy /FFT),
            so NAS sub-second timestamp truncation no longer forces a re-copy every run.
        - Exclude-directory matching respects path boundaries (no false prefix matches).
        - Plex registry export wrapped in error handling with suppressed console output.
    Carried over from v1.1:
        - Per-source state file skips unchanged folders entirely.
        - Pre-flight destination check; lock file prevents concurrent runs.
        - Parses and logs Robocopy copy/skip/fail counts per task.
        - Force parameter bypasses all skip logic.
.PARAMETER Force
    Bypass change-detection and back up all sources regardless of state.
.PARAMETER ThrottleLimit
    Number of parallel Robocopy jobs (network-bound). Default 4.
.PARAMETER ScanThrottleLimit
    Number of parallel change-detection scans (local-disk-bound). Default 8.
    Lower to 2-3 if your source files live on a spinning HDD.
.EXAMPLE
    .\Auto-Backup.ps1
.EXAMPLE
    .\Auto-Backup.ps1 -Force
.NOTES
    Version: 1.2
    Requires: PowerShell 7+ (ForEach-Object -Parallel)
#>
[CmdletBinding()]
param(
	[switch]$Force,
	[int]$ThrottleLimit = 4,
	[int]$ScanThrottleLimit = 8,
	[string]$Destination = "\\10.10.10.2\hddnas\AutoBackup"
)
$runStart = [datetime]::UtcNow
# ── 1. Load logging function ────────────────────────────────────────────────
$ErrorActionPreference = 'Stop'
if (-not (Get-Command -Name New-Log -ErrorAction SilentlyContinue)) {
	try {
		$newLogUrl = 'https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1'
		. ([ScriptBlock]::Create((Invoke-WebRequest -Uri $newLogUrl -UseBasicParsing -ErrorAction Stop).Content))
	}
	catch {
		Write-Warning "Could not load New-Log from GitHub: $($_.Exception.Message) -- using built-in fallback."
		function New-Log {
			[CmdletBinding()]
			param(
				[Parameter(ValueFromPipeline, Position = 0)]$Message,
				[Parameter(Position = 1)][string]$Level = 'INFO',
				[Parameter()][string]$LogFilePath,
				[Parameter()]$ErrorObject
			)
			process {
				$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
				$color = switch ($Level) {
					'ERROR' { 'Red' }
					'EXTENDEDERROR' { 'Red' }
					'WARNING' { 'Yellow' }
					'SUCCESS' { 'Green' }
					'DEBUG' { 'Blue' }
					'VERBOSE' { 'Cyan' }
					default { 'White' }
				}
				$suffix = if ($ErrorObject) { " [$($ErrorObject.Exception.Message)]" } else { '' }
				$logLine = "[$timestamp][$Level] $Message$suffix"
				Write-Host $logLine -ForegroundColor $color
				if ($LogFilePath) {
					try {
						Add-Content -Path $LogFilePath -Value $logLine -ErrorAction Stop
					}
					catch {
						Write-Warning "Failed to write to log file '$LogFilePath': $($_.Exception.Message)"
					}
				}
			}
		}
	}
}
$ErrorActionPreference = 'Continue'
# ── 2. Pre-flight: verify destination ────────────────────────────────────────
if (-not (Test-Path -LiteralPath $Destination -PathType Container)) {
	New-Log "Destination $Destination is not available. Aborting backup." -Level ERROR
	return
}
# ── 3. Lock file - prevent concurrent runs ──────────────────────────────────
$lockFilePath = Join-Path -Path $Destination -ChildPath '.backup.lock'
if (Test-Path -LiteralPath $lockFilePath) {
	try {
		$lockInfo = Get-Content -LiteralPath $lockFilePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
		$lockAge = (Get-Date) - [datetime]$lockInfo.Started
		if ($lockAge.TotalHours -lt 2) {
			New-Log "Another backup is running (PID $($lockInfo.PID), started $($lockInfo.Started)). Aborting." -Level WARNING
			return
		}
		New-Log "Stale lock file found ($([int]$lockAge.TotalMinutes) min old). Overriding." -Level WARNING
	}
	catch {
		New-Log "Could not read lock file, overriding. Error: $($_.Exception.Message)" -Level WARNING
	}
}
try {
	@{ Started = (Get-Date).ToString('o'); PID = $PID; Computer = $env:COMPUTERNAME } |
		ConvertTo-Json | Set-Content -LiteralPath $lockFilePath -Force -ErrorAction Stop
}
catch {
	New-Log "Could not create lock file: $($_.Exception.Message)" -Level WARNING
}
# ── 4. State file for change detection ──────────────────────────────────────
$stateFilePath = Join-Path -Path $Destination -ChildPath '.backup-state.json'
$previousState = @{}
if (Test-Path -LiteralPath $stateFilePath) {
	try {
		$raw = Get-Content -LiteralPath $stateFilePath -Raw -ErrorAction Stop
		$loaded = $raw | ConvertFrom-Json -ErrorAction Stop
		foreach ($prop in $loaded.PSObject.Properties) {
			$previousState[$prop.Name] = [datetime]$prop.Value
		}
	}
	catch {
		New-Log "Could not read state file, will back up all sources." -Level WARNING
		$previousState = @{}
	}
}
# ── 5. Define sources ───────────────────────────────────────────────────────
$sources = @(
	"C:\Users\Martin\GitHub-Harze2k"
	"C:\Users\Martin\AppData\Roaming\zen\Profiles\zsisfer9.Default (release)"
	"C:\Users\Martin\AppData\Local\Plex Media Server"
	"C:\Users\Martin\AppData\Local\qBittorrent"
	"C:\Users\Martin\AppData\Local\zen\Profiles\zsisfer9.Default (release)"
	"C:\Users\Martin\AppData\Roaming\Code - Insiders"
	"C:\Users\Martin\AppData\Roaming\SVP4"
	"C:\Users\Martin\AppData\Roaming\Claude"
	"C:\Program Files\Plex\Plex Media Server"
	"C:\Users\Martin\Documents\PowerShell"
	"C:\Toolkit\Toolkit_v13.7\Data"
	"C:\Toolkit\Toolkit_v13.7\Custom"
	"C:\Temp"
	"C:\Program Files\PowerShell\7\Modules"
	"B:\UFM"
	"C:\Program Files\SVP4"
	"C:\Users\Martin\.claude"
	"C:\Users\Martin\Pictures\Screenshots"
)
$files = @(
	"C:\Users\Martin\.claude.json"
	"C:\Users\Martin\AppData\Roaming\Code - Insiders\User\settings.json"
	"C:\Users\Martin\AppData\Local\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
	"C:\Users\Martin\AppData\Roaming\zen\installs.ini"
	"C:\Users\Martin\AppData\Roaming\zen\profiles.ini"
)
# Export registry keys into C:\Temp so they are captured by the C:\Temp source above.
try {
	reg export 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{688e1d8f-188e-49cd-83ca-2669a7e3f8cc}_is1' C:\Temp\plex_backup_hklm.reg /y *> $null
	if ($LASTEXITCODE -ne 0) { New-Log "reg export (HKLM Plex key) returned exit code $LASTEXITCODE." -Level WARNING } else { New-Log "reg export (HKLM Plex key) backed up" -Level SUCCESS }
	reg export 'HKCU\Software\Plex, Inc.' C:\Temp\plex_backup.reg /y *> $null
	if ($LASTEXITCODE -ne 0) { New-Log "reg export (HKCU Plex key) returned exit code $LASTEXITCODE." -Level WARNING } else { New-Log "reg export (HKCU Plex key) backed up" -Level SUCCESS }
}
catch {
	New-Log "Failed to export Plex registry keys: $($_.Exception.Message)" -Level WARNING
}
# ── 6. Exclusions ───────────────────────────────────────────────────────────
$excludeDirs = @(
	"C:\Temp\Sorted-FB"
	"C:\Users\Martin\AppData\Local\Plex Media Server\Cache"
	"C:\Users\Martin\AppData\Local\Plex Media Server\Logs"
	"C:\Users\Martin\AppData\Local\Plex Media Server\Crash Reports"
	"C:\Users\Martin\AppData\Local\Plex Media Server\Updates"
)
$excludeFiles = @(
	"parent.lock"
	"*.log"
	"*.tmp"
	"*.temp"
	"*.cache"
	"*.bak"
	"~*"
	"Thumbs.db"
	"desktop.ini"
)
# ── 7. Change detection ─────────────────────────────────────────────────────
# Returns $true as soon as ANY file newer than $Since is found (early exit).
# Uses .NET EnumerateFiles for memory efficiency - doesn't load all files into memory.
# Skips System + ReparsePoint entries so junctions/symlinks are not followed.
function Test-HasChanges {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string]$Path,
		[Parameter(Mandatory)][datetime]$Since,
		[string[]]$ExcludeDirPaths,
		[string[]]$ExcludeFilePatterns
	)
	try {
		$enumOptions = [System.IO.EnumerationOptions]@{
			RecurseSubdirectories = $true
			IgnoreInaccessible    = $true
			AttributesToSkip      = [System.IO.FileAttributes]::System -bor [System.IO.FileAttributes]::ReparsePoint
		}
		foreach ($file in [System.IO.DirectoryInfo]::new($Path).EnumerateFiles('*', $enumOptions)) {
			$inExcludedDir = $false
			foreach ($exDir in $ExcludeDirPaths) {
				$exDirPrefix = if ($exDir.EndsWith('\')) { $exDir } else { $exDir + '\' }
				if ($file.FullName.StartsWith($exDirPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
					$inExcludedDir = $true
					break
				}
			}
			if ($inExcludedDir) { continue }
			$skip = $false
			foreach ($pattern in $ExcludeFilePatterns) {
				if ($file.Name -like $pattern) { $skip = $true; break }
			}
			if ($skip) { continue }
			if ($file.LastWriteTimeUtc -gt $Since) {
				return $true
			}
		}
		return $false
	}
	catch {
		return $true
	}
}
# ── 7b. Destination path mapping ─────────────────────────────────────────────
# Maps a full source path to its path relative to the backup root.
# C: keeps the existing layout (drive letter stripped); any other drive becomes a
# top-level folder named after the drive letter (e.g. "B:\UFM" -> "B\UFM"), which
# avoids the invalid colon that produced "...\AutoBackup\B:\UFM".
function Get-RelativeBackupPath {
	[CmdletBinding()]
	param([Parameter(Mandatory)][string]$SourcePath)
	if ($SourcePath -match '^[Cc]:\\') {
		return ($SourcePath -replace '^[Cc]:\\', '')
	}
	return ($SourcePath -replace '^([A-Za-z]):\\', '$1\')
}
# ── 8. Build task list with PARALLEL change detection ────────────────────────
# Functions are not inherited by ForEach-Object -Parallel runspaces, so inject
# Test-HasChanges as text and re-create it inside each runspace.
$testHasChangesDef = ${function:Test-HasChanges}.ToString()
$copyTasks = [System.Collections.Generic.List[pscustomobject]]::new()
$skippedCount = 0
if (-not $Force) {
	New-Log "Scanning $($sources.Count) source(s) for changes (parallel, throttle $ScanThrottleLimit)..."
}
$scanResults = $sources | ForEach-Object -Parallel {
	${function:Test-HasChanges} = $using:testHasChangesDef
	$src = $_
	if (-not (Test-Path -LiteralPath $src)) {
		return [pscustomobject]@{ Source = $src; ShouldCopy = $true }
	}
	if ($using:Force) {
		return [pscustomobject]@{ Source = $src; ShouldCopy = $true }
	}
	$lastBackupTime = ($using:previousState)[$src]
	if ($null -eq $lastBackupTime) {
		return [pscustomobject]@{ Source = $src; ShouldCopy = $true }
	}
	$hasChanges = Test-HasChanges -Path $src -Since $lastBackupTime -ExcludeDirPaths $using:excludeDirs -ExcludeFilePatterns $using:excludeFiles
	return [pscustomobject]@{ Source = $src; ShouldCopy = $hasChanges }
} -ThrottleLimit $ScanThrottleLimit
foreach ($result in $scanResults) {
	if ($result.ShouldCopy) {
		$copyTasks.Add([pscustomobject]@{
				Source      = $result.Source
				Destination = Join-Path -Path $Destination -ChildPath (Get-RelativeBackupPath -SourcePath $result.Source)
			})
	}
	else {
		$skippedCount++
	}
}
# ── 9. Execute backup ───────────────────────────────────────────────────────
$completed = 0
$total = $copyTasks.Count
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
New-Log "Starting backup: $total task(s) to process, $skippedCount skipped (unchanged)$(if ($Force) { ' [FORCE mode]' })"
$successfulSources = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
if ($total -gt 0) {
	$copyTasks | ForEach-Object -Parallel {
		$source = $_.Source
		$dest = $_.Destination
		if (-not (Test-Path -LiteralPath $source)) {
			return [pscustomobject]@{
				Source   = $source
				ExitCode = -1
				Status   = 'SOURCE_MISSING'
				Copied   = 0
				Skipped  = 0
				Failed   = 0
				Output   = ''
			}
		}
		if (-not (Test-Path -LiteralPath $dest)) {
			New-Item -Path $dest -ItemType Directory -Force | Out-Null
		}
		# Robocopy arguments
		# /E      recurse incl. empty dirs but do NOT purge (deletions are not mirrored)
		# /XJ     skip junctions/symlinks (avoids AppData loops & redundant copies)
		# /R:1 /W:2  fail fast on locked files instead of stalling
		# /FFT /DST  NAS-friendly timestamp comparison (correctly skips same files)
		# /J      unbuffered I/O (good for large files)
		# /NJH suppresses header, /NJS is NOT used so we can parse the summary stats
		$roboArgs = @(
			$source, $dest,
			"/E", "/XJ", "/R:1", "/W:2", "/MT:4", "/J", "/FFT", "/DST", "/NP", "/NFL", "/NDL", "/NJH", "/BYTES"
		)
		if ($using:excludeDirs.Count -gt 0) {
			$roboArgs += "/XD"
			$roboArgs += $using:excludeDirs
		}
		if ($using:excludeFiles.Count -gt 0) {
			$roboArgs += "/XF"
			$roboArgs += $using:excludeFiles
		}
		$roboOut = & robocopy $roboArgs 2>&1
		$exitCode = $LASTEXITCODE
		$outText = ($roboOut | Out-String)
		$copied = 0; $skippedFiles = 0; $failed = 0
		if ($outText -match 'Files\s*:\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)') {
			$copied = [int]$Matches[2]
			$skippedFiles = [int]$Matches[3]
			$failed = [int]$Matches[5]
		}
		if ($exitCode -lt 8) {
			($using:successfulSources).Add($source)
		}
		return [pscustomobject]@{
			Source   = $source
			ExitCode = $exitCode
			Status   = if ($exitCode -lt 8) { 'OK' } else { 'FAILED' }
			Copied   = $copied
			Skipped  = $skippedFiles
			Failed   = $failed
			Output   = $outText
		}
	} -ThrottleLimit $ThrottleLimit | ForEach-Object {
		$completed++
		Write-Progress -Activity "Backup in progress" -Status "[$completed/$total] $($_.Source)" -PercentComplete (($completed / $total) * 100)
		if ($_.Status -eq 'OK') {
			$detail = "copied=$($_.Copied) skipped=$($_.Skipped)"
			if ($_.Copied -eq 0 -and $_.Failed -eq 0) {
				New-Log "[$completed/$total] OK (no changes): $($_.Source)"
			}
			else {
				New-Log "[$completed/$total] OK ($detail): $($_.Source)"
			}
		}
		elseif ($_.Status -eq 'SOURCE_MISSING') {
			New-Log "[$completed/$total] Source not found: $($_.Source)" -Level WARNING
		}
		else {
			New-Log "[$completed/$total] FAILED (exit $($_.ExitCode), failed=$($_.Failed)): $($_.Source)" -Level ERROR
			if ($_.Output) {
				New-Log "  Robocopy output: $($_.Output.Trim())" -Level ERROR
			}
		}
	}
}
# ── 10. Copy individual files (with change detection) ───────────────────────
if ($files.Count -gt 0) {
	New-Log "Checking $($files.Count) individual file(s)"
	foreach ($filePath in $files) {
		$fileDest = Join-Path -Path $Destination -ChildPath (Get-RelativeBackupPath -SourcePath $filePath)
		$fileDestDir = Split-Path -Path $fileDest -Parent
		if (-not (Test-Path -LiteralPath $filePath)) {
			New-Log "File not found: $filePath" -Level WARNING
			continue
		}
		try {
			$srcFile = Get-Item -LiteralPath $filePath -ErrorAction Stop
			$needsCopy = $true
			if (-not $Force -and (Test-Path -LiteralPath $fileDest)) {
				$destFile = Get-Item -LiteralPath $fileDest -ErrorAction Stop
				$timeDiffSeconds = [Math]::Abs(($srcFile.LastWriteTimeUtc - $destFile.LastWriteTimeUtc).TotalSeconds)
				if ($timeDiffSeconds -le 2 -and $srcFile.Length -eq $destFile.Length) {
					$needsCopy = $false
				}
			}
			if ($needsCopy) {
				if (-not (Test-Path -LiteralPath $fileDestDir)) {
					New-Item -Path $fileDestDir -ItemType Directory -Force | Out-Null
				}
				Copy-Item -LiteralPath $filePath -Destination $fileDest -Force
				New-Log "OK (copied): $filePath"
			}
			else {
				New-Log "OK (unchanged, skipped): $filePath"
			}
		}
		catch {
			New-Log "Failed to copy: $filePath -- $($_.Exception.Message)" -Level ERROR
		}
	}
}
# ── 11. Update state file ───────────────────────────────────────────────────
# Use $runStart (captured before scanning) as the watermark for sources backed up
# this run, so changes made during the backup are still detected next time.
try {
	$newState = @{}
	foreach ($kvp in $previousState.GetEnumerator()) {
		$newState[$kvp.Key] = $kvp.Value
	}
	foreach ($src in $successfulSources) {
		$newState[$src] = $runStart
	}
	$newState | ConvertTo-Json -Depth 1 | Set-Content -LiteralPath $stateFilePath -Encoding UTF8 -Force
}
catch {
	New-Log "Failed to write state file: $($_.Exception.Message)" -Level WARNING
}
# ── 12. Cleanup and summary ─────────────────────────────────────────────────
Remove-Item -LiteralPath $lockFilePath -Force -ErrorAction SilentlyContinue
$stopwatch.Stop()
Write-Progress -Activity "Backup in progress" -Completed
New-Log "Backup complete. Processed: $total, Skipped (unchanged): $skippedCount. Total runtime: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"