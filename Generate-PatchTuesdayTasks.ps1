#Requires -Modules ScheduledTasks
<#
.NOTES
	Author: Harze2k
	Date:   2025-05-15
	Version: 1.2 (Public release.)
		-Cleaned up the functions.
		-Verified it works.
		-Will calculate and created a task that runs a PS file when a Microsoft patch Tuesday has occurred. (Or any other triggers and script for any reason ofc.)
	Sample output from a script using these functions:

	--- Example: Patch Tuesday High Prio Servers ---
	Defining triggers for 'PatchTuesdayHighPrioServers'...
	Successfully created scheduled task 'PatchTuesdayHighPrioServers' at path '\'.

	--- Verifying 'PatchTuesdayHighPrioServers' Triggers (if created) ---
	Date for patch tuesday with custom offset (UTC): 2025-05-16T02:00:00+02:00
	... (more dates) ...
	Date for patch tuesday with custom offset (UTC): 2026-12-11T03:00:00+02:00
#>
function Get-TargetDateForSecondTuesday {
	<#
	.SYNOPSIS
		Calculates the date and time for the second Tuesday of a given month and year, with optional offsets.
	.DESCRIPTION
		This function determines the specific date of the second Tuesday within a given month and year.
		It then applies an optional day offset and an optional hour offset to this calculated date and time.
		The initial calculation for the second Tuesday results in a time of 00:00:00 (midnight).
		The HourOffset is applied to this midnight time.
		The returned DateTime object will have its Kind property set to Unspecified, indicating that its
		time zone is not yet defined and should be interpreted by the calling context.
	.PARAMETER Month
		The target month (an integer from 1 to 12) for which to find the second Tuesday. This parameter is mandatory.
	.PARAMETER Year
		The target year (e.g., 2023) for which to find the second Tuesday. This parameter is mandatory.
	.PARAMETER DayOffset
		An optional integer representing the number of days to add to (positive value) or subtract from (negative value)
		the calculated second Tuesday. Defaults to 0 if not specified.
		For example, a DayOffset of 1 will result in the Wednesday after the second Tuesday.
	.PARAMETER HourOffset
		An optional integer representing the number of hours to add to (positive value) or subtract from (negative value)
		the time of the (potentially day-offsetted) second Tuesday. The base time is midnight (00:00:00) of the
		offsetted day before this HourOffset is applied. Defaults to 0 if not specified.
	.EXAMPLE
		Get-TargetDateForSecondTuesday -Month 10 -Year 2023
		# Returns the date of the second Tuesday of October 2023, at 00:00:00, with Kind=Unspecified.
	.EXAMPLE
		Get-TargetDateForSecondTuesday -Month 10 -Year 2023 -DayOffset 1 -HourOffset 14
		# Returns the date of the Wednesday following the second Tuesday of October 2023, at 14:00:00, with Kind=Unspecified.
	.EXAMPLE
		Get-TargetDateForSecondTuesday -Month 3 -Year 2024 -DayOffset -1 -HourOffset 22
		# Returns the date of the Monday before the second Tuesday of March 2024, at 22:00:00 (10 PM), with Kind=Unspecified.
	.OUTPUTS
		[System.DateTime]
		A DateTime object representing the calculated target date and time. The Kind property of this
		DateTime object is set to [System.DateTimeKind]::Unspecified.
	#>
	[CmdletBinding()]
	[OutputType([System.DateTime])]
	param (
		[Parameter(Mandatory)][ValidateRange(1, 12)][int]$Month,
		[Parameter(Mandatory)][ValidateRange(1900, 2100)][int]$Year,
		[int]$DayOffset = 0,
		[int]$HourOffset = 0
	)
	Write-Verbose "Get-TargetDateForSecondTuesday: Calculating for $Month/$Year, DayOffset: $DayOffset, HourOffset: $HourOffset"
	try {
		$firstDayOfMonth = Get-Date -Year $Year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0 -ErrorAction Stop
	}
	catch {
		Write-Error "Get-TargetDateForSecondTuesday: Invalid Year/Month: $Year/$Month. Error: $($_.Exception.Message)"
		throw
	}
	$daysToAddForFirstTuesday = ([System.DayOfWeek]::Tuesday - $firstDayOfMonth.DayOfWeek + 7) % 7
	$firstTuesdayDate = $firstDayOfMonth.AddDays($daysToAddForFirstTuesday)
	$secondTuesdayBaseDate = $firstTuesdayDate.AddDays(7)
	$finalTargetDate = $secondTuesdayBaseDate.AddDays($DayOffset).AddHours($HourOffset)
	$finalTargetDate = [DateTime]::SpecifyKind($finalTargetDate, [System.DateTimeKind]::Unspecified)
	Write-Verbose "Get-TargetDateForSecondTuesday: Result: $($finalTargetDate.ToString('o')) (Kind: $($finalTargetDate.Kind))"
	return $finalTargetDate
}
function New-MonthlyScheduledTaskTrigger {
	<#
	.SYNOPSIS
		Generates an array of Scheduled Task Trigger objects based on the second Tuesday of the month,
		spanning a specified range of months and years.
	.DESCRIPTION
		This function creates multiple New-ScheduledTaskTrigger CIM objects, each designed for a "once" execution.
		It calculates a specific date and time for each month within a defined period (from TriggerStartMonth/TriggerStartYear
		to TriggerEndMonth/TriggerEndYear). The calculation is based on the second Tuesday of the month,
		to which day and hour offsets (SecondTuesdayDayOffset, SecondTuesdayHourOffset) are applied using
		the Get-TargetDateForSecondTuesday function. A final task execution hour (ScheduledTaskHour) is then added.
		The resulting local time is converted to UTC using the specified TimeZoneId before creating the trigger.
		The loop for months is specific:
		- For the TriggerStartYear, it iterates from TriggerStartMonth up to TriggerEndMonth.
		- For subsequent years (until TriggerEndYear), it iterates from month 1 (January) up to TriggerEndMonth.
	.PARAMETER TriggerStartYear
		The starting year for generating triggers (e.g., 2023). Defaults to the current year if not specified.
	.PARAMETER TriggerStartMonth
		The starting month (an integer from 1 to 12) for generating triggers in the TriggerStartYear. This parameter is mandatory.
	.PARAMETER ExecutionTimeLimitHours
		The maximum time (in hours) the task is allowed to run when launched by these triggers. This parameter is mandatory.
		The value is an integer (e.g., 1 for one hour).
	.PARAMETER SecondTuesdayDayOffset
		An optional integer representing the number of days to offset from the actual second Tuesday of the month.
		This is passed directly to Get-TargetDateForSecondTuesday's DayOffset parameter. Defaults to 0.
	.PARAMETER SecondTuesdayHourOffset
		An optional integer representing the number of hours to offset from midnight of the (potentially day-offsetted)
		second Tuesday. This is passed directly to Get-TargetDateForSecondTuesday's HourOffset parameter. Defaults to 0.
	.PARAMETER ScheduledTaskHour
		The hour of the day (an integer from 0 to 23) when the task should be scheduled to run. This hour is added to
		the date and time returned by Get-TargetDateForSecondTuesday. This parameter is mandatory.
	.PARAMETER TriggerEndYear
		The ending year for generating triggers. This parameter is mandatory.
	.PARAMETER TriggerEndMonth
		The ending month (an integer from 1 to 12) for generating triggers. This month boundary applies to all years
		processed by the function, including the TriggerStartYear and TriggerEndYear. This parameter is mandatory.
	.PARAMETER TimeZoneId
		The ID string of the time zone in which the SecondTuesdayDayOffset, SecondTuesdayHourOffset, and ScheduledTaskHour
		are to be interpreted (e.g., 'W. Europe Standard Time', 'Eastern Standard Time'). The calculated local time
		will be converted from this time zone to UTC. This parameter is mandatory.
	.EXAMPLE
		New-MonthlyScheduledTaskTrigger -TriggerStartMonth 10 -TriggerStartYear 2023 -ExecutionTimeLimitHours 2 -ScheduledTaskHour 20 -TriggerEndYear 2023 -TriggerEndMonth 12 -TimeZoneId 'Central Standard Time'
		# Creates triggers for the 2nd Tuesday of Oct, Nov, Dec 2023, all scheduled at 20:00 Central Time (converted to UTC).
		# Assumes default SecondTuesdayDayOffset=0 and SecondTuesdayHourOffset=0.
	.EXAMPLE
		$triggerParams = @{
			TriggerStartYear        = (Get-Date).Year
			TriggerStartMonth       = (Get-Date).Month
			TriggerEndYear          = (Get-Date).Year + 1
			TriggerEndMonth         = 6 # Only up to June each year
			TimeZoneId              = 'UTC' # Times are already UTC, no conversion needed other than Kind specification
			ExecutionTimeLimitHours = 4
			ScheduledTaskHour       = 10
			SecondTuesdayDayOffset  = 1 # Wednesday after 2nd Tueday
			SecondTuesdayHourOffset = -2 # Two hours before midnight of that Wednesday
		}
		New-MonthlyScheduledTaskTrigger @triggerParams
		# Creates triggers for the Wednesday after the 2nd Tuesday, at 08:00 UTC (10 AM - 2 AM), from current month/year up to June of next year.
	.OUTPUTS
		[Microsoft.Management.Infrastructure.CimInstance[]]
		An array of CIMInstance objects, where each object represents a configured scheduled task trigger
		(specifically, the type returned by New-ScheduledTaskTrigger). Returns an empty array if no triggers
		are generated or if input parameters result in an invalid date range.
	#>
	[CmdletBinding()]
	[OutputType([Microsoft.Management.Infrastructure.CimInstance[]])]
	param (
		[Parameter()][int]$TriggerStartYear = (Get-Date).Year,
		[Parameter(Mandatory)][ValidateRange(1, 12)][int]$TriggerStartMonth = (Get-Date).Month,
		[Parameter(Mandatory)][int]$ExecutionTimeLimitHours = 1,
		[int]$SecondTuesdayDayOffset = 0,
		[int]$SecondTuesdayHourOffset = 0,
		[Parameter(Mandatory)][ValidateRange(0, 23)][int]$ScheduledTaskHour = 20,
		[Parameter(Mandatory)][ValidateRange(1900, 2100)][int]$TriggerEndYear = ((Get-Date).Year + 2),
		[Parameter(Mandatory)][ValidateRange(1, 12)][int]$TriggerEndMonth = 12,
		[Parameter(Mandatory)][string]$TimeZoneId = 'W. Europe Standard Time'
	)
	Write-Verbose "New-MonthlyScheduledTaskTrigger: Generating triggers from $TriggerStartMonth/$TriggerStartYear to $TriggerEndMonth/$TriggerEndYear for TimeZone '$TimeZoneId'."
	Write-Verbose "New-MonthlyScheduledTaskTrigger: Base offsets: Day=$SecondTuesdayDayOffset, Hour=$SecondTuesdayHourOffset. Scheduled Hour: $ScheduledTaskHour."
	try {
		$targetTimeZoneInfo = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
	}
	catch {
		Write-Error "New-MonthlyScheduledTaskTrigger: Invalid TimeZoneId '$TimeZoneId'. Error: $($_.Exception.Message)"
		throw
	}
	$generatedTriggers = [System.Collections.Generic.List[Microsoft.Management.Infrastructure.CimInstance]]::new()
	$currentProcessingYear = $TriggerStartYear
	$currentProcessingStartMonth = $TriggerStartMonth
	if ($TriggerStartYear -gt $TriggerEndYear -or ($TriggerStartYear -eq $TriggerEndYear -and $TriggerStartMonth -gt $TriggerEndMonth) ) {
		Write-Warning "New-MonthlyScheduledTaskTrigger: Start date ($TriggerStartMonth/$TriggerStartYear) is after end date ($TriggerEndMonth/$TriggerEndYear). No triggers generated."
		return @()
	}
	do {
		Write-Verbose "New-MonthlyScheduledTaskTrigger: Processing Year: $currentProcessingYear. Months from $currentProcessingStartMonth to $TriggerEndMonth."
		for ($currentMonthInLoop = $currentProcessingStartMonth; $currentMonthInLoop -le $TriggerEndMonth; $currentMonthInLoop++) {
			if ($currentProcessingYear -eq $TriggerEndYear -and $currentMonthInLoop -gt $TriggerEndMonth) {
				Write-Verbose "New-MonthlyScheduledTaskTrigger: Reached TriggerEndMonth ($TriggerEndMonth) in TriggerEndYear ($TriggerEndYear). Stopping month loop."
				break
			}
			Write-Verbose "New-MonthlyScheduledTaskTrigger: Getting target date for $currentMonthInLoop/$currentProcessingYear."
			try {
				$targetDate = Get-TargetDateForSecondTuesday -Month $currentMonthInLoop -Year $currentProcessingYear `
					-DayOffset $SecondTuesdayDayOffset -HourOffset $SecondTuesdayHourOffset -ErrorAction Stop
			}
			catch {
				Write-Warning "New-MonthlyScheduledTaskTrigger: Could not calculate base date for $currentMonthInLoop/$currentProcessingYear. Skipping. Error: $($_.Exception.Message)"
				continue
			}
			# $targetDate has Kind Unspecified from Get-TargetDateForSecondTuesday
			# AddHours preserves the Kind.
			$taskScheduledLocalTime = $targetDate.AddHours($ScheduledTaskHour)
			Write-Verbose "New-MonthlyScheduledTaskTrigger: Calculated local time for $currentMonthInLoop/${currentProcessingYear}: $($taskScheduledLocalTime.ToString("o")) (Kind: $($taskScheduledLocalTime.Kind)) in specified TimeZone '$TimeZoneId'"
			# For ConvertTimeToUtc with an explicit sourceTimeZoneInfo, the dateTime.Kind must be Unspecified.
			# Ensure it is, as AddHours should have preserved it from $targetDate.
			$taskScheduledLocalTimeForConversion = [DateTime]::SpecifyKind($taskScheduledLocalTime, [System.DateTimeKind]::Unspecified)
			$utcTime = $null
			try {
				$utcTime = [TimeZoneInfo]::ConvertTimeToUtc($taskScheduledLocalTimeForConversion, $targetTimeZoneInfo)
				Write-Verbose "New-MonthlyScheduledTaskTrigger: UTC equivalent: $($utcTime.ToString("o"))"
			}
			catch {
				Write-Error "New-MonthlyScheduledTaskTrigger: Failed to convert '$($taskScheduledLocalTimeForConversion.ToString("o"))' (Kind: $($taskScheduledLocalTimeForConversion.Kind)) to UTC using TimeZone '$($targetTimeZoneInfo.Id)'. Error: $($_.Exception.Message)"
				continue
			}
			try {
				$trigger = New-ScheduledTaskTrigger -Once -At $utcTime -ErrorAction Stop
				$trigger.ExecutionTimeLimit = "PT$($ExecutionTimeLimitHours)H"
				$generatedTriggers.Add($trigger)
			}
			catch {
				Write-Error "New-MonthlyScheduledTaskTrigger: Failed to create CIM trigger for UTC time $($utcTime.ToString("o")). Error: $($_.Exception.Message)"
			}
		}
		$currentProcessingStartMonth = 1
		$currentProcessingYear++
	} while ($currentProcessingYear -le $TriggerEndYear)
	if ($generatedTriggers.Count -eq 0) {
		Write-Warning "New-MonthlyScheduledTaskTrigger: No triggers were generated for the specified period and criteria."
	}
	return $generatedTriggers.ToArray()
}
function Manage-ScheduledTask {
	<#
	.SYNOPSIS
		Creates, updates, or deletes a Windows Scheduled Task.
	.DESCRIPTION
		This function provides a comprehensive way to manage scheduled tasks. It can register a new task with
		specified triggers, actions, and settings. It also supports overwriting an existing task or simply
		deleting an existing task without creating a new one. The function leverages standard PowerShell
		cmdlets for task scheduling and includes error handling and verbose output.
	.PARAMETER TaskName
		The name for the scheduled task (e.g., "MyDailyReport"). Defaults to "CustomTask_" followed by a random
		5-digit number if not specified. While not strictly mandatory due to the default, providing a meaningful name is highly recommended.
	.PARAMETER Triggers
		An array of CIMInstance objects representing task triggers. These are typically generated by functions
		like New-ScheduledTaskTrigger or New-MonthlyScheduledTaskTrigger. This parameter is required if
		-DeleteOnly is not specified and a new task is being created or an existing one overwritten.
	.PARAMETER TaskExecutionTimeLimitHours
		An optional integer specifying the maximum time (in hours) the task is allowed to run. This is set in
		the task's overall settings. Defaults to 1 hour if not specified.
	.PARAMETER ActionExecutable
		An optional string specifying the path to the executable to be run by the task's action
		(e.g., 'powershell.exe', 'pwsh.exe', 'C:\Scripts\MyBatch.bat').
		Defaults to 'pwsh.exe'.
	.PARAMETER ActionArguments
		An optional string containing the arguments to pass to the ActionExecutable.
		Defaults to a sample command launching a PowerShell script from the temp directory.
		It's crucial to customize this for your specific needs.
	.PARAMETER RunAsUser
		An optional string specifying the user account under which the task will run (e.g., 'NT AUTHORITY\SYSTEM',
		'DOMAIN\User', '.\LocalUser'). Defaults to 'NT AUTHORITY\SYSTEM'.
		Ensure this account has "Log on as a batch job" rights.
	.PARAMETER TaskPath
		An optional string specifying the path (folder) in Task Scheduler where the task will be created
		(e.g., '\MyCustomTasks\', '\Microsoft\Windows\MyTasks'). Defaults to the root path '\'.
		The path must exist or be creatable.
	.PARAMETER OverwriteExisting
		A switch parameter. If specified, and a task with the same TaskName already exists at the specified TaskPath,
		the existing task will be stopped, unregistered, and then a new task will be created with the provided settings.
	.PARAMETER DeleteOnly
		A switch parameter. If specified, the function will only attempt to find, stop, and unregister an existing task
		with the given TaskName at the specified TaskPath. No new task will be created. The -Triggers parameter is
		not required if -DeleteOnly is used.
	.EXAMPLE
		$myTriggers = New-MonthlyScheduledTaskTrigger -TriggerStartMonth 1 -TriggerEndMonth 12 -ScheduledTaskHour 2 -TriggerEndYear 2024 -TimeZoneId 'Eastern Standard Time'
		Manage-ScheduledTask -TaskName "MonthlyMaintenance" -Triggers $myTriggers -ActionExecutable "C:\Scripts\Maintenance.ps1" -RunAsUser "NT AUTHORITY\SYSTEM" -OverwriteExisting
		# Creates or overwrites a task named "MonthlyMaintenance" to run C:\Scripts\Maintenance.ps1 as SYSTEM,
		# triggered on the 2nd Tuesday (default offsets) at 2 AM Eastern Time, monthly for 2024.
	.EXAMPLE
		Manage-ScheduledTask -TaskName "OldBackupJob" -TaskPath "\MyCompany\Backups" -DeleteOnly
		# Attempts to delete the scheduled task named "OldBackupJob" located in the "\MyCompany\Backups" folder.
	.EXAMPLE
		$action = New-ScheduledTaskAction -Execute "notepad.exe"
		$trigger = New-ScheduledTaskTrigger -Daily -At "3am"
		Manage-ScheduledTask -TaskName "OpenNotepadDaily" -Triggers $trigger -Action $action -TaskExecutionTimeLimitHours 2
		# Creates a new task to open Notepad daily at 3 AM. Note: This example shows providing pre-created Action and Trigger objects,
		# though the function internally uses New-ScheduledTaskAction if only ActionExecutable/Arguments are given.
		# For simplicity, directly passing triggers is the primary design for the -Triggers parameter.
	.OUTPUTS
		None.
		This function does not return any objects to the pipeline. It writes status messages (hosts, warnings, errors, verbose)
		to the appropriate PowerShell streams.
	#>
	[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
	param (
		[Parameter()][string]$TaskName = "CustomTask_$(Get-Random -Minimum 10000 -Maximum 99999)",
		[Parameter()][Microsoft.Management.Infrastructure.CimInstance[]]$Triggers,
		[int]$TaskExecutionTimeLimitHours = 1,
		[string]$ActionExecutable = 'pwsh.exe',
		[string]$ActionArguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File ""$($env:temp)\Script.ps1""",
		[string]$RunAsUser = 'NT AUTHORITY\SYSTEM',
		[string]$TaskPath = '\',
		[switch]$OverwriteExisting,
		[switch]$DeleteOnly
	)
	if (-not $DeleteOnly -and (-not $Triggers -or $Triggers.Count -eq 0)) {
		Write-Error "Parameter -Triggers is required and must not be empty when not using -DeleteOnly."
		return
	}
	$existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
	if ($DeleteOnly.IsPresent) {
		if ($existingTask) {
			if ($PSCmdlet.ShouldProcess("Task '$TaskName' at path '$TaskPath'", "Stop and Unregister")) {
				try {
					Write-Verbose "Stopping task '$TaskName'..."
					Stop-ScheduledTask -InputObject $existingTask -ErrorAction Stop
					Write-Verbose "Unregistering task '$TaskName'..."
					Unregister-ScheduledTask -InputObject $existingTask -Confirm:$false -ErrorAction Stop
					Write-Host "Successfully deleted task '$TaskName' from path '$TaskPath'." -ForegroundColor Green
				}
				catch {
					Write-Error "Failed to stop and/or remove existing task '$TaskName' from path '$TaskPath'. Error: $($_.Exception.Message)"
				}
			}
		}
		else {
			Write-Warning "Task '$TaskName' not found at path '$TaskPath'. Nothing to delete."
		}
		return
	}
	if ($existingTask) {
		if ($OverwriteExisting.IsPresent) {
			if ($PSCmdlet.ShouldProcess("Existing task '$TaskName' at path '$TaskPath'", "Stop, Unregister (due to -OverwriteExisting), and then Recreate")) {
				try {
					Write-Verbose "Overwriting existing task '$TaskName'. Stopping and unregistering..."
					Stop-ScheduledTask -InputObject $existingTask -ErrorAction SilentlyContinue
					Unregister-ScheduledTask -InputObject $existingTask -Confirm:$false -ErrorAction Stop
					Write-Host "Successfully removed existing task '$TaskName' from path '$TaskPath' for overwrite." -ForegroundColor Yellow
				}
				catch {
					Write-Error "Failed to remove existing task '$TaskName' for overwrite. Error: $($_.Exception.Message)"
					return
				}
			}
			else {
				Write-Warning "Overwrite of task '$TaskName' at path '$TaskPath' cancelled by user."
				return
			}
		}
		else {
			Write-Warning "Task '$TaskName' already exists at path '$TaskPath'. Use -OverwriteExisting to replace it or -DeleteOnly to remove it. Task not created."
			return
		}
	}
	if ($PSCmdlet.ShouldProcess("New task '$TaskName' at path '$TaskPath'", "Register Scheduled Task")) {
		Write-Verbose "Defining new scheduled task '$TaskName'."
		try {
			$taskAction = New-ScheduledTaskAction -Execute $ActionExecutable -Argument $ActionArguments -ErrorAction Stop
			$taskPrincipal = New-ScheduledTaskPrincipal -UserId $RunAsUser -RunLevel Highest -ErrorAction Stop
            # Priority 3 is higher than default 7. 0 is highest, 10 is lowest.
			$taskSettings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours $TaskExecutionTimeLimitHours) -Priority 3 -ErrorAction Stop
			$taskDefinition = New-ScheduledTask -Action $taskAction -Principal $taskPrincipal -Trigger $Triggers -Settings $taskSettings -ErrorAction Stop
			Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -InputObject $taskDefinition -Force -ErrorAction Stop | Out-Null
			Write-Host "Successfully created scheduled task '$TaskName' at path '$TaskPath'." -ForegroundColor Green
		}
		catch {
			Write-Error "Failed during definition or registration of scheduled task '$TaskName'. Error: $($_.Exception.Message)"
		}
	}
	else {
		Write-Warning "Creation of task '$TaskName' at path '$TaskPath' cancelled by user."
	}
}
# --- Example Usage (adapted from your original) ---
<#
Write-Host "`n--- Example: Patch Tuesday High Prio Servers ---" -ForegroundColor Cyan
$commonTriggerSettings = @{
	TriggerStartYear  = (Get-Date).Year
	TriggerStartMonth = (Get-Date).Month
	TriggerEndYear    = 2027
	TriggerEndMonth   = 12
	TimeZoneId        = 'W. Europe Standard Time'
}
$commonTaskSettings = @{
	RunAsUser         = 'Domain\User' #Need to set this, domian\user or device\user
	OverwriteExisting = $true
	TaskPath          = '\'
}
$psFileToRun = "C:\path\to\ps\file"
Write-Host "Defining triggers for 'PatchTuesdayHighPrioServers'..."
$trigHighPrio = New-MonthlyScheduledTaskTrigger @commonTriggerSettings -ExecutionTimeLimitHours 6 -ScheduledTaskHour 23 -SecondTuesdayHourOffset 3 -SecondTuesdayDayOffset 2
if ($trigHighPrio) {
	Manage-ScheduledTask @commonTaskSettings -TaskName 'PatchTuesdayHighPrioServers' -Triggers $trigHighPrio -TaskExecutionTimeLimitHours 6 -ActionExecutable 'pwsh.exe' -ActionArguments "-WindowStyle Hidden -NoProfile -ExecutionPolicy bypass -file ""$psFileToRun"" -Wait"
}
else {
	Write-Warning "No triggers generated for PatchTuesdayHighPrioServers."
}
Write-Host "`n--- Verifying 'PatchTuesdayHighPrioServers' Triggers (if created) ---" -ForegroundColor Cyan
$taskToCheck = Get-ScheduledTask -TaskName 'PatchTuesdayHighPrioServers' -TaskPath $commonTaskSettings.TaskPath -ErrorAction SilentlyContinue
if ($taskToCheck) {
	$taskToCheck.Triggers | ForEach-Object {
		Write-Host "Date for patch tuesday with custom offset (UTC): $($_.StartBoundary)"
	}
}
else {
	Write-Warning "Task 'PatchTuesdayHighPrioServers' not found at path '$($commonTaskSettings.TaskPath)'."
}
#>