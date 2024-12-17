function Get-DateForSecondTuesday {
    param (
        [ValidateRange(1, 12)]
        [int]$Month = 10,
        [int]$Year = 2023,
        [int]$OffsetDays = 0,
        [int]$OffsetHours = 0
    )
    $firstDayOfMonth = Get-Date -Year $year -Month $month -Day 1
    $firstTuesday = [enum]::GetValues([System.DayOfWeek]) | Where-Object { $_ -eq 'Tuesday' } | ForEach-Object { [DateTime]::new($year, $month, (1..7)[($_ - $firstDayOfMonth.DayOfWeek + 7) % 7]) }
    $secondTuesday = $firstTuesday.AddDays(7 + $offsetDays).AddHours($offsetHours)
    return $secondTuesday
}
function New-CustomTrigger {
    param (
        [int]$StartYear = 2023,
        [ValidateRange(1, 12)]
        [int]$StartMonth = 10,
        [string]$ExecutionTimeLimit = '1',
        [int]$OffsetDays = 0,
        [int]$OffsetHours = 0,
        [ValidateRange(0, 23)]
        [int]$StartTime = 20,
        [ValidateRange(2023, 2034)]
        [int]$EndYear = 2026,
        [ValidateRange(1, 12)]
        [int]$EndMonth = 12,
        [string]$TimeZone = 'W. Europe Standard Time'
    )
    $tmZone = [TimeZoneInfo]::FindSystemTimeZoneById("$($timeZone)")
    $triggers = @()
    [int]$year = 0
    $year = $startYear
    do {
        for ($month = $startMonth; $month -le $endMonth; $month++) {
            $secondTuesday = Get-DateForSecondTuesday -month $month -Year $year -OffsetDays $offsetDays -OffsetHours $offsetHours
            $secondTuesday = $secondTuesday.AddHours($startTime)  # Set the time to 20:00
            $utcTime = [TimeZoneInfo]::ConvertTimeToUtc($secondTuesday, $tmZone)
            try {
                $trigger = New-ScheduledTaskTrigger -Once -At $utcTime -ErrorAction Stop
                $trigger.ExecutionTimeLimit = 'PT' + $($executionTimeLimit) + 'H'
                $triggers += $trigger
            }
            catch {
                Write-Error "Failed to create trigger for $utcTime. $_"
                Break
            }
        }
        $startMonth = 1
        $year ++
    } while ($year -le $endYear)
    return $triggers
}
function New-CustomTask {
    param (
        [string]$TaskName = "Task_$(Get-Random -Count 1)",
        [Microsoft.Management.Infrastructure.CimInstance[]]$Triggers,
        [string]$ExecutionTimeLimit = '1',
        [string]$Executable = 'powershell.exe',
        [string]$Arguments = '-File ',
        [string]$ScriptRunner = 'internal\SVC_MEMCM_CP',
        [Switch]$EndAndRemove,
        [switch]$NoCreateJustDelete
    )
    if ($NoCreateJustDelete.IsPresent) {
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            try {
                Stop-ScheduledTask -InputObject $existingTask -ErrorAction Stop
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
                Write-Host "Successfully deleted task with name: $taskname." -ForegroundColor Green
                return
            }
            catch {
                Write-Error "Failed to stop and remove existing task with name: $taskName. $_"
                return
            }
        }
        else {
            Write-Error "Found no existing task with name: $taskname."
            return
        }
    }
    if ($EndAndRemove.IsPresent) {
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            try {
                Stop-ScheduledTask -InputObject $existingTask -ErrorAction Stop
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
                Write-Host "Successfully deleted task with name: $taskname." -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to stop and remove existing task with name: $taskName. $_"
                return
            }
        }
        else {
            Write-Warning "Found no existing task with name: $taskname."
        }
    }
    else {
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Write-Warning "Found existing task with name $taskname."
            Write-Warning 'Need to use the parameter -EndAndRemove to override. Task not created.'
            return
        }
    }
    $action = New-ScheduledTaskAction -Execute $executable -Argument $arguments
    $principal = New-ScheduledTaskPrincipal -UserId $scriptRunner -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours $ExecutionTimeLimit) -Priority 1
    $task = New-ScheduledTask -Action $action -Principal $principal -Trigger $Triggers -Settings $settings
    try {
        Register-ScheduledTask -TaskName $TaskName -TaskPath '\' -InputObject $task -Force -ErrorAction Stop | Out-Null
        Write-Host "Successfully created task with name: $taskname." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to register task $TaskName. $_"
        Break
    }
}
$trigHighPrio = New-CustomTrigger -ExecutionTimeLimit 6 -StartYear (Get-Date).Year -StartMonth (Get-Date).Month -EndYear 2026 -EndMonth 12 -StartTime 23 -OffsetHours 3 -OffsetDays 2
$trigRestartCheckHighPrioServers = New-CustomTrigger -ExecutionTimeLimit 2 -StartYear (Get-Date).Year -StartMonth (Get-Date).Month -EndYear 2026 -EndMonth 12 -StartTime 23 -OffsetHours 7 -OffsetDays 2
$trigAnalyzeHighPrio = New-CustomTrigger -ExecutionTimeLimit 2 -StartYear (Get-Date).Year -StartMonth (Get-Date).Month -EndYear 2026 -EndMonth 12 -OffsetDays 3 -OffsetHours -8
New-CustomTask -TaskName 'PatchTuesdayHighPrioServers' -Triggers $trigHighPrio -ExecutionTimeLimit 6 -EndAndRemove -ScriptRunner 'internal\SVC_MEMCM_CP' -Executable 'pwsh.exe' -Arguments '-WindowStyle Hidden -NoProfile -ExecutionPolicy bypass -file "E:\Program Files\WindowsUpdateForServers\Scripts\RunWUonServersandReportBack_v1.4_HighPrioServers.ps1" -Wait'
New-CustomTask -TaskName 'RestartCheckHighPrioServers' -Triggers $trigRestartCheckHighPrioServers -ExecutionTimeLimit 2 -EndAndRemove -ScriptRunner 'internal\SVC_MEMCM_CP' -Executable 'pwsh.exe' -Arguments '-WindowStyle Hidden -NoProfile -ExecutionPolicy bypass -file "E:\Program Files\WindowsUpdateForServers\Scripts\RestartCheck_HighPrioServers_v1.0.ps1" -Wait'
New-CustomTask -TaskName 'AnalyzeWUResultHighPrioServers' -Triggers $trigAnalyzeHighPrio -ExecutionTimeLimit 3 -EndAndRemove -ScriptRunner 'internal\SVC_MEMCM_CP' -Executable 'pwsh.exe' -Arguments '-WindowStyle Hidden -NoProfile -ExecutionPolicy bypass -file "E:\Program Files\WindowsUpdateForServers\Scripts\WindowsUpdateReport_v1.4_HighPrioServers.ps1" -Wait'
$trigHighPrio | ForEach-Object {
    Write-Host "Date for patch tuseday with custom offset: $($_.StartBoundary)"
}