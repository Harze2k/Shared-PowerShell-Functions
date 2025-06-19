function Test-DNSServers {
	<#
	.SYNOPSIS
		Tests and compares DNS server performance for resolving specified websites.
	.DESCRIPTION
		The Test-DNSServers function is a PowerShell tool designed to evaluate and compare the performance
		of multiple DNS servers in resolving specified websites. It conducts a series of DNS resolution tests
		for each combination of website and DNS server, calculates response time statistics, and provides
		a comprehensive performance analysis.
		This function is particularly useful for:
		- Network administrators optimizing DNS configurations
		- Developers troubleshooting DNS-related performance issues
		- Users seeking to identify the fastest DNS servers for their frequently visited websites
		- Benchmarking DNS performance across different providers
	.PARAMETER Websites
		An array of websites to test. Each website will be tested against all specified DNS servers.
		Example: @('google.com', 'microsoft.com', 'github.com')
	.PARAMETER TestsCount
		The number of DNS resolution tests to conduct for each combination of website and DNS server.
		Higher values provide more accurate results but take longer to complete.
		Default: 10
	.PARAMETER DNSServers
		A hashtable of DNS servers and their respective IP addresses to test.
		Default includes popular DNS providers: Google, Cloudflare, Quad9, and CleanBrowsing.
	.PARAMETER RecordType
		The DNS record type to query for each test.
		Valid values: A, AAAA, ANY, CNAME, MX, NS, SOA, SRV, TXT
		Default: A
	.PARAMETER Timeout
		The timeout in milliseconds for each DNS resolution test.
		Default: 2000 (2 seconds)
	.PARAMETER PingTimeout
		The timeout in milliseconds for ping operations when IncludePingTest is enabled.
		Default: 50 (0.05 seconds)
	.PARAMETER Parallel
		When specified, runs tests in parallel to improve overall execution time.
		Note: This may affect measurement accuracy on systems with limited resources.
	.PARAMETER ParallelThrottleLimit
		Controls the maximum number of concurrent DNS tests when running in parallel mode.
		Lower values may improve accuracy but increase execution time.
		Higher values may decrease accuracy but improve execution time.
		Valid range: 1-20
		Default: 8
	.PARAMETER IncludePingTest
		When specified, the function will also measure round-trip ping time to the resolved IP addresses.
		This provides additional latency metrics beyond just DNS resolution time.
	.PARAMETER IncludeSystemDNS
		When specified, includes the system's current DNS servers in the comparison.
	.PARAMETER OutputFormat
		Specifies the format for outputting detailed results.
		Valid values: Console, CSV, JSON, GridView
		Default: Console
	.PARAMETER OutputPath
		The file path for saving the test results when using CSV or JSON output formats.
		If not specified, results will only be displayed and not saved.
	.PARAMETER ShowProgress
		When specified, displays a progress bar during the tests.
	.PARAMETER ReturnResults
		When specified, returns results object.
	.EXAMPLE
		Test-DNSServers -Websites @('reddit.com', 'youtube.com', 'github.com') -TestsCount 20
		Performs 20 DNS resolution tests for each combination of the specified websites and default DNS servers.
	.EXAMPLE
		Test-DNSServers -Websites @('microsoft.com') -RecordType MX -IncludeSystemDNS -OutputFormat GridView
		Tests MX record resolution for microsoft.com using default DNS servers plus the system's current DNS servers,
		and displays results in a sortable GridView.
	.EXAMPLE
		Test-DNSServers -Websites @('google.com', 'amazon.com') -Parallel -OutputFormat CSV -OutputPath "C:\Temp\DNSResults.csv"
		Tests resolution for google.com and amazon.com using parallel execution for speed,
		and saves detailed results to a CSV file.
	.OUTPUTS
		Displays DNS server performance statistics on the console and optionally exports detailed results
		to a file based on the specified OutputFormat.
	.NOTES
		Author: Harze2k
		Date:   2025-05-15
		Version: 1.4 (Initial release.)

		Change Log:
		- 1.4: Initial release
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0)][ValidateNotNullOrEmpty()][string[]]$Websites = @('google.com', 'microsoft.com', 'github.com'),
		[Parameter(Position = 1)][ValidateRange(1, 1000)][int]$TestsCount = 10,
		[Parameter(Position = 2)]
		[hashtable]$DNSServers = @{
			'Google'         = '8.8.8.8'
			'Google2'        = '8.8.4.4'
			'Cloudflare'     = '1.1.1.1'
			'Cloudflare2'    = '1.0.0.1'
			'Quad9'          = '9.9.9.9'
			'Quad92'         = '149.112.112.112'
			'CleanBrowsing'  = '185.228.168.9'
			'CleanBrowsing2' = '185.228.169.9'
		},
		[Parameter()][ValidateSet('A', 'AAAA', 'ANY', 'CNAME', 'MX', 'NS', 'SOA', 'SRV', 'TXT')][string]$RecordType = 'A',
		[Parameter()][ValidateRange(20, 10000)][int]$Timeout = 50,
		[Parameter()][ValidateRange(50, 5000)][int]$PingTimeout = 50,
		[Parameter()][switch]$Parallel,
		[Parameter()][ValidateRange(1, 64)][int]$ParallelThrottleLimit = 8,
		[Parameter()][switch]$IncludePingTest,
		[Parameter()][switch]$IncludeSystemDNS,
		[Parameter()][ValidateSet('Console', 'CSV', 'JSON', 'GridView')][string]$OutputFormat = 'Console',
		[Parameter()][string]$OutputPath,
		[Parameter()][switch]$ShowProgress,
		[Parameter()][switch]$ReturnResults
	)
	# Function initialization and diagnostics
	$startTime = Get-Date
	$functionVersion = "1.0.0"
	Write-Host "==================================================================" -ForegroundColor Cyan
	Write-Host "Test-DNSServers v$functionVersion - DNS Performance Testing Tool" -ForegroundColor Cyan
	Write-Host "Started at: $startTime" -ForegroundColor Cyan
	Write-Host "==================================================================" -ForegroundColor Cyan
	# Check if Resolve-DnsName is available
	$resolveDnsNameAvailable = $null -ne (Get-Command -Name Resolve-DnsName -ErrorAction SilentlyContinue)
	if (-not $resolveDnsNameAvailable) {
		Write-Warning "Resolve-DnsName cmdlet not available. This function works best with Windows PowerShell 5.1+ or PowerShell 7+ on Windows."
		Write-Host "Will attempt to use alternative DNS resolution methods." -ForegroundColor Yellow
	}
	# Validate OutputPath when using file output formats
	if (($OutputFormat -in @('CSV', 'JSON')) -and (-not $OutputPath)) {
		Write-Warning "OutputPath parameter is required when using CSV or JSON output formats. Defaulting to Console output."
		$OutputFormat = 'Console'
	}
	# Display overview of testing parameters
	Write-Host "Testing Parameters:" -ForegroundColor Yellow
	Write-Host "- Websites to test: $($Websites.Count)" -ForegroundColor White
	Write-Host "- DNS servers to test: $($DNSServers.Count)" -ForegroundColor White
	Write-Host "- Tests per combination: $TestsCount" -ForegroundColor White
	Write-Host "- Record type: $RecordType" -ForegroundColor White
	Write-Host "- Execution mode: $(if ($Parallel) { 'Parallel' } else { 'Sequential' })" -ForegroundColor White
	if ($Parallel) {
		Write-Host "  - Parallel throttle limit: $ParallelThrottleLimit" -ForegroundColor White
	}
	if ($IncludePingTest) {
		Write-Host "- Ping round-trip testing: Enabled (Timeout: $PingTimeout ms)" -ForegroundColor White
	}
	# Important note about parallel execution and timing
	if ($Parallel) {
		Write-Host "`nIMPORTANT NOTE:" -ForegroundColor Yellow
		Write-Host "Parallel execution typically shows higher response times than sequential execution" -ForegroundColor Yellow
		Write-Host "due to resource contention, network stack behavior, and thread management overhead." -ForegroundColor Yellow
		Write-Host "For the most accurate timing measurements, use sequential mode." -ForegroundColor Yellow
		Write-Host "Parallel mode is best used for faster overall execution when exact timing is not critical." -ForegroundColor Yellow
	}
	Write-Host ""
	# Add system DNS if requested
	if ($IncludeSystemDNS) {
		$systemDNSServers = Get-DnsClientServerAddress |
			Where-Object { $_.AddressFamily -eq 2 -and $_.ServerAddresses } |
			Select-Object -ExpandProperty ServerAddresses -Unique
		if ($systemDNSServers) {
			$i = 1
			foreach ($dns in $systemDNSServers) {
				$DNSServers["SystemDNS$i"] = $dns
				$i++
			}
			Write-Host "Added system DNS servers: $($systemDNSServers -join ', ')" -ForegroundColor Yellow
			Write-Host ""
		}
		else {
			Write-Warning "No system DNS servers found."
		}
	}
	# Initialize results collection
	$globalResults = @()
	$overallResults = @{}
	$totalTests = $Websites.Count * $DNSServers.Count
	$currentTest = 0
	# Test each website
	foreach ($website in $Websites) {
		Write-Host "Testing website: $website" -ForegroundColor Cyan
		# Test each DNS server for the current website
		$dnsResults = @()
		$dnsTestJobs = @()
		# Setup for parallel or sequential testing
		if ($Parallel -and ($PSVersionTable.PSVersion.Major -ge 7)) {
			Write-Host "Running tests in parallel mode..." -ForegroundColor Yellow
			$dnsTestScriptBlock = {
				param($Website, $RecordType, $DNSProvider, $DNSServer, $TestsCount, $Timeout, $IncludePingTest, $PingTimeout)
				$results = @()
				$resolvedCount = 0
				$failedCount = 0
				$responseTimeTotal = 0
				$pingTimeTotal = 0
				$pingCount = 0
				$responseTimesArray = @()
				$pingTimesArray = @()
				for ($i = 1; $i -le $TestsCount; $i++) {
					try {
						# Measure individual test time
						$sw = [System.Diagnostics.Stopwatch]::StartNew()
						$resolved = Resolve-DnsName -Name $Website -Type $RecordType -Server $DNSServer -ErrorAction Stop
						$sw.Stop()
						$resolvedCount++
						$responseTime = $sw.ElapsedMilliseconds
						$responseTimeTotal += $responseTime
						$responseTimesArray += $responseTime
						# Perform ping test if requested
						$pingTime = $null
						if ($IncludePingTest) {
							$ipAddress = $null
							# Get the IP address based on the record type
							if ($RecordType -eq 'A' -or $RecordType -eq 'AAAA') {
								$ipAddress = ($resolved | Where-Object { $_.Type -eq $RecordType } | Select-Object -First 1).IPAddress
							}
							# If we have a CNAME or other record type, try to get the A record
							elseif ($resolved | Where-Object { $_.Type -eq 'A' }) {
								$ipAddress = ($resolved | Where-Object { $_.Type -eq 'A' } | Select-Object -First 1).IPAddress
							}
							if ($ipAddress) {
								try {
									$ping = New-Object System.Net.NetworkInformation.Ping
									# Use a shorter timeout for ping to prevent hanging
									$pingTimeout = [Math]::Min($Timeout, 500) # Maximum 500ms timeout for ping
									$pingResult = $ping.Send($ipAddress, $pingTimeout)
									if ($pingResult.Status -eq 'Success') {
										$pingTime = $pingResult.RoundtripTime
										$pingTimeTotal += $pingTime
										$pingTimesArray += $pingTime
										$pingCount++
									}
								}
								catch {
									# Just continue if ping fails
									Write-Verbose "Ping failed for ${ipAddress}: $($_.Exception.Message)"
								}
							}
						}
						$results += [PSCustomObject]@{
							Website      = $Website
							DNSProvider  = $DNSProvider
							TestNumber   = $i
							ResponseTime = $responseTime
							PingTime     = $pingTime
							Status       = "Success"
							IPsResolved  = ($resolved | Where-Object { $_.Type -eq $RecordType } | Measure-Object).Count
						}
					}
					catch {
						$failedCount++
						$results += [PSCustomObject]@{
							Website      = $Website
							DNSProvider  = $DNSProvider
							TestNumber   = $i
							ResponseTime = $Timeout
							PingTime     = $null
							Status       = "Failed: $($_.Exception.Message)"
							IPsResolved  = 0
						}
					}
				}
				# Calculate statistics
				$successRate = if ($TestsCount -gt 0) { ($resolvedCount / $TestsCount) * 100 } else { 0 }
				$averageResponseTime = if ($resolvedCount -gt 0) { $responseTimeTotal / $resolvedCount } else { $Timeout }
				# Calculate additional statistics if we have successful resolutions
				$minResponseTime = if ($responseTimesArray.Count -gt 0) { ($responseTimesArray | Measure-Object -Minimum).Minimum } else { $Timeout }
				$maxResponseTime = if ($responseTimesArray.Count -gt 0) { ($responseTimesArray | Measure-Object -Maximum).Maximum } else { $Timeout }
				# Calculate median
				$medianResponseTime = if ($responseTimesArray.Count -gt 0) {
					$sorted = $responseTimesArray | Sort-Object
					if ($sorted.Count % 2 -eq 0) {
						($sorted[($sorted.Count / 2) - 1] + $sorted[$sorted.Count / 2]) / 2
					}
					else {
						$sorted[($sorted.Count - 1) / 2]
					}
				}
				else { $Timeout }
				# Calculate standard deviation
				$stdDevResponseTime = if ($responseTimesArray.Count -gt 1) {
					$avg = $responseTimeTotal / $responseTimesArray.Count
					[Math]::Sqrt(($responseTimesArray | ForEach-Object { [Math]::Pow(($_ - $avg), 2) } | Measure-Object -Average).Average)
				}
				else { 0 }
				# Create the result object
				$result = [PSCustomObject]@{
					Website             = $Website
					DNSProvider         = $DNSProvider
					DNSServer           = $DNSServer
					TestsRun            = $TestsCount
					SuccessCount        = $resolvedCount
					FailedCount         = $failedCount
					SuccessRate         = $successRate
					AverageResponseTime = $averageResponseTime
					MedianResponseTime  = $medianResponseTime
					MinResponseTime     = $minResponseTime
					MaxResponseTime     = $maxResponseTime
					StdDevResponseTime  = $stdDevResponseTime
					DetailedResults     = $results
				}
				# Add ping statistics if ping tests were performed
				if ($IncludePingTest) {
					if ($pingCount -gt 0) {
						$result | Add-Member -NotePropertyName AveragePingTime -NotePropertyValue ($pingTimeTotal / $pingCount)
						$result | Add-Member -NotePropertyName MinPingTime -NotePropertyValue ($pingTimesArray | Measure-Object -Minimum).Minimum
						$result | Add-Member -NotePropertyName MaxPingTime -NotePropertyValue ($pingTimesArray | Measure-Object -Maximum).Maximum
					}
					else {
						$result | Add-Member -NotePropertyName AveragePingTime -NotePropertyValue $null
						$result | Add-Member -NotePropertyName MinPingTime -NotePropertyValue $null
						$result | Add-Member -NotePropertyName MaxPingTime -NotePropertyValue $null
					}
				}
				return $result
			}
			# Submit parallel jobs
			foreach ($dns in $DNSServers.GetEnumerator()) {
				$dnsTestJobs += Start-ThreadJob -ScriptBlock $dnsTestScriptBlock -ArgumentList $website, $RecordType, $dns.Key, $dns.Value, $TestsCount, $Timeout, $IncludePingTest, $PingTimeout -ThrottleLimit $ParallelThrottleLimit
			}
			# Wait for all jobs to complete
			if ($ShowProgress) {
				$jobsComplete = 0
				$totalJobs = $dnsTestJobs.Count
				while ($jobsComplete -lt $totalJobs) {
					$jobsComplete = ($dnsTestJobs | Where-Object { $_.State -eq 'Completed' }).Count
					$percentComplete = ($jobsComplete / $totalJobs) * 100
					Write-Progress -Activity "Testing DNS resolution for $website" -Status "$jobsComplete of $totalJobs DNS servers tested" -PercentComplete $percentComplete
					Start-Sleep -Milliseconds 100
				}
				Write-Progress -Activity "Testing DNS resolution for $website" -Completed
			}
			else {
				$null = $dnsTestJobs | Wait-Job
			}
			# Process job results
			foreach ($job in $dnsTestJobs) {
				$jobResult = Receive-Job -Job $job -Wait -AutoRemoveJob
				$dnsResults += $jobResult
				$globalResults += $jobResult
				# Update overall results
				$dnsProvider = $jobResult.DNSProvider
				if (-not $overallResults.ContainsKey($dnsProvider)) {
					$overallResults[$dnsProvider] = @{
						TotalResponseTime = 0
						WebsitesTested    = 0
						SuccessRate       = 0
						TotalSuccessRate  = 0
						TotalPingTime     = 0
						PingTestsCount    = 0
					}
				}
				$overallResults[$dnsProvider].TotalResponseTime += $jobResult.AverageResponseTime
				$overallResults[$dnsProvider].TotalSuccessRate += $jobResult.SuccessRate
				$overallResults[$dnsProvider].WebsitesTested++
				# Add ping data if available
				if ($IncludePingTest -and $jobResult.AveragePingTime) {
					$overallResults[$dnsProvider].TotalPingTime += $jobResult.AveragePingTime
					$overallResults[$dnsProvider].PingTestsCount++
				}
			}
		}
		else {
			# Sequential testing
			foreach ($dns in $DNSServers.GetEnumerator()) {
				$currentTest++
				if ($ShowProgress) {
					$percentComplete = ($currentTest / $totalTests) * 100
					Write-Progress -Activity "Testing DNS Servers" -Status "Testing $website with $($dns.Key)" -PercentComplete $percentComplete
				}
				Write-Verbose "Testing $($dns.Key) ($($dns.Value)) for $website..."
				# Initialize metrics
				$responseTimesArray = @()
				$pingTimesArray = @()
				$resolvedCount = 0
				$failedCount = 0
				$pingCount = 0
				$detailedResults = @()
				# Run the tests
				for ($i = 1; $i -le $TestsCount; $i++) {
					try {
						# Measure individual test time
						$sw = [System.Diagnostics.Stopwatch]::StartNew()
						$resolved = Resolve-DnsName -Name $website -Type $RecordType -Server $dns.Value -ErrorAction Stop
						$sw.Stop()
						$resolvedCount++
						$responseTime = $sw.ElapsedMilliseconds
						$responseTimesArray += $responseTime
						# Perform ping test if requested
						$pingTime = $null
						if ($IncludePingTest) {
							$ipAddress = $null
							# Get the IP address based on the record type
							if ($RecordType -eq 'A' -or $RecordType -eq 'AAAA') {
								$ipAddress = ($resolved | Where-Object { $_.Type -eq $RecordType } | Select-Object -First 1).IPAddress
							}
							# If we have a CNAME or other record type, try to get the A record
							elseif ($resolved | Where-Object { $_.Type -eq 'A' }) {
								$ipAddress = ($resolved | Where-Object { $_.Type -eq 'A' } | Select-Object -First 1).IPAddress
							}
							if ($ipAddress) {
								try {
									$ping = New-Object System.Net.NetworkInformation.Ping
									$pingResult = $ping.Send($ipAddress, 1000)
									if ($pingResult.Status -eq 'Success') {
										$pingTime = $pingResult.RoundtripTime
										$pingTimesArray += $pingTime
										$pingCount++
									}
								}
								catch {
									# Just continue if ping fails
									Write-Verbose "Ping failed for ${ipAddress}: $($_.Exception.Message)"
								}
							}
						}
						$detailedResults += [PSCustomObject]@{
							Website      = $website
							DNSProvider  = $dns.Key
							TestNumber   = $i
							ResponseTime = $responseTime
							PingTime     = $pingTime
							Status       = "Success"
							IPsResolved  = ($resolved | Where-Object { $_.Type -eq $RecordType } | Measure-Object).Count
						}
					}
					catch {
						$failedCount++
						$detailedResults += [PSCustomObject]@{
							Website      = $website
							DNSProvider  = $dns.Key
							TestNumber   = $i
							ResponseTime = $Timeout
							PingTime     = $null
							Status       = "Failed: $($_.Exception.Message)"
							IPsResolved  = 0
						}
					}
				}
				# Calculate statistics
				$successRate = if ($TestsCount -gt 0) { ($resolvedCount / $TestsCount) * 100 } else { 0 }
				$averageResponseTime = if ($resolvedCount -gt 0) { ($responseTimesArray | Measure-Object -Average).Average } else { $Timeout }
				$minResponseTime = if ($responseTimesArray.Count -gt 0) { ($responseTimesArray | Measure-Object -Minimum).Minimum } else { $Timeout }
				$maxResponseTime = if ($responseTimesArray.Count -gt 0) { ($responseTimesArray | Measure-Object -Maximum).Maximum } else { $Timeout }
				# Calculate median
				$medianResponseTime = if ($responseTimesArray.Count -gt 0) {
					$sorted = $responseTimesArray | Sort-Object
					if ($sorted.Count % 2 -eq 0) {
						($sorted[($sorted.Count / 2) - 1] + $sorted[$sorted.Count / 2]) / 2
					}
					else {
						$sorted[($sorted.Count - 1) / 2]
					}
				}
				else { $Timeout }
				# Calculate standard deviation
				$stdDevResponseTime = if ($responseTimesArray.Count -gt 1) {
					$avg = ($responseTimesArray | Measure-Object -Average).Average
					[Math]::Sqrt(($responseTimesArray | ForEach-Object { [Math]::Pow(($_ - $avg), 2) } | Measure-Object -Average).Average)
				}
				else { 0 }
				# Create result object
				$result = [PSCustomObject]@{
					Website             = $website
					DNSProvider         = $dns.Key
					DNSServer           = $dns.Value
					TestsRun            = $TestsCount
					SuccessCount        = $resolvedCount
					FailedCount         = $failedCount
					SuccessRate         = $successRate
					AverageResponseTime = $averageResponseTime
					MedianResponseTime  = $medianResponseTime
					MinResponseTime     = $minResponseTime
					MaxResponseTime     = $maxResponseTime
					StdDevResponseTime  = $stdDevResponseTime
					DetailedResults     = $detailedResults
				}
				# Add ping statistics if ping tests were performed
				if ($IncludePingTest) {
					if ($pingCount -gt 0) {
						$result | Add-Member -NotePropertyName AveragePingTime -NotePropertyValue ($pingTimesArray | Measure-Object -Average).Average
						$result | Add-Member -NotePropertyName MinPingTime -NotePropertyValue ($pingTimesArray | Measure-Object -Minimum).Minimum
						$result | Add-Member -NotePropertyName MaxPingTime -NotePropertyValue ($pingTimesArray | Measure-Object -Maximum).Maximum
					}
					else {
						$result | Add-Member -NotePropertyName AveragePingTime -NotePropertyValue $null
						$result | Add-Member -NotePropertyName MinPingTime -NotePropertyValue $null
						$result | Add-Member -NotePropertyName MaxPingTime -NotePropertyValue $null
					}
				}
				$dnsResults += $result
				$globalResults += $result
				# Update overall results
				if (-not $overallResults.ContainsKey($dns.Key)) {
					$overallResults[$dns.Key] = @{
						TotalResponseTime = 0
						WebsitesTested    = 0
						SuccessRate       = 0
						TotalSuccessRate  = 0
						TotalPingTime     = 0
						PingTestsCount    = 0
					}
				}
				$overallResults[$dns.Key].TotalResponseTime += $averageResponseTime
				$overallResults[$dns.Key].TotalSuccessRate += $successRate
				$overallResults[$dns.Key].WebsitesTested++
				# Add ping data if available
				if ($IncludePingTest -and $pingCount -gt 0) {
					$overallResults[$dns.Key].TotalPingTime += ($pingTimesArray | Measure-Object -Average).Average
					$overallResults[$dns.Key].PingTestsCount++
				}
			}
		}
		if ($ShowProgress) {
			Write-Progress -Activity "Testing DNS Servers" -Completed
		}
		# Display results for this website
		Write-Host "`nResults for $website ($RecordType record):" -ForegroundColor Green
		if ($IncludePingTest) {
			$sortedResults = $dnsResults | Sort-Object AverageResponseTime |
				Select-Object DNSProvider, DNSServer, SuccessRate,
				@{Name = "AverageResponseTime"; Expression = { "{0:N2}" -f $_.AverageResponseTime } },
				@{Name = "MedianResponseTime"; Expression = { "{0:N2}" -f $_.MedianResponseTime } },
				@{Name = "MinResponseTime"; Expression = { "{0:N2}" -f $_.MinResponseTime } },
				@{Name = "MaxResponseTime"; Expression = { "{0:N2}" -f $_.MaxResponseTime } },
				@{Name = "AveragePingTime"; Expression = { if ($_.AveragePingTime) { "{0:N2}" -f $_.AveragePingTime } else { "N/A" } } }
		}
		else {
			$sortedResults = $dnsResults | Sort-Object AverageResponseTime |
				Select-Object DNSProvider, DNSServer, SuccessRate,
				@{Name = "AverageResponseTime"; Expression = { "{0:N2}" -f $_.AverageResponseTime } },
				@{Name = "MedianResponseTime"; Expression = { "{0:N2}" -f $_.MedianResponseTime } },
				@{Name = "MinResponseTime"; Expression = { "{0:N2}" -f $_.MinResponseTime } },
				@{Name = "MaxResponseTime"; Expression = { "{0:N2}" -f $_.MaxResponseTime } }
		}
		$sortedResults | Format-Table -AutoSize
		# Add spacing between website results
		Write-Host "`n"
	}
	# Calculate overall averages
	$overallAverages = $overallResults.GetEnumerator() | ForEach-Object {
		$provider = [PSCustomObject]@{
			DNSProvider                = $_.Key
			OverallAverageResponseTime = $_.Value.TotalResponseTime / $_.Value.WebsitesTested
			OverallSuccessRate         = $_.Value.TotalSuccessRate / $_.Value.WebsitesTested
			WebsitesTested             = $_.Value.WebsitesTested
		}
		# Add ping stats if available
		if ($IncludePingTest -and $_.Value.TotalPingTime -gt 0 -and $_.Value.PingTestsCount -gt 0) {
			$provider | Add-Member -NotePropertyName OverallAveragePingTime -NotePropertyValue ($_.Value.TotalPingTime / $_.Value.PingTestsCount)
		}
		return $provider
	} | Sort-Object OverallAverageResponseTime
	# Display overall results header with execution mode note
	Write-Host "==================================================================" -ForegroundColor Cyan
	Write-Host "OVERALL RESULTS" -ForegroundColor Cyan
	if ($Parallel) {
		Write-Host "(Note: Times in parallel mode are typically higher than sequential mode)" -ForegroundColor Yellow
	}
	Write-Host "==================================================================" -ForegroundColor Cyan
	Write-Host "`nAll DNS Providers Ranked by Average Response Time:" -ForegroundColor Yellow
	if ($IncludePingTest) {
		$formattedAverages = $overallAverages | Select-Object DNSProvider,
		@{Name = "OverallAverageResponseTime"; Expression = { "{0:N2}" -f $_.OverallAverageResponseTime } },
		@{Name = "OverallSuccessRate"; Expression = { "{0:N2}" -f $_.OverallSuccessRate } },
		@{Name = "OverallAveragePingTime"; Expression = { if ($_.OverallAveragePingTime) { "{0:N2}" -f $_.OverallAveragePingTime } else { "N/A" } } },
		WebsitesTested
	}
	else {
		$formattedAverages = $overallAverages | Select-Object DNSProvider,
		@{Name = "OverallAverageResponseTime"; Expression = { "{0:N2}" -f $_.OverallAverageResponseTime } },
		@{Name = "OverallSuccessRate"; Expression = { "{0:N2}" -f $_.OverallSuccessRate } },
		WebsitesTested
	}
	$formattedAverages | Format-Table -AutoSize
	# Export results if requested
	if ($OutputFormat -ne 'Console' -and $globalResults.Count -gt 0) {
		switch ($OutputFormat) {
			'CSV' {
				# Flatten the detailed results for CSV export
				$flatResults = @()
				foreach ($result in $globalResults) {
					$flatResults += [PSCustomObject]@{
						Website             = $result.Website
						DNSProvider         = $result.DNSProvider
						DNSServer           = $result.DNSServer
						TestsRun            = $result.TestsRun
						SuccessCount        = $result.SuccessCount
						FailedCount         = $result.FailedCount
						SuccessRate         = $result.SuccessRate
						AverageResponseTime = $result.AverageResponseTime
						MedianResponseTime  = $result.MedianResponseTime
						MinResponseTime     = $result.MinResponseTime
						MaxResponseTime     = $result.MaxResponseTime
						StdDevResponseTime  = $result.StdDevResponseTime
					}
					# Add ping metrics if available
					if ($IncludePingTest) {
						$flatResults[-1] | Add-Member -NotePropertyName AveragePingTime -NotePropertyValue $result.AveragePingTime
						$flatResults[-1] | Add-Member -NotePropertyName MinPingTime -NotePropertyValue $result.MinPingTime
						$flatResults[-1] | Add-Member -NotePropertyName MaxPingTime -NotePropertyValue $result.MaxPingTime
					}
				}
				$flatResults | Export-Csv -Path $OutputPath -NoTypeInformation
				Write-Host "`nResults exported to CSV file: $OutputPath" -ForegroundColor Green
			}
			'JSON' {
				$globalResults | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputPath
				Write-Host "`nResults exported to JSON file: $OutputPath" -ForegroundColor Green
			}
			'GridView' {
				# Flatten the detailed results for GridView
				$flatResults = @()
				foreach ($result in $globalResults) {
					$gridViewItem = [PSCustomObject]@{
						Website             = $result.Website
						DNSProvider         = $result.DNSProvider
						DNSServer           = $result.DNSServer
						TestsRun            = $result.TestsRun
						SuccessCount        = $result.SuccessCount
						FailedCount         = $result.FailedCount
						SuccessRate         = "$($result.SuccessRate.ToString("F2"))%"
						AverageResponseTime = "$($result.AverageResponseTime.ToString("F3")) ms"
						MedianResponseTime  = "$($result.MedianResponseTime.ToString("F3")) ms"
						MinResponseTime     = "$($result.MinResponseTime.ToString("F3")) ms"
						MaxResponseTime     = "$($result.MaxResponseTime.ToString("F3")) ms"
						StdDevResponseTime  = "$($result.StdDevResponseTime.ToString("F3")) ms"
					}
					# Add ping metrics if available
					if ($IncludePingTest) {
						$pingAvg = if ($result.AveragePingTime) { "$($result.AveragePingTime.ToString("F3")) ms" } else { "N/A" }
						$pingMin = if ($result.MinPingTime) { "$($result.MinPingTime.ToString("F3")) ms" } else { "N/A" }
						$pingMax = if ($result.MaxPingTime) { "$($result.MaxPingTime.ToString("F3")) ms" } else { "N/A" }
						$gridViewItem | Add-Member -NotePropertyName AveragePingTime -NotePropertyValue $pingAvg
						$gridViewItem | Add-Member -NotePropertyName MinPingTime -NotePropertyValue $pingMin
						$gridViewItem | Add-Member -NotePropertyName MaxPingTime -NotePropertyValue $pingMax
					}
					$flatResults += $gridViewItem
				}
				$flatResults | Out-GridView -Title "DNS Server Performance Test Results"
			}
		}
	}
	# Display execution summary
	$endTime = Get-Date
	$duration = $endTime - $startTime
	Write-Host "`n==================================================================" -ForegroundColor Cyan
	Write-Host "Test-DNSServers Execution Summary:" -ForegroundColor Cyan
	Write-Host "==================================================================" -ForegroundColor Cyan
	Write-Host "Started: $startTime" -ForegroundColor White
	Write-Host "Ended: $endTime" -ForegroundColor White
	Write-Host "Duration: $($duration.TotalSeconds.ToString("F2")) seconds" -ForegroundColor White
	Write-Host "Websites Tested: $($Websites.Count)" -ForegroundColor White
	Write-Host "DNS Servers Tested: $($DNSServers.Count)" -ForegroundColor White
	Write-Host "Tests Per Website: $TestsCount" -ForegroundColor White
	Write-Host "Total DNS Queries: $($Websites.Count * $DNSServers.Count * $TestsCount)" -ForegroundColor White
	Write-Host "==================================================================" -ForegroundColor Cyan
	# Display top 3 providers
	$top3 = $overallAverages | Select-Object -First 3
	Write-Host "`n`nTop 3 Overall DNS Providers:" -ForegroundColor Cyan
	$medals = @(
		@{Name = 'Gold 🥇'; ANSI = "`e[38;2;255;215;0m"; Legacy = 'Yellow' }
		@{Name = 'Silver 🥈'; ANSI = "`e[38;2;192;192;192m"; Legacy = 'Gray' }
		@{Name = 'Bronze 🥉'; ANSI = "`e[38;2;205;127;50m"; Legacy = "DarkYellow" }
	)
	for ($i = 0; $i -lt 3 -and $i -lt $top3.Count; $i++) {
		$message = "$($medals[$i].Name): $($top3[$i].DNSProvider) - Average Response Time: $($top3[$i].OverallAverageResponseTime.ToString("F3")) ms - Success Rate: $($top3[$i].OverallSuccessRate.ToString("F2"))%"
		if ($PSVersionTable.PSVersion.Major -gt 6) {
			Write-Host "$($medals[$i].ANSI)$message`e[0m"
		}
		else {
			Write-Host $message -ForegroundColor $medals[$i].Legacy
		}
	}
	Write-Host "`n"
	# Return the overall average results object for pipeline usage
	if ($ReturnResults.IsPresent) {
		return $overallAverages
	}
}