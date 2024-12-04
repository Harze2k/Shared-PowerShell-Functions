<#
.DESCRIPTION
The Test-DNSServers function is a PowerShell tool designed to evaluate and compare the performance of multiple DNS servers in resolving specified websites. It conducts a series of DNS resolution tests for each combination of website and DNS server, calculates average response times, and provides a comprehensive performance analysis.
This function is particularly useful for:
Network administrators optimizing DNS configurations
Developers troubleshooting DNS-related performance issues
Users seeking to identify the fastest DNS servers for their frequently visited websites
.PARAMETER Websites
-Websites: Accepts an array of websites to test.
.PARAMETER TestsCount
-TestsCount: Accepts the number of DNS resolution tests to conduct for each combination of website and DNS server.
.PARAMETER DNSServers
-DNSServers: Accepts a hashtable of DNS servers and their respective IP addresses. The default is Google, Google2, Cloudflare, Cloudflare2, Quad9, Quad92, CleanBrowsing, and CleanBrowsing2.
.EXAMPLE
Test-DNSServers -Websites @('reddit.com', 'youtube.com', 'github.com', 'microsoft.com') -TestsCount 20
.OUTPUTS
Top 3 Overall DNS Providers:
Gold: Quad92 - Average Response Time: 2.265145 ms
Silver: Quad9 - Average Response Time: 2.330935 ms
Bronze: Cloudflare2 - Average Response Time: 2.9119 ms
#>
function Test-DNSServers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string[]]$Websites,
        [Parameter(Mandatory)][int]$TestsCount,
        [Parameter(Mandatory)][hashtable]$DNSServers = @{
            'Google'         = '8.8.8.8'
            'Google2'        = '8.8.4.4'
            'Cloudflare'     = '1.1.1.1'
            'Cloudflare2'    = '1.0.0.1'
            'Quad9'          = '9.9.9.9'
            'Quad92'         = '149.112.112.112'
            'CleanBrowsing'  = '185.228.168.9'
            'CleanBrowsing2' = '185.228.169.9'
        }
    )
    $overallResults = @{}
    foreach ($website in $Websites) {
        Write-Host "Testing website: $website" -ForegroundColor Cyan
        $dnsResults = @()
        foreach ($dns in $DNSServers.GetEnumerator()) {
            $totalResponseTime = Measure-Command {
                for ($i = 1; $i -le $TestsCount; $i++) {
                    try {
                        Resolve-DnsName -Name $website -Server $dns.Value -ErrorAction Stop | Out-Null
                    }
                    catch {
                        Write-Warning "Failed to resolve $website using $($dns.Key) DNS server"
                    }
                }
            }
            $averageResponseTime = $totalResponseTime.TotalMilliseconds / $TestsCount
            $dnsResults += [PSCustomObject]@{
                DNSProvider         = $dns.Key
                AverageResponseTime = $averageResponseTime
            }
            if (-not $overallResults.ContainsKey($dns.Key)) {
                $overallResults[$dns.Key] = @{
                    TotalResponseTime = 0
                    WebsitesTested    = 0
                }
            }
            $overallResults[$dns.Key].TotalResponseTime += $averageResponseTime
            $overallResults[$dns.Key].WebsitesTested++
        }
        $sortedResults = $dnsResults | Sort-Object AverageResponseTime
        Write-Host "`nAverage DNS Provider Response Times (in milliseconds) for $website after $TestsCount tests:" -ForegroundColor Green
        $sortedResults | Format-Table -AutoSize
    }
    $overallAverages = $overallResults.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{
            DNSProvider                = $_.Key
            OverallAverageResponseTime = $_.Value.TotalResponseTime / $_.Value.WebsitesTested
        }
    } | Sort-Object OverallAverageResponseTime
    $top3 = $overallAverages | Select-Object -First 3
    Write-Host "`nTop 3 Overall DNS Providers:" -ForegroundColor Cyan
    $colors = @(
        @{Name = 'Gold'; ANSI = "`e[38;2;255;215;0m"; Legacy = 'Yellow' }
        @{Name = 'Silver'; ANSI = "`e[38;2;192;192;192m"; Legacy = 'Gray' }
        @{Name = 'Bronze'; ANSI = "`e[38;2;205;127;50m"; Legacy = 'DarkRed' }
    )
    for ($i = 0; $i -lt 3; $i++) {
        $message = "$($colors[$i].Name): $($top3[$i].DNSProvider) - Average Response Time: $($top3[$i].OverallAverageResponseTime) ms"
        if ($PSVersionTable.PSVersion.Major -gt 6) {
            Write-Host "$($colors[$i].ANSI)$message`e[0m"
        }
        else {
            Write-Host $message -ForegroundColor $colors[$i].Legacy
        }
    }
}