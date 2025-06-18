function Get-RandomHeader {
    <#
    .SYNOPSIS
        Generates a reliable set of HTTP headers to mimic real browser behavior.
    .DESCRIPTION
        The Get-RandomHeader function generates HTTP headers that mimic real browser behavior while prioritizing reliability and compatibility.
        This version is more conservative than highly randomized alternatives, focusing on widely-compatible headers that work across different APIs and services.
        The function supports various use cases, including generating headers for direct use with Invoke-RestMethod/Invoke-WebRequest, or creating a pre-configured HttpClient object.
    .PARAMETER GetHTTPClient
        If specified, the function returns a System.Net.Http.HttpClient object with the random headers pre-configured. Otherwise, it returns a hashtable of headers.
    .PARAMETER ChromeChannel
        Specifies the Chrome channel to fetch the version for. Valid values are 'Stable', 'Beta', 'Dev', 'Canary'. Defaults to 'Stable'.
    .EXAMPLE
        Example 1: Get a hashtable of reliable headers
        $headers = Get-RandomHeader -Verbose
        Invoke-RestMethod -Uri "https://api.example.com/data" -Headers $headers
    .EXAMPLE
        Example 2: Get a pre-configured HttpClient object
        $httpClient = Get-RandomHeader -GetHTTPClient -Verbose
        $httpClient.Dispose()
    .NOTES
        Author: Harze2k
        Date:   2025-06-19
        Version: 1.5.0 (Reliability-focused modifications)
        System.Net.Http.HttpClient is available in PowerShell 5.1 and later.
    #>
    [CmdletBinding()]
    param (
        [switch]$GetHTTPClient,
        [ValidateSet('Stable', 'Beta', 'Dev', 'Canary')][string]$ChromeChannel = 'Stable'
    )
    Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue
    $languageMap = @{
        'us' = 'en-US,en;q=0.9'; 'gb' = 'en-GB,en;q=0.9'; 'ca' = 'en-CA,en;q=0.9'
        'au' = 'en-AU,en;q=0.9'; 'de' = 'de-DE,de;q=0.9,en;q=0.8'; 'fr' = 'fr-FR,fr;q=0.9,en;q=0.8'
        'es' = 'es-ES,es;q=0.9,en;q=0.8'; 'it' = 'it-IT,it;q=0.9,en;q=0.8'; 'jp' = 'ja-JP,ja;q=0.9,en;q=0.8'
        'br' = 'pt-BR,pt;q=0.9,en;q=0.8'; 'in' = 'en-IN,en;q=0.9'; 'se' = 'sv-SE,sv;q=0.9,en;q=0.8'
        'nl' = 'nl-NL,nl;q=0.9,en;q=0.8'
    }
    $defaultLanguage = 'en-US,en;q=0.9'
    $uaProfiles = @(
        @{ UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{ChromeVersion} Safari/537.36'; SecChUa = '"Chromium";v="{ChromeVersionBase}", "Google Chrome";v="{ChromeVersionBase}", "Not A;Brand";v="99"'; SecChUaMobile = '?0'; SecChUaPlatform = '"Windows"'; OS = 'Windows'; Browser = 'Chrome'; Weight = 40 },
        @{ UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:{FirefoxVersion}) Gecko/20100101 Firefox/{FirefoxVersion}'; SecChUa = ''; SecChUaMobile = '?0'; SecChUaPlatform = '"Windows"'; OS = 'Windows'; Browser = 'Firefox'; Weight = 25 },
        @{ UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{ChromeVersion} Safari/537.36 Edg/{EdgeVersion}'; SecChUa = '"Microsoft Edge";v="{EdgeVersionBase}", "Chromium";v="{ChromeVersionBase}", "Not A;Brand";v="99"'; SecChUaMobile = '?0'; SecChUaPlatform = '"Windows"'; OS = 'Windows'; Browser = 'Edge'; Weight = 20 },
        @{ UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{ChromeVersion} Safari/537.36'; SecChUa = '"Chromium";v="{ChromeVersionBase}", "Google Chrome";v="{ChromeVersionBase}", "Not A;Brand";v="99"'; SecChUaMobile = '?0'; SecChUaPlatform = '"macOS"'; OS = 'Mac'; Browser = 'Chrome'; Weight = 10 },
        @{ UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/{SafariVersion} Safari/605.1.15'; SecChUa = ''; SecChUaMobile = '?0'; SecChUaPlatform = '"macOS"'; OS = 'Mac'; Browser = 'Safari'; Weight = 5 }
    )
    $referrers = @('https://www.google.com/', 'https://www.bing.com/', 'https://github.com/', 'https://reddit.com/', $null, $null, $null, $null)
    $acceptHeaders = @( 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7', 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8','application/json,text/plain,*/*', 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8' )
    $defaultLatestChromeMajor = 136
    switch ($ChromeChannel) {
        'Dev' { $defaultLatestChromeMajor = 138 }
        'Beta' { $defaultLatestChromeMajor = 137 }
        'Canary' { $defaultLatestChromeMajor = 138 }
        default { $defaultLatestChromeMajor = 136 }
    }
    $defaultMinChromeMajorOffset = 8
    $latestChromeMajorVersion = $null
    Write-Verbose "Attempting to fetch the latest '$ChromeChannel' Chrome version from Chrome for Testing (CfT) JSON endpoint..."
    try {
        $cftJsonUrl = 'https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json'
        $cftData = Invoke-RestMethod -Uri $cftJsonUrl -TimeoutSec 5 -ErrorAction Stop
        $channelVersionString = $cftData.channels.($ChromeChannel).version
        if ($channelVersionString -and $channelVersionString -match '^\d+\.\d+\.\d+\.\d+$') {
            $majorVersionString = ($channelVersionString -split '\.')[0]
            if ($majorVersionString -match '^\d+$') {
                $latestChromeMajorVersion = [int]$majorVersionString
                Write-Verbose "Successfully fetched latest '$ChromeChannel' Chrome version: $channelVersionString (Major: $latestChromeMajorVersion)"
            }
        }
    }
    catch {
        Write-Verbose "Could not fetch Chrome version from CfT, using default: $($_.Exception.Message)"
    }
    if (-not $latestChromeMajorVersion) {
        $latestChromeMajorVersion = $defaultLatestChromeMajor
        Write-Verbose "Using default Chrome major version: $latestChromeMajorVersion"
    }
    $minChromeMajorVersion = $latestChromeMajorVersion - $defaultMinChromeMajorOffset
    if ($minChromeMajorVersion -lt 120) {
        $minChromeMajorVersion = 120
    }
    $firefoxMinMajor = 120
    $firefoxMaxMajor = [Math]::Max($firefoxMinMajor, ($latestChromeMajorVersion + 2))
    $safariMinMajor = 16
    $safariMaxMajor = 17
    Write-Verbose "Version ranges - Chrome: $minChromeMajorVersion-$latestChromeMajorVersion, Firefox: $firefoxMinMajor-$firefoxMaxMajor, Safari: $safariMinMajor-$safariMaxMajor"
    $countryCode = $null
    try {
        $response = Invoke-RestMethod -Method 'GET' -Uri 'http://ip-api.com/json/?fields=countryCode' -TimeoutSec 3 -ErrorAction Stop
        $countryCode = $response.countryCode.ToLower()
        Write-Verbose "Detected country code: $countryCode"
    }
    catch {
        $countryCode = 'us'
        Write-Verbose "Could not detect location, defaulting to US"
    }
    $acceptLanguage = $languageMap[$countryCode]
    if (-not $acceptLanguage) {
        $acceptLanguage = $defaultLanguage
    }
    $weightedProfiles = @()
    foreach ($profile in $uaProfiles) {
        for ($i = 0; $i -lt $profile.Weight; $i++) {
            $weightedProfiles += $profile
        }
    }
    $selectedProfile = Get-Random -InputObject $weightedProfiles
    $finalChromeVersion = "0.0.0.0"
    $finalEdgeVersion = "0.0.0.0"
    $finalFirefoxVersion = "0.0"
    $finalSafariVersion = "0.0.0"
    switch ($selectedProfile.Browser) {
        'Chrome' {
            $chromeVersionBase = Get-Random -Minimum $minChromeMajorVersion -Maximum ($latestChromeMajorVersion + 1)
            $finalChromeVersion = "$chromeVersionBase.0.$(Get-Random -Minimum 5000 -Maximum 7000).$(Get-Random -Minimum 100 -Maximum 200)"
        }
        'Edge' {
            $edgeVersionBase = Get-Random -Minimum $minChromeMajorVersion -Maximum ($latestChromeMajorVersion + 1)
            $chromeVersionBase = $edgeVersionBase
            $finalEdgeVersion = "$edgeVersionBase.0.$(Get-Random -Minimum 1800 -Maximum 2200).$(Get-Random -Minimum 50 -Maximum 99)"
            $finalChromeVersion = "$chromeVersionBase.0.$(Get-Random -Minimum 5000 -Maximum 7000).$(Get-Random -Minimum 100 -Maximum 200)"
        }
        'Firefox' {
            $firefoxVersionBase = Get-Random -Minimum $firefoxMinMajor -Maximum ($firefoxMaxMajor + 1)
            $finalFirefoxVersion = "$firefoxVersionBase.0"
        }
        'Safari' {
            $safariVersionBase = Get-Random -Minimum $safariMinMajor -Maximum ($safariMaxMajor + 1)
            $finalSafariVersion = "$safariVersionBase.$(Get-Random -Minimum 0 -Maximum 3).$(Get-Random -Minimum 0 -Maximum 2)"
        }
    }
    $userAgentString = $selectedProfile.UA -replace '{ChromeVersion}', $finalChromeVersion `
        -replace '{FirefoxVersion}', $finalFirefoxVersion `
        -replace '{EdgeVersion}', $finalEdgeVersion `
        -replace '{SafariVersion}', $finalSafariVersion
    Write-Verbose "Generated User-Agent: $userAgentString"
    $Header = @{
        'User-Agent'                = $userAgentString
        'Accept'                    = Get-Random -InputObject $acceptHeaders
        'Accept-Language'           = $acceptLanguage
        'Accept-Encoding'           = 'gzip, deflate, br'
        'Connection'                = 'keep-alive'
        'Upgrade-Insecure-Requests' = '1'
    }
    if ($selectedProfile.Browser -match 'Chrome|Edge' -and $selectedProfile.SecChUa) {
        $secChUaString = $selectedProfile.SecChUa -replace '{ChromeVersionBase}', $chromeVersionBase -replace '{EdgeVersionBase}', $edgeVersionBase
        if ($secChUaString -and $secChUaString -notmatch '{.*}') {
            $Header['Sec-Ch-Ua'] = $secChUaString
            $Header['Sec-Ch-Ua-Mobile'] = $selectedProfile.SecChUaMobile
            $Header['Sec-Ch-Ua-Platform'] = $selectedProfile.SecChUaPlatform
        }
    }
    $selectedReferrer = Get-Random -InputObject $referrers
    if ($selectedReferrer) {
        $Header['Referer'] = $selectedReferrer
    }
    if ((Get-Random -Maximum 10) -gt 7) {
        $Header['DNT'] = '1'
    }
    Write-Verbose "Generated headers for $($selectedProfile.Browser) on $($selectedProfile.OS)"
    try {
        if ($GetHTTPClient.IsPresent) {
            Write-Verbose "Creating HttpClient with reliable headers"
            $httpClient = [System.Net.Http.HttpClient]::new()
            foreach ($key in $Header.Keys) {
                try {
                    $httpClient.DefaultRequestHeaders.TryAddWithoutValidation($key, $Header[$key]) | Out-Null
                    Write-Verbose "Added header: $key"
                }
                catch {
                    Write-Warning "Could not add header '$key': $($_.Exception.Message)"
                }
            }
            return $httpClient
        }
        else {
            return $Header
        }
    }
    catch {
        Write-Error "Failed to create headers: $($_.Exception.Message)"
        return $null
    }
}