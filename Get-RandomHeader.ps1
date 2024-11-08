<#
.DESCRIPTION
The Get-RandomHeader function is a versatile utility designed to generate randomized HTTP headers for web requests in PowerShell scripts. It creates a set of headers that mimic real browser behavior, helping to avoid detection as an automated script. The function supports various use cases, including generating headers for direct use, or creating pre-configured HttpClient object.
The function attempts to determine the user's country code using an IP geolocation service, falling back to a random country if the service is unavailable. It then uses this information to set appropriate language headers. The function generates a wide variety of realistic user agents, referrers, and other HTTP headers to simulate different browsers and devices.

.PARAMETER GetHTTPClient
-GetHTTPClient: When specified, returns a System.Net.Http.HttpClient object with the random headers pre-configured.

.EXAMPLE
Example 1: Get a hashtable of random headers
$header = Get-RandomHeader
$header # Is compatible with Invoke-RestMethod and Invoke-WebRequest.

Example 2: Get a pre-configured HttpClient object
$httpClient = Get-RandomHeader -GetHTTPClient

.NOTES
The function includes a wide range of user agents, referrers, and accept headers to provide variety in the generated headers. It also adapts the language settings based on the detected or randomly selected country code.
The function handles errors gracefully, particularly when attempting to determine the user's location. If the IP geolocation service is unavailable, it falls back to selecting a random country code from a predefined list.
This function is particularly useful for web scraping tasks, API interactions, or any scenario where simulating a real browser's behavior is beneficial to avoid rate limiting or blocking by web servers.
#>
function Get-RandomHeader {
    [CmdletBinding()]
    param (
        [switch]$GetHTTPClient
    )
    try {
        $response = Invoke-RestMethod -Method 'GET' -Uri 'http://ip-api.com/json/' -TimeoutSec 5
        $countryCode = $response.countryCode.ToLower()
    }
    catch {
        $countryCode = Get-Random -InputObject @('us', 'gb', 'ca', 'au', 'de', 'fr', 'jp', 'br', 'in', 'ru', 'se', 'no', 'fi', 'dk', 'ch')
    }
    $userAgents = @(
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:91.0) Gecko/20100101 Firefox/91.0',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:92.0) Gecko/20100101 Firefox/92.0',
        'Mozilla/5.0 (X11; Linux x86_64; rv:92.0) Gecko/20100101 Firefox/92.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:91.0) Gecko/20100101 Firefox/91.0',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36',
        'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.61 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.61 Safari/537.36 Edg/94.0.992.31',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36 Edg/93.0.961.47',
        'Mozilla/5.0 (iPhone; CPU iPhone OS 14_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.2 Mobile/15E148 Safari/604.1',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.2 Safari/605.1.15',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36 OPR/79.0.4143.72',
        'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.61 Safari/537.36 OPR/80.0.4170.40',
        'Mozilla/5.0 (Windows NT 10.0; Trident/7.0; rv:11.0) like Gecko',
        'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0)',
        'Mozilla/5.0 (Linux; Android 10; SM-A205U) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.105 Mobile Safari/537.36',
        'Mozilla/5.0 (Linux; Android 10; SM-A205U) AppleWebKit/537.36 (KHTML, like Gecko) SamsungBrowser/11.0 Chrome/75.0.3770.143 Mobile Safari/537.36'
    )
    $referrer = @(
        'https://www.google.com',
        'https://www.bing.com',
        'https://search.yahoo.com',
        'https://duckduckgo.com',
        'https://www.ecosia.org',
        'https://www.qwant.com',
        'https://www.startpage.com'
    )
    $languageMap = @{
        'us' = 'en-US,en;q=0.9'
        'gb' = 'en-GB,en;q=0.9'
        'ca' = 'en-CA,en;q=0.9,fr-CA;q=0.8'
        'au' = 'en-AU,en;q=0.9'
        'de' = 'de-DE,de;q=0.9,en;q=0.8'
        'fr' = 'fr-FR,fr;q=0.9,en;q=0.8'
        'es' = 'es-ES,es;q=0.9,en;q=0.8'
        'it' = 'it-IT,it;q=0.9,en;q=0.8'
        'jp' = 'ja-JP,ja;q=0.9,en;q=0.8'
        'kr' = 'ko-KR,ko;q=0.9,en;q=0.8'
        'br' = 'pt-BR,pt;q=0.9,en;q=0.8'
        'in' = 'hi-IN,hi;q=0.9,en;q=0.8'
        'ru' = 'ru-RU,ru;q=0.9,en;q=0.8'
        'nl' = 'nl-NL,nl;q=0.9,en;q=0.8'
        'pl' = 'pl-PL,pl;q=0.9,en;q=0.8'
        'se' = 'sv-SE,sv;q=0.9,en;q=0.8'
        'no' = 'nb-NO,nb;q=0.9,en;q=0.8'
        'fi' = 'fi-FI,fi;q=0.9,en;q=0.8'
        'dk' = 'da-DK,da;q=0.9,en;q=0.8'
        'ch' = 'de-CH,de;q=0.9,fr;q=0.8,it;q=0.7,en;q=0.6'
    }
    $acceptLanguage = if ($languageMap.ContainsKey($countryCode)) {
        $languageMap[$countryCode]
    }
    else {
        'en-US,en;q=0.5'
    }
    $Header = @{
        'User-Agent'      = $userAgents | Get-Random
        'Accept'          = (Get-Random -InputObject @(
                'text/html,application/xhtml+xml,application/xml,application/json;q=0.9,image/webp,*/*;q=0.8',
                'text/html,application/xhtml+xml,application/xml,application/json;q=0.9,image/webp,image/apng,*/*;q=0.8',
                'text/html,application/xhtml+xml,application/xml,application/json;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'text/html,application/xhtml+xml,application/xml,application/json;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'
            ))
        'Accept-Language' = $acceptLanguage
        'Referer'         = $referrer | Get-Random
    }
    try {
        if ($GetHTTPClient) {
            Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue
            $clientName = "HTTPClient" + (Get-Random -Minimum 100 -Maximum 999)
            Set-Variable -Name $clientName -Value ([System.Net.Http.HttpClient]::new())
            foreach ($key in $Header.Keys) {
                Write-Verbose -Message "Adding $key : $($Header[$key]) to $clientName"
                (Get-Variable -Name $clientName -ValueOnly).DefaultRequestHeaders.Add($key, $Header[$key])
            }
            Write-Verbose -Message "Returned HTTPClient named: $clientName"
            return Get-Variable -Name $clientName -ValueOnly
        }
        else {
            return $Header
        }
    }
    catch {
        Write-Error "Failed creating return object. Error: $($_.Exception.Message)"
    }
}