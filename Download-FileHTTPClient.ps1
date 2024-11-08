function Download-FileHTTPClient {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][string]$FileName,
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $false)]$HTTPClient,
        [Parameter(Mandatory = $false)][int]$BufferFactor = 2, # Use 2 when speed is ~100/mbit. Use 3-5 when speed is ~1000/mbit.
        [switch]$DisposeClient
    )
    function Get-OptimalBufferSize {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)][long]$FileSize,
            [Parameter(Mandatory = $false)][int]$BufferFactor = 2 # Use 2 when speed is ~100/mbit. Use 3-5 when speed is ~1000/mbit.
        )
        $availableMemory = (Get-CimInstance Win32_OperatingSystem -Verbose:$false).FreePhysicalMemory * 1KB
        $maxBuffer = [Math]::Min(($availableMemory * 0.01), 8MB)
        $baseBuffer = switch ($fileSize) {
            { $_ -lt 1MB } {
                8KB
                break
            }      # Tiny files
            { $_ -lt 10MB } {
                64KB
                break
            }       # Small files
            { $_ -lt 100MB } {
                256KB
                break
            }       # Medium files
            { $_ -lt 1GB } {
                512KB
                break
            }       # Large files
            { $_ -lt 10GB } {
                1MB
                break
            }       # Very large files
            default {
                2MB
                break
            }       # Huge files
        }
        $calculatedBuffer = [Math]::Min(($baseBuffer * $bufferFactor), $maxBuffer)
        $powerOf2 = [Math]::Log($calculatedBuffer, 2)
        $roundedPowerOf2 = [Math]::Ceiling($powerOf2)
        $finalBuffer = [Math]::Pow(2, $roundedPowerOf2)
        $finalBuffer = [Math]::Min($finalBuffer, $maxBuffer)
        return [pscustomobject]@{
            "AvailableRAMMemory"       = $(Format-FileSize $availableMemory)
            "MaximumSafeBuffer"        = $(Format-FileSize $maxBuffer)
            "BaseBufferforFileSize"    = $(Format-FileSize $baseBuffer)
            "BufferFactorApplied"      = $bufferFactor
            "FinalBufferSizeFormatted" = $(Format-FileSize $finalBuffer)
            "FinalBufferSize"          = $finalBuffer
        }
    }
    function Format-FileSize {
        param (
            [Parameter(Mandatory = $true)][long]$Bytes
        )
        $sizes = 'B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB'
        $index = [Math]::Floor([Math]::Log($bytes, 1000))
        $index = [Math]::Min($index, $sizes.Count - 1)
        $value = $bytes / [Math]::Pow(1000, $index)
        return "{0:N2} {1}" -f $value, $sizes[$index]
    }
    if (!($FileName)) {
        $fileName = Split-Path -Path $Url -Leaf
        if (-not $fileName -or [string]::IsNullOrEmpty([IO.Path]::GetExtension($fileName))) {
            Write-Host "No file name provided and no filename with an extension could be extracted from the URL."
            return
        }
    }
    if (!($HTTPClient)) {
        Write-Host "No HTTPClient provided. Creating a new one."
        $HTTPClient = [System.Net.Http.HttpClient]::new()
    }
    try {
        $outputFile = Join-Path -Path $FilePath -ChildPath $FileName
        $request = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Get, $url)
        $response = $HTTPClient.SendAsync($request, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
        $contentStream = $response.Content.ReadAsStreamAsync().Result
        $totalLength = $response.Content.Headers.ContentLength
        $totalLengthFormatted = Format-FileSize -Bytes $totalLength
        $fileStream = [System.IO.File]::Create($outputFile)
        $bufferSize = Get-OptimalBufferSize -FileSize $totalLength -BufferFactor $BufferFactor
        if ($VerbosePreference) {
            $bufferSize | Format-List | Out-Host
        }
        #Write-Host "Starting download of $fileName (Total size: $totalLengthFormatted) using a buffer size of $($bufferSize.FinalBufferSizeFormatted)" -ForegroundColor Cyan
        $buffer = New-Object byte[] $bufferSize.FinalBufferSize
        $totalBytesRead = 0
        $lastProgressTime = Get-Date
        $progressUpdateInterval = [TimeSpan]::FromMilliseconds(1000) # Update every 1000ms
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        while (($bytesRead = $contentStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $fileStream.Write($buffer, 0, $bytesRead)
            $totalBytesRead += $bytesRead
            $progressPercentage = [math]::Floor(($totalBytesRead / $totalLength) * 100)
            $now = Get-Date
            if (($now - $lastProgressTime) -ge $progressUpdateInterval) {
                $downloadedFormatted = Format-FileSize -Bytes $totalBytesRead
                $currentSpeed = Format-FileSize -Bytes ($totalBytesRead / $stopwatch.Elapsed.TotalSeconds)
                Write-Progress -Activity "Downloading $fileName" `
                    -Status "$downloadedFormatted of $totalLengthFormatted ($progressPercentage%) - Current Speed: $currentSpeed/s" `
                    -PercentComplete $progressPercentage
                $lastProgressTime = $now
            }
        }
        $stopwatch.Stop()
        $elapsedSeconds = $stopwatch.Elapsed.TotalSeconds
        if ($elapsedSeconds -lt 0.001) {
            $elapsedSeconds = 0.001
        }
        $downloadSpeed = ($totalLength / $elapsedSeconds)
        $downloadSpeedFormatted = Format-FileSize -Bytes $downloadSpeed
        #Write-Host "Downloaded $outputFile successfully" -ForegroundColor Green
        return [pscustomobject]@{
            "Status"       = "Completed"
            "FileName"     = $fileName
            "FileSize"     = $totalLengthFormatted
            "FilePath"     = $outputFile
            "TimeTaken"    = "$($stopwatch.Elapsed.ToString('mm\:ss\.fff')) (minutes:seconds.milliseconds)"
            "AverageSpeed" = "$downloadSpeedFormatted/s"
        }
    }
    catch {
        Write-Host "Error during download: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        Write-Progress -Activity "Downloading $outputFile" -Completed
        if ($contentStream) {
            $contentStream.Dispose()
        }
        if ($fileStream) {
            $fileStream.Dispose()
        }
        if ($DisposeClient.IsPresent) {
            if (Get-Variable -Name 'HTTPClient*' -ErrorAction SilentlyContinue) {
                Remove-Variable -Name 'HTTPClient*' -Force -ErrorAction SilentlyContinue -Verbose:$VerbosePreference
            }
        }
    }
}