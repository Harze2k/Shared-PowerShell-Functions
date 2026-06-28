#Requires -Version 7.0
#Requires -modules Microsoft.PowerShell.ThreadJob
#Requires -RunAsAdministrator
<#
Author:  Harze2k
Date:    2026-06-28
Version: 4.2
    - Added -ReplaceLockedModuleOnReboot and -AllowVersionedSubfolder to Update-Modules.
    - Fixed parallel retrieval of online module metadata and per-module repository filtering.
    - Fixed cleanup for modules outside PSResourceGet-managed locations, including $PSHOME.
    - Reduced normal-path console output and consolidated operation summaries.
    - Fixed primary-manifest filtering, parallel result counts, and cleanup status reporting.
#>
#region Check-PSResourceRepository
function Check-PSResourceRepository {
	<#
    .SYNOPSIS
    Ensures PSResourceGet and the required repositories are available.
    .DESCRIPTION
    Loads Microsoft.PowerShell.PSResourceGet and configures PSGallery, NuGetGallery,
    and NuGet as trusted repositories with the priorities used by this script.
    .PARAMETER ImportDependencies
    Reimports Microsoft.PowerShell.PSResourceGet even when its commands are available.
    .PARAMETER ForceInstall
    Reinstalls Microsoft.PowerShell.PSResourceGet when running Windows PowerShell 5.1.
    .PARAMETER TimeoutSeconds
    Maximum time allowed for dependency and repository setup operations.
    .INPUTS
    None.
    .OUTPUTS
    System.Boolean. Returns false when setup cannot continue; successful setup writes no value.
    .EXAMPLE
    Check-PSResourceRepository -ImportDependencies
    Loads PSResourceGet and verifies the repository configuration.
    .NOTES
    Requires administrator rights and network access to the configured repositories.
    .LINK
    Register-PSResourceRepository
    #>
	[CmdletBinding()]
	param (
		[switch]$ImportDependencies,
		[switch]$ForceInstall,
		[int]$TimeoutSeconds = 30
	)
	$isPSCore = $PSVersionTable.PSVersion.Major -ge 6
	$hasPSResourceGet = [bool](Get-Command -Name 'Get-PSResourceRepository' -ErrorAction SilentlyContinue)
	New-Log "PowerShell version: $($PSVersionTable.PSVersion) | PSCore: $isPSCore | PSResourceGet available: $hasPSResourceGet"
	function Invoke-WithTimeout {
		[CmdletBinding()]
		param (
			[Parameter(Mandatory)][scriptblock]$ScriptBlock,
			[int]$Timeout = 30,
			[string]$OperationName = 'Operation'
		)
		$runspace = $null
		$powershell = $null
		try {
			$runspace = [runspacefactory]::CreateRunspace()
			$runspace.Open()
			$powershell = [powershell]::Create()
			$powershell.Runspace = $runspace
			[void]$powershell.AddScript($ScriptBlock)
			$handle = $powershell.BeginInvoke()
			$completed = $handle.AsyncWaitHandle.WaitOne($Timeout * 1000)
			if (-not $completed) {
				New-Log "$OperationName timed out after $Timeout seconds." -Level WARNING
				$powershell.Stop()
				return $null
			}
			if ($powershell.HadErrors) {
				$errorMsg = $powershell.Streams.Error | ForEach-Object { $_.ToString() } | Join-String -Separator '; '
				New-Log "$OperationName had errors: $errorMsg" -Level WARNING
			}
			return $powershell.EndInvoke($handle)
		}
		catch {
			New-Log "$OperationName failed" -Level ERROR
			return $null
		}
		finally {
			if ($powershell) { $powershell.Dispose() }
			if ($runspace) { $runspace.Close(); $runspace.Dispose() }
		}
	}
	function Set-TlsProtocol {
		try {
			$existingProtocols = [Net.ServicePointManager]::SecurityProtocol
			$tls12Enum = [Net.SecurityProtocolType]::Tls12
			if (-not ($existingProtocols -band $tls12Enum)) {
				[Net.ServicePointManager]::SecurityProtocol = $existingProtocols -bor $tls12Enum
				New-Log "TLS 1.2 security protocol enabled."
			}
			else {
				New-Log "TLS 1.2 already enabled."
			}
			return $true
		}
		catch {
			New-Log "Unable to set TLS 1.2" -Level ERROR
			return $false
		}
	}
	function Install-PSResourceGetForPS5 {
		[CmdletBinding()]
		param (
			[int]$Timeout = 30,
			[switch]$Force
		)
		New-Log "Attempting to install Microsoft.PowerShell.PSResourceGet for PS 5.1$(if ($Force) { ' (Force)' })..."
		try {
			$psGalleryScript = { Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue }
			$psGallery = Invoke-WithTimeout -ScriptBlock $psGalleryScript -Timeout $Timeout -OperationName "Get-PSRepository PSGallery"
			if ($null -eq $psGallery) {
				New-Log "Could not query PSGallery repository - may need manual registration." -Level WARNING
			}
			elseif (-not $psGallery.Trusted) {
				New-Log "Setting PSGallery to Trusted..."
				$setRepoScript = { Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction Stop }
				Invoke-WithTimeout -ScriptBlock $setRepoScript -Timeout $Timeout -OperationName "Set-PSRepository Trusted" | Out-Null
				New-Log "PSGallery set to Trusted." -Level SUCCESS
			}
			else {
				New-Log "PSGallery is already trusted."
			}
		}
		catch {
			New-Log "Error configuring PSGallery" -Level ERROR
		}
		$forceFlag = $Force.IsPresent
		$installScript = [scriptblock]::Create(@"
            `$ErrorActionPreference = 'Stop'
            Install-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Repository 'PSGallery' -Scope AllUsers -Force:$forceFlag -AllowClobber -AcceptLicense -SkipPublisherCheck -Confirm:`$false
"@)
		New-Log "Installing Microsoft.PowerShell.PSResourceGet via Install-Module..."
		Invoke-WithTimeout -ScriptBlock $installScript -Timeout ($Timeout * 2) -OperationName "Install-Module PSResourceGet" | Out-Null
		try {
			Import-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Force -ErrorAction Stop -Verbose:$false
			New-Log "Successfully imported Microsoft.PowerShell.PSResourceGet." -Level SUCCESS
			return $true
		}
		catch {
			New-Log "Failed to import Microsoft.PowerShell.PSResourceGet" -Level ERROR
			return $false
		}
	}
	function Import-PSResourceGetModule {
		[CmdletBinding()]
		param ([switch]$Force)
		$action = if ($Force) { "Force importing" } else { "Importing" }
		New-Log "$action Microsoft.PowerShell.PSResourceGet module..."
		try {
			Import-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Force:$Force -ErrorAction Stop -Verbose:$false
			New-Log "Successfully imported PSResourceGet." -Level SUCCESS
			return $true
		}
		catch {
			New-Log "Failed to import PSResourceGet" -Level ERROR
			return $false
		}
	}
	function Register-RepositoryPSResourceGet {
		[CmdletBinding()]
		param (
			[Parameter(Mandatory)][string]$Name,
			[string]$Uri,
			[Parameter(Mandatory)][int]$Priority,
			[string]$ApiVersion = 'v3',
			[switch]$IsPSGallery
		)
		try {
			$repository = Get-PSResourceRepository -Name $Name -ErrorAction SilentlyContinue
			$needsUpdate = ($null -eq $repository) -or ($repository.Priority -ne $Priority) -or (-not $repository.Trusted)
			if (-not $IsPSGallery -and $Uri -and $repository) {
				$currentUri = if ($repository.Uri) { $repository.Uri.AbsoluteUri } else { $null }
				$needsUpdate = $needsUpdate -or ($currentUri -ne $Uri)
			}
			if ($needsUpdate) {
				if ($IsPSGallery) {
					New-Log "Configuring PSGallery (Priority: $Priority, Trusted: True)."
					Set-PSResourceRepository -Name $Name -Priority $Priority -Trusted -ErrorAction Stop
				}
				else {
					New-Log "Registering repository '$Name' (Uri: $Uri, Priority: $Priority)."
					$registerParams = @{
						Name        = $Name
						Uri         = $Uri
						Priority    = $Priority
						Trusted     = $true
						Force       = $true
						PassThru    = $false
						ErrorAction = 'Stop'
					}
					if ($ApiVersion -eq 'v2') {
						$registerParams.ApiVersion = 'v2'
						New-Log "Using API Version V2 for '$Name'."
					}
					Register-PSResourceRepository @registerParams
				}
				New-Log "Successfully configured '$Name' repository." -Level SUCCESS
			}
			else {
				New-Log "'$Name' repository already configured correctly."
			}
			return $true
		}
		catch {
			New-Log "Failed to configure '$Name'" -Level ERROR
			return $false
		}
	}
	Set-TlsProtocol | Out-Null
	$needsDependencyWork = (-not $hasPSResourceGet) -or $ImportDependencies.IsPresent
	if ($needsDependencyWork) {
		if ($ImportDependencies.IsPresent -and $hasPSResourceGet) {
			New-Log "-ImportDependencies specified. Re-importing PSResourceGet module..."
		}
		if ($isPSCore) {
			if (-not (Import-PSResourceGetModule -Force:$ImportDependencies.IsPresent)) {
				New-Log "Could not import PSResourceGet in PS7." -Level ERROR
				return $false
			}
		}
		else {
			$existingModule = Get-Module -Name 'Microsoft.PowerShell.PSResourceGet' -ListAvailable -ErrorAction SilentlyContinue
			if ($ForceInstall.IsPresent -or -not $existingModule) {
				if ($ForceInstall.IsPresent -and $existingModule) {
					New-Log "-ForceInstall specified. Reinstalling PSResourceGet..."
				}
				if (-not (Install-PSResourceGetForPS5 -Timeout $TimeoutSeconds -Force:$ForceInstall.IsPresent)) {
					New-Log "Could not install PSResourceGet. Cannot continue." -Level ERROR
					return $false
				}
			}
			else {
				if (-not (Import-PSResourceGetModule -Force:$ImportDependencies.IsPresent)) {
					New-Log "Could not import existing PSResourceGet module." -Level ERROR
					return $false
				}
			}
		}
		$hasPSResourceGet = [bool](Get-Command -Name 'Get-PSResourceRepository' -ErrorAction SilentlyContinue)
	}
	if (-not $hasPSResourceGet) {
		New-Log "PSResourceGet cmdlets still not available. Aborting." -Level ERROR
		return $false
	}
	New-Log "PSResourceGet cmdlets are available. Configuring repositories..."
	$repositories = @(
		@{ Name = 'PSGallery'; Uri = $null; Priority = 30; IsPSGallery = $true }
		@{ Name = 'NuGetGallery'; Uri = 'https://api.nuget.org/v3/index.json'; Priority = 40 }
		@{ Name = 'NuGet'; Uri = 'https://www.nuget.org/api/v2'; Priority = 50; ApiVersion = 'v2' }
	)
	$overallSuccess = $true
	foreach ($repo in $repositories) {
		$splatParams = @{
			Name        = $repo.Name
			Priority    = $repo.Priority
			IsPSGallery = [bool]$repo.IsPSGallery
		}
		if ($repo.Uri) { $splatParams.Uri = $repo.Uri }
		if ($repo.ApiVersion) { $splatParams.ApiVersion = $repo.ApiVersion }
		if (-not (Register-RepositoryPSResourceGet @splatParams)) {
			$overallSuccess = $false
		}
	}
	if ($overallSuccess) {
		New-Log "All repositories configured successfully." -Level SUCCESS
	}
	else {
		New-Log "Some repositories could not be configured." -Level WARNING
	}
}
#endregion Check-PSResourceRepository
#region Get-ModuleInfo
function Get-ModuleInfo {
	<#
    .SYNOPSIS
    Builds an inventory of installed PowerShell modules.
    .DESCRIPTION
    Scans the supplied directories for module manifests and PSGetModuleInfo.xml files,
    parses their metadata in parallel, normalizes installation paths, and groups the
    results by module name.
    .PARAMETER Paths
    Directories to scan recursively. Paths from PSModulePath are typically supplied.
    .PARAMETER IgnoredModules
    Module names to exclude from the returned inventory.
    .PARAMETER ThrottleLimit
    Maximum number of files processed concurrently.
    .INPUTS
    None.
    .OUTPUTS
    System.Collections.Specialized.OrderedDictionary. Keys are module names and values
    are arrays of installation records containing version, path, prerelease, and author data.
    .EXAMPLE
    $paths = $env:PSModulePath -split [IO.Path]::PathSeparator
    $inventory = Get-ModuleInfo -Paths $paths -IgnoredModules 'BurntToast'
    Scans standard module paths and excludes BurntToast.
    .NOTES
    Requires PowerShell 7 or later for ForEach-Object -Parallel.
    .LINK
    Get-ModuleUpdateStatus
    #>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string[]]$Paths,
		[string[]]$IgnoredModules = @(),
		[int]$ThrottleLimit = ([System.Environment]::ProcessorCount * 2)
	)
	$helperFunctionDefinitions = @{
		"New-Log"                 = ${function:New-Log}.ToString()
		"Parse-ModuleVersion"     = ${function:Parse-ModuleVersion}.ToString()
		"Get-ManifestVersionInfo" = ${function:Get-ManifestVersionInfo}.ToString()
		"Resolve-ModuleVersion"   = ${function:Resolve-ModuleVersion}.ToString()
		"Get-ModuleformPath"      = ${function:Get-ModuleformPath}.ToString()
		"Get-ModuleInfoFromXml"   = ${function:Get-ModuleInfoFromXml}.ToString()
		"Test-IsResourceFile"     = ${function:Test-IsResourceFile}.ToString()
	}
	foreach ($funcName in $helperFunctionDefinitions.Keys) {
		if ([string]::IsNullOrWhiteSpace($helperFunctionDefinitions[$funcName])) {
			Write-Error "Helper function '$funcName' could not be found. It must be loaded."
			return
		}
	}
	$scanPaths = @($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
	New-Log "Scanning $($scanPaths.Count) module path(s) for manifests and PSGetModuleInfo.xml files..."
	$fileDiscoveryStartTime = Get-Date
	$allPotentialFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
	foreach ($dir in $scanPaths) {
		try {
			if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
				New-Log "Skipping missing module path '$dir'." -Level WARNING
				continue
			}
			$psd1Files = @(Get-ChildItem -LiteralPath $dir -Recurse -File -Filter "*.psd1" -ErrorAction Stop)
			$xmlFiles = @(Get-ChildItem -LiteralPath $dir -Recurse -File -Filter "PSGetModuleInfo.xml" -ErrorAction Stop)
			foreach ($file in $psd1Files) { $allPotentialFiles.Add($file) }
			foreach ($file in $xmlFiles) { $allPotentialFiles.Add($file) }
		}
		catch {
			New-Log "Could not scan module path '$dir'" -Level ERROR
		}
	}
	$allPotentialFiles = @($allPotentialFiles | Sort-Object FullName -Unique)
	$fileDiscoveryDuration = (Get-Date) - $fileDiscoveryStartTime
	if ($allPotentialFiles.Count -eq 0) {
		New-Log "No potential module files found to process." -Level WARNING
		return [ordered]@{}
	}
	New-Log "Found $($allPotentialFiles.Count) candidate file(s) in $([math]::Round($fileDiscoveryDuration.TotalSeconds, 2)) seconds; parsing with throttle $ThrottleLimit." -Level VERBOSE
	$allFoundModulesFromParallel = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
	$skippedFiles = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
	$fallbackFiles = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
	$failedFiles = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
	$null = $allPotentialFiles | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
		$fileInfo = $_
		$filePath = $fileInfo.FullName
		$fileExtension = $fileInfo.Extension
		$VerbosePreference = $using:VerbosePreference
		$allFoundModulesFromParallel = $using:allFoundModulesFromParallel
		$skippedFiles = $using:skippedFiles
		$fallbackFiles = $using:fallbackFiles
		$failedFiles = $using:failedFiles
		$localHelperFunctionDefinitions = $using:helperFunctionDefinitions
		foreach ($funcName in $localHelperFunctionDefinitions.Keys) {
			$funcDefinition = $localHelperFunctionDefinitions[$funcName]
			if (-not [string]::IsNullOrWhiteSpace($funcDefinition)) {
				Set-Item -Path "function:global:$funcName" -Value ([scriptblock]::Create($funcDefinition))
			}
		}
		if (Test-IsResourceFile -Path $filePath) {
			$skippedFiles.Add($filePath)
			return
		}
		if ($fileExtension -eq '.psd1') {
			$manifestInfoObj = $null
			$testManifestOutput = $null
			$usedFallback = $false
			try {
				$testManifestOutput = Test-ModuleManifest -Path $filePath -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false
			}
			catch {
				$testManifestOutput = $null
				$usedFallback = $true
			}
			if ($testManifestOutput) {
				try {
					$manifestInfoObj = Get-ManifestVersionInfo -ResData $testManifestOutput -Quick -ErrorAction Stop -WarningAction SilentlyContinue
				}
				catch {
					$manifestInfoObj = $null
					$usedFallback = $true
				}
			}
			if (-not $manifestInfoObj) {
				$usedFallback = $true
				try {
					$manifestInfoObj = Get-ManifestVersionInfo -ModuleFilePath $filePath -ErrorAction Stop -WarningAction SilentlyContinue
				}
				catch {
					$manifestInfoObj = $null
				}
			}
			if ($usedFallback -and $manifestInfoObj) {
				$fallbackFiles.Add($filePath)
			}
			$manifestInfosToProcess = @()
			if ($manifestInfoObj) {
				if ($manifestInfoObj -is [array] -or $manifestInfoObj -is [System.Collections.IList]) {
					$manifestInfosToProcess = $manifestInfoObj
				}
				else {
					$manifestInfosToProcess = @($manifestInfoObj)
				}
			}
			foreach ($mInfo in $manifestInfosToProcess) {
				if ($mInfo -and $mInfo.ModuleVersion -and $mInfo.ModuleName) {
					$allFoundModulesFromParallel.Add([PSCustomObject]@{
							ModuleName          = $mInfo.ModuleName
							ModuleVersion       = $mInfo.ModuleVersion
							ModuleVersionString = $mInfo.ModuleVersionString
							BasePath            = $mInfo.BasePath
							isPreRelease        = $mInfo.isPreRelease
							PreReleaseLabel     = $mInfo.PreReleaseLabel
							Author              = $mInfo.Author
						})
				}
			}
			if ($manifestInfosToProcess.Count -eq 0) {
				$failedFiles.Add($filePath)
			}
		}
		elseif ($fileExtension -eq '.xml') {
			$xmlInfo = Get-ModuleInfoFromXml -XmlFilePath $filePath -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
			if ($xmlInfo -and $xmlInfo.ModuleName -and $xmlInfo.ModuleVersion) {
				$allFoundModulesFromParallel.Add([PSCustomObject]@{
						ModuleName          = $xmlInfo.ModuleName
						ModuleVersion       = $xmlInfo.ModuleVersion
						ModuleVersionString = $xmlInfo.ModuleVersionString
						BasePath            = $xmlInfo.BasePath
						isPreRelease        = $xmlInfo.isPreRelease
						PreReleaseLabel     = $xmlInfo.PreReleaseLabel
						Author              = $xmlInfo.Author
					})
			}
			else {
				$failedFiles.Add($filePath)
			}
		}
	}
	$allFoundModulesArray = $allFoundModulesFromParallel.ToArray()
	if ($allFoundModulesArray.Count -eq 0) {
		New-Log "No valid module data collected after parallel processing." -Level WARNING
		return [ordered]@{}
	}
	$uniqueModules = $allFoundModulesArray | Group-Object -Property ModuleName, BasePath, ModuleVersionString | ForEach-Object { $_.Group[0] }
	$resultModules = [ordered]@{}
	$modulesGroupedByName = $uniqueModules | Where-Object { $null -ne $_.ModuleName -and $_.ModuleName -notmatch '^\d+(\.\d+)+$' } | Group-Object ModuleName
	foreach ($nameGroup in $modulesGroupedByName) {
		$moduleName = $nameGroup.Name
		$groupWithNormalizedPaths = $nameGroup.Group | Where-Object { $null -ne $_ } | ForEach-Object {
			$newObject = $_ | Select-Object *
			if ($null -ne $newObject.BasePath -and -not ([string]::IsNullOrWhiteSpace($moduleName))) {
				$currentBasePath = $newObject.BasePath
				$normalizedCurrentPath = $currentBasePath.TrimEnd('\', '/') -replace '/', '\'
				$expectedEnding = "\$moduleName"
				if (-not $normalizedCurrentPath.EndsWith($expectedEnding, [System.StringComparison]::OrdinalIgnoreCase)) {
					$leafName = Split-Path $normalizedCurrentPath -Leaf
					$parentOfCurrent = Split-Path $normalizedCurrentPath -Parent -ErrorAction SilentlyContinue
					$parentLeafName = if ($parentOfCurrent) { Split-Path $parentOfCurrent -Leaf -ErrorAction SilentlyContinue } else { $null }
					if ($parentLeafName -eq $moduleName -and $leafName -match '^\d+(\.\d+){1,3}(-.+)?$') {
						$newObject.BasePath = $parentOfCurrent
					}
					elseif ($leafName -eq 'Modules' -or $leafName -eq 'Documents') {
						$newObject.BasePath = Join-Path -Path $normalizedCurrentPath -ChildPath $moduleName -ErrorAction SilentlyContinue
					}
					else {
						$newObject.BasePath = $normalizedCurrentPath
					}
				}
				else {
					$newObject.BasePath = $normalizedCurrentPath
				}
			}
			$newObject
		}
		$modulesGroupedByBasePath = $groupWithNormalizedPaths | Where-Object { $null -ne $_.BasePath } | Group-Object -Property BasePath
		$finalModuleLocations = [System.Collections.Generic.List[object]]::new()
		foreach ($basePathGroup in $modulesGroupedByBasePath) {
			$currentBasePath = $basePathGroup.Name
			$versionsInPathGroup = $basePathGroup.Group | Group-Object -Property @{ Expression = { if ($_.ModuleVersion -is [version]) { $_.ModuleVersion } else { $_.ModuleVersionString } } }
			foreach ($versionGroup in $versionsInPathGroup) {
				$representativeEntry = $versionGroup.Group | Sort-Object -Property @{Expression = { $_.ModuleVersion -is [version] }; Descending = $true }, ModuleVersionString | Select-Object -First 1
				if ($representativeEntry) {
					$outputObject = [PSCustomObject]@{
						ModuleName          = $moduleName
						ModuleVersion       = $representativeEntry.ModuleVersion
						ModuleVersionString = $representativeEntry.ModuleVersionString
						BasePath            = $currentBasePath
						IsPreRelease        = $representativeEntry.IsPreRelease
						PreReleaseLabel     = $representativeEntry.PreReleaseLabel
						Author              = $representativeEntry.Author
					}
					$finalModuleLocations.Add($outputObject)
				}
			}
		}
		if ($finalModuleLocations.Count -gt 0) {
			$sortedLocations = $finalModuleLocations | Sort-Object BasePath, @{Expression = { $_.ModuleVersion }; Ascending = $true }
			$resultModules[$moduleName] = $sortedLocations
		}
	}
	$finalSortedModules = [ordered]@{}
	foreach ($key in ($resultModules.Keys | Sort-Object)) {
		if ($IgnoredModules -notcontains $key) {
			$finalSortedModules[$key] = $resultModules[$key]
		}
		else {
			New-Log "Skipping module '$key' as it is in the IgnoredModules list (final filter)." -Level VERBOSE
		}
	}
	$totalFunctionDuration = (Get-Date) - $fileDiscoveryStartTime
	$installationCount = @($finalSortedModules.Values | ForEach-Object { $_ }).Count
	$inventoryLevel = if ($failedFiles.Count -gt 0) { 'WARNING' } else { 'SUCCESS' }
	New-Log "Inventory complete: $($allPotentialFiles.Count) candidates produced $installationCount installation record(s) for $($finalSortedModules.Keys.Count) module(s) in $([math]::Round($totalFunctionDuration.TotalSeconds, 2)) seconds; skipped $($skippedFiles.Count) support file(s), used fallback metadata for $($fallbackFiles.Count), and could not parse $($failedFiles.Count)." -Level $inventoryLevel
	return $finalSortedModules
}
#endregion Get-ModuleInfo
#region Get-ModuleUpdateStatus
function Get-ModuleUpdateStatus {
	<#
    .SYNOPSIS
    Finds newer module versions in registered PSResource repositories.
    .DESCRIPTION
    Fetches stable and prerelease metadata in parallel, applying repository exclusions
    separately for each module. It compares the online versions with the local inventory
    and returns one record for each module that has an update.
    .PARAMETER ModuleInventory
    Module inventory returned by Get-ModuleInfo.
    .PARAMETER Repositories
    Registered PSResource repositories to query, in priority order.
    .PARAMETER ThrottleLimit
    Maximum number of concurrent repository lookups and comparisons.
    .PARAMETER TimeoutSeconds
    Hard time limit, in seconds, for the complete online lookup operation.
    .PARAMETER FindModuleTimeoutSeconds
    Soft time budget, in seconds, for repository queries for one module.
    .PARAMETER BlackList
    Maps module names to excluded repositories. Use '*' to exclude a module completely.
    .PARAMETER MatchAuthor
    Reports an update only when normalized local and repository author names match.
    .INPUTS
    None.
    .OUTPUTS
    System.Management.Automation.PSCustomObject. Each record contains the module name,
    repository, local and online versions, prerelease data, and outdated installations.
    .EXAMPLE
    $updates = Get-ModuleUpdateStatus -ModuleInventory $inventory -Repositories 'PSGallery' -MatchAuthor
    Checks PSGallery and requires matching author names.
    .NOTES
    Requires Microsoft.PowerShell.PSResourceGet and network access to the repositories.
    .LINK
    Find-PSResource
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][hashtable]$ModuleInventory,
		[string[]]$Repositories = @('PSGallery', 'NuGet'),
		[int]$ThrottleLimit = ([Environment]::ProcessorCount * 2),
		[ValidateRange(1, 3600)][int]$TimeoutSeconds = 30,
		[ValidateRange(1, 60)][int]$FindModuleTimeoutSeconds = 10,
		[hashtable]$BlackList = @{},
		[switch]$MatchAuthor
	)
	if ($PSVersionTable.PSVersion.Major -lt 7) {
		New-Log "This function requires PowerShell 7 or later. Current version: $($PSVersionTable.PSVersion)" -Level ERROR
		return
	}
	$allModuleNames = $ModuleInventory.Keys | Where-Object { $_ -and $_.Trim() } | Sort-Object -Unique
	if ($allModuleNames.Count -eq 0) {
		New-Log "Module inventory is empty. Nothing to check."
		return @()
	}
	$moduleDataArray = @()
	foreach ($moduleNameInLoop in $allModuleNames) {
		$localModulesInput = $ModuleInventory[$moduleNameInLoop]
		if ($localModulesInput -is [PSCustomObject]) {
			$localModulesInput = @($localModulesInput)
		}
		$parsedVersions = $localModulesInput | Where-Object { $_ -and ($_.PSObject.Properties.Name -contains 'ModuleVersion' -or $_.PSObject.Properties.Name -contains 'ModuleVersionString') -and $_.PSObject.Properties.Name -contains 'BasePath' } | ForEach-Object {
			[PSCustomObject]@{
				ModuleVersion       = $_.ModuleVersion
				ModuleVersionString = $_.ModuleVersionString
				PreReleaseLabel     = $_.PreReleaseLabel
				BasePath            = $_.BasePath
				IsPreRelease        = $_.IsPrerelease
				Author              = $_.Author
			}
		}
		if ($parsedVersions.Count -gt 0) {
			$highestLocalVersionInstall = $parsedVersions | Sort-Object -Property @{E = { $_.ModuleVersion }; Descending = $true }, @{E = { $_.IsPreRelease }; Ascending = $true } | Select-Object -First 1
			$moduleDataArray += [PSCustomObject]@{
				ModuleName          = $moduleNameInLoop
				HighestLocalInstall = $highestLocalVersionInstall
				AllParsedVersions   = $parsedVersions
			}
		}
	}
	$validModuleCountForProcessing = $moduleDataArray.Count
	if ($validModuleCountForProcessing -eq 0) {
		New-Log "No valid modules remaining after pre-processing local inventory." -Level WARNING
		return @()
	}
	New-Log "Prepared $validModuleCountForProcessing modules from local inventory." -Level VERBOSE
	New-Log "Starting online version pre-fetching for $($moduleDataArray.Count) modules (Throttle: $ThrottleLimit, max ${TimeoutSeconds}s)..."
	$overallOperationStartTime = Get-Date
	$onlineModuleVersionsCache = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
	$NewLogDef = ${function:New-Log}.ToString()
	# Fetch metadata concurrently and recalculate repository exclusions for each module.
	$moduleDataArray | ForEach-Object -ThrottleLimit $ThrottleLimit -TimeoutSeconds $TimeoutSeconds -Parallel {
		$moduleNameToFetch = $_.ModuleName
		${function:New-Log} = $using:NewLogDef
		$VerbosePreference = $using:VerbosePreference
		$allRepositories = @($using:Repositories)
		$blackList = $using:BlackList
		$perModuleBudget = $using:FindModuleTimeoutSeconds
		$cache = $using:onlineModuleVersionsCache
		$currentRepositories = $allRepositories
		if ($blackList -and $blackList.ContainsKey($moduleNameToFetch)) {
			$setting = $blackList[$moduleNameToFetch]
			if ($setting -eq '*') {
				$cache[$moduleNameToFetch] = [pscustomobject]@{ ModuleName = $moduleNameToFetch; Stable = $null; PreRelease = $null; ErrorFetching = $null; Skipped = $true }
				New-Log "[$moduleNameToFetch] Pre-fetch: Blacklisted ('*'). Skipping online check." -Level VERBOSE
				return
			}
			elseif ($setting -is [array]) { $currentRepositories = @($allRepositories | Where-Object { $setting -notcontains $_ }) }
			elseif ($setting -is [string]) { $currentRepositories = @($allRepositories | Where-Object { $_ -ne $setting }) }
		}
		if (@($currentRepositories).Count -eq 0) {
			$cache[$moduleNameToFetch] = [pscustomobject]@{ ModuleName = $moduleNameToFetch; Stable = $null; PreRelease = $null; ErrorFetching = $null; Skipped = $true }
			New-Log "[$moduleNameToFetch] Pre-fetch: No repositories left to check after blacklist exclusion. Skipping." -Level VERBOSE
			return
		}
		$stableResult = $null
		$prereleaseResult = $null
		$fetchError = $null
		$swModule = [System.Diagnostics.Stopwatch]::StartNew()
		foreach ($repo in $currentRepositories) {
			if ($swModule.Elapsed.TotalSeconds -gt $perModuleBudget) { $fetchError = "Per-module time budget (${perModuleBudget}s) exceeded during stable search."; break }
			try {
				$found = Find-PSResource -Name $moduleNameToFetch -Repository $repo -ErrorAction SilentlyContinue -Verbose:$false | Sort-Object -Property Version -Descending | Select-Object -First 1
				if ($found) { $stableResult = $found; break }
			}
			catch { $fetchError = "Stable search error in '$repo': $($_.Exception.Message)" }
		}
		$prereleaseRepos = if ($stableResult) { @($stableResult.Repository) } else { $currentRepositories }
		foreach ($repo in $prereleaseRepos) {
			if ($swModule.Elapsed.TotalSeconds -gt $perModuleBudget) { $fetchError = (@($fetchError, "Per-module time budget exceeded during prerelease search.") -join '; ').Trim('; ', ' '); break }
			try {
				$found = Find-PSResource -Name $moduleNameToFetch -Prerelease -Repository $repo -ErrorAction SilentlyContinue -Verbose:$false
				| Where-Object { $_.IsPrerelease } | Sort-Object -Property Version -Descending | Select-Object -First 1
				if ($found) { $prereleaseResult = $found; break }
			}
			catch { $fetchError = (@($fetchError, "Prerelease search error in '$repo': $($_.Exception.Message)") -join '; ').Trim('; ', ' ') }
		}
		$cache[$moduleNameToFetch] = [pscustomobject]@{
			ModuleName    = $moduleNameToFetch
			Stable        = $stableResult
			PreRelease    = $prereleaseResult
			ErrorFetching = $fetchError
			Skipped       = $false
		}
	}
	$preFetchTimeouts = 0
	foreach ($moduleEntry in $moduleDataArray) {
		if (-not $onlineModuleVersionsCache.ContainsKey($moduleEntry.ModuleName)) {
			New-Log "[$($moduleEntry.ModuleName)] No pre-fetched data found (timed out or stage cap reached). Marking as error." -Level WARNING
			$onlineModuleVersionsCache[$moduleEntry.ModuleName] = [pscustomobject]@{
				ModuleName    = $moduleEntry.ModuleName
				Stable        = $null
				PreRelease    = $null
				ErrorFetching = "Data not found in pre-fetch cache (timeout)."
				Skipped       = $false
			}
			$preFetchTimeouts++
		}
	}
	$prefetchSync = @{ timeouts = $preFetchTimeouts; completed = ($onlineModuleVersionsCache.Count - $preFetchTimeouts); total = $moduleDataArray.Count }
	$preFetchDuration = (Get-Date) - $overallOperationStartTime
	$results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
	$comparisonErrors = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
	$NewLogDef = ${function:New-Log}.ToString()
	$CompareModuleVersionDef = ${function:Compare-ModuleVersion}.ToString()
	$moduleDataArray | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
		$script:ErrorActionPreference = 'Continue'
		$moduleData = $_
		$moduleName = $moduleData.ModuleName
		$comparisonErrors = $using:comparisonErrors
		$VerbosePreference = $using:VerbosePreference
		${function:New-Log} = $using:NewLogDef
		${function:Compare-ModuleVersion} = $using:CompareModuleVersionDef
		$results = $using:results
		$matchAuthor = $using:MatchAuthor.IsPresent
		$onlineCache = $using:onlineModuleVersionsCache
		try {
			$highestLocalInstall = $moduleData.HighestLocalInstall
			if (-not $highestLocalInstall) {
				New-Log "[$moduleName] Pre-processed HighestLocalInstall object missing. Skipping." -Level WARNING
				$comparisonErrors.Add("$moduleName`: missing local version data")
				return
			}
			$onlineModuleData = $null
			if (-not $onlineCache.TryGetValue($moduleName, [ref]$onlineModuleData)) {
				New-Log "[$moduleName] Could not retrieve pre-fetched data from cache. Skipping." -Level WARNING
				$comparisonErrors.Add("$moduleName`: lookup cache entry missing")
				return
			}
			if ($onlineModuleData.Skipped -or (-not $onlineModuleData.Stable -and -not $onlineModuleData.PreRelease -and $onlineModuleData.ErrorFetching)) {
				if ($onlineModuleData.Skipped) {
					New-Log "[$moduleName] Skipped by repository exclusion." -Level VERBOSE
				}
				else {
					$comparisonErrors.Add("$moduleName`: $($onlineModuleData.ErrorFetching)")
					New-Log "[$moduleName] Repository lookup failed: $($onlineModuleData.ErrorFetching)" -Level WARNING
				}
				return
			}
			$stableModule = $onlineModuleData.Stable
			$preReleaseModule = $onlineModuleData.PreRelease
			$galleryModule = $null
			if ($stableModule -and $preReleaseModule) {
				$stableVerStr = $stableModule.Version.ToString()
				$prereleaseVerStr = $preReleaseModule.Version.ToString()
				$prereleaseLbl = $preReleaseModule.PSObject.Properties['PreRelease'].Value
				$fullPrereleaseVerStr = if (-not [string]::IsNullOrEmpty($prereleaseLbl)) { "$prereleaseVerStr-$prereleaseLbl" } else { $prereleaseVerStr }
				try {
					$isPreReleaseNewer = Compare-ModuleVersion -VersionA $stableVerStr -VersionB $fullPrereleaseVerStr -ReturnBoolean
					$galleryModule = if ($isPreReleaseNewer) { $preReleaseModule } else { $stableModule }
				}
				catch {
					$galleryModule = $stableModule
					$comparisonErrors.Add("$moduleName`: could not compare stable and prerelease versions")
					New-Log "[$moduleName] Could not compare stable '$stableVerStr' with prerelease '$fullPrereleaseVerStr'; using stable" -Level ERROR
				}
			}
			elseif ($preReleaseModule) {
				$galleryModule = $preReleaseModule
			}
			elseif ($stableModule) {
				$galleryModule = $stableModule
			}
			if (-not $galleryModule) {
				return
			}
			[string]$latestOnlineStr = if ($galleryModule.PSObject.Properties['PreRelease'].Value) {
				"$($galleryModule.Version)-$($galleryModule.PSObject.Properties['PreRelease'].Value)".Trim()
			}
			else {
				"$($galleryModule.Version)".Trim()
			}
			[string]$highestLocalStr = if ($highestLocalInstall.IsPrerelease -and $highestLocalInstall.PreReleaseLabel) {
				"$($highestLocalInstall.ModuleVersion)-$($highestLocalInstall.PreReleaseLabel)"
			}
			else {
				"$($highestLocalInstall.ModuleVersion)"
			}
			if ([string]::IsNullOrWhiteSpace($latestOnlineStr) -or [string]::IsNullOrWhiteSpace($highestLocalStr)) {
				New-Log "[$moduleName] Invalid online ('$latestOnlineStr') or local ('$highestLocalStr') version string. Skipping." -Level WARNING
				$comparisonErrors.Add("$moduleName`: invalid version data")
				return
			}
			$needsOverallUpdate = $false
			try {
				if (Compare-ModuleVersion -VersionA $highestLocalStr -VersionB $latestOnlineStr -ReturnBoolean) {
					$needsOverallUpdate = $true
				}
			}
			catch {
				New-Log "[$moduleName] Error comparing versions $highestLocalStr and $latestOnlineStr." -Level ERROR
				$comparisonErrors.Add("$moduleName`: version comparison failed")
				return
			}
			if ($needsOverallUpdate -and $matchAuthor) {
				$localAuthor = $highestLocalInstall.Author
				$galleryAuthor = $galleryModule.Author
				$authorsMatch = $false
				$normalizedLocalAuthor = [Regex]::Replace([string]$localAuthor, '[^a-zA-Z0-9]', '')
				$normalizedGalleryAuthor = [Regex]::Replace([string]$galleryAuthor, '[^a-zA-Z0-9]', '')
				if ($normalizedLocalAuthor -and $normalizedGalleryAuthor -and $normalizedGalleryAuthor -eq $normalizedLocalAuthor) {
					$authorsMatch = $true
				}
				if (-not $authorsMatch) {
					New-Log "[$moduleName] Skipping update: -MatchAuthor specified and authors do not match (Local: '$localAuthor', Online: '$galleryAuthor')." -Level VERBOSE
					$needsOverallUpdate = $false
				}
			}
			if ($needsOverallUpdate) {
				$outdatedInstallationsDetailed = @()
				$allLocalInstalls = $moduleData.AllParsedVersions
				$installsByPath = $allLocalInstalls | Group-Object -Property BasePath
				foreach ($pathGroup in $installsByPath) {
					$versionsInThisPath = $pathGroup.Group
					$latestOnlineVersionFoundInThisPath = $false
					foreach ($installedVersionEntry in $versionsInThisPath) {
						if ($installedVersionEntry.ModuleVersionString -eq $latestOnlineStr) {
							$latestOnlineVersionFoundInThisPath = $true; break
						}
					}
					if (-not $latestOnlineVersionFoundInThisPath) {
						foreach ($outdatedInstall in $versionsInThisPath) {
							$outdatedInstallationsDetailed += [PSCustomObject]@{
								Path             = $outdatedInstall.BasePath
								InstalledVersion = $outdatedInstall.ModuleVersionString
							}
						}
					}
				}
				if ($outdatedInstallationsDetailed.Count -gt 0) {
					$uniqueOutdatedModules = $outdatedInstallationsDetailed | Sort-Object Path, InstalledVersion -Unique
					$resultObject = [PSCustomObject]@{
						ModuleName          = $moduleName
						Repository          = $galleryModule.Repository
						IsPreview           = if ($galleryModule.PSObject.Properties['PreRelease'].Value) { $true } else { $false }
						PreReleaseVersion   = $galleryModule.PSObject.Properties['PreRelease'].Value
						HighestLocalVersion = $highestLocalInstall.ModuleVersion
						LatestVersion       = [version]($galleryModule.Version.ToString() -replace '-.*$', '')
						LatestVersionString = $latestOnlineStr
						OutdatedModules     = $uniqueOutdatedModules
						Author              = $highestLocalInstall.Author
						GalleryAuthor       = $galleryModule.Author
					}
					$results.Add($resultObject)
					New-Log "[$moduleName] Update found: Local '$highestLocalStr' -> Online '$latestOnlineStr'. $($uniqueOutdatedModules.Count) outdated paths." -Level SUCCESS
				}
				else {
					New-Log "[$moduleName] Version '$latestOnlineStr' is newer, but no installation path requiring it was found." -Level WARNING
				}
			}
		}
		catch {
			$comparisonErrors.Add("$moduleName`: $($_.Exception.Message)")
			New-Log "[$moduleName] Unhandled comparison error." -Level ERROR
		}
	}
	$moduleObjects = @($results)
	$finalOverallTime = (Get-Date) - $overallOperationStartTime
	$comparisonDuration = $finalOverallTime - $preFetchDuration
	$updateCheckLevel = if ($preFetchTimeouts -gt 0 -or $comparisonErrors.Count -gt 0) { 'WARNING' } else { 'SUCCESS' }
	New-Log "Update check complete: $validModuleCountForProcessing module(s) checked in $([math]::Round($finalOverallTime.TotalSeconds, 2)) seconds (lookup $([math]::Round($preFetchDuration.TotalSeconds, 2))s, comparison $([math]::Round($comparisonDuration.TotalSeconds, 2))s); found $($moduleObjects.Count) update(s), $preFetchTimeouts timeout(s), and $($comparisonErrors.Count) comparison error(s)." -Level $updateCheckLevel
	if ($prefetchSync.timeouts -gt 0) {
		New-Log "$($prefetchSync.timeouts) module pre-fetch checks timed out." -Level WARNING
	}
	if ($comparisonErrors.Count -gt 0) {
		New-Log "Comparison errors: $(@($comparisonErrors) -join '; ')" -Level WARNING
	}
	return $moduleObjects | Sort-Object ModuleName
}
#endregion Get-ModuleUpdateStatus
#region Update-Modules
function Update-Modules {
	<#
    .SYNOPSIS
    Installs available module updates.
    .DESCRIPTION
    Validates update records and ShouldProcess decisions on the caller thread, installs
    approved modules concurrently, and optionally removes older versions sequentially.
    Flat module layouts are preserved unless a locked module policy is selected.
    .PARAMETER OutdatedModules
    Update records returned by Get-ModuleUpdateStatus. Accepts pipeline input.
    .PARAMETER Clean
    Removes older versions after the new version is installed successfully.
    .PARAMETER UseProgressBar
    Displays progress while installation results are processed.
    .PARAMETER PreRelease
    Allows prerelease targets supplied by the input records.
    .PARAMETER ReplaceLockedModuleOnReboot
    For a locked flat module, stages replacement files and schedules the in-place update
    for the next Windows restart.
    .PARAMETER AllowVersionedSubfolder
    For a locked flat module, installs the update in a versioned subfolder instead of
    preserving the flat layout.
    .INPUTS
    System.Management.Automation.PSCustomObject.
    .OUTPUTS
    System.Management.Automation.PSCustomObject. Each result reports updated, failed,
    cleaned, and pending-reboot paths.
    .EXAMPLE
    Get-ModuleUpdateStatus -ModuleInventory $inventory |
        Update-Modules -Clean -ReplaceLockedModuleOnReboot
    Installs updates, cleans older versions, and schedules locked flat modules for restart.
    .NOTES
    ReplaceLockedModuleOnReboot and AllowVersionedSubfolder are mutually exclusive.
    Administrative rights may be required for system module paths and reboot scheduling.
    .LINK
    Get-ModuleUpdateStatus
    #>
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[Parameter(ValueFromPipeline, Mandatory)][Object[]]$OutdatedModules,
		[switch]$Clean,
		[switch]$UseProgressBar,
		[switch]$PreRelease,
		[switch]$ReplaceLockedModuleOnReboot,
		[switch]$AllowVersionedSubfolder
	)
	begin {
		if ($ReplaceLockedModuleOnReboot -and $AllowVersionedSubfolder) {
			throw "Parameters -ReplaceLockedModuleOnReboot and -AllowVersionedSubfolder are mutually exclusive. Specify at most one (or neither)."
		}
		$aggregateResults = [System.Collections.Generic.List[object]]::new()
		$batchModules = @()
		$updateStartTime = Get-Date
	}
	process {
		foreach ($module in $OutdatedModules) {
			$batchModules += $module
		}
	}
	end {
		if ($batchModules.Count -eq 0) {
			New-Log "No modules provided for update. Exiting."
			return
		}
		$total = $batchModules.Count
		New-Log "Preparing $total module update(s)..."
		$modulesToInstall = [System.Collections.Generic.List[object]]::new()
		$current = 0
		foreach ($module in $batchModules) {
			$moduleName = $module.ModuleName
			$current++
			[string]$targetVersionString = $($module.LatestVersion)
			[version]$latestVer = $null
			if ($module.isPreView) {
				$parsedTargetVersion = Parse-ModuleVersion -VersionString "$($targetVersionString)-$($module.PreReleaseVersion)" -ErrorAction SilentlyContinue
				[string]$preReleaseVersion = $parsedTargetVersion.PreReleaseVersion
			}
			else {
				$parsedTargetVersion = Parse-ModuleVersion -VersionString $targetVersionString -ErrorAction SilentlyContinue
				[string]$preReleaseVersion = $null
			}
			if (-not $parsedTargetVersion -or -not $parsedTargetVersion.ModuleVersion) {
				New-Log "[$current/$total] Skipping module [$moduleName]: Could not parse Target Version String '$targetVersionString' using Parse-ModuleVersion. Will skip." -Level WARNING
				$aggregateResults.Add([PSCustomObject]@{
						ModuleName           = $moduleName
						NewVersionPreRelease = if ($module.IsPreview) { "$($targetVersionString)-$($module.PreReleaseVersion)" }
						NewVersion           = $targetVersionString
						UpdatedPaths         = @()
						FailedPaths          = @("Version parsing failed: $targetVersionString")
						PendingRebootPaths   = @()
						PendingReboot        = $false
						OverallSuccess       = $false
						CleanedPaths         = @()
						CleanStatus          = if ($Clean.IsPresent) { 'NotRun' } else { 'NotRequested' }
					})
				continue
			}
			$latestVer = $parsedTargetVersion.ModuleVersion
			[string]$baseVerStr = $latestVer.ToString()
			$repository = $module.Repository
			$installAsPreview = $parsedTargetVersion.IsPrerelease
			if ($installAsPreview -and -not $PreRelease.IsPresent) {
				New-Log "[$moduleName] Skipping prerelease '$preReleaseVersion'; specify -PreRelease to install it." -Level WARNING
				$aggregateResults.Add([PSCustomObject]@{
						ModuleName           = $moduleName
						NewVersionPreRelease = $preReleaseVersion
						NewVersion           = $baseVerStr
						UpdatedPaths         = @()
						FailedPaths          = @('Prerelease updates require -PreRelease')
						PendingRebootPaths   = @()
						PendingReboot        = $false
						OverallSuccess       = $false
						CleanedPaths         = @()
						CleanStatus          = if ($Clean.IsPresent) { 'NotRun' } else { 'NotRequested' }
					})
				continue
			}
			$outdatedPaths = @($module.OutdatedModules | Where-Object { $null -ne $_.Path } | Select-Object -ExpandProperty Path -Unique | Where-Object { $_ -and (Test-Path $_ -PathType Container -Verbose:$false) })
			$outdatedVersions = @($module.OutdatedModules.InstalledVersion | Select-Object -Unique)
			if ($outdatedPaths.Count -eq 0) {
				New-Log "[$moduleName][$current/$total] Skipping module: No valid outdated base paths found where old versions exist. (Checked: $($module.OutdatedModules.Path -join '; '))." -Level WARNING
				$aggregateResults.Add([PSCustomObject]@{
						ModuleName           = $moduleName
						NewVersionPreRelease = if ($module.IsPreview) { $preReleaseVersion }
						NewVersion           = $baseVerStr
						UpdatedPaths         = @()
						FailedPaths          = @("No valid source paths provided or accessible")
						PendingRebootPaths   = @()
						PendingReboot        = $false
						OverallSuccess       = $false
						CleanedPaths         = @()
						CleanStatus          = if ($Clean.IsPresent) { 'NotRun' } else { 'NotRequested' }
					})
				continue
			}
			New-Log "[$moduleName] Target base paths based on outdated locations: $($outdatedPaths -join '; ')" -Level VERBOSE
			$displayVersion = if ($installAsPreview -and $preReleaseVersion) { $preReleaseVersion } else { $targetVersionString }
			# ShouldProcess must run before work is dispatched to parallel runspaces.
			if ($PSCmdlet.ShouldProcess("$moduleName v$displayVersion", "Install from repository '$repository' to paths: $($outdatedPaths -join ', ')")) {
				$modulesToInstall.Add([PSCustomObject]@{
						ModuleName          = $moduleName
						TargetVersionString = $targetVersionString
						PreReleaseVersion   = $preReleaseVersion
						BaseVerStr          = $baseVerStr
						Repository          = $repository
						InstallAsPreview    = $installAsPreview
						OutdatedPaths       = $outdatedPaths
						OutdatedVersions    = $outdatedVersions
						LatestVer           = $latestVer
						IsPreview           = [bool]$module.IsPreview
						CleanApproved       = if ($Clean.IsPresent) {
							$PSCmdlet.ShouldProcess("$moduleName (Versions: $($outdatedVersions -join ', '))", "Remove from paths: $($outdatedPaths -join ', ')")
						}
						else { $false }
					})
			}
			else {
				New-Log "[$moduleName][$current/$total] Skipped update due to ShouldProcess user choice." -Level WARNING
				$aggregateResults.Add([PSCustomObject]@{
						ModuleName           = $moduleName
						NewVersionPreRelease = if ($module.IsPreview) { $preReleaseVersion }
						NewVersion           = $baseVerStr
						UpdatedPaths         = @()
						FailedPaths          = @("Skipped by ShouldProcess")
						PendingRebootPaths   = @()
						PendingReboot        = $false
						OverallSuccess       = $false
						CleanedPaths         = @()
						CleanStatus          = if ($Clean.IsPresent) { 'NotRun' } else { 'NotRequested' }
					})
			}
		}
		if ($modulesToInstall.Count -eq 0) {
			New-Log "No modules approved for installation after pre-processing. Exiting."
			return $aggregateResults
		}
		$installThrottleLimit = [Math]::Min($modulesToInstall.Count, [Math]::Max(4, [System.Environment]::ProcessorCount * 2))
		New-Log "Installing $($modulesToInstall.Count) module update(s) in parallel (throttle $installThrottleLimit)..."
		$parallelInstallResults = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
		$parallelStartTime = Get-Date
		$NewLogDef = ${function:New-Log}.ToString()
		$InstallPSModuleDef = ${function:Install-PSModule}.ToString()
		$replaceLockedOnReboot = $ReplaceLockedModuleOnReboot.IsPresent
		$allowVersionedSubfolder = $AllowVersionedSubfolder.IsPresent
		$useClean = $Clean.IsPresent
		$modulesToInstall | ForEach-Object -ThrottleLimit $installThrottleLimit -Parallel {
			$script:ErrorActionPreference = 'Continue'
			$preparedModule = $_
			$moduleName = $preparedModule.ModuleName
			$VerbosePreference = $using:VerbosePreference
			$parallelInstallResults = $using:parallelInstallResults
			${function:New-Log} = $using:NewLogDef
			${function:Install-PSModule} = $using:InstallPSModuleDef
			Import-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Global -Force -Verbose:$false
			try {
				New-Log "[$moduleName] Installing version [$($preparedModule.TargetVersionString)]..." -Level VERBOSE
				$installResult = Install-PSModule -ModuleName $moduleName -TargetVersionString $preparedModule.TargetVersionString -PreReleaseVersion $preparedModule.PreReleaseVersion -RepositoryName $preparedModule.Repository -IsPreview $preparedModule.InstallAsPreview -Destinations $preparedModule.OutdatedPaths -ReplaceLockedModuleOnReboot $using:replaceLockedOnReboot -AllowVersionedSubfolder $using:allowVersionedSubfolder -ErrorAction SilentlyContinue
				$pendingRebootPaths = @($installResult.PendingRebootPaths)
				$finalResult = [PSCustomObject]@{
					ModuleName           = $moduleName
					NewVersionPreRelease = if ($preparedModule.IsPreview) { $preparedModule.PreReleaseVersion }
					NewVersion           = $preparedModule.BaseVerStr
					UpdatedPaths         = @($installResult.UpdatedPaths)
					FailedPaths          = if ($installResult.FailedPaths) { @($installResult.FailedPaths) } else { @() }
					PendingRebootPaths   = $pendingRebootPaths
					PendingReboot        = ($pendingRebootPaths.Count -gt 0)
					OverallSuccess       = (@($installResult.FailedPaths).Count -eq 0 -and @($installResult.UpdatedPaths).Count -gt 0)
					CleanedPaths         = @()
					CleanStatus          = 'NotRequested'
					_OutdatedVersions    = $preparedModule.OutdatedVersions
					_LatestVer           = $preparedModule.LatestVer
					_PreReleaseVersion   = $preparedModule.PreReleaseVersion
					_CleanApproved       = $preparedModule.CleanApproved
				}
				$parallelInstallResults.Add($finalResult)
				if ($finalResult.PendingReboot) {
					New-Log "[$moduleName] Staged $($preparedModule.TargetVersionString) for replacement on restart at: $($finalResult.PendingRebootPaths -join '; ')." -Level WARNING
				}
				elseif ($finalResult.OverallSuccess) {
					New-Log "[$moduleName] Installed $($preparedModule.TargetVersionString) to $(@($finalResult.UpdatedPaths).Count) path(s)." -Level SUCCESS
				}
				elseif (@($finalResult.UpdatedPaths).Count -gt 0) {
					New-Log "[$moduleName] Partially installed $($preparedModule.TargetVersionString); failed paths: $($finalResult.FailedPaths -join '; ')." -Level WARNING
				}
				else {
					New-Log "[$moduleName] Installation failed for: $($finalResult.FailedPaths -join '; ')" -Level WARNING
				}
			}
			catch {
				New-Log "[$moduleName] Unhandled error during parallel install." -Level ERROR
				$parallelInstallResults.Add([PSCustomObject]@{
						ModuleName           = $moduleName
						NewVersionPreRelease = if ($preparedModule.IsPreview) { $preparedModule.PreReleaseVersion }
						NewVersion           = $preparedModule.BaseVerStr
						UpdatedPaths         = @()
						FailedPaths          = @("Parallel install error: $($_.Exception.Message)")
						PendingRebootPaths   = @()
						PendingReboot        = $false
						OverallSuccess       = $false
						CleanedPaths         = @()
						CleanStatus          = if ($using:useClean) { 'NotRun' } else { 'NotRequested' }
						_OutdatedVersions    = $preparedModule.OutdatedVersions
						_LatestVer           = $preparedModule.LatestVer
						_PreReleaseVersion   = $preparedModule.PreReleaseVersion
						_CleanApproved       = $false
					})
			}
		}
		$parallelDuration = (Get-Date) - $parallelStartTime
		$installResultsArray = @($parallelInstallResults)
		$installPendingCount = @($installResultsArray | Where-Object PendingReboot).Count
		$installSuccessCount = @($installResultsArray | Where-Object { $_.OverallSuccess -and -not $_.PendingReboot }).Count
		$installFailureCount = $installResultsArray.Count - $installSuccessCount - $installPendingCount
		New-Log "Installation complete in $([math]::Round($parallelDuration.TotalSeconds, 2)) seconds: $installSuccessCount succeeded, $installPendingCount pending reboot, $installFailureCount failed or partial."
		$postProcessIndex = 0
		foreach ($finalResult in $installResultsArray) {
			$moduleName = $finalResult.ModuleName
			$postProcessIndex++
			$finalResult.CleanedPaths = @($finalResult.CleanedPaths)
			$finalResult.CleanStatus = if ($useClean) { 'NotRun' } else { 'NotRequested' }
			if ($UseProgressBar.IsPresent) {
				$progressParams = @{
					Activity         = "Updating PowerShell Modules"
					Status           = "Post-processing: $moduleName"
					PercentComplete  = [math]::Round(($postProcessIndex / $installResultsArray.Count) * 100)
					CurrentOperation = if ($useClean -and $finalResult.OverallSuccess) { "[$moduleName] Cleaning old versions.." } else { "[$moduleName] Finalizing.." }
				}
				Write-Progress @progressParams
			}
			if ($finalResult.OverallSuccess -and $useClean) {
				if ($finalResult._CleanApproved) {
					$cleanedPathsResult = @(Remove-OutdatedVersions -ModuleName $moduleName -ModuleBasePaths $finalResult.UpdatedPaths -LatestVersion $finalResult._LatestVer -PreReleaseVersion $finalResult._PreReleaseVersion -ErrorAction SilentlyContinue | Where-Object { $_ -is [string] })
					if ($cleanedPathsResult.Count -gt 0) {
						$finalResult.CleanedPaths = $cleanedPathsResult
						$finalResult.CleanStatus = 'Succeeded'
						New-Log "[$moduleName] Successfully cleaned $($cleanedPathsResult.Count) old items: $($cleanedPathsResult -join '; ')" -Level SUCCESS
					}
					else {
						$finalResult.CleanStatus = 'NoOldVersions'
						New-Log "[$moduleName] No old version folders required removal." -Level VERBOSE
					}
				}
				else {
					$finalResult.CleanStatus = 'Skipped'
					New-Log "[$moduleName] Skipped cleaning due to ShouldProcess user choice." -Level WARNING
				}
			}
			elseif ($useClean -and -not $finalResult.OverallSuccess) {
				$finalResult.CleanStatus = 'NotRun'
				New-Log "[$moduleName] Skipping cleaning as the update was not fully successful (Failed Paths: $($finalResult.FailedPaths -join '; '))." -Level VERBOSE
			}
			$finalResult.PSObject.Properties.Remove('_OutdatedVersions')
			$finalResult.PSObject.Properties.Remove('_LatestVer')
			$finalResult.PSObject.Properties.Remove('_PreReleaseVersion')
			$finalResult.PSObject.Properties.Remove('_CleanApproved')
			$aggregateResults.Add($finalResult)
		}
		$pendingRebootCount = ($aggregateResults | Where-Object PendingReboot).Count
		$successCount = ($aggregateResults | Where-Object { $_.OverallSuccess -and -not $_.PendingReboot }).Count
		$failCount = $total - $successCount - $pendingRebootCount
		$totalUpdateDuration = (Get-Date) - $updateStartTime
		$updateSummaryLevel = if ($failCount -gt 0 -or $pendingRebootCount -gt 0) { 'WARNING' } else { 'SUCCESS' }
		New-Log "Module update complete: $total processed in $([math]::Round($totalUpdateDuration.TotalSeconds, 2)) seconds; $successCount succeeded, $pendingRebootCount pending reboot, $failCount failed or partial." -Level $updateSummaryLevel
		if ($pendingRebootCount -gt 0) {
			New-Log "$pendingRebootCount module(s) staged; the flat update will be applied on the next reboot: $((($aggregateResults | Where-Object PendingReboot).ModuleName) -join ', ')" -Level WARNING
		}
		if ($failCount -gt 0) {
			New-Log "Modules with failures or partial updates:" -Level WARNING
			$aggregateResults | Where-Object {
				-not $_.OverallSuccess -and -not $_.PendingReboot
			} | ForEach-Object {
				$failReason = if ($_.FailedPaths) {
					"[$($_.ModuleName)] Failed Paths: $($_.FailedPaths -join '; ')"
				}
				else {
					"[$($_.ModuleName)] Unknown reason"
				}
				New-Log "$failReason" -Level WARNING
			}
		}
		if ($Clean.IsPresent) {
			$cleanedModules = @($aggregateResults | Where-Object CleanStatus -EQ 'Succeeded')
			$cleanNotNeeded = @($aggregateResults | Where-Object CleanStatus -EQ 'NoOldVersions')
			$cleanSkipped = @($aggregateResults | Where-Object CleanStatus -EQ 'Skipped')
			$cleanNotRun = @($aggregateResults | Where-Object CleanStatus -EQ 'NotRun')
			New-Log "Cleanup summary: $($cleanedModules.Count) cleaned, $($cleanNotNeeded.Count) already clean, $($cleanSkipped.Count) skipped, $($cleanNotRun.Count) not run."
		}
		return $aggregateResults
	}
}
#endregion Update-Modules
#region Install-PSModule
function Install-PSModule {
	<#
    .SYNOPSIS
    Installs one module version into the requested locations.
    .DESCRIPTION
    Uses Save-PSResource for each destination and falls back to Install-PSResource or
    Install-Module when needed. Flat installations are updated in place. Locked flat
    installations can be staged for restart or redirected to a versioned subfolder.
    .PARAMETER ModuleName
    Name of the module to install.
    .PARAMETER TargetVersionString
    Exact version to install.
    .PARAMETER RepositoryName
    Registered repository that provides the module.
    .PARAMETER IsPreview
    Indicates that the target is a prerelease.
    .PARAMETER Destinations
    Module base paths that require the target version.
    .PARAMETER PreReleaseVersion
    Full prerelease version string when the target is a prerelease.
    .PARAMETER ReplaceLockedModuleOnReboot
    Schedules replacement of locked flat-module files for the next Windows restart.
    .PARAMETER AllowVersionedSubfolder
    Installs a locked flat module into a versioned subfolder.
    .INPUTS
    None.
    .OUTPUTS
    System.Management.Automation.PSCustomObject with UpdatedPaths, FailedPaths, and
    PendingRebootPaths properties.
    .EXAMPLE
    Install-PSModule -ModuleName Pester -TargetVersionString 5.7.1 -RepositoryName PSGallery -IsPreview $false -Destinations 'C:\Program Files\PowerShell\Modules\Pester'
    Installs Pester 5.7.1 into the requested module location.
    .NOTES
    Internal helper used by Update-Modules. Reboot scheduling is Windows-specific.
    .LINK
    Save-PSResource
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$ModuleName,
		[Parameter(Mandatory)][string]$TargetVersionString,
		[Parameter(Mandatory)][string]$RepositoryName,
		[Parameter(Mandatory)][bool]$IsPreview,
		[Parameter(Mandatory)][string[]]$Destinations,
		[string]$PreReleaseVersion,
		[bool]$ReplaceLockedModuleOnReboot = $false,
		[bool]$AllowVersionedSubfolder = $false
	)
	function Test-FlatTargetLocked {
		param(
			[Parameter(Mandatory)][string]$BasePath
		)
		foreach ($f in @(Get-ChildItem -LiteralPath $BasePath -Recurse -File -Force -ErrorAction SilentlyContinue)) {
			try { $fs = [System.IO.File]::Open($f.FullName, 'Open', 'ReadWrite', 'None'); $fs.Close(); $fs.Dispose() }
			catch {
				$ex = $_.Exception
				while ($ex) {
					if ($ex -is [System.IO.IOException] -and $ex -isnot [System.IO.FileNotFoundException]) { return $true }
					$ex = $ex.InnerException
				}
			}
		}
		return $false
	}
	function Register-RebootFlatReplace {
		param(
			[Parameter(Mandatory)][string]$StagingDir,
			[Parameter(Mandatory)][string]$DestRoot
		)
		if (-not ('Win32UpdMod.PendingMove' -as [type])) {
			Add-Type -Namespace 'Win32UpdMod' -Name 'PendingMove' -MemberDefinition '[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Unicode)] public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);' -ErrorAction Stop
		}
		$flags = 5 # MOVEFILE_REPLACE_EXISTING | MOVEFILE_DELAY_UNTIL_REBOOT
		$base = $StagingDir.TrimEnd('\')
		$scheduled = 0; $failed = 0
		foreach ($f in @(Get-ChildItem -LiteralPath $StagingDir -Recurse -File -Force -ErrorAction SilentlyContinue)) {
			$rel = $f.FullName.Substring($base.Length).TrimStart('\')
			$dest = Join-Path $DestRoot $rel
			$destDir = Split-Path $dest -Parent
			if ($destDir -and -not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force -ErrorAction SilentlyContinue | Out-Null }
			if ([Win32UpdMod.PendingMove]::MoveFileEx($f.FullName, $dest, $flags)) { $scheduled++ } else { $failed++ }
		}
		return [pscustomobject]@{ Scheduled = $scheduled; Failed = $failed }
	}
	$result = [PSCustomObject]@{
		UpdatedPaths       = [System.Collections.Generic.List[string]]::new()
		FailedPaths        = [System.Collections.Generic.List[string]]::new()
		PendingRebootPaths = [System.Collections.Generic.List[string]]::new()
	}
	if ($PreReleaseVersion) {
		$TargetVersionStringOrig = $TargetVersionString
		$TargetVersionString = $PreReleaseVersion
	}
	else {
		$TargetVersionStringOrig = $TargetVersionString
	}
	$commonSaveParams = @{
		Name                = $ModuleName
		Version             = $TargetVersionString
		Repository          = $RepositoryName
		TrustRepository     = $true
		IncludeXml          = $true
		SkipDependencyCheck = $true
		AcceptLicense       = $true
		Confirm             = $false
		PassThru            = $true
		Verbose             = $false
		ErrorAction         = 'Stop'
		WarningAction       = 'SilentlyContinue'
	}
	if ($IsPreview) { $commonSaveParams['PreRelease'] = $true }
	$pathsToRetry = [System.Collections.Generic.List[string]]::new()
	foreach ($destinationBasePath in $Destinations) {
		# Preserve flat installs by staging the target version and copying it into the module root.
		if (Test-Path -LiteralPath (Join-Path $destinationBasePath "$ModuleName.psd1") -PathType Leaf) {
			$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("UpdMod_{0}" -f [guid]::NewGuid().ToString('N'))
			$newContentDir = $null; $savedItemFlat = $null
			try {
				$null = New-Item -ItemType Directory -Path $tempRoot -Force -ErrorAction Stop
				$savedItemFlat = (Save-PSResource @commonSaveParams -Path $tempRoot) | Select-Object -Last 1
				$savedModuleDir = Join-Path $tempRoot $ModuleName
				$newContentDir = Get-ChildItem -LiteralPath $savedModuleDir -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d+(\.\d+)' } | Sort-Object Name | Select-Object -Last 1
			}
			catch { New-Log "[$moduleName] Flat update: Save-PSResource to temp failed." -Level ERROR }
			$expPreF = if ($PreReleaseVersion) { ($PreReleaseVersion -split '-')[-1] } else { '' }
			$actPreF = if ($savedItemFlat -and $savedItemFlat.Prerelease) { $savedItemFlat.Prerelease } else { '' }
			if ($newContentDir -and $savedItemFlat -and "$($savedItemFlat.Version)" -eq "$TargetVersionStringOrig" -and $actPreF -eq $expPreF) {
				if (-not (Test-FlatTargetLocked -BasePath $destinationBasePath)) {
					try {
						Copy-Item -Path (Join-Path $newContentDir.FullName '*') -Destination $destinationBasePath -Recurse -Force -ErrorAction Stop
						New-Log "[$moduleName] Updated FLAT (unversioned) install at '$destinationBasePath' to [$TargetVersionString] in place (no version subfolder)." -Level SUCCESS
						$result.UpdatedPaths.Add($destinationBasePath)
					}
					catch { New-Log "[$moduleName] Flat copy into '$destinationBasePath' failed." -Level ERROR; $result.FailedPaths.Add($destinationBasePath) }
				}
				elseif ($ReplaceLockedModuleOnReboot) {
					$stagingDir = Join-Path $env:ProgramData ("UpdateModules\PendingReboot\{0}_{1}" -f $ModuleName, [guid]::NewGuid().ToString('N'))
					try {
						$null = New-Item -ItemType Directory -Path (Split-Path $stagingDir -Parent) -Force -ErrorAction SilentlyContinue
						Move-Item -LiteralPath $newContentDir.FullName -Destination $stagingDir -Force -ErrorAction Stop
						$sched = Register-RebootFlatReplace -StagingDir $stagingDir -DestRoot $destinationBasePath
						if ($sched.Scheduled -gt 0 -and $sched.Failed -eq 0) {
							New-Log "[$moduleName] '$ModuleName' is in use; scheduled $($sched.Scheduled) file(s) to replace the flat install at '$destinationBasePath' on next reboot (no version subfolder created)." -Level SUCCESS
							$result.PendingRebootPaths.Add($destinationBasePath)
						}
						else {
							New-Log "[$moduleName] Could not schedule reboot replacement for '$destinationBasePath' (scheduled=$($sched.Scheduled), failed=$($sched.Failed)). Administrator rights are required." -Level WARNING
							$result.FailedPaths.Add($destinationBasePath)
							if (Test-Path -LiteralPath $stagingDir) { Remove-Item -LiteralPath $stagingDir -Recurse -Force -ErrorAction SilentlyContinue }
						}
					}
					catch {
						New-Log "[$moduleName] Reboot-replace staging/scheduling for '$destinationBasePath' failed." -Level ERROR
						$result.FailedPaths.Add($destinationBasePath)
						if (Test-Path -LiteralPath $stagingDir) { Remove-Item -LiteralPath $stagingDir -Recurse -Force -ErrorAction SilentlyContinue }
					}
				}
				elseif ($AllowVersionedSubfolder) {
					$saveParentF = Split-Path $destinationBasePath -Parent -ErrorAction SilentlyContinue
					try {
						$svItem = (Save-PSResource @commonSaveParams -Path $saveParentF) | Select-Object -Last 1
						if ($svItem -and "$($svItem.Version)" -eq "$TargetVersionStringOrig") {
							New-Log "[$moduleName] '$ModuleName' is in use; installed [$TargetVersionString] into a version subfolder under '$destinationBasePath' (-AllowVersionedSubfolder). The flat layout could not be kept while the module is loaded." -Level WARNING
							$result.UpdatedPaths.Add($destinationBasePath)
						}
						else { $result.FailedPaths.Add($destinationBasePath) }
					}
					catch { New-Log "[$moduleName] Versioned-subfolder fallback for '$destinationBasePath' failed." -Level ERROR; $result.FailedPaths.Add($destinationBasePath) }
				}
				else {
					New-Log "[$moduleName] Cannot update '$destinationBasePath' in place: '$ModuleName' is currently in use (files locked) and its layout is flat (unversioned). Skipped to preserve the layout. Re-run from a session that does not load '$ModuleName', or pass -ReplaceLockedModuleOnReboot or -AllowVersionedSubfolder." -Level WARNING
					$result.FailedPaths.Add($destinationBasePath)
				}
			}
			else {
				New-Log "[$moduleName] Flat update: could not obtain version [$TargetVersionString] from '$RepositoryName' for '$destinationBasePath' (saved='$($savedItemFlat.Version)'). Marking path as failed." -Level WARNING
				$result.FailedPaths.Add($destinationBasePath)
			}
			if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
			continue
		}
		$saveTargetDir = Split-Path $destinationBasePath -Parent -ErrorAction SilentlyContinue
		if (-not $saveTargetDir -or -not (Test-Path $saveTargetDir -PathType Container)) {
			New-Log "[$moduleName] Install-PSModule: Cannot determine valid parent directory from destination '$destinationBasePath'. Skipping Save-PSResource for this path." -Level WARNING
			$pathsToRetry.Add($destinationBasePath)
			continue
		}
		New-Log "[$moduleName] Attempting Save-PSResource for version [$TargetVersionString] to '$saveTargetDir'..." -Level VERBOSE
		$saveRes = $null
		try {
			$saveRes = Save-PSResource @commonSaveParams -Path $saveTargetDir
		}
		catch {
			New-Log "[$moduleName] Save-PSResource failed for v$TargetVersionString to '$saveTargetDir'" -Level ERROR
			$saveRes = $null
		}
		$savedItem = $saveRes | Select-Object -Last 1
		$expectedPrerelease = if ($PreReleaseVersion) { ($PreReleaseVersion -split '-')[-1] } else { '' }
		$actualPrerelease = if ($savedItem -and $savedItem.Prerelease) { $savedItem.Prerelease } else { '' }
		if ($savedItem -and "$($savedItem.Version)" -eq "$TargetVersionStringOrig" -and $actualPrerelease -eq $expectedPrerelease) {
			$result.UpdatedPaths.Add($destinationBasePath)
		}
		else {
			$pathsToRetry.Add($destinationBasePath)
		}
	}
	if ($pathsToRetry.Count -gt 0) {
		$fallbackInstallSucceeded = $false
		$installedLocationFallback = $null
		try {
			$psResourceParams = @{
				Name                = $ModuleName
				Version             = $TargetVersionString
				Scope               = 'AllUsers'
				AcceptLicense       = $true
				SkipDependencyCheck = $true
				Confirm             = $false
				Reinstall           = $true
				TrustRepository     = $true
				Repository          = $RepositoryName
				PassThru            = $true
				Verbose             = $false
				ErrorAction         = 'Stop'
				WarningAction       = 'SilentlyContinue'
			}
			if ($IsPreview) { $psResourceParams['PreRelease'] = $true }
			New-Log "[$moduleName] Attempting Install-PSResource for version [$($psResourceParams.Version)]..." -Level VERBOSE
			$installRes1 = $null
			$installRes1 = Install-PSResource @psResourceParams
			$installedItem1 = $installRes1 | Select-Object -Last 1
			$expectedPrerelease1 = if ($PreReleaseVersion) { ($PreReleaseVersion -split '-')[-1] } else { '' }
			$actualPrerelease1 = if ($installedItem1 -and $installedItem1.Prerelease) { $installedItem1.Prerelease } else { '' }
			if ($installedItem1 -and "$($installedItem1.Version)" -eq "$TargetVersionStringOrig" -and $actualPrerelease1 -eq $expectedPrerelease1) {
				$installedLocationFallback = $installedItem1.InstalledLocation
				$fallbackInstallSucceeded = $true
			}
		}
		catch {
			New-Log "[$moduleName] Install-PSResource failed for v${TargetVersionString}" -Level ERROR
		}
		if (-not $fallbackInstallSucceeded) {
			try {
				$installModuleParams = @{
					Name               = $ModuleName
					RequiredVersion    = $TargetVersionString
					Scope              = 'AllUsers'
					Force              = $true
					AcceptLicense      = $true
					SkipPublisherCheck = $true
					AllowClobber       = $true
					PassThru           = $true
					Repository         = $RepositoryName
					Verbose            = $false
					ErrorAction        = 'Stop'
					WarningAction      = 'SilentlyContinue'
					Confirm            = $false
				}
				if ($IsPreview) {
					$installModuleCmd = Get-Command -Name 'Install-Module' -ErrorAction SilentlyContinue
					if ($installModuleCmd -and $installModuleCmd.Parameters.ContainsKey('AllowPrerelease')) {
						$installModuleParams['AllowPrerelease'] = $true
					}
				}
				try {
					Set-PSResourceRepository -Name $RepositoryName -Trusted -ErrorAction Stop -Verbose:$false
				}
				catch {
					New-Log "[$moduleName] Could not mark repository '$RepositoryName' as trusted before the Install-Module fallback" -Level ERROR
				}
				New-Log "[$moduleName] Attempting Install-Module for version [$($installModuleParams.RequiredVersion)]..." -Level DEBUG
				$installRes2 = $null
				$installRes2 = Install-Module @installModuleParams
				$installedItem2 = $installRes2 | Select-Object -Last 1
				if ($installedItem2 -and $($installedItem2.Version) -eq $TargetVersionStringOrig -and $installedItem2.PSObject.Properties.Match('Prerelease').Value -eq $(($PreReleaseVersion -split '-')[-1])) {
					$installedLocationFallback = $installedItem2.InstalledLocation
					$fallbackInstallSucceeded = $true
				}
			}
			catch {
				New-Log "[$moduleName] Install-Module (last fallback) failed" -Level ERROR
			}
		}
		if ($fallbackInstallSucceeded -and $installedLocationFallback) {
			$fallbackBaseInstallPath = Split-Path $installedLocationFallback -Parent -Verbose:$false
			foreach ($retryPath in $pathsToRetry) {
				if ($retryPath -eq $fallbackBaseInstallPath) {
					if ($result.UpdatedPaths -notcontains $retryPath) {
						$result.UpdatedPaths.Add($retryPath)
						New-Log "[$moduleName] Fallback installation to '$installedLocationFallback' satisfied the intended destination base path '$retryPath'." -Level DEBUG
					}
				}
			}
		}
		foreach ($retryPath in $pathsToRetry) {
			if ($result.UpdatedPaths -notcontains $retryPath) {
				if ($result.FailedPaths -notcontains $retryPath) {
					$result.FailedPaths.Add($retryPath)
					New-Log "[$moduleName] Marking path '$retryPath' as failed after Save-PSResource and fallback install attempts." -Level Debug
				}
			}
		}
	}
	# Keep empty result collections as empty arrays rather than @($null).
	$result.UpdatedPaths = @($result.UpdatedPaths | Select-Object -Unique -Verbose:$false)
	$result.PendingRebootPaths = @($result.PendingRebootPaths | Select-Object -Unique -Verbose:$false)
	$result.FailedPaths = @($result.FailedPaths | Select-Object -Unique -Verbose:$false)
	$result.FailedPaths = @($result.FailedPaths | Where-Object { $result.UpdatedPaths -notcontains $_ -and $result.PendingRebootPaths -notcontains $_ })
	$updatedCount = @($result.UpdatedPaths).Count
	$pendingCount = @($result.PendingRebootPaths).Count
	if ($updatedCount -eq 0 -and $pendingCount -eq 0) {
		$result.FailedPaths = @($Destinations | Select-Object -Unique)
	}
	return $result
}
#endregion Install-PSModule
#region Remove-OutdatedVersions
function Remove-OutdatedVersions {
	<#
    .SYNOPSIS
    Removes older installed versions of a module.
    .DESCRIPTION
    Preserves the requested version and removes other version folders from each module
    base path. Managed CurrentUser and AllUsers installations use Uninstall-PSResource
    first; other locations are removed directly with retry handling.
    .PARAMETER ModuleName
    Name of the module to clean.
    .PARAMETER ModuleBasePaths
    Module root directories containing version subfolders.
    .PARAMETER LatestVersion
    Numeric version that must be preserved.
    .PARAMETER DoNotClean
    Module names that must never be cleaned.
    .PARAMETER PreReleaseVersion
    Full prerelease version string that must be preserved, when applicable.
    .INPUTS
    None.
    .OUTPUTS
    System.String. Paths of version directories removed successfully.
    .EXAMPLE
    Remove-OutdatedVersions -ModuleName Pester -ModuleBasePaths 'C:\Modules\Pester' -LatestVersion 5.7.1
    Removes Pester version folders other than 5.7.1.
    .NOTES
    Supports ShouldProcess. Modules outside PSResourceGet-managed locations, including
    $PSHOME, bypass Uninstall-PSResource and use direct filesystem cleanup.
    .LINK
    Uninstall-PSResource
    #>
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[Parameter(Mandatory)][string]$ModuleName,
		[Parameter(Mandatory)][string[]]$ModuleBasePaths,
		[Parameter(Mandatory)][version]$LatestVersion,
		[string[]]$DoNotClean = @('PowerShellGet', 'Microsoft.PowerShell.PSResourceGet'),
		[string]$PreReleaseVersion = $null
	)
	if ($ModuleName -in $DoNotClean) {
		New-Log "[$moduleName] Skipping cleaning as it is in the DoNotClean list."
		return @()
	}
	[string]$latestVersionString = $LatestVersion.ToString()
	[string]$latestVersionFullString = if ($PreReleaseVersion) { $PreReleaseVersion } else { $latestVersionString }
	New-Log "[$moduleName] Starting cleanup of old versions (keeping v$latestVersionFullString)..." -Level VERBOSE
	$cleanedItems = [System.Collections.Generic.List[string]]::new()
	Remove-Module -Name $ModuleName -Force -ErrorAction SilentlyContinue
	foreach ($basePath in $ModuleBasePaths) {
		if (-not (Test-Path $basePath -PathType Container)) {
			New-Log "[$moduleName] Base path '$basePath' for cleaning does not exist or is not a directory. Skipping." -Level WARNING
			continue
		}
		$versionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue -Verbose:$false | Where-Object {
			$folderName = $_.Name
			if ($folderName -notmatch '^\d+(\.\d+){1,4}(-.+)?$') { return $false }
			if ($folderName -eq $latestVersionFullString -or $folderName -eq $latestVersionString) { return $false }
			$folderBase = ($folderName -split '-', 2)[0]
			$keepBase = ($latestVersionFullString -split '-', 2)[0]
			$folderVer = $null; $keepVer = $null
			if ([version]::TryParse($folderBase, [ref]$folderVer) -and [version]::TryParse($keepBase, [ref]$keepVer)) {
				if ($folderVer -eq $keepVer) {
					$folderPre = if ($folderName -match '-(.+)$') { $Matches[1] } else { $null }
					$keepPre = if ($latestVersionFullString -match '-(.+)$') { $Matches[1] } else { $null }
					if ($folderPre -eq $keepPre) { return $false }
				}
			}
			return $true
		}
		if ($versionFolders) {
			foreach ($versionFolder in $versionFolders) {
				$folderPath = $versionFolder.FullName
				$versionString = $versionFolder.Name
				$removed = $false
				if (-not $removed) {
					if ($PSCmdlet.ShouldProcess($folderPath, "Remove module '$ModuleName' version '$versionString'")) {
						$uninstalledViaCmdlet = $false
						$uninstallAttempted = $false
						# Uninstall-PSResource only manages the standard CurrentUser and AllUsers locations.
						$auModRoot = (Join-Path $env:ProgramFiles 'PowerShell\Modules') + '\'
						$cuModRoot = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules') + '\'
						$managedScope = if ($folderPath.StartsWith($auModRoot, [System.StringComparison]::OrdinalIgnoreCase)) { 'AllUsers' }
						elseif ($folderPath.StartsWith($cuModRoot, [System.StringComparison]::OrdinalIgnoreCase)) { 'CurrentUser' }
						else { $null }
						if ($managedScope) {
							$uninstallAttempted = $true
							try {
								Uninstall-PSResource -Name $ModuleName -Version $versionString -Scope $managedScope -Confirm:$false -Verbose:$false -SkipDependencyCheck -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
								$uninstalledViaCmdlet = $true
							}
							catch {
								New-Log "[$moduleName] Uninstall-PSResource failed for version '$versionString'" -Level ERROR
							}
						}
						else {
							New-Log "[$moduleName] '$folderPath' is in an unmanaged location that Uninstall-PSResource cannot target (it only manages the CurrentUser/AllUsers scopes). Removing the version folder directly." -Level VERBOSE
						}
						if ($uninstallAttempted) {
							$uninstallDeadline = [DateTime]::UtcNow.AddSeconds(1)
							while ((Test-Path -LiteralPath $folderPath -PathType Container) -and [DateTime]::UtcNow -lt $uninstallDeadline) {
								Start-Sleep -Milliseconds 100
							}
						}
						if (-not (Test-Path -LiteralPath $folderPath -PathType Container)) {
							$cleanedItems.Add($folderPath)
							$removed = $true
						}
						else {
							if ($uninstalledViaCmdlet) {
								New-Log "[$moduleName] Uninstall-PSResource returned successfully for '$versionString', but '$folderPath' still exists; using filesystem cleanup." -Level WARNING
							}
							else {
								New-Log "[$moduleName] Using filesystem cleanup for '$folderPath'." -Level VERBOSE
							}
						}
						if (-not $removed -and (Test-Path -LiteralPath $folderPath -PathType Container -Verbose:$false)) {
							New-Log "[$moduleName] Attempting Remove-Item -Recurse -Force on '$folderPath'..." -Level DEBUG
							$lastRemoveError = $null
							for ($attempt = 1; ($attempt -le 3) -and (-not $removed); $attempt++) {
								try {
									Get-ChildItem -LiteralPath $folderPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
										if ($_.Attributes -band [System.IO.FileAttributes]::ReadOnly) { try { $_.Attributes = [System.IO.FileAttributes]::Normal } catch { $null = $_ } }
									}
									Remove-Item -LiteralPath $folderPath -Recurse -Force -ErrorAction Stop -Verbose:$false | Out-Null
								}
								catch { $lastRemoveError = $_.Exception.Message }
								if (-not (Test-Path -LiteralPath $folderPath -PathType Container)) {
									$removed = $true
									break
								}
								if ($attempt -lt 3) {
									Start-Sleep -Milliseconds ($attempt * 200)
								}
							}
							if ($removed) {
								$cleanedItems.Add($folderPath)
							}
							else {
								$failureDetail = if ($lastRemoveError) { $lastRemoveError } else { 'folder still exists after three attempts' }
								New-Log "[$moduleName] Failed to remove '$folderPath': $failureDetail" -Level WARNING
							}
						}
					}
					else {
						New-Log "[$moduleName] Skipped removal of '$folderPath' due to ShouldProcess user choice." -Level DEBUG
					}
				}
			}
		}
	}
	$cleanedPathsToReturn = $cleanedItems | Where-Object { $_ -is [string] } | Select-Object -Unique -Verbose:$false
	return $cleanedPathsToReturn
}
#endregion Remove-OutdatedVersions
#region Get-ManifestVersionInfo
function Get-ManifestVersionInfo {
	<#
    .SYNOPSIS
    Extracts normalized module metadata from a manifest or manifest path.
    .DESCRIPTION
    Converts Test-ModuleManifest output into a module inventory record. When parsed data
    is unavailable, it infers the module name, version, path, and author from the file path.
    .PARAMETER ResData
    Object returned by Test-ModuleManifest.
    .PARAMETER Quick
    Uses ResData directly instead of reading ModuleFilePath.
    .PARAMETER ModuleFilePath
    Path to a module manifest used for fallback metadata discovery.
    .INPUTS
    None.
    .OUTPUTS
    System.Management.Automation.PSCustomObject containing module name, version, base
    path, prerelease label, and author; or null when metadata cannot be resolved.
    .EXAMPLE
    $manifest = Test-ModuleManifest 'C:\Modules\Example\1.0.0\Example.psd1'
    Get-ManifestVersionInfo -ResData $manifest -Quick
    Converts validated manifest data into an inventory record.
    .NOTES
    Internal helper used by Get-ModuleInfo.
    .LINK
    Test-ModuleManifest
    #>
	[CmdletBinding()]
	param (
		[Parameter()][object]$ResData,
		[switch]$Quick,
		[string]$ModuleFilePath
	)
	if (-not $ResData -and -not $Quick.IsPresent -and $ModuleFilePath) {
		$module = Get-ModuleformPath -Path $ModuleFilePath
		if ($module) {
			return [PSCustomObject]@{
				ModuleName          = $module.ModuleName
				ModuleVersion       = $module.ModuleVersion
				ModuleVersionString = $module.ModuleVersionString
				IsPreRelease        = $module.IsPrerelease
				PreReleaseLabel     = $module.PreReleaseLabel
				BasePath            = $module.BasePath
				Author              = $module.Author
			}
		}
	}
	if ($Quick.IsPresent -and $ResData) {
		$quickVersionString = $ResData.Version.ToString()
		$quickVersionStringPreRelease = if ($ResData.PrivateData.PSData.Prerelease) { $ResData.PrivateData.PSData.Prerelease } else { $null }
		if ($quickVersionString -and $quickVersionStringPreRelease) {
			$parsedQuickVersion = Parse-ModuleVersion -VersionString "$($quickVersionString)-$($quickVersionStringPreRelease)" -ErrorAction SilentlyContinue
		}
		else {
			$parsedQuickVersion = Parse-ModuleVersion -VersionString $($quickVersionString) -ErrorAction SilentlyContinue
		}
		[version]$quickVersion = if ($parsedQuickVersion) { $parsedQuickVersion.ModuleVersion } else { $null }
		[bool]$quickIsPre = if ($parsedQuickVersion) { $parsedQuickVersion.IsPrerelease } else { $false }
		[string]$quickPreLabel = if ($parsedQuickVersion) { $parsedQuickVersion.PreReleaseLabel } else { $null }
		$quickBasePath = if ($ResData.ModuleBase) { $ResData.ModuleBase } else { $null }
		$quickModuleName = if ($ResData.Name) { $ResData.Name } else { $null }
		if (!$quickVersion) {
			New-Log "Quick mode: Could not parse version '$quickVersionString' using Parse-ModuleVersion." -Level DEBUG
		}
		return [PSCustomObject]@{
			ModuleName          = $quickModuleName
			ModuleVersion       = $quickVersion
			ModuleVersionString = "$($parsedQuickVersion.ModuleVersionString)"
			IsPreRelease        = $quickIsPre
			PreReleaseLabel     = $quickPreLabel
			BasePath            = "$($quickBasePath)"
			Author              = $resData.Author
		}
	}
}
#endregion Get-ManifestVersionInfo
#region Test-IsResourceFile
function Test-IsResourceFile {
	<#
    .SYNOPSIS
    Identifies likely localization or resource files.
    .DESCRIPTION
    Tests a path against common culture-directory, resource-directory, and resource-file
    naming patterns so those files are not treated as module manifests.
    .PARAMETER Path
    File path to evaluate.
    .INPUTS
    None.
    .OUTPUTS
    System.Boolean.
    .EXAMPLE
    Test-IsResourceFile 'C:\Modules\Example\en-US\Example.strings.psd1'
    Returns true for a path under a culture directory.
    .NOTES
    Pattern-based detection can classify unusually named files as resources.
    .LINK
    Get-ModuleInfo
    #>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path
	)
	$normalizedPath = $Path -replace '/', '\'
	$fileName = Split-Path -Path $normalizedPath -Leaf
	$fileBaseName = [IO.Path]::GetFileNameWithoutExtension($fileName)
	$parentPath = Split-Path -Path $normalizedPath -Parent
	$parentName = Split-Path -Path $parentPath -Leaf
	$moduleDirectoryName = if ($parentName -match '^\d+(\.\d+){1,3}(-.+)?$') {
		Split-Path -Path (Split-Path -Path $parentPath -Parent) -Leaf
	}
	else {
		$parentName
	}
	# A manifest named after its flat or versioned module directory is always primary.
	if ($fileBaseName -ieq $moduleDirectoryName) {
		return $false
	}
	if ($normalizedPath -match '\\(Resources?|Localization|Localizations|Languages|Lang|Cultures?|i18n|l10n|DSCResources|Settings|Templates?|Tests?|TestData|Data|Runtimes?|InstallerHelpFiles)\\') {
		return $true
	}
	if ($normalizedPath -match '\\[a-z]{2}(-[a-z]{2})?\\[^\\]+\.(psd1|xml)$') {
		return $true
	}
	if ($fileBaseName -match '(^|[._-])(Resources?|Strings|Localized|Messages|Text|Errors|Labels|UI)$') {
		return $true
	}
	if ($fileName -match '^[a-z]{2,3}-[A-Z]{2,3}\.(psd1|xml)$') {
		return $true
	}
	return $false
}
#endregion Test-IsResourceFile
#region Resolve-ModuleVersion
function Resolve-ModuleVersion {
	<#
    .SYNOPSIS
    Resolves version data for a module installation path.
    .DESCRIPTION
    Parses the supplied version string and, when possible, replaces it with metadata from
    the matching module returned by Get-Module -ListAvailable.
    .PARAMETER VersionString
    Initial version string to parse.
    .PARAMETER ModuleName
    Name of the module to resolve.
    .PARAMETER Path
    File or directory path used to select the matching installation.
    .INPUTS
    None.
    .OUTPUTS
    System.Management.Automation.PSCustomObject with Module and VersionPattern properties.
    .EXAMPLE
    Resolve-ModuleVersion -VersionString 1.2.0 -ModuleName Example -Path 'C:\Modules\Example\1.2.0\Example.psd1'
    Resolves the version against the matching installed module.
    .NOTES
    Internal helper used by Get-ModuleformPath.
    .LINK
    Get-Module
    #>
	[CmdletBinding()]
	param (
		[Parameter()][string]$VersionString,
		[Parameter(Mandatory)][string]$ModuleName,
		[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path
	)
	$versionPattern = Parse-ModuleVersion -VersionString $VersionString
	$module = Get-Module -Name $ModuleName -All -ListAvailable -ErrorAction SilentlyContinue -Verbose:$false | Where-Object { $Path -like ($_.ModuleBase + '\*') } | Sort-Object -Property Version -Descending -ErrorAction SilentlyContinue | Select-Object -First 1 -ErrorAction SilentlyContinue
	if ($module) {
		$versionPattern = Parse-ModuleVersion -VersionString $module.Version
	}
	return [psCustomObject]@{
		Module         = $module
		VersionPattern = $versionPattern
	}
}
#endregion Resolve-ModuleVersion
#region Get-ModuleformPath
function Get-ModuleformPath {
	<#
    .SYNOPSIS
    Infers module metadata from an installation path.
    .DESCRIPTION
    Recognizes common versioned and flat PowerShell module layouts, then resolves the
    module name, version, base path, prerelease label, and author.
    .PARAMETER Path
    Manifest or directory path to inspect.
    .INPUTS
    None.
    .OUTPUTS
    System.Management.Automation.PSCustomObject containing inferred module metadata.
    Unresolved properties are null.
    .EXAMPLE
    Get-ModuleformPath 'C:\Modules\Example\1.2.0\Example.psd1'
    Returns metadata inferred from the versioned module path.
    .NOTES
    Internal helper. The function name is retained for compatibility.
    .LINK
    Resolve-ModuleVersion
    #>
	[CmdletBinding()]
	[OutputType([PSCustomObject])]
	param (
		[Parameter(Mandatory)][string]$Path
	)
	if ([string]::IsNullOrWhiteSpace($Path)) {
		New-Log "Input path cannot be null or empty. Will abort." -Level WARNING
		return [PSCustomObject]@{
			ModuleName          = $null
			BasePath            = $null
			ModuleVersion       = $null
			ModuleVersionString = $null
			IsPrerelease        = $null
			PreReleaseLabel     = $null
			Author              = $null
		}
	}
	$trimmedPath = $Path.TrimEnd('\/')
	$normalizedPath = $trimmedPath -replace '/', '\'
	$versionPattern = '\d+(\.\d+){1,3}(-.+)?'
	$regexVersionOnly = "^$versionPattern$"
	$regexPattern1 = "(?i)^(.*\\Modules\\([^\\]+))\\($versionPattern)(?:\\.*)?$"
	New-Log "Trying Pattern 1 regex: '$regexPattern1' against '$normalizedPath'" -Level VERBOSE
	if ($normalizedPath -match $regexPattern1) {
		if ($Matches -ne $null -and $Matches.Count -ge 3) {
			[string]$modulePath = $Matches[1]
			[string]$moduleName = $Matches[2]
			[string]$matchedVersion = $Matches[3]
			$moduleData = Resolve-ModuleVersion -VersionString $matchedVersion -ModuleName $moduleName -Path $Path
			$versionPattern = $moduleData.VersionPattern
			$module = $moduleData.Module
			return [PSCustomObject]@{
				ModuleName          = $($moduleName)
				BasePath            = $($modulePath)
				ModuleVersion       = $versionPattern.ModuleVersion
				ModuleVersionString = $versionPattern.ModuleVersionString
				IsPrerelease        = $versionPattern.IsPrerelease
				PreReleaseLabel     = $versionPattern.PreReleaseLabel
				Author              = if ($module.Author) { $module.Author } else { $null }
			}
		}
	}
	New-Log "Pattern 1 did NOT match." -Level VERBOSE
	$potentialModuleBasePath = $null
	if (Test-Path -Path $normalizedPath -PathType Container -Verbose:$false) {
		$potentialModuleBasePath = $normalizedPath
		New-Log "Pattern 2: Input '$normalizedPath' is a directory. Assuming it's the ModulePath." -Level VERBOSE
	}
	else {
		$potentialModuleBasePath = Split-Path -Path $normalizedPath -Parent -ErrorAction SilentlyContinue -Verbose:$false
		New-Log "Pattern 2: Input '$normalizedPath' is a file. Assuming parent '$potentialModuleBasePath' is the ModulePath." -Level VERBOSE
	}
	if ($potentialModuleBasePath) {
		$potentialModuleName = Split-Path -Path $potentialModuleBasePath -Leaf -ErrorAction SilentlyContinue -Verbose:$false
		New-Log "Pattern 2: Potential ModuleName (leaf of ModulePath) is '$potentialModuleName'." -Level VERBOSE
		if ($potentialModuleName -match $regexVersionOnly) {
			New-Log "Pattern 2: determined ModuleName '$potentialModuleName' which looks like a version. This is usually incorrect." -Level WARNING
			$parentOfVersionDir = Split-Path -Path $potentialModuleBasePath -Parent -ErrorAction SilentlyContinue -Verbose:$false
			if ($parentOfVersionDir) {
				$betterModuleName = Split-Path -Path $parentOfVersionDir -Leaf -ErrorAction SilentlyContinue -Verbose:$false
				if ($betterModuleName -and $betterModuleName -ne 'Modules') {
					$moduleData = Resolve-ModuleVersion -VersionString $potentialModuleName -ModuleName $betterModuleName -Path $Path
					$versionPattern = $moduleData.VersionPattern
					$module = $moduleData.Module
					return [PSCustomObject]@{
						ModuleName          = $($betterModuleName)
						BasePath            = $($parentOfVersionDir)
						ModuleVersion       = $versionPattern.ModuleVersion
						ModuleVersionString = $versionPattern.ModuleVersionString
						IsPrerelease        = $versionPattern.IsPrerelease
						PreReleaseLabel     = $versionPattern.PreReleaseLabel
						Author              = if ($module.Author) { $module.Author } else { $null }
					}
				}
			}
		}
		elseif ($potentialModuleName -and $potentialModuleName -ne 'Modules') {
			New-Log "Pattern 2: Found potential module name '$potentialModuleName', attempting to find version information" -Level VERBOSE
			$version = $null
			$pathParts = $normalizedPath -split '\\'
			foreach ($part in $pathParts) {
				if ($part -match $regexVersionOnly) {
					$version = $part
					New-Log "Pattern 2: Found version '$version' in path component" -Level VERBOSE
					break
				}
			}
			if (-not $version -and $normalizedPath -like "*.psd1") {
				try {
					$manifestContent = Get-Content -Path $normalizedPath -Raw -ErrorAction SilentlyContinue -Verbose:$false
					if ($manifestContent -match "ModuleVersion\s*=\s*['`"]([^'`"]+)") {
						$version = $Matches[1]
						New-Log "Pattern 2: Found version '$version' in module manifest" -Level VERBOSE
					}
				}
				catch {
					New-Log "Pattern 2: Error reading module manifest." -Level VERBOSE
				}
			}
			$moduleData = Resolve-ModuleVersion -VersionString $version -ModuleName $potentialModuleName -Path $Path
			$versionPattern = $moduleData.VersionPattern
			$module = $moduleData.Module
			return [PSCustomObject]@{
				ModuleName          = $($potentialModuleName)
				BasePath            = $($potentialModuleBasePath)
				ModuleVersion       = $versionPattern.ModuleVersion
				ModuleVersionString = $versionPattern.ModuleVersionString
				IsPrerelease        = $versionPattern.IsPrerelease
				PreReleaseLabel     = $versionPattern.PreReleaseLabel
				Author              = if ($module -and $module.Author) { $module.Author } else { $null }
			}
		}
	}
	return [PSCustomObject]@{
		ModuleName          = $null
		BasePath            = $null
		ModuleVersion       = $null
		ModuleVersionString = $null
		IsPrerelease        = $null
		PreReleaseLabel     = $null
		Author              = $null
	}
}
#endregion Get-ModuleformPath
#region Get-ModuleInfoFromXml
function Get-ModuleInfoFromXml {
	<#
    .SYNOPSIS
    Reads module metadata from PSGetModuleInfo.xml.
    .DESCRIPTION
    Parses PowerShell-serialized module metadata, including normalized prerelease data,
    and returns an inventory record compatible with Get-ModuleInfo.
    .PARAMETER XmlFilePath
    Path to a PSGetModuleInfo.xml file.
    .INPUTS
    None.
    .OUTPUTS
    System.Management.Automation.PSCustomObject containing module name, version, base
    path, prerelease label, and author; or null when required data is missing.
    .EXAMPLE
    Get-ModuleInfoFromXml 'C:\Modules\Example\1.2.0\PSGetModuleInfo.xml'
    Reads the saved module metadata.
    .NOTES
    Internal helper used by Get-ModuleInfo.
    .LINK
    Get-ModuleInfo
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$XmlFilePath
	)
	try {
		[xml]$xmlContent = Get-Content -Path $XmlFilePath -Raw -ErrorAction Stop -Verbose:$false
		$nsManager = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
		$nsManager.AddNamespace("ps", "http://schemas.microsoft.com/powershell/2004/04")
		$nameNode = $xmlContent.SelectSingleNode("//ps:S[@N='Name']", $nsManager)
		$nameValue = if ($nameNode) { $nameNode.'#text' } else { $null }
		$authorNode = $xmlContent.SelectSingleNode("//ps:S[@N='Author']", $nsManager)
		$authorValue = if ($authorNode) { $authorNode.'#text' } else { $null }
		$versionNode = $xmlContent.SelectSingleNode("//ps:S[@N='Version']", $nsManager)
		$versionValue = if ($versionNode) { $versionNode.'#text' } else { $null }
		$locationNode = $xmlContent.SelectSingleNode("//ps:S[@N='InstalledLocation']", $nsManager)
		$locationValue = if ($locationNode) { $locationNode.'#text' } else { $null }
		$normalizedVersionNode = $xmlContent.SelectSingleNode("//ps:Obj[@N='AdditionalMetadata']/MS/ps:S[@N='NormalizedVersion']", $nsManager)
		if (-not $normalizedVersionNode) {
			$normalizedVersionNode = $xmlContent.SelectSingleNode("//ps:S[@N='NormalizedVersion']", $nsManager)
		}
		$normalizedVersionValue = if ($normalizedVersionNode) { $normalizedVersionNode.'#text' } else { $null }
		$prereleaseBoolNode = $xmlContent.SelectSingleNode("//ps:B[@N='IsPrerelease']", $nsManager)
		$prereleaseBoolText = if ($prereleaseBoolNode) { $prereleaseBoolNode.'#text' } else { $null }
		$prereleaseStringNode = $xmlContent.SelectSingleNode("//ps:Obj[@N='AdditionalMetadata']/MS/ps:S[@N='IsPrerelease']", $nsManager)
		if (-not $prereleaseStringNode) {
			$prereleaseStringNode = $xmlContent.SelectSingleNode("//ps:S[@N='IsPrerelease']", $nsManager)
		}
		$prereleaseStringText = if ($prereleaseStringNode) { $prereleaseStringNode.'#text' } else { $null }
		$parsedVersionInfo = $null
		$isPrereleaseFromParse = $false
		$preReleaseLabelFromParse = $null
		$moduleVersionObject = $null
		if ($normalizedVersionValue) {
			New-Log "XML: Attempting to parse NormalizedVersion '$normalizedVersionValue'" -Level VERBOSE
			$parsedVersionInfo = Parse-ModuleVersion -VersionString $normalizedVersionValue -ErrorAction SilentlyContinue
			if ($parsedVersionInfo) {
				New-Log "XML: Parsed NormalizedVersion '$normalizedVersionValue'. IsPre: $($parsedVersionInfo.IsPrerelease), Label: '$($parsedVersionInfo.PreReleaseLabel)'" -Level VERBOSE
			}
			else {
				New-Log "XML: Could not parse NormalizedVersion string '$normalizedVersionValue' using Parse-ModuleVersion." -Level DEBUG
			}
		}
		if (-not $parsedVersionInfo -and $versionValue) {
			New-Log "XML: Attempting to parse Version '$versionValue' (NormalizedVersion failed or N/A)" -Level VERBOSE
			$parsedVersionInfo = Parse-ModuleVersion -VersionString $versionValue -ErrorAction SilentlyContinue
			if ($parsedVersionInfo) {
				New-Log "XML: Parsed Version '$versionValue'. IsPre: $($parsedVersionInfo.IsPrerelease), Label: '$($parsedVersionInfo.PreReleaseLabel)'" -Level VERBOSE
			}
			else {
				New-Log "XML: Could not parse Version string '$versionValue' using Parse-ModuleVersion." -Level WARNING
			}
		}
		if ($parsedVersionInfo) {
			$moduleVersionObject = $parsedVersionInfo.ModuleVersion
			$isPrereleaseFromParse = $parsedVersionInfo.IsPrerelease
			$preReleaseLabelFromParse = $parsedVersionInfo.PreReleaseLabel
		}
		else {
			if ($versionValue) {
				try {
					$moduleVersionObject = [System.Version]$versionValue
					New-Log "XML: Fallback - Created System.Version from '$versionValue' as Parse-ModuleVersion failed for all inputs." -Level DEBUG
				}
				catch {
					New-Log "XML: Fallback - Could not create System.Version from '$versionValue' either." -Level VERBOSE
				}
			}
		}
		$isPrereleaseFromXmlFlag = $false
		if (($prereleaseBoolText -is [string] -and $prereleaseBoolText.ToLowerInvariant() -eq 'true') -or
			($prereleaseStringText -is [string] -and $prereleaseStringText.ToLowerInvariant() -eq 'true')) {
			$isPrereleaseFromXmlFlag = $true
			New-Log "XML: An explicit IsPrerelease flag in XML is true (Bool: '$prereleaseBoolText', String: '$prereleaseStringText')." -Level VERBOSE
		}
		$finalIsPrerelease = $isPrereleaseFromParse -or $isPrereleaseFromXmlFlag
		$finalPreReleaseLabel = $preReleaseLabelFromParse
		if ($nameValue -and $versionValue) {
			$basePathValue = $null
			if ($locationValue -and $nameValue) {
				try {
					$basePathValue = Join-Path -Path $locationValue -ChildPath $nameValue -ErrorAction Stop -Verbose:$false
				}
				catch {
					New-Log "XML: Error constructing BasePath from Location '$locationValue' and Name '$nameValue'." -Level ERROR
				}
			}
			else {
				New-Log "XML: Cannot construct BasePath - InstalledLocation or Name missing. Location: '$locationValue', Name: '$nameValue'" -Level VERBOSE
			}
			$result = [PSCustomObject]@{
				ModuleName          = $nameValue
				ModuleVersion       = $moduleVersionObject
				ModuleVersionString = $versionValue
				BasePath            = if ($basePathValue) { "$basePathValue" } else { $null }
				isPreRelease        = $finalIsPrerelease
				PreReleaseLabel     = if ($finalIsPrerelease -and $finalPreReleaseLabel) { $finalPreReleaseLabel } else { $null }
				Author              = if ($authorValue) { $authorValue } else { $null }
			}
			New-Log "XML Parsed: Name='$($result.ModuleName)', VersionObj='$($result.ModuleVersion)', OrigVerStr='$($result.ModuleVersionString)', BasePath='$($result.BasePath)', IsPre=$($result.isPreRelease), Label='$($result.PreReleaseLabel)'" -Level VERBOSE
			return $result
		}
		else {
			New-Log "Could not find required 'Name' or 'Version' node in XML: $XmlFilePath. Name: '$nameValue', Version: '$versionValue'" -Level WARNING
			return $null
		}
	}
	catch {
		New-Log "Error parsing XML file '$XmlFilePath'" -Level ERROR
		return $null
	}
}
#endregion Get-ModuleInfoFromXml
#region Parse-ModuleVersion
function Parse-ModuleVersion {
	<#
    .SYNOPSIS
    Parses numeric and prerelease module version strings.
    .DESCRIPTION
    Separates a two-to-four-part numeric version from an optional SemVer-style prerelease
    label and returns both the original and normalized values.
    .PARAMETER VersionString
    Version string to parse.
    .INPUTS
    None.
    .OUTPUTS
    System.Management.Automation.PSCustomObject containing ModuleVersion,
    ModuleVersionString, IsPrerelease, PreReleaseLabel, and related normalized values;
    or null when the numeric version is invalid.
    .EXAMPLE
    Parse-ModuleVersion '2.0.0-preview.3'
    Returns version 2.0.0 with prerelease label preview.3.
    .NOTES
    Build metadata is not parsed.
    .LINK
    Compare-ModuleVersion
    #>
	[CmdletBinding()]
	param (
		[Parameter()][string]$VersionString
	)
	if ([string]::IsNullOrWhiteSpace($VersionString)) { return $null }
	[version]$version = $null
	[string]$baseVersionString = $VersionString
	[string]$prereleasePart = $null
	[bool]$isSemVerStyle = $false
	[bool]$isPrerelease = $false
	$semVerRegex = '^(\d+(\.\d+){1,3})-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)$'
	if ($VersionString -match $semVerRegex) {
		$baseVersionString = $Matches[1]
		$prereleasePart = $Matches[3]
		$isSemVerStyle = $true
		$isPrerelease = $true
		New-Log "Parse-ModuleVersion: Matched SemVer pattern. Base:'$baseVersionString', Label:'$prereleasePart'" -Level VERBOSE
	}
	else {
		$baseVersionString = $VersionString
	}
	$isParseable = [version]::TryParse($baseVersionString, [ref]$version)
	if ($isParseable) {
		return [PSCustomObject]@{
			ModuleVersionStringNoPrefix = $baseVersionString
			ModuleVersionString         = $VersionString
			ModuleVersion               = $version
			IsSemVer                    = $isSemVerStyle
			IsPrerelease                = $isPrerelease
			PreReleaseLabel             = if (![string]::IsNullOrEmpty($prereleasePart)) { $prereleasePart } else { $null }
			PreReleaseVersion           = if (![string]::IsNullOrEmpty($prereleasePart)) { "$($version)-$($prereleasePart)" } else { $null }
		}
	}
	else {
		New-Log "Parse-ModuleVersion: Could not parse base version string '$baseVersionString' (derived from original '$VersionString') into a [System.Version] object." -Level WARNING
		return $null
	}
}
#endregion Parse-ModuleVersion
#region Compare-ModuleVersion
function Compare-ModuleVersion {
	<#
    .SYNOPSIS
    Compares two module version strings.
    .DESCRIPTION
    Compares numeric versions first, then prerelease labels. For equal numeric versions,
    a prerelease is treated as newer than a stable version. Recognized label priority is
    dev, alpha, beta, preview, then rc; numeric label suffixes break ties.
    .PARAMETER VersionA
    Current or first version.
    .PARAMETER VersionB
    Candidate or second version.
    .PARAMETER ReturnBoolean
    Returns true only when VersionB is newer than VersionA. Without this switch, returns
    the version string considered newer.
    .INPUTS
    None.
    .OUTPUTS
    System.Boolean when ReturnBoolean is used; otherwise System.String.
    .EXAMPLE
    Compare-ModuleVersion -VersionA '1.0.0-beta1' -VersionB '1.0.0-beta2' -ReturnBoolean
    Returns true.
    .NOTES
    The prerelease ordering is script-specific and intentionally differs from SemVer,
    where a stable release normally has higher precedence than its prereleases.
    .LINK
    Parse-ModuleVersion
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$VersionA,
		[Parameter(Mandatory)][string]$VersionB,
		[Parameter()][switch]$ReturnBoolean
	)
	if ([string]::IsNullOrWhiteSpace($VersionA) -or [string]::IsNullOrWhiteSpace($VersionB)) {
		New-Log "Compare-ModuleVersion: One or both versions are null or empty." -Level VERBOSE
		if ($ReturnBoolean) { return $false } else { return $VersionA }
	}
	$typePriority = @{
		'dev'     = 5
		'alpha'   = 4
		'beta'    = 3
		'preview' = 2
		'rc'      = 1
	}
	$baseVersionA, $prereleaseA = $VersionA -split '-', 2
	$baseVersionB, $prereleaseB = $VersionB -split '-', 2
	try {
		$versionObjectA = [System.Version]::new($baseVersionA)
		$versionObjectB = [System.Version]::new($baseVersionB)
		$baseComparison = $versionObjectA.CompareTo($versionObjectB)
		if ($baseComparison -ne 0) {
			New-Log "Compare-ModuleVersion: Base versions differ - A=$baseVersionA, B=$baseVersionB. Result=$baseComparison" -Level VERBOSE
			if ($ReturnBoolean) {
				return $baseComparison -lt 0
			}
			else {
				return $(if ($baseComparison -gt 0) { $VersionA } else { $VersionB })
			}
		}
	}
	catch {
		New-Log "Compare-ModuleVersion: Error parsing version strings. Falling back to string comparison." -Level VERBOSE
		if ($ReturnBoolean) { return $false } else { return $VersionA }
	}
	if ($prereleaseA -and -not $prereleaseB) {
		New-Log "Compare-ModuleVersion: VersionA has prerelease ($prereleaseA), VersionB doesn't. A is newer." -Level VERBOSE
		if ($ReturnBoolean) { return $false } else { return $VersionA }
	}
	elseif (-not $prereleaseA -and $prereleaseB) {
		New-Log "Compare-ModuleVersion: VersionB has prerelease ($prereleaseB), VersionA doesn't. B is newer." -Level VERBOSE
		if ($ReturnBoolean) { return $true } else { return $VersionB }
	}
	elseif (-not $prereleaseA -and -not $prereleaseB) {
		New-Log "Compare-ModuleVersion: Neither version has prerelease. Versions are equal." -Level VERBOSE
		if ($ReturnBoolean) { return $false } else { return $VersionA }
	}
	$prereleaseA = $prereleaseA.ToLower().Trim('.- ')
	$prereleaseB = $prereleaseB.ToLower().Trim('.- ')
	$regex = "^($(($typePriority.Keys | ForEach-Object {[regex]::Escape($_)}) -join '|'))(?:[\.\-_]?(\d+))?$"
	$matchA = [regex]::Match($prereleaseA, $regex)
	$matchB = [regex]::Match($prereleaseB, $regex)
	if (-not $matchA.Success -or -not $matchB.Success) {
		New-Log "Compare-ModuleVersion: Non-standard prerelease format. Performing simple string comparison." -Level VERBOSE
		$comparisonResult = $prereleaseA.CompareTo($prereleaseB)
		if ($ReturnBoolean) {
			return $comparisonResult -lt 0
		}
		else {
			return $(if ($comparisonResult -ge 0) { $VersionA } else { $VersionB })
		}
	}
	$typeA = $matchA.Groups[1].Value
	$numberA = if ([string]::IsNullOrEmpty($matchA.Groups[2].Value)) { 0 } else { [int]$matchA.Groups[2].Value }
	$typeB = $matchB.Groups[1].Value
	$numberB = if ([string]::IsNullOrEmpty($matchB.Groups[2].Value)) { 0 } else { [int]$matchB.Groups[2].Value }
	New-Log "Compare-ModuleVersion: Comparing prerelease A='$typeA'($numberA) vs B='$typeB'($numberB)" -Level VERBOSE
	if ($typeA -ne $typeB) {
		$priorityA = $typePriority[$typeA]
		$priorityB = $typePriority[$typeB]
		New-Log "Prerelease types differ. Priority A=$priorityA, B=$priorityB." -Level VERBOSE
		if ($ReturnBoolean) {
			return $priorityB -gt $priorityA
		}
		else {
			return $(if ($priorityA -gt $priorityB) { $VersionA } else { $VersionB })
		}
	}
	New-Log "Prerelease types are the same ('$typeA'). Comparing numbers: A=$numberA, B=$numberB." -Level VERBOSE
	if ($ReturnBoolean) {
		return $numberB -gt $numberA
	}
	else {
		return $(if ($numberA -ge $numberB) { $VersionA } else { $VersionB })
	}
}
#endregion Compare-ModuleVersion
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
Check-PSResourceRepository
$ignoredModules = @('Example2.Diagnostics', 'BurntToast')
$blackList = @{
	'Microsoft.Graph.Beta' = @("Nuget", "NugetGallery")
	'Microsoft.Graph'      = @("Nuget", "NugetGallery")
}
$paths = $env:PSModulePath.Split(';') | Where-Object { $_ -notmatch '.vscode' -and $_ -notmatch 'System32' }
$moduleInfo = Get-ModuleInfo -Paths $paths -IgnoredModules $ignoredModules
$outdated = Get-ModuleUpdateStatus -ModuleInventory $moduleInfo -TimeoutSeconds 120 -Repositories @("PSGallery", "Nuget", "NugetGallery") -MatchAuthor -BlackList $blackList
if ($outdated) {
	$outdated | Update-Modules -Clean -PreRelease
}
else {
	New-Log "No outdated modules to update" -Level SUCCESS
}