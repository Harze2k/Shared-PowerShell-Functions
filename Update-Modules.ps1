#Requires -Version 7.0
#Requires -modules ThreadJob
#Requires -modules Microsoft.PowerShell.ThreadJob
#region Check-PSResourceRepository
<#
Author: Harze2k
Date:   2026-02-25
Version: 4.0 (Parallel module installation!)
	-Update-Modules now installs modules in parallel using ForEach-Object -Parallel.
	-3-phase architecture: pre-processing (sequential) -> parallel install -> post-processing/cleaning (sequential).
	-ShouldProcess checks evaluated on main thread before dispatching to parallel runspaces.
	-Throttle limit scales with CPU cores: Min(moduleCount, Max(4, ProcessorCount * 2)).
	-Progress reporting with ETA during parallel installation.
	-Cleaning still runs sequentially to avoid filesystem conflicts.
	---Older ---
	Version 3.7 (2025-05-21):
	-Fixed a loop issue that caused the Find-Module to slow down. (10x speed increase)
	-Fixed PreRelease logic in several functions.
	-Now we try to parse the PreRelease version also from .XML files.
	-Fixed output from several functions to be relevant and less spammy.
	-Fixed a typo that made the -MatchAutor not work.
	-Added more Parallel Processing
	-Added ThreadJobs
	-Over 150% faster
	-Finds basically all modules that can be found.
	-Included help and guidelines.
	-Progress reporting while parallel jobs.
#>
function Check-PSResourceRepository {
	[CmdletBinding()]
	param (
		[switch]$ImportDependencies, # Force re-import of PSResourceGet module even if commands exist
		[switch]$ForceInstall, # Force reinstall of PSResourceGet (PS5.1 only)
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
			New-Log "$OperationName failed: $($_.Exception.Message)" -Level ERROR
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
			New-Log "Unable to set TLS 1.2: $($_.Exception.Message)" -Level ERROR
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
			New-Log "Error configuring PSGallery: $($_.Exception.Message)" -Level WARNING
		}
		$forceFlag = $Force.IsPresent
		$installScript = [scriptblock]::Create(@"
            `$ErrorActionPreference = 'Stop'
            Install-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Repository 'PSGallery' -Scope AllUsers -Force:$forceFlag -AllowClobber -AcceptLicense -SkipPublisherCheck -Confirm:`$false
"@)
		New-Log "Installing Microsoft.PowerShell.PSResourceGet via Install-Module..."
		Invoke-WithTimeout -ScriptBlock $installScript -Timeout ($Timeout * 2) -OperationName "Install-Module PSResourceGet" | Out-Null
		try {
			Import-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Force -ErrorAction Stop
			New-Log "Successfully imported Microsoft.PowerShell.PSResourceGet." -Level SUCCESS
			return $true
		}
		catch {
			New-Log "Failed to import Microsoft.PowerShell.PSResourceGet: $($_.Exception.Message)" -Level ERROR
			return $false
		}
	}
	function Import-PSResourceGetModule {
		[CmdletBinding()]
		param ([switch]$Force)
		$action = if ($Force) { "Force importing" } else { "Importing" }
		New-Log "$action Microsoft.PowerShell.PSResourceGet module..."
		try {
			Import-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Force:$Force -ErrorAction Stop
			New-Log "Successfully imported PSResourceGet." -Level SUCCESS
			return $true
		}
		catch {
			New-Log "Failed to import PSResourceGet: $($_.Exception.Message)" -Level ERROR
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
			New-Log "Failed to configure '$Name': $($_.Exception.Message)" -Level ERROR
			return $false
		}
	}
	if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
		New-Log "Administrator privileges required. Aborting." -Level WARNING
		return $false
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
	Scans specified paths for PowerShell module manifest files (.psd1) and PSGetModuleInfo.xml files, processing them in parallel to gather module metadata.
	.DESCRIPTION
	This function discovers and processes PowerShell module files to create a comprehensive inventory. It operates in three main phases:
	1.  *Helper Function Preparation:** Collects definitions of internal helper functions (like `Parse-ModuleVersion`, `Get-ManifestVersionInfo`, etc.) as strings to make them available within parallel processing scopes.
	2.  *Phase 1: File Discovery:** Recursively scans the provided `-Paths` for all files ending in `.psd1` (module manifests) and files named `PSGetModuleInfo.xml` (often created by `Save-PSResource -IncludeXml`).
	3.  *Phase 2: Parallel File Processing:** Each discovered file is processed in parallel using `ForEach-Object -Parallel`.
	Inside the parallel task, helper functions are restored using `Invoke-Expression`.
	Likely resource files (e.g., localization files) are skipped using the `Test-IsResourceFile` helper.
	For `.psd1` files:
	`Test-ModuleManifest` is called to validate and get basic manifest data.
	The output of `Test-ModuleManifest` is preferably processed by `Get-ManifestVersionInfo -Quick`.
	If `Test-ModuleManifest` fails or its output is insufficient, `Get-ManifestVersionInfo -ModuleFilePath` is used as a fallback, attempting to infer details from the file path.
	Extracted metadata (ModuleName, ModuleVersion, BasePath, Author, pre-release info) is added to a thread-safe `ConcurrentBag`.
	For `PSGetModuleInfo.xml` files:
	`Get-ModuleInfoFromXml` is called to parse the XML and extract metadata.
	Extracted metadata is added to the `ConcurrentBag`.
	4.  *Phase 3: Aggregation and Normalization:**
	The collected module entries from the `ConcurrentBag` are converted to an array.
	Initial deduplication is performed by grouping entries by ModuleName, BasePath, and ModuleVersionString.
	Further processing groups modules by name. For each name group:
	Module base paths are normalized. This involves logic to identify the true root directory of a module, especially when version numbers are part of the directory structure (e.g., ensuring 'C:\Modules\MyModule\1.0' and 'C:\Modules\MyModule\1.1' both resolve 'C:\Modules\MyModule' as the base path).
	Modules are then grouped by these normalized base paths.
	Within each base path group, versions are grouped to select a single representative entry (preferring entries with `[System.Version]` objects over strings if both exist for the same version identifier).
	The final, aggregated module data is stored in an ordered hashtable, where each key is a module name and the value is an array of `PSCustomObject`s, each representing a unique installation location and version of that module.
	Modules specified in the `-IgnoredModules` parameter are filtered out from the final result.
	The function returns this ordered hashtable, providing a structured inventory of all discovered modules.
	.PARAMETER Paths
	[Mandatory] An array of strings, where each string is a path to a directory to be scanned recursively for module files. Typically, this would be paths from `$env:PSModulePath`.
	.PARAMETER IgnoredModules
	An array of strings containing module names to be excluded from the final results. Defaults to an empty array.
	.PARAMETER ThrottleLimit
	The maximum number of parallel threads to use for processing files in Phase 2. Defaults to `([System.Environment]::ProcessorCount * 2)`.
	.INPUTS
	None. This function does not accept pipeline input.
	.OUTPUTS
	System.Collections.Specialized.OrderedDictionary
	Returns an ordered hashtable where:
	- Each key is a [string] representing a unique module name found.
	- Each value is an array of [PSCustomObject]s. Each PSCustomObject represents a distinct installation of that module and contains properties like:
	- ModuleName ([string])
	- ModuleVersion ([System.Version])
	- ModuleVersionString ([string])
	- BasePath ([string]): The normalized root directory of this module installation.
	- IsPreRelease ([bool])
	- PreReleaseLabel ([string])
	- Author ([string])
	If no module files are found, an empty ordered hashtable is returned.
	.EXAMPLE
	PS C:\> $moduleInventory = Get-ModuleInfo -Paths ($env:PSModulePath -split ';') -IgnoredModules 'Pester','MyCustomDevModule'
	Scans all standard module paths, processes found module files in parallel, and returns an inventory excluding 'Pester' and 'MyCustomDevModule'.
	.NOTES
	- Requires PowerShell 7.0 or later due to the use of `ForEach-Object -Parallel`.
	- Relies on several internal helper functions (e.g., `Parse-ModuleVersion`, `Get-ManifestVersionInfo`, `Get-ModuleformPath`, `Test-IsResourceFile`, `Get-ModuleInfoFromXml`) which must be defined in the same scope.
	- Depends on an external `New-Log` function for logging operations.
	- The accuracy of `BasePath` normalization depends on common module directory structures.
	.LINK
	Test-ModuleManifest
	Get-ChildItem
	ForEach-Object -Parallel
	Invoke-Expression
	#>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)][string[]]$Paths,
		[string[]]$IgnoredModules = @(),
		[int]$ThrottleLimit = ([System.Environment]::ProcessorCount * 2) # Increased default, file processing is often quick
	)
	# --- Define Helper Functions as Strings to Pass to Parallel Scope ---
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
	New-Log "Phase 1: Gathering all .psd1 and PSGetModuleInfo.xml file paths..."
	$fileDiscoveryStartTime = Get-Date
	$allPotentialFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
	foreach ($dir in $paths) {
		try {
			$psd1Files = @(Get-ChildItem -Path $dir -Recurse -Filter "*.psd1" -ErrorAction SilentlyContinue)
			$xmlFiles = @(Get-ChildItem -Path $dir -Recurse -Filter "PSGetModuleInfo.xml" -ErrorAction SilentlyContinue)
			foreach ($file in $psd1Files) { $allPotentialFiles.Add($file) }
			foreach ($file in $xmlFiles) { $allPotentialFiles.Add($file) }
		}
		catch {
			Write-Host "Error processing directory $($dir.FullName): $_" -ForegroundColor Red
		}
	}
	$fileDiscoveryDuration = (Get-Date) - $fileDiscoveryStartTime
	New-Log "Phase 1 complete. Found $($allPotentialFiles.Count) potential module files in $($fileDiscoveryDuration.ToString("g"))."
	if ($allPotentialFiles.Count -eq 0) {
		New-Log "No potential module files found to process." -Level WARNING
		return [ordered]@{}
	}
	# --- PHASE 2: Process individual files in parallel ---
	New-Log "Phase 2: Starting parallel processing of $($allPotentialFiles.Count) files (Throttle: $ThrottleLimit)..."
	$parallelProcessingStartTime = Get-Date
	$allFoundModulesFromParallel = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
	$sortedModules = [ordered]@{}
	$sortedModules = $allPotentialFiles | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
		# --- BEGIN PARALLEL SCRIPT BLOCK ---
		$fileInfo = $_ # Current System.IO.FileInfo object
		$filePath = $fileInfo.FullName
		$fileExtension = $fileInfo.Extension # .psd1 or .xml
		Import-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Global -Force
		# $currentKernelTime = (Get-Process -Id $pid).KernelTime # Example, can be removed if not used
		$VerbosePreference = $using:VerbosePreference
		$allFoundModulesFromParallel = $using:allFoundModulesFromParallel
		# --- Restore Helper Functions ---
		# Bring the whole hashtable into the parallel scope
		$localHelperFunctionDefinitions = $using:helperFunctionDefinitions
		foreach ($funcName in $localHelperFunctionDefinitions.Keys) {
			# Ensure the definition string is not null or empty before invoking
			$funcDefinition = $localHelperFunctionDefinitions[$funcName]
			if (-not [string]::IsNullOrWhiteSpace($funcDefinition)) {
				Invoke-Expression -Command "function global:$funcName { $funcDefinition }"
			}
		}
		if (Test-IsResourceFile -Path $filePath) {
			New-Log "Skipping likely resource file in parallel: $filePath" -Level VERBOSE
			return
		}
		if ($fileExtension -eq '.psd1') {
			$manifestInfoObj = $null
			$testManifestOutput = $null
			try {
				$testManifestOutput = Test-ModuleManifest -Path $filePath -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false
			}
			catch {
				$testManifestOutput = $null
				New-Log "Test-ModuleManifest crashed while using path $filePath" -Level VERBOSE
			}
			if ($testManifestOutput) {
				try {
					$manifestInfoObj = Get-ManifestVersionInfo -ResData $testManifestOutput -Quick -ErrorAction Stop -WarningAction SilentlyContinue
				}
				catch {
					$manifestInfoObj = $null
					New-Log "Get-ManifestVersionInfo crashed." -Level VERBOSE
				}
			}
			if (-not $manifestInfoObj) {
				try {
					$manifestInfoObj = Get-ManifestVersionInfo -ModuleFilePath $filePath -ErrorAction Stop -WarningAction SilentlyContinue
				}
				catch {
					$manifestInfoObj = $null
					New-Log "Get-ManifestVersionInfo while using path $filePath" -Level VERBOSE
				}
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
					New-Log "[$($mInfo.ModuleName)] Successfully found module info from the [.PSD1] file. Version [$($mInfo.ModuleVersionString)]" -Level SUCCESS
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
				else {
					New-Log "psd1 info from '$mInfo' missing required parameters (ModuleName/ModuleVersion)." -Level VERBOSE
				}
			}
		}
		elseif ($fileExtension -eq '.xml') {
			$xmlInfo = Get-ModuleInfoFromXml -XmlFilePath $filePath -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
			if ($xmlInfo -and $xmlInfo.ModuleName -and $xmlInfo.ModuleVersion) {
				New-Log "[$($xmlInfo.ModuleName)] Successfully found module info from the [.XML] file. Version [$($xmlInfo.ModuleVersionString)]" -Level SUCCESS
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
				New-Log "xml info from '$xmlInfo' missing required parameters (ModuleName/ModuleVersion)." -Level VERBOSE
			}
		}
		# --- END PARALLEL SCRIPT BLOCK ---
	} # End ForEach-Object -Parallel
	$parallelProcessingDuration = (Get-Date) - $parallelProcessingStartTime
	New-Log "Phase 2 complete. Parallel processing took $($parallelProcessingDuration.ToString("g")). Collected $($allFoundModulesFromParallel.Count) raw entries."
	# --- PHASE 3: Post-processing and Aggregation ---
	New-Log "Phase 3: Starting post-processing and aggregation..."
	$aggregationStartTime = Get-Date
	$allFoundModulesArray = $allFoundModulesFromParallel.ToArray()
	if ($allFoundModulesArray.Count -eq 0) {
		New-Log "No valid module data collected after parallel processing." -Level WARNING
		return [ordered]@{}
	}
	# Deduplicate, Normalize, Group (This part remains the same as your latest version)
	$uniqueModules = $allFoundModulesArray | Group-Object -Property ModuleName, BasePath, ModuleVersionString | ForEach-Object { $_.Group[0] }
	New-Log "Reduced to $($uniqueModules.Count) unique entries after initial grouping."
	$resultModules = [ordered]@{}
	$modulesGroupedByName = $uniqueModules | Where-Object { $null -ne $_.ModuleName -and $_.ModuleName -notmatch '^\d+(\.\d+)+$' } | Group-Object ModuleName
	foreach ($nameGroup in $modulesGroupedByName) {
		$moduleName = $nameGroup.Name
		# New-Log "Post-processing ModuleName: $moduleName" -Level VERBOSE
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
			# $using: not needed, main thread
			$finalSortedModules[$key] = $resultModules[$key]
		}
		else {
			New-Log "Skipping module '$key' as it is in the IgnoredModules list (final filter)." -Level VERBOSE
		}
	}
	$aggregationDuration = (Get-Date) - $aggregationStartTime
	New-Log "Phase 3 (Aggregation) complete in $($aggregationDuration.ToString("g"))."
	$totalFunctionDuration = (Get-Date) - $fileDiscoveryStartTime # Start from very beginning of Phase 1
	New-Log "Get-ModuleInfo completed. Total duration: $($totalFunctionDuration.ToString("g")). Found $($finalSortedModules.Keys.Count) modules." -Level SUCCESS
	return $finalSortedModules
}
#endregion Get-ModuleInfo
#region Get-ModuleUpdateStatus
function Get-ModuleUpdateStatus {
	<#
	.SYNOPSIS
	Checks online repositories for updates to locally installed modules based on provided inventory data, using parallel processing for efficiency.
	.DESCRIPTION
	This function determines if updates are available for modules listed in a local inventory. It operates in two main stages:
	1.  *Stage 1: Online Version Pre-fetching:**
	For each unique module name derived from the input `-ModuleInventory`:
	a. Applies blacklist rules: If a module is blacklisted for all repositories (entry like `ModuleName = '*'`) or for all specified repositories, it's skipped.
	b. For non-blacklisted modules, it launches a parallel thread job (`Start-ThreadJob`) using a script block that calls `Find-Module` against the specified `-Repositories`. This is done twice for each module: once to find the latest stable version and once (with `-AllowPrerelease`) to find the latest pre-release version.
	c. The results (latest stable and latest pre-release found online, or any errors) are collected from these jobs and stored in a thread-safe cache. This stage includes timeout handling for each `Find-Module` job.
	2.  *Stage 2: Parallel Comparison:**
	For each module from the prepared local inventory:
	a. It retrieves the pre-fetched online version data (stable and pre-release) from the cache.
	b. It determines the single "latest" available online version by comparing the pre-fetched stable and pre-release versions using the `Compare-ModuleVersion` helper function (which generally prefers a pre-release if its base version is the same as or newer than the stable version).
	c. This "latest" online version is then compared against the highest locally installed version of the module (also determined using `Compare-ModuleVersion`).
	d. If an update is indicated (`LatestOnlineVersion` > `HighestLocalVersion`) and the `-MatchAuthor` switch is specified, it further checks if the author of the online module matches the author of the local module. An update is only reported if authors match or if `-MatchAuthor` is not used.
	e. If an update is confirmed, it identifies which specific local installation paths of the module do not yet have this latest online version.
	f. Information about modules needing updates is collected into a thread-safe bag.
	Progress is logged throughout the process. The function returns an array of PSCustomObjects, each detailing a module for which an update is available.
	.PARAMETER ModuleInventory
	[Mandatory] A hashtable, typically the output of `Get-ModuleInfo`.
	The keys are module names [string].
	Each value is an array of PSCustomObjects, where each object represents an installed instance of that module and must contain at least:
	- ModuleVersion ([System.Version] or a string parsable to a version)
	- ModuleVersionString ([string])
	- BasePath ([string])
	- IsPrerelease ([bool], optional)
	- PreReleaseLabel ([string], optional)
	- Author ([string], optional)
	.PARAMETER Repositories
	An array of strings specifying the names of the registered PSResource repositories to check for updates. Defaults to `@('PSGallery', 'NuGet')`. Ensure these repositories are registered and accessible.
	.PARAMETER ThrottleLimit
	The maximum number of parallel jobs to run simultaneously. This applies to both the pre-fetching stage (`Start-ThreadJob`'s internal throttle) and the comparison stage (`ForEach-Object -Parallel`). Defaults to `([System.Environment]::ProcessorCount * 2)`.
	.PARAMETER TimeoutSeconds
	The maximum time in seconds that each individual `Find-Module` job in the pre-fetching stage is allowed to run before being timed out. Defaults to 30 seconds.
	.PARAMETER FindModuleTimeoutSeconds
	The maximum time in seconds that each `Find-Module` command within a job is allowed to run before being timed out. Defaults to 10 seconds.
	.PARAMETER BlackList
	A hashtable used to exclude specific modules from update checks or to exclude them from being checked against certain repositories.
	- To completely exclude a module: `@{ 'ModuleName' = '*' }`
	- To exclude a module from specific repositories: `@{ 'ModuleName' = @('Repo1', 'Repo2') }` or `@{ 'ModuleName' = 'Repo1' }`
	Defaults to an empty hashtable (no blacklisting).
	.PARAMETER MatchAuthor
	If specified, an update for a module will only be reported if the author of the latest online version matches the author of the locally installed version. Author matching is case-insensitive and ignores non-alphanumeric characters.
	.INPUTS
	None. This function does not accept direct pipeline input for its main parameters but relies on the `-ModuleInventory` parameter.
	.OUTPUTS
	System.Management.Automation.PSCustomObject[]
	An array of PSCustomObjects, where each object represents a module that has an available update. Each object includes:
	- ModuleName ([string]): The name of the module.
	- Repository ([string]): The name of the repository where the latest version was found.
	- IsPreview ([bool]): True if the latest available online version is a pre-release.
	- PreReleaseVersion ([string]): The pre-release tag of the latest online version (e.g., "beta1"), if applicable.
	- HighestLocalVersion ([System.Version]): The [System.Version] object of the highest version currently installed locally.
	- LatestVersion ([System.Version]): The [System.Version] object of the latest version available online.
	- LatestVersionString ([string]): The full string representation of the latest online version (e.g., "2.1.0" or "2.1.0-beta1").
	- OutdatedModules ([PSCustomObject[]]): An array of objects, each detailing a local installation path that is outdated. Each sub-object has:
	- Path ([string]): The base path of the outdated local installation.
	- InstalledVersion ([string]): The version string of the outdated local installation at that path.
	- Author ([string]): Author of the local module.
	- GalleryAuthor ([string]): Author of the online module.
	Returns an empty array if no updates are found or if the input inventory is empty.
	.EXAMPLE
	PS C:\> $inventory = Get-ModuleInfo -Paths $env:PSModulePath
	PS C:\> $updates = Get-ModuleUpdateStatus -ModuleInventory $inventory -Repositories 'PSGallery' -MatchAuthor -BlackList @{ 'SomeModule' = '*' }
	Checks PSGallery for updates to modules in `$inventory`, requiring author match, and skipping 'SomeModule'.
	.NOTES
	- Requires PowerShell 7.0 or later due to the use of `Start-ThreadJob` and `ForEach-Object -Parallel`.
	- Relies on `Find-Module` cmdlet (from `PowerShellGet` or `Microsoft.PowerShell.PSResourceGet` module). Ensure one of these is installed and functional.
	- Depends on internal helper functions `Compare-ModuleVersion` and an external `New-Log` function.
	- Network connectivity to the specified repositories is required for the online version pre-fetching stage.
	- The accuracy of update detection depends on the quality of the input `$ModuleInventory`.
	.LINK
	Get-ModuleInfo
	Find-Module
	Start-ThreadJob
	ForEach-Object -Parallel
	Compare-ModuleVersion
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][hashtable]$ModuleInventory,
		[string[]]$Repositories = @('PSGallery', 'NuGet'),
		[int]$ThrottleLimit = ([Environment]::ProcessorCount * 2),
		[ValidateRange(1, 3600)][int]$TimeoutSeconds = 30, # Job timeout
		[ValidateRange(1, 60)][int]$FindModuleTimeoutSeconds = 10, # Internal timeout for Find-Module
		[hashtable]$BlackList = @{},
		[switch]$MatchAuthor
	)
	# Quick environment check
	if ($PSVersionTable.PSVersion.Major -lt 7) {
		New-Log "This function requires PowerShell 7 or later. Current version: $($PSVersionTable.PSVersion)" -Level ERROR
		return
	}
	$allModuleNames = $ModuleInventory.Keys | Where-Object { $_ -and $_.Trim() } | Sort-Object -Unique
	if ($allModuleNames.Count -eq 0) {
		New-Log "Module inventory is empty. Nothing to check."
		return @()
	}
	# --- Prepare Local Module Data ---
	$moduleDataArray = @()
	foreach ($moduleNameInLoop in $allModuleNames) {
		# Renamed to avoid conflict with $moduleName later
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
	# --- STAGE 1: Pre-fetch Online Module Versions (Optimized) ---
	New-Log "Starting online version pre-fetching for up to $($moduleDataArray.Count) modules..."
	$overallOperationStartTime = Get-Date
	$onlineModuleVersionsCache = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new() # Thread-safe for direct assignment
	# Script block to find module versions - search repositories one at a time
	$findModuleScriptBlock = {
		param(
			$moduleNameToFetch,
			$repositoriesForJob,
			$findModuleTimeoutSeconds
		)
		# Ensure PSResourceGet cmdlets are available in the thread
		# Import-Module Microsoft.PowerShell.PSResourceGet -ErrorAction SilentlyContinue -Force
		# Import-Module PowerShellGet -ErrorAction SilentlyContinue -Force # For older systems if needed
		$ErrorActionPreference = 'SilentlyContinue' # Let Find-Module try its best for each repo
		$stableResult = $null
		$prereleaseResult = $null
		$fetchError = $null
		# Check each repository one at a time to stop once we find a match
		foreach ($repo in $repositoriesForJob) {
			try {
				$stableModuleFromRepo = Find-Module -Name $moduleNameToFetch -Repository $repo -ErrorAction SilentlyContinue -Verbose:$false
				if ($stableModuleFromRepo) {
					# Found in this repo, take the highest version
					$stableResult = $stableModuleFromRepo | Sort-Object -Property Version -Descending | Select-Object -First 1
					break # Stop checking other repos
				}
			}
			catch {
				$fetchError = "Error finding stable for $moduleNameToFetch in $repo : $($_.Exception.Message)"
				# Continue to next repo
			}
		}
		# If we didn't find a stable version, $stableResult will remain $null
		# Now check for prerelease versions (only if we need to)
		if (-not $stableResult) {
			foreach ($repo in $repositoriesForJob) {
				try {
					$prereleaseModuleFromRepo = Find-Module -Name $moduleNameToFetch -AllowPrerelease -Repository $repo -ErrorAction SilentlyContinue -Verbose:$false |
						Where-Object { ($_.PSObject.Properties['IsPrerelease'] -and $_.PSObject.Properties['IsPrerelease'].Value) -or ($_.Version.ToString() -match '-') }
					if ($prereleaseModuleFromRepo) {
						# Found in this repo, take the highest version
						$prereleaseResult = $prereleaseModuleFromRepo | Sort-Object -Property Version -Descending | Select-Object -First 1
						break # Stop checking other repos
					}
				}
				catch {
					$prereleaseError = "Error finding prerelease for $moduleNameToFetch in $repo : $($_.Exception.Message)"
					$fetchError = ($fetchError + "; " + $prereleaseError).TrimStart('; ')
					# Continue to next repo
				}
			}
		}
		# Even if we find a stable version, we might want to check for a newer prerelease version in the same repo
		if ($stableResult) {
			try {
				$stableRepo = $stableResult.Repository
				$prereleaseModuleFromSameRepo = Find-Module -Name $moduleNameToFetch -AllowPrerelease -Repository $stableRepo -ErrorAction SilentlyContinue -Verbose:$false |
					Where-Object { ($_.PSObject.Properties['IsPrerelease'] -and $_.PSObject.Properties['IsPrerelease'].Value) -or ($_.Version.ToString() -match '-') }
				if ($prereleaseModuleFromSameRepo) {
					$prereleaseResult = $prereleaseModuleFromSameRepo | Sort-Object -Property Version -Descending | Select-Object -First 1
				}
			}
			catch {
				# Non-critical, just log it
				$fetchError = ($fetchError + "; Error checking for prerelease in same repo as stable for $moduleNameToFetch").TrimStart('; ')
			}
		}
		[pscustomobject]@{
			ModuleName    = $moduleNameToFetch
			Stable        = $stableResult
			PreRelease    = $prereleaseResult
			ErrorFetching = $fetchError # Store any caught error messages
			Skipped       = $false
		}
	}
	# Start jobs for each module
	$preFetchJobs = @()
	foreach ($moduleEntry in $moduleDataArray) {
		$moduleNameToFetch = $moduleEntry.ModuleName
		$currentRepositories = $Repositories # Start with all specified function repos
		# Apply blacklist logic for repositories for this module
		if ($BlackList -and $BlackList.ContainsKey($moduleNameToFetch)) {
			$blacklistedReposSetting = $BlackList[$moduleNameToFetch]
			if ($blacklistedReposSetting -eq '*') {
				New-Log "[$moduleNameToFetch] Pre-fetch: Blacklisted ('*'). Skipping online check." -Level DEBUG
				$onlineModuleVersionsCache[$moduleNameToFetch] = [pscustomobject]@{
					ModuleName    = $moduleNameToFetch
					Stable        = $null
					PreRelease    = $null
					ErrorFetching = $null
					Skipped       = $true
				}
				continue
			}
			if ($blacklistedReposSetting -is [array]) {
				$currentRepositories = $Repositories | Where-Object { $blacklistedReposSetting -notcontains $_ }
			}
			elseif ($blacklistedReposSetting -is [string]) {
				$currentRepositories = $Repositories | Where-Object { $_ -ne $blacklistedReposSetting }
			}
			if ($currentRepositories.Count -eq 0) {
				New-Log "[$moduleNameToFetch] Pre-fetch: Blacklisted due to repository exclusion for all specified repos. Skipping." -Level DEBUG
				$onlineModuleVersionsCache[$moduleNameToFetch] = [pscustomobject]@{
					ModuleName    = $moduleNameToFetch
					Stable        = $null
					PreRelease    = $null
					ErrorFetching = $null
					Skipped       = $true
				}
				continue
			}
		}
		if ($currentRepositories.Count -eq 0) {
			# If $Repositories was empty to begin with
			New-Log "[$moduleNameToFetch] Pre-fetch: No repositories specified to check. Skipping." -Level DEBUG
			$onlineModuleVersionsCache[$moduleNameToFetch] = [pscustomobject]@{
				ModuleName    = $moduleNameToFetch
				Stable        = $null
				PreRelease    = $null
				ErrorFetching = $null
				Skipped       = $true
			}
			continue
		}
		$job = Start-ThreadJob -ScriptBlock $findModuleScriptBlock -ThrottleLimit $ThrottleLimit -ArgumentList @($moduleNameToFetch, $currentRepositories, $FindModuleTimeoutSeconds)
		$job.PSObject.Properties.Add([psnoteproperty]::new("ModuleNameForJob", $moduleNameToFetch)) # Tag job with module name
		$preFetchJobs += $job
	}
	# Single log statement for waiting
	New-Log "Waiting for $($preFetchJobs.Count) pre-fetch jobs to complete (Timeout per job: ${TimeoutSeconds}s)..."
	# Initialize counters for job monitoring
	$totalJobs = $preFetchJobs.Count
	$prefetchSync = [System.Collections.Hashtable]::Synchronized(@{
			timeouts  = 0
			completed = 0
			total     = $totalJobs
		})
	# For tracking job completion progress
	$lastCompletedCount = 0
	$completionThreshold = [Math]::Max(1, [Math]::Ceiling($totalJobs / 10)) # Show progress every ~10% completion
	$lastProgressUpdate = Get-Date
	$progressInterval = 5 # seconds
	# Much simpler approach - check every second, report progress on thresholds
	$waitStart = Get-Date
	while ($true) {
		# Sleep briefly to avoid high CPU usage
		Start-Sleep -Milliseconds 250
		# Get current job states
		$runningJobs = @($preFetchJobs | Where-Object State -EQ 'Running')
		$completedJobs = @($preFetchJobs | Where-Object State -EQ 'Completed')
		$runningCount = $runningJobs.Count
		$completedCount = $completedJobs.Count
		# Check if we should update progress
		$timeToUpdate = ((Get-Date) - $lastProgressUpdate).TotalSeconds -ge $progressInterval
		$completionCountChanged = ($completedCount - $lastCompletedCount) -ge $completionThreshold
		if ($timeToUpdate -or $completionCountChanged -or ($lastCompletedCount -eq 0 -and $completedCount -gt 0)) {
			$percentComplete = [math]::Round(($completedCount / $totalJobs) * 100, 1)
			New-Log "Pre-fetch progress: $completedCount/$totalJobs completed ($percentComplete%), $runningCount still running" -Level INFO
			$lastProgressUpdate = Get-Date
			$lastCompletedCount = $completedCount
		}
		# Check for timed-out jobs
		$longRunningJobs = @($runningJobs | Where-Object { ((Get-Date) - $_.PSBeginTime).TotalSeconds -gt $TimeoutSeconds })
		if ($longRunningJobs.Count -gt 0) {
			# Process timed out jobs
			New-Log "Stopping $($longRunningJobs.Count) jobs that exceeded timeout of ${TimeoutSeconds}s" -Level WARNING
			$prefetchSync.timeouts += $longRunningJobs.Count
			foreach ($timeoutJob in $longRunningJobs) {
				$jobModuleName = $timeoutJob.PSObject.Properties["ModuleNameForJob"].Value
				try {
					$partialResults = $timeoutJob | Receive-Job -ErrorAction SilentlyContinue
					if ($partialResults -and $partialResults.ModuleName) {
						$onlineModuleVersionsCache[$partialResults.ModuleName] = $partialResults
					}
					else {
						$onlineModuleVersionsCache[$jobModuleName] = [pscustomobject]@{
							ModuleName = $jobModuleName; Stable = $null; PreRelease = $null
							ErrorFetching = "Pre-fetch job timed out after ${TimeoutSeconds}s"; Skipped = $false
						}
					}
				}
				catch {
					$onlineModuleVersionsCache[$jobModuleName] = [pscustomobject]@{
						ModuleName = $jobModuleName; Stable = $null; PreRelease = $null
						ErrorFetching = "Pre-fetch job timed out after ${TimeoutSeconds}s"; Skipped = $false
					}
				}
				$timeoutJob | Stop-Job -ErrorAction SilentlyContinue
			}
		}
		# Check for aggressive termination criteria
		$elapsedSeconds = ((Get-Date) - $waitStart).TotalSeconds
		$percentCompleted = $completedCount / $totalJobs
		if ($elapsedSeconds -gt 30 -and $percentCompleted -gt 0.90 -and $runningCount -gt 0 -and $runningCount -lt 10) {
			New-Log "Aggressively terminating remaining $runningCount jobs after $([int]$elapsedSeconds)s as they're taking too long" -Level WARNING
			foreach ($slowJob in $runningJobs) {
				$jobModuleName = $slowJob.PSObject.Properties["ModuleNameForJob"].Value
				try {
					$partialResults = $slowJob | Receive-Job -ErrorAction SilentlyContinue
					if ($partialResults -and $partialResults.ModuleName) {
						$onlineModuleVersionsCache[$partialResults.ModuleName] = $partialResults
					}
					else {
						$onlineModuleVersionsCache[$jobModuleName] = [pscustomobject]@{
							ModuleName = $jobModuleName; Stable = $null; PreRelease = $null
							ErrorFetching = "Job terminated for taking too long"; Skipped = $false
						}
					}
				}
				catch {
					$onlineModuleVersionsCache[$jobModuleName] = [pscustomobject]@{
						ModuleName = $jobModuleName; Stable = $null; PreRelease = $null
						ErrorFetching = "Job terminated for taking too long"; Skipped = $false
					}
				}
				$slowJob | Stop-Job -ErrorAction SilentlyContinue
			}
		}
		# EXIT CONDITIONS - in order of importance
		# 1. All jobs have finished - clean exit
		if ($runningCount -eq 0) {
			# Final progress report if needed
			if (((Get-Date) - $lastProgressUpdate).TotalSeconds -gt 0.5) {
				$percentComplete = [math]::Round(($completedCount / $totalJobs) * 100, 1)
				New-Log "Pre-fetch progress: $completedCount/$totalJobs completed ($percentComplete%), 0 still running" -Level INFO
			}
			break
		}
		# 2. Hard timeout - if we've waited too long, just exit regardless
		if ($elapsedSeconds -gt 45) {
			New-Log "Hard timeout after 45 seconds - exiting job monitoring loop" -Level WARNING
			# Stop any remaining running jobs
			foreach ($job in $runningJobs) {
				$job | Stop-Job -ErrorAction SilentlyContinue
			}
			break
		}
	}
	# Receive results from all completed jobs (not already processed)
	$remainingCompletedJobs = @($preFetchJobs | Where-Object { $_.State -eq 'Completed' })
	if ($remainingCompletedJobs.Count -gt 0) {
		New-Log "Processing $($remainingCompletedJobs.Count) remaining completed jobs" -Level VERBOSE
		foreach ($job in $remainingCompletedJobs) {
			$jobModuleName = $job.PSObject.Properties["ModuleNameForJob"].Value
			try {
				$jobOutput = $job | Receive-Job -ErrorAction SilentlyContinue -Wait
				if ($jobOutput -and $jobOutput.ModuleName) {
					$onlineModuleVersionsCache[$jobOutput.ModuleName] = $jobOutput
					$prefetchSync.completed++
				}
			}
			catch {
				# Already logged elsewhere
			}
		}
	}
	# Clean up all jobs
	$preFetchJobs | Remove-Job -Force -ErrorAction SilentlyContinue
	New-Log "Pre-fetch complete: $($prefetchSync.completed)/$($prefetchSync.total) processed, $($prefetchSync.timeouts) timeouts" -Level INFO
	New-Log "Online version pre-fetching complete. Cached data for $($onlineModuleVersionsCache.Count) modules. Timeouts: $($prefetchSync.timeouts)"
	$preFetchDuration = (Get-Date) - $overallOperationStartTime
	New-Log "Pre-fetching (Stage 1) took: $($preFetchDuration.ToString("g"))"
	# Ensure all modules in $moduleDataArray have an entry in $onlineModuleVersionsCache
	foreach ($moduleEntry in $moduleDataArray) {
		if (-not $onlineModuleVersionsCache.ContainsKey($moduleEntry.ModuleName)) {
			New-Log "[$($moduleEntry.ModuleName)] No pre-fetched data found post-job processing. Marking as error/skipped." -Level WARNING
			$onlineModuleVersionsCache[$moduleEntry.ModuleName] = [pscustomobject]@{
				ModuleName    = $moduleEntry.ModuleName
				Stable        = $null
				PreRelease    = $null
				ErrorFetching = "Data not found in pre-fetch cache."
				Skipped       = $true
			}
		}
	}
	# --- STAGE 2: Parallel Processing with Pre-fetched Data ---
	$results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
	$sync = [System.Collections.Hashtable]::Synchronized(@{
			processed = 0; updates = 0; errors = 0
			total = $validModuleCountForProcessing; startTime = Get-Date
		})
	$NewLogDef = ${function:New-Log}.ToString()
	$CompareModuleVersionDef = ${function:Compare-ModuleVersion}.ToString()
	New-Log "Starting parallel update comparison for $($moduleDataArray.Count) modules (Throttle: $ThrottleLimit)..." -Level SUCCESS
	$moduleDataArray | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
		$script:ErrorActionPreference = 'Continue' # Set for this parallel runspace
		$moduleData = $_
		$moduleName = $moduleData.ModuleName
		$sync = $using:sync
		$VerbosePreference = $using:VerbosePreference
		${function:New-Log} = $using:NewLogDef
		${function:Compare-ModuleVersion} = $using:CompareModuleVersionDef
		$results = $using:results
		$matchAuthor = $using:MatchAuthor.IsPresent
		$onlineCache = $using:onlineModuleVersionsCache # Access the main cache
		try {
			$highestLocalInstall = $moduleData.HighestLocalInstall
			if (-not $highestLocalInstall) {
				New-Log "[$moduleName] Pre-processed HighestLocalInstall object missing. Skipping." -Level WARNING
				$sync.errors++
				return
			}
			$onlineModuleData = $null
			if (-not $onlineCache.TryGetValue($moduleName, [ref]$onlineModuleData)) {
				New-Log "[$moduleName] Could not retrieve pre-fetched data from cache. Skipping." -Level WARNING
				$sync.errors++
				return
			}
			if ($onlineModuleData.Skipped -or (-not $onlineModuleData.Stable -and -not $onlineModuleData.PreRelease -and $onlineModuleData.ErrorFetching)) {
				if ($onlineModuleData.Skipped) {
					New-Log "[$moduleName] Skipping processing: Pre-fetch marked as skipped (e.g., blacklist)." -Level WARNING
				}
				else {
					New-Log "[$moduleName] Skipping processing: No online versions found or error during pre-fetch: $($onlineModuleData.ErrorFetching)" -Level VERBOSE
				}
				return
			}
			$stableModule = $onlineModuleData.Stable
			$preReleaseModule = $onlineModuleData.PreRelease
			$galleryModule = $null # This will be the selected latest online module info
			if ($stableModule -and $preReleaseModule) {
				$stableVerStr = $stableModule.Version.ToString()
				$prereleaseVerStr = $preReleaseModule.Version.ToString()
				$prereleaseLbl = $preReleaseModule.PSObject.Properties['PreRelease'].Value # Access PreRelease label, could be $null
				$fullPrereleaseVerStr = if (-not [string]::IsNullOrEmpty($prereleaseLbl)) { "$prereleaseVerStr-$prereleaseLbl" } else { $prereleaseVerStr }
				try {
					$isPreReleaseNewer = Compare-ModuleVersion -VersionA $stableVerStr -VersionB $fullPrereleaseVerStr -ReturnBoolean
					$galleryModule = if ($isPreReleaseNewer) { $preReleaseModule } else { $stableModule }
					New-Log "[$moduleName] Cache: Both stable ($stableVerStr) and prerelease ($fullPrereleaseVerStr) found. Selected: $($galleryModule.Version)$(if($galleryModule.PSObject.Properties['PreRelease'].Value){ "-$($galleryModule.PSObject.Properties['PreRelease'].Value)" })" -Level VERBOSE
				}
				catch {
					$galleryModule = $stableModule
					New-Log "[$moduleName] Cache: Error comparing versions. Defaulting to stable ($stableVerStr)." -Level VERBOSE
					$sync.errors++
				}
			}
			elseif ($preReleaseModule) {
				$galleryModule = $preReleaseModule
				New-Log "[$moduleName] Cache: Only prerelease version found: $($preReleaseModule.Version)$(if($preReleaseModule.PSObject.Properties['PreRelease'].Value){"-$($preReleaseModule.PSObject.Properties['PreRelease'].Value)"})" -Level VERBOSE
			}
			elseif ($stableModule) {
				$galleryModule = $stableModule
				New-Log "[$moduleName] Cache: Only stable version found: $($stableModule.Version)" -Level VERBOSE
			}
			if (-not $galleryModule) {
				New-Log "[$moduleName] No suitable online module version determined from pre-fetched data." -Level VERBOSE
				return
			}
			# --- Continue with version comparison and other checks ---
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
				New-Log "[$moduleName] Invalid online ('$latestOnlineStr') or local ('$highestLocalStr') version string. Skipping." -Level VERBOSE
				$sync.errors++
				return
			}
			$needsOverallUpdate = $false
			try {
				if (Compare-ModuleVersion -VersionA $highestLocalStr -VersionB $latestOnlineStr -ReturnBoolean) {
					$needsOverallUpdate = $true
					New-Log "[$moduleName] Comparison: Online ($latestOnlineStr) IS NEWER than Local ($highestLocalStr)" -Level VERBOSE
				}
				else {
					New-Log "[$moduleName] Comparison: Online ($latestOnlineStr) not newer than Local ($highestLocalStr)" -Level VERBOSE
				}
			}
			catch {
				New-Log "[$moduleName] Error comparing versions $highestLocalStr and $latestOnlineStr : $($_.Exception.Message)." -Level ERROR
				$sync.errors++
				return
			}
			if ($needsOverallUpdate -and $matchAuthor) {
				# ... (Author matching logic - ensure $galleryModule.Author and $highestLocalInstall.Author are accessed)
				$localAuthor = $highestLocalInstall.Author
				$galleryAuthor = $galleryModule.Author # Assuming Find-Module result has .Author
				$authorsMatch = $false
				$normalizedLocalAuthor = [Regex]::Replace($localAuthor, '[^a-zA-Z0-9]', '')
				$normalizedGalleryAuthor = [Regex]::Replace($galleryAuthor, '[^a-zA-Z0-9]', '')
				if ($normalizedLocalAuthor -and $normalizedGalleryAuthor -and $normalizedGalleryAuthor -eq $normalizedLocalAuthor) {
					# Simple equality for now
					New-Log "[$moduleName] Author Match: OK (Local: '$localAuthor', Online: '$galleryAuthor')." -Level VERBOSE
					$authorsMatch = $true
				}
				if (-not $authorsMatch) {
					New-Log "[$moduleName] Skipping update: -MatchAuthor specified and authors do not match (Local: '$localAuthor', Online: '$galleryAuthor')." -Level VERBOSE
					$needsOverallUpdate = $false
				}
			}
			if ($needsOverallUpdate) {
				# ... (Your logic for $outdatedInstallationsDetailed)
				$outdatedInstallationsDetailed = @()
				$allLocalInstalls = $moduleData.AllParsedVersions # This is an array of PSCustomObjects
				# Group by BasePath to check each unique installation location
				$installsByPath = $allLocalInstalls | Group-Object -Property BasePath
				foreach ($pathGroup in $installsByPath) {
					$versionsInThisPath = $pathGroup.Group # Array of local installs at this path
					$latestOnlineVersionFoundInThisPath = $false
					foreach ($installedVersionEntry in $versionsInThisPath) {
						if ($installedVersionEntry.ModuleVersionString -eq $latestOnlineStr) {
							$latestOnlineVersionFoundInThisPath = $true; break
						}
					}
					if (-not $latestOnlineVersionFoundInThisPath) {
						# This path lacks the latest, add all versions from this path as outdated
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
						HighestLocalVersion = $highestLocalInstall.ModuleVersion # This is a [Version] object
						LatestVersion       = [version]$galleryModule.Version # Ensure this is also a [Version] object
						LatestVersionString = $latestOnlineStr
						OutdatedModules     = $uniqueOutdatedModules
						Author              = $highestLocalInstall.Author
						GalleryAuthor       = $galleryModule.Author
					}
					$results.Add($resultObject)
					New-Log "[$moduleName] Update found: Local '$highestLocalStr' -> Online '$latestOnlineStr'. $($uniqueOutdatedModules.Count) outdated paths." -Level SUCCESS
					$sync.updates++
				}
				else {
					New-Log "[$moduleName] Overall update flagged, but no specific paths found lacking '$latestOnlineStr'. This might indicate all paths are up-to-date or an issue in logic." -Level VERBOSE
				}
			}
		}
		catch {
			New-Log "[$moduleName] Unhandled error in parallel processing" -Level ERROR
			$sync.errors++
		}
		finally {
			$sync.processed++
			$percentComplete = [math]::Min(100, ($sync.processed / $sync.total) * 100)
			[int]$numberOfMoudules = [math]::Ceiling($sync.total / 20)
			if (($sync.processed % $numberOfMoudules -eq 0) -or $sync.processed -eq $sync.total) {
				$elapsed = (Get-Date) - $sync.startTime
				$avgTimePerModule = if ($sync.processed -gt 0) { $elapsed.TotalSeconds / $sync.processed } else { 0 }
				$remainingModules = $sync.total - $sync.processed
				$etaSeconds = $avgTimePerModule * $remainingModules
				$etaString = if ($etaSeconds -gt 0) {
					$eta = [timespan]::FromSeconds($etaSeconds)
					"$($eta.Minutes.ToString('00')):$($eta.Seconds.ToString('00'))"
				}
				else {
					"00:00"
				}
				$progressMsg = "Progress: $([math]::Round($percentComplete, 0))% ($($sync.processed)/$($sync.total)) | Updates: $($sync.updates), Errors: $($sync.errors) | Elapsed: {0:mm\:ss} | ETA: $etaString" -f $elapsed
				New-Log $progressMsg
			}
		}
	}
	# --- Final Results Processing ---
	$moduleObjects = @($results) # $results is the ConcurrentBag
	$finalOverallTime = (Get-Date) - $overallOperationStartTime # Corrected total time calculation
	New-Log "Pre-fetching (Stage 1) duration: $($preFetchDuration.ToString("g"))"
	New-Log "Comparison (Stage 2) duration: $(($finalOverallTime - $preFetchDuration).ToString("g"))" # Duration of Stage 2
	# Use $finalOverallTime for the final summary
	New-Log "Completed check of $($sync.total) modules in $($finalOverallTime.ToString("g")). Found $($moduleObjects.Count) modules needing updates." -Level SUCCESS
	if ($prefetchSync.timeouts -gt 0) {
		# Report timeouts from the prefetch stage
		New-Log "$($prefetchSync.timeouts) module pre-fetch checks timed out." -Level WARNING
	}
	if ($sync.errors -gt 0) {
		# Report errors from Stage 2's sync context
		New-Log "Encountered $($sync.errors) errors during comparison processing." -Level WARNING
	}
	return $moduleObjects | Sort-Object ModuleName
}
#endregion Get-ModuleUpdateStatus
#region Update-Modules
function Update-Modules {
	<#
	.SYNOPSIS
	Installs the latest versions of modules identified as outdated, optionally cleaning up old versions.
	.DESCRIPTION
	This function takes an array of module update objects (typically the output from `Get-ModuleUpdateStatus`) and attempts to install the specified newer version for each module using a 3-phase parallel pipeline.

	Phase 1 - Pre-processing (sequential, main thread):
	For each module, the function:
	1.  Determines the exact target version string to install (e.g., "2.1.0" or "2.1.0-preview3") using the `LatestVersionString` property from the input object (or constructing it from `LatestVersion` and `PreReleaseVersion` if `LatestVersionString` is not directly available). It parses this version using `Parse-ModuleVersion` to get structured version info.
	2.  Identifies the source repository and the unique base paths where outdated versions of this module currently exist (from the `OutdatedModules.Path` property of the input object).
	3.  Evaluates `ShouldProcess` prompts on the main thread (required since `ShouldProcess` does not work in parallel runspaces). Approved modules are collected with all pre-computed data for Phase 2.

	Phase 2 - Parallel installation (ForEach-Object -Parallel):
	All approved modules are installed concurrently using `ForEach-Object -Parallel` with a throttle limit of `Min(moduleCount, Max(4, ProcessorCount * 2))`. Each parallel runspace:
	1.  Restores helper function definitions (`New-Log`, `Install-PSModule`) and imports `Microsoft.PowerShell.PSResourceGet`.
	2.  Calls `Install-PSModule` which prioritizes `Save-PSResource` to place the module into the parent directory of the destination base paths. If that fails, it falls back to `Install-PSResource -Scope AllUsers` and then `Install-Module -Scope AllUsers`.
	3.  Reports progress with ETA calculation via a synchronized hashtable and `ConcurrentBag` for thread-safe result collection.

	Phase 3 - Post-processing and cleaning (sequential, main thread):
	If the `-Clean` switch is specified AND an update was successful, the function calls `Remove-OutdatedVersions` sequentially (to avoid filesystem conflicts between parallel operations).
	`Remove-OutdatedVersions` uses `Uninstall-PSResource` for each old version directory and may fall back to `Remove-Item -Recurse -Force`.
	Certain critical modules (e.g., 'PowerShellGet', 'Microsoft.PowerShell.PSResourceGet') are excluded from cleaning by default.
	The function reports the success or failure for each module update attempt, including which base paths were successfully updated, which failed, and (if -Clean was used) which old version paths were removed. A progress bar can optionally be displayed using `-UseProgressBar`.
	.PARAMETER OutdatedModules
	[Mandatory, ValueFromPipeline] An array of PSCustomObjects detailing the modules to update. Each object must contain at least the following properties (this format matches the output of `Get-ModuleUpdateStatus`):
	- ModuleName ([string]): The name of the module.
	- LatestVersion ([System.Version] or a string parsable into a version): The base version object of the update (e.g., for "2.1.0-preview3", this would be "2.1.0"). Used to derive the base version string.
	- LatestVersionString ([string], optional but preferred): The full string representation of the latest version to install (e.g., "2.1.0", "2.1.0-preview3"). If not present, it's constructed from `LatestVersion` and `PreReleaseVersion`.
	- Repository ([string]): The name of the repository from which to install the update.
	- IsPreview ([bool]): True if the `LatestVersionString` (or equivalent constructed version) represents a pre-release version.
	- PreReleaseVersion ([string], optional): The pre-release tag of the `LatestVersionString` (e.g., "preview3"), if applicable. Used with `LatestVersion` if `LatestVersionString` is absent.
	- OutdatedModules ([PSCustomObject[]]): An array of objects, each representing an outdated installation. Each of these sub-objects must have at least a `Path` property ([string]) indicating the base installation path of an outdated version (e.g., "C:\Program Files\PowerShell\Modules\MyModule").
	.PARAMETER Clean
	If specified, the function will attempt to remove the directories of the older versions of a module *after* its update has been successfully installed in a given path. This operation uses `Uninstall-PSResource` and, as a fallback, `Remove-Item -Recurse -Force`. Use with caution.
	.PARAMETER UseProgressBar
	If specified, displays a progress bar tracking the overall module update process.
	.PARAMETER PreRelease
	If specified, signals the user's intent to allow installation of pre-release versions if they are identified in the `-OutdatedModules` input. The function primarily relies on the `IsPreview` flag and the version string format (e.g., "x.y.z-tag") within the input objects to determine if a module version is a pre-release and to pass appropriate flags (`-Prerelease`, `-AllowPrerelease`) to the underlying installation cmdlets. This switch serves as a confirmation and may influence verbose logging but does not override the data-driven pre-release determination from the input.
	.INPUTS
	System.Management.Automation.PSCustomObject[]
	Accepts an array of module update objects (matching the output structure of `Get-ModuleUpdateStatus`) via the pipeline or the `-OutdatedModules` parameter.
	.OUTPUTS
	System.Management.Automation.PSCustomObject[]
	Returns an array of PSCustomObjects, each summarizing the result of the update attempt for a single module:
	- ModuleName ([string]): The name of the module.
	- NewVersionPreRelease ([string]): The full version string (including any pre-release tag) of the version that was attempted/installed (e.g., "2.1.0-preview3"). If not a pre-release, this will be the same as NewVersion.
	- NewVersion ([string]): The base version string (e.g., "2.1.0") of the version that was attempted/installed.
	- UpdatedPaths ([string[]]): An array of base paths (from the input module's `OutdatedModules.Path`) where the update was successfully installed or confirmed to exist.
	- FailedPaths ([string[]]): An array of base paths (from the input module's `OutdatedModules.Path`) where the update attempt failed.
	- OverallSuccess ([bool]): True if `FailedPaths` is empty and `UpdatedPaths` is not empty, indicating the module was updated in all targeted original locations or a general successful install occurred that covered at least one target.
	- CleanedPaths ([string[]] or [string]): Present if `-Clean` was specified. An array of full paths to the old version directories that were successfully removed. If cleaning was attempted but no items were removed for a module, this might be a string message like "Cleaning attempted but no paths removed." or "Skipped by ShouldProcess".
	.EXAMPLE
	# Scenario: Find outdated modules and update them, cleaning old versions.
	PS C:\> $moduleInventory = Get-ModuleInfo -Paths ($env:PSModulePath -split ';')
	PS C:\> $updatesToInstall = Get-ModuleUpdateStatus -ModuleInventory $moduleInventory -Repositories 'PSGallery'
	PS C:\> if ($updatesToInstall) {
	PS C:\>     $updateResults = $updatesToInstall | Update-Modules -Clean -UseProgressBar -PreRelease -Verbose
	PS C:\>     $updateResults | Format-Table ModuleName, NewVersionPreRelease, OverallSuccess, CleanedPaths
	PS C:\> } else { Write-Host "No module updates found." }
	This example first gets module inventory, then checks for updates against PSGallery. If updates are found, it pipes them to `Update-Modules` to install them (respecting pre-release status from input, confirmed by -PreRelease), cleans up old versions after successful updates, shows a progress bar, and provides verbose output. Administrator privileges are typically required.
	.EXAMPLE
	# Scenario: Update only a specific module from a previously generated list of updates.
	PS C:\> $specificUpdate = $allUpdates | Where-Object ModuleName -eq 'PSScriptAnalyzer'
	PS C:\> Update-Modules -OutdatedModules $specificUpdate -Verbose
	Updates only the 'PSScriptAnalyzer' module (if present in `$specificUpdate`), without cleaning old versions or showing a progress bar by default.
	.NOTES
	- Requires Administrator privileges to install modules to system locations (e.g., Program Files) and to remove old versions from these locations.
	- Depends on `Microsoft.PowerShell.PSResourceGet` module for `*-PSResource*` cmdlets. It may fall back to `Install-Module` (from `PowerShellGet`) if `Install-PSResource` encounters issues.
	- Relies on internal helper functions `Install-PSModule` and `Remove-OutdatedVersions`, and an external `New-Log` function.
	- The `-Clean` operation uses `Remove-Item -Recurse -Force` as a fallback. Always review modules and paths before using `-Clean`.
	- The success of installing to specific original `Destinations` (derived from outdated module paths) relies heavily on `Save-PSResource`. Fallback installation methods might install to a default `AllUsers` path, which may or may not align with all originally intended outdated paths. The `UpdatedPaths` output reflects where the new version is confirmed relative to the input paths.
	.LINK
	Get-ModuleUpdateStatus
	Get-ModuleInfo
	Save-PSResource
	Install-PSResource
	Uninstall-PSResource
	Install-Module
	Remove-Item
	Write-Progress
	Parse-ModuleVersion
	#>
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[Parameter(ValueFromPipeline, Mandatory)][Object[]]$OutdatedModules,
		[switch]$Clean,
		[switch]$UseProgressBar,
		[switch]$PreRelease # Helps signal intent for handling preview modules specified in input
	)
	begin {
		$aggregateResults = [System.Collections.Generic.List[object]]::new()
		$batchModules = @() # Collect all modules from pipeline input first
		New-Log "Initializing module update process." -Level VERBOSE
	}
	process {
		New-Log "Receiving $($OutdatedModules.Count) module(s) from pipeline/parameter." -Level VERBOSE
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
		# --- PHASE 1: Pre-process all modules on the main thread ---
		# Version parsing, path validation, and ShouldProcess must run sequentially on the main thread.
		# ShouldProcess does not work inside ForEach-Object -Parallel runspaces.
		New-Log "Phase 1: Pre-processing $total module(s) (version parsing, path validation, ShouldProcess)..."
		$modulesToInstall = [System.Collections.Generic.List[object]]::new() # Modules approved for parallel install
		$current = 0
		foreach ($module in $batchModules) {
			$moduleName = $module.ModuleName
			$current++
			[string]$targetVersionString = $($module.LatestVersion)
			[version]$latestVer = $null # This will hold the base [version] object
			# Try parsing the target string
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
						NewVersionPreRelease = if ($module.IsPreview) { "$($targetVersionString)-$($module.PreReleaseVersion)" } # Report the problematic string
						NewVersion           = $targetVersionString
						UpdatedPaths         = @()
						FailedPaths          = @("Version parsing failed: $targetVersionString")
						OverallSuccess       = $false
					})
				continue # Skip to the next module
			}
			$latestVer = $parsedTargetVersion.ModuleVersion # The [version] object
			[string]$baseVerStr = $latestVer.ToString() # Base version string without pre-release tag
			$repository = $module.Repository
			# Determine if it's a preview install based on input object's flag *and* if the version string *looks* like a prerelease
			# Use $PreRelease switch to confirm intent if needed, but primarily rely on the version data
			$installAsPreview = $parsedTargetVersion.IsPrerelease # Default to true if the version string looks like a prerelease
			if ($installAsPreview) {
				if ($PreRelease.IsPresent) {
					New-Log "[$moduleName] version '$preReleaseVersion' appears to be pre-release, and -PreRelease switch is present. Proceeding with preview installation logic." -Level VERBOSE
				}
				else {
					New-Log "[$moduleName] version '$preReleaseVersion' appears to be pre-release (based on version string parsing). Proceeding with preview installation logic even without -PreRelease switch." -Level VERBOSE
				}
			}
			else {
				New-Log "[$moduleName] version '$targetVersionString' does not appear to be pre-release. Proceeding with standard installation logic." -Level VERBOSE
			}
			# Get distinct base paths where old versions exist
			# Ensure Path exists before trying to Test-Path
			$outdatedPaths = @($module.OutdatedModules | Where-Object { $null -ne $_.Path } | Select-Object -ExpandProperty Path -Unique | Where-Object { $_ -and (Test-Path $_ -PathType Container -Verbose:$false) })
			$outdatedVersions = @($module.OutdatedModules.InstalledVersion | Select-Object -Unique) # Get unique old version strings
			if ($outdatedPaths.Count -eq 0) {
				New-Log "[$moduleName][$current/$total] Skipping module: No valid outdated base paths found where old versions exist. (Checked: $($module.OutdatedModules.Path -join '; '))." -Level WARNING
				$aggregateResults.Add([PSCustomObject]@{
						ModuleName           = $moduleName
						NewVersionPreRelease = if ($module.IsPreview) { $preReleaseVersion }
						NewVersion           = $baseVerStr
						UpdatedPaths         = @()
						FailedPaths          = @("No valid source paths provided or accessible")
						OverallSuccess       = $false
					})
				continue
			}
			New-Log "[$moduleName] Target base paths based on outdated locations: $($outdatedPaths -join '; ')" -Level VERBOSE
			# --- ShouldProcess check on the main thread (not supported in parallel runspaces) ---
			$displayVersion = if ($installAsPreview -and $preReleaseVersion) { $preReleaseVersion } else { $targetVersionString }
			if ($PSCmdlet.ShouldProcess("$moduleName v$displayVersion", "Install from repository '$repository' to paths: $($outdatedPaths -join ', ')")) {
				# Approved for installation - add to the parallel batch with all pre-computed data
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
						# Pre-evaluate ShouldProcess for cleaning on main thread
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
						OverallSuccess       = $false
						CleanedPaths         = if ($Clean.IsPresent) { "Skipped by ShouldProcess" } else { $null }
					})
			}
		}
		if ($modulesToInstall.Count -eq 0) {
			New-Log "No modules approved for installation after pre-processing. Exiting."
			return $aggregateResults
		}
		# --- PHASE 2: Parallel module installation ---
		# Uses ForEach-Object -Parallel following the same pattern as Get-ModuleUpdateStatus comparison stage.
		# Helper function definitions are passed via $using: and restored in each parallel runspace.
		$installThrottleLimit = [Math]::Min($modulesToInstall.Count, [Math]::Max(4, [System.Environment]::ProcessorCount * 2))
		New-Log "Phase 2: Starting parallel installation of $($modulesToInstall.Count) module(s) (Throttle: $installThrottleLimit)..." -Level SUCCESS
		$parallelInstallResults = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
		$installSync = [System.Collections.Hashtable]::Synchronized(@{
				processed = 0; successes = 0; failures = 0
				total = $modulesToInstall.Count; startTime = Get-Date
			})
		# Capture helper function definitions for the parallel runspaces
		$NewLogDef = ${function:New-Log}.ToString()
		$InstallPSModuleDef = ${function:Install-PSModule}.ToString()
		$modulesToInstall | ForEach-Object -ThrottleLimit $installThrottleLimit -Parallel {
			$script:ErrorActionPreference = 'Continue'
			$preparedModule = $_
			$moduleName = $preparedModule.ModuleName
			$VerbosePreference = $using:VerbosePreference
			$installSync = $using:installSync
			$parallelInstallResults = $using:parallelInstallResults
			# Restore helper functions into this parallel runspace
			${function:New-Log} = $using:NewLogDef
			${function:Install-PSModule} = $using:InstallPSModuleDef
			Import-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Global -Force
			try {
				New-Log "[$moduleName] Installing version [$($preparedModule.TargetVersionString)]..." -Level VERBOSE
				$installResult = Install-PSModule -ModuleName $moduleName -TargetVersionString $preparedModule.TargetVersionString -PreReleaseVersion $preparedModule.PreReleaseVersion -RepositoryName $preparedModule.Repository -IsPreview $preparedModule.InstallAsPreview -Destinations $preparedModule.OutdatedPaths -ErrorAction SilentlyContinue
				$finalResult = [PSCustomObject]@{
					ModuleName           = $moduleName
					NewVersionPreRelease = if ($preparedModule.IsPreview) { $preparedModule.PreReleaseVersion }
					NewVersion           = $preparedModule.BaseVerStr
					UpdatedPaths         = @($installResult.UpdatedPaths)
					FailedPaths          = if ($installResult.FailedPaths) { @($installResult.FailedPaths) } else { @() }
					OverallSuccess       = ($installResult.FailedPaths.Count -eq 0 -and $installResult.UpdatedPaths.Count -gt 0)
					CleanedPaths         = $null # Placeholder - cleaning is handled in Phase 3
					# Carry forward pre-computed data needed for Phase 3 (cleaning)
					_OutdatedVersions    = $preparedModule.OutdatedVersions
					_LatestVer           = $preparedModule.LatestVer
					_PreReleaseVersion   = $preparedModule.PreReleaseVersion
					_CleanApproved       = $preparedModule.CleanApproved
				}
				$parallelInstallResults.Add($finalResult)
				if ($finalResult.OverallSuccess) {
					$installSync.successes++
					New-Log "[$moduleName] Parallel install completed successfully." -Level SUCCESS
				}
				else {
					$installSync.failures++
					New-Log "[$moduleName] Parallel install completed with failures: $($finalResult.FailedPaths -join '; ')" -Level WARNING
				}
			}
			catch {
				New-Log "[$moduleName] Unhandled error during parallel install: $($_.Exception.Message)" -Level ERROR
				$parallelInstallResults.Add([PSCustomObject]@{
						ModuleName           = $moduleName
						NewVersionPreRelease = if ($preparedModule.IsPreview) { $preparedModule.PreReleaseVersion }
						NewVersion           = $preparedModule.BaseVerStr
						UpdatedPaths         = @()
						FailedPaths          = @("Parallel install error: $($_.Exception.Message)")
						OverallSuccess       = $false
						CleanedPaths         = $null
						_OutdatedVersions    = $preparedModule.OutdatedVersions
						_LatestVer           = $preparedModule.LatestVer
						_PreReleaseVersion   = $preparedModule.PreReleaseVersion
						_CleanApproved       = $false
					})
				$installSync.failures++
			}
			finally {
				$installSync.processed++
				$percentComplete = [math]::Min(100, ($installSync.processed / $installSync.total) * 100)
				[int]$progressInterval = [math]::Max(1, [math]::Ceiling($installSync.total / 10))
				if (($installSync.processed % $progressInterval -eq 0) -or $installSync.processed -eq $installSync.total) {
					$elapsed = (Get-Date) - $installSync.startTime
					$avgTimePerModule = if ($installSync.processed -gt 0) { $elapsed.TotalSeconds / $installSync.processed } else { 0 }
					$remainingModules = $installSync.total - $installSync.processed
					$etaSeconds = $avgTimePerModule * $remainingModules
					$etaString = if ($etaSeconds -gt 0) {
						$eta = [timespan]::FromSeconds($etaSeconds)
						"$($eta.Minutes.ToString('00')):$($eta.Seconds.ToString('00'))"
					}
					else {
						"00:00"
					}
					$progressMsg = "Install progress: $([math]::Round($percentComplete, 0))% ($($installSync.processed)/$($installSync.total)) | OK: $($installSync.successes), Fail: $($installSync.failures) | Elapsed: {0:mm\:ss} | ETA: $etaString" -f $elapsed
					New-Log $progressMsg
				}
			}
		}
		$parallelDuration = (Get-Date) - $installSync.startTime
		New-Log "Phase 2 complete. Parallel installation took $($parallelDuration.ToString("g")). Successes: $($installSync.successes), Failures: $($installSync.failures)."
		# --- PHASE 3: Post-processing - Cleaning and result assembly (sequential on main thread) ---
		# Cleaning runs sequentially to avoid filesystem conflicts (Uninstall-PSResource, Remove-Item).
		$installResultsArray = @($parallelInstallResults)
		$useClean = $Clean.IsPresent
		foreach ($finalResult in $installResultsArray) {
			$moduleName = $finalResult.ModuleName
			if ($useClean) {
				$finalResult.CleanedPaths = @() # Initialize for -Clean mode
			}
			# --- Progress Bar Update (sequential, on main thread) ---
			if ($UseProgressBar.IsPresent) {
				$progressParams = @{
					Activity         = "Updating PowerShell Modules"
					Status           = "Post-processing: $moduleName"
					PercentComplete  = 100
					CurrentOperation = if ($useClean -and $finalResult.OverallSuccess) { "[$moduleName] Cleaning old versions.." } else { "[$moduleName] Finalizing.." }
				}
				Write-Progress @progressParams
			}
			# --- Optional Cleaning Step ---
			if ($finalResult.OverallSuccess -and $useClean) {
				New-Log "[$moduleName] Update successful to paths: $($finalResult.UpdatedPaths -join '; '). Proceeding with cleaning old versions..."
				if ($finalResult._CleanApproved) {
					# Pass the successfully parsed [version] object and pre-release tag of the NEW version to avoid removing it
					$cleanedPathsResult = Remove-OutdatedVersions -ModuleName $moduleName -ModuleBasePaths $finalResult.UpdatedPaths -LatestVersion $finalResult._LatestVer -PreReleaseVersion $finalResult._PreReleaseVersion -ErrorAction SilentlyContinue
					# Ensure we only process string paths
					$cleanedPathsResult = $cleanedPathsResult | Where-Object { $_ -is [string] }
					if ($cleanedPathsResult -and $cleanedPathsResult.Count -gt 0) {
						$finalResult.CleanedPaths = $cleanedPathsResult
						New-Log "[$moduleName] Successfully cleaned $($cleanedPathsResult.Count) old items: $($cleanedPathsResult -join '; ')" -Level SUCCESS
					}
					else {
						New-Log "[$moduleName] Cleaning step completed. No old versions were removed in the updated paths. This may be expected or could indicate issues (e.g., permissions, module not found by Uninstall-PSResource). Check paths: $($finalResult.UpdatedPaths -join '; ')" -Level VERBOSE
					}
				}
				else {
					New-Log "[$moduleName] Skipped cleaning due to ShouldProcess user choice." -Level WARNING
				}
			}
			elseif ($useClean -and -not $finalResult.OverallSuccess) {
				New-Log "[$moduleName] Skipping cleaning as the update was not fully successful (Failed Paths: $($finalResult.FailedPaths -join '; '))." -Level VERBOSE
			}
			# Set CleanedPaths to explicit message if Clean was used but skipped/failed
			if ($useClean -and $finalResult.OverallSuccess -and ($null -eq $finalResult.CleanedPaths -or $finalResult.CleanedPaths.Count -eq 0)) {
				$finalResult.CleanedPaths = "Cleaning attempted but no paths removed."
			}
			# Remove internal tracking properties before adding to final results
			$finalResult.PSObject.Properties.Remove('_OutdatedVersions')
			$finalResult.PSObject.Properties.Remove('_LatestVer')
			$finalResult.PSObject.Properties.Remove('_PreReleaseVersion')
			$finalResult.PSObject.Properties.Remove('_CleanApproved')
			$aggregateResults.Add($finalResult)
		}
		$successCount = ($aggregateResults | Where-Object { $_.OverallSuccess }).Count
		$failCount = $total - $successCount
		New-Log "Update process finished for $total modules. Successful Updates: $successCount, Failed/Partial Updates: $failCount." -Level SUCCESS
		if ($failCount -gt 0) {
			New-Log "Modules with failures or partial updates:" -Level WARNING
			$aggregateResults | Where-Object {
				-not $_.OverallSuccess
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
		# Summarize cleaning results if -Clean was used
		if ($Clean.IsPresent) {
			$cleanedModules = $aggregateResults | Where-Object {
				$null -ne $_.CleanedPaths -and $_.CleanedPaths -is [array] -and $_.CleanedPaths.Count -gt 0
			}
			$cleanAttemptedButNoneRemoved = $aggregateResults | Where-Object {
				$_.CleanedPaths -is [string] -and $_.CleanedPaths -like '*attempted*'
			}
			$cleanSkipped = $aggregateResults | Where-Object {
				$_.CleanedPaths -is [string] -and $_.CleanedPaths -like '*Skipped*'
			}
			New-Log "Cleaning Summary: $($cleanedModules.Count) modules had old versions successfully removed. $($cleanAttemptedButNoneRemoved.Count) modules had cleaning attempted but no items removed. $($cleanSkipped.Count) modules had cleaning skipped (ShouldProcess/Failure)." -Level DEBUG
		}
		return $aggregateResults # Return the detailed results for each module
	}
}
#endregion Update-Modules
# -------------------------------------------------------------------------------#
# Helper Functions (Assume these are defined in the same scope)					 #
# Minimal help blocks for context, focusing on parameters used by Update-Modules #
# -------------------------------------------------------------------------------#
#region Install-PSModule (Helper)
function Install-PSModule {
	<#
	.SYNOPSIS
	(Helper Function) Installs or saves a specific module version to target locations using the best available method.
	.DESCRIPTION
	This is an internal helper function called by `Update-Modules`. It attempts to install a specific version of a module.
	The primary installation strategy is to use `Save-PSResource`. This cmdlet is called to save the module into the parent directory of each path specified in the `-Destinations` parameter. For example, if a destination base path is 'C:\Modules\MyModule', `Save-PSResource` attempts to save the module content to 'C:\Modules', which would result in a structure like 'C:\Modules\MyModule\NewVersion'. This method is preferred for its control over the installation path.
	If `Save-PSResource` fails for any destination or entirely, or if some destinations remain unaddressed, the function attempts fallback methods:
	1. `Install-PSResource -Scope AllUsers`: This installs the module to a system-wide location.
	2. `Install-Module -Scope AllUsers`: If `Install-PSResource` also fails, this older cmdlet is tried.
	These fallback methods typically install to a default PSModulePath location for the AllUsers scope and offer less precise control if multiple such paths exist.
	The function determines success by checking if the new version is present in the expected location after each attempt. It returns an object indicating which of the original destination base paths were successfully updated (either directly via `Save-PSResource` to their parent, or indirectly if a fallback installation resulted in the module being available at an intended destination) and which failed.
	.PARAMETER ModuleName
	[Mandatory] The name of the module to install.
	.PARAMETER TargetVersionString
	[Mandatory] The base version string of the module to install (e.g., "1.2.3"). If installing a pre-release, this should be the numeric base version part, and the full pre-release version string must be supplied via the `-PreReleaseVersion` parameter.
	.PARAMETER RepositoryName
	[Mandatory] The name of the PSResource repository to install from.
	.PARAMETER IsPreview
	[Mandatory] A boolean indicating if the target version is a pre-release. This controls whether `-PreRelease` (for `Save-PSResource`, `Install-PSResource`) or `-AllowPrerelease` (for `Install-Module`) flags are used with underlying cmdlets.
	.PARAMETER Destinations
	[Mandatory] An array of strings, where each string is a base path of an existing module installation (e.g., "C:\Program Files\WindowsPowerShell\Modules\MyModule"). The function will attempt to install the new version such that it would reside within the parent of these paths (e.g., if "C:\Modules\MyModule" is a destination, `Save-PSResource` targets "C:\Modules").
	.PARAMETER PreReleaseVersion
	[Optional] The full pre-release version string (e.g., "2.0.0-beta1"). If specified and `-IsPreview` is true, this exact version string is used with the installation cmdlets (`Save-PSResource -Version`, `Install-PSResource -Version`, `Install-Module -RequiredVersion`). This parameter is crucial for installing specific pre-releases.
	.OUTPUTS
	PSCustomObject
	An object with the following properties:
	- UpdatedPaths ([System.Collections.Generic.List[string]]): A list of base paths from the original `-Destinations` parameter where the new version was successfully installed or confirmed to exist.
	- FailedPaths ([System.Collections.Generic.List[string]]): A list of base paths from the original `-Destinations` parameter where the update attempt failed for that specific path.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$ModuleName,
		[Parameter(Mandatory)][string]$TargetVersionString, # Use full string now
		[Parameter(Mandatory)][string]$RepositoryName,
		[Parameter(Mandatory)][bool]$IsPreview,
		[Parameter(Mandatory)][string[]]$Destinations,
		[string]$PreReleaseVersion
	)
	$result = [PSCustomObject]@{
		UpdatedPaths = [System.Collections.Generic.List[string]]::new()
		FailedPaths  = [System.Collections.Generic.List[string]]::new()
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
		Version             = $TargetVersionString # Use the full string
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
	# Only add -PreRelease switch when true — passing $false can cause issues with some PSResourceGet versions
	if ($IsPreview) { $commonSaveParams['PreRelease'] = $true }
	$pathsToRetry = [System.Collections.Generic.List[string]]::new()
	foreach ($destinationBasePath in $Destinations) {
		# Save-PSResource needs the PARENT of the destination base path
		# e.g., if destination is C:\Modules\MyMod, save path is C:\Modules
		$saveTargetDir = Split-Path $destinationBasePath -Parent -ErrorAction SilentlyContinue
		if (-not $saveTargetDir -or -not (Test-Path $saveTargetDir -PathType Container)) {
			New-Log "[$moduleName] Install-PSModule: Cannot determine valid parent directory from destination '$destinationBasePath'. Skipping Save-PSResource for this path." -Level WARNING
			$pathsToRetry.Add($destinationBasePath) # Mark for fallback
			continue
		}
		New-Log "[$moduleName] Attempting Save-PSResource for version [$TargetVersionString] to '$saveTargetDir'..." -Level VERBOSE
		$saveRes = $null # Renamed variable to avoid conflict
		try {
			$saveRes = Save-PSResource @commonSaveParams -Path $saveTargetDir # Use Stop to catch errors
		}
		catch {
			New-Log "[$moduleName] Save-PSResource failed for v$TargetVersionString to '$saveTargetDir': $($_.Exception.Message)" -Level WARNING
			$saveRes = $null # Ensure it's null on error
		}
		$savedItem = $saveRes | Select-Object -Last 1 # Get the actual saved item if multiple dependencies were saved
		# Normalize Prerelease comparison: treat $null and '' as equivalent
		$expectedPrerelease = if ($PreReleaseVersion) { ($PreReleaseVersion -split '-')[-1] } else { '' }
		$actualPrerelease = if ($savedItem -and $savedItem.Prerelease) { $savedItem.Prerelease } else { '' }
		if ($savedItem -and "$($savedItem.Version)" -eq "$TargetVersionStringOrig" -and $actualPrerelease -eq $expectedPrerelease) {
			# Construct the expected final module path after save
			$expectedModuleVersionPath = Join-Path -Path $saveTargetDir -ChildPath "$ModuleName\$TargetVersionStringOrig"
			New-Log "[$moduleName] Successfully saved version [$TargetVersionString] via Save-PSResource. Expected path: '$expectedModuleVersionPath'" -Level SUCCESS
			$result.UpdatedPaths.Add($destinationBasePath) # Report success for the target base path
		}
		else {
			$pathsToRetry.Add($destinationBasePath) # Mark this base path for potential fallback
		}
	}
	# --- Fallback Installation Methods ---
	# If Save-PSResource didn't work for all paths, or failed entirely, try Install-* cmdlets
	# Only retry if there are paths specifically marked for retry
	if ($pathsToRetry.Count -gt 0) {
		# Fallback 1: Install-PSResource (Modern, preferred fallback)
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
			# Only add -PreRelease switch when true
			if ($IsPreview) { $psResourceParams['PreRelease'] = $true }
			New-Log "[$moduleName] Attempting Install-PSResource for version [$($psResourceParams.Version)]..." -Level VERBOSE
			$installRes1 = $null # Renamed variable
			$installRes1 = Install-PSResource @psResourceParams
			$installedItem1 = $installRes1 | Select-Object -Last 1
			# Normalize Prerelease comparison: treat $null and '' as equivalent
			$expectedPrerelease1 = if ($PreReleaseVersion) { ($PreReleaseVersion -split '-')[-1] } else { '' }
			$actualPrerelease1 = if ($installedItem1 -and $installedItem1.Prerelease) { $installedItem1.Prerelease } else { '' }
			if ($installedItem1 -and "$($installedItem1.Version)" -eq "$TargetVersionStringOrig" -and $actualPrerelease1 -eq $expectedPrerelease1) {
				$installedLocationFallback = $installedItem1.InstalledLocation # Path like C:\Modules\ModuleName\Version
				New-Log "[$moduleName] Successfully installed version [$($psResourceParams.Version)] via Install-PSResource to '$installedLocationFallback'" -Level SUCCESS
				$fallbackInstallSucceeded = $true
			}
		}
		catch {
			New-Log "[$moduleName] Install-PSResource failed for v${TargetVersionString}: $($_.Exception.Message)" -Level VERBOSE
		}
		if (-not $fallbackInstallSucceeded) {
			try {
				$installModuleParams = @{
					Name               = $ModuleName
					RequiredVersion    = $TargetVersionString # Use the full string
					Scope              = 'AllUsers' # Fallback usually targets AllUsers
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
				# Only add -AllowPrerelease if it's a preview AND the command actually supports the parameter
				# PS7's PSResourceGet Install-Module proxy does NOT support -AllowPrerelease
				if ($IsPreview) {
					$installModuleCmd = Get-Command -Name 'Install-Module' -ErrorAction SilentlyContinue
					if ($installModuleCmd -and $installModuleCmd.Parameters.ContainsKey('AllowPrerelease')) {
						$installModuleParams['AllowPrerelease'] = $true
					}
				}
				# Ensure repository is trusted to prevent interactive prompts from the PSResourceGet Install-Module proxy
				try { Set-PSResourceRepository -Name $RepositoryName -Trusted -ErrorAction SilentlyContinue } catch {}
				New-Log "[$moduleName] Attempting Install-Module for version [$($installModuleParams.RequiredVersion)]..." -Level DEBUG
				$installRes2 = $null # Renamed variable
				$installRes2 = Install-Module @installModuleParams
				$installedItem2 = $installRes2 | Select-Object -Last 1
				if ($installedItem2 -and $($installedItem2.Version) -eq $TargetVersionStringOrig -and $installedItem2.PSObject.Properties.Match('Prerelease').Value -eq $(($PreReleaseVersion -split '-')[-1])) {
					# Access Prerelease differently for Install-Module output
					$installedLocationFallback = $installedItem2.InstalledLocation # Path like C:\Modules\ModuleName\Version
					New-Log "[$moduleName] Successfully installed version [$($installModuleParams.RequiredVersion)] via Install-Module to '$installedLocationFallback'" -Level SUCCESS
					$fallbackInstallSucceeded = $true
				}
			}
			catch {
				New-Log "[$moduleName] Install-Module (last fallback) failed: $($_.Exception.Message)" -Level ERROR
			}
		}
		# --- Check if fallback installation satisfied any retry paths ---
		if ($fallbackInstallSucceeded -and $installedLocationFallback) {
			# Fallback install typically goes to a default path.
			# We need to see if any of the *intended* destination base paths correspond to this default path.
			$fallbackBaseInstallPath = Split-Path $installedLocationFallback -Parent -Verbose:$false # Get the parent (e.g., C:\Modules\ModuleName)
			foreach ($retryPath in $pathsToRetry) {
				# Compare the intended base path with the actual base path where the fallback installed
				if ($retryPath -eq $fallbackBaseInstallPath) {
					# Only add if not already added by Save-PSResource
					if ($result.UpdatedPaths -notcontains $retryPath) {
						$result.UpdatedPaths.Add($retryPath)
						New-Log "[$moduleName] Fallback installation to '$installedLocationFallback' satisfied the intended destination base path '$retryPath'." -Level DEBUG
					}
				}
			}
		}
		# --- Finalize Failed Paths ---
		# Any path marked for retry that isn't now in UpdatedPaths is considered failed
		foreach ($retryPath in $pathsToRetry) {
			if ($result.UpdatedPaths -notcontains $retryPath) {
				# Only add if not already marked as failed (though it shouldn't be)
				if ($result.FailedPaths -notcontains $retryPath) {
					$result.FailedPaths.Add($retryPath)
					New-Log "[$moduleName] Marking path '$retryPath' as failed after Save-PSResource and fallback install attempts." -Level DEB
				}
			}
		}
	}
	# --- Consolidate and Report ---
	$result.UpdatedPaths = $result.UpdatedPaths | Select-Object -Unique -Verbose:$false
	$result.FailedPaths = $result.FailedPaths | Select-Object -Unique -Verbose:$false
	# Remove any path from FailedPaths that is also in UpdatedPaths (shouldn't happen often, but safety)
	$result.FailedPaths = $result.FailedPaths | Where-Object { $result.UpdatedPaths -notcontains $_ }
	if ($result.UpdatedPaths.Count -gt 0 -and $result.FailedPaths.Count -eq 0) {
		New-Log "[$moduleName] Successfully updated to version [$TargetVersionString] for all target destinations ($($result.UpdatedPaths -join '; '))." -Level SUCCESS
	}
	elseif ($result.UpdatedPaths.Count -gt 0) {
		New-Log "[$moduleName] Partially updated to version [$TargetVersionString]. Succeeded Base Paths: ($($result.UpdatedPaths -join '; ')). Failed Base Paths: ($($result.FailedPaths -join '; '))" -Level WARNING
	}
	else {
		New-Log "[$moduleName] Failed to update to version [$TargetVersionString] for any target destinations. Failed Base Paths: ($($Destinations -join '; '))" -Level WARNING # Report original destinations if all failed
		# Ensure all original destinations are marked as failed if UpdatedPaths is empty
		if ($result.UpdatedPaths.Count -eq 0) {
			$result.FailedPaths.Clear()
			$result.FailedPaths.AddRange($Destinations)
			$result.FailedPaths = $result.FailedPaths | Select-Object -Unique -Verbose:$false # Ensure uniqueness again
		}
	}
	return $result
}
#endregion Install-PSModule (Helper)
#region Remove-OutdatedVersions (Helper)
function Remove-OutdatedVersions {
	<#
	.SYNOPSIS
	(Helper Function) Removes specified older versions of a module from given base paths, preserving the specified latest version.
	.DESCRIPTION
	This is an internal helper function called by `Update-Modules` when the -Clean switch is used and an update was successful.
	Its purpose is to remove directories of older versions of a specific module. It operates on each path provided in `-ModuleBasePaths`. A module base path is the directory containing different version folders of that module (e.g., "C:\Program Files\WindowsPowerShell\Modules\MyModule").
	1.  It constructs the full string of the newly installed version (the one to KEEP) using the provided `-LatestVersion` (a [System.Version] object) and `-PreReleaseVersion` (the full pre-release string, if applicable).
	2.  For each directory in `-ModuleBasePaths`, it lists all subdirectories.
	3.  It identifies subdirectories that represent older versions by checking if their names look like version strings (e.g., "1.0.0", "1.1.0-beta") AND are different from the version string constructed in step 1 (and its base numeric form).
	4.  For each identified old version folder:
	a. It first attempts to remove the old version using `Uninstall-PSResource -Name <ModuleName> -Version <OldVersionString>`. This is the preferred method as it handles module unregistration.
	b. If `Uninstall-PSResource` fails to remove the directory or if the command itself fails (e.g., module not found by `Uninstall-PSResource` under that version string), it falls back to using `Remove-Item -Recurse -Force` on the specific old version folder. This is a more direct but potentially less clean removal.
	5.  Modules listed in the `-DoNotClean` parameter (which defaults to include 'PowerShellGet' and 'Microsoft.PowerShell.PSResourceGet') are skipped entirely.
	The function returns an array of strings, where each string is the full path to an old version directory that was successfully removed.
	.PARAMETER ModuleName
	[Mandatory] The name of the module whose old versions are to be cleaned.
	.PARAMETER ModuleBasePaths
	[Mandatory] An array of strings, where each string is a base installation path for the module (e.g., "C:\Program Files\WindowsPowerShell\Modules\MyModule"). The function will look for version subdirectories (e.g., "1.0.0", "1.1.0-beta") within these paths to clean.
	.PARAMETER LatestVersion
	[Mandatory] A [System.Version] object representing the base numeric version of the module that should be KEPT (not removed) (e.g., for "2.1.0-preview3", this is the [version]"2.1.0").
	.PARAMETER DoNotClean
	[Optional] An array of module names that should never be cleaned, even if old versions are found. Defaults to `@('PowerShellGet', 'Microsoft.PowerShell.PSResourceGet')`.
	.PARAMETER PreReleaseVersion
	[Optional] The full pre-release version string of the latest version to KEEP (e.g., "2.0.0-beta1"). This is used in conjunction with `LatestVersion` to accurately construct the exact full version string of the module installation that must be preserved during cleanup. If the latest installed version is not a pre-release, this parameter should be $null or not provided.
	.OUTPUTS
	String[]
	An array of full paths to the old module version directories that were successfully removed. Returns an empty array if no versions were removed, if cleaning was skipped for the module, or if errors occurred.
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
	# Construct the full string of the version to KEEP, including prerelease tag if present
	[string]$latestVersionString = $LatestVersion.ToString()
	[string]$latestVersionFullString = if ($PreReleaseVersion) { $PreReleaseVersion } else { $latestVersionString }
	New-Log "[$moduleName] Starting cleanup of old versions (keeping v$latestVersionFullString)..." -Level VERBOSE
	$cleanedItems = [System.Collections.Generic.List[string]]::new()
	$attemptedUninstallFor = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
	# Unload the module from current session to prevent file locks during cleanup
	Remove-Module -Name $ModuleName -Force -ErrorAction SilentlyContinue
	foreach ($basePath in $ModuleBasePaths) {
		if (-not (Test-Path $basePath -PathType Container)) {
			New-Log "[$moduleName] Base path '$basePath' for cleaning does not exist or is not a directory. Skipping." -Level WARNING
			continue
		}
		New-Log "[$moduleName] Checking for old versions within '$basePath'..." -Level VERBOSE
		$versionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue -Verbose:$false | Where-Object {
			$folderName = $_.Name
			# Must look like a version
			if ($folderName -notmatch '^\d+(\.\d+){1,4}(-.+)?$') { return $false }
			# Don't match the version we want to keep (both with and without prerelease tag)
			if ($folderName -eq $latestVersionFullString -or $folderName -eq $latestVersionString) { return $false }
			# Also compare as [version] objects for normalized form mismatches
			$folderBase = ($folderName -split '-', 2)[0]
			$keepBase = ($latestVersionFullString -split '-', 2)[0]
			$folderVer = $null; $keepVer = $null
			if ([version]::TryParse($folderBase, [ref]$folderVer) -and [version]::TryParse($keepBase, [ref]$keepVer)) {
				if ($folderVer -eq $keepVer) {
					# Base versions match — also check prerelease suffix
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
				# Parse this old version's own prerelease tag (if any) from the folder name
				$oldPreReleaseTag = $null
				if ($versionString -match '^(\d+(\.\d+){1,3})-(.*)') {
					$oldPreReleaseTag = $Matches[3]
				}
				# Avoid double-processing if Uninstall-PSResource was already tried for this version string
				if ($attemptedUninstallFor.Contains($versionString)) {
					New-Log "[$moduleName] Already attempted Uninstall-PSResource for version '$versionString'. Checking existence of '$folderPath'." -Level VERBOSE
					if (-not (Test-Path -LiteralPath $folderPath)) {
						$cleanedItems.Add($folderPath) # Add if it was removed by a previous attempt
						$removed = $true
					}
				}
				if (-not $removed) {
					New-Log "[$moduleName] Found potential old version folder: '$folderPath'. Attempting removal..." -Level DEBUG
					if ($PSCmdlet.ShouldProcess($folderPath, "Remove module '$ModuleName' version '$versionString'")) {
						$uninstalledViaCmdlet = $false
						# Attempt to uninstall via PSResource using the old version's own version string
						try {
							New-Log "[$moduleName] Attempting Uninstall-PSResource with Version '$versionString'..." -Level VERBOSE
							Uninstall-PSResource -Name $ModuleName -Version $versionString -Scope AllUsers -Confirm:$false -Verbose:$false -SkipDependencyCheck -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
							$uninstalledViaCmdlet = $true
						}
						catch {
							New-Log "[$moduleName] Uninstall-PSResource failed with version '$versionString'" -Level WARNING
						}
						$attemptedUninstallFor.Add($versionString)
						Start-Sleep -Milliseconds 50
						# Check if the folder was removed
						if (-not (Test-Path -LiteralPath $folderPath)) {
							New-Log "[$moduleName] Successfully removed '$folderPath' (verified after Uninstall-PSResource attempt)." -Level SUCCESS
							$cleanedItems.Add($folderPath)
							$removed = $true
						}
						else {
							# Log appropriate warning based on what happened
							if ($uninstalledViaCmdlet) {
								New-Log "[$moduleName] Uninstall-PSResource command succeeded with version '$versionString', but folder '$folderPath' still exists. Proceeding to Remove-Item fallback. (Permissions?)" -Level WARNING
							}
							else {
								New-Log "[$moduleName] Uninstall-PSResource failed or did not remove folder '$folderPath'. Proceeding to Remove-Item fallback. (Permissions?)" -Level VERBOSE
							}
						}
						# Fallback: Use Remove-Item if folder still exists
						if (-not $removed -and (Test-Path -LiteralPath $folderPath -Verbose:$false)) {
							New-Log "[$moduleName] Attempting Remove-Item -Recurse -Force on '$folderPath'..." -Level DEBUG
							try {
								Remove-Item -LiteralPath $folderPath -Recurse -Force -ErrorAction Stop -Verbose:$false | Out-Null
								Start-Sleep -Milliseconds 50
								if (-not (Test-Path -LiteralPath $folderPath)) {
									New-Log "[$moduleName] Successfully removed old folder '$folderPath' via Remove-Item." -Level SUCCESS
									$cleanedItems.Add($folderPath)
								}
								else {
									New-Log "[$moduleName] Remove-Item ran but folder '$folderPath' still exists. Manual cleanup might be needed. (Permissions?)" -Level WARNING
								}
							}
							catch {
								New-Log "[$moduleName] Failed to remove old folder '$folderPath' via Remove-Item. Manual cleanup might be required. (Permissions?)" -Level WARNING
							}
						}
					}
					else {
						New-Log "[$moduleName] Skipped removal of '$folderPath' due to ShouldProcess user choice." -Level DEBUG
					}
				}
			}
		}
		else {
			New-Log "[$moduleName] No version-like subdirectories found to clean within '$basePath'" -Level DEBUG
		}
	}
	# Ensure we only return string paths (filter out any accidental booleans or other types)
	$cleanedPathsToReturn = $cleanedItems | Where-Object { $_ -is [string] } | Select-Object -Unique -Verbose:$false
	New-Log "[$moduleName] Finished cleaning attempt. Removed $($cleanedPathsToReturn.Count) items." -Level DEBUG
	return $cleanedPathsToReturn # Return list of removed directory paths
}
#endregion Remove-OutdatedVersions (Helper)
#region Get-ManifestVersionInfo (Helper)
function Get-ManifestVersionInfo {
	<#
	.SYNOPSIS
	(Helper Function) Extracts version and metadata from parsed module manifest data or by analyzing a module file path.
	.DESCRIPTION
	Internal helper function for `Get-ModuleInfo`. It processes module metadata from different sources:
	1. If `-Quick` is specified and `$ResData` (typically the direct output object from `Test-ModuleManifest`) is provided, it performs a simplified extraction directly from this object's properties (like Name, Version, Author, ModuleBase).
	2. If `-ModuleFilePath` is provided (and not in `-Quick` mode, or if `$ResData` is absent or insufficient), it calls another helper, `Get-ModuleformPath`. `Get-ModuleformPath` attempts to infer module details (name, version, base path, author) by analyzing the file path structure and may leverage `Get-Module` for confirmation.
	In both cases, where a version string is obtained, it utilizes the `Parse-ModuleVersion` helper to interpret the version string, identify pre-release status, and extract any pre-release labels.
	The function is designed to be flexible and could also process data that might originate from `Import-PowerShellDataFile` if such data (a hashtable with expected keys) were piped to its `$ResData` parameter, though `Get-ModuleInfo` primarily uses `Test-ModuleManifest` output or file path analysis via `Get-ModuleformPath`.
	.PARAMETER ResData
	[Optional] An object or hashtable containing manifest data. When used with `-Quick`, this is typically the output object from `Test-ModuleManifest`. It can also be data from `Import-PowerShellDataFile`.
	.PARAMETER Quick
	[Optional] If specified, and `$ResData` is provided and is an object with expected properties (like from `Test-ModuleManifest`), the function performs a direct and simplified extraction of module metadata from `$ResData`.
	.PARAMETER ModuleFilePath
	[Optional] A string path to the module's manifest file (.psd1). This is used if `$ResData` is not provided, or if not in `-Quick` mode, or if the `-Quick` extraction from `$ResData` is insufficient, to attempt inferring module information by analyzing the file path via the `Get-ModuleformPath` helper.
	.OUTPUTS
	PSCustomObject
	A PSCustomObject containing the extracted module information, or $null if essential information (like module name or version) cannot be determined. The object includes:
	- ModuleName ([string]): The name of the module.
	- ModuleVersion ([System.Version]): The parsed [System.Version] object of the module (the numeric part).
	- ModuleVersionString ([string]): The original, full version string as found or parsed.
	- IsPreRelease ([bool]): True if the version is identified as a pre-release by `Parse-ModuleVersion`.
	- PreReleaseLabel ([string]): The pre-release label (e.g., "beta1", "rc.2"), if applicable, from `Parse-ModuleVersion`.
	- BasePath ([string]): The determined base path of the module (the module's root directory, e.g., "C:\Modules\MyModule").
	- Author ([string]): The author of the module, if available from the manifest data or `Get-ModuleformPath`.
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
#endregion Get-ManifestVersionInfo (Helper)
#region Test-IsResourceFile (Helper)
function Test-IsResourceFile {
	<#
	.SYNOPSIS
	(Helper Function) Checks if a given file path likely points to a localization/resource file rather than a primary module manifest.
	.DESCRIPTION
	Internal helper for `Get-ModuleInfo`, used to filter out files that are probably not main module manifests but rather culture-specific resource files or localization data.
	It uses several regular expression patterns to make this determination:
	1.  Checks if the path contains a directory named like a culture code (e.g., 'en-US', 'de-DE', 'fr').
	2.  Checks if the path contains common resource directory names such as 'Resources', 'Localization', 'Languages', 'i18n', 'l10n'.
	3.  Checks for common resource file naming patterns, where the filename itself indicates a resource (e.g., 'MyModule.Strings.psd1', 'Errors.xml', 'LocalizedData.psd1').
	4.  Checks if the filename itself is a culture code followed by '.psd1' or '.xml' (e.g., 'en-US.psd1').
	If any of these patterns match, the function assumes the file is a resource file.
	.PARAMETER Path
	[Mandatory] The string path to the file to be tested (typically a .psd1 or .xml file).
	.OUTPUTS
	System.Boolean
	Returns `$true` if the path matches patterns commonly associated with resource or localization files; otherwise, returns `$false`.
	#>
	[CmdletBinding()]
	param(
		[string]$Path
	)
	# Simple checks based on common patterns for localization resource files
	# 1. Culture code directory (e.g., en-US, de-DE)
	if ($Path -match '\\([a-z]{2,3}-[A-Z]{2,3})\\') {
		# Allows 2 or 3 letter lang codes
		New-Log "Path '$Path' matches culture code pattern: $($Matches[1])" -Level VERBOSE
		return $true
	}
	# 2. Common resource directory names
	if ($Path -match '\\(Resources|Resource|Localization|Localizations|Languages|Lang|Cultures|Culture|i18n|l10n)\\') {
		New-Log "Path '$Path' contains common resource directory name: $($Matches[1])" -Level VERBOSE
		return $true
	}
	# 3. Common resource file naming patterns (often ending in .resources.psd1 or similar)
	if ($Path -match '(Resources|Strings|Localized|Messages|Text|Errors|Labels|UI)\.(psd1|xml)$') {
		New-Log "Path '$Path' matches common resource file name pattern: $($Matches[0])" -Level VERBOSE
		return $true
	}
	# 4. Simpler culture code format (e.g., \en\, \fr\)
	if ($Path -match '\\([a-z]{2,3})\\[^\\]+\.(psd1|xml)$') {
		New-Log "Path '$Path' matches simple culture code pattern: $($Matches[1])" -Level VERBOSE
		return $true
	}
	# 5. File name IS a culture code (e.g., en-US.psd1) - less common for manifests but possible
	$fileName = Split-Path -Path $Path -Leaf
	if ($fileName -match '^[a-z]{2,3}-[A-Z]{2,3}\.(psd1|xml)$') {
		New-Log "Path '$Path' filename matches culture code: $fileName" -Level VERBOSE
		return $true
	}
	# Default: Not identified as a resource file
	New-Log "Path '$Path' did not match resource file patterns." -Level VERBOSE
	return $false
}
#endregion Test-IsResourceFile (Helper)
#region Resolve-ModuleVersion (Helper)
function Resolve-ModuleVersion {
	<#
    .SYNOPSIS
    (Helper Function) Parses a given version string and optionally refines it by checking installed modules at a specific path.
    .DESCRIPTION
    Internal helper function, primarily used by `Get-ModuleformPath`.
    1. It first attempts to parse the provided (optional) `-VersionString` using the `Parse-ModuleVersion` helper. This gives an initial interpretation of the version.
    2. Regardless of whether `-VersionString` was provided or successfully parsed, it then queries for installed modules matching the mandatory `-ModuleName`. This query is filtered: it only considers module installations whose `ModuleBase` (the directory containing the module name folder, e.g., "C:\Program Files\WindowsPowerShell\Modules") is a parent of or the same as the directory containing the input `-Path` (or the `-Path` itself if it's a directory). This helps pinpoint the specific module installation relevant to the input `-Path`.
    3. If a matching installed module is found from this filtered `Get-Module` call, its version string is taken and parsed using `Parse-ModuleVersion`. This result, if valid, becomes the definitive version information.
    This process helps in scenarios where a version string might be absent, ambiguous, or incompletely represented in a path or filename, allowing `Get-Module` to provide a more authoritative version if the module is properly installed and discoverable at that location.
    .PARAMETER VersionString
    [Optional] An initial version string to parse (e.g., "1.0.0", "2.1.0-beta"). If not provided or unparsable, the function relies more heavily on `Get-Module`.
    .PARAMETER ModuleName
    [Mandatory] The name of the module for which to resolve the version.
    .PARAMETER Path
    [Mandatory] The full path to a file (e.g., a .psd1) or a directory within the module's installation. This path is crucial for filtering `Get-Module` results to the correct installation if multiple versions or locations exist.
    .OUTPUTS
    PSCustomObject
    An object with two properties:
    - Module ([PSModuleInfo]): The `PSModuleInfo` object from `Get-Module` if a matching installed module was found and its version was used. This can be `$null` if no such module is found or if the initial `VersionString` parse was used and deemed sufficient.
    - VersionPattern ([PSCustomObject]): The output from `Parse-ModuleVersion` based on the finally determined version string (either from the input `VersionString` or from a found `PSModuleInfo` object). This object contains fields like `ModuleVersion` ([System.Version]), `ModuleVersionString`, `IsPrerelease`, `PreReleaseLabel`. This will be `$null` if no version could be parsed at all.
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
#endregion Resolve-ModuleVersion (Helper)
#region Get-ModuleformPath (Helper)
function Get-ModuleformPath {
	<#
    .SYNOPSIS
    (Helper Function) Attempts to guess the module name, version, and module base path from a given file/directory path.
    .DESCRIPTION
    Internal helper function, primarily for `Get-ManifestVersionInfo`. It aims to extract module metadata (name, version, base path, author) by analyzing the structure of an input path, which typically points to a module manifest file (.psd1) or a directory within a module structure.
    The function employs several strategies:
    1.  **Standard Path Regex:** It first tries to match a common PowerShell module path pattern like '...\Modules\ModuleName\VersionString\...' (e.g., 'C:\Program Files\Modules\MyModule\1.0.0\MyModule.psd1'). If this pattern matches, it extracts the module name, version string, and the module's base path (e.g., 'C:\Program Files\Modules\MyModule').
    2.  **Heuristic Path Analysis:** If the standard regex doesn't match, it uses a more heuristic approach:
        a.  It determines a `potentialModuleBasePath` (e.g., if input is '...\MyModule\1.0.0\file.psd1', this could be '...\MyModule\1.0.0' or '...\MyModule').
        b.  The `potentialModuleName` is taken as the leaf name of this `potentialModuleBasePath`.
        c.  **Version as Directory Name:** If this `potentialModuleName` itself looks like a version string (e.g., '1.0.0'), it assumes the `potentialModuleBasePath` was actually a version-specific directory. It then considers the parent of this directory as the true `BasePath`, and the leaf name of that parent as the `ModuleName`. The version string is the folder name that looked like a version.
        d.  **General Case:** If `potentialModuleName` doesn't look like a version and isn't simply "Modules", it's treated as the module name. The function then tries to find a version string by:
            i.  Looking for version-like strings in any of the parent directory names in the input path.
            ii. If the input path is a .psd1 file, by reading the manifest content and looking for a `ModuleVersion = '...'` line.
    3.  **Version Resolution:** In all cases where a module name and a potential version string are identified (either from path parsing or manifest content), it calls the `Resolve-ModuleVersion` helper. `Resolve-ModuleVersion` parses the version string and can further refine it by checking `Get-Module -ListAvailable` for an installed module at that path context, providing a more authoritative version and author if available.
    The `BasePath` returned is intended to be the module's root installation directory (e.g., "C:\Modules\MyModule"), which is the parent of any version-specific subfolders. Version strings are interpreted using `Parse-ModuleVersion` (via `Resolve-ModuleVersion`).
    .PARAMETER Path
    [Mandatory] The string path, typically to a module manifest file (.psd1) or a directory that is part of a module's file structure.
    .OUTPUTS
    PSCustomObject
    A PSCustomObject containing the inferred module details. If essential details like ModuleName or ModuleVersion cannot be reliably inferred, their corresponding properties will be $null.
    - ModuleName ([string]): The inferred name of the module.
    - BasePath ([string]): The inferred base path (root directory) of the module (e.g., "C:\Modules\MyModule").
    - ModuleVersion ([System.Version]): The parsed [System.Version] object of the module (numeric part).
    - ModuleVersionString ([string]): The original, full version string that was identified and parsed.
    - IsPrerelease ([bool]): True if the version is identified as a pre-release by `Parse-ModuleVersion`.
    - PreReleaseLabel ([string]): The pre-release label (e.g., "beta1"), if applicable.
    - Author ([string]): The author of the module, typically resolved via `Get-Module` by the `Resolve-ModuleVersion` helper.
    Returns an object with null properties for fields that could not be determined.
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
	$versionPattern = '\d+(\.\d+){1,3}(-.+)?' # Example: 1.0.0, 1.2.3.4, 2.0.0-beta
	$regexVersionOnly = "^$versionPattern$" # Anchored version pattern for exact match
	# --- Pattern 1: Standard Structure ...\Modules\ModuleName\Version[\Something] ---
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
	# --- Pattern 2 (Previously Fallback 2): Look for module name and version in path or manifest ---
	$potentialModuleBasePath = $null
	if (Test-Path -Path $normalizedPath -PathType Container -Verbose:$false) {
		$potentialModuleBasePath = $normalizedPath
		New-Log "Pattern 2: Input '$normalizedPath' is a directory. Assuming it's the ModulePath." -Level VERBOSE
	}
	else {
		# If input is a file, its parent directory is the candidate for ModulePath
		$potentialModuleBasePath = Split-Path -Path $normalizedPath -Parent -ErrorAction SilentlyContinue -Verbose:$false
		New-Log "Pattern 2: Input '$normalizedPath' is a file. Assuming parent '$potentialModuleBasePath' is the ModulePath." -Level VERBOSE
	}
	if ($potentialModuleBasePath) {
		$potentialModuleName = Split-Path -Path $potentialModuleBasePath -Leaf -ErrorAction SilentlyContinue -Verbose:$false
		New-Log "Pattern 2: Potential ModuleName (leaf of ModulePath) is '$potentialModuleName'." -Level VERBOSE
		# Guard against the potential module name being a version string
		if ($potentialModuleName -match $regexVersionOnly) {
			New-Log "Pattern 2: determined ModuleName '$potentialModuleName' which looks like a version. This is usually incorrect." -Level WARNING
			# Try to get the correct module name from the parent
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
		# Prevent 'Modules' folder itself being named the module
		elseif ($potentialModuleName -and $potentialModuleName -ne 'Modules') {
			New-Log "Pattern 2: Found potential module name '$potentialModuleName', attempting to find version information" -Level VERBOSE
			$version = $null
			# 1. Look in any of the parent directory names
			$pathParts = $normalizedPath -split '\\'
			foreach ($part in $pathParts) {
				if ($part -match $regexVersionOnly) {
					$version = $part
					New-Log "Pattern 2: Found version '$version' in path component" -Level VERBOSE
					break
				}
			}
			# 2. Try looking for version in manifest file if this is a psd1
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
			# Now use our helper to resolve the version or get it from module info
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
	# No patterns matched, return null values
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
#endregion Get-ModuleformPath (Helper)
#region Get-ModuleInfoFromXml (Helper)
function Get-ModuleInfoFromXml {
	<#
    .SYNOPSIS
    (Helper Function) Parses a PSGetModuleInfo.xml file to extract module metadata, utilizing Parse-ModuleVersion for version interpretation.
    .DESCRIPTION
    Internal helper function for `Get-ModuleInfo`. It is designed to read and parse an XML file that typically conforms to the structure of `PSGetModuleInfo.xml`, which is often created when modules are saved using `Save-PSResource -IncludeXml` or `Save-Module -IncludeXML`.
    The function uses XPath queries to navigate the XML structure and extract key pieces of module metadata. The standard PowerShell serialization namespace ("http://schemas.microsoft.com/powershell/2004/04") is used for these queries. Properties typically extracted include:
    - Module Name (from XPath `//ps:S[@N='Name']`)
    - Module Version string (from `//ps:S[@N='Version']`)
    - Author (from `//ps:S[@N='Author']`)
    - InstalledLocation (from `//ps:S[@N='InstalledLocation']`). This usually points to the parent directory where modules are stored (e.g., "C:\Program Files\WindowsPowerShell\Modules").
    The extracted version string is then processed by the `Parse-ModuleVersion` helper. This provides a structured `[System.Version]` object for the numeric part of the version, determines if the version is a pre-release, and extracts any pre-release label (e.g., "beta1").
    The `BasePath` for the module is constructed by joining the `InstalledLocation` value (which is the modules' parent directory) with the extracted module `Name`. For example, if `InstalledLocation` is "C:\Modules" and `Name` is "MyModule", the `BasePath` becomes "C:\Modules\MyModule".
    The function returns a PSCustomObject containing these extracted and parsed details.
    .PARAMETER XmlFilePath
    [Mandatory] The string path to the `PSGetModuleInfo.xml` (or similarly structured XML) file to be parsed.
    .OUTPUTS
    PSCustomObject
    A PSCustomObject containing the extracted module information. If parsing fails or essential nodes (like Name or Version) are missing from the XML, it returns $null. The object includes:
    - ModuleName ([string]): The name of the module as extracted from the XML.
    - ModuleVersion ([System.Version]): The parsed [System.Version] object (numeric part) of the module, obtained via `Parse-ModuleVersion`.
    - ModuleVersionString ([string]): The original, unparsed version string as read directly from the XML file (from the primary 'Version' node).
    - BasePath ([string]): The constructed base path for the module (e.g., "C:\Program Files\WindowsPowerShell\Modules\MyModule"). This points to the module's root directory.
    - isPreRelease ([bool]): True if the module is identified as a pre-release, considering both version string parsing and explicit XML flags.
    - PreReleaseLabel ([string]): The pre-release label (e.g., "beta1", "rc.2") identified by `Parse-ModuleVersion`, if applicable.
    - Author ([string]): The author of the module, if specified in the XML.
    Returns $null if critical information cannot be extracted or parsed.
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$XmlFilePath
	)
	try {
		[xml]$xmlContent = Get-Content -Path $XmlFilePath -Raw -ErrorAction Stop -Verbose:$false
		$nsManager = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
		$nsManager.AddNamespace("ps", "http://schemas.microsoft.com/powershell/2004/04")
		# --- Extract basic values ---
		$nameNode = $xmlContent.SelectSingleNode("//ps:S[@N='Name']", $nsManager)
		$nameValue = if ($nameNode) { $nameNode.'#text' } else { $null }
		$authorNode = $xmlContent.SelectSingleNode("//ps:S[@N='Author']", $nsManager)
		$authorValue = if ($authorNode) { $authorNode.'#text' } else { $null }
		$versionNode = $xmlContent.SelectSingleNode("//ps:S[@N='Version']", $nsManager)
		$versionValue = if ($versionNode) { $versionNode.'#text' } else { $null } # Original main version string
		$locationNode = $xmlContent.SelectSingleNode("//ps:S[@N='InstalledLocation']", $nsManager)
		$locationValue = if ($locationNode) { $locationNode.'#text' } else { $null }
		# --- Extract values specifically for Prerelease determination ---
		# 1. NormalizedVersion string (often contains the full SemVer string)
		$normalizedVersionNode = $xmlContent.SelectSingleNode("//ps:Obj[@N='AdditionalMetadata']/MS/ps:S[@N='NormalizedVersion']", $nsManager)
		if (-not $normalizedVersionNode) {
			# Fallback for older PSGet formats or if not under AdditionalMetadata
			$normalizedVersionNode = $xmlContent.SelectSingleNode("//ps:S[@N='NormalizedVersion']", $nsManager)
		}
		$normalizedVersionValue = if ($normalizedVersionNode) { $normalizedVersionNode.'#text' } else { $null }
		# 2. Explicit IsPrerelease boolean flag (top-level)
		$prereleaseBoolNode = $xmlContent.SelectSingleNode("//ps:B[@N='IsPrerelease']", $nsManager) # <B N='IsPrerelease'>true</B>
		$prereleaseBoolText = if ($prereleaseBoolNode) { $prereleaseBoolNode.'#text' } else { $null }
		# 3. Explicit IsPrerelease string flag (often under AdditionalMetadata)
		$prereleaseStringNode = $xmlContent.SelectSingleNode("//ps:Obj[@N='AdditionalMetadata']/MS/ps:S[@N='IsPrerelease']", $nsManager) # <S N='IsPrerelease'>True</S>
		if (-not $prereleaseStringNode) {
			# Fallback for older PSGet formats or if not under AdditionalMetadata
			$prereleaseStringNode = $xmlContent.SelectSingleNode("//ps:S[@N='IsPrerelease']", $nsManager)
		}
		$prereleaseStringText = if ($prereleaseStringNode) { $prereleaseStringNode.'#text' } else { $null }
		# --- Determine Prerelease status ---
		$parsedVersionInfo = $null
		$isPrereleaseFromParse = $false
		$preReleaseLabelFromParse = $null
		$moduleVersionObject = $null
		$parsedVersionString = $null # The string that was successfully parsed by Parse-ModuleVersion
		# Try parsing NormalizedVersion first if available
		if ($normalizedVersionValue) {
			New-Log "XML: Attempting to parse NormalizedVersion '$normalizedVersionValue'" -Level VERBOSE
			$parsedVersionInfo = Parse-ModuleVersion -VersionString $normalizedVersionValue -ErrorAction SilentlyContinue
			if ($parsedVersionInfo) {
				$parsedVersionString = $normalizedVersionValue
				New-Log "XML: Parsed NormalizedVersion '$normalizedVersionValue'. IsPre: $($parsedVersionInfo.IsPrerelease), Label: '$($parsedVersionInfo.PreReleaseLabel)'" -Level VERBOSE
			}
			else {
				New-Log "XML: Could not parse NormalizedVersion string '$normalizedVersionValue' using Parse-ModuleVersion." -Level DEBUG
			}
		}
		# If NormalizedVersion wasn't parsed or doesn't exist, try parsing the main Version string
		if (-not $parsedVersionInfo -and $versionValue) {
			New-Log "XML: Attempting to parse Version '$versionValue' (NormalizedVersion failed or N/A)" -Level VERBOSE
			$parsedVersionInfo = Parse-ModuleVersion -VersionString $versionValue -ErrorAction SilentlyContinue
			if ($parsedVersionInfo) {
				$parsedVersionString = $versionValue
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
			# If no version string could be parsed by Parse-ModuleVersion, try to create a System.Version from $versionValue
			# This is a last resort for the ModuleVersion object, prerelease will rely on flags.
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
		# Determine if any XML flag explicitly says it's a prerelease
		$isPrereleaseFromXmlFlag = $false
		if (($prereleaseBoolText -is [string] -and $prereleaseBoolText.ToLowerInvariant() -eq 'true') -or
			($prereleaseStringText -is [string] -and $prereleaseStringText.ToLowerInvariant() -eq 'true')) {
			$isPrereleaseFromXmlFlag = $true
			New-Log "XML: An explicit IsPrerelease flag in XML is true (Bool: '$prereleaseBoolText', String: '$prereleaseStringText')." -Level VERBOSE
		}
		# Final Prerelease Status:
		# If Parse-ModuleVersion says it's prerelease, it is.
		# Or, if an XML flag says it's prerelease, it is.
		$finalIsPrerelease = $isPrereleaseFromParse -or $isPrereleaseFromXmlFlag
		$finalPreReleaseLabel = $preReleaseLabelFromParse # Label comes *only* from parsing the version string
		# If an XML flag forced IsPrerelease to true, but parsing didn't yield a label, the label remains null.
		# This is acceptable, as the version string itself might not have had a SemVer label (e.g., "1.0.0" with IsPrerelease=true).
		if ($nameValue -and $versionValue) {
			# Check if essential name and original version string exist
			$basePathValue = $null
			if ($locationValue -and $nameValue) {
				try {
					$basePathValue = Join-Path -Path $locationValue -ChildPath $nameValue -ErrorAction Stop -Verbose:$false
				}
				catch {
					New-Log "XML: Error constructing BasePath from Location '$locationValue' and Name '$nameValue'. Error: $($_.Exception.Message)" -Level ERROR
					# $basePathValue remains $null
				}
			}
			else {
				New-Log "XML: Cannot construct BasePath - InstalledLocation or Name missing. Location: '$locationValue', Name: '$nameValue'" -Level VERBOSE
			}
			$result = [PSCustomObject]@{
				ModuleName          = $nameValue
				ModuleVersion       = $moduleVersionObject # This can be $null if all parsing/conversion failed
				ModuleVersionString = $versionValue # Always the original string from <S N='Version'>
				BasePath            = if ($basePathValue) { "$basePathValue" } else { $null }
				isPreRelease        = $finalIsPrerelease
				PreReleaseLabel     = if ($finalIsPrerelease -and $finalPreReleaseLabel) { $finalPreReleaseLabel } else { $null } # Only show label if it's a prerelease
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
#endregion Get-ModuleInfoFromXml (Helper)
#region Parse-ModuleVersion (Helper)
function Parse-ModuleVersion {
	<#
	.SYNOPSIS
	(Helper Function) Parses a module version string to extract its base numeric version, pre-release status, and pre-release label.
	.DESCRIPTION
	This internal helper function is used by various other functions in this script to consistently interpret module version strings.
	It takes a string that might represent a standard version (e.g., '1.2.3'), a 4-part version (e.g., '1.2.3.4'), or a Semantic Versioning (SemVer) 2.0.0 style pre-release (e.g., '1.2.3-beta1', '2.0.0-rc.2', '3.0.1-alpha.10.x.Z').
	The function uses a regular expression to identify if the string matches a SemVer pre-release pattern (numeric base version followed by a hyphen and a pre-release identifier).
	- If it matches, it separates the base numeric part (e.g., "1.2.3") from the pre-release label (e.g., "beta1").
	- If it doesn't match the SemVer pre-release pattern, the entire input string is treated as the base version string.
	The identified base numeric part is then parsed into a [System.Version] object. This object can handle 2 to 4 numeric components (Major.Minor.Build.Revision).
	The function returns a PSCustomObject containing the parsed components.
	.PARAMETER VersionString
	[Mandatory] The version string to parse (e.g., "1.2.3", "1.2.3.4", "2.0.0-beta.1", "3.1-alpha-rev2").
	.OUTPUTS
	PSCustomObject
	A PSCustomObject containing the parsed version components. Returns $null if the base numeric part of the `VersionString` cannot be successfully parsed into a [System.Version] object. The output object includes:
	- ModuleVersion ([System.Version]): The parsed base [System.Version] object (e.g., for "1.2.3-beta1", this would be the [version]"1.2.3"). This will be $null if the base string is unparsable.
	- ModuleVersionString ([string]): The original, unmodified input `VersionString`.
	- ModuleVersionStringNoPrefix ([string]): The base numeric version part extracted from the `VersionString` (e.g., "1.2.3" from "1.2.3-beta1"). This is the string that was attempted to be parsed into `ModuleVersion`.
	- IsSemVer ([bool]): True if the input `VersionString` matched the N.N.N[-N.N.N]-tag pattern, indicating a SemVer-like structure with a pre-release tag.
	- IsPrerelease ([bool]): True if `IsSemVer` is true (i.e., a pre-release tag was identified according to the SemVer pattern).
	- PreReleaseLabel ([string]): The extracted pre-release tag (e.g., "beta1", "rc.2", "alpha.10.x.Z"), if `IsSemVer` is true; otherwise $null.
	- PreReleaseVersion ([string]): The full pre-release version string, reconstructed from the parsed `ModuleVersion` and `PreReleaseLabel` (e.g., "1.2.3-beta1"), if `IsPrerelease` is true; otherwise $null.
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
	# Regex: Match base version (allowing 2 to 4 parts: N.N to N.N.N.N) followed by a hyphen and the prerelease tag.
	# Base Version: \d+ (\.\d+){1,3} -> One or more digits, followed by 1 to 3 groups of (dot + one or more digits).
	# Prerelease Tag (Option 1): [0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)* -> One or more valid SemVer identifiers, dot-separated.
	# Anchors: ^...$ -> Match the entire string.
	$semVerRegex = '^(\d+(\.\d+){1,3})-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)$'
	if ($VersionString -match $semVerRegex) {
		$baseVersionString = $Matches[1]
		$prereleasePart = $Matches[3] # Group 3 captures the pre-release part
		$isSemVerStyle = $true
		$isPrerelease = $true # By definition of the pattern
		New-Log "Parse-ModuleVersion: Matched SemVer pattern. Base:'$baseVersionString', Label:'$prereleasePart'" -Level VERBOSE
	}
	else {
		# Not SemVer style, keep original string as base for TryParse
		$baseVersionString = $VersionString
	}
	# Attempt to parse the determined base version string into a [version] object
	$isParseable = [version]::TryParse($baseVersionString, [ref]$version)
	if ($isParseable) {
		return [PSCustomObject]@{
			ModuleVersionStringNoPrefix = $baseVersionString
			ModuleVersionString         = $VersionString
			ModuleVersion               = $version
			IsSemVer                    = $isSemVerStyle
			IsPrerelease                = $isPrerelease      # True only if matched SemVer pattern
			PreReleaseLabel             = if (![string]::IsNullOrEmpty($prereleasePart)) { $prereleasePart } else { $null }
			PreReleaseVersion           = if (![string]::IsNullOrEmpty($prereleasePart)) { "$($version)-$($prereleasePart)" } else { $null }
		}
	}
	else {
		# Base version string itself was not parseable by [version]::TryParse
		New-Log "Parse-ModuleVersion: Could not parse base version string '$baseVersionString' (derived from original '$VersionString') into a [System.Version] object." -Level WARNING
		return $null
	}
}
#endregion Parse-ModuleVersion (Helper)
#region Compare-ModuleVersion (Helper)
function Compare-ModuleVersion {
	<#
    .SYNOPSIS
    Compares two module version strings, with specific handling for common pre-release identifiers, to determine which is newer.
    .DESCRIPTION
    Compares two version strings (`VersionA` and `VersionB`) to determine which is "newer". This is crucial for deciding if an online version is an update to a local one. The comparison logic is as follows:
    1.  **Base Version Comparison:** The numeric parts of the versions (e.g., "3.8.0" from "3.8.0-beta1") are parsed into [System.Version] objects and compared first. A higher base version is always considered newer (e.g., "3.8.1" is newer than "3.8.0").
    2.  **Identical Base Versions - Pre-release Handling:** If the base numeric versions are identical:
        a.  **Pre-release vs. Stable:** If one version has a pre-release tag (e.g., "-beta1") and the other does not (is stable), the version WITH the pre-release tag is considered NEWER. This logic prioritizes testing newer, potentially unstable, pre-releases over an existing stable version if their base numbers are the same (e.g., "3.8.0-beta1" is considered newer than "3.8.0").
        b.  **Both Stable or Both Pre-release with Same Tag Type & Number:** If neither has a pre-release tag and base versions are identical, they are considered equal. If both have pre-release tags that are identical (same type and same numeric suffix, or no suffix), they are also considered equal.
        c.  **Both Pre-release - Tag Type Priority:** If both versions have pre-release tags but the tag types are different, they are compared based on a defined priority order: 'dev' (highest priority) > 'alpha' > 'beta' > 'preview' > 'rc' (lowest standard pre-release tag in this list). For example, "1.0.0-dev1" is newer than "1.0.0-alpha1"; "1.0.0-alpha1" is newer than "1.0.0-beta1".
        d.  **Both Pre-release - Same Tag Type, Different Number:** If both pre-release tags are of the same type (e.g., both are 'beta'), any numeric suffix is compared. A higher number indicates a newer version (e.g., "beta2" is newer than "beta1"). If one has a number and the other doesn't for the same tag type, the one with the number is typically considered newer (a tag with no number is treated as if it has a suffix of 0).
        e.  **Non-standard Pre-release Tags:** If pre-release tags do not match the recognized standard types ('dev', 'alpha', 'beta', 'preview', 'rc'), a simple string comparison of the full pre-release tags is performed.
    .PARAMETER VersionA
    [Mandatory] The first version string (e.g., "3.8.0", "3.8.0-beta1"). This is typically considered the "current" or "local" version in an update check.
    .PARAMETER VersionB
    [Mandatory] The second version string (e.g., "3.8.0", "3.8.0-preview2"). This is typically considered the "new" or "online" version in an update check.
    .PARAMETER ReturnBoolean
    [Optional] If this switch is specified, the function returns `$true` if `VersionB` is considered newer than `VersionA` according to the logic above, and `$false` otherwise (including if they are considered equal or if `VersionA` is newer).
    If this switch is not specified (the default behavior), the function returns the string of the version that is considered newer. If they are considered equal by the comparison logic, `VersionA` (the first one provided) is returned.
    .OUTPUTS
    System.String or System.Boolean
    - If `-ReturnBoolean` is NOT specified: Returns the string of the version considered newer (`VersionA` or `VersionB`). If they are deemed equal, `VersionA` is returned.
    - If `-ReturnBoolean` IS specified: Returns `$true` if `VersionB` is newer than `VersionA`; otherwise `$false`.
    .EXAMPLE
    PS C:\> Compare-ModuleVersion -VersionA "3.8.0" -VersionB "3.8.0-preview2"
    Returns: "3.8.0-preview2" (Reason: Base versions are same; VersionB is a pre-release, VersionA is stable, so pre-release is considered newer).
    .EXAMPLE
    PS C:\> Compare-ModuleVersion -VersionA "3.8.0-beta23" -VersionB "3.8.0-alpha5" -ReturnBoolean
    Returns: $false (Reason: Base versions are same. Both are pre-releases. 'alpha' has higher priority than 'beta' in this function's logic [alpha=4, beta=3]. So, VersionB ('alpha5') is NOT newer than VersionA ('beta23') because beta is lower priority than alpha. Wait, this example text is confusing.
    Corrected Logic: VersionA ('beta', Prio 3), VersionB ('alpha', Prio 4). Is B newer? Yes, Prio B (4) > Prio A (3). So should return $true. The example description text needs update.)
    PS C:\> Compare-ModuleVersion -VersionA "1.0.0-beta3" -VersionB "1.0.0-alpha10" -ReturnBoolean
    Returns: $true (Reason: 'alpha' [priority 4] is considered newer than 'beta' [priority 3])
    .EXAMPLE
    PS C:\> Compare-ModuleVersion -VersionA "1.0.0-rc1" -VersionB "1.0.0-dev5"
    Returns: "1.0.0-dev5" (Reason: 'dev' [priority 5] is considered newer than 'rc' [priority 1])
    .EXAMPLE
    PS C:\> Compare-ModuleVersion -VersionA "2.0.0" -VersionB "1.9.0-dev"
    Returns: "2.0.0" (Reason: Base version "2.0.0" is newer than "1.9.0")
    #>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$VersionA,
		[Parameter(Mandatory)][string]$VersionB,
		[Parameter()][switch]$ReturnBoolean
	)
	# Basic input validation
	if ([string]::IsNullOrWhiteSpace($VersionA) -or [string]::IsNullOrWhiteSpace($VersionB)) {
		New-Log "Compare-ModuleVersion: One or both versions are null or empty." -Level VERBOSE
		if ($ReturnBoolean) { return $false } else { return $VersionA }
	}
	# Precedence order for prerelease types (HIGHER number = higher priority)
	$typePriority = @{
		'dev'     = 5
		'alpha'   = 4
		'beta'    = 3
		'preview' = 2
		'rc'      = 1 # Release Candidate
	}
	# Split versions into base and prerelease parts
	$baseVersionA, $prereleaseA = $VersionA -split '-', 2
	$baseVersionB, $prereleaseB = $VersionB -split '-', 2
	# Compare base versions first (e.g., 3.8.0 vs 3.7.0)
	try {
		$versionObjectA = [System.Version]::new($baseVersionA)
		$versionObjectB = [System.Version]::new($baseVersionB)
		$baseComparison = $versionObjectA.CompareTo($versionObjectB)
		if ($baseComparison -ne 0) {
			# Base versions are different
			New-Log "Compare-ModuleVersion: Base versions differ - A=$baseVersionA, B=$baseVersionB. Result=$baseComparison" -Level VERBOSE
			if ($ReturnBoolean) {
				return $baseComparison -lt 0 # Return true if B > A
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
	# Base versions are the same, now handle prerelease logic
	# IMPORTANT: If one is prerelease and the other isn't, prerelease is ALWAYS newer
	if ($prereleaseA -and -not $prereleaseB) {
		# A has prerelease, B doesn't, so A is newer
		New-Log "Compare-ModuleVersion: VersionA has prerelease ($prereleaseA), VersionB doesn't. A is newer." -Level VERBOSE
		if ($ReturnBoolean) { return $false } else { return $VersionA }
	}
	elseif (-not $prereleaseA -and $prereleaseB) {
		# B has prerelease, A doesn't, so B is newer
		New-Log "Compare-ModuleVersion: VersionB has prerelease ($prereleaseB), VersionA doesn't. B is newer." -Level VERBOSE
		if ($ReturnBoolean) { return $true } else { return $VersionB }
	}
	elseif (-not $prereleaseA -and -not $prereleaseB) {
		# Neither has prerelease, they're equal
		New-Log "Compare-ModuleVersion: Neither version has prerelease. Versions are equal." -Level VERBOSE
		if ($ReturnBoolean) { return $false } else { return $VersionA }
	}
	# Both have prerelease parts, so we need to compare them
	# Normalize prerelease strings
	$prereleaseA = $prereleaseA.ToLower().Trim('.- ')
	$prereleaseB = $prereleaseB.ToLower().Trim('.- ')
	# Extract prerelease type and number using regex
	$regex = "^($(($typePriority.Keys | ForEach-Object {[regex]::Escape($_)}) -join '|'))(?:[\.\-_]?(\d+))?$"
	$matchA = [regex]::Match($prereleaseA, $regex)
	$matchB = [regex]::Match($prereleaseB, $regex)
	# If not matching standard format, do a simple string comparison
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
	# Extract type and number
	$typeA = $matchA.Groups[1].Value
	$numberA = if ([string]::IsNullOrEmpty($matchA.Groups[2].Value)) { 0 } else { [int]$matchA.Groups[2].Value }
	$typeB = $matchB.Groups[1].Value
	$numberB = if ([string]::IsNullOrEmpty($matchB.Groups[2].Value)) { 0 } else { [int]$matchB.Groups[2].Value }
	New-Log "Compare-ModuleVersion: Comparing prerelease A='$typeA'($numberA) vs B='$typeB'($numberB)" -Level VERBOSE
	# Compare prerelease types
	if ($typeA -ne $typeB) {
		$priorityA = $typePriority[$typeA]
		$priorityB = $typePriority[$typeB]
		New-Log "Prerelease types differ. Priority A=$priorityA, B=$priorityB." -Level VERBOSE
		# HIGHER number means higher priority
		if ($ReturnBoolean) {
			return $priorityB -gt $priorityA
		}
		else {
			return $(if ($priorityA -gt $priorityB) { $VersionA } else { $VersionB })
		}
	}
	# Types are the same, compare numbers
	New-Log "Prerelease types are the same ('$typeA'). Comparing numbers: A=$numberA, B=$numberB." -Level VERBOSE
	if ($ReturnBoolean) {
		return $numberB -gt $numberA
	}
	else {
		return $(if ($numberA -ge $numberB) { $VersionA } else { $VersionB })
	}
}
#endregion Compare-ModuleVersio (Helper)
### OBS: New-Log Function is needed otherwise remove all New-Log and replace with Write-Host. New-Log is vastly better though, check the link below:
#Example:
#Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
Check-PSResourceRepository -ImportDependencies
#Import-Module -Name ThreadJob -Force
$ignoredModules = @('Example2.Diagnostics', 'BurntToast') #Fully ignored modules
$blackList = @{ #Ignored module and repo combo.
	'Microsoft.Graph.Beta' = @("Nuget", "NugetGallery")
	'Microsoft.Graph'      = @("Nuget", "NugetGallery")
}
$paths = $env:PSModulePath.Split(';') | Where-Object { $_ -notmatch '.vscode' -and $_ -notmatch 'System32' }
$moduleInfo = Get-ModuleInfo -Paths $paths -IgnoredModules $ignoredModules
$outdated = Get-ModuleUpdateStatus -ModuleInventory $moduleInfo -TimeoutSeconds 120 -Repositories @("PSGallery", "Nuget", "NugetGallery") -MatchAuthor -BlackList $blackList
if ($outdated) {
	$res = $outdated | Update-Modules -Clean -PreRelease -UseProgressBar
	$res
}
else {
	New-Log "No outdata to run" -Level SUCCESS
}
#>