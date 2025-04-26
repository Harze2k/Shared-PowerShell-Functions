#Requires -Version 7.4
<#
Author: Harze2k
Date:   2025-04-26 (First public release)
Version: 2.2

Sample output:
[2025-04-26 04:12:47.289][DEBUG] Processing potential module 'WindowsErrorReporting' at 'C:\Windows\System32\WindowsPowerShell\v1.0\Modules\WindowsErrorReporting'...
[2025-04-26 04:12:47.320][DEBUG] Processing potential module 'WindowsSearch' at 'C:\Windows\System32\WindowsPowerShell\v1.0\Modules\WindowsSearch'...
[2025-04-26 04:12:47.328][DEBUG] Processing potential module 'WindowsUpdate' at 'C:\Windows\System32\WindowsPowerShell\v1.0\Modules\WindowsUpdate'...
[2025-04-26 04:12:47.335][DEBUG] Processing potential module 'WinHttpProxy' at 'C:\Windows\System32\WindowsPowerShell\v1.0\Modules\WinHttpProxy'...

[2025-04-26 04:14:20.432][INFO] Starting parallel module check for 173 modules (ThrottleLimit: 64)...
[2025-04-26 04:14:20.446][DEBUG] Found 173 modules with valid version information

[2025-04-26 04:14:20.469][WARNING] Found update for ActiveDirectory but that module and repository combo: NuGetGallery is blacklisted.
[2025-04-26 04:14:20.859][WARNING] Found update for iSCSI but that module and repository combo: NuGetGallery is blacklisted.
[2025-04-26 04:14:20.908][WARNING] Found update for Microsoft.Graph but that module and repository combo: NuGetGallery is blacklisted.
[2025-04-26 04:14:20.952][WARNING] Found update for Microsoft.Graph.Beta but that module and repository combo: NuGetGallery is blacklisted.

[2025-04-26 04:14:21.254][SUCCESS] [3%] Processed 5 of 173 modules, ETA: Calculating..., Updates: 0, Errors: 0
[2025-04-26 04:14:21.258][SUCCESS] [6%] Processed 10 of 173 modules, ETA: 00:13, Updates: 0, Errors: 0
[2025-04-26 04:14:21.261][SUCCESS] [9%] Processed 15 of 173 modules, ETA: 00:08, Updates: 0, Errors: 0
[2025-04-26 04:14:21.423][SUCCESS] [12%] Processed 20 of 173 modules, ETA: 00:07, Updates: 0, Errors: 0
[2025-04-26 04:14:21.448][SUCCESS] [14%] Processed 25 of 173 modules, ETA: 00:05, Updates: 0, Errors: 0
[2025-04-26 04:14:21.471][SUCCESS] [17%] Processed 30 of 173 modules, ETA: 00:04, Updates: 0, Errors: 0
[2025-04-26 04:14:21.543][SUCCESS] [20%] Processed 35 of 173 modules, ETA: 00:04, Updates: 0, Errors: 0
[2025-04-26 04:14:21.596][SUCCESS] [23%] Processed 40 of 173 modules, ETA: 00:03, Updates: 0, Errors: 0
[2025-04-26 04:14:21.647][SUCCESS] [26%] Processed 45 of 173 modules, ETA: 00:03, Updates: 0, Errors: 0
[2025-04-26 04:14:21.757][SUCCESS] [29%] Processed 50 of 173 modules, ETA: 00:03, Updates: 0, Errors: 0
[2025-04-26 04:14:21.781][SUCCESS] [32%] Processed 55 of 173 modules, ETA: 00:02, Updates: 0, Errors: 0
[2025-04-26 04:14:21.853][SUCCESS] [35%] Processed 60 of 173 modules, ETA: 00:02, Updates: 0, Errors: 0
[2025-04-26 04:14:21.898][SUCCESS] [38%] Processed 65 of 173 modules, ETA: 00:02, Updates: 0, Errors: 0
[2025-04-26 04:14:21.898][SUCCESS] [38%] Processed 65 of 173 modules, ETA: 00:02, Updates: 0, Errors: 0
[2025-04-26 04:14:21.926][SUCCESS] [40%] Processed 70 of 173 modules, ETA: 00:02, Updates: 0, Errors: 0
[2025-04-26 04:14:22.015][SUCCESS] [43%] Processed 75 of 173 modules, ETA: 00:02, Updates: 0, Errors: 0

[2025-04-26 04:14:22.028][WARNING] Found update for ScheduledTasks but that module and repository combo: NuGetGallery is blacklisted.

[2025-04-26 04:14:22.035][SUCCESS] [46%] Processed 80 of 173 modules, ETA: 00:01, Updates: 0, Errors: 0
[2025-04-26 04:14:22.096][SUCCESS] [49%] Processed 85 of 173 modules, ETA: 00:01, Updates: 0, Errors: 0
[2025-04-26 04:14:22.174][SUCCESS] [52%] Processed 90 of 173 modules, ETA: 00:01, Updates: 0, Errors: 0
[2025-04-26 04:14:22.294][SUCCESS] [55%] Processed 95 of 173 modules, ETA: 00:01, Updates: 0, Errors: 0
[2025-04-26 04:14:22.364][SUCCESS] [58%] Processed 100 of 173 modules, ETA: 00:01, Updates: 0, Errors: 0
[2025-04-26 04:14:22.396][SUCCESS] [61%] Processed 105 of 173 modules, ETA: 00:01, Updates: 0, Errors: 0
[2025-04-26 04:14:22.474][SUCCESS] [64%] Processed 110 of 173 modules, ETA: 00:01, Updates: 0, Errors: 0
[2025-04-26 04:14:22.488][SUCCESS] [66%] Processed 115 of 173 modules, ETA: 00:01, Updates: 0, Errors: 0
[2025-04-26 04:14:22.528][SUCCESS] [69%] Processed 120 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:22.589][SUCCESS] [72%] Processed 125 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:22.629][SUCCESS] [75%] Processed 130 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:22.644][SUCCESS] [78%] Processed 135 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:22.774][SUCCESS] [81%] Processed 140 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:22.808][SUCCESS] [84%] Processed 145 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:22.904][SUCCESS] [87%] Processed 150 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:22.986][SUCCESS] [90%] Processed 155 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:23.113][SUCCESS] [92%] Processed 160 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:23.444][SUCCESS] [95%] Processed 165 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0
[2025-04-26 04:14:23.733][SUCCESS] [98%] Processed 170 of 173 modules, ETA: 00:00, Updates: 0, Errors: 0

[2025-04-26 04:14:24.915][DEBUG] Completed check of 173 modules in 4,5 seconds. Found 0 modules needing updates.
[2025-04-26 04:14:24.915][DEBUG] Encountered 0 errors during processing.

[2025-04-26 04:10:52.211][DEBUG] [1/2] Updating [Microsoft.PowerShell.Security] to version [7.6.0] (Preview=True).
[2025-04-26 04:10:52.824][SUCCESS] Found module [Microsoft.PowerShell.Security] version [7.6.0] in repository [NuGetGallery]
[2025-04-26 04:10:53.378][WARNING] Failed to update [Microsoft.PowerShell.Security] to [7.6.0] via Save-PSResource to C:\Program Files\PowerShell\7\Modules\Microsoft.PowerShell.Security
[2025-04-26 04:10:53.935][WARNING] Failed to update [Microsoft.PowerShell.Security] to [7.6.0] via Save-PSResource to C:\Windows\System32\WindowsPowerShell\v1.0\Modules\Microsoft.PowerShell.Security
[2025-04-26 04:10:53.937][DEBUG] Will try a different way...
[2025-04-26 04:10:54.493][WARNING] Install-PSResource failed for [Microsoft.PowerShell.Security]. Trying Install-Module... with -AllowClobber
[2025-04-26 04:10:56.259][WARNING] Last chance managed to install the latest version in a path where it was not present before.
[2025-04-26 04:10:56.263][SUCCESS] Microsoft.PowerShell.Security with version 7.6.0 is now present at: C:\Program Files\PowerShell\Modules\Microsoft.PowerShell.Security

[2025-04-26 04:11:05.787][DEBUG] [2/2] Updating [HtmlAgilityPack] to version [1.12.1] (Preview=True).
[2025-04-26 04:11:06.587][SUCCESS] Found module [HtmlAgilityPack] version [1.12.1] in repository [NuGetGallery]
[2025-04-26 04:11:08.045][SUCCESS] Successfully updated [HtmlAgilityPack] to [1.12.1] via Save-PSResource to C:\Program Files\WindowsPowerShell\Modules\HtmlAgilityPack
[2025-04-26 04:11:09.578][SUCCESS] Successfully updated [HtmlAgilityPack] to [1.12.1] via Save-PSResource to C:\Windows\System32\WindowsPowerShell\v1.0\Modules\HtmlAgilityPack
[2025-04-26 04:11:09.580][SUCCESS] Successfully updated [HtmlAgilityPack] to [1.12.1] on all destinations!

ModuleName     : Microsoft.PowerShell.Security
NewVersion     : 7.6.0
UpdatedPaths   : {C:\Program Files\PowerShell\Modules\Microsoft.PowerShell.Security}
FailedPaths    : {C:\Program Files\PowerShell\7\Modules\Microsoft.PowerShell.Security, C:\Windows\System32\WindowsPowerShell\v1.0\Modules\Microsoft.PowerShell.Security}
OverallSuccess : False

ModuleName     : HtmlAgilityPack
NewVersion     : 1.12.1
UpdatedPaths   : {C:\Program Files\WindowsPowerShell\Modules\HtmlAgilityPack, C:\Windows\System32\WindowsPowerShell\v1.0\Modules\HtmlAgilityPack}
FailedPaths    : {}
OverallSuccess : True

#>
function Check-PSResourceRepository {
	[CmdletBinding()]
	param (
		[switch]$ImportDependencies,
		[switch]$ForceInstall
	)
	function Install-RequiredModulesInternal {
		param (
			[switch]$ForceReinstall
		)
		try {
			$psGallery = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
			if ($null -eq $psGallery -or -not $psGallery.Trusted) {
				New-Log "Setting PSGallery repository to Trusted." -Level INFO
				Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction Stop
				New-Log "PSGallery repository set to trusted." -Level SUCCESS
			}
			else {
				New-Log "PSGallery repository is already trusted." -Level INFO
			}
			$moduleList = @(
				@{ Name = 'Microsoft.PowerShell.PSResourceGet'; Prerelease = $true }
				@{ Name = 'PowerShellGet'; Prerelease = $true }
			)
		}
		catch {
			New-Log "Failed to set PSGallery repository to Trusted." -Level ERROR
		}
		$usePSResourceCmdlets = Get-Command Install-PSResource -ErrorAction SilentlyContinue
		$moduleSuccess = $true
		foreach ($moduleInfo in $moduleList) {
			$moduleName = $moduleInfo.Name
			$installPrerelease = $moduleInfo.Prerelease
			$foundModule = Find-Module -Name $moduleName -Repository @('PSGallery') -AllowPrerelease:$installPrerelease -ErrorAction SilentlyContinue
			$isInstalled = Get-module -Name $moduleName -All -ListAvailable | Where-Object {$_.Version -eq $foundModule.Version} | Select-Object -First 1
			if ($ForceReinstall -or !$isInstalled) {
				New-Log "Attempting to install/update module '$moduleName' ($($ForceReinstall ? 'Forced Reinstall' : ($isInstalled ? 'Update' : 'Install')))..." -Level INFO
				$commonInstallParams = @{
					Name          = $moduleName
					Scope         = 'AllUsers'
					AcceptLicense = $true
					Confirm       = $false
					PassThru = $true
					ErrorAction   = 'SilentlyContinue'
					WarningAction = 'SilentlyContinue'
				}
				if ($usePSResourceCmdlets) {
					New-Log "Using Install-PSResource for '$moduleName'." -Level INFO
					$installParams = @{
						Reinstall       = $ForceReinstall
						TrustRepository = $true
						Repository 	  =	@('PSGallery')
					} + $commonInstallParams
					if ($installPrerelease) { $installParams.Add('Prerelease', $true) }
					$res = Install-PSResource @installParams
				}
				if (!$res -or !$usePSResourceCmdlets) {
					New-Log "Trying with Install-Module for '$moduleName'." -Level INFO
					$installParams = @{
						Force = $ForceReinstall
					} + $commonInstallParams
					if ($installPrerelease) { $installParams.Add('AllowPrerelease', $true) }
					$res = Install-Module @installParams
				}
				if ($res) {
					New-Log "Successfully installed/updated module '$moduleName'." -Level SUCCESS
					$moduleSuccess = $true
				}
				elseif (!$res -and $isInstalled) {
					New-Log "Could not force an reinstall/update of '$moduleName'. Latest version [$($foundModule.Version)] is installed though."
					$moduleSuccess = $true
				}
				else {
					New-Log "Could not install/update '$moduleName'" -Level WARNING
					$moduleSuccess = $false
				}
			}
			else {
				New-Log "Module '$moduleName' is already available." -Level INFO
			}
			try {
				Import-Module -Name $moduleName -Force -ErrorAction Stop
				New-Log "Successfully imported module '$moduleName'." -Level SUCCESS
			}
			catch {
				New-Log "Failed to import-module $moduleName." -Level ERROR
				$moduleSuccess = $false
			}
		}
		return $moduleSuccess
	}
	function Register-ConfigureRepositoryInternal {
		param (
			[Parameter(Mandatory = $true)][string]$Name,
			[string]$Uri,
			[Parameter(Mandatory = $true)][int]$Priority,
			[string]$ApiVersion = 'v3',
			[switch]$IsPSGallery
		)
		try {
			$repository = Get-PSResourceRepository -Name $Name -ErrorAction SilentlyContinue
			$needsUpdate = ($null -eq $repository) -or ($repository.Priority -ne $Priority) -or (-not $repository.Trusted) -or (!$IsPSGallery -and $repository.Uri.AbsoluteUri -ne $Uri)
			if ($needsUpdate) {
				New-Log "Registering/Updating repository '$Name' (Priority: $Priority, Trusted: True)." -Level INFO
				if ($IsPSGallery) {
					Register-PSResourceRepository -PSGallery -Force -Trusted -Priority $Priority -Confirm:$false -ErrorAction Stop
				}
				else {
					$registerParams = @{
						Name        = $Name
						Uri         = $Uri
						Trusted     = $true
						Priority    = $Priority
						Force       = $true
						ErrorAction = 'Stop'
						Confirm     = $false
					}
					if ($ApiVersion -eq 'v2') {
						$registerParams.Add('ApiVersion', $ApiVersion)
						New-Log "Using API Version V2 for '$Name'." -Level INFO
					}
					Register-PSResourceRepository @registerParams
				}
				New-Log "Successfully registered/updated '$Name' repository resource." -Level SUCCESS
			}
			else {
				New-Log "'$Name' repository is already registered and configured correctly." -Level INFO
			}
			return $true
		}
		catch {
			New-Log "Failed to register/configure '$Name' repository." -Level ERROR
			return $false
		}
	}
	if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
		New-Log "Administrator privileges are required to install modules with -Scope AllUsers. Module installation will fail. Aborting." -Level WARNING
		return
	}
	$overallSuccess = $true
	try {
		$existingProtocols = [Net.ServicePointManager]::SecurityProtocol
		$tls12Enum = [Net.SecurityProtocolType]::Tls12
		if (-not ($existingProtocols -band $tls12Enum)) {
			[Net.ServicePointManager]::SecurityProtocol = $existingProtocols -bor $tls12Enum
			New-Log "TLS 1.2 security protocol enabled for this session." -Level INFO
		}
		else {
			New-Log "TLS 1.2 security protocol was already enabled." -Level INFO
		}
	}
	catch {
		New-Log "Unable to set TLS 1.2 security protocol. Network operations might fail." -Level ERROR
	}
	$psResourceCmdletAvailable = Get-Command Register-PSResourceRepository -ErrorAction SilentlyContinue
	if (!$psResourceCmdletAvailable -or $ImportDependencies.IsPresent) {
		if (!$psResourceCmdletAvailable) {
			New-Log "Required command 'Register-PSResourceRepository' not found. Attempting to install dependencies..." -Level WARNING
		}
		else {
			New-Log "Parameter -ImportDependencies specified. Checking/installing dependencies..." -Level INFO
		}
		if (-not (Install-RequiredModulesInternal -ForceReinstall:$ForceInstall.IsPresent)) {
			New-Log "Failed to install or import required modules (PowerShellGet/Microsoft.PowerShell.PSResourceGet). Cannot configure repositories." -Level WARNING
			return
		}
		if (-not (Get-Command Register-PSResourceRepository -ErrorAction SilentlyContinue)) {
			New-Log "Required command 'Register-PSResourceRepository' is still not available after installation attempt. Aborting." -Level WARNING
			return
		}
		New-Log "Required module commands are now available." -Level SUCCESS
	}
	else {
		New-Log "Required module commands (PSResourceGet) are already available." -Level INFO
	}
	$repositories = @(
		@{ Name = 'PSGallery'; Uri = ''; Priority = 30; IsPSGallery = $true }
		@{ Name = 'NuGetGallery'; Uri = 'https://api.nuget.org/v3/index.json'; Priority = 40 }
		@{ Name = 'NuGet'; Uri = 'http://www.nuget.org/api/v2'; Priority = 50; ApiVersion = 'v2' }
	)
	New-Log "Starting repository configuration..." -Level INFO
	foreach ($repo in $repositories) {
		if (-not (Register-ConfigureRepositoryInternal @repo)) {
			$overallSuccess = $false
		}
	}
	if ($overallSuccess) {
		New-Log "All specified repositories have been successfully registered and configured." -Level SUCCESS
	}
	else {
		New-Log "One or more repositories could not be configured correctly. Please check previous error messages." -Level WARNING
	}
}
function Get-ModuleInfo {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string[]]$Paths,
		[int]$WarningThreshold = 200,
		[string[]]$IgnoredModules = @(),
		[switch]$Force,
		[switch]$IncludeResourceFiles
	)
	$allFoundModules = [System.Collections.Generic.List[object]]::new()
	$processedManifests = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
	$processedModuleRoots = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
	foreach ($startPath in $Paths) {
		if (-not (Test-Path $startPath)) {
			New-Log "Input path '$startPath' does not exist. Skipping." -Level WARNING
			continue
		}
		$potentialModuleDirs = Get-ChildItem -Path $startPath -Directory -ErrorAction SilentlyContinue | Where-Object {
			$_.Name -notmatch '^\.|^__'
		}
		if (-not $potentialModuleDirs) {
			New-Log "No potential module directories found directly under '$startPath'." -Level VERBOSE
			continue
		}
		foreach ($moduleDir in $potentialModuleDirs) {
			$moduleRootPath = $moduleDir.FullName
			$moduleRootName = $moduleDir.Name
			if (-not $processedModuleRoots.Add($moduleRootPath)) {
				New-Log "Skipping already processed module root: '$moduleRootPath'" -Level VERBOSE
				continue
			}
			if ($IgnoredModules -contains $moduleRootName) {
				New-Log "Skipping ignored module root '$moduleRootName' at '$moduleRootPath'." -Level VERBOSE
				continue
			}
			New-Log "Processing potential module '$moduleRootName' at '$moduleRootPath'..." -Level DEBUG
			$psd1Files = Get-ChildItem -Path $moduleRootPath -Recurse -Filter *.psd1 -File -ErrorAction SilentlyContinue
			$xmlFiles = Get-ChildItem -Path $moduleRootPath -Recurse -Filter PSGetModuleInfo.xml -File -ErrorAction SilentlyContinue
			if (($psd1Files.Count + $xmlFiles.Count) -gt $WarningThreshold) {
				New-Log "Processing $($psd1Files.Count) manifests and $($xmlFiles.Count) XML files in '$moduleRootPath'. This might take a while..." -Level WARNING
			}
			foreach ($file in $psd1Files) {
				$filePath = $file.FullName
				if (-not $processedManifests.Add($filePath)) {
					New-Log "Skipping already processed manifest '$filePath'" -Level VERBOSE
					continue
				}
				$moduleInfo = Get-ModuleformPath -Path $file.DirectoryName
				$isResourceFile = $false
				if ($filePath -like '*Resource*' -or $filePath -like '*String*' -or $filePath -like '*Localization*' -or $filePath -like '*Msg*') {
					$isResourceFile = $true
				}
				if ($isResourceFile) {
					if (-not $IncludeResourceFiles) {
						New-Log "Skipping resource file (IncludeResourceFiles=$false): '$filePath'" -Level VERBOSE
						continue
					}
					$version = Get-ModuleVersionFromManifest -Path $filePath
					if (-not $version) {
						$version = [version]'0.0.0'
					}
					$moduleName = $moduleInfo.ModuleName
					$moduleBasePath = $moduleInfo.ModulePath
				}
				else {
					$version = Get-ModuleVersionFromManifest -Path $filePath
					if (-not $version) {
						New-Log "Could not determine version for manifest '$filePath'. Skipping." -Level WARNING
						continue
					}
					$moduleName = $moduleInfo.ModuleName
					$moduleBasePath = $moduleInfo.ModulePath
					if (-not $Force -and ($IgnoredModules -contains $moduleName)) {
						New-Log "Skipping ignored module '$moduleName' found in '$filePath'." -Level SUCCESS
						continue
					}
				}
				if ($moduleName -and $version -and $moduleBasePath) {
					$allFoundModules.Add([PSCustomObject]@{
							Name           = $moduleName
							Version        = $version
							BasePath       = $moduleBasePath
							ManifestPath   = $filePath
							IsResourceFile = $isResourceFile
							Source         = 'Manifest'
						})
				}
				else {
					New-Log "Could not reliably determine module info for '$filePath'. Skipping." -Level WARNING
				}
			}
			foreach ($xmlFile in $xmlFiles) {
				$moduleInfo = Get-ModuleformPath -Path $xmlFile.DirectoryName
				$existingManifest = $allFoundModules | Where-Object { $_.Source -eq 'Manifest' -and $_.BasePath -eq $moduleInfo.ModulePath } | Sort-Object Version -Descending | Select-Object -First 1
				$xmlInfo = Get-ModuleInfoFromXml -XmlFilePath $xmlFile.FullName
				if ($xmlInfo) {
					if ($IgnoredModules -contains $xmlInfo.Name) {
						New-Log "Skipping ignored module '$($xmlInfo.Name)' found in XML '$($xmlFile.FullName)'." -Level SUCCESS
						continue
					}
					if (-not $existingManifest -or $xmlInfo.Version -gt $existingManifest.Version) {
						$existingXmlIndex = $allFoundModules.FindIndex({ param($m) $m.Source -eq 'XML' -and $m.Name -eq $xmlInfo.Name -and $m.BasePath -eq $moduleInfo.ModulePath })
						if ($existingXmlIndex -ge 0) { $allFoundModules.RemoveAt($existingXmlIndex) }
						$allFoundModules.Add([PSCustomObject]@{
								Name           = $xmlInfo.Name
								Version        = $xmlInfo.Version
								BasePath       = $moduleInfo.ModulePath
								ManifestPath   = $xmlFile.FullName
								IsResourceFile = $false
								Source         = 'XML'
							})
						New-Log "Added module '$($xmlInfo.Name)' version '$($xmlInfo.Version)' from XML '$($xmlFile.FullName)' for base path $($moduleInfo.ModulePath)." -Level VERBOSE
					}
					else {
						New-Log "Skipping XML info for '$($xmlInfo.Name)' version '$($xmlInfo.Version)' because a manifest version '$($existingManifest.Version)' exists for base path $($moduleInfo.ModulePath)." -Level VERBOSE
					}
				}
			}
		}
	}
	$resultModules = @{}
	# Group by Module Name first, then by Base Path
	$modulesGroupedByName = $allFoundModules | Where-Object { $_.Name -ne $null } | Group-Object Name
	foreach ($nameGroup in $modulesGroupedByName) {
		$moduleName = $nameGroup.Name
		if ($IgnoredModules -contains $moduleName) { continue }
		$modulesGroupedByPath = $nameGroup.Group | Group-Object BasePath
		$moduleVersionsList = [System.Collections.Generic.List[object]]::new()
		foreach ($pathGroup in $modulesGroupedByPath) {
			# Find the highest version within this specific BasePath for this Module Name
			$highestVersionEntry = $pathGroup.Group | Sort-Object Version -Descending | Select-Object -First 1
			if ($highestVersionEntry) {
				$outputObject = [PSCustomObject]@{
					Version = $highestVersionEntry.Version.ToString()
					Path    = $highestVersionEntry.BasePath
				}
				if ($highestVersionEntry.IsResourceFile) {
					$outputObject | Add-Member -NotePropertyName Type -NotePropertyValue 'Resource'
				}
				elseif ($highestVersionEntry.Source -eq 'XML') {
					$outputObject | Add-Member -NotePropertyName Source -NotePropertyValue 'XML'
				}
				$moduleVersionsList.Add($outputObject)
			}
		}
		if ($moduleVersionsList.Count -gt 0) {
			$resultModules[$moduleName] = $moduleVersionsList | Sort-Object Path
		}
	}
	$sortedModules = [ordered]@{}
	foreach ($key in ($resultModules.Keys | Sort-Object)) {
		# Filter out keys that look like version numbers (can happen if folder structure is odd)
		if ($key -notmatch '^\d+(\.\d+)+$') {
			$sortedModules[$key] = $resultModules[$key]
		}
		else {
			New-Log "Some issue with the module name: $key" -Level WARNING
		}
	}
	Return $sortedModules
}
function Get-ModuleVersionFromManifest {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]
		[string]$Path
	)
	try {
		$manifestMetadata = Test-ModuleManifest -Path $Path -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
		if ($manifestMetadata -and $manifestMetadata.Version) {
			return $manifestMetadata.Version
		}
		$manifestData = Import-PowerShellDataFile -Path $Path -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
		if ($manifestData) {
			if ($manifestData.PSObject.Properties.Name -contains 'ModuleVersion') {
				return [version]$manifestData.ModuleVersion
			}
			if ($manifestData.PSObject.Properties.Name -contains 'Version') {
				return [version]$manifestData.Version
			}
		}
	}
	catch { }
	return $null
}
function Get-ModuleformPath {
	param (
		[Parameter(Mandatory = $true)][string]$Path
	)
	if ($Path -match '\\Modules\\([^\\]+)(?:\\.*)?$') {
		$moduleName = $Matches[1]
		$modulePath = $Path -replace '(\\Modules\\[^\\]+).*$', '$1'
		return [PSCustomObject]@{
			ModuleName = $moduleName
			ModulePath = $modulePath
		}
	}
	$currentPath = $Path
	$leaf = Split-Path -Path $currentPath -Leaf
	while (($leaf -match '^[a-z]{2,3}(-[A-Z]{2,4})?$' -or $leaf -match '^\d+(\.\d+)+$') -and $leaf -ne 'Modules') {
		$currentPath = Split-Path -Path $currentPath -Parent
		$leaf = Split-Path -Path $currentPath -Leaf
	}
	if ($leaf -eq 'Modules') {
		$parentPath = Split-Path -Path $currentPath -Parent
		$directoryAfterModules = Split-Path -Path $Path -Leaf
		return [PSCustomObject]@{
			ModuleName = $directoryAfterModules
			ModulePath = "$currentPath\$directoryAfterModules"
		}
	}
	return [PSCustomObject]@{
		ModuleName = $leaf
		ModulePath = $currentPath
	}
}
function Get-ModuleInfoFromXml {
	param (
		[Parameter(Mandatory = $true)][string]$XmlFilePath
	)
	if (-not (Test-Path -Path $XmlFilePath)) {
		Write-Error "XML file not found: $XmlFilePath"
		return $null
	}
	[xml]$xmlContent = Get-Content -Path $XmlFilePath -Raw
	$nsManager = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
	$nsManager.AddNamespace("ps", "http://schemas.microsoft.com/powershell/2004/04")
	$nameNode = $xmlContent.SelectSingleNode("//ps:S[@N='Name']", $nsManager)
	$versionNode = $xmlContent.SelectSingleNode("//ps:S[@N='Version']", $nsManager)
	if ($nameNode -eq $null -or $versionNode -eq $null) {
		$nameNode = $xmlContent.SelectSingleNode("//S[@N='Name']")
		$versionNode = $xmlContent.SelectSingleNode("//S[@N='Version']")
	}
	$moduleName = if ($nameNode) { $nameNode.'#text' } else { $null }
	$moduleVersion = if ($versionNode) { $versionNode.'#text' } else { $null }
	$result = [PSCustomObject]@{
		Name    = $moduleName
		Version = $moduleVersion
	}
	return $result
}
function Get-ModuleUpdateStatus {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][hashtable]$ModuleInventory,
		[string[]]$Repositories = @('PSGallery', 'NuGet'),
		[int]$ThrottleLimit = [Environment]::ProcessorCount * 2,
		[ValidateRange(1, 300)][int]$TimeoutSeconds = 180,
		[hashtable]$BlackList = @{}
	)
	if ($PSVersionTable.PSVersion.Major -lt 7) {
		throw "This function requires PowerShell 7 or later. Current version: $($PSVersionTable.PSVersion)"
	}
	function Parse-ModuleVersion {
		param (
			[string]$VersionString
		)
		$version = $null
		$isParseable = [version]::TryParse($VersionString, [ref]$version)
		if ($isParseable) {
			return [PSCustomObject]@{
				OriginalString  = $VersionString
				Version         = $version
				IsSemVer        = $false
				IsPrerelease    = $false
				PreReleaseLabel = $null
			}
		}
		if ($VersionString -match '^(\d+(?:\.\d+){2,3})-(.+)$') {
			$versionPart = $Matches[1]
			$prereleasePart = $Matches[2]
			if ([version]::TryParse($versionPart, [ref]$version)) {
				return [PSCustomObject]@{
					OriginalString  = $VersionString
					Version         = $version
					IsSemVer        = $true
					IsPrerelease    = $true
					PreReleaseLabel = $prereleasePart
				}
			}
		}
		if ($VersionString -match '^(\d+(?:\.\d+){2,3})') {
			$versionPart = $Matches[0]
			if ([version]::TryParse($versionPart, [ref]$version)) {
				return [PSCustomObject]@{
					OriginalString  = $VersionString
					Version         = $version
					IsSemVer        = $false
					IsPrerelease    = $false
					PreReleaseLabel = $null
				}
			}
		}
		return $null
	}
	function Compare-PrereleaseVersion {
		[CmdletBinding()]
		param (
			[Parameter(Mandatory)][string]$VersionA,
			[Parameter(Mandatory)][string]$VersionB,
			[Parameter()][switch]$ReturnBoolean
		)
		$typePriority = @{
			'alpha'   = 1
			'beta'    = 2
			'preview' = 3
		}
		$VersionA = $VersionA -replace '.*-', ''
		$VersionB = $VersionB -replace '.*-', ''
		$VersionA = $VersionA.ToLower()
		$VersionB = $VersionB.ToLower()
		$regex = "^(alpha|beta|preview)(\d*)$"
		$matchA = [regex]::Match($VersionA, $regex)
		$matchB = [regex]::Match($VersionB, $regex)
		if (-not $matchA.Success -or -not $matchB.Success) {
			Write-Verbose "Invalid prerelease format. Expected format like 'alpha1', 'beta2', or 'preview3'"
			if ($ReturnBoolean) {
				return $false
			}
			else {
				if ($matchA.Success) { return $VersionA }
				elseif ($matchB.Success) { return $VersionB }
				else { return $VersionA }
			}
		}
		$typeA = $matchA.Groups[1].Value
		$numberA = if ($matchA.Groups[2].Value -eq '') { 0 } else { [int]$matchA.Groups[2].Value }
		$typeB = $matchB.Groups[1].Value
		$numberB = if ($matchB.Groups[2].Value -eq '') { 0 } else { [int]$matchB.Groups[2].Value }
		if ($typeA -ne $typeB) {
			if ($ReturnBoolean) {
				return $typePriority[$typeB] -gt $typePriority[$typeA]
			}
			else {
				return $(if ($typePriority[$typeA] -gt $typePriority[$typeB]) { $VersionA } else { $VersionB })
			}
		}
		if ($ReturnBoolean) {
			return $numberB -gt $numberA
		}
		else {
			return $(if ($numberA -ge $numberB) { $VersionA } else { $VersionB })
		}
	}
	$allModuleNames = $ModuleInventory.Keys | Sort-Object -Unique | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
	$totalModules = $allModuleNames.Count
	$results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
	New-Log "Starting parallel module check for $totalModules modules (ThrottleLimit: $ThrottleLimit)..."
	$moduleDataArray = @()
	foreach ($moduleName in $allModuleNames) {
		$localModules = $ModuleInventory[$moduleName]
		if ($localModules -isnot [array]) {
			$localModules = @($localModules)
		}
		$parsedVersions = @()
		foreach ($module in $localModules) {
			$versionStr = $module.Version
			$parsedVersion = Parse-ModuleVersion -VersionString $versionStr
			if ($parsedVersion) {
				$parsedVersions += [PSCustomObject]@{
					ParsedVersion   = $parsedVersion
					OriginalVersion = $versionStr
					Path            = $module.Path
				}
			}
			else {
				New-Log "Failed to parse version: $versionStr in module: $moduleName" -Level WARNING
			}
		}
		if ($parsedVersions.Count -gt 0) {
			if ($parsedVersions.Count -eq 1) {
				$highestVersionInfo = $parsedVersions[0]
			}
			else {
				$highestVersionInfo = $parsedVersions | Sort-Object -Property { $_.ParsedVersion.Version } -Descending | Select-Object -First 1
			}
			$moduleDataArray += [PSCustomObject]@{
				Name                    = $moduleName
				HighestLocalVersionInfo = $highestVersionInfo
				AllVersions             = $parsedVersions
				PreReleaseLabel         = if ($highestVersionInfo.ParsedVersion.PreReleaseLabel) { $highestVersionInfo.ParsedVersion.PreReleaseLabel } else { $null }
			}
		}
	}
	$validModuleCount = $moduleDataArray.Count
	New-Log "Found $validModuleCount modules with valid version information" -Level DEBUG
	# Create synchronized hashtable for counters
	$counters = [hashtable]::Synchronized(@{
			Processed = 0
			Updates   = 0
			Errors    = 0
			StartTime = Get-Date
			Total     = $validModuleCount
		})
	$ComparePrereleaseVersion = ${function:Compare-PrereleaseVersion}.ToString()
	$NewLog = ${function:New-Log}.ToString()
	$moduleDataArray | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
		$moduleData = $_
		$galleryModule = $null
		$moduleName = $moduleData.Name
		$repositories = $using:Repositories
		$counters = $using:counters
		$results = $using:results
		$timeoutSeconds = $using:TimeoutSeconds
		$blackList = $using:BlackList
		${function:Compare-PrereleaseVersion} = $using:ComparePrereleaseVersion
		$isBlacklisted = $false
		$repositoriesToCheck = $repositories
		if ($blackList.ContainsKey($moduleName)) {
			$blacklistedRepos = $blackList[$moduleName]
			if ($blacklistedRepos -is [array] -or $blacklistedRepos -is [System.Collections.ArrayList]) {
				if ($blacklistedRepos.Count -eq 0 -or ($repositories | Where-Object { $blacklistedRepos -notcontains $_ }).Count -eq 0) {
					$isBlacklisted = $true
				}
				else {
					$repositoriesToCheck = $repositories | Where-Object { $blacklistedRepos -notcontains $_ }
				}
			}
			elseif ($blacklistedRepos -is [string]) {
				if ($blacklistedRepos -eq "*" -or $repositories -contains $blacklistedRepos) {
					$isBlacklisted = $true
				}
				else {
					$repositoriesToCheck = $repositories | Where-Object { $_ -ne $blacklistedRepos }
				}
			}
			else {
				$isBlacklisted = $true
			}
		}
		${function:New-Log} = $using:NewLog
		if (-not $isBlacklisted -and $repositoriesToCheck.Count -gt 0) {
			$galleryModule = Find-Module -Name $moduleName -AllowPrerelease -Repository $repositories -ErrorAction SilentlyContinue | Sort-Object -Property version -Descending -ErrorAction SilentlyContinue | Select-Object -First 1 -ErrorAction SilentlyContinue
			try {
				if ($galleryModule) {
					$highestLocal = $moduleData.HighestLocalVersionInfo.ParsedVersion.Version
					$localPreReleaseLabel = $moduleData.PreReleaseLabel
					$compared = $false
					if ($localPreReleaseLabel -and $galleryModule.Prerelease) {
						$compared = Compare-PrereleaseVersion -VersionA $localPreReleaseLabel -VersionB $galleryModule.Prerelease -ReturnBoolean
					}
					if ($galleryModule.Version -gt $highestLocal -or ($compared -and $galleryModule.Version -eq $highestLocal)) {
						$outdatedPaths = $moduleData.AllVersions | Where-Object { $_.ParsedVersion.Version -lt [version]$galleryModule.Version } | Select-Object -ExpandProperty Path -Unique
						if ($outdatedPaths) {
							$results.Add([PSCustomObject]@{
									ModuleName        = $moduleName
									LatestVersion     = [version]$galleryModule.Version
									LocalVersion      = $highestLocal
									IsPreview         = [bool]$galleryModule.Prerelease
									Repository        = $galleryModule.Repository
									OutdatedPaths     = $outdatedPaths
									PreReleaseVersion = $galleryModule.Prerelease
								})
							$counters.Updates++
						}
					}
				}
			}
			catch {
				$counters.Errors++
			}
		}
		else {
			New-log "Found update for $moduleName but that module and repository combo: $blacklistedRepos is blacklisted." -Level WARNING
		}
		$currentCount = ++$counters.Processed
		if ($currentCount % 5 -eq 0 -or $currentCount -eq $counters.Total) {
			$elapsed = (Get-Date) - $counters.StartTime
			$percentComplete = [Math]::Round(($currentCount / $counters.Total) * 100)
			$eta = "Calculating..."
			if ($currentCount -gt 5) {
				$avgTimePerModule = $elapsed.TotalSeconds / $currentCount
				$remainingModules = $counters.Total - $currentCount
				$remainingSeconds = $avgTimePerModule * $remainingModules
				if ($remainingSeconds -gt 0) {
					$eta = "{0:mm\:ss}" -f [timespan]::FromSeconds($remainingSeconds)
				}
				else {
					$eta = "00:00"
				}
			}
			$progressMsg = "[$percentComplete%] Processed $currentCount of $($counters.Total) modules, ETA: $eta, Updates: $($counters.Updates), Errors: $($counters.Errors)"
			New-Log $progressMsg -Level SUCCESS
		}
	}
	$finalResults = @($results)
	$totalTime = (Get-Date) - $counters.StartTime
	$secondsElapsed = $totalTime.TotalSeconds.ToString('0.0')
	New-Log "Completed check of $validModuleCount modules in $secondsElapsed seconds. Found $($counters.Updates) modules needing updates." -Level DEBUG
	New-Log "Encountered $($counters.Errors) errors during processing." -Level DEBUG
	return $finalResults
}
function Update-PSModules {
	[CmdletBinding()]
	param (
		[Parameter(ValueFromPipeline)][Object[]]$OutdatedModules,
		[switch]$Clean,
		[switch]$UseProgressBar,
		[switch]$PreRelease
	)
	begin {
		$moduleResults = @()
		$batchModules = @()
	}
	process {
		foreach ($module in $OutdatedModules) {
			$batchModules += $module
		}
	}
	end {
		if(!$batchModules){
			New-Log "No modules to update, will aboart."
			return
		}
		$total = $batchModules.Count
		$current = 0
		foreach ($module in $batchModules) {
			$script:moduleResult = [PSCustomObject]@{}
			$current++
			if ($UseProgressBar.IsPresent) {
				Write-Progress -Activity "Updating PowerShell Modules" -Status $module.ModuleName -PercentComplete (($current / $total) * 100)
			}
			$moduleName = $module.ModuleName
			$latestVer = $module.LatestVersion
			$localVers = $module.LocalVersion
			$isPreview = $PreRelease.IsPresent -or $module.IsPreview
			$paths = $module.OutdatedPaths
			$repository = $module.Repository
			New-Log "[$current/$total] Updating [$moduleName] to version [$latestVer] (Preview=$isPreview)." -Level DEBUG
			$script:moduleResult = [PSCustomObject]@{
				ModuleName     = $moduleName
				NewVersion     = $latestVer
				UpdatedPaths   = @()
				FailedPaths    = @()
				OverallSuccess = $false
			}
			try {
				$foundModule = Find-Module -Name $moduleName -Repository $repository -AllowPrerelease:$isPreview -ErrorAction Stop
				New-Log "Found module [$moduleName] version [$($foundModule.Version)] in repository [$repository]" -Level SUCCESS
			}
			catch {
				try {
					$foundModule = Find-Module -Name $moduleName -Repository $repository -ErrorAction Stop
					New-Log "Found module [$moduleName] version [$($foundModule.Version)] in repository [$repository]" -Level SUCCESS
				}
				catch {
					New-Log "Failed to find module [$moduleName] in repository [$repository]" -Level ERROR
					$script:moduleResult.FailedPaths = "Repository lookup failed"
					$moduleResults += $script:moduleResult
					continue
				}
			}
			$uniquePaths = $paths | Group-Object { ($_ -split '\\Modules\\')[0] } | ForEach-Object { $_.Group[0] }
			if ($uniquePaths.Count -gt 0) {
				Install-PSModule -ModuleName $moduleName -IsPreview $isPreview -Repository $repository -FoundModule $foundModule -Destinations $uniquePaths
			}
			$script:moduleResult.OverallSuccess = ($script:moduleResult.FailedPaths.Count -eq 0)
			if ($script:moduleResult.UpdatedPaths.Count -ge 1 -and $Clean.IsPresent -and $script:moduleResult.UpdatedPaths -in $uniquePaths) {
				$cleanedPaths = Remove-OutdatedVersions -ModuleName $moduleName -Paths $script:moduleResult.UpdatedPaths -LatestVersion $latestVer -LocalVersions $localVers
				if ($cleanedPaths) {
					$script:moduleResult | Add-Member -NotePropertyName "CleanedPaths" -NotePropertyValue $cleanedPaths
				}
			}
			$moduleResults += $script:moduleResult
		}
		if ($UseProgressBar) {
			Write-Progress -Activity "Updating PowerShell Modules" -Completed
		}
		return $moduleResults
	}
}
function Install-PSModule {
	[CmdletBinding()]
	param (
		[string]$ModuleName,
		[bool]$IsPreview,
		[string[]]$Repository = @('PSGallery'),
		[PSObject]$FoundModule,
		[string[]]$Destinations
	)
	$commonParams = @{
		Name                = $ModuleName
		Version             = $($FoundModule.Version)
		IncludeXml          = $true
		TrustRepository     = $true
		SkipDependencyCheck = $true
		AcceptLicense       = $true
		Confirm             = $false
		PassThru            = $true
		Prerelease          = $IsPreview
	}
	$max = $Destinations.Count
	$current = 0
	foreach ($Destination in $Destinations) {
		$res = $null
		$path = ($destination -replace "$ModuleName", "")
		$res = Save-PSResource @commonParams -Repository $Repository -Path $path -ErrorAction SilentlyContinue
		if ($res -and $res.Version -eq $FoundModule.Version) {
			New-Log "Successfully updated [$ModuleName] to [$($res.Version)] via Save-PSResource to $destination" -Level SUCCESS
			$script:moduleResult.UpdatedPaths += $destination
			$current++
		}
		else {
			New-Log "Failed to update [$ModuleName] to [$($FoundModule.Version)] via Save-PSResource to $destination" -Level WARNING
			$script:moduleResult.FailedPaths += $destination
		}
	}
	if ($max -eq $current) {
		New-Log "Successfully updated [$ModuleName] to [$($FoundModule.Version)] on all destinations!" -Level SUCCESS
		return
	}
	else {
		New-Log "Will try a different way..." -Level DEBUG
	}
	try {
		$psResourceParams = @{
			Name                = $ModuleName
			Version             = $($FoundModule.Version)
			Scope               = 'AllUsers'
			AcceptLicense       = $true
			SkipDependencyCheck = $true
			Confirm             = $false
			Reinstall           = $true
			TrustRepository     = $true
			Repository          = $Repository
			PassThru            = $true
			ErrorAction         = 'Stop'
			WarningAction       = 'SilentlyContinue'
			Prerelease          = $IsPreview
		}
		$res = Install-PSResource @psResourceParams
		$installLocation = $res.InstalledLocation
		$reCheck = "$installLocation\$ModuleName"
		$installedVersion = $res.Version.ToString()
		if ($res -and $installLocation -and $installedVersion -and $reCheck -in $script:moduleResult.FailedPaths) {
			New-Log "Successfully updated [$ModuleName] to [$installedVersion] via Install-PSResource to $reCheck" -Level SUCCESS
			$script:moduleResult.FailedPaths = $script:moduleResult.FailedPaths | Where-Object { $_ -ne $reCheck }
			$script:moduleResult.UpdatedPaths += $reCheck
		}
		elseif ($res -and $installLocation -and $installedVersion) {
			New-Log "Install-PSResource managed to install the latest version in a path where it was not present before." -Level WARNING
			New-Log "$moduleName with version $installedVersion is now present at: $recheck" -Level SUCCESS
			$script:moduleResult.UpdatedPaths += $reCheck
		}
	}
	catch {
		New-Log "Install-PSResource failed for [$ModuleName]. Trying Install-Module... with -AllowClobber" -Level WARNING
		try {
			$installModuleParams = @{
				Name               = $ModuleName
				Scope              = 'AllUsers'
				Force              = $true
				AcceptLicense      = $true
				SkipPublisherCheck = $true
				AllowClobber       = $true
				PassThru           = $true
				ErrorAction        = 'Stop'
				WarningAction      = 'SilentlyContinue'
				Confirm            = $false
				AllowPrerelease    = $IsPreview
			}
			$res = Install-Module @installModuleParams
			$installLocation = $res.InstalledLocation
			$reCheck = "$installLocation\$ModuleName"
			$installedVersion = $res.Version.ToString()
			if ($res -and $installLocation -and $installedVersion -and $reCheck -in $script:moduleResult.FailedPaths) {
				New-Log "Last chance successfully updated [$ModuleName] to [$installedVersion] via Install-PSResource to $reCheck" -Level SUCCESS
				$script:moduleResult.FailedPaths = $script:moduleResult.FailedPaths | Where-Object { $_ -ne $reCheck }
				$script:moduleResult.UpdatedPaths += $reCheck
			}
			elseif ($res -and $installLocation -and $installedVersion) {
				New-Log "Last chance managed to install the latest version in a path where it was not present before." -Level WARNING
				New-Log "$moduleName with version $installedVersion is now present at: $recheck" -Level SUCCESS
				$script:moduleResult.UpdatedPaths += $reCheck
			}
		}
		catch {
			New-Log "Failed last chance to update [$ModuleName].." -Level ERROR
		}
	}
}
function Remove-OutdatedVersions {
	[CmdletBinding()]
	param (
		[string]$ModuleName,
		[string[]]$Paths,
		[version]$LatestVersion,
		[string[]]$DoNotClean = @('PowerShellGet'),
		[string[]]$LocalVersions
	)
	if ($ModuleName -in $DoNotClean) {
		New-Log "Skipping cleaning of [$ModuleName] since it's blacklisted"
		return
	}
	New-Log "Cleaning old folders for [$ModuleName]..." -Level DEBUG
	$cleanedPaths = @()
	foreach ($outdatedPath in $Paths) {
		$leaf = Split-Path $outdatedPath -Leaf
		$parentPath = Split-Path $outdatedPath -Parent
		$versionFolders = Get-ChildItem -Path $parentPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d+(\.\d+)+$' -and $_.Name -ne $LatestVersion }
		if (!$versionFolders) {
			$versionFolders = Get-ChildItem -Path $outdatedPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d+(\.\d+)+$' -and $_.Name -ne $LatestVersion } | Sort-Object -Property Name -Descending #| Select-Object -First 1
		}
		try {
			if (-not $versionFolders -or $versionFolders.Count -eq 0) {
				if ($LocalVersions) {
					foreach ($LocalVersion in ($LocalVersions | Where-Object { $_.ToString() -ne $latestVer }) ) {
						try {
							Uninstall-PSResource -Name $ModuleName -Version $LocalVersion.ToString() -Scope AllUsers -Confirm:$false -ErrorAction Stop
							Start-Sleep -Seconds 1
							if (!(Test-Path -Path $outdatedPath)) {
								New-Log "Successfully removed $moduleName version $($LocalVersion.ToString())" -Level SUCCESS
								$cleanedPaths += $outdatedPath
							}
						}
						catch {
							New-Log "Failed to remove $moduleName with version $($LocalVersion.ToString())" -Level WARNING
						}
					}
				}
				else {
					New-Log "Skipping removal of $moduleName : No version folders found in $parentPath" -Level WARNING
				}
				continue
			}
			foreach ($versionFolder in $versionFolders) {
				if (Test-Path -Path $versionFolder.FullName) {
					$leaf = $versionFolder.Name
					Uninstall-PSResource -Name $ModuleName -Version $leaf -Scope AllUsers -Confirm:$false -ErrorAction SilentlyContinue
					Start-Sleep -Seconds 1
					if (Test-Path -Path $versionFolder.FullName) {
						Remove-Item -LiteralPath $versionFolder.FullName -Recurse -Force -ErrorAction Stop | Out-Null
					}
					$cleanedPaths += $versionFolder.FullName
					New-Log "Successfully removed old modulepath: $($versionFolder.FullName)" -Level SUCCESS
				}
			}
		}
		catch {
			New-Log "Failed to remove outdated folder [$outdatedPath]" -Level ERROR
		}
	}
	return $cleanedPaths
}
### OBS: New-Log Function is needed otherwise remove all New-Log and replace with Write-Host. New-Log is vastly better though, check the link below:
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
<# Example usecase:

Check-PSResourceRepository -ImportDependencies
$ignoredModules = @('Defender', 'DnsClient', 'BurntToast') #Fully ignored modules
$blackList = @{ #Ignored module and repo combo.
    'Microsoft.Graph.Beta' = 'NuGetGallery'
    'Microsoft.Graph' = 'NuGetGallery'
	'ScheduledTasks' = 'NuGetGallery'
	'iSCSI' = 'NuGetGallery'
	'ActiveDirectory' = 'NuGetGallery'
}
$paths = $env:PSModulePath.Split(';') | Where-Object {$_ -inotmatch '.vscode'}
$moduleInfo = Get-ModuleInfo -Paths $paths -IncludeResourceFiles -IgnoredModules $ignoredModules
$outdated = Get-ModuleUpdateStatus -ModuleInventory $moduleInfo -TimeoutSeconds 180 -Repositories @("PSGallery", "Nuget", "NugetGallery") -BlackList $blackList
$res = $outdated | Update-PSModules -Clean -UseProgressBar -PreRelease
#>