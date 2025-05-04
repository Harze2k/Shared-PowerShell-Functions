#Requires -Version 7.0
<#
Author: Harze2k
Date:   2025-05-05
Version: 2.6 (Added some helpful comments.)

Sample output:

[2025-05-04 23:21:28.486][INFO] TLS 1.2 security protocol enabled for this session.
[2025-05-04 23:21:28.527][INFO] Parameter -ImportDependencies specified. Checking/installing dependencies...
[2025-05-04 23:21:28.598][INFO] PSGallery repository is already trusted.
[2025-05-04 23:21:28.846][INFO] Module 'Microsoft.PowerShell.PSResourceGet' version [1.1.1] is already installed and available.
[2025-05-04 23:21:28.864][SUCCESS] Successfully imported module 'Microsoft.PowerShell.PSResourceGet'.
[2025-05-04 23:21:29.833][INFO] Module 'PowerShellGet' version [3.0.23] is already installed and available.
[2025-05-04 23:21:29.839][SUCCESS] Successfully imported module 'PowerShellGet'.
[2025-05-04 23:21:29.842][SUCCESS] Required module commands are now available.
[2025-05-04 23:21:29.844][INFO] Starting repository configuration...
[2025-05-04 23:21:29.853][INFO] 'PSGallery' repository is already registered and configured correctly.
[2025-05-04 23:21:29.860][INFO] 'NuGetGallery' repository is already registered and configured correctly.
[2025-05-04 23:21:29.863][INFO] 'NuGet' repository is already registered and configured correctly.
[2025-05-04 23:21:29.864][SUCCESS] All specified repositories appear to be registered and configured.

[2025-05-04 23:23:44.370][DEBUG] Finished processing ActiveDirectory. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.372][DEBUG] Finished processing ADEssentials. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.373][DEBUG] Finished processing AOVPNTools. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.374][DEBUG] Finished processing AppBackgroundTask. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.376][DEBUG] Finished processing AppLocker. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.379][DEBUG] Finished processing AppvClient. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.381][DEBUG] Finished processing Appx. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.383][DEBUG] Finished processing Az.Accounts. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.385][DEBUG] Finished processing AzureAD. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.388][DEBUG] Finished processing AzureStackHCI. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.389][DEBUG] Finished processing BestPractices. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.391][DEBUG] Finished processing BitLocker. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.393][DEBUG] Finished processing BitsTransfer. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.396][DEBUG] Finished processing BranchCache. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.401][DEBUG] Finished processing BurntToast. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.403][DEBUG] Finished processing CimCmdlets. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.407][DEBUG] Finished processing ClusterAwareUpdating. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.410][DEBUG] Finished processing CommonStuff. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.412][DEBUG] Finished processing CompletionPredictor. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.434][DEBUG] Finished processing ConfigCI. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.447][DEBUG] Finished processing ConfigDefender. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.453][DEBUG] Finished processing ConfigDefenderPerformance. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.458][DEBUG] Finished processing Defender. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.462][DEBUG] Finished processing DefenderPerformance. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.471][DEBUG] Finished processing DeliveryOptimization. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.474][DEBUG] Finished processing DependencySearch. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.478][DEBUG] Finished processing dfsn. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.481][DEBUG] Finished processing DFSR. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.484][DEBUG] Finished processing DhcpServer. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.487][DEBUG] Finished processing DirectAccessClientComponents. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.490][DEBUG] Finished processing Dism. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.492][DEBUG] Finished processing DnsClient. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.494][DEBUG] Finished processing DnsServer. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.497][DEBUG] Finished processing DSCFileDownloadManager. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.500][DEBUG] Finished processing EditorServicesCommandSuite. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.502][DEBUG] Finished processing EnhancedPSTools. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.504][DEBUG] Finished processing EventTracingManagement. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.507][DEBUG] Finished processing Example2.Diagnostics. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.511][DEBUG] Finished processing ExchangeOnlineManagement. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.514][DEBUG] Finished processing Get-NetView. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.517][DEBUG] Finished processing GroupPolicy. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.519][DEBUG] Finished processing GroupSet. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.521][DEBUG] Finished processing HtmlAgilityPack. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.524][DEBUG] Finished processing ImportExcel. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.526][DEBUG] Finished processing Indented.Net.IP. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.528][DEBUG] Finished processing Indented.ScriptAnalyzerRules. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.531][DEBUG] Finished processing International. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.534][DEBUG] Finished processing IntuneStuff. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.537][DEBUG] Finished processing IpamServer. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.539][DEBUG] Finished processing iSCSI. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.541][DEBUG] Finished processing IscsiTarget. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.543][DEBUG] Finished processing Kds. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.547][DEBUG] Finished processing LanguagePackManagement. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.549][DEBUG] Finished processing LAPS. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.552][DEBUG] Finished processing LSUClient. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.553][DEBUG] Finished processing Microsoft.Graph. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.555][DEBUG] Finished processing Microsoft.Graph.Authentication. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.557][DEBUG] Finished processing Microsoft.Graph.Beta. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.559][DEBUG] Finished processing Microsoft.Graph.Beta.DeviceManagement. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.560][DEBUG] Finished processing Microsoft.Graph.Beta.DeviceManagement.Actions. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.563][DEBUG] Finished processing Microsoft.Graph.Beta.DeviceManagement.Administration. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.566][DEBUG] Finished processing Microsoft.Graph.Beta.DeviceManagement.Enrollment. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.568][DEBUG] Finished processing Microsoft.Graph.Beta.DeviceManagement.Functions. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.572][DEBUG] Finished processing Microsoft.Graph.Beta.Devices.CorporateManagement. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.577][DEBUG] Finished processing Microsoft.Graph.Beta.Groups. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.579][DEBUG] Finished processing Microsoft.Graph.Beta.Identity.DirectoryManagement. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.581][DEBUG] Finished processing Microsoft.Graph.Beta.Users. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.584][DEBUG] Finished processing Microsoft.Graph.Beta.Users.Actions. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.586][DEBUG] Finished processing Microsoft.Graph.Beta.Users.Functions. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.588][DEBUG] Finished processing Microsoft.Graph.Core. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.590][DEBUG] Finished processing Microsoft.Graph.DeviceManagement. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.591][DEBUG] Finished processing Microsoft.Graph.DeviceManagement.Actions. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.594][DEBUG] Finished processing Microsoft.Graph.DeviceManagement.Administration. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.596][DEBUG] Finished processing Microsoft.Graph.DeviceManagement.Enrollment. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.599][DEBUG] Finished processing Microsoft.Graph.DeviceManagement.Functions. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.605][DEBUG] Finished processing Microsoft.Graph.Devices.CorporateManagement. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.609][DEBUG] Finished processing Microsoft.Graph.DirectoryObjects. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.611][DEBUG] Finished processing Microsoft.Graph.Groups. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.613][DEBUG] Finished processing Microsoft.Graph.Identity.DirectoryManagement. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.617][DEBUG] Finished processing Microsoft.Graph.Intune. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.620][DEBUG] Finished processing Microsoft.PowerApps.Administration.PowerShell. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.622][DEBUG] Finished processing Microsoft.PowerShell.Archive. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.624][DEBUG] Finished processing Microsoft.PowerShell.Diagnostics. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.626][DEBUG] Finished processing Microsoft.PowerShell.Host. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.627][DEBUG] Finished processing Microsoft.PowerShell.Management. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.629][DEBUG] Finished processing Microsoft.PowerShell.ODataUtils. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.632][DEBUG] Finished processing Microsoft.PowerShell.Operation.Validation. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.636][DEBUG] Finished processing Microsoft.PowerShell.PSResourceGet. Found 4 unique BasePath/Version combinations.
[2025-05-04 23:23:44.640][DEBUG] Finished processing Microsoft.PowerShell.Security. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.642][DEBUG] Finished processing Microsoft.PowerShell.Utility. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.644][DEBUG] Finished processing Microsoft.ReFsDedup.Commands. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.647][DEBUG] Finished processing Microsoft.WSMan.Management. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.649][DEBUG] Finished processing MMAgent. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.652][DEBUG] Finished processing MSAL.PS. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.655][DEBUG] Finished processing MsDtc. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.657][DEBUG] Finished processing MSFT_NfsMappedIdentity. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.659][DEBUG] Finished processing MSGraphStuff. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.661][DEBUG] Finished processing NetAdapter. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.664][DEBUG] Finished processing NetConnection. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.666][DEBUG] Finished processing NetEventPacketCapture. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.668][DEBUG] Finished processing NetLbfo. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.670][DEBUG] Finished processing NetLldpAgent. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.674][DEBUG] Finished processing NetNat. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.676][DEBUG] Finished processing NetQos. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.678][DEBUG] Finished processing NetSwitchTeam. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.680][DEBUG] Finished processing NetworkConnectivityStatus. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.682][DEBUG] Finished processing NetworkControllerDiagnostics. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.685][DEBUG] Finished processing NetworkControllerFc. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.687][DEBUG] Finished processing NetworkLoadBalancingClusters. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.689][DEBUG] Finished processing NetworkSwitchManager. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.691][DEBUG] Finished processing NetworkTransition. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.692][DEBUG] Finished processing nfs. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.694][DEBUG] Finished processing OsConfiguration. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.696][DEBUG] Finished processing PackageManagement. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.698][DEBUG] Finished processing PcsvDevice. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.700][DEBUG] Finished processing PersistentMemory. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.706][DEBUG] Finished processing Pester. Found 4 unique BasePath/Version combinations.
[2025-05-04 23:23:44.709][DEBUG] Finished processing pki. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.710][DEBUG] Finished processing PnpDevice. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.712][DEBUG] Finished processing PolicyFileEditor. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.715][DEBUG] Finished processing PowerHTML. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.717][DEBUG] Finished processing PowerQuest. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.721][DEBUG] Finished processing PowerShellGet. Found 12 unique BasePath/Version combinations.
[2025-05-04 23:23:44.723][DEBUG] Finished processing PrintManagement. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.725][DEBUG] Finished processing ProcessMitigations. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.727][DEBUG] Finished processing ProcessSet. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.729][DEBUG] Finished processing provisioning. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.732][DEBUG] Finished processing ps2exe. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.736][DEBUG] Finished processing PSDesiredStateConfiguration. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.738][DEBUG] Finished processing PSDev. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.740][DEBUG] Finished processing PSDiagnostics. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.742][DEBUG] Finished processing PSEventViewer. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.744][DEBUG] Finished processing PSFramework. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.746][DEBUG] Finished processing PSModuleDevelopment. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.747][DEBUG] Finished processing PSPreworkout. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.751][DEBUG] Finished processing PSReadLine. Found 4 unique BasePath/Version combinations.
[2025-05-04 23:23:44.753][DEBUG] Finished processing PSScheduledJob. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.756][DEBUG] Finished processing PSSharedGoods. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.758][DEBUG] Finished processing PSWindowsUpdate. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.760][DEBUG] Finished processing PSWorkflow. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.762][DEBUG] Finished processing PSWorkflowUtility. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.764][DEBUG] Finished processing PSWriteHTML. Found 3 unique BasePath/Version combinations.
[2025-05-04 23:23:44.768][DEBUG] Finished processing Refactor. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.770][DEBUG] Finished processing RemoteAccess. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.772][DEBUG] Finished processing RemoteDesktop. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.774][DEBUG] Finished processing ScheduledTasks. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.775][DEBUG] Finished processing SecureBoot. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.777][DEBUG] Finished processing ServerManager. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.778][DEBUG] Finished processing ServerManagerTasks. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.780][DEBUG] Finished processing ServiceSet. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.781][DEBUG] Finished processing SharePointPnPPowerShellOnline. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.783][DEBUG] Finished processing SmbShare. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.786][DEBUG] Finished processing SmbWitness. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.787][DEBUG] Finished processing SplitPipeline. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.789][DEBUG] Finished processing StartLayout. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.790][DEBUG] Finished processing StorageMigrationService. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.791][DEBUG] Finished processing storagereplica. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.794][DEBUG] Finished processing string. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.796][DEBUG] Finished processing SystemInsights. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.800][DEBUG] Finished processing ThreadJob. Found 4 unique BasePath/Version combinations.
[2025-05-04 23:23:44.802][DEBUG] Finished processing tls. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.804][DEBUG] Finished processing TroubleshootingPack. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.806][DEBUG] Finished processing TrustedPlatformModule. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.807][DEBUG] Finished processing UEV. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.808][DEBUG] Finished processing VMDirectStorage. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.809][DEBUG] Finished processing VpnClient. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.814][DEBUG] Finished processing Wdac. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.816][DEBUG] Finished processing WebDownloadManager. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.817][DEBUG] Finished processing Whea. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.820][DEBUG] Finished processing WindowsAutoPilotIntune. Found 2 unique BasePath/Version combinations.
[2025-05-04 23:23:44.822][DEBUG] Finished processing WindowsDeveloperLicense. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.823][DEBUG] Finished processing WindowsErrorReporting. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.827][DEBUG] Finished processing WindowsFeatureSet. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.828][DEBUG] Finished processing WindowsOptionalFeatureSet. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.830][DEBUG] Finished processing WindowsPackageCab. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.832][DEBUG] Finished processing WindowsSearch. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.833][DEBUG] Finished processing WindowsUpdate. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.836][DEBUG] Finished processing WinHttpProxy. Found 1 unique BasePath/Version combinations.
[2025-05-04 23:23:44.842][WARNING] Skipping Example2.Diagnostics since it's on the ignorelist.
[2025-05-04 23:23:44.845][WARNING] Skipping string since it's on the ignorelist.

[2025-05-04 23:24:53.586][SUCCESS] Starting parallel module update check for 176 modules (Throttle: 24, Timeout: 90s)...
[2025-05-04 23:24:53.648][DEBUG] Prepared 176 modules with valid version info for checking.
[2025-05-04 23:24:56.468][SUCCESS] Job 2 completed. Found 'AOVPNTools' version 1.9.4.
[2025-05-04 23:24:56.468][SUCCESS] Job 1 completed. Found 'ADEssentials' version 0.0.237.
[2025-05-04 23:24:56.529][SUCCESS] Job 19 completed. Found 'AzureAD' version 2.0.2.182.
[2025-05-04 23:24:57.022][SUCCESS] Job 15 completed. Found 'Az.Accounts' version 4.1.0.
[2025-05-04 23:24:57.145][SUCCESS] Job 29 completed. Found 'BurntToast' version 1.0.0.
[2025-05-04 23:24:57.153][WARNING] Filtering out BurntToast since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:24:57.377][SUCCESS] Job 35 completed. Found 'CommonStuff' version 1.0.22.
[2025-05-04 23:24:57.869][SUCCESS] Job 41 completed. Found 'CompletionPredictor' version 0.1.1.
[2025-05-04 23:24:58.025][SUCCESS] Job 47 completed. Found 'Defender' version 6.0.0.
[2025-05-04 23:24:58.026][WARNING] Filtering out Defender since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:24:59.132][SUCCESS] Job 51 completed. Found 'DependencySearch' version 1.1.7.
[2025-05-04 23:24:59.558][SUCCESS] Job 63 completed. Found 'DnsClient' version 1.8.0.
[2025-05-04 23:24:59.563][WARNING] Filtering out DnsClient since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:24:59.943][SUCCESS] Job 69 completed. Found 'EditorServicesCommandSuite' version 1.0.0.
[2025-05-04 23:24:59.994][DEBUG] Comparing pre-release versions: beta4 not newer then beta4
[2025-05-04 23:25:00.007][SUCCESS] Job 65 completed. Found 'DnsServer' version 1.0.0.
[2025-05-04 23:25:00.015][WARNING] Filtering out DnsServer since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:25:00.146][SUCCESS] Job 85 completed. Found 'ImportExcel' version 7.8.10.
[2025-05-04 23:25:00.155][SUCCESS] Job 73 completed. Found 'ExchangeOnlineManagement' version 3.8.0.
[2025-05-04 23:25:00.164][SUCCESS] Online pre-release 'Preview2' is newer than local 'Preview1' for version 3.8.0.
[2025-05-04 23:25:00.194][SUCCESS] Update found for 'ExchangeOnlineManagement': Local '3.8.0-Preview1' -> Online '3.8.0'
[2025-05-04 23:25:00.571][SUCCESS] Job 71 completed. Found 'EnhancedPSTools' version 0.0.38.
[2025-05-04 23:25:00.648][SUCCESS] Job 77 completed. Found 'Get-NetView' version 2025.2.26.254.
[2025-05-04 23:25:00.653][SUCCESS] Update found for 'Get-NetView': Local '2023.2.7.226' -> Online '2025.2.26.254'
[2025-05-04 23:25:00.943][SUCCESS] Job 87 completed. Found 'Indented.Net.IP' version 6.3.2.
[2025-05-04 23:25:01.320][SUCCESS] Job 83 completed. Found 'HtmlAgilityPack' version 1.12.1.
[2025-05-04 23:25:01.320][SUCCESS] Job 93 completed. Found 'IntuneStuff' version 1.6.2.
[2025-05-04 23:25:03.627][SUCCESS] Job 97 completed. Found 'iSCSI' version 1.5.5.1.
[2025-05-04 23:25:03.665][WARNING] Filtering out iSCSI since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:25:04.574][SUCCESS] Job 109 completed. Found 'Microsoft.Graph' version 2.27.0.
[2025-05-04 23:25:04.602][SUCCESS] Job 115 completed. Found 'Microsoft.Graph.Beta' version 2.27.0.
[2025-05-04 23:25:04.618][SUCCESS] Job 107 completed. Found 'LSUClient' version 1.7.1.
[2025-05-04 23:25:05.214][SUCCESS] Job 111 completed. Found 'Microsoft.Graph.Authentication' version 2.27.0.
[2025-05-04 23:25:05.262][SUCCESS] Job 113 completed. Found 'Microsoft.Graph.Beta.DeviceManagement.Actions' version 2.25.0.
[2025-05-04 23:25:05.359][SUCCESS] Update found for 'Microsoft.Graph.Beta.DeviceManagement.Actions': Local '2.24.0' -> Online '2.25.0'
[2025-05-04 23:25:05.378][SUCCESS] Update found for 'Microsoft.Graph.Beta.DeviceManagement.Actions': Local '2.24.0' -> Online '2.25.0'
[2025-05-04 23:25:05.419][SUCCESS] Job 125 completed. Found 'Microsoft.Graph.Beta.Groups' version 2.27.0.
[2025-05-04 23:25:05.529][SUCCESS] Job 121 completed. Found 'Microsoft.Graph.Beta.DeviceManagement.Enrollment' version 2.27.0.
[2025-05-04 23:25:05.710][SUCCESS] Job 119 completed. Found 'Microsoft.Graph.Beta.DeviceManagement.Administration' version 2.27.0.
[2025-05-04 23:25:05.762][SUCCESS] Job 123 completed. Found 'Microsoft.Graph.Beta.Devices.CorporateManagement' version 2.27.0.
[2025-05-04 23:25:05.973][SUCCESS] Job 129 completed. Found 'Microsoft.Graph.Beta.DeviceManagement.Functions' version 2.27.0.
[2025-05-04 23:25:06.039][SUCCESS] Job 117 completed. Found 'Microsoft.Graph.Beta.DeviceManagement' version 2.27.0.
[2025-05-04 23:25:06.044][SUCCESS] Job 133 completed. Found 'Microsoft.Graph.Beta.Users.Actions' version 2.27.0.
[2025-05-04 23:25:06.090][SUCCESS] Job 135 completed. Found 'Microsoft.Graph.Beta.Users.Functions' version 2.27.0.
[2025-05-04 23:25:06.099][SUCCESS] Job 127 completed. Found 'Microsoft.Graph.Beta.Identity.DirectoryManagement' version 2.27.0.
[2025-05-04 23:25:06.356][SUCCESS] Job 131 completed. Found 'Microsoft.Graph.Beta.Users' version 2.27.0.
[2025-05-04 23:25:07.468][SUCCESS] Job 137 completed. Found 'Microsoft.Graph.Core' version 3.2.4.
[2025-05-04 23:25:07.684][SUCCESS] Job 143 completed. Found 'Microsoft.Graph.DeviceManagement.Administration' version 2.27.0.
[2025-05-04 23:25:07.721][SUCCESS] Job 139 completed. Found 'Microsoft.Graph.DeviceManagement' version 2.27.0.
[2025-05-04 23:25:07.821][SUCCESS] Job 141 completed. Found 'Microsoft.Graph.DeviceManagement.Actions' version 2.25.0.
[2025-05-04 23:25:07.823][SUCCESS] Update found for 'Microsoft.Graph.DeviceManagement.Actions': Local '2.24.0' -> Online '2.25.0'
[2025-05-04 23:25:07.852][SUCCESS] Update found for 'Microsoft.Graph.DeviceManagement.Actions': Local '2.24.0' -> Online '2.25.0'
[2025-05-04 23:25:08.031][SUCCESS] Job 145 completed. Found 'Microsoft.Graph.DeviceManagement.Enrollment' version 2.27.0.
[2025-05-04 23:25:08.152][SUCCESS] Job 147 completed. Found 'Microsoft.Graph.DeviceManagement.Functions' version 2.27.0.
[2025-05-04 23:25:08.867][SUCCESS] Job 155 completed. Found 'Microsoft.Graph.Identity.DirectoryManagement' version 2.27.0.
[2025-05-04 23:25:09.249][SUCCESS] Job 153 completed. Found 'Microsoft.Graph.Groups' version 2.27.0.
[2025-05-04 23:25:09.482][SUCCESS] Job 151 completed. Found 'Microsoft.Graph.DirectoryObjects' version 2.27.0.
[2025-05-04 23:25:09.802][SUCCESS] Job 159 completed. Found 'Microsoft.PowerApps.Administration.PowerShell' version 2.0.210.
[2025-05-04 23:25:09.807][SUCCESS] Job 149 completed. Found 'Microsoft.Graph.Devices.CorporateManagement' version 2.27.0.
[2025-05-04 23:25:10.322][SUCCESS] Job 157 completed. Found 'Microsoft.Graph.Intune' version 6.1907.1.0.
[2025-05-04 23:25:10.664][SUCCESS] Job 161 completed. Found 'Microsoft.PowerShell.Archive' version 1.2.5.
[2025-05-04 23:25:10.666][SUCCESS] Update found for 'Microsoft.PowerShell.Archive': Local '1.0.1.0' -> Online '1.2.5'
[2025-05-04 23:25:11.214][SUCCESS] Job 173 completed. Found 'Microsoft.PowerShell.PSResourceGet' version 1.1.1.
[2025-05-04 23:25:12.188][SUCCESS] Job 187 completed. Found 'MSAL.PS' version 4.37.0.0.
[2025-05-04 23:25:12.238][SUCCESS] Job 177 completed. Found 'Microsoft.PowerShell.Security' version 7.6.0.
[2025-05-04 23:25:12.250][DEBUG] Comparing pre-release versions: preview.4 not newer then preview.4
[2025-05-04 23:25:12.254][WARNING] Filtering out Microsoft.PowerShell.Security since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:25:12.447][SUCCESS] Job 181 completed. Found 'Microsoft.WSMan.Management' version 7.6.0.
[2025-05-04 23:25:12.483][DEBUG] Comparing pre-release versions: preview.4 not newer then preview.4
[2025-05-04 23:25:12.492][WARNING] Filtering out Microsoft.WSMan.Management since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:25:12.553][SUCCESS] Job 189 completed. Found 'MSGraphStuff' version 1.1.1.
[2025-05-04 23:25:16.711][SUCCESS] Job 229 completed. Found 'Pester' version 6.0.0.
[2025-05-04 23:25:16.724][DEBUG] Comparing pre-release versions: alpha5 not newer then alpha5
[2025-05-04 23:25:16.727][SUCCESS] Update found for 'Pester': Local '3.4.0' -> Online '6.0.0'
[2025-05-04 23:25:16.917][SUCCESS] Job 225 completed. Found 'PackageManagement' version 1.4.8.1.
[2025-05-04 23:25:16.939][SUCCESS] Update found for 'PackageManagement': Local '1.0.0.1' -> Online '1.4.8.1'
[2025-05-04 23:25:17.474][SUCCESS] Job 237 completed. Found 'PolicyFileEditor' version 3.0.1.
[2025-05-04 23:25:17.922][SUCCESS] Job 239 completed. Found 'PowerHTML' version 0.2.0.
[2025-05-04 23:25:18.510][SUCCESS] Job 241 completed. Found 'PowerQuest' version 0.4.0.
[2025-05-04 23:25:18.805][SUCCESS] Job 243 completed. Found 'PowerShellGet' version 3.0.23.
[2025-05-04 23:25:18.810][SUCCESS] Update found for 'PowerShellGet': Local '2.2.5' -> Online '3.0.23'
[2025-05-04 23:25:18.819][SUCCESS] Update found for 'PowerShellGet': Local '2.2.5' -> Online '3.0.23'
[2025-05-04 23:25:18.827][DEBUG] Comparing pre-release versions: beta23 not newer then beta23
[2025-05-04 23:25:18.829][SUCCESS] Update found for 'PowerShellGet': Local '1.0.0.1' -> Online '3.0.23'
[2025-05-04 23:25:18.845][SUCCESS] Update found for 'PowerShellGet': Local '2.2.5' -> Online '3.0.23'
[2025-05-04 23:25:18.852][DEBUG] Comparing pre-release versions: beta23 not newer then beta23
[2025-05-04 23:25:18.859][SUCCESS] Update found for 'PowerShellGet': Local '2.2.5' -> Online '3.0.23'
[2025-05-04 23:25:18.967][SUCCESS] Job 257 completed. Found 'PSDev' version 1.8.0.
[2025-05-04 23:25:19.071][SUCCESS] Job 247 completed. Found 'ProcessMitigations' version 1.0.7.
[2025-05-04 23:25:19.073][WARNING] Filtering out ProcessMitigations since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:25:19.165][SUCCESS] Job 255 completed. Found 'PSDesiredStateConfiguration' version 2.0.7.
[2025-05-04 23:25:19.167][SUCCESS] Update found for 'PSDesiredStateConfiguration': Local '1.1' -> Online '2.0.7'
[2025-05-04 23:25:19.383][SUCCESS] Job 251 completed. Found 'ps2exe' version 1.0.15.
[2025-05-04 23:25:20.126][SUCCESS] Job 263 completed. Found 'PSEventViewer' version 2.4.3.
[2025-05-04 23:25:20.650][SUCCESS] Job 261 completed. Found 'PSFramework' version 1.12.346.
[2025-05-04 23:25:20.735][SUCCESS] Job 267 completed. Found 'PSModuleDevelopment' version 2.2.13.176.
[2025-05-04 23:25:20.825][SUCCESS] Update found for 'PSModuleDevelopment': Local '2.2.12.172' -> Online '2.2.13.176'
[2025-05-04 23:25:20.918][SUCCESS] Job 269 completed. Found 'PSReadLine' version 2.4.2.
[2025-05-04 23:25:20.931][DEBUG] Comparing pre-release versions: beta2 not newer then beta2
[2025-05-04 23:25:21.054][SUCCESS] Job 265 completed. Found 'PSPreworkout' version 1.8.3.
[2025-05-04 23:25:21.180][SUCCESS] Job 271 completed. Found 'PSSharedGoods' version 0.0.307.
[2025-05-04 23:25:21.820][SUCCESS] Job 279 completed. Found 'PSWriteHTML' version 1.28.0.
[2025-05-04 23:25:22.001][SUCCESS] Job 277 completed. Found 'PSWindowsUpdate' version 2.2.1.5.
[2025-05-04 23:25:22.800][SUCCESS] Job 283 completed. Found 'Refactor' version 1.2.25.
[2025-05-04 23:25:23.195][SUCCESS] Job 287 completed. Found 'RemoteDesktop' version 0.0.0.2.
[2025-05-04 23:25:23.199][WARNING] Filtering out RemoteDesktop since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:25:23.411][SUCCESS] Job 299 completed. Found 'SharePointPnPPowerShellOnline' version 3.29.2101.0.
[2025-05-04 23:25:23.909][SUCCESS] Job 291 completed. Found 'ScheduledTasks' version 1.0.2.
[2025-05-04 23:25:23.911][WARNING] Filtering out ScheduledTasks since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:25:24.725][SUCCESS] Job 305 completed. Found 'SplitPipeline' version 2.0.0.
[2025-05-04 23:25:25.397][SUCCESS] Job 313 completed. Found 'ThreadJob' version 2.1.0.
[2025-05-04 23:25:26.426][SUCCESS] Job 319 completed. Found 'tls' version 1.0.2.
[2025-05-04 23:25:26.430][WARNING] Filtering out tls since -MatchAuthor is used and the Authors doesn't match.
[2025-05-04 23:25:27.099][SUCCESS] Job 337 completed. Found 'WindowsAutoPilotIntune' version 5.7.
[2025-05-04 23:25:54.408][SUCCESS] Completed check of 176 modules in 30,8 seconds. Found 5 updates.

GalleryAuthor       : Microsoft Corporation
PreReleaseVersion   :
HighestLocalVersion : 1.4.8.1
OutdatedModules     : @{Path=C:\Program Files\WindowsPowerShell\Modules\PackageManagement; InstalledVersion=1.0.0.1}
Author              : Microsoft Corporation
ModuleName          : PackageManagement
IsPreview           : False
LatestVersionString : 1.4.8.1
Repository          : PSGallery
LatestVersion       : 1.4.8.1

GalleryAuthor       : Microsoft Corporation
PreReleaseVersion   :
HighestLocalVersion : 1.2.5
OutdatedModules     : @{Path=C:\Windows\System32\WindowsPowerShell\v1.0\Modules\Microsoft.PowerShell.Archive; InstalledVersion=1.0.1.0}
Author              : Microsoft Corporation
ModuleName          : Microsoft.PowerShell.Archive
IsPreview           : False
LatestVersionString : 1.2.5
Repository          : PSGallery
LatestVersion       : 1.2.5

GalleryAuthor       : Microsoft Corporation
PreReleaseVersion   :
HighestLocalVersion : 2.25.0
OutdatedModules     : @{Path=C:\Windows\System32\WindowsPowerShell\v1.0\Modules\Microsoft.Graph.Beta.DeviceManagement.Actions; InstalledVersion=2.24.0}
Author              : Microsoft Corporation
ModuleName          : Microsoft.Graph.Beta.DeviceManagement.Actions
IsPreview           : False
LatestVersionString : 2.25.0
Repository          : PSGallery
LatestVersion       : 2.25.0

GalleryAuthor       : Microsoft Corporation
PreReleaseVersion   :
HighestLocalVersion : 2.25.0
OutdatedModules     : @{Path=C:\Windows\System32\WindowsPowerShell\v1.0\Modules\Microsoft.Graph.DeviceManagement.Actions; InstalledVersion=2.24.0}
Author              : Microsoft Corporation
ModuleName          : Microsoft.Graph.DeviceManagement.Actions
IsPreview           : False
LatestVersionString : 2.25.0
Repository          : PSGallery
LatestVersion       : 2.25.0

GalleryAuthor       : Friedrich Weinmann
PreReleaseVersion   :
HighestLocalVersion : 2.2.12.172
OutdatedModules     : @{Path=C:\Program Files\PowerShell\Modules\PSModuleDevelopment; InstalledVersion=2.2.12.172}
Author              : Friedrich Weinmann
ModuleName          : PSModuleDevelopment
IsPreview           : False
LatestVersionString : 2.2.13.176
Repository          : PSGallery
LatestVersion       : 2.2.13.176

[2025-05-04 23:27:20.880][INFO] Starting update process for 1 modules.
[2025-05-04 23:27:20.887][INFO] [1/1] Processing update for [PackageManagement] to version [1.4.8.1] (Preview=False) from [PSGallery].
[2025-05-04 23:27:20.896][DEBUG] Attempting update for 'PackageManagement' in base paths: C:\Program Files\WindowsPowerShell\Modules\PackageManagement
[2025-05-04 23:27:20.914][DEBUG] Attempting Save-PSResource for [PackageManagement] version [v1.4.8.1] to 'C:\Program Files\WindowsPowerShell\Modules'...
[2025-05-04 23:27:21.746][SUCCESS] Successfully saved [PackageManagement] version [v1.4.8.1] via Save-PSResource to 'C:\Program Files\WindowsPowerShell\Modules\PackageManagement'
[2025-05-04 23:27:21.753][SUCCESS] Successfully updated [PackageManagement] v1.4.8.1 for all target destinations.
[2025-05-04 23:27:21.755][INFO] Update successful for 'PackageManagement'. Proceeding with cleaning old versions...
[2025-05-04 23:27:21.758][INFO] Starting cleanup of old versions for [PackageManagement] (keeping v1.4.8.1)...
[2025-05-04 23:27:21.760][DEBUG] Checking for old versions within 'C:\Program Files\WindowsPowerShell\Modules\PackageManagement'...
[2025-05-04 23:27:21.765][DEBUG] Found old version folder: 'C:\Program Files\WindowsPowerShell\Modules\PackageManagement\1.0.0.1'. Attempting removal...
[2025-05-04 23:27:21.987][WARNING] Uninstall-PSResource ran for '1.0.0.1' but folder 'C:\Program Files\WindowsPowerShell\Modules\PackageManagement\1.0.0.1' still exists. Will attempt Remove-Item.
[2025-05-04 23:27:22.225][SUCCESS] Successfully removed old folder 'C:\Program Files\WindowsPowerShell\Modules\PackageManagement\1.0.0.1' via Remove-Item.
[2025-05-04 23:27:22.231][SUCCESS] Successfully cleaned 1 old items for 'PackageManagement'.
[2025-05-04 23:27:22.237][SUCCESS] Update process finished. Successful: 1, Failed/Partial: 0 (of 1).

ModuleName           : PackageManagement
NewVersionPreRelease : 1.4.8.1
NewVersion           : 1.4.8.1
UpdatedPaths         : C:\Program Files\WindowsPowerShell\Modules\PackageManagement
FailedPaths          :
OverallSuccess       : True
CleanedPaths         : C:\Program Files\WindowsPowerShell\Modules\PackageManagement\1.0.0.1

#>
# Assume New-Log function is defined elsewhere or via:
# Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
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
			$latestVersion = ($foundModule | Sort-Object Version -Descending | Select-Object -First 1).Version
			$isInstalled = $false
			if ($latestVersion) {
				$isInstalled = Get-module -Name $moduleName -ListAvailable | Where-Object { $_.Version -eq $latestVersion } | Select-Object -First 1
			}
			else {
				New-Log "Could not find module $moduleName in PSGallery to check installation status." -Level WARNING
			}
			if ($ForceReinstall -or !$isInstalled) {
				New-Log "Attempting to install/update module '$moduleName' ($($ForceReinstall ? 'Forced Reinstall' : ($isInstalled ? 'Update' : 'Install')))..." -Level INFO
				$commonInstallParams = @{
					Name          = $moduleName
					Scope         = 'AllUsers'
					AcceptLicense = $true
					Confirm       = $false
					PassThru      = $true
					ErrorAction   = 'SilentlyContinue'
					WarningAction = 'SilentlyContinue'
				}
				$res = $null
				if ($usePSResourceCmdlets) {
					New-Log "Using Install-PSResource for '$moduleName'." -Level INFO
					$installParams = @{
						Reinstall       = $ForceReinstall
						TrustRepository = $true
						Repository      =	@('PSGallery')
					} + $commonInstallParams
					if ($installPrerelease) { $installParams.Add('Prerelease', $true) }
					$res = Install-PSResource @installParams
				}
				if (!$res -or !$usePSResourceCmdlets) {
					if ($usePSResourceCmdlets -and !$res) {
						New-Log "Install-PSResource failed or returned no result. Trying Install-Module..." -Level WARNING
					}
					else {
						New-Log "Trying with Install-Module for '$moduleName'." -Level INFO
					}
					$installParams = @{
						Force = $ForceReinstall
					} + $commonInstallParams
					if ($installPrerelease) { $installParams.Add('AllowPrerelease', $true) }
					$res = Install-Module @installParams
				}
				if ($res) {
					New-Log "Successfully installed/updated module '$moduleName'." -Level SUCCESS
				}
				elseif (!$res -and $isInstalled) {
					New-Log "Could not force an reinstall/update of '$moduleName'. Target version [$($latestVersion)] is already installed." -Level INFO
				}
				else {
					New-Log "Could not install/update '$moduleName'" -Level WARNING
					$moduleSuccess = $false
				}
			}
			else {
				New-Log "Module '$moduleName' version [$latestVersion] is already installed and available." -Level INFO
			}
			try {
				Import-Module -Name $moduleName -Force -ErrorAction Stop
				New-Log "Successfully imported module '$moduleName'." -Level SUCCESS
			}
			catch {
				New-Log "Failed to import module '$moduleName' after check/install attempt. Error: $($_.Exception.Message)" -Level ERROR
				$moduleSuccess
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
			$needsUpdate = ($null -eq $repository) -or ($repository.Priority -ne $Priority) -or (-not $repository.Trusted)
			if (-not $IsPSGallery -and $Uri -and $repository) {
				$needsUpdate = $needsUpdate -or ($repository.Uri.AbsoluteUri -ne $Uri)
			}
			if ($needsUpdate) {
				New-Log "Registering/Updating repository '$Name' (Priority: $Priority, Trusted: True)." -Level INFO
				$commonRegisterParams = @{
					Name        = $Name
					Force       = $true
					Trusted     = $true
					Priority    = $Priority
					Confirm     = $false
					ErrorAction = 'Stop'
				}
				$registerParams = @{} + $commonRegisterParams
				if ($IsPSGallery) {
					Set-PSResourceRepository -Name $Name -Priority $Priority -InstallationPolicy Trusted -ErrorAction Stop
				}
				else {
					$registerParams.Uri = $Uri
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
			New-Log "Failed to register/configure '$Name' repository. Error: $($_.Exception.Message)" -Level ERROR
			return $false
		}
	}
	if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
		New-Log "Administrator privileges are required to manage repositories and install modules with -Scope AllUsers. Aborting." -Level WARNING
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
			New-Log "Failed to install or import required modules (PowerShellGet/Microsoft.PowerShell.PSResourceGet). Repository configuration might fail." -Level WARNING
			if (-not (Get-Command Register-PSResourceRepository -ErrorAction SilentlyContinue)) {
				New-Log "Required command 'Register-PSResourceRepository' is STILL not available after installation attempt. Aborting repository configuration." -Level ERROR
				return
			}
		}
		if (-not (Get-Command Register-PSResourceRepository -ErrorAction SilentlyContinue)) {
			New-Log "Required command 'Register-PSResourceRepository' is still not available after installation attempt. Aborting repository configuration." -Level ERROR
			return
		}
		New-Log "Required module commands are now available." -Level SUCCESS
	}
	else {
		New-Log "Required module commands (PSResourceGet) are already available." -Level INFO
	}
	$repositories = @(
		@{ Name = 'PSGallery'; Uri = 'https://www.powershellgallery.com/api/v2'; Priority = 30; IsPSGallery = $true }
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
		New-Log "All specified repositories appear to be registered and configured." -Level SUCCESS
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
		[string[]]$IgnoredModules = @(),
		[switch]$Force
	)
	function Get-ManifestVersionInfo {
		[CmdletBinding()]
		param (
			[Parameter(Mandatory = $true, ValueFromPipeline)][object]$ManifestData,
			[switch]$Quick,
			[object]$ResData,
			[string]$ModuleFilePath
		)
		begin {
			function Add-VersionCandidate {
				[CmdletBinding()]
				param(
					[string]$VersionString
				)
				if (-not [string]::IsNullOrWhiteSpace($VersionString)) {
					try {
						$baseVersionString = $VersionString -replace '-.*$' # Remove potential prerelease tag for parsing
						$versionObject = [version]$baseVersionString
						$candidateVersions.Add([PSCustomObject]@{
								VersionObject = $versionObject # Parsed [version] for comparison
								VersionString = "$VersionString" # Original full string
							})
					}
					catch {
						New-Log "Could not parse version string '$VersionString'" -Level ERROR
					}
				}
			}
		}
		process {
			if ($Quick -and $ManifestData) {
				try {
					$higestVersion = [version]$ManifestData.Version
				}
				catch {
					$higestVersion = $ManifestData.Version
				}
				return [PSCustomObject]@{
					ModuleName           = if ($ManifestData.Name) { $ManifestData.Name } else { $null }
					HighestVersion       = if ($higestVersion) { $higestVersion } else { $null } # The [version] object
					HighestVersionString = if ($ManifestData.Version) { "$($ManifestData.Version)" } else { $null } # The original string
					IsPreRelease         = if ($ManifestData.PreRelease) { $true } else { $false }
					PreReleaseLabel      = if ($ManifestData.PreRelease) { $ManifestData.PreRelease } else { $null }
					BasePath             = if ($ManifestData.ModuleBase) { $ManifestData.ModuleBase } else { $null }
				}
			}
			if (!$ManifestData.Name) {
				$moduleName = $resData.Name
			}
			else {
				$moduleName = $ManifestData.Name
			}
			$candidateVersions = [System.Collections.Generic.List[object]]::new()
			if ($ManifestData.ModuleVersion) {
				Add-VersionCandidate -VersionString $ManifestData.ModuleVersion
			}
			if ($ManifestData.Version) {
				Add-VersionCandidate -VersionString $ManifestData.Version
			}
			if ($ManifestData.PreRelease) {
				Add-VersionCandidate -VersionString $ManifestData.PreRelease
			}
			$psData = $null
			$privatePreReleaseValue = $null
			try {
				$check = $ManifestData.PrivateData['PSData'] -and $ManifestData.PrivateData -is [hashtable]
			}
			catch {
				$check = $null
			}
			if ($check) {
				if ($ManifestData.PrivateData['PSData'].PreRelease) {
					$privatePreReleaseValue = $ManifestData.PrivateData['PSData'].PreRelease
				}
				if ($ManifestData.PrivateData['PSData'].Version) {
					Add-VersionCandidate -VersionString $ManifestData.PrivateData['PSData'].Version
				}
				elseif ($ManifestData.PrivateData['PSData'].ModuleVersion) {
					Add-VersionCandidate -VersionString $ManifestData.PrivateData['PSData'].ModuleVersion
				}
			}
			if ($candidateVersions.Count -eq 0) {
				New-Log "No valid version information found in the manifest data for $moduleName" -Level VERBOSE
				return $null
			}
			$highestVersionCandidate = $candidateVersions | Sort-Object -Property VersionObject -Descending | Select-Object -First 1
			$isPreRelease = $false
			$preReleaseLabel = $null
			$highestVersionString = $highestVersionCandidate.VersionString
			if ($null -ne $privatePreReleaseValue) {
				if ($privatePreReleaseValue -is [string] -and -not [string]::IsNullOrWhiteSpace($privatePreReleaseValue)) {
					$preReleaseLabel = $privatePreReleaseValue
					$isPreRelease = $true
				}
				else {
					$isPreRelease = $false
					$preReleaseLabel = $null
				}
			}
			else {
				Write-Verbose "No explicit PSData.Prerelease key found. Checking version string '$highestVersionString' for prerelease tag."
				if ($highestVersionString -match '^(\d+(?:\.\d+){2,3})-(.+)$') {
					$prereleasePart = $Matches[2]
					$isPreRelease = $true
					$preReleaseLabel = $prereleasePart
					Write-Verbose "Prerelease status set to TRUE based on version string suffix: '$preReleaseLabel'"
				}
				else {
					Write-Verbose "No prerelease tag found in version string."
				}
			}
			try {
				$moduleNameRootModule = $null
				$moduleNameRootModule = ($ManifestData.RootModule -split '\\')[-1]
				$moduleNameRootModuleString = ($ManifestData.RootModule -split '\.')[-1]
				if ($moduleNameRootModule -match '\.([a-zA-Z0-9]{3,4})$') {
					$moduleNameRootModule = [System.IO.Path]::GetFileNameWithoutExtension($moduleNameRootModule)
				}
			}
			catch {
				$moduleNameRootModule = $null
			}
			if (!$resData.ModuleBase -or !$resData.Name) {
				$module = $null
				try {
					$module = Get-Module -Name $moduleNameRootModule -All -ListAvailable -ErrorAction Stop | Where-Object { "$($_.Version)" -eq $ManifestData.ModuleVersion -and $ModuleFilePath.StartsWith($_.ModuleBase) } | Sort-Object -Property Version -Descending -Unique
				}
				catch {
					$module = $null
				}
				if (Test-Path -Path $module.Path -ErrorAction SilentlyContinue) {
					$moduleBase = $module.ModuleBase
					$moduleName = $module.Name
				}
				else {
					try {
						$module = Get-Module -Name $moduleNameRootModuleString -All -ListAvailable -ErrorAction Stop | Where-Object { "$($_.Version)" -eq $ManifestData.ModuleVersion -and $ModuleFilePath.StartsWith($_.ModuleBase) } | Sort-Object -Property Version -Descending -Unique
					}
					catch {
						$module = $null
					}
					if (Test-Path -Path $module.Path -ErrorAction SilentlyContinue) {
						$moduleBase = $module.ModuleBase
						$moduleName = $module.Name
					}
				}
			}
			return [PSCustomObject]@{
				ModuleName           = if ($resData.Name) {	$resData.Name } elseif ( $moduleName ) { $moduleName } else { $null }
				HighestVersion       = $highestVersionCandidate.VersionObject
				HighestVersionString = $highestVersionString
				IsPreRelease         = $isPreRelease
				PreReleaseLabel      = $preReleaseLabel
				BasePath             = if ($resData.ModuleBase) { $resData.ModuleBase } elseif ( $moduleBase ) { $moduleBase } else { $null }
			}
		}
	}
	function Test-IsResourceFile {
		param([string]$Path)
		if ($Path -match '\\([a-z]{2}-[A-Z]{2})\\') {
			return $true # Check for culture/language code directory pattern (like 'en-US', 'fr-FR', etc.)
		}
		if ($Path -match '\\(Resources|Localization|Localizations|Languages|Cultures)\\') {
			return $true # Check if in any directory that might contain UI culture resources
		}
		if ($Path -match '(Resources|Strings|Localized|Messages|Text|Errors|Labels)\.(psd1|xml)$') {
			return $true # Check for common resource file naming patterns
		}
		if ($Path -match '\\([a-z]{2})\\[^\\]+\.(psd1|xml)$') {
			return $true # Check for alternative culture code format (like 'en', 'fr', etc.)
		}
		return $false
	}
	$allFoundModules = [System.Collections.Generic.List[object]]::new()
	$processedManifests = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
	$processedModuleRoots = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
	foreach ($startPath in $Paths) {
		if (-not (Test-Path $startPath -PathType Container)) {
			New-Log "Input path '$startPath' is not a valid directory. Skipping." -Level WARNING
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
			$moduleRootNameGuess = $moduleDir.Name
			if (-not $processedModuleRoots.Add($moduleRootPath)) {
				continue
			}
			$psd1Files = Get-ChildItem -Path $moduleRootPath -Recurse -Filter *.psd1 -File -ErrorAction SilentlyContinue
			$xmlFiles = Get-ChildItem -Path $moduleRootPath -Recurse -Filter PSGetModuleInfo.xml -File -ErrorAction SilentlyContinue
			foreach ($file in $psd1Files) {
				$filePath = $file.FullName
				if (-not $processedManifests.Add($filePath)) {
					continue
				}
				if (Test-IsResourceFile -Path $filePath) {
					continue
				}
				try {
					$manifestInfo, $manifestInfo1, $manifestInfo2, $res1, $res2 = $null
					try {
						$res1 = Test-ModuleManifest -Path $filePath -ErrorAction Stop -WarningAction SilentlyContinue
					}
					catch {
						$res1 = $null
					}
					if ($res1) {
						$manifestInfo = Get-ManifestVersionInfo -ManifestData $res1 -Quick -ErrorAction Stop -WarningAction SilentlyContinue # Use SilentlyContinue as it might be invalid/resource
						if ($manifestInfo) {
							$manifestInfo1 = [psCustomObject]@{
								ModuleName           = $manifestInfo.ModuleName
								ModuleVersion        = $manifestInfo.HighestVersion
								HighestVersionString = $manifestInfo.HighestVersionString
								isPreRelease         = $manifestInfo.IsPreRelease
								PreReleaseLabel      = $manifestInfo.PreReleaseLabel
								BasePath             = if ($manifestInfo.BasePath.EndsWith($manifestInfo.ModuleName)) { "$($manifestInfo.BasePath)" } elseif ($manifestInfo.BasePath.EndsWith("Modules")) { "$(Join-Path ($manifestInfo.BasePath) ($manifestInfo.ModuleName) -Erroraction SilentlyContinue)" } else { "$(Split-Path $manifestInfo.BasePath -Parent -ErrorAction SilentlyContinue)" }
								Author               = if ($res1.Author) { $res1.Author } else { $null }
							}
						}
					}
				}
				catch {
					New-Log "Test-ModuleManifest failed for '$filePath'.. Will try other methods." -Level ERROR
				}
				try {
					$res2 = Import-PowerShellDataFile -Path $filePath -ErrorAction Stop -WarningAction SilentlyContinue
				}
				catch {
					$res2 = $null
				}
				if ($res1 -and $res2 ) {
					try {
						$manifestInfo = Get-ManifestVersionInfo -ManifestData $res2 -ResData $res1 -ModuleFilePath $filePath -ErrorAction Stop -WarningAction SilentlyContinue
						$manifestInfo = $manifestInfo | Where-Object { $_ -ne $null }
						$manifestInfo2 = [psCustomObject]@{
							ModuleName           = $manifestInfo.ModuleName
							ModuleVersion        = $manifestInfo.HighestVersion
							HighestVersionString = $manifestInfo.HighestVersionString
							isPreRelease         = $manifestInfo.IsPreRelease
							PreReleaseLabel      = $manifestInfo.PreReleaseLabel
							BasePath             = if ($manifestInfo.BasePath.EndsWith($manifestInfo.ModuleName)) { $($manifestInfo.BasePath) } elseif ($manifestInfo.BasePath.EndsWith("Modules")) { "$(Join-Path ($manifestInfo.BasePath) ($manifestInfo.ModuleName) -Erroraction SilentlyContinue)" } else { "$(Split-Path $manifestInfo.BasePath -Parent -ErrorAction SilentlyContinue)" }
							Author               = if ($res2.Author) { $res2.Author } else { $null }
						}
					}
					catch {
						New-Log "Failed to get manifest." -Level ERROR
					}
				}
				elseif ($res2) {
					try {
						$manifestInfo = Get-ManifestVersionInfo -ManifestData $res2 -ModuleFilePath $filePath -ErrorAction Stop -WarningAction SilentlyContinue
						$manifestInfo = $manifestInfo | Where-Object { $_ -ne $null }
						if ($manifestInfo) {
							$manifestInfo2 = [psCustomObject]@{
								ModuleName           = $manifestInfo.ModuleName
								ModuleVersion        = $manifestInfo.HighestVersion
								HighestVersionString = $manifestInfo.HighestVersionString
								isPreRelease         = $manifestInfo.IsPreRelease
								PreReleaseLabel      = $manifestInfo.PreReleaseLabel
								BasePath             = if ($manifestInfo.BasePath.EndsWith($manifestInfo.ModuleName)) { "$($manifestInfo.BasePath)" } elseif ($manifestInfo.BasePath.EndsWith("Modules")) { "$(Join-Path ($manifestInfo.BasePath) ($manifestInfo.ModuleName) -Erroraction SilentlyContinue)" } else { "$(Split-Path $manifestInfo.BasePath -Parent -ErrorAction SilentlyContinue)" }
								Author               = if ($res2.Author) { $res2.Author } else { $null }
							}
						}
					}
					catch { }
				}
				$manifestInfo = @()
				if ($manifestInfo1) {
					$manifestInfo += $manifestInfo1
					Write-Verbose "Added manifestInfo1 to merged array."
				}
				if (![string]::IsNullOrEmpty($manifestInfo2.ModuleName) -and ($manifestInfo2.BasePath -ne $manifestInfo1.BasePath -or $manifestInfo2.ModuleName -ne $manifestInfo1.ModuleName -or $manifestInfo2.HighestVersionString -ne $manifestInfo1.HighestVersionString)) {
					$manifestInfo += $manifestInfo2
					Write-Verbose "Added manifestInfo2 to merged array."
				}
				foreach ($mInfo in $manifestInfo) {
					if ($mInfo -and $mInfo.ModuleVersion -and $mInfo.ModuleName) {
						$allFoundModules.Add([PSCustomObject]@{
								ModuleName           = $mInfo.ModuleName
								ModuleVersion        = $mInfo.ModuleVersion # Store as [version] object
								HighestVersionString = $mInfo.HighestVersionString
								BasePath             = if ($mInfo.BasePath) { $minfo.BasePath }  else { $null }
								isPreRelease         = $mInfo.isPreRelease
								PreReleaseLabel      = $mInfo.PreReleaseLabel
								Author               = if ($mInfo.Author) { $mInfo.Author } else { $null }
							})
					}
					else {
						New-Log "mInfo doesnt contain all required parameters: $mInfo" -Level DEBUG
					}
				}
			}
			foreach ($xmlFile in $xmlFiles) {
				$filePath = $xmlFile.FullName
				if (-not $processedManifests.Add($filePath)) {
					continue
				}
				$xmlInfo = Get-ModuleInfoFromXml -XmlFilePath $filePath -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				$pathInfo = Get-ModuleformPath -Path $xmlFile.DirectoryName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				if ($xmlInfo -and $xmlInfo.ModuleName -and $xmlInfo.ModuleVersion -and ($xmlInfo.BasePath -or $pathInfo.BasePath)) {
					$allFoundModules.Add([PSCustomObject]@{
							ModuleName      = $xmlInfo.ModuleName
							ModuleVersion   = $xmlInfo.ModuleVersion
							BasePath        = if ($xmlInfo.BasePath) { $xmlInfo.BasePath } elseif ($pathInfo.BasePath) { $pathInfo.BasePath } else { $null }
							isPreRelease    = $xmlInfo.isPreRelease
							PreReleaseLabel = $xmlInfo.PreReleaseLabel
							Author          = if ($xmlInfo.Author) { $xmlInfo.Author } else { $null }
						})
				}
				else {
					New-Log "xmlInfo doesnt contain all required parameters: $xmlInfo and $pathInfo" -Level DEBUG
				}
			}
		}
	}
	$uniqueModules = $allFoundModules | Group-Object -Property ModuleName, ModuleVersion, BasePath | ForEach-Object { $_.Group[0] }
	$allFoundModules = [System.Collections.Generic.List[object]]::new($uniqueModules)
	$resultModules = @{}
	if ($allFoundModules.Count -eq 0) {
		New-Log "No module files found in the specified paths." -Level WARNING
		return $resultModules
	}
	$resultModules = [ordered]@{}
	$modulesGroupedByName = $allFoundModules | Where-Object { $null -ne $_.ModuleName } | Group-Object ModuleName
	foreach ($nameGroup in $modulesGroupedByName) {
		$moduleName = $nameGroup.Name
		Write-Verbose "Processing ModuleName: $moduleName"
		$groupWithNormalizedPaths = $nameGroup.Group | Where-Object { $null -ne $_ } | ForEach-Object {
			$newObject = $_ | Select-Object *
			if ($null -ne $newObject.BasePath -and -not ([string]::IsNullOrWhiteSpace($moduleName))) {
				$currentBasePath = $newObject.BasePath
				$normalizedCurrentPath = $currentBasePath.TrimEnd('\', '/') -replace '/', '\'
				$expectedEnding = "\$moduleName"
				if (-not $normalizedCurrentPath.EndsWith($expectedEnding, [System.StringComparison]::OrdinalIgnoreCase)) {
					$newObject.BasePath = Join-Path -Path $normalizedCurrentPath -ChildPath $moduleName
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
			Write-Verbose " Processing Grouped BasePath: $currentBasePath (Count: $($basePathGroup.Count))"
			$versionsInPathGroup = $basePathGroup.Group | Group-Object -Property ModuleVersion
			foreach ($versionGroup in $versionsInPathGroup) {
				$representativeEntry = $versionGroup.Group | Select-Object -First 1
				if ($representativeEntry) {
					$outputObject = [PSCustomObject]@{
						ModuleVersion       = $representativeEntry.ModuleVersion
						ModuleVersionString = $representativeEntry.ModuleVersion.ToString()
						BasePath            = $currentBasePath
						IsPreRelease        = $representativeEntry.IsPreRelease
						PreReleaseLabel     = $representativeEntry.PreReleaseLabel
						Author              = $representativeEntry.Author
					}
					$finalModuleLocations.Add($outputObject)
					Write-Verbose "  Added unique Version: $($outputObject.ModuleVersionString) for BasePath: $currentBasePath"
				}
				else {
					New-Log "Not a representative entry: $currentBasePath" -Level WARNING
				}
			}
		}
		if ($finalModuleLocations.Count -gt 0) {
			$sortedLocations = $finalModuleLocations | Sort-Object BasePath, ModuleVersion
			$resultModules[$moduleName] = $sortedLocations
			New-Log "Finished processing $moduleName. Found $($finalModuleLocations.Count) unique BasePath/Version combinations." -Level DEBUG
		}
		else {
			$finalModuleLocations
			New-Log "FinalModuleLocations count was 0 on module: $moduleName" -Level WARNING
		}
	}
	$sortedModules = [ordered]@{}
	foreach ($key in ($resultModules.Keys | Sort-Object)) {
		if ($key -in $IgnoredModules) {
			New-Log "Skipping $key since it's on the ignorelist." -Level WARNING
			continue
		}
		if ($key -notmatch '^\d+(\.\d+)+$') {
			$sortedModules[$key] = $resultModules[$key]
		}
		else {
			New-Log "Skipping potential module entry with an invalid name (looks like a version): $key" -Level WARNING
		}
	}
	Return $sortedModules
}
function Get-ModuleformPath {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)][string]$Path
	)
	$normalizedPath = $Path -replace '/', '\'
	if ($normalizedPath -match '\\Modules\\([^\\]+)(?:\\|$)') {
		$moduleName = $Matches[1]
		$searchPattern = '^(.*\\Modules\\' + [regex]::Escape($moduleName) + ')(?:\\.*)?$'
		$modulePath = $normalizedPath -replace $searchPattern, '$1'
		return [PSCustomObject]@{
			ModuleName = $moduleName
			ModulePath = $modulePath
		}
	}
	$leaf = Split-Path -Path $normalizedPath -Leaf
	$parent = Split-Path -Path $normalizedPath -Parent
	if ($leaf -eq 'Modules' -and $parent) {
		$parentLeaf = Split-Path -Path $parent -Leaf
		$parentParent = Split-Path -Path $parent -Parent
		if ($parentLeaf -and $parentParent) {
			return [PSCustomObject]@{
				ModuleName = $parentLeaf
				ModulePath = $parent
			}
		}
	}
	elseif ($leaf -and $parent) {
		return [PSCustomObject]@{
			ModuleName = $leaf
			ModulePath = $parent
		}
	}
	else {
		return [PSCustomObject]@{
			ModuleName = $normalizedPath
			ModulePath = $null
		}
	}
}
function Get-ModuleInfoFromXml {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)][string]$XmlFilePath
	)
	if (-not (Test-Path -Path $XmlFilePath -PathType Leaf)) {
		New-Log "XML file not found: $XmlFilePath" -Level WARNING
		return $null
	}
	try {
		[xml]$xmlContent = Get-Content -Path $XmlFilePath -Raw
		$nsManager = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
		$nsManager.AddNamespace("ps", "http://schemas.microsoft.com/powershell/2004/04")
		$nameNode = ($xmlContent.SelectSingleNode("//ps:S[@N='Name']", $nsManager)).'#text'
		$authorNode = ($xmlContent.SelectSingleNode("//ps:S[@N='Author']", $nsManager)).'#text'
		$versionNode = ($xmlContent.SelectSingleNode("//ps:S[@N='Version']", $nsManager)).'#text'
		$locationNode = ($xmlContent.SelectSingleNode("//ps:S[@N='InstalledLocation']", $nsManager)).'#text'
		$prerelease = ($xmlContent.SelectSingleNode("//ps:B[@N='IsPrerelease']", $nsManager)).'#text'
		$prerelease2 = ($xmlContent.SelectSingleNode("//ps:S[@N='IsPrerelease']", $nsManager)).'#text'
		$normalizedVersion = ($xmlContent.SelectSingleNode("//ps:S[@N='NormalizedVersion']", $nsManager)).'#text'
		[bool]$IsPrerelease = $false
		if ($prerelease -ieq 'true' -or $prerelease2 -ieq 'true') {
			$IsPrerelease = $true
		}
		if ($nameNode -and $versionNode) {
			$result = [PSCustomObject]@{
				ModuleName      = $nameNode
				ModuleVersion   = $versionNode
				BasePath        = if ($locationNode) { "$(Join-Path -Path $locationNode -ChildPath $nameNode)" } else { $null }
				isPreRelease    = $IsPrerelease
				PreReleaseLabel = if ($IsPrerelease -and $normalizedVersion) { $normalizedVersion } elseif ($IsPrerelease -and $versionNode) { $versionNode } else { $null }
				Author          = if ($authorNode) { $authorNode } else { $null }
			}
			return $result
		}
		else {
			New-Log "Could not find Name or Version node in XML: $XmlFilePath" -Level VERBOSE
			return $null
		}
	}
	catch {
		New-Log "Error parsing XML file '$XmlFilePath': $($_.Exception.Message)" -Level WARNING
		return $null
	}
}
function Get-ModuleUpdateStatus {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][hashtable]$ModuleInventory,
		[string[]]$Repositories = @('PSGallery', 'NuGet'),
		[int]$ThrottleLimit = [Environment]::ProcessorCount * 2,
		[ValidateRange(1, 300)][int]$TimeoutSeconds = 30,
		[hashtable]$BlackList = @{},
		[switch]$MatchAuthor
	)
	if ($PSVersionTable.PSVersion.Major -lt 7) {
		New-Log "This function requires PowerShell 7 or later. Current version: $($PSVersionTable.PSVersion)" -Level ERROR
		return
	}
	function Parse-ModuleVersion {
		param (
			[string]$VersionString
		)
		$version = $null
		$isParseable = [version]::TryParse($VersionString, [ref]$version)
		if ($isParseable) {
			return [PSCustomObject]@{
				ModuleVersionStringNoPrefix = $VersionString
				ModuleVersionString         = $VersionString
				ModuleVersion               = [version]$version
				IsSemVer                    = $false
				IsPrerelease                = $false
				PreReleaseLabel             = $null
			}
		}
		if ($VersionString -match '^(\d+(?:\.\d+){2,3})-(.+)$') {
			[string]$versionPart = $Matches[1]
			[string]$prereleasePart = $Matches[2]
			if ([version]::TryParse($versionPart, [ref]$version)) {
				return [PSCustomObject]@{
					ModuleVersionStringNoPrefix = $versionPart
					ModuleVersionString         = [string]$VersionString
					ModuleVersion               = [version]$version
					IsSemVer                    = $true
					IsPrerelease                = $true
					PreReleaseLabel             = $prereleasePart
				}
			}
		}
		if ($VersionString -match '^(\d+(?:\.\d+){2,3})') {
			[string]$versionPart = $Matches[0]
			if ([version]::TryParse($versionPart, [ref]$version)) {
				return [PSCustomObject]@{
					ModuleVersionStringNoPrefix = $VersionString
					ModuleVersionString         = $VersionString
					ModuleVersion               = [version]$version
					IsSemVer                    = $false
					IsPrerelease                = $false
					PreReleaseLabel             = $null
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
			'dev'     = 0
			'alpha'   = 1
			'beta'    = 2
			'preview' = 3
			'rc'      = 4
		}
		if ($VersionA.gettype().Name -ne 'String' -or $VersionA.Count -gt 1) {
			$VersionA = $VersionA[0].ToString()
		}
		if ($VersionB.gettype().Name -ne 'String' -or $VersionB.Count -gt 1) {
			$VersionB = $VersionB[0].ToString()
		}
		$preReleaseA = ($VersionA -split '-', 2)[-1]
		$preReleaseB = ($VersionB -split '-', 2)[-1]
		if (-not $preReleaseA -or -not $preReleaseB) {
			Write-Verbose "Could not extract pre-release identifier from '$VersionA' or '$VersionB'."
			if ($ReturnBoolean) { return $false } else { return $VersionA }
		}
		$preReleaseA = $preReleaseA.ToLower()
		$preReleaseB = $preReleaseB.ToLower()
		$regex = "^(dev|alpha|beta|preview|rc)(\d*)$"
		$matchA = [regex]::Match($preReleaseA, $regex)
		$matchB = [regex]::Match($preReleaseB, $regex)
		if (-not $matchA.Success -or -not $matchB.Success) {
			Write-Verbose "Invalid prerelease format found. Comparing '$preReleaseA' and '$preReleaseB'. Expected format like 'dev','alpha1', 'beta2', 'preview3' or 'rc'."
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
	$allModuleNames = $ModuleInventory.Keys | Sort-Object -Unique | Where-Object { $_ -and $_.Trim() }
	$totalModules = $allModuleNames.Count
	if ($totalModules -eq 0) {
		New-Log "Module inventory is empty. Nothing to check." -Level INFO
		return @()
	}
	$results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
	New-Log "Starting parallel module update check for $totalModules modules (Throttle: $ThrottleLimit, Timeout: ${TimeoutSeconds}s)..." -Level SUCCESS
	$moduleDataArray = @()
	foreach ($moduleName in $allModuleNames) {
		$localModules = $ModuleInventory[$moduleName]
		if ($localModules -isnot [array]) {
			$localModules = @($localModules)
		}
		$res = @()
		foreach ($module in $localModules) {
			$moduleVersionCheck = Parse-ModuleVersion -VersionString $module.ModuleVersion
			$res += [PSCustomObject]@{
				ModuleVersion               = $moduleVersionCheck.ModuleVersion
				ModuleVersionString         = $moduleVersionCheck.ModuleVersionString
				ModuleVersionStringNoPrefix = $moduleVersionCheck.ModuleVersionStringNoPrefix
				PreReleaseLabel             = $moduleVersionCheck.PreReleaseLabel
				BasePath                    = $module.BasePath
				isPreRelease                = $module.IsPreRelease -or $moduleVersionCheck.IsPrerelease
				Author                      = $module.Author
			}
		}
		if ($res.Count -gt 0) {
			$preview = $res | Where-Object { $_.isPreRelease -eq $true }
			if ($preview) {
				$highestLocalVersion = $preview | Sort-Object -Property { $_.ModuleVersion } -Descending | Select-Object -First 1
				$res = $res | Where-Object { # Keep items that have different BasePath OR different ModuleVersionStringNoPrefix. Only keep one version of the same module+path+version, preferring the prerelease one
					$_.BasePath -ne $highestLocalVersion.BasePath -or $_.ModuleVersionStringNoPrefix -ne $highestLocalVersion.ModuleVersionStringNoPrefix -or ($_.BasePath -eq $highestLocalVersion.BasePath -and $_.ModuleVersionStringNoPrefix -eq $highestLocalVersion.ModuleVersionStringNoPrefix -and $_.ModuleVersionString -eq $highestLocalVersion.ModuleVersionString)
				}
			}
			else {
				$highestLocalVersion = $res | Sort-Object -Property { $_.ModuleVersion } -Descending | Select-Object -First 1
			}
			$moduleDataArray += [PSCustomObject]@{
				ModuleName                  = $moduleName
				HighestLocalVersion         = $highestLocalVersion.ModuleVersion
				HighestLocalVersionString   = $highestLocalVersion.ModuleVersionString
				HighestLocalVersionNoPrefix = $highestLocalVersion.ModuleVersionStringNoPrefix
				AllVersions                 = $res
			}
		}
	}
	$validModuleCount = $moduleDataArray.Count
	New-Log "Prepared $validModuleCount modules with valid version info for checking." -Level DEBUG
	$counters = [hashtable]::Synchronized(@{ Processed = 0; Updates = 0; Errors = 0; Timeouts = 0; StartTime = Get-Date; Total = $validModuleCount })
	$NewLogDef = ${function:New-Log}.ToString()
	$ComparePrereleaseVersionDef = ${function:Compare-PrereleaseVersion}.ToString()
	$job = $null
	$moduleDataArray | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
		$job = $null
		try {
			$moduleData = $_
			$moduleName = $moduleData.ModuleName
			$counters = $using:counters
			$matchAuthor = $using:MatchAuthor.IsPresent
			$results = $using:results
			$blackList = $using:BlackList
			$repositories = $using:Repositories
			$timeoutSeconds = $using:TimeoutSeconds
			${function:New-Log} = $using:NewLogDef
			${function:Compare-PrereleaseVersion} = $using:ComparePrereleaseVersionDef
			$isBlacklisted = $false
			$repositoriesToCheck = $repositories
			if ($blacklist -and $blackList.ContainsKey($moduleName)) {
				$blacklistedRepos = $blackList[$moduleName]
				if ($blacklistedRepos -eq '*') { $isBlacklisted = $true }
				elseif ($blacklistedRepos -is [array]) { $repositoriesToCheck = $repositories | Where-Object { $blacklistedRepos -notcontains $_ }; if ($repositoriesToCheck.Count -eq 0) { $isBlacklisted = $true } }
				elseif ($blacklistedRepos -is [string]) { $repositoriesToCheck = $repositories | Where-Object { $_ -ne $blacklistedRepos }; if ($repositoriesToCheck.Count -eq 0) { $isBlacklisted = $true } }
				else { $isBlacklisted = $true }
			}
			$galleryModule = $null
			if (-not $isBlacklisted -and $repositoriesToCheck.Count -gt 0) {
				try {
					$jobScriptBlock = {
						param($modName, $repos)
						Find-Module -Name $modName -AllowPrerelease -Repository $repos -ErrorAction SilentlyContinue | Sort-Object -Property @{E = { [version]$_.Version }; Descending = $true } | Select-Object -First 1
					}
					$job = Start-Job -ScriptBlock $jobScriptBlock -ArgumentList $moduleName, $repositoriesToCheck
					New-Log "Started job $($job.Id) to find module '$moduleName' in $($repositoriesToCheck -join ', ')" -Level VERBOSE
					$waitResult = $job | Wait-Job -Timeout $timeoutSeconds -ErrorAction SilentlyContinue
					if ($job.State -eq 'Running') {
						New-Log "Timeout ($($timeoutSeconds)s) waiting for Find-Module job for '$moduleName'." -Level WARNING
						$counters.Timeouts++
						try { $job | Stop-Job -ErrorAction SilentlyContinue } catch {}
					}
					elseif ($job.State -eq 'Completed') {
						$galleryModule = $job | Receive-Job
						if ($galleryModule) { New-Log "Job $($job.Id) completed. Found '$moduleName' version $($galleryModule.Version)." -Level SUCCESS }
						else { New-Log "Job $($job.Id) completed. Module '$moduleName' not found in specified repositories." -Level VERBOSE }
					}
					else {
						$reason = ($job | Select-Object -ExpandProperty JobStateInfo).Reason
						New-Log "Find-Module job $($job.Id) for '$moduleName' failed or ended unexpectedly. State: $($job.State). Reason: $reason" -Level WARNING
						$counters.Errors++
					}
				}
				catch {
					New-Log "Error managing Find-Module job for '$moduleName'" -Level ERROR
					$counters.Errors++
				}
				if ($galleryModule) {
					try {
						$modueResults = @()
						$filtered = @()
						foreach ($currentModule in $moduleData.AllVersions) {
							if ($matchAuthor -and $galleryModule.Author -notmatch $currentModule.Author) {
								if($moduleName -notin $filtered){
									New-Log "Filtering out '$moduleName' since -MatchAuthor is used and the Authors doesn't match." -Level WARNING
								}
								$filtered += $moduleName
								continue
							}
							$needsUpdate = $false
							[version]$highestLocal = $currentModule.ModuleVersion
							[string]$highestLocalStr = $currentModule.ModuleVersionString
							[version]$latestOnline = $galleryModule.Version
							[string]$latestOnlineStr = $($galleryModule.Version)
							if ($latestOnline -gt $highestLocal) {
								$needsUpdate = $true
							}
							if (-not $needsUpdate) {
								[string]$onlinePreLabel = $($galleryModule.Prerelease)
								[string]$localPreLabel = $currentModule.PreReleaseLabel
								if ($localPreLabel -and $onlinePreLabel) {
									$preRelease = $false
									$needsUpdate = Compare-PrereleaseVersion -VersionA $localPreLabel -VersionB $onlinePreLabel -ReturnBoolean
									if ($needsUpdate) {
										New-Log "Online pre-release '$onlinePreLabel' is newer than local '$localPreLabel' for version $latestOnlineStr." -Level SUCCESS
										$preRelease = $true
									}
									else {
										New-Log "Comparing pre-release versions: $onlinePreLabel not newer then $localPreLabel" -Level DEBUG
									}
								}
							}
							if ($needsUpdate) {
								New-Log "Update found for '$moduleName': Local '$highestLocalStr' -> Online '$latestOnlineStr'" -Level SUCCESS
								$modulesByPath = $moduleData.AllVersions | Group-Object -Property BasePath
								$outdatedPaths = @()
								foreach ($pathGroup in $modulesByPath) {
									$hasUpToDateVersion = $pathGroup.Group | Where-Object {	$_.ModuleVersion -eq $latestOnline -or ($_.ModuleVersion -eq $latestOnline -and $_.PreReleaseLabel -eq $galleryModule.Prerelease) } | Select-Object -First 1 # Check if any module in this path is up-to-date
									if (-not $hasUpToDateVersion) {
										$outdatedPaths += $pathGroup.Name # If no up-to-date version exists in this path, add it to outdated paths
									}
								}
								if ($outdatedPaths) {
									$modueResults += ([PSCustomObject]@{
											ModuleName          = $moduleName
											HighestLocalVersion = $moduleData.HighestLocalVersion
											LatestVersion       = $latestOnline
											LatestVersionString = $latestOnlineStr
											LocalVersion        = $highestLocal
											LocalVersionString  = $highestLocalStr
											IsPreview           = [bool]$galleryModule.Prerelease
											Repository          = $galleryModule.Repository
											OutdatedPaths       = $outdatedPaths
											PreReleaseVersion   = $galleryModule.Prerelease
											GalleryData         = $galleryModule
											LocalData           = $moduleData
											Author              = $currentModule.Author
											GalleryAuthor       = $galleryModule.Author
										})
								}
							}
						}
						if ($modueResults.Count -gt 0) {
							$results.Add($modueResults)
							$counters.Updates++
						}
					}
					catch {
						New-Log "Error comparing versions for '$moduleName'" -Level ERROR
						$counters.Errors++
					}
				}
			}
			elseif ($isBlacklisted) {
				New-Log "Skipping check for '$moduleName' due to blacklist." -Level WARNING
			}
			else {
				New-Log "Skipping check for '$moduleName' as no repositories remained after blacklist filter." -Level WARNING
			}
		}
		catch {
			New-Log "Unhandled error during processing for module '$moduleName'" -Level ERROR
			$counters.Errors++
		}
		finally {
			if ($job -ne $null -and $job.State -ne 'Removed') {
				New-Log "Performing final cleanup for job $($job.Id) (State: $($job.State))..." -Level VERBOSE
				try {
					$job | Remove-Job -Force -ErrorAction Stop
					New-Log "Job $($job.Id) removed." -Level VERBOSE
				}
				catch {
					New-Log "Error during final removal of job '$($job.Id)'" -Level ERROR
				}
			}
			$currentCount = [System.Threading.Interlocked]::Increment([ref]$counters.Processed)
			if ($currentCount % 20 -eq 0 -or $currentCount -eq $counters.Total) {
				try { ${function:New-Log} = $using:NewLogDef } catch {}
				$elapsed = (Get-Date) - $counters.StartTime
				$percentComplete = [Math]::Round(($currentCount / $counters.Total) * 100, 0)
				$eta = "N/A"
				if ($elapsed.TotalSeconds -gt 5 -and $currentCount -gt 0) {
					$avgTimePerModule = $elapsed.TotalSeconds / $currentCount
					$remainingModules = $counters.Total - $currentCount
					$remainingSeconds = $avgTimePerModule * $remainingModules
					if ($remainingSeconds -lt 0) { $remainingSeconds = 0 }
					$eta = "{0:mm\:ss}" -f [timespan]::FromSeconds($remainingSeconds)
				}
				$progressMsg = "[$percentComplete% ($currentCount/$($counters.Total))] Updates: $($counters.Updates), Timeouts: $($counters.Timeouts), Errors: $($counters.Errors), ETA: $eta"
				New-Log $progressMsg -Level SUCCESS
			}
		}
	}
	$moduleSummary = @($results)
	$totalTime = (Get-Date) - $counters.StartTime
	$secondsElapsed = $totalTime.TotalSeconds.ToString('F1')
	New-Log "Completed check of $validModuleCount modules in $secondsElapsed seconds. Found $($counters.Updates) updates." -Level SUCCESS
	if ($counters.Timeouts -gt 0) { New-Log "$($counters.Timeouts) module checks timed out." -Level WARNING }
	if ($counters.Errors -gt 0) { New-Log "Encountered $($counters.Errors) errors during processing." -Level WARNING }
	if ($moduleSummary) {
		$finalResults = @{}
		foreach ($module in $moduleSummary) {
			$moduleName = ($module.ModuleName | Select-Object -Unique)
			if ($finalResults.ContainsKey($moduleName)) {
				continue
			}
			$outdatedModulesInfo = $module | Select-Object @{Name = 'Path'; Expression = { $_.OutdatedPaths } }, @{Name = 'InstalledVersion'; Expression = { $_.LocalVersionString } } | Sort-Object Path, InstalledVersion -Unique
			$moduleValue = [PSCustomObject]@{
				Repository          = ($module.Repository | Select-Object -Unique)
				IsPreview           = $module.IsPreview | Select-Object -Unique
				PreReleaseVersion   = ($module.PreReleaseVersion | Select-Object -Unique)
				HighestLocalVersion = ($module.HighestLocalVersion | Select-Object -Unique)
				LatestVersion       = ($module.LatestVersion | Select-Object -Unique)
				LatestVersionString = ($module.LatestVersionString | Select-Object -Unique)
				OutdatedModules     = $outdatedModulesInfo
				Author              = ($module.Author | Select-Object -First 1 -Unique)
				GalleryAuthor       = ($module.GalleryAuthor | Select-Object -First 1 -Unique)
			}
			if (-not $finalResults.ContainsKey($moduleName)) {
				$finalResults[$moduleName] = $moduleValue
			}
		}
		$moduleObjects = $finalResults.GetEnumerator() | ForEach-Object {
			$moduleName = $_.Key
			$moduleData = $_.Value
			$properties = @{
				ModuleName = $moduleName
			}
			foreach ($property in $moduleData.PSObject.Properties) {
				$properties[$property.Name] = $property.Value
			}
			[PSCustomObject]$properties
		}
	}
	return $moduleObjects
}
function Update-Modules {
	[CmdletBinding()]
	param (
		[Parameter(ValueFromPipeline)][Object[]]$OutdatedModules,
		[switch]$Clean,
		[switch]$UseProgressBar,
		[switch]$PreRelease
	)
	begin {
		$aggregateResults = [System.Collections.Generic.List[object]]::new()
		$batchModules = @()
	}
	process {
		foreach ($module in $OutdatedModules) {
			$batchModules += $module
		}
	}
	end {
		if ($batchModules.Count -eq 0) {
			New-Log "No modules provided for update. Aborting." -Level INFO
			return
		}
		$total = $batchModules.Count
		New-Log "Starting update process for $total modules." -Level INFO
		$current = 0
		foreach ($module in $batchModules) {
			$current++
			$moduleName = $module.ModuleName
			[version]$latestVer = $module.LatestVersion
			[string]$latestVerStr = $module.LatestVersionString
			[string]$preReleaseVersion = $module.PreReleaseVersion
			$outdatedVersions = @($module.OutdatedModules.InstalledVersion | Select-Object -Unique)
			$repository = $module.Repository
			$isPreview = $module.IsPreview -or ($PreRelease.IsPresent -and $module.IsPreview)
			$outdatedPaths = @($module.OutdatedModules.Path) # Paths where OLD versions exist
			if ($UseProgressBar.IsPresent) {
				$progressParams = @{
					Activity         = "Updating PowerShell Modules"
					Status           = "[$current/$total] Updating $moduleName to $latestVerStr"
					PercentComplete  = (($current / $total) * 100)
					CurrentOperation = "Calling Install-PSModule for $moduleName"
				}
				Write-Progress @progressParams
			}
			New-Log "[$current/$total] Processing update for [$moduleName] to version [$latestVerStr] (Preview=$isPreview) from [$repository]." -Level INFO
			$uniqueBasePaths = $outdatedPaths | ForEach-Object {
				$currentPath = $_
				if ($currentPath -match '^(.*\\Modules\\[^\\]+).*') {
					$Matches[1]
				}
				else {
					Split-Path -Path $currentPath -Parent
				}
			} | Select-Object -Unique | Where-Object { $_ -and $_.Trim() }
			$installResult = $null
			if ($uniqueBasePaths.Count -gt 0) {
				New-Log "Attempting update for '$moduleName' in base paths: $($uniqueBasePaths -join '; ')" -Level DEBUG
				$installResult = Install-PSModule -ModuleName $moduleName -TargetVersion $latestVer -RepositoryName $repository -IsPreview $isPreview -PreReleaseVersion $preReleaseVersion -Destinations $uniqueBasePaths -ErrorAction SilentlyContinue
			}
			else {
				New-Log "Could not determine valid unique base installation paths for module [$moduleName] from OutdatedPaths: $($outdatedPaths -join ', '). Skipping installation." -Level WARNING
				$installResult = [PSCustomObject]@{
					UpdatedPaths = @()
					FailedPaths  = @("Path determination failed")
				}
			}
			$finalResult = [PSCustomObject]@{
				ModuleName           = $moduleName
				NewVersionPreRelease = if ($PreReleaseVersion -and $IsPreview) { "$latestVerStr-$preReleaseVersion" } else { $latestVerStr }
				NewVersion           = $latestVerStr
				UpdatedPaths         = $installResult.UpdatedPaths
				FailedPaths          = $installResult.FailedPaths
				OverallSuccess       = ($installResult.FailedPaths.Count -eq 0 -and $installResult.UpdatedPaths.Count -gt 0)
			}
			if ($finalResult.OverallSuccess -and $Clean.IsPresent) {
				if ($UseProgressBar.IsPresent) {
					$progressParams.CurrentOperation = "Cleaning old versions of $moduleName"
					Write-Progress @progressParams
				}
				New-Log "Update successful for '$moduleName'. Proceeding with cleaning old versions..." -Level INFO
				$cleanedPaths = Remove-OutdatedVersions -ModuleName $moduleName -OutdatedVersions $outdatedVersions -ModuleBasePaths $finalResult.UpdatedPaths -LatestVersion $latestVer -PreReleaseTag $preReleaseVersion -ErrorAction SilentlyContinue
				if ($cleanedPaths -and $cleanedPaths.Count -gt 0) {
					$finalResult | Add-Member -NotePropertyName "CleanedPaths" -NotePropertyValue $cleanedPaths -Force
					New-Log "Successfully cleaned $($cleanedPaths.Count) old items for '$moduleName'." -Level SUCCESS
				}
				else {
					New-Log "Cleaning step completed for '$moduleName'. No old versions found or removed." -Level INFO
				}
			}
			elseif ($Clean.IsPresent -and -not $finalResult.OverallSuccess) {
				New-Log "Skipping cleaning for '$moduleName' as the update was not fully successful." -Level INFO
			}
			$aggregateResults.Add($finalResult)
		}
		if ($UseProgressBar) {
			Write-Progress -Activity "Updating PowerShell Modules" -Completed
		}
		$successCount = ($aggregateResults | Where-Object { $_.OverallSuccess }).Count
		$failCount = $total - $successCount
		New-Log "Update process finished. Successful: $successCount, Failed/Partial: $failCount (of $total)." -Level SUCCESS
		return $aggregateResults
	}
}
function Install-PSModule {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$ModuleName,
		[Parameter(Mandatory)][version]$TargetVersion,
		[Parameter(Mandatory)][string]$RepositoryName,
		[Parameter(Mandatory)][bool]$IsPreview,
		[Parameter(Mandatory)][string[]]$Destinations,
		[string]$PreReleaseVersion = $null
	)
	$result = [PSCustomObject]@{
		UpdatedPaths = [System.Collections.Generic.List[string]]::new()
		FailedPaths  = [System.Collections.Generic.List[string]]::new()
	}
	$targetVersionString = $TargetVersion.ToString()
	$commonSaveParams = @{
		Name                = $ModuleName
		Version             = if ($PreReleaseVersion -and $IsPreview) { "$targetVersionString-$preReleaseVersion" } else { $targetVersionString }
		Repository          = $RepositoryName
		TrustRepository     = $true
		IncludeXml          = $true
		SkipDependencyCheck = $true
		AcceptLicense       = $true
		Confirm             = $false
		PassThru            = $true
		ErrorAction         = 'SilentlyContinue'
		WarningAction       = 'SilentlyContinue'
	}
	if ($IsPreview) { $commonSaveParams.Prerelease = $true }
	$pathsToRetry = [System.Collections.Generic.List[string]]::new()
	$saveSucceededOnce = $false
	foreach ($destinationBasePath in $Destinations) {
		$savePath = Split-Path $destinationBasePath -Parent
		$saveTargetDir = $savePath
		New-Log "Attempting Save-PSResource for [$ModuleName] version [v$($commonSaveParams.Version)] to '$saveTargetDir'..." -Level DEBUG
		$res = $null
		$res = Save-PSResource @commonSaveParams -Path $saveTargetDir
		if ($res -and $res[-1].Version -eq $TargetVersion -and $res[-1].PreRelease -eq $PreReleaseVersion ) {
			$savedModulePath = Join-Path -Path $res[-1].InstalledLocation -ChildPath $ModuleName # Verify the actual saved path
			New-Log "Successfully saved [$ModuleName] version [v$($commonSaveParams.Version)] via Save-PSResource to '$savedModulePath'" -Level SUCCESS
			$result.UpdatedPaths.Add($destinationBasePath) # Report success for the target base path
			$saveSucceededOnce = $true
		}
		else {
			New-Log "Save-PSResource failed for [$ModuleName] v$targetVersionString to '$savePath'" -Level WARNING
			$pathsToRetry.Add($destinationBasePath) # Mark this base path for potential fallback
		}
	}
	if ($pathsToRetry.Count -gt 0 -or (-not $saveSucceededOnce -and $Destinations.Count -gt 0)) {
		New-Log "Save-PSResource did not succeed for all target paths. Attempting fallback install methods..." -Level INFO
		$fallbackInstallSucceeded = $false
		try {
			$psResourceParams = @{
				Name                = $ModuleName
				Version             = if ($PreReleaseVersion -and $IsPreview) { "$targetVersionString-$preReleaseVersion" } else { $targetVersionString }
				Scope               = 'AllUsers'
				AcceptLicense       = $true
				SkipDependencyCheck = $true
				Confirm             = $false
				Reinstall           = $true
				TrustRepository     = $true
				Repository          = $RepositoryName
				PassThru            = $true
				ErrorAction         = 'Stop'
				WarningAction       = 'SilentlyContinue'
			}
			if ($IsPreview) { $psResourceParams.Prerelease = $true }
			New-Log "Attempting Install-PSResource for [$ModuleName] version [v$($psResourceParams.Version)]..." -Level DEBUG
			$res = $null
			$res = Install-PSResource @psResourceParams
			if ($res -and $res[-1].InstalledLocation -and $res[-1].Version -eq $TargetVersion -and $res[-1].PreRelease -eq $PreReleaseVersion) {
				New-Log "Successfully installed [$ModuleName] version [v$($psResourceParams.Version)] via Install-PSResource to '$($res[-1].InstalledLocation)'" -Level SUCCESS
				$fallbackInstallSucceeded = $true
				foreach ($retryPath in $pathsToRetry) {
					if ($($res[-1].InstalledLocation) -eq $retryPath) {
						$result.UpdatedPaths.Add($retryPath)
					}
				}
			}
			else {
				New-Log "Install-PSResource did not return expected result or version for [$ModuleName]." -Level VERBOSE
			}
		}
		catch {
			New-Log "Install-PSResource failed for [$ModuleName]. Trying Install-Module..." -Level ERROR
		}
		if (-not $fallbackInstallSucceeded) {
			try {
				$installModuleParams = @{
					Name               = $ModuleName
					RequiredVersion    = if ($PreReleaseVersion -and $IsPreview) { "$targetVersionString-$preReleaseVersion" } else { $targetVersionString }
					Scope              = 'AllUsers'
					Force              = $true
					AcceptLicense      = $true
					SkipPublisherCheck = $true
					AllowClobber       = $true
					PassThru           = $true
					Repository         = $RepositoryName
					ErrorAction        = 'Stop'
					WarningAction      = 'SilentlyContinue'
					Confirm            = $false
				}
				if ($IsPreview) { $installModuleParams.AllowPrerelease = $true }
				New-Log "Attempting Install-Module for [$ModuleName] version [v$($installModuleParams.RequiredVersion)]..." -Level DEBUG
				$res = $null
				$res = Install-Module @installModuleParams
				if ($res -and $res[-1].InstalledLocation -and $res[-1].Version -eq $TargetVersion -and $res[-1].PreRelease -eq $PreReleaseVersion) {
					New-Log "Successfully installed [$ModuleName] version [v$($installModuleParams.RequiredVersion)] via Install-Module to '$($res[-1].InstalledLocation)'" -Level SUCCESS
					$fallbackInstallSucceeded = $true
					foreach ($retryPath in $pathsToRetry) {
						if ($($res[-1].InstalledLocation) -eq $retryPath) {
							$result.UpdatedPaths.Add($retryPath)
						}
					}
				}
				else {
					New-Log "Install-Module did not return expected result or version for [$ModuleName]." -Level VERBOSE
				}
			}
			catch {
				New-Log "Install-Module (fallback) failed for [$ModuleName]." -Level ERROR
			}
		}
		foreach ($retryPath in $pathsToRetry) {
			if ($result.UpdatedPaths -notcontains $retryPath) {
				$result.FailedPaths.Add($retryPath)
			}
		}
	}
	$result.UpdatedPaths = $result.UpdatedPaths | Select-Object -Unique
	$result.FailedPaths = $result.FailedPaths | Select-Object -Unique
	$result.FailedPaths = $result.FailedPaths | Where-Object { $result.UpdatedPaths -notcontains $_ }
	if ($result.UpdatedPaths.Count -gt 0 -and $result.FailedPaths.Count -eq 0) {
		New-Log "Successfully updated [$ModuleName] v$targetVersionString for all target destinations." -Level SUCCESS
	}
	elseif ($result.UpdatedPaths.Count -gt 0) {
		New-Log "Partially updated [$ModuleName] v$targetVersionString. Succeeded: $($result.UpdatedPaths -join '; '). Failed: $($result.FailedPaths -join '; ')" -Level WARNING
	}
	else {
		New-Log "Failed to update [$ModuleName] v$targetVersionString for destinations: $($result.FailedPaths -join '; ')" -Level ERROR
	}
	return $result
}
function Remove-OutdatedVersions {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)][string]$ModuleName,
		[Parameter(Mandatory)][string[]]$ModuleBasePaths,
		[Parameter(Mandatory)][version]$LatestVersion,
		[string[]]$OutdatedVersions,
		[string[]]$DoNotClean = @('PowerShellGet', 'Microsoft.PowerShell.PSResourceGet'),
		[string]$PreReleaseTag = $null
	)
	if ($ModuleName -in $DoNotClean) {
		New-Log "Skipping cleaning for module [$ModuleName] as it is in the DoNotClean list." -Level INFO
		return @()
	}
	New-Log "Starting cleanup of old versions for [$ModuleName] (keeping v$($LatestVersion.ToString()))..." -Level INFO
	$cleanedItems = [System.Collections.Generic.List[string]]::new()
	$latestVersionString = $LatestVersion.ToString()
	foreach ($basePath in $ModuleBasePaths) {
		if (-not (Test-Path $basePath -PathType Container)) {
			New-Log "Base path '$basePath' for cleaning does not exist or is not a directory. Skipping." -Level WARNING
			continue
		}
		New-Log "Checking for old versions within '$basePath'..." -Level DEBUG
		$foundOldVersionDirs = $false
		$versionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue | Where-Object {
			$_.Name -match '^\d+(\.\d+){1,4}(-.+)?$' -and
			$_.Name -ne $latestVersionString -and
			$_.Name -ne "$latestVersionString-$preReleaseVersion"
		}
		if ($versionFolders) {
			$foundOldVersionDirs = $true
			foreach ($versionFolder in $versionFolders) {
				$folderPath = $versionFolder.FullName
				$versionString = $versionFolder.Name
				New-Log "Found old version folder: '$folderPath'. Attempting removal..." -Level DEBUG
				$removed = $false
				if ($PreReleaseTag) {
					Uninstall-PSResource -Name $ModuleName -Version "$versionString-$PreReleaseTag" -Scope AllUsers -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
					Uninstall-PSResource -Name $ModuleName -Version $versionString -Scope AllUsers -Prerelease -SkipDependencyCheck -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
				}
				if (Test-Path -LiteralPath $folderPath) {
					Uninstall-PSResource -Name $ModuleName -Version $versionString -Scope AllUsers -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
				}
				Start-Sleep -Milliseconds 200
				if (-not (Test-Path -LiteralPath $folderPath)) {
					New-Log "Successfully removed '$folderPath' via Uninstall-PSResource." -Level SUCCESS
					$cleanedItems.Add($folderPath)
					$removed = $true
				}
				else {
					New-Log "Uninstall-PSResource ran for '$versionString' but folder '$folderPath' still exists. Will attempt Remove-Item." -Level WARNING
				}
				if (-not $removed -and (Test-Path -LiteralPath $folderPath)) {
					try {
						Remove-Item -LiteralPath $folderPath -Recurse -Force -ErrorAction Stop | Out-Null
						Start-Sleep -Milliseconds 200
						if (-not (Test-Path -LiteralPath $folderPath)) {
							New-Log "Successfully removed old folder '$folderPath' via Remove-Item." -Level SUCCESS
							$cleanedItems.Add($folderPath)
						}
						else {
							New-Log "Remove-Item ran but folder '$folderPath' still exists. Manual cleanup might be needed." -Level WARNING
						}
					}
					catch {
						New-Log "Failed to remove old folder '$folderPath' via Remove-Item. Error: $($_.Exception.Message). Manual cleanup might be required." -Level ERROR
						$Error.Clear()
					}
				}
			}
		}
		if (-not $versionFolders) {
			New-Log "No specific version folders found in '$basePath'. Trying Uninstall-PSResource with provided local versions..." -Level DEBUG
			foreach ($localVer in ($OutdatedVersions | Where-Object { $_ -ne $latestVersionString -or $_ -ne "$latestVersionString-$PreReleaseTag" })) {
				try {
					New-Log "Attempting Uninstall-PSResource for '$ModuleName' version 'v$localVer'..." -Level DEBUG
					Uninstall-PSResource -Name $ModuleName -Version $localVer -Scope AllUsers -Confirm:$false -SkipDependencyCheck -ErrorAction Stop
					New-Log "Successfully uninstalled '$ModuleName' version 'v$localVer' via fallback." -Level SUCCESS
				}
				catch {
					Uninstall-PSResource -Name $ModuleName -Version "$localVer-$PreReleaseTag" -Scope AllUsers -SkipDependencyCheck -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
					try {
						Uninstall-PSResource -Name $ModuleName -Version $localVer -Scope AllUsers -SkipDependencyCheck -Prerelease -Confirm:$false -ErrorAction Stop | Out-Null
						New-Log "Successfully uninstalled '$ModuleName' version 'v$localVer-$PreReleaseTag' via fallback." -Level SUCCESS
					}
					catch {
						New-Log "Failed with backup way to Uninstall-PSResource for module '$ModuleName'." -Level WARNING
					}
				}
			}
		}
		return $cleanedItems
	}
}
### OBS: New-Log Function is needed otherwise remove all New-Log and replace with Write-Host. New-Log is vastly better though, check the link below:
#Example:
<#
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Harze2k/Shared-PowerShell-Functions/main/New-Log.ps1" -UseBasicParsing -MaximumRedirection 1 | Select-Object -ExpandProperty Content | Invoke-Expression
Check-PSResourceRepository -ImportDependencies
$ignoredModules = @('Example2.Diagnostics') #Fully ignored modules
$blackList = @{ #Ignored module and repo combo.
	'Microsoft.Graph.Beta' = 'NuGetGallery'
	'Microsoft.Graph'      = @("Nuget", "NugetGallery")
}
$paths = $env:PSModulePath.Split(';') | Where-Object { $_ -inotmatch '.vscode' }
$moduleInfo = Get-ModuleInfo -Paths $paths -IgnoredModules $ignoredModules
$outdated = Get-ModuleUpdateStatus -ModuleInventory $moduleInfo -TimeoutSeconds 30 -Repositories @("PSGallery", "Nuget", "NugetGallery") -BlackList $blackList -MatchAuthor
if ($outdated) {
	$res = $outdated | Update-Modules -Clean -UseProgressBar
	$res
}
else {
	New-Log "No outdated modules found!" -Level SUCCESS
}
#>