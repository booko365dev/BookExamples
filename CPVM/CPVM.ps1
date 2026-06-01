
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsGraphSdk_LoginWithSecret
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$ClientSecret -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $ClientID, $securePW

	Connect-MgGraph -TenantId $TenantName `
					-ClientSecretCredential $myCredentials
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpGraphSdk_GetTenantConfiguration
{
	# App Registration type: Graph 
	# App Registration permissions: SharePointTenantSettings.ReadWrite.All
	
	PsGraphSdk_LoginWithSecret -TenantName $cnfTenantName `
							   -ClientID $cnfClientIdWithSecret `
							   -ClientSecret $cnfClientSecret

	$myConfigs = Get-MgAdminSharepointSetting
	Write-Host "Loop Is Enabled - " $myConfigs.IsLoopEnabled
    
	Disconnect-MgGraph
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpGraphSdk_UpdateTenantConfiguration
{
	# App Registration type: Graph 
	# App Registration permissions: SharePointTenantSettings.ReadWrite.All
	
	PsGraphSdk_LoginWithSecret -TenantName $cnfTenantName `
							   -ClientID $cnfClientIdWithSecret `
							   -ClientSecret $cnfClientSecret

	$myConfigs = @{
		IsLoopEnabled = $false
	}
	Update-MgAdminSharepointSetting -BodyParameter $myConfigs
    
	Disconnect-MgGraph
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpGraphSdk_GetAppsInCatalog
{
	# App Registration type: Graph 
	# App Registration permissions: AppCatalog.ReadWrite.All
	
	PsGraphSdk_LoginWithSecret -TenantName $cnfTenantName `
							   -ClientID $cnfClientIdWithSecret `
							   -ClientSecret $cnfClientSecret

	Get-MgAppCatalogTeamApp

	Disconnect-MgGraph
}
#gavdcodeend 003

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 003 ***

#region ConfigValuesCS.config
[xml]$config = Get-Content -Path "C:\Projects\ConfigValuesCS.config"
$cnfUserName               = $config.SelectSingleNode("//add[@key='UserName']").value
$cnfUserPw                 = $config.SelectSingleNode("//add[@key='UserPw']").value
$cnfTenantUrl              = $config.SelectSingleNode("//add[@key='TenantUrl']").value     # https://domain.onmicrosoft.com
$cnfSiteBaseUrl            = $config.SelectSingleNode("//add[@key='SiteBaseUrl']").value   # https://domain.sharepoint.com
$cnfSiteAdminUrl           = $config.SelectSingleNode("//add[@key='SiteAdminUrl']").value  # https://domain-admin.sharepoint.com
$cnfSiteCollUrl            = $config.SelectSingleNode("//add[@key='SiteCollUrl']").value   # https://domain.sharepoint.com/sites/TestSite
$cnfTenantName             = $config.SelectSingleNode("//add[@key='TenantName']").value
$cnfClientIdWithAccPw      = $config.SelectSingleNode("//add[@key='ClientIdWithAccPw']").value
$cnfClientIdWithSecret     = $config.SelectSingleNode("//add[@key='ClientIdWithSecret']").value
$cnfClientSecret           = $config.SelectSingleNode("//add[@key='ClientSecret']").value
$cnfClientIdWithCert       = $config.SelectSingleNode("//add[@key='ClientIdWithCert']").value
$cnfCertificateThumbprint  = $config.SelectSingleNode("//add[@key='CertificateThumbprint']").value
$cnfCertificateFilePath    = $config.SelectSingleNode("//add[@key='CertificateFilePath']").value
$cnfCertificateFilePw      = $config.SelectSingleNode("//add[@key='CertificateFilePw']").value
#endregion ConfigValuesCS.config

#PsSpGraphSdk_GetTenantConfiguration
#PsSpGraphSdk_UpdateTenantConfiguration
#PsSpGraphSdk_GetAppsInCatalog  # Apps for Teams (including also SharePoint Catalog apps)

Write-Host "Done" 

