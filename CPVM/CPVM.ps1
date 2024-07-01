
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
	
	PsGraphSdk_LoginWithSecret -TenantName $configFile.appsettings.TenantName `
								 -ClientID $configFile.appsettings.ClientIdWithSecret `
								 -ClientSecret $configFile.appsettings.ClientSecret

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
	
	PsGraphSdk_LoginWithSecret -TenantName $configFile.appsettings.TenantName `
								 -ClientID $configFile.appsettings.ClientIdWithSecret `
								 -ClientSecret $configFile.appsettings.ClientSecret

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
	
	PsGraphSdk_LoginWithSecret -TenantName $configFile.appsettings.TenantName `
								 -ClientID $configFile.appsettings.ClientIdWithSecret `
								 -ClientSecret $configFile.appsettings.ClientSecret

	Get-MgAppCatalogTeamApp

	Disconnect-MgGraph
}
#gavdcodeend 003

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: xxx ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#PsSpGraphSdk_GetTenantConfiguration
#PsSpGraphSdk_UpdateTenantConfiguration
#PsSpGraphSdk_GetAppsInCatalog  # Apps for Teams (including also SharePoint Catalog apps)

Write-Host "Done" 

