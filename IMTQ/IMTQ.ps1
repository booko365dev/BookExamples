
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function LoginPsPnPPowerShell()
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 01
function PsPnPSharePoint_GetTenantProps
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenant
}
#gavdcodeend 01

#gavdcodebegin 02
function PsPnPSharePoint_SetTenantProps
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Set-PnPTenant -PreventExternalUsersFromResharing $false
}
#gavdcodeend 02

#gavdcodebegin 03
function PsPnPSharePoint_GetTenantId
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantId
}
#gavdcodeend 03

#gavdcodebegin 04
function PsPnPSharePoint_GetTenantInstance
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantInstance
}
#gavdcodeend 04

#gavdcodebegin 05
function PsPnPSharePoint_GetSitesInRecycleBin
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$mySiteColls = Get-PnPTenantRecycleBinItem
	foreach($oneSiteColl in $mySiteColls)
	{
		Write-Host $oneSiteColl.SiteId + " - " + $oneSiteColl.Url
	}
}
#gavdcodeend 05

#gavdcodebegin 06
function PsPnPSharePoint_RestoreSiteFromRecycleBin
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Restore-PnPTenantRecycleBinItem `
						-Url "https://domain.sharepoint.com/sites/TestToRecycleBin" `
						-Force
}
#gavdcodeend 06

#gavdcodebegin 07
function PsPnPSharePoint_ClearSitesFromRecycleBin
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Clear-PnPTenantRecycleBinItem `
						-Url "https://domain.sharepoint.com/sites/TestToRecycleBin" `
						-Force
}
#gavdcodeend 07

#gavdcodebegin 08
function PsPnPSharePoint_GetTenantTheme
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantTheme
	#Get-PnPTenantTheme -Name "myTheme"
}
#gavdcodeend 08

#gavdcodebegin 09
function PsPnPSharePoint_AddTenantTheme
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSitesWrite.Read
	
	$themeColors = @{
	  "black" = "#000000";
	  "white" = "#ffffff";
	  "primaryBackground" = "#a5a5a5";
	 }
	Add-PnPTenantTheme -Identity "myTheme" `
						-Palette $themeColors `
						-IsInverted $false `
						-Overwrite
}
#gavdcodeend 09

#gavdcodebegin 10
function PsPnPSharePoint_DeleteTenantTheme
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Remove-PnPTenantTheme -Identity "myTheme"
}
#gavdcodeend 10

#gavdcodebegin 11
function PsPnPSharePoint_GetTenantAppCatalog
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantAppCatalogUrl
}
#gavdcodeend 11

#gavdcodebegin 12
function PsPnPSharePoint_CreateTenantAppCatalog
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Set-PnPTenantAppCatalogUrl -Url "https://domain.sharepoint.com/sites/AppCatalog"
}
#gavdcodeend 12

#gavdcodebegin 13
function PsPnPSharePoint_ClearTenantAppCatalog
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Clear-PnPTenantAppCatalogUrl
}
#gavdcodeend 13

#gavdcodebegin 14
function PsPnPSharePoint_IsCDNAvailable
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantCdnEnabled -CdnType Public
	#Get-PnPTenantCdnEnabled -CdnType Private
}
#gavdcodeend 14

#gavdcodebegin 15
function PsPnPSharePoint_SetCDNAvailable
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Set-PnPTenantCdnEnabled -Enable $true -CdnType Public
	#Set-PnPTenantCdnEnabled -Enable $true -CdnType Private
}
#gavdcodeend 15

#gavdcodebegin 16
function PsPnPSharePoint_GetCDNOrigens
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantCdnOrigin -CdnType Public
	#Get-PnPTenantCdnOrigin -CdnType Private
}
#gavdcodeend 16

#gavdcodebegin 17
function PsPnPSharePoint_GetCDNPolicy
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantCdnPolicies -CdnType Public
	#Get-PnPTenantCdnPolicies -CdnType Private
}
#gavdcodeend 17

#gavdcodebegin 18
function PsPnPSharePoint_CreateCDNOrigen
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Add-PnPTenantCdnOrigin -OriginUrl "/sites/USSales/myCDN" -CdnType Public
	#Add-PnPTenantCdnOrigin -OriginUrl /sites/SiteColl/myCDN -CdnType Private
}
#gavdcodeend 18

#gavdcodebegin 19
function PsPnPSharePoint_UpdateCDNPolicy
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Set-PnPTenantCdnPolicy -CdnType Public `
						   -PolicyType ExcludeRestrictedSiteClassifications `
						   -PolicyValue "Confidential,Restricted"
}
#gavdcodeend 19

#gavdcodebegin 20
function PsPnPSharePoint_DeleteCDNOrigen
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Remove-PnPTenantCdnOrigin -OriginUrl "/sites/USSales/myCDN" -CdnType Public
	#Remove-PnPTenantCdnOrigin -OriginUrl /sites/SiteColl/myCDN -CdnType Private
}
#gavdcodeend 20

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the PnP PowerShell module --------
# Connect to Office 365
$spCtx = LoginPsPnPPowerShell

#PsPnPSharePoint_GetTenantProps
#PsPnPSharePoint_GetTenantId
#PsPnPSharePoint_GetTenantInstance
#PsPnPSharePoint_SetTenantProps
#PsPnPSharePoint_GetSitesInRecycleBin
#PsPnPSharePoint_RestoreSiteFromRecycleBin
#PsPnPSharePoint_ClearSitesFromRecycleBin
#PsPnPSharePoint_GetTenantTheme
#PsPnPSharePoint_AddTenantTheme
#PsPnPSharePoint_DeleteTenantTheme
#PsPnPSharePoint_GetTenantAppCatalog
#PsPnPSharePoint_CreateTenantAppCatalog
#PsPnPSharePoint_ClearTenantAppCatalog
#PsPnPSharePoint_IsCDNAvailable
#PsPnPSharePoint_SetCDNAvailable
#PsPnPSharePoint_GetCDNOrigens
#PsPnPSharePoint_GetCDNPolicy
#PsPnPSharePoint_CreateCDNOrigen
#PsPnPSharePoint_UpdateCDNPolicy
#PsPnPSharePoint_DeleteCDNOrigen

# Disconnect from Office 365
Disconnect-PnPOnline


Write-Host "Done" 
