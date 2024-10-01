
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsSpPnP_LoginWithAccPw()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithAccPw `
					  -Credentials $myCredentials
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function SpPsPnPSharePoint_GetTenantProps
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenant
}
#gavdcodeend 001

#gavdcodebegin 002
function SpPsPnPSharePoint_SetTenantProps
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Set-PnPTenant -PreventExternalUsersFromResharing $false
}
#gavdcodeend 002

#gavdcodebegin 003
function SpPsPnPSharePoint_GetTenantId
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantId
}
#gavdcodeend 003

#gavdcodebegin 004
function SpPsPnPSharePoint_GetTenantInstance
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantInstance
}
#gavdcodeend 004

#gavdcodebegin 005
function SpPsPnPSharePoint_GetSitesInRecycleBin
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
#gavdcodeend 005

#gavdcodebegin 006
function SpPsPnPSharePoint_RestoreSiteFromRecycleBin
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Restore-PnPTenantRecycleBinItem `
						-Url "https://domain.sharepoint.com/sites/TestToRecycleBin" `
						-Force
}
#gavdcodeend 006

#gavdcodebegin 007
function SpPsPnPSharePoint_ClearSitesFromRecycleBin
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Clear-PnPTenantRecycleBinItem `
						-Url "https://domain.sharepoint.com/sites/TestToRecycleBin" `
						-Force
}
#gavdcodeend 007

#gavdcodebegin 008
function SpPsPnPSharePoint_GetTenantTheme
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantTheme
	#Get-PnPTenantTheme -Name "myTheme"
}
#gavdcodeend 008

#gavdcodebegin 009
function SpPsPnPSharePoint_AddTenantTheme
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
#gavdcodeend 009

#gavdcodebegin 010
function SpPsPnPSharePoint_DeleteTenantTheme
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Remove-PnPTenantTheme -Identity "myTheme"
}
#gavdcodeend 010

#gavdcodebegin 011
function SpPsPnPSharePoint_GetTenantAppCatalog
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantAppCatalogUrl
}
#gavdcodeend 011

#gavdcodebegin 012
function SpPsPnPSharePoint_CreateTenantAppCatalog
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Set-PnPTenantAppCatalogUrl -Url "https://domain.sharepoint.com/sites/AppCatalog"
}
#gavdcodeend 012

#gavdcodebegin 013
function SpPsPnPSharePoint_ClearTenantAppCatalog
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Clear-PnPTenantAppCatalogUrl
}
#gavdcodeend 013

#gavdcodebegin 014
function SpPsPnPSharePoint_IsCDNAvailable
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantCdnEnabled -CdnType Public
	#Get-PnPTenantCdnEnabled -CdnType Private
}
#gavdcodeend 014

#gavdcodebegin 015
function SpPsPnPSharePoint_SetCDNAvailable
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Set-PnPTenantCdnEnabled -Enable $true -CdnType Public
	#Set-PnPTenantCdnEnabled -Enable $true -CdnType Private
}
#gavdcodeend 015

#gavdcodebegin 016
function SpPsPnPSharePoint_GetCDNOrigens
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantCdnOrigin -CdnType Public
	#Get-PnPTenantCdnOrigin -CdnType Private
}
#gavdcodeend 016

#gavdcodebegin 017
function SpPsPnPSharePoint_GetCDNPolicy
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	Get-PnPTenantCdnPolicies -CdnType Public
	#Get-PnPTenantCdnPolicies -CdnType Private
}
#gavdcodeend 017

#gavdcodebegin 018
function SpPsPnPSharePoint_CreateCDNOrigen
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Add-PnPTenantCdnOrigin -OriginUrl "/sites/USSales/myCDN" -CdnType Public
	#Add-PnPTenantCdnOrigin -OriginUrl /sites/SiteColl/myCDN -CdnType Private
}
#gavdcodeend 018

#gavdcodebegin 019
function SpPsPnPSharePoint_UpdateCDNPolicy
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Set-PnPTenantCdnPolicy -CdnType Public `
						   -PolicyType ExcludeRestrictedSiteClassifications `
						   -PolicyValue "Confidential,Restricted"
}
#gavdcodeend 019

#gavdcodebegin 020
function SpPsPnPSharePoint_DeleteCDNOrigen
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	Remove-PnPTenantCdnOrigin -OriginUrl "/sites/USSales/myCDN" -CdnType Public
	#Remove-PnPTenantCdnOrigin -OriginUrl /sites/SiteColl/myCDN -CdnType Private
}
#gavdcodeend 020

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the PnP PowerShell module --------
# Connect to Office 365
$spCtx = PsSpPnP_LoginWithAccPw

#SpPsPnPSharePoint_GetTenantProps
#SpPsPnPSharePoint_GetTenantId
#SpPsPnPSharePoint_GetTenantInstance
#SpPsPnPSharePoint_SetTenantProps
#SpPsPnPSharePoint_GetSitesInRecycleBin
#SpPsPnPSharePoint_RestoreSiteFromRecycleBin
#SpPsPnPSharePoint_ClearSitesFromRecycleBin
#SpPsPnPSharePoint_GetTenantTheme
#SpPsPnPSharePoint_AddTenantTheme
#SpPsPnPSharePoint_DeleteTenantTheme
#SpPsPnPSharePoint_GetTenantAppCatalog
#SpPsPnPSharePoint_CreateTenantAppCatalog
#SpPsPnPSharePoint_ClearTenantAppCatalog
#SpPsPnPSharePoint_IsCDNAvailable
#SpPsPnPSharePoint_SetCDNAvailable
#SpPsPnPSharePoint_GetCDNOrigens
#SpPsPnPSharePoint_GetCDNPolicy
#SpPsPnPSharePoint_CreateCDNOrigen
#SpPsPnPSharePoint_UpdateCDNPolicy
#SpPsPnPSharePoint_DeleteCDNOrigen

# Disconnect from Office 365
Disconnect-PnPOnline


Write-Host "Done" 
