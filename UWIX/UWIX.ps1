Function LoginPsPnP()  #*** LEGACY CODE *** 
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

Function LoginPsPnP($fullSiteUrl)  #*** LEGACY CODE *** 
{
	if($fullSiteUrl -ne $null) {
		[SecureString]$securePW = ConvertTo-SecureString -String `
				$configFile.appsettings.UserPw -AsPlainText -Force

		$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
				-argumentlist $configFile.appsettings.UserName, $securePW
		Connect-PnPOnline -Url $fullSiteUrl -Credentials $myCredentials
	}
}

#----------------------------------------------------------------------------------------

#gavdcodebegin 001
Function SpPsPnp_CreateOneSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP"
	New-PnPSite -Type CommunicationSite `
				-Title "NewSiteCollectionModernPsPnP" `
				-Url $fullSiteUrl `
				-SiteDesign "Showcase"
}
#gavdcodeend 001

#gavdcodebegin 002
Function SpPsPnp_CreateOneSiteCollection01()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	New-PnPTenantSite -Title "NewSiteCollModernPsPnP01" `
					  -Url $fullSiteUrl `
					  -Owner $configFile.appsettings.UserName `
					  -Template STS#3 `
					  -TimeZone 4
}
#gavdcodeend 002

#gavdcodebegin 003
Function SpPsPnp_GetAllSiteCollections()  #*** LEGACY CODE *** 
{
	Get-PnPTenantSite
}
#gavdcodeend 003

#gavdcodebegin 004
Function SpPsPnp_GetOneSiteCollection()  #*** LEGACY CODE *** 
{
	Get-PnPSite
}
#gavdcodeend 004

#gavdcodebegin 005
Function SpPsPnp_GetAllSiteCollectionsFiltered()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP"
	Get-PnPTenantSite -Template "SITEPAGEPUBLISHING#0" -Detailed
}
#gavdcodeend 005

#gavdcodebegin 006
Function SpPsPnp_GetHubSiteCollection()  #*** LEGACY CODE *** 
{
	Get-PnPHubSite
}
#gavdcodeend 006

#gavdcodebegin 007
Function SpPsPnp_UpdateOneSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	LoginPsPnP $fullSiteUrl
	Set-PnPSite -CommentsOnSitePagesDisabled
}
#gavdcodeend 007

#gavdcodebegin 008
Function SpPsPnp_UpdateOneSiteCollection01()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	Set-PnPTenantSite -Url $fullSiteUrl -Title "NewSiteCollModernPsPnP01_Updated"
}
#gavdcodeend 008

#gavdcodebegin 009
Function SpPsPnp_DeleteOneSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	Remove-PnPTenantSite -Url $fullSiteUrl -Force -SkipRecycleBin
}
#gavdcodeend 009

#gavdcodebegin 010
Function SpPsPnp_RegisterHubSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	Register-PnPHubSite -Site $fullSiteUrlHub
}
#gavdcodeend 010

#gavdcodebegin 011
Function SpPsPnp_UnregisterHubSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	Unregister-PnPHubSite -Site $fullSiteUrlHub
}
#gavdcodeend 011

#gavdcodebegin 012
Function SpPsPnp_AddSiteToHubSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	Add-PnPHubSiteAssociation -Site $fullSiteUrl -HubSite $fullSiteUrlHub
}
#gavdcodeend 012

#gavdcodebegin 013
Function SpPsPnp_RemoveSiteFromHubSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	Remove-PnPHubSiteAssociation -Site $fullSiteUrl
}
#gavdcodeend 013

#gavdcodebegin 014
Function SpPsPnp_GetAdminsInSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Get-PnPSiteCollectionAdmin 
}
#gavdcodeend 014

#gavdcodebegin 015
Function SpPsPnp_AddAdminsToSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Add-PnPSiteCollectionAdmin -Owners "domain@domain.onmicrosoft.com"
}
#gavdcodeend 015

#gavdcodebegin 016
Function SpPsPnp_RemoveAdminsFromSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Remove-PnPSiteCollectionAdmin -Owners "domain@user.onmicrosoft.com"
}
#gavdcodeend 016

#gavdcodebegin 017
Function SpPsPnp_CreateWebInSiteCollection()  #*** LEGACY CODE *** 
{
	New-PnPWeb -Title "NewWebSiteModernPsPnP" `
			   -Url "NewWebSiteModernPsPnP" `
			   -Description "NewWebSiteModernPsPnP Description" `
			   -Locale "1033" `
			   -Template "STS#3"
}
#gavdcodeend 017

#gavdcodebegin 018
Function SpPsPnp_GetOneWebInSiteCollection()  #*** LEGACY CODE *** 
{
	Get-PnPWeb
}
#gavdcodeend 018

#gavdcodebegin 019
Function SpPsPnp_GetWebsInSiteCollection()  #*** LEGACY CODE *** 
{
	Get-PnPSubWebs -Recurse
}
#gavdcodeend 019

#gavdcodebegin 020
Function SpPsPnp_UpdateOneWebInSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsPnP"
	LoginPsPnP $fullSiteUrl

	Set-PnPWeb -Description "NewWebSiteModernPsPnP Description Updated"
}
#gavdcodeend 020

#gavdcodebegin 021
Function SpPsPnp_AddPermissionsInWebInSiteCollection()  #*** LEGACY CODE *** 
{
	Set-PnPWebPermission -Url "NewWebSiteModernPsPnP" `
						 -User 'user@domain.onmicrosoft.com' `
						 -AddRole 'Contribute'
}
#gavdcodeend 021

#gavdcodebegin 022
Function SpPsPnp_RemoveOneWebFromSiteCollection()  #*** LEGACY CODE *** 
{
	Remove-PnPWeb -Url "NewWebSiteModernPsPnP"
}
#gavdcodeend 022

#----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$spCtx = LoginPsPnP

#SpPsPnp_CreateOneSiteCollection
#SpPsPnp_CreateOneSiteCollection01
#SpPsPnp_GetAllSiteCollections
#SpPsPnp_GetOneSiteCollection
#SpPsPnp_GetAllSiteCollectionsFiltered
#SpPsPnp_GetHubSiteCollection
#SpPsPnp_UpdateOneSiteCollection
#SpPsPnp_UpdateOneSiteCollection01
#SpPsPnp_DeleteOneSiteCollection
#SpPsPnp_RegisterHubSiteCollection
#SpPsPnp_UnregisterHubSiteCollection
#SpPsPnp_AddSiteToHubSiteCollection
#SpPsPnp_RemoveSiteFromHubSiteCollection
#SpPsPnp_GetAdminsInSiteCollection
#SpPsPnp_AddAdminsToSiteCollection
#SpPsPnp_RemoveAdminsFromSiteCollection
#SpPsPnp_GrantRightsHubSiteCollection

#SpPsPnp_CreateWebInSiteCollection
#SpPsPnp_GetOneWebInSiteCollection
#SpPsPnp_GetWebsInSiteCollection
#SpPsPnp_UpdateOneWebInSiteCollection
#SpPsPnp_AddPermissionsInWebInSiteCollection
#SpPsPnp_RemoveOneWebFromSiteCollection

Write-Host "Done"