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

#gavdcodebegin 01
Function SpPsPnpCreateOneSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP"
	New-PnPSite -Type CommunicationSite `
				-Title "NewSiteCollectionModernPsPnP" `
				-Url $fullSiteUrl `
				-SiteDesign "Showcase"
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsPnpCreateOneSiteCollection01()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	New-PnPTenantSite -Title "NewSiteCollModernPsPnP01" `
					  -Url $fullSiteUrl `
					  -Owner $configFile.appsettings.UserName `
					  -Template STS#3 `
					  -TimeZone 4
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsPnpGetAllSiteCollections()  #*** LEGACY CODE *** 
{
	Get-PnPTenantSite
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsPnpGetOneSiteCollection()  #*** LEGACY CODE *** 
{
	Get-PnPSite
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsPnpGetAllSiteCollectionsFiltered()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP"
	Get-PnPTenantSite -Template "SITEPAGEPUBLISHING#0" -Detailed
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsPnpGetHubSiteCollection()  #*** LEGACY CODE *** 
{
	Get-PnPHubSite
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsPnpUpdateOneSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	LoginPsPnP $fullSiteUrl
	Set-PnPSite -CommentsOnSitePagesDisabled
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsPnpUpdateOneSiteCollection01()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	Set-PnPTenantSite -Url $fullSiteUrl -Title "NewSiteCollModernPsPnP01_Updated"
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsPnpDeleteOneSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	Remove-PnPTenantSite -Url $fullSiteUrl -Force -SkipRecycleBin
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsPnpRegisterHubSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	Register-PnPHubSite -Site $fullSiteUrlHub
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsPnpUnregisterHubSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	Unregister-PnPHubSite -Site $fullSiteUrlHub
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsPnpAddSiteToHubSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	Add-PnPHubSiteAssociation -Site $fullSiteUrl -HubSite $fullSiteUrlHub
}
#gavdcodeend 12

#gavdcodebegin 13
Function SpPsPnpRemoveSiteFromHubSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	Remove-PnPHubSiteAssociation -Site $fullSiteUrl
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsPnpGetAdminsInSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Get-PnPSiteCollectionAdmin 
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsPnpAddAdminsToSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Add-PnPSiteCollectionAdmin -Owners "domain@domain.onmicrosoft.com"
}
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsPnpRemoveAdminsFromSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Remove-PnPSiteCollectionAdmin -Owners "domain@user.onmicrosoft.com"
}
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsPnpCreateWebInSiteCollection()  #*** LEGACY CODE *** 
{
	New-PnPWeb -Title "NewWebSiteModernPsPnP" `
			   -Url "NewWebSiteModernPsPnP" `
			   -Description "NewWebSiteModernPsPnP Description" `
			   -Locale "1033" `
			   -Template "STS#3"
}
#gavdcodeend 17

#gavdcodebegin 18
Function SpPsPnpGetOneWebInSiteCollection()  #*** LEGACY CODE *** 
{
	Get-PnPWeb
}
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsPnpGetWebsInSiteCollection()  #*** LEGACY CODE *** 
{
	Get-PnPSubWebs -Recurse
}
#gavdcodeend 19

#gavdcodebegin 20
Function SpPsPnpUpdateOneWebInSiteCollection()  #*** LEGACY CODE *** 
{
	$fullSiteUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsPnP"
	LoginPsPnP $fullSiteUrl

	Set-PnPWeb -Description "NewWebSiteModernPsPnP Description Updated"
}
#gavdcodeend 20

#gavdcodebegin 21
Function SpPsPnpAddPermissionsInWebInSiteCollection()  #*** LEGACY CODE *** 
{
	Set-PnPWebPermission -Url "NewWebSiteModernPsPnP" `
						 -User 'user@domain.onmicrosoft.com' `
						 -AddRole 'Contribute'
}
#gavdcodeend 21

#gavdcodebegin 22
Function SpPsPnpRemoveOneWebFromSiteCollection()  #*** LEGACY CODE *** 
{
	Remove-PnPWeb -Url "NewWebSiteModernPsPnP"
}
#gavdcodeend 22

#----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$spCtx = LoginPsPnP

#SpPsPnpCreateOneSiteCollection
#SpPsPnpCreateOneSiteCollection01
#SpPsPnpGetAllSiteCollections
#SpPsPnpGetOneSiteCollection
#SpPsPnpGetAllSiteCollectionsFiltered
#SpPsPnpGetHubSiteCollection
#SpPsPnpUpdateOneSiteCollection
#SpPsPnpUpdateOneSiteCollection01
#SpPsPnpDeleteOneSiteCollection
#SpPsPnpRegisterHubSiteCollection
#SpPsPnpUnregisterHubSiteCollection
#SpPsPnpAddSiteToHubSiteCollection
#SpPsPnpRemoveSiteFromHubSiteCollection
#SpPsPnpGetAdminsInSiteCollection
#SpPsPnpAddAdminsToSiteCollection
#SpPsPnpRemoveAdminsFromSiteCollection
#SpPsPnpGrantRightsHubSiteCollection

#SpPsPnpCreateWebInSiteCollection
#SpPsPnpGetOneWebInSiteCollection
#SpPsPnpGetWebsInSiteCollection
#SpPsPnpUpdateOneWebInSiteCollection
#SpPsPnpAddPermissionsInWebInSiteCollection
#SpPsPnpRemoveOneWebFromSiteCollection

Write-Host "Done"