Function LoginPsPnP()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.spUrl -Credentials $myCredentials
}

Function LoginPsPnP($fullSiteUrl)
{
	if($fullSiteUrl -ne $null) {
		[SecureString]$securePW = ConvertTo-SecureString -String `
				$configFile.appsettings.spUserPw -AsPlainText -Force

		$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
				-argumentlist $configFile.appsettings.spUserName, $securePW
		Connect-PnPOnline -Url $fullSiteUrl -Credentials $myCredentials
	}
}

#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsPnpCreateOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP"
	New-PnPSite -Type CommunicationSite `
				-Title "NewSiteCollectionModernPsPnP" `
				-Url $fullSiteUrl `
				-SiteDesign "Showcase"
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsPnpCreateOneSiteCollection01()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	New-PnPTenantSite -Title "NewSiteCollModernPsPnP01" `
					  -Url $fullSiteUrl `
					  -Owner $configFile.appsettings.spUserName `
					  -Template STS#3 `
					  -TimeZone 4
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsPnpGetAllSiteCollections()
{
	Get-PnPTenantSite
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsPnpGetOneSiteCollection()
{
	Get-PnPSite
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsPnpGetAllSiteCollectionsFiltered()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP"
	Get-PnPTenantSite -Template "SITEPAGEPUBLISHING#0" -Detailed
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsPnpGetHubSiteCollection()
{
	Get-PnPHubSite
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsPnpUpdateOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	LoginPsPnP $fullSiteUrl
	Set-PnPSite -CommentsOnSitePagesDisabled
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsPnpUpdateOneSiteCollection01()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	Set-PnPTenantSite -Url $fullSiteUrl -Title "NewSiteCollModernPsPnP01_Updated"
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsPnpDeleteOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	Remove-PnPTenantSite -Url $fullSiteUrl -Force -SkipRecycleBin
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsPnpRegisterHubSiteCollection()
{
	$fullSiteUrlHub = $configFile.appsettings.spBaseUrl + "/sites/NewHubSite"
	Register-PnPHubSite -Site $fullSiteUrlHub
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsPnpUnregisterHubSiteCollection()
{
	$fullSiteUrlHub = $configFile.appsettings.spBaseUrl + "/sites/NewHubSite"
	Unregister-PnPHubSite -Site $fullSiteUrlHub
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsPnpAddSiteToHubSiteCollection()
{
	$fullSiteUrlHub = $configFile.appsettings.spBaseUrl + "/sites/NewHubSite"
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/OneSite"
	Add-PnPHubSiteAssociation -Site $fullSiteUrl -HubSite $fullSiteUrlHub
}
#gavdcodeend 12

#gavdcodebegin 13
Function SpPsPnpRemoveSiteFromHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/OneSite"
	Remove-PnPHubSiteAssociation -Site $fullSiteUrl
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsPnpGetAdminsInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Get-PnPSiteCollectionAdmin 
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsPnpAddAdminsToSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Add-PnPSiteCollectionAdmin -Owners "domain@domain.onmicrosoft.com"
}
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsPnpRemoveAdminsFromSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Remove-PnPSiteCollectionAdmin -Owners "domain@user.onmicrosoft.com"
}
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsPnpCreateWebInSiteCollection()
{
	New-PnPWeb -Title "NewWebSiteModernPsPnP" `
			   -Url "NewWebSiteModernPsPnP" `
			   -Description "NewWebSiteModernPsPnP Description" `
			   -Locale "1033" `
			   -Template "STS#3"
}
#gavdcodeend 17

#gavdcodebegin 18
Function SpPsPnpGetOneWebInSiteCollection()
{
	Get-PnPWeb
}
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsPnpGetWebsInSiteCollection()
{
	Get-PnPSubWebs -Recurse
}
#gavdcodeend 19

#gavdcodebegin 20
Function SpPsPnpUpdateOneWebInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsPnP"
	LoginPsPnP $fullSiteUrl

	Set-PnPWeb -Description "NewWebSiteModernPsPnP Description Updated"
}
#gavdcodeend 20

#gavdcodebegin 21
Function SpPsPnpAddPermissionsInWebInSiteCollection()
{
	Set-PnPWebPermission -Url "NewWebSiteModernPsPnP" `
						 -User 'user@domain.onmicrosoft.com' `
						 -AddRole 'Contribute'
}
#gavdcodeend 21

#gavdcodebegin 22
Function SpPsPnpRemoveOneWebFromSiteCollection()
{
	Remove-PnPWeb -Url "NewWebSiteModernPsPnP"
}
#gavdcodeend 22

#----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

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