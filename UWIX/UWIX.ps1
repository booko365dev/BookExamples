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

Function SpPsPnpCreateOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP"
	New-PnPSite -Type CommunicationSite `
				-Title "NewSiteCollectionModernPsPnP" `
				-Url $fullSiteUrl `
				-SiteDesign "Showcase"
}

Function SpPsPnpCreateOneSiteCollection01()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	New-PnPTenantSite -Title "NewSiteCollModernPsPnP01" `
					  -Url $fullSiteUrl `
					  -Owner $configFile.appsettings.spUserName `
					  -Template STS#3 `
					  -TimeZone 4
}

Function SpPsPnpGetAllSiteCollections()
{
	Get-PnPTenantSite
}

Function SpPsPnpGetOneSiteCollection()
{
	Get-PnPSite
}

Function SpPsPnpGetAllSiteCollectionsFiltered()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP"
	Get-PnPTenantSite -Template "SITEPAGEPUBLISHING#0" -Detailed
}

Function SpPsPnpGetHubSiteCollection()
{
	Get-PnPHubSite
}

Function SpPsPnpUpdateOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	LoginPsPnP $fullSiteUrl
	Set-PnPSite -CommentsOnSitePagesDisabled
}

Function SpPsPnpUpdateOneSiteCollection01()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	Set-PnPTenantSite -Url $fullSiteUrl -Title "NewSiteCollModernPsPnP01_Updated"
}

Function SpPsPnpDeleteOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	Remove-PnPTenantSite -Url $fullSiteUrl -Force -SkipRecycleBin
}

Function SpPsPnpRegisterHubSiteCollection()
{
	$fullSiteUrlHub = $configFile.appsettings.spBaseUrl + "/sites/NewHubSite"
	Register-PnPHubSite -Site $fullSiteUrlHub
}

Function SpPsPnpUnregisterHubSiteCollection()
{
	$fullSiteUrlHub = $configFile.appsettings.spBaseUrl + "/sites/NewHubSite"
	Unregister-PnPHubSite -Site $fullSiteUrlHub
}

Function SpPsPnpAddSiteToHubSiteCollection()
{
	$fullSiteUrlHub = $configFile.appsettings.spBaseUrl + "/sites/NewHubSite"
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/OneSite"
	Add-PnPHubSiteAssociation -Site $fullSiteUrl -HubSite $fullSiteUrlHub
}

Function SpPsPnpRemoveSiteFromHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/OneSite"
	Remove-PnPHubSiteAssociation -Site $fullSiteUrl
}

Function SpPsPnpGetAdminsInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Get-PnPSiteCollectionAdmin 
}

Function SpPsPnpAddAdminsToSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Add-PnPSiteCollectionAdmin -Owners "domain@domain.onmicrosoft.com"
}

Function SpPsPnpRemoveAdminsFromSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/classicsitecoll"
	LoginPsPnP $fullSiteUrl

	Remove-PnPSiteCollectionAdmin -Owners "domain@user.onmicrosoft.com"
}

Function SpPsPnpCreateWebInSiteCollection()
{
	New-PnPWeb -Title "NewWebSiteModernPsPnP" `
			   -Url "NewWebSiteModernPsPnP" `
			   -Description "NewWebSiteModernPsPnP Description" `
			   -Locale "1033" `
			   -Template "STS#3"
}

Function SpPsPnpGetOneWebInSiteCollection()
{
	Get-PnPWeb
}

Function SpPsPnpGetWebsInSiteCollection()
{
	Get-PnPSubWebs -Recurse
}

Function SpPsPnpUpdateOneWebInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsPnP"
	LoginPsPnP $fullSiteUrl

	Set-PnPWeb -Description "NewWebSiteModernPsPnP Description Updated"
}

Function SpPsPnpAddPermissionsInWebInSiteCollection()
{
	Set-PnPWebPermission -Url "NewWebSiteModernPsPnP" `
						 -User 'user@domain.onmicrosoft.com' `
						 -AddRole 'Contribute'
}

Function SpPsPnpRemoveOneWebFromSiteCollection()
{
	Remove-PnPWeb -Url "NewWebSiteModernPsPnP"
}

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
