
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function LoginPsPnPPowerShellWithAccPwDefault
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

Function LoginPsPnPPowerShellWithAccPw($FullSiteUrl)
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	if($fullSiteUrl -ne $null) {
		[SecureString]$securePW = ConvertTo-SecureString -String `
				$configFile.appsettings.UserPw -AsPlainText -Force

		$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
				-argumentlist $configFile.appsettings.UserName, $securePW
		Connect-PnPOnline -Url $FullSiteUrl -Credentials $myCredentials
	}
}

Function LoginPsPnPPowerShellWithInteraction
{
	# Using user interaction and the Azure AD PnP App Registration (Delegated)
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}

Function LoginPsPnPPowerShellWithCertificate
{
	# Using a Digital Certificate and Azure App Registration (Application)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificatePath "[PathForThePfxCertificateFile]" `
					  -CertificatePassword $securePW
}

Function LoginPsPnPPowerShellWithCertificateBase64
{
	# Using a Digital Certificate and Azure App Registration (Application)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificateBase64Encoded "[Base64EncodedValue]" `
					  -CertificatePassword $securePW
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 01
function SpPsPnpCreateOneSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP"
	New-PnPSite -Type CommunicationSite `
				-Title "NewSiteCollectionModernPsPnP" `
				-Url $fullSiteUrl `
				-SiteDesign "Showcase"

	Disconnect-PnPOnline
}
#gavdcodeend 01

#gavdcodebegin 02
function SpPsPnpCreateOneSiteCollection01
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	New-PnPTenantSite -Title "NewSiteCollectionModernPsPnP01" `
					  -Url $fullSiteUrl `
					  -Owner $configFile.appsettings.UserName `
					  -Template STS#3 `
					  -TimeZone 4

	Disconnect-PnPOnline
}
#gavdcodeend 02

#gavdcodebegin 03
function SpPsPnpGetAllSiteCollections
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPTenantSite

	Disconnect-PnPOnline
}
#gavdcodeend 03

#gavdcodebegin 04
function SpPsPnpGetOneSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPSite

	Disconnect-PnPOnline
}
#gavdcodeend 04

#gavdcodebegin 05
function SpPsPnpGetAllSiteCollectionsFiltered
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPTenantSite -Template "SITEPAGEPUBLISHING#0" -Detailed

	Disconnect-PnPOnline
}
#gavdcodeend 05

#gavdcodebegin 06
function SpPsPnpGetHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPHubSite

	Disconnect-PnPOnline
}
#gavdcodeend 06

#gavdcodebegin 07
function SpPsPnpUpdateOneSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrl
	Set-PnPSite -CommentsOnSitePagesDisabled $true

	Disconnect-PnPOnline
}
#gavdcodeend 07

#gavdcodebegin 08
function SpPsPnpUpdateOneSiteCollection01
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrl
	Set-PnPTenantSite -Url $fullSiteUrl -Title "NewSiteCollModernPsPnP01_Updated"

	Disconnect-PnPOnline
}
#gavdcodeend 08

#gavdcodebegin 09
function SpPsPnpDeleteOneSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSitesWrite.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrl
	Remove-PnPTenantSite -Url $fullSiteUrl -Force -SkipRecycleBin

	Disconnect-PnPOnline
}
#gavdcodeend 09

#gavdcodebegin 10
function SpPsPnpRegisterHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrlHub
	Register-PnPHubSite -Site $fullSiteUrlHub

	Disconnect-PnPOnline
}
#gavdcodeend 10

#gavdcodebegin 11
function SpPsPnpUnregisterHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrlHub
	Unregister-PnPHubSite -Site $fullSiteUrlHub

	Disconnect-PnPOnline
}
#gavdcodeend 11

#gavdcodebegin 12
function SpPsPnpAddSiteToHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrlHub
	Add-PnPHubSiteAssociation -Site $fullSiteUrl -HubSite $fullSiteUrlHub

	Disconnect-PnPOnline
}
#gavdcodeend 12

#gavdcodebegin 13
function SpPsPnpRemoveSiteFromHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrl
	Remove-PnPHubSiteAssociation -Site $fullSiteUrl

	Disconnect-PnPOnline
}
#gavdcodeend 13

#gavdcodebegin 14
function SpPsPnpGetAdminsInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrl
	Get-PnPSiteCollectionAdmin

	Disconnect-PnPOnline
}
#gavdcodeend 14

#gavdcodebegin 15
function SpPsPnpAddAdminsToSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrl
	Add-PnPSiteCollectionAdmin -Owners "user@domain.OnMicrosoft.com"

	Disconnect-PnPOnline
}
#gavdcodeend 15

#gavdcodebegin 16
function SpPsPnpRemoveAdminsFromSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrl
	Remove-PnPSiteCollectionAdmin -Owners "user@domain.OnMicrosoft.com"

	Disconnect-PnPOnline
}
#gavdcodeend 16

#gavdcodebegin 17
function SpPsPnpCreateWebInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	New-PnPWeb -Title "NewWebSiteModernPsPnP" `
			   -Url "NewWebSiteModernPsPnP" `
			   -Description "NewWebSiteModernPsPnP Description" `
			   -Locale "1033" `
			   -Template "STS#3"

	Disconnect-PnPOnline
}
#gavdcodeend 17

#gavdcodebegin 18
function SpPsPnpGetOneWebInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPWeb

	Disconnect-PnPOnline
}
#gavdcodeend 18

#gavdcodebegin 19
function SpPsPnpGetWebsInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPSubWeb -Recurse

	Disconnect-PnPOnline
}
#gavdcodeend 19

#gavdcodebegin 20
function SpPsPnpUpdateOneWebInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsPnP"
	LoginPsPnPPowerShellWithAccPw $fullSiteUrl
	Set-PnPWeb -Description "NewWebSiteModernPsPnP Description Updated"

	Disconnect-PnPOnline
}
#gavdcodeend 20

#gavdcodebegin 21
Function SpPsPnpAddPermissionsInWebInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Set-PnPWebPermission -Identity "NewWebSiteModernPsPnP" `
						 -User "AdeleV@M365x17625301.OnMicrosoft.com" `
						 -AddRole "Contribute"

	Disconnect-PnPOnline
}
#gavdcodeend 21

#gavdcodebegin 22
Function SpPsPnpRemoveOneWebFromSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Remove-PnPWeb -Identity "NewWebSiteModernPsPnP"

	Disconnect-PnPOnline
}
#gavdcodeend 22

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using PnP PowerShell --------
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
#SpPsPnpCreateWebInSiteCollection
#SpPsPnpGetOneWebInSiteCollection
#SpPsPnpGetWebsInSiteCollection
#SpPsPnpUpdateOneWebInSiteCollection
#SpPsPnpAddPermissionsInWebInSiteCollection
#SpPsPnpRemoveOneWebFromSiteCollection

Write-Host "Done" 
