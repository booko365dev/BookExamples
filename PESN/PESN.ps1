
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

#gavdcodebegin 001
function SpPsPnpPowerShell_CreateOneSiteCollection
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
#gavdcodeend 001

#gavdcodebegin 002
function SpPsPnpPowerShell_CreateOneSiteCollection01
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
#gavdcodeend 002

#gavdcodebegin 003
function SpPsPnpPowerShell_GetAllSiteCollections
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPTenantSite

	Disconnect-PnPOnline
}
#gavdcodeend 003

#gavdcodebegin 004
function SpPsPnpPowerShell_GetOneSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPSite

	Disconnect-PnPOnline
}
#gavdcodeend 004

#gavdcodebegin 005
function SpPsPnpPowerShell_GetAllSiteCollectionsFiltered
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPTenantSite -Template "SITEPAGEPUBLISHING#0" -Detailed

	Disconnect-PnPOnline
}
#gavdcodeend 005

#gavdcodebegin 006
function SpPsPnpPowerShell_GetHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPHubSite

	Disconnect-PnPOnline
}
#gavdcodeend 006

#gavdcodebegin 007
function SpPsPnpPowerShell_UpdateOneSiteCollection
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
#gavdcodeend 007

#gavdcodebegin 008
function SpPsPnpPowerShell_UpdateOneSiteCollection01
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
#gavdcodeend 008

#gavdcodebegin 009
function SpPsPnpPowerShell_DeleteOneSiteCollection
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
#gavdcodeend 009

#gavdcodebegin 010
function SpPsPnpPowerShell_RegisterHubSiteCollection
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
#gavdcodeend 010

#gavdcodebegin 011
function SpPsPnpPowerShell_UnregisterHubSiteCollection
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
#gavdcodeend 011

#gavdcodebegin 012
function SpPsPnpPowerShell_AddSiteToHubSiteCollection
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
#gavdcodeend 012

#gavdcodebegin 013
function SpPsPnpPowerShell_RemoveSiteFromHubSiteCollection
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
#gavdcodeend 013

#gavdcodebegin 014
function SpPsPnpPowerShell_GetAdminsInSiteCollection
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
#gavdcodeend 014

#gavdcodebegin 015
function SpPsPnpPowerShell_AddAdminsToSiteCollection
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
#gavdcodeend 015

#gavdcodebegin 016
function SpPsPnpPowerShell_RemoveAdminsFromSiteCollection
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
#gavdcodeend 016

#gavdcodebegin 017
function SpPsPnpPowerShell_CreateWebInSiteCollection
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
#gavdcodeend 017

#gavdcodebegin 018
function SpPsPnpPowerShell_GetOneWebInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPWeb

	Disconnect-PnPOnline
}
#gavdcodeend 018

#gavdcodebegin 019
function SpPsPnpPowerShell_GetWebsInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Get-PnPSubWeb -Recurse

	Disconnect-PnPOnline
}
#gavdcodeend 019

#gavdcodebegin 020
function SpPsPnpPowerShell_UpdateOneWebInSiteCollection
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
#gavdcodeend 020

#gavdcodebegin 021
Function SpPsPnpPowerShell_AddPermissionsInWebInSiteCollection
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
#gavdcodeend 021

#gavdcodebegin 022
Function SpPsPnpPowerShell_RemoveOneWebFromSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	Remove-PnPWeb -Identity "NewWebSiteModernPsPnP"

	Disconnect-PnPOnline
}
#gavdcodeend 022

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using PnP PowerShell --------
#SpPsPnpPowerShell_CreateOneSiteCollection
#SpPsPnpPowerShell_CreateOneSiteCollection01
#SpPsPnpPowerShell_GetAllSiteCollections
#SpPsPnpPowerShell_GetOneSiteCollection
#SpPsPnpPowerShell_GetAllSiteCollectionsFiltered
#SpPsPnpPowerShell_GetHubSiteCollection
#SpPsPnpPowerShell_UpdateOneSiteCollection
#SpPsPnpPowerShell_UpdateOneSiteCollection01
#SpPsPnpPowerShell_DeleteOneSiteCollection
#SpPsPnpPowerShell_RegisterHubSiteCollection
#SpPsPnpPowerShell_UnregisterHubSiteCollection
#SpPsPnpPowerShell_AddSiteToHubSiteCollection
#SpPsPnpPowerShell_RemoveSiteFromHubSiteCollection
#SpPsPnpPowerShell_GetAdminsInSiteCollection
#SpPsPnpPowerShell_AddAdminsToSiteCollection
#SpPsPnpPowerShell_RemoveAdminsFromSiteCollection
#SpPsPnpPowerShell_CreateWebInSiteCollection
#SpPsPnpPowerShell_GetOneWebInSiteCollection
#SpPsPnpPowerShell_GetWebsInSiteCollection
#SpPsPnpPowerShell_UpdateOneWebInSiteCollection
#SpPsPnpPowerShell_AddPermissionsInWebInSiteCollection
#SpPsPnpPowerShell_RemoveOneWebFromSiteCollection

Write-Host "Done" 
