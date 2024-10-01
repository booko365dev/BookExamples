
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsSpPnpPowerShell_LoginWithAccPwDefault
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithAccPw `
					  -Credentials $myCredentials
}

function PsSpPnpPowerShell_LoginWithAccPw($FullSiteUrl)
{
	if($fullSiteUrl -ne $null) {
		[SecureString]$securePW = ConvertTo-SecureString -String `
				$configFile.appsettings.UserPw -AsPlainText -Force

		$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
				-argumentlist $configFile.appsettings.UserName, $securePW
		Connect-PnPOnline -Url $FullSiteUrl `
						  -ClientId $configFile.appsettings.ClientIdWithAccPw `
						  -Credentials $myCredentials
	}
}

function PsSpPnpPowerShell_LoginWithInteraction
{
	# Using user interaction and the Azure AD PnP App Registration (Delegated)
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}

function PsSpPnpPowerShell_LoginWithCertificate
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

function PsSpPnpPowerShell_LoginWithCertificateBase64
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
function PsSpPnpPowerShell_CreateOneSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP"
	New-PnPSite -Type CommunicationSite `
				-Title "NewSiteCollectionModernPsPnP" `
				-Url $fullSiteUrl `
				-SiteDesign "Showcase"

	Disconnect-PnPOnline
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpPnpPowerShell_CreateOneSiteCollection01
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
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
function PsSpPnpPowerShell_GetAllSiteCollections
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	Get-PnPTenantSite

	Disconnect-PnPOnline
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpPnpPowerShell_GetOneSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	Get-PnPSite

	Disconnect-PnPOnline
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpPnpPowerShell_GetAllSiteCollectionsFiltered
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	Get-PnPTenantSite -Template "SITEPAGEPUBLISHING#0" -Detailed

	Disconnect-PnPOnline
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpPnpPowerShell_GetHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	Get-PnPHubSite

	Disconnect-PnPOnline
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpPnpPowerShell_UpdateOneSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrl
	Set-PnPSite -CommentsOnSitePagesDisabled $true

	Disconnect-PnPOnline
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpPnpPowerShell_UpdateOneSiteCollection01
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrl
	Set-PnPTenantSite -Url $fullSiteUrl -Title "NewSiteCollModernPsPnP01_Updated"

	Disconnect-PnPOnline
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpPnpPowerShell_DeleteOneSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSitesWrite.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsPnP01"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrl
	Remove-PnPTenantSite -Url $fullSiteUrl -Force -SkipRecycleBin

	Disconnect-PnPOnline
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpPnpPowerShell_RegisterHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrlHub
	Register-PnPHubSite -Site $fullSiteUrlHub

	Disconnect-PnPOnline
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpPnpPowerShell_UnregisterHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrlHub
	Unregister-PnPHubSite -Site $fullSiteUrlHub

	Disconnect-PnPOnline
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpPnpPowerShell_AddSiteToHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrlHub = $configFile.appsettings.SiteBaseUrl + "/sites/NewHubSite"
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrlHub
	Add-PnPHubSiteAssociation -Site $fullSiteUrl -HubSite $fullSiteUrlHub

	Disconnect-PnPOnline
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpPnpPowerShell_RemoveSiteFromHubSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrl
	Remove-PnPHubSiteAssociation -Site $fullSiteUrl

	Disconnect-PnPOnline
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpPnpPowerShell_GetAdminsInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrl
	Get-PnPSiteCollectionAdmin

	Disconnect-PnPOnline
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpPnpPowerShell_AddAdminsToSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrl
	Add-PnPSiteCollectionAdmin -Owners "user@domain.OnMicrosoft.com"

	Disconnect-PnPOnline
}
#gavdcodeend 015

#gavdcodebegin 016
function PsSpPnpPowerShell_RemoveAdminsFromSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/OneSite"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrl
	Remove-PnPSiteCollectionAdmin -Owners "user@domain.OnMicrosoft.com"

	Disconnect-PnPOnline
}
#gavdcodeend 016

#gavdcodebegin 017
function PsSpPnpPowerShell_CreateWebInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	New-PnPWeb -Title "NewWebSiteModernPsPnP" `
			   -Url "NewWebSiteModernPsPnP" `
			   -Description "NewWebSiteModernPsPnP Description" `
			   -Locale "1033" `
			   -Template "STS#3"

	Disconnect-PnPOnline
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpPnpPowerShell_GetOneWebInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	Get-PnPWeb

	Disconnect-PnPOnline
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpPnpPowerShell_GetWebsInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	Get-PnPSubWeb -Recurse

	Disconnect-PnPOnline
}
#gavdcodeend 019

#gavdcodebegin 020
function PsSpPnpPowerShell_UpdateOneWebInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	$fullSiteUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsPnP"
	PsSpPnpPowerShell_LoginWithAccPw $fullSiteUrl
	Set-PnPWeb -Description "NewWebSiteModernPsPnP Description Updated"

	Disconnect-PnPOnline
}
#gavdcodeend 020

#gavdcodebegin 021
function PsSpPnpPowerShell_AddPermissionsInWebInSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	Set-PnPWebPermission -Identity "NewWebSiteModernPsPnP" `
						 -User "AdeleV@M365x17625301.OnMicrosoft.com" `
						 -AddRole "Contribute"

	Disconnect-PnPOnline
}
#gavdcodeend 021

#gavdcodebegin 022
function PsSpPnpPowerShell_RemoveOneWebFromSiteCollection
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	
	Remove-PnPWeb -Identity "NewWebSiteModernPsPnP"

	Disconnect-PnPOnline
}
#gavdcodeend 022

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 022 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using PnP PowerShell --------
#PsSpPnpPowerShell_CreateOneSiteCollection
#PsSpPnpPowerShell_CreateOneSiteCollection01
#PsSpPnpPowerShell_GetAllSiteCollections
#PsSpPnpPowerShell_GetOneSiteCollection
#PsSpPnpPowerShell_GetAllSiteCollectionsFiltered
#PsSpPnpPowerShell_GetHubSiteCollection
#PsSpPnpPowerShell_UpdateOneSiteCollection
#PsSpPnpPowerShell_UpdateOneSiteCollection01
#PsSpPnpPowerShell_DeleteOneSiteCollection
#PsSpPnpPowerShell_RegisterHubSiteCollection
#PsSpPnpPowerShell_UnregisterHubSiteCollection
#PsSpPnpPowerShell_AddSiteToHubSiteCollection
#PsSpPnpPowerShell_RemoveSiteFromHubSiteCollection
#PsSpPnpPowerShell_GetAdminsInSiteCollection
#PsSpPnpPowerShell_AddAdminsToSiteCollection
#PsSpPnpPowerShell_RemoveAdminsFromSiteCollection
#PsSpPnpPowerShell_CreateWebInSiteCollection
#PsSpPnpPowerShell_GetOneWebInSiteCollection
#PsSpPnpPowerShell_GetWebsInSiteCollection
#PsSpPnpPowerShell_UpdateOneWebInSiteCollection
#PsSpPnpPowerShell_AddPermissionsInWebInSiteCollection
#PsSpPnpPowerShell_RemoveOneWebFromSiteCollection

Write-Host "Done" 
