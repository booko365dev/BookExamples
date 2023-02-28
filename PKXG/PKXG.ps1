Function LoginPsSPO()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.SiteAdminUrl -Credential $myCredentials
}

#----------------------------------------------------------------------------------------

#gavdcodebegin 001
Function SpPsSpo_CreateOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	New-SPOSite -Url $fullSiteUrl `
				-Title "NewSiteCollModernPsSPO" `
				-Owner $configFile.appsettings.UserName `
				-Template "STS#3" `
				-LocaleID "1033" `
				-StorageQuota "1000" `
				-CompatibilityLevel "15" `
				-TimeZoneId "13"
	Disconnect-SPOService
}
#gavdcodeend 001

#gavdcodebegin 002
Function SpPsSpo_TestOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Test-SPOSite $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 002

#gavdcodebegin 003
Function SpPsSpo_RepairOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Repair-SPOSite $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 003

#gavdcodebegin 004
Function SpPsSpo_GetTemplates()
{
	Get-SPOWebTemplate
	Disconnect-SPOService
}
#gavdcodeend 004

#gavdcodebegin 005
Function SpPsSpo_GetSiteCollections()
{
	Get-SPOSite
	Disconnect-SPOService
}
#gavdcodeend 005

#gavdcodebegin 006
Function SpPsSpo_UpdateOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOSite -Identity $fullSiteUrl -Title "NewSiteCollModernPsSPO Updated"
	Disconnect-SPOService
}
#gavdcodeend 006

#gavdcodebegin 007
Function SpPsSpo_DeleteOneSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 007

#gavdcodebegin 008
Function SpPsSpo_EnumereDeletedOneSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 008

#gavdcodebegin 009
Function SpPsSpo_RestoreOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Restore-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 009

#gavdcodebegin 010
Function SpPsSpo_RemoveDeletedOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 010

#gavdcodebegin 011
Function SpPsSpo_RegisterHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Register-SPOHubSite -Site $fullSiteUrl -Principals $null
	Disconnect-SPOService
}
#gavdcodeend 011

#gavdcodebegin 012
Function SpPsSpo_GetHubSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOHubSite
	Disconnect-SPOService
}
#gavdcodeend 012

#gavdcodebegin 013
Function SpPsSpo_UpdateHubSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOHubSite -Identity $fullSiteUrl -Description "NewSiteCollModernPsSPO Descr."
	Disconnect-SPOService
}
#gavdcodeend 013

#gavdcodebegin 014
Function SpPsSpo_SetSiteInHubSiteCollections()
{
	$fullHubSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/modernsitecoll"
	Add-SPOHubSiteAssociation -HubSite $fullHubSiteUrl `
							  -Site $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 014

#gavdcodebegin 015
Function SpPsSpo_RemoveSiteFromHubSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/modernsitecoll"
	Remove-SPOHubSiteAssociation -Site $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 015

#gavdcodebegin 016
Function SpPsSpo_UnregisterHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Unregister-SPOHubSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 016

#gavdcodebegin 017
Function SpPsSpo_AddUserToSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Add-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com" `
				-Group "NewSiteCollModernPsSPO Visitors"
	Disconnect-SPOService
}
#gavdcodeend 017

#gavdcodebegin 018
Function SpPsSpo_GetAllUsersInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 018

#gavdcodebegin 019
Function SpPsSpo_AllUsersInGroupSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl `
				-Group "NewSiteCollModernPsSPO Visitors "
	Disconnect-SPOService
}
#gavdcodeend 019

#gavdcodebegin 020
Function SpPsSpo_OneUserInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com"
	Disconnect-SPOService
}
#gavdcodeend 020

#gavdcodebegin 021
Function SpPsSpo_SetUserAsAdminInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com" `
				-IsSiteCollectionAdmin $true
	Disconnect-SPOService
}
#gavdcodeend 021

#gavdcodebegin 022
Function SpPsSpo_RemoveOneUserFromSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOUser -Site $fullSiteUrl `
				   -LoginName "user@domain.onmicrosoft.com"
	Disconnect-SPOService
}
#gavdcodeend 022

#gavdcodebegin 023
Function SpPsSpo_AddSecurityGroupToSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	New-SPOSiteGroup -Site $fullSiteUrl `
				     -Group "New SPO Group" `
					 -PermissionLevels "Design"
	Disconnect-SPOService
}
#gavdcodeend 023

#gavdcodebegin 024
Function SpPsSpo_GetSecurityGroupsInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOSiteGroup -Site $fullSiteUrl `
				     -Group "New SPO Group"
	Disconnect-SPOService
}
#gavdcodeend 024

#gavdcodebegin 025
Function SpPsSpo_UpdateSecurityGroupInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOSiteGroup -Site $fullSiteUrl `
				     -Identity "New SPO Group" `
					 -PermissionLevelsToRemove "Design" `
					 -PermissionLevelsToAdd "Full Control"
	Disconnect-SPOService
}
#gavdcodeend 025

#gavdcodebegin 026
Function SpPsSpo_RemoveOneSecurityGroupFromSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOSiteGroup -Site $fullSiteUrl `
				        -Identity "New SPO Group" 
	Disconnect-SPOService
}
#gavdcodeend 026

#----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

LoginPsSPO

#SpPsSpo_CreateOneSiteCollection
#SpPsSpo_TestOneSiteCollection
#SpPsSpo_RepairOneSiteCollection
#SpPsSpo_GetTemplates
#SpPsSpo_GetSiteCollections
#SpPsSpo_UpdateOneSiteCollection
#SpPsSpo_DeleteOneSiteCollection
#SpPsSpo_EnumereDeletedOneSiteCollections
#SpPsSpo_RestoreOneSiteCollection
#SpPsSpo_RemoveDeletedOneSiteCollection
#SpPsSpo_RegisterHubSiteCollection
#SpPsSpo_GetHubSiteCollections
#SpPsSpo_UpdateHubSiteCollections
#SpPsSpo_SetSiteInHubSiteCollections
#SpPsSpo_RemoveSiteFromHubSiteCollections
#SpPsSpo_UnregisterHubSiteCollection
#SpPsSpo_AddUserToSiteCollection
#SpPsSpo_GetAllUsersInSiteCollection
#SpPsSpo_AllUsersInGroupSiteCollection
#SpPsSpo_OneUserInSiteCollection
#SpPsSpo_SetUserAsAdminInSiteCollection
#SpPsSpo_RemoveOneUserFromSiteCollection
#SpPsSpo_AddSecurityGroupToSiteCollection
#SpPsSpo_GetSecurityGroupsInSiteCollection
#SpPsSpo_UpdateSecurityGroupInSiteCollection
#SpPsSpo_RemoveOneSecurityGroupFromSiteCollection

Write-Host "Done"
