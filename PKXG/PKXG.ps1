﻿
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsSpSpo_Login
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.SiteAdminUrl -Credential $myCredentials
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpSpo_CreateOneSiteCollection
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
function PsSpSpo_TestOneSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Test-SPOSite $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpSpo_RepairOneSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Repair-SPOSite $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpSpo_GetTemplates
{
	Get-SPOWebTemplate
	Disconnect-SPOService
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpSpo_GetSiteCollections
{
	Get-SPOSite
	Disconnect-SPOService
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpSpo_UpdateOneSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOSite -Identity $fullSiteUrl -Title "NewSiteCollModernPsSPO Updated"
	Disconnect-SPOService
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpSpo_DeleteOneSiteCollections
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpSpo_EnumereDeletedOneSiteCollections
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpSpo_RestoreOneSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Restore-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpSpo_RemoveDeletedOneSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpSpo_RegisterHubSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Register-SPOHubSite -Site $fullSiteUrl -Principals $null
	Disconnect-SPOService
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpSpo_GetHubSiteCollections
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOHubSite
	Disconnect-SPOService
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpSpo_UpdateHubSiteCollections
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOHubSite -Identity $fullSiteUrl -Description "NewSiteCollModernPsSPO Descr."
	Disconnect-SPOService
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpSpo_SetSiteInHubSiteCollections
{
	$fullHubSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/modernsitecoll"
	Add-SPOHubSiteAssociation -HubSite $fullHubSiteUrl `
							  -Site $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpSpo_RemoveSiteFromHubSiteCollections
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/modernsitecoll"
	Remove-SPOHubSiteAssociation -Site $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 015

#gavdcodebegin 016
function PsSpSpo_UnregisterHubSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Unregister-SPOHubSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 016

#gavdcodebegin 017
function PsSpSpo_AddUserToSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Add-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com" `
				-Group "NewSiteCollModernPsSPO Visitors"
	Disconnect-SPOService
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpSpo_GetAllUsersInSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpSpo_AllUsersInGroupSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl `
				-Group "NewSiteCollModernPsSPO Visitors "
	Disconnect-SPOService
}
#gavdcodeend 019

#gavdcodebegin 020
function PsSpSpo_OneUserInSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com"
	Disconnect-SPOService
}
#gavdcodeend 020

#gavdcodebegin 021
function PsSpSpo_SetUserAsAdminInSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com" `
				-IsSiteCollectionAdmin $true
	Disconnect-SPOService
}
#gavdcodeend 021

#gavdcodebegin 022
function PsSpSpo_RemoveOneUserFromSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOUser -Site $fullSiteUrl `
				   -LoginName "user@domain.onmicrosoft.com"
	Disconnect-SPOService
}
#gavdcodeend 022

#gavdcodebegin 023
function PsSpSpo_AddSecurityGroupToSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	New-SPOSiteGroup -Site $fullSiteUrl `
				     -Group "New SPO Group" `
					 -PermissionLevels "Design"
	Disconnect-SPOService
}
#gavdcodeend 023

#gavdcodebegin 024
function PsSpSpo_GetSecurityGroupsInSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOSiteGroup -Site $fullSiteUrl `
				     -Group "New SPO Group"
	Disconnect-SPOService
}
#gavdcodeend 024

#gavdcodebegin 025
function PsSpSpo_UpdateSecurityGroupInSiteCollection
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
function PsSpSpo_RemoveOneSecurityGroupFromSiteCollection
{
	$fullSiteUrl = $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOSiteGroup -Site $fullSiteUrl `
				        -Identity "New SPO Group" 
	Disconnect-SPOService
}
#gavdcodeend 026

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 026 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

PsSpSpo_Login

#PsSpSpo_CreateOneSiteCollection
#PsSpSpo_TestOneSiteCollection
#PsSpSpo_RepairOneSiteCollection
#PsSpSpo_GetTemplates
#PsSpSpo_GetSiteCollections
#PsSpSpo_UpdateOneSiteCollection
#PsSpSpo_DeleteOneSiteCollection
#PsSpSpo_EnumereDeletedOneSiteCollections
#PsSpSpo_RestoreOneSiteCollection
#PsSpSpo_RemoveDeletedOneSiteCollection
#PsSpSpo_RegisterHubSiteCollection
#PsSpSpo_GetHubSiteCollections
#PsSpSpo_UpdateHubSiteCollections
#PsSpSpo_SetSiteInHubSiteCollections
#PsSpSpo_RemoveSiteFromHubSiteCollections
#PsSpSpo_UnregisterHubSiteCollection
#PsSpSpo_AddUserToSiteCollection
#PsSpSpo_GetAllUsersInSiteCollection
#PsSpSpo_AllUsersInGroupSiteCollection
#PsSpSpo_OneUserInSiteCollection
#PsSpSpo_SetUserAsAdminInSiteCollection
#PsSpSpo_RemoveOneUserFromSiteCollection
#PsSpSpo_AddSecurityGroupToSiteCollection
#PsSpSpo_GetSecurityGroupsInSiteCollection
#PsSpSpo_UpdateSecurityGroupInSiteCollection
#PsSpSpo_RemoveOneSecurityGroupFromSiteCollection

Write-Host "Done"
