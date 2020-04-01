Function LoginPsSPO()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.spAdminUrl -Credential $myCredentials
}

#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsSpoCreateOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	New-SPOSite -Url $fullSiteUrl `
				-Title "NewSiteCollModernPsSPO" `
				-Owner $configFile.appsettings.spUserName `
				-Template "STS#3" `
				-LocaleID "1033" `
				-StorageQuota "1000" `
				-CompatibilityLevel "15" `
				-TimeZoneId "13"
	Disconnect-SPOService
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsSpoTestOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Test-SPOSite $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsSpoRepairOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Repair-SPOSite $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsSpoGetTemplates()
{
	Get-SPOWebTemplate
	Disconnect-SPOService
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsSpoGetSiteCollections()
{
	Get-SPOSite
	Disconnect-SPOService
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsSpoUpdateOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOSite -Identity $fullSiteUrl -Title "NewSiteCollModernPsSPO Updated"
	Disconnect-SPOService
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsSpoDeleteOneSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsSpoEnumereDeletedOneSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsSpoRestoreOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Restore-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsSpoRemoveDeletedOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsSpoRegisterHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Register-SPOHubSite -Site $fullSiteUrl -Principals $null
	Disconnect-SPOService
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsSpoGetHubSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOHubSite
	Disconnect-SPOService
}
#gavdcodeend 12

#gavdcodebegin 13
Function SpPsSpoUpdateHubSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOHubSite -Identity $fullSiteUrl -Description "NewSiteCollModernPsSPO Descr."
	Disconnect-SPOService
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsSpoSetSiteInHubSiteCollections()
{
	$fullHubSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/modernsitecoll"
	Add-SPOHubSiteAssociation -HubSite $fullHubSiteUrl `
							  -Site $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsSpoRemoveSiteFromHubSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/modernsitecoll"
	Remove-SPOHubSiteAssociation -Site $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsSpoUnregisterHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Unregister-SPOHubSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 16

#gavdcodebegin 16
Function SpPsSpoUnregisterHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Unregister-SPOHubSite -Identity $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsSpoAddUserToSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Add-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com" `
				-Group "NewSiteCollModernPsSPO Visitors"
	Disconnect-SPOService
}
#gavdcodeend 17

#gavdcodebegin 18
Function SpPsSpoGetAllUsersInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl
	Disconnect-SPOService
}
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsSpoAllUsersInGroupSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl `
				-Group "NewSiteCollModernPsSPO Visitors "
	Disconnect-SPOService
}
#gavdcodeend 19

#gavdcodebegin 20
Function SpPsSpoOneUserInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com"
	Disconnect-SPOService
}
#gavdcodeend 20

#gavdcodebegin 21
Function SpPsSpoSetUserAsAdminInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com" `
				-IsSiteCollectionAdmin $true
	Disconnect-SPOService
}
#gavdcodeend 21

#gavdcodebegin 22
Function SpPsSpoRemoveOneUserFromSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOUser -Site $fullSiteUrl `
				   -LoginName "user@domain.onmicrosoft.com"
	Disconnect-SPOService
}
#gavdcodeend 22

#gavdcodebegin 23
Function SpPsSpoAddSecurityGroupToSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	New-SPOSiteGroup -Site $fullSiteUrl `
				     -Group "New SPO Group" `
					 -PermissionLevels "Design"
	Disconnect-SPOService
}
#gavdcodeend 23

#gavdcodebegin 24
Function SpPsSpoGetSecurityGroupsInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOSiteGroup -Site $fullSiteUrl `
				     -Group "New SPO Group"
	Disconnect-SPOService
}
#gavdcodeend 24

#gavdcodebegin 25
Function SpPsSpoUpdateSecurityGroupInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOSiteGroup -Site $fullSiteUrl `
				     -Identity "New SPO Group" `
					 -PermissionLevelsToRemove "Design" `
					 -PermissionLevelsToAdd "Full Control"
	Disconnect-SPOService
}
#gavdcodeend 25

#gavdcodebegin 25
Function SpPsSpoRemoveOneSecurityGroupFromSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOSiteGroup -Site $fullSiteUrl `
				        -Identity "New SPO Group" 
	Disconnect-SPOService
}
#gavdcodeend 25

#----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

LoginPsSPO

#SpPsSpoCreateOneSiteCollection
#SpPsSpoTestOneSiteCollection
#SpPsSpoRepairOneSiteCollection
#SpPsSpoGetTemplates
#SpPsSpoGetSiteCollections
#SpPsSpoUpdateOneSiteCollection
#SpPsSpoDeleteOneSiteCollection
#SpPsSpoEnumereDeletedOneSiteCollections
#SpPsSpoRestoreOneSiteCollection
#SpPsSpoRemoveDeletedOneSiteCollection
#SpPsSpoRegisterHubSiteCollection
#SpPsSpoGetHubSiteCollections
#SpPsSpoUpdateHubSiteCollections
#SpPsSpoSetSiteInHubSiteCollections
#SpPsSpoRemoveSiteFromHubSiteCollections
#SpPsSpoUnregisterHubSiteCollection
#SpPsSpoAddUserToSiteCollection
#SpPsSpoGetAllUsersInSiteCollection
#SpPsSpoAllUsersInGroupSiteCollection
#SpPsSpoOneUserInSiteCollection
#SpPsSpoSetUserAsAdminInSiteCollection
#SpPsSpoRemoveOneUserFromSiteCollection
#SpPsSpoAddSecurityGroupToSiteCollection
#SpPsSpoGetSecurityGroupsInSiteCollection
#SpPsSpoUpdateSecurityGroupInSiteCollection
#SpPsSpoRemoveOneSecurityGroupFromSiteCollection

Write-Host "Done"
