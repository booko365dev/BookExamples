Function LoginPsSPO()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.spAdminUrl -Credential $myCredentials
}

#----------------------------------------------------------------------------------------

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

Function SpPsSpoTestOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Test-SPOSite $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoRepairOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Repair-SPOSite $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoGetTemplates()
{
	Get-SPOWebTemplate
	Disconnect-SPOService
}

Function SpPsSpoGetSiteCollections()
{
	Get-SPOSite
	Disconnect-SPOService
}

Function SpPsSpoUpdateOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOSite -Identity $fullSiteUrl -Title "NewSiteCollModernPsSPO Updated"
	Disconnect-SPOService
}

Function SpPsSpoDeleteOneSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOSite -Identity $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoEnumereDeletedOneSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoRestoreOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Restore-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoRemoveDeletedOneSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPODeletedSite -Identity $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoRegisterHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Register-SPOHubSite -Site $fullSiteUrl -Principals $null
	Disconnect-SPOService
}

Function SpPsSpoGetHubSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOHubSite
	Disconnect-SPOService
}

Function SpPsSpoUpdateHubSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOHubSite -Identity $fullSiteUrl -Description "NewSiteCollModernPsSPO Descr."
	Disconnect-SPOService
}

Function SpPsSpoSetSiteInHubSiteCollections()
{
	$fullHubSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/modernsitecoll"
	Add-SPOHubSiteAssociation -HubSite $fullHubSiteUrl `
							  -Site $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoRemoveSiteFromHubSiteCollections()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/modernsitecoll"
	Remove-SPOHubSiteAssociation -Site $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoUnregisterHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Unregister-SPOHubSite -Identity $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoUnregisterHubSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Unregister-SPOHubSite -Identity $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoAddUserToSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Add-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com" `
				-Group "NewSiteCollModernPsSPO Visitors"
	Disconnect-SPOService
}

Function SpPsSpoGetAllUsersInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl
	Disconnect-SPOService
}

Function SpPsSpoAllUsersInGroupSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl `
				-Group "NewSiteCollModernPsSPO Visitors "
	Disconnect-SPOService
}

Function SpPsSpoOneUserInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com"
	Disconnect-SPOService
}

Function SpPsSpoSetUserAsAdminInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOUser -Site $fullSiteUrl `
				-LoginName "user@domain.onmicrosoft.com" `
				-IsSiteCollectionAdmin $true
	Disconnect-SPOService
}

Function SpPsSpoRemoveOneUserFromSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOUser -Site $fullSiteUrl `
				   -LoginName "user@domain.onmicrosoft.com"
	Disconnect-SPOService
}

Function SpPsSpoAddSecurityGroupToSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	New-SPOSiteGroup -Site $fullSiteUrl `
				     -Group "New SPO Group" `
					 -PermissionLevels "Design"
	Disconnect-SPOService
}

Function SpPsSpoGetSecurityGroupsInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Get-SPOSiteGroup -Site $fullSiteUrl `
				     -Group "New SPO Group"
	Disconnect-SPOService
}

Function SpPsSpoUpdateSecurityGroupInSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Set-SPOSiteGroup -Site $fullSiteUrl `
				     -Identity "New SPO Group" `
					 -PermissionLevelsToRemove "Design" `
					 -PermissionLevelsToAdd "Full Control"
	Disconnect-SPOService
}

Function SpPsSpoRemoveOneSecurityGroupFromSiteCollection()
{
	$fullSiteUrl = $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollModernPsSPO"
	Remove-SPOSiteGroup -Site $fullSiteUrl `
				        -Identity "New SPO Group" 
	Disconnect-SPOService
}

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

