
Function LoginPsPowerPlatform()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.ppUserPw -AsPlainText -Force

	Add-PowerAppsAccount -Username $configFile.appsettings.ppUserName -Password $securePW
}

#----------------------------------------------------------------------------------------

Function PowerAppPsAdminEnumerateApps()
{
	Get-AdminPowerApp
}

Function PowerAppPsAdminFindOneApps()
{
	Get-AdminPowerApp "NameApp"
}

Function PowerAppsPsAdminUserDetails()
{
	Get-AdminPowerAppsUserDetails `
		-OutputFilePath "C:\Temporary\UsersPA.json" `
		–UserPrincipalName "user@domain.onmicrosoft.com"
}

Function PowerAppsPsAdminSetOwner()
{
	Set-AdminPowerAppOwner `
		–AppName "01d96b0e-f371-4ced-91c4-bc53acb5dbcf" `
		-AppOwner "092b1237-a428-45a7-b76b-310fdd6e7246" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0"
}

Function PowerAppsPsAdminSetFeatured()
{
	Get-AdminPowerApp "NameApp" | Set-AdminPowerAppAsFeatured
}

Function PowerAppsPsAdminSetHero()
{
	Get-AdminPowerApp "NameApp" | Set-AdminPowerAppAsHero
}

Function PowerAppsPsAdminDeleteFeatured()
{
	Get-AdminPowerApp "NameApp" | Clear-AdminPowerAppAsFeatured
}

Function PowerAppsPsAdminDeleteHero()
{
	Get-AdminPowerApp "NameApp" | Clear-AdminPowerAppAsHero
}

Function PowerAppsPsAdminDeleteApp()
{
	Remove-AdminPowerApp `
		–AppName "01d96b0e-f371-4ced-91c4-bc53acb5dbcf" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0"
}

Function PowerAppsPsAdminFindRoles()
{
	Get-AdminPowerAppRoleAssignment `
		–UserId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}

Function PowerAppsPsAdminAddRoles()
{
	Set-AdminPowerAppRoleAssignment `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		-RoleName CanEdit `
		-PrincipalType User `
		-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}

Function PowerAppsPsAdminDeleteRoles()
{
	$myRoleId = "/providers/Microsoft.PowerApps/scopes/admin/apps/" + 
				"fa014c64-efe7-4301-bea2-9034bb7b51fd/permissions/" + 
				"959ae10e-0015-4948-b602-fbf7fccfe2a3"

	Remove-AdminPowerAppRoleAssignment `
		–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		–AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-RoleId $myRoleId
}

Function PowerAppsPsMakerEnumerateEnvironments()
{
	Get-PowerAppEnvironment
}

Function PowerAppsPsMakerEnumerateApps()
{
	Get-PowerApp
}

Function PowerAppsPsMakerSetDisplayName()
{
	Set-PowerAppDisplayName `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-AppDisplayName "NameChangedApp"
}

Function PowerAppsPsMakerGetNotifications()
{
	Get-PowerAppsNotification
}

Function PowerAppsPsMakerPublishApp()
{
	Publish-PowerApp `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd"
}

Function PowerAppsPsMakerEnumerateVersions()
{
	Get-PowerAppVersion `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505"
}

Function PowerAppsPsMakerRestoreVersion()
{
	Restore-PowerAppVersion `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505" `
		-AppVersionName "20191215T131114Z"
}

Function PowerAppsPsMakerDeleteApp()
{
	Remove-PowerApp `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505"
}

Function PowerAppsPsMakerFindRoles()
{
	Get-PowerAppRoleAssignment `
		–AppName "c7965df9-a921-4a23-a21d-02ff19fca82d"
}

Function PowerAppsPsMakerAddRoles()
{
	Set-PowerAppRoleAssignment `
		-AppName "c7965df9-a921-4a23-a21d-02ff19fca82d" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		-RoleName CanEdit `
		-PrincipalType User `
		-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}

Function PowerAppsPsMakerDeleteRoles()
{
	$myRoleId = "/providers/Microsoft.PowerApps/apps/" + 
				"c7965df9-a921-4a23-a21d-02ff19fca82d/permissions/" + 
				"092b1237-a428-45a7-b76b-310fdd6e7246"

	Remove-PowerAppRoleAssignment `
		–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		–AppName "c7965df9-a921-4a23-a21d-02ff19fca82d" `
		-RoleId $myRoleId
}

#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ppPs.values.config"

LoginPsPowerPlatform

#PowerAppPsAdminEnumerateApps
#PowerAppPsAdminFindOneApps
#PowerAppsPsAdminUserDetails
#PowerAppsPsAdminSetOwner
#PowerAppsPsAdminSetFeatured
#PowerAppsPsAdminSetHero
#PowerAppsPsAdminDeleteApp
#PowerAppsPsAdminDeleteHero
#PowerAppsPsAdminDeleteFeatured
#PowerAppsPsAdminFindRoles
#PowerAppsPsAdminAddRoles
#PowerAppsPsAdminDeleteRoles
#PowerAppsPsMakerEnumerateEnvironments
#PowerAppsPsMakerEnumerateApps
#PowerAppsPsMakerSetDisplayName
#PowerAppsPsMakerGetNotifications
#PowerAppsPsMakerPublishApp
#PowerAppsPsMakerEnumerateVersions
#PowerAppsPsMakerRestoreVersion
#PowerAppsPsMakerDeleteApp
#PowerAppsPsMakerFindRoles
#PowerAppsPsMakerAddRoles
#PowerAppsPsMakerDeleteRoles

Write-Host "Done"  

