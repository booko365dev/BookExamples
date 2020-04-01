
#gavdcodebegin 01
Function LoginPsPowerPlatform()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.ppUserPw -AsPlainText -Force

	Add-PowerAppsAccount -Username $configFile.appsettings.ppUserName -Password $securePW
}
#gavdcodeend 01

#----------------------------------------------------------------------------------------

#gavdcodebegin 02
Function PowerAppPsAdminEnumerateApps()
{
	Get-AdminPowerApp
}
#gavdcodeend 02

#gavdcodebegin 03
Function PowerAppPsAdminFindOneApps()
{
	Get-AdminPowerApp "NameApp"
}
#gavdcodeend 03

#gavdcodebegin 04
Function PowerAppsPsAdminUserDetails()
{
	Get-AdminPowerAppsUserDetails `
		-OutputFilePath "C:\Temporary\UsersPA.json" `
		–UserPrincipalName "user@domain.onmicrosoft.com"
}
#gavdcodeend 04

#gavdcodebegin 05
Function PowerAppsPsAdminSetOwner()
{
	Set-AdminPowerAppOwner `
		–AppName "01d96b0e-f371-4ced-91c4-bc53acb5dbcf" `
		-AppOwner "092b1237-a428-45a7-b76b-310fdd6e7246" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0"
}
#gavdcodeend 05

#gavdcodebegin 06
Function PowerAppsPsAdminSetFeatured()
{
	Get-AdminPowerApp "NameApp" | Set-AdminPowerAppAsFeatured
}
#gavdcodeend 06

#gavdcodebegin 07
Function PowerAppsPsAdminSetHero()
{
	Get-AdminPowerApp "NameApp" | Set-AdminPowerAppAsHero
}
#gavdcodeend 07

#gavdcodebegin 09
Function PowerAppsPsAdminDeleteFeatured()
{
	Get-AdminPowerApp "NameApp" | Clear-AdminPowerAppAsFeatured
}
#gavdcodeend 09

#gavdcodebegin 10
Function PowerAppsPsAdminDeleteHero()
{
	Get-AdminPowerApp "NameApp" | Clear-AdminPowerAppAsHero
}
#gavdcodeend 10

#gavdcodebegin 08
Function PowerAppsPsAdminDeleteApp()
{
	Remove-AdminPowerApp `
		–AppName "01d96b0e-f371-4ced-91c4-bc53acb5dbcf" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0"
}
#gavdcodeend 08

#gavdcodebegin 11
Function PowerAppsPsAdminFindRoles()
{
	Get-AdminPowerAppRoleAssignment `
		–UserId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 11

#gavdcodebegin 12
Function PowerAppsPsAdminAddRoles()
{
	Set-AdminPowerAppRoleAssignment `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		-RoleName CanEdit `
		-PrincipalType User `
		-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 12

#gavdcodebegin 13
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
#gavdcodeend 13

#gavdcodebegin 14
Function PowerAppsPsMakerEnumerateEnvironments()
{
	Get-PowerAppEnvironment
}
#gavdcodeend 14

#gavdcodebegin 15
Function PowerAppsPsMakerEnumerateApps()
{
	Get-PowerApp
}
#gavdcodeend 15

#gavdcodebegin 16
Function PowerAppsPsMakerSetDisplayName()
{
	Set-PowerAppDisplayName `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-AppDisplayName "NameChangedApp"
}
#gavdcodeend 16

#gavdcodebegin 17
Function PowerAppsPsMakerGetNotifications()
{
	Get-PowerAppsNotification
}
#gavdcodeend 17

#gavdcodebegin 18
Function PowerAppsPsMakerPublishApp()
{
	Publish-PowerApp `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd"
}
#gavdcodeend 18

#gavdcodebegin 19
Function PowerAppsPsMakerEnumerateVersions()
{
	Get-PowerAppVersion `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505"
}
#gavdcodeend 19

#gavdcodebegin 20
Function PowerAppsPsMakerRestoreVersion()
{
	Restore-PowerAppVersion `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505" `
		-AppVersionName "20191215T131114Z"
}
#gavdcodeend 20

#gavdcodebegin 21
Function PowerAppsPsMakerDeleteApp()
{
	Remove-PowerApp `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505"
}
#gavdcodeend 21

#gavdcodebegin 22
Function PowerAppsPsMakerFindRoles()
{
	Get-PowerAppRoleAssignment `
		–AppName "c7965df9-a921-4a23-a21d-02ff19fca82d"
}
#gavdcodeend 22

#gavdcodebegin 23
Function PowerAppsPsMakerAddRoles()
{
	Set-PowerAppRoleAssignment `
		-AppName "c7965df9-a921-4a23-a21d-02ff19fca82d" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		-RoleName CanEdit `
		-PrincipalType User `
		-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 23

#gavdcodebegin 24
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
#gavdcodeend 24

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
