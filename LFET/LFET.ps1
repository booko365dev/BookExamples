
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


#gavdcodebegin 001
Function LoginPsPowerPlatform
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	Add-PowerAppsAccount -Username $configFile.appsettings.UserName -Password $securePW
}
#gavdcodeend 001

Function LoginPsCLI
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

Function LoginPsPnPPowerShellWithAccPwDefault
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


##==> Routines for PowerShell Admin and Maker cmdlets

#gavdcodebegin 002
Function PapPsAdmin_EnumerateApps
{
	Get-AdminPowerApp
}
#gavdcodeend 002

#gavdcodebegin 003
Function PapPsAdmin_FindOneApps
{
	Get-AdminPowerApp "a4978531-2218-406c-9158-0b9353334c6d"
}
#gavdcodeend 003

#gavdcodebegin 004
Function PapPsAdmin_UserDetails
{
	Get-AdminPowerAppsUserDetails `
		-OutputFilePath "C:\Temporary\UsersPA.json" `
		–UserPrincipalName "user@domain.onmicrosoft.com"
}
#gavdcodeend 004

#gavdcodebegin 005
Function PapPsAdmin_SetOwner
{
	Set-AdminPowerAppOwner `
		–AppName "01d96b0e-f371-4ced-91c4-bc53acb5dbcf" `
		-AppOwner "092b1237-a428-45a7-b76b-310fdd6e7246" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0"
}
#gavdcodeend 005

#gavdcodebegin 006
Function PapPsAdmin_SetFeatured
{
	Get-AdminPowerApp "NameApp" | Set-AdminPowerAppAsFeatured
}
#gavdcodeend 006

#gavdcodebegin 007
Function PapPsAdmin_SetHero
{
	Get-AdminPowerApp "NameApp" | Set-AdminPowerAppAsHero
}
#gavdcodeend 007

#gavdcodebegin 009
Function PapPsAdmin_DeleteFeatured
{
	Get-AdminPowerApp "NameApp" | Clear-AdminPowerAppAsFeatured
}
#gavdcodeend 009

#gavdcodebegin 010
Function PapPsAdmin_DeleteHero
{
	Get-AdminPowerApp "NameApp" | Clear-AdminPowerAppAsHero
}
#gavdcodeend 010

#gavdcodebegin 008
Function PapPsAdmin_DeleteApp
{
	Remove-AdminPowerApp `
		–AppName "01d96b0e-f371-4ced-91c4-bc53acb5dbcf" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0"
}
#gavdcodeend 008

#gavdcodebegin 011
Function PapPsAdmin_FindRoles
{
	Get-AdminPowerAppRoleAssignment `
		–UserId "acc28fcb-5261-47f8-960b-715d2f98a431"
}
#gavdcodeend 011

#gavdcodebegin 012
Function PapPsAdmin_AddRoles
{
	Set-AdminPowerAppRoleAssignment `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		-RoleName CanEdit `
		-PrincipalType User `
		-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 012

#gavdcodebegin 013
Function PapPsAdmin_DeleteRoles
{
	$myRoleId = "/providers/Microsoft.PowerApps/scopes/admin/apps/" + 
				"fa014c64-efe7-4301-bea2-9034bb7b51fd/permissions/" + 
				"959ae10e-0015-4948-b602-fbf7fccfe2a3"

	Remove-AdminPowerAppRoleAssignment `
		–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		–AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-RoleId $myRoleId
}
#gavdcodeend 013

#gavdcodebegin 035
Function PapPsAdmin_GetDeletedApps
{
	Get-AdminDeletedPowerAppsList `
						-EnvironmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca"
}
#gavdcodeend 035

#gavdcodebegin 036
Function PapPsAdmin_RecoverDeletedApps
{
	Get-AdminRecoverDeletedPowerApp `
						-AppName "a4978531-2218-406c-9158-0b9353334c6d" `
						-EnvironmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca"
}
#gavdcodeend 036

#gavdcodebegin 037
Function PapPsAdmin_GetTenantSettings
{
	Get-TenantSettings
}
#gavdcodeend 037

#gavdcodebegin 038
Function PapPsAdmin_ModifyTenantSettings
{
	$mySettings = Get-TenantSettings 
	$mySettings.disableTrialEnvironmentCreationByNonAdminUsers = $true 
	Set-TenantSettings -RequestBody $mySettings
}
#gavdcodeend 038

#gavdcodebegin 039
Function PapPsAdmin_GetAllEnvironments
{
	Get-AdminPowerAppEnvironment
}
#gavdcodeend 039

#gavdcodebegin 040
Function PapPsAdmin_GetOneEnvironments
{
	Get-AdminPowerAppEnvironment –Default
	Get-AdminPowerAppEnvironment –EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444"
}
#gavdcodeend 040

#gavdcodebegin 041
Function PapPsAdmin_GetAdToken
{
	Get-JwtToken "https://service.powerapps.com/"
}
#gavdcodeend 041

#gavdcodebegin 042
Function PapPsAdmin_SelectEnvironment
{
	Select-CurrentEnvironment -Default
	Select-CurrentEnvironment -EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444"
}
#gavdcodeend 042

#gavdcodebegin 043
Function PapPsAdmin_GetAllCustomConnectors
{
	Get-AdminPowerAppConnector
}
#gavdcodeend 043

#gavdcodebegin 044
Function PapPsMaker_GetAllConnectors
{
	Get-PowerAppConnector
}
#gavdcodeend 044

#gavdcodebegin 045
Function PapPsAdmin_DeleteCustomConnector
{
	Remove-AdminPowerAppConnector `
		-EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
		-ConnectorName "shared_chapter17getmails-5f8f636b4042be0adb-5f8a8be452227400db"
}
#gavdcodeend 045

#gavdcodebegin 046
Function PapPsMaker_DeleteConnector
{
	Remove-PowerAppConnector `
		-EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
		-ConnectorName "shared_zvanuparvaldnieks"
}
#gavdcodeend 046

#gavdcodebegin 014
Function PapPsMaker_EnumerateEnvironments
{
	Get-PowerAppEnvironment
}
#gavdcodeend 014

#gavdcodebegin 015
Function PapPsMaker_EnumerateApps
{
	Get-PowerApp
}
#gavdcodeend 015

#gavdcodebegin 016
Function PapPsMaker_SetDisplayName
{
	Set-PowerAppDisplayName `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-AppDisplayName "NameChangedApp"
}
#gavdcodeend 016

#gavdcodebegin 017
Function PapPsMaker_GetNotifications
{
	Get-PowerAppsNotification
}
#gavdcodeend 017

#gavdcodebegin 018
Function PapPsMaker_PublishApp
{
	Publish-PowerApp `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd"
}
#gavdcodeend 018

#gavdcodebegin 019
Function PapPsMaker_EnumerateVersions
{
	Get-PowerAppVersion `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505"
}
#gavdcodeend 019

#gavdcodebegin 020
Function PapPsMaker_RestoreVersion
{
	Restore-PowerAppVersion `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505" `
		-AppVersionName "20191215T131114Z"
}
#gavdcodeend 020

#gavdcodebegin 021
Function PapPsMaker_DeleteApp
{
	Remove-PowerApp `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505"
}
#gavdcodeend 021

#gavdcodebegin 022
Function PapPsMaker_FindRoles
{
	Get-PowerAppRoleAssignment `
		–AppName "c7965df9-a921-4a23-a21d-02ff19fca82d"
}
#gavdcodeend 022

#gavdcodebegin 023
Function PapPsMaker_AddRoles
{
	Set-PowerAppRoleAssignment `
		-AppName "c7965df9-a921-4a23-a21d-02ff19fca82d" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		-RoleName CanEdit `
		-PrincipalType User `
		-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 023

#gavdcodebegin 024
Function PapPsMaker_DeleteRoles
{
	$myRoleId = "/providers/Microsoft.PowerApps/apps/" + 
				"c7965df9-a921-4a23-a21d-02ff19fca82d/permissions/" + 
				"092b1237-a428-45a7-b76b-310fdd6e7246"

	Remove-PowerAppRoleAssignment `
		–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		–AppName "c7965df9-a921-4a23-a21d-02ff19fca82d" `
		-RoleId $myRoleId
}
#gavdcodeend 024

#-----------------------------------------------------------------------------------------

##==> Routines for CLI

#gavdcodebegin 025
Function PapPsCli_GetAllApps
{
	LoginPsCLI
	
	m365 pa app list

	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
Function PapPsCli_GetAllAppsByEnvironment
{
	LoginPsCLI
	
	m365 pa app list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					 --asAdmin

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
Function PapPsCli_GetOneApp
{
	LoginPsCLI
	
	m365 pa app get --name "d2b01511-bff7-4dbb-849d-6a482541fa4d"
	m365 pa app get --displayName "TestApp01"

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
Function PapPsCli_DeleteOneApp
{
	LoginPsCLI
	
	m365 pa app remove --name "d2b01511-bff7-4dbb-849d-6a482541fa4d" --confirm

	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 029
Function PapPsCli_GetAllEnvironment
{
	LoginPsCLI
	
	m365 pa environment list

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
Function PapPsCli_GetOneEnvironment
{
	LoginPsCLI
	
	m365 pa environment get --name "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
Function PapPsCli_GetAllConnectors
{
	LoginPsCLI
	
	m365 pa connector list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
Function PapPsCli_ExportOneConnectors
{
	LoginPsCLI
	
	m365 pa connector export --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
							 --connector "sh_con-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa" `
							 --outputFolder "C:\Temp\MyConnector"

	m365 logout
}
#gavdcodeend 032

#gavdcodebegin 033
Function PapPsCli_ScaffoldSolution
{
	LoginPsCLI
	
	m365 pa solution init --publisherName "GuitacaSolution" --publisherPrefix "ypn"

	m365 logout
}
#gavdcodeend 033

#gavdcodebegin 034
Function PapPsCli_ScaffoldComponent
{
	LoginPsCLI
	
	m365 pa pcf init --namespace "GuitacaNameSpace" `
					 --name "GuitacaDataset" `
					 --template "Dataset"

	m365 logout
}
#gavdcodeend 034

#-----------------------------------------------------------------------------------------


##==> Routines for PnPPowerShell

# No cmdlets for Power Apps in the PnP PowerShell module


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

##==> PowerShell Admin and Maker cmdlets
#LoginPsPowerPlatform
#PapPsAdmin_EnumerateApps
#PapPsAdmin_FindOneApps
#PapPsAdmin_UserDetails
#PapPsAdmin_SetOwner
#PapPsAdmin_SetFeatured
#PapPsAdmin_SetHero
#PapPsAdmin_DeleteApp
#PapPsAdmin_DeleteHero
#PapPsAdmin_DeleteFeatured
#PapPsAdmin_FindRoles
#PapPsAdmin_AddRoles
#PapPsAdmin_DeleteRoles
#PapPsAdmin_GetDeletedApps
#PapPsAdmin_RecoverDeletedApps
#PapPsAdmin_GetTenantSettings
#PapPsAdmin_ModifyTenantSettings
#PapPsAdmin_GetAllEnvironments
#PapPsAdmin_GetOneEnvironments
#PapPsAdmin_GetAdToken
#PapPsAdmin_SelectEnvironment
#PapPsAdmin_GetAllCustomConnectors
#PapPsMaker_GetAllConnectors
#PapPsAdmin_DeleteCustomConnector
#PapPsMaker_DeleteConnector
#PapPsMaker_EnumerateEnvironments
#PapPsMaker_EnumerateApps
#PapPsMaker_SetDisplayName
#PapPsMaker_GetNotifications
#PapPsMaker_PublishApp
#PapPsMaker_EnumerateVersions
#PapPsMaker_RestoreVersion
#PapPsMaker_DeleteApp
#PapPsMaker_FindRoles
#PapPsMaker_AddRoles
#PapPsMaker_DeleteRoles

##==> CLI
#PapPsCli_GetAllApps
#PapPsCli_GetAllAppsByEnvironment
#PapPsCli_GetOneApp
#PapPsCli_DeleteOneApp
#PapPsCli_GetAllEnvironment
#PapPsCli_GetOneEnvironment
#PapPsCli_GetAllConnectors
#PapPsCli_ExportOneConnectors
#PapPsCli_ScaffoldComponent

##==> PnPPowerShell
#PapPsPnpPowerShell_

Write-Host "Done"  
