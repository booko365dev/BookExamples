
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


#gavdcodebegin 001
Function PsPpPs_LoginWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$UserPw,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
			$UserPw -AsPlainText -Force

	Add-PowerAppsAccount -Username $UserName -Password $securePW
}
#gavdcodeend 001

Function PsCliM365_LoginWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientIdWithAccPw
	)

	m365 login --authType password `
			   --appId $ClientIdWithAccPw `
			   --userName $UserName `
			   --password $UserPw
}

function PsCliM365_LoginWithCertificateFile
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientId,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateFilePath,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateFilePw
	)

	m365 login --authType certificate `
			   --tenant $TenantName --appId $ClientId `
			   --certificateFile $CertificateFilePath --password $CertificateFilePw
}

Function PsPnpPowerShell_LoginWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw,
 
		[Parameter(Mandatory=$True)]
		[String]$SiteCollUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientIdWithAccPw
	)

	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $UserName, $securePW
	Connect-PnPOnline -Url $SiteCollUrl `
					  -ClientId $ClientIdWithAccPw `
					  -Credentials $myCredentials
}

function PsPnPPowerShell_LoginGraphWithCertificateThumbprint
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$SiteBaseUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientId,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateThumbprint
	)

	Connect-PnPOnline -Url $SiteBaseUrl -Tenant $TenantName -ClientId $ClientId `
					  -Thumbprint $CertificateThumbprint
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


##==> Routines for PowerShell Admin and Maker cmdlets

#gavdcodebegin 002
Function PsPapAdmin_EnumerateApps
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerApp
}
#gavdcodeend 002

#gavdcodebegin 003
Function PsPapAdmin_FindOneApps
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerApp "2762f714-52b4-4766-9fea-e1a8c72d2f36"
}
#gavdcodeend 003

#gavdcodebegin 004
Function PsPapAdmin_UserDetails
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerAppsUserDetails `
		-OutputFilePath "C:\Temporary\UsersPA.json" `
		–UserPrincipalName "user@domain.onmicrosoft.com"
}
#gavdcodeend 004

#gavdcodebegin 005
Function PsPapAdmin_SetOwner
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Set-AdminPowerAppOwner `
		–AppName "01d96b0e-f371-4ced-91c4-bc53acb5dbcf" `
		-AppOwner "092b1237-a428-45a7-b76b-310fdd6e7246" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0"
}
#gavdcodeend 005

#gavdcodebegin 006
Function PsPapAdmin_SetFeatured
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerApp "NameApp" | Set-AdminPowerAppAsFeatured
}
#gavdcodeend 006

#gavdcodebegin 007
Function PsPapAdmin_SetHero
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerApp "NameApp" | Set-AdminPowerAppAsHero
}
#gavdcodeend 007

#gavdcodebegin 009
Function PsPapAdmin_DeleteFeatured
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerApp "NameApp" | Clear-AdminPowerAppAsFeatured
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsPapAdmin_DeleteHero
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerApp "NameApp" | Clear-AdminPowerAppAsHero
}
#gavdcodeend 010

#gavdcodebegin 008
Function PsPapAdmin_DeleteApp
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Remove-AdminPowerApp `
		–AppName "01d96b0e-f371-4ced-91c4-bc53acb5dbcf" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0"
}
#gavdcodeend 008

#gavdcodebegin 011
Function PsPapAdmin_FindRoles
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerAppRoleAssignment `
		–UserId "acc28fcb-5261-47f8-960b-715d2f98a431"
}
#gavdcodeend 011

#gavdcodebegin 012
Function PsPapAdmin_AddRoles
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Set-AdminPowerAppRoleAssignment `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		-RoleName CanEdit `
		-PrincipalType User `
		-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 012

#gavdcodebegin 013
Function PsPapAdmin_DeleteRoles
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

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
Function PsPapAdmin_GetDeletedApps
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminDeletedPowerAppsList `
						-EnvironmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca"
}
#gavdcodeend 035

#gavdcodebegin 036
Function PsPapAdmin_RecoverDeletedApps
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminRecoverDeletedPowerApp `
						-AppName "a4978531-2218-406c-9158-0b9353334c6d" `
						-EnvironmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca"
}
#gavdcodeend 036

#gavdcodebegin 037
Function PsPapAdmin_GetTenantSettings
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-TenantSettings
}
#gavdcodeend 037

#gavdcodebegin 038
Function PsPapAdmin_ModifyTenantSettings
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	$mySettings = Get-TenantSettings 
	$mySettings.disableTrialEnvironmentCreationByNonAdminUsers = $true 
	Set-TenantSettings -RequestBody $mySettings
}
#gavdcodeend 038

#gavdcodebegin 039
Function PsPapAdmin_GetAllEnvironments
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerAppEnvironment
}
#gavdcodeend 039

#gavdcodebegin 040
Function PsPapAdmin_GetOneEnvironments
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerAppEnvironment –Default
	Get-AdminPowerAppEnvironment –EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444"
}
#gavdcodeend 040

#gavdcodebegin 041
Function PsPapAdmin_GetAdToken
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-JwtToken "https://service.powerapps.com/"
}
#gavdcodeend 041

#gavdcodebegin 042
Function PsPapAdmin_SelectEnvironment
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Select-CurrentEnvironment -Default
	Select-CurrentEnvironment -EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444"
}
#gavdcodeend 042

#gavdcodebegin 043
Function PsPapAdmin_GetAllCustomConnectors
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminPowerAppConnector
}
#gavdcodeend 043

#gavdcodebegin 044
Function PsPapMaker_GetAllConnectors
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-PowerAppConnector
}
#gavdcodeend 044

#gavdcodebegin 045
Function PsPapAdmin_DeleteCustomConnector
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Remove-AdminPowerAppConnector `
		-EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
		-ConnectorName "shared_chapter17getmails-5f8f636b4042be0adb-5f8a8be452227400db"
}
#gavdcodeend 045

#gavdcodebegin 046
Function PsPapMaker_DeleteConnector
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Remove-PowerAppConnector `
		-EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
		-ConnectorName "shared_zvanuparvaldnieks"
}
#gavdcodeend 046

#gavdcodebegin 014
Function PsPapMaker_EnumerateEnvironments
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-PowerAppEnvironment
}
#gavdcodeend 014

#gavdcodebegin 015
Function PsPapMaker_EnumerateApps
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-PowerApp
}
#gavdcodeend 015

#gavdcodebegin 016
Function PsPapMaker_SetDisplayName
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Set-PowerAppDisplayName `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd" `
		-AppDisplayName "NameChangedApp"
}
#gavdcodeend 016

#gavdcodebegin 017
Function PsPapMaker_GetNotifications
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-PowerAppsNotification
}
#gavdcodeend 017

#gavdcodebegin 018
Function PsPapMaker_PublishApp
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Publish-PowerApp `
		-AppName "fa014c64-efe7-4301-bea2-9034bb7b51fd"
}
#gavdcodeend 018

#gavdcodebegin 019
Function PsPapMaker_EnumerateVersions
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-PowerAppVersion `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505"
}
#gavdcodeend 019

#gavdcodebegin 020
Function PsPapMaker_RestoreVersion
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Restore-PowerAppVersion `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505" `
		-AppVersionName "20191215T131114Z"
}
#gavdcodeend 020

#gavdcodebegin 021
Function PsPapMaker_DeleteApp
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Remove-PowerApp `
		-AppName "c9a52c61-a550-4c5f-ac2c-b3c36032a505"
}
#gavdcodeend 021

#gavdcodebegin 022
Function PsPapMaker_FindRoles
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-PowerAppRoleAssignment `
		–AppName "c7965df9-a921-4a23-a21d-02ff19fca82d"
}
#gavdcodeend 022

#gavdcodebegin 023
Function PsPapMaker_AddRoles
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Set-PowerAppRoleAssignment `
		-AppName "c7965df9-a921-4a23-a21d-02ff19fca82d" `
		-EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
		-RoleName CanEdit `
		-PrincipalType User `
		-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 023

#gavdcodebegin 024
Function PsPapMaker_DeleteRoles
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

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
Function PsPapCli_GetAllApps
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 pa app list

	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
Function PsPapCli_GetAllAppsByEnvironment
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 pa app list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					 --asAdmin

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
Function PsPapCli_GetOneApp
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 pa app get --name "d2b01511-bff7-4dbb-849d-6a482541fa4d"
	m365 pa app get --displayName "TestApp01"

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
Function PsPapCli_DeleteOneApp
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 pa app remove --name "d2b01511-bff7-4dbb-849d-6a482541fa4d" --confirm

	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 029
Function PsPapCli_GetAllEnvironment
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 pa environment list

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
Function PsPapCli_GetOneEnvironment
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 pa environment get --name "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
Function PsPapCli_GetAllConnectors
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 pa connector list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
Function PsPapCli_ExportOneConnectors
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 pa connector export --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
							 --connector "sh_con-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa" `
							 --outputFolder "C:\Temp\MyConnector"

	m365 logout
}
#gavdcodeend 032

#gavdcodebegin 033
Function PsPapCli_ScaffoldSolution
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 pa solution init --publisherName "GuitacaSolution" --publisherPrefix "ypn"

	m365 logout
}
#gavdcodeend 033

#gavdcodebegin 034
Function PsPapCli_ScaffoldComponent
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
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

# *** Latest Source Code Index: 046 ***

#region ConfigValuesCS.config
[xml]$config = Get-Content -Path "C:\Projects\ConfigValuesCS.config"
$cnfUserName               = $config.SelectSingleNode("//add[@key='UserName']").value
$cnfUserPw                 = $config.SelectSingleNode("//add[@key='UserPw']").value
$cnfTenantUrl              = $config.SelectSingleNode("//add[@key='TenantUrl']").value     # https://domain.onmicrosoft.com
$cnfSiteBaseUrl            = $config.SelectSingleNode("//add[@key='SiteBaseUrl']").value   # https://domain.sharepoint.com
$cnfSiteAdminUrl           = $config.SelectSingleNode("//add[@key='SiteAdminUrl']").value  # https://domain-admin.sharepoint.com
$cnfSiteCollUrl            = $config.SelectSingleNode("//add[@key='SiteCollUrl']").value   # https://domain.sharepoint.com/sites/TestSite
$cnfTenantName             = $config.SelectSingleNode("//add[@key='TenantName']").value
$cnfClientIdWithAccPw      = $config.SelectSingleNode("//add[@key='ClientIdWithAccPw']").value
$cnfClientIdWithSecret     = $config.SelectSingleNode("//add[@key='ClientIdWithSecret']").value
$cnfClientSecret           = $config.SelectSingleNode("//add[@key='ClientSecret']").value
$cnfClientIdWithCert       = $config.SelectSingleNode("//add[@key='ClientIdWithCert']").value
$cnfCertificateThumbprint  = $config.SelectSingleNode("//add[@key='CertificateThumbprint']").value
$cnfCertificateFilePath    = $config.SelectSingleNode("//add[@key='CertificateFilePath']").value
$cnfCertificateFilePw      = $config.SelectSingleNode("//add[@key='CertificateFilePw']").value
#endregion ConfigValuesCS.config

##==> PowerShell Admin and Maker cmdlets
#PsPapAdmin_EnumerateApps
#PsPapAdmin_FindOneApps
#PsPapAdmin_UserDetails
#PsPapAdmin_SetOwner
#PsPapAdmin_SetFeatured
#PsPapAdmin_SetHero
#PsPapAdmin_DeleteApp
#PsPapAdmin_DeleteHero
#PsPapAdmin_DeleteFeatured
#PsPapAdmin_FindRoles
#PsPapAdmin_AddRoles
#PsPapAdmin_DeleteRoles
#PsPapAdmin_GetDeletedApps
#PsPapAdmin_RecoverDeletedApps
#PsPapAdmin_GetTenantSettings
#PsPapAdmin_ModifyTenantSettings
#PsPapAdmin_GetAllEnvironments
#PsPapAdmin_GetOneEnvironments
#PsPapAdmin_GetAdToken
#PsPapAdmin_SelectEnvironment
#PsPapAdmin_GetAllCustomConnectors
#PsPapMaker_GetAllConnectors
#PsPapAdmin_DeleteCustomConnector
#PsPapMaker_DeleteConnector
#PsPapMaker_EnumerateEnvironments
#PsPapMaker_EnumerateApps
#PsPapMaker_SetDisplayName
#PsPapMaker_GetNotifications
#PsPapMaker_PublishApp
#PsPapMaker_EnumerateVersions
#PsPapMaker_RestoreVersion
#PsPapMaker_DeleteApp
#PsPapMaker_FindRoles
#PsPapMaker_AddRoles
#PsPapMaker_DeleteRoles

##==> CLI
#PsPapCli_GetAllApps
#PsPapCli_GetAllAppsByEnvironment
#PsPapCli_GetOneApp
#PsPapCli_DeleteOneApp
#PsPapCli_GetAllEnvironment
#PsPapCli_GetOneEnvironment
#PsPapCli_GetAllConnectors
#PsPapCli_ExportOneConnectors
#PsPapCli_ScaffoldComponent

##==> PnPPowerShell
#PapPsPnpPowerShell_

Write-Host "Done"  
