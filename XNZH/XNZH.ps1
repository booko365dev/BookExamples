
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 01
Function LoginPsPowerPlatform()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	Add-PowerAppsAccount -Username $configFile.appsettings.UserName -Password $securePW
}
#gavdcodeend 01

Function LoginPsCLI()
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

#gavdcodebegin 02
Function PaPsAdmin_EnumerateFlows()
{
	Get-AdminFlow
}
#gavdcodeend 02

#gavdcodebegin 03
Function PaPsAdmin_OwnerRole()
{
	Get-AdminFlowOwnerRole `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e"
}
#gavdcodeend 03

#gavdcodebegin 04
Function PaPsAdmin_UserDetails()
{
	Get-AdminFlowUserDetails –UserId "acc28fcb-5261-47f8-960b-715d2f98a431"
}
#gavdcodeend 04

#gavdcodebegin 05
Function PaPsAdmin_DisableFlow()
{
	Disable-AdminFlow `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e"
}
#gavdcodeend 05

#gavdcodebegin 06
Function PaPsAdmin_EnableFlow()
{
	Enable-AdminFlow `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e"
}
#gavdcodeend 06

#gavdcodebegin 07
Function PaPsAdmin_DeleteFlow()
{
	Remove-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}
#gavdcodeend 07

#gavdcodebegin 40
Function PaPsAdmin_RestoreFlow()
{
	Restore-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}
#gavdcodeend 40

#gavdcodebegin 08
Function PaPsAdmin_DeleteApprovalFlows()
{
	Remove-AdminFlowApprovals `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 08

#gavdcodebegin 09
Function PaPsAdmin_AddRoleUser()
{
	Set-AdminFlowOwnerRole `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e" `
					-PrincipalType User `
					-PrincipalObjectId "bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb" `
					-RoleName CanEdit
}
#gavdcodeend 09

#gavdcodebegin 10
Function PaPsAdmin_DeleteRoleUser()
{
	$myRoleId = "/providers/Microsoft.ProcessSimple/environments/" + 
				"909ee029-5b74-4b2f-a9ee-b6b5158f630b/flows/" +
				"4355d741-c54b-4372-9bb0-eb5b49285333/permissions/" + 
				"959ae10e-0015-4948-b602-fbf7fccfe2a3"

	Remove-AdminFlowOwnerRole `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "4355d741-c54b-4372-9bb0-eb5b49285333" `
					-RoleId $myRoleId
}
#gavdcodeend 10

#gavdcodebegin 11
Function PaPsAdmin_DeleteUserDetails()
{
	Remove-AdminFlowUserDetails –UserId "092b1237-a428-45a7-b76b-310fdd6e7246"
}
#gavdcodeend 11

#gavdcodebegin 41
Function PaPsAdmin_CallApi()
{
	$myApi = 
	"https://api.met.no/weatherapi/locationforecast/2.0/classic?lat=52.353218&lon=5.0027695"
	$myBody = ""

	$myResult = InvokeApiNoParseContent `
									-Method GET `
									-Route $myApi `
									-Body $body `
									-ThrowOnFailure
	Write-Host $myResult.Content
}
#gavdcodeend 41

#gavdcodebegin 42
Function PaPsAdmin_GetTenantSettings()
{
	Get-TenantSettings
}
#gavdcodeend 42

#gavdcodebegin 43
Function PaPsAdmin_GetTenantDetails()
{
	Get-TenantDetailsFromGraph
}
#gavdcodeend 43

#gavdcodebegin 44
Function PaPsAdmin_GetUsersGroups()
{
	Get-UsersOrGroupsFromGraph -SearchString "Admin"
	Get-UsersOrGroupsFromGraph -ObjectId "Admin@tenant.onmicrosoft.com"
}
#gavdcodeend 44

#gavdcodebegin 12
Function PaPs_RemoveFromSharePoint()
{
	# It works only for SharePoint Teams sites created without Group
	# Use PowerShell 5.x

	$AdminSiteURL="https://[domain]-admin.sharepoint.com/"
	$SiteURL="https://[domain].sharepoint.com/sites/[TeamSiteWithoutGroup]"
 
	Connect-SPOService -Url $AdminSiteURL -Credential (Get-Credential)
 
	Set-SPOSite -Identity $SiteURL -DisableFlows Disabled
 
	#To Enable the button: 
	#Set-SPOSite -Identity $SiteURL -DisableFlows NotDisabled
}
#gavdcodeend 12

#gavdcodebegin 13
Function PaPsMaker_EnumarateEnvironment()
{
	Get-FlowEnvironment
}
#gavdcodeend 13

#gavdcodebegin 14
Function PaPsMaker_EnumarateFlows()
{
	Get-Flow
	Write-Host "-------------------------"
	Get-Flow -EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444"
}
#gavdcodeend 14

#gavdcodebegin 15
Function PaPsMaker_GetOneFlow()
{
	Get-FlowRun –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}
#gavdcodeend 15

#gavdcodebegin 16
Function PaPsMaker_DisableFlow()
{
	Disable-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}
#gavdcodeend 16

#gavdcodebegin 17
Function PaPsMaker_EnableFlow()
{
	Enable-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}
#gavdcodeend 17

#gavdcodebegin 18
Function PaPsMaker_DeleteFlow()
{
	Remove-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b" -Confirm:$false
}
#gavdcodeend 18

#gavdcodebegin 19
Function PaPsMaker_EnumarateFlowApprovals()
{
	Get-FlowApproval –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 19

#gavdcodebegin 20
Function PaPsMaker_EnumarateFlowApprovalRequests()
{
	Get-FlowApprovalRequest –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 20

#gavdcodebegin 21
Function PaPsMaker_ApproveFlows()
{
	$myApprovals = Get-FlowApprovalRequest `
					-EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"

	foreach($oneApproval in $myApprovals) {
		Approve-FlowApprovalRequest `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					-Comments "Approved" `
					-ApprovalId $oneApproval.ApprovalId `
					-ApprovalRequestId $oneApproval.ApprovalRequestId
	}
}
#gavdcodeend 21

#gavdcodebegin 22
Function PaPsMaker_RejectFlows()
{
	$myApprovals = Get-FlowApprovalRequest `
					-EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"

	foreach($oneApproval in $myApprovals) {
		Deny-FlowApprovalRequest `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					-Comments "Rejected" `
					-ApprovalId $oneApproval.ApprovalId `
					-ApprovalRequestId $oneApproval.ApprovalRequestId
	}
}
#gavdcodeend 22

#gavdcodebegin 23
Function PaPsMaker_OwnerRole()
{
	Get-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a"
}
#gavdcodeend 23

#gavdcodebegin 24
Function PaPsMaker_AddRoleUser()
{
	Set-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a" `
					-PrincipalType User `
					-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 24

#gavdcodebegin 25
Function PaPsMaker_DeleteRoleUser()
{
	$myRoleId = "/providers/Microsoft.ProcessSimple/environments/" + 
				"Default-03d561bf-4472-41e0-b2d6-ee506471e9d0/flows/" + 
				"28327980-4786-4107-9f2e-80674c3cb98a/owners/" + 
				"959ae10e-0015-4948-b602-fbf7fccfe2a3"

	Remove-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a" `
					-RoleId $myRoleId
}
#gavdcodeend 25

#-----------------------------------------------------------------------------------------

##==> Routines for CLI

#gavdcodebegin 26
Function PaPsCli_GetAllFlowsByEnvironment()
{
	LoginPsCLI
	
	m365 flow list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
				   --asAdmin

	m365 logout
}
#gavdcodeend 26

#gavdcodebegin 27
Function PaPsCli_GetOneFlow()
{
	LoginPsCLI
	
	m365 flow get --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
				  --name "3126a4e4-71b9-49d8-802d-734a71534ff4" `
				  --asAdmin

	m365 logout
}
#gavdcodeend 27

#gavdcodebegin 28
Function PaPsCli_ExportOneFlow()
{
	LoginPsCLI
	
	m365 flow export --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					 --id "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					 --format "zip" `
					 --packageDisplayName "MyTestFlow" `
					 --packageDescription "It is a test flow" `
					 --packageCreatedBy "Guitaca" `
					 --packageSourceEnvironment "My Default Environment" `
					 --path 'C:\Temporary\MyTestFlow.zip'

	m365 logout
}
#gavdcodeend 28

#gavdcodebegin 29
Function PaPsCli_DisableOneFlow()
{
	LoginPsCLI
	
	m365 flow disable --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					  --name "3126a4e4-71b9-49d8-802d-734a71534ff4"

	m365 logout
}
#gavdcodeend 29

#gavdcodebegin 30
Function PaPsCli_EnableOneFlow()
{
	LoginPsCLI
	
	m365 flow enable --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					 --name "3126a4e4-71b9-49d8-802d-734a71534ff4"

	m365 logout
}
#gavdcodeend 30

#gavdcodebegin 31
Function PaPsCli_DeleteOneFlow()
{
	LoginPsCLI
	
	m365 flow remove --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					 --name "e7817682-71ce-4bee-b435-b98e0edfad12"

	m365 logout
}
#gavdcodeend 31

#gavdcodebegin 32
Function PaPsCli_GetAllEnvironment()
{
	LoginPsCLI
	
	m365 flow environment list

	m365 logout
}
#gavdcodeend 32

#gavdcodebegin 33
Function PaPsCli_GetOneEnvironment()
{
	LoginPsCLI
	
	m365 flow environment get --name "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 33

#gavdcodebegin 34
Function PaPsCli_GetAllConnectors()
{
	LoginPsCLI
	
	m365 flow connector list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 34

#gavdcodebegin 35
Function PaPsCli_ExportOneConnectors()
{
	LoginPsCLI
	
	m365 flow connector export --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
							   --connector "sh_con-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa" `
							   --outputFolder "C:\Temp\MyConnector"

	m365 logout
}
#gavdcodeend 35

#gavdcodebegin 36
Function PaPsCli_GetRunsOneFlow()
{
	LoginPsCLI
	
	m365 flow run list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					   --flow "3126a4e4-71b9-49d8-802d-734a71534ff4"

	m365 logout
}
#gavdcodeend 36

#gavdcodebegin 37
Function PaPsCli_GetOneRunOneFlow()
{
	LoginPsCLI
	
	m365 flow run get --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					  --flow "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					  --name "08585583736314022523222750393CU146"

	m365 logout
}
#gavdcodeend 37

#gavdcodebegin 38
Function PaPsCli_ResubmitOneRunOneFlow()
{
	LoginPsCLI
	
	m365 flow run resubmit --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					       --flow "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					       --name "08585583736314022523222750393CU146"

	m365 logout
}
#gavdcodeend 38

#gavdcodebegin 39
Function PaPsCli_CancelOneRunOneFlow()
{
	LoginPsCLI
	
	m365 flow run cancel --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					     --flow "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					     --name "08585583736314022523222750393CU146"

	m365 logout
}
#gavdcodeend 39

#-----------------------------------------------------------------------------------------


##==> Routines for PnPPowerShell

#gavdcodebegin 45
function SpPsPnpPowerShell_GetEnvironment
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $true

	Disconnect-PnPOnline
}
#gavdcodeend 45

#gavdcodebegin 46
function SpPsPnpPowerShell_GetAllFlowsInEnvironment
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $true
	Get-PnPFlow -Environment $myEnvironment

	Disconnect-PnPOnline
}
#gavdcodeend 46

#gavdcodebegin 47
function SpPsPnpPowerShell_GetOneFlowInEnvironment
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Get-PnPFlow -Environment $myEnvironment `
				-Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 47

#gavdcodebegin 48
function SpPsPnpPowerShell_GetOneFlowRuns
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Get-PnPFlowRun -Environment $myEnvironment `
				   -Flow "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 48

#gavdcodebegin 49
function SpPsPnpPowerShell_DisableOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Disable-PnPFlow -Environment $myEnvironment `
				    -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 49

#gavdcodebegin 50
function SpPsPnpPowerShell_EnableOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Enable-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 50

#gavdcodebegin 51
function SpPsPnpPowerShell_StopOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Stop-PnPFlowRun -Environment $myEnvironment `
				    -Flow "1ecbcc34-ec34-4220-b39e-bde8c961655e" `
				    -Identity "08585321999201590891763396367CU168" `
					-Force

	Disconnect-PnPOnline
}
#gavdcodeend 51

#gavdcodebegin 52
function SpPsPnpPowerShell_RestartOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Restart-PnPFlowRun -Environment $myEnvironment `
					   -Flow "1ecbcc34-ec34-4220-b39e-bde8c961655e" `
				       -Identity "08585321999201590891763396367CU168" `
					   -Force

	Disconnect-PnPOnline
}
#gavdcodeend 52

#gavdcodebegin 53
function SpPsPnpPowerShell_ExportOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Export-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 53

#gavdcodebegin 54
function SpPsPnpPowerShell_ExportOneFlowZip
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Export-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e" `
				   -AsZipPackage `
				   -OutPath "C:\Temporary\myFlow.zip" `
				   -Verbose

	Disconnect-PnPOnline
}
#gavdcodeend 54

#gavdcodebegin 55
function SpPsPnpPowerShell_DeleteOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Remove-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 55


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

##==> PowerShell Admin and Maker cmdlets
#LoginPsPowerPlatform

#PaPsAdmin_EnumerateFlows
#PaPsAdmin_OwnerRole
#PaPsAdmin_UserDetails
#PaPsAdmin_DisableFlow
#PaPsAdmin_EnableFlow
#PaPsAdmin_DeleteFlow
#PaPsAdmin_RestoreFlow
#PaPsAdmin_DeleteApprovalFlows
#PaPsAdmin_AddRoleUser
#PaPsAdmin_DeleteRoleUser
#PaPsAdmin_DeleteUserDetails
#PaPs_RemoveFromSharePoint
#PaPsMaker_EnumarateEnvironment
#PaPsMaker_EnumarateFlows
#PaPsMaker_GetOneFlow
#PaPsMaker_DisableFlow
#PaPsMaker_EnableFlow
#PaPsMaker_DeleteFlow
#PaPsMaker_EnumarateFlowApprovals
#PaPsMaker_EnumarateFlowApprovalRequests
#PaPsMaker_ApproveFlows
#PaPsMaker_RejectFlows
#PaPsMaker_OwnerRole
#PaPsMaker_AddRoleUser
#PaPsMaker_DeleteRoleUser
#PaPsAdmin_CallApi
#PaPsAdmin_GetTenantDetails
#PaPsAdmin_GetTenantSettings
#PaPsAdmin_GetUsersGroups

##==> CLI
#PaPsCli_GetAllFlowsByEnvironment
#PaPsCli_GetOneFlow
#PaPsCli_ExportOneFlow
#PaPsCli_DisableOneFlow
#PaPsCli_EnableOneFlow
#PaPsCli_DeleteOneFlow
#PaPsCli_GetAllEnvironment
#PaPsCli_GetOneEnvironment
#PaPsCli_GetAllConnectors
#PaPsCli_ExportOneConnectors
#PaPsCli_GetRunsOneFlow
#PaPsCli_GetOneRunOneFlow
#PaPsCli_ResubmitOneRunOneFlow
#PaPsCli_CancelOneRunOneFlow

##==> PnPPowerShell
#SpPsPnpPowerShell_GetEnvironment
#SpPsPnpPowerShell_GetAllFlowsInEnvironment
#SpPsPnpPowerShell_GetOneFlowInEnvironment
#SpPsPnpPowerShell_GetOneFlowRuns
#SpPsPnpPowerShell_DisableOneFlow
#SpPsPnpPowerShell_EnableOneFlow
#SpPsPnpPowerShell_StopOneFlow
#SpPsPnpPowerShell_RestartOneFlow
#SpPsPnpPowerShell_ExportOneFlow
#SpPsPnpPowerShell_ExportOneFlowZip
#SpPsPnpPowerShell_DeleteOneFlow

Write-Host "Done"  
