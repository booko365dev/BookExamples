
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
			   --appId $configFile.appsettings.ClientIdWithAccPw `
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
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithAccPw `
					  -Credentials $myCredentials
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


##==> Routines for PowerShell Admin and Maker cmdlets

#gavdcodebegin 002
Function PauPsAdmin_EnumerateFlows
{
	Get-AdminFlow
}
#gavdcodeend 002

#gavdcodebegin 003
Function PauPsAdmin_OwnerRole
{
	Get-AdminFlowOwnerRole `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e"
}
#gavdcodeend 003

#gavdcodebegin 004
Function PauPsAdmin_UserDetails
{
	Get-AdminFlowUserDetails –UserId "acc28fcb-5261-47f8-960b-715d2f98a431"
}
#gavdcodeend 004

#gavdcodebegin 005
Function PauPsAdmin_DisableFlow
{
	Disable-AdminFlow `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e"
}
#gavdcodeend 005

#gavdcodebegin 006
Function PauPsAdmin_EnableFlow
{
	Enable-AdminFlow `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e"
}
#gavdcodeend 006

#gavdcodebegin 007
Function PauPsAdmin_DeleteFlow
{
	Remove-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}
#gavdcodeend 007

#gavdcodebegin 040
Function PauPsAdmin_RestoreFlow
{
	Restore-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}
#gavdcodeend 040

#gavdcodebegin 008
Function PauPsAdmin_DeleteApprovalFlows
{
	Remove-AdminFlowApprovals `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 008

#gavdcodebegin 009
Function PauPsAdmin_AddRoleUser
{
	Set-AdminFlowOwnerRole `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e" `
					-PrincipalType User `
					-PrincipalObjectId "bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb" `
					-RoleName CanEdit
}
#gavdcodeend 009

#gavdcodebegin 010
Function PauPsAdmin_DeleteRoleUser
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
#gavdcodeend 010

#gavdcodebegin 011
Function PauPsAdmin_DeleteUserDetails
{
	Remove-AdminFlowUserDetails –UserId "092b1237-a428-45a7-b76b-310fdd6e7246"
}
#gavdcodeend 011

#gavdcodebegin 041
Function PauPsAdmin_CallApi
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
#gavdcodeend 041

#gavdcodebegin 042
Function PauPsAdmin_GetTenantSettings
{
	Get-TenantSettings
}
#gavdcodeend 042

#gavdcodebegin 043
Function PauPsAdmin_GetTenantDetails
{
	Get-TenantDetailsFromGraph
}
#gavdcodeend 043

#gavdcodebegin 044
Function PauPsAdmin_GetUsersGroups
{
	Get-UsersOrGroupsFromGraph -SearchString "Admin"
	Get-UsersOrGroupsFromGraph -ObjectId "Admin@tenant.onmicrosoft.com"
}
#gavdcodeend 044

#gavdcodebegin 012
Function PauPs_RemoveFromSharePoint
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
#gavdcodeend 012

#gavdcodebegin 013
Function PauPsMaker_EnumarateEnvironment
{
	Get-FlowEnvironment
}
#gavdcodeend 013

#gavdcodebegin 014
Function PauPsMaker_EnumarateFlows
{
	Get-Flow
	Write-Host "-------------------------"
	Get-Flow -EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444"
}
#gavdcodeend 014

#gavdcodebegin 015
Function PauPsMaker_GetOneFlow
{
	Get-FlowRun –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}
#gavdcodeend 015

#gavdcodebegin 016
Function PauPsMaker_DisableFlow
{
	Disable-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}
#gavdcodeend 016

#gavdcodebegin 017
Function PauPsMaker_EnableFlow
{
	Enable-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}
#gavdcodeend 017

#gavdcodebegin 018
Function PauPsMaker_DeleteFlow
{
	Remove-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b" -Confirm:$false
}
#gavdcodeend 018

#gavdcodebegin 019
Function PauPsMaker_EnumarateFlowApprovals
{
	Get-FlowApproval –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 019

#gavdcodebegin 020
Function PauPsMaker_EnumarateFlowApprovalRequests
{
	Get-FlowApprovalRequest –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 020

#gavdcodebegin 021
Function PauPsMaker_ApproveFlows
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
#gavdcodeend 021

#gavdcodebegin 022
Function PauPsMaker_RejectFlows
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
#gavdcodeend 022

#gavdcodebegin 023
Function PauPsMaker_OwnerRole
{
	Get-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a"
}
#gavdcodeend 023

#gavdcodebegin 024
Function PauPsMaker_AddRoleUser
{
	Set-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a" `
					-PrincipalType User `
					-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 024

#gavdcodebegin 025
Function PauPsMaker_DeleteRoleUser
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
#gavdcodeend 025

#-----------------------------------------------------------------------------------------

##==> Routines for CLI

#gavdcodebegin 026
Function PauPsCli_GetAllFlowsByEnvironment
{
	LoginPsCLI
	
	m365 flow list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
				   --asAdmin

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
Function PauPsCli_GetOneFlow
{
	LoginPsCLI
	
	m365 flow get --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
				  --name "3126a4e4-71b9-49d8-802d-734a71534ff4" `
				  --asAdmin

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
Function PauPsCli_ExportOneFlow
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
#gavdcodeend 028

#gavdcodebegin 029
Function PauPsCli_DisableOneFlow
{
	LoginPsCLI
	
	m365 flow disable --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					  --name "3126a4e4-71b9-49d8-802d-734a71534ff4"

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
Function PauPsCli_EnableOneFlow
{
	LoginPsCLI
	
	m365 flow enable --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					 --name "3126a4e4-71b9-49d8-802d-734a71534ff4"

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
Function PauPsCli_DeleteOneFlow
{
	LoginPsCLI
	
	m365 flow remove --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					 --name "e7817682-71ce-4bee-b435-b98e0edfad12"

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
Function PauPsCli_GetAllEnvironment
{
	LoginPsCLI
	
	m365 flow environment list

	m365 logout
}
#gavdcodeend 032

#gavdcodebegin 033
Function PauPsCli_GetOneEnvironment
{
	LoginPsCLI
	
	m365 flow environment get --name "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 033

#gavdcodebegin 034
Function PauPsCli_GetAllConnectors
{
	LoginPsCLI
	
	m365 flow connector list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 034

#gavdcodebegin 035
Function PauPsCli_ExportOneConnectors
{
	LoginPsCLI
	
	m365 flow connector export --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
							   --connector "sh_con-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa" `
							   --outputFolder "C:\Temp\MyConnector"

	m365 logout
}
#gavdcodeend 035

#gavdcodebegin 036
Function PauPsCli_GetRunsOneFlow
{
	LoginPsCLI
	
	m365 flow run list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					   --flow "3126a4e4-71b9-49d8-802d-734a71534ff4"

	m365 logout
}
#gavdcodeend 036

#gavdcodebegin 037
Function PauPsCli_GetOneRunOneFlow
{
	LoginPsCLI
	
	m365 flow run get --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					  --flow "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					  --name "08585583736314022523222750393CU146"

	m365 logout
}
#gavdcodeend 037

#gavdcodebegin 038
Function PauPsCli_ResubmitOneRunOneFlow
{
	LoginPsCLI
	
	m365 flow run resubmit --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					       --flow "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					       --name "08585583736314022523222750393CU146"

	m365 logout
}
#gavdcodeend 038

#gavdcodebegin 039
Function PauPsCli_CancelOneRunOneFlow
{
	LoginPsCLI
	
	m365 flow run cancel --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					     --flow "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					     --name "08585583736314022523222750393CU146"

	m365 logout
}
#gavdcodeend 039

#-----------------------------------------------------------------------------------------


##==> Routines for PnPPowerShell

#gavdcodebegin 045
function PauPsPnpPowerShell_GetEnvironment
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $true

	Disconnect-PnPOnline
}
#gavdcodeend 045

#gavdcodebegin 046
function PauPsPnpPowerShell_GetAllFlowsInEnvironment
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $true
	Get-PnPFlow -Environment $myEnvironment

	Disconnect-PnPOnline
}
#gavdcodeend 046

#gavdcodebegin 047
function PauPsPnpPowerShell_GetOneFlowInEnvironment
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Get-PnPFlow -Environment $myEnvironment `
				-Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 047

#gavdcodebegin 048
function PauPsPnpPowerShell_GetOneFlowRuns
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Get-PnPFlowRun -Environment $myEnvironment `
				   -Flow "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 048

#gavdcodebegin 049
function PauPsPnpPowerShell_DisableOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Disable-PnPFlow -Environment $myEnvironment `
				    -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 049

#gavdcodebegin 050
function PauPsPnpPowerShell_EnableOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Enable-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 050

#gavdcodebegin 051
function PauPsPnpPowerShell_StopOneFlow
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
#gavdcodeend 051

#gavdcodebegin 052
function PauPsPnpPowerShell_RestartOneFlow
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
#gavdcodeend 052

#gavdcodebegin 053
function PauPsPnpPowerShell_ExportOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Export-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 053

#gavdcodebegin 054
function PauPsPnpPowerShell_ExportOneFlowZip
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
#gavdcodeend 054

#gavdcodebegin 055
function PauPsPnpPowerShell_DeleteOneFlow
{
	# App Registration type: Office 365 Online 
	# App Registration permissions: Azure: management.azure.com
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault $false
	Remove-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 055


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

##==> PowerShell Admin and Maker cmdlets
#LoginPsPowerPlatform

#PauPsAdmin_EnumerateFlows
#PauPsAdmin_OwnerRole
#PauPsAdmin_UserDetails
#PauPsAdmin_DisableFlow
#PauPsAdmin_EnableFlow
#PauPsAdmin_DeleteFlow
#PauPsAdmin_RestoreFlow
#PauPsAdmin_DeleteApprovalFlows
#PauPsAdmin_AddRoleUser
#PauPsAdmin_DeleteRoleUser
#PauPsAdmin_DeleteUserDetails
#PauPs_RemoveFromSharePoint
#PauPsMaker_EnumarateEnvironment
#PauPsMaker_EnumarateFlows
#PauPsMaker_GetOneFlow
#PauPsMaker_DisableFlow
#PauPsMaker_EnableFlow
#PauPsMaker_DeleteFlow
#PauPsMaker_EnumarateFlowApprovals
#PauPsMaker_EnumarateFlowApprovalRequests
#PauPsMaker_ApproveFlows
#PauPsMaker_RejectFlows
#PauPsMaker_OwnerRole
#PauPsMaker_AddRoleUser
#PauPsMaker_DeleteRoleUser
#PauPsAdmin_CallApi
#PauPsAdmin_GetTenantDetails
#PauPsAdmin_GetTenantSettings
#PauPsAdmin_GetUsersGroups

##==> CLI
#PauPsCli_GetAllFlowsByEnvironment
#PauPsCli_GetOneFlow
#PauPsCli_ExportOneFlow
#PauPsCli_DisableOneFlow
#PauPsCli_EnableOneFlow
#PauPsCli_DeleteOneFlow
#PauPsCli_GetAllEnvironment
#PauPsCli_GetOneEnvironment
#PauPsCli_GetAllConnectors
#PauPsCli_ExportOneConnectors
#PauPsCli_GetRunsOneFlow
#PauPsCli_GetOneRunOneFlow
#PauPsCli_ResubmitOneRunOneFlow
#PauPsCli_CancelOneRunOneFlow

##==> PnPPowerShell
#PauPsPnpPowerShell_GetEnvironment
#PauPsPnpPowerShell_GetAllFlowsInEnvironment
#PauPsPnpPowerShell_GetOneFlowInEnvironment
#PauPsPnpPowerShell_GetOneFlowRuns
#PauPsPnpPowerShell_DisableOneFlow
#PauPsPnpPowerShell_EnableOneFlow
#PauPsPnpPowerShell_StopOneFlow
#PauPsPnpPowerShell_RestartOneFlow
#PauPsPnpPowerShell_ExportOneFlow
#PauPsPnpPowerShell_ExportOneFlowZip
#PauPsPnpPowerShell_DeleteOneFlow

Write-Host "Done"  
