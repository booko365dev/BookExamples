
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
Function PsPauAdmin_EnumerateFlows
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminFlow
}
#gavdcodeend 002

#gavdcodebegin 003
Function PsPauAdmin_OwnerRole
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminFlowOwnerRole `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e"
}
#gavdcodeend 003

#gavdcodebegin 004
Function PsPauAdmin_UserDetails
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-AdminFlowUserDetails –UserId "acc28fcb-5261-47f8-960b-715d2f98a431"
}
#gavdcodeend 004

#gavdcodebegin 005
Function PsPauAdmin_DisableFlow
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Disable-AdminFlow `
					–EnvironmentName "ade56059-89c0-4594-90c3-e4772a8168ca" `
					–FlowName "ee720c41-2eef-4997-a471-539e343ec0f8"
}
#gavdcodeend 005

#gavdcodebegin 006
Function PsPauAdmin_EnableFlow
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Enable-AdminFlow `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e"
}
#gavdcodeend 006

#gavdcodebegin 007
Function PsPauAdmin_DeleteFlow
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Remove-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}
#gavdcodeend 007

#gavdcodebegin 040
Function PsPauAdmin_RestoreFlow
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Restore-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}
#gavdcodeend 040

#gavdcodebegin 008
Function PsPauAdmin_DeleteApprovalFlows
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Remove-AdminFlowApprovals `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 008

#gavdcodebegin 009
Function PsPauAdmin_AddRoleUser
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Set-AdminFlowOwnerRole `
					–EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444" `
					–FlowName "1ecbcc34-ec34-4220-b39e-bde8c961655e" `
					-PrincipalType User `
					-PrincipalObjectId "bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb" `
					-RoleName CanEdit
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsPauAdmin_DeleteRoleUser
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

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
Function PsPauAdmin_DeleteUserDetails
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Remove-AdminFlowUserDetails –UserId "092b1237-a428-45a7-b76b-310fdd6e7246"
}
#gavdcodeend 011

#gavdcodebegin 041
Function PsPauAdmin_CallApi
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
Function PsPauAdmin_GetTenantSettings
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-TenantSettings
}
#gavdcodeend 042

#gavdcodebegin 043
Function PsPauAdmin_GetTenantDetails
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-TenantDetailsFromGraph
}
#gavdcodeend 043

#gavdcodebegin 044
Function PsPauAdmin_GetUsersGroups
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-UsersOrGroupsFromGraph -SearchString "Admin"
	Get-UsersOrGroupsFromGraph -ObjectId "Admin@tenant.onmicrosoft.com"
}
#gavdcodeend 044

#gavdcodebegin 012
Function PsPau_RemoveFromSharePoint
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

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
Function PsPauMaker_EnumarateEnvironment
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-FlowEnvironment
}
#gavdcodeend 013

#gavdcodebegin 014
Function PsPauMaker_EnumarateFlows
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-Flow
	Write-Host "-------------------------"
	Get-Flow -EnvironmentName "c336e3a2-5a73-e274-b5ac-94dbc5a41444"
}
#gavdcodeend 014

#gavdcodebegin 015
Function PsPauMaker_GetOneFlow
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-FlowRun –FlowName "a1396da8-8184-4341-81da-2a7642fc1e8e"
}
#gavdcodeend 015

#gavdcodebegin 016
Function PsPauMaker_DisableFlow
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Disable-Flow –FlowName "a1396da8-8184-4341-81da-2a7642fc1e8e"
}
#gavdcodeend 016

#gavdcodebegin 017
Function PsPauMaker_EnableFlow
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Enable-Flow –FlowName "a1396da8-8184-4341-81da-2a7642fc1e8e"
}
#gavdcodeend 017

#gavdcodebegin 018
Function PsPauMaker_DeleteFlow
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Remove-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b" -Confirm:$false
}
#gavdcodeend 018

#gavdcodebegin 019
Function PsPauMaker_EnumarateFlowApprovals
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-FlowApproval –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 019

#gavdcodebegin 020
Function PsPauMaker_EnumarateFlowApprovalRequests
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-FlowApprovalRequest –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 020

#gavdcodebegin 021
Function PsPauMaker_ApproveFlows
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

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
Function PsPauMaker_RejectFlows
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

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
Function PsPauMaker_OwnerRole
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Get-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a"
}
#gavdcodeend 023

#gavdcodebegin 024
Function PsPauMaker_AddRoleUser
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

	Set-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a" `
					-PrincipalType User `
					-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 024

#gavdcodebegin 025
Function PsPauMaker_DeleteRoleUser
{
	PsPpPs_LoginWithAccPw -UserPw $cnfUserPw -UserName $cnfUserName

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
Function PsPauCli_GetAllFlowsByEnvironment
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw

	m365 flow list --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
				   --asAdmin

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
Function PsPauCli_GetOneFlow
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow get --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
				  --name "ee720c41-2eef-4997-a471-539e343ec0f8" `
				  --asAdmin

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
Function PsPauCli_ExportOneFlow
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow export --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
					 --name "ee720c41-2eef-4997-a471-539e343ec0f8" `
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
Function PsPauCli_DisableOneFlow
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow disable --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
					  --name "ee720c41-2eef-4997-a471-539e343ec0f8"

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
Function PsPauCli_EnableOneFlow
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow enable --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
					 --name "ee720c41-2eef-4997-a471-539e343ec0f8"

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
Function PsPauCli_DeleteOneFlow
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow remove --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
					 --name "ee720c41-2eef-4997-a471-539e343ec0f8"

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
Function PsPauCli_GetAllEnvironment
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow environment list

	m365 logout
}
#gavdcodeend 032

#gavdcodebegin 033
Function PsPauCli_GetOneEnvironment
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow environment get --name "Default-ade56059-89c0-4594-90c3-e4772a8168ca"

	m365 logout
}
#gavdcodeend 033

#gavdcodebegin 034
Function PsPauCli_GetAllConnectors
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow connector list --environmentName `
									"Default-ade56059-89c0-4594-90c3-e4772a8168ca"

	m365 logout
}
#gavdcodeend 034

#gavdcodebegin 035
Function PsPauCli_ExportOneConnectors
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow connector export `
					--environmentName "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					--connector "sh_con-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa" `
					--outputFolder "C:\Temp\MyConnector"

	m365 logout
}
#gavdcodeend 035

#gavdcodebegin 036
Function PsPauCli_GetRunsOneFlow
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow run list --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
					   --flowName "ee720c41-2eef-4997-a471-539e343ec0f8"

	m365 logout
}
#gavdcodeend 036

#gavdcodebegin 037
Function PsPauCli_GetOneRunOneFlow
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow run get --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
					  --flowName "ee720c41-2eef-4997-a471-539e343ec0f8" `
					  --name "08584461474529939526833772059CU01"

	m365 logout
}
#gavdcodeend 037

#gavdcodebegin 038
Function PsPauCli_ResubmitOneRunOneFlow
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow run resubmit --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
					       --flowName "ee720c41-2eef-4997-a471-539e343ec0f8" `
					       --name "08584461474529939526833772059CU01"

	m365 logout
}
#gavdcodeend 038

#gavdcodebegin 039
Function PsPauCli_CancelOneRunOneFlow
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 flow run cancel --environmentName "Default-ade56059-89c0-4594-90c3-e4772a8168ca" `
					     --flowName "ee720c41-2eef-4997-a471-539e343ec0f8" `
					     --name "08584461474529939526833772059CU01"

	m365 logout
}
#gavdcodeend 039

#-----------------------------------------------------------------------------------------


##==> Routines for PnPPowerShell

#gavdcodebegin 045
function PsPauPnpPowerShell_GetEnvironment
{
	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$true

	Write-Host "Environment Name: " $myEnvironment.Name
	Write-Host "Environment Id: " $myEnvironment.Id

	Disconnect-PnPOnline
}
#gavdcodeend 045

#gavdcodebegin 046
function PsPauPnpPowerShell_GetAllFlowsInEnvironment
{
	# Get-PnPFlow does not support pure Application permissions, even if you grant them, 
	#    the cmdlet will refuse to run in app‑only mode.

	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$true
	Get-PnPFlow -Environment $myEnvironment

	Disconnect-PnPOnline
}
#gavdcodeend 046

#gavdcodebegin 047
function PsPauPnpPowerShell_GetOneFlowInEnvironment
{
	# Get-PnPFlow does not support pure Application permissions, even if you grant them, 
	#    the cmdlet will refuse to run in app‑only mode.

	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint

	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$false
	Get-PnPFlow -Environment $myEnvironment `
				-Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 047

#gavdcodebegin 048
function PsPauPnpPowerShell_GetOneFlowRuns
{
	# Get-PnPFlowRun does not support pure Application permissions, even if you grant them, 
	#    the cmdlet will refuse to run in app‑only mode.

	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$false
	Get-PnPFlowRun -Environment $myEnvironment `
				   -Flow "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 048

#gavdcodebegin 049
function PsPauPnpPowerShell_DisableOneFlow
{
	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$false
	Disable-PnPFlow -Environment $myEnvironment `
				    -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 049

#gavdcodebegin 050
function PsPauPnpPowerShell_EnableOneFlow
{
	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$false
	Enable-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 050

#gavdcodebegin 051
function PsPauPnpPowerShell_StopOneFlow
{
	# Stop-PnPFlowRun does not support pure Application permissions, even if you grant them, 
	#    the cmdlet will refuse to run in app‑only mode.

	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint
	
	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$false
	Stop-PnPFlowRun -Environment $myEnvironment `
				    -Flow "1ecbcc34-ec34-4220-b39e-bde8c961655e" `
				    -Identity "08585321999201590891763396367CU168" `
					-Force

	Disconnect-PnPOnline
}
#gavdcodeend 051

#gavdcodebegin 052
function PsPauPnpPowerShell_RestartOneFlow
{	# Restart-PnPFlowRun does not support pure Application permissions, even if you grant them, 
	#    the cmdlet will refuse to run in app‑only mode.

	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint

	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$false
	Restart-PnPFlowRun -Environment $myEnvironment `
					   -Flow "1ecbcc34-ec34-4220-b39e-bde8c961655e" `
				       -Identity "08585321999201590891763396367CU168" `
					   -Force

	Disconnect-PnPOnline
}
#gavdcodeend 052

#gavdcodebegin 053
function PsPauPnpPowerShell_ExportOneFlow
{
	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint

	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$false
	Export-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 053

#gavdcodebegin 054
function PsPauPnpPowerShell_ExportOneFlowZip
{
	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint

	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$false
	Export-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e" `
				   -AsZipPackage `
				   -OutPath "C:\Temporary\myFlow.zip" `
				   -Verbose

	Disconnect-PnPOnline
}
#gavdcodeend 054

#gavdcodebegin 055
function PsPauPnpPowerShell_DeleteOneFlow
{
	$spCtx = PsPnPPowerShell_LoginGraphWithCertificateThumbprint `
									-SiteBaseUrl $cnfSiteCollUrl `
									-TenantName $cnfTenantName `
									-ClientId $cnfClientIdWithCert `
									-CertificateThumbprint $cnfCertificateThumbprint

	$myEnvironment = Get-PnPPowerPlatformEnvironment -IsDefault:$false
	Remove-PnPFlow -Environment $myEnvironment `
				   -Identity "1ecbcc34-ec34-4220-b39e-bde8c961655e"

	Disconnect-PnPOnline
}
#gavdcodeend 055


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 055 ***

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
#PsPauAdmin_EnumerateFlows
#PsPauAdmin_OwnerRole
#PsPauAdmin_UserDetails
#PsPauAdmin_DisableFlow
#PsPauAdmin_EnableFlow
#PsPauAdmin_DeleteFlow
#PsPauAdmin_RestoreFlow
#PsPauAdmin_DeleteApprovalFlows
#PsPauAdmin_AddRoleUser
#PsPauAdmin_DeleteRoleUser
#PsPauAdmin_DeleteUserDetails
#PsPauAdmin_CallApi
#PsPauAdmin_GetTenantSettings
#PsPauAdmin_GetTenantDetails
#PsPauAdmin_GetUsersGroups
#PsPau_RemoveFromSharePoint
#PsPauMaker_EnumarateEnvironment
#PsPauMaker_EnumarateFlows
#PsPauMaker_GetOneFlow
#PsPauMaker_DisableFlow
#PsPauMaker_EnableFlow
#PsPauMaker_DeleteFlow
#PsPauMaker_EnumarateFlowApprovals
#PsPauMaker_EnumarateFlowApprovalRequests
#PsPauMaker_ApproveFlows
#PsPauMaker_RejectFlows
#PsPauMaker_OwnerRole
#PsPauMaker_AddRoleUser
#PsPauMaker_DeleteRoleUser

##==> CLI
#PsPauCli_GetAllFlowsByEnvironment
#PsPauCli_GetOneFlow
#PsPauCli_ExportOneFlow
#PsPauCli_DisableOneFlow
#PsPauCli_EnableOneFlow
#PsPauCli_DeleteOneFlow
#PsPauCli_GetAllEnvironment
#PsPauCli_GetOneEnvironment
#PsPauCli_GetAllConnectors
#PsPauCli_ExportOneConnectors
#PsPauCli_GetRunsOneFlow
#PsPauCli_GetOneRunOneFlow
#PsPauCli_ResubmitOneRunOneFlow
#PsPauCli_CancelOneRunOneFlow

##==> PnPPowerShell
#PsPauPnpPowerShell_GetEnvironment
#PsPauPnpPowerShell_GetAllFlowsInEnvironment
#PsPauPnpPowerShell_GetOneFlowInEnvironment
#PsPauPnpPowerShell_GetOneFlowRuns
#PsPauPnpPowerShell_DisableOneFlow
#PsPauPnpPowerShell_EnableOneFlow
#PsPauPnpPowerShell_StopOneFlow
#PsPauPnpPowerShell_RestartOneFlow
#PsPauPnpPowerShell_ExportOneFlow
#PsPauPnpPowerShell_ExportOneFlowZip
#PsPauPnpPowerShell_DeleteOneFlow

Write-Host "Done"  
