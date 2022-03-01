
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

#----------------------------------------------------------------------------------------

##==> Routines for PowerShell

#gavdcodebegin 02
Function PowerAutomatePsAdminEnumarateFlows()
{
	Get-AdminFlow
}
#gavdcodeend 02

#gavdcodebegin 03
Function PowerAutomatePsAdminOwnerRole()
{
	Get-AdminFlowOwnerRole `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "4355d741-c54b-4372-9bb0-eb5b49285333"
}
#gavdcodeend 03

#gavdcodebegin 04
Function PowerAutomatePsAdminUserDetails()
{
	Get-AdminFlowUserDetails –UserId "092b1237-a428-45a7-b76b-310fdd6e7246"
}
#gavdcodeend 04

#gavdcodebegin 05
Function PowerAutomatePsAdminDisableFlow()
{
	Disable-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}
#gavdcodeend 05

#gavdcodebegin 06
Function PowerAutomatePsAdminEnableFlow()
{
	Enable-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}
#gavdcodeend 06

#gavdcodebegin 07
Function PowerAutomatePsAdminDeleteFlow()
{
	Remove-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}
#gavdcodeend 07

#gavdcodebegin 08
Function PowerAutomatePsAdminDeleteApprovalFlows()
{
	Remove-AdminFlowApprovals `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 08

#gavdcodebegin 09
Function PowerAutomatePsAdminAddRoleUser()
{
	Set-AdminFlowOwnerRole `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "4355d741-c54b-4372-9bb0-eb5b49285333" `
					-PrincipalType User `
					-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3" `
					-RoleName CanEdit
}
#gavdcodeend 09

#gavdcodebegin 10
Function PowerAutomatePsAdminDeleteRoleUser()
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
Function PowerAutomatePsAdminDeleteUserDetails()
{
	Remove-AdminFlowUserDetails –UserId "092b1237-a428-45a7-b76b-310fdd6e7246"
}
#gavdcodeend 11

#gavdcodebegin 12
Function PowerAutomatePsRemoveFromSharePoint()
{
	$AdminSiteURL="https://[domain]-admin.sharepoint.com/"
	$SiteURL="https://[domain].sharepoint.com/sites/[TeamSiteWithoutGroup]"
 
	Connect-SPOService -Url $AdminSiteURL -Credential (Get-Credential)
 
	Set-SPOSite -Identity $SiteURL -DisableFlows Disabled
 
	#To Enable the button: 
	#Set-SPOSite -Identity $SiteURL -DisableFlows NotDisabled
}
#gavdcodeend 12

#gavdcodebegin 13
Function PowerAutomatePsMakerEnumarateEnvironment()
{
	Get-FlowEnvironment
}
#gavdcodeend 13

#gavdcodebegin 14
Function PowerAutomatePsMakerEnumarateFlows()
{
	Get-Flow
}
#gavdcodeend 14

#gavdcodebegin 15
Function PowerAutomatePsMakerGetOneFlow()
{
	Get-FlowRun –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}
#gavdcodeend 15

#gavdcodebegin 16
Function PowerAutomatePsMakerDisableFlow()
{
	Disable-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}
#gavdcodeend 16

#gavdcodebegin 17
Function PowerAutomatePsMakerEnableFlow()
{
	Enable-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}
#gavdcodeend 17

#gavdcodebegin 18
Function PowerAutomatePsMakerDeleteFlow()
{
	Remove-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b" -Confirm:$false
}
#gavdcodeend 18

#gavdcodebegin 19
Function PowerAutomatePsMakerEnumarateFlowApprovals()
{
	Get-FlowApproval –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 19

#gavdcodebegin 20
Function PowerAutomatePsMakerEnumarateFlowApprovalRequests()
{
	Get-FlowApprovalRequest –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}
#gavdcodeend 20

#gavdcodebegin 21
Function PowerAutomatePsMakerApproveFlows()
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
Function PowerAutomatePsMakerRejectFlows()
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
Function PowerAutomatePsMakerOwnerRole()
{
	Get-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a"
}
#gavdcodeend 23

#gavdcodebegin 24
Function PowerAutomatePsMakerAddRoleUser()
{
	Set-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a" `
					-PrincipalType User `
					-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}
#gavdcodeend 24

#gavdcodebegin 25
Function PowerAutomatePsMakerDeleteRoleUser()
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
Function PowerAutomatePsCliGetAllFlowsByEnvironment()
{
	LoginPsCLI
	
	m365 flow list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
				   --asAdmin

	m365 logout
}
#gavdcodeend 26

#gavdcodebegin 27
Function PowerAutomatePsCliGetOneFlow()
{
	LoginPsCLI
	
	m365 flow get --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
				  --name "3126a4e4-71b9-49d8-802d-734a71534ff4" `
				  --asAdmin

	m365 logout
}
#gavdcodeend 27

#gavdcodebegin 28
Function PowerAutomatePsCliExportOneFlow()
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
Function PowerAutomatePsCliDisableOneFlow()
{
	LoginPsCLI
	
	m365 flow disable --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					  --name "3126a4e4-71b9-49d8-802d-734a71534ff4"

	m365 logout
}
#gavdcodeend 29

#gavdcodebegin 30
Function PowerAutomatePsCliEnableOneFlow()
{
	LoginPsCLI
	
	m365 flow enable --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					 --name "3126a4e4-71b9-49d8-802d-734a71534ff4"

	m365 logout
}
#gavdcodeend 30

#gavdcodebegin 31
Function PowerAutomatePsCliDeleteOneFlow()
{
	LoginPsCLI
	
	m365 flow remove --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					 --name "e7817682-71ce-4bee-b435-b98e0edfad12"

	m365 logout
}
#gavdcodeend 31

#gavdcodebegin 32
Function PowerAutomatePsCliGetAllEnvironment()
{
	LoginPsCLI
	
	m365 flow environment list

	m365 logout
}
#gavdcodeend 32

#gavdcodebegin 33
Function PowerAutomatePsCliGetOneEnvironment()
{
	LoginPsCLI
	
	m365 flow environment get --name "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 33

#gavdcodebegin 34
Function PowerAutomatePsCliGetAllConnectors()
{
	LoginPsCLI
	
	m365 flow connector list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052"

	m365 logout
}
#gavdcodeend 34

#gavdcodebegin 35
Function PowerAutomatePsCliExportOneConnectors()
{
	LoginPsCLI
	
	m365 flow connector export --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
							   --connector "sh_con-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa" `
							   --outputFolder "C:\Temp\MyConnector"

	m365 logout
}
#gavdcodeend 35

#gavdcodebegin 36
Function PowerAutomatePsCliGetRunsOneFlow()
{
	LoginPsCLI
	
	m365 flow run list --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					   --flow "3126a4e4-71b9-49d8-802d-734a71534ff4"

	m365 logout
}
#gavdcodeend 36

#gavdcodebegin 37
Function PowerAutomatePsCliGetOneRunOneFlow()
{
	LoginPsCLI
	
	m365 flow run get --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					  --flow "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					  --name "08585583736314022523222750393CU146"

	m365 logout
}
#gavdcodeend 37

#gavdcodebegin 38
Function PowerAutomatePsCliResubmitOneRunOneFlow()
{
	LoginPsCLI
	
	m365 flow run resubmit --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					       --flow "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					       --name "08585583736314022523222750393CU146"

	m365 logout
}
#gavdcodeend 38

#gavdcodebegin 39
Function PowerAutomatePsCliCancelOneRunOneFlow()
{
	LoginPsCLI
	
	m365 flow run cancel --environment "default-021ee864-951d-4f25-a5c3-b6d4412c4052" `
					     --flow "3126a4e4-71b9-49d8-802d-734a71534ff4" `
					     --name "08585583736314022523222750393CU146"

	m365 logout
}
#gavdcodeend 39

#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

##==> PowerShell
#LoginPsPowerPlatform

#PowerAutomatePsAdminEnumarateFlows
#PowerAutomatePsAdminOwnerRole
#PowerAutomatePsAdminUserDetails
#PowerAutomatePsAdminDisableFlow
#PowerAutomatePsAdminEnableFlow
#PowerAutomatePsAdminDeleteFlow
#PowerAutomatePsAdminDeleteApprovalFlows
#PowerAutomatePsAdminAddRoleUser
#PowerAutomatePsAdminDeleteRoleUser
#PowerAutomatePsAdminDeleteUserDetails
#PowerAutomatePsRemoveFromSharePoint
#PowerAutomatePsMakerEnumarateEnvironment
#PowerAutomatePsMakerEnumarateFlows
#PowerAutomatePsMakerGetOneFlow
#PowerAutomatePsMakerDisableFlow
#PowerAutomatePsMakerEnableFlow
#PowerAutomatePsMakerDeleteFlow
#PowerAutomatePsMakerEnumarateFlowApprovals
#PowerAutomatePsMakerEnumarateFlowApprovalRequests
#PowerAutomatePsMakerApproveFlows
#PowerAutomatePsMakerRejectFlows
#PowerAutomatePsMakerOwnerRole
#PowerAutomatePsMakerAddRoleUser
#PowerAutomatePsMakerDeleteRoleUser

##==> CLI
#PowerAutomatePsCliGetAllFlowsByEnvironment
#PowerAutomatePsCliGetOneFlow
#PowerAutomatePsCliExportOneFlow
#PowerAutomatePsCliDisableOneFlow
#PowerAutomatePsCliEnableOneFlow
#PowerAutomatePsCliDeleteOneFlow
#PowerAutomatePsCliGetAllEnvironment
#PowerAutomatePsCliGetOneEnvironment
#PowerAutomatePsCliGetAllConnectors
#PowerAutomatePsCliExportOneConnectors
#PowerAutomatePsCliGetRunsOneFlow
#PowerAutomatePsCliGetOneRunOneFlow
#PowerAutomatePsCliResubmitOneRunOneFlow
#PowerAutomatePsCliCancelOneRunOneFlow

Write-Host "Done"  
