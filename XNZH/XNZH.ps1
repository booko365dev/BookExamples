
Function LoginPsPowerPlatform()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.ppUserPw -AsPlainText -Force

	Add-PowerAppsAccount -Username $configFile.appsettings.ppUserName -Password $securePW
}

#----------------------------------------------------------------------------------------

Function PowerAutomatePsAdminEnumarateFlows()
{
	Get-AdminFlow
}

Function PowerAutomatePsAdminOwnerRole()
{
	Get-AdminFlowOwnerRole `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "4355d741-c54b-4372-9bb0-eb5b49285333"
}

Function PowerAutomatePsAdminUserDetails()
{
	Get-AdminFlowUserDetails –UserId "092b1237-a428-45a7-b76b-310fdd6e7246"
}

Function PowerAutomatePsAdminDisableFlow()
{
	Disable-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}

Function PowerAutomatePsAdminEnableFlow()
{
	Enable-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}

Function PowerAutomatePsAdminDeleteFlow()
{
	Remove-AdminFlow `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "9824f3b9-17ad-49fb-aa39-f54edcc0fd81"
}

Function PowerAutomatePsAdminDeleteApprovalFlows()
{
	Remove-AdminFlowApprovals `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}

Function PowerAutomatePsAdminAddRoleUser()
{
	Set-AdminFlowOwnerRole `
					–EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b" `
					–FlowName "4355d741-c54b-4372-9bb0-eb5b49285333" `
					-PrincipalType User `
					-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3" `
					-RoleName CanEdit
}

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

Function PowerAutomatePsAdminDeleteUserDetails()
{
	Remove-AdminFlowUserDetails –UserId "092b1237-a428-45a7-b76b-310fdd6e7246"
}

Function PowerAutomatePsRemoveFromSharePoint()
{
	$AdminSiteURL="https://[domain]-admin.sharepoint.com/"
	$SiteURL="https://[domain].sharepoint.com/sites/[TeamSiteWithoutGroup]"
 
	Connect-SPOService -Url $AdminSiteURL -Credential (Get-Credential)
 
	Set-SPOSite -Identity $SiteURL -DisableFlows Disabled
 
	#To Enable the button: 
	#Set-SPOSite -Identity $SiteURL -DisableFlows NotDisabled
}

Function PowerAutomatePsMakerEnumarateEnvironment()
{
	Get-FlowEnvironment
}

Function PowerAutomatePsMakerEnumarateFlows()
{
	Get-Flow
}

Function PowerAutomatePsMakerGetOneFlow()
{
	Get-FlowRun –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}

Function PowerAutomatePsMakerDisableFlow()
{
	Disable-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}

Function PowerAutomatePsMakerEnableFlow()
{
	Enable-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b"
}

Function PowerAutomatePsMakerDeleteFlow()
{
	Remove-Flow –FlowName "756899c1-0b22-40f0-b170-931698fd615b" -Confirm:$false
}

Function PowerAutomatePsMakerEnumarateFlowApprovals()
{
	Get-FlowApproval –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}

Function PowerAutomatePsMakerEnumarateFlowApprovalRequests()
{
	Get-FlowApprovalRequest –EnvironmentName "909ee029-5b74-4b2f-a9ee-b6b5158f630b"
}

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

Function PowerAutomatePsMakerOwnerRole()
{
	Get-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a"
}

Function PowerAutomatePsMakerAddRoleUser()
{
	Set-FlowOwnerRole `
					–EnvironmentName "Default-03d561bf-4472-41e0-b2d6-ee506471e9d0" `
					–FlowName "28327980-4786-4107-9f2e-80674c3cb98a" `
					-PrincipalType User `
					-PrincipalObjectId "959ae10e-0015-4948-b602-fbf7fccfe2a3"
}

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

#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ppPs.values.config"

LoginPsPowerPlatform

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

Write-Host "Done"  

