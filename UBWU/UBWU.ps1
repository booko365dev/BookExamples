
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


Function PsGraphRestApi_GetAzureTokenDelegationWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	 $LoginUrl = "https://login.microsoftonline.com"
	 $ScopeUrl = "https://graph.microsoft.com/.default"

	 $myBody  = @{ Scope = $ScopeUrl; `
					grant_type = "Password"; `
					client_id = $ClientID; `
					Username = $UserName; `
					Password = $UserPw }

	 $myOAuth = Invoke-RestMethod `
					-Method Post `
					-Uri $LoginUrl/$TenantName/oauth2/v2.0/token `
					-Body $myBody

	return $myOAuth
}

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

function PsGraphPowerShellSdk_LoginWithCertificateThumbprint
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateThumbprint
	)

	Connect-MgGraph -TenantId $TenantName -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


##==> Graph

#gavdcodebegin 001
Function PsPlannerGraphRestApi_GetAllPlansForOneGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$grpId = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $grpId + "/planner/plans"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$groupObject = ConvertFrom-Json –InputObject $myResult
	$groupObject.value.subject
}
#gavdcodeend 001 

#gavdcodebegin 002
Function PsPlannerGraphRestApi_CreateOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$grpId = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'owner':'" + $grpId + "', `
			     'title':'GraphPlan' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 002 

#gavdcodebegin 003
Function PsPlannerGraphRestApi_GetOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$planId = "JwJcILO6k0aFnTR5iljJ3JgAA8no"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$planObject = ConvertFrom-Json –InputObject $myResult
	$planObject.value.subject
}
#gavdcodeend 003 

#gavdcodebegin 004
Function PsPlannerGraphRestApi_UpdateOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$grpId = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	$planId = "JwJcILO6k0aFnTR5iljJ3JgAA8no"
	$eTag = 'W/"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'title':'GraphPlanUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 004 

#gavdcodebegin 005
Function PsPlannerGraphRestApi_GetOnePlanDetails
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$planId = "JwJcILO6k0aFnTR5iljJ3JgAA8no"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId + "/details"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$planObject = ConvertFrom-Json –InputObject $myResult
	$planObject.value.subject
}
#gavdcodeend 005 

#gavdcodebegin 006
Function PsPlannerGraphRestApi_UpdateOnePlanDetails
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$grpId = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$eTag = 'W/"JzEtUGxhbkRldGFpbHMgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId + "/details"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'sharedWith': { 'bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb':true, `
								 '3ce80572-cd16-4ddd-8e35-30782cf0db9d':true}, `
				 'categoryDescriptions': { 'category1': 'myLabel', `
										   'category2': null } }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 006 

#gavdcodebegin 007
Function PsPlannerGraphRestApi_GetAllBucketsInOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$planId = "JwJcILO6k0aFnTR5iljJ3JgAA8no"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId + "/buckets"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$bucketsObject = ConvertFrom-Json –InputObject $myResult
	$bucketsObject.value.subject
}
#gavdcodeend 007 

#gavdcodebegin 008
Function PsPlannerGraphRestApi_GetOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$bucketId = "_Mk8LLnUEkOVStBwuAdZtZgAFUFU"
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$bucketsObject = ConvertFrom-Json –InputObject $myResult
	$bucketsObject.value.subject
}
#gavdcodeend 008 

#gavdcodebegin 009
Function PsPlannerGraphRestApi_CreateOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$planId = "JwJcILO6k0aFnTR5iljJ3JgAA8no"
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'planId':'" + $planId + "', `
			     'name':'BucketOne', ` 
				 'orderHint': ' !' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 009 

#gavdcodebegin 010
Function PsPlannerGraphRestApi_UpdateOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$bucketId = "D3hsxxsWiUuk0As8CVdpk5gAAugr"
	$eTag = 'W/"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'name':'BucketOneUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 010 

#gavdcodebegin 011
Function PsPlannerGraphRestApi_DeleteOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$bucketId = "0e70kPhvY0ueVmVlf50hbZgAOOxj"
	$eTag = 'W/"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId

	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 011 

#gavdcodebegin 012
Function PsPlannerGraphRestApi_DeleteOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$eTag = 'W/"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBAUCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId

	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 012 

#gavdcodebegin 013
Function PsPlannerGraphRestApi_GetAllTasksInOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$bucketId = "iRlt3u8lc0C0rUHO-q35JpgAOw_F"
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId + "/tasks"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$tasksObject = ConvertFrom-Json –InputObject $myResult
	$tasksObject.value.subject
}
#gavdcodeend 013 

#gavdcodebegin 014
Function PsPlannerGraphRestApi_GetOneTask
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$taskId = "dp3JAX9my0uDbpoV0_26gpgALTDm"
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$taskObject = ConvertFrom-Json –InputObject $myResult
	$taskObject.value.subject
}
#gavdcodeend 014 

#gavdcodebegin 015
Function PsPlannerGraphRestApi_CreateOneTask
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$planId = "JwJcILO6k0aFnTR5iljJ3JgAA8no"
	$bucketId = "iRlt3u8lc0C0rUHO-q35JpgAOw_F"
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'planId':'" + $planId + "', `
			     'bucketId':'" + $bucketId + "', ` 
			     'title':'TaskOne', `
				 'assignments': {
					'acc28fcb-5261-47f8-960b-715d2f98a431': {
					'@odata.type': '#microsoft.graph.plannerAssignment',
					'orderHint': ' !' }}}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 015 

#gavdcodebegin 016
Function PsPlannerGraphRestApi_UpdateOneTask
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$taskId = "kqTGSJuN_kC-lUCGi7ruDJgAF46C"
	$eTag = 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'percentComplete':1 }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 016 

#gavdcodebegin 017
Function PsPlannerGraphRestApi_DeleteOneTask
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$taskId = "dp3JAX9my0uDbpoV0_26gpgALTDm"
	$eTag = 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId

	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 017 

#gavdcodebegin 053
Function PsPlannerGraphRestApi_GetOneTaskDetails
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$taskId = "tJqIvX1FwE6ixDnExZnYCpgAHErb"
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId + "/details"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$planObject = ConvertFrom-Json –InputObject $myResult
	$planObject.value.subject
}
#gavdcodeend 053 

#gavdcodebegin 054
Function PsPlannerGraphRestApi_UpdateOneTaskDetails
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$taskId = "tJqIvX1FwE6ixDnExZnYCpgAHErb"
	$eTag = 'W/"JzEtVGFza0RldGFpbHMgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId + "/details"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'checklist': {'12345678-1234-1234-1234-1234567890ab': { `
								'@odata.type': 'microsoft.graph.plannerChecklistItem', `
								'title': 'My Task details', `
								'isChecked': true } }, `
				 'description': 'This is a Task' , `
				 'previewType': 'noPreview' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 054 

#gavdcodebegin 055
Function PsPlannerGraphRestApi_GetTasksOneUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/me/planner/tasks"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$planObject = ConvertFrom-Json –InputObject $myResult
	$planObject.value.subject
}
#gavdcodeend 055

#-----------------------------------------------------------------------------------------

##==> CLI

#gavdcodebegin 018
Function PsPlannerCliM365_GetAllPlans
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner plan list --ownerGroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	Write-Host ("-------")
	m365 planner plan list --ownerGroupName "Chapter18"

	m365 logout
}
#gavdcodeend 018

#gavdcodebegin 019
Function PsPlannerCliM365_GetPlansByQuery
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner plan list --ownerGroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" `
						   --output json `
						   --query "[?title == 'Plan01']"

	m365 logout
}
#gavdcodeend 019

#gavdcodebegin 020
Function PsPlannerCliM365_GetOnePlan
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner plan get --id "QP5LN-ygA0GnuHHDLKjY-JgACTnA"

	m365 logout
}
#gavdcodeend 020

#gavdcodebegin 021
Function PsPlannerCliM365_CreateOnePlan
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner plan add --title "PlanCreatedWithCLI" `
						  --ownerGroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"

	m365 logout
}
#gavdcodeend 021

#gavdcodebegin 056
Function PsPlannerCliM365_UpdateOnePlan
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner plan set --id "whr8psBnkkuZ6QPl4EfBc5gABQpF" `
						  --newTitle "PlanUpdatedWithCLI" `
						  --shareWithUserNames "user@tenant.onmicrosoft.com" `
						  --category1 "My Category"

	m365 logout
}
#gavdcodeend 056

#gavdcodebegin 057
Function PsPlannerCliM365_DeletePlan
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner plan remove --id "whr8psBnkkuZ6QPl4EfBc5gABQpF"

	#m365 planner plan remove --title "PlanUpdatedWithCLI" `
	#						 --ownerGroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" `
	#						 --confirm

	m365 logout
}
#gavdcodeend 057

#gavdcodebegin 022
Function PsPlannerCliM365_GetAllBuckets
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner bucket list --planId "JwJcILO6k0aFnTR5iljJ3JgAA8no"
	Write-Host ("-------")
	m365 planner bucket list --planTitle "PlanUpdatedWithCLI" `
							 --ownerGroupName "Chapter18"

	m365 logout
}
#gavdcodeend 022

#gavdcodebegin 023
Function PsPlannerCliM365_GetBucketsByQuery
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner bucket list --planId "whr8psBnkkuZ6QPl4EfBc5gABQpF" `
						     --output json `
						     --query "[?name == 'To do']"

	m365 logout
}
#gavdcodeend 023

#gavdcodebegin 024
Function PsPlannerCliM365_CreateOneBucket
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner bucket add --name "BucketCreatedWithCLI" `
						    --planId "whr8psBnkkuZ6QPl4EfBc5gABQpF" `
							--orderHint " !"

	m365 logout
}
#gavdcodeend 024

#gavdcodebegin 058
Function PsPlannerCliM365_GetOneBucket
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner bucket get --id "r5bnTWCfnUGasiYREdqWPpgAOrFy"
	Write-Host ("-------")
	m365 planner bucket get --name "BucketCreatedWithCLI" `
							--planId "whr8psBnkkuZ6QPl4EfBc5gABQpF"
	Write-Host ("-------")
	m365 planner bucket get --name "BucketCreatedWithCLI" `
							--planTitle "PlanUpdatedWithCLI" `
							--ownerGroupName "Chapter18"

	m365 logout
}
#gavdcodeend 058

#gavdcodebegin 059
Function PsPlannerCliM365_UpdateOneBucket
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner bucket set --id "r5bnTWCfnUGasiYREdqWPpgAOrFy" `
							--newName "BucketUpdatedWithCLI"

	m365 logout
}
#gavdcodeend 059

#gavdcodebegin 060
Function PsPlannerCliM365_DeleteBucket
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner bucket remove --id "r5bnTWCfnUGasiYREdqWPpgAOrFy"

	#m365 planner bucket remove --name "BucketUpdatedWithCLI" `
	#						    --planId "whr8psBnkkuZ6QPl4EfBc5gABQpF"

	#m365 planner bucket remove --name "BucketUpdatedWithCLI" `
	#						    --planTitle "PlanUpdatedWithCLI" `
	#						    --ownerGroupName "Chapter18"

	m365 logout
}
#gavdcodeend 060

#gavdcodebegin 025
Function PsPlannerCliM365_GetAllTasks
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task list --planId "whr8psBnkkuZ6QPl4EfBc5gABQpF" `
						   --bucketId "r5bnTWCfnUGasiYREdqWPpgAOrFy"
	Write-Host ("-------")
	m365 planner task list --ownerGroupName "Chapter18" `
						   --planTitle "PlanUpdatedWithCLI" `
						   --bucketName "BucketUpdatedWithCLI"

	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
Function PsPlannerCliM365_GetTasksByQuery
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task list --planId "whr8psBnkkuZ6QPl4EfBc5gABQpF" `
						   --output json `
						   --query "[?title == 'Task01']"

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 061
Function PsPlannerCliM365_CreateOneTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task add --title "TaskCreatedWithCLI" `
						  --ownerGroupName "Chapter18" `
						  --planTitle "PlanUpdatedWithCLI" `
						  --bucketId "r5bnTWCfnUGasiYREdqWPpgAOrFy" `
						  --percentComplete 45 `
						  --priority "Important" `
						  --assignedToUserNames "a@d.onmicrosoft.com,b@d.onmicrosoft.com"

	m365 logout
}
#gavdcodeend 061

#gavdcodebegin 062
Function PsPlannerCliM365_GetOneTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task get --id "9lrWLlboRkyy3A3cBdQ-9ZgADGQY"
	Write-Host ("-------")
	m365 planner task get --title "TaskCreatedWithCLI" `
						  --bucketName "BucketUpdatedWithCLI" `
						  --planTitle "PlanUpdatedWithCLI" `
						  --ownerGroupName "Chapter18"

	m365 logout
}
#gavdcodeend 062

#gavdcodebegin 063
Function PsPlannerCliM365_UpdateOneTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task set --id "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
						  --title "TaskUpdatedWithCLI" `
						  --percentComplete 100 `
						  --appliedCategories "category1,category2"


	m365 logout
}
#gavdcodeend 063

#gavdcodebegin 064
Function PsPlannerCliM365_GetAllCheckListInTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task checklistitem list --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY"

	m365 logout
}
#gavdcodeend 064

#gavdcodebegin 065
Function PsPlannerCliM365_AddOneCheckListToTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task checklistitem add --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
										--title "Checklist Item CLI" `
										--isChecked

	m365 logout
}
#gavdcodeend 065

#gavdcodebegin 066
Function PsPlannerCliM365_DeleteOneCheckListFromTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task checklistitem remove --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
										   --id "8fe00b42-68b5-459e-9a71-8ff56bc96bee" `
										   --confirm

	m365 logout
}
#gavdcodeend 066

#gavdcodebegin 067
Function PsPlannerCliM365_GetAllAttachmentsInTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task reference list --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY"

	m365 logout
}
#gavdcodeend 067

#gavdcodebegin 068
Function PsPlannerCliM365_AddOneAttachmentToTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task reference add --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
									--url "https://guitaca.com" `
									--type "Other" `
									--alias "Guitaca Publishers"

	m365 logout
}
#gavdcodeend 068

#gavdcodebegin 069
Function PsPlannerCliM365_DeleteOneAttachmentFromTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task reference remove --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
									   --alias "Guitaca Publishers" `
									   --confirm
	
	#m365 planner task reference remove --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
	#								   --url "https://guitaca.com" 

	m365 logout
}
#gavdcodeend 069

#gavdcodebegin 070
Function PsPlannerCliM365_DeleteOneTask
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 planner task remove --id "9lrWLlboRkyy3A3cBdQ-9ZgADGQY"

	#m365 planner task remove --title "TaskUpdatedWithCLI" `
	#						 --bucketName "BucketUpdatedWithCLI" `
	#						 --planTitle "PlanUpdatedWithCLI" `
	#						 --ownerGroupName "Chapter18"

	m365 logout
}
#gavdcodeend 070

#-----------------------------------------------------------------------------------------

##==> PnP

#gavdcodebegin 027
Function PsPlannerPnP_GetPlansByGroup
{
	# App Registration permissions: Group.Read.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw

	Get-PnPPlannerPlan -Group "Chapter18"
}
#gavdcodeend 027

#gavdcodebegin 028
Function PsPlannerPnP_GetPlansByGroupAndPlan
{
	# App Registration permissions: Group.Read.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Get-PnPPlannerPlan -Group "Chapter18" -Identity "Plan01"
}
#gavdcodeend 028

#gavdcodebegin 029
Function PsPlannerPnP_CreatePlan
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	New-PnPPlannerPlan -Group "Chapter18" -Title "PlanCreatedWithPnP"
}
#gavdcodeend 029

#gavdcodebegin 030
Function PsPlannerPnP_UpdatePlanByPlan
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Set-PnPPlannerPlan -Group "Chapter18" `
					   -Plan "PlanCreatedWithPnP" `
					   -Title "PlanUpdatedWithPnP"
}
#gavdcodeend 030

#gavdcodebegin 031
Function PsPlannerPnP_UpdatePlanById
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Set-PnPPlannerPlan -PlanId "O3tqA06E_kaVBs0lVOAgbZgAAsQq" `
					   -Title "PlanUpdatedWithPnPById"
}
#gavdcodeend 031

#gavdcodebegin 032
Function PsPlannerPnP_DeletePlan
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Remove-PnPPlannerPlan -Group "Chapter18" `
						  -Identity "O3tqA06E_kaVBs0lVOAgbZgAAsQq"
}
#gavdcodeend 032

#gavdcodebegin 033
Function PsPlannerPnP_GetBucketsByGroupAndPlan
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Get-PnPPlannerBucket -Group "Chapter18" -Plan "AnotherPlanCreatedWithPnP"
}
#gavdcodeend 033

#gavdcodebegin 034
Function PsPlannerPnP_GetBucketsById
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Get-PnPPlannerBucket -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF"
}
#gavdcodeend 034

#gavdcodebegin 035
Function PsPlannerPnP_CreateBucketById
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Add-PnPPlannerBucket -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF" `
						 -Name "BucketCreatedWithPnP"
}
#gavdcodeend 035

#gavdcodebegin 036
Function PsPlannerPnP_CreateBucketByGroupAndPlan
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Add-PnPPlannerBucket -Group "Chapter18" `
						 -Plan "AnotherPlanCreatedWithPnP" `
						 -Name "BucketCreatedWithPnPByGroupAndPlan"
}
#gavdcodeend 036

#gavdcodebegin 037
Function PsPlannerPnP_UpdateBucketById
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Set-PnPPlannerBucket -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF" `
						 -Bucket "BucketCreatedWithPnP" `
						 -Name "BucketUpdatedWithPnP"
}
#gavdcodeend 037

#gavdcodebegin 038
Function PsPlannerPnP_UpdateBucketByGroupAndPlan
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Set-PnPPlannerBucket -Group "Chapter18" `
						 -Plan "AnotherPlanCreatedWithPnP" `
						 -Bucket "BucketCreatedWithPnPByGroupAndPlan" `
						 -Name "BucketUpdatedWithPnPByGroupAndPlan"
}
#gavdcodeend 038

#039 removed from the source code (update 2023-03)

#gavdcodebegin 040
Function PsPlannerPnP_DeleteBucketByGroupAndPlan
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Remove-PnPPlannerBucket -Group "Chapter18" `
							-Plan "AnotherPlanCreatedWithPnP" `
							-Identity "BucketUpdatedWithPnPByGroupAndPlan"
}
#gavdcodeend 040

#gavdcodebegin 041
Function PsPlannerPnP_GetTasksByPlanId
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Get-PnPPlannerTask -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF"
}
#gavdcodeend 041

#gavdcodebegin 042
Function PsPlannerPnP_GetTasksByBucketId
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Get-PnPPlannerTask -Bucket "BsZfkoF2iEmWB23IdPDxEJgAKDhP"
}
#gavdcodeend 042

#gavdcodebegin 043
Function PsPlannerPnP_GetTasksByTaskId
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Get-PnPPlannerTask -TaskId "tVlliyvnL0GqamF0wkzGeZgAD7vF" `
					   -IncludeDetails `
					   -ResolveUserDisplayNames
}
#gavdcodeend 043

#gavdcodebegin 044
Function PsPlannerPnP_GetTasksByGroupAndPlan
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Get-PnPPlannerTask -Group "Chapter18" `
					   -Plan "AnotherPlanCreatedWithPnP"
}
#gavdcodeend 044

#gavdcodebegin 045
Function PsPlannerPnP_CreateTaskByGroupAndPlan
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Add-PnPPlannerTask -Group "Chapter18" `
					   -Plan "AnotherPlanCreatedWithPnP" `
					   -Bucket "BucketUpdatedWithPnP" `
					   -Title "TaskCreatedWithPnP" `
					   -AssignedTo "user@domain.onmicrosoft.com"
}
#gavdcodeend 045

#gavdcodebegin 046
Function PsPlannerPnP_CreateTaskById
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Add-PnPPlannerTask -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF" `
					   -Bucket "BucketUpdatedWithPnP" `
					   -Title "TaskCreatedWithPnPById" `
}
#gavdcodeend 046

#gavdcodebegin 047
Function PsPlannerPnP_UpdateTaskById
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Set-PnPPlannerTask -TaskId "tVlliyvnL0GqamF0wkzGeZgAD7vF" `
					   -Title "TaskUpdatedWithPnP" `
					   -AssignedTo "user1@dom.onmicrosoft.com","user2@dom.onmicrosoft.com"
}
#gavdcodeend 047

#gavdcodebegin 048
Function PsPlannerPnP_DeleteTaskById
{
	# App Registration permissions: Group.ReadWrite.All

	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Remove-PnPPlannerTask -Task "tVlliyvnL0GqamF0wkzGeZgAD7vF"
}
#gavdcodeend 048

#gavdcodebegin 049
Function PsPlannerPnP_GetUserPolicy
{
	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Get-PnPPlannerUserPolicy -Identity "user@domain.onmicrosoft.com"
}
#gavdcodeend 049

#gavdcodebegin 050
Function PsPlannerPnP_SetUserPolicy
{
	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Set-PnPPlannerUserPolicy -Identity "user@domain.onmicrosoft.com" `
							 -BlockDeleteTasksNotCreatedBySelf $true
}
#gavdcodeend 050

#gavdcodebegin 051
Function PsPlannerPnP_GetConfiguration
{
	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Get-PnPPlannerConfiguration
}
#gavdcodeend 051

#gavdcodebegin 052
Function PsPlannerPnP_SetConfiguration
{
	PsPnpPowerShell_LoginWithAccPw `
				$cnfUserName $cnfUserPw $cnfSiteCollUrl $cnfClientIdWithAccPw
	
	Set-PnPPlannerConfiguration -AllowCalendarSharing $false
}
#gavdcodeend 052

#-----------------------------------------------------------------------------------------

##==> Graph SDK

#gavdcodebegin 071
Function PsPlannerGraphSdk_GetAllPlansInGroup
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgGroupPlannerPlan -GroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"

	Disconnect-MgGraph
}
#gavdcodeend 071

#gavdcodebegin 072
Function PsPlannerGraphSdk_GetOnePlan
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgGroupPlannerPlan -GroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" `
						   -PlannerPlanId "QP5LN-ygA0GnuHHDLKjY-JgACTnA"

	Get-MgPlannerPlan -PlannerPlanId "QP5LN-ygA0GnuHHDLKjY-JgACTnA"

	Disconnect-MgGraph
}
#gavdcodeend 072

#gavdcodebegin 073
Function PsPlannerGraphSdk_CreatePlanWithBodyParameters
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	$PlanParameters = @{
		Owner = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
		Title = "PlanCreatedWithGraphSDKBodyParams"
	}
	New-MgPlannerPlan -BodyParameter $PlanParameters

	Disconnect-MgGraph
}
#gavdcodeend 073

#gavdcodebegin 074
Function PsPlannerGraphSdk_CreatePlanWithParameters
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	New-MgPlannerPlan -Owner "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" `
					  -Title "PlanCreatedWithGraphSDKParams"

	Disconnect-MgGraph
}
#gavdcodeend 074

#gavdcodebegin 075
Function PsPlannerGraphSdk_UpdatePlanWithBodyParameters
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	$PlanParameters = @{
		Title = "PlanUpdatedWithGraphSDKParams"
	}
	Update-MgPlannerPlan -PlannerPlanId "465_rAMzX0eeXcvMIR8945gADbIY" `
						 -BodyParameter $PlanParameters

	Disconnect-MgGraph
}
#gavdcodeend 075

#gavdcodebegin 076
Function PsPlannerGraphSdk_DeletePlan
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	$planId = "465_rAMzX0eeXcvMIR8945gADbIY"
	$myPlan = Get-MgPlannerPlan -PlannerPlanId $planId
	$myEtag = $myPlan.AdditionalProperties.'@odata.etag'
	Remove-MgPlannerPlan -PlannerPlanId $planId `
						 -IfMatch $myEtag

	Disconnect-MgGraph
}
#gavdcodeend 076

#gavdcodebegin 077
Function PsPlannerGraphSdk_GetAllBucketsInPlan
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgPlannerPlanBucket -PlannerPlanId "kCJFZHc0jEqn5wBDnPRQv5gAH1_n"

	Disconnect-MgGraph
}
#gavdcodeend 077

#gavdcodebegin 078
Function PsPlannerGraphSdk_GetOneBucket
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgPlannerBucket -PlannerBucketId "sJQ8XSFfRUqvL6k_0OG1xpgAGUgE"

	Disconnect-MgGraph
}
#gavdcodeend 078

#gavdcodebegin 079
Function PsPlannerGraphSdk_CreateBucketWithBodyParameters
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	$bucketParameters = @{
		Name = "BucketCreatedWithGraphSDKBodyParams"
		PlanId = "kCJFZHc0jEqn5wBDnPRQv5gAH1_n"
		OrderHint = " !"
	}
	New-MgPlannerBucket -BodyParameter $bucketParameters

	Disconnect-MgGraph
}
#gavdcodeend 079

#gavdcodebegin 080
Function PsPlannerGraphSdk_CreateBucketWithParameters
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	New-MgPlannerBucket -PlanId "kCJFZHc0jEqn5wBDnPRQv5gAH1_n" `
						-Name "BucketCreatedWithGraphSDKParamsAA"

	Disconnect-MgGraph
}
#gavdcodeend 080

#gavdcodebegin 081
Function PsPlannerGraphSdk_UpdateBucketWithBodyParameters
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	$bucketParameters = @{
		Name = "BucketUpdatedWithGraphSDKBodyParams"
	}
	Update-MgPlannerBucket -PlannerBucketId "RqSFxZYBF06Y8TslksbV9ZgAP9-6" `
						   -BodyParameter $bucketParameters

	Disconnect-MgGraph
}
#gavdcodeend 081

#gavdcodebegin 082
Function PsPlannerGraphSdk_DeleteBucket
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	$bucketId = "RqSFxZYBF06Y8TslksbV9ZgAP9-6"
	$myBucket = Get-MgPlannerBucket -PlannerBucketId $bucketId
	$myEtag = $myBucket.AdditionalProperties.'@odata.etag'
	Remove-MgPlannerBucket -PlannerBucketId $bucketId `
						   -IfMatch $myEtag

	Disconnect-MgGraph
}
#gavdcodeend 082

#gavdcodebegin 083
Function PsPlannerGraphSdk_GetAllTasksInPlan
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgPlannerPlanTask -PlannerPlanId "kCJFZHc0jEqn5wBDnPRQv5gAH1_n"

	Disconnect-MgGraph
}
#gavdcodeend 083

#gavdcodebegin 084
Function PsPlannerGraphSdk_GetAllTasksInBucket
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgPlannerBucketTask -PlannerBucketId "1JSb34_9e0y71BgN-Z7jeJgAGdI5"

	Disconnect-MgGraph
}
#gavdcodeend 084

#gavdcodebegin 085
Function PsPlannerGraphSdk_GetOneTask
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgPlannerTask -PlannerTaskId "h-WwrIBusEW906Q9asal8ZgALG8c"

	Disconnect-MgGraph
}
#gavdcodeend 085

#gavdcodebegin 086
Function PsPlannerGraphSdk_GetOneTaskDetail
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgPlannerTaskDetail -PlannerTaskId "h-WwrIBusEW906Q9asal8ZgALG8c"

	Disconnect-MgGraph
}
#gavdcodeend 086

#gavdcodebegin 087
Function PsPlannerGraphSdk_CreateTaskWithBodyParameters
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	$bucketParameters = @{
		PlanId = "kCJFZHc0jEqn5wBDnPRQv5gAH1_n"
		BucketId = "1JSb34_9e0y71BgN-Z7jeJgAGdI5"
		Title = "BucketCreatedWithGraphSDKBodyParams"
	}
	New-MgPlannerTask -BodyParameter $bucketParameters

	Disconnect-MgGraph
}
#gavdcodeend 087

#gavdcodebegin 088
Function PsPlannerGraphSdk_CreateTaskWithParameters
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	New-MgPlannerTask -PlanId "kCJFZHc0jEqn5wBDnPRQv5gAH1_n" `
					  -BucketId "1JSb34_9e0y71BgN-Z7jeJgAGdI5" `
					  -Title "BucketCreatedWithGraphSDKParams"

	Disconnect-MgGraph
}
#gavdcodeend 088

#gavdcodebegin 089
Function PsPlannerGraphSdk_UpdateTaskWithBodyParameters
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	$bucketParameters = @{
		Title = "BucketUpdatedWithGraphSDKBodyParams"
		AppliedCategories = @{
			Category10 = $true
			Category11 = $false
		}
	}
	Update-MgPlannerTask -PlannerTaskId "cs8qYcpUSEKaIjuzePRtQJgAKt0x" `
						 -BodyParameter $bucketParameters

	Disconnect-MgGraph
}
#gavdcodeend 089

#gavdcodebegin 090
Function PsPlannerGraphSdk_DeleteTask
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	$taskId = "cs8qYcpUSEKaIjuzePRtQJgAKt0x"
	$myTask = Get-MgPlannerTask -PlannerTaskId $TaskId
	$myEtag = $myTask.AdditionalProperties.'@odata.etag'
	Remove-MgPlannerTask -PlannerTaskId $taskId `
						 -IfMatch $myEtag

	Disconnect-MgGraph
}
#gavdcodeend 090

#gavdcodebegin 091
Function PsPlannerGraphSdk_GetPlannerForUser
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgUserPlanner -UserId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	Get-MgUserPlanner -UserId "user@domain.onmicrosoft.com"

	Disconnect-MgGraph
}
#gavdcodeend 091

#gavdcodebegin 092
Function PsPlannerGraphSdk_GetPlansForUser
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgUserPlannerPlan -UserId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	Get-MgUserPlannerPlan -UserId "user@domain.onmicrosoft.com"

	Disconnect-MgGraph
}
#gavdcodeend 092

#gavdcodebegin 093
Function PsPlannerGraphSdk_GetTasksForUser
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint `
							 -TenantName $cnfTenantName `
							 -ClientID $cnfClientIdWithCert `
							 -CertificateThumbprint $cnfCertificateThumbprint

	Get-MgUserPlannerTask -UserId "acc28fcb-5261-47f8-960b-715d2f98a431"
	Get-MgUserPlannerTask -UserId "user@domain.onmicrosoft.com"

	Disconnect-MgGraph
}
#gavdcodeend 093


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

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

# *** Latest Source Code Index: 093 ***

#------------------------ Using Microsoft Graph PowerShell

#PsPlannerGraphRestApi_GetAllPlansForOneGroup
#PsPlannerGraphRestApi_CreateOnePlan
#PsPlannerGraphRestApi_GetOnePlan
#PsPlannerGraphRestApi_UpdateOnePlan
#PsPlannerGraphRestApi_GetOnePlanDetails
#PsPlannerGraphRestApi_UpdateOnePlanDetails
#PsPlannerGraphRestApi_GetAllBucketsInOnePlan
#PsPlannerGraphRestApi_GetOneBucket
#PsPlannerGraphRestApi_CreateOneBucket
#PsPlannerGraphRestApi_UpdateOneBucket
#PsPlannerGraphRestApi_DeleteOneBucket
#PsPlannerGraphRestApi_GetAllTasksInOneBucket
#PsPlannerGraphRestApi_GetOneTask
#PsPlannerGraphRestApi_GetOneTaskDetails
#PsPlannerGraphRestApi_UpdateOneTaskDetails
#PsPlannerGraphRestApi_GetTasksOneUser
#PsPlannerGraphRestApi_CreateOneTask
#PsPlannerGraphRestApi_UpdateOneTask
#PsPlannerGraphRestApi_DeleteOneTask
#PsPlannerGraphRestApi_DeleteOnePlan

#------------------------ Using PnP CLI

#PsPlannerCliM365_GetAllPlans
#PsPlannerCliM365_GetPlansByQuery
#PsPlannerCliM365_GetOnePlan
#PsPlannerCliM365_CreateOnePlan
#PsPlannerCliM365_UpdateOnePlan
#PsPlannerCliM365_DeletePlan
#PsPlannerCliM365_GetAllBuckets
#PsPlannerCliM365_GetBucketsByQuery
#PsPlannerCliM365_CreateOneBucket
#PsPlannerCliM365_GetOneBucket
#PsPlannerCliM365_UpdateOneBucket
#PsPlannerCliM365_DeleteBucket
#PsPlannerCliM365_GetAllTasks
#PsPlannerCliM365_GetTasksByQuery
#PsPlannerCliM365_CreateOneTask
#PsPlannerCliM365_GetOneTask
#PsPlannerCliM365_UpdateOneTask
#PsPlannerCliM365_GetAllCheckListInTask
#PsPlannerCliM365_AddOneCheckListToTask
#PsPlannerCliM365_DeleteOneCheckListFromTask
#PsPlannerCliM365_GetAllAttachmentsInTask
#PsPlannerCliM365_AddOneAttachmentToTask
#PsPlannerCliM365_DeleteOneAttachmentFromTask
#PsPlannerCliM365_DeleteOneTask

#------------------------ Using PowerShell PnP

#PsPlannerPnP_GetPlansByGroup
#PsPlannerPnP_GetPlansByGroupAndPlan
#PsPlannerPnP_CreatePlan
#PsPlannerPnP_UpdatePlanByPlan
#PsPlannerPnP_UpdatePlanById
#PsPlannerPnP_DeletePlan
#PsPlannerPnP_GetBucketsByGroupAndPlan
#PsPlannerPnP_GetBucketsById
#PsPlannerPnP_CreateBucketById
#PsPlannerPnP_CreateBucketByGroupAndPlan
#PsPlannerPnP_UpdateBucketById
#PsPlannerPnP_UpdateBucketByGroupAndPlan
#PsPlannerPnP_DeleteBucketByGroupAndPlan
#PsPlannerPnP_GetTasksByPlanId
#PsPlannerPnP_GetTasksByBucketId
#PsPlannerPnP_GetTasksByTaskId
#PsPlannerPnP_GetTasksByGroupAndPlan
#PsPlannerPnP_CreateTaskByGroupAndPlan
#PsPlannerPnP_CreateTaskById
#PsPlannerPnP_UpdateTaskById
#PsPlannerPnP_DeleteTaskById
#PsPlannerPnP_GetUserPolicy
#PsPlannerPnP_SetUserPolicy
#PsPlannerPnP_GetConfiguration
#PsPlannerPnP_SetConfiguration
#PsPlannerPnP_CreatePlannerRoster

#------------------------ Using Microsoft Graph PowerShell SDK

#PsPlannerGraphSdk_GetAllPlansInGroup
#PsPlannerGraphSdk_GetOnePlan
#PsPlannerGraphSdk_CreatePlanWithBodyParameters
#PsPlannerGraphSdk_CreatePlanWithParameters
#PsPlannerGraphSdk_UpdatePlanWithBodyParameters
#PsPlannerGraphSdk_DeletePlan
#PsPlannerGraphSdk_GetAllBucketsInPlan
#PsPlannerGraphSdk_GetOneBucket
#PsPlannerGraphSdk_CreateBucketWithBodyParameters
#PsPlannerGraphSdk_CreateBucketWithParameters
#PsPlannerGraphSdk_UpdateBucketWithBodyParameters
#PsPlannerGraphSdk_DeleteBucket
#PsPlannerGraphSdk_GetAllTasksInPlan
#PsPlannerGraphSdk_GetAllTasksInBucket
#PsPlannerGraphSdk_GetOneTask
#PsPlannerGraphSdk_GetOneTaskDetail
#PsPlannerGraphSdk_CreateTaskWithBodyParameters
#PsPlannerGraphSdk_CreateTaskWithParameters
#PsPlannerGraphSdk_UpdateTaskWithBodyParameters
#PsPlannerGraphSdk_DeleteTask
#PsPlannerGraphSdk_GetPlannerForUser
#PsPlannerGraphSdk_GetPlansForUser
#PsPlannerGraphSdk_GetTasksForUser

Write-Host "Done" 

