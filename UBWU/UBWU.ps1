
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


Function Get-AzureTokenApplication
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret,
 
		[Parameter(Mandatory=$False)]
		[String]$TenantName
	)
   
	 $LoginUrl = "https://login.microsoftonline.com"
	 $ScopeUrl = "https://graph.microsoft.com/.default"
	 
	 $myBody  = @{ Scope = $ScopeUrl; `
					grant_type = "client_credentials"; `
					client_id = $ClientID; `
					client_secret = $ClientSecret }

	 $myOAuth = Invoke-RestMethod `
					-Method Post `
					-Uri $LoginUrl/$TenantName/oauth2/v2.0/token `
					-Body $myBody

	return $myOAuth
}

Function Get-AzureTokenDelegation
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

Function LoginPsCLI
{
	m365 login --authType password `
			   --appId $configFile.appsettings.ClientIdWithAccPw `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

Function LoginPsPnPPowerShell
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteBaseUrl `
					  -ClientId $configFile.appsettings.ClientIdWithAccPw `
					  -Credentials $myCredentials
}

Function LoginPsGraphSDKWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $UserName, $securePW

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -UserCredential $myCredentials 

	Connect-Graph -AccessToken $myToken.AccessToken
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


##==> Graph

#gavdcodebegin 001
Function PlannerPsGraph_GetAllPlansForOneGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$grpId = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $grpId + "/planner/plans"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$groupObject = ConvertFrom-Json –InputObject $myResult
	$groupObject.value.subject
}
#gavdcodeend 001 

#gavdcodebegin 002
Function PlannerPsGraph_CreateOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$grpId = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
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
Function PlannerPsGraph_GetOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$planObject = ConvertFrom-Json –InputObject $myResult
	$planObject.value.subject
}
#gavdcodeend 003 

#gavdcodebegin 004
Function PlannerPsGraph_UpdateOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$grpId = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$eTag = 'W/"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $configFile.appsettings.ClientIdWithAccPw `
										-TenantName $configFile.appsettings.TenantName `
										-UserName $configFile.appsettings.UserName `
										-UserPw $configFile.appsettings.UserPw
	
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
Function PlannerPsGraph_GetOnePlanDetails
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId + "/details"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$planObject = ConvertFrom-Json –InputObject $myResult
	$planObject.value.subject
}
#gavdcodeend 005 

#gavdcodebegin 006
Function PlannerPsGraph_UpdateOnePlanDetails
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$grpId = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$eTag = 'W/"JzEtUGxhbkRldGFpbHMgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId + "/details"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
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
Function PlannerPsGraph_GetAllBucketsInOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId + "/buckets"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$bucketsObject = ConvertFrom-Json –InputObject $myResult
	$bucketsObject.value.subject
}
#gavdcodeend 007 

#gavdcodebegin 008
Function PlannerPsGraph_GetOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$bucketId = "0e70kPhvY0ueVmVlf50hbZgAOOxj"
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$bucketsObject = ConvertFrom-Json –InputObject $myResult
	$bucketsObject.value.subject
}
#gavdcodeend 008 

#gavdcodebegin 009
Function PlannerPsGraph_CreateOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
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
Function PlannerPsGraph_UpdateOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$bucketId = "D3hsxxsWiUuk0As8CVdpk5gAAugr"
	$eTag = 'W/"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
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
Function PlannerPsGraph_DeleteOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$bucketId = "0e70kPhvY0ueVmVlf50hbZgAOOxj"
	$eTag = 'W/"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId

	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 011 

#gavdcodebegin 012
Function PlannerPsGraph_DeleteOnePlan
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$eTag = 'W/"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBAUCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId

	$myOAuth = Get-AzureTokenDelegation -ClientID $configFile.appsettings.ClientIdWithAccPw `
										-TenantName $configFile.appsettings.TenantName `
										-UserName $configFile.appsettings.UserName `
										-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 012 

#gavdcodebegin 013
Function PlannerPsGraph_GetAllTasksInOneBucket
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$bucketId = "D3hsxxsWiUuk0As8CVdpk5gAAugr"
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId + "/tasks"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$tasksObject = ConvertFrom-Json –InputObject $myResult
	$tasksObject.value.subject
}
#gavdcodeend 013 

#gavdcodebegin 014
Function PlannerPsGraph_GetOneTask
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$taskId = "dp3JAX9my0uDbpoV0_26gpgALTDm"
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$taskObject = ConvertFrom-Json –InputObject $myResult
	$taskObject.value.subject
}
#gavdcodeend 014 

#gavdcodebegin 015
Function PlannerPsGraph_CreateOneTask
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$planId = "FNSaSwSeOkWKEkJ-l50klpgAHmCj"
	$bucketId = "D3hsxxsWiUuk0As8CVdpk5gAAugr"
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
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
Function PlannerPsGraph_UpdateOneTask
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$taskId = "kqTGSJuN_kC-lUCGi7ruDJgAF46C"
	$eTag = 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $configFile.appsettings.ClientIdWithAccPw `
										-TenantName $configFile.appsettings.TenantName `
										-UserName $configFile.appsettings.UserName `
										-UserPw $configFile.appsettings.UserPw
	
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
Function PlannerPsGraph_DeleteOneTask
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$taskId = "dp3JAX9my0uDbpoV0_26gpgALTDm"
	$eTag = 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId

	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 017 

#gavdcodebegin 053
Function PlannerPsGraph_GetOneTaskDetails
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$taskId = "tJqIvX1FwE6ixDnExZnYCpgAHErb"
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId + "/details"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$planObject = ConvertFrom-Json –InputObject $myResult
	$planObject.value.subject
}
#gavdcodeend 053 

#gavdcodebegin 054
Function PlannerPsGraph_UpdateOneTaskDetails
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$taskId = "tJqIvX1FwE6ixDnExZnYCpgAHErb"
	$eTag = 'W/"JzEtVGFza0RldGFpbHMgQEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId + "/details"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
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
Function PlannerPsGraph_GetTasksOneUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/me/planner/tasks"
	
	$myOAuth = Get-AzureTokenDelegation `
									-ClientID $configFile.appsettings.ClientIdWithAccPw `
									-TenantName $configFile.appsettings.TenantName `
									-UserName $configFile.appsettings.UserName `
									-UserPw $configFile.appsettings.UserPw
	
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
Function PlannerPsCli_GetAllPlans
{
	LoginPsCLI
	
	m365 planner plan list --ownerGroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	Write-Host ("-------")
	m365 planner plan list --ownerGroupName "Chapter18"

	m365 logout
}
#gavdcodeend 018

#gavdcodebegin 019
Function PlannerPsCli_GetPlansByQuery
{
	LoginPsCLI
	
	m365 planner plan list --ownerGroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" `
						   --output json `
						   --query "[?title == 'Plan01']"

	m365 logout
}
#gavdcodeend 019

#gavdcodebegin 020
Function PlannerPsCli_GetOnePlan
{
	LoginPsCLI
	
	m365 planner plan get --id "QP5LN-ygA0GnuHHDLKjY-JgACTnA"

	m365 logout
}
#gavdcodeend 020

#gavdcodebegin 021
Function PlannerPsCli_CreateOnePlan
{
	LoginPsCLI
	
	m365 planner plan add --title "PlanCreatedWithCLI" `
						  --ownerGroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"

	m365 logout
}
#gavdcodeend 021

#gavdcodebegin 056
Function PlannerPsCli_UpdateOnePlan
{
	LoginPsCLI
	
	m365 planner plan set --id "whr8psBnkkuZ6QPl4EfBc5gABQpF" `
						  --newTitle "PlanUpdatedWithCLI" `
						  --shareWithUserNames "user@tenant.onmicrosoft.com" `
						  --category1 "My Category"

	m365 logout
}
#gavdcodeend 056

#gavdcodebegin 057
Function PlannerPsCli_DeletePlan
{
	LoginPsCLI
	
	m365 planner plan remove --id "whr8psBnkkuZ6QPl4EfBc5gABQpF"

	#m365 planner plan remove --title "PlanUpdatedWithCLI" `
	#						 --ownerGroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" `
	#						 --confirm

	m365 logout
}
#gavdcodeend 057

#gavdcodebegin 022
Function PlannerPsCli_GetAllBuckets
{
	LoginPsCLI
	
	m365 planner bucket list --planId "whr8psBnkkuZ6QPl4EfBc5gABQpF"
	Write-Host ("-------")
	m365 planner bucket list --planTitle "PlanUpdatedWithCLI" `
							 --ownerGroupName "Chapter18"

	m365 logout
}
#gavdcodeend 022

#gavdcodebegin 023
Function PlannerPsCli_GetBucketsByQuery
{
	LoginPsCLI
	
	m365 planner bucket list --planId "whr8psBnkkuZ6QPl4EfBc5gABQpF" `
						     --output json `
						     --query "[?name == 'To do']"

	m365 logout
}
#gavdcodeend 023

#gavdcodebegin 024
Function PlannerPsCli_CreateOneBucket
{
	LoginPsCLI
	
	m365 planner bucket add --name "BucketCreatedWithCLI" `
						    --planId "whr8psBnkkuZ6QPl4EfBc5gABQpF" `
							--orderHint " !"

	m365 logout
}
#gavdcodeend 024

#gavdcodebegin 058
Function PlannerPsCli_GetOneBucket
{
	LoginPsCLI
	
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
Function PlannerPsCli_UpdateOneBucket
{
	LoginPsCLI
	
	m365 planner bucket set --id "r5bnTWCfnUGasiYREdqWPpgAOrFy" `
							--newName "BucketUpdatedWithCLI"

	m365 logout
}
#gavdcodeend 059

#gavdcodebegin 060
Function PlannerPsCli_DeleteBucket
{
	LoginPsCLI
	
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
Function PlannerPsCli_GetAllTasks
{
	LoginPsCLI
	
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
Function PlannerPsCli_GetTasksByQuery
{
	LoginPsCLI
	
	m365 planner task list --planId "whr8psBnkkuZ6QPl4EfBc5gABQpF" `
						   --output json `
						   --query "[?title == 'Task01']"

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 061
Function PlannerPsCli_CreateOneTask
{
	LoginPsCLI
	
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
Function PlannerPsCli_GetOneTask
{
	LoginPsCLI
	
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
Function PlannerPsCli_UpdateOneTask
{
	LoginPsCLI
	
	m365 planner task set --id "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
						  --title "TaskUpdatedWithCLI" `
						  --percentComplete 100 `
						  --appliedCategories "category1,category2"


	m365 logout
}
#gavdcodeend 063

#gavdcodebegin 064
Function PlannerPsCli_GetAllCheckListInTask
{
	LoginPsCLI
	
	m365 planner task checklistitem list --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY"

	m365 logout
}
#gavdcodeend 064

#gavdcodebegin 065
Function PlannerPsCli_AddOneCheckListToTask
{
	LoginPsCLI
	
	m365 planner task checklistitem add --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
										--title "Checklist Item CLI" `
										--isChecked

	m365 logout
}
#gavdcodeend 065

#gavdcodebegin 066
Function PlannerPsCli_DeleteOneCheckListFromTask
{
	LoginPsCLI
	
	m365 planner task checklistitem remove --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
										   --id "8fe00b42-68b5-459e-9a71-8ff56bc96bee" `
										   --confirm

	m365 logout
}
#gavdcodeend 066

#gavdcodebegin 067
Function PlannerPsCli_GetAllAttachmentsInTask
{
	LoginPsCLI
	
	m365 planner task reference list --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY"

	m365 logout
}
#gavdcodeend 067

#gavdcodebegin 068
Function PlannerPsCli_AddOneAttachmentToTask
{
	LoginPsCLI
	
	m365 planner task reference add --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
									--url "https://guitaca.com" `
									--type "Other" `
									--alias "Guitaca Publishers"

	m365 logout
}
#gavdcodeend 068

#gavdcodebegin 069
Function PlannerPsCli_DeleteOneAttachmentFromTask
{
	LoginPsCLI
	
	m365 planner task reference remove --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
									   --alias "Guitaca Publishers" `
									   --confirm
	
	#m365 planner task reference remove --taskId "9lrWLlboRkyy3A3cBdQ-9ZgADGQY" `
	#								   --url "https://guitaca.com" 

	m365 logout
}
#gavdcodeend 069

#gavdcodebegin 070
Function PlannerPsCli_DeleteOneTask
{
	LoginPsCLI
	
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
Function PlannerPsPnP_GetPlansByGroup
{
	# App Registration permissions: Group.Read.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerPlan -Group "Chapter18"
}
#gavdcodeend 027

#gavdcodebegin 028
Function PlannerPsPnP_GetPlansByGroupAndPlan
{
	# App Registration permissions: Group.Read.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerPlan -Group "Chapter18" -Identity "Plan01"
}
#gavdcodeend 028

#gavdcodebegin 029
Function PlannerPsPnP_CreatePlan
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	New-PnPPlannerPlan -Group "Chapter18" -Title "PlanCreatedWithPnP"
}
#gavdcodeend 029

#gavdcodebegin 030
Function PlannerPsPnP_UpdatePlanByPlan
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerPlan -Group "Chapter18" `
					   -Plan "PlanCreatedWithPnP" `
					   -Title "PlanUpdatedWithPnP"
}
#gavdcodeend 030

#gavdcodebegin 031
Function PlannerPsPnP_UpdatePlanById
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerPlan -PlanId "O3tqA06E_kaVBs0lVOAgbZgAAsQq" `
					   -Title "PlanUpdatedWithPnPById"
}
#gavdcodeend 031

#gavdcodebegin 032
Function PlannerPsPnP_DeletePlan
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Remove-PnPPlannerPlan -Group "Chapter18" `
						  -Identity "O3tqA06E_kaVBs0lVOAgbZgAAsQq"
}
#gavdcodeend 032

#gavdcodebegin 033
Function PlannerPsPnP_GetBucketsByGroupAndPlan
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerBucket -Group "Chapter18" -Plan "AnotherPlanCreatedWithPnP"
}
#gavdcodeend 033

#gavdcodebegin 034
Function PlannerPsPnP_GetBucketsById
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerBucket -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF"
}
#gavdcodeend 034

#gavdcodebegin 035
Function PlannerPsPnP_CreateBucketById
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Add-PnPPlannerBucket -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF" `
						 -Name "BucketCreatedWithPnP"
}
#gavdcodeend 035

#gavdcodebegin 036
Function PlannerPsPnP_CreateBucketByGroupAndPlan
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Add-PnPPlannerBucket -Group "Chapter18" `
						 -Plan "AnotherPlanCreatedWithPnP" `
						 -Name "BucketCreatedWithPnPByGroupAndPlan"
}
#gavdcodeend 036

#gavdcodebegin 037
Function PlannerPsPnP_UpdateBucketById
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerBucket -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF" `
						 -Bucket "BucketCreatedWithPnP" `
						 -Name "BucketUpdatedWithPnP"
}
#gavdcodeend 037

#gavdcodebegin 038
Function PlannerPsPnP_UpdateBucketByGroupAndPlan
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerBucket -Group "Chapter18" `
						 -Plan "AnotherPlanCreatedWithPnP" `
						 -Bucket "BucketCreatedWithPnPByGroupAndPlan" `
						 -Name "BucketUpdatedWithPnPByGroupAndPlan"
}
#gavdcodeend 038

#039 removed from the source code (update 2023-03)

#gavdcodebegin 040
Function PlannerPsPnP_DeleteBucketByGroupAndPlan
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Remove-PnPPlannerBucket -Group "Chapter18" `
							-Plan "AnotherPlanCreatedWithPnP" `
							-Identity "BucketUpdatedWithPnPByGroupAndPlan"
}
#gavdcodeend 040

#gavdcodebegin 041
Function PlannerPsPnP_GetTasksByPlanId
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerTask -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF"
}
#gavdcodeend 041

#gavdcodebegin 042
Function PlannerPsPnP_GetTasksByBucketId
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerTask -Bucket "BsZfkoF2iEmWB23IdPDxEJgAKDhP"
}
#gavdcodeend 042

#gavdcodebegin 043
Function PlannerPsPnP_GetTasksByTaskId
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerTask -TaskId "tVlliyvnL0GqamF0wkzGeZgAD7vF" `
					   -IncludeDetails `
					   -ResolveUserDisplayNames
}
#gavdcodeend 043

#gavdcodebegin 044
Function PlannerPsPnP_GetTasksByGroupAndPlan
{
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerTask -Group "Chapter18" `
					   -Plan "AnotherPlanCreatedWithPnP"
}
#gavdcodeend 044

#gavdcodebegin 045
Function PlannerPsPnP_CreateTaskByGroupAndPlan
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Add-PnPPlannerTask -Group "Chapter18" `
					   -Plan "AnotherPlanCreatedWithPnP" `
					   -Bucket "BucketUpdatedWithPnP" `
					   -Title "TaskCreatedWithPnP" `
					   -AssignedTo "user@domain.onmicrosoft.com"
}
#gavdcodeend 045

#gavdcodebegin 046
Function PlannerPsPnP_CreateTaskById
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Add-PnPPlannerTask -PlanId "lgyRpYpGEUmmXr0fzohpCZgAB_GF" `
					   -Bucket "BucketUpdatedWithPnP" `
					   -Title "TaskCreatedWithPnPById" `
}
#gavdcodeend 046

#gavdcodebegin 047
Function PlannerPsPnP_UpdateTaskById
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerTask -TaskId "tVlliyvnL0GqamF0wkzGeZgAD7vF" `
					   -Title "TaskUpdatedWithPnP" `
					   -AssignedTo "user1@dom.onmicrosoft.com","user2@dom.onmicrosoft.com"
}
#gavdcodeend 047

#gavdcodebegin 048
Function PlannerPsPnP_DeleteTaskById
{
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Remove-PnPPlannerTask -Task "tVlliyvnL0GqamF0wkzGeZgAD7vF"
}
#gavdcodeend 048

#gavdcodebegin 049
Function PlannerPsPnP_GetUserPolicy
{
	LoginPsPnPPowerShell
	
	Get-PnPPlannerUserPolicy -Identity "user@domain.onmicrosoft.com"
}
#gavdcodeend 049

#gavdcodebegin 050
Function PlannerPsPnP_SetUserPolicy
{
	LoginPsPnPPowerShell
	
	Set-PnPPlannerUserPolicy -Identity "user@domain.onmicrosoft.com" `
							 -BlockDeleteTasksNotCreatedBySelf $true
}
#gavdcodeend 050

#gavdcodebegin 051
Function PlannerPsPnP_GetConfiguration
{
	LoginPsPnPPowerShell
	
	Get-PnPPlannerConfiguration
}
#gavdcodeend 051

#gavdcodebegin 052
Function PlannerPsPnP_SetConfiguration
{
	LoginPsPnPPowerShell
	
	Set-PnPPlannerConfiguration -AllowCalendarSharing $false
}
#gavdcodeend 052

#-----------------------------------------------------------------------------------------

##==> Graph SDK

#gavdcodebegin 071
Function PlannerPsGraphSdk_GetAllPlansInGroup
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgGroupPlannerPlan -GroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"

	Disconnect-MgGraph
}
#gavdcodeend 071

#gavdcodebegin 072
Function PlannerPsGraphSdk_GetOnePlan
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgGroupPlannerPlan -GroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" `
						   -PlannerPlanId "QP5LN-ygA0GnuHHDLKjY-JgACTnA"

	Get-MgPlannerPlan -PlannerPlanId "QP5LN-ygA0GnuHHDLKjY-JgACTnA"

	Disconnect-MgGraph
}
#gavdcodeend 072

#gavdcodebegin 073
Function PlannerPsGraphSdk_CreatePlanWithBodyParameters
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	$PlanParameters = @{
		Owner = "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
		Title = "PlanCreatedWithGraphSDKBodyParams"
	}
	New-MgPlannerPlan -BodyParameter $PlanParameters

	Disconnect-MgGraph
}
#gavdcodeend 073

#gavdcodebegin 074
Function PlannerPsGraphSdk_CreatePlanWithParameters
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	New-MgPlannerPlan -Owner "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" `
					  -Title "PlanCreatedWithGraphSDKParams"

	Disconnect-MgGraph
}
#gavdcodeend 074

#gavdcodebegin 075
Function PlannerPsGraphSdk_UpdatePlanWithBodyParameters
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	$PlanParameters = @{
		Title = "PlanUpdatedWithGraphSDKParams"
	}
	Update-MgPlannerPlan -PlannerPlanId "465_rAMzX0eeXcvMIR8945gADbIY" `
						 -BodyParameter $PlanParameters

	Disconnect-MgGraph
}
#gavdcodeend 075

#gavdcodebegin 076
Function PlannerPsGraphSdk_DeletePlan
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	$planId = "465_rAMzX0eeXcvMIR8945gADbIY"
	$myPlan = Get-MgPlannerPlan -PlannerPlanId $planId
	$myEtag = $myPlan.AdditionalProperties.'@odata.etag'
	Remove-MgPlannerPlan -PlannerPlanId $planId `
						 -IfMatch $myEtag

	Disconnect-MgGraph
}
#gavdcodeend 076

#gavdcodebegin 077
Function PlannerPsGraphSdk_GetAllBucketsInPlan
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgPlannerPlanBucket -PlannerPlanId "kCJFZHc0jEqn5wBDnPRQv5gAH1_n"

	Disconnect-MgGraph
}
#gavdcodeend 077

#gavdcodebegin 078
Function PlannerPsGraphSdk_GetOneBucket
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgPlannerBucket -PlannerBucketId "sJQ8XSFfRUqvL6k_0OG1xpgAGUgE"

	Disconnect-MgGraph
}
#gavdcodeend 078

#gavdcodebegin 079
Function PlannerPsGraphSdk_CreateBucketWithBodyParameters
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

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
Function PlannerPsGraphSdk_CreateBucketWithParameters
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	New-MgPlannerBucket -PlanId "kCJFZHc0jEqn5wBDnPRQv5gAH1_n" `
						-Name "BucketCreatedWithGraphSDKParamsAA"

	Disconnect-MgGraph
}
#gavdcodeend 080

#gavdcodebegin 081
Function PlannerPsGraphSdk_UpdateBucketWithBodyParameters
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	$bucketParameters = @{
		Name = "BucketUpdatedWithGraphSDKBodyParams"
	}
	Update-MgPlannerBucket -PlannerBucketId "RqSFxZYBF06Y8TslksbV9ZgAP9-6" `
						   -BodyParameter $bucketParameters

	Disconnect-MgGraph
}
#gavdcodeend 081

#gavdcodebegin 082
Function PlannerPsGraphSdk_DeleteBucket
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	$bucketId = "RqSFxZYBF06Y8TslksbV9ZgAP9-6"
	$myBucket = Get-MgPlannerBucket -PlannerBucketId $bucketId
	$myEtag = $myBucket.AdditionalProperties.'@odata.etag'
	Remove-MgPlannerBucket -PlannerBucketId $bucketId `
						   -IfMatch $myEtag

	Disconnect-MgGraph
}
#gavdcodeend 082

#gavdcodebegin 083
Function PlannerPsGraphSdk_GetAllTasksInPlan
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgPlannerPlanTask -PlannerPlanId "kCJFZHc0jEqn5wBDnPRQv5gAH1_n"

	Disconnect-MgGraph
}
#gavdcodeend 083

#gavdcodebegin 084
Function PlannerPsGraphSdk_GetAllTasksInBucket
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgPlannerBucketTask -PlannerBucketId "1JSb34_9e0y71BgN-Z7jeJgAGdI5"

	Disconnect-MgGraph
}
#gavdcodeend 084

#gavdcodebegin 085
Function PlannerPsGraphSdk_GetOneTask
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgPlannerTask -PlannerTaskId "h-WwrIBusEW906Q9asal8ZgALG8c"

	Disconnect-MgGraph
}
#gavdcodeend 085

#gavdcodebegin 086
Function PlannerPsGraphSdk_GetOneTaskDetail
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgPlannerTaskDetail -PlannerTaskId "h-WwrIBusEW906Q9asal8ZgALG8c"

	Disconnect-MgGraph
}
#gavdcodeend 086

#gavdcodebegin 087
Function PlannerPsGraphSdk_CreateTaskWithBodyParameters
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

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
Function PlannerPsGraphSdk_CreateTaskWithParameters
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	New-MgPlannerTask -PlanId "kCJFZHc0jEqn5wBDnPRQv5gAH1_n" `
					  -BucketId "1JSb34_9e0y71BgN-Z7jeJgAGdI5" `
					  -Title "BucketCreatedWithGraphSDKParams"

	Disconnect-MgGraph
}
#gavdcodeend 088

#gavdcodebegin 089
Function PlannerPsGraphSdk_UpdateTaskWithBodyParameters
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

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
Function PlannerPsGraphSdk_DeleteTask
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	$taskId = "cs8qYcpUSEKaIjuzePRtQJgAKt0x"
	$myTask = Get-MgPlannerTask -PlannerTaskId $TaskId
	$myEtag = $myTask.AdditionalProperties.'@odata.etag'
	Remove-MgPlannerTask -PlannerTaskId $taskId `
						 -IfMatch $myEtag

	Disconnect-MgGraph
}
#gavdcodeend 090

#gavdcodebegin 091
Function PlannerPsGraphSdk_GetPlannerForUser
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgUserPlanner -UserId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	Get-MgUserPlanner -UserId "user@domain.onmicrosoft.com"

	Disconnect-MgGraph
}
#gavdcodeend 091

#gavdcodebegin 092
Function PlannerPsGraphSdk_GetPlansForUser
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgUserPlannerPlan -UserId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4"
	Get-MgUserPlannerPlan -UserId "user@domain.onmicrosoft.com"

	Disconnect-MgGraph
}
#gavdcodeend 092

#gavdcodebegin 093
Function PlannerPsGraphSdk_GetTasksForUser
{
	LoginPsGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							 -ClientID $configFile.appsettings.ClientIdWithAccPw `
							 -UserName $configFile.appsettings.UserName `
							 -UserPw $configFile.appsettings.UserPw

	Get-MgUserPlannerTask -UserId "acc28fcb-5261-47f8-960b-715d2f98a431"
	Get-MgUserPlannerTask -UserId "user@domain.onmicrosoft.com"

	Disconnect-MgGraph
}
#gavdcodeend 093


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

# *** Latest Source Code Index: 093 ***

#------------------------ Using Microsoft Graph PowerShell

#PlannerPsGraph_GetAllPlansForOneGroup
#PlannerPsGraph_CreateOnePlan
#PlannerPsGraph_GetOnePlan
#PlannerPsGraph_UpdateOnePlan
#PlannerPsGraph_GetOnePlanDetails
#PlannerPsGraph_UpdateOnePlanDetails
#PlannerPsGraph_GetAllBucketsInOnePlan
#PlannerPsGraph_GetOneBucket
#PlannerPsGraph_CreateOneBucket
#PlannerPsGraph_UpdateOneBucket
#PlannerPsGraph_DeleteOneBucket
#PlannerPsGraph_GetAllTasksInOneBucket
#PlannerPsGraph_GetOneTask
#PlannerPsGraph_GetOneTaskDetails
#PlannerPsGraph_UpdateOneTaskDetails
#PlannerPsGraph_GetTasksOneUser
#PlannerPsGraph_CreateOneTask
#PlannerPsGraph_UpdateOneTask
#PlannerPsGraph_DeleteOneTask
#PlannerPsGraph_DeleteOnePlan

#------------------------ Using PnP CLI

#PlannerPsCli_GetAllPlans
#PlannerPsCli_GetPlansByQuery
#PlannerPsCli_GetOnePlan
#PlannerPsCli_CreateOnePlan
#PlannerPsCli_UpdateOnePlan
#PlannerPsCli_DeletePlan
#PlannerPsCli_GetAllBuckets
#PlannerPsCli_GetBucketsByQuery
#PlannerPsCli_CreateOneBucket
#PlannerPsCli_GetOneBucket
#PlannerPsCli_UpdateOneBucket
#PlannerPsCli_DeleteBucket
#PlannerPsCli_GetAllTasks
#PlannerPsCli_GetTasksByQuery
#PlannerPsCli_CreateOneTask
#PlannerPsCli_GetOneTask
#PlannerPsCli_UpdateOneTask
#PlannerPsCli_GetAllCheckListInTask
#PlannerPsCli_AddOneCheckListToTask
#PlannerPsCli_DeleteOneCheckListFromTask
#PlannerPsCli_GetAllAttachmentsInTask
#PlannerPsCli_AddOneAttachmentToTask
#PlannerPsCli_DeleteOneAttachmentFromTask
#PlannerPsCli_DeleteOneTask

#------------------------ Using PowerShell PnP

#PlannerPsPnP_GetPlansByGroup
#PlannerPsPnP_GetPlansByGroupAndPlan
#PlannerPsPnP_CreatePlan
#PlannerPsPnP_UpdatePlanByPlan
#PlannerPsPnP_UpdatePlanById
#PlannerPsPnP_DeletePlan
#PlannerPsPnP_GetBucketsByGroupAndPlan
#PlannerPsPnP_GetBucketsById
#PlannerPsPnP_CreateBucketById
#PlannerPsPnP_CreateBucketByGroupAndPlan
#PlannerPsPnP_UpdateBucketById
#PlannerPsPnP_UpdateBucketByGroupAndPlan
#PlannerPsPnP_DeleteBucketByGroupAndPlan
#PlannerPsPnP_GetTasksByPlanId
#PlannerPsPnP_GetTasksByBucketId
#PlannerPsPnP_GetTasksByTaskId
#PlannerPsPnP_GetTasksByGroupAndPlan
#PlannerPsPnP_CreateTaskByGroupAndPlan
#PlannerPsPnP_CreateTaskById
#PlannerPsPnP_UpdateTaskById
#PlannerPsPnP_DeleteTaskById
#PlannerPsPnP_GetUserPolicy
#PlannerPsPnP_SetUserPolicy
#PlannerPsPnP_GetConfiguration
#PlannerPsPnP_SetConfiguration
#PlannerPsPnP_CreatePlannerRoster

#------------------------ Using Microsoft Graph PowerShell SDK

#PlannerPsGraphSdk_GetAllPlansInGroup
#PlannerPsGraphSdk_GetOnePlan
#PlannerPsGraphSdk_CreatePlanWithBodyParameters
#PlannerPsGraphSdk_CreatePlanWithParameters
#PlannerPsGraphSdk_UpdatePlanWithBodyParameters
#PlannerPsGraphSdk_DeletePlan
#PlannerPsGraphSdk_GetAllBucketsInPlan
#PlannerPsGraphSdk_GetOneBucket
#PlannerPsGraphSdk_CreateBucketWithBodyParameters
#PlannerPsGraphSdk_CreateBucketWithParameters
#PlannerPsGraphSdk_UpdateBucketWithBodyParameters
#PlannerPsGraphSdk_DeleteBucket
#PlannerPsGraphSdk_GetAllTasksInPlan
#PlannerPsGraphSdk_GetAllTasksInBucket
#PlannerPsGraphSdk_GetOneTask
#PlannerPsGraphSdk_GetOneTaskDetail
#PlannerPsGraphSdk_CreateTaskWithBodyParameters
#PlannerPsGraphSdk_CreateTaskWithParameters
#PlannerPsGraphSdk_UpdateTaskWithBodyParameters
#PlannerPsGraphSdk_DeleteTask
#PlannerPsGraphSdk_GetPlannerForUser
#PlannerPsGraphSdk_GetPlansForUser
#PlannerPsGraphSdk_GetTasksForUser

Write-Host "Done" 

