
Function Get-AzureTokenApplication(){
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

Function Get-AzureTokenDelegation(){
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

#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function GrPsGetAllPlansForOneGroup()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$grpId = "6f45c2f8-9b25-47a6-a2d4-9323cbb094d7"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $grpId + "/planner/plans"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$groupObject = ConvertFrom-Json –InputObject $myResult
	$groupObject.value.subject
}
#gavdcodeend 01 

#gavdcodebegin 02
Function GrPsCreateOnePlan()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$grpId = "6f45c2f8-9b25-47a6-a2d4-9323cbb094d7"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'owner':'" + $grpId + "', `
			     'title':'GraphPlan' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 02 

#gavdcodebegin 03
Function GrPsGetOnePlan()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$planId = "JuRApYInUEOXTyhAOo0iCpgACiCh"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$planObject = ConvertFrom-Json –InputObject $myResult
	$planObject.value.subject
}
#gavdcodeend 03 

#gavdcodebegin 04
Function GrPsUpdateOnePlan()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$grpId = "6f45c2f8-9b25-47a6-a2d4-9323cbb094d7"
	$planId = "JuRApYInUEOXTyhAOo0iCpgACiCh"
	$eTag = 'W/"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBATCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'title':'GraphPlanUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 04 

#gavdcodebegin 05
Function GrPsGetOnePlanDetails()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$planId = "JuRApYInUEOXTyhAOo0iCpgACiCh"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId + "/details"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$planObject = ConvertFrom-Json –InputObject $myResult
	$planObject.value.subject
}
#gavdcodeend 05 

#gavdcodebegin 06
Function GrPsUpdateOnePlanDetails()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$grpId = "6f45c2f8-9b25-47a6-a2d4-9323cbb094d7"
	$planId = "JuRApYInUEOXTyhAOo0iCpgACiCh"
	$eTag = 'W/"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBATCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId + "/details"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'sharedWith': { '2288575c-d043-4bfd-acf4-fddd769846e5':true, `
								 'c651f419-5dbb-44f1-9e2f-cb39c6300211':true}, `
				 'categoryDescriptions': { 'category1': 'myLabel', `
										   'category2': null } }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 06 

#gavdcodebegin 07
Function GrPsGetAllBucketsInOnePlan()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$planId = "JuRApYInUEOXTyhAOo0iCpgACiCh"
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId + "/buckets"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$bucketsObject = ConvertFrom-Json –InputObject $myResult
	$bucketsObject.value.subject
}
#gavdcodeend 07 

#gavdcodebegin 08
Function GrPsGetOneBucket()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$bucketId = "Ujuq39ZqMEic1oxIf3iRBpgAKDfs"
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$bucketsObject = ConvertFrom-Json –InputObject $myResult
	$bucketsObject.value.subject
}
#gavdcodeend 08 

#gavdcodebegin 09
Function GrPsCreateOneBucket()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$planId = "JuRApYInUEOXTyhAOo0iCpgACiCh"
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'planId':'" + $planId + "', `
			     'name':'BucketOne', ` 
				 'orderHint': ' !' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 09 

#gavdcodebegin 10
Function GrPsUpdateOneBucket()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$bucketId = "xrcV4uOhtU2eFAihFm8vIJgAJ_yH"
	$eTag = 'W/"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $bucketId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'name':'BucketOneUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 10 

#gavdcodebegin 11
Function GrPsDeleteOneBucket()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$bucketId = "xrcV4uOhtU2eFAihFm8vIJgAJ_yH"
	$eTag = 'W/"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 11 

#gavdcodebegin 12
Function GrPsDeleteOnePlan()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$planId = "JuRApYInUEOXTyhAOo0iCpgACiCh"
	$eTag = 'W/"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBATCc="'
	$Url = "https://graph.microsoft.com/v1.0/planner/plans/" + $planId

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 12 

#gavdcodebegin 13
Function GrPsGetAllTasksInOneBucket()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$bucketId = "xrcV4uOhtU2eFAihFm8vIJgAJ_yH"
	$Url = "https://graph.microsoft.com/v1.0/planner/buckets/" + $bucketId + "/tasks"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$tasksObject = ConvertFrom-Json –InputObject $myResult
	$tasksObject.value.subject
}
#gavdcodeend 13 

#gavdcodebegin 14
Function GrPsGetOneTask()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$taskId = "G3jS0IU1b0Wh9WVR66N9NpgAOXfj"
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$taskObject = ConvertFrom-Json –InputObject $myResult
	$taskObject.value.subject
}
#gavdcodeend 14 

#gavdcodebegin 15
Function GrPsCreateOneTask()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$planId = "JuRApYInUEOXTyhAOo0iCpgACiCh"
	$bucketId = "xrcV4uOhtU2eFAihFm8vIJgAJ_yH"
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'planId':'" + $planId + "', `
			     'bucketId':'" + $bucketId + "', ` 
			     'title':'TaskOne', `
				 'assignments': {
					'2288575c-d043-4bfd-acf4-fddd769846e5': {
					'@odata.type': '#microsoft.graph.plannerAssignment',
					'orderHint': ' !' }}}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 15 

#gavdcodebegin 16
Function GrPsUpdateOneTask()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$taskId = "Ed3Iti8rY0WqR0xfpKebIZgAJZNU"
	$eTag = 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc"'
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'percentComplete':1 }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 16 

#gavdcodebegin 17
Function GrPsDeleteOneTask()
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$taskId = "Ed3Iti8rY0WqR0xfpKebIZgAJZNU"
	$eTag = 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc"'
	$Url = "https://graph.microsoft.com/v1.0/planner/tasks/" + $taskId

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 17 

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\grPs.values.config"

$ClientIDApp = $configFile.appsettings.ClientIdApp
$ClientSecretApp = $configFile.appsettings.ClientSecretApp
$ClientIDDel = $configFile.appsettings.ClientIdDel
$TenantName = $configFile.appsettings.TenantName
$UserName = $configFile.appsettings.UserName
$UserPw = $configFile.appsettings.UserPw

#GrPsGetAllPlansForOneGroup
#GrPsCreateOnePlan
#GrPsGetOnePlan
#GrPsUpdateOnePlan
#GrPsDeleteOnePlan
#GrPsGetOnePlanDetails
#GrPsUpdateOnePlanDetails
#GrPsGetAllBucketsInOnePlan
#GrPsGetOneBucket
#GrPsCreateOneBucket
#GrPsUpdateOneBucket
#GrPsDeleteOneBucket
#GrPsGetAllTasksInOneBucket
#GrPsGetOneTask
#GrPsCreateOneTask
#GrPsUpdateOneTask
#GrPsDeleteOneTask

Write-Host "Done" 

