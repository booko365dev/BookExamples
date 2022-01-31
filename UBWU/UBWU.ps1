
# Functions to login in Azure

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

Function LoginPsCLI()
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

#----------------------------------------------------------------------------------------

Function LoginPsPnPPowerShell()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.TenantUrl -Credentials $myCredentials
}

#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function PlannerPsGraphGetAllPlansForOneGroup()
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
Function PlannerPsGraphCreateOnePlan()
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
Function PlannerPsGraphGetOnePlan()
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
Function PlannerPsGraphUpdateOnePlan()
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
Function PlannerPsGraphGetOnePlanDetails()
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
Function PlannerPsGraphUpdateOnePlanDetails()
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
Function PlannerPsGraphGetAllBucketsInOnePlan()
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
Function PlannerPsGraphGetOneBucket()
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
Function PlannerPsGraphCreateOneBucket()
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
Function PlannerPsGraphUpdateOneBucket()
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
Function PlannerPsGraphDeleteOneBucket()
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
Function PlannerPsGraphDeleteOnePlan()
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
Function PlannerPsGraphGetAllTasksInOneBucket()
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
Function PlannerPsGraphGetOneTask()
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
Function PlannerPsGraphCreateOneTask()
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
Function PlannerPsGraphUpdateOneTask()
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
Function PlannerPsGraphDeleteOneTask()
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

#gavdcodebegin 18
Function PlannerPsCliGetAllPlans(){
	LoginPsCLI
	
	m365 planner plan list --ownerGroupId "ad111b29-e670-4ed5-9091-a3f38e96699c"
	Write-Host ("-------")
	m365 planner plan list --ownerGroupName "Sales and Marketing"

	m365 logout
}
#gavdcodeend 18

#gavdcodebegin 19
Function PlannerPsCliGetPlansByQuery(){
	LoginPsCLI
	
	m365 planner plan list --ownerGroupId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						   --output json `
						   --query "[?title == 'Product Launch Event']"

	m365 logout
}
#gavdcodeend 19

#gavdcodebegin 20
Function PlannerPsCliGetOnePlan(){
	LoginPsCLI
	
	m365 planner plan get --id "UkAjM0Z1kkCuBbMiW_faZWUAHAwS"

	m365 logout
}
#gavdcodeend 20

#gavdcodebegin 21
Function PlannerPsCliCreateOnePlan(){
	LoginPsCLI
	
	m365 planner plan add --title "PlanCreatedWithCLI" `
						  --ownerGroupId "7bfa5f9c-6bd5-453b-a028-d8d57dfb48ba"

	m365 logout
}
#gavdcodeend 21

#gavdcodebegin 22
Function PlannerPsCliGetAllBuckets(){
	LoginPsCLI
	
	m365 planner bucket list --planId "17lVrTdRhU26Jzdylapl0mUAEoMx"
	Write-Host ("-------")
	m365 planner bucket list --planName "PlanCreatedWithCLI" `
							 --ownerGroupName "GroupForPlanner"

	m365 logout
}
#gavdcodeend 22

#gavdcodebegin 23
Function PlannerPsCliGetBucketsByQuery(){
	LoginPsCLI
	
	m365 planner bucket list --planId "17lVrTdRhU26Jzdylapl0mUAEoMx" `
						     --output json `
						     --query "[?name == 'To do']"

	m365 logout
}
#gavdcodeend 23

#gavdcodebegin 24
Function PlannerPsCliCreateOneBucket(){
	LoginPsCLI
	
	m365 planner bucket add --name "BucketCreatedWithCLI" `
						    --planId "17lVrTdRhU26Jzdylapl0mUAEoMx" `
							--orderHint " !"

	m365 logout
}
#gavdcodeend 24

#gavdcodebegin 25
Function PlannerPsCliGetAllTasks(){
	LoginPsCLI
	
	m365 planner task list --planId "17lVrTdRhU26Jzdylapl0mUAEoMx" `
						   --bucketId "WGqJZ3Q-zEeI4-oFTKeMAWUACUt1"

	m365 logout
}
#gavdcodeend 25

#gavdcodebegin 26
Function PlannerPsCliGetTasksByQuery(){
	LoginPsCLI
	
	m365 planner task list --planId "17lVrTdRhU26Jzdylapl0mUAEoMx" `
						   --output json `
						   --query "[?title == 'Task added manually']"

	m365 logout
}
#gavdcodeend 26

#gavdcodebegin 27
Function PlannerPsPnPGetPlansByGroup(){
	# App Registration permissions: Group.Read.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerPlan -Group "MyPlan"
}
#gavdcodeend 27

#gavdcodebegin 28
Function PlannerPsPnPGetPlansByGroupAndPlan(){
	# App Registration permissions: Group.Read.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerPlan -Group "MyPlan" -Identity "MyOtherPlan"
}
#gavdcodeend 28

#gavdcodebegin 29
Function PlannerPsPnPCreatePlan(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	New-PnPPlannerPlan -Group "MyPlan" -Title "PlanCreatedWithPnP"
}
#gavdcodeend 29

#gavdcodebegin 30
Function PlannerPsPnPUpdatePlanByPlan(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerPlan -Group "MyPlan" `
					   -Plan "PlanCreatedWithPnP" `
					   -Title "PlanUpdatedWithPnP"
}
#gavdcodeend 30

#gavdcodebegin 31
Function PlannerPsPnPUpdatePlanById(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerPlan -PlanId "aX4Jj_e8LU-oEh2bmfyvTJgAGgYr" `
					   -Title "PlanUpdatedWithPnPById"
}
#gavdcodeend 31

#gavdcodebegin 32
Function PlannerPsPnPDeletePlan(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Remove-PnPPlannerPlan -Group "MyPlan" `
						  -Identity "aX4Jj_e8LU-oEh2bmfyvTJgAGgYr"
}
#gavdcodeend 32

#gavdcodebegin 33
Function PlannerPsPnPGetBucketsByGroupAndPlan(){
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerBucket -Group "MyPlan" -Plan "MyOtherPlan"
}
#gavdcodeend 33

#gavdcodebegin 34
Function PlannerPsPnPGetBucketsById(){
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerBucket -PlanId "D8_FXAjY-kK5pfc4bk6ImZgAEYJ1"
}
#gavdcodeend 34

#gavdcodebegin 35
Function PlannerPsPnPCreateBucketById(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Add-PnPPlannerBucket -PlanId "D8_FXAjY-kK5pfc4bk6ImZgAEYJ1" `
						 -Name "BucketCreatedWithPnP"
}
#gavdcodeend 35

#gavdcodebegin 36
Function PlannerPsPnPCreateBucketByGroupAndPlan(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Add-PnPPlannerBucket -Group "MyPlan" `
						 -Plan "MyOtherPlan" `
						 -Name "BucketCreatedWithPnPByGroupAndPlan"
}
#gavdcodeend 36

#gavdcodebegin 37
Function PlannerPsPnPUpdateBucketById(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerBucket -PlanId "D8_FXAjY-kK5pfc4bk6ImZgAEYJ1" `
						 -Bucket "BucketCreatedWithPnP" `
						 -Name "BucketUpdatedWithPnP"
}
#gavdcodeend 37

#gavdcodebegin 38
Function PlannerPsPnPUpdateBucketByGroupAndPlan(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerBucket -Group "MyPlan" `
						 -Plan "MyOtherPlan" `
						 -Bucket "BucketCreatedWithPnPByGroupAndPlan" `
						 -Name "BucketUpdatedWithPnPByGroupAndPlan"
}
#gavdcodeend 38

#gavdcodebegin 39
Function PlannerPsPnPDeleteBucketById(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	# Note: It doesn't work --> Issue raised to Microsoft
	#Remove-PnPPlannerBucket -BucketId "D8_FXAjY-kK5pfc4bk6ImZgAEYJ1" `
	#						-Identity "BucketUpdatedWithPnPByGroupAndPlan"
}
#gavdcodeend 39

#gavdcodebegin 40
Function PlannerPsPnPDeleteBucketByGroupAndPlan(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Remove-PnPPlannerBucket -Group "MyPlan" `
							-Plan "MyOtherPlan" `
							-Identity "BucketUpdatedWithPnPByGroupAndPlan"
}
#gavdcodeend 40

#gavdcodebegin 41
Function PlannerPsPnPGetTasksByPlanId(){
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerTask -PlanId "D8_FXAjY-kK5pfc4bk6ImZgAEYJ1"
}
#gavdcodeend 41

#gavdcodebegin 42
Function PlannerPsPnPGetTasksByBucketId(){
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerTask -Bucket "JsiP6IO10UGm3yzUsjmiJ5gAGKcq"
}
#gavdcodeend 42

#gavdcodebegin 43
Function PlannerPsPnPGetTasksByTaskId(){
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerTask -TaskId "pzWcc9A4xECghOOhBQFh5ZgAM5Ni" `
					   -IncludeDetails `
					   -ResolveUserDisplayNames
}
#gavdcodeend 43

#gavdcodebegin 44
Function PlannerPsPnPGetTasksByGroupAndPlan(){
	# App Registration permissions: Group.Read.All or Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Get-PnPPlannerTask -Group "MyPlan" `
					   -Plan "MyOtherPlan"
}
#gavdcodeend 44

#gavdcodebegin 45
Function PlannerPsPnPCreateTaskByGroupAndPlan(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Add-PnPPlannerTask -Group "MyPlan" `
					   -Plan "MyOtherPlan" `
					   -Bucket "MyBucket" `
					   -Title "TaskCreatedWithPnP" `
					   -AssignedTo "user@tenant.onmicrosoft.com"
}
#gavdcodeend 45

#gavdcodebegin 46
Function PlannerPsPnPCreateTaskById(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Add-PnPPlannerTask -PlanId "D8_FXAjY-kK5pfc4bk6ImZgAEYJ1" `
					   -Bucket "MyBucket" `
					   -Title "TaskCreatedWithPnPById" `
}
#gavdcodeend 46

#gavdcodebegin 47
Function PlannerPsPnPUpdateTaskById(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Set-PnPPlannerTask -TaskId "pzWcc9A4xECghOOhBQFh5ZgAM5Ni" `
					   -Title "TaskUpdatedWithPnP" `
					   -AssignedTo "user1@dom.onmicrosoft.com","user2@dom.onmicrosoft.com"
}
#gavdcodeend 47

#gavdcodebegin 48
Function PlannerPsPnPDeleteTaskById(){
	# App Registration permissions: Group.ReadWrite.All

	LoginPsPnPPowerShell
	
	Remove-PnPPlannerTask -Task "pzWcc9A4xECghOOhBQFh5ZgAM5Ni"
}
#gavdcodeend 48

#gavdcodebegin 49
Function PlannerPsPnPGetUserPolicy(){
	LoginPsPnPPowerShell
	
	Get-PnPPlannerUserPolicy -Identity "user@domain.onmicrosoft.com"
}
#gavdcodeend 49

#gavdcodebegin 50
Function PlannerPsPnPSetUserPolicy(){
	LoginPsPnPPowerShell
	
	Set-PnPPlannerUserPolicy -Identity "user@domain.onmicrosoft.com" `
							 -BlockDeleteTasksNotCreatedBySelf $true
}
#gavdcodeend 50

#gavdcodebegin 51
Function PlannerPsPnPGetConfiguration(){
	LoginPsPnPPowerShell
	
	Get-PnPPlannerConfiguration
}
#gavdcodeend 51

#gavdcodebegin 52
Function PlannerPsPnPSetConfiguration(){
	LoginPsPnPPowerShell
	
	Set-PnPPlannerConfiguration -AllowCalendarSharing $false
}
#gavdcodeend 52

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\grPs.values.config"

#------------------------ Using Microsoft Graph PowerShell for Teams

#$ClientIDApp = $configFile.appsettings.ClientIdApp
#$ClientSecretApp = $configFile.appsettings.ClientSecretApp
#$ClientIDDel = $configFile.appsettings.ClientIdDel
#$TenantName = $configFile.appsettings.TenantName
#$UserName = $configFile.appsettings.UserName
#$UserPw = $configFile.appsettings.UserPw

#PlannerPsGraphGetAllPlansForOneGroup
#PlannerPsGraphCreateOnePlan
#PlannerPsGraphGetOnePlan
#PlannerPsGraphUpdateOnePlan
#PlannerPsGraphDeleteOnePlan
#PlannerPsGraphGetOnePlanDetails
#PlannerPsGraphUpdateOnePlanDetails
#PlannerPsGraphGetAllBucketsInOnePlan
#PlannerPsGraphGetOneBucket
#PlannerPsGraphCreateOneBucket
#PlannerPsGraphUpdateOneBucket
#PlannerPsGraphDeleteOneBucket
#PlannerPsGraphGetAllTasksInOneBucket
#PlannerPsGraphGetOneTask
#PlannerPsGraphCreateOneTask
#PlannerPsGraphUpdateOneTask
#PlannerPsGraphDeleteOneTask

#------------------------ Using PnP CLI for Teams

#PlannerPsCliGetAllPlans
#PlannerPsCliGetPlansByQuery
#PlannerPsCliGetOnePlan
#PlannerPsCliCreateOnePlan
#PlannerPsCliGetAllBuckets
#PlannerPsCliGetBucketsByQuery
#PlannerPsCliCreateOneBucket
#PlannerPsCliGetAllTasks
#PlannerPsCliGetTasksByQuery

#------------------------ Using PowerShell PnP for Teams

#PlannerPsPnPGetPlansByGroup
#PlannerPsPnPGetPlansByGroupAndPlan
#PlannerPsPnPCreatePlan
#PlannerPsPnPUpdatePlanByPlan
#PlannerPsPnPUpdatePlanById
#PlannerPsPnPDeletePlan
#PlannerPsPnPGetBucketsByGroupAndPlan
#PlannerPsPnPGetBucketsById
#PlannerPsPnPCreateBucketById
#PlannerPsPnPCreateBucketByGroupAndPlan
#PlannerPsPnPUpdateBucketById
#PlannerPsPnPUpdateBucketByGroupAndPlan
#PlannerPsPnPDeleteBucketById (It doesn't work --> Issue raised to Microsoft)
#PlannerPsPnPDeleteBucketByGroupAndPlan
#PlannerPsPnPGetTasksByPlanId
#PlannerPsPnPGetTasksByBucketId
#PlannerPsPnPGetTasksByTaskId
#PlannerPsPnPGetTasksByGroupAndPlan
#PlannerPsPnPCreateTaskByGroupAndPlan
#PlannerPsPnPCreateTaskById
#PlannerPsPnPUpdateTaskById
#PlannerPsPnPDeleteTaskById
#PlannerPsPnPGetUserPolicy
#PlannerPsPnPSetUserPolicy
#PlannerPsPnPGetConfiguration
#PlannerPsPnPSetConfiguration

Write-Host "Done" 

