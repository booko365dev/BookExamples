
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
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


##==> Graph

#gavdcodebegin 001
Function ToDoPsGraph_GetAllListsMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$listObject = ConvertFrom-Json –InputObject $myResult
	$listObject.value.subject
}
#gavdcodeend 001 

#gavdcodebegin 002
Function ToDoPsGraph_GetAllListsUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
										$configFile.appsettings.UserName + "/todo/lists"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$listObject = ConvertFrom-Json –InputObject $myResult
	$listObject.value.subject
}
#gavdcodeend 002 

#gavdcodebegin 003
Function ToDoPsGraph_CreateOneListMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName':'ListCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 003 

#gavdcodebegin 004
Function ToDoPsGraph_GetOneListMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$listObject = ConvertFrom-Json –InputObject $myResult
	$listObject.value.subject
}
#gavdcodeend 004 

#gavdcodebegin 005
Function ToDoPsGraph_UpdateOneListMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName':'ListUpdatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 005 

#gavdcodebegin 006
Function ToDoPsGraph_DeleteOneListMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 006 

#gavdcodebegin 007
Function ToDoPsGraph_GetAllTasksInOneListMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + "/tasks"
	
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
#gavdcodeend 007 

#gavdcodebegin 008
Function ToDoPsGraph_GetOneTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw0JdAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
																	"/tasks/" + $taskId
	
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
#gavdcodeend 008 

#gavdcodebegin 009
Function ToDoPsGraph_CreateOneTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + "/tasks"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'title':'TaskCreatedWithGraph', `
				 'categories': ['Important'], `
				 'status': 'inProgress', `
				 'body': {
					'content':'This is the body', `
					'contentType':'text' }}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 009 

#gavdcodebegin 010
Function ToDoPsGraph_UpdateOneTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw0JeAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
																	"/tasks/" + $taskId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'title':'TaskUpdatedWithGraph', `
				 'body': {
					'content':'This is the body updated', `
					'contentType':'text' }}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 010 

#gavdcodebegin 011
Function ToDoPsGraph_DeleteOneTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOnAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw0JeAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
																	"/tasks/" + $taskId

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 011 

#gavdcodebegin 012
Function ToDoPsGraph_CreateOneLinkedResourceMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOoAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + "/tasks"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'title':'LinkedResourceCreatedWithGraph', `
				 'linkedResources': [{
					'webUrl':'https://guitaca.com', `
					'applicationName':'Guitaca', `
					'displayName':'Guitaca Publishers site', `
				    'externalId': 'myExternalId' }]}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 012 

#gavdcodebegin 013
Function ToDoPsGraph_GetAllLinkedResourcesInOneTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOoAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOoAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw0pIAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
																	"/tasks/" + $taskId
	
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
Function ToDoPsGraph_GetOneLinkedResourceInOneTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOoAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOoAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw0pIAAA="
	$linkedResourceId = "844b2936-0c58-416f-b42c-9aba1bd3ab0a"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
							"/tasks/" + $taskId + "/linkedResources/" + $linkedResourceId
	
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
Function ToDoPsGraph_UpdateOneLinkedResourceMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOoAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOoAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw0pIAAA="
	$linkedResourceId = "844b2936-0c58-416f-b42c-9aba1bd3ab0a"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
							"/tasks/" + $taskId + "/linkedResources/" + $linkedResourceId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName':'Guitaca Publishers site Updated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 015 

#gavdcodebegin 016
Function ToDoPsGraph_DeleteOneLinkedResourceMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOoAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOoAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw0pIAAA="
	$linkedResourceId = "844b2936-0c58-416f-b42c-9aba1bd3ab0a"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
							"/tasks/" + $taskId + "/linkedResources/" + $linkedResourceId

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 016 

#gavdcodebegin 017
Function ToDoPsGraph_CreateOneListWithExtensionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName':'ListExtensionCreatedWithGraph',
				  'extensions': [{
					  '@odata.type':'microsoft.graph.openTypeExtension',
					  'extensionName':'Com.Guitaca.MessageList',
					  'companyName':'Guitaca Publishers',
					  'expirationDate':'2055-12-30T01:00:00.000Z',
					  'myValue':123 }]}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 017 

#gavdcodebegin 018
Function ToDoPsGraph_GetOneListExtensionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$extensionName = "Com.Guitaca.MessageList"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
														"/extensions/" + $extensionName
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$listObject = ConvertFrom-Json –InputObject $myResult
	$listObject.value.subject
}
#gavdcodeend 018 

#gavdcodebegin 019
Function ToDoPsGraph_CreateOneTaskWithExtensionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + "/tasks"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'title':'TaskExtensionCreatedWithGraph', 
				 'body': {
					'content':'This is the body', 
					'contentType':'text' }, 
				 'extensions': [{ 
					'@odata.type':'microsoft.graph.openTypeExtension',
					'extensionName':'Com.Guitaca.MessageTask',
					'companyName':'Guitaca Publishers',
					'expirationDate':'2035-12-30T01:00:00.000Z',
					'myValue':456 }]}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 019 

#gavdcodebegin 020
Function ToDoPsGraph_GetOneTaskExtensionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IyAAA%3D"
	$extensionName = "Com.Guitaca.MessageTask"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
									"/tasks/" + $taskId + "/extensions/" + $extensionName
	
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
#gavdcodeend 020 

#gavdcodebegin 031
Function ToDoPsGraph_GetAllAttachmentsInTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
									"/tasks/" + $taskId + "/attachments"
	
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
#gavdcodeend 031

#gavdcodebegin 032
Function ToDoPsGraph_GetOneAttachmentInTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$attachmentId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAABEgAQANItUM24OuhJqMFkZ" + `
							"BSxfy8="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
									"/tasks/" + $taskId + "/attachments/" + $attachmentId
	
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
#gavdcodeend 032

#gavdcodebegin 033
Function ToDoPsGraph_UploadAttachmentSmallToTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$fileInfo = New-Object System.IO.FileInfo("C:\Temporary\TestDocument01.docx")
	$fileName = $fileInfo.Name
	$fileContentB64 = [System.IO.File]::ReadAllBytes($fileInfo.FullName)
	$fileContent = [System.Convert]::ToBase64String($fileContentB64)

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
									"/tasks/" + $taskId + "/attachments"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw

	$myBody = "{ '@odata.type':'#microsoft.graph.taskFileAttachment', 
				 'name':'" + $fileName + "', 
				 'contentBytes':'" + $fileContent + "' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 033

#gavdcodebegin 034
Function ToDoPsGraph_UploadAttachmentLargeToTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$fileInfo = New-Object System.IO.FileInfo("C:\Temporary\TestDocument01.docx")
	$fileName = $fileInfo.Name
	$fileSize = $fileInfo.Length
	$fileContentB64 = [System.IO.File]::ReadAllBytes($fileInfo.FullName)
	$fileContent = [System.Convert]::ToBase64String($fileContentB64)

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
								"/tasks/" + $taskId + "/attachments/createUploadSession"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw

	# Get an upload session
	$UrlSess = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
									"/tasks/" + $taskId + "/attachments"
	$myBodySess = "{ 'attachmentInfo': { 
				     'attachmentType':'file', 
				     'name':'" + $fileName + "',
				     'size':" + $fileSize + " }}"
	$myContentTypeSess = "application/json"
	$myHeaderSess = @{'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"}
	
	$myResultSess = Invoke-WebRequest -Headers $myHeaderSess -Uri $UrlSess -Method Post `
									  -Body $myBodySess -ContentType $myContentTypeSess

	Write-Host $myResultSess

	# Make a loop to upload each chunk of range
	$UrlUpl = $myResultSess.uploadUrl
	$myBodyUpl = "{ '" + $fileContent + "' }}"
	$myContentTypeUpl = "application/octet-stream"
	$myHeaderUpl = @{'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"
				  'Content-Length' = $($fileInfo.Length)
				  'Content-Range' = "bytes 0-$($fileInfo.Length)/$($fileInfo.Length)" }
	
	$myResultUpl = Invoke-WebRequest -Headers $myHeaderUpl -Uri $UrlUpl -Method Put `
									-Body $myBodyUpl -ContentType $myContentTypeUpl

	Write-Host $myResultUpl
}
#gavdcodeend 034

#gavdcodebegin 035
Function ToDoPsGraph_DeleteOneSessionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$sessionId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAABEgAQANItUM24OuhJqMFkZ" + `
							"BSxfy8="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
							"/tasks/" + $taskId + "/attachmentSessions/" + $sessionId

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 035

#gavdcodebegin 036
Function ToDoPsGraph_DeleteOneAttachmentsFromTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$attachmentId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAABEgAQANItUM24OuhJqMFkZ" + `
							"BSxfy8="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
									"/tasks/" + $taskId + "/attachments/" + $attachmentId

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 036

#gavdcodebegin 037
Function ToDoPsGraph_GetAllStepsInTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
									"/tasks/" + $taskId + "/checklistItems"
	
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
#gavdcodeend 037

#gavdcodebegin 038
Function ToDoPsGraph_GetOneStepInTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$stepId = "5a6d2d5c-35a8-43e7-a188-82863b7c1c03"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
								"/tasks/" + $taskId + "/checklistItems/" + $stepId
	
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
#gavdcodeend 038

#gavdcodebegin 039
Function ToDoPsGraph_CreateStepInTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
									"/tasks/" + $taskId + "/checklistItems"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw

	$myBody = "{ 'displayName':'Step created with Graph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 039

#gavdcodebegin 040
Function ToDoPsGraph_UpdateOneStepInTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$stepId = "f6a2337f-96ef-4cb1-b6e6-5c97d57bc0b7"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
								"/tasks/" + $taskId + "/checklistItems/" + $stepId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName':'Step updated with Graph',
				 'isChecked':'True' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 040

#gavdcodebegin 041
Function ToDoPsGraph_DeleteOneStepFromTaskMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAAA="
	$taskId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
							"cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAACKwqOpAA" + `
							"C1vtBLB-F9SJ2ZDb7Xo-OrAACKw1IzAAA="
	$stepId = "f6a2337f-96ef-4cb1-b6e6-5c97d57bc0b7"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
								"/tasks/" + $taskId + "/checklistItems/" + $stepId

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 041

#-----------------------------------------------------------------------------------------

##==> CLI

#gavdcodebegin 021
Function ToDoPsCli_GetAllLists
{
	LoginPsCLI
	
	m365 todo list list

	m365 logout
}
#gavdcodeend 021

#gavdcodebegin 022
Function ToDoPsCli_GetListsByQuery
{
	LoginPsCLI
	
	m365 todo list list --output json `
						--query "[?displayName == 'My ToDo']"

	m365 logout
}
#gavdcodeend 022

#gavdcodebegin 023
Function ToDoPsCli_AddOneList
{
	LoginPsCLI
	
	m365 todo list add --name "ToDoCreatedWithCLI"

	m365 logout
}
#gavdcodeend 023

#gavdcodebegin 024
Function ToDoPsCli_UpdateOneList
{
	LoginPsCLI
	
	m365 todo list set --name "ToDoCreatedWithCLI" `
					   --newName "ToDoUpdatedWithCLI"

	m365 logout
}
#gavdcodeend 024

#gavdcodebegin 025
Function ToDoPsCli_DeleteOneList
{
	LoginPsCLI
	
	m365 todo list remove --name "ToDoUpdatedWithCLI" `
						  --confirm

	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
Function ToDoPsCli_GetAllTasks
{
	LoginPsCLI
	
	m365 todo task list --listName "ToDoCreatedWithCLI"

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
Function ToDoPsCli_GetTasksByQuery
{
	LoginPsCLI
	
	m365 todo task list --listName "ToDoCreatedWithCLI" `
						--output json `
						--query "[?title == 'Task number one']"

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
Function ToDoPsCli_AddOneTask
{
	LoginPsCLI
	
	m365 todo task add --listName "ToDoCreatedWithCLI" `
					   --title "ToDoTaskCreatedWithCLI"

	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 029
Function ToDoPsCli_UpdateOneTask
{
	LoginPsCLI
	
	m365 todo task set --listName "ToDoCreatedWithCLI" `
					   --id "AAMkAGE0ODQ3NTc1LT ... 1vtBLB-F9SJ2ZDb7Xo-OrAACMOMRsAAA=" `
					   --title "ToDoTaskUpdatedWithCLI" `
					   --status "deferred"

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
Function ToDoPsCli_DeleteOneTask
{
	LoginPsCLI
	
	m365 todo task remove --listName "ToDoCreatedWithCLI" `
						  --id "AAMkAGE0ODQ3NTc1LT ... 1vtBLB-F9SJ2ZDb7Xo-OrAACMRsAAA=" `
						  --confirm

	m365 logout
}
#gavdcodeend 030


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

# *** Latest Source Code Index: 41 ***

#------------------------ Using Microsoft Graph PowerShell

#ToDoPsGraph_GetAllListsMe
#ToDoPsGraph_GetAllListsUser
#ToDoPsGraph_CreateOneListMe
#ToDoPsGraph_GetOneListMe
#ToDoPsGraph_UpdateOneListMe
#ToDoPsGraph_DeleteOneListMe
#ToDoPsGraph_GetAllTasksInOneListMe
#ToDoPsGraph_GetOneTaskMe
#ToDoPsGraph_CreateOneTaskMe
#ToDoPsGraph_UpdateOneTaskMe
#ToDoPsGraph_DeleteOneTaskMe
#ToDoPsGraph_CreateOneLinkedResourceMe
#ToDoPsGraph_GetAllLinkedResourcesInOneTaskMe
#ToDoPsGraph_GetOneLinkedResourceInOneTaskMe
#ToDoPsGraph_UpdateOneLinkedResourceMe
#ToDoPsGraph_DeleteOneLinkedResourceMe
#ToDoPsGraph_CreateOneListWithExtensionMe
#ToDoPsGraph_GetOneListExtensionMe
#ToDoPsGraph_CreateOneTaskWithExtensionMe
#ToDoPsGraph_GetOneTaskExtensionMe
#ToDoPsGraph_GetAllAttachmentsInTaskMe
#ToDoPsGraph_GetOneAttachmentInTaskMe
#ToDoPsGraph_UploadAttachmentSmallToTaskMe
#ToDoPsGraph_UploadAttachmentLargeToTaskMe
#ToDoPsGraph_DeleteOneSessionMe
#ToDoPsGraph_DeleteOneAttachmentsFromTaskMe
#ToDoPsGraph_GetAllStepsInTaskMe
#ToDoPsGraph_GetOneStepInTaskMe
#ToDoPsGraph_CreateStepInTaskMe
#ToDoPsGraph_UpdateOneStepInTaskMe
#ToDoPsGraph_DeleteOneStepFromTaskMe

#------------------------ Using PnP CLI

#ToDoPsCli_GetAllLists
#ToDoPsCli_GetListsByQuery
#ToDoPsCli_AddOneList
#ToDoPsCli_UpdateOneList
#ToDoPsCli_DeleteOneList
#ToDoPsCli_GetAllTasks
#ToDoPsCli_GetTasksByQuery
#ToDoPsCli_AddOneTask
#ToDoPsCli_UpdateOneTask
#ToDoPsCli_DeleteOneTask

Write-Host "Done" 