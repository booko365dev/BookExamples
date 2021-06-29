
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
Function TdPsGetAllListsMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$listObject = ConvertFrom-Json –InputObject $myResult
	$listObject.value.subject
}
#gavdcodeend 01 

#gavdcodebegin 02
Function TdPsGetAllListsUser()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + $UserName + "/todo/lists"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$listObject = ConvertFrom-Json –InputObject $myResult
	$listObject.value.subject
}
#gavdcodeend 02 

#gavdcodebegin 03
Function TdPsCreateOneListMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'displayName':'ListFromPowerShell' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 03 

#gavdcodebegin 04
Function TdPsGetOneListMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5KAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$listObject = ConvertFrom-Json –InputObject $myResult
	$listObject.value.subject
}
#gavdcodeend 04 

#gavdcodebegin 05
Function TdPsUpdateOneListMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5KAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'displayName':'ListFromPowerShellUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 05 

#gavdcodebegin 06
Function TdPsDeleteOneListMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5KAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 06 

#gavdcodebegin 07
Function TdPsGetAllTasksInOneListMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + "/tasks"
	
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
#gavdcodeend 07 

#gavdcodebegin 08
Function TdPsGetOneTaskMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$taskId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgBGAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLBwD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAD-ny0YZ9qbRaGi" + `
				"xdDNLZGSAABEgEEZAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
																	"/tasks/" + $taskId
	
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
#gavdcodeend 08 

#gavdcodebegin 09
Function TdPsCreateOneTaskMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + "/tasks"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'title':'TaskOne', `
				 'body': {
					'content':'This is the body', `
					'contentType':'text' }}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 09 

#gavdcodebegin 10
Function TdPsUpdateOneTaskMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$taskId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgBGAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLBwD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAD-ny0YZ9qbRaGi" + `
				"xdDNLZGSAABHPwyKAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
																	"/tasks/" + $taskId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'title':'TaskOneUpdated', `
				 'body': {
					'content':'This is the body updated', `
					'contentType':'text' }}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 10 

#gavdcodebegin 11
Function TdPsDeleteOneTaskMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$taskId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgBGAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLBwD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAD-ny0YZ9qbRaGi" + `
				"xdDNLZGSAABHPwyKAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
																	"/tasks/" + $taskId

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 11 

#gavdcodebegin 12
Function TdPsCreateOneLinkedResourceMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + "/tasks"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'title':'Task With linkedResource', `
				 'linkedResources': [{
					'webUrl':'https://guitaca.com', `
					'applicationName':'Guitaca', `
					'displayName':'Blog in Guitaca Publishers', `
				    'externalId': 'myExternalId' }]}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 12 

#gavdcodebegin 13
Function TdPsGetAllLinkedResourcesInOneTaskMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$taskId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgBGAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLBwD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAD-ny0YZ9qbRaGi" + `
				"xdDNLZGSAABHox4hAAA%3D"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
																	"/tasks/" + $taskId
	
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
Function TdPsGetOneLinkedResourceInOneTaskMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$taskId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgBGAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLBwD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAD-ny0YZ9qbRaGi" + `
				"xdDNLZGSAABHox4hAAA%3D"
	$linkedResourceId = "880db57c-d2c9-4fcd-a9bf-f6171ea9a390"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
							"/tasks/" + $taskId + "/linkedResources/" + $linkedResourceId
	
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
Function TdPsUpdateOneLinkedResourceMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$taskId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgBGAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLBwD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAD-ny0YZ9qbRaGi" + `
				"xdDNLZGSAABHox4hAAA%3D"
	$linkedResourceId = "880db57c-d2c9-4fcd-a9bf-f6171ea9a390"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
							"/tasks/" + $taskId + "/linkedResources/" + $linkedResourceId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'displayName':'Blog in Guitaca Publishers updated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 15 

#gavdcodebegin 16
Function TdPsDeleteOneLinkedResourceMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$taskId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgBGAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLBwD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAD-ny0YZ9qbRaGi" + `
				"xdDNLZGSAABHox4hAAA%3D"
	$linkedResourceId = "880db57c-d2c9-4fcd-a9bf-f6171ea9a390"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
							"/tasks/" + $taskId + "/linkedResources/" + $linkedResourceId

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 16 

#gavdcodebegin 17
Function TdPsCreateOneListWithExtensionMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'displayName':'ListFromPowerShell With Extension',
				  'extensions': [{
					  '@odata.type':'microsoft.graph.openTypeExtension',
					  'extensionName':'Com.Guitaca.MessageList',
					  'companyName':'Guitaca Publishers',
					  'expirationDate':'2035-12-30T01:00:00.000Z',
					  'myValue':123 }]}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 17 

#gavdcodebegin 18
Function TdPsGetOneListExtensionMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABHo0nDAAA%3D"
	$extensionName = "Com.Guitaca.MessageList"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
														"/extensions/" + $extensionName
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$listObject = ConvertFrom-Json –InputObject $myResult
	$listObject.value.subject
}
#gavdcodeend 18 

#gavdcodebegin 19
Function TdPsCreateOneTaskWithExtensionMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + ` 
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + "/tasks"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'title':'Task With Extension', 
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
#gavdcodeend 19 

#gavdcodebegin 20
Function TdPsGetOneTaskExtensionMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.Read, Tasks.ReadWrite

	$listId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgAuAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLAQD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAA%3D"
	$taskId = "AAMkAGFjYmFkODk2LTE0ZTEtNGFhOC04YWEzLWVlOTJmN2U2MzM0NgBGAAAAAAAXIQ2" + `
				"MI48_TqhxzHlnbwlLBwD-ny0YZ9qbRaGixdDNLZGSAABEf-5JAAD-ny0YZ9qbRaGi" + `
				"xdDNLZGSAABHox4iAAA%3D"
	$extensionName = "Com.Guitaca.MessageTask"
	$Url = "https://graph.microsoft.com/v1.0/me/todo/lists/" + $listId + `
									"/tasks/" + $taskId + "/extensions/" + $extensionName
	
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
#gavdcodeend 20 

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\grPs.values.config"

$ClientIDApp = $configFile.appsettings.ClientIdApp
$ClientSecretApp = $configFile.appsettings.ClientSecretApp
$ClientIDDel = $configFile.appsettings.ClientIdDel
$TenantName = $configFile.appsettings.TenantName
$UserName = $configFile.appsettings.UserName
$UserPw = $configFile.appsettings.UserPw

#TdPsGetAllListsMe
#TdPsGetAllListsUser
#TdPsCreateOneListMe
#TdPsGetOneListMe
#TdPsUpdateOneListMe
#TdPsDeleteOneListMe
#TdPsGetAllTasksInOneListMe
#TdPsGetOneTaskMe
#TdPsCreateOneTaskMe
#TdPsUpdateOneTaskMe
#TdPsDeleteOneTaskMe
#TdPsCreateOneLinkedResourceMe
#TdPsGetAllLinkedResourcesInOneTaskMe
#TdPsGetOneLinkedResourceInOneTaskMe
#TdPsUpdateOneLinkedResourceMe
#TdPsDeleteOneLinkedResourceMe
#TdPsCreateOneListWithExtensionMe
#TdPsGetOneListExtensionMe
#TdPsCreateOneTaskWithExtensionMe
#TdPsGetOneTaskExtensionMe

Write-Host "Done" 

