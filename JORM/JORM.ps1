
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

#gavdcodebegin 01
Function TodoPsGraphGetAllListsMe()
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
Function TodoPsGraphGetAllListsUser()
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
Function TodoPsGraphCreateOneListMe()
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
Function TodoPsGraphGetOneListMe()
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
Function TodoPsGraphUpdateOneListMe()
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
Function TodoPsGraphDeleteOneListMe()
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
Function TodoPsGraphGetAllTasksInOneListMe()
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
Function TodoPsGraphGetOneTaskMe()
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
Function TodoPsGraphCreateOneTaskMe()
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
Function TodoPsGraphUpdateOneTaskMe()
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
Function TodoPsGraphDeleteOneTaskMe()
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
Function TodoPsGraphCreateOneLinkedResourceMe()
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
Function TodoPsGraphGetAllLinkedResourcesInOneTaskMe()
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
Function TodoPsGraphGetOneLinkedResourceInOneTaskMe()
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
Function TodoPsGraphUpdateOneLinkedResourceMe()
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
Function TodoPsGraphDeleteOneLinkedResourceMe()
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
Function TodoPsGraphCreateOneListWithExtensionMe()
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
Function TodoPsGraphGetOneListExtensionMe()
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
Function TodoPsGraphCreateOneTaskWithExtensionMe()
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
Function TodoPsGraphGetOneTaskExtensionMe()
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

#gavdcodebegin 21
Function TodoPsCliGetAllLists(){
	LoginPsCLI
	
	m365 todo list list

	m365 logout
}
#gavdcodeend 21

#gavdcodebegin 22
Function TodoPsCliGetListsByQuery(){
	LoginPsCLI
	
	m365 todo list list --output json `
						--query "[?displayName == 'My ToDo']"

	m365 logout
}
#gavdcodeend 22

#gavdcodebegin 23
Function TodoPsCliAddOneList(){
	LoginPsCLI
	
	m365 todo list add --name "ToDoCreatedWithCLI"

	m365 logout
}
#gavdcodeend 23

#gavdcodebegin 24
Function TodoPsCliUpdateOneList(){
	LoginPsCLI
	
	m365 todo list set --name "ToDoCreatedWithCLI" `
					   --newName "ToDoUpdatedWithCLI"

	m365 logout
}
#gavdcodeend 24

#gavdcodebegin 25
Function TodoPsCliDeleteOneList(){
	LoginPsCLI
	
	m365 todo list remove --name "ToDoUpdatedWithCLI" `
						  --confirm

	m365 logout
}
#gavdcodeend 25

#gavdcodebegin 26
Function TodoPsCliGetAllTasks(){
	LoginPsCLI
	
	m365 todo task list --listName "ToDoCreatedWithCLI"

	m365 logout
}
#gavdcodeend 26

#gavdcodebegin 27
Function TodoPsCliGetTasksByQuery(){
	LoginPsCLI
	
	m365 todo task list --listName "ToDoCreatedWithCLI" `
						--output json `
						--query "[?title == 'Task number one']"

	m365 logout
}
#gavdcodeend 27

#gavdcodebegin 28
Function TodoPsCliAddOneTask(){
	LoginPsCLI
	
	m365 todo task add --listName "ToDoCreatedWithCLI" `
					   --title "ToDoTaskCreatedWithCLI"

	m365 logout
}
#gavdcodeend 28

#gavdcodebegin 29
Function TodoPsCliUpdateOneTask(){
	LoginPsCLI
	
	m365 todo task set --listName "ToDoCreatedWithCLI" `
					   --id "AAMkAGRiNjdkMjA5LTNkOWMtNDkxMS...A1yrhXAAA=" `
					   --title "ToDoTaskUpdatedWithCLI" `
					   --status "deferred"

	m365 logout
}
#gavdcodeend 29

#gavdcodebegin 30
Function TodoPsCliDeleteOneTask(){
	LoginPsCLI
	
	m365 todo task remove --listName "ToDoCreatedWithCLI" `
						  --id "AAMkAGRiNjdkMjA5LTNkOWMtNDkxMS...A1yrhXAAA=" `
						  --confirm

	m365 logout
}
#gavdcodeend 30

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

#TodoPsGraphGetAllListsMe
#TodoPsGraphGetAllListsUser
#TodoPsGraphCreateOneListMe
#TodoPsGraphGetOneListMe
#TodoPsGraphUpdateOneListMe
#TodoPsGraphDeleteOneListMe
#TodoPsGraphGetAllTasksInOneListMe
#TodoPsGraphGetOneTaskMe
#TodoPsGraphCreateOneTaskMe
#TodoPsGraphUpdateOneTaskMe
#TodoPsGraphDeleteOneTaskMe
#TodoPsGraphCreateOneLinkedResourceMe
#TodoPsGraphGetAllLinkedResourcesInOneTaskMe
#TodoPsGraphGetOneLinkedResourceInOneTaskMe
#TodoPsGraphUpdateOneLinkedResourceMe
#TodoPsGraphDeleteOneLinkedResourceMe
#TodoPsGraphCreateOneListWithExtensionMe
#TodoPsGraphGetOneListExtensionMe
#TodoPsGraphCreateOneTaskWithExtensionMe
#TodoPsGraphGetOneTaskExtensionMe

#------------------------ Using Microsoft PnP CLI for Teams

#TodoPsCliGetAllLists
#TodoPsCliGetListsByQuery
#TodoPsCliAddOneList
#TodoPsCliUpdateOneList
#TodoPsCliDeleteOneList
#TodoPsCliGetAllTasks
#TodoPsCliGetTasksByQuery
#TodoPsCliAddOneTask
#TodoPsCliUpdateOneTask
#TodoPsCliDeleteOneTask

Write-Host "Done" 