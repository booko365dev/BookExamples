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
 
		[Parameter(Mandatory=$False)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret
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

Function LoginGraphSDKWithAccPw
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
Function OneNotePsGraph_GetAllNotebooksMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 001 

#gavdcodebegin 002
Function OneNotePsGraph_GetAllNotebooksByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/onenote/notebooks"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 002 

#gavdcodebegin 027
Function OneNotePsGraph_GetAllNotebooksByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$groupId = "CF7E4CB9-E929-43D9-84BA-BD7C123DAAE9"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + `
							$groupId + "/onenote/notebooks"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 027 

#gavdcodebegin 028
Function OneNotePsGraph_GetAllNotebooksBySite
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$siteId = "FCB9425A-E423-4988-8611-ACACEA52400B"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + `
							$siteId + "/onenote/notebooks"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 028 

#gavdcodebegin 003
Function OneNotePsGraph_CreateOneNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName':'NotebookCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 003 

#gavdcodebegin 004
Function OneNotePsGraph_GetOneNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 004 

#gavdcodebegin 005
Function OneNotePsGraph_GetRecentNotebooksMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + `
					"getRecentNotebooks(includePersonalNotebooks=true)"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 005 

#gavdcodebegin 006
Function OneNotePsGraph_CopyOneNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + `
				"/copyNotebook"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'renameAs':'CopyNotebookCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 006

Function OneNotePsGraph_UpdateOneNotebookMe # No methods in Graph API for OneNote
{
	# ATTENTION: There are no methods in the OneNote Graph API to modify one Book
	#	It needs to be done modifying the file in OneDrive

	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'renameAs':'NotebookUpdatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}

Function OneNotePsGraph_DeleteOneNotebookMe # # No methods in Graph API for OneNote
{
	# ATTENTION: There are no methods in the OneNote Graph API to modify one Book
	#	It needs to be done modifying the file in OneDrive

	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}

#gavdcodebegin 007
Function OneNotePsGraph_GetAllSectionsMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 007 

#gavdcodebegin 008
Function OneNotePsGraph_GetAllSectionsOneNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + "/sections"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 008

#gavdcodebegin 009
Function OneNotePsGraph_GetOneSectionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionId = "1-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 009

#gavdcodebegin 010
Function OneNotePsGraph_CreateOneSectionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + "/sections"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName':'SectionCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 010 

#gavdcodebegin 011
Function OneNotePsGraph_UpdateOneSectionMe # It doesn't work
{
	# ATTENTION: The routine gives no error, but the display name is not changed at all

	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite

	$sectionId = "1-47e53239-8c05-4e9b-9eed-7124fce1be0d"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'renameAs':'SectionUpdatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 011 

Function OneNotePsGraph_DeleteOneSectionMe # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$sectionId = "1-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}

#gavdcodebegin 012
Function OneNotePsGraph_CopyOneSectionToNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$sectionId = "1-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
				"/copyToNotebook"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'renameAs':'CopySectionFromCreatedWithGraph', `
			     'id':'$($bookId)' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 012

#gavdcodebegin 013
Function OneNotePsGraph_GetAllSectionGroupsMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 013 

#gavdcodebegin 014
Function OneNotePsGraph_GetOneSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionGroupId = "1-27b4559e-63a9-4beb-b6d7-772df43e5f19"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups/" + $sectionGroupId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 014

#gavdcodebegin 015
Function OneNotePsGraph_CreateOneSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + `
					"/sectionGroups"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName':'SectionGroupCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 015

#gavdcodebegin 016
Function OneNotePsGraph_GetAllSectionsOneSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionGroupId = "1-31330f7c-afe6-4554-aa7e-7c0a73a64c09"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups/" + `
						$sectionGroupId + "/sections"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 016

#gavdcodebegin 017
Function OneNotePsGraph_CreateOneSectionInOneSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$sectionGroupId = "1-31330f7c-afe6-4554-aa7e-7c0a73a64c09"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups/" + `
						$sectionGroupId + "/sections"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName':'SectionInSectionGroupCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 017

#gavdcodebegin 018
Function OneNotePsGraph_CopyOneSectionToSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$sectionGroupId = "1-31330f7c-afe6-4554-aa7e-7c0a73a64c09"
	$sectionId = "1-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
				"/copyToSectionGroup"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'renameAs':'CopySectionToSectionGroupWithGraph', `
				 'id':'$($sectionGroupId)' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 018

#gavdcodebegin 019
Function OneNotePsGraph_GetAllPagesInOneNoteMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 019

#gavdcodebegin 020
Function OneNotePsGraph_GetAllPagesInOneSectionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionId = "1-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
					"/pages"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 020

#gavdcodebegin 021
Function OneNotePsGraph_GetOnePageMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-8c31cf8ccc4a410ea79eb50c6f04f963!14-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 021

#gavdcodebegin 022
Function OneNotePsGraph_CreatePageInOneSectionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$sectionId = "1-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
					"/pages"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "<!DOCTYPE html>" + `
			  "	<html>" + `
			  "	  <head>" + `
			  "		<title>PageInSectionCreatedWithGraph</title>" + `
			  "	  </head>" + `
			  "	  <body>" + `
			  "		<p>Content of the page</p>" + `
			  "	  </body>" + `
			  "	</html>"
	
	$myContentType = "application/xhtml+xml"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 022

#gavdcodebegin 023
Function OneNotePsGraph_GetOnePageContentMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-367fa14618504079824c3320806db6a2!6-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId + "/content"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 023

#gavdcodebegin 024
Function OneNotePsGraph_GetOnePageContentWithIdsMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-367fa14618504079824c3320806db6a2!6-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId + `
					"/content?includeIDs=true"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 024

#gavdcodebegin 025
Function OneNotePsGraph_UpdateOnePageContentMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite

	$pageId = "1-367fa14618504079824c3320806db6a2!6-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$targetId = "p:{1d122d8f-b983-4903-8c1a-43e2f2a0cae5}{101}"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $spageId + "/content"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "[{ 'target':'$($targetId)', `
				  'action':'append', `
				  'position':'after', `
				  'content':' - <b>Updated with Graph</b>' }]"

	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 025 

#gavdcodebegin 026
Function OneNotePsGraph_DeleteOnePageMe # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-367fa14618504079824c3320806db6a2!6-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}
#gavdcodeend 026

#-----------------------------------------------------------------------------------------

##==> CLI

#gavdcodebegin 029
Function OneNotePsCli_GetAllNotebooks
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginPsCLI
	
	m365 onenote notebook list

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
Function OneNotePsCli_GetAllNotebooksByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginPsCLI
	
	m365 onenote notebook list --groupName "MyM365Group"
	m365 onenote notebook list --groupId "B5FD0AE8-0695-489E-B142-3A2C36AC43B2"

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
Function OneNotePsCli_GetAllNotebooksByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginPsCLI
	
	m365 onenote notebook list --userName $configFile.appsettings.UserName
	m365 onenote notebook list --userId "765764B2-3C30-40DA-B293-0CB28B8E8148"

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
Function OneNotePsCli_GetAllNotebooksBySite
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginPsCLI
	
	m365 onenote notebook list --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 032

#gavdcodebegin 033
Function OneNotePsCli_GetAllPagess
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginPsCLI
	
	m365 onenote page list

	m365 logout
}
#gavdcodeend 033

#gavdcodebegin 034
Function OneNotePsCli_GetAllPagesByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginPsCLI
	
	m365 onenote page list --groupName "MyM365Group"
	m365 onenote page list --groupId "B5FD0AE8-0695-489E-B142-3A2C36AC43B2"

	m365 logout
}
#gavdcodeend 034

#gavdcodebegin 035
Function OneNotePsCli_GetAllPagesByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginPsCLI
	
	m365 onenote page list --userName $configFile.appsettings.UserName
	m365 onenote page list --userId "765764B2-3C30-40DA-B293-0CB28B8E8148"

	m365 logout
}
#gavdcodeend 035

#gavdcodebegin 036
Function OneNotePsCli_GetAllPagesBySite
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginPsCLI
	
	m365 onenote page list --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 036

#-----------------------------------------------------------------------------------------

##==> Graph SDK

#gavdcodebegin 037
Function OneNotePsGraphSdk_GetAllNotebooksByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenoteNotebook -UserId $configFile.appsettings.UserName
	Get-MgUserOnenoteNotebook -UserId "B193A6C2-CBDD-43A8-B080-09BB93FDF8A1"

	Disconnect-MgGraph
}
#gavdcodeend 037

#gavdcodebegin 038
Function OneNotePsGraphSdk_GetOneNotebookByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenoteNotebook -UserId $configFile.appsettings.UserName `
							  -NotebookId "1-0a6ab636-237a-48c5-95c9-f59d042ff776"

	Disconnect-MgGraph
}
#gavdcodeend 038

#gavdcodebegin 039
Function OneNotePsGraphSdk_GetAllNotebooksByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgGroupOnenoteNotebook -GroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" 

	Disconnect-MgGraph
}
#gavdcodeend 039

#gavdcodebegin 040
Function OneNotePsGraphSdk_GetAllNotebooksBySite
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteOnenoteNotebook -SiteId "1ec84def-070a-4a94-a26d-2845b3bf6ff3"

	Disconnect-MgGraph
}
#gavdcodeend 040

#gavdcodebegin 041
Function OneNotePsGraphSdk_CreateOneNotebookByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	$NotebookParameters = @{
		DisplayName = "NotebookCreatedWithGraphSDKBodyParams"
	}

	New-MgUserOnenoteNotebook -UserId $configFile.appsettings.UserName `
							  -BodyParameter $NotebookParameters

	Disconnect-MgGraph
}
#gavdcodeend 041

#gavdcodebegin 042
Function OneNotePsGraphSdk_CreateOneNotebookByUserWithParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	New-MgUserOnenoteNotebook -UserId $configFile.appsettings.UserName `
							  -DisplayName "NotebookCreatedWithGraphSDKParams"

	Disconnect-MgGraph
}
#gavdcodeend 042

#gavdcodebegin 043
Function OneNotePsGraphSdk_UpdateOneNotebookByUserWithBodyParameters # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	$NotebookParameters = @{
		DisplayName = "NotebookUpdatedWithGraphSDKBodyParams"
	}

	Update-MgUserOnenoteNotebook -UserId $configFile.appsettings.UserName `
								 -NotebookId "1-18896ef3-0d27-47ad-a1da-271b03bc5eb4" `
								 -BodyParameter $NotebookParameters

	Disconnect-MgGraph
}
#gavdcodeend 043

#gavdcodebegin 044
Function OneNotePsGraphSdk_DeleteOneNotebookByUser # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Remove-MgUserOnenoteNotebook -UserId $configFile.appsettings.UserName `
								 -NotebookId "1-18896ef3-0d27-47ad-a1da-271b03bc5eb4"

	Disconnect-MgGraph
}
#gavdcodeend 044

#gavdcodebegin 045
Function OneNotePsGraphSdk_GetAllSectionsByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenoteSection -UserId $configFile.appsettings.UserName
	Get-MgUserOnenoteSection -UserId "B193A6C2-CBDD-43A8-B080-09BB93FDF8A1"

	Disconnect-MgGraph
}
#gavdcodeend 045

#gavdcodebegin 046
Function OneNotePsGraphSdk_GetAllSectionsOneNotebookByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenoteNotebookSection -UserId $configFile.appsettings.UserName `
									 -NotebookId "1-41dc1f07-d7c3-4913-96a5-7c27a973050a"

	Disconnect-MgGraph
}
#gavdcodeend 046

#gavdcodebegin 047
Function OneNotePsGraphSdk_CreateOneSectionInNotebookByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	$SectionParameters = @{
		DisplayName = "SectionCreatedWithGraphSDKBodyParams"
	}

	New-MgUserOnenoteNotebookSection -UserId $configFile.appsettings.UserName `
								-NotebookId "1-41dc1f07-d7c3-4913-96a5-7c27a973050a" `
								-BodyParameter $SectionParameters

	Disconnect-MgGraph
}
#gavdcodeend 047

#gavdcodebegin 048
Function OneNotePsGraphSdk_UpdateOneSectionByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	$SectionParameters = @{
		DisplayName = "SectionUpdatedWithGraphSDKBodyParams"
	}

	Update-MgUserOnenoteSection -UserId $configFile.appsettings.UserName `
							-OnenoteSectionId "1-f2474782-6ea0-41f4-a187-bc1d1715cc3f" `
							-BodyParameter $SectionParameters

	Disconnect-MgGraph
}
#gavdcodeend 048

#gavdcodebegin 049
Function OneNotePsGraphSdk_DeleteOneSectionByUser # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Remove-MgUserOnenoteSection -UserId $configFile.appsettings.UserName `
							-OnenoteSectionId "1-f2474782-6ea0-41f4-a187-bc1d1715cc3f"

	Disconnect-MgGraph
}
#gavdcodeend 049

#gavdcodebegin 050
Function OneNotePsGraphSdk_GetAllSectionGroupsByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenoteSectionGroup -UserId $configFile.appsettings.UserName
	Get-MgUserOnenoteSectionGroup -UserId "B193A6C2-CBDD-43A8-B080-09BB93FDF8A1"

	Disconnect-MgGraph
}
#gavdcodeend 050

#gavdcodebegin 051
Function OneNotePsGraphSdk_GetAllSectionGroupsOneNotebookByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenoteNotebookSectionGroup -UserId $configFile.appsettings.UserName `
									 -NotebookId "1-41dc1f07-d7c3-4913-96a5-7c27a973050a"

	Disconnect-MgGraph
}
#gavdcodeend 051

#gavdcodebegin 052
Function OneNotePsGraphSdk_CreateOneSectionGroupInNotebookByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	$SectionParameters = @{
		DisplayName = "SectionGroupCreatedWithGraphSDKBodyParams"
	}

	New-MgUserOnenoteNotebookSectionGroup -UserId $configFile.appsettings.UserName `
								-NotebookId "1-41dc1f07-d7c3-4913-96a5-7c27a973050a" `
								-BodyParameter $SectionParameters

	Disconnect-MgGraph
}
#gavdcodeend 052

#gavdcodebegin 053
Function OneNotePsGraphSdk_UpdateOneSectionGroupByUserWithBodyParameters # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	$SectionParameters = @{
		DisplayName = "SectionGroupUpdatedWithGraphSDKBodyParams"
	}

	Update-MgUserOnenoteSectionGroup -UserId $configFile.appsettings.UserName `
						-SectionGroupId "1-60a12e43-8562-4919-8e77-4f7b633012a2" `
						-BodyParameter $SectionParameters

	Disconnect-MgGraph
}
#gavdcodeend 053

#gavdcodebegin 054
Function OneNotePsGraphSdk_DeleteOneSectionGroupByUser # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Remove-MgUserOnenoteSectionGroup -UserId $configFile.appsettings.UserName `
							    -SectionGroupId "1-60a12e43-8562-4919-8e77-4f7b633012a2"

	Disconnect-MgGraph
}
#gavdcodeend 054

#gavdcodebegin 055
Function OneNotePsGraphSdk_GetAllPagesByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenotePage -UserId $configFile.appsettings.UserName
	Get-MgUserOnenotePage -UserId "B193A6C2-CBDD-43A8-B080-09BB93FDF8A1"

	Disconnect-MgGraph
}
#gavdcodeend 055

#gavdcodebegin 056
Function OneNotePsGraphSdk_GetAllPagesOneSectionByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenoteSectionPage -UserId $configFile.appsettings.UserName `
							-OnenoteSectionId "1-f2474782-6ea0-41f4-a187-bc1d1715cc3f"

	Disconnect-MgGraph
}
#gavdcodeend 056

#gavdcodebegin 057
Function OneNotePsGraphSdk_GetOnePageContentByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenotePageContent -UserId $configFile.appsettings.UserName `
								 -OnenotePageId "1-ffea745...1d1715cc3f" `
								 -OutFile "C:\Temporary\page.txt"

	Disconnect-MgGraph
}
#gavdcodeend 057

#gavdcodebegin 058
Function OneNotePsGraphSdk_GetOnePageResourceByUser #It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenoteResource -UserId $configFile.appsettings.UserName `
							  -OnenoteResourceId "1-0e975ac85...d1715cc3f"

	Disconnect-MgGraph
}
#gavdcodeend 058

#gavdcodebegin 059
Function OneNotePsGraphSdk_GetOnePageResourceConentByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Get-MgUserOnenoteResourceContent -UserId $configFile.appsettings.UserName `
									 -OnenoteResourceId "1-0e975ac8...715cc3f" `
									 -OutFile "C:\Temporary\pagecontent.txt"

	Disconnect-MgGraph
}
#gavdcodeend 059

#gavdcodebegin 060
Function OneNotePsGraphSdk_CreateOnePageInSectionByUserWithBodyParameters #It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	$PageParameters = @{
		Title = "PageCreatedWithGraphSDKBodyParams"
	}

	New-MgUserOnenoteSectionPage -UserId $configFile.appsettings.UserName `
							-OnenoteSectionId "1-f2474782-6ea0-41f4-a187-bc1d1715cc3f" `
							-BodyParameter $PageParameters

	Disconnect-MgGraph
}
#gavdcodeend 060

#gavdcodebegin 061
Function OneNotePsGraphSdk_CreateOnePageInSectionByUserWithParameters # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	New-MgUserOnenoteSectionPage -UserId $configFile.appsettings.UserName `
							-OnenoteSectionId "1-f2474782-6ea0-41f4-a187-bc1d1715cc3f" `
							-Title "PageCreatedWithGraphSDKBodyParams"

	Disconnect-MgGraph
}
#gavdcodeend 061

#gavdcodebegin 062
Function OneNotePsGraphSdk_UpdateOnePageByUserWithBodyParameters # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	$PageParameters = @{
		Title = "PageUpdatedWithGraphSDKBodyParams"
	}

	Update-MgUserOnenotePage -UserId $configFile.appsettings.UserName `
							 -OnenotePageId "1-ffea745f7d6e44...1715cc3f" `
							 -BodyParameter $PageParameters

	Disconnect-MgGraph
}
#gavdcodeend 062

#gavdcodebegin 063
Function OneNotePsGraphSdk_DeleteOnePageByUser # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	LoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
						   -ClientID $configFile.appsettings.ClientIdWithAccPw `
						   -UserName $configFile.appsettings.UserName `
						   -UserPw $configFile.appsettings.UserPw

	Remove-MgUserOnenotePage -UserId $configFile.appsettings.UserName `
							 -OnenotePageId "1-ffea745f7d6e44...715cc3f"

	Disconnect-MgGraph
}
#gavdcodeend 063


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

# *** Latest Source Code Index: 63 ***

#------------------------ Using Microsoft Graph PowerShell

#OneNotePsGraph_GetAllNotebooksMe
#OneNotePsGraph_GetAllNotebooksByUser
#OneNotePsGraph_GetAllNotebooksByGroup
#OneNotePsGraph_GetAllNotebooksBySite
#OneNotePsGraph_CreateOneNotebookMe
#OneNotePsGraph_GetOneNotebookMe
#OneNotePsGraph_GetRecentNotebooksMe
#OneNotePsGraph_CopyOneNotebookMe
#OneNotePsGraph_UpdateOneNotebookMe #==> No methods in Graph API for OneNote
#OneNotePsGraph_DeleteOneNotebookMe #==> No methods in Graph API for OneNote
#OneNotePsGraph_GetAllSectionsMe
#OneNotePsGraph_GetAllSectionsOneNotebookMe
#OneNotePsGraph_GetOneSectionMe
#OneNotePsGraph_CreateOneSectionMe
#OneNotePsGraph_UpdateOneSectionMe #==> It doesn't work
#OneNotePsGraph_DeleteOneSectionMe #==> It doesn't work
#OneNotePsGraph_CopyOneSectionToNotebookMe
#OneNotePsGraph_GetAllSectionGroupsMe
#OneNotePsGraph_GetOneSectionGroupMe
#OneNotePsGraph_CreateOneSectionGroupMe
#OneNotePsGraph_GetAllSectionsOneSectionGroupMe
#OneNotePsGraph_CreateOneSectionInOneSectionGroupMe
#OneNotePsGraph_CopyOneSectionToSectionGroupMe
#OneNotePsGraph_GetAllPagesInOneNoteMe
#OneNotePsGraph_GetAllPagesInOneSectionMe
#OneNotePsGraph_GetOnePageMe
#OneNotePsGraph_CreatePageInOneSectionMe
#OneNotePsGraph_GetOnePageContentMe
#OneNotePsGraph_GetOnePageContentWithIdsMe
#OneNotePsGraph_UpdateOnePageContentMe
#OneNotePsGraph_DeleteOnePageMe

#------------------------ Using PnP CLI

#OneNotePsCli_GetAllNotebooks
#OneNotePsCli_GetAllNotebooksByGroup
#OneNotePsCli_GetAllNotebooksByUser
#OneNotePsCli_GetAllNotebooksBySite
#OneNotePsCli_GetAllPagesByGroup
#OneNotePsCli_GetAllPagesByUser
#OneNotePsCli_GetAllPagesBySite

#------------------------ Using Graph SDK

#OneNotePsGraphSdk_GetAllNotebooksByUser
#OneNotePsGraphSdk_GetOneNotebookByUser
#OneNotePsGraphSdk_GetAllNotebooksByGroup
#OneNotePsGraphSdk_GetAllNotebooksBySite
#OneNotePsGraphSdk_CreateOneNotebookByUserWithBodyParameters
#OneNotePsGraphSdk_CreateOneNotebookByUserWithParameters
#OneNotePsGraphSdk_UpdateOneNotebookByUserWithBodyParameters
#OneNotePsGraphSdk_DeleteOneNotebookByUser
#OneNotePsGraphSdk_GetAllSectionsByUser
#OneNotePsGraphSdk_GetAllSectionsOneNotebookByUser
#OneNotePsGraphSdk_CreateOneSectionInNotebookByUserWithBodyParameters
#OneNotePsGraphSdk_UpdateOneSectionByUserWithBodyParameters
#OneNotePsGraphSdk_DeleteOneSectionByUser
#OneNotePsGraphSdk_GetAllSectionGroupsByUser
#OneNotePsGraphSdk_GetAllSectionGroupsOneNotebookByUser
#OneNotePsGraphSdk_CreateOneSectionGroupInNotebookByUserWithBodyParameters
#OneNotePsGraphSdk_UpdateOneSectionGroupByUserWithBodyParameters
#OneNotePsGraphSdk_DeleteOneSectionGroupByUser
#OneNotePsGraphSdk_GetAllPagesByUser
#OneNotePsGraphSdk_GetAllPagesOneSectionByUser
#OneNotePsGraphSdk_GetOnePageContentByUser
#OneNotePsGraphSdk_GetOnePageResourceByUser
#OneNotePsGraphSdk_GetOnePageResourceConentByUser
#OneNotePsGraphSdk_CreateOnePageInSectionByUserWithBodyParameters
#OneNotePsGraphSdk_CreateOnePageInSectionByUserWithParameters
#OneNotePsGraphSdk_UpdateOnePageByUserWithBodyParameters
#OneNotePsGraphSdk_DeleteOnePageByUser

Write-Host "Done" 

