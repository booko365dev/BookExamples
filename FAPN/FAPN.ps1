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

Function PsGraphSDK_LoginWithAccPw
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

	[SecureString]$secureToken = ConvertTo-SecureString -String `
											$myToken.AccessToken -AsPlainText -Force

	Connect-Graph -AccessToken $secureToken
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


##==> Graph

#gavdcodebegin 001
Function PsOneNoteGraphRestApi_GetAllNotebooksMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 001 

#gavdcodebegin 002
Function PsOneNoteGraphRestApi_GetAllNotebooksByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
									$cnfUserName + "/onenote/notebooks"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 002 

#gavdcodebegin 027
Function PsOneNoteGraphRestApi_GetAllNotebooksByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$groupId = "CF7E4CB9-E929-43D9-84BA-BD7C123DAAE9"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + `
							$groupId + "/onenote/notebooks"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 027 

#gavdcodebegin 028
Function PsOneNoteGraphRestApi_GetAllNotebooksBySite
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$siteId = "FCB9425A-E423-4988-8611-ACACEA52400B"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + `
							$siteId + "/onenote/notebooks"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 028 

#gavdcodebegin 003
Function PsOneNoteGraphRestApi_CreateOneNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'displayName':'NotebookCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 003 

#gavdcodebegin 004
Function PsOneNoteGraphRestApi_GetOneNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 004 

#gavdcodebegin 005
Function PsOneNoteGraphRestApi_GetRecentNotebooksMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + `
					"getRecentNotebooks(includePersonalNotebooks=true)"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 005 

#gavdcodebegin 006
Function PsOneNoteGraphRestApi_CopyOneNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + `
				"/copyNotebook"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'renameAs':'CopyNotebookCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 006

#gavdcodebegin 064
Function PsOneNoteGraphRestApi_GetDriverForNotebooksMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
	# $Url = "https://graph.microsoft.com/v1.0/users/" + $cnfUserName + "/drive/root/children"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 064 

#gavdcodebegin 065
Function PsOneNoteGraphRestApi_UpdateOneNotebookOneUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite

	$driveId = "01ASKQTCCWR7GXCRMEAJHKCMBVIOQ35LVC"
	$notebookName = "Test_Notebook_01"
	$Url = "https://graph.microsoft.com/v1.0/users/" + $cnfUserName + "/drive/items/" + `
									$driveId + "/children/" + $notebookName
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'renameAs':'NotebookUpdatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 065

#gavdcodebegin 066
Function PsOneNoteGraphRestApi_DeleteOneNotebookOneUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$driveId = "01ASKQTCCWR7GXCRMEAJHKCMBVIOQ35LVC"
	$notebookName = "Test_Notebook_01"
	$Url = "https://graph.microsoft.com/v1.0/users/" + $cnfUserName + "/drive/items/" + `
									$driveId + "/children/" + $notebookName
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}
#gavdcodeend 066

#gavdcodebegin 007
Function PsOneNoteGraphRestApi_GetAllSectionsMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 007 

#gavdcodebegin 008
Function PsOneNoteGraphRestApi_GetAllSectionsOneNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + "/sections"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 008

#gavdcodebegin 009
Function PsOneNoteGraphRestApi_GetOneSectionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionId = "1-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsOneNoteGraphRestApi_CreateOneSectionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-2bf37229-95e5-4a1c-80ab-12db361f1719"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + "/sections"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'displayName':'SectionCreatedWithGraphRestApi' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 010 

#gavdcodebegin 011
Function PsOneNoteGraphRestApi_UpdateOneSectionMe # It doesn't work
{
	# ATTENTION: The routine gives no error, but the display name is not changed at all

	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite

	$sectionId = "1-e4161eda-bb8e-4b3b-b9d9-d88aa09be232"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'renameAs':'SectionUpdatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 011 

Function PsOneNoteGraphRestApi_DeleteOneSectionMe # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$sectionId = "1-e4161eda-bb8e-4b3b-b9d9-d88aa09be232"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}

#gavdcodebegin 012
Function PsOneNoteGraphRestApi_CopyOneSectionToNotebookMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$sectionId = "1-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
				"/copyToNotebook"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
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
Function PsOneNoteGraphRestApi_GetAllSectionGroupsMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 013 

#gavdcodebegin 014
Function PsOneNoteGraphRestApi_GetOneSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionGroupId = "1-27b4559e-63a9-4beb-b6d7-772df43e5f19"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups/" + $sectionGroupId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 014

#gavdcodebegin 015
Function PsOneNoteGraphRestApi_CreateOneSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-bcaad78a-23d3-4a8e-bdc8-3a79165b5bbe"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + `
					"/sectionGroups"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'displayName':'SectionGroupCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 015

#gavdcodebegin 016
Function PsOneNoteGraphRestApi_GetAllSectionsOneSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionGroupId = "1-31330f7c-afe6-4554-aa7e-7c0a73a64c09"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups/" + `
						$sectionGroupId + "/sections"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 016

#gavdcodebegin 017
Function PsOneNoteGraphRestApi_CreateOneSectionInOneSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$sectionGroupId = "1-31330f7c-afe6-4554-aa7e-7c0a73a64c09"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups/" + `
						$sectionGroupId + "/sections"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myBody = "{ 'displayName':'SectionInSectionGroupCreatedWithGraph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 017

#gavdcodebegin 018
Function PsOneNoteGraphRestApi_CopyOneSectionToSectionGroupMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$sectionGroupId = "1-31330f7c-afe6-4554-aa7e-7c0a73a64c09"
	$sectionId = "1-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
				"/copyToSectionGroup"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
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
Function PsOneNoteGraphRestApi_GetAllPagesInOneNoteMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 019

#gavdcodebegin 020
Function PsOneNoteGraphRestApi_GetAllPagesInOneSectionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionId = "1-e4161eda-bb8e-4b3b-b9d9-d88aa09be232"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
					"/pages"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 020

#gavdcodebegin 021
Function PsOneNoteGraphRestApi_GetOnePageMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-8c31cf8ccc4a410ea79eb50c6f04f963!14-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 021

#gavdcodebegin 022
Function PsOneNoteGraphRestApi_CreatePageInOneSectionMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$sectionId = "1-e4161eda-bb8e-4b3b-b9d9-d88aa09be232"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
					"/pages"

	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
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
Function PsOneNoteGraphRestApi_GetOnePageContentMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-367fa14618504079824c3320806db6a2!6-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId + "/content"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 023

#gavdcodebegin 024
Function PsOneNoteGraphRestApi_GetOnePageContentWithIdsMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-367fa14618504079824c3320806db6a2!6-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId + `
					"/content?includeIDs=true"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 024

#gavdcodebegin 025
Function PsOneNoteGraphRestApi_UpdateOnePageContentMe
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite

	$pageId = "1-367fa14618504079824c3320806db6a2!6-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$targetId = "p:{1d122d8f-b983-4903-8c1a-43e2f2a0cae5}{101}"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $spageId + "/content"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
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
Function PsOneNoteGraphRestApi_DeleteOnePageMe # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-367fa14618504079824c3320806db6a2!6-ec9017de-dd76-4492-8dcf-d83687235cb3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegationWithAccPw `
									-ClientID $cnfClientIdWithAccPw `
									-TenantName $cnfTenantName `
									-UserName $cnfUserName `
									-UserPw $cnfUserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}
#gavdcodeend 026

#-----------------------------------------------------------------------------------------

##==> M365 CLI

#gavdcodebegin 029
Function PsOneNoteM365Cli_GetAllNotebooks
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 onenote notebook list

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
Function PsOneNoteM365Cli_GetAllNotebooksByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 onenote notebook list --groupName "MyM365Group"
	m365 onenote notebook list --groupId "B5FD0AE8-0695-489E-B142-3A2C36AC43B2"

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
Function PsOneNoteM365Cli_GetAllNotebooksByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 onenote notebook list --userName $configFile.appsettings.UserName
	m365 onenote notebook list --userId "765764B2-3C30-40DA-B293-0CB28B8E8148"

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
Function PsOneNoteM365Cli_GetAllNotebooksBySite
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 onenote notebook list --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 032

#gavdcodebegin 067
Function PsOneNoteM365Cli_CreateOneNotebook
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 onenote notebook add --name "Test_Notebook_03"
	#m365 onenote notebook add --name "Test_Notebook_03" --groupId "[GroupId]"
	#m365 onenote notebook add --name "Test_Notebook_03" --groupName "[GroupName]"
	#m365 onenote notebook add --name "Test_Notebook_03" --userId "[UserId]"
	#m365 onenote notebook add --name "Test_Notebook_03" --userName "[UeserName]"
	#m365 onenote notebook add --name "Test_Notebook_03" `
	#						   --webUrl "https://domain.sharepoint.com/sites/SiteName"

	m365 logout
}
#gavdcodeend 067

#gavdcodebegin 033
Function PsOneNoteM365Cli_GetAllPages
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 onenote page list

	m365 logout
}
#gavdcodeend 033

#gavdcodebegin 034
Function PsOneNoteM365Cli_GetAllPagesByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 onenote page list --groupName "MyM365Group"
	m365 onenote page list --groupId "B5FD0AE8-0695-489E-B142-3A2C36AC43B2"

	m365 logout
}
#gavdcodeend 034

#gavdcodebegin 035
Function PsOneNoteM365Cli_GetAllPagesByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 onenote page list --userName $configFile.appsettings.UserName
	m365 onenote page list --userId "765764B2-3C30-40DA-B293-0CB28B8E8148"

	m365 logout
}
#gavdcodeend 035

#gavdcodebegin 036
Function PsOneNoteM365Cli_GetAllPagesBySite
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
	m365 onenote page list --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 036

#-----------------------------------------------------------------------------------------

##==> Graph PowerShell SDK

#gavdcodebegin 037
Function PsOneNoteGraphSdk_GetAllNotebooksByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenoteNotebook -UserId $cnfUserName
	# Get-MgUserOnenoteNotebook -UserId "B193A6C2-CBDD-43A8-B080-09BB93FDF8A1"

	Disconnect-MgGraph
}
#gavdcodeend 037

#gavdcodebegin 038
Function PsOneNoteGraphSdk_GetOneNotebookByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenoteNotebook -UserId $cnfUserName `
							  -NotebookId "1-83eec81d-be98-41b0-aee1-177e6e967f48"

	Disconnect-MgGraph
}
#gavdcodeend 038

#gavdcodebegin 039
Function PsOneNoteGraphSdk_GetAllNotebooksByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgGroupOnenoteNotebook -GroupId "5f41785a-87f6-4c70-9e5f-20da7e0e7ba4" 

	Disconnect-MgGraph
}
#gavdcodeend 039

#gavdcodebegin 040
Function PsOneNoteGraphSdk_GetAllNotebooksBySite
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgSiteOnenoteNotebook -SiteId "1ec84def-070a-4a94-a26d-2845b3bf6ff3"

	Disconnect-MgGraph
}
#gavdcodeend 040

#gavdcodebegin 041
Function PsOneNoteGraphSdk_CreateOneNotebookByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	$NotebookParameters = @{
		DisplayName = "NotebookCreatedWithGraphSDKBodyParams"
	}

	New-MgUserOnenoteNotebook -UserId $cnfUserName `
							  -BodyParameter $NotebookParameters

	Disconnect-MgGraph
}
#gavdcodeend 041

#gavdcodebegin 042
Function PsOneNoteGraphSdk_CreateOneNotebookByUserWithParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	New-MgUserOnenoteNotebook -UserId $cnfUserName `
							  -DisplayName "NotebookCreatedWithGraphSDKParams"

	Disconnect-MgGraph
}
#gavdcodeend 042

#gavdcodebegin 043
Function PsOneNoteGraphSdk_UpdateOneNotebookByUserWithBodyParameters # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	$NotebookParameters = @{
		DisplayName = "NotebookUpdatedWithGraphSDKBodyParams"
	}

	Update-MgUserOnenoteNotebook -UserId $cnfUserName `
								 -NotebookId "1-18e3b013-0991-48db-926c-deda8f5a53f9" `
								 -BodyParameter $NotebookParameters

	Disconnect-MgGraph
}
#gavdcodeend 043

#gavdcodebegin 044
Function PsOneNoteGraphSdk_DeleteOneNotebookByUser # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Remove-MgUserOnenoteNotebook -UserId $cnfUserName `
								 -NotebookId "1-18e3b013-0991-48db-926c-deda8f5a53f9"

	Disconnect-MgGraph
}
#gavdcodeend 044

#gavdcodebegin 045
Function PsOneNoteGraphSdk_GetAllSectionsByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenoteSection -UserId $cnfUserName
	# Get-MgUserOnenoteSection -UserId "B193A6C2-CBDD-43A8-B080-09BB93FDF8A1"

	Disconnect-MgGraph
}
#gavdcodeend 045

#gavdcodebegin 046
Function PsOneNoteGraphSdk_GetAllSectionsOneNotebookByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenoteNotebookSection -UserId $cnfUserName `
									 -NotebookId "1-18e3b013-0991-48db-926c-deda8f5a53f9"

	Disconnect-MgGraph
}
#gavdcodeend 046

#gavdcodebegin 047
Function PsOneNoteGraphSdk_CreateOneSectionInNotebookByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	$SectionParameters = @{
		DisplayName = "SectionCreatedWithGraphSDKBodyParams"
	}

	New-MgUserOnenoteNotebookSection -UserId $cnfUserName `
								-NotebookId "1-18e3b013-0991-48db-926c-deda8f5a53f9" `
								-BodyParameter $SectionParameters

	Disconnect-MgGraph
}
#gavdcodeend 047

#gavdcodebegin 048
Function PsOneNoteGraphSdk_UpdateOneSectionByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	$SectionParameters = @{
		DisplayName = "SectionUpdatedWithGraphSDKBodyParams"
	}

	Update-MgUserOnenoteSection -UserId $cnfUserName `
							-OnenoteSectionId "1-7fce7c56-b39a-4b51-adc3-798050310c3b" `
							-BodyParameter $SectionParameters

	Disconnect-MgGraph
}
#gavdcodeend 048

#gavdcodebegin 049
Function PsOneNoteGraphSdk_DeleteOneSectionByUser # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Remove-MgUserOnenoteSection -UserId $cnfUserName `
							-OnenoteSectionId "1-7fce7c56-b39a-4b51-adc3-798050310c3b"

	Disconnect-MgGraph
}
#gavdcodeend 049

#gavdcodebegin 050
Function PsOneNoteGraphSdk_GetAllSectionGroupsByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenoteSectionGroup -UserId $cnfUserName
	Get-MgUserOnenoteSectionGroup -UserId "B193A6C2-CBDD-43A8-B080-09BB93FDF8A1"

	Disconnect-MgGraph
}
#gavdcodeend 050

#gavdcodebegin 051
Function PsOneNoteGraphSdk_GetAllSectionGroupsOneNotebookByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenoteNotebookSectionGroup -UserId $cnfUserName `
									 -NotebookId "1-41dc1f07-d7c3-4913-96a5-7c27a973050a"

	Disconnect-MgGraph
}
#gavdcodeend 051

#gavdcodebegin 052
Function PsOneNoteGraphSdk_CreateOneSectionGroupInNotebookByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	$SectionParameters = @{
		DisplayName = "SectionGroupCreatedWithGraphSDKBodyParams"
	}

	New-MgUserOnenoteNotebookSectionGroup -UserId $cnfUserName `
								-NotebookId "1-41dc1f07-d7c3-4913-96a5-7c27a973050a" `
								-BodyParameter $SectionParameters

	Disconnect-MgGraph
}
#gavdcodeend 052

#gavdcodebegin 053
Function PsOneNoteGraphSdk_UpdateOneSectionGroupByUserWithBodyParameters # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	$SectionParameters = @{
		DisplayName = "SectionGroupUpdatedWithGraphSDKBodyParams"
	}

	Update-MgUserOnenoteSectionGroup -UserId $cnfUserName `
						-SectionGroupId "1-60a12e43-8562-4919-8e77-4f7b633012a2" `
						-BodyParameter $SectionParameters

	Disconnect-MgGraph
}
#gavdcodeend 053

#gavdcodebegin 054
Function PsOneNoteGraphSdk_DeleteOneSectionGroupByUser # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Remove-MgUserOnenoteSectionGroup -UserId $cnfUserName `
							    -SectionGroupId "1-60a12e43-8562-4919-8e77-4f7b633012a2"

	Disconnect-MgGraph
}
#gavdcodeend 054

#gavdcodebegin 055
Function PsOneNoteGraphSdk_GetAllPagesByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenotePage -UserId $cnfUserName
	# Get-MgUserOnenotePage -UserId "B193A6C2-CBDD-43A8-B080-09BB93FDF8A1"

	Disconnect-MgGraph
}
#gavdcodeend 055

#gavdcodebegin 056
Function PsOneNoteGraphSdk_GetAllPagesOneSectionByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenoteSectionPage -UserId $cnfUserName `
							-OnenoteSectionId "1-f2474782-6ea0-41f4-a187-bc1d1715cc3f"

	Disconnect-MgGraph
}
#gavdcodeend 056

#gavdcodebegin 057
Function PsOneNoteGraphSdk_GetOnePageContentByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenotePageContent -UserId $cnfUserName `
								 -OnenotePageId "1-ffea745...1d1715cc3f" `
								 -OutFile "C:\Temporary\page.txt"

	Disconnect-MgGraph
}
#gavdcodeend 057

#gavdcodebegin 058
Function PsOneNoteGraphSdk_GetOnePageResourceByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenoteResource -UserId $cnfUserName `
							  -OnenoteResourceId "1-0e975ac85...d1715cc3f"

	Disconnect-MgGraph
}
#gavdcodeend 058

#gavdcodebegin 059
Function PsOneNoteGraphSdk_GetOnePageResourceConentByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Get-MgUserOnenoteResourceContent -UserId $cnfUserName `
									 -OnenoteResourceId "1-0e975ac8...715cc3f" `
									 -OutFile "C:\Temporary\pagecontent.txt"

	Disconnect-MgGraph
}
#gavdcodeend 059

#gavdcodebegin 060
Function PsOneNoteGraphSdk_CreateOnePageInSectionByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	$PageParameters = @{
		Title = "PageCreatedWithGraphSDKBodyParams"
	}

	New-MgUserOnenoteSectionPage -UserId $cnfUserName `
							-OnenoteSectionId "1-7fce7c56-b39a-4b51-adc3-798050310c3b" `
							-BodyParameter $PageParameters

	Disconnect-MgGraph
}
#gavdcodeend 060

#gavdcodebegin 061
Function PsOneNoteGraphSdk_CreateOnePageInSectionByUserWithParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	New-MgUserOnenoteSectionPage -UserId $cnfUserName `
							-OnenoteSectionId "1-f2474782-6ea0-41f4-a187-bc1d1715cc3f" `
							-Title "PageCreatedWithGraphSDKBodyParams"

	Disconnect-MgGraph
}
#gavdcodeend 061

#gavdcodebegin 062
Function PsOneNoteGraphSdk_UpdateOnePageByUserWithBodyParameters
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	$PageParameters = @{
		Title = "PageUpdatedWithGraphSDKBodyParams"
	}

	Update-MgUserOnenotePage -UserId $cnfUserName `
							 -OnenotePageId "1-0c0a271ca063459-d88aa09be232" `
							 -BodyParameter $PageParameters

	Disconnect-MgGraph
}
#gavdcodeend 062

#gavdcodebegin 063
Function PsOneNoteGraphSdk_DeleteOnePageByUser
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						   -ClientID $cnfClientIdWithAccPw `
						   -UserName $cnfUserName `
						   -UserPw $cnfUserPw

	Remove-MgUserOnenotePage -UserId $cnfUserName `
							 -OnenotePageId "1-0c0a271ca06345d697b375ebaab1f51b!231-e4161eda-bb8e-4b3b-b9d9-d88aa09be232"

	Disconnect-MgGraph
}
#gavdcodeend 063


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

# *** Latest Source Code Index: 67 ***

#------------------------ Using Microsoft Graph PowerShell

#PsOneNoteGraphRestApi_GetAllNotebooksMe
#PsOneNoteGraphRestApi_GetAllNotebooksByUser
#PsOneNoteGraphRestApi_GetAllNotebooksByGroup
#PsOneNoteGraphRestApi_GetAllNotebooksBySite
#PsOneNoteGraphRestApi_CreateOneNotebookMe
#PsOneNoteGraphRestApi_GetOneNotebookMe
#PsOneNoteGraphRestApi_GetRecentNotebooksMe
#PsOneNoteGraphRestApi_CopyOneNotebookMe
#PsOneNoteGraphRestApi_GetDriverForNotebooksMe
#PsOneNoteGraphRestApi_UpdateOneNotebookOneUser
#PsOneNoteGraphRestApi_DeleteOneNotebookOneUser
#PsOneNoteGraphRestApi_GetAllSectionsMe
#PsOneNoteGraphRestApi_GetAllSectionsOneNotebookMe
#PsOneNoteGraphRestApi_GetOneSectionMe
#PsOneNoteGraphRestApi_CreateOneSectionMe
#PsOneNoteGraphRestApi_UpdateOneSectionMe #==> It doesn't work
#PsOneNoteGraphRestApi_DeleteOneSectionMe #==> It doesn't work
#PsOneNoteGraphRestApi_CopyOneSectionToNotebookMe
#PsOneNoteGraphRestApi_GetAllSectionGroupsMe
#PsOneNoteGraphRestApi_GetOneSectionGroupMe
#PsOneNoteGraphRestApi_CreateOneSectionGroupMe
#PsOneNoteGraphRestApi_GetAllSectionsOneSectionGroupMe
#PsOneNoteGraphRestApi_CreateOneSectionInOneSectionGroupMe
#PsOneNoteGraphRestApi_CopyOneSectionToSectionGroupMe
#PsOneNoteGraphRestApi_GetAllPagesInOneNoteMe
#PsOneNoteGraphRestApi_GetAllPagesInOneSectionMe
#PsOneNoteGraphRestApi_GetOnePageMe
#PsOneNoteGraphRestApi_CreatePageInOneSectionMe
#PsOneNoteGraphRestApi_GetOnePageContentMe
#PsOneNoteGraphRestApi_GetOnePageContentWithIdsMe
#PsOneNoteGraphRestApi_UpdateOnePageContentMe
#PsOneNoteGraphRestApi_DeleteOnePageMe

#------------------------ Using Microsoft 365 CLI

#PsOneNoteM365Cli_GetAllNotebooks
#PsOneNoteM365Cli_GetAllNotebooksByGroup
#PsOneNoteM365Cli_GetAllNotebooksByUser
#PsOneNoteM365Cli_GetAllNotebooksBySite
#PsOneNoteM365Cli_CreateOneNotebook
#PsOneNoteM365Cli_GetAllPages
#PsOneNoteM365Cli_GetAllPagesByGroup
#PsOneNoteM365Cli_GetAllPagesByUser
#PsOneNoteM365Cli_GetAllPagesBySite

#------------------------ Using Graph PowerShell SDK

#PsOneNoteGraphSdk_GetAllNotebooksByUser
#PsOneNoteGraphSdk_GetOneNotebookByUser
#PsOneNoteGraphSdk_GetAllNotebooksByGroup
#PsOneNoteGraphSdk_GetAllNotebooksBySite
#PsOneNoteGraphSdk_CreateOneNotebookByUserWithBodyParameters
#PsOneNoteGraphSdk_CreateOneNotebookByUserWithParameters
#PsOneNoteGraphSdk_UpdateOneNotebookByUserWithBodyParameters
#PsOneNoteGraphSdk_DeleteOneNotebookByUser
#PsOneNoteGraphSdk_GetAllSectionsByUser
#PsOneNoteGraphSdk_GetAllSectionsOneNotebookByUser
#PsOneNoteGraphSdk_CreateOneSectionInNotebookByUserWithBodyParameters
#PsOneNoteGraphSdk_UpdateOneSectionByUserWithBodyParameters
#PsOneNoteGraphSdk_DeleteOneSectionByUser
#PsOneNoteGraphSdk_GetAllSectionGroupsByUser
#PsOneNoteGraphSdk_GetAllSectionGroupsOneNotebookByUser
#PsOneNoteGraphSdk_CreateOneSectionGroupInNotebookByUserWithBodyParameters
#PsOneNoteGraphSdk_UpdateOneSectionGroupByUserWithBodyParameters
#PsOneNoteGraphSdk_DeleteOneSectionGroupByUser
#PsOneNoteGraphSdk_GetAllPagesByUser
#PsOneNoteGraphSdk_GetAllPagesOneSectionByUser
#PsOneNoteGraphSdk_GetOnePageContentByUser
#PsOneNoteGraphSdk_GetOnePageResourceByUser
#PsOneNoteGraphSdk_GetOnePageResourceConentByUser
#PsOneNoteGraphSdk_CreateOnePageInSectionByUserWithBodyParameters
#PsOneNoteGraphSdk_CreateOnePageInSectionByUserWithParameters
#PsOneNoteGraphSdk_UpdateOnePageByUserWithBodyParameters
#PsOneNoteGraphSdk_DeleteOnePageByUser

Write-Host "Done" 

