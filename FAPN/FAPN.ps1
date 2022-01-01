
# Functions to login in Azure

Function Get-AzureTokenApplication(){
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
Function OneNotePsGraphGetAllNotebooksMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 01 

#gavdcodebegin 02
Function OneNotePsGraphGetAllNotebooksUser()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + $UserName + "/onenote/notebooks"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 02 

#gavdcodebegin 03
Function OneNotePsGraphCreateOneNotebookMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'displayName':'NotebookFromPowerShell' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 03 

#gavdcodebegin 04
Function OneNotePsGraphGetOneNotebookMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$bookId = "1-2d02d4f3-f665-4965-a5bc-eb698bd161ad"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 04 

#gavdcodebegin 05
Function OneNotePsGraphGetRecentNotebooksMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + `
					"getRecentNotebooks(includePersonalNotebooks=true)"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 05 

#gavdcodebegin 06
Function OneNotePsGraphCopyOneNotebookMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-2d02d4f3-f665-4965-a5bc-eb698bd161ad"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + `
				"/copyNotebook"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'renameAs':'CopyNotebookFromPowerShell' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 06

Function OneNotePsGraphUpdateOneNotebookMe() # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite

	$bookId = "1-2d02d4f3-f665-4965-a5bc-eb698bd161ad"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'renameAs':'NotebookFromPowerShellUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}

Function OneNotePsGraphDeleteOneNotebookMe() # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$bookId = "1-2d02d4f3-f665-4965-a5bc-eb698bd161ad"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}

#gavdcodebegin 07
Function OneNotePsGraphGetAllSectionsMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 07 

#gavdcodebegin 08
Function OneNotePsGraphGetSectionsOneNotebookMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$bookId = "1-2d02d4f3-f665-4965-a5bc-eb698bd161ad"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + "/sections"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 08

#gavdcodebegin 09
Function OneNotePsGraphGetOneSectionMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionId = "1-042349ca-f9ae-4c56-84d7-a58454db88cb"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 09

#gavdcodebegin 10
Function OneNotePsGraphCreateOneSectionMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-2d02d4f3-f665-4965-a5bc-eb698bd161ad"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + "/sections"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'displayName':'SectionFromPowerShell' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 10 

#gavdcodebegin 11
Function OneNotePsGraphUpdateOneSectionMe() # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite

	$sectionId = "1-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'renameAs':'SectionFromPowerShellUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 11 

Function OneNotePsGraphDeleteOneSectionMe() # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$sectionId = "1-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}

#gavdcodebegin 12
Function OneNotePsGraphCopyOneSectionToNotebookMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-2d02d4f3-f665-4965-a5bc-eb698bd161ad"
	$sectionId = "1-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
				"/copyToNotebook"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'renameAs':'CopySectionFromPowerShell', `
			     'id':'$($bookId)' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 12

#gavdcodebegin 13
Function OneNotePsGraphGetAllSectionGroupsMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 13 

#gavdcodebegin 14
Function OneNotePsGraphGetOneSectionGroupMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionGroupId = "1-66238e38-e4a9-48bf-bddb-aece7d19e4a6"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups/" + $sectionGroupId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 14

#gavdcodebegin 15
Function OneNotePsGraphCreateOneSectionGroupMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$bookId = "1-2d02d4f3-f665-4965-a5bc-eb698bd161ad"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks/" + $bookId + `
					"/sectionGroups"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'displayName':'SectionGroupFromPowerShell' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 15

#gavdcodebegin 16
Function OneNotePsGraphGetAllSectionsOneSectionGroupMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionGroupId = "1-66238e38-e4a9-48bf-bddb-aece7d19e4a6"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups/" + $sectionGroupId + `
					"/sections"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 16

#gavdcodebegin 17
Function OneNotePsGraphCreateOneSectionInOneSectionGroupMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$sectionGroupId = "1-99c94cb0-10a4-43fd-b595-4024e2f3d4c3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sectionGroups/" + $sectionGroupId + `
					"/sections"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'displayName':'SectionInSectionGroupFromPowerShell' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 17

#gavdcodebegin 18
Function OneNotePsGraphCopyOneSectionToSectionGroupMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$sectionId = "1-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$sectionGroupId = "1-99c94cb0-10a4-43fd-b595-4024e2f3d4c3"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
				"/copyToSectionGroup"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ 'renameAs':'CopySectionToSectionGroupFromPowerShell', `
				 'id':'$($sectionGroupId)' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 18

#gavdcodebegin 19
Function OneNotePsGraphGetAllPagesInOneNoteMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 19

#gavdcodebegin 20
Function OneNotePsGraphGetAllPagesInOneSectionMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$sectionId = "1-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
					"/pages"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 20

#gavdcodebegin 21
Function OneNotePsGraphGetOnePageMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-3a296a44f7584a1fa5f81234ecfb026b!11-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 21

#gavdcodebegin 22
Function OneNotePsGraphCreatePageInOneSectionMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite, Notes.Create

	$sectionId = "1-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/sections/" + $sectionId + `
					"/pages"

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "<!DOCTYPE html>" + `
			  "	<html>" + `
			  "	  <head>" + `
			  "		<title>PageInSectionFromPowerShell</title>" + `
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
#gavdcodeend 22

#gavdcodebegin 23
Function OneNotePsGraphGetOnePageContentMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-3a296a44f7584a1fa5f81234ecfb026b!11-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId + "/content"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 23

#gavdcodebegin 24
Function OneNotePsGraphGetOnePageContentWithIdsMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.Read, Notes.ReadWrite

	$pageId = "1-3a296a44f7584a1fa5f81234ecfb026b!11-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId + `
					"/content?includeIDs=true"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 24

#gavdcodebegin 25
Function OneNotePsGraphUpdateOnePageContentMe()
{
	# App Registration type:		Delegation
	# App Registration permissions: Notes.ReadWrite

	$pageId = "1-3a296a44f7584a1fa5f81234ecfb026b!11-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$targetId = "p:{bbdd7818-ab8e-4563-bfc8-d5f337b5b2b2}{22}"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $spageId + "/content"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "[{ 'target':'$($targetId)', `
				  'action':'append', `
				  'position':'after', `
				  'content':' - <b>Updated with PowerShell</b>' }]"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 25 

#gavdcodebegin 26
Function OneNotePsGraphDeleteOnePageMe() # It doesn't work
{
	# App Registration type:		Delegation
	# App Registration permissions: Tasks.ReadWrite

	$pageId = "1-98d37a3eabef4293b8f2d5fe7d9fc3db!49-1c90e552-0a4a-44e1-a6cb-daaaade2769a"
	$Url = "https://graph.microsoft.com/v1.0/me/onenote/pages/" + $pageId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}
#gavdcodeend 26

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\grPs.values.config"

#$ClientIDApp = $configFile.appsettings.ClientIdApp
#$ClientSecretApp = $configFile.appsettings.ClientSecretApp
$ClientIDDel = $configFile.appsettings.ClientIdDel
$TenantName = $configFile.appsettings.TenantName
$UserName = $configFile.appsettings.UserName
$UserPw = $configFile.appsettings.UserPw

#------------------------ Using Microsoft Graph PowerShell for Teams

#OneNotePsGraphGetAllNotebooksMe
#OneNotePsGraphGetAllNotebooksUser
#OneNotePsGraphCreateOneNotebookMe
#OneNotePsGraphGetOneNotebookMe
#OneNotePsGraphGetRecentNotebooksMe
#OneNotePsGraphCopyOneNotebookMe
#OneNotePsGraphUpdateOneNotebookMe ==> It doesn't work
#OneNotePsGraphDeleteOneNotebookMe ==> It doesn't work
#OneNotePsGraphGetAllSectionsMe
#OneNotePsGraphGetSectionsOneNotebookMe
#OneNotePsGraphGetOneSectionMe
#OneNotePsGraphCreateOneSectionMe
#OneNotePsGraphUpdateOneSectionMe ==> It doesn't work
#OneNotePsGraphDeleteOneSectionMe ==> It doesn't work
#OneNotePsGraphCopyOneSectionToNotebookMe
#OneNotePsGraphGetAllSectionGroupsMe
#OneNotePsGraphGetOneSectionGroupMe
#OneNotePsGraphCreateOneSectionGroupMe
#OneNotePsGraphGetAllSectionsOneSectionGroupMe
#OneNotePsGraphCreateOneSectionInOneSectionGroupMe
#OneNotePsGraphCopyOneSectionToSectionGroupMe
#OneNotePsGraphGetAllPagesInOneNoteMe
#OneNotePsGraphGetAllPagesInOneSectionMe
#OneNotePsGraphGetOnePageMe
#OneNotePsGraphCreatePageInOneSectionMe
#OneNotePsGraphGetOnePageContentMe
#OneNotePsGraphGetOnePageContentWithIdsMe
#OneNotePsGraphUpdateOnePageContentMe
#OneNotePsGraphDeleteOnePageMe

#------------------------ Using Microsoft PnP CLI for Teams

#OneNotePsCliGetAllLists

Write-Host "Done" 

