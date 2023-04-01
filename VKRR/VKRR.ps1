 
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

#----------------------------------------------------------------------------------------

#gavdcodebegin 001
Function GrPs_GetAllSites
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/sites"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 001 

#gavdcodebegin 002
Function GrPs_GetOneSiteById
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "ec8c7d0f-c887-4318-8c0b-b2b88b12bc29"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 002 

#gavdcodebegin 003
Function GrPs_GetOneSiteByPath
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SitePath = "[domain].sharepoint.com:/sites/[SiteName]"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SitePath
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 003 

#gavdcodebegin 004
Function GrPs_GetSitesBySearch
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/sites?search='My Site'"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 004 

#gavdcodebegin 016
Function GrPs_GetOneSiteAnalytics
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "ec8c7d0f-c887-4318-8c0b-b2b88b12bc29"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/analytics"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 016 

#gavdcodebegin 005
Function GrPs_GetSitesFollowed
{
	# App Registration type:		Application
	# App Registration permissions: Files.Read.All, Files.ReadWrite.All, Sites.Read.All,	
	#								Sites.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/followedSites"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 005 

#gavdcodebegin 006
Function GrPs_SitesUnfollow
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$WebId = "cfaf31a6-f73f-4d7d-af24-6e530c022b5c"
	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/followedSites/remove"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = '{ "value": [{ "id": "' + $TenantName + ',' + $SiteId + ',' + $WebId + '" 
				}]}'
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 006 

#gavdcodebegin 007
Function GrPs_SitesFollow()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$WebId = "cfaf31a6-f73f-4d7d-af24-6e530c022b5c"
	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/followedSites/add"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = '{ "value": [{ "id": "' + $TenantName + ',' + $SiteId + ',' + $WebId + '" 
				}]}'
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 007 

#gavdcodebegin 008
Function GrPs_GetAllListsInSite
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "870ae987-120f-45ed-aa6e-b4a6b7bc226e"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 008 

#gavdcodebegin 009
Function GrPs_GetOneListInSite
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListTitle = "Documents"   #Use the DisplayName or the ListId
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListTitle
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 009 

#gavdcodebegin 010
Function GrPs_CreateList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ 'displayName' = 'MyList'; 
				 'list' = @{ 'template' = 'genericList' 
				}} | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 010 

#gavdcodebegin 011
Function GrPs_GetAllItemsInList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "Documents"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
																				"/items"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 011 

#gavdcodebegin 012
Function GrPs_GetOneItemInList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "Documents"
	$ItemId = "11"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
																 "/items/" + $ItemId
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 012 

#gavdcodebegin 017
Function GrPs_GetOneItemAnalytics
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "Documents"
	$ItemId = "11"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
													"/items/" + $ItemId + "/analytics"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 017 

#gavdcodebegin 013
Function GrPs_CreateItem
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "MyList"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
																 "/items"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ 'fields' = @{ 'Title' = 'MyItem' 
				}} | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 013 

#gavdcodebegin 014
Function GrPs_UpdateItem
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "MyList"
	$ItemId = "1"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
													"/items/" + $ItemId + "/fields"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ 'Description' = 'MyDescription' 
				} | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 014 

#gavdcodebegin 015
Function GrPs_DeleteItem
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "MyList"
	$ItemId = "1"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
													"/items/" + $ItemId

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ 'Description' = 'MyDescription' 
				} | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 015 

#----------------------------------------------------------------------------------------

## Running the Functions

# *** Latest Source Code Index: 17 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$ClientIDApp = $configFile.appsettings.ClientIdWithSecret
$ClientSecretApp = $configFile.appsettings.ClientSecret
$TenantName = $configFile.appsettings.TenantName
$UserName = $configFile.appsettings.UserName

#GrPs_GetAllSites 
#GrPs_GetOneSiteById
#GrPs_GetOneSiteByPath
#GrPs_GetSitesBySearch
#GrPs_GetOneSiteAnalytics
#GrPs_GetSitesFollowed
#GrPs_SitesUnfollow
#GrPs_SitesFollow
#GrPs_GetAllListsInSite
#GrPs_GetOneListInSite
#GrPs_CreateList
#GrPs_GetAllItemsInList
#GrPs_GetOneItemInList
#GrPs_GetOneItemAnalytics
#GrPs_CreateItem
#GrPs_UpdateItem
#GrPs_DeleteItem

Write-Host "Done" 
