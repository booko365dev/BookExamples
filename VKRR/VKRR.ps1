 
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
Function GrPsGetAllSites()
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
#gavdcodeend 01 

#gavdcodebegin 02
Function GrPsGetOneSiteById()
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
#gavdcodeend 02 

#gavdcodebegin 03
Function GrPsGetOneSiteByPath()
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
#gavdcodeend 03 

#gavdcodebegin 04
Function GrPsGetSitesBySearch()
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
#gavdcodeend 04 

#gavdcodebegin 16
Function GrPsGetOneSiteAnalytics()
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
#gavdcodeend 16 

#gavdcodebegin 05
Function GrPsGetSitesFollowed()
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
#gavdcodeend 05 

#gavdcodebegin 06
Function GrPsSitesUnfollow()
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
#gavdcodeend 06 

#gavdcodebegin 07
Function GrPsSitesFollow()
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
#gavdcodeend 07 

#gavdcodebegin 08
Function GrPsGetAllListsInSite()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
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
#gavdcodeend 08 

#gavdcodebegin 09
Function GrPsGetOneListInSite()
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
#gavdcodeend 09 

#gavdcodebegin 10
Function GrPsCreateList()
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
#gavdcodeend 10 

#gavdcodebegin 11
Function GrPsGetAllItemsInList()
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
#gavdcodeend 11 

#gavdcodebegin 12
Function GrPsGetOneItemInList()
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
#gavdcodeend 12 

#gavdcodebegin 17
Function GrPsGetOneItemAnalytics()
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
#gavdcodeend 17 

#gavdcodebegin 13
Function GrPsCreateItem()
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
#gavdcodeend 13 

#gavdcodebegin 14
Function GrPsUpdateItem()
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
#gavdcodeend 14 

#gavdcodebegin 15
Function GrPsDeleteItem()
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
#gavdcodeend 15 

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\grPs.values.config"

$ClientIDApp = $configFile.appsettings.ClientIdApp
$ClientSecretApp = $configFile.appsettings.ClientSecretApp
$TenantName = $configFile.appsettings.TenantName
$UserName = $configFile.appsettings.UserName

#GrPsGetAllSites 
#GrPsGetOneSiteById
#GrPsGetOneSiteByPath
#GrPsGetSitesBySearch
#GrPsGetOneSiteAnalytics
#GrPsGetSitesFollowed
#GrPsSitesUnfollow
#GrPsSitesFollow
#GrPsGetAllListsInSite
#GrPsGetOneListInSite
#GrPsCreateList
#GrPsGetAllItemsInList
#GrPsGetOneItemInList
#GrPsGetOneItemAnalytics
#GrPsCreateItem
#GrPsUpdateItem
#GrPsDeleteItem

Write-Host "Done" 
