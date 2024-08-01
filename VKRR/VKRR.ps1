 
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsSpGraphRest_GetAzureTokenApplication
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

function PsSpGraphRest_GetAzureTokenDelegation
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


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpGraphRest_GetAllSites
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/sites"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_GetOneSiteById
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "ec8c7d0f-c887-4318-8c0b-b2b88b12bc29"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_GetOneSiteByPath
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SitePath = "[domain].sharepoint.com:/sites/[SiteName]"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SitePath
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_GetSitesBySearch
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/sites?search='My Site'"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_GetOneSiteAnalytics
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "ec8c7d0f-c887-4318-8c0b-b2b88b12bc29"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/analytics"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_GetSitesFollowed
{
	# App Registration type:		Application
	# App Registration permissions: Files.Read.All, Files.ReadWrite.All, Sites.Read.All,	
	#								Sites.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/followedSites"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_SitesUnfollow
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$WebId = "cfaf31a6-f73f-4d7d-af24-6e530c022b5c"
	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/followedSites/remove"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_SitesFollow()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$WebId = "cfaf31a6-f73f-4d7d-af24-6e530c022b5c"
	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/followedSites/add"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_GetAllListsInSite
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_GetOneListInSite
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListTitle = "Documents"   #Use the DisplayName or the ListId
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListTitle
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_CreateList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ 'displayName' = 'List Created with GraphApi'; 
				 'list' = @{ 'template' = 'genericList' 
				}} | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 010 

#gavdcodebegin 018
function PsSpGraphRest_UpdateList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "List Created with GraphApi"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = '{  "__metadata": {
					"type": "SP.List"
				  },
				  "Description": "Description updated"
				}'

	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 018 

#gavdcodebegin 019
function PsSpGraphRest_DeleteList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "List Created with GraphApi"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ } | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 019

#gavdcodebegin 020
function PsSpGraphRest_GetAllColumnsInLists
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + 
								$ListId + "/columns"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 020

#gavdcodebegin 021
function PsSpGraphRest_GetOneColumnInList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$ColumnId = "Title"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + 
								$ListId + "/columns/" + $ColumnId
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 021

#gavdcodebegin 022
function PsSpGraphRest_CreateColumnInList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + 
														$ListId + "/columns"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName

	$myBody = @{  'description' = 'My Column Description';
				  'enforceUniqueValues' = 'false';
				  'hidden' = 'false';
				  'indexed' = 'false';
				  'name' = 'MyTextColumn';
				  'text' = @{
					'allowMultipleLines' = 'false';
					'appendChangesToExistingText' = 'false';
					'linesForEditing' = '0';
					'maxLength' = '255'
				  }} | ConvertTo-Json

	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 022

#gavdcodebegin 023
function PsSpGraphRest_UpdateColumn
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$ColumnId = "MyTextColumn"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{  'description' = 'My Column Updated'
				  } | ConvertTo-Json

	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 023

#gavdcodebegin 024
function PsSpGraphRest_DeleteColumn
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$ColumnId = "MyTextColumn"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
													"/columns/" + $ColumnId

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ } | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 024

#gavdcodebegin 025
function PsSpGraphRest_GetAllContentTypesInLists
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + 
								$ListId + "/contentTypes"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 025

#gavdcodebegin 026
function PsSpGraphRest_GetOneContentTypeInList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$ContentTypeId = "0x01010088C1A9D197313E45A4B9DD5AC6447A05"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + 
								$ListId + "/contentTypes/" + $ContentTypeId
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 026

#gavdcodebegin 027
function PsSpGraphRest_CreateContentTypeInList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + 
														$ListId + "/contentTypes"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{  'name' = 'myContentType';
				  'description' = 'My custom ContentType';
				  'base' = @{
					'name' = 'Document';
					'id' = '0x010101'
				  };
				  'group' = 'Document Content Types'
				} | ConvertTo-Json

	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 027

#gavdcodebegin 028
function PsSpGraphRest_UpdateContentType
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$ContentTypeId = "MyContentType"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
													"/contentTypes/" + $ContentTypeId

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{  'description' = 'My ContentType Updated'
				  } | ConvertTo-Json

	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 028

#gavdcodebegin 029
function PsSpGraphRest_DeleteContentType
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "Documents"
	$ContentTypeId = "MyContentType"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
													"/contentType/" + $ContentTypeId

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ } | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 029

#gavdcodebegin 011
function PsSpGraphRest_GetAllItemsInList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "TestList"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
																				"/items"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_GetOneItemInList
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "TestList"
	$ItemId = "11"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
																 "/items/" + $ItemId
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_GetOneItemAnalytics
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "TestList"
	$ItemId = "11"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
													"/items/" + $ItemId + "/analytics"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_CreateItem
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "TestList"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
																 "/items"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_UpdateItem
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "TestList"
	$ItemId = "1"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
													"/items/" + $ItemId + "/fields"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
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
function PsSpGraphRest_DeleteItem
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "3d93e562-aeb0-4316-a2b1-914aff04ad1a"
	$ListId = "TestList"
	$ItemId = "1"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $ListId + 
													"/items/" + $ItemId

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ } | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 015 

#gavdcodebegin 030
function PsSpGraphRest_GetDocumentsDriveInSite
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drive"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 030 

#gavdcodebegin 031
function PsSpGraphRest_GetDrivesInSite
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drives"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 031

#gavdcodebegin 032
function PsSpGraphRest_UploadFileToDocumentsLibrary
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DocPath = "C:\Temporary\TestDocument.docx"
	$DocName = "TestDocument.docx"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drive/root:/" + `
												$DocName + ":/content"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = Get-Content $DocPath -Raw

	$myContentType = "application/octet-stream"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Put `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 032

#gavdcodebegin 033
function PsSpGraphRest_UploadFileToLibrary
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$DocPath = "C:\Temporary\TestDocument.docx"
	$DocName = "TestDocument.docx"
	$Url = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + `
								$DocName + ":/content"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = Get-Content $DocPath -Raw

	$myContentType = "application/octet-stream"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Put `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 033

#gavdcodebegin 034
function PsSpGraphRest_DownloadFileFromDocumentsLibraryInSite
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DocName = "TestDocument.docx"
	$DownloadPath = "C:\Temporary"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drive/root:/" + `
													   $DocName + ":/content"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -OutFile $DownloadPath
	
	Write-Host $myResult
}
#gavdcodeend 034

#gavdcodebegin 035
function PsSpGraphRest_DownloadFileFromDriveInSite
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$DocName = "TestDocument.docx"
	$DownloadPath = "C:\Temporary"
	$Url = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + `
								$DocName + ":/content"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -OutFile $DownloadPath
	
	Write-Host $myResult
}
#gavdcodeend 035

#gavdcodebegin 036
function PsSpGraphRest_CheckOutFileInDocumentsLibrary
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$ItemId = "b699ae4a-7fcb-4aab-91ef-b4745123c25f"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drives/" + `
								$DriveId + "/items/" + $ItemId + "/checkout"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ } | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 036 

#gavdcodebegin 037
function PsSpGraphRest_CheckInFileInDocumentsLibrary
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$ItemId = "b699ae4a-7fcb-4aab-91ef-b4745123c25f"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drives/" + `
								$DriveId + "/items/" + $ItemId + "/checkin"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ } | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 037

#gavdcodebegin 038
function PsSpGraphRest_CreateFolderInLibrary
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drives/" + `
													   $DriveId + "/root/children"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ 'name' = 'NewFolderGraphApi'; 
				 'folder' = { };
				 '@microsoft.graph.conflictBehavior' = 'rename'
				} | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 038

#gavdcodebegin 039
function PsSpGraphRest_GetFoldersInLibrary
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drives/" + `
													   $DriveId + "/root/children"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 039

#gavdcodebegin 040
function PsSpGraphRest_GetFileVersions
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$ItemId = "b699ae4a-7fcb-4aab-91ef-b4745123c25f"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drives/" + `
													   $DriveId + "/items/" + `
													   $ItemId + "/versions"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 040

#gavdcodebegin 041
function PsSpGraphRest_GetFilePermissions
{
	# App Registration type:		Application
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$ItemId = "b699ae4a-7fcb-4aab-91ef-b4745123c25f"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drives/" + `
													   $DriveId + "/items/" + `
													   $ItemId + "/permissions"
	
	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$myObject = ConvertFrom-Json –InputObject $myResult
	$myObject.value.subject
}
#gavdcodeend 041

#gavdcodebegin 042
function PsSpGraphRest_AddPermissionsToFile
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$ItemId = "b699ae4a-7fcb-4aab-91ef-b4745123c25f"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drives/" + `
													   $DriveId + "/items/" + `
													   $ItemId + "/invite"

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName

	$myBody = @{ 'recipients' = @(@{
					  'email' = 'adelev@guitacadev.onmicrosoft.com'
					});
				 'message' = 'Sending invite for file';
				 'requireSignIn' = 'true';
				 'sendInvitation' = 'true';
				 'roles' = @( 'write' );
				 'password' = 'password123';
				 'expirationDateTime' = '2024-12-31T59:00:00.000Z'
				} | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 042

#gavdcodebegin 043
function PsSpGraphRest_DeletePermissionsFromFile
{
	# App Registration type:		Application
	# App Registration permissions: Sites.FullControl.All

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$DriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$ItemId = "b699ae4a-7fcb-4aab-91ef-b4745123c25f"
	$PermissionId = "VGVzdF9HdWl0YWNhIFZpc2l0b3Jz"
	$Url = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/drives/" + `
													   $DriveId + "/items/" + `
													   $ItemId + "/permissions" + `
													   $PermissionId

	$myOAuth = PsSpGraphRest_GetAzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = @{ } | ConvertTo-Json
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 043 


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 043 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$ClientIDApp = $configFile.appsettings.ClientIdWithSecret
$ClientSecretApp = $configFile.appsettings.ClientSecret
$TenantName = $configFile.appsettings.TenantName
$UserName = $configFile.appsettings.UserName

#PsSpGraphRest_GetAllSites 
#PsSpGraphRest_GetOneSiteById
#PsSpGraphRest_GetOneSiteByPath
#PsSpGraphRest_GetSitesBySearch
#PsSpGraphRest_GetOneSiteAnalytics
#PsSpGraphRest_GetSitesFollowed
#PsSpGraphRest_SitesUnfollow
#PsSpGraphRest_SitesFollow
#PsSpGraphRest_GetAllListsInSite
#PsSpGraphRest_GetOneListInSite
#PsSpGraphRest_CreateList
#PsSpGraphRest_UpdateList
#PsSpGraphRest_DeleteList
#PsSpGraphRest_GetAllColumnsInLists
#PsSpGraphRest_GetOneColumnInList
#PsSpGraphRest_CreateColumnInList
#PsSpGraphRest_UpdateColumn
#PsSpGraphRest_DeleteColumn
#PsSpGraphRest_GetAllContentTypesInLists
#PsSpGraphRest_GetOneContentTypeInList
#PsSpGraphRest_CreateContentTypeInList
#PsSpGraphRest_UpdateContentType
#PsSpGraphRest_DeleteContentType
#PsSpGraphRest_GetAllItemsInList
#PsSpGraphRest_GetOneItemInList
#PsSpGraphRest_GetOneItemAnalytics
#PsSpGraphRest_CreateItem
#PsSpGraphRest_UpdateItem
#PsSpGraphRest_DeleteItem
#PsSpGraphRest_GetDocumentsDriveInSite
#PsSpGraphRest_GetDrivesInSite
#PsSpGraphRest_UploadFileToDocumentsLibrary
#PsSpGraphRest_UploadFileToLibrary
#PsSpGraphRest_DownloadFileFromDocumentsLibraryInSite
#PsSpGraphRest_DownloadFileFromDriveInSite
#PsSpGraphRest_CheckOutFileInDocumentsLibrary
#PsSpGraphRest_CheckInFileInDocumentsLibrary
#PsSpGraphRest_CreateFolderInLibrary
#PsSpGraphRest_GetFoldersInLibrary
#PsSpGraphRest_GetFileVersions
#PsSpGraphRest_GetFilePermissions
#PsSpGraphRest_AddPermissionsToFile
#PsSpGraphRest_DeletePermissionsFromFile

Write-Host "Done" 
