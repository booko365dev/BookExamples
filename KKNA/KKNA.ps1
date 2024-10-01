
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


#*** Getting the Azure token with REST --------------------------------------------------
#gavdcodebegin 001
function PsRest_GetAzureTokenWithSecret
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
#gavdcodeend 001 

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 002
function PsSpGraphRest_ActivateSpEmbeddedContainer 
{
    $containerId = "b!A7pQX90BaEWOOSirXHOIrs_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
												$containerId + "/activate"
    
    $myBody = "{  }"
    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Post


    Write-Host $myResult
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpGraphRest_GetAllSpEmbeddedContainers 
{
    $containerTypeId = "86d347fa-234c-4e53-973e-a1d0100d6fe4"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers?" + `
                                "filter=(containerTypeId eq " + $containerTypeId + ")"
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpGraphRest_GetOneSpEmbeddedContainer
{
    $containerId = "b!A7pQX90BaEWOOSirXHOIrs_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
                                                        $containerId
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 004

#gavdcodebegin 023
function PsSpGraphRest_GetDriveSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
													$containerId + "/drive"
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 023

#gavdcodebegin 005
function PsSpGraphRest_CreateSpEmbeddedContainer 
{
    $containerTypeId = "86d347fa-234c-4e53-973e-a1d0100d6fe4"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers"
    
    $myBody = "{  'displayName': 'My Test Storage Container', `
                  'description': 'It is a Test Storage Container', `
                  'containerTypeId': '" + $containerTypeId + "' `
               }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Post

    Write-Host $myResult
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpGraphRest_UpdateSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
														$containerId
    
    $myBody = "{  'displayName': 'My Test Storage Container Updated', `
                  'description': 'It is a Test Storage Container Updated' `
               }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Patch

    Write-Host $myResult
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpGraphRest_DeleteSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
														$containerId
    
    $myBody = "{  }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Delete

    Write-Host $myResult
}
#gavdcodeend 007

#gavdcodebegin 024
function PsSpGraphRest_RestoreSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    #IMPORTANT: The EntryPoint is still in Beta version [2024-10]
    $Url = "https://graph.microsoft.com/beta/deletedStorageContainers/" + `
							$containerId + "/restore"

    $myBody = "{  }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Post

    Write-Host $myResult
}
#gavdcodeend 024

#gavdcodebegin 008
function PsSpGraphRest_DeletePermanentlySpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
													$containerId + "/permanentDelete"
    
    $myBody = "{  }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Post

    Write-Host $myResult
}
#gavdcodeend 008

#gavdcodebegin 025
function PsSpGraphRest_DeleteFromRecycleSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    #IMPORTANT: The EntryPoint is still in Beta version [2024-10]
    $Url = "https://graph.microsoft.com/beta/deletedStorageContainers/" + `
							$containerId

    $myBody = "{  }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Delete

    Write-Host $myResult
}
#gavdcodeend 025

#gavdcodebegin 009
function PsSpGraphRest_GetAllCustomPropertiesSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
													$containerId + "/customProperties"
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpGraphRest_GetOneCustomPropertieSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
    $customPropertyName = "MyCustomProperty"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
							$containerId + "/customProperties/" + $customPropertyName
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpGraphRest_CreateCustomPropertySpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
							$containerId + "/customProperties/"
    
    $myBody = "{  'MyCustomProperty': { `
                        'value': 'The value of the Property' `
                  } `
               }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Patch

    Write-Host $myResult
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpGraphRest_UpdateCustomPropertySpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
							$containerId + "/customProperties/"
    
    $myBody = "{  'MyCustomProperty': { `
                        'value': 'The value of the Property Updated' `
                  } `
               }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Patch

    Write-Host $myResult
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpGraphRest_DeleteCustomPropertySpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
							$containerId + "/customProperties/"
    
    $myBody = "{  'MyCustomProperty': null `
               }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Patch

    Write-Host $myResult
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpGraphRest_GetAllColumnsSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    #IMPORTANT: The EntryPoint is still in Beta version [2024-10]
    $Url = "https://graph.microsoft.com/beta/storage/fileStorage/containers/" + `
													$containerId + "/columns"
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpGraphRest_GetOneColumnSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
    $columnId = "fa564e0f-0c70-4ab9-b863-0177e6ddd247"

    #IMPORTANT: The EntryPoint is still in Beta version [2024-10]
    $Url = "https://graph.microsoft.com/beta/storage/fileStorage/containers/" + `
							$containerId + "/columns/" + $columnId
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 015

#gavdcodebegin 016
function PsSpGraphRest_CreateColumnSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    #IMPORTANT: The EntryPoint is still in Beta version [2024-10]
    $Url = "https://graph.microsoft.com/beta/storage/fileStorage/containers/" + `
							$containerId + "/columns/"
    
    $myBody = "{  'description': 'It is my Column', `
                  'displayName': 'MyColumn', `
                  'enforceUniqueValues': false, `
                  'hidden': false, `
                  'indexed': false, `
                  'name': 'MyColumn', `
                  'text': { `
                    'allowMultipleLines': true, `
                    'appendChangesToExistingText': false, `
                    'linesForEditing': 0, `
                    'maxLength': 255 `
                  } `
               }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Post

    Write-Host $myResult
}
#gavdcodeend 016

#gavdcodebegin 017
function PsSpGraphRest_UpdateColumnSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
    $columnId = "558c4b10-b64c-4334-9a07-4ea28ed3385b"

    #IMPORTANT: The EntryPoint is still in Beta version [2024-10]
    $Url = "https://graph.microsoft.com/beta/storage/fileStorage/containers/" + `
							$containerId + "/columns/" + $columnId
    
    $myBody = "{  'description': 'It is my Column Updated', `
                  'displayName': 'MyColumnUpdated' `
               }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Patch

    Write-Host $myResult
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpGraphRest_DeleteColumnSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
    $columnId = "558c4b10-b64c-4334-9a07-4ea28ed3385b"

    #IMPORTANT: The EntryPoint is still in Beta version [2024-10]
    $Url = "https://graph.microsoft.com/beta/storage/fileStorage/containers/" + `
							$containerId + "/columns/" + $columnId
    
    $myBody = "{  }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Delete

    Write-Host $myResult
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpGraphRest_GetAllPermissionsSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
													$containerId + "/permissions"
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 019

#gavdcodebegin 020
function PsSpGraphRest_AddPermissionSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
							$containerId + "/permissions/"
    
    $myBody = "{  'roles': ['reader'], `
                  'grantedToV2': { `
                    'user': { `
                      'userPrincipalName': 'user@domain.onmicrosoft.com' `
                    } `
                  } `
               }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Post

    Write-Host $myResult
}
#gavdcodeend 020

#gavdcodebegin 021
function PsSpGraphRest_UpdatePermissionSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
    $permissionId = "X2k6MCMuZnxtZW1iZXJzaGlwfGFkZWxldkBndWl0YWNhZGV2Lm9ubWljcm9zb2Z0LmNvbQ"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
							$containerId + "/permissions/" + $permissionId
    
    $myBody = "{  'roles': ['owner'] `
               }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Patch

    Write-Host $myResult
}
#gavdcodeend 021

#gavdcodebegin 022
function PsSpGraphRest_DeletePermissionSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
    $permissionId = "X2k6MCMuZnxtZW1iZXJzaGlwfGFkZWxldkBndWl0YWNhZGV2Lm9ubWljcm9zb2Z0LmNvbQ"

    $Url = "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + `
							$containerId + "/permissions/" + $permissionId
    
    $myBody = "{  }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Delete

    Write-Host $myResult
}
#gavdcodeend 022

#gavdcodebegin 026
function PsSpGraphRest_UploadFileToSpEmbeddedContainer
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
	$DocPath = "C:\Temporary\TestWordFile.docx"
	$DocName = "TestWordFile.docx"

	$Url = "https://graph.microsoft.com/v1.0/drives/" + $ContainerId + "/root:/" + `
								$DocName + ":/content"
	
    $myBody = Get-Content -Path $DocPath -AsByteStream -Raw

	$myContentType = "application/octet-stream"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Put `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 026

#gavdcodebegin 027
function PsSpGraphRest_GetAllFilesInSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"

    $Url = "https://graph.microsoft.com/v1.0/drives/" + $containerId + `
                                            "/items/root/children"
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 027

#gavdcodebegin 028
function PsSpGraphRest_DownloadFileFromSpEmbeddedContainer
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
    $itemId = "01VKOT7Y7WUIZ3IX2KJVDLUSALD27TOLHZ"
	$DocName = "TestWordFile01.docx"
	$DownloadPath = "C:\Temporary"

	$Url = "https://graph.microsoft.com/v1.0/drives/" + $containerId + "/items/" + `
								                        $itemId
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url 

    $myResultJson = $myResult.Content | ConvertFrom-Json
    $downloadUrl = $myResultJson."@microsoft.graph.downloadUrl"

    # Download the file using the download URL given by the previous request
    Invoke-WebRequest -Uri $downloadUrl -OutFile ($DownloadPath + "\" + $DocName)

	Write-Host $myResult
}
#gavdcodeend 028

#gavdcodebegin 029
function PsSpGraphRest_GetAllMetadataItemSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
    $itemId = "01VKOT7Y2KGNUBEU3LEJGLAQ53JUJSV5AS"

    $Url = "https://graph.microsoft.com/v1.0/drives/" + $containerId + `
                                            "/items/" + $itemId + "/listitem/fields"
    
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
    
    Write-Host $myResult
}
#gavdcodeend 029

#gavdcodebegin 030
function PsSpGraphRest_DeleteFileFromSpEmbeddedContainer 
{
    $containerId = "b!xQSmRvOosUCUOhlgdNyyk8_2HzkZmN9Gonxt4GLB1FspsW2Ful26RqiYjT8Qtn88"
    $itemId = "01VKOT7Y7P4O2NV3JZU5B2WW65NIJTSYCS"

    $Url = "https://graph.microsoft.com/v1.0/drives/" + `
							$containerId + "/items/" + $itemId
    
    $myBody = "{  }"

    $myContentType = "application/json" 
    $myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" } 
    $myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
                                  -ContentType $myContentType -Method Delete

    Write-Host $myResult
}
#gavdcodeend 030


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 030 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

# Get the token from Entra
$myOAuth = PsRest_GetAzureTokenWithSecret `
                -ClientID "3d16c7bc-a14e-454e-81cd-da571cf2a8e3" `
                -ClientSecret "IIX8Q~M5q-FCotlzO7GA4bLAQBPtF9dHwr.CcbsN" `
                -TenantName "ade56059-89c0-4594-90c3-e4772a8168ca"

#PsSpGraphRest_ActivateSpEmbeddedContainer
#PsSpGraphRest_GetAllSpEmbeddedContainers
#PsSpGraphRest_GetOneSpEmbeddedContainer
#PsSpGraphRest_CreateSpEmbeddedContainer
#PsSpGraphRest_UpdateSpEmbeddedContainer
#PsSpGraphRest_DeleteSpEmbeddedContainer
#PsSpGraphRest_RestoreSpEmbeddedContainer
#PsSpGraphRest_DeleteFromRecycleSpEmbeddedContainer
#PsSpGraphRest_DeletePermanentlySpEmbeddedContainer
#PsSpGraphRest_GetAllCustomPropertiesSpEmbeddedContainer
#PsSpGraphRest_GetOneCustomPropertieSpEmbeddedContainer
#PsSpGraphRest_CreateCustomPropertySpEmbeddedContainer
#PsSpGraphRest_UpdateCustomPropertySpEmbeddedContainer
#PsSpGraphRest_DeleteCustomPropertySpEmbeddedContainer
#PsSpGraphRest_GetAllColumnsSpEmbeddedContainer
#PsSpGraphRest_GetOneColumnSpEmbeddedContainer
#PsSpGraphRest_CreateColumnSpEmbeddedContainer
#PsSpGraphRest_UpdateColumnSpEmbeddedContainer
#PsSpGraphRest_DeleteColumnSpEmbeddedContainer
#PsSpGraphRest_GetAllPermissionsSpEmbeddedContainer
#PsSpGraphRest_AddPermissionSpEmbeddedContainer
#PsSpGraphRest_UpdatePermissionSpEmbeddedContainer
#PsSpGraphRest_DeletePermissionSpEmbeddedContainer
#PsSpGraphRest_GetDriveSpEmbeddedContainer
#PsSpGraphRest_UploadFileToSpEmbeddedContainer
#PsSpGraphRest_GetAllFilesInSpEmbeddedContainer
#PsSpGraphRest_DownloadFileFromSpEmbeddedContainer
#PsSpGraphRest_GetAllMetadataItemSpEmbeddedContainer
#PsSpGraphRest_DeleteFileFromSpEmbeddedContainer

Write-Host "Done" 
