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


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


#gavdcodebegin 001
Function ExcelPsGraph_GetSpreadsheetId
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
											$itemId + "/driveitem/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " > SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 001 

#gavdcodebegin 002
Function ExcelPsGraph_GetDriveId
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$tenantBaseUrl = "[tenant].sharepoint.com"
	$siteCollName = "Chapter03"

	# Get the Drive ID
	$Url = $graphBaseUrl + "/sites/" + $tenantBaseUrl + `
											":/sites/" + $siteCollName + ":/drives"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name - " $xlsxObject.value[0].name " > Drive ID - `
															" $xlsxObject.value[0].id
}
#gavdcodeend 002 

#gavdcodebegin 003
Function ExcelPsGraph_GetSpreadsheetByDriveId
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$tenantBaseUrl = "[tenant].sharepoint.com"
	$driveId = "b!0br1f8VWqEWObdlRrScqOvjdCbTLvTFDuDobRNveAdAa5_SrQGIvSqk_ezm-VTvq"

	# Get the Document ID
	$Url = $graphBaseUrl + "/sites/" + $tenantBaseUrl + "/drives/" + $driveId `
																	+ "/root/children"  
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " > DocumentID- " `
															$xlsxObject.value[0].id
}
#gavdcodeend 003 

#gavdcodebegin 004
Function ExcelPsGraph_GetSpreadsheetByDriveIdAndItemId
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$tenantBaseUrl = "[tenant].sharepoint.com"
	$driveId = "b!0br1f8VWqEWObdlRrScqOvjdCbTLvTFDuDobRNveAdAa5_SrQGIvSqk_ezm-VTvq"
	$itemId = "015XX4O2PIQCTK65N6QBDIFW2VWDDBPEIV"

	$Url = $graphBaseUrl + "/sites/" + $tenantBaseUrl + "/drives/" + $driveId + `
											"/items/" + $itemId + "/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " > SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 004 

#gavdcodebegin 005
Function ExcelPsGraph_GetSpreadsheetInOnedriveByName
{
	# App Registration type:		Delegated
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$xlsxName = "TestBook.xlsx"

	$Url = $graphBaseUrl + "users/" + $configFile.appsettings.UserName + `
								"/drive/root:/" + $xlsxName + ":/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " > SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 005 

#gavdcodebegin 006
Function ExcelPsGraph_GetSpreadsheetInOnedriveById
{
	# App Registration type:		Delegated
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$itemId = "01Y5GMQHD7LKQV6L5URRC32IH5XNQN34WP"

	# To find the spreadsheet ID ($itemId)
	#$Url = $graphBaseUrl + "users/" + $configFile.appsettings.UserName + `
	#														"/drive/root/children/"		

	$Url = $graphBaseUrl + "users/" + $configFile.appsettings.UserName + `
									"/drive/items/" + $itemId + "/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " > SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 006 

#gavdcodebegin 007
Function ExcelPsGraph_GetSpreadsheetInOnedriveByMe
{
	# App Registration type:		Delegate
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$xlsxName = "TestBook.xlsx"

	$Url = $graphBaseUrl + "me/drive/root:/" + $xlsxName + ":/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " > SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 007 

#gavdcodebegin 008
Function ExcelPsGraph_GetAllWorksheets
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
											$itemId + "/driveitem/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	foreach($oneValue in $xlsxObject.value) {
		Write-Host "Name- " $oneValue.name " >> SheetID- " $oneValue.id
	}
}
#gavdcodeend 008 

#gavdcodebegin 009
Function ExcelPsGraph_GetOneWorksheetByName
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName
	
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
Function ExcelPsGraph_CreateWorksheet
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
											$itemId + "/driveitem/workbook/worksheets"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'name': 'PowerShellSheet' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 010 

#gavdcodebegin 011
Function ExcelPsGraph_UpdateWorksheet
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'name': 'PsSheetUpdated', 'position': 1 }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 011 

#gavdcodebegin 012
Function ExcelPsGraph_DeleteWorksheet
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheetUpdated"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 012 

#gavdcodebegin 013
Function ExcelPsGraph_CallFunction
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
										$itemId + "/driveitem/workbook/functions/arabic"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'text' : 'MLVI' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 013 

#gavdcodebegin 014
Function ExcelPsGraph_GetAllComments
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
											$itemId + "/driveitem/workbook/comments"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	foreach($oneValue in $xlsxObject.value) {
		Write-Host "Name- " $oneValue.name " > SheetID- " $oneValue.id
	}
}
#gavdcodeend 014 

#gavdcodebegin 015
Function ExcelPsGraph_GetAllReplaysOneComment
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$commentId = "{BC2DBDF6-F028-4B56-A87A-C322B9BCA5B4}"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
					$itemId + "/driveitem/workbook/comments/" + $commentId + "/replies"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	foreach($oneValue in $xlsxObject.value) {
		Write-Host "Name- " $oneValue.name " > SheetID- " $oneValue.id
	}
}
#gavdcodeend 015 

#gavdcodebegin 016
Function GrPsCreateReply
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$commentId = "{11075FE6-6B52-4B18-8E4B-A4387CECBDD5}"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
					$itemId + "/driveitem/workbook/comments/" + $commentId + "/replies"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'content': 'Reply from PowerShell',
				 'contentType': 'plain' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 016 

#gavdcodebegin 017
Function ExcelPsGraph_InsertRange
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/range(address='B2:C3')/insert"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'shift': 'Right' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 017 

#gavdcodebegin 018
Function ExcelPsGraph_GetRange
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/range(address='B2:E7')"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 018 

#gavdcodebegin 019
Function ExcelPsGraph_UpdateRange
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/range(address='A1:B2')"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'values' : [['Excel', '456'],['1/1/2021', null]],
				 'formulas' : [[null, null], [null, '=B1*5']],
				 'numberFormat' : [[null,null], ['m-ddd', null]] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 019 

#gavdcodebegin 020
Function ExcelPsGraph_ClearRange
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/range(address='B2:C3')/clear"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'applyTo': 'All' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 020 

#gavdcodebegin 021
Function ExcelPsGraph_DeleteRange
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/range(address='B2:C3')/delete"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'applyTo': 'All' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 021 

#gavdcodebegin 022
Function ExcelPsGraph_CreateNamedRange
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/names/add"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'name': 'GraphNamedRange',
			     'reference': '=B2:C3',
			     'comment': 'Named range set by Graph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 022 

#gavdcodebegin 023
Function ExcelPsGraph_GetAllNamedRanges
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/names"
	
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
Function ExcelPsGraph_GetOneNamedRange
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$namedRange = "GraphNamedRange"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/names/" + $namedRange
	
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
Function ExcelPsGraph_UpdateNamedRange
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$namedRange = "GraphNamedRange"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/names/" + $namedRange

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'comment': 'Named Range Updated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 025 

#gavdcodebegin 026
Function ExcelPsGraph_CreateTable
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/add"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'address': 'F5:I9',
			     'hasHeaders': true }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 026 

#gavdcodebegin 027
Function ExcelPsGraph_GetAllTables
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables"
	
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
Function ExcelPsGraph_UpdateTable
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "Table1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + $tableName

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'name': 'GraphTable', 
				 'showTotals': false, 
				 'style': 'TableStyleMedium3' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 028 

#gavdcodebegin 029
Function ExcelPsGraph_GetTableRows
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + $tableName + "/rows"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 029 

#gavdcodebegin 030
Function ExcelPsGraph_GetTableColumns
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + $tableName + "/columns"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 030 

#gavdcodebegin 031
Function ExcelPsGraph_GetTableOneColumn
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = "3"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + `
							$tableName + "/columns/" + $columnIndex
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 031 

#gavdcodebegin 032
Function ExcelPsGraph_CreateTableRow
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$rowIndex = 2

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + $tableName + "/rows/add"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'values': [ [11, 'ab', 22, 'cd'] ],
			     'index': " + $rowIndex + " }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 032 

#gavdcodebegin 033
Function ExcelPsGraph_CreateTableColumn
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 2

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + $tableName + "/columns/add"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'values': [ ['myCol'], [11], ['ab'], [22], ['cd'], [33] ],
			     'index': " + $columnIndex + " }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 033 

#gavdcodebegin 034
Function ExcelPsGraph_DeleteColumn
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 3

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + `
							$tableName + "/columns/" + $columnIndex

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}
#gavdcodeend 034 

#gavdcodebegin 035
Function ExcelPsGraph_TableSort
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 2

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + $tableName + "/sort/apply"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'fields' : [
				  { 'key': 0,
				   'ascending': true
				  }
				] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 035 

#gavdcodebegin 036
Function ExcelPsGraph_TableFilter
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 1

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + `
							$tableName + "/columns/" + $columnIndex + "/filter/apply"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'criteria' : 
				  { 'filterOn': 'custom',
				    'criterion1': '>1',
				    'operator': 'and',
				    'criterion2': '<8'
				  } }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 036 

#gavdcodebegin 037
Function ExcelPsGraph_TableClearFilter
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 1

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/tables/" + `
							$tableName + "/columns/" + $columnIndex + "/filter/clear"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 037 

#gavdcodebegin 038
Function ExcelPsGraph_CreateChart
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/charts/add"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'type': 'ColumnStacked',
				 'sourceData': 'F5:G7',
				 'seriesBy': 'Auto' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 038 

#gavdcodebegin 039
Function ExcelPsGraph_GetAllCharts
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/charts"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 039 

#gavdcodebegin 040
Function ExcelPsGraph_GetAllChartsGetChartImage
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$chartName = "Chart 1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/charts/" + $chartName + `
							"/Image(width=0,height=0,fittingMode='fit')"
	
	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 040 

#gavdcodebegin 041
Function ExcelPsGraph_UpdateChartSource
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$chartName = "Chart 1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/charts/" + $chartName + "/setData"

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'type': 'ColumnStacked',
				 'sourceData': 'F5:G8',
				 'seriesBy': 'Auto' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 041 

#gavdcodebegin 042
Function ExcelPsGraph_UpdateChartProperties
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$chartName = "Chart 1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/charts/" + $chartName

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'name': 'ChartGraph', 
				 'top': 10, 'left': 10, 
				 'height': 200.0, 'width': 300.0 }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 042 

#gavdcodebegin 043
Function ExcelPsGraph_DeleteChart
{
	# App Registration type:		Delegated
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "7ff5bad1-56c5-45a8-8e6d-d951ad272a3a"  # Site Collection ID
	$listId = "ad7b2787-1494-4940-9692-a0080a105af0"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$chartName = "ChartGraph"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + `
							$worksheetName + "/charts/" + $chartName

	$myOAuth = Get-AzureTokenDelegation `
								-ClientID $configFile.appsettings.ClientIdWithAccPw `
								-TenantName $configFile.appsettings.TenantName `
								-UserName $configFile.appsettings.UserName `
								-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 043 


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

# *** Latest Source Code Index: 043 ***

#ExcelPsGraph_GetSpreadsheetId
#ExcelPsGraph_GetDriveId
#ExcelPsGraph_GetSpreadsheetByDriveId
#ExcelPsGraph_GetSpreadsheetByDriveIdAndItemId
#ExcelPsGraph_GetSpreadsheetInOnedriveByName
#ExcelPsGraph_GetSpreadsheetInOnedriveById
#ExcelPsGraph_GetSpreadsheetInOnedriveByMe
#ExcelPsGraph_GetAllWorksheets
#ExcelPsGraph_GetOneWorksheetByName
#ExcelPsGraph_CreateWorksheet
#ExcelPsGraph_UpdateWorksheet
#ExcelPsGraph_DeleteWorksheet
#ExcelPsGraph_CallFunction
#ExcelPsGraph_GetAllComments
#ExcelPsGraph_GetAllReplaysOneComment
#GrPsCreateReply ==> Not used in book
#ExcelPsGraph_InsertRange
#ExcelPsGraph_GetRange
#ExcelPsGraph_UpdateRange
#ExcelPsGraph_ClearRange
#ExcelPsGraph_DeleteRange
#ExcelPsGraph_CreateNamedRange
#ExcelPsGraph_GetAllNamedRanges
#ExcelPsGraph_GetOneNamedRange
#ExcelPsGraph_UpdateNamedRange
#ExcelPsGraph_CreateTable
#ExcelPsGraph_GetAllTables
#ExcelPsGraph_UpdateTable
#ExcelPsGraph_GetTableRows
#ExcelPsGraph_GetTableColumns
#ExcelPsGraph_GetTableOneColumn
#ExcelPsGraph_CreateTableRow
#ExcelPsGraph_CreateTableColumn
#ExcelPsGraph_DeleteColumn
#ExcelPsGraph_TableSort
#ExcelPsGraph_TableFilter
#ExcelPsGraph_TableClearFilter
#ExcelPsGraph_CreateChart
#ExcelPsGraph_GetAllCharts
#ExcelPsGraph_GetAllChartsGetChartImage
#ExcelPsGraph_UpdateChartSource
#ExcelPsGraph_UpdateChartProperties
#ExcelPsGraph_DeleteChart

Write-Host "Done" 
