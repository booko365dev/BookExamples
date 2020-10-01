
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
Function GrPsCallGraphExcelSP01()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
											$itemId + "/driveitem/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " >> SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 01 

#gavdcodebegin 02
Function GrPsCallGraphExcelSP02a()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$tenantBaseUrl = "m365x895762.sharepoint.com"
	$siteCollName = "Test_Guitaca"

	# Get the Drive ID
	$Url = $graphBaseUrl + "/sites/" + $tenantBaseUrl + ":/sites/" + $siteCollName + ":/drives"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name - " $xlsxObject.value[0].name " >> Drive ID - " $xlsxObject.value[0].id
}
#gavdcodeend 02 

#gavdcodebegin 03
Function GrPsCallGraphExcelSP02b()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$tenantBaseUrl = "m365x895762.sharepoint.com"
	$driveId = "b!ur1VuWGXGUGRf2vMfESaWqYxr88_931NryRuUwwCK1y5nbpnSMD-RbeQda0yM1AQ"

	# Get the Document ID
	$Url = $graphBaseUrl + "/sites/" + $tenantBaseUrl + "/drives/" + $driveId `
																	+ "/root/children"  
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " >> DocumentID- " $xlsxObject.value[0].id
}
#gavdcodeend 03 

#gavdcodebegin 04
Function GrPsCallGraphExcelSP02c()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$tenantBaseUrl = "m365x895762.sharepoint.com"
	$driveId = "b!ur1VuWGXGUGRf2vMfESaWqYxr88_931NryRuUwwCK1y5nbpnSMD-RbeQda0yM1AQ"
	$itemId = "015XX4O2PIQCTK65N6QBDIFW2VWDDBPEIV"

	$Url = $graphBaseUrl + "/sites/" + $tenantBaseUrl + "/drives/" + $driveId + `
											"/items/" + $itemId + "/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " >> SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 04 

#gavdcodebegin 05
Function GrPsCallGraphExcelOD01()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$xlsxName = "TestBook.xlsx"

	$Url = $graphBaseUrl + "users/" + $userName + "/drive/root:/" + $xlsxName + `
																":/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " >> SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 05 

#gavdcodebegin 06
Function GrPsCallGraphExcelOD02()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$itemId = "01WE2M3BBMWEXOJ7XABZH3LO6YSVSXTLBH"

	# To find the spreadsheet ID ($itemId)
	#$Url = $graphBaseUrl + "users/" + $userName + "/drive/root/children/"		

	$Url = $graphBaseUrl + "users/" + $userName + "/drive/items/" + $itemId + `
																"/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " >> SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 06 

#gavdcodebegin 07
Function GrPsCallGraphExcelMeOD()
{
	# App Registration type:		Delegate
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$xlsxName = "TestBook.xlsx"

	$Url = $graphBaseUrl + "me/drive/root:/" + $xlsxName + ":/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	Write-Host "Name- " $xlsxObject.value[0].name " >> SheetID- " $xlsxObject.value[0].id
}
#gavdcodeend 07 

#gavdcodebegin 08
Function GrPsGetAllWorksheets()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
											$itemId + "/driveitem/workbook/worksheets"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	foreach($oneValue in $xlsxObject.value) {
		Write-Host "Name- " $oneValue.name " >> SheetID- " $oneValue.id
	}
}
#gavdcodeend 08 

#gavdcodebegin 09
Function GrPsGetOneWorksheetByName()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 09 

#gavdcodebegin 10
Function GrPsCreateWorksheet()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
											$itemId + "/driveitem/workbook/worksheets"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'name': 'PsSheet' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 10 

#gavdcodebegin 11
Function GrPsUpdateWorksheet()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'name': 'PsSheetUpdated', 'position': 1 }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 11 

#gavdcodebegin 12
Function GrPsDeleteWorksheet()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 12 

#gavdcodebegin 13
Function GrPsCallFunction()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
										$itemId + "/driveitem/workbook/functions/arabic"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'text' : 'MLVI' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 13 

#gavdcodebegin 14
Function GrPsGetAllComments()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
											$itemId + "/driveitem/workbook/comments"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	foreach($oneValue in $xlsxObject.value) {
		Write-Host "Name- " $oneValue.name " >> SheetID- " $oneValue.id
	}
}
#gavdcodeend 14 

#gavdcodebegin 15
Function GrPsGetAllReplaysOneComment()
{
	# App Registration type:		Application
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
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$xlsxObject = ConvertFrom-Json –InputObject $myResult
	foreach($oneValue in $xlsxObject.value) {
		Write-Host "Name- " $oneValue.name " >> SheetID- " $oneValue.id
	}
}
#gavdcodeend 15 

#gavdcodebegin 16
Function GrPsCreateReply()
{
	# App Registration type:		Application
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
#gavdcodeend 16 

#gavdcodebegin 17
Function GrPsInsertRange()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/range(address='B2:C3')/insert"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'shift': 'Right' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 17 

#gavdcodebegin 18
Function GrPsGetRange()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/range(address='B2:E7')"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 18 

#gavdcodebegin 19
Function GrPsUpdateRange()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/range(address='A1:B2')"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'values' : [['Excel', '456'],['1/1/2021', null]],
				 'formulas' : [[null, null], [null, '=B1*5']],
				 'numberFormat' : [[null,null], ['m-ddd', null]] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 19 

#gavdcodebegin 20
Function GrPsClearRange()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/range(address='B2:C3')/clear"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'applyTo': 'All' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 20 

#gavdcodebegin 21
Function GrPsDeleteRange()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/range(address='B2:C3')/delete"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'applyTo': 'All' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 21 

#gavdcodebegin 22
Function GrPsCreateNamedRange()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/names/add"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'name': 'GraphNamedRange',
			     'reference': '=B2:C3',
			     'comment': 'Named range set by Graph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 22 

#gavdcodebegin 23
Function GrPsGetAllNamedRanges()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/names"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 23 

#gavdcodebegin 24
Function GrPsGetOneNamedRange()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$namedRange = "GraphNamedRange"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/names/" + $namedRange
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 24 

#gavdcodebegin 25
Function GrPsUpdateNamedRange()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$namedRange = "GraphNamedRange"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/names/" + $namedRange

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'comment': 'Named Range Updated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 25 

#gavdcodebegin 26
Function GrPsCreateTable()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/add"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'address': 'F5:I9',
			     'hasHeaders': true }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 26 

#gavdcodebegin 27
Function GrPsGetAllTables()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 27 

#gavdcodebegin 28
Function GrPsUpdateTable()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "Table1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'name': 'GraphTable', 
				 'showTotals': false, 
				 'style': 'TableStyleMedium3' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 28 

#gavdcodebegin 29
Function GrPsGetTableRows()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName + "/rows"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 29 

#gavdcodebegin 30
Function GrPsGetTableColumns()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName + "/columns"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 30 

#gavdcodebegin 31
Function GrPsGetTableOneColumn()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = "3"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName + "/columns/" + $columnIndex
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 31 

#gavdcodebegin 32
Function GrPsCreateTableRow()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$rowIndex = 2

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName + "/rows/add"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'values': [ [11, 'ab', 22, 'cd'] ],
			     'index': " + $rowIndex + " }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 32 

#gavdcodebegin 33
Function GrPsCreateTableColumn()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 2

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName + "/columns/add"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'values': [ ['myCol'], [11], ['ab'], [22], ['cd'], [33], ['ef'], [44] ],
			     'index': " + $columnIndex + " }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 33 

#gavdcodebegin 34
Function GrPsDeleteColumn()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 3

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName + "/columns/" + $columnIndex

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}
#gavdcodeend 34 

#gavdcodebegin 35
Function GrPsTableSort()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 2

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName + "/sort/apply"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
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
#gavdcodeend 35 

#gavdcodebegin 36
Function GrPsTableFilter()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 1

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName + "/columns/" + $columnIndex + `
							"/filter/apply"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
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
#gavdcodeend 36 

#gavdcodebegin 37
Function GrPsTableClearFilter()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$tableName = "GraphTable"
	$columnIndex = 1

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/tables/" + $tableName + "/columns/" + $columnIndex + `
							"/filter/clear"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 37 

#gavdcodebegin 38
Function GrPsCreateChart()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/charts/add"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'type': 'ColumnStacked',
				 'sourceData': 'F5:G7',
				 'seriesBy': 'Auto' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 38 

#gavdcodebegin 39
Function GrPsGetAllCharts()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/charts"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 39 

#gavdcodebegin 40
Function GrPsGetChartImage()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$chartName = "Chart 1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/charts/" + $chartName + `
							"/Image(width=0,height=0,fittingMode='fit')"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 40 

#gavdcodebegin 41
Function GrPsUpdateChartSource()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$chartName = "Chart 1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/charts/" + $chartName + "/setData"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'type': 'ColumnStacked',
				 'sourceData': 'F5:G8',
				 'seriesBy': 'Auto' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 41 

#gavdcodebegin 42
Function GrPsUpdateChartProperties()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$chartName = "Chart 1"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/charts/" + $chartName

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'name': 'ChartGraph', 
				 'top': 10, 'left': 10, 
				 'height': 200.0, 'width': 300.0 }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 42 

#gavdcodebegin 43
Function GrPsDeleteChart()
{
	# App Registration type:		Application
	# App Registration permissions: Sites.ReadWrite.All, Files.ReadWrite.All

	$graphBaseUrl = "https://graph.microsoft.com/v1.0/"
	$siteId = "b955bdba-9761-4119-917f-6bcc7c449a5a"  # Site Collection ID
	$listId = "67ba9db9-c048-45fe-b790-75ad32335010"
	$itemId = "1"
	$worksheetName = "PsSheet"
	$chartName = "ChartGraph"

	$Url = $graphBaseUrl + "sites/" + $siteId + "/lists/" + $listId + "/items/" + `
							$itemId + "/driveitem/workbook/worksheets/" + $worksheetName + `
							"/charts/" + $chartName

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 43 

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\grPs.values.config"

$ClientIDApp = $configFile.appsettings.ClientIdApp
$ClientSecretApp = $configFile.appsettings.ClientSecretApp
$ClientIDDel = $configFile.appsettings.ClientIdDel
$TenantName = $configFile.appsettings.TenantName
$UserName = $configFile.appsettings.UserName
$UserPw = $configFile.appsettings.UserPw

#GrPsCallGraphExcelSP01
#GrPsCallGraphExcelSP02a
#GrPsCallGraphExcelSP02b
#GrPsCallGraphExcelSP02c
#GrPsCallGraphExcelOD01
#GrPsCallGraphExcelOD02
#GrPsCallGraphExcelMeOD
#GrPsGetAllWorksheets
#GrPsCreateWorksheet
#GrPsGetOneWorksheetByName
#GrPsUpdateWorksheet
#GrPsDeleteWorksheet
#GrPsCallFunction
#GrPsGetAllComments
#GrPsGetAllReplaysOneComment
#GrPsCreateReply
#GrPsInsertRange
#GrPsGetRange
#GrPsUpdateRange
#GrPsClearRange
#GrPsDeleteRange
#GrPsCreateNamedRange
#GrPsGetAllNamedRanges
#GrPsGetOneNamedRange
#GrPsUpdateNamedRange
#GrPsCreateTable
#GrPsGetAllTables
#GrPsUpdateTable
#GrPsGetTableRows
#GrPsGetTableColumns
#GrPsGetTableOneColumn
#GrPsCreateTableRow
#GrPsCreateTableColumn
#GrPsDeleteColumn
#GrPsTableSort
#GrPsTableFilter
#GrPsTableClearFilter
#GrPsCreateChart
#GrPsGetAllCharts
#GrPsGetChartImage
#GrPsUpdateChartSource
#GrPsUpdateChartProperties
#GrPsDeleteChart

Write-Host "Done" 
