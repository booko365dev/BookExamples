
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
Function GrPsGetAllMessages()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 01 

#gavdcodebegin 02
Function GrPsGetAllMessagesMe()
{
	# App Registration type:		Delegate
	# App Registration permissions: Mail.ReadBasic

	$Url = "https://graph.microsoft.com/v1.0/me/messages"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 02 

#gavdcodebegin 03
Function GrPsCreateMessage()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'subject':'Test message created by Graph Application', `
			     'body':{ 'contentType':'Text', `
						  'content':'This is a test mail' }, `
			     'toRecipients':[{ 'emailAddress':{ 'address':'gustavo@gavd.net' } }] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 03 

#gavdcodebegin 04
Function GrPsGetOneMessageById()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$messageId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwBGAAAAAA" + `
		"CU5EKYHcFZTYUMWJ21mOXMBwC9tJp4R0ZpSbxvhS3KDF4kAAAAAAEPAAC9tJp4R0ZpSbxvhS3K" + `
		"DF4kAAA9xPPpAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages/" + `
																			$messageId
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 04 

#gavdcodebegin 05
Function GrPsGetOneMessageByText()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$messageText = "This is a test mail"

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages" + `
													"`?`$search=`"" + $messageText + "`""
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 05 

#gavdcodebegin 06
Function GrPsUpdateMessage()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwBGAAAAAA" + `
		"CU5EKYHcFZTYUMWJ21mOXMBwC9tJp4R0ZpSbxvhS3KDF4kAAAAAAEPAAC9tJp4R0ZpSbxvhS3K" + `
		"DF4kAAA9xPPpAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages/" + `
																			$messageId

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'subject':'Test message updated by Graph Application' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 06 

#gavdcodebegin 07
Function GrPsDeleteMessage()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwBGAAAAAA" + `
		"CU5EKYHcFZTYUMWJ21mOXMBwC9tJp4R0ZpSbxvhS3KDF4kAAAAAAEPAAC9tJp4R0ZpSbxvhS3K" + `
		"DF4kAAA9xPPpAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages/" + `
																			$messageId

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 07 

#gavdcodebegin 08
Function GrPsSendMessage()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Send

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/sendMail"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'message':{ 'subject':'Test message sent by Graph Application', `
			     'body':{ 'contentType':'Text', `
			     'content':'This is a test mail' }, `
			     'toRecipients':[{ 'emailAddress':{ 'address':'gustavo@gavd.net' } }]} }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 08 

#gavdcodebegin 09
Function GrPsCopyMessage()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwBGAAAAAACU5" + 
		"EKYHcFZTYUMWJ21mOXMBwC9tJp4R0ZpSbxvhS3KDF4kAAAAAAEMAAC9tJp4R0ZpSbx" + 
		"vhS3KDF4kAAA9xQtjAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages/" + `
																	$messageId + "/copy"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'destinationId': 'drafts' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 09 

#gavdcodebegin 10
Function GrPsMoveMessage()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwBGAAAAAACU5" + 
		"EKYHcFZTYUMWJ21mOXMBwC9tJp4R0ZpSbxvhS3KDF4kAAAAAAEMAAC9tJp4R0ZpSbx" + 
		"vhS3KDF4kAAA9xQtjAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages/" + `
																	$messageId + "/move"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'destinationId': 'junkemail' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 10 

#gavdcodebegin 11
Function GrPsReplyMessage()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Send

	$messageId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwBGAAAAAACU5" + 
		"EKYHcFZTYUMWJ21mOXMBwC9tJp4R0ZpSbxvhS3KDF4kAAAAAAEMAAC9tJp4R0ZpSbx" + 
		"vhS3KDF4kAAA9xQtkAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages/" + `
																	$messageId + "/reply"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'comment':'Email received' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 11 

#gavdcodebegin 12
Function GrPsForwardMessage()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Send

	$messageId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwBGAAAAAACU5" + 
		"EKYHcFZTYUMWJ21mOXMBwC9tJp4R0ZpSbxvhS3KDF4kAAAAAAEMAAC9tJp4R0ZpSbx" + 
		"vhS3KDF4kAAA9xQtkAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/messages/" + `
																	$messageId + "/forward"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'comment': 'Please review this email', `
				 'toRecipients': [{ `
						'emailAddress': { 'name': 'gustavo', `
										  'address': 'gustavo@gavd.net' }}] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 12 

#gavdcodebegin 13
Function GrPsCreateOverride()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + `
												"/inferenceClassification/overrides"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'classifyAs': 'other', `
				 'senderEmailAddress': { `
						'name': 'mysender', `
						'address': 'mysender@domain.com' }}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 13 

#gavdcodebegin 14
Function GrPsGetAllOverrides()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Read

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + `
												"/inferenceClassification/overrides"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 14 

#gavdcodebegin 15
Function GrPsUpdateOverride()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$overrideId = "70b8bfc8-a564-48cd-847a-6e84b5804235"

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + `
									"/inferenceClassification/overrides/" + $overrideId

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'classifyAs': 'focused' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 15 

#gavdcodebegin 16
Function GrPsDeleteOverride()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$overrideId = "70b8bfc8-a564-48cd-847a-6e84b5804235"

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + `
									"/inferenceClassification/overrides/" + $overrideId

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 16 

#gavdcodebegin 17
Function GrPsGetAllFolders()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.displayname
}
#gavdcodeend 17 

#gavdcodebegin 18
Function GrPsCreateFolder()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'displayName': 'MyFolder' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 18 

#gavdcodebegin 19
Function GrPsGetOneFolderById()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$folderId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwAuAAAAAACU" + 
							"5EKYHcFZTYUMWJ21mOXMAQC9tJp4R0ZpSbxvhS3KDF4kAAA9xNaaAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders/" + `
																			$folderId
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 19 

#gavdcodebegin 20
Function GrPsCreateChildFolder()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwAuAAAAAACU" + 
							"5EKYHcFZTYUMWJ21mOXMAQC9tJp4R0ZpSbxvhS3KDF4kAAA9xNaaAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders/" + `
															$folderId + "/childFolders"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'displayName': 'MyChildFolder' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 20 

#gavdcodebegin 21
Function GrPsGetChildFoldersInOneFolderById()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$folderId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwAuAAAAAACU" + 
							"5EKYHcFZTYUMWJ21mOXMAQC9tJp4R0ZpSbxvhS3KDF4kAAA9xNaaAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders/" + `
															$folderId + "/childFolders"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 21 

#gavdcodebegin 22
Function GrPsCreateMessageInFolder()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwAuAAAAAACU" + 
							"5EKYHcFZTYUMWJ21mOXMAQC9tJp4R0ZpSbxvhS3KDF4kAAA9xNaaAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders/" + `
															$folderId + "/messages"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'subject':'Test message created by Graph Application', `
			     'body':{ 'contentType':'Text', `
						  'content':'This is a test mail in a folder' }, `
			     'toRecipients':[{ 'emailAddress':{ 'address':'gustavo@gavd.net' } }] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 22 

#gavdcodebegin 23
Function GrPsUpdateFolder()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwAuAAAAAACU" + 
							"5EKYHcFZTYUMWJ21mOXMAQC9tJp4R0ZpSbxvhS3KDF4kAAA9xNaaAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders/" + `
																			$folderId

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'displayName': 'MyFolderUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 23 

#gavdcodebegin 24
Function GrPsCopyFolder()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwAuAAAAAACU" + 
							"5EKYHcFZTYUMWJ21mOXMAQC9tJp4R0ZpSbxvhS3KDF4kAAA9xNaaAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders/" + `
																	$folderId + "/copy"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'destinationId': 'drafts' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 24 

#gavdcodebegin 25
Function GrPsMoveFolder()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwAuAAAAAACU" + 
							"5EKYHcFZTYUMWJ21mOXMAQC9tJp4R0ZpSbxvhS3KDF4kAAA9xNaaAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders/" + `
																	$folderId + "/move"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myBody = "{ 'destinationId': 'junkemail' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 25 

#gavdcodebegin 26
Function GrPsDeleteFolder()
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkADE3MTIzYWNhLTNhZWItNDJkNC1iMGViLWZlYTI4ODUxNjRiYwAuAAAAAACU" + 
							"5EKYHcFZTYUMWJ21mOXMAQC9tJp4R0ZpSbxvhS3KDF4kAAA9xNaaAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + $userName + "/mailFolders/" + `
																			$folderId

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										 -ClientSecret $ClientSecretApp `
										 -TenantName $TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 26 

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\grPs.values.config"

$ClientIDApp = $configFile.appsettings.ClientIdApp
$ClientSecretApp = $configFile.appsettings.ClientSecretApp
$ClientIDDel = $configFile.appsettings.ClientIdDel
$TenantName = $configFile.appsettings.TenantName
$UserName = $configFile.appsettings.UserName
$UserPw = $configFile.appsettings.UserPw

#GrPsGetAllMessages
#GrPsGetAllMessagesMe
#GrPsCreateMessage
#GrPsGetOneMessageById
#GrPsGetOneMessageByText
#GrPsUpdateMessage
#GrPsDeleteMessage
#GrPsSendMessage
#GrPsCopyMessage
#GrPsMoveMessage
#GrPsReplyMessage
#GrPsForwardMessage
#GrPsCreateOverride
#GrPsGetAllOverrides
#GrPsUpdateOverride
#GrPsDeleteOverride
#GrPsGetAllFolders
#GrPsCreateFolder
#GrPsGetOneFolderById
#GrPsGetOneFolderByText
#GrPsCreateChildFolder
#GrPsGetChildFoldersInOneFolderById
#GrPsCreateMessageInFolder
#GrPsUpdateFolder
#GrPsCopyFolder
#GrPsMoveFolder
#GrPsDeleteFolder

##### ***** Quitar "gustavo@gavd.net" de todas partes  ******

Write-Host "Done" 
