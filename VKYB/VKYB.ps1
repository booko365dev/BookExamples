##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsGraphRestApi_GetAzureTokenApplication{
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

Function PsGraphRestApi_GetAzureTokenDelegation{
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
Function PsExchangeGraphRestApi_GetAllMessages
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/messages"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 001 

#gavdcodebegin 002
Function PsExchangeGraphRestApi_GetAllMessagesMe
{
	# App Registration type:		Delegate
	# App Registration permissions: Mail.ReadBasic

	$Url = "https://graph.microsoft.com/v1.0/me/messages"

	# To use "me" you need to use Delegation rights	
	$myOAuth = PsGraphRestApi_GetAzureTokenDelegation `
							-ClientID $cnfClientIdWithAccPw `
							-TenantName $cnfTenantName `
							-UserName $cnfUserName `
							-UserPw $cnfUserPw

	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 002 

#gavdcodebegin 003
Function PsExchangeGraphRestApi_CreateMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/messages"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'subject':'Test message created by Graph Application', `
			     'body':{ 'contentType':'Text', `
						  'content':'This is a test mail' }, `
			     'toRecipients':[{ 'emailAddress':{ 'address':'user@domain.com' } }] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 003 

#gavdcodebegin 004
Function PsExchangeGraphRestApi_GetOneMessageById
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + 
				 "cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEPAAC1vtBLB-F" + 
				 "9SJ2ZDb7Xo-OrAALATnwvAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/messages/" + $messageId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 004 

#gavdcodebegin 005
Function PsExchangeGraphRestApi_GetOneMessageByText
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$messageText = "This is a test mail"

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/messages" + `
							"`?`$search=`"" + $messageText + "`""
	
	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 005 

#gavdcodebegin 006
Function PsExchangeGraphRestApi_UpdateMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAA" + 
				 "AAAADcxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEP" + 
				 "AAC1vtBLB-F9SJ2ZDb7Xo-OrAALATnwvAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/messages/" + $messageId

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'subject':'Test message updated by Graph Application' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 006 

#gavdcodebegin 007
Function PsExchangeGraphRestApi_DeleteMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAA" + 
				 "AAAADcxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEP" + 
				 "AAC1vtBLB-F9SJ2ZDb7Xo-OrAALATnwvAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/messages/" + $messageId

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 007 

#gavdcodebegin 008
Function PsExchangeGraphRestApi_SendMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Send

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/sendMail"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'message':{ 'subject':'Test message sent by Graph Application', `
			     'body':{ 'contentType':'Text', `
			     'content':'This is a test mail' }, `
			     'toRecipients':[{ 'emailAddress':{ 'address':'user@domain.com' } }]} }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 008 

#gavdcodebegin 009
Function PsExchangeGraphRestApi_CopyMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAA" + 
				 "AAADcxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEMAA" + 
				 "C1vtBLB-F9SJ2ZDb7Xo-OrAALATWn7AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName+ "/messages/" + `
							$messageId + "/copy"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'destinationId': 'drafts' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 009 

#gavdcodebegin 010
Function PsExchangeGraphRestApi_MoveMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAA" + 
				 "AAADcxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEMAA" + 
				 "C1vtBLB-F9SJ2ZDb7Xo-OrAALATWn7AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/messages/" + `
							$messageId + "/move"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'destinationId': 'junkemail' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 010 

#gavdcodebegin 011
Function PsExchangeGraphRestApi_ReplyMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Send

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAA" + 
				 "AAADcxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEMAA" + 
				 "C1vtBLB-F9SJ2ZDb7Xo-OrAALATWn8AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/messages/" + `
							$messageId + "/reply"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'comment':'Email received' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 011 

#gavdcodebegin 012
Function PsExchangeGraphRestApi_ForwardMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Send

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAA" + 
				 "AAADcxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEMAA" + 
				 "C1vtBLB-F9SJ2ZDb7Xo-OrAALATWn8AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/messages/" + `
							$messageId + "/forward"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'comment': 'Please review this email', `
				 'toRecipients': [{ `
						'emailAddress': { 'name': 'user', `
										  'address': 'user@domain.com' }}] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 012 

#gavdcodebegin 013
Function PsExchangeGraphRestApi_CreateOverride
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + `
							"/inferenceClassification/overrides"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'classifyAs': 'other', `
				 'senderEmailAddress': { `
						'name': 'mysender', `
						'address': 'mysender@domain.com' }}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 013 

#gavdcodebegin 014
Function PsExchangeGraphRestApi_GetAllOverrides
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Read

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + `
							"/inferenceClassification/overrides"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 014 

#gavdcodebegin 015
Function PsExchangeGraphRestApi_UpdateOverride
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$overrideId = "70b8bfc8-a564-48cd-847a-6e84b5804235"

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + `
							"/inferenceClassification/overrides/" + $overrideId

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'classifyAs': 'focused' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 015 

#gavdcodebegin 016
Function PsExchangeGraphRestApi_DeleteOverride
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$overrideId = "70b8bfc8-a564-48cd-847a-6e84b5804235"

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + `
							"/inferenceClassification/overrides/" + $overrideId

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 016 

#gavdcodebegin 017
Function PsExchangeGraphRestApi_GetAllFolders
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.displayname
}
#gavdcodeend 017 

#gavdcodebegin 018
Function PsExchangeGraphRestApi_CreateFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'displayName': 'MyFolder01' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 018 

#gavdcodebegin 019
Function PsExchangeGraphRestApi_GetOneFolderById
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAA" + 
				"AADcxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAALATT7_AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders/" + `
							$folderId
	
	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 019 

#gavdcodebegin 020
Function PsExchangeGraphRestApi_CreateChildFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAA" + 
				"AADcxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAALATT7_AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders/" + `
							$folderId + "/childFolders"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'displayName': 'MyChildFolder' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 020 

#gavdcodebegin 021
Function PsExchangeGraphRestApi_GetChildFoldersInOneFolderById
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read
	
	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAA" + 
				"AADcxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAALATT7_AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders/" + `
							$folderId + "/childFolders"
	
	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 021 

#gavdcodebegin 022
Function PsExchangeGraphRestApi_CreateMessageInFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAA" + 
				"AADcxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAALATT7_AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders/" + `
							$folderId + "/messages"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'subject':'Test message created by Graph Application', `
			     'body':{ 'contentType':'Text', `
						  'content':'This is a test mail in a folder' }, `
			     'toRecipients':[{ 'emailAddress':{ 'address':'user@domain.com' } }] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 022 

#gavdcodebegin 023
Function PsExchangeGraphRestApi_UpdateFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAA" + 
				"AADcxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAALATT7_AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders/" + `
							$folderId

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'displayName': 'MyFolderUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 023 

#gavdcodebegin 024
Function PsExchangeGraphRestApi_CopyFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAA" + 
				"AADcxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAALATT7_AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders/" + `
							$folderId + "/copy"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'destinationId': 'drafts' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 024 

#gavdcodebegin 025
Function PsExchangeGraphRestApi_MoveFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAA" + 
				"AADcxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAALATT7_AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders/" + `
							$folderId + "/move"

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myBody = "{ 'destinationId': 'junkemail' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 025 

#gavdcodebegin 026
Function PsExchangeGraphRestApi_DeleteFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAA" + 
				"AADcxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAALATT7_AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$cnfUserName + "/mailFolders/" + `
							$folderId

	$myOAuth = PsGraphRestApi_GetAzureTokenApplication `
							-ClientID $cnfClientIdWithSecret `
							-ClientSecret $cnfClientSecret `
							-TenantName $cnfTenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 026 


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 026 ***

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

#PsExchangeGraphRestApi_GetAllMessages
#PsExchangeGraphRestApi_GetAllMessagesMe
#PsExchangeGraphRestApi_CreateMessage
#PsExchangeGraphRestApi_GetOneMessageById
#PsExchangeGraphRestApi_GetOneMessageByText
#PsExchangeGraphRestApi_UpdateMessage
#PsExchangeGraphRestApi_DeleteMessage
#PsExchangeGraphRestApi_SendMessage
#PsExchangeGraphRestApi_CopyMessage
#PsExchangeGraphRestApi_MoveMessage
#PsExchangeGraphRestApi_ReplyMessage
#PsExchangeGraphRestApi_ForwardMessage
#PsExchangeGraphRestApi_CreateOverride
#PsExchangeGraphRestApi_GetAllOverrides
#PsExchangeGraphRestApi_UpdateOverride
#PsExchangeGraphRestApi_DeleteOverride
#PsExchangeGraphRestApi_GetAllFolders
#PsExchangeGraphRestApi_CreateFolder
#PsExchangeGraphRestApi_GetOneFolderById
#PsExchangeGraphRestApi_GetOneFolderByText
#PsExchangeGraphRestApi_CreateChildFolder
#PsExchangeGraphRestApi_GetChildFoldersInOneFolderById
#PsExchangeGraphRestApi_CreateMessageInFolder
#PsExchangeGraphRestApi_UpdateFolder
#PsExchangeGraphRestApi_CopyFolder
#PsExchangeGraphRestApi_MoveFolder
#PsExchangeGraphRestApi_DeleteFolder

Write-Host "Done" 
