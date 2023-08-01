
Function Get-AzureTokenApplication{
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

Function Get-AzureTokenDelegation{
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
Function ExchangePsGraph_GetAllMessages
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/messages"
	
	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 001 

#gavdcodebegin 002
Function ExchangePsGraph_GetAllMessagesMe
{
	# App Registration type:		Delegate
	# App Registration permissions: Mail.ReadBasic

	$Url = "https://graph.microsoft.com/v1.0/me/messages"

	# To use "me" you need to use Delegation rights	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw

	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 002 

#gavdcodebegin 003
Function ExchangePsGraph_CreateMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/messages"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
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
Function ExchangePsGraph_GetOneMessageById
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
				 "cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEPAAC1vtBLB-" + `
				 "F9SJ2ZDb7Xo-OrAADcqoD-AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/messages/" + $messageId
	
	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 004 

#gavdcodebegin 005
Function ExchangePsGraph_GetOneMessageByText
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$messageText = "This is a test mail"

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/messages" + `
							"`?`$search=`"" + $messageText + "`""
	
	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 005 

#gavdcodebegin 006
Function ExchangePsGraph_UpdateMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
				 "cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEPAAC1vtBLB-" + `
				 "F9SJ2ZDb7Xo-OrAADcqoD-AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/messages/" + $messageId

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'subject':'Test message updated by Graph Application' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 006 

#gavdcodebegin 007
Function ExchangePsGraph_DeleteMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
				 "cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEPAAC1vtBLB-" + `
				 "F9SJ2ZDb7Xo-OrAADcqoD-AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/messages/" + $messageId

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 007 

#gavdcodebegin 008
Function ExchangePsGraph_SendMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Send

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/sendMail"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
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
Function ExchangePsGraph_CopyMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
				 "cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEMAAC1vtBLB-" + `
				 "F9SJ2ZDb7Xo-OrAACiEhw0AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName+ "/messages/" + `
							$messageId + "/copy"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'destinationId': 'drafts' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 009 

#gavdcodebegin 010
Function ExchangePsGraph_MoveMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
				 "cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEMAAC1vtBLB-" + `
				 "F9SJ2ZDb7Xo-OrAACiEhw0AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/messages/" + `
							$messageId + "/move"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'destinationId': 'junkemail' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 010 

#gavdcodebegin 011
Function ExchangePsGraph_ReplyMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Send

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
				 "cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEMAAC1vtBLB-" + `
				 "F9SJ2ZDb7Xo-OrAACiEhw0AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/messages/" + `
							$messageId + "/reply"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'comment':'Email received' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 011 

#gavdcodebegin 012
Function ExchangePsGraph_ForwardMessage
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Send

	$messageId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQBGAAAAAAD" + `
				 "cxoIkHT46T678SPCidFpEBwC1vtBLB-F9SJ2ZDb7Xo-OrAAAAAAEMAAC1vtBLB-" + `
				 "F9SJ2ZDb7Xo-OrAACiEhw0AAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/messages/" + `
							$messageId + "/forward"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
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
Function ExchangePsGraph_CreateOverride
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + `
							"/inferenceClassification/overrides"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
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
Function ExchangePsGraph_GetAllOverrides
{
	# App Registration type:		Application
	# App Registration permissions: Mail.Read

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + `
							"/inferenceClassification/overrides"
	
	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.subject
}
#gavdcodeend 014 

#gavdcodebegin 015
Function ExchangePsGraph_UpdateOverride
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$overrideId = "70b8bfc8-a564-48cd-847a-6e84b5804235"

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + `
							"/inferenceClassification/overrides/" + $overrideId

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'classifyAs': 'focused' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 015 

#gavdcodebegin 016
Function ExchangePsGraph_DeleteOverride
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$overrideId = "70b8bfc8-a564-48cd-847a-6e84b5804235"

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + `
							"/inferenceClassification/overrides/" + $overrideId

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 016 

#gavdcodebegin 017
Function ExchangePsGraph_GetAllFolders
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read, Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders"
	
	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult

	$mailObject = ConvertFrom-Json –InputObject $myResult
	$mailObject.value.displayname
}
#gavdcodeend 017 

#gavdcodebegin 018
Function ExchangePsGraph_CreateFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'displayName': 'MyFolder' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 018 

#gavdcodebegin 019
Function ExchangePsGraph_GetOneFolderById
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
				"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAADcqmldAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders/" + `
							$folderId
	
	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 019 

#gavdcodebegin 020
Function ExchangePsGraph_CreateChildFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
				"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAADcqmldAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders/" + `
							$folderId + "/childFolders"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'displayName': 'MyChildFolder' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 020 

#gavdcodebegin 021
Function ExchangePsGraph_GetChildFoldersInOneFolderById
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadBasic, Mail.Read

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
				"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAADcqmldAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders/" + `
							$folderId + "/childFolders"
	
	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 021 

#gavdcodebegin 022
Function ExchangePsGraph_CreateMessageInFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
				"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAADcqmldAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders/" + `
							$folderId + "/messages"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
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
Function ExchangePsGraph_UpdateFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
				"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAADcqmldAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders/" + `
							$folderId

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'displayName': 'MyFolderUpdated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 023 

#gavdcodebegin 024
Function ExchangePsGraph_CopyFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
				"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAADcqmldAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders/" + `
							$folderId + "/copy"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'destinationId': 'drafts' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 024 

#gavdcodebegin 025
Function ExchangePsGraph_MoveFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
				"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAADcqmldAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders/" + `
							$folderId + "/move"

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myBody = "{ 'destinationId': 'junkemail' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
}
#gavdcodeend 025 

#gavdcodebegin 026
Function ExchangePsGraph_DeleteFolder
{
	# App Registration type:		Application
	# App Registration permissions: Mail.ReadWrite

	$folderId = "AAMkAGE0ODQ3NTc1LTZkM2ItNDk5Ny1iZDlkLTM5ODUxNWJkYmIwZQAuAAAAAAD" + `
				"cxoIkHT46T678SPCidFpEAQC1vtBLB-F9SJ2ZDb7Xo-OrAADcqmldAAA="

	$Url = "https://graph.microsoft.com/v1.0/users/" + `
							$configFile.appsettings.UserName + "/mailFolders/" + `
							$folderId

	$myOAuth = Get-AzureTokenApplication `
							-ClientID $configFile.appsettings.ClientIdWithSecret `
							-ClientSecret $configFile.appsettings.ClientSecret `
							-TenantName $configFile.appsettings.TenantName
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 026 

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

#ExchangePsGraph_GetAllMessages
#ExchangePsGraph_GetAllMessagesMe
#ExchangePsGraph_CreateMessage
#ExchangePsGraph_GetOneMessageById
#ExchangePsGraph_GetOneMessageByText
#ExchangePsGraph_UpdateMessage
#ExchangePsGraph_DeleteMessage
#ExchangePsGraph_SendMessage
#ExchangePsGraph_CopyMessage
#ExchangePsGraph_MoveMessage
#ExchangePsGraph_ReplyMessage
#ExchangePsGraph_ForwardMessage
#ExchangePsGraph_CreateOverride
#ExchangePsGraph_GetAllOverrides
#ExchangePsGraph_UpdateOverride
#ExchangePsGraph_DeleteOverride
#ExchangePsGraph_GetAllFolders
#ExchangePsGraph_CreateFolder
#ExchangePsGraph_GetOneFolderById
#ExchangePsGraph_GetOneFolderByText
#ExchangePsGraph_CreateChildFolder
#ExchangePsGraph_GetChildFoldersInOneFolderById
#ExchangePsGraph_CreateMessageInFolder
#ExchangePsGraph_UpdateFolder
#ExchangePsGraph_CopyFolder
#ExchangePsGraph_MoveFolder
#ExchangePsGraph_DeleteFolder

Write-Host "Done" 
