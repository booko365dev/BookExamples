
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


#*** Getting the Azure token with REST ---------------------------------------------------

#gavdcodebegin 007
Function Get-AzureTokenWithAccPw
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
#gavdcodeend 007 
 
#gavdcodebegin 001
Function Get-AzureTokenWithSecret
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

#gavdcodebegin 022
Function Get-AzureTokenWithCertificate
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateThumbprint
	)

	$Scope = "https://graph.microsoft.com/.default"
	$myCertificatePath = "Cert:\CurrentUser\My\" + $CertificateThumbprint
	$myCertificate = Get-Item $myCertificatePath

	$CertificateBase64Hash = 
						[System.Convert]::ToBase64String($myCertificate.GetCertHash())

	$StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()
	$JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate `
						-End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
	$JWTExpiration = [math]::Round($JWTExpirationTimeSpan,0)

	$NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate `
									-End ((Get-Date).ToUniversalTime())).TotalSeconds
	$NotBefore = [math]::Round($NotBeforeExpirationTimeSpan,0)

	$JWTHeader = @{
		alg = "RS256"
		typ = "JWT"
		x5t = $CertificateBase64Hash -replace '\+','-' -replace '/','_' -replace '='
	}

	$JWTPayLoad = @{
		aud = "https://login.microsoftonline.com/$TenantName/oauth2/token"
		exp = $JWTExpiration
		iss = $ClientID
		jti = [guid]::NewGuid()
		nbf = $NotBefore
		sub = $ClientID
	}

	$JWTHeaderToByte = [System.Text.Encoding]::`
										UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
	$EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

	$JWTPayLoadToByte =  [System.Text.Encoding]::`
										UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
	$EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

	$JWT = $EncodedHeader + "." + $EncodedPayload

	$PrivateKey = $myCertificate.PrivateKey

	$RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
	$HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

	$Signature = [Convert]::ToBase64String(
		$PrivateKey.SignData([System.Text.Encoding]::`
										UTF8.GetBytes($JWT),$HashAlgorithm,$RSAPadding)
	) -replace '\+','-' -replace '/','_' -replace '='

	$JWT = $JWT + "." + $Signature

	$myBody = @{
		client_id = $ClientID
		client_assertion = $JWT
		client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
		scope = $Scope
		grant_type = "client_credentials"
	}

	$myUrl = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"

	$myHeader = @{
		Authorization = "Bearer $JWT"
	}

	$myOAuth = Invoke-RestMethod `
					-Method Post `
					-ContentType "application/x-www-form-urlencoded" `
					-Uri $myUrl `
					-Headers $myHeader `
					-Body $myBody

	return $myOAuth
}
#gavdcodeend 022 

#*** Using Classic PowerShell cmdlets and REST -------------------------------------------

#gavdcodebegin 002
Function PsGrClassical_GetTeam
{
	$Url = "https://graph.microsoft.com/v1.0/teams/dd1223a2-28a7-47d4-afc2-f42eae94f037"

	# Requires Delegated rights for Team.ReadBasic.All
	$myOAuth = Get-AzureTokenWithAccPw `
					-ClientID $configFile.appsettings.ClientIdWithAccPw `
					-TenantName $configFile.appsettings.TenantName `
					-UserName $configFile.appsettings.UserName `
					-UserPw $configFile.appsettings.UserPw

	<#
	# Requires Application rights for Team.ReadBasic.All
	$myOAuth = Get-AzureTokenWithSecret `
					-ClientID $configFile.appsettings.ClientIdWithSecret `
					-ClientSecret $configFile.appsettings.ClientSecret `
					-TenantName $configFile.appsettings.TenantName

	# Requires Application rights for Team.ReadBasic.All
	$myOAuth = Get-AzureTokenWithCertificate `
					-ClientID $configFile.appsettings.ClientIdWithCert `
					-TenantName $configFile.appsettings.TenantName `
					-CertificateThumbprint $configFile.appsettings.CertificateThumbprint
	#>

	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult
}
#gavdcodeend 002 

#gavdcodebegin 003
Function PsGrClassical_CreateChannel
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c" + 
							"/channels"
	

	# Requires Delegated rights for Channel.Create
	$myOAuth = Get-AzureTokenWithAccPw `
					-ClientID $configFile.appsettings.ClientIdWithAccPw `
									   -TenantName $configFile.appsettings.TenantName `
									   -UserName $configFile.appsettings.UserName `
									   -UserPw $configFile.appsettings.UserPw
	<#
	# Requires Application rights for Channel.Create
	$myOAuth = Get-AzureTokenWithSecret `
					-ClientID $configFile.appsettings.ClientIdWithSecret `
					-ClientSecret $configFile.appsettings.ClientSecret `
					-TenantName $configFile.appsettings.TenantName
    
	# Requires Application rights for Channel.Create
	$myOAuth = Get-AzureTokenWithCertificate `
					-ClientID $configFile.appsettings.ClientIdWithCert `
					-TenantName $configFile.appsettings.TenantName `
					-CertificateThumbprint $configFile.appsettings.CertificateThumbprint
	#>

	$myBody = "{ 'displayName':'Graph Channel 27', `
                 'description':'Channel created with Graph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
											-ContentType $myContentType -Method Post

	Write-Host $myResult
}
#gavdcodeend 003 

#gavdcodebegin 004
Function PsGrClassical_GetChannel
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c" +
							"/channels/19:012cd6295faa400db7aa1b869150feb0@thread.tacv2"
	

	# Requires Delegated rights for ChannelSettings.Read.All
	$myOAuth = Get-AzureTokenWithAccPw `
					-ClientID $configFile.appsettings.ClientIdWithAccPw `
					-TenantName $configFile.appsettings.TenantName `
					-UserName $configFile.appsettings.UserName `
					-UserPw $configFile.appsettings.UserPw
	<#
	# Requires Application rights for ChannelSettings.Read.All
	$myOAuth = Get-AzureTokenWithSecret `
					-ClientID $configFile.appsettings.ClientIdWithSecret `
					-ClientSecret $configFile.appsettings.ClientSecret `
					-TenantName $configFile.appsettings.TenantName
	
	# Requires Application rights for ChannelSettings.Read.All
	$myOAuth = Get-AzureTokenWithCertificate `
					-ClientID $configFile.appsettings.ClientIdWithCert `
					-TenantName $configFile.appsettings.TenantName `
					-CertificateThumbprint $configFile.appsettings.CertificateThumbprint
	#>

	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult
}
#gavdcodeend 004 

#gavdcodebegin 005
Function PsGrClassical_UpdateChannel
{
	$Url = 
		"https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c" +
							"/channels/19:012cd6295faa400db7aa1b869150feb0@thread.tacv2"


	# Requires Delegated rights for ChannelSettings.ReadWrite.All
	$myOAuth = Get-AzureTokenWithAccPw `
					-ClientID $configFile.appsettings.ClientIdWithAccPw `
					-TenantName $configFile.appsettings.TenantName `
					-UserName $configFile.appsettings.UserName `
					-UserPw $configFile.appsettings.UserPw
	<#
	# Requires Application rights for ChannelSettings.ReadWrite.All
	$myOAuth = Get-AzureTokenWithSecret `
					-ClientID $configFile.appsettings.ClientIdWithSecret `
					-ClientSecret $configFile.appsettings.ClientSecret `
					-TenantName $configFile.appsettings.TenantName
    
	# Requires Application rights for ChannelSettings.ReadWrite.All
	$myOAuth = Get-AzureTokenWithCertificate `
					-ClientID $configFile.appsettings.ClientIdWithCert `
					-TenantName $configFile.appsettings.TenantName `
					-CertificateThumbprint $configFile.appsettings.CertificateThumbprint
	#>

	$myBody = "{ 'description':'Channel Description Updated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)";
				   'IF-MATCH' = '*' }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
											-ContentType $myContentType -Method Patch

	Write-Host $myResult
}
#gavdcodeend 005

#gavdcodebegin 006
Function PsGrClassical_DeleteChannel
{
	$Url = 
		"https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c" + 
							"/channels/19:012cd6295faa400db7aa1b869150feb0@thread.tacv2"
	

	# Requires Delegated rights for Channel.Delete.All
	$myOAuth = Get-AzureTokenWithAccPw `
					-ClientID $configFile.appsettings.ClientIdWithAccPw `
					-TenantName $configFile.appsettings.TenantName `
					-UserName $configFile.appsettings.UserName `
									   -UserPw $configFile.appsettings.UserPw
	<#
	# Requires Application rights for Channel.Delete.All
	$myOAuth = Get-AzureTokenWithSecret `
					-ClientID $configFile.appsettings.ClientIdWithSecret `
					-ClientSecret $configFile.appsettings.ClientSecret `
					-TenantName $configFile.appsettings.TenantName
	
	# Requires Application rights for Channel.Delete.All
	$myOAuth = Get-AzureTokenWithCertificate `
					-ClientID $configFile.appsettings.ClientIdWithCert `
					-TenantName $configFile.appsettings.TenantName `
					-CertificateThumbprint $configFile.appsettings.CertificateThumbprint
	#>

	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}
#gavdcodeend 006

#*** Logging in using the Graph PowerShell SDK cmdlets -----------------------------------

#gavdcodebegin 008
Function PsGrSdk_LoginGraphWithInteraction
{
	Connect-Graph
}
#gavdcodeend 008

#gavdcodebegin 023
Function PsGrSdk_GetContextInfo
{
	Get-MgContext
}
#gavdcodeend 023

#gavdcodebegin 024
Function PsGrSdk_GetMe
{
	Get-MgUser -UserId "user@domain.onmicrosoft.com"
}
#gavdcodeend 024

#gavdcodebegin 025
Function PsGrSdk_ConnectDisconnect
{
	Connect-Graph -TenantId "021ee864-951d-4f25-a5c3-b6d4412c4052"
	Get-MgUser -UserId "user@domain.onmicrosoft.com"
	Disconnect-MgGraph
}
#gavdcodeend 025

#gavdcodebegin 031
Function PsGrSdk_CheckRights
{
	PsGrSdk_LoginGraphWithSecret
	(Get-MgContext).Scopes
	Disconnect-MgGraph
}
#gavdcodeend 031

#gavdcodebegin 026
Function PsGrSdk_SetVersion
{
	Select-MgProfile -Name "beta"
	Select-MgProfile -Name "v1.0"
}
#gavdcodeend 026

#gavdcodebegin 009
Function PsGrSdk_AssignRights
{
	Connect-Graph -Scopes "Directory.AccessAsUser.All, Directory.ReadWrite.All"
	Get-MgUser
	Disconnect-MgGraph
}
#gavdcodeend 009

#gavdcodebegin 032
Function PsGrSdk_CheckAvailableRights
{
	Find-MgGraphPermission "user" -PermissionType Application
}
#gavdcodeend 032


#gavdcodebegin 027
Function PsGrSdk_LoginGraphWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $UserName, $securePW

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -UserCredential $myCredentials 

	Connect-Graph -AccessToken $myToken.AccessToken
}
#gavdcodeend 027

#gavdcodebegin 028
Function PsGrSdk_GetUserWithAccPw
{
	# Requires Delegated rights for Directory.Read.All
	PsGrSdk_LoginGraphWithAccPw -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw
	Get-MgUser -UserId "user@domain.onmicrosoft.com"
	Disconnect-MgGraph
}
#gavdcodeend 028


#gavdcodebegin 029
Function PsGrSdk_LoginGraphWithSecret
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret
	)

	[SecureString]$secureSecret = ConvertTo-SecureString -String `
								$ClientSecret -AsPlainText -Force

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -ClientSecret ($secureSecret)

	Connect-Graph -AccessToken $myToken.AccessToken
}
#gavdcodeend 029

#gavdcodebegin 030
Function PsGrSdk_GetUsersWithSecret
{
	# Requires Application rights for Directory.Read.All
	PsGrSdk_LoginGraphWithSecret -TenantName $configFile.appsettings.TenantName `
								 -ClientID $configFile.appsettings.ClientIdWithSecret `
								 -ClientSecret $configFile.appsettings.ClientSecret
	Get-MgUser
	Disconnect-MgGraph
}
#gavdcodeend 030


#gavdcodebegin 033
Function PsGrSdk_LoginGraphWithCertificate
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateThumbprint
	)

	Connect-MgGraph -TenantId $TenantName `
					-ClientId $ClientId `
					-CertificateThumbprint $CertificateThumbprint
}
#gavdcodeend 033

#gavdcodebegin 034
Function PsGrSdk_LoginGraphWithCertificateFile
{
	[SecureString]$secureCertPw = ConvertTo-SecureString -String `
						$configFile.appSettings.CertificateFilePw -AsPlainText -Force

	$myCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(`
							 $configFile.appSettings.CertificateFilePath, $secureCertPw)
	
	Connect-MgGraph -TenantId $configFile.appsettings.TenantName `
					-ClientId $configFile.appsettings.ClientIdWithCert `
					-Certificate $myCert 
}
#gavdcodeend 034

#gavdcodebegin 035
Function PsGrSdk_GetUsersWithCertificate
{
	# Requires Application rights for Directory.Read.All
	PsGrSdk_LoginGraphWithCertificate `
					-TenantName $configFile.appsettings.TenantName `
					-ClientID $configFile.appsettings.ClientIdWithCert `
					-CertificateThumbprint $configFile.appsettings.CertificateThumbprint

	Get-MgUser -Property Id, DisplayName, BusinessPhones | `
										Format-Table Id, DisplayName, BusinessPhones

	Disconnect-MgGraph
}
#gavdcodeend 035


#gavdcodebegin 011
Function GrPsGetGroupsSelect #Not Used
{
	Get-MgGroup | Select-Object id, DisplayName, GroupTypes
}
#gavdcodeend 011

#*** Using MSAL.PS module to get the token -----------------------------------------------

#gavdcodebegin 036
Function PsMsal_LoginGraphWithInteraction
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID
	)

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -Interactive  # -Silent   # -ForceRefresh
							 #-RedirectUri "https://localhost"

	#Write-Host $myToken.AccessToken

	return $myToken
}
#gavdcodeend 036

#gavdcodebegin 037
Function PsMsal_LoginGraphWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
								-argumentlist $UserName, $securePW

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -UserCredential $myCredentials 
	
	#Write-Host $myToken.AccessToken

	return $myToken
}
#gavdcodeend 037

#gavdcodebegin 038
Function PsMsal_LoginGraphWithSecret
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret
	)

	[SecureString]$secureSecret = ConvertTo-SecureString -String `
								$ClientSecret -AsPlainText -Force

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -ClientSecret $secureSecret
	
	#Write-Host $myToken.AccessToken

	return $myToken
}
#gavdcodeend 038

#gavdcodebegin 039
Function PsMsal_LoginGraphWithCertificate
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateThumbprint
	)

	$myCertificatePath = "Cert:\CurrentUser\My\" + `
											$CertificateThumbprint
	$myCertificate = Get-Item $myCertificatePath

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -ClientCertificate $myCertificate
	
	#Write-Host $myToken.AccessToken

	return $myToken
}
#gavdcodeend 039

#gavdcodebegin 040
Function PsMsal_LoginGraph_GetTeamWithAccPw
{
	$Url = "https://graph.microsoft.com/v1.0/teams/dd1223a2-28a7-47d4-afc2-f42eae94f037"
	
	$myToken = GrPsLoginGraphMsalWithAccPw `
						-TenantName	$configFile.appsettings.TenantName `
						-ClientId $configFile.appsettings.ClientIdWithAccPw `
						-UserName $configFile.appsettings.UserName `
						-UserPw $configFile.appsettings.UserPw

	$myHeader = @{ 'Authorization' = "$($myToken.TokenType) $($myToken.AccessToken)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult
}
#gavdcodeend 040

#gavdcodebegin 041
Function PsMsal_LoginGraph_GetUsersWithSecret
{
	$myToken = PsMsal_LoginGraphWithSecret `
						-TenantName	$configFile.appsettings.TenantName `
						-ClientId $configFile.appsettings.ClientIdWithSecret `
						-ClientSecret $configFile.appsettings.ClientSecret

	Connect-Graph -AccessToken $myToken.AccessToken

	Get-MgUser

	Disconnect-MgGraph
}
#gavdcodeend 041

#*** Using PowerShell-MicrosoftGraphAPI module (Other Modules, not from MS) --------------

#gavdcodebegin 012
Function PsGrFrea_GetToken
{
	Import-Module .\MicrosoftGraph.psm1

	$myCredential = New-Object System.Management.Automation.PSCredential(`
		$myClientIdWithSecret,(ConvertTo-SecureString `
			$myClientSecret -AsPlainText -Force))
	$myToken = Get-MSGraphAuthToken -Credential $myCredential -TenantID $myTenantName

	Return $myToken
}
#gavdcodeend 012

#gavdcodebegin 013
Function PsGrFrea_GetTeamWithModule
{
	$myToken = PsGrFrea_GetToken
	Invoke-MSGraphQuery `
	  -URI "https://graph.microsoft.com/v1.0/teams/dd1223a2-28a7-47d4-afc2-f42eae94f037" `
	  -Token $myToken
}
#gavdcodeend 013

#gavdcodebegin 014
Function PsGrFrea_CreateChannelWithModule
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels"
	
	$myToken = PsGrFrea_GetToken
	$myBody = "{ 'displayName':'Graph Channel 40', `
				 'description':'Channel created with Graph' }"
	Invoke-MSGraphQuery `
		-URI $Url `
		-Body $myBody `
		-Token $myToken `
		-Meth Post
}
#gavdcodeend 014

#gavdcodebegin 015
Function PsGrFrea_UpdateChannelWithModule
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels/19:bb17af0c3a894262809c5412606f09f3@thread.tacv2"
	
	$myToken = PsGrFrea_GetToken
	$myBody = "{ 'description':'Channel Description Updated' }"
	Invoke-MSGraphQuery `
		-URI $Url `
		-Body $myBody `
		-Token $myToken `
		-Meth Patch
}
#gavdcodeend 015

#gavdcodebegin 016
Function PsGrFrea_DeleteChannelWithModule
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels/19:bb17af0c3a894262809c5412606f09f3@thread.tacv2"
	
	$myToken = PsGrFrea_GetToken
	$myBody = "{ 'description':'Channel Description Updated' }"
	Invoke-MSGraphQuery `
		-URI $Url `
		-Body $myBody `
		-Token $myToken `
		-Meth Delete
}
#gavdcodeend 016

#*** Using PnP Graph PowerShell ----------------------------------------------------------

#gavdcodebegin 042
Function PsGrPnp_LoginGraphWithInteraction
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$SiteBaseUrl
	)

	Connect-PnPOnline -Url $SiteBaseUrl -Interactive

	#Disconnect-PnPOnline
}
#gavdcodeend 042

#gavdcodebegin 043
Function PsGrPnp_LoginGraphWithInteractionMFA
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$SiteBaseUrl
	)

	Connect-PnPOnline -Url $SiteBaseUrl -DeviceLogin -LaunchBrowser

	#Disconnect-PnPOnline
}
#gavdcodeend 043

#gavdcodebegin 018
Function PsGrPnp_GetTeamUsersWithInteraction
{
	PsGrPnp_LoginGraphWithInteractionMFA -SiteBaseUrl $configFile.appsettings.SiteBaseUrl
	
	Get-PnPTeamsUser -Team "Design"

	Disconnect-PnPOnline
}
#gavdcodeend 018

#gavdcodebegin 021
Function PsGrPnp_GetToken
{
	Connect-PnPOnline -Url $configFile.appsettings.SiteBaseUrl -DeviceLogin -LaunchBrowser
	Get-PnPGraphAccessToken -Decoded

	Disconnect-PnPOnline
}
#gavdcodeend 021

#gavdcodebegin 020
Function PsGrPnp_LoginGraphWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$SiteUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)
	Write-Host $SiteUrl " - " $UserName " - " $UserPw
	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object System.Management.Automation.PSCredential `
								-argumentlist $UserName, $securePW

	Connect-PnPOnline -Url $SiteUrl -Credentials $myCredentials
}
#gavdcodeend 020

#gavdcodebegin 047
Function PsGrPnp_LoginGraphWithAccPwAndClientId
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$SiteBaseUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientId,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object System.Management.Automation.PSCredential `
								-argumentlist $UserName, $securePW

	Connect-PnPOnline -Url $SiteBaseUrl -ClientId $ClientId -Credentials $myCredentials
}
#gavdcodeend 047

#gavdcodebegin 044
Function PsGrPnp_GetContextWithAccPw
{
	PsGrPnp_LoginGraphWithAccPwAndClientId `
					-SiteBaseUrl $configFile.appsettings.SiteBaseUrl `
					-ClientId $configFile.appSettings.ClientIdWithAccPw `
					-UserName $configFile.appSettings.UserName `
					-UserPw $configFile.appSettings.UserPw
	
	Get-PnPContext

	Disconnect-PnPOnline
}
#gavdcodeend 044


#gavdcodebegin 045
Function PsGrPnp_LoginGraphWithSecret
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientId,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret	
	)

	$myToken = GrPsLoginGraphMsalWithSecret `
								-TenantName	$configFile.appsettings.TenantName `
								-ClientId $configFile.appsettings.ClientIdWithSecret `
								-ClientSecret $configFile.appsettings.ClientSecret

	Connect-PnPOnline -Url $configFile.appsettings.SiteBaseUrl `
					  -AccessToken $myToken.AccessToken

	# Does not work anymore
	#Connect-PnPOnline -Url $SiteBaseUrl -ClientId $ClientId -ClientSecret $ClientSecret
}
#gavdcodeend 045

#gavdcodebegin 046
Function PsGrPnp_GetTeamUsersWithSecret
{
	PsGrPnp_LoginGraphWithSecret `
					-TenantName	$configFile.appsettings.TenantName `
					-ClientId $configFile.appsettings.ClientIdWithSecret `
					-ClientSecret $configFile.appsettings.ClientSecret

	Get-PnPTeamsUser -Team "Design"

	Disconnect-PnPOnline
}
#gavdcodeend 046


#gavdcodebegin 048
Function PsGrPnp_LoginGraphWithCertificate
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$SiteBaseUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientId,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateThumbprint
	)

	Connect-PnPOnline -Url $SiteBaseUrl `
					  -Tenant $TenantName `
					  -ClientId $ClientId `
					  -Thumbprint $CertificateThumbprint
}
#gavdcodeend 048

#gavdcodebegin 049
Function PsGrPnp_LoginGraphWithCertificateFile
{
	[SecureString]$secureCertPw = ConvertTo-SecureString -String `
						$configFile.appSettings.CertificateFilePw -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteBaseUrl `
					  -Tenant $configFile.appsettings.TenantName `
					  -ClientId $configFile.appSettings.ClientIdWithCert `
					  -CertificatePath $configFile.appSettings.CertificateFilePath `
					  -CertificatePassword $certPw 
}
#gavdcodeend 049

#gavdcodebegin 050
Function PsGrPnp_GetLanguagesWithCertificate
{
	PsGrPnp_LoginGraphWithCertificate `
					-SiteBaseUrl $configFile.appsettings.SiteBaseUrl `
					-TenantName $configFile.appsettings.TenantName `
					-ClientId $configFile.appSettings.ClientIdWithCert `
					-CertificateThumbprint $configFile.appSettings.CertificateThumbprint
	
	Get-PnPAvailableLanguage

	Disconnect-PnPOnline
}
#gavdcodeend 050


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 050 ***

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

$myClientIdWithSecret = $configFile.appsettings.ClientIdWithSecret
$myClientSecret = $configFile.appsettings.ClientSecret
$myClientIdWithAccPw = $configFile.appsettings.ClientIdWithAccPw
$myTenantName = $configFile.appsettings.TenantName
$myTenantUrl = $configFile.appsettings.TenantUrl
$myClientIdWithCert = $configFile.appsettings.ClientIdWithCert
$myCertificateThumbprint = $configFile.appsettings.CertificateThumbprint
$myCertificateFilePath = $configFile.appSettings.CertificateFilePath
$myCertificateFilePw = $configFile.appSettings.CertificateFilePw
$myUserName = $configFile.appsettings.UserName
$myUserPw = $configFile.appsettings.UserPw
$mySiteCollUrl = $configFile.appsettings.SiteCollUrl
$mySiteBaseUrl = $configFile.appsettings.SiteBaseUrl

#*** Using Classic PowerShell cmdlets
#PsGrClassical_GetTeam
#PsGrClassical_CreateChannel
#PsGrClassical_GetChannel
#PsGrClassical_UpdateChannel
#PsGrClassical_DeleteChannel

#*** Using Microsoft Graph PowerShell SDK cmdlets
#PsGrSdk_LoginGraphWithInteraction
#PsGrSdk_GetContextInfo
#PsGrSdk_GetMe
#PsGrSdk_ConnectDisconnect
#PsGrSdk_CheckRights
#PsGrSdk_SetVersion
#PsGrSdk_AssignRights
#PsGrSdk_CheckAvailableRights

#PsGrSdk_LoginGraphWithAccPw $myTenantName $myClientIdWithAccPw $myUserName $myUserPw
#PsGrSdk_GetUserWithAccPw

#PsGrSdk_LoginGraphWithSecret $myTenantName $myClientIdWithSecret $myClientSecret
#PsGrSdk_GetUsersWithSecret

#PsGrSdk_LoginGraphWithCertificate $myTenantName $myClientIdWithCert $myCertificateThumbprint
#PsGrSdk_LoginGraphWithCertificateFile
#PsGrSdk_GetUsersWithCertificate

#GrPsGetGroupsSelect #Not Used

#*** Using MSAL.PS module to get the token
#PsMsal_LoginGraphWithInteraction $myTenantName $myClientIdWithAccPw
#PsMsal_LoginGraphWithAccPw $myTenantName $myClientIdWithAccPw $myUserName $myUserPw
#PsMsal_LoginGraphWithSecret $myTenantName $myClientIdWithSecret $myClientSecret
#PsMsal_LoginGraphWithCertificate $myTenantName $myClientIdWithCert $myCertificateThumbprint
#PsMsal_LoginGraph_GetTeamWithAccPw
#PsMsal_LoginGraph_GetUsersWithSecret

#*** Using PowerShell-MicrosoftGraphAPI module (Other modules, not MS)
#PsGrFrea_GetToken
#PsGrFrea_GetTeamWithModule
#PsGrFrea_CreateChannelWithModule
#PsGrFrea_UpdateChannelWithModule
#PsGrFrea_DeleteChannelWithModule

#*** Using PnP Graph PowerShell
#PsGrPnp_LoginGraphWithInteraction $mySiteBaseUrl
#PsGrPnp_LoginGraphWithInteractionMFA $mySiteBaseUrl
#PsGrPnp_GetTeamUsersWithInteraction
#PsGrPnp_GetToken

#PsGrPnp_LoginGraphWithAccPw $mySiteCollUrl $myUserName $myUserPw
#PsGrPnp_LoginGraphWithAccPwAndClientId $mySiteBaseUrl $myClientIdWithAccPw $myUserName $myUserPw
#PsGrPnp_GetContextWithAccPw

#PsGrPnp_LoginGraphWithSecret $myTenantName $myClientIdWithSecret $myClientSecret
#PsGrPnp_GetTeamUsersWithSecret

#PsGrPnp_LoginGraphWithCertificate $mySiteBaseUrl $myClientIdWithCert $myCertificateThumbprint
#PsGrPnp_LoginGraphWithCertificateFile
#PsGrPnp_GetLanguagesWithCertificate

Write-Host "Done" 
