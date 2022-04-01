
#*** Getting the Azure token with REST ---------------------------------------------------

#gavdcodebegin 07
Function Get-AzureTokenWithAccPw(){
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
#gavdcodeend 07 
 
#gavdcodebegin 01
Function Get-AzureTokenWithSecret(){
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
#gavdcodeend 01 

#gavdcodebegin 22
Function Get-AzureTokenWithCertificate(){
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

	$CertificateBase64Hash = [System.Convert]::ToBase64String($myCertificate.GetCertHash())

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
#gavdcodeend 22 

#*** Using Classic PowerShell cmdlets and REST -------------------------------------------

#gavdcodebegin 02
Function GrPsGetTeam()
{
	$Url = "https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c"

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
#gavdcodeend 02 

#gavdcodebegin 03
Function GrPsCreateChannel()
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
#gavdcodeend 03 

#gavdcodebegin 04
Function GrPsGetChannel()
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
#gavdcodeend 04 

#gavdcodebegin 05
Function GrPsUpdateChannel()
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
#gavdcodeend 05

#gavdcodebegin 06
Function GrPsDeleteChannel()
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
#gavdcodeend 06

#*** Logging in using the Graph PowerShell SDK cmdlets ---------------------------------------

#gavdcodebegin 08
Function GrPsLoginGraphSDKWithInteraction()
{
	Connect-Graph
}
#gavdcodeend 08

#gavdcodebegin 23
Function GrPsLoginGraphSDKGetContextInfo()
{
	Get-MgContext
}
#gavdcodeend 23

#gavdcodebegin 24
Function GrPsLoginGraphSDK_GetMe()
{
	Get-MgUser -UserId "user@domain.onmicrosoft.com"
}
#gavdcodeend 24

#gavdcodebegin 25
Function GrPsLoginGraphSDKConnectDisconnect()
{
	Connect-Graph -TenantId "021ee864-951d-4f25-a5c3-b6d4412c4052"
	Get-MgUser -UserId "user@domain.onmicrosoft.com"
	Disconnect-MgGraph
}
#gavdcodeend 25

#gavdcodebegin 26
Function GrPsLoginGraphSDKSetVersion()
{
	Select-MgProfile -Name "beta"
	Select-MgProfile -Name "v1.0"
}
#gavdcodeend 26

#gavdcodebegin 09
Function GrPsLoginGraphSDKAssignRights()
{
	Connect-Graph -Scopes "Directory.AccessAsUser.All, Directory.ReadWrite.All"
	Get-MgUser
	Disconnect-MgGraph
}
#gavdcodeend 09

#gavdcodebegin 27
Function GrPsLoginGraphSDKWithAccPw()
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
#gavdcodeend 27

#gavdcodebegin 28
Function GrPsLoginGraphSDK_GetUserWithAccPw()
{
	# Requires Delegated rights for Directory.Read.All
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw
	Get-MgUser -UserId "user@domain.onmicrosoft.com"
	Disconnect-MgGraph
}
#gavdcodeend 28

#gavdcodebegin 29
Function GrPsLoginGraphSDKWithSecret()
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
#gavdcodeend 29

#gavdcodebegin 30
Function GrPsLoginGraphSDK_GetUsersWithSecret()
{
	# Requires Application rights for Directory.Read.All
	GrPsLoginGraphSDKWithSecret -TenantName $configFile.appsettings.TenantName `
								-ClientID $configFile.appsettings.ClientIdWithSecret `
								-ClientSecret $configFile.appsettings.ClientSecret
	Get-MgUser
	Disconnect-MgGraph
}
#gavdcodeend 30

#gavdcodebegin 31
Function GrPsLoginGraphSDKCheckRights()
{
	GrPsLoginGraphSDKWithSecret
	(Get-MgContext).Scopes
	Disconnect-MgGraph
}
#gavdcodeend 31

#gavdcodebegin 32
Function GrPsLoginGraphSDKCheckAvailableRights()
{
	Find-MgGraphPermission "user" -PermissionType Application
}
#gavdcodeend 32

#gavdcodebegin 33
Function GrPsLoginGraphSDKWithCertificate()
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
#gavdcodeend 33

#gavdcodebegin 34
Function GrPsLoginGraphSDKWithCertificateFile()
{
	[SecureString]$secureCertPw = ConvertTo-SecureString -String `
							$configFile.appSettings.CertificateFilePw -AsPlainText -Force

	$myCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(`
							  $configFile.appSettings.CertificateFilePath, $secureCertPw)
	
	Connect-MgGraph -TenantId $configFile.appsettings.TenantName `
					-ClientId $configFile.appsettings.ClientIdWithCert `
					-Certificate $myCert 
}
#gavdcodeend 34

#gavdcodebegin 35
Function GrPsLoginGraphSDK_GetUsersWithCertificate()
{
	# Requires Application rights for Directory.Read.All
	GrPsLoginGraphSDKWithCertificate `
					-TenantName $configFile.appsettings.TenantName `
					-ClientID $configFile.appsettings.ClientIdWithCert `
					-CertificateThumbprint $configFile.appsettings.CertificateThumbprint

	Get-MgUser -Property Id, DisplayName, BusinessPhones | `
										Format-Table Id, DisplayName, BusinessPhones

	Disconnect-MgGraph
}
#gavdcodeend 35

#gavdcodebegin 11
Function GrPsGetGroupsSelect()
{
	Get-MgGroup | Select-Object id, DisplayName, GroupTypes
}
#gavdcodeend 11

#*** Using MSAL.PS module to get the token -----------------------------------------------

#gavdcodebegin 36
Function GrPsLoginGraphMsalWithInteraction()
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
#gavdcodeend 36

#gavdcodebegin 37
Function GrPsLoginGraphMsalWithAccPw()
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
#gavdcodeend 37

#gavdcodebegin 38
Function GrPsLoginGraphMsalWithSecret()
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
#gavdcodeend 38

#gavdcodebegin 39
Function GrPsLoginGraphMsalWithCertificate()
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
#gavdcodeend 39

#gavdcodebegin 40
Function GrPsLoginGraphMsal_GetTeamWithAccPw()
{
	$Url = "https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c"
	
	$myToken = GrPsLoginGraphMsalWithAccPw `
						-TenantName	$configFile.appsettings.TenantName `
						-ClientId $configFile.appsettings.ClientIdWithAccPw `
						-UserName $configFile.appsettings.UserName `
						-UserPw $configFile.appsettings.UserPw

	$myHeader = @{ 'Authorization' = "$($myToken.TokenType) $($myToken.AccessToken)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult
}
#gavdcodeend 40

#gavdcodebegin 41
Function GrPsLoginGraphMsal_GetUsersWithSecret()
{
	$myToken = GrPsLoginGraphMsalWithSecret `
						-TenantName	$configFile.appsettings.TenantName `
						-ClientId $configFile.appsettings.ClientIdWithSecret `
						-ClientSecret $configFile.appsettings.ClientSecret

	Connect-Graph -AccessToken $myToken.AccessToken

	Get-MgUser

	Disconnect-MgGraph
}
#gavdcodeend 41

#*** Using PowerShell-MicrosoftGraphAPI module (Other Modules, not from MS) --------------

#gavdcodebegin 12
Function GrPsGetToken()
{
	Import-Module .\MicrosoftGraph.psm1

	$myCredential = New-Object System.Management.Automation.PSCredential(`
			$ClientIDApp,(ConvertTo-SecureString $ClientSecretApp -AsPlainText -Force))
	$myToken = Get-MSGraphAuthToken -Credential $myCredential -TenantID $TenantID

	Return $myToken
}
#gavdcodeend 12

#gavdcodebegin 13
Function GrPsGetTeamWithModule()
{
	$myToken = GrPsGetToken
	Invoke-MSGraphQuery `
		-URI "https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" `
		-Token $myToken
}
#gavdcodeend 13

#gavdcodebegin 14
Function GrPsCreateChannelWithModule()
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels"
	
	$myToken = GrPsGetToken
	$myBody = "{ 'displayName':'Graph Channel 40', `
				 'description':'Channel created with Graph' }"
	Invoke-MSGraphQuery `
		-URI $Url `
		-Body $myBody `
		-Token $myToken `
		-Meth Post
}
#gavdcodeend 14

#gavdcodebegin 15
Function GrPsUpdateChannelWithModule()
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels/19:bb17af0c3a894262809c5412606f09f3@thread.tacv2"
	
	$myToken = GrPsGetToken
	$myBody = "{ 'description':'Channel Description Updated' }"
	Invoke-MSGraphQuery `
		-URI $Url `
		-Body $myBody `
		-Token $myToken `
		-Meth Patch
}
#gavdcodeend 15

#gavdcodebegin 16
Function GrPsDeleteChannelWithModule()
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels/19:bb17af0c3a894262809c5412606f09f3@thread.tacv2"
	
	$myToken = GrPsGetToken
	$myBody = "{ 'description':'Channel Description Updated' }"
	Invoke-MSGraphQuery `
		-URI $Url `
		-Body $myBody `
		-Token $myToken `
		-Meth Delete
}
#gavdcodeend 16

#*** Using PnP Graph PowerShell ----------------------------------------------------------

#gavdcodebegin 42
Function GrPsLoginGraphPnPWithInteraction()
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantUrl
	)

	Connect-PnPOnline -Url $TenantUrl -Interactive

	#Disconnect-PnPOnline
}
#gavdcodeend 42

#gavdcodebegin 43
Function GrPsLoginGraphPnPWithInteractionMFA()
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantUrl
	)

	Connect-PnPOnline -Url $TenantUrl -DeviceLogin -LaunchBrowser

	#Disconnect-PnPOnline
}
#gavdcodeend 43

#gavdcodebegin 18
Function GrPsLoginGraphPnP_GetTeamUsersWithInteraction()
{
	GrPsLoginGraphPnPWithInteractionMFA -TenantUrl $configFile.appsettings.TenantUrl
	
	Get-PnPTeamsUser -Team "Design"

	Disconnect-PnPOnline
}
#gavdcodeend 18

#gavdcodebegin 21
Function GrPsLoginGraphPnPGetToken()
{
	Connect-PnPOnline -Url $configFile.appsettings.TenantUrl -DeviceLogin -LaunchBrowser
	Get-PnPGraphAccessToken -Decoded

	Disconnect-PnPOnline
}
#gavdcodeend 21

#gavdcodebegin 20
Function GrPsLoginGraphPnPWithAccPw()
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object System.Management.Automation.PSCredential `
								-argumentlist $UserName, $securePW

	Connect-PnPOnline -Url $TenantUrl -Credentials $myCredentials
}
#gavdcodeend 20

#gavdcodebegin 47
Function GrPsLoginGraphPnPWithAccPwAndClientId()
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantUrl,
 
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

	Connect-PnPOnline -Url $TenantUrl -ClientId $ClientId -Credentials $myCredentials
}
#gavdcodeend 47

#gavdcodebegin 44
Function GrPsLoginGraphPnP_GetContextWithAccPw()
{
	GrPsLoginGraphPnPWithAccPwAndClientId `
					-TenantUrl $configFile.appsettings.TenantUrl `
					-ClientId $configFile.appSettings.ClientIdWithAccPw `
					-UserName $configFile.appSettings.UserName `
					-UserPw $configFile.appSettings.UserPw
	
	Get-PnPContext

	Disconnect-PnPOnline
}
#gavdcodeend 44

#gavdcodebegin 45
Function GrPsLoginGraphPnPWithSecret()
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

	Connect-PnPOnline -Url $configFile.appsettings.TenantUrl `
					  -AccessToken $myToken.AccessToken

	# Does not work anymore
	#Connect-PnPOnline -Url $TenantUrl -ClientId $ClientId -ClientSecret $ClientSecret
}
#gavdcodeend 45

#gavdcodebegin 46
Function GrPsLoginGraphPnP_GetTeamUsersWithSecret()
{
	GrPsLoginGraphPnPWithSecret `
					-TenantName	$configFile.appsettings.TenantName `
					-ClientId $configFile.appsettings.ClientIdWithSecret `
					-ClientSecret $configFile.appsettings.ClientSecret

	Get-PnPTeamsUser -Team "Design"

	Disconnect-PnPOnline
}
#gavdcodeend 46

#gavdcodebegin 48
Function GrPsLoginGraphPnPWithCertificate()
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientId,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateThumbprint
	)

	Connect-PnPOnline -Url $TenantUrl `
					  -Tenant $TenantName `
					  -ClientId $ClientId `
					  -Thumbprint $CertificateThumbprint
}
#gavdcodeend 48

#gavdcodebegin 49
Function GrPsLoginGraphPnPWithCertificateFile()
{
	[SecureString]$secureCertPw = ConvertTo-SecureString -String `
							$configFile.appSettings.CertificateFilePw -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.TenantUrl `
					  -Tenant $configFile.appsettings.TenantName `
					  -ClientId $configFile.appSettings.ClientIdWithCert `
					  -CertificatePath $configFile.appSettings.CertificateFilePath `
					  -CertificatePassword $certPw 
}
#gavdcodeend 49

#gavdcodebegin 50
Function GrPsLoginGraphPnP_GetLanguagesWithCertificate()
{
	GrPsLoginGraphPnPWithCertificate `
					-TenantUrl $configFile.appsettings.TenantUrl `
					-TenantName $configFile.appsettings.TenantName `
					-ClientId $configFile.appSettings.ClientIdWithCert `
					-CertificateThumbprint $configFile.appSettings.CertificateThumbprint
	
	Get-PnPAvailableLanguage

	Disconnect-PnPOnline
}
#gavdcodeend 50

#----------------------------------------------------------------------------------------

## Running the Functions
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#$myClientIdWithSecret = $configFile.appsettings.ClientIdWithSecret
#$myClientSecret = $configFile.appsettings.ClientSecret
#$myClientIdWithAccPw = $configFile.appsettings.ClientIdWithAccPw
#$myTenantName = $configFile.appsettings.TenantName
#$myTenantUrl = $configFile.appsettings.TenantUrl
#$myClientIdWithCert = $configFile.appsettings.ClientIdWithCert
#$myCertificateThumbprint = $configFile.appsettings.CertificateThumbprint
#$myCertificateFilePath = $configFile.appSettings.CertificateFilePath
#$myCertificateFilePw = $configFile.appSettings.CertificateFilePw
#$myUserName = $configFile.appsettings.UserName
#$myUserPw = $configFile.appsettings.UserPw

#*** Using Classic PowerShell cmdlets
#GrPsGetTeam
#GrPsCreateChannel
#GrPsGetChannel
#GrPsUpdateChannel
#GrPsDeleteChannel

#*** Using Microsoft Graph PowerShell SDK cmdlets
#GrPsLoginGraphSDKWithInteraction
#GrPsLoginGraphSDKAssignRights
#GrPsLoginGraphSDKApplication
#GrPsLoginGraphSDKGetContextInfo
#GrPsLoginGraphSDK_GetMe
#GrPsLoginGraphSDKConnectDisconnect
#GrPsLoginGraphSDKCheckRights
#GrPsLoginGraphSDKCheckAvailableRights
#GrPsLoginGraphSDKSetVersion
#GrPsLoginGraphSDKAssignRights
#GrPsLoginGraphSDKWithAccPw
#GrPsLoginGraphSDK_GetUserWithAccPw
#GrPsLoginGraphSDKWithSecret
#GrPsLoginGraphSDK_GetUsersWithSecret
#GrPsLoginGraphSDKWithCertificate
#GrPsLoginGraphSDKWithCertificateFile
#GrPsLoginGraphSDK_GetUsersWithCertificate
#GrPsGetGroupsSelect

#*** Using MSAL.PS module to get the token
#GrPsLoginGraphMsalWithInteraction
#GrPsLoginGraphMsalWithAccPw
#GrPsLoginGraphMsalWithSecret
#GrPsLoginGraphMsalWithCertificate
#GrPsLoginGraphMsal_GetTeamWithAccPw
#GrPsLoginGraphMsal_GetUsersWithSecret

#*** Using PowerShell-MicrosoftGraphAPI module (Other modules, not MS)
#GrPsGetToken
#GrPsGetTeamWithModule
#GrPsCreateChannelWithModule
#GrPsUpdateChannelWithModule
#GrPsDeleteChannelWithModule

#*** Using PnP Graph PowerShell
#GrPsLoginGraphPnPWithInteraction
#GrPsLoginGraphPnPWithInteractionMFA
#GrPsLoginGraphPnPGetToken
#GrPsLoginGraphPnP_GetTeamUsersWithInteraction
#GrPsLoginGraphPnPWithAccPw
#GrPsLoginGraphPnPWithAccPwAndClientId
#GrPsLoginGraphPnP_GetContextWithAccPw
#GrPsLoginGraphPnPWithSecret
#GrPsLoginGraphPnP_GetTeamUsersWithSecret
#GrPsLoginGraphPnPWithCertificate
#GrPsLoginGraphPnPWithCertificateFile
#GrPsLoginGraphPnP_GetLanguagesWithCertificate

Write-Host "Done" 
