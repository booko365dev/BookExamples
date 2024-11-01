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
    $ClientID = $configFile.appsettings.ClientIdWithSecret
    $ClientSecret = $configFile.appsettings.ClientSecret
    $TenantName = $configFile.appsettings.TenantName
   
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

function PsGraphSdk_LoginWithSecret
{
    $ClientID = $configFile.appsettings.ClientIdWithSecret
    $ClientSecret = $configFile.appsettings.ClientSecret
    $TenantName = $configFile.appsettings.TenantName

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$ClientSecret -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $ClientID, $securePW

	Connect-MgGraph -TenantId $TenantName `
					-ClientSecretCredential $myCredentials
}
#gavdcodeend 001 


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#-- Using the Microsoft Graph REST API to manage App Registrations ----------------------

#gavdcodebegin 002
function PsM365GraphRest_GetAllAppRegistrations 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
    $graphApiUrl = "https://graph.microsoft.com/v1.0"

    $myUri = "$($graphApiUrl)/applications"
    $myHeaders = @{
        "Authorization" = "Bearer $($myAccessToken)"
        "Content-Type"  = "application/json"
    }
    
    $myResponse = Invoke-RestMethod -Uri $myUri `
                                    -Headers $myHeaders `
                                    -Method Get

    $myResponseJson = $myResponse | ConvertTo-Json -Depth 10

    Write-Host $myResponseJson
}
#gavdcodeend 002

#gavdcodebegin 003
function PsM365GraphRest_GetOneAppRegistrationByObjectId 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"

	$myAppObjectId = "824741c8-88da-4414-808e-a2d0181cd1c4" # Object ID, not Client ID

	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)"
	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}
	
    $myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Method Get

    $myResponseJson = $myResponse | ConvertTo-Json -Depth 10

    Write-Host $myResponseJson
}
#gavdcodeend 003

#gavdcodebegin 004
function PsM365GraphRest_GetOneAppRegistrationByClientId
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"

	$myAppClientId = "5a84f9ed-d0be-4f7e-9fe8-42efb58acd2a" # Client ID, not Object ID

	$myUri = "$($graphApiUrl)/applications(appId='$($myAppClientId)')"
	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}
	
    $myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Method Get

    $myResponseJson = $myResponse | ConvertTo-Json -Depth 10

    Write-Host $myResponseJson
}
#gavdcodeend 004

#gavdcodebegin 005
function PsM365GraphRest_GetOneAppRegistrationByObjectIdByProperties
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"

	$myAppObjectId = "5279baca-6361-4b17-a885-2a00cd2fd73a" # Object ID, not Client ID

	# It can be used also by Client ID and for all App Registrations
	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)" + `
												"?`$select=displayName,appId,id"
	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}
												
    $myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Method Get

    $myResponseJson = $myResponse | ConvertTo-Json -Depth 10

    Write-Host $myResponseJson
}
#gavdcodeend 005

#gavdcodebegin 006
function PsM365GraphRest_CreateAppRegistrationGraphApi
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myUri = "$($graphApiUrl)/applications"
	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}
	
	# displayName is required, other properties are optional
	$myBody = @{
		displayName = "Test_MyAppRegFromGraphApi"
		web = @{
			redirectUris = @()
			homePageUrl = $null
			logoutUrl = $null
			implicitGrantSettings = @{
				enableIdTokenIssuance = $false
				enableAccessTokenIssuance = $false
			}
		}
	} | ConvertTo-Json

	$myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Body $myBody `
								    -Method Post

	Write-Host "Client ID - " $myResponse.appId
}
#gavdcodeend 006

#gavdcodebegin 007
function PsM365GraphRest_AddOwnerToAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myAppClientId = "cec8b03d-f21d-48c2-ac87-c6afd4bc4dbc" # Client ID
	$myAppObjectId = "d11874ad-129e-4f65-a53f-91e5e3e75bf2" # Object ID
	$myUserEmail = "user@domain.onmicrosoft.com"

	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}
	
	# Find the User ID by Email
	$myUri = "$($graphApiUrl)/users?`$filter=mail eq '$($myUserEmail)'"

	$myUser = Invoke-RestMethod -Uri $myUri `
								-Headers $myHeaders `
								-Method Get
	$myUserId = $myUser.value.id

	# Create a Service Principal for the Application
	$myUri = "$($graphApiUrl)/servicePrincipals"
	
	$myBody = @{
	  	"appId" = "$myAppClientId"
	} | ConvertTo-Json

	$myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Body $myBody `
								    -Method Post

	Write-Host "Service Principal ID - " $myResponse.id

	# Add the User as an Owner of the App Registration
	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)/owners/`$ref"

	$myBody = @{
		"@odata.id" = `
			"$($graphApiUrl)/directoryObjects/$($myUserId)"
	} | ConvertTo-Json

	$myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Body $myBody `
								    -Method Post

	Write-Host "User set as owner"
}
#gavdcodeend 007

#gavdcodebegin 008
function PsM365GraphRest_AddDelegatedClaimsToAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"

	$myAppClientId = "cec8b03d-f21d-48c2-ac87-c6afd4bc4dbc" # Client ID
	$myClaimName = "User.ReadWrite.All"

	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}

	# Get the client service principal
	$myUri = "$($graphApiUrl)/servicePrincipals?`$filter=appId eq '$($myAppClientId)'"
	$clientServicePrincipal = Invoke-RestMethod -Uri $myUri `
												-Headers $myHeaders `
												-Method Get
	
	# Get the service principal for Microsoft Graph
	$myUri = "$($graphApiUrl)/servicePrincipals?`$filter=displayName eq 'Microsoft Graph'"
	$graphServicePrincipal = Invoke-RestMethod -Uri $myUri `
											   -Headers $myHeaders `
											   -Method Get 

	# Get the oAuth2PermissionScope for Delegated
	$claimScope = $graphServicePrincipal.value[0].oauth2PermissionScopes `
								| Where-Object { $_.value -eq $($myClaimName) }

	# Grant the Delegated permission
	$body = @{
		clientId     = $($clientServicePrincipal.value[0].id)
		consentType  = "AllPrincipals"
		principalId  = $null
		resourceId   = $($graphServicePrincipal.value[0].id)
		scope        = $($claimScope.value)
	} | ConvertTo-Json -Depth 10

	# Set the righst for Delegated
	$myUri = "$($graphApiUrl)/oauth2PermissionGrants"
	$myResponse = Invoke-RestMethod -Uri $myUri `
									-Headers $myHeaders `
									-Body $body `
									-Method Post

    $myResponseJson = $myResponse | ConvertTo-Json -Depth 10

    Write-Host $myResponseJson
}
#gavdcodeend 008

#gavdcodebegin 009
function PsM365GraphRest_DeleteDelegatedClaimsFromAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myAppClientId = "cec8b03d-f21d-48c2-ac87-c6afd4bc4dbc" # Client ID
	$myClaimName = "User.ReadWrite.All"

	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}

	# Get the client service principal
	$myUri = "$($graphApiUrl)/servicePrincipals?`$filter=appId eq '$($myAppClientId)'"
	$clientServicePrincipal = Invoke-RestMethod -Uri $myUri `
												-Headers $myHeaders `
												-Method Get

	# Get the scope for the Delegated claim
	$myUri = "$($graphApiUrl)/oauth2PermissionGrants"
	$myPermissionGrants = Invoke-RestMethod -Uri $myUri `
											-Headers $myHeaders `
											-Method Get
	$myClaim = $myPermissionGrants.value `
			| Where-Object { ($_.clientId -eq "$($clientServicePrincipal.value[0].id)") `
						-and ($_.scope -eq $($myClaimName)) }	
	
	$myUri = "$($graphApiUrl)/oauth2PermissionGrants/$($myClaim.id)"
	$myResponse = Invoke-RestMethod -Uri $myUri `
									-Headers $myHeaders `
									-Method Delete

    $myResponseJson = $myResponse | ConvertTo-Json -Depth 10

    Write-Host $myResponseJson
}
#gavdcodeend 009

#gavdcodebegin 010
function PsM365GraphRest_AddApplicationClaimsToAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myAppClientId = "cec8b03d-f21d-48c2-ac87-c6afd4bc4dbc" # Client ID
	$myClaimName = "AuditLog.Read.All"

	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}

	# Get the client service principal
	$myUri = "$($graphApiUrl)/servicePrincipals?`$filter=appId eq '$($myAppClientId)'"
	$clientServicePrincipal = Invoke-RestMethod -Uri $myUri `
												-Headers $myHeaders `
												-Method Get
		
	# Get the service principal for Microsoft Graph
	$myUri = "$($graphApiUrl)/servicePrincipals?`$filter=displayName eq 'Microsoft Graph'"
	$graphServicePrincipal = Invoke-RestMethod -Uri $myUri `
											   -Headers $myHeaders `
											   -Method Get

	# Get the oAuth2PermissionScope for Application
	$claimRole = $graphServicePrincipal.value[0].appRoles `
								| Where-Object { $_.value -eq $($myClaimName) }

	# Assign the Application permission
	$body = @{
		principalId   = $clientServicePrincipal.value[0].id
		resourceId    = $graphServicePrincipal.value[0].id
		appRoleId     = $claimRole.id
	} | ConvertTo-Json -Depth 10

	# Set the righst for Application
	$myUri = "$($graphApiUrl)/servicePrincipals" + `
							"/$($clientServicePrincipal.value[0].id)/appRoleAssignedTo"
	$myResponse = Invoke-RestMethod -Uri $myUri `
									-Headers $myHeaders `
									-Body $body `
									-Method Post

    $myResponseJson = $myResponse | ConvertTo-Json -Depth 10

    Write-Host $myResponseJson
}
#gavdcodeend 010

#gavdcodebegin 011
function PsM365GraphRest_DeleteApplicationClaimsFromAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myAppClientId = "cec8b03d-f21d-48c2-ac87-c6afd4bc4dbc" # Client ID
	$myClaimName = "AuditLog.Read.All"

	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}

	# Get the client service principal
	$myUri = "$($graphApiUrl)/servicePrincipals?`$filter=appId eq '$($myAppClientId)'"
	$clientServicePrincipal = Invoke-RestMethod -Uri $myUri `
												-Headers $myHeaders `
												-Method Get
		
	# Get the service principal for Microsoft Graph
	$myUri = "$($graphApiUrl)/servicePrincipals?`$filter=displayName eq 'Microsoft Graph'"
	$graphServicePrincipal = Invoke-RestMethod -Uri $myUri `
											   -Headers $myHeaders `
											   -Method Get

	# Get the oAuth2PermissionScope for Application
	$myGraphRole = $graphServicePrincipal.value[0].appRoles `
								| Where-Object { $_.value -eq $($myClaimName) }
	
	# Get the Application Role Assignment
	$myUri = "$($graphApiUrl)/servicePrincipals" + `
						"/$($clientServicePrincipal.value[0].id)/appRoleAssignments"
	$myAllRoles = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Method Get
	$myRole = $myAllRoles.value | Where-Object { $_.appRoleId -eq "$($myGraphRole.id)" }

	# Remove the righst for Application
	$myUri = "$($graphApiUrl)/servicePrincipals" + `
			"/$($clientServicePrincipal.value[0].id)/appRoleAssignments/$($myRole.id)"
	$myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Method Delete

    Write-Host $myResponse
}
#gavdcodeend 011

#gavdcodebegin 012
function PsM365GraphRest_AddSecretToAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myAppObjectId = "d11874ad-129e-4f65-a53f-91e5e3e75bf2" # Object ID

	# The values for the Secret
	$mySecretName = "My AppReg Secret"
	$mySecretDurationInMonths = 26

	#Add the Secret to the App Registration
	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)/addPassword"

	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}
	
	$myBody = @{
		"passwordCredential" = @{
            "displayName" = "$($mySecretName)"
            "endDateTime" = "$(Get-Date -format o (Get-Date)" + `
											".AddMonths($mySecretDurationInMonths))"
        }	
	} | ConvertTo-Json

	$myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Body $myBody `
								    -Method Post
	
	Write-Host "Secret -" $myResponse.secretText
}
#gavdcodeend 012

#gavdcodebegin 013
function PsM365GraphRest_DeleteSecretFromAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myAppObjectId = "d11874ad-129e-4f65-a53f-91e5e3e75bf2" # Object ID

	# The values for the Secret
	$mySecretName = "My AppReg Secret"

	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}
	
	# Get the application details
	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)"
	$myAppRegistration = Invoke-RestMethod -Uri $myUri `
										   -Headers $myHeaders `
										   -Method Get

	# Find the secret to remove
	$mySecret = $myAppRegistration.passwordCredentials `
						| Where-Object { $_.displayName -eq $($mySecretName) }

	# Remove the Secret from the App Registration
	$myBody = @{
		"keyId" = "$($mySecret.keyId)"
	} | ConvertTo-Json

	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)/removePassword"
	$myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Body $myBody `
								    -Method Post

    $myResponseJson = $myResponse | ConvertTo-Json -Depth 10

    Write-Host $myResponseJson
}
#gavdcodeend 013

#gavdcodebegin 014
function PsM365GraphRest_AddCertificateToAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myAppObjectId = "d11874ad-129e-4f65-a53f-91e5e3e75bf2" # Object ID
	$myTenantId = $configFile.appsettings.TenantName

	# Create a Self-Signed Certificate
	$myCertPathPublic = "C:\Temporary\MyCertificate.cer"
	$myCertPathPrivate = "C:\Temporary\MyCertificate.pfx"
	$myCertPrivatePwd = "MyPassword"
	$myCertName = "CN=MyGraphApiCert"
	$myCertFriendlyName = "My Graph Api Cert"
	$myCertDurationInMonths = 23  # Max duration: 24 months

	$myCert = New-SelfSignedCertificate `
					-Subject "$($myCertName)" `
					-CertStoreLocation "Cert:\CurrentUser\My" `
					-KeyExportPolicy "Exportable" `
					-KeySpec "Signature" `
					-KeyLength 2048 `
					-KeyAlgorithm "RSA" `
					-HashAlgorithm "SHA256" `
					-DnsName "$($myTenantId)" `
					-NotAfter (Get-Date).AddMonths($($myCertDurationInMonths)) `
					-FriendlyName "$($myCertFriendlyName)" `
					-Provider "Microsoft Enhanced RSA and AES Cryptographic Provider"

	# Export the Certificate public key to a file
	Export-Certificate -Cert $myCert -FilePath $myCertPathPublic | Out-Null

	# Export the Certificate private key to a file
	$myPw = ConvertTo-SecureString -String $myCertPrivatePwd -Force -AsPlainText 
	Export-PfxCertificate -Cert $MyCert -FilePath $myCertPathPrivate `
										-Password $myPw | Out-Null
	
	# Get the Certificate’s en Thumbprint Base64 Strings
	$myCertB64 = [System.Convert]::ToBase64String([System.IO.File]::`
													ReadAllBytes($myCertPathPublic))
	$myCertThumbp = (Get-PfxCertificate -FilePath $myCertPathPublic).Thumbprint
	$myCertThumbpB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::`
													UTF8.GetBytes($myCertThumbp))

	#Add the Certificate to the App Registration
	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)"

	$myHeaders = @{
		"Authorization" = "Bearer $myAccessToken"
		"Content-Type"  = "application/json"
	}

	$myBody = @{
		"keyCredentials" = @(@{
			"displayName" = "$($myCertFriendlyName)"
			"type" = "AsymmetricX509Cert"
			"usage" = "Verify"
			"key" = "$($myCertB64)"
        })	
 	} | ConvertTo-Json

	$myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Body $myBody `
								    -Method Patch
	
	Write-Host "Thumbprint -" $myCertThumbp "-" $myCertThumbpB64
}
#gavdcodeend 014

#gavdcodebegin 015
function PsM365GraphRest_DeleteCertificateFromAppRegistrationAndComputer
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myAppObjectId = "d11874ad-129e-4f65-a53f-91e5e3e75bf2" # Object ID

	# The values for the Certificate
	$myCertThumbp = "2E01C7B224FEF0B7118EAF9D2C49ECD83104F135" # Thumbprint

	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}
	
	# Get the application details
	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)"
	$myAppRegistration = Invoke-RestMethod -Uri $myUri `
										   -Headers $myHeaders `
										   -Method Get

	# Find the certificate to remove
	$myCert = $myAppRegistration.keyCredentials `
						| Where-Object { $_.customKeyIdentifier -eq $($myCertThumbp) }

	# Remove the Secret from the App Registration
	$myBody = @{
		"keyId" = "$($myCert.keyId)"
	} | ConvertTo-Json

	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)/removeKey"
	$myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Body $myBody `
								    -Method Post
	
	Write-Host $myResponse

	# Delete Certificate from the Windows Certificate Store
	# Access the Local Machine's My store
	$myStore = New-Object System.Security.Cryptography.X509Certificates.X509Store`
									("My", "LocalMachine")
	$myStore.Open("ReadWrite")

	# Find the certificate by thumbprint
	$myCertLocal = $myStore.Certificates `
					| Where-Object { $_.Thumbprint -eq $($myCertThumbp) }

	$myStore.Remove($myCertLocal)
	
	Write-Host "Certificate removed from the Windows Certificate Store"

	# Delete the Certificate files if necessary
}
#gavdcodeend 015

#gavdcodebegin 016
function PsM365GraphRest_DeleteAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"
	
	$myAppObjectId = "d11874ad-129e-4f65-a53f-91e5e3e75bf2" # Object ID, not Client ID

	$myUri = "$($graphApiUrl)/applications/$($myAppObjectId)"
	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}

    $myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Method Delete
	
	Write-Host "App Registration deleted"
}
#gavdcodeend 016

#gavdcodebegin 017
function PsM365GraphRest_OtherRecipesForAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    $myAccessToken = (PsRest_GetAzureTokenWithSecret).access_token
	$graphApiUrl = "https://graph.microsoft.com/v1.0"

	$myHeaders = @{
		"Authorization" = "Bearer $($myAccessToken)"
		"Content-Type"  = "application/json"
	}

	# Find app registration by DisplayName
	$myUri = "$($graphApiUrl)/applications?`$filter=displayName " + `
													"eq 'Test_AppRegFromGraphApi'"

    $myResponse = Invoke-RestMethod -Uri $myUri `
								    -Headers $myHeaders `
								    -Method Get
	
	$myResponse

	# Get the Certificate Thumbprint Value from a .pfx file
	Get-PfxCertificate -FilePath "C:\PathToThePfxFile.pfx"
}
#gavdcodeend 017

#-- Using the Microsoft Graph PowerShell SDK to manage App Registrations ----------------

#gavdcodebegin 018
function PsM365GraphSdk_GetAllAppRegistrations
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

	Get-MgApplication
}
#gavdcodeend 018

#gavdcodebegin 019
function PsM365GraphSdk_GetOneAppRegistrationByObjectId 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

	$myAppObjectId = "824741c8-88da-4414-808e-a2d0181cd1c4" # Object ID, not Client ID

	Get-MgApplication -ApplicationId $myAppObjectId
}
#gavdcodeend 019

#gavdcodebegin 020
function PsM365GraphSdk_GetOneAppRegistrationByClientIdId 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

	$myAppClientId = "5a84f9ed-d0be-4f7e-9fe8-42efb58acd2a" # Client ID

	Get-MgApplication -Filter "AppId eq '$($myAppClientId)'"
}
#gavdcodeend 020

#gavdcodebegin 021
function PsM365GraphSdk_GetOneAppRegistrationByObjectIdByProperties 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

	$myAppObjectId = "824741c8-88da-4414-808e-a2d0181cd1c4" # Object ID, not Client ID

	Get-MgApplication -ApplicationId $myAppObjectId | Select-Object id, DisplayName
}
#gavdcodeend 021

#gavdcodebegin 022
function PsM365GraphSdk_CreateAppRegistrationGraphApi 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

	$myBody = @{
		displayName = "Test_MyAppRegFromGraphPsSdk"
	}

	$myApp =  New-MgApplication -BodyParameter $myBody

	Write-Host "Client ID - " $myApp.AppId
}
#gavdcodeend 022

#gavdcodebegin 023
function PsM365GraphSdk_AddOwnerToAppRegistration 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myAppClientId = "6f42c95c-0afd-453a-b356-2d6def30a2d5" # Client ID
    $myAppObjectId = "4f57cf4e-75dd-4895-a38d-1e540dd13383" # Object ID
    $myUserEmail = "user@domain.onmicrosoft.com"

    # Find the User ID by Email
    $myUser = Get-MgUser -ConsistencyLevel eventual -Count userCount `
						 -Search "UserPrincipalName:$($myUserEmail)"
    $myUserId = $myUser.Id;
    Write-Host("User ID - " + $myUserId)

    # Create a Service Principal for the Application
	$requestBody = @{
		appId = $myAppClientId
	}

	$myResponse = New-MgServicePrincipal -BodyParameter $requestBody
    Write-Host("Service Principal ID - " + $myResponse.Id)

    # Add the User as an Owner of the App Registration
	$myBody = @{
		"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/" + $myUserId
	}

	New-MgApplicationOwnerByRef -ApplicationId $myAppObjectId -BodyParameter $myBody

    Write-Host("User set as owner")
}
#gavdcodeend 023

#gavdcodebegin 024
function PsM365GraphSdk_AddDelegatedClaimsToAppRegistration 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myAppClientId = "6f42c95c-0afd-453a-b356-2d6def30a2d5" # Client ID
    $myClaimName = "User.ReadWrite.All"

    # Get the client service principal
	$clientServicePrincipal = Get-MgServicePrincipal `
									-Filter "appId eq '$($myAppClientId)'"
	Write-Host("Client Service Principal ID - " + $clientServicePrincipal.Id)

    # Get the service principal for Microsoft Graph
	$graphServicePrincipal = Get-MgServicePrincipal `
									-Filter "displayName eq 'Microsoft Graph'"
	Write-Host("Graph Service Principal ID - " + $graphServicePrincipal.Id)

    # Get the oAuth2PermissionScope for Delegated
	$claimScope = $graphServicePrincipal.oauth2PermissionScopes | Where-Object { $_.value -eq $myClaimName }
	write-host("Claim Scope - " + $claimScope.value)

    # Grant the Delegated permission
	$myOAuth2PermissionGrant = @{
        clientId = $clientServicePrincipal.id
        consentType = "AllPrincipals"
        principalId = $null
        resourceId = $graphServicePrincipal.id
        scope = $claimScope.value
    }
	New-MgOauth2PermissionGrant -BodyParameter $myOAuth2PermissionGrant

    Write-Host("Delegated permission granted");
}
#gavdcodeend 024

#gavdcodebegin 025
function PsM365GraphSdk_DeleteDelegatedClaimsFromAppRegistration 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myAppClientId = "6f42c95c-0afd-453a-b356-2d6def30a2d5" # Client ID
    $myClaimName = "User.ReadWrite.All"

    # Get the client service principal
	$clientServicePrincipal = Get-MgServicePrincipal `
									-Filter "appId eq '$($myAppClientId)'"
	Write-Host("Client Service Principal ID - " + $clientServicePrincipal.Id)

    # Get the scope for the Delegated claim
	$myPermissionGrants = Get-MgOauth2PermissionGrant
	$myClaim = $myPermissionGrants `
			| Where-Object { ($_.ClientId -eq "$($clientServicePrincipal.id)") `
								-and ($_.Scope -eq $myClaimName) }

    # Delete the Delegated permission
	Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $myClaim.id

    Write-Host("Delegated permission deleted");
}
#gavdcodeend 025

#gavdcodebegin 026
function PsM365GraphSdk_AddApplicationClaimsToAppRegistration 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myAppClientId = "6f42c95c-0afd-453a-b356-2d6def30a2d5" # Client ID
    $myClaimName = "AuditLog.Read.All"

    # Get the client service principal
	$clientServicePrincipal = Get-MgServicePrincipal `
									-Filter "appId eq '$($myAppClientId)'"
	Write-Host("Client Service Principal ID - " + $clientServicePrincipal.Id)

    # Get the service principal for Microsoft Graph
	$graphServicePrincipal = Get-MgServicePrincipal `
									-Filter "displayName eq 'Microsoft Graph'"
	Write-Host("Graph Service Principal ID - " + $graphServicePrincipal.Id)

    # Get the Role
	$claimRole = $graphServicePrincipal.AppRoles | `
									Where-Object { $_.value -eq $myClaimName }
	write-host("Role - " + $claimRole.Id)

    # Grant the Application permission
	$myAppRoleAssignment = @{
        principalId = $clientServicePrincipal.Id
        resourceId = $graphServicePrincipal.Id
        appRoleId = $claimRole.Id
    }
	New-MgServicePrincipalAppRoleAssignedTo `
							-ServicePrincipalId $clientServicePrincipal.Id `
							-BodyParameter $myAppRoleAssignment

    Write-Host("Application permission granted");
}
#gavdcodeend 026

#gavdcodebegin 027
function PsM365GraphSdk_DeleteApplicationClaimsFromAppRegistration 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myAppClientId = "6f42c95c-0afd-453a-b356-2d6def30a2d5" # Client ID
    $myClaimName = "AuditLog.Read.All"

    # Get the client service principal
	$clientServicePrincipal = Get-MgServicePrincipal `
									-Filter "appId eq '$($myAppClientId)'"
	Write-Host("Client Service Principal ID - " + $clientServicePrincipal.Id)

    # Get the service principal for Microsoft Graph
	$graphServicePrincipal = Get-MgServicePrincipal `
									-Filter "displayName eq 'Microsoft Graph'"
	Write-Host("Graph Service Principal ID - " + $graphServicePrincipal.Id)

    # Get the oAuth2PermissionScope for Application
	$myGraphRole = $graphServicePrincipal.AppRoles | `
							Where-Object { $_.value -eq $myClaimName }

    # Get the Application Role Assignment
	$myAllRoles = Get-MgServicePrincipalAppRoleAssignment `
							-ServicePrincipalId $clientServicePrincipal.Id
	$myRole = $myAllRoles | Where-Object { $_.AppRoleId -eq $myGraphRole.Id } | `
							Select-Object -First 1

    # Delete the Application permission
	Remove-MgServicePrincipalAppRoleAssignment `
							-ServicePrincipalId $clientServicePrincipal.Id `
							-AppRoleAssignmentId $myRole.Id

    Write-Host("Application permission deleted");
}
#gavdcodeend 027

#gavdcodebegin 028
function PsM365GraphSdk_AddSecretToAppRegistration 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myAppObjectId = "4f57cf4e-75dd-4895-a38d-1e540dd13383" # Object ID

    # The values for the Secret
    $mySecretName = "My AppReg Secret"
    $mySecretDurationInMonths = 26

    # Add the Secret to the App Registration
	$myBody = @{
		passwordCredential = @{
			displayName = $mySecretName
			endDateTime = (Get-Date).AddMonths($mySecretDurationInMonths)
		}
	}

	$myResponse = Add-MgApplicationPassword -ApplicationId $myAppObjectId `
											-BodyParameter $myBody

    Write-Host("Secret - " + $myResponse.SecretText);
}
#gavdcodeend 028

#gavdcodebegin 029
function PsM365GraphSdk_DeleteSecretFromAppRegistration 
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myAppObjectId = "4f57cf4e-75dd-4895-a38d-1e540dd13383" # Object ID

    # The values for the Secret
    $mySecretName = "My AppReg Secret"

    # Get the application details
	$myAppRegistration = Get-MgApplication -ApplicationId $myAppObjectId

    # Find the secret to remove
	$mySecret = $myAppRegistration.PasswordCredentials | `
						Where-Object { $_.DisplayName -eq $mySecretName } | `
						Select-Object -First 1
    
	# Remove the secret
	$myBody = @{
		keyId = $mySecret.KeyId
	}
	Remove-MgApplicationPassword -ApplicationId $myAppObjectId -BodyParameter $myBody

    Write-Host("Secret deleted");
}
#gavdcodeend 029

#gavdcodebegin 030
function PsM365GraphSdk_AddCertificateToAppRegistration
{
    # Requires Application.Read.All, AppRoleAssignment.ReadWrite.All,
    # Application.ReadWrite.OwnedBy and Directory.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myAppObjectId = "4f57cf4e-75dd-4895-a38d-1e540dd13383" # Object ID
	$myTenantId = $configFile.appsettings.TenantName

	# Create a Self-Signed Certificate
	$myCertPathPublic = "C:\Temporary\MyCertificate.cer"
	$myCertPathPrivate = "C:\Temporary\MyCertificate.pfx"
	$myCertPrivatePwd = "MyPassword"
	$myCertName = "CN=MyGraphApiCert"
	$myCertFriendlyName = "My Graph Api Cert"
	$myCertDurationInMonths = 23  # Max duration: 24 months

	$myCert = New-SelfSignedCertificate `
					-Subject "$($myCertName)" `
					-CertStoreLocation "Cert:\CurrentUser\My" `
					-KeyExportPolicy "Exportable" `
					-KeySpec "Signature" `
					-KeyLength 2048 `
					-KeyAlgorithm "RSA" `
					-HashAlgorithm "SHA256" `
					-DnsName "$($myTenantId)" `
					-NotAfter (Get-Date).AddMonths($($myCertDurationInMonths)) `
					-FriendlyName "$($myCertFriendlyName)" `
					-Provider "Microsoft Enhanced RSA and AES Cryptographic Provider"

	# Export the Certificate public key to a file
	Export-Certificate -Cert $myCert -FilePath $myCertPathPublic | Out-Null

	# Export the Certificate private key to a file
	$myPw = ConvertTo-SecureString -String $myCertPrivatePwd -Force -AsPlainText 
	Export-PfxCertificate -Cert $MyCert -FilePath $myCertPathPrivate `
										-Password $myPw | Out-Null
	
	# Get the Certificate’s en Thumbprint Base64 Strings
	$myCertB64 = [System.Convert]::ToBase64String([System.IO.File]::`
													ReadAllBytes($myCertPathPublic))
	$myCertThumbp = (Get-PfxCertificate -FilePath $myCertPathPublic).Thumbprint
	$myCertThumbpB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::`
													UTF8.GetBytes($myCertThumbp))

	#Add the Certificate to the App Registration
	$myBody = @{
		"keyCredentials" = @(@{
			"displayName" = "$($myCertFriendlyName)"
			"type" = "AsymmetricX509Cert"
			"usage" = "Verify"
			"key" = "$($myCertB64)"
        })	
 	} #| ConvertTo-Json

	$myResponse = Add-MgApplicationKey -ApplicationId $myAppObjectId `
									   -BodyParameter $myBody

	Write-Host "Thumbprint -" $myCertThumbp "-" $myCertThumbpB64
}
#gavdcodeend 030

#gavdcodebegin 031
function PsM365GraphSdk_DeleteCertificateFromAppRegistrationAndComputer
{
    # Requires Application.Read.All, AppRoleAssignment.ReadWrite.All,
    # Application.ReadWrite.OwnedBy and Directory.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myAppObjectId = "4f57cf4e-75dd-4895-a38d-1e540dd13383" # Object ID
    $myAppClientId = "d86afffc-eb8d-4ac5-856f-6ddd9a347033" # Client ID
	$myTenantId = $configFile.appsettings.TenantName
	$myCertPathPrivate = "C:\Temporary\MyCertificate.pfx"
	$myCertPrivatePwd = "MyPassword"
	$myTenantId = $configFile.appsettings.TenantName

    # The values for the Certificate Thumbprint
	$privCertificate = New-Object `
			System.Security.Cryptography.X509Certificates.X509Certificate2(`
				$myCertPathPrivate, $myCertPrivatePwd)
    $securityKey = New-Object `
			System.IdentityModel.Tokens.X509SecurityKey($privCertificate)
    $myCertThumbp = $privCertificate.Thumbprint

    # Get the application details
	$myAppRegistration = Get-MgApplication -ApplicationId $myAppObjectId
    
	# Find the certificate to remove
    KeyCredential myCert = myAppRegistration.KeyCredentials
                    .Where(crt => crt.CustomKeyIdentifier != null &&
                           crt.CustomKeyIdentifier.SequenceEqual(
                               Convert.FromBase64String(myCertThumbp))).FirstOrDefault();

	$myCert = $myAppRegistration.KeyCredentials | `
		Where-Object { $_.CustomKeyIdentifier -ne $null -and `
					   $_.CustomKeyIdentifier -eq `
							[Convert]::FromBase64String($myCertThumbp) } `
					| Select-Object -First 1

	$myBody = @{
		keyId = $myCert.KeyId
	}

	Remove-MgApplicationKey -ApplicationId $clientServicePrincipal.Id `
							-BodyParameter $myBody

	Write-Host "Certificate deleted"

    # Delete Certificate from the Windows Certificate Store
    # Access the Local Machine's My store
	$winStore = New-Object `
		System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
	$winStore.Open("ReadWrite")

    # Find the certificate by thumbprint
	$myCertLocal = $winStore.Certificates.Find(`
							"FindByThumbprint", $myCertThumbp, $false) | `
				   Select-Object -First 1

	$winStore.Remove($myCertLocal)

    Write-Host "Certificate removed from the Windows Certificate Store"

    # Delete the Certificate files if necessary
}
#gavdcodeend 031

#gavdcodebegin 032
function PsM365GraphSdk_DeleteAppRegistration
{
    # Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

	PsGraphSdk_LoginWithSecret

	$myAppObjectId = "824741c8-88da-4414-808e-a2d0181cd1c4" # Object ID, not Client ID
	
	Remove-MgApplication -ApplicationId $myAppObjectId

    Write-Host "App Registration deleted"
}
#gavdcodeend 032


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 032 *** 

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#-- Using the Microsoft Graph REST API to manage App Registrations ----------------------
#PsM365GraphRest_GetAllAppRegistrations
#PsM365GraphRest_GetOneAppRegistrationByObjectId
#PsM365GraphRest_GetOneAppRegistrationByClientId
#PsM365GraphRest_GetOneAppRegistrationByObjectIdByProperties
#PsM365GraphRest_CreateAppRegistrationGraphApi
#PsM365GraphRest_AddOwnerToAppRegistration
#PsM365GraphRest_AddDelegatedClaimsToAppRegistration
#PsM365GraphRest_DeleteDelegatedClaimsFromAppRegistration
#PsM365GraphRest_AddApplicationClaimsToAppRegistration
#PsM365GraphRest_DeleteApplicationClaimsFromAppRegistration
#PsM365GraphRest_AddSecretToAppRegistration
#PsM365GraphRest_DeleteSecretFromAppRegistration
#PsM365GraphRest_AddCertificateToAppRegistration
#PsM365GraphRest_DeleteCertificateFromAppRegistrationAndComputer
#PsM365GraphRest_DeleteAppRegistration
#PsM365GraphRest_OtherRecipesForAppRegistration

#-- Using the Microsoft Graph PowerShell SDK to manage App Registrations ----------------
#PsM365GraphSdk_GetAllAppRegistrations
#PsM365GraphSdk_GetOneAppRegistrationByObjectId
#PsM365GraphSdk_GetOneAppRegistrationByClientIdId
#PsM365GraphSdk_GetOneAppRegistrationByObjectIdByProperties
#PsM365GraphSdk_CreateAppRegistrationGraphApi
#PsM365GraphSdk_AddOwnerToAppRegistration
#PsM365GraphSdk_AddDelegatedClaimsToAppRegistration
#PsM365GraphSdk_DeleteDelegatedClaimsFromAppRegistration
#PsM365GraphSdk_AddApplicationClaimsToAppRegistration
#PsM365GraphSdk_DeleteApplicationClaimsFromAppRegistration
#PsM365GraphSdk_AddSecretToAppRegistration
#PsM365GraphSdk_DeleteSecretFromAppRegistration
#PsM365GraphSdk_AddCertificateToAppRegistration
#PsM365GraphSdk_DeleteCertificateFromAppRegistrationAndComputer
#PsM365GraphSdk_DeleteAppRegistration

Write-Host "Done"
