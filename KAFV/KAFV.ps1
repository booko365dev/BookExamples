##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsSpRestApiMsal_LoginWithCertificateFile
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateFilePath,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateFilePw,

		[Parameter(Mandatory=$True)]
		[String]$SiteBaseUrl
	)

    $myAuthority = "https://login.microsoftonline.com/$TenantName"
    $myScopes = @("$SiteBaseUrl/.default")

    $myCertificate = New-Object `
						System.Security.Cryptography.X509Certificates.X509Certificate2 `
						-ArgumentList $CertificateFilePath, $CertificateFilePw

    $myApp = New-MsalClientApplication -ClientId $ClientId `
                                       -ClientCertificate $myCertificate `
                                       -Authority $myAuthority

    $myToken = Get-MsalToken -ConfidentialClientApplication $myApp -Scopes $myScopes

    return $myToken.AccessToken
}

function PsMsal_GetAzureTokenWithAccPw
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

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
								-argumentlist $UserName, $securePW

	$myOAuth = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -UserCredential $myCredentials 
	
	#Write-Host $myToken.AccessToken

	return $myOAuth
}

function PsPnPPowerShell_LoginGraphWithCertificateFile
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$SiteBaseUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientId,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateFilePath,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateFilePw
	)

	[SecureString]$secureCertPw = ConvertTo-SecureString -String `
						$CertificateFilePw -AsPlainText -Force

	Connect-PnPOnline -Url $SiteBaseUrl -Tenant $TenantName -ClientId $ClientId `
				-CertificatePath $CertificateFilePath -CertificatePassword $secureCertPw 
}

function PsCliM365_LoginWithCertificateFile
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientId,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateFilePath,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateFilePw
	)

	m365 login --authType certificate `
			   --tenant $TenantName --appId $ClientId `
			   --certificateFile $CertificateFilePath --password $CertificateFilePw
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#==> Using the SharePoint REST API to work with Webhooks
#gavdcodebegin 001
function PsSpRestApiMsal_GetSubscriptions
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	$accessToken = PsSpRestApiMsal_LoginWithCertificateFile -TenantName $cnfTenantName `
									-ClientID $cnfClientIdWithCert `
									-CertificateFilePath $cnfCertificateFilePath `
									-CertificateFilePw $cnfCertificateFilePw `
									-SiteBaseUrl $cnfSiteBaseUrl

    $myHeader = @{
		"Authorization" = "Bearer $accessToken"
		"Accept"        = "application/json;odata=verbose"
	}

	$endpointUrl = $cnfSiteCollUrl + 
						"/_api/web/lists/GetByTitle('TestList')/subscriptions"

	$response = Invoke-RestMethod -Method Get `
						-Headers $myHeader `
						-Uri $endpointUrl `
						-ContentType "application/json;odata=verbose"

	foreach ($oneSubscription in $response.d.results) {
		Write-Host $oneSubscription.notificationUrl
	}
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpRestApiMsal_CreateSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	$accessToken = PsSpRestApiMsal_LoginWithCertificateFile -TenantName $cnfTenantName `
									-ClientID $cnfClientIdWithCert `
									-CertificateFilePath $cnfCertificateFilePath `
									-CertificateFilePw $cnfCertificateFilePw `
									-SiteBaseUrl $cnfSiteBaseUrl
									
	$myHeader = @{
					"Authorization" = "Bearer $accessToken"
					"Accept"        = "application/json;odata=verbose"
				}
	$myPayload = @{
					"__metadata" = @{ "type" = "SP.ListItem" };
					"resource" = "$cnfSiteCollUrl" + 
								"/_api/web/lists/GetByTitle('TestList')/subscriptions";
					"notificationUrl" = "https://domain.azurewebsites.net" + 
								"/api/SharePointWebhookReceiver";
					"expirationDateTime" = "2025-05-31T00:00:00Z"
				} | ConvertTo-Json

	$endpointUrl = $cnfSiteCollUrl + 
						"/_api/web/lists/GetByTitle('TestList')/subscriptions"
			
	$response = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $response
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpRestApiMsal_GetOneSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	$accessToken = PsSpRestApiMsal_LoginWithCertificateFile -TenantName $cnfTenantName `
									-ClientID $cnfClientIdWithCert `
									-CertificateFilePath $cnfCertificateFilePath `
									-CertificateFilePw $cnfCertificateFilePw `
									-SiteBaseUrl $cnfSiteBaseUrl

    $myHeader = @{
		"Authorization" = "Bearer $accessToken"
		"Accept"        = "application/json;odata=verbose"
	}

	$endpointUrl = $cnfSiteCollUrl + 
						"/_api/web/lists/GetByTitle('TestList')/" + 
						"subscriptions('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')"

	$response = Invoke-RestMethod -Method Get `
						-Headers $myHeader `
						-Uri $endpointUrl `
						-ContentType "application/json;odata=verbose"

	$responseString = $response | ConvertTo-Json

	Write-Host $responseString
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpRestApiMsal_UpdateSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	$accessToken = PsSpRestApiMsal_LoginWithCertificateFile -TenantName $cnfTenantName `
									-ClientID $cnfClientIdWithCert `
									-CertificateFilePath $cnfCertificateFilePath `
									-CertificateFilePw $cnfCertificateFilePw `
									-SiteBaseUrl $cnfSiteBaseUrl
									
	$myHeader = @{
					"Authorization" = "Bearer $accessToken"
					"Accept"        = "application/json;odata=verbose"
					"If-Match" 	    = "*"
				}
	$myPayload = @{
					"__metadata" = @{ "type" = "SP.ListItem" };
					"notificationUrl" = "https://domain.azurewebsites.net/" + 
											"api/SharePointWebhookReceiver";
					"expirationDateTime" = "2025-06-01T02:00:00Z"
				} | ConvertTo-Json

	$endpointUrl = $cnfSiteCollUrl + 
						"/_api/web/lists/GetByTitle('TestList')/" + 
						"subscriptions('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')"
			
	$response = Invoke-WebRequest -Method Patch `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $response
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpRestApiMsal_DeleteSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	$accessToken = PsSpRestApiMsal_LoginWithCertificateFile -TenantName $cnfTenantName `
									-ClientID $cnfClientIdWithCert `
									-CertificateFilePath $cnfCertificateFilePath `
									-CertificateFilePw $cnfCertificateFilePw `
									-SiteBaseUrl $cnfSiteBaseUrl

	$myHeader = @{
					"Authorization" = "Bearer $accessToken"
					"Accept"        = "application/json;odata=verbose"
					"If-Match" 	    = "*"
				}
					
	$endpointUrl = $cnfSiteCollUrl + 
				"/_api/web/lists/GetByTitle('TestList')/" + 
				"subscriptions('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')"
	
	$response = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $response
}
#gavdcodeend 005

#==> Using the PnP PowerShell module to work with Webhooks
#gavdcodebegin 006
function PsPnPPowerShell_GetSubscriptions
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	PsPnPPowerShell_LoginGraphWithCertificateFile -SiteBaseUrl $cnfSiteCollUrl `
				-TenantName $cnfTenantName -ClientId $cnfClientIdWithCert `
				-CertificateFilePath $cnfCertificateFilePath `
				-CertificateFilePw $cnfCertificateFilePw

	Get-PnPWebhookSubscription -List "TestList"

	Disconnect-PnPOnline
}
#gavdcodeend 006

#gavdcodebegin 007
function PsPnPPowerShell_CreateSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	PsPnPPowerShell_LoginGraphWithCertificateFile -SiteBaseUrl $cnfSiteCollUrl `
				-TenantName $cnfTenantName -ClientId $cnfClientIdWithCert `
				-CertificateFilePath $cnfCertificateFilePath `
				-CertificateFilePw $cnfCertificateFilePw

	# To subscribe to the SpWebHookListener function in the local environment
	#	you need to use ngrok to create a tunnel to your local machine
	Add-PnPWebhookSubscription -List "TestList" `
							-NotificationUrl "https://8014-62--141.ngrok-free.app/" + 
														"api/SpWebhookListener" `
							-ExpirationDate "2025-05-31T00:00:00Z" `
							-ClientState "guitacaClientState"

	## To subscribe to the SharePointWebhookReceiver function in Azure
	# Add-PnPWebhookSubscription -List "TestList" `
	# 						-NotificationUrl "https://domain.azurewebsites.net/" + 
	#										"api/SharePointWebhookReceiver" `
	# 						-ExpirationDate "2025-05-31T00:00:00Z" #-ClientState "myState"

	Disconnect-PnPOnline
}
#gavdcodeend 007

#gavdcodebegin 008
function PsPnPPowerShell_UpdateSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	PsPnPPowerShell_LoginGraphWithCertificateFile -SiteBaseUrl $cnfSiteCollUrl `
				-TenantName $cnfTenantName -ClientId $cnfClientIdWithCert `
				-CertificateFilePath $cnfCertificateFilePath `
				-CertificateFilePw $cnfCertificateFilePw

	Set-PnPWebhookSubscription -List "TestList" `
							-Subscription "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
							-ExpirationDate "2025-06-01T02:00:00Z"

	Disconnect-PnPOnline
}
#gavdcodeend 008

#gavdcodebegin 009
function PsPnPPowerShell_DeleteSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	PsPnPPowerShell_LoginGraphWithCertificateFile -SiteBaseUrl $cnfSiteCollUrl `
				-TenantName $cnfTenantName -ClientId $cnfClientIdWithCert `
				-CertificateFilePath $cnfCertificateFilePath `
				-CertificateFilePw $cnfCertificateFilePw

	Remove-PnPWebhookSubscription -List "TestList" `
							-Identity "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Force

	Disconnect-PnPOnline
}
#gavdcodeend 009

#==> Using the CLI for Microsoft 365 module to work with Webhooks
#gavdcodebegin 010
function PsCliM365_GetSubscriptions
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	PsCliM365_LoginWithCertificateFile -TenantName $cnfTenantName `
							  -ClientId $cnfClientIdWithCert `
							  -CertificateFilePath $cnfCertificateFilePath `
							  -CertificateFilePw $cnfCertificateFilePw
	
	m365 spo list webhook list --webUrl $cnfSiteCollUrl --listTitle "TestList"

	m365 logout
}
#gavdcodeend 010

#gavdcodebegin 011
function PsCliM365_CreateSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	PsCliM365_LoginWithCertificateFile -TenantName $cnfTenantName `
							  -ClientId $cnfClientIdWithCert `
							  -CertificateFilePath $cnfCertificateFilePath `
							  -CertificateFilePw $cnfCertificateFilePw
	
	m365 spo list webhook add --webUrl $cnfSiteCollUrl --listTitle "TestList" `
				--notificationUrl "https://domain.azurewebsites.net/" + 
										"api/SharePointWebhookReceiver" `
				--expirationDateTime "2025-05-31T00:00:00Z" #-clientState "myState"

	m365 logout
}
#gavdcodeend 011

#gavdcodebegin 012
function PsCliM365_GetOneSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	PsCliM365_LoginWithCertificateFile -TenantName $cnfTenantName `
							  -ClientId $cnfClientIdWithCert `
							  -CertificateFilePath $cnfCertificateFilePath `
							  -CertificateFilePw $cnfCertificateFilePw
	
	m365 spo list webhook get --webUrl $cnfSiteCollUrl --listTitle "TestList" `
								--id "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
	m365 logout
}
#gavdcodeend 012

#gavdcodebegin 013
function PsCliM365_UpdateSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	PsCliM365_LoginWithCertificateFile -TenantName $cnfTenantName `
							  -ClientId $cnfClientIdWithCert `
							  -CertificateFilePath $cnfCertificateFilePath `
							  -CertificateFilePw $cnfCertificateFilePw
	
	m365 spo list webhook set --webUrl $cnfSiteCollUrl --listTitle "TestList" `
							--id "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
							--expirationDateTime "2025-06-01T02:00:00Z"
	m365 logout
}
#gavdcodeend 013

#gavdcodebegin 014
function PsCliM365_DeleteSubscription
{
	# Require permissions: Sites.Manage.All (Delegated), Sites.Manage.All (Application)
	PsCliM365_LoginWithCertificateFile -TenantName $cnfTenantName `
							  -ClientId $cnfClientIdWithCert `
							  -CertificateFilePath $cnfCertificateFilePath `
							  -CertificateFilePw $cnfCertificateFilePw
	
	m365 spo list webhook remove --webUrl $cnfSiteCollUrl --listTitle "TestList" `
							--id "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

	m365 logout
}
#gavdcodeend 014


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 014 ***

#region ConfigValuesPS.config
[xml]$config = Get-Content -Path "C:\Projects\ConfigValuesCS.config"
$cnfUserName               = $config.SelectSingleNode("//add[@key='UserName']").value
$cnfUserPw                 = $config.SelectSingleNode("//add[@key='UserPw']").value
$cnfTenantUrl              = $config.SelectSingleNode("//add[@key='TenantUrl']").value
$cnfSiteBaseUrl            = $config.SelectSingleNode("//add[@key='SiteBaseUrl']").value
$cnfSiteAdminUrl           = $config.SelectSingleNode("//add[@key='SiteAdminUrl']").value
$cnfSiteCollUrl            = $config.SelectSingleNode("//add[@key='SiteCollUrl']").value
$cnfTenantName             = $config.SelectSingleNode("//add[@key='TenantName']").value
$cnfClientIdWithAccPw      = $config.SelectSingleNode("//add[@key='ClientIdWithAccPw']").value
$cnfClientIdWithSecret     = $config.SelectSingleNode("//add[@key='ClientIdWithSecret']").value
$cnfClientSecret           = $config.SelectSingleNode("//add[@key='ClientSecret']").value
$cnfClientIdWithCert       = $config.SelectSingleNode("//add[@key='ClientIdWithCert']").value
$cnfCertificateThumbprint  = $config.SelectSingleNode("//add[@key='CertificateThumbprint']").value
$cnfCertificateFilePath    = $config.SelectSingleNode("//add[@key='CertificateFilePath']").value
$cnfCertificateFilePw      = $config.SelectSingleNode("//add[@key='CertificateFilePw']").value
#endregion ConfigValuesCS.config

#PsSpRestApiMsal_GetSubscriptions
#PsSpRestApiMsal_CreateSubscription
#PsSpRestApiMsal_GetOneSubscription
#PsSpRestApiMsal_UpdateSubscription
#PsSpRestApiMsal_DeleteSubscription

#PsPnPPowerShell_GetSubscriptions
#PsPnPPowerShell_CreateSubscription
#PsPnPPowerShell_UpdateSubscription
#PsPnPPowerShell_DeleteSubscription

#PsCliM365_GetSubscriptions
#PsCliM365_CreateSubscription
#PsCliM365_GetOneSubscription
#PsCliM365_UpdateSubscription
#PsCliM365_DeleteSubscription

Write-Host "Done" 
