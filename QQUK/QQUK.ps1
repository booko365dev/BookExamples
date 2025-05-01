##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsGraphRestApi_GetAzureTokenWithCertificateThumbprint
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

	$myToken = Invoke-RestMethod `
					-Method Post `
					-ContentType "application/x-www-form-urlencoded" `
					-Uri $myUrl `
					-Headers $myHeader `
					-Body $myBody

	$accessToken = $myToken.access_token
	return $accessToken
}

function PsPnPPowerShell_LoginGraphWithCertificateThumbprint
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

	Connect-PnPOnline -Url $SiteBaseUrl -Tenant $TenantName -ClientId $ClientId `
					  -Thumbprint $CertificateThumbprint
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

#gavdcodebegin 001
function PsGraphRestApiMs_GetUnifiedAuditLogs
{
	$accessToken = PsGraphRestApi_GetAzureTokenWithCertificateThumbprint `
							-ClientID $cnfClientIdWithCert `
							-TenantName $cnfTenantName `
							-CertificateThumbprint $cnfCertificateThumbprint

	$myHeader = @{
		"Authorization" = "Bearer $accessToken"
	}
					
	$fromDate = (Get-Date).AddDays(-7).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
	$toDate   = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
	$myPayload = @{
		"filterStartDateTime" = $fromDate
		"filterEndDateTime" = $toDate
		#"operationFilters" = @("sharePointFileOperation")
		#"limit" = 1000
	} | ConvertTo-Json
	
	$endpointUrl = "https://graph.microsoft.com/beta/security/auditLog/queries"
			
	$response = Invoke-RestMethod -Method Post `
							-Uri $endpointUrl `
							-Headers $myHeader `
							-Body $myPayload

	if ($response.value) {
		$response.value | Format-Table -Wrap -AutoSize
	} else {
		$response | Format-List
	}
}
#gavdcodeend 001

#gavdcodebegin 002
function PsGraphRestApiMs_CheckUnifiedAuditLogs
{
	$accessToken = PsGraphRestApi_GetAzureTokenWithCertificateThumbprint `
						-ClientID $cnfClientIdWithCert `
						-TenantName $cnfTenantName `
						-CertificateThumbprint $cnfCertificateThumbprint

	$myHeader = @{
		"Authorization" = "Bearer $accessToken"
	}
			
	$searchID = "5775c911-8dd0-4ad2-96ac-c7e0cda766aa"
	$endpointUrl = `
			"https://graph.microsoft.com/beta/security/auditLog/queries/$searchID"

	$response = Invoke-RestMethod -Method Get `
						-Headers $myHeader `
						-Uri $endpointUrl `
						-ContentType "application/json;odata=verbose"

	$responseJson = $response | ConvertTo-Json -Depth 10
	Write-Host $responseJson
}
#gavdcodeend 002

#gavdcodebegin 003
function PsGraphRestApiMs_ReadUnifiedAuditLogs
{
	$accessToken = PsGraphRestApi_GetAzureTokenWithCertificateThumbprint `
						-ClientID $cnfClientIdWithCert `
						-TenantName $cnfTenantName `
						-CertificateThumbprint $cnfCertificateThumbprint

	$myHeader = @{
		"Authorization" = "Bearer $accessToken"
	}
					
	$searchID = "5775c911-8dd0-4ad2-96ac-c7e0cda766aa"
	$endpointUrl = `
			"https://graph.microsoft.com/beta/security/auditLog/queries/$searchID/records"

	$response = Invoke-RestMethod -Method Get `
						-Headers $myHeader `
						-Uri $endpointUrl `
						-ContentType "application/json;odata=verbose"

	$responseJson = $response | ConvertTo-Json -Depth 10
	Write-Host $responseJson
}
#gavdcodeend 003

#gavdcodebegin 004
function PsGraphRestApiPurview_GetPurviewAuditLogs
{
	$accessToken = PsGraphRestApi_GetAzureTokenWithCertificateThumbprint `
							-ClientID $cnfClientIdWithCert `
							-TenantName $cnfTenantName `
							-CertificateThumbprint $cnfCertificateThumbprint

	$myHeader = @{
		"Authorization" = "Bearer $accessToken"
	}

	$dateWindow = `
			(Get-Date).AddDays(-2).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

	$endpointUrl = "https://graph.microsoft.com/v1.0/auditLogs/" + `
							"directoryAudits?`$filter=activityDateTime ge $dateWindow"

	# To get all logs use:
	#$endpointUrl = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits"  

	$response = Invoke-RestMethod -Method Get `
						-Headers $myHeader `
						-Uri $endpointUrl `
						-ContentType "application/json;odata=verbose"

	$responseJson = $response | ConvertTo-Json -Depth 10
	Write-Host $responseJson
}
#gavdcodeend 004

#gavdcodebegin 005
function PsPnPPowerShellPurview_GetPurviewAuditLogs
{
	# Required rights: Microsoft Office 365 Management API: 
	#	ActivityFeed.Read, Microsoft Office 365 Management API: ActivityFeed.ReadDlp
	
	PsPnPPowerShell_LoginGraphWithCertificateThumbprint -SiteBaseUrl $cnfSiteBaseUrl `
				-TenantName $cnfTenantName -ClientId $cnfClientIdWithCert `
				-CertificateThumbprint $cnfCertificateThumbprint
	
	$myStartTime = `
			(Get-Date).AddHours(-23).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
	$myEndTime = `
			(Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
		
	$myResults = Get-PnPUnifiedAuditLog -ContentType "SharePoint" `
				-StartTime $myStartTime `
				-EndTime $myEndTime

	foreach ($oneResult in $myResults) {
		Write-Output "CreationTime: $($oneResult.CreationTime), 
						Operation: $($oneResult.Operation), 
						ClientIP: $($oneResult.ClientIP)"
	}
	
	Disconnect-PnPOnline
}
#gavdcodeend 005

#gavdcodebegin 006
function PsCliM365Purview_GetPurviewAuditLogs
{
	PsCliM365_LoginWithCertificateFile -TenantName $cnfTenantName `
								-ClientId $cnfClientIdWithCert `
								-CertificateFilePath $cnfCertificateFilePath `
								-CertificateFilePw $cnfCertificateFilePw
		
	$myStartTime = (Get-Date).AddHours(-23).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
	$myEndTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
						  
	$m365Results = m365 purview auditlog list --contentType "SharePoint" `
									--startTime $myStartTime `
									--endTime $myEndTime --debug | ConvertFrom-Json

	$m365Results | ForEach-Object {
		[PSCustomObject]@{
			CreationTime = $_.CreationTime
			Operation    = $_.Operation
			ClientIP     = $_.ClientIP
		}
	} | Format-Table -AutoSize
	
	m365 logout
}
#gavdcodeend 006


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 006 ***

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

#PsGraphRestApiMs_GetUnifiedAuditLogs
#PsGraphRestApiMs_CheckUnifiedAuditLogs
#PsGraphRestApiMs_ReadUnifiedAuditLogs

#PsGraphRestApiPurview_GetPurviewAuditLogs
#PsPnPPowerShellPurview_GetPurviewAuditLogs
#PsCliM365Purview_GetPurviewAuditLogs

Write-Host "Done" 

