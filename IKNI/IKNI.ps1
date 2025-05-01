##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

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

function PsGraphPowerShellSdk_LoginWithCertificateThumbprint
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateThumbprint
	)

	Connect-MgGraph -TenantId $TenantName -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpCsom_GetChangeLogSite
{
	PsPnPPowerShell_LoginGraphWithCertificateThumbprint -SiteBaseUrl $cnfSiteCollUrl `
				-TenantName $cnfTenantName -ClientId $cnfClientIdWithCert `
				-CertificateThumbprint $cnfCertificateThumbprint

	$ctx = Get-PnPContext

	$changeQuery = New-Object Microsoft.SharePoint.Client.ChangeQuery
	$changeQuery.Site = $true
	$changeQuery.Add = $true
	$changeQuery.Update = $true
	$changeQuery.DeleteObject = $true
	$changeQuery.FetchLimit = 1000  # Optionally limit the number of changes returned

	$changesSite = $ctx.Site.GetChanges($changeQuery)
	$ctx.Load($changesSite)
	$ctx.ExecuteQuery()

	$startDate = (Get-Date).AddDays(-90)

	$filteredChanges = $changesSite | Where-Object {
		$_.PSObject.Properties['Time'] -and ($_.Time -ge $startDate)
	}

	$filteredChanges | Format-Table -Wrap -AutoSize
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpCsom_GetChangeLogWeb
{
	PsPnPPowerShell_LoginGraphWithCertificateThumbprint -SiteBaseUrl $cnfSiteCollUrl `
				-TenantName $cnfTenantName -ClientId $cnfClientIdWithCert `
				-CertificateThumbprint $cnfCertificateThumbprint

	$ctx = Get-PnPContext

	$changeQuery = New-Object Microsoft.SharePoint.Client.ChangeQuery
	$changeQuery.Web = $true
	$changeQuery.Add = $true
	$changeQuery.Update = $true
	$changeQuery.DeleteObject = $true
	$changeQuery.View = $true
	$changeQuery.FetchLimit = 1000  # Optionally limit the number of changes returned

	$changesWeb = $ctx.Web.GetChanges($changeQuery)
	$ctx.Load($changesWeb)
	$ctx.ExecuteQuery()

	$startDate = (Get-Date).AddDays(-90)

	$filteredChanges = $changesWeb | Where-Object {
		$_.PSObject.Properties['Time'] -and ($_.Time -ge $startDate)
	}

	$filteredChanges | Format-Table -Wrap -AutoSize
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpCsom_GetChangeLogList
{
	PsPnPPowerShell_LoginGraphWithCertificateThumbprint -SiteBaseUrl $cnfSiteCollUrl `
				-TenantName $cnfTenantName -ClientId $cnfClientIdWithCert `
				-CertificateThumbprint $cnfCertificateThumbprint

	$ctx = Get-PnPContext

	$changeQuery = New-Object Microsoft.SharePoint.Client.ChangeQuery
	$changeQuery.Add = $true
	$changeQuery.Update = $true
	$changeQuery.DeleteObject = $true
	$changeQuery.Item = $true
	$changeQuery.List = $true
	$changeQuery.FetchLimit = 1000  # Optionally limit the number of changes returned

	$list = Get-PnPList -Identity "TestList"
	$changes = $list.GetChanges($changeQuery)
	$ctx.Load($changes)
	$ctx.ExecuteQuery()

	$startDate = (Get-Date).AddDays(-90)

	$filteredChanges = $changes | Where-Object {
		$_.PSObject.Properties['Time'] -and ($_.Time -ge $startDate)
	}

	$filteredChanges | Format-Table -Wrap -AutoSize
}
#gavdcodeend 003

#gavdcodebegin 004
function PsGraphRestApi_GetTenantLogChanges
{
	$accessToken = PsGraphRestApi_GetAzureTokenWithCertificateThumbprint `
					-ClientID $cnfClientIdWithCert `
					-TenantName $cnfTenantName `
					-CertificateThumbprint $cnfCertificateThumbprint

	$myHeader = @{
		"Authorization" = "Bearer $accessToken"
	}

	$endpointUrl = "https://graph.microsoft.com/v1.0/sites/delta"

	$response = Invoke-RestMethod -Method Get `
						-Headers $myHeader `
						-Uri $endpointUrl `
						-ContentType "application/json;odata=verbose"

	Write-Output "=== Response tenantDelta - " ($response | ConvertTo-Json -Depth 10)
	$response.value | ForEach-Object {
		[PSCustomObject]@{
			Id              = $_.id
			Name            = $_.name
			CreatedDateTime = $_.createdDateTime
		}
	} | Format-Table -AutoSize
	Write-Output $response.'@odata.deltaLink'
}
#gavdcodeend 004

#gavdcodebegin 005
function PsGraphRestApi_GetSiteCollectionLogChanges
{
	$accessToken = PsGraphRestApi_GetAzureTokenWithCertificateThumbprint `
					-ClientID $cnfClientIdWithCert `
					-TenantName $cnfTenantName `
					-CertificateThumbprint $cnfCertificateThumbprint

	$myHeader = @{
	"Authorization" = "Bearer $accessToken"
	}

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$endpointUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root/delta"

	$response = Invoke-RestMethod -Method Get `
						-Headers $myHeader `
						-Uri $endpointUrl `
						-ContentType "application/json;odata=verbose"
		
	Write-Output "=== Response siteDelta - " ($response | ConvertTo-Json -Depth 10)
	$response.value | ForEach-Object {
		[PSCustomObject]@{
			Id              = $_.id
			Name            = $_.name
			CreatedDateTime = $_.createdDateTime
		}
	} | Format-Table -AutoSize
	Write-Output $response.'@odata.deltaLink'
}
#gavdcodeend 005

#gavdcodebegin 006
function PsGraphRestApi_GetListLogChanges
{
	$accessToken = PsGraphRestApi_GetAzureTokenWithCertificateThumbprint `
					-ClientID $cnfClientIdWithCert `
					-TenantName $cnfTenantName `
					-CertificateThumbprint $cnfCertificateThumbprint

	$myHeader = @{
	"Authorization" = "Bearer $accessToken"
	}

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"
	$endpointUrl = `
			"https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$ListId/items/delta"

	$response = Invoke-RestMethod -Method Get `
						-Headers $myHeader `
						-Uri $endpointUrl `
						-ContentType "application/json;odata=verbose"
			
	Write-Output "=== Response listDelta - " ($response | ConvertTo-Json -Depth 10)
	$response.value | ForEach-Object {
		[PSCustomObject]@{
			Id              	 = $_.id
			CreatedDateTime 	 = $_.createdDateTime
			LastModifiedDateTime = $_.lastModifiedDateTime
		}
	} | Format-Table -AutoSize
	Write-Output $response.'@odata.deltaLink'
}
#gavdcodeend 006

#gavdcodebegin 007
function PsGraphPowerShellSdk_GetTenantLogChanges
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint -TenantName $cnfTenantName `
					-ClientID $cnfClientIdWithCert `
					-CertificateThumbprint $cnfCertificateThumbprint

	$lastWeek = (Get-Date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ")

	$myDelta = Get-MgSiteDelta -Filter "CreatedDateTime ge $lastWeek"
	#$myDelta = Get-MgSiteDelta -Token "latest"

	Write-Output $myDelta

	$myDeltaJson = $myDelta | ConvertTo-Json -Depth 3
	Write-Output $myDeltaJson
			
	Disconnect-MgGraph
}
#gavdcodeend 007

#gavdcodebegin 008
function PsGraphPowerShellSdk_GetListLogChanges
{
	PsGraphPowerShellSdk_LoginWithCertificateThumbprint -TenantName $cnfTenantName `
					-ClientID $cnfClientIdWithCert `
					-CertificateThumbprint $cnfCertificateThumbprint
		
	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$ListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"
	$lastWeek = (Get-Date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ")

	$myDelta = Get-MgSiteListItemDelta -SiteId $SiteId `
										-ListId $ListId `
										-Filter "lastModifiedDateTime ge $lastWeek"
	#$myDelta = Get-MgSiteDelta -Token "latest"

	Write-Output $myDelta

	$myDeltaJson = $myDelta | ConvertTo-Json -Depth 3
	Write-Output $myDeltaJson
			
	Disconnect-MgGraph
}
#gavdcodeend 008

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 008 ***


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

#PsSpCsom_GetChangeLogSite
#PsSpCsom_GetChangeLogWeb
#PsSpCsom_GetChangeLogList

#PsGraphRestApi_GetTenantLogChanges
#PsGraphRestApi_GetSiteCollectionLogChanges
#PsGraphRestApi_GetListLogChanges

#PsGraphPowerShellSdk_GetTenantLogChanges
#PsGraphPowerShellSdk_GetListLogChanges

Write-Host "Done" 
