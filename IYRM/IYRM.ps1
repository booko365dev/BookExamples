
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------


function PsSpCsom_LoginAdmin  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.UserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.SiteAdminUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

#----------------------------------------------------------------------------------------

function PsSpSpo_Login  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.SiteAdminUrl -Credential $myCredentials
}

#----------------------------------------------------------------------------------------

function PsSpSpo_InvokeRest  #*** LEGACY CODE ***
{
	Param (
		[Parameter(Mandatory=$True)]
		[String]$Url,
 
		[Parameter(Mandatory=$False)]
		[Microsoft.PowerShell.Commands.WebRequestMethod]$Method = `
								[Microsoft.PowerShell.Commands.WebRequestMethod]::Get,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$False)]
		[String]$Password,
 
		[Parameter(Mandatory=$False)]
		[String]$Metadata,

		[Parameter(Mandatory=$False)]
		[System.Byte[]]$Body,
 
		[Parameter(Mandatory=$False)]
		[String]$RequestDigest,
 
		[Parameter(Mandatory=$False)]
		[String]$ETag,
 
		[Parameter(Mandatory=$False)]
		[String]$XHTTPMethod,

		[Parameter(Mandatory=$False)]
		[System.String]$Accept = "application/json;odata=verbose",

		[Parameter(Mandatory=$False)]
		[String]$ContentType = "application/json;odata=verbose",

		[Parameter(Mandatory=$False)]
		[Boolean]$BinaryStringResponseBody = $False
	)
 
	if([string]::IsNullOrEmpty($Password)) {
		$SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString 
	}
	else {
		$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
	}
 
	$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials(`
															$UserName, $SecurePassword)
	$request = [System.Net.WebRequest]::Create($Url)
	$request.Credentials = $credentials
	$request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
	$request.ContentType = $ContentType
	$request.Accept = $Accept
	$request.Method=$Method
 
	if($RequestDigest) { 
		$request.Headers.Add("X-RequestDigest", $RequestDigest)
	}
	if($ETag) { 
		$request.Headers.Add("If-Match", $ETag)
	}
	if($XHTTPMethod) { 
		$request.Headers.Add("X-HTTP-Method", $XHTTPMethod)
	}
	if($Metadata -or $Body) {
		if($Metadata) {     
			$Body = [byte[]][char[]]$Metadata
		}      
		$request.ContentLength = $Body.Length 
		$stream = $request.GetRequestStream()
		$stream.Write($Body, 0, $Body.Length)
	}
	else {
		$request.ContentLength = 0
	}

	$response = $request.GetResponse()
	try {
		if($BinaryStringResponseBody -eq $False) {    
			$streamReader = New-Object System.IO.StreamReader $response.GetResponseStream()
			try {
				$data=$streamReader.ReadToEnd()
				$results = $data.ToString().Replace("ID", "_ID") | ConvertFrom-Json
				$results.d 
			}
			finally {
				$streamReader.Dispose()
			}
		}
		else {
			$dataStream = New-Object System.IO.MemoryStream
			try {
			Stream-CopyTo -Source $response.GetResponseStream() -Destination $dataStream
			$dataStream.ToArray()
			}
			finally {
				$dataStream.Dispose()
			} 
		}
	}
	finally {
		$response.Dispose()
	}
}
 
function PsSpSpo_GetContextInfo  #*** LEGACY CODE ***
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$WebUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$False)]
		[String]$Password
	)
   
	$Url = $WebUrl + "/_api/contextinfo"
	PsSpSpo_InvokeRest $Url Post $UserName $Password
}

function Stream-CopyTo([System.IO.Stream]$Source, [System.IO.Stream]$Destination)   #*** LEGACY CODE ***
{
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}

function PsSpPnP_LoginWithAccPw
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithAccPw `
					  -Credentials $myCredentials
}
#----------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpCsom_GetPropertiesTenant($spAdminCtx)  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)

    foreach ($oneProperty in $myTenant.GetType().GetProperties()) {
        Write-Host($oneProperty.Name)
    }
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpCsom_GetValuePropertyTenant($spAdminCtx)  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)

    $spAdminCtx.Load($myTenant)
    $spAdminCtx.ExecuteQuery()

    $myAccessDevices = $myTenant.BlockAccessOnUnmanagedDevices
    Write-Host($myAccessDevices)
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpCsom_UpdateValuePropertyTenant($spAdminCtx)#*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)

    $myTenant.BlockAccessOnUnmanagedDevices = $false
    $myTenant.Update()
    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpRest_FindAppCatalog  #*** LEGACY CODE ***
{
    $endpointUrl = $webBaseUrl + "/_api/SP_TenantSettings_Current"
	$contextInfo = PsSpSpo_GetContextInfo -WebUrl $webBaseUrl -UserName $userName `
																-Password $password
    $data = PsSpSpo_InvokeRest -Url $endpointUrl -Method GET -UserName $userName `
				-Password  $password -RequestDigest `
				$contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpRest_FindTenantProps  #*** LEGACY CODE ***
{
    $catalogUrl = $webBaseUrl + "/sites/appcatalog"
    $endpointUrl = $webBaseUrl + "/_api/web/GetStorageEntity('SomeKey')"
	$contextInfo = PsSpSpo_GetContextInfo -WebUrl $webBaseUrl -UserName $userName `
																-Password $password
    $data = PsSpSpo_InvokeRest -Url $endpointUrl -Method GET -UserName $userName `
				-Password $password `
				-RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}
#gavdcodeend 005

#gavdcodebegin 104
function PsSpRest_FindAppCatalogAD
{
	PsSpPnP_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $endpointUrl = $configFile.appsettings.SiteBaseUrl + "/_api/SP_TenantSettings_Current"
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	Write-Host "Catalog Url - " $dataObject.d.CorporateCatalogUrl
}
#gavdcodeend 104

#gavdcodebegin 105
function PsSpRest_FindTenantPropsAD
{
	PsSpPnP_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $endpointUrl = $configFile.appsettings.SiteBaseUrl + `
									"/sites/appcatalog/_api/web/AllProperties"
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 105

#gavdcodebegin 006
function PsSpSpo_GetTenant  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Get-SPOTenant
	Disconnect-SPOService
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpSpo_ModifyTenantProperties  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Set-SPOTenant -NoAccessRedirectUrl $configFile.appsettings.SiteBaseUrl
	Disconnect-SPOService
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpSpo_GetTenantLogs  #*** LEGACY CODE ***
{

	Get-SPOTenantLogEntry
	Disconnect-SPOService
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpSpo_GetTenantLogsLastEntryTime  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	# This cmdlet has been changed by Microsoft in the latest release of the module, 
	#    and it is now reserved for internal Microsoft use.
	Get-SPOTenantLogLastAvailableTimeInUtc
	Disconnect-SPOService
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpSpo_GetCdnEnabled  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Get-SPOTenantCdnEnabled -CdnType Public
	Disconnect-SPOService
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpSpo_GetCdnOrigins  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Get-SPOTenantCdnOrigins -CdnType Public
	Disconnect-SPOService
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpSpo_GetCdnPolicies  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Get-SPOTenantCdnPolicies -CdnType Public
	Disconnect-SPOService
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpSpo_EnableCdn  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Set-SPOTenantCdnEnabled -CdnType public -Enable $false
	Disconnect-SPOService
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpSpo_CdnPolicy  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Set-SPOTenantCdnPolicy -CdnType Public `
						   -PolicyType ExcludeRestrictedSiteClassifications `
						   -PolicyValue "Confidential,Restricted"
	Disconnect-SPOService
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpSpo_AddCdn  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Add-SPOTenantCdnOrigin -CdnType Public `
						   -OriginUrl "/sites/[sitename]/[library]"
	Disconnect-SPOService
}
#gavdcodeend 015

#gavdcodebegin 016
function PsSpSpo_RemoveCdn  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Remove-SPOTenantCdnOrigin -CdnType Public `
							  -OriginUrl "/sites/[sitename]/[library]"
	Disconnect-SPOService
}
#gavdcodeend 016

#gavdcodebegin 017
function PsSpSpo_SetKey  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$appCatUrl = $configFile.appsettings.SiteBaseUrl + "/sites/appcatalog"
	Set-SPOStorageEntity -Site $appCatUrl -Key "MyPropertyKey" `
						 -Value "ValueOfMyKey" -Description "This is my key" `
						 -Comments "Comments for my key"
	Disconnect-SPOService
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpSpo_GetKey  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$appCatUrl = $configFile.appsettings.SiteBaseUrl + "sites/appcatalog"
	Get-SPOStorageEntity -Site $appCatUrl -Key "MyPropertyKey"
	Disconnect-SPOService
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpSpo_DeleteKey  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$appCatUrl = $configFile.appsettings.SiteBaseUrl + "/sites/appcatalog"
	Remove-SPOStorageEntity -Site $appCatUrl -Key "MyPropertyKey"
	Disconnect-SPOService
}
#gavdcodeend 019


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 105 ***

#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.Client.Tenant.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

#$spAdminCtx = PsSpCsom_LoginAdmin
#PsSpCsom_GetPropertiesTenant $spAdminCtx
#PsSpCsom_GetValuePropertyTenant $spAdminCtx
#PsSpCsom_UpdateValuePropertyTenant $spAdminCtx

#$webBaseUrl = $configFile.appsettings.SiteBaseUrl
#$userName = $configFile.appsettings.UserName
#$password = $configFile.appsettings.UserPw
#PsSpRest_FindAppCatalog
#PsSpRest_FindTenantProps
#PsSpRest_FindAppCatalogAD
#PsSpRest_FindTenantPropsAD

#PsSpSpo_Login
#PsSpSpo_GetTenant
#PsSpSpo_ModifyTenantProperties
#PsSpSpo_GetTenantLogs
#PsSpSpo_GetTenantLogsLastEntryTime
#PsSpSpo_GetCdnEnabled
#PsSpSpo_GetCdnOrigins
#PsSpSpo_GetCdnPolicies
#PsSpSpo_EnableCdn
#PsSpSpo_CdnPolicy
#PsSpSpo_RemoveCdn
#PsSpSpo_SetKey
#PsSpSpo_GetKey
#PsSpSpo_DeleteKey

Write-Host "Done"
