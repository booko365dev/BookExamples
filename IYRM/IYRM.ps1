Function LoginAdminCsom()  #*** USE POWERSHELL 5.x, NOT 7.x ***
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

Function LoginPsSPO()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.SiteAdminUrl -Credential $myCredentials
}

#----------------------------------------------------------------------------------------

Function Invoke-RestSPO()  #*** LEGACY CODE ***
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
 
Function Get-SPOContextInfo()  #*** LEGACY CODE ***
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
	Invoke-RestSPO $Url Post $UserName $Password
}

Function Stream-CopyTo([System.IO.Stream]$Source, [System.IO.Stream]$Destination)   #*** LEGACY CODE ***
{
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}

Function LoginPsPnPPowerShellWithAccPwDefault
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}
#----------------------------------------------------------------------------------------

#gavdcodebegin 001
Function SpPsCsom_GetPropertiesTenant($spAdminCtx)  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)

    foreach ($oneProperty in $myTenant.GetType().GetProperties()) {
        Write-Host($oneProperty.Name)
    }
}
#gavdcodeend 001

#gavdcodebegin 002
Function SpPsCsom_GetValuePropertyTenant($spAdminCtx)  #*** USE POWERSHELL 5.x, NOT 7.x ***
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
Function SpPsCsom_UpdateValuePropertyTenant($spAdminCtx)#*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)

    $myTenant.BlockAccessOnUnmanagedDevices = $false
    $myTenant.Update()
    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 003

#gavdcodebegin 004
Function SpPsRest_FindAppCatalog()  #*** LEGACY CODE ***
{
    $endpointUrl = $webBaseUrl + "/_api/SP_TenantSettings_Current"
	$contextInfo = Get-SPOContextInfo -WebUrl $webBaseUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}
#gavdcodeend 004

#gavdcodebegin 005
Function SpPsRest_FindTenantProps()  #*** LEGACY CODE ***
{
    $catalogUrl = $webBaseUrl + "/sites/appcatalog"
    $endpointUrl = $webBaseUrl + "/_api/web/GetStorageEntity('SomeKey')"
	$contextInfo = Get-SPOContextInfo -WebUrl $webBaseUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}
#gavdcodeend 005

#gavdcodebegin 104
Function SpPsRest_FindAppCatalogAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRest_FindTenantPropsAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

    $endpointUrl = $configFile.appsettings.SiteBaseUrl + "/sites/appcatalog/_api/web/AllProperties"
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
Function SpPsSpo_GetTenant()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Get-SPOTenant
	Disconnect-SPOService
}
#gavdcodeend 006

#gavdcodebegin 007
Function SpPsSpo_ModifyTenantProperties()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Set-SPOTenant -NoAccessRedirectUrl $configFile.appsettings.SiteBaseUrl
	Disconnect-SPOService
}
#gavdcodeend 007

#gavdcodebegin 008
Function SpPsSpo_GetTenantLogs()  #*** LEGACY CODE ***
{

	Get-SPOTenantLogEntry
	Disconnect-SPOService
}
#gavdcodeend 008

#gavdcodebegin 009
Function SpPsSpo_GetTenantLogsLastEntryTime()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	# This cmdlet has been changed by Microsoft in the latest release of the module, 
	#    and it is now reserved for internal Microsoft use.
	Get-SPOTenantLogLastAvailableTimeInUtc
	Disconnect-SPOService
}
#gavdcodeend 009

#gavdcodebegin 010
Function SpPsSpo_GetCdnEnabled()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Get-SPOTenantCdnEnabled -CdnType Public
	Disconnect-SPOService
}
#gavdcodeend 010

#gavdcodebegin 011
Function SpPsSpo_GetCdnOrigins()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Get-SPOTenantCdnOrigins -CdnType Public
	Disconnect-SPOService
}
#gavdcodeend 011

#gavdcodebegin 012
Function SpPsSpo_GetCdnPolicies()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Get-SPOTenantCdnPolicies -CdnType Public
	Disconnect-SPOService
}
#gavdcodeend 012

#gavdcodebegin 013
Function SpPsSpo_EnableCdn()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Set-SPOTenantCdnEnabled -CdnType public -Enable $false
	Disconnect-SPOService
}
#gavdcodeend 013

#gavdcodebegin 014
Function SpPsSpo_CdnPolicy()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Set-SPOTenantCdnPolicy -CdnType Public `
						   -PolicyType ExcludeRestrictedSiteClassifications `
						   -PolicyValue "Confidential,Restricted"
	Disconnect-SPOService
}
#gavdcodeend 014

#gavdcodebegin 015
Function SpPsSpo_AddCdn()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Add-SPOTenantCdnOrigin -CdnType Public `
						   -OriginUrl "/sites/[sitename]/[library]"
	Disconnect-SPOService
}
#gavdcodeend 015

#gavdcodebegin 016
Function SpPsSpo_RemoveCdn()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	Remove-SPOTenantCdnOrigin -CdnType Public `
							  -OriginUrl "/sites/[sitename]/[library]"
	Disconnect-SPOService
}
#gavdcodeend 016

#gavdcodebegin 017
Function SpPsSpo_SetKey()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$appCatUrl = $configFile.appsettings.SiteBaseUrl + "/sites/appcatalog"
	Set-SPOStorageEntity -Site $appCatUrl -Key "MyPropertyKey" `
						 -Value "ValueOfMyKey" -Description "This is my key" `
						 -Comments "Comments for my key"
	Disconnect-SPOService
}
#gavdcodeend 017

#gavdcodebegin 018
Function SpPsSpo_GetKey()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$appCatUrl = $configFile.appsettings.SiteBaseUrl + "sites/appcatalog"
	Get-SPOStorageEntity -Site $appCatUrl -Key "MyPropertyKey"
	Disconnect-SPOService
}
#gavdcodeend 018

#gavdcodebegin 019
Function SpPsSpo_DeleteKey()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$appCatUrl = $configFile.appsettings.SiteBaseUrl + "/sites/appcatalog"
	Remove-SPOStorageEntity -Site $appCatUrl -Key "MyPropertyKey"
	Disconnect-SPOService
}
#gavdcodeend 019

#-----------------------------------------------------------------------------------------

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.Client.Tenant.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

#$spAdminCtx = LoginAdminCsom
#SpPsCsom_GetPropertiesTenant $spAdminCtx
#SpPsCsom_GetValuePropertyTenant $spAdminCtx
#SpPsCsom_UpdateValuePropertyTenant $spAdminCtx

#$webBaseUrl = $configFile.appsettings.SiteBaseUrl
#$userName = $configFile.appsettings.UserName
#$password = $configFile.appsettings.UserPw
#SpPsRest_FindAppCatalog
#SpPsRest_FindTenantProps
#SpPsRest_FindAppCatalogAD
#SpPsRest_FindTenantPropsAD

#LoginPsSPO
#SpPsSpo_GetTenant
#SpPsSpo_ModifyTenantProperties
#SpPsSpo_GetTenantLogs
#SpPsSpo_GetTenantLogsLastEntryTime
#SpPsSpo_GetCdnEnabled
#SpPsSpo_GetCdnOrigins
#SpPsSpo_GetCdnPolicies
#SpPsSpo_EnableCdn
#SpPsSpo_CdnPolicy
#SpPsSpo_RemoveCdn
#SpPsSpo_SetKey
#SpPsSpo_GetKey
#SpPsSpo_DeleteKey

Write-Host "Done"