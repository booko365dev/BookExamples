Function LoginAdminCsom()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.spUserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.spAdminUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

#----------------------------------------------------------------------------------------

Function LoginPsSPO()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.spAdminUrl -Credential $myCredentials
}

#----------------------------------------------------------------------------------------

Function Invoke-RestSPO() {
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
 
Function Get-SPOContextInfo(){
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

Function Stream-CopyTo([System.IO.Stream]$Source, [System.IO.Stream]$Destination) {
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}

#----------------------------------------------------------------------------------------

Function SpPsCsomGetPropertiesTenant($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)

    foreach ($oneProperty in $myTenant.GetType().GetProperties()) {
        Write-Host($oneProperty.Name)
    }
}

Function SpPsCsomGetValuePropertyTenant($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)

    $spAdminCtx.Load($myTenant)
    $spAdminCtx.ExecuteQuery()

    $myAccessDevices = $myTenant.BlockAccessOnUnmanagedDevices
    Write-Host($myAccessDevices)
}

Function SpPsCsomUpdateValuePropertyTenant($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)

    $myTenant.BlockAccessOnUnmanagedDevices = $false
    $myTenant.Update()
    $spAdminCtx.ExecuteQuery()
}

Function SpPsRestFindAppCatalog()
{
    $endpointUrl = $webBaseUrl + "/_api/SP_TenantSettings_Current"
	$contextInfo = Get-SPOContextInfo -WebUrl $webBaseUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}

Function SpPsRestFindTenantProps()
{
    $catalogUrl = $webBaseUrl + "/sites/appcatalog"
    $endpointUrl = $webBaseUrl + "/_api/web/GetStorageEntity('SomeKey')"
	$contextInfo = Get-SPOContextInfo -WebUrl $webBaseUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}

Function SpPsSpoGetTenant()
{
	Get-SPOTenant
	Disconnect-SPOService
}

Function SpPsSpoModifyTenantProperties()
{
	Set-SPOTenant -NoAccessRedirectUrl $configFile.appsettings.spBaseUrl
	Disconnect-SPOService
}

Function SpPsSpoGetTenantLogs()
{
	Get-SPOTenantLogEntry
	Disconnect-SPOService
}

Function SpPsSpoGetTenantLogsLastEntryTime()
{
	Get-SPOTenantLogLastAvailableTimeInUtc
	Disconnect-SPOService
}

Function SpPsSpoGetCdnEnabled()
{
	Get-SPOTenantCdnEnabled -CdnType Public
	Disconnect-SPOService
}

Function SpPsSpoGetCdnOrigins()
{
	Get-SPOTenantCdnOrigins -CdnType Public
	Disconnect-SPOService
}

Function SpPsSpoGetCdnPolicies()
{
	Get-SPOTenantCdnPolicies -CdnType Public
	Disconnect-SPOService
}

Function SpPsSpoEnableCdn()
{
	Set-SPOTenantCdnEnabled -CdnType public -Enable $false
	Disconnect-SPOService
}

Function SpPsSpoCdnPolicy()
{
	Set-SPOTenantCdnPolicy -CdnType Public `
						   -PolicyType ExcludeRestrictedSiteClassifications `
						   -PolicyValue "Confidential,Restricted"
	Disconnect-SPOService
}

Function SpPsSpoAddCdn()
{
	Add-SPOTenantCdnOrigin -CdnType Public `
						   -OriginUrl "/sites/[sitename]/[library]"
	Disconnect-SPOService
}

Function SpPsSpoRemoveCdn()
{
	Remove-SPOTenantCdnOrigin -CdnType Public `
							  -OriginUrl "/sites/[sitename]/[library]"
	Disconnect-SPOService
}

Function SpPsSpoSetKey()
{
	$appCatUrl = $configFile.appsettings.spBaseUrl + "/sites/appcatalog"
	Set-SPOStorageEntity -Site $appCatUrl -Key "MyPropertyKey" `
						 -Value "ValueOfMyKey" -Description "This is my key" `
						 -Comments "Comments for my key"
	Disconnect-SPOService
}

Function SpPsSpoGetKey()
{
	$appCatUrl = $configFile.appsettings.spBaseUrl + "/sites/appcatalog"
	Get-SPOStorageEntity -Site $appCatUrl -Key "MyPropertyKey"
	Disconnect-SPOService
}

Function SpPsSpoDeleteKey()
{
	$appCatUrl = $configFile.appsettings.spBaseUrl + "/sites/appcatalog"
	Remove-SPOStorageEntity -Site $appCatUrl -Key "MyPropertyKey"
	Disconnect-SPOService
}

#-----------------------------------------------------------------------------------------

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.Client.Tenant.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

#$spAdminCtx = LoginAdminCsom
#SpPsCsomGetPropertiesTenant $spAdminCtx
#SpPsCsomGetValuePropertyTenant $spAdminCtx
#SpPsCsomUpdateValuePropertyTenant $spAdminCtx

#$webBaseUrl = $configFile.appsettings.spBaseUrl
#$userName = $configFile.appsettings.spUserName
#$password = $configFile.appsettings.spUserPw
#SpPsRestFindAppCatalog
#SpPsRestFindTenantProps

#LoginPsSPO
#SpPsSpoGetTenant
#SpPsSpoModifyTenantProperties
#SpPsSpoGetTenantLogs
#SpPsSpoGetTenantLogsLastEntryTime
#SpPsSpoGetCdnEnabled
#SpPsSpoGetCdnOrigins
#SpPsSpoGetCdnPolicies
#SpPsSpoEnableCdn
#SpPsSpoCdnPolicy
#SpPsSpoRemoveCdn
#SpPsSpoSetKey
#SpPsSpoGetKey
#SpPsSpoDeleteKey

Write-Host "Done"
