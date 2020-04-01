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

#gavdcodebegin 01
Function SpPsRestCreateOneCommunicationSiteCollection()
{
	$endpointUrl = $webBaseUrl + "/_api/sitepages/communicationsite/create"
	$myPayload = 
		"{
			'request':  {
				'lcid': '1033',
				'Title': 'NewWebSiteModernPsRest01',
				'Url': '" + $webBaseUrl + "/sites/NewWebSiteModernPsRest01',
				'Description': 'NewWebSiteModernPsRest01 Description',
				'__metadata': {
								'type': 'SP.Publishing.CommunicationSiteCreationRequest'
							},
				'SiteDesignId': '6142d2a0-63a5-4ba0-aede-d9fefca2c767',
				'Classification': '',
				'AllowFileSharingForGuestUsers': 'false'
			}
		}"
	$contextInfo = Get-SPOContextInfo -WebUrl $webBaseUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 01 

#gavdcodebegin 02
Function SpPsRestCreateOneSiteCollection()
{
	$endpointUrl = $webBaseUrl + "/_api/SPSiteManager/create"
	$myPayload = 
		"{
		'request':  {
			'__metadata': { 'type': 'Microsoft.SharePoint.Portal.SPSiteCreationRequest' },
			'Title': 'NewWebSiteModernPsRest02',
			'Lcid': 1033,
			'Description': 'NewWebSiteModernPsRest02 Description',
			'Classification': '',
			'ShareByEmailEnabled': false,
			'SiteDesignId': '00000000-0000-0000-0000-000000000000',
			'Url': '" + $webBaseUrl + "/sites/NewWebSiteModernPsRest02',
			'WebTemplate': 'SITEPAGEPUBLISHING#0',
			'WebTemplateExtensionId': '00000000-0000-0000-0000-000000000000' 
			}
		}"
	$contextInfo = Get-SPOContextInfo -WebUrl $webBaseUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsRestCreateOneWebInSiteCollection()
{
	$endpointUrl = $webUrl + "/_api/web/webs/add"
	$myPayload = @{
			request = @{
			__metadata = @{ 'type' = 'SP.ListItem' }
			Title = 'NewWebSiteModernPsRest'
            Description = 'NewWebSiteModernPsRest Description'
            Url = 'NewWebSiteModernPsRest'
            UseSamePermissionsAsParentSite = 'true'
            WebTemplate = 'STS#3'
			}} | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpCsRestReadAllSiteCollections()
{
    $endpointUrl = $webBaseUrl +
                    "/_api/search/query?querytext='contentclass:sts_site'"
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpCsRestReadAllWebsInSiteCollection()
{
	$endpointUrl = $webUrl + "/_api/web/webs"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload 
    $data | ConvertTo-Json
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpCsRestUpdateOneWeb()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"
    $myPayload = @{
        __metadata = @{ type = "SP.Web" }
        Description = "NewWebSiteModernPsRest Description Updated"
		} | ConvertTo-Json
    $endpointUrl = $subWebUrl + "/_api/web"
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue -ETag "*" `
						-XHTTPMethod "MERGE"
	$data | ConvertTo-Json
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpCsRestDeleteOneWebFromSiteCollection()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web";
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue -ETag "*" `
						-XHTTPMethod "DELETE"
	$data | ConvertTo-Json
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpCsRestGetRoleDefinitionsWeb()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web/roledefinitions"
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue
	$data | ConvertTo-Json
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpCsRestFindUserPermissionsWeb()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web/" +
                                "doesuserhavepermissions(@v)?@v=" +
                                "{'High':'2147483647', 'Low':'4294967295'}"
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue
	$data | ConvertTo-Json
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpCsRestFindOtherUserPermissionsWeb()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web/" +
                                "getusereffectivepermissions(@v)?@v=" +
                                "'i%3A0%23.f%7Cmembership%7C" + $userName + "'"
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue
	$data | ConvertTo-Json
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpCsRestBreakSecurityInheritanceWeb()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web" +
                    "/breakroleinheritance(copyRoleAssignments=false," +
                    "clearSubscopes=true)"
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue -ETag "*" `
						-XHTTPMethod "MERGE"
	$data | ConvertTo-Json
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpCsRestResetSecurityInheritanceWeb()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web/resetroleinheritance"
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue
	$data | ConvertTo-Json
}
#gavdcodeend 12

#gavdcodebegin 13
Function SpCsRestAddUserToSecurityRoleInWeb()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"

	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $subWebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Find the RoleDefinitions
	$endpointUrl = $subWebUrl + "/_api/web/roledefinitions?$select=Id&" +
	                                "$filter=startswith(Name,'Full Control')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $roleId = $data.results[0].Id
	$data | ConvertTo-Json

    # Add the User in the RoleDefinion to the List
	$endpointUrl = $subWebUrl + "/_api/web/" +
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + 
			$roleId + ")"
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
							-Password $password -RequestDigest `
							$contextInfo.GetContextWebInformation.FormDigestValue `
							-ETag "*" -XHTTPMethod "MERGE"
	$data | ConvertTo-Json
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpCsRestUpdateUserSecurityRoleInWeb()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"

	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $subWebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Find the RoleDefinitions
	$endpointUrl = $subWebUrl + "/_api/web/roledefinitions?$select=Id&" +
	                                "$filter=startswith(Name,'Full Control')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $roleId = $data.results[0].Id
	$data | ConvertTo-Json

    # Add the User in the RoleDefinion to the List
	$endpointUrl = $subWebUrl + "/_api/web/" +
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + 
			$roleId + ")"
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "MERGE"
	$data | ConvertTo-Json
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpCsRestDeleteUserFromSecurityRoleInWeb()
{
    $subWebUrl = $webUrl + "/NewWebSiteModernPsRest"

    # Find the User
	$endpointUrl = $subWebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Remove the User from the List
	$endpointUrl = $subWebUrl + "/_api/web/" +
					     "roleassignments/getbyprincipalid(principalid=" + $userId + ")";
	$contextInfo = Get-SPOContextInfo -WebUrl $subWebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "DELETE"
	$data | ConvertTo-Json
}
#gavdcodeend 15

#----------------------------------------------------------------------------------------

## Running the Functions
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

$webUrl = $configFile.appsettings.spUrl
$webBaseUrl = $configFile.appsettings.spBaseUrl
$userName = $configFile.appsettings.spUserName
$password = $configFile.appsettings.spUserPw

#SpPsRestCreateOneCommunicationSiteCollection
#SpPsRestCreateOneSiteCollection
#SpPsRestCreateOneWebInSiteCollection
#SpCsRestReadAllSiteCollections
#SpCsRestReadAllWebsInSiteCollection
#SpCsRestUpdateOneWeb
#SpCsRestDeleteOneWebFromSiteCollection
#SpCsRestGetRoleDefinitionsWeb
#SpCsRestFindUserPermissionsWeb
#SpCsRestFindOtherUserPermissionsWeb
#SpCsRestBreakSecurityInheritanceWeb
#SpCsRestResetSecurityInheritanceWeb
#SpCsRestAddUserToSecurityRoleInWeb
#SpCsRestUpdateUserSecurityRoleInWeb
#SpCsRestDeleteUserFromSecurityRoleInWeb

Write-Host "Done" 
