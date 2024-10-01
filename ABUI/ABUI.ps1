
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function Invoke-RestSPO {  #*** LEGACY CODE *** 
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
 
function Get-SPOContextInfo{  #*** LEGACY CODE *** 
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

function Stream-CopyTo([System.IO.Stream]$Source, [System.IO.Stream]$Destination) 
 #*** LEGACY CODE *** 
 {
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}

function PsSpPnpPowerShell_LoginWithAccPw
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithAccPw `
					  -Credentials $myCredentials
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

# Using Classic Authentication, this is Legacy code, and cannot be used anymore

#gavdcodebegin 001
function PsSpRest_CreateOneCommunicationSiteCollection  #*** LEGACY CODE *** 
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
#gavdcodeend 001 

#gavdcodebegin 002
function PsSpRest_CreateOneSiteCollection   #*** LEGACY CODE *** 
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
#gavdcodeend 002

#gavdcodebegin 003
function PsSpRest_CreateOneWebInSiteCollection   #*** LEGACY CODE *** 
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
#gavdcodeend 003

#gavdcodebegin 004
function PsSpRest_GetAllSiteCollections   #*** LEGACY CODE *** 
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
#gavdcodeend 004

#gavdcodebegin 005
function PsSpRest_GetAllWebsInSiteCollection   #*** LEGACY CODE *** 
{
	$endpointUrl = $webUrl + "/_api/web/webs"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload 
    $data | ConvertTo-Json
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpRest_UpdateOneWeb   #*** LEGACY CODE *** 
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
#gavdcodeend 006

#gavdcodebegin 007
function PsSpRest_DeleteOneWebFromSiteCollection   #*** LEGACY CODE *** 
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
#gavdcodeend 007

#gavdcodebegin 008
function PsSpRest_GetRoleDefinitionsWeb   #*** LEGACY CODE *** 
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
#gavdcodeend 008

#gavdcodebegin 009
function PsSpRest_GetUserPermissionsWeb   #*** LEGACY CODE *** 
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
#gavdcodeend 009

#gavdcodebegin 010
function PsSpRest_GetOtherUserPermissionsWeb   #*** LEGACY CODE *** 
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
#gavdcodeend 010

#gavdcodebegin 011
function PsSpRest_BreakSecurityInheritanceWeb   #*** LEGACY CODE *** 
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
#gavdcodeend 011

#gavdcodebegin 012
function PsSpRest_ResetSecurityInheritanceWeb   #*** LEGACY CODE *** 
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
#gavdcodeend 012

#gavdcodebegin 013
function PsSpRest_AddUserToSecurityRoleInWeb   #*** LEGACY CODE *** 
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
#gavdcodeend 013

#gavdcodebegin 014
function PsSpRest_UpdateUserSecurityRoleInWeb   #*** LEGACY CODE *** 
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
#gavdcodeend 014

#gavdcodebegin 015
function PsSpRest_DeleteUserFromSecurityRoleInWeb   #*** LEGACY CODE *** 
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
#gavdcodeend 015

#----------------------------------------------------------------------------------------

# Using Azure AD Authentication through Connect-PnPOnline and an Account/PW App Registration

#gavdcodebegin 101
function PsSpRest_CreateOneCommunicationSiteCollectionAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
											"/_api/sitepages/communicationsite/create"
	$myPayload = 
		"{
			'request':  {
				'lcid': '1033',
				'Title': 'NewWebSiteModernPsRest03',
				'Url': '" + $configFile.appsettings.SiteBaseUrl + `
														"/sites/NewWebSiteModernPsRest03',
				'Description': 'NewWebSiteModernPsRest03 Description',
				'__metadata': {
								'type': 'SP.Publishing.CommunicationSiteCreationRequest'
							},
				'SiteDesignId': '6142d2a0-63a5-4ba0-aede-d9fefca2c767',
				'Classification': '',
				'AllowFileSharingForGuestUsers': 'false'
			}
		}"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 101 

#gavdcodebegin 102
function PsSpRest_CreateOneSiteCollectionAD 
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + "/_api/SPSiteManager/create"
	$myPayload = 
		"{
		'request':  {
			'__metadata': { 'type': 'Microsoft.SharePoint.Portal.SPSiteCreationRequest' },
			'Title': 'NewWebSiteModernPsRest04',
			'Lcid': 1033,
			'Description': 'NewWebSiteModernPsRest04 Description',
			'Classification': '',
			'ShareByEmailEnabled': false,
			'SiteDesignId': '00000000-0000-0000-0000-000000000000',
			'Url': '" + $configFile.appsettings.SiteBaseUrl + "/sites/NewWebSiteModernPsRest04',
			'WebTemplate': 'SITEPAGEPUBLISHING#0',
			'WebTemplateExtensionId': '00000000-0000-0000-0000-000000000000' 
			}
		}"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 102

#gavdcodebegin 103
function PsSpRest_CreateOneWebInSiteCollectionAD 
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/webinfos/add"
	$myPayload = 
		"{
		'parameters':  {
			'__metadata': { 'type': 'SP.WebInfoCreationInformation' },
			'Title': 'NewWebSiteModernPsRest',
			'Description': 'NewWebSiteModernPsRest Description',
			'Url': 'NewWebSiteModernPsRest',
			'WebTemplate': 'STS#3',
			'UseUniquePermissions': 'true' 
			}
		}"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 103

#gavdcodebegin 104
function PsSpRest_GetAllSiteCollectionsAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl +
                "/_api/search/query?querytext='contentclass:sts_site'"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 104

#gavdcodebegin 105
function PsSpRest_GetAllWebsInSiteCollectionAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/webs"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 105

#gavdcodebegin 106
function PsSpRest_UpdateOneWebAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $subWebUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web"

	$myPayload = 
		"{
			'__metadata': { 'type': 'SP.Web' },
			'Description': 'NewWebSiteModernPsRest Description Updated'
		}"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 106

#gavdcodebegin 107
function PsSpRest_DeleteOneWebFromSiteCollectionAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $subWebUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 107

#gavdcodebegin 108
function PsSpRest_GetRoleDefinitionsSiteAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/roledefinitions"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 108

#gavdcodebegin 109
function PsSpRest_GetUserPermissionsSiteAD 
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/" +
                                "doesuserhavepermissions(@v)?@v=" +
                                "{'High':'2147483647', 'Low':'4294967295'}"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 109

#gavdcodebegin 110
function PsSpRest_GetOtherUserPermissionsSiteAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$userName = $configFile.appsettings.UserName
    $endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/" +
                                "getusereffectivepermissions(@v)?@v=" +
                                "'i%3A0%23.f%7Cmembership%7C" + $userName + "'"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 110

#gavdcodebegin 111
function PsSpRest_BreakSecurityInheritanceWebAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $subWebUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web" +
                    "/breakroleinheritance(copyRoleAssignments=false," +
                    "clearSubscopes=true)"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 111

#gavdcodebegin 112
function PsSpRest_ResetSecurityInheritanceWebAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $subWebUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsRest"
    $endpointUrl = $subWebUrl + "/_api/web/resetroleinheritance"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 112

#gavdcodebegin 113
function PsSpRest_AddUserToSecurityRoleInWebAD 
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	# For sub-Webs, inheritance MUST be broken
    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/siteusers?`$select=Id&" +
	                                "`$filter=startswith(Title,'Ade')"
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$userId = $dataObject.d.results[0].Id
	Write-Host "UserId - " $userId

    # Find the RoleDefinitions
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
										"/_api/web/roledefinitions?`$select=Id&" + `
										"`$filter=startswith(Name,'Full Control')"
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$roleId = $dataObject.d.results[0].Id
	Write-Host "RoleId - " $roleId

    # Add the User into the RoleDefinion to a Sub-Web
	$subWebUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsRest"
	$endpointUrl = $subWebUrl + "/_api/web/" + `
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + `
			$roleId + ")"
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 113

#gavdcodebegin 114
function PsSpRest_UpdateUserSecurityRoleInWebAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	# For sub-Webs, inheritance MUST be broken
    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/siteusers?`$select=Id&" +
	                                "`$filter=startswith(Title,'Ade')"
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$userId = $dataObject.d.results[0].Id
	Write-Host "UserId - " $userId

    # Find the RoleDefinitions
	$endpointUrl = $configFile.appsettings.SiteCollUrl + 
										"/_api/web/roledefinitions?`$select=Id&" + 
										"`$filter=startswith(Name,'Edit')"
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$roleId = $dataObject.d.results[0].Id
	Write-Host "RoleId - " $roleId

    # Update the User into the RoleDefinion to a Sub-Web
	$subWebUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsRest"
	$endpointUrl = $subWebUrl + "/_api/web/" + `
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + `
			$roleId + ")"
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 114

#gavdcodebegin 115
function PsSpRest_DeleteUserFromSecurityRoleInWebAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	# For sub-Webs, inheritance MUST be broken
    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/siteusers?`$select=Id&" +
	                                "`$filter=startswith(Title,'Ade')"
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$userId = $dataObject.d.results[0].Id
	Write-Host "UserId - " $userId

    # Remove the User from the Sub-Web
	$subWebUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsRest"
	$endpointUrl = $subWebUrl + "/_api/web/" +
					     "roleassignments/getbyprincipalid(principalid=" + $userId + ")";
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 115

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 115 ***

## Running the functions
[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

## Using Legacy Classic Authentication
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#$webUrl = $configFile.appsettings.SiteCollUrl
#$webBaseUrl = $configFile.appsettings.SiteBaseUrl
#$userName = $configFile.appsettings.UserName
#$password = $configFile.appsettings.UserPw
#PsSpRest_CreateOneCommunicationSiteCollection
#PsSpRest_CreateOneSiteCollection
#PsSpRest_CreateOneWebInSiteCollection
#PsSpRest_GetAllSiteCollections
#PsSpRest_GetAllWebsInSiteCollection
#PsSpRest_UpdateOneWeb
#PsSpRest_DeleteOneWebFromSiteCollection
#PsSpRest_GetRoleDefinitionsWeb
#PsSpRest_GetUserPermissionsWeb
#PsSpRest_GetOtherUserPermissionsWeb
#PsSpRest_BreakSecurityInheritanceWeb
#PsSpRest_ResetSecurityInheritanceWeb
#PsSpRest_AddUserToSecurityRoleInWeb
#PsSpRest_UpdateUserSecurityRoleInWeb
#PsSpRest_DeleteUserFromSecurityRoleInWeb

## Using Azure AD Authentication
#PsSpRest_CreateOneCommunicationSiteCollectionAD
#PsSpRest_CreateOneSiteCollectionAD
#PsSpRest_CreateOneWebInSiteCollectionAD
#PsSpRest_GetAllSiteCollectionsAD
#PsSpRest_GetAllWebsInSiteCollectionAD
#PsSpRest_UpdateOneWebAD
#PsSpRest_DeleteOneWebFromSiteCollectionAD
#PsSpRest_GetRoleDefinitionsSiteAD
#PsSpRest_GetUserPermissionsSiteAD
#PsSpRest_GetOtherUserPermissionsSiteAD
#PsSpRest_BreakSecurityInheritanceWebAD
#PsSpRest_ResetSecurityInheritanceWebAD
#PsSpRest_AddUserToSecurityRoleInWebAD
#PsSpRest_UpdateUserSecurityRoleInWebAD
#PsSpRest_DeleteUserFromSecurityRoleInWebAD

Write-Host "Done" 
