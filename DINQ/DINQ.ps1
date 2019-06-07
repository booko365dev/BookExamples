
Function ExecuteRestGetQuery($SiteBaseUrl, $BaseRestQuery)
{
    $SiteRestUrl = $SiteBaseUrl + $BaseRestQuery

    $myCredentials = GetCredentials

    $myWebClient = New-Object System.Net.WebClient
    $myWebClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
    $myWebClient.Credentials = $myCredentials
    $myWebClient.Headers.Add("ContentType", "application/json;odata=verbose")
    $myWebClient.Headers.Add("Accept", "application/json;odata=verbose")

    $myResult = $myWebClient.DownloadString($SiteRestUrl)

	$resultObject = $myResult | ConvertFrom-Json
	$resultValue = ConvertTo-Json $resultObject.d

	return $resultValue
}

Function ExecuteRestQuery($SiteBaseUrl, $BaseRestQuery, $PostRestQuery, $RequestType)
{
    if ($RequestType -eq [TypeRequest]::GET)
    { return ExecuteRestGetQuery $SiteBaseUrl $BaseRestQuery }

    $SiteRestUrl = $SiteBaseUrl + $BaseRestQuery;

    $myCookies = GetAuthCookies $SiteBaseUrl
    $myFormDigest = GetFormDigest $SiteBaseUrl $myCookies

    $myWebReq = GetRequest $SiteRestUrl $myCookies $PostRestQuery.Length
    $myWebReq.Headers.Add("X-RequestDigest", $myFormDigest)

	if($RequestType -eq [TypeRequest]::MERGE) {
		$myWebReq.Headers.Add("IF-MATCH", "*")
		$myWebReq.Headers.Add("X-HTTP-Method", "MERGE")
	}
	elseif($RequestType -eq [TypeRequest]::DELETE) {
		$myWebReq.Headers.Add("IF-MATCH", "*")
		$myWebReq.Headers.Add("X-HTTP-Method", "DELETE")
	}

    $myReqStream = New-Object System.IO.StreamWriter($myWebReq.GetRequestStream())
    $myReqStream.Write($PostRestQuery)
    $myReqStream.Flush()

    $resultJson = GetResult $myWebReq
    if ($resultJson -ne $null) { 
		$resultObject = $resultJson | ConvertFrom-Json
		$resultValue = $resultObject.d
		
		return $resultValue
	}
    else
		{ return string.Empty; }
}

Function GetAuthCookies($SiteBaseUrl)
{
    $myCredentials = GetCredentials

    $authCookie = $myCredentials.GetAuthenticationCookie($SiteBaseUrl)
    $myCookies = New-Object System.Net.CookieContainer
    $myCookies.SetCookies($SiteBaseUrl, $authCookie)

    return $myCookies
}

Function GetCredentials()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
		$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.spUserName, $securePW)

    return $myCredentials
}

Function GetFormDigest($SiteBaseUrl, $Cookies)
{
    $resourceUrl = $SiteBaseUrl + "/_api/contextinfo"
    $myWebReq = GetRequest $resourceUrl $Cookies 0

    $resultJson = GetResult $myWebReq

	$resultObject = $resultJson | ConvertFrom-Json
	$formDigestValue = $resultObject.d.GetContextWebInformation.FormDigestValue

    return $formDigestValue
}

Function GetRequest($ReqUrl, $Cookies, $ContentLenght)
{
    $myWebReq = [System.Net.WebRequest]::Create($ReqUrl)
    $myWebReq.CookieContainer = $Cookies
    $myWebReq.Method = "POST"
    $myWebReq.Accept = "application/json;odata=verbose"
    $myWebReq.ContentLength = $ContentLenght
    $myWebReq.ContentType = "application/json;odata=verbose"

    return $myWebReq
}

Function GetResult($WebRequest)
{
    $myResult = ""
    $myWebResp = $WebRequest.GetResponse()

	Using-Object ($myRespStream = `
				New-Object System.IO.StreamReader($myWebResp.GetResponseStream())) {
		$myResult = $myRespStream.ReadToEnd()
		return $myResult
	}
}

function Using-Object
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowEmptyCollection()]
        [AllowNull()]
        [Object]
        $InputObject,
 
        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock
    )
 
    try {
        . $ScriptBlock
    }
    finally {
        if ($null -ne $InputObject -and $InputObject -is [System.IDisposable]) {
            $InputObject.Dispose()
        }
    }
}

enum TypeRequest { GET 
				   POST 
				   MERGE 
				   DELETE }

#-----------------------------------------------------------------------------------------

Function SpPsRestCreateOneList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
	$BaseRestQuery = "/_api/web/lists"
	$PostRestQuery = "{ '__metadata': { 'type': 'SP.List' }, " +
                            "'Title': 'NewListPowerShellRest', " +
                            "'BaseTemplate': 100, " +
                            "'Description': 'New List created using PowerShell REST' }"
	$TypeRequest = [TypeRequest]::POST

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestReadeAllLists()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists?$select=Title,Id";
	$PostRestQuery = ""
	$TypeRequest = [TypeRequest]::GET

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestReadeOneList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')"
	$PostRestQuery = ""
	$TypeRequest = [TypeRequest]::GET

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestUpdateOneList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')"
    $PostRestQuery = "{ '__metadata': { 'type': 'SP.List' }, " +
                            "'Description': 'New List Description' }"
	$TypeRequest = [TypeRequest]::MERGE

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestDeleteOneList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')"
    $PostRestQuery = $null
	$TypeRequest = [TypeRequest]::DELETE

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestAddOneFieldToList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')/fields"
    $PostRestQuery = "{ '__metadata': { 'type': 'SP.Field' }, " +
                            "'Title': 'MyMultilineField', " +
                            "'FieldTypeKind': 3 }"
	$TypeRequest = [TypeRequest]::POST

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestReadAllFieldsFromList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')/fields"
    $PostRestQuery = $null
	$TypeRequest = [TypeRequest]::GET

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestReadOneFieldFromList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')/fields/" +
                                "getbytitle('MyMultilineField')"
    $PostRestQuery = $null
	$TypeRequest = [TypeRequest]::GET

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestUpdateOneFieldInList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')/fields/" +
        "                               getbytitle('MyMultilineField')"
    $PostRestQuery = "{ '__metadata': { 'type': 'SP.Field' }, " +
                            "'Description': 'New Field Description' }"
	$TypeRequest = [TypeRequest]::MERGE

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestDeleteOneFieldFromList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')/fields/" +
										"getbytitle('MyMultilineField')"
    $PostRestQuery = $null
	$TypeRequest = [TypeRequest]::DELETE

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SppsRestBreakSecurityInheritanceList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')/" +
				"breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)"
    $PostRestQuery = $null
	$TypeRequest = [TypeRequest]::POST

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestResetSecurityInheritanceList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $BaseRestQuery = "/_api/lists/getbytitle('NewListPowerShellRest')/" +
        "resetroleinheritance"
    $PostRestQuery = $null
	$TypeRequest = [TypeRequest]::POST

	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestAddUserToSecurityRoleInList()
{
	# Inheritance MUST be broken
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $PostRestQuery = $null
    $BaseRestQuery = $null
    $ResponseResult = $null
	$TypeRequest = $null
    $resultJson = $null

    # Find the User
    $BaseRestQuery = "/_api/web/siteusers?$select=Id&" +
                                "$filter=startswith(Title,'MOD')"
	$TypeRequest = [TypeRequest]::GET
	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest
	$resultObject = $ResponseResult | ConvertFrom-Json
    $userId = $resultObject.results[0].Id

    # Find the RoleDefinitions
    $BaseRestQuery = "/_api/web/roledefinitions?$select=Id&" +
                                "$filter=startswith(Name,'Full Control')";
	$TypeRequest = [TypeRequest]::GET
	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest
	$resultObject = $ResponseResult | ConvertFrom-Json
    $roleId = $resultObject.results[0].Id

    # Add the User in the RoleDefinion to the List
    $BaseRestQuery = "/_api/web/lists/getbytitle('NewListPowerShellRest')/roleassignments/" +
        "addroleassignment(principalid=" + $userId + ",roledefid=" + $roleId + ")"
	$TypeRequest = [TypeRequest]::POST
	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestUpdateUserSecurityRoleInList()
{
	# Inheritance MUST be broken
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $PostRestQuery = $null
    $BaseRestQuery = $null
    $ResponseResult = $null
	$TypeRequest = $null
    $resultJson = $null

    # Find the User
    $BaseRestQuery = "/_api/web/siteusers?$select=Id&" +
                                "$filter=startswith(Title,'MOD')"
	$TypeRequest = [TypeRequest]::GET
	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest
	$resultObject = $ResponseResult | ConvertFrom-Json
    $userId = $resultObject.results[0].Id

    # Find the RoleDefinitions
    $BaseRestQuery = "/_api/web/roledefinitions/getbyname('Edit')/id"
	$TypeRequest = [TypeRequest]::GET
	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest
	$resultObject = $ResponseResult | ConvertFrom-Json
    $roleId = $resultObject.Id

    # Add the User in the RoleDefinion to the List
    $BaseRestQuery = "/_api/web/lists/getbytitle('NewListPowerShellRest')/roleassignments/" +
        "addroleassignment(principalid=" + $userId + ",roledefid=" + $roleId + ")"
	$TypeRequest = [TypeRequest]::MERGE
	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

Function SpPsRestDeleteUserFromSecurityRoleInList()
{
	$SiteBaseUrl = $configFile.appsettings.spUrl
    $PostRestQuery = $null
    $BaseRestQuery = $null
    $ResponseResult = $null
	$TypeRequest = $null
    $resultJson = $null

    # Find the User
    $BaseRestQuery = "/_api/web/siteusers?$select=Id&" +
                                "$filter=startswith(Title,'MOD')"
	$TypeRequest = [TypeRequest]::GET
	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest
	$resultObject = $ResponseResult | ConvertFrom-Json
    $userId = $resultObject.results[0].Id

    # Remove the User from the List
    $BaseRestQuery = "/_api/web/lists/GetByTitle('NewListPowerShellRest')/roleassignments/" +
        "getbyprincipalid(principalid=" + $userId + ")";
	$TypeRequest = [TypeRequest]::DELETE
	$ResponseResult = ExecuteRestQuery $SiteBaseUrl `
                                         $BaseRestQuery `
                                         $PostRestQuery `
                                         $TypeRequest

	Write-Host $ResponseResult
}

#-----------------------------------------------------------------------------------------

## Running the Functions
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

#SpPsRestCreateOneList
#SpPsRestReadeAllLists
#SpPsRestReadeOneList
#SpPsRestUpdateOneList
#SpPsRestDeleteOneList
#SpPsRestAddOneFieldToList
#SpPsRestReadAllFieldsFromList
#SpPsRestReadOneFieldFromList
#SpPsRestUpdateOneFieldInList
#SpPsRestDeleteOneFieldFromList
#SpPsRestBreakSecurityInheritanceList
#SpPsRestResetSecurityInheritanceList
#SpPsRestAddUserToSecurityRoleInList
#SpPsRestUpdateUserSecurityRoleInList
#SpPsRestDeleteUserFromSecurityRoleInList

Write-Host "Done" 

