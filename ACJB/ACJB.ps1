
Function LoginPsCsom()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.spUserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.spUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

Function LoginPsPSO()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.spAdminUrl -Credential $myCredentials
}

Function LoginPsPnP()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.spUrl -Credentials $myCredentials
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

## Running the Functions
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

$spCtx = LoginPsCsom
$rootWeb = $spCtx.Web
$spCtx.Load($rootWeb)
$spCtx.ExecuteQuery()
Write-Host $rootWeb.Created  

LoginPsPSO
Test-SPOSite $configFile.appsettings.spUrl

LoginPsPnP
$myWeb = Get-PnPWeb
Write-Host $myWeb.Title
#Get-PnPTenant

$ResutlRest = ExecuteRestQuery $configFile.appsettings.spUrl "/_api/web/Created"
Write-Host $ResutlRest

Write-Host ""  

