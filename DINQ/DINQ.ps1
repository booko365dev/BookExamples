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
				$results = $data | ConvertFrom-Json
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

Function SpPsRestCreateOneList()
{
	$endpointUrl = $WebUrl + "/_api/web/lists"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.List' }; 
				Title = 'NewListRestPs'; 
				BaseTemplate = 100; 
				Description = 'Test NewListRest'; 
				AllowContentTypes = $true;
				ContentTypesEnabled = $true
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}

Function SpPsRestReadAllLists()
{
	$endpointUrl = $WebUrl + "/_api/lists?$select=Title,Id"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password

	$data | ConvertTo-Json
}

Function SpPsRestReadOneList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password

	$data | ConvertTo-Json
}

Function SpPsRestUpdateOneList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.List' }; 
				Description = 'New List Description' 
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}

Function SpPsRestDeleteOneList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}

Function SpPsRestAddOneFieldToList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/fields"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.Field' }; 
				Title = 'MyMultilineField'; 
				FieldTypeKind = 3 
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}

Function SpPsRestReadAllFieldsFromList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/fields"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password

	$data | ConvertTo-Json
}

Function SpPsRestReadOneFieldFromList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/fields/" +
					                                "getbytitle('MyMultilineField')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password

	$data | ConvertTo-Json
}

Function SpPsRestUpdateOneFieldInList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/fields/" +
													"getbytitle('MyMultilineField')"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.Field' }; 
				Description = 'New Field Description' 
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}

Function SpPsRestDeleteOneFieldFromList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/fields/" +
													"getbytitle('MyMultilineField')"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}

Function SpPsRestBreakSecurityInheritanceList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/" +
				"breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}

Function SpPsRestResetSecurityInheritanceList()
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/" +
				"resetroleinheritance"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue

	$data | ConvertTo-Json
}

Function SpPsRestAddUserToSecurityRoleInList()
{
	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $WebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Find the RoleDefinitions
	$endpointUrl = $WebUrl + "/_api/web/roledefinitions?$select=Id&" +
	                                "$filter=startswith(Name,'Full Control')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $roleId = $data.results[0].Id
	$data | ConvertTo-Json

    # Add the User in the RoleDefinion to the List
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/roleassignments/" +
        "addroleassignment(principalid=" + $userId + ",roledefid=" + $roleId + ")"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
							-Password $password -RequestDigest `
							$contextInfo.GetContextWebInformation.FormDigestValue `
							-ETag "*" -XHTTPMethod "MERGE"
	$data | ConvertTo-Json
}

Function SpPsRestUpdateUserSecurityRoleInList()
{
	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $WebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Find the RoleDefinitions
	$endpointUrl = $WebUrl + "/_api/web/roledefinitions?$select=Id&" +
	                                "$filter=startswith(Name,'Full Control')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $roleId = $data.results[0].Id
	$data | ConvertTo-Json

    # Add the User in the RoleDefinion to the List
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/roleassignments/" +
        "addroleassignment(principalid=" + $userId + ",roledefid=" + $roleId + ")"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "MERGE"
	$data | ConvertTo-Json
}

Function SpPsRestDeleteUserFromSecurityRoleInList()
{
    # Find the User
	$endpointUrl = $WebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Remove the User from the List
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/roleassignments/" +
					        "getbyprincipalid(principalid=" + $userId + ")";
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "DELETE"
	$data | ConvertTo-Json
}

#----------------------------------------------------------------------------------------

## Running the Functions
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

$webUrl = $configFile.appsettings.spUrl
$userName = $configFile.appsettings.spUserName
$password = $configFile.appsettings.spUserPw

#SpPsRestCreateOneList
#SpPsRestReadAllLists
#SpPsRestReadOneList
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

