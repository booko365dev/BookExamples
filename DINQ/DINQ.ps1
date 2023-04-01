Function Invoke-RestSPO 
{  #*** LEGACY CODE ***
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
 
Function Get-SPOContextInfo
{  #*** LEGACY CODE ***
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

Function Stream-CopyTo([System.IO.Stream]$Source, [System.IO.Stream]$Destination) 
{
	  #*** LEGACY CODE ***
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

# Using Basic Authentication, this is Legacy code, and cannot be used anymore

#gavdcodebegin 001
Function SpPsRest_CreateOneList  #*** LEGACY CODE ***
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
#gavdcodeend 001 

#gavdcodebegin 002
Function SpPsRest_ReadAllLists  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists?$select=Title,Id"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password

	$data | ConvertTo-Json
}
#gavdcodeend 002

#gavdcodebegin 003
Function SpPsRest_ReadOneList  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password

	$data | ConvertTo-Json
}
#gavdcodeend 003

#gavdcodebegin 004
Function SpPsRest_UpdateOneList  #*** LEGACY CODE ***
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
#gavdcodeend 004

#gavdcodebegin 005
Function SpPsRest_DeleteOneList  #*** LEGACY CODE ***
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
#gavdcodeend 005

#gavdcodebegin 006
Function SpPsRest_AddOneFieldToList  #*** LEGACY CODE ***
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
#gavdcodeend 006

#gavdcodebegin 007
Function SpPsRest_ReadAllFieldsFromList  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/fields"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password

	$data | ConvertTo-Json
}
#gavdcodeend 007

#gavdcodebegin 008
Function SpPsRest_ReadOneFieldFromList  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('NewListRestPs')/fields/" +
					                                "getbytitle('MyMultilineField')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password

	$data | ConvertTo-Json
}
#gavdcodeend 008

#gavdcodebegin 009
Function SpPsRest_UpdateOneFieldInList  #*** LEGACY CODE ***
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
#gavdcodeend 009

#gavdcodebegin 010
Function SpPsRest_DeleteOneFieldFromList  #*** LEGACY CODE ***
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
#gavdcodeend 010

#gavdcodebegin 011
Function SpPsRest_BreakSecurityInheritanceList  #*** LEGACY CODE ***
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
#gavdcodeend 011

#gavdcodebegin 012
Function SpPsRest_ResetSecurityInheritanceList  #*** LEGACY CODE ***
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
#gavdcodeend 012

#gavdcodebegin 013
Function SpPsRest_AddUserToSecurityRoleInList  #*** LEGACY CODE ***
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
#gavdcodeend 013

#gavdcodebegin 014
Function SpPsRest_UpdateUserSecurityRoleInList  #*** LEGACY CODE ***
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
#gavdcodeend 014

#gavdcodebegin 015
Function SpPsRest_DeleteUserFromSecurityRoleInList  #*** LEGACY CODE ***
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
#gavdcodeend 015

#----------------------------------------------------------------------------------------

# Using Azure AD Authentication through Connect-PnPOnline and an Account/PW App Registration

#gavdcodebegin 101
Function SpPsRest_CreateOneListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/lists"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.List' }; 
				Title = 'NewListRestPs'; 
				BaseTemplate = 100; 
				Description = 'Test NewListRest'; 
				AllowContentTypes = $true;
				ContentTypesEnabled = $true
			   } | ConvertTo-Json
	
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
Function SpPsRest_ReadAllListsAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/lists?$select=Title,Id"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 102

#gavdcodebegin 103
Function SpPsRest_ReadOneListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
												"/_api/lists/getbytitle('NewListRestPs')"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 103

#gavdcodebegin 104
Function SpPsRest_UpdateOneListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
												"/_api/lists/getbytitle('NewListRestPs')"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.List' }; 
				Description = 'New List Description' 
			   } | ConvertTo-Json
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 104

#gavdcodebegin 105
Function SpPsRest_DeleteOneListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
												"/_api/lists/getbytitle('NewListRestPs')"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 105

#gavdcodebegin 106
Function SpPsRest_AddOneFieldToListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
										"/_api/lists/getbytitle('NewListRestPs')/fields"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.Field' }; 
				Title = 'MyMultilineField'; 
				FieldTypeKind = 3 
			   } | ConvertTo-Json
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 106

#gavdcodebegin 107
Function SpPsRest_ReadAllFieldsFromListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
										"/_api/lists/getbytitle('NewListRestPs')/fields"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 107

#gavdcodebegin 108
Function SpPsRest_ReadOneFieldFromListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/lists/getbytitle('NewListRestPs')/fields/" + `
					                                "getbytitle('MyMultilineField')"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 108

#gavdcodebegin 109
Function SpPsRest_UpdateOneFieldInListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/lists/getbytitle('NewListRestPs')/fields/" + `
													"getbytitle('MyMultilineField')"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.Field' }; 
				Description = 'New Field Description' 
			   } | ConvertTo-Json

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 109

#gavdcodebegin 110
Function SpPsRest_DeleteOneFieldFromListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/lists/getbytitle('NewListRestPs')/fields/" + `
													"getbytitle('MyMultilineField')"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 110

#gavdcodebegin 111
Function SpPsRest_BreakSecurityInheritanceListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
				"/_api/lists/getbytitle('NewListRestPs')/" + `
				"breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)"
	
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
Function SpPsRest_ResetSecurityInheritanceListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('NewListRestPs')/resetroleinheritance"
	
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
Function SpPsRest_AddUserToSecurityRoleInListAD
{
	# Inheritance MUST be broken
    # Find the User
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/siteusers?`$select=Id&" + `
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
	                                "'$filter=startswith(Name,'Full Control')"
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$roleId = $dataObject.d.results[0].Id
	Write-Host "RoleId - " $roleId

    # Add the User in the RoleDefinion to the List
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
		"/_api/lists/getbytitle('NewListRestPs')/roleassignments/" + `
        "addroleassignment(principalid=" + $userId + ",roledefid=" + $roleId + ")"
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 113

#gavdcodebegin 114
Function SpPsRest_UpdateUserSecurityRoleInListAD
{
	# Inheritance MUST be broken
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/siteusers?`$select=Id&" + `
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
	                                "`$filter=startswith(Name,'Edit')"
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$roleId = $dataObject.d.results[0].Id
	Write-Host "RoleId - " $roleId

    # Update the User in the RoleDefinion to the List
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
		"/_api/lists/getbytitle('NewListRestPs')/roleassignments/" + `
        "addroleassignment(principalid=" + $userId + ",roledefid=" + $roleId + ")"
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 114

#gavdcodebegin 115
Function SpPsRest_DeleteUserFromSecurityRoleInListAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/siteusers?`$select=Id&" + `
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

    # Remove the User from the List
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('NewListRestPs')/roleassignments/" + `
						"getbyprincipalid(principalid=" + $userId + ")";
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 115

#----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

## Using Legacy Classic Authentication
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#$webUrl = $configFile.appsettings.spUrl
#$userName = $configFile.appsettings.spUserName
#$password = $configFile.appsettings.spUserPw
#SpPsRest_CreateOneList
#SpPsRest_ReadAllLists
#SpPsRest_ReadOneList
#SpPsRest_UpdateOneList
#SpPsRest_DeleteOneList
#SpPsRest_AddOneFieldToList
#SpPsRest_ReadAllFieldsFromList
#SpPsRest_ReadOneFieldFromList
#SpPsRest_UpdateOneFieldInList
#SpPsRest_DeleteOneFieldFromList
#SpPsRest_BreakSecurityInheritanceList
#SpPsRest_ResetSecurityInheritanceList
#SpPsRest_AddUserToSecurityRoleInList
#SpPsRest_UpdateUserSecurityRoleInList
#SpPsRest_DeleteUserFromSecurityRoleInList

## Using Azure AD Authentication
#SpPsRest_CreateOneListAD
#SpPsRest_ReadAllListsAD
#SpPsRest_ReadOneListAD
#SpPsRest_UpdateOneListAD
#SpPsRest_DeleteOneListAD
#SpPsRest_AddOneFieldToListAD
#SpPsRest_ReadAllFieldsFromListAD
#SpPsRest_ReadOneFieldFromListAD
#SpPsRest_UpdateOneFieldInListAD
#SpPsRest_DeleteOneFieldFromListAD
#SpPsRest_BreakSecurityInheritanceListAD
#SpPsRest_ResetSecurityInheritanceListAD
#SpPsRest_AddUserToSecurityRoleInListAD
#SpPsRest_UpdateUserSecurityRoleInListAD
#SpPsRest_DeleteUserFromSecurityRoleInListAD

Write-Host "Done" 
