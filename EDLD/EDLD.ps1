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
Function SpPsRest_CreateOneListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestList')/items"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'NewListItemRestPs'
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
Function SpPsRest_UploadOneDocument  #*** LEGACY CODE ***
{
	$FileInfo = New-Object System.IO.FileInfo("C:\Temporary\TestDocument01.docx")
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath

	$endpointUrl = $WebUrl + "/_api/web/GetFolderByServerRelativeUrl('" + $WebUrlRel + 
				"/TestLibrary')/Files/add(url='" + $FileInfo.Name + "',overwrite=true)"
	$FileContent = [System.IO.File]::ReadAllBytes($FileInfo.FullName)
	$contextInfo = Get-SPOContextInfo  $WebUrl $UserName $Password
	$data = Invoke-RestSPO -Url $endpointUrl -Method Post -UserName $userName -Password `
						$password -Body $FileContent -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue

	$data | ConvertTo-Json
}
#gavdcodeend 002

#gavdcodebegin 003
Function SpPsRest_DownloadOneDocument  #*** LEGACY CODE ***
{
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath

	$endpointUrl = $WebUrl + "/_api/web/GetFileByServerRelativeUrl('" + $WebUrlRel + 
				"/TestLibrary/TestDocument01.docx')/`$value"
	$fileContent = Invoke-RestSPO -Url $endpointUrl -Method Get -UserName $UserName `
									-Password $Password -BinaryStringResponseBody $True
	$fileName = [System.IO.Path]::GetFileName("TestDwload.docx")
	$downloadFilePath = [System.IO.Path]::Combine("C:\Temporary", $fileName)
	[System.IO.File]::WriteAllBytes($downloadFilePath, $fileContent)
}
#gavdcodeend 003

#gavdcodebegin 004
Function SpPsRest_ReadAllListItems  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestList')/items?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 004

#gavdcodebegin 005
Function SpPsRest_ReadOneListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestList')/items(26)?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 005

#gavdcodebegin 006
Function SpPsRest_ReadAllLibraryDocs  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestLibrary')/items?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 006

#gavdcodebegin 007
Function SpPsRest_ReadOneLibraryDoc  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestLibrary')/items(25)?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 007

#gavdcodebegin 008
Function SpPsRest_UpdateOneListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(26)"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'NewListItemCsRest_Updated'
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue -ETag "*" `
						-XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}
#gavdcodeend 008

#gavdcodebegin 009
Function SpPsRest_UpdateOneLibraryDoc  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestLibrary')/items(25)"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'TestDocument01_Updated.docx'
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue -ETag "*" `
						-XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}
#gavdcodeend 009

#gavdcodebegin 010
Function SpPsRest_DeleteOneListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(26)"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 010

#gavdcodebegin 011
Function SpPsRest_DeleteOneLibraryDoc  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestLibrary')/items(25)"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 011

#gavdcodebegin 012
Function SpPsRest_BreakSecurityInheritanceListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/" +
						"items(27)/breakroleinheritance(copyRoleAssignments=false," +
						"clearSubscopes=true)"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}
#gavdcodeend 012

#gavdcodebegin 013
Function SpPsRest_ResetSecurityInheritanceListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/" +
						"items(27)/resetroleinheritance"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue

	$data | ConvertTo-Json
}
#gavdcodeend 013

#gavdcodebegin 014
Function SpPsRest_AddUserToSecurityRoleInListItem  #*** LEGACY CODE ***
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
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(27)/" +
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + 
			$roleId + ")"
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
Function SpPsRest_UpdateUserSecurityRoleInListItem  #*** LEGACY CODE ***
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
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(27)/" +
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + 
			$roleId + ")"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "MERGE"
	$data | ConvertTo-Json
}
#gavdcodeend 015

#gavdcodebegin 016
Function SpPsRest_DeleteUserFromSecurityRoleInListItem  #*** LEGACY CODE ***
{
    # Find the User
	$endpointUrl = $WebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Remove the User from the List
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(27)/" +
					     "roleassignments/getbyprincipalid(principalid=" + $userId + ")";
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "DELETE"
	$data | ConvertTo-Json
}
#gavdcodeend 016

#gavdcodebegin 017
Function SpPsRest_CreateOneFolder  #*** LEGACY CODE ***
{
    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $webUrl + "/_api/web/Folders"
    $myPayload = @{
				__metadata = @{ 'type' = 'SP.Folder' };
				ServerRelativeUrl = $myServerRelativeUrl
			} | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue
    
	$data | ConvertTo-Json
}
#gavdcodeend 017 

#gavdcodebegin 018
Function SpPsRest_ReadAllFolders  #*** LEGACY CODE ***
{
    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $webUrl + "/_api/web/GetFolderByServerRelativeUrl('" + 
                            $myServerRelativeUrl + "')/ListItemAllFields"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 018

#gavdcodebegin 019
Function SpPsRest_RenameOneFolder  #*** LEGACY CODE ***
{
    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $webUrl + "/_api/web/GetFolderByServerRelativeUrl('" + 
                            $myServerRelativeUrl + "')/ListItemAllFields"
    $myPayload = @{
        __metadata = @{ 'type' = 'SP.Data.TestDocumentsItem' };
        Title = 'RestFolderPSRenamed';
        FileLeafRef = 'RestFolderPSRenamend'
    } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue -ETag "*" `
						-XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}
#gavdcodeend 019 

#gavdcodebegin 020
Function SpPsRest_DeleteOneFolder  #*** LEGACY CODE ***
{
    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $webUrl + "/_api/web/GetFolderByServerRelativeUrl('" + 
                                                $myServerRelativeUrl + "')"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 020 

#gavdcodebegin 021
Function SpPsRest_CreateOneAttachment  #*** LEGACY CODE ***
{
	$FileInfo = New-Object System.IO.FileInfo("C:\Temporary\Test.csv")
	$FileContent = [System.IO.File]::ReadAllBytes($FileInfo.FullName)

    $endpointUrl = $webUrl + "/_api/lists/GetByTitle('TestList')" + `
                        "/items(3)/AttachmentFiles/add(FileName='" + $FileInfo.Name + "')"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Body $FileContent -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue
    
	$data | ConvertTo-Json
}
#gavdcodeend 021

#gavdcodebegin 022
Function SpPsRest_ReadAllAttachments  #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/lists/GetByTitle('TestList')" + `
                                                        "/items(3)/AttachmentFiles"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 022

#gavdcodebegin 023
Function SpPsRest_DownloadOneAttachmentByFileName  #*** LEGACY CODE ***
{
	$myFileName = "Test.csv"

	$endpointUrl = $WebUrl + "/_api/lists/GetByTitle('TestList')" +
                                      "/items(3)/AttachmentFiles('" + $myFileName + "')" +
                                      "/$value"
	$fileContent = Invoke-RestSPO -Url $endpointUrl -Method Get -UserName $UserName `
									-Password $Password -BinaryStringResponseBody $True
	$fileName = [System.IO.Path]::GetFileName($myFileName)
	$downloadFilePath = [System.IO.Path]::Combine("C:\Temporary", $fileName)
	[System.IO.File]::WriteAllBytes($downloadFilePath, $fileContent)
}
#gavdcodeend 023

#gavdcodebegin 024
Function SpPsRest_DeleteOneAttachmentByFileName  #*** LEGACY CODE ***
{
	$myFileName = "Test.csv"

	$endpointUrl = $WebUrl + "/_api/lists/GetByTitle('TestList')" + `
                                      "/items(3)/AttachmentFiles('" + $myFileName + "')"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 024

#----------------------------------------------------------------------------------------

# Using Azure AD Authentication through Connect-PnPOnline and an Account/PW App Registration

#gavdcodebegin 101
Function SpPsRest_CreateOneListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
										"/_api/web/lists/getbytitle('TestList')/items"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'NewListItemRestPs'
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
Function SpPsRest_UploadOneDocumentAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$FileInfo = New-Object System.IO.FileInfo("C:\Temporary\TestDocument01.docx")
	$WebUri = [System.Uri]$configFile.appsettings.SiteCollUrl
	$WebUrlRel = $WebUri.AbsolutePath

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
				"/_api/web/GetFolderByServerRelativeUrl('" + $WebUrlRel + `
				"/TestLibrary')/Files/add(url='" + $FileInfo.Name + "',overwrite=true)"
	$FileContent = [System.IO.File]::ReadAllBytes($FileInfo.FullName)

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
Function SpPsRest_DownloadOneDocumentAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$WebUri = [System.Uri]$configFile.appsettings.SiteCollUrl
	$WebUrlRel = $WebUri.AbsolutePath

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
				"/_api/web/GetFileByServerRelativeUrl('" + $WebUrlRel + `
				"/TestLibrary/TestDocument01.docx')/`$value"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$fileContent = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$fileName = [System.IO.Path]::GetFileName("TestDowload.docx")
	$downloadFilePath = [System.IO.Path]::Combine("C:\Temporary", $fileName)
	[System.IO.File]::WriteAllBytes($downloadFilePath, $fileContent)
}
#gavdcodeend 103

#gavdcodebegin 104
Function SpPsRest_ReadAllListItemsAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/web/lists/getbytitle('TestList')/items?" + `
						"$select=Title,Id"

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
Function SpPsRest_ReadOneListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/web/lists/getbytitle('TestList')/items(2)?" + `
						"$select=Title,Id"

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
Function SpPsRest_ReadAllLibraryDocsAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/web/lists/getbytitle('TestLibrary')/items?" + `
						"$select=Title,Id"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 106

#gavdcodebegin 107
Function SpPsRest_ReadOneLibraryDocAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/web/lists/getbytitle('TestLibrary')/items(1)?" + 
						"$select=Title,Id"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 107

#gavdcodebegin 108
Function SpPsRest_UpdateOneListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestList')/items(2)"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'NewListItemCsRest_Updated'
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
#gavdcodeend 108

#gavdcodebegin 109
Function SpPsRest_UpdateOneLibraryDocAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestLibrary')/items(1)"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'TestDocument01_Updated.docx'
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
Function SpPsRest_DeleteOneListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestList')/items(2)"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 110

#gavdcodebegin 111
Function SpPsRest_DeleteOneLibraryDocAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestLibrary')/items(1)"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 111

#gavdcodebegin 112
Function SpPsRest_BreakSecurityInheritanceListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestList')/" + `
						"items(3)/breakroleinheritance(copyRoleAssignments=false," +
						"clearSubscopes=true)"
	
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
Function SpPsRest_ResetSecurityInheritanceListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestList')/" + `
						"items(3)/resetroleinheritance"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 113

#gavdcodebegin 114
Function SpPsRest_AddUserToSecurityRoleInListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
							"/_api/web/siteusers?$select=Id&" + `
	                        "$filter=startswith(Title,'Ade')"
	
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
							"/_api/web/roledefinitions?$select=Id&" + `
	                        "$filter=startswith(Name,'Full Control')"
	
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$roleId = $dataObject.d.results[0].Id
	Write-Host "RoleId - " $roleId

    # Add the User in the RoleDefinion to the Item
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
			"/_api/lists/getbytitle('TestList')/items(3)/" + `
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + `
			$roleId + ")"
	
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 114

#gavdcodebegin 115
Function SpPsRest_UpdateUserSecurityRoleInListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/siteusers?$select=Id&" + `
	                                "$filter=startswith(Title,'Ade')"
	
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
									"/_api/web/roledefinitions?$select=Id&" + `
	                                "$filter=startswith(Name,'Edit')"
	
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$roleId = $dataObject.d.results[0].Id
	Write-Host "RoleId - " $roleId

    # Update the User in the RoleDefinion to the Item
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
			"/_api/lists/getbytitle('TestList')/items(3)/" + `
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + `
			$roleId + ")"
	
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 115

#gavdcodebegin 116
Function SpPsRest_DeleteUserFromSecurityRoleInListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/siteusers?$select=Id&" + `
	                                "$filter=startswith(Title,'MOD')"
	
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
						"/_api/lists/getbytitle('TestList')/items(3)/" + `
					    "roleassignments/getbyprincipalid(principalid=" + $userId + ")";

	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 116

#gavdcodebegin 117
Function SpPsRest_CreateOneFolderAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/Folders"
    $myPayload = @{
				__metadata = @{ 'type' = 'SP.Folder' };
				ServerRelativeUrl = $myServerRelativeUrl
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
#gavdcodeend 117 

#gavdcodebegin 118
Function SpPsRest_ReadAllFoldersAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
							"/_api/web/GetFolderByServerRelativeUrl('" + `
                            $myServerRelativeUrl + "')/ListItemAllFields"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 118

#gavdcodebegin 119
Function SpPsRest_RenameOneFolderAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
							"/_api/web/GetFolderByServerRelativeUrl('" + `
                            $myServerRelativeUrl + "')/ListItemAllFields"
    $myPayload = @{
        __metadata = @{ 'type' = 'SP.Data.TestLibraryItem' };
        Title = 'RestFolderPSRenamed';
        FileLeafRef = 'RestFolderPSRenamed'
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
#gavdcodeend 119 

#gavdcodebegin 120
Function SpPsRest_DeleteOneFolderAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/GetFolderByServerRelativeUrl('" + `
                                    $myServerRelativeUrl + "')"
	
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
#gavdcodeend 120 

#gavdcodebegin 121
Function SpPsRest_CreateOneAttachmentAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$FileInfo = New-Object System.IO.FileInfo("C:\Temporary\Test.csv")
	$FileContent = [System.IO.File]::ReadAllBytes($FileInfo.FullName)

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
					  "/_api/lists/GetByTitle('TestList')" + `
                      "/items(3)/AttachmentFiles/add(FileName='" + $FileInfo.Name + "')"
	$myPayload = $FileContent

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 121

#gavdcodebegin 122
Function SpPsRest_ReadAllAttachmentsAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
								"/_api/lists/GetByTitle('TestList')" + `
                                "/items(3)/AttachmentFiles"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 122

#gavdcodebegin 123
Function SpPsRest_DownloadOneAttachmentByFileNameAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$myFileName = "Test.csv"

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/lists/GetByTitle('TestList')" + `
                                    "/items(3)/AttachmentFiles('" + $myFileName + "')" + `
                                    "/$value"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$fileContent = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$fileName = [System.IO.Path]::GetFileName($myFileName)
	$downloadFilePath = [System.IO.Path]::Combine("C:\Temporary", $fileName)
	[System.IO.File]::WriteAllBytes($downloadFilePath, $fileContent)
}
#gavdcodeend 123

#gavdcodebegin 124
Function SpPsRest_DeleteOneAttachmentByFileNameAD
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken

	$myFileName = "Test.csv"

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
										"/_api/lists/GetByTitle('TestList')" + `
										"/items(3)/AttachmentFiles('" + $myFileName + "')"
	
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
#gavdcodeend 124

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

## Using Legacy Classic Authentication
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#$webUrl = $configFile.appsettings.spUrl
#$userName = $configFile.appsettings.spUserName
#$password = $configFile.appsettings.spUserPw
#SpPsRest_CreateOneListItem
#SpPsRest_UploadOneDocument
#SpPsRest_DownloadOneDocument
#SpPsRest_ReadAllListItems
#SpPsRest_ReadOneListItem
#SpPsRest_ReadAllLibraryDocs
#SpPsRest_ReadOneLibraryDoc
#SpPsRest_UpdateOneListItem
#SpPsRest_UpdateOneLibraryDoc
#SpPsRest_DeleteOneListItem
#SpPsRest_DeleteOneLibraryDoc
#SpPsRest_BreakSecurityInheritanceListItem
#SpPsRest_ResetSecurityInheritanceListItem
#SpPsRest_AddUserToSecurityRoleInListItem
#SpPsRest_UpdateUserSecurityRoleInListItem
#SpPsRest_DeleteUserFromSecurityRoleInListItem
#SpPsRest_CreateOneFolder
#SpPsRest_ReadAllFolders
#SpPsRest_RenameOneFolder
#SpPsRest_DeleteOneFolder
#SpPsRest_CreateOneAttachment
#SpPsRest_ReadAllAttachments
#SpPsRest_DownloadOneAttachmentByFileName
#SpPsRest_DeleteOneAttachmentByFileName

## Using Azure AD Authentication
#SpPsRest_CreateOneListItemAD
#SpPsRest_UploadOneDocumentAD
#SpPsRest_DownloadOneDocumentAD
#SpPsRest_ReadAllListItemsAD
#SpPsRest_ReadOneListItemAD
#SpPsRest_ReadAllLibraryDocsAD
#SpPsRest_ReadOneLibraryDocAD
#SpPsRest_UpdateOneListItemAD
#SpPsRest_UpdateOneLibraryDocAD
#SpPsRest_DeleteOneListItemAD
#SpPsRest_DeleteOneLibraryDocAD
#SpPsRest_BreakSecurityInheritanceListItemAD
#SpPsRest_ResetSecurityInheritanceListItemAD
#SpPsRest_AddUserToSecurityRoleInListItemAD
#SpPsRest_UpdateUserSecurityRoleInListItemAD
#SpPsRest_DeleteUserFromSecurityRoleInListItemAD
#SpPsRest_CreateOneFolderAD
#SpPsRest_ReadAllFoldersAD
#SpPsRest_RenameOneFolderAD
#SpPsRest_DeleteOneFolderAD
#SpPsRest_CreateOneAttachmentAD
#SpPsRest_ReadAllAttachmentsAD
#SpPsRest_DownloadOneAttachmentByFileNameAD
#SpPsRest_DeleteOneAttachmentByFileNameAD

Write-Host "Done" 
